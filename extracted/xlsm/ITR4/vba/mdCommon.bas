Attribute VB_Name = "mdCommon"
Option Explicit

Public Msgbox_I As String
Public ActiveCellRange As String
Public DefinedgridNameRange As Variant
Public msgError As Variant
Public vbMessageCaption As Variant

Public Total_SheetCount As Long
Public ScheduleName As Variant
Public EnabledSchedule As Variant
Public PrintYN As Variant
Public CalculateTaxCheckFlag As Boolean
'Local variables
Dim strCurrActiveCellRange, strNewActiveCellRange As String
Dim pwd As String

Sub CloseMsg()
End
End Sub
Function GetMergedAddressCell(rnge As Range, i As Long) As Variant
    Dim sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue As String
    Dim sTempCellValueInt As String

    sTempCellValue = Replace(rnge.AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
    sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
    sTempLastCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C6
    sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
    
    sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + i - 1)
    GetMergedAddressCell = sNewCellValue
End Function

Function UVCase(valV1 As Variant) As Variant
    If (valV1 = "") Then
        valV1 = 0
    End If
        UVCase = valV1
End Function

Function isdropdownblank(field As Variant) As Boolean
    isdropdownblank = False
    If Mid(field, 1, 1) = "(" Then
        isdropdownblank = True
    End If
    If Mid(field, 1, 1) = "" Or IsEmpty(field) Then
        isdropdownblank = True
    End If
End Function
'Ayush_06/05/2025
Function isdropdownblankDD(field As Variant) As Boolean
    isdropdownblankDD = False
    If Mid(field, 1, 2) = "(S" Then
        isdropdownblankDD = True
    End If
    If Mid(field, 1, 1) = "" Or IsEmpty(field) Then
        isdropdownblankDD = True
    End If
End Function

Function checkfieldspecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "|", "`", "_") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter = False
            Exit Function
        End If
        Next
    Next
End Function

'Ayush_Leagcy3
Function checkfieldspecialcharacter80DD_80U(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter80DD_80U = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "|", "`", "_", ".", ",") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter80DD_80U = False
            Exit Function
        End If
        Next
    Next
End Function

Function checkfieldspecialcharacter3(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter3 = True
    Dim arr As Variant
    arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter3 = False
            Exit Function
        End If
        Next
    Next
End Function
Function checkfieldSuperSpecialcharactername(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharactername = True
    Dim arr As Variant
    arr = Array(">", "<", "&")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharactername = False
            Exit Function
        End If
        Next
    Next
End Function

Function chkCompulsory(field As Variant) As Boolean
    chkCompulsory = True
    If Len(Trim(field)) <= 0 Then
        chkCompulsory = False
    End If
    If IsEmpty(field) Then
    chkCompulsory = False
    End If
End Function

Function getmsgstate() As String
    pwd = Sheet7.Range("pwd")
    getmsgstate = pwd
End Function

Function calculateAge(dob As Variant) As Long
On Error Resume Next
'Change-08.12.2022.103.23
    'calculateAge = (2022) - val(Mid(dob, 7, 4))
    'Change Year from 2023 to 2024 Bindu
    'Ayush_25-26
'    calculateAge = (2024) - val(Mid(dob, 7, 4))
'     calculateAge = (2025) - val(Mid(dob, 7, 4))
    'Ayush_26-27
    calculateAge = (Sheet8.Range("DOB_Year").value) - val(Mid(dob, 7, 4))
    
    If 4 < val(Mid(dob, 4, 2)) Then
        calculateAge = calculateAge - 1
    ElseIf val(Mid(dob, 4, 2)) = 4 And 1 < val(Mid(dob, 1, 2)) Then
        calculateAge = calculateAge - 1
    End If
End Function


Function ValidateIFSCList(IFSC As String) As Boolean
    ValidateIFSCList = False
    Dim IFSCCodeSet1 As Range
    Dim IFSCCodeSet2 As Range
    Dim IFSCCodeSet3 As Range
    Dim IFSCCodeSet4 As Range
    Dim IFSCCodeSet5 As Range
    
    Dim IFSC1 As Range
    Dim IFSC2 As Range
    Dim IFSC3 As Range
    Dim IFSC4 As Range
    Dim IFSC5 As Range


    
    Set IFSCCodeSet1 = Range("IFSCCodeSet1")
    Set IFSCCodeSet2 = Range("IFSCCodeSet2")
    Set IFSCCodeSet3 = Range("IFSCCodeSet3")
    Set IFSCCodeSet4 = Range("IFSCCodeSet4")
    Set IFSCCodeSet5 = Range("IFSCCodeSet5")
    
    
    
    For Each IFSC1 In IFSCCodeSet1
        If IFSC1.value = IFSC Then
        ValidateIFSCList = True
        Exit For
        End If
    Next

    If Not ValidateIFSCList Then
        For Each IFSC2 In IFSCCodeSet2
            If IFSC2.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    If Not ValidateIFSCList Then
        For Each IFSC3 In IFSCCodeSet3
            If IFSC3.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    If Not ValidateIFSCList Then
        For Each IFSC4 In IFSCCodeSet4
            If IFSC4.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    
    If Not ValidateIFSCList Then
        For Each IFSC5 In IFSCCodeSet5
            If IFSC5.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    ValidateIFSCList = True
End Function
Public Function ChkAlphabet(Char As String) As Boolean
    ChkAlphabet = True
    If ((asc(Char) < 65) Or (asc(Char) > 90)) Then
        ChkAlphabet = False
    End If
End Function
Function ChkAlphabetP(Char As String) As Boolean
    ChkAlphabetP = True
    If ((asc(Char) <> 80)) Then
        ChkAlphabetP = False
    End If
End Function


'Public Function CheckDate(dateEntered As Variant, Optional HUF As Variant = 0) As Long
''0: Date Ok
''1: Invalid Date Format
''2: Date not in Range
'
'On Error Resume Next
'    CheckDate = 0
'    If Trim(dateEntered) = "" Or Not IsEmpty(dateEntered) Then
'        If Not FormatNCheckDate(dateEntered, HUF) Then
'            CheckDate = 1
'            Exit Function
'        End If
'
'        If Not ChkMaxDate(dateEntered, "02/04/2022") Then
'            CheckDate = 2
'            Exit Function
'        Else
'
'        End If
'    End If
' End Function

Public Function FormatNCheckDate(ByRef dt As Variant, Optional HUF As Variant = 0, Optional Year As String = "2022") As Boolean
On Error Resume Next
    FormatNCheckDate = True
    If Len(dt) > 0 Then
        If Mid(dt, 3, 1) <> "/" Then 'Checking between dd/mm
           FormatNCheckDate = False
           Exit Function
        End If
        
        If Mid(dt, 6, 1) <> "/" Then 'Checking between mm/yyyy
            FormatNCheckDate = False
            Exit Function
        End If
        
        If val(Mid(dt, 1, 2)) < 0 Or val(Mid(dt, 1, 2)) > 31 Then FormatNCheckDate = False
        If val(Mid(dt, 4, 2)) < 0 Or val(Mid(dt, 4, 2)) > 12 Then FormatNCheckDate = False
        If (HUF = 1) Then
            If val(Mid(dt, 7, 4)) < 1 Then FormatNCheckDate = False
        End If
        If Not IsDate(dt) Then FormatNCheckDate = False
    End If
End Function
'Public Function ChkMaxDatePrev(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
'On Error Resume Next
'Dim Year, month, dat As Variant
'
'     ChkMaxDatePrev = True
'     If Len(dateEntered) > 0 Then
'
'     Year = val(Mid(dateEntered, 7, 4))
'     month = val(Mid(dateEntered, 4, 2))
'     dat = val(Mid(dateEntered, 1, 2))
'
'        If Year > 2021 Then
'            ChkMaxDatePrev = False
'            Exit Function
'        Else
'            If Year = 2021 Then
'                If month > 4 Then
'                    ChkMaxDatePrev = False
'                    Exit Function
'                Else
'                    If month = 4 Then
'                        If dat > 1 Then
'                            ChkMaxDatePrev = False
'                            Exit Function
'                        Else
'                            If dat = 1 Then
'                                ChkMaxDatePrev = False
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'     End If
'End Function

'False : greater
'Public Function ChkMaxDate(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
'On Error Resume Next
'Dim Year, month, dat As Variant
'
'     ChkMaxDate = True
'     If Len(dateEntered) > 0 Then
'
'     Year = val(Mid(dateEntered, 7, 4))
'     month = val(Mid(dateEntered, 4, 2))
'     dat = val(Mid(dateEntered, 1, 2))
'
'        If Year > 2022 Then
'            ChkMaxDate = False
'            Exit Function
'        Else
'            If Year = 2022 Then
'                If month > 4 Then
'                    ChkMaxDate = False
'                    Exit Function
'                Else
'                    If month = 4 Then
'                        If dat > 1 Then
'                            ChkMaxDate = False
'                            Exit Function
'                        Else
'                            If dat = 1 Then
'                                ChkMaxDate = False
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'     End If
'End Function

Public Function CheckSpecialCharacter(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacter = True
    specialCharArray = Array("*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacter = False
            Exit Function
        End If
        Next
    Next
End Function

Public Function CheckEmptyDDList(item As Variant) As Boolean
On Error Resume Next
CheckEmptyDDList = False

If IsEmpty(item) Or Mid(item, 1, 1) = "(" Or item = "" Then CheckEmptyDDList = True

End Function
Function insertRowUnderSectionWithFormulaOne(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        'nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         nRows = 1
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    
    
    'Ayush_21
       LinkCheckBoxes
    Dim chk As CheckBox
    For Each chk In ActiveSheet.CheckBoxes
        If mdCommon.onlyDigits(chk.LinkedCell, "I") = mdCommon.onlyDigits(sNewCellValue, "I") Then
           chk.value = False
        End If
    Next
    
    insertRowUnderSectionWithFormulaOne = nRows

If ActiveSheet.name = "80G" Then
    AutoSelectionforStateCode
End If

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function
Function insertRowUnderSectionWithFormula(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
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
   sPassword = mdCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula = nRows

If ActiveSheet.name = "80G" Then
    AutoSelectionforStateCode
End If

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function
'Ayush_26-27_80CCC_80CCD1_80CCD1B
Function insertRowUnderSectionWithFormula_IncomeDetails_80CCC(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
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
   sPassword = mdCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            'Ayush_26-27
            If sTempFirstCellValue = "" Then
            sRangeValue = sTempCellValue + ":" + sNewCellValue
            End If
            '----------
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
             'Ayush_26-27
            If sTempFirstCellValue = "" Then
            sRangeValue = sTempCellValue + ":" + sNewCellValue
            End If
            '----------
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_IncomeDetails_80CCC = nRows

If ActiveSheet.name = "80G" Then
    AutoSelectionforStateCode
End If

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function
'-------------------------------

Function insertRowUnderSectionWithFormulaTDS(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
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
   sPassword = mdCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Ayush_03/06
        
        Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
     
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
            
            
            'Ayush_legacy38
                Range("I" & rr.row).Locked = False
                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
                '-----------------
                
                
                Range("k" & rr.row).Locked = True 'J to K as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
                Range("k" & rr.row).Interior.Color = "&HD8D8D8" 'J to K as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
                Range("l" & rr.row).Locked = True ' K to L as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
                Range("l" & rr.row).Interior.Color = "&HD8D8D8" 'K to L as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
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
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
            If gridRange(iCount) = "TDS2.TdsCredit" Or _
                gridRange(iCount) = "TDS3.TdsCredit" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
                gridRange(iCount) = "TDSoth2.DeductedYear" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
            If gridRange(iCount) = "TDS2.TdsCredit" Or _
                gridRange(iCount) = "TDS3.TdsCredit" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
                gridRange(iCount) = "TDSoth2.DeductedYear" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormulaTDS = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function

'Ayush_03/06
Function insertRowUnderSectionWithFormulaTDS3(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
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
   sPassword = mdCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Ayush_03/06
        
        Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
     
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
            
            
            'Ayush_legacy38
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
                '-----------------
                
                
                Range("L" & rr.row).Locked = True 'J to K as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
                Range("L" & rr.row).Interior.Color = "&HD8D8D8" 'J to K as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
                Range("M" & rr.row).Locked = True ' K to L as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
                Range("M" & rr.row).Interior.Color = "&HD8D8D8" 'K to L as new field added in TDS2(i) on Feb 3rd 2025 by Bindu
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
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
            If gridRange(iCount) = "TDS2.TdsCredit" Or _
                gridRange(iCount) = "TDS3.TdsCredit" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
                gridRange(iCount) = "TDSoth2.DeductedYear" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
            If gridRange(iCount) = "TDS2.TdsCredit" Or _
                gridRange(iCount) = "TDS3.TdsCredit" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
                gridRange(iCount) = "TDSoth2.DeductedYear" Then
                Range(sAddingRangeValue).value = "(Select)"
            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormulaTDS3 = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function

Function searchLastRow(ByVal gridRangeName As String) As String
On Error Resume Next
 strCurrActiveCellRange = Replace(ActiveSheet.Range(gridRangeName).AddressLocal, "$", "")
 strNewActiveCellRange = Mid(strCurrActiveCellRange, InStr(1, strCurrActiveCellRange, ":") + 1, Len(strCurrActiveCellRange))
 ActiveSheet.Range(strNewActiveCellRange).Select
 searchLastRow = strCurrActiveCellRange
End Function


Function onlyDigits(ByVal s As String, ExtracType As String) As String
On Error Resume Next
    Dim retval As String
    Dim i As Long

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


Function ValidateTantype_text(strName As Variant) As Boolean
    ValidateTantype_text = True
    Dim len1, j As Long
    Dim s1 As String
    len1 = Len(strName)
    For j = 1 To len1
        s1 = Mid(strName, j, 1)
        If (((asc(s1) >= 65) And (asc(s1) <= 90)) Or (asc(s1) = 45)) Then
        Else
            ValidateTantype_text = False
        End If
    Next
End Function




Sub CreateArray()
    Dim i, s As Long

    Total_SheetCount = 0

    For i = 9 To 20
        s = i
        If Not Sheet9.Range("AO" & s).value = "" Then
           Total_SheetCount = Total_SheetCount + 1
        End If
    Next
    Total_SheetCount = Total_SheetCount + 8

    ReDim ScheduleName(Total_SheetCount)
        For i = 9 To Total_SheetCount
        s = i
        ScheduleName(i) = Sheet9.Range("AO" & s)
    Next
    
    ReDim PrintYN(Total_SheetCount)
        For i = 9 To Total_SheetCount
            s = i
            PrintYN(i) = Sheet9.Range("AP" & s)
        Next
End Sub

Function printWorkSheet()
    Dim PrintMsgOP As Long
    Dim sourceSheet As Worksheet

    PrintMsgOP = MsgBox("Do you want to preview the workbook for printing?", vbOKCancel, "Print preview")
    If PrintMsgOP = 1 Then
           With Sheet1.PageSetup
                 .PrintArea = Sheet1.Range("Print_Area").Address
                 .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = 100
                .Orientation = xlLandscape
            End With
'            With Sheet2.PageSetup
'                  .BlackAndWhite = True
'                .CenterHorizontally = True
'                .CenterVertically = False
'                .LeftMargin = 0.4
'                .RightMargin = 0.4
'                .PaperSize = xlPaperA4
'                .FitToPagesWide = 1
'                .Orientation = xlLandscape
'            End With
             With Sheet3.PageSetup
              .PrintArea = Sheet3.Range("Print_Area").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet4.PageSetup
              .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
              With Sheet5.PageSetup
              .PrintArea = Sheet5.Range("Print_Area").Address
               .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            
              With Sheet6.PageSetup
                  .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
        
            With Sheet13.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
                .Zoom = False
            End With
            
            
            'Added by Aavula Naresh
            
            'Part A Gen_139(8A)
            
            With Sheet201.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
                
            End With
            
            'PartB-ATI
            With Sheet202.PageSetup
                .PrintArea = Sheet202.Range("Print_Area").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            '80GGC
            With Sheet11.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            '80DD-80U
            With Sheet14.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            'EA 10(13A))
             With Sheet15.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            '80D
             With Sheet16.PageSetup
                .PrintArea = Sheet16.Range("Print_Area").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            'TCS
            With Sheet17.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            'IT
             With Sheet18.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            '80C
            With Sheet19.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            '80E,80EE,80EEA,80EEB
            With Sheet20.PageSetup
            
                .PrintArea = Sheet20.Range("Print_Area").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            '24(b)
            With Sheet21.PageSetup
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = False
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
                 
     ' rowCount = ThisWorkbook.Worksheets.count
    
    ThisWorkbook.PrintOut , Preview:=True
  '  From:=1, To:=rowcount
 Else
 End If
 'Sheet1.DisplayPageBreaks = False
 'Sheet2.DisplayPageBreaks = False
 'Sheet3.DisplayPageBreaks = False
 'Sheet4.DisplayPageBreaks = False
 'Sheet5.DisplayPageBreaks = False
 'Sheet6.DisplayPageBreaks = False
 'Sheet5.DisplayPageBreaks = False
 
 
  
 
End Function


Function CheckVerPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckVerPAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckVerPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckVerPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckVerPAN = False
            Exit Function
        End If
        
        'Ayush_PAN===
'       If Not Mid(PAN, 4, 1) = "P" Then
'            CheckVerPAN = False
'            Exit Function
'        End If
'=================
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckVerPAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckVerPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckVerPAN = False
            Exit Function
        End If
               
    End If
End Function


Function checkfieldSuperSpecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharacter = True
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



Sub Lock_TDS()
Dim sPassword As String
'On Error Resume Next
'------------------Un Protect--------------------
  sPassword = mdCommon.getmsgstate
   Sheet17.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------

Dim rangecells As Range ' rangecells1, rangecells2, rangecells3 As Range
Dim mIntCount, i As Long ', mIntCount1, mIntCount2, mIntCount3, i As Long
i = 0
'Set rangecells = Sheet4.Range("TDSoth.TD_PAN").Cells
'mIntCount = Sheet4.Range("TDSoth.TD_PAN").Count
'Set rangecells1 = Sheet4.Range("TDSoth.TDS_PAN").Cells
'mIntCount1 = Sheet4.Range("TDSoth.TDS_PAN").Count
'Set rangecells2 = Sheet4.Range("TDSoth2.TD_PAN").Cells
'mIntCount2 = Sheet4.Range("TDSoth2.TD_PAN").Count
'Set rangecells3 = Sheet4.Range("TDSoth2.TDS_PAN").Cells
'mIntCount3 = Sheet4.Range("TDSoth2.TDS_PAN").Count
'
'
'For i = 1 To mIntCount
'    rangecells.item(i).value = ""
'    rangecells.item(i).Locked = True
'    rangecells.item(i).Interior.Color = (&HD8D8D8)
'Next
'
'For i = 1 To mIntCount1
'    rangecells1.item(i).value = ""
'    rangecells1.item(i).Locked = True
'    rangecells1.item(i).Interior.Color = (&HD8D8D8)
'Next
'
'For i = 1 To mIntCount
'    rangecells2.item(i).value = ""
'    rangecells2.item(i).Locked = True
'    rangecells2.item(i).Interior.Color = (&HD8D8D8)
'Next
'
'For i = 1 To mIntCount1
'    rangecells3.item(i).value = ""
'    rangecells3.item(i).Locked = True
'    rangecells3.item(i).Interior.Color = (&HD8D8D8)
'Next



'Set rangecells = Sheet17.Range("TCS.AmtClaimedBySpouse").Cells
'mIntCount = Sheet17.Range("TCS.AmtClaimedBySpouse").Count
'
'For i = 1 To mIntCount
'    rangecells.item(i).value = ""
'    rangecells.item(i).Locked = True
'    rangecells.item(i).Interior.Color = (&HD8D8D8)
'Next
        
'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
  Sheet17.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub


Sub UnLock_TDS()
'On Error Resume Next
Dim sPassword As String
'------------------Un Protect--------------------
   sPassword = mdCommon.getmsgstate
   Sheet17.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------

Dim rangecells As Range ', rangecells1, rangecells2, rangecells3 As Range
Dim mIntCount, i As Long ', mIntCount1, mIntCount2, mIntCount3, i As Long
i = 0
'Set rangecells = Sheet4.Range("TDSoth.TD_PAN").Cells
'mIntCount = Sheet4.Range("TDSoth.TD_PAN").Count
'Set rangecells1 = Sheet4.Range("TDSoth.TDS_PAN").Cells
'mIntCount1 = Sheet4.Range("TDSoth.TDS_PAN").Count
'
'Set rangecells2 = Sheet4.Range("TDSoth2.TD_PAN").Cells
'mIntCount2 = Sheet4.Range("TDSoth2.TD_PAN").Count
'Set rangecells3 = Sheet4.Range("TDSoth2.TDS_PAN").Cells
'mIntCount3 = Sheet4.Range("TDSoth2.TDS_PAN").Count

'For i = 1 To mIntCount
'    rangecells.item(i).Locked = False
'    rangecells.item(i).Interior.Color = (&HCCFFCC)
'Next
'
'For i = 1 To mIntCount1
'    rangecells1.item(i).Locked = False
'    rangecells1.item(i).Interior.Color = (&HCCFFCC)
'Next
'
'For i = 1 To mIntCount2
'    rangecells2.item(i).Locked = False
'    rangecells2.item(i).Interior.Color = (&HCCFFCC)
'Next
'
'For i = 1 To mIntCount3
'    rangecells3.item(i).Locked = False
'    rangecells3.item(i).Interior.Color = (&HCCFFCC)
'Next
'Set rangecells = Sheet17.Range("TCS.AmtClaimedBySpouse").Cells
'mIntCount = Sheet17.Range("TCS.AmtClaimedBySpouse").Count
'
'For i = 1 To mIntCount
'    rangecells.item(i).Locked = False
'    rangecells.item(i).Interior.Color = (&HCCFFCC)
'Next
        
'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
  Sheet17.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub
Function CheckDoneePAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 1 is for individuals .So,4th character of PAN should be "P"

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




Sub Worksht_Active()
On Error Resume Next
Application.EnableEvents = False
Dim sPassword As String
 '------------------Un Protect--------------------
   sPassword = mdCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------
     
'    If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") Then
'
'                'Sheet5.Range("Ver.AssesseeVerName").value = ""
'                'Sheet5.Range("Ver.PAN").value = ""
'
'                Sheet1.Range("sheet1.PortugeseCC5A").value = "(Select)"
''                Sheet1.Range("AD16:AH16").Locked = True
''                Sheet1.Range("AD16:AH16").Interior.Color = (&HD8D8D8)
'
'                'Sheet1.Range("sheet1.aadharFlag").value = "(Select)"
'                'Sheet1.Range("U21:AE21").Locked = True
'                'Sheet1.Range("U21:AE21").Interior.Color = (&HD8D8D8)
'
'                'Sheet1.Range("sheet1.aadharValue").value = ""
'                'Sheet1.Range("AM21:AT21").Locked = True
'                'Sheet1.Range("AM21:AT21").Interior.Color = (&HD8D8D8)
'
'                Sheet1.Range("sheet1.PANOFSPOUSE").value = ""
''                Sheet1.Range("AP16:AT16").Locked = True
''                Sheet1.Range("AP16:AT16").Interior.Color = (&HD8D8D8)
'
''            Else
''                'Range("sheet1.aadharFlag").value = "(Select)"
''
''                Range("AD16:AH16").Locked = False
''                Range("AD16:AH16").Interior.Color = (&HCCFFCC)
''
''                Range("U21:AE21").Locked = False
''                Range("U21:AE21").Interior.Color = (&HCCFFCC)
'
'            End If
    
    

    
    
    'If Sheet1.Range("sheet1.aadharFlag").value <> "Yes" Then
    '    Sheet1.Range("sheet1.aadharValue").value = ""
    '    Sheet1.Range("AM21:AT21").Locked = True
    '    Sheet1.Range("AM21:AT21").Interior.Color = (&HD8D8D8)
    'Else
    '    Sheet1.Range("AM21:AT21").Locked = False
    '    Sheet1.Range("AM21:AT21").Interior.Color = (&HCCFFCC)
    'End If
      
    
'    If Sheet1.Range("sheet1.PortugeseCC5A").value <> "Yes" Then
'
'        mdCommon.Lock_TDS
'
'        Sheet1.Range("sheet1.PANOFSPOUSE").value = ""
''        Sheet1.Range("AP16:AT16").Locked = True
''        Sheet1.Range("AP16:AT16").Interior.Color = (&HD8D8D8)
'
''        Sheet4.Range("TDSoth.TD_PAN").value = ""
''        Sheet4.Range("TDSoth.TD_PAN").Locked = True
''        Sheet4.Range("TDSoth.TD_PAN").Interior.Color = (&HD8D8D8)
''
''        Sheet4.Range("TDSoth.TDS_PAN").value = ""
''        Sheet4.Range("TDSoth.TDS_PAN").Locked = True
''        Sheet4.Range("TDSoth.TDS_PAN").Interior.Color = (&HD8D8D8)
'
'
'    Else
'
'        mdCommon.UnLock_TDS
'
''        Sheet1.Range("AP16:AT16").Locked = False
''        Sheet1.Range("AP16:AT16").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("TDSoth.TD_PAN").Locked = False
'        Sheet1.Range("TDSoth.TD_PAN").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("TDSoth.TDS_PAN").Locked = False
'        Sheet1.Range("TDSoth.TDS_PAN").Interior.Color = (&HCCFFCC)
'
'
'    End If
      
      
'   If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17 Then
'        'Range("sheet1.ReturnType1").value = "Revised"
'    ElseIf (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 13 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 14 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 15 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 16 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 20) Then
'        'Range("sheet1.ReturnType1").value = "Original"
'    ElseIf Range("sheet1.ReturnFileSec1").value = "(Select)" Then
'        Range("sheet1.ReturnType1").value = "(Select)"
'    End If
    
    
    If ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 13) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 14) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 15) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 16)) Then
        Range("sheet1.ReceiptNo").value = ""
'        Range("U18:AE18").Locked = True
'        Range("U18:AE18").Interior.Color = (&HD8D8D8)
        
        Range("sheet1.OrigRetFiledDate").value = ""
'        Range("AM18:AT18").Locked = True
'        Range("AM18:AT18").Interior.Color = (&HD8D8D8)
        
        Range("sheet1.NoticeNo").value = ""
'        Range("U19:AE19").Locked = True
'        Range("U19:AE19").Interior.Color = (&HD8D8D8)
        
'        Range("AM19:AT19").Locked = False
'        Range("AM19:AT19").Interior.Color = (&HCCFFCC)
    End If
    
    
    If ((Range("sheet1.ReturnFileSec1").value = "(Select)") Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18)) Then
'        Range("U18:AE18").Locked = False
'        Range("U18:AE18").Interior.Color = (&HCCFFCC)
        
'        Range("AM18:AT18").Locked = False
'        Range("AM18:AT18").Interior.Color = (&HCCFFCC)
'
'        Range("U19:AE19").Locked = False
'        Range("U19:AE19").Interior.Color = (&HCCFFCC)
'
'        Range("AM19:AT19").Locked = False
'        Range("AM19:AT19").Interior.Color = (&HCCFFCC)
        
    End If
    
    
    If ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 20)) Then
        Range("sheet1.ReceiptNo").value = ""
'        Range("U18:AE18").Locked = True
'        Range("U18:AE18").Interior.Color = (&HD8D8D8)
'
        Range("sheet1.OrigRetFiledDate").value = ""
'        Range("AM18:AT18").Locked = True
'        Range("AM18:AT18").Interior.Color = (&HD8D8D8)
        
        Range("sheet1.NoticeNo").value = ""
'        Range("U19:AE19").Locked = True
'        Range("U19:AE19").Interior.Color = (&HD8D8D8)
        
        Range("sheet1.NoticeDate").value = ""
'        Range("AM19:AT19").Locked = True
'        Range("AM19:AT19").Interior.Color = (&HD8D8D8)
    End If
    
    
    
    If ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17)) Then
        
'        Range("U18:AE18").Locked = False
'        Range("U18:AE18").Interior.Color = (&HCCFFCC)
'
'        Range("AM18:AT18").Locked = False
'        Range("AM18:AT18").Interior.Color = (&HCCFFCC)
        
        Range("sheet1.NoticeNo").value = ""
'        Range("U19:AE19").Locked = True
'        Range("U19:AE19").Interior.Color = (&HD8D8D8)
'
'        Range("AM19:AT19").Locked = True
'        Range("AM19:AT19").Interior.Color = (&HD8D8D8)
    End If
      
      
      
      
        
'----------------Lock Password-------------------START---
  'Sheet4.Protect Password:=sPassword
  Sheet1.Protect Password:=sPassword
'----------------Lock Password-------------------END-----


'If Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = 17 Then
'        'Sheet1.Range("sheet1.ReturnType1").value = "Revised"
'    ElseIf (Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = 11 Or Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = 12 Or Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = 13) Then
'        'Sheet1.Range("sheet1.ReturnType1").value = "Original"
'End If

If (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR") Then
                Lock44AD
            Else
                Unlock44AD
End If




Sheet1.Activate
        

Application.EnableEvents = True
End Sub


Sub AutoSelectionforStateCode()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
'Grid 1
 mIntCells = Sheet6.Range("Per10080G.StateCode").count
 Set rangecells = Sheet6.Range("Per10080G.StateCode").Cells

 For mIntCtr = 1 To mIntCells
    If rangecells.item(mIntCtr) = "" Then rangecells.item(mIntCtr) = "(Select)"
 Next
 
'Grid 2
 mIntCells = Sheet6.Range("PerNO5080G.StateCode").count
 Set rangecells = Sheet6.Range("PerNO5080G.StateCode").Cells

 For mIntCtr = 1 To mIntCells
    If rangecells.item(mIntCtr) = "" Then rangecells.item(mIntCtr) = "(Select)"
 Next
 
'Grid 3
 mIntCells = Sheet6.Range("PerYES10080G.StateCode").count
 Set rangecells = Sheet6.Range("PerYES10080G.StateCode").Cells

 For mIntCtr = 1 To mIntCells
    If rangecells.item(mIntCtr) = "" Then rangecells.item(mIntCtr) = "(Select)"
 Next
 
'Grid 4
 mIntCells = Sheet6.Range("Per5080G.StateCode").count
 Set rangecells = Sheet6.Range("Per5080G.StateCode").Cells

 For mIntCtr = 1 To mIntCells
    If rangecells.item(mIntCtr) = "" Then rangecells.item(mIntCtr) = "(Select)"
 Next

End Sub




Function checkspecialcharacterForPinCode(field As Variant) As Boolean
    Dim i, j As Long
    checkspecialcharacterForPinCode = True
    Dim arr As Variant
    arr = Array(".", "@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkspecialcharacterForPinCode = False
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


Public Function CheckSpecialCharacterdot(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacterdot = True
    specialCharArray = Array(".")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacterdot = False
            Exit Function
        End If
        Next
    Next
End Function

Private Sub ButtonHyperlink()
On Error Resume Next
'fmsgboxsmall "After e-payment of Tax, the details of amount paid, Challan No., BSR Code, etc. should be filled in Schedule IT of Income Tax Return before submission of the return to claim the challan" ', vbInformation, "Alert"
fmsgboxsmall "You will be redirected to the to e-Pay Tax. After e-payment of Tax, the details of amount paid, Challan No., BSR Code, etc. should be filled in Schedule IT of ITR before submission of the return to claim the challan .."
ActiveWorkbook.FollowHyperlink _
Address:="https://eportal.incometax.gov.in/iec/foservices/#/e-pay-tax-prelogin/user-details" 'Changed by Ayush on 04/06/2025
'Address:="https://onlineservices.tin.egov-nsdl.com/etaxnew/tdsnontds.jsp"
End Sub
'https://eportal.incometax.gov.in/iec/foservices/#/e-pay-tax-prelogin/user-details
Sub LinkCheckBoxes()
Dim chk As CheckBox
Dim lCol As Long
lCol = 7 'number of columns to the right for link
For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = _
         .TopLeftCell.Offset(0, lCol).Address
   End With
Next chk
End Sub
'newly added by Sai for 80GGC Addrow
Function insertRowUnderSectionWithFormula80GGC(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
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
   sPassword = mdCommon.getmsgstate
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
                Range("E" & rr.row).Locked = False
                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
                Range("F" & rr.row).Locked = False
                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
                Range("G" & rr.row).Locked = False
                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
                Range("K" & rr.row).Locked = False
                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("K" & rr.row).Locked = True
'                Range("K" & rr.row).Interior.Color = "&HD8D8D8"
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
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
    insertRowUnderSectionWithFormula80GGC = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function
'Ayush_09/04
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
   sPassword = mdCommon.getmsgstate
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
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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

'Ayush_09/04
Function insertRowUnderSectionWithFormula80C(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
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
   sPassword = mdCommon.getmsgstate
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
'                Range("E" & rr.row).Locked = False
'                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
                Range("F" & rr.row).Locked = False
                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
                Range("G" & rr.row).Locked = False
                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("H" & rr.row).Locked = False
'                Range("H" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("I" & rr.row).Locked = False
'                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("J" & rr.row).Locked = False
'                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("K" & rr.row).Locked = False
'                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("L" & rr.row).Locked = False
'                Range("L" & rr.row).Interior.Color = (&HCCFFCC)
'
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
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
    insertRowUnderSectionWithFormula80C = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function
'Ayush_09/04
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
   sPassword = mdCommon.getmsgstate
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
'                Range("J" & rr.row).Locked = False
'                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
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
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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

Function state_Validation(ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    state_Validation = True
    
    Dim PinCode As Range
        
            If Range(pin_trgt_adrs).value <> "" Then
              Dim state1 As String
                state1 = Application.VLookup(Range(pin_trgt_adrs).value, Sheet8.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Range(state_trgt_adrs).value, 4)) Or mdCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value)) Then
                    state_Validation = False
    
                    'Range(pin_trgt_adrs).value = ""
                End If
            End If
         
                   
End Function
 
                 
Function PinState_codes_validation(sheetname1 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String, ByVal country_trgt_adrs As String, ByVal Zip_trgt_adrs As String)
   
Dim PinCode1 As String
Dim DupFlag As Boolean

DupFlag = False
PinCode1 = Application.IsError(Application.VLookup(sheetname1.Range(pin_trgt_adrs).value, Sheet8.Range("Duplicate_Pincode_list"), 1, False))
If PinCode1 = "False" Then DupFlag = True

If DupFlag = True Then
   If sheetname1.Range(state_trgt_adrs).value <> "" Then
       If Not mdCommon.StateMatchesPin(Range(state_trgt_adrs).value, sheetname1.Range(pin_trgt_adrs).value) Then
           sheetname1.Range(state_trgt_adrs).value = "(Select)"
       End If
   End If
End If
                 
If DupFlag = False Then
   Dim StateName
   Dim StateCode

   StateName = Application.VLookup(sheetname1.Range(pin_trgt_adrs).value, Sheet8.Range("All_PinCode_V"), 2, False)
   If IsError(StateName) Then
       If sheetname1.Range(pin_trgt_adrs).value <> "" Then
           
           MsgBox ("Invalid Pincode")
           sheetname1.Range(pin_trgt_adrs).ClearContents
       End If
   Else
       StateCode = Application.VLookup(StateName, Sheet8.Range("EW:EX"), 2, False)
       sheetname1.Range(state_trgt_adrs).value = StateCode
       sheetname1.Range(country_trgt_adrs).value = "91-INDIA"
       sheetname1.Unprotect Password:=mdCommon.getmsgstate
   End If
End If


End Function



Function StateMatchesPin(StateCode As Variant, Pincode_pag2) As Boolean
    StateMatchesPin = False
    Dim state1, state2 As Variant
    
    
    state1 = Application.VLookup(Pincode_pag2, Sheet8.Range("Duplicate_V"), 2, False)
    state2 = Application.VLookup(Pincode_pag2, Sheet8.Range("Duplicate_V"), 3, False)

    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin = False
    Exit Function
    End If
    
    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = state2 Then
    StateMatchesPin = True
    End If
    
End Function




 

'Ayush_03/06
Function ValidateMobileNumber() As Boolean
    Dim MobileNumber As String
    Dim i As Long
    ValidateMobileNumber = True
    MobileNumber = Sheet1.Range("sheet1.Mobileno").value
    For i = 1 To Len(MobileNumber)
        If Not IsNumeric(Mid(MobileNumber, i, 1)) Then
             fmsgbox ("* Primary Mobile No. must contain only digits from 0 to 9 in Sheet Income Details")
            ValidateMobileNumber = False
            Exit Function
        End If
     
    Next
End Function
'sheet1.MobileNoSec
Function ValidateMobileNumberSec() As Boolean
    Dim MobileNumber As String
    Dim i As Long
    ValidateMobileNumberSec = True
    MobileNumber = Sheet1.Range("sheet1.MobileNoSec").value
    For i = 1 To Len(MobileNumber)
        If Not IsNumeric(Mid(MobileNumber, i, 1)) Then
             fmsgbox ("* Secondary Mobile No. must contain only digits from 0 to 9 in Sheet Income Details")
            ValidateMobileNumberSec = False
            Exit Function
        End If
     
    Next
End Function
'Ayush_13/01/2026_V0.2
Function ValidateMobileNumber_Rep() As Boolean
    Dim MobileNumber_Rep As String
    Dim i As Long
    ValidateMobileNumber_Rep = True
    MobileNumber_Rep = Sheet1.Range("sheet1.MobileRepAssessee").value
    For i = 1 To Len(MobileNumber_Rep)
        If Not IsNumeric(Mid(MobileNumber_Rep, i, 1)) Then
             fmsgbox ("* Mobile No. in contact No. of representative assessee must contain only digits from 0 to 9 in Sheet Income Details")
            ValidateMobileNumber_Rep = False
            Exit Function
        End If
     
    Next
End Function
'--------------------
Function insertRowUnderSectionWithFormula_80CCC(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
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
   sPassword = mdCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=1).Rows(1).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(mdCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
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
                   
            sTempCellValueInt = mdCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_80CCC = nRows

If ActiveSheet.name = "80G" Then
    AutoSelectionforStateCode
End If

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function

'Ayush_16/01/2026
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
Sub LockUnlockCoOwnersDetails(HpCount As Long)
On Error Resume Next
'Dim HpCount  As Long
'For HpCount = 1 To Sheet3.Range("HousePropertySectionCount").value

   If Sheet22.Range("HP.CoOwnedYN" & HpCount).value = "No" Then
            
                 Sheet22.Range("Co_OwnnerBlock" & HpCount).ClearContents
                 Sheet22.Range("Co_OwnnerBlock" & HpCount).Locked = True
                 Sheet22.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HD8D8D8)
                 Sheet22.Range("HP.SharePercent" & HpCount).value = "100"
                 Sheet22.Range("HP.SharePercent" & HpCount).Locked = True
                 Sheet22.Range("HP.SharePercent" & HpCount).Interior.ColorIndex = 2
                 
  ElseIf Sheet22.Range("HP.CoOwnedYN" & HpCount).value = "Yes" Then
                 Sheet22.Range("Co_OwnnerBlock" & HpCount).Locked = False
                 Sheet22.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HCCFFCC)
                 Sheet22.Range("HP.SharePercent" & HpCount).value = WorksheetFunction.Max(0, 100 - WorksheetFunction.Sum(Sheet22.Range("HP.Co.Share" & HpCount).value))
                 Sheet22.Range("HP.SharePercent" & HpCount).Locked = True
                 Sheet22.Range("HP.SharePercent" & HpCount).Interior.ColorIndex = 2
   ElseIf Sheet22.Range("HP.CoOwnedYN" & HpCount).value <> "No" Then
            
                Sheet22.Range("Co_OwnnerBlock" & HpCount).ClearContents
                Sheet22.Range("Co_OwnnerBlock" & HpCount).Locked = False
                Sheet22.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.SharePercent" & HpCount).value = ""
                Sheet22.Range("HP.SharePercent" & HpCount).Locked = False
                Sheet22.Range("HP.SharePercent" & HpCount).Interior.Color = (&HCCFFCC)
    End If
'Next
End Sub
'Ayush_27/01
Function CheckAtoZ(chr1) As Boolean
    CheckAtoZ = True
    If ((asc(chr1) < 65) Or (asc(chr1) > 90)) Then
        CheckAtoZ = False
    End If
End Function
'Ayush_17/04/2026
Function checkfieldspecialcharacter_pran(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_pran = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "|", "`", "_", ",", ".") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_pran = False
            Exit Function
        End If
        Next
    Next
End Function
