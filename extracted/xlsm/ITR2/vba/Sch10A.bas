Attribute VB_Name = "Sch10A"
Option Explicit

Public DedFromUndertaking_SEZA10 As Variant
Public AssmtYrUnit_SEZ10 As Variant

Public TotalDedUs10Sub_SEZA10 As Variant
Public TotalDedUs10A_SEZA10 As Variant
Public DedFromUndertaking_AA10 As Variant
Public AssmtYrUnit_AA10 As Variant
Public TotalDedUs10Sub_AA10 As Variant

Public msgValidateSheetSEZA10 As Variant
Public msgValidateSheetAA10 As Variant

Public end_SEZA10 As Variant
Public rngname_SEZA10 As Variant

Public end_AA10 As Variant
Public rngname_AA10 As Variant

Public end_SEZA10_2 As Variant
Dim rngname_SEZA10_2 As Variant

Public end_AA10_2 As Variant
Dim rngname_AA10_2 As Variant


Sub AddRows_AA10()
    Dim numberofrows As Long
    setTblinfo_AA10
    SelectLastRow ("AA10.DedFromUndertaking")
    numberofrows = InsertRowsAndFillFormulas(1)
    Call ExendRangeNameToTable(numberofrows, rngname_AA10)
    
      

'     SelectLastRow ("DI_10AA_Date")
'     numberofrows = InsertRowsAndFillFormulas(1)
'    Call ExendRangeNameToTable(numberofrows, "DI_10AA_Amount;DI_10AA_Date;DI_10AA_YearOfClaim;")
    Sheet18.Activate
End Sub

Sub setTblinfo_AA10()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet18.Range("AA10.DedFromUndertaking").count
    Set rangecells = Sheet18.Range("AA10.DedFromUndertaking").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
            Else
            Exit For
        End If
    Next
    end_AA10 = ccount
    rngname_AA10 = "AA10.AssmtYrUnit;AA10.DedFromUndertaking;AA10_undertaking;"
End Sub

Sub ValidateSheet10AClick()
ValidateSheet10A
fmsgboxStatus "Sheet 10AA is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheet10A()

subProcCaption = "Validating 10AA"
    If Not ValidatesheetAA10 Then
        Sheet18.Activate
        fmsgboxsmall msgValidateSheetAA10 ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub

Function ValidatesheetAA10() As Boolean
    ValidatesheetAA10 = True
    If Not ValidateDedFromUndertaking_AA10() Then ValidatesheetAA10 = False
    If Not ValidateTotalDedUs10Sub_AA10() Then ValidatesheetAA10 = False
    
    setTblinfo_AA10_2
    
    
End Function

Sub setTblinfo_AA10_2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim Year As Variant
    
    ccount = 0
    mIntCells = Sheet18.Range("AA10.AssmtYrUnit").count
    Set rangecells = Sheet18.Range("AA10.AssmtYrUnit").Cells
    For mIntCtr = 1 To mIntCells
        Year = rangecells.item(mIntCtr).value
        If isdropdownblank(Year) Then
            Year = ""
        End If
        If Not Year = "" Then
            ccount = ccount + 1
            Else
            Exit For
        End If
    Next
    
    end_AA10_2 = ccount
    rngname_AA10_2 = "AA10.AssmtYrUnit;AA10.DedFromUndertaking;"
End Sub

Function ValidateTotalDedUs10Sub_SEZA10() As Boolean
 ValidateTotalDedUs10Sub_SEZA10 = True
 TotalDedUs10Sub_SEZA10 = Sheet18.Range("SEZA10.TotalDedUs10Sub").value
 If Len(TotalDedUs10Sub_SEZA10) > 14 Then
 msgValidateSheetSEZA10 = "* Schedule 10A.Total of deductions cannot be greater than 14 digits"
 ValidateTotalDedUs10Sub_SEZA10 = False
 End If
End Function

Function ValidateTotalDedUs10A_SEZA10() As Boolean
    ValidateTotalDedUs10A_SEZA10 = True
    'TotalDedUs10A_SEZA10 = Sheet18.Range("SEZA10.TotalDedUs10A").value
End Function

Function ValidateDedFromUndertaking_AA10() As Boolean
    Dim i As Long
    Dim rangecells As Range
    Dim rangecells2 As Range
    
    ValidateDedFromUndertaking_AA10 = True
    setTblinfo_AA10_2
    Set rangecells = Sheet18.Range("AA10.DedFromUndertaking").Cells
    Set rangecells2 = Sheet18.Range("AA10.AssmtYrUnit").Cells
    ReDim DedFromUndertaking_AA10(rangecells2.Cells.count)
    ReDim AssmtYrUnit_AA10(rangecells.Cells.count)
    noOfProcessSub = rangecells.Cells.count
    For i = 1 To rangecells.Cells.count
    
        DedFromUndertaking_AA10(i) = rangecells.item(i).value
        AssmtYrUnit_AA10(i) = rangecells2.item(i).value
        
        If isdropdownblank(AssmtYrUnit_AA10(i)) Then
            AssmtYrUnit_AA10(i) = ""
        End If
        
        If (DedFromUndertaking_AA10(i) < 0) Then
            msgValidateSheetAA10 = "* Amount of deduction at Sr. No  " & i & "   should be Non negative, no decimal, upto 99,999,999,999,999 in Sheet 10AA"
            ValidateDedFromUndertaking_AA10 = False
            Exit Function
        End If
   
         If Not isdropdownblank(AssmtYrUnit_AA10(i)) Or DedFromUndertaking_AA10(i) <> "" Then
            If (DedFromUndertaking_AA10(i) = "") Then
                msgValidateSheetAA10 = "* Please enter Amount of deduction at Sr. No  " & i & "  in Sheet 10AA " & Chr(13)
                ValidateDedFromUndertaking_AA10 = False
                Exit Function
            End If
            If isdropdownblank(AssmtYrUnit_AA10(i)) Then
                msgValidateSheetAA10 = "* Please select an Assessment Year from the Drop down at Sr. No  " & i & "  in Sheet 10AA " & Chr(13)
                ValidateDedFromUndertaking_AA10 = False
            Exit Function
            End If

         End If
         UpdateProgressBar
    Next
End Function

Function ValidateTotalDedUs10Sub_AA10() As Boolean
    ValidateTotalDedUs10Sub_AA10 = True
    TotalDedUs10Sub_AA10 = Sheet18.Range("AA10.TotalDedUs10Sub").value
    
    If Len(TotalDedUs10Sub_AA10) > 14 Then
        msgValidateSheetAA10 = "* Total Ded U/s 10AA in Sch 10-AA should not be greater than 14 digits "
        ValidateTotalDedUs10Sub_AA10 = False
        Exit Function
    End If
End Function

Function DefaultDedFromUndertaking_SEZA10() As String
    DefaultDedFromUndertaking_SEZA10 = "0"
End Function

Function DefaultTotalDedUs10Sub_SEZA10() As String
    DefaultTotalDedUs10Sub_SEZA10 = "0"
End Function

Function DefaultTotalDedUs10A_SEZA10() As String
    DefaultTotalDedUs10A_SEZA10 = "0"
End Function

Function DefaultDedFromUndertaking_AA10() As String
    DefaultDedFromUndertaking_AA10 = "0"
End Function

Function DefaultTotalDedUs10Sub_AA10() As String
    DefaultTotalDedUs10Sub_AA10 = "0"
End Function


