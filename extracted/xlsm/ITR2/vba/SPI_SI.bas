Attribute VB_Name = "SPI_SI"
Option Explicit

Public personName_SPI, IncmNature_SPI, amt_SPI, Pan_SPI, Aadhaar_SPI As Variant
Dim endname_SPI, endname_S, endAmt_SPI, endPan_SPI, endAadhaar_SPI As Long 'Changed by Shrutika(SIT-92753)
Dim endRelation_SPI, endIncmNature_SPI As Long

Dim endname_IF, endPan_IF, endisLiable_IF As Long
Dim endProfitShare_IF, endProfitAmt_IF, endCap_IF As Long

Public firmName_IF, firmPan_IF, firmLiability_IF, firm92E_IF As Variant
Public firmProfitShare_IF, firmProfitAmt_IF, firmCapital_IF As Variant
Public relation_SPI, personName As Variant
Public end_IF, rngname_IF As Variant
Public end_SPI, rngname_SPI

Dim MsgIFSheet, MsgSPISheet As String

Sub Cmd_AddRows_IF_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("IF.FirmName||IF.FirmPAN||IF.IsLiableToAudit||IF.Is92EApplicable||IF.ProfitSharePercent||IF.ProfitShareAmt||IF.FirmCapBalOn31Mar||isLiableForAudit||LiableFor92E")
ActiveCellRange = EfilingCommon.searchLastRow("IF.FirmName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_SPI_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("SPI.SpecifiedPersonName||SPI.PANofSpecPerson||SPI.AadhaarofSpecPerson||SPI.ReltnShip||SPI.NatureOfInc||SPI.AmtIncluded")
ActiveCellRange = EfilingCommon.searchLastRow("SPI.SpecifiedPersonName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_Validate_SPI_Click()
SPI_SI.ValidateSPI_IF
fmsgboxStatus "Sheet SPI-SI is ok"
End Sub

Sub ValidateSPI_IF()
subProcCaption = "Validating SPI"
If Not ValidateSheetSPI Then
     Sheet21.Activate
     fmsgboxsmall MsgSPISheet ', vbOKOnly, "Error(s)!"
     CloseMsg
End If

ValidateSI
'subProcCaption = "Validating IF"
'If Not ValidateSheetIF Then
'     Sheet21.Activate
'     fmsgboxsmall MsgIFSheet ', vbOKOnly, "Error(s)!"
'     CloseMsg
'End If

End Sub

Function ValidateSheetSPI() As Boolean
'On Error Resume Next
MsgSPISheet = ""

ValidateSheetSPI = True

setTblinfoName_SPI
setTblinfoRelation_SPI
setTblinfoIncomrNature_SPI
setTblinfoAmount_SPI
setTblinfoPAN_SPI
setTblinfoAadhaar_SPI


'Shrutika(SIT-92753)
endname_SPI = Application.WorksheetFunction.Max(endname_SPI, endRelation_SPI, endIncmNature_SPI, endAmt_SPI, endPan_SPI, endAadhaar_SPI)
If Not ValidateSpecifiedPersonName_SPI Then ValidateSheetSPI = False
    If Len(Sheet21.Range("SPI.SpecifiedPersonName").item(1).value) > 0 Then
            If Not ValidateRetnship_SPI Then ValidateSheetSPI = False
            If Not ValidateNatureOfIncm_SPI Then ValidateSheetSPI = False
            If Not ValidateAmt_SPI Then ValidateSheetSPI = False
            If Not ValidatePAN_SPI Then ValidateSheetSPI = False
    End If
        
'Commented by Riyaz on 08/05/2025
'If ((endname_SPI <> endRelation_SPI) Or (endname_SPI <> endIncmNature_SPI) Or (endname_SPI <> endAmt_SPI)) Then
'            MsgSPISheet = MsgSPISheet + "* * Enter All Mandatory Fields in Schedule SPI."
'            ValidateSheetSPI = False
    
If Len(Sheet21.Range("SPI.Total").value) > 14 Then
            MsgSPISheet = MsgSPISheet + "* * Total Amount cannot be greater than 14 digits in Schedule SPI ."
            ValidateSheetSPI = False
 End If
    
    

End Function

Function ValidateSpecifiedPersonName_SPI() As Boolean

ValidateSpecifiedPersonName_SPI = True
'    setTblinfoName_SPI
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("SPI.SpecifiedPersonName").Cells
    Dim i As Long
    
    ReDim personName_SPI(endname_SPI)
    noOfProcessSub = endname_SPI
    For i = 1 To endname_SPI
        personName_SPI(i) = rangecells.item(i).value
        
        If Not chkCompulsory(personName_SPI(i)) Then
            'MsgSPISheet = MsgSPISheet + "* * Please enter the Name of Person at Sr. No " & i & " in Schedule SPI" & Chr(13)
            MsgSPISheet = MsgSPISheet + "* ""Please enter Name of the person in Schedule SPI in Sr. No " & i & ".""" & Chr(13) 'Added by Shrutika(SIT-93915-04/09/2025)
            ValidateSpecifiedPersonName_SPI = False
            Exit Function
        End If
        UpdateProgressBar
    Next

End Function

Function ValidateRetnship_SPI() As Boolean

ValidateRetnship_SPI = True
'    setTblinfoName_SPI
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("SPI.ReltnShip").Cells
    Dim i As Long
    ReDim relation_SPI(endname_SPI)
    For i = 1 To endname_SPI
        relation_SPI(i) = rangecells.item(i).value
        
        If Not chkCompulsory(relation_SPI(i)) Then
            MsgSPISheet = MsgSPISheet + "*  Please enter Relationship at Sr. No " & i & " in Schedule SPI" & Chr(13)
            ValidateRetnship_SPI = False
            Exit Function
        End If
    Next

End Function

Function ValidateNatureOfIncm_SPI() As Boolean

ValidateNatureOfIncm_SPI = True
'    setTblinfoName_SPI
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("SPI.NatureOfInc").Cells
    Dim i As Long
    ReDim IncmNature_SPI(endname_SPI)
    For i = 1 To endname_SPI
        IncmNature_SPI(i) = rangecells.item(i).value
        
        If isdropdownblank(IncmNature_SPI(i)) Then
            MsgSPISheet = MsgSPISheet + "*  Please select Head of Income in which included in Sl.No " & i & " in Schedule SPI " & Chr(13)
            ValidateNatureOfIncm_SPI = False
            Exit Function
        End If
    Next

End Function

Function ValidatePAN_SPI() As Boolean

ValidatePAN_SPI = True
Dim rangecells, rangecells1 As Range
Dim i, PanCount_SPI, AadhaarCount_SPI As Long
Set rangecells = Sheet21.Range("SPI.PANofSpecPerson").Cells
PanCount_SPI = Sheet21.Range("SPI.PANofSpecPerson").count

Set rangecells1 = Sheet21.Range("SPI.AadhaarofSpecPerson").Cells
AadhaarCount_SPI = Sheet21.Range("SPI.AadhaarofSpecPerson").count



ReDim Pan_SPI(PanCount_SPI)
For i = 1 To PanCount_SPI
If Not rangecells.item(i).value = "" Then
    Pan_SPI(i) = rangecells.item(i).value

        If Not checkfieldspecialcharacter(Pan_SPI(i)) Then
             MsgSPISheet = MsgSPISheet + "* * Please enter PAN No at Sr. No  " & i & "Cannot Contain Special Characters in Schedule SPI" & Chr(13)
             ValidatePAN_SPI = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(Pan_SPI(i))) Then
            MsgSPISheet = MsgSPISheet + "* * PAN at Sr. No  " & i & "is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.in Schedule SPI " & Chr(13)
            ValidatePAN_SPI = False
            Exit Function
        End If
End If
Next

ReDim Aadhaar_SPI(AadhaarCount_SPI)
For i = 1 To AadhaarCount_SPI
If Not rangecells1.item(i).value = "" Then
    Aadhaar_SPI(i) = rangecells1.item(i).value

         If Aadhaar_SPI(i) <> "" Then
        If Not IsNumeric(Aadhaar_SPI(i)) Then
             MsgSPISheet = MsgSPISheet + Chr(13) & "Please enter valid Aadhaar Number in schedule SPI" & Chr(13)
                ValidatePAN_SPI = False
                Exit Function
            End If
       
    
        If Aadhaar_SPI(i) = "000000000000" Then
           MsgSPISheet = MsgSPISheet + Chr(13) & "Please enter valid Aadhaar Number in schedule SPI" & Chr(13)
                ValidatePAN_SPI = False
                Exit Function
        End If
        
        If Aadhaar_SPI(i) = "111111111111" Then
            MsgSPISheet = MsgSPISheet + Chr(13) & "Please enter valid Aadhaar Number in schedule SPI" & Chr(13)
                ValidatePAN_SPI = False
                Exit Function
        End If
        
    
        If Len(Aadhaar_SPI(i)) <> 12 Then
            MsgSPISheet = MsgSPISheet + Chr(13) & "Please enter valid 12 digit Aadhaar Number in schedule SPI" & Chr(13)
                ValidatePAN_SPI = False
                Exit Function
        End If
         End If
End If
Next



End Function

Function ValidateAmt_SPI() As Boolean

ValidateAmt_SPI = True
'    setTblinfoName_SPI
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("SPI.AmtIncluded").Cells
    Dim i As Long
    ReDim amt_SPI(endname_SPI)
    For i = 1 To endname_SPI
        amt_SPI(i) = rangecells.item(i).value
        
        If Not chkCompulsory(amt_SPI(i)) Then
            'MsgSPISheet = MsgSPISheet + "*  Please enter Amount at Sr. No " & i & " in Schedule SPI" & Chr(13)
           MsgSPISheet = MsgSPISheet + "* ""Please enter Amount of Income in Schedule SPI in Sl.no.1.""" & Chr(13) 'Updated by Shrutika (11/9/2025)
            ValidateAmt_SPI = False
            Exit Function
        End If
    Next

End Function

Function ValidateSheetIF() As Boolean
On Error Resume Next

ValidateSheetIF = True
MsgIFSheet = ""

setTblinfoName_IF
setTblinfoPan_IF
setTblinfoIsLiable_IF
setTblinfoProfitAmt_IF
setTblinfoProfitShare_IF
setTblinfoCap_IF

If Mid(Sheet2.Range("Sheet2.Code1").value, 1, 5) = "00001" Or Mid(Sheet2.Range("Sheet2.Code2").value, 1, 5) = "00001" Or Mid(Sheet2.Range("Sheet.Code3").value, 1, 5) = "00001" Then
If Not endname_IF > 0 Then
    MsgIFSheet = MsgIFSheet + "* Atleast one row is Mandatory in Schedule IF."
    ValidateSheetIF = False
    Exit Function
End If
End If

If Not ValidateFirmName_IF Then ValidateSheetIF = False
 If Len(Sheet21.Range("IF.FirmName").item(1).value) > 0 Then
            If Not ValidateFirmPan_IF Then ValidateSheetIF = False
            If Not ValidateFirmLiability_IF Then ValidateSheetIF = False
            If Not ValidateFirm92E_IF Then ValidateSheetIF = False
            If Not ValidateFirmProfitShare_IF Then ValidateSheetIF = False
            If Not ValidateFirmProfitAmt_IF Then ValidateSheetIF = False
            If Not ValidateFirmCapital_IF Then ValidateSheetIF = False
            If Not validateTotalIF Then ValidateSheetIF = False
 End If

        
If ((endname_IF <> endPan_IF) Or (endname_IF <> endisLiable_IF) Or (endname_IF <> endProfitShare_IF) Or _
            (endname_IF <> endProfitAmt_IF) Or (endname_IF <> endCap_IF)) Then
            MsgIFSheet = MsgIFSheet + "* Enter All Mandatory Fields in Schedule IF."
            ValidateSheetIF = False
            Exit Function
 End If
        
End Function
Function ValidateFirmName_IF() As Boolean

ValidateFirmName_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.FirmName").Cells
    Dim i As Long
    ReDim firmName_IF(endname_IF)
    noOfProcessSub = endname_IF
    For i = 1 To endname_IF
        firmName_IF(i) = rangecells.item(i).value
        If Not checkfieldspecialcharacter(firmName_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Please enter the Name of the Firm at Sr. No " & i & " characters < > & ' " & Chr(34) & " are not allowed in Schedule IF " & Chr(13)
            ValidateFirmName_IF = False
            Exit Function
        End If
        
        If Not chkCompulsory(firmName_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Please enter the Name of the Firm at Sr. No " & i & " in Schedule IF " & Chr(13)
            ValidateFirmName_IF = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateFirmPan_IF() As Boolean

ValidateFirmPan_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.FirmPAN").Cells
    Dim i As Long
    ReDim firmPan_IF(endname_IF)
    For i = 1 To endname_IF
        firmPan_IF(i) = rangecells.item(i).value
        
        If Not chkCompulsory(firmPan_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Please enter PAN of the Firm at Sr. No " & i & " in Schedule IF" & Chr(13)
            ValidateFirmPan_IF = False
            Exit Function
        End If
        
        If Not CheckPANIF(firmPan_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* PAN of the Firm at Sr. No " & i & " in Schedule IF is Invalid, PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet, fourth character must be 'F'" & Chr(13)
            ValidateFirmPan_IF = False
            Exit Function
        End If
    Next
End Function
Function CheckPANIF(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 3 is for individuals .So,4th character of PAN should be "P" or "H"
'status = Sheet1.Range("sheet1.Status").value
PAN = UCase(PAN)
    CheckPANIF = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPANIF = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPANIF = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPANIF = False
            Exit Function
        End If
        
      If Not Mid(PAN, 4, 1) = "F" Then
            CheckPANIF = False
            Exit Function
        End If
            
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPANIF = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPANIF = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPANIF = False
            Exit Function
        End If
    End If
End Function
Function ValidateFirmLiability_IF() As Boolean

ValidateFirmLiability_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.IsLiableToAudit").Cells
    Dim i As Long
    ReDim firmLiability_IF(endname_IF)
    For i = 1 To endname_IF
        firmLiability_IF(i) = rangecells.item(i).value
        
        If isdropdownblank(firmLiability_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Whether the firm is liable for audit at Sr. No " & i & "is mandatory in Schedule IF " & Chr(13)
            ValidateFirmLiability_IF = False
            Exit Function
        End If
    Next
End Function

Function ValidateFirm92E_IF() As Boolean

ValidateFirm92E_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.Is92EApplicable").Cells
    Dim i As Long
    ReDim firm92E_IF(endname_IF)
    For i = 1 To endname_IF
        firm92E_IF(i) = rangecells.item(i).value
        
        If isdropdownblank(firm92E_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Whether the section 92E is applicable to firm at Sr. No " & i & " is manadtory in Schedule IF." & Chr(13)
            ValidateFirm92E_IF = False
            Exit Function
        End If
    Next
End Function

Function ValidateFirmProfitShare_IF() As Boolean

ValidateFirmProfitShare_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.ProfitSharePercent").Cells
    Dim i As Long
    ReDim firmProfitShare_IF(endname_IF)
    For i = 1 To endname_IF
        firmProfitShare_IF(i) = rangecells.item(i).text
        
        If Not chkCompulsory(firmProfitShare_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Percentage share in profit of the firm at Sr. No " & i & " is mandatory in Schedule IF" & Chr(13)
            ValidateFirmProfitShare_IF = False
            Exit Function
        End If
        
        If firmProfitShare_IF(i) = "0" Or firmProfitShare_IF(i) = "0.00" Then
            MsgIFSheet = MsgIFSheet + "* Percentage share in profit of the firm cannot be zero at Sr. No " & i & " in Schedule IF" & Chr(13)
            ValidateFirmProfitShare_IF = False
            Exit Function
        End If
        
    Next
End Function

Function ValidateFirmProfitAmt_IF() As Boolean

ValidateFirmProfitAmt_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.ProfitShareAmt").Cells
    Dim i As Long
    ReDim firmProfitAmt_IF(endname_IF)
    For i = 1 To endname_IF
        firmProfitAmt_IF(i) = rangecells.item(i).value
        
        If Not chkCompulsory(firmProfitAmt_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Amount of share in profit at Sr. No " & i & " is mandatory in Schedule IF" & Chr(13)
            ValidateFirmProfitAmt_IF = False
            Exit Function
        End If
    Next
End Function


Function ValidateFirmCapital_IF() As Boolean

ValidateFirmCapital_IF = True
    setTblinfoName_IF
    Dim rangecells As Range
    Set rangecells = Sheet21.Range("IF.FirmCapBalOn31Mar").Cells
    Dim i As Long
    ReDim firmCapital_IF(endname_IF)
    For i = 1 To endname_IF
        firmCapital_IF(i) = rangecells.item(i).value
        
        If Not chkCompulsory(firmCapital_IF(i)) Then
            MsgIFSheet = MsgIFSheet + "* Capital Balance as on 31st March in the firm at Sr. No " & i & " is mandatory in Schedule IF" & Chr(13)
            ValidateFirmCapital_IF = False
            Exit Function
        End If
    Next
End Function

Function validateTotalIF() As Boolean
validateTotalIF = True
Dim total1, total2 As Variant
    total1 = Sheet21.Range("IF.TotalProfitShareAmt").value
    total2 = Sheet21.Range("IF.TotalFirmCapBalOn31Mar").value
    If Len(total1) > 14 Then
    MsgIFSheet = MsgIFSheet + "* Total of Amount of Share of Profit cannot be greater than 14 digits. in Schedule IF" + Chr(13)
    validateTotalIF = False
    End If
    If Len(total2) > 14 Then
    MsgIFSheet = MsgIFSheet + "* Total of Capital balance on 31 March in the firm cannot be greater than 14 digits. in Schedule IF" + Chr(13)
    validateTotalIF = False
    End If
End Function

'==========================
'For IF

Sub setTblinfoName_IF()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("IF.FirmName").count
 Set rangecells = Sheet21.Range("IF.FirmName").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endname_IF = ccount
End Sub

Sub setTblinfoPan_IF()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("IF.FirmPAN").count
 Set rangecells = Sheet21.Range("IF.FirmPAN").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endPan_IF = ccount
End Sub

Sub setTblinfoIsLiable_IF()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("IF.IsLiableToAudit").count
 Set rangecells = Sheet21.Range("IF.IsLiableToAudit").Cells
 For mIntCtr = 1 To mIntCells
     If Not ((rangecells.item(mIntCtr).value = "(Select)") Or (rangecells.item(mIntCtr).value = "")) Then
         ccount = ccount + 1
     End If
 Next
 endisLiable_IF = ccount
End Sub

Sub setTblinfoProfitShare_IF()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("IF.ProfitSharePercent").count
 Set rangecells = Sheet21.Range("IF.ProfitSharePercent").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endProfitShare_IF = ccount
End Sub

Sub setTblinfoProfitAmt_IF()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("IF.ProfitShareAmt").count
 Set rangecells = Sheet21.Range("IF.ProfitShareAmt").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endProfitAmt_IF = ccount
End Sub

Sub setTblinfoCap_IF()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("IF.FirmCapBalOn31Mar").count
 Set rangecells = Sheet21.Range("IF.FirmCapBalOn31Mar").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endCap_IF = ccount
End Sub

'For SPI

Sub setTblinfoName_SPI()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("SPI.SpecifiedPersonName").count
 Set rangecells = Sheet21.Range("SPI.SpecifiedPersonName").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endname_SPI = ccount
 'Malli------------
 'AY_2024_25 OLD
'rngname_SPI = "SPI.SpecifiedPersonName;SPI.PANofSpecPerson;SPI.ReltnShip;SPI.NatureOfInc;SPI.AmtIncluded;"
'Malli_AY_2025_26
rngname_SPI = "SPI.SpecifiedPersonName;SPI.PANofSpecPerson;SPI.AadhaarofSpecPerson;SPI.ReltnShip;SPI.NatureOfInc;SPI.AmtIncluded;"
'--------------------------
End Sub


Sub setTblinfoRelation_SPI()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("SPI.ReltnShip").count
 Set rangecells = Sheet21.Range("SPI.ReltnShip").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endRelation_SPI = ccount
End Sub

Sub setTblinfoIncomrNature_SPI()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("SPI.NatureOfInc").count
 Set rangecells = Sheet21.Range("SPI.NatureOfInc").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
         ccount = ccount + 1
     End If
 Next
 endIncmNature_SPI = ccount
End Sub

Sub setTblinfoAmount_SPI()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("SPI.AmtIncluded").count
 Set rangecells = Sheet21.Range("SPI.AmtIncluded").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endAmt_SPI = ccount
End Sub

Sub setTblinfoPAN_SPI()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("SPI.PANofSpecPerson").count
 Set rangecells = Sheet21.Range("SPI.PANofSpecPerson").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endPan_SPI = ccount
End Sub




Sub setTblinfoAadhaar_SPI()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet21.Range("SPI.AadhaarofSpecPerson").count
 Set rangecells = Sheet21.Range("SPI.AadhaarofSpecPerson").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 endAadhaar_SPI = ccount
End Sub



