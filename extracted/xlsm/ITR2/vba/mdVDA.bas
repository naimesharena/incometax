Attribute VB_Name = "mdVDA"
Option Explicit
Public msgValidateSheetVDA As String

'Added by Malli
'Name_vdigitalasset  end_Name_vdigitalasset
Public Name_vdigitalasset As Variant
Public end_Name_vdigitalasset As Variant
'----------




'Change.21.03.2023.102.IDS.68
'Sub CmdScheduleVDAAdd_Click()
'    Sheets("VDA").Activate
'    EfilingCommon.DefinedgridNameRange = "VDA.DateOfAcq||VDA.DateOfTrans||VDA.HeadIncome||VDA.CostOfAcq||VDA.ConsiderationReceived||VDA.TotalIncome"
'    ActiveCellRange = EfilingCommon.searchLastRow("VDA.DateOfAcq")
'    EfilingCommon.insertRowUnderSection
'End Sub
Sub CmdScheduleVDAAdd_Click()
Dim vRows As Long
    Sheets("VDA").Activate
    'VDA_E1_AY_2023_24 old
    'EfilingCommon.DefinedgridNameRange = "VDA.DateOfAcq||VDA.DateOfTrans||VDA.HeadIncome||VDA.CostOfAcq||VDA.ConsiderationReceived||VDA.TotalIncome"
    'VDA_C1_AY_2023_24  Change
    EfilingCommon.DefinedgridNameRange = "VDA.DateOfAcq||VDA.DateOfTrans||VDA.HeadIncome||VDA.CostOfAcq||VDA.ConsiderationReceived||VDA.TotalIncome"
    'EfilingCommon.DefinedgridNameRange = "VDA.DateOfAcq||VDA.Name_vdigitalasset||VDA.DateOfTrans||VDA.HeadIncome||VDA.CostOfAcq||VDA.ConsiderationReceived||VDA.TotalIncome"

    ActiveCellRange = EfilingCommon.searchLastRow("VDA.DateOfAcq")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'End Change.IDS.68

Sub CmdScheduleVDAAdd_Import(DiffRows As Long)
    Dim rngName_VDA
    Sheets("VDA").Activate
    rngName_VDA = "VDA.DateOfAcq||VDA.DateOfTrans||VDA.HeadIncome||VDA.CostOfAcq||VDA.ConsiderationReceived||VDA.TotalIncome"
    SelectLastRow ("VDA.DateOfAcq")
    InsertDiffRowsAndFillFormulas (DiffRows)
    'Malli----------------
    'AY_2023_24 OLd
      'Call ExendRangeNameToTable(Diffrows, rngName_VDA)
      'AY_2024_25 change
      Call ExendRangeNameToTable_M_VDA(DiffRows, rngName_VDA)
      '-----------------------
End Sub

Sub Cmd_Validate_VDA_Click()
ValidateScheduleVDA
fmsgboxStatus "Sheet VDA  is ok" ', vbOKOnly, "ITR 3"
End Sub

Sub ValidateScheduleVDA()
    subProcCaption = "Validating VDA"
    If Not ValidateScheduleVDA1 Then
        Sheet47.Activate
        fmsgboxsmall msgValidateSheetVDA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub


Function ValidateScheduleVDA1() As Boolean
ValidateScheduleVDA1 = True

If Not ValidateMandatoryVDA() Then ValidateScheduleVDA1 = False

End Function
Function msgbox_VDA(strmsg As String) As String
     msgValidateSheetVDA = msgValidateSheetVDA & strmsg & Chr(13)
End Function


Function ValidateMandatoryVDA() As Boolean
ValidateMandatoryVDA = True
Dim i As Long

For i = 2 To Sheet47.Range("VDA.DateOfAcq").count
    If (Sheet47.Range("VDA.DateOfAcq").item(i).value <> "") Then
        If Sheet47.Range("VDA.DateOfAcq").item(i - 1).value = "" Then
            msgbox_VDA ("* Please fill all the mandatory fields of the row " & i - 1 & " before filling the next row in Schedule VDA" & Chr(13))
            ValidateMandatoryVDA = False
            Exit Function
        End If
    End If
Next



For i = 1 To Sheet47.Range("VDA.DateOfAcq").count

'    If Sheet47.Range("VDA.DateOfAcq").Cells(i, 1).value <> "" Or Sheet47.Range("VDA.DateOfTrans").Cells(i, 1).value <> "" Or _
'        (Sheet47.Range("VDA.HeadIncome").Cells(i, 1).value <> "" And Sheet47.Range("VDA.HeadIncome").Cells(i, 1).value <> "(Select)") Or _
'        Sheet47.Range("VDA.CostOfAcq").Cells(i, 1).value <> "" Or Sheet47.Range("VDA.ConsiderationReceived").Cells(i, 1).value <> "" Or Sheet47.Range("VDA.Name_vdigitalasset").Cells(i, 1).value <> "" Then
'
'
   If Sheet47.Range("VDA.DateOfAcq").Cells(i, 1).value <> "" Or Sheet47.Range("VDA.DateOfTrans").Cells(i, 1).value <> "" Or _
        (Sheet47.Range("VDA.HeadIncome").Cells(i, 1).value <> "" And Sheet47.Range("VDA.HeadIncome").Cells(i, 1).value <> "(Select)") Or _
        Sheet47.Range("VDA.CostOfAcq").Cells(i, 1).value <> "" Or Sheet47.Range("VDA.ConsiderationReceived").Cells(i, 1).value <> "" Then
             
        
        
        'Some thing is filled in the row
        
        If Sheet47.Range("VDA.DateOfAcq").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please select the date of Acquisition at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA.DateOfTrans").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please select the date of Transfer at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA.HeadIncome").Cells(i, 1).value = "" Or Sheet47.Range("VDA.HeadIncome").Cells(i, 1).value = "(Select)" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please select head of income at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA.CostOfAcq").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please enter Cost of Acquisition at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA.ConsiderationReceived").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please enter Consideration Received at Sr. No. " & i)
        End If
        
        'VDA_C1_AY_2023_24  ADDEd 'Malli

'        If Sheet47.Range("VDA.Name_vdigitalasset").Cells(i, 1).value = "" Then
'            ValidateMandatoryVDA = False
'            msgbox_VDA ("Please enter Name of the virtual digital asset at Sr. No. " & i)
'        End If
        
    End If
Next i


' If Not ValidateName_vdigitalasset() Then ValidateMandatoryVDA = False



End Function

''''''''''''''''''Malli added
'VDA_C1_AY_2023_24  ADDEd 'Malli
Sub setTblinfo_Name_vdigitalasset()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet47.Range("VDA.Name_vdigitalasset").count
    Set rangecells = Sheet47.Range("VDA.Name_vdigitalasset").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Name_vdigitalasset = ccount
     
End Sub

'-------------------------------
'VDA_C1_AY_2023_24  ADDEd 'Malli
Function ValidateName_vdigitalasset() As Boolean
    ValidateName_vdigitalasset = True
    setTblinfo_Name_vdigitalasset
    Dim rangecells As Range
    Dim cellrange As Variant
    
    Dim i As Long
   
    Set rangecells = Sheet47.Range("VDA.Name_vdigitalasset").Cells
    ReDim Name_vdigitalasset(end_Name_vdigitalasset)
    
    For i = 1 To end_Name_vdigitalasset
        cellrange = GetMergedAddressCell(rangecells, i)
        Name_vdigitalasset(i) = UCase(Sheet47.Range(cellrange).value)
        
        If Not chkCompulsory(Name_vdigitalasset(i)) Then
            msgbox_VDA ("Please enter Name of the virtual digital asset at Sr. No. " & i)
            ValidateName_vdigitalasset = False
            Exit Function
        End If
        
        If Len(Name_vdigitalasset(i)) > 125 Then
            msgbox_VDA ("Name of the virtual digital asset should not excced 125 characters at Sr. No. " & i)
            ValidateName_vdigitalasset = False
            Exit Function
        End If
    
    If Not checkfieldspecialcharacter1(Name_vdigitalasset(i)) Then
                msgbox_VDA ("Special characters are not allowed in Name of the virtual digital asset  at Sr. No. " & i)
                ValidateName_vdigitalasset = False
            Exit Function
        End If
    
    UpdateProgressBar
    Next
End Function
