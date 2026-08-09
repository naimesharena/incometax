Attribute VB_Name = "mdEI"
Option Explicit

Public sceiInterestInc As Variant
Public sceiDividendInc As Variant
Public sceiLTCGWhereSTTPaid As Variant
Public sceiNetAgriIncOrOthrIncRule7 As Variant
Public sceiShareOfProfitFirmAOP As Variant
Public sceiOthers As Variant
Public sceiTotalExemptInc As Variant
Public sceiExpenditureOnAgriculture As Variant
Public sceiUnabsorbedAgriculturalloss As Variant
Public sceiAgriInc As Variant
Public sceiNetAgriculturalIncome As Variant
Public NI_EI, Description_EI, Amt_EI, Subcat_EI As Variant
Public NameOfDistrict_EI, Pincode_EI, AgriculturalLand_In_Acre_EI, AgriculturalLand_Owned_On_Held_onlease_EI, AgriculturalLand_irrigated_or_rainfed_EI As Variant
Public AmountOfIncome_EIDTAA, NatureOfIncome_EIDTAA, CountryNameA1_EIDTAA, CountryCodeA1_EIDTAA, Article_EIDTAA, HeadOfIncome_EIDTAA, TRC_EIDTAA As Variant

'Malli--------------
Public ColCount_Name_of_district As Variant
'-------------------
Dim ColCount2, ColCount2_1 As Long
Dim ColCount_EI1, ColCount_EI2, ColCount_EI3, ColCount_EI4, ColCount_EI5 As Long
Dim ColCount_EIDTAA_1, ColCount_EIDTAA_2, ColCount_EIDTAA_3, ColCount_EIDTAA_4, ColCount_EIDTAA_5, ColCount_EIDTAA_6 As Long
Dim msgvalidateEIError As Variant

Sub ValidateSheetEI()

If Not ValidateShtEI Then
Sheet24.Activate
fmsgboxsmall msgvalidateEIError ', vbOKOnly, "Error(s)!"
CloseMsg
End If
End Sub
Sub ValidateEIClick()
ValidateSheetEI
If msgvalidateEIError = "" Then fmsgboxStatus "Sheet EI is OK"
End Sub
Sub Cmd_AddRows_EI_Click()
Dim vRows As Long
'EfilingCommon.DefinedgridNameRange = ("EI.slno||EI.NatureOfIncome||EI.Description||EI.Amount")
'EfilingCommon.DefinedgridNameRange = ("EI.slno||Others.NOI||Nature_Others||EI.Amount")

EfilingCommon.DefinedgridNameRange = ("EI.slno||Others.NOI||Nature_Others||EI.Description||EI.Amount") 'Added by Naresh for V0.4 Changes
ActiveCellRange = EfilingCommon.searchLastRow("Others.NOI")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_EI_Agriculture_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("Name_of_district||SchEI_Pincode||SchEI_AgriculturalLand||SchEI_OwedOrLease||SchEI_IrrigatedOrRainFed")
ActiveCellRange = EfilingCommon.searchLastRow("Name_of_district")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_DTAA_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("SCHEI_DTAA_AmountOfIncome||SCHEI_DTAA_NatureOfIncome||SCHEI_DTAA_CountryCode||SCHEI_DTAA_Article||SCHEI_DTAA_HeadOfIncome||SCHEI_TRCObtained")
ActiveCellRange = EfilingCommon.searchLastRow("SCHEI_DTAA_AmountOfIncome")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

'Function ValidateNI_EI() As Boolean
'ValidateNI_EI = True
'setTableInfo_EI1
'Dim rangecells As Range
'Dim rangecells1 As Range
''Set rangecells = Sheet24.Range("EI.NatureOfIncome").Cells
''Set rangecells1 = Sheet24.Range("EI.Description").Cells
'Set rangecells = Sheet24.Range("Others.NOI").Cells
'Set rangecells1 = Sheet24.Range("Nature_Others").Cells
'Dim i As Long
'ReDim NI_EI(ColCount2)
'ReDim Description_EI(ColCount2)
'
'noOfProcessSub = ColCount2
'For i = 1 To ColCount2
'NI_EI(i) = rangecells.item(i).value
'Description_EI(i) = rangecells1.item(i).value
'
'If Len(NI_EI(i)) > 0 Then
'End If
'
'
'
'If Not chkCompulsory(NI_EI(i)) Then
'  '  msgvalidateEIError = msgvalidateEIError + "* Nature Of Income at Sr.NO " & i & " in Sheet EI, is Compulsory" & Chr(13)
'     msgvalidateEIError = msgvalidateEIError + "* ""Please select category from the drop down in Exempt Income""" & Chr(13)
'
'    ValidateNI_EI = False
'    Exit Function
'End If
'
''If NI_EI(i) = "Any other" Then
' If Not chkCompulsory(Description_EI(i)) Then
'
'   'JYOTI55
'   'msgvalidateEIError = msgvalidateEIError + "* Description at Sr.NO " & i & " in Sheet EI, is Compulsory" & Chr(13)
'
'  ' msgvalidateEIError = msgvalidateEIError + "* Description is mandatory in schedule EI at Sl. NO " & i & "" & Chr(13)
'   msgvalidateEIError = msgvalidateEIError + "* ""Please select sub-category from the drop down in Exempt Income""" & Chr(13)
'
'    ValidateNI_EI = False
'    Exit Function
' End If
''End If
'
'
''If Len(Description_EI(i)) > 125 Then 'SIT-97873 As per DEsheet we have to restrict if more than 50
''If Len(Description_EI(i)) > 50 Then
''
''   ' msgvalidateEIError = msgvalidateEIError + "* Description at Serial No " & i & " cannot be more than 125 Characters" & Chr(10)
''    msgvalidateEIError = msgvalidateEIError + "* Description at Serial No " & i & " cannot be more than 50 Characters" & Chr(10) 'SIT-97873 As per DEsheet we have to restrict if more than 50
''
''    ValidateNI_EI = False
''    Exit Function
''End If
'UpdateProgressBar
'Next
'End Function

Function ValidateNI_EI() As Boolean
ValidateNI_EI = True
setTableInfo_EI1
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range

'Set rangecells = Sheet24.Range("EI.NatureOfIncome").Cells

Set rangecells2 = Sheet24.Range("EI.Description").Cells
Set rangecells = Sheet24.Range("Others.NOI").Cells
Set rangecells1 = Sheet24.Range("Nature_Others").Cells
Dim i As Long
ReDim NI_EI(ColCount2)
ReDim Subcat_EI(ColCount2)
ReDim Description_EI(ColCount2)

noOfProcessSub = ColCount2
For i = 1 To ColCount2
NI_EI(i) = rangecells.item(i).value
Subcat_EI(i) = rangecells1.item(i).value
Description_EI(i) = rangecells2.item(i).value

If Len(NI_EI(i)) > 0 Then
End If



If Not chkCompulsory(NI_EI(i)) Then
     msgvalidateEIError = msgvalidateEIError + "* ""Please select category from the drop down in Exempt Income""" & Chr(13)

    ValidateNI_EI = False
    Exit Function
End If

' If Not chkCompulsory(Subcat_EI(i)) Then

  If isdropdownblank(Subcat_EI(i)) Then
  
   msgvalidateEIError = msgvalidateEIError + "* ""Please select sub-category from the drop down in Exempt Income""" & Chr(13)
    ValidateNI_EI = False
    Exit Function
  End If

 

If Subcat_EI(i) = "Income exempt as per CBDT Circular" Then
If Not chkCompulsory(Description_EI(i)) Then
   msgvalidateEIError = msgvalidateEIError + "* ""Please enter the relevant Circular No. for other exempt income for sub-category “Income exempt as per CBDT Circular” in Schedule EI""" & Chr(13)
   ValidateNI_EI = False
    Exit Function
 End If
End If

If Subcat_EI(i) = "Income exempt as per CBDT Notification" Then
If Not chkCompulsory(Description_EI(i)) Then
   msgvalidateEIError = msgvalidateEIError + "* ""Please enter the relevant Notification No. for other exempt income for sub-category “Income exempt as per CBDT Notification” in Schedule EI""" & Chr(13)
   ValidateNI_EI = False
    Exit Function
 End If
End If


If Subcat_EI(i) = "Receipts not in the nature of income" Then
If Not chkCompulsory(Description_EI(i)) Then
   msgvalidateEIError = msgvalidateEIError + "* ""Please enter the relevant Act/Section reference for other exempt income for sub-category “Receipts not in the nature of income” in Schedule EI""" & Chr(13)
   ValidateNI_EI = False
    Exit Function
 End If
End If

UpdateProgressBar
Next
End Function

Function ValidateAmount_EI() As Boolean
ValidateAmount_EI = True
setTableInfo_EI1
Dim rangecells As Range
Set rangecells = Sheet24.Range("EI.Amount").Cells
Dim i As Long
ReDim Amt_EI(ColCount2)
For i = 1 To ColCount2
Amt_EI(i) = rangecells.item(i).value
If Len(Amt_EI(i)) = 0 Then
End If

If Not chkCompulsory(Amt_EI(i)) Then
   'JYOTI56
   ' msgvalidateEIError = msgvalidateEIError + "* Amount at Sr.NO " & i & " in Sheet EI, is Compulsory" & Chr(13)
   '  msgvalidateEIError = msgvalidateEIError + "* Amount is mandatory in schedule EI at Sl.NO " & i & "" & Chr(13)
    'Added by Shrutika(16/05/2025)
     msgvalidateEIError = msgvalidateEIError + "* Amount is mandatory in schedule EI at Sl.No " & i & "" & Chr(13)
    
    ValidateAmount_EI = False
    Exit Function
End If


If Not IsNumeric(Amt_EI(i)) Then
    msgvalidateEIError = msgvalidateEIError + "* Amount at Sr.NO " & i & " in Sheet EI, can allow only Numeric Values" & Chr(13)
    ValidateAmount_EI = False
    Exit Function
End If

If Len(Amt_EI(i)) > 14 Then
    msgvalidateEIError = msgvalidateEIError + "* Amount at Sr.NO " & i & " in Sheet EI, cannot be more than 14 digits" & Chr(13)
    ValidateAmount_EI = False
    Exit Function
End If
Next
End Function


Sub setTableInfo_EI1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("Others.NOI").count
    Set rangecells = Sheet24.Range("Others.NOI").Cells
    For mIntCtr = 1 To mIntCells
            If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
               ccount = ccount + 1
           End If
    Next
    ColCount2 = ccount
End Sub

Sub setTableInfo1_EI2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("EI.Amount").count
    Set rangecells = Sheet24.Range("EI.Amount").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_1 = ccount
End Sub

'Added by Shrutika- 23/04/2026 V0.3 AY-26-27


Sub setTableInfo1_EI3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("Nature_Others").count
    Set rangecells = Sheet24.Range("Nature_Others").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_2 = ccount
End Sub

Function ValidateOthersEI() As Boolean
ValidateOthersEI = True

    
   setTableInfo_EI1
   setTableInfo1_EI3
   
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellrange1 As String
    Dim i As Long
    
    Set rangecells = Range("Others.NOI").Cells
 
    ReDim Others_NOI(ColCount2_1)

    For i = 1 To ColCount2_1
   
    cellrange = GetMergedAddressCell(rangecells, i)
   
    Others_NOI(i) = Sheet24.Range(cellrange).value

        
            If Others_NOI(i) = "(Select)" Or Others_NOI(i) = "" Then
'             msgError = msgError & "Please select an option from the drop down in Exempt Income at Sr. No  " & i & "  in Sheet Taxes paid and verification" & Chr(13)
                  msgvalidateEIError = msgvalidateEIError + "* ""Please select category from the drop down in Exempt Income""" & Chr(13)
             ValidateOthersEI = False
             Exit Function
         End If

   Next


End Function



Function ValidateShtEI() As Boolean
On Error Resume Next

ValidateShtEI = True

sceiInterestInc = Sheet24.Range("Sheet20.scei.InterestInc").value
sceiDividendInc = Sheet24.Range("Sheet20.scei.DividendInc").value
'sceiLTCGWhereSTTPaid = Sheet24.Range("Sheet20.scei.LTCGWhereSTTPaid").value
sceiNetAgriIncOrOthrIncRule7 = Sheet24.Range("Sheet20.scei.NetAgriIncOrOthrIncRule7").value
sceiShareOfProfitFirmAOP = Sheet24.Range("Sheet20.scei.ShareOfProfitFirmAOP").value
sceiExpenditureOnAgriculture = Sheet24.Range("Sheet20.ExpenditureOnAgriculture").value
sceiUnabsorbedAgriculturalloss = Sheet24.Range("Sheet20.UnabsorbedAgriculturalloss").value
sceiNetAgriculturalIncome = Sheet24.Range("Sheet20.NetAgriculturalIncome").value
sceiAgriInc = Sheet24.Range("Sheet20.Agriculturalinc").value
sceiOthers = Sheet24.Range("Sheet20.scei.Others").value
sceiTotalExemptInc = Sheet24.Range("Sheet20.scei.TotalExemptInc").value

If Len(sceiTotalExemptInc) > 14 Then msgvalidateEIError = "Total should be less than 14 digits" & Chr(13)

If (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" And sceiDividendInc > 1000000) Then
    msgvalidateEIError = msgvalidateEIError + "* Dividend income from domestic company (amount not exceeding Rs. 10 lakh) should not be entered more than Rs. 10 lakh for Residential Status- " & Sheet1.Range("sheet1.ResidentialStatus1").value & " in Sheet EI" & Chr(13)
    ValidateShtEI = False
    Exit Function
End If


    setTableInfo_EI1
    setTableInfo1_EI2
    setTableInfo1_EI3

  '  If Len((Sheet24.Range("EI.NatureOfIncome").item(1).value) > 0) Then
        If Not ValidateNI_EI Then ValidateShtEI = False
        If Not ValidateAmount_EI Then ValidateShtEI = False
        If Not ValidateOthersEI Then ValidateShtEI = False
        
 '   End If
    
    
   ' If ((ColCount2 <> ColCount2_1)) Then
   ' If ((ColCount2 <> ColCount2_1)) Or ((ColCount2_2 <> ColCount2_2)) Then
   'Jyoti
   ' msgvalidateEIError = msgvalidateEIError + "* Enter All Mandatory fields in Exempt Income"
   '  msgvalidateEIError = msgvalidateEIError + "* Nature of Income is mandatory in schedule EI at Sl. No. 3" & Chr(13)
'         msgvalidateEIError = msgvalidateEIError + "* ""Please select category from the drop down in Exempt Income""" & Chr(13)
'
'    ValidateShtEI = False
'    End If
'    End If
 
   


'If Sheet24.Range("Sheet20.NetAgriculturalIncome").value > 500000 Then 'Commented by Bindu as per this defect
'New Table


    setTableInfo_EI_Agriculture1
    setTableInfo_EI_Agriculture2
    setTableInfo_EI_Agriculture3
    setTableInfo_EI_Agriculture4
    setTableInfo_EI_Agriculture5
  
 'Dim ColCount_Name_of_district As Variant
 
     ColCount_Name_of_district = Application.WorksheetFunction.Max(0, ColCount_EI1, ColCount_EI2, ColCount_EI3, ColCount_EI4, ColCount_EI5)
 
    'AY_2023_24 OLD
    'If Len((Sheet24.Range("Name_of_district").item(1).value) > 0) Then
    'AY_2024_25 New
    If Len((Sheet24.Range("Name_of_district").item(1).value) > 0) Or ColCount_Name_of_district > 0 Then
    
        If Not ValidateNameOfDistrict_EI Then ValidateShtEI = False
        If Not ValidatePincode_EI Then ValidateShtEI = False
        If Not Validate_AgriculturalLand_In_Acre_EI Then ValidateShtEI = False
        If Not Validate_AgriculturalLand_Owned_On_Held_onlease_EI Then ValidateShtEI = False
        If Not Validate_AgriculturalLand_irrigated_or_rainfed_EI Then ValidateShtEI = False
        
    End If
    
     
'End If ---By Bindu SIT-91932

    'Malli comented-------------------------
'    setTableInfo_EI_Agriculture2
'    setTableInfo_EI_Agriculture3
'    setTableInfo_EI_Agriculture4
'    setTableInfo_EI_Agriculture5
    '------------------------------------------
    
    
    'comented by Malli---------
    If Sheet24.Range("Sheet20.NetAgriculturalIncome").value > 500000 Then
       'If ColCount_EI1 = 0 Then
        If ColCount_Name_of_district = 0 Then
       'Commented by Shrutika
       'msgvalidateEIError = msgvalidateEIError + "* Please enter the Name of district at point 2(vi)(a)" & Chr(13)
       'Added by Shrutika
       'msgvalidateEIError = msgvalidateEIError + "* ""Please enter the Name of district at point 3(v)(a)"" in Schedule EI.""" & Chr(13)
        'added by Malli
        msgvalidateEIError = msgvalidateEIError + "* ""Please enter the Name of district at point 2(v)(a) in Schedule EI.""" & Chr(13)
        msgvalidateEIError = msgvalidateEIError + "* ""Please enter the pin code at point 2(v)(a) in Schedule EI.""" & Chr(13)
        msgvalidateEIError = msgvalidateEIError + "* ""Please enter the Measurement of agricultural land in Acre at point 2(v)(b) in Schedule EI.""" & Chr(13)
        msgvalidateEIError = msgvalidateEIError + "* ""Please select an option from the drop down at point 2(v)(c) in schedule EI""" & Chr(13)
        msgvalidateEIError = msgvalidateEIError + "* ""Please select an option from the drop down at point 2(v)(d) in Schedule EI.""" & Chr(13)
        msgvalidateEIError = msgvalidateEIError + "* ""Enter All Mandatory fields in Table 2(V)""" & Chr(13)
       
       ValidateShtEI = False
       Exit Function
       End If


    End If

    '--------------------------
   
    If ((ColCount_EI1 <> ColCount_EI2) Or (ColCount_EI1 <> ColCount_EI3) Or (ColCount_EI1 <> ColCount_EI4) Or (ColCount_EI1 <> ColCount_EI5)) Then
    'Ankita
 '  msgvalidateEIError = msgvalidateEIError + "* Enter All Mandatory fields in Table 3(Vi)" & Chr(13)
    msgvalidateEIError = msgvalidateEIError + "* Enter All Mandatory fields in Table 2(V)" & Chr(13)
    ValidateShtEI = False
    Exit Function
    End If
    
    If Sheet9.Range("sheet11.TotExempIncPL").value > (Sheet24.Range("EI.Total").value) Then
    msgvalidateEIError = msgvalidateEIError + "* In schedule BP Amount of exempt income reduced from Profits and Gains of Business and Profession does not tally with income offered in Schedule EI. " & Chr(13)
    ValidateShtEI = False
    Exit Function
    End If
    
setTableInfo_EIDTAA_1
setTableInfo_EIDTAA_2
setTableInfo_EIDTAA_3
setTableInfo_EIDTAA_4
setTableInfo_EIDTAA_5
setTableInfo_EIDTAA_6

ColCount_EIDTAA_1 = WorksheetFunction.Max(0, ColCount_EIDTAA_1, ColCount_EIDTAA_2, ColCount_EIDTAA_3, ColCount_EIDTAA_4, ColCount_EIDTAA_5, ColCount_EIDTAA_6)
   
  If Not ValidateEIDTAA_AmountOfIncome Then ValidateShtEI = False
        If Not ValidateEIDTAA_NatureOfIncome Then ValidateShtEI = False
        If Not ValidateEIDTAA_CountryNameAndCode Then ValidateShtEI = False
        If Not ValidateEIDTAA_ArticleOfDTAA Then ValidateShtEI = False
        If Not ValidateEIDTAA_HeadOfIncome Then ValidateShtEI = False
        If Not ValidateEIDTAA_TRCObtained Then ValidateShtEI = False
 
If msgvalidateEIError <> "" Then
 ValidateShtEI = False

End If

End Function

Sub setTableInfo_EIDTAA_1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SCHEI_DTAA_AmountOfIncome").count
    Set rangecells = Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EIDTAA_1 = ccount
    End Sub
    
    Sub setTableInfo_EIDTAA_2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SCHEI_DTAA_NatureOfIncome").count
    Set rangecells = Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EIDTAA_2 = ccount
    End Sub
    
    Sub setTableInfo_EIDTAA_3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SCHEI_DTAA_CountryCode").count
    Set rangecells = Sheet24.Range("SCHEI_DTAA_CountryCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EIDTAA_3 = ccount
    End Sub
    
    Sub setTableInfo_EIDTAA_4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SCHEI_DTAA_Article").count
    Set rangecells = Sheet24.Range("SCHEI_DTAA_Article").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EIDTAA_4 = ccount
    End Sub
    
    Sub setTableInfo_EIDTAA_5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SCHEI_DTAA_HeadOfIncome").count
    Set rangecells = Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Cells
    For mIntCtr = 1 To mIntCells
            If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EIDTAA_5 = ccount
    End Sub
    
    Sub setTableInfo_EIDTAA_6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SCHEI_TRCObtained").count
    Set rangecells = Sheet24.Range("SCHEI_TRCObtained").Cells
    For mIntCtr = 1 To mIntCells
            If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EIDTAA_6 = ccount
    End Sub


Sub setTableInfo_EI_Agriculture1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("Name_of_district").count
    Set rangecells = Sheet24.Range("Name_of_district").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EI1 = ccount
    End Sub
    
    Sub setTableInfo_EI_Agriculture2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SchEI_Pincode").count
    Set rangecells = Sheet24.Range("SchEI_Pincode").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_EI2 = ccount
    End Sub
    
    Sub setTableInfo_EI_Agriculture3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet24.Range("SchEI_AgriculturalLand").count
 Set rangecells = Sheet24.Range("SchEI_AgriculturalLand").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
 Next
 ColCount_EI3 = ccount
 End Sub
 
 Sub setTableInfo_EI_Agriculture4()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet24.Range("SchEI_OwedOrLease").count
 Set rangecells = Sheet24.Range("SchEI_OwedOrLease").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 ColCount_EI4 = ccount
 End Sub

 Sub setTableInfo_EI_Agriculture5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet24.Range("SchEI_IrrigatedOrRainFed").count
    Set rangecells = Sheet24.Range("SchEI_IrrigatedOrRainFed").Cells
    For mIntCtr = 1 To mIntCells
            If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
    Next
    ColCount_EI5 = ccount
    End Sub
    
    
    Function ValidateNameOfDistrict_EI() As Boolean
ValidateNameOfDistrict_EI = True
setTableInfo_EI_Agriculture1
Dim rangecells As Range
Set rangecells = Sheet24.Range("Name_of_district").Cells
Dim i As Long
ReDim NameOfDistrict_EI(ColCount_EI1)
noOfProcessSub = ColCount_EI1

'Malli------------------

If (ColCount_EI1 <= 0 And ColCount_Name_of_district > 0) Then
    msgvalidateEIError = msgvalidateEIError + "* ""Please enter the Name of district at point 2(v)(a) in Schedule EI""" & Chr(13)
    ValidateNameOfDistrict_EI = False
    Exit Function
End If

'------------------------
For i = 1 To ColCount_EI1
NameOfDistrict_EI(i) = rangecells.item(i).value



If Not chkCompulsory(NameOfDistrict_EI(i)) Then
'Ankita
  ' msgvalidateEIError = msgvalidateEIError + "* Please enter the Name of district at point 3(vi)(a) at Sr.NO " & i & " in schedule EI" & Chr(13)
    msgvalidateEIError = msgvalidateEIError + "* ""Please enter the Name of district at point 2(v)(a) in Schedule EI""" & Chr(13)
    ValidateNameOfDistrict_EI = False
    Exit Function
End If


If Len(NameOfDistrict_EI(i)) > 125 Then
    msgvalidateEIError = msgvalidateEIError + "* ""NameOfDistrict at Serial No " & i & " cannot be more than 125 Characters in schedule EI""" & Chr(13)
    ValidateNameOfDistrict_EI = False
    Exit Function
End If
UpdateProgressBar
Next
End Function

Function ValidatePincode_EI() As Boolean
ValidatePincode_EI = True
setTableInfo_EI_Agriculture1
Dim rangecells As Range
Set rangecells = Sheet24.Range("SchEI_Pincode").Cells
Dim i As Long
ReDim Pincode_EI(ColCount_EI1)
noOfProcessSub = ColCount_EI1
For i = 1 To ColCount_EI1
Pincode_EI(i) = rangecells.item(i).value



If Not chkCompulsory(Pincode_EI(i)) Then
           'Ankita
           'msgvalidateEIError = msgvalidateEIError + "* Please enter the pin code at point 2(vi)(a) at Sr. No  " & i & "  in schedule EI" & Chr(13)
            msgvalidateEIError = msgvalidateEIError + "* ""Please enter the pin code at point 2(v)(a) in Schedule EI.""" & Chr(13)
            ValidatePincode_EI = False
            Exit Function
        End If
        
        If Not chkNumeric(Pincode_EI(i)) Then
            msgvalidateEIError = msgvalidateEIError + "* ""PinCode at Sr. No  " & i & "  in Sheet EI  only digits 0 to 9 allowed in schedule EI""" & Chr(13)
            ValidatePincode_EI = False
            Exit Function
        End If
UpdateProgressBar
Next
End Function


Function Validate_AgriculturalLand_In_Acre_EI() As Boolean
Validate_AgriculturalLand_In_Acre_EI = True
setTableInfo_EI_Agriculture1
Dim rangecells As Range
Set rangecells = Sheet24.Range("SchEI_AgriculturalLand").Cells
Dim i As Long
ReDim AgriculturalLand_In_Acre_EI(ColCount_EI1)
noOfProcessSub = ColCount_EI1
For i = 1 To ColCount_EI1
AgriculturalLand_In_Acre_EI(i) = rangecells.item(i).value

        If Not chkCompulsory(AgriculturalLand_In_Acre_EI(i)) Then
   'Ankita
   'msgvalidateEIError = msgvalidateEIError + "* Please enter the Measurement of agricultural land in Acre at point 2(vi)(b) at Sr.NO " & i & " in schedule EI" & Chr(13)
    msgvalidateEIError = msgvalidateEIError + "* ""Please enter the Measurement of agricultural land in Acre at point 2(v)(b) in Schedule EI.""" & Chr(13)
    Validate_AgriculturalLand_In_Acre_EI = False
    Exit Function
End If


If Not IsNumeric(AgriculturalLand_In_Acre_EI(i)) Then
    msgvalidateEIError = msgvalidateEIError + "* ""Measurement of agricultural land in Acre at Sr.NO " & i & "can allow only Numeric Values in schedule EI""" & Chr(13)
    Validate_AgriculturalLand_In_Acre_EI = False
    Exit Function
End If

If Len(AgriculturalLand_In_Acre_EI(i)) > 17 Then
    msgvalidateEIError = msgvalidateEIError + "* ""Measurement of agricultural land in Acre at Sr.NO " & i & "cannot be more than 14 digits in schedule EI""" & Chr(13)
    Validate_AgriculturalLand_In_Acre_EI = False
    Exit Function
End If
UpdateProgressBar
Next
End Function


Function Validate_AgriculturalLand_Owned_On_Held_onlease_EI() As Boolean
Validate_AgriculturalLand_Owned_On_Held_onlease_EI = True
setTableInfo_EI_Agriculture1
Dim rangecells As Range
Set rangecells = Sheet24.Range("SchEI_OwedOrLease").Cells
Dim i As Long
ReDim AgriculturalLand_Owned_On_Held_onlease_EI(ColCount_EI1)
noOfProcessSub = ColCount_EI1
For i = 1 To ColCount_EI1
AgriculturalLand_Owned_On_Held_onlease_EI(i) = rangecells.item(i).value

       If isdropdownblank(AgriculturalLand_Owned_On_Held_onlease_EI(i)) Then
          'Ankita
          'msgvalidateEIError = msgvalidateEIError + "*  Please select an option from the drop down at point 2(vi)(c ) at Sr. No  " & i & " in schedule EI" & Chr(13)
           msgvalidateEIError = msgvalidateEIError + "* ""Please select an option from the drop down at point 2(v)(c ) at Sr. No  " & i & " in schedule EI""" & Chr(13)
            Validate_AgriculturalLand_Owned_On_Held_onlease_EI = False
            Exit Function
        End If


UpdateProgressBar
Next
End Function

Function Validate_AgriculturalLand_irrigated_or_rainfed_EI() As Boolean
Validate_AgriculturalLand_irrigated_or_rainfed_EI = True
setTableInfo_EI_Agriculture1
Dim rangecells As Range
Set rangecells = Sheet24.Range("SchEI_IrrigatedOrRainFed").Cells
Dim i As Long
ReDim AgriculturalLand_irrigated_or_rainfed_EI(ColCount_EI1)
noOfProcessSub = ColCount_EI1
For i = 1 To ColCount_EI1
AgriculturalLand_irrigated_or_rainfed_EI(i) = rangecells.item(i).value

       If isdropdownblank(AgriculturalLand_irrigated_or_rainfed_EI(i)) Then
                    
          'Ankita
        '  msgvalidateEIError = msgvalidateEIError + "* Please select an option from the drop down at point 2(vi)(d) at Sr. No  " & i & " in schedule EI" & Chr(13)
           msgvalidateEIError = msgvalidateEIError + "* ""Please select an option from the drop down at point 2(v)(d) in Schedule EI.""" & Chr(13)
            Validate_AgriculturalLand_irrigated_or_rainfed_EI = False
            Exit Function
        End If


UpdateProgressBar
Next
End Function


 Function ValidateEIDTAA_AmountOfIncome()
 ValidateEIDTAA_AmountOfIncome = True

Dim rangecells As Range
Set rangecells = Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Cells
Dim i As Long
ReDim AmountOfIncome_EIDTAA(ColCount_EIDTAA_1)
For i = 1 To ColCount_EIDTAA_1
AmountOfIncome_EIDTAA(i) = rangecells.item(i).value


If Not chkCompulsory(AmountOfIncome_EIDTAA(i)) Then
    'JYOTI57
   ' msgvalidateEIError = msgvalidateEIError + "* Please enter the Amount of Income at Sr.NO " & i & " in schedule EI" & Chr(13)
   ' msgvalidateEIError = msgvalidateEIError + "*Please enter amount of income in Point " & i & " of DTAA table of Schedule EI " & Chr(13)
 'Added by Shrutika(16/05/2025)
    msgvalidateEIError = msgvalidateEIError + "*Please enter amount of income in Point " & 4 & " of DTAA table of Schedule EI " & Chr(13)
    
    ValidateEIDTAA_AmountOfIncome = False
    Exit Function
End If


If Not IsNumeric(AmountOfIncome_EIDTAA(i)) Then
    msgvalidateEIError = msgvalidateEIError + "* Amount of Income at Sr.NO " & i & " can allow only Numeric Values in schedule EI" & Chr(13)
    ValidateEIDTAA_AmountOfIncome = False
    Exit Function
End If

If Len(AmountOfIncome_EIDTAA(i)) > 14 Then
    msgvalidateEIError = msgvalidateEIError + "* Amount of Income at Sr.NO " & i & "cannot be more than 14 digits in schedule EI" & Chr(13)
    ValidateEIDTAA_AmountOfIncome = False
    Exit Function
End If
Next
End Function

 Function ValidateEIDTAA_NatureOfIncome()
 ValidateEIDTAA_NatureOfIncome = True

Dim rangecells As Range
Set rangecells = Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Cells
Dim i As Long
ReDim NatureOfIncome_EIDTAA(ColCount_EIDTAA_1)
For i = 1 To ColCount_EIDTAA_1
NatureOfIncome_EIDTAA(i) = rangecells.item(i).value


If Not chkCompulsory(NatureOfIncome_EIDTAA(i)) Then
    'Jyoti58
    'msgvalidateEIError = msgvalidateEIError + "* Please enterNature of Income at Sr.NO " & i & " in schedule EI" & Chr(13)
  '  msgvalidateEIError = msgvalidateEIError + "* Please enter nature of income in Point " & i & " of  DTAA table of Schedule EI" & Chr(13)
   'Added by Shrutika(16/05/2025)
    msgvalidateEIError = msgvalidateEIError + "* Please enter nature of income in Point " & 4 & " of  DTAA table of Schedule EI" & Chr(13)
    ValidateEIDTAA_NatureOfIncome = False
    Exit Function
End If

If Len(NatureOfIncome_EIDTAA(i)) > 75 Then
    msgvalidateEIError = msgvalidateEIError + "* Nature of Income at Sr.NO " & i & " cannot be more than 75 characters in schedule EI" & Chr(13)
    ValidateEIDTAA_NatureOfIncome = False
    Exit Function
End If


Next
End Function

Function ValidateEIDTAA_CountryNameAndCode()
 ValidateEIDTAA_CountryNameAndCode = True

Dim rangecells As Range
Set rangecells = Sheet24.Range("SCHEI_DTAA_CountryCode").Cells
Dim i As Long
Dim CountryName As Variant
ReDim CountryNameA1_EIDTAA(ColCount_EIDTAA_1)
ReDim CountryCodeA1_EIDTAA(ColCount_EIDTAA_1)
For i = 1 To ColCount_EIDTAA_1

CountryName = rangecells.item(i).value
        
        If isdropdownblank(CountryName) Then
          'Jyoti59
           ' msgvalidateEIError = msgvalidateEIError + "* Please enter the Country Name and Code at Sr. No  " & i & " in schedule EI" & Chr(13)
       '     msgvalidateEIError = msgvalidateEIError + "*Please select Country name and code in Point  " & i & " of DTAA table of Schedule EI " & Chr(13)
          'Added by Shrutika(16/05/2025)
            msgvalidateEIError = msgvalidateEIError + "*Please select Country name and code in Point  " & 4 & " of DTAA table of Schedule EI " & Chr(13)
            ValidateEIDTAA_CountryNameAndCode = False
            Exit Function
        End If
        
        If CountryName <> "" Then
        CountryNameA1_EIDTAA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
         CountryCodeA1_EIDTAA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
         End If
        
        
        If Not checkfieldspecialcharacter(CountryCodeA1_EIDTAA(i)) Then
            msgvalidateEIError = msgvalidateEIError + "* Country Name at Sr. No  " & i & " in Sheet EI, characters < > & ' " & Chr(34) & " are not allowed in schedule EI" & Chr(13)
            ValidateEIDTAA_CountryNameAndCode = False
            Exit Function
        End If
        UpdateProgressBar
Next
End Function

    
Function ValidateEIDTAA_ArticleOfDTAA()
 ValidateEIDTAA_ArticleOfDTAA = True

Dim rangecells As Range
Set rangecells = Sheet24.Range("SCHEI_DTAA_Article").Cells
Dim i As Long
ReDim Article_EIDTAA(ColCount_EIDTAA_1)
For i = 1 To ColCount_EIDTAA_1
Article_EIDTAA(i) = rangecells.item(i).value


If Not chkCompulsory(Article_EIDTAA(i)) Then
    'Jyoti60
   ' msgvalidateEIError = msgvalidateEIError + "* Please enter Article of DTAA at Sr.NO " & i & " in schedule EI" & Chr(13)
  '  msgvalidateEIError = msgvalidateEIError + "* Please enter Article of DTAA in Point " & i & " of DTAA table of Schedule EI" & Chr(13)
    'Added by Shrutika(16/05/2025)
     msgvalidateEIError = msgvalidateEIError + "* Please enter Article of DTAA in Point " & 4 & " of DTAA table of Schedule EI" & Chr(13)
    ValidateEIDTAA_ArticleOfDTAA = False
    Exit Function
End If

If Len(Article_EIDTAA(i)) > 16 Then
    msgvalidateEIError = msgvalidateEIError + "* Article of DTAA at Sr.NO " & i & "cannot be more than 16 characters in schedule EI" & Chr(13)
    ValidateEIDTAA_ArticleOfDTAA = False
    Exit Function
End If


Next
End Function

Function ValidateEIDTAA_HeadOfIncome()
 ValidateEIDTAA_HeadOfIncome = True

Dim rangecells As Range
Set rangecells = Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Cells
Dim i As Long
ReDim HeadOfIncome_EIDTAA(ColCount_EIDTAA_1)
For i = 1 To ColCount_EIDTAA_1
HeadOfIncome_EIDTAA(i) = rangecells.item(i).value


If isdropdownblank(HeadOfIncome_EIDTAA(i)) Then
   'Jyoti61
   ' msgvalidateEIError = msgvalidateEIError + "* Please enter the Head Of Income at Sr.NO " & i & " in schedule EI" & Chr(13)
  '  msgvalidateEIError = msgvalidateEIError + "* Please select head of income in Point  " & i & " of DTAA table of Schedule EI" & Chr(13)
'Added by Shrutika(16/05/2025)
    msgvalidateEIError = msgvalidateEIError + "* Please select head of income in Point  " & 4 & " of DTAA table of Schedule EI" & Chr(13)

    ValidateEIDTAA_HeadOfIncome = False
    Exit Function
End If



Next
End Function

Function ValidateEIDTAA_TRCObtained()
 ValidateEIDTAA_TRCObtained = True

Dim rangecells As Range
Set rangecells = Sheet24.Range("SCHEI_TRCObtained").Cells
Dim i As Long
ReDim TRC_EIDTAA(ColCount_EIDTAA_1)
For i = 1 To ColCount_EIDTAA_1
TRC_EIDTAA(i) = rangecells.item(i).value


If isdropdownblank(TRC_EIDTAA(i)) Then
    
    'Jyoti62
   'msgvalidateEIError = msgvalidateEIError + "* Please enter Whether TRC Obtained at Sr.NO " & i & "in schedule EI" & Chr(13)
'    msgvalidateEIError = msgvalidateEIError + "* Please select whether TRC obtained in Point  " & i & " of DTAA table of Schedule EI " & Chr(13)
   'Added by Shrutika(16/05/2025)
    msgvalidateEIError = msgvalidateEIError + "* Please select whether TRC obtained in Point  " & 4 & " of DTAA table of Schedule EI " & Chr(13)
    
    ValidateEIDTAA_TRCObtained = False
    Exit Function
End If



Next
End Function


Sub DPMAddComment45(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption is allowed to the extent such income does not exceed one thousand five hundred rupees in respect of each minor child whose income is so includible;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment46(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26) is available only to certain categories of tax payers in NER and Ladakh"""
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment47(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26AAA) is available only to certain categories of Sikkimese tax payers"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment48(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed sixty per cent of the total amount payable  at the time of such closure or opting out of the scheme"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment49(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent that it does not exceed sixty per cent. of the individual corpus"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment50(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment51(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub
Sub DPMAddComment52(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption u/s 10(12C) can only be claimed by nominees of Agniveer or Disabled Agniveer under the Agnipath Scheme"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub



Sub DPMADeleteComment45(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
Sheet24.Protect Password:=sPassword
End Sub

    
