Attribute VB_Name = "SchTPSA"
Public BSR_TPSC, BankName_TPSC, Date_TPSC, SrlNoOfChaln_TPSC, Amount_TPSC, endTPSC, endTPSC1, endTPSC2, endTPSC3, endTPSC4, rngname_TPSC As Variant
Public MsgBox_TPSC As Variant
Sub ValidateSheetTPSC_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 2: AY: 2024-25" 'dpk1601
 
 ValidateTPSC
 
fmsgboxStatus "Sheet TPSA is OK"

End Sub
Function ValidateTPSC()
ValidateTPSC = True


If Not Chk92CEAmount Then ValidateTPSC = False
If Sheet5.Range("sheet7.Section92CE_Flag").value = "Yes" Then
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    mIntCells = Range("TPSA.Amount").count
    Set rangecells = Range("TPSA.Amount").Cells
    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'        If rangecells.item(mIntCtr).value = "" Then
'            MsgBox_TPSC = MsgBox_TPSC + "*Please enter Amount of primary adjustment on which option u/s 92CE(2A) is exercised & such excess money has not been repatriated within the prescribed time" & Chr(13)
'            ValidateTPSC = False
'            Exit For
'        End If
'         If Not rangecells.item(mIntCtr).value = "" Then
'          If Range("FinYr").Cells.item(mIntCtr).value = "" Then
'            MsgBox_TPSC = MsgBox_TPSC + "*Please enter Please select the dropdown against the  Amount of primary adjustment on which option u/s 92CE(2A) is exercised & such excess money has not been repatriated within the prescribed time" & Chr(13)
'            ValidateTPSC = False
'            Exit For
'          End If
'        End If
'    Next
End If
If Not Validate_TPSC_TaxDetail Then ValidateTPSC = False
If Not ValidateTPSC Then
fmsgboxsmall MsgBox_TPSC
Sheet57.Activate
End
End If

End Function

Function Chk92CEAmount() As Boolean
Chk92CEAmount = True


If Sheet5.Range("sheet7.Section92CE_Flag").value = "Yes" Then
If Sheet57.Range("TPSA_92CE_Amount").value = "" Then
    MsgBox_TPSC = MsgBox_TPSC + "*please enter amount of primary adjustment on which option u/s 92CE(2A) is exercised & such excess money has not been repatriated within the prescribed time in schedule TPSA" & Chr(13)

        Chk92CEAmount = False
        
        
        Exit Function
End If
End If
 

End Function

Sub AddRows_TPSC()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("TPSC_BSRCode||TPSC_BankName||TPSC_DateDep||TPSC_SrlNoOfChaln||TPSC_Amt")
ActiveCellRange = EfilingCommon.searchLastRow("TPSC_BSRCode")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Function Validate_TPSC_TaxDetail() As Boolean
    Validate_TPSC_TaxDetail = True

        setTableInfo_TPSC
        setTableInfo_TPSC1
        setTableInfo_TPSC2
        setTableInfo_TPSC3
        setTableInfo_TPSC4

               
    endTPSC = WorksheetFunction.Max(0, endTPSC, endTPSC1, endTPSC2, endTPSC3, endTPSC4)
    
            If Not ValidateBSRCode_TPSC Then Validate_TPSC_TaxDetail = False
            If Not ValidateBankName_TPSC Then Validate_TPSC_TaxDetail = False
            If Not Validate_date_TPSC Then Validate_TPSC_TaxDetail = False
            If Not Validate_SrChallan_TPSC Then Validate_TPSC_TaxDetail = False
            If Not Validate_Amount_TPSC Then Validate_TPSC_TaxDetail = False
       
End Function
Function ValidateBSRCode_TPSC() As Boolean
    ValidateBSRCode_TPSC = True
    Dim rangecells As Range
    Set rangecells = Sheet57.Range("TPSC_BSRCode").Cells
    Dim i As Long
    ReDim BSR_TPSC(endTPSC)
    
    For i = 1 To endTPSC
        BSR_TPSC(i) = rangecells.item(i).value
        
        If Not chkCompulsory(BSR_TPSC(i)) Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter the BSR code  at Sr. No " & i & " in Schedule TPSA" & Chr(13)
            ValidateBSRCode_TPSC = False
            Exit Function
        End If
        
        If Not Len(BSR_TPSC(i)) = 0 Then
            If Not (chkNumeric(Mid(BSR_TPSC(i), 1, 3)) And checkfieldspecialcharacter(Mid(BSR_TPSC(i), 4, 4))) Then
            
                MsgBox_TPSC = MsgBox_TPSC + "*Please enter a valid 7 digit BSR Code at Sr. No  " & i & " in Schedule TPSA " & Chr(13)
                ValidateBSRCode_TPSC = False
                Exit Function
                End If
          End If

        
    Next
End Function
Function ValidateBankName_TPSC() As Boolean
    ValidateBankName_TPSC = True
    Dim rangecells As Range
    Set rangecells = Sheet57.Range("TPSC_BankName").Cells
    Dim i As Long
    ReDim BankName_TPSC(endTPSC)
    
    For i = 1 To endTPSC
        BankName_TPSC(i) = rangecells.item(i).value
        
        If Not chkCompulsory(BankName_TPSC(i)) Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter name of bank and branch at Sr. No " & i & " in Schedule TPSA" & Chr(13)
            ValidateBankName_TPSC = False
            Exit Function
        End If
        
        If Len(BankName_TPSC(i)) > 125 Then
                MsgBox_TPSC = MsgBox_TPSC + "*bank and branch name cannot exceed 125 charcters at Sr. No  " & i & " in Schedule TPSA " & Chr(13)
                ValidateBankName_TPSC = False
                Exit Function
                End If
        

        
    Next
End Function
Function Validate_date_TPSC() As Boolean
    Validate_date_TPSC = True
    Dim rangecells As Range
    Set rangecells = Sheet57.Range("TPSC_DateDep").Cells
    Dim i As Long
    ReDim Date_TPSC(endTPSC)
    
    For i = 1 To endTPSC
        Date_TPSC(i) = rangecells.item(i).value
        If Len(Date_TPSC(i)) = 0 Then
        End If
        
        If Not chkCompulsory(Date_TPSC(i)) Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter the date of Deposit at Sr. No " & i & " in Schedule TPSA" & Chr(13)
            Validate_date_TPSC = False
            Exit Function
        End If
    
        
    Next
End Function
Function Validate_SrChallan_TPSC() As Boolean
    Validate_SrChallan_TPSC = True
    Dim rangecells As Range
    Set rangecells = Sheet57.Range("TPSC_SrlNoOfChaln").Cells
    Dim i As Long
    ReDim SrlNoOfChaln_TPSC(endTPSC)
    
    For i = 1 To endTPSC
        SrlNoOfChaln_TPSC(i) = rangecells.item(i).value
        
        If Not chkCompulsory(SrlNoOfChaln_TPSC(i)) Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter the serial number of Challan at Sr. No " & i & " in Schedule TPSA" & Chr(13)
            Validate_SrChallan_TPSC = False
            Exit Function
        End If
    
        If Not IsNumeric(SrlNoOfChaln_TPSC(i)) Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter the serial number of Challan in Numeric Values at Sr. No " & i & " in Schedule TPSA" & Chr(13)
            Validate_SrChallan_TPSC = False
            Exit Function
        End If
    Next
End Function
Function Validate_Amount_TPSC() As Boolean
    Validate_Amount_TPSC = True
    Dim rangecells As Range
    Set rangecells = Sheet57.Range("TPSC_Amt").Cells
    Dim i As Long
    ReDim Amount_TPSC(endTPSC)
    
    For i = 1 To endTPSC
        Amount_TPSC(i) = rangecells.item(i).value
        If Len(Amount_TPSC(i)) = 0 Then
        End If
        
        If Not chkCompulsory(Amount_TPSC(i)) Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter the amount at Sr. No " & i & " in Schedule TPSA" & Chr(13)
            Validate_Amount_TPSC = False
            Exit Function
        End If
    Next
    
    If Sheet57.Range("TPSA_92CE_Amount").value > 0 Then
        If Sheet57.Range("TPSC_Amount_Deposited").value = 0 Then
            MsgBox_TPSC = MsgBox_TPSC + "*Please enter the amount in Schedule TPSA" & Chr(13)
            Validate_Amount_TPSC = False
            Exit Function
        End If
    End If
    
    
End Function

Sub setTableInfo_TPSC()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet57.Range("TPSC_BSRCode").Cells
    mIntCells = Sheet57.Range("TPSC_BSRCode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    endTPSC = ccount
    'rngname_TPSC = "TPSC_BSRCode;TPSC_BankName;TPSC_DateDep;TPSC_SrlNoOfChaln;TPSC_Amt;TPSC_Amount_Deposited;"
    rngname_TPSC = "TPSC_BSRCode;TPSC_BankName;TPSC_DateDep;TPSC_SrlNoOfChaln;TPSC_Amt;"
End Sub

Sub setTableInfo_TPSC1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet57.Range("TPSC_BankName").Cells
    mIntCells = Sheet57.Range("TPSC_BankName").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    endTPSC1 = ccount
End Sub

Sub setTableInfo_TPSC2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet57.Range("TPSC_DateDep").Cells
    mIntCells = Sheet57.Range("TPSC_DateDep").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    endTPSC2 = ccount
End Sub

Sub setTableInfo_TPSC3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet57.Range("TPSC_SrlNoOfChaln").Cells
    mIntCells = Sheet57.Range("TPSC_SrlNoOfChaln").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    endTPSC3 = ccount
End Sub

Sub setTableInfo_TPSC4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet57.Range("TPSC_Amt").Cells
    mIntCells = Sheet57.Range("TPSC_Amt").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    endTPSC4 = ccount
End Sub


