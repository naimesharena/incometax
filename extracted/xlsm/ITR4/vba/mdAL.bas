Attribute VB_Name = "mdAL"
'Option Explicit
'
'Dim msgboxAL As String
'
'Public IsImmovableAsset As Variant
'Public ImmovableAssetDesc As Variant
'Public ImmovableAssetAddress As Variant
'Public ImmovableAssetAddress1 As Variant
'Public ImmovableAssetAddress2 As Variant
'Public ImmovableAssetAddress3 As Variant
'Public ImmovableAssetAddress4 As Variant
'Public ImmovableAssetAddress5 As Variant
'Public ImmovableAssetAddress6 As Variant
'Public ImmovableAssetAddress7 As Variant
'Public ImmovableAssetAddress8 As Variant
'Public ImmovableAssetAddress9 As Variant
'Public ImmovableAssetAmount As Variant
'
'Public count_AL1, count_AL2, count_AL3, count_AL4, count_AL5, count_AL6, count_AL7, count_AL8, count_AL9 As Variant
'Public range_AL1, range_AL2, range_AL3 As Variant
'
'Public JewelleryBullionEtc_AL As Variant
'Public ArchCollDrawPaintSulpArt_AL As Variant
'Public VehiclYachtsBoatsAircrafts_AL As Variant
'
'Public DepositsInBank_AL As Variant
'Public SharesAndSecurities_AL As Variant
'Public InsurancePolicies_AL As Variant
'Public LoansAndAdvancesGiven_AL As Variant
'Public CashInHand_AL As Variant
'
'Public IsInterestHeld As Variant
'Public InterestHeldNameOfFirm As Variant
'Public InterestHeldAddressOfFirm As Variant
'Public InterestHeldAddressOfFirm1 As Variant
'Public InterestHeldAddressOfFirm2 As Variant
'Public InterestHeldAddressOfFirm3 As Variant
'Public InterestHeldAddressOfFirm4 As Variant
'Public InterestHeldAddressOfFirm5 As Variant
'Public InterestHeldAddressOfFirm6 As Variant
'Public InterestHeldAddressOfFirm7 As Variant
'Public InterestHeldAddressOfFirm8 As Variant
'Public InterestHeldAddressOfFirm9 As Variant
'Public InterestHeldPANOfFirm As Variant
'Public InterestHeldInvestmentOfFirm As Variant
'
'Public count_AL2_1, count_AL2_2, count_AL2_3, count_AL2_4, count_AL2_5, count_AL2_6, count_AL2_7, count_AL2_8, count_AL2_9, count_AL2_10 As Variant
'Public range_AL2_1, range_AL2_2, range_AL2_3, range_AL2_4 As Variant
'Public LiabilityInRelatAssets_AL As Variant
'
'Public TotalIncome_8b As Variant
'
'Sub ValidateAL_Click()
'ValidateSchAL
'MsgBox "Sheet is ok", vbOKOnly, "ITR-4"
'End Sub
'
'
'Sub AddRows_ImmovableAssets()
'    Sheets("AL").Activate
'    mdCommon.DefinedgridNameRange = "SchAL.A.Description||SchAL.A.Address_Flat||SchAL.A.Address_Village||SchAL.A.Address_Road||SchAL.A.Address_Area||SchAL.A.Address_City||SchAL.A.Address_State||SchAL.A.Address_Country||SchAL.A.Address_Pin||SchAL.A.Address_Zip||SchAL.A.Amount"
'    ActiveCellRange = mdCommon.searchLastRow("SchAL.A.Description")
'    mdCommon.insertRowUnderSectionWithFormula
'End Sub
'
'Sub AddRows_InterestHeld()
'    Sheets("AL").Activate
'    mdCommon.DefinedgridNameRange = "SchAL.C.Name||SchAL.C.Address_Flat||SchAL.C.Address_Village||SchAL.C.Address_Road||SchAL.C.Address_Area||SchAL.C.Address_City||SchAL.C.Address_State||SchAL.C.Address_Country||SchAL.C.Address_Pin||SchAL.C.Address_Zip||SchAL.C.PAN||SchAL.C.Investment"
'    ActiveCellRange = mdCommon.searchLastRow("SchAL.C.Name")
'    mdCommon.insertRowUnderSectionWithFormula
'End Sub
'
'
'Sub ValidateSchAL()
'
' If Not ValidateSheetSchAL_1 Then
'    Sheet13.Activate
'    MsgBox msgboxAL, vbOKOnly, "Error(s)!"
'    CloseMsg
'End If
'
'End Sub
'
'Function ValidateSheetSchAL_1() As Boolean
'ValidateSheetSchAL_1 = True
'
'    msgboxAL = ""
'    Dim status_6 As Variant
'
'    status_6 = Mid(Sheet1.Range("sheet1.Status"), 1, 2)
'
'    If Trim(status_6) = "F" Then Exit Function
'
'    TotalIncome_8b = Sheet1.Range("IncD.TotalIncome").value
'    IsImmovableAsset = Sheet13.Range("IsImmovableAsset").value
'
'    If (TotalIncome_8b > 5000000 And isdropdownblank(IsImmovableAsset)) Then
'        msgboxAL = msgboxAL + "Do you own any immovable asset? is Mandatory.Please select an option from the dropdown in Sheet AL" & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'    If UCase(IsImmovableAsset) = "YES" Then
'        setTableInfo_AL1
'        setTableInfo_AL2
'        setTableInfo_AL3
'        setTableInfo_AL4
'        setTableInfo_AL5
'        setTableInfo_AL6
'        setTableInfo_AL7
'        setTableInfo_AL8
'        setTableInfo_AL9
'
'        If (count_AL1 <> count_AL2) Or _
'            (count_AL1 <> count_AL3) Or _
'            (count_AL1 <> count_AL4) Or _
'            (count_AL1 <> count_AL5) Or _
'            (count_AL1 <> count_AL7) Or _
'            (count_AL1 <> count_AL8) Then
'
'            msgboxAL = msgboxAL + "Details of immovable asset: Please fill all the details" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'        If (count_AL1 = 0) Then
'           msgboxAL = msgboxAL + "Atleast one details of immovable asset requires in Sheet AL." & Chr(13)
'           ValidateSheetSchAL_1 = False
'           Exit Function
'        End If
'
'        If Not ValidateImmovableAssetsDesc_AL Then ValidateSheetSchAL_1 = False
'        If Not ValidateImmovableAssetAddress_AL Then ValidateSheetSchAL_1 = False
'        If Not ValidateImmovableAssetAmount_AL Then ValidateSheetSchAL_1 = False
'    End If
'
'
'    IsInterestHeld = Sheet13.Range("IsInterestHeld").value
'    If (TotalIncome_8b > 5000000 And isdropdownblank(IsInterestHeld)) Then
'        msgboxAL = msgboxAL + "Do you have any Interest held in the assets of a firm or AOP as a partner or member thereof? is Mandatory.Please select an option from the dropdown in Sheet AL" & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If UCase(IsInterestHeld) = "YES" Then
'        setTableInfo_AL2_1
'        setTableInfo_AL2_2
'        setTableInfo_AL2_3
'        setTableInfo_AL2_4
'        setTableInfo_AL2_5
'        setTableInfo_AL2_6
'        setTableInfo_AL2_7
'        setTableInfo_AL2_8
'        setTableInfo_AL2_9
'        setTableInfo_AL2_10
'
'        If (count_AL2_1 <> count_AL2_2) Or _
'            (count_AL2_1 <> count_AL2_3) Or _
'            (count_AL2_1 <> count_AL2_4) Or _
'            (count_AL2_1 <> count_AL2_5) Or _
'            (count_AL2_1 <> count_AL2_6) Or _
'            (count_AL2_1 <> count_AL2_8) Or _
'            (count_AL2_1 <> count_AL2_9) Then
'            msgboxAL = msgboxAL + "Interest held in the assets of a firm or AOP: Please fill all the details" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'        If (count_AL2_1 = 0) Then
'           msgboxAL = msgboxAL + "Atleast one details Interest held in the assets of a firm or AOP requires in Sheet AL." & Chr(13)
'           ValidateSheetSchAL_1 = False
'           Exit Function
'        End If
'
'        If Not ValidateInterestHeldNameOfFirm_AL Then ValidateSheetSchAL_1 = False
'        If Not ValidateInterestHeldAddressOfFirm_AL Then ValidateSheetSchAL_1 = False
'        If Not ValidateInterestHeldPANOfFirm_AL Then ValidateSheetSchAL_1 = False
'        If Not ValidateInterestHeldInvestmentOfFirm_AL Then ValidateSheetSchAL_1 = False
'    End If
'
'
'    JewelleryBullionEtc_AL = Sheet13.Range("SchAL.B.Jwellery").value
'    ArchCollDrawPaintSulpArt_AL = Sheet13.Range("SchAL.B.Art").value
'    VehiclYachtsBoatsAircrafts_AL = Sheet13.Range("SchAL.B.Vehicles").value
'
'    DepositsInBank_AL = Sheet13.Range("SchAL.B.Bank").value
'    SharesAndSecurities_AL = Sheet13.Range("SchAL.B.Share").value
'    InsurancePolicies_AL = Sheet13.Range("SchAL.B.Insurance").value
'    LoansAndAdvancesGiven_AL = Sheet13.Range("SchAL.B.Loans").value
'    CashInHand_AL = Sheet13.Range("SchAL.B.Cash").value
'
'    LiabilityInRelatAssets_AL = Sheet13.Range("SchAL.Liability").value
'
'
' Dim blankChkA As Boolean
'    Dim blankChkCntB As Long
'    Dim blankChkC As Boolean
'    Dim blankChkBD As Boolean
'
'    blankChkA = False
'    blankChkC = False
'    blankChkBD = False
'    blankChkCntB = 0
'
'    If isdropdownblank(IsImmovableAsset) Then
'        blankChkA = True
'    End If
'
'    If isdropdownblank(IsInterestHeld) Then
'        blankChkC = True
'    End If
'
'    If JewelleryBullionEtc_AL = "" Or _
'        ArchCollDrawPaintSulpArt_AL = "" Or _
'        VehiclYachtsBoatsAircrafts_AL = "" Or _
'        DepositsInBank_AL = "" Or _
'        SharesAndSecurities_AL = "" Or _
'        InsurancePolicies_AL = "" Or _
'        LoansAndAdvancesGiven_AL = "" Or _
'        CashInHand_AL = "" Or _
'        LiabilityInRelatAssets_AL = "" Then
'        blankChkBD = True
'    End If
'
'    If JewelleryBullionEtc_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If ArchCollDrawPaintSulpArt_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If VehiclYachtsBoatsAircrafts_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If DepositsInBank_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If SharesAndSecurities_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If InsurancePolicies_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If LoansAndAdvancesGiven_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If CashInHand_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'    If LiabilityInRelatAssets_AL <> "" Then
'        blankChkCntB = blankChkCntB + 1
'    End If
'
'
'    If Not (blankChkA) Then
'        If (blankChkBD) Then
'            msgboxAL = msgboxAL + "Please fill the amount fields at Point B or / and  Point D. If not applicable, enter zeros, in Schedule AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'
'        If (blankChkC) Then
'            msgboxAL = msgboxAL + "Please select an option at Point C in Schedule AL." & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'        Exit Function
'    End If
'
'    If Not (blankChkBD) Or blankChkCntB >= 1 Then
'        If (blankChkA) Then
'            msgboxAL = msgboxAL + "Please select an option at Point A in Schedule AL." & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'
'        If blankChkCntB <> 9 Then
'            msgboxAL = msgboxAL + "Please fill the amount fields at Point B or / and  Point D. If not applicable, enter zeros, in Schedule AL." & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'
'        If (blankChkC) Then
'            msgboxAL = msgboxAL + "Please select an option at Point C in Schedule AL." & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'        Exit Function
'    End If
'
'    If Not (blankChkC) Then
'        If (blankChkA) Then
'            msgboxAL = msgboxAL + "Please select an option at Point A in Schedule AL." & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'
'        If (blankChkBD) Then
'            msgboxAL = msgboxAL + "Please fill the amount fields at Point B or / and  Point D. If not applicable, enter zeros, in Schedule AL." & Chr(13)
'            ValidateSheetSchAL_1 = False
'            'Exit Function
'        End If
'        Exit Function
'    End If
'
'
'    If JewelleryBullionEtc_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Movable Asset(Jewellery,bullion etc.) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If ArchCollDrawPaintSulpArt_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Archaeological collections, drawings, painting, sculpture or any work of art should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If VehiclYachtsBoatsAircrafts_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Movable Asset(Vehicles,yachts,boats and aircrafts) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'
'    If DepositsInBank_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in bank(including all deposits) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If SharesAndSecurities_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in shares and securities should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If LoansAndAdvancesGiven_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Loans and advances given should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If InsurancePolicies_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Insurance policies should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If CashInHand_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Cash in hand should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'    If LiabilityInRelatAssets_AL < 0 Then
'        msgboxAL = msgboxAL + "Amount in Liability in relation to Assets should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidateSheetSchAL_1 = False
'    End If
'
'
'    '50lkh Check
'    If ((TotalIncome_8b > 5000000)) Then
'
'        If DepositsInBank_AL = "" Or IsEmpty(DepositsInBank_AL) Then
'            msgboxAL = msgboxAL + "Amount of Deposits in bank is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'
'        If SharesAndSecurities_AL = "" Or IsEmpty(SharesAndSecurities_AL) Then
'            msgboxAL = msgboxAL + "Investment in Shares and Securities is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'
'        If InsurancePolicies_AL = "" Or IsEmpty(InsurancePolicies_AL) Then
'            msgboxAL = msgboxAL + "Investment in Insurance Policies is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'        If LoansAndAdvancesGiven_AL = "" Or IsEmpty(LoansAndAdvancesGiven_AL) Then
'            msgboxAL = msgboxAL + "Investment in Loans and Advances Given is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'
'        If CashInHand_AL = "" Or IsEmpty(CashInHand_AL) Then
'             msgboxAL = msgboxAL + "Details of Cash in hand is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'
'        If JewelleryBullionEtc_AL = "" Or IsEmpty(JewelleryBullionEtc_AL) Then
'             msgboxAL = msgboxAL + "Details of Jewellery / Bullion etc is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'
'        If ArchCollDrawPaintSulpArt_AL = "" Or IsEmpty(ArchCollDrawPaintSulpArt_AL) Then
'             msgboxAL = msgboxAL + "Details of Arch Collections etc is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'
'        If VehiclYachtsBoatsAircrafts_AL = "" Or IsEmpty(VehiclYachtsBoatsAircrafts_AL) Then
'             msgboxAL = msgboxAL + "Details of Vehicles, Yachts etc is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'
'        If LiabilityInRelatAssets_AL = "" Or IsEmpty(LiabilityInRelatAssets_AL) Then
'            msgboxAL = msgboxAL + "Details of Liabilityin relation at assets is Mandatory in Sheet AL" & Chr(13)
'            ValidateSheetSchAL_1 = False
'        End If
'    End If
'End Function
'
'Sub setTableInfo_AL1()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Description").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Description").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL1 = ccount
'    range_AL1 = "SchAL.A.Description;SchAL.A.Address_Flat;SchAL.A.Address_Village;SchAL.A.Address_Road;SchAL.A.Address_Area;SchAL.A.Address_City;SchAL.C.Address_State;SchAL.A.Address_Country;SchAL.A.Address_Pin;SchAL.A.Address_Zip;SchAL.A.Amount;"
'End Sub
'
'Sub setTableInfo_AL2()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_Flat").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_Flat").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2 = ccount
'End Sub
'
'
'Sub setTableInfo_AL3()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Amount").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Amount").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL3 = ccount
'End Sub
'
'
'Sub setTableInfo_AL4()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_Area").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_Area").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL4 = ccount
'End Sub
'
'Sub setTableInfo_AL5()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_City").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_City").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL5 = ccount
'End Sub
'
'Sub setTableInfo_AL6()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_Pin").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_Pin").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL6 = ccount
'End Sub
'
'Sub setTableInfo_AL7()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_State").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_State").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL7 = ccount
'End Sub
'
'Sub setTableInfo_AL8()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_Country").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_Country").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL8 = ccount
'End Sub
'
'
'Sub setTableInfo_AL9()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.A.Address_Zip").Count
'    Set rangecells = Sheet13.Range("SchAL.A.Address_Zip").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL9 = ccount
'End Sub
'
'Function ValidateImmovableAssetsDesc_AL() As Boolean
'ValidateImmovableAssetsDesc_AL = True
'    Dim rangecells, rangecells_code, rangecells2 As Range
'    Dim a As Long
'    Dim i As Long
'
'    setTableInfo_AL1
'
'    Set rangecells = Sheet13.Range("SchAL.A.Description").Cells
'    ReDim ImmovableAssetDesc(count_AL1)
'
'    For i = 1 To count_AL1
'        ImmovableAssetDesc(i) = rangecells.item(i).value
'            If ImmovableAssetDesc(i) = "" Then
'                msgboxAL = msgboxAL + "Description of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateImmovableAssetsDesc_AL = False
'                Exit Function
'            End If
'    Next
'End Function
'
'
'Function ValidateImmovableAssetAddress_AL() As Boolean
'ValidateImmovableAssetAddress_AL = True
'    Dim rangecells, rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8 As Range
'    Dim a As Long
'    Dim i As Long
'
'    setTableInfo_AL1
'
'    Set rangecells = Sheet13.Range("SchAL.A.Address_Flat").Cells
'    Set rangecells1 = Sheet13.Range("SchAL.A.Address_Village").Cells
'    Set rangecells2 = Sheet13.Range("SchAL.A.Address_Road").Cells
'    Set rangecells3 = Sheet13.Range("SchAL.A.Address_Area").Cells
'    Set rangecells4 = Sheet13.Range("SchAL.A.Address_City").Cells
'    Set rangecells5 = Sheet13.Range("SchAL.A.Address_State").Cells
'    Set rangecells6 = Sheet13.Range("SchAL.A.Address_Country").Cells
'    Set rangecells7 = Sheet13.Range("SchAL.A.Address_Pin").Cells
'    Set rangecells8 = Sheet13.Range("SchAL.A.Address_Zip").Cells
'
'    ReDim ImmovableAssetAddress(count_AL1)
'    ReDim ImmovableAssetAddress1(count_AL1)
'    ReDim ImmovableAssetAddress2(count_AL1)
'    ReDim ImmovableAssetAddress3(count_AL1)
'    ReDim ImmovableAssetAddress4(count_AL1)
'    ReDim ImmovableAssetAddress5(count_AL1)
'    ReDim ImmovableAssetAddress6(count_AL1)
'    ReDim ImmovableAssetAddress7(count_AL1)
'    ReDim ImmovableAssetAddress8(count_AL1)
'    ReDim ImmovableAssetAddress9(count_AL1)
'
'    For i = 1 To count_AL1
'        ImmovableAssetAddress1(i) = rangecells.item(i).value
'        ImmovableAssetAddress2(i) = rangecells1.item(i).value
'        ImmovableAssetAddress3(i) = rangecells2.item(i).value
'        ImmovableAssetAddress4(i) = rangecells3.item(i).value
'        ImmovableAssetAddress5(i) = rangecells4.item(i).value
'        ImmovableAssetAddress6(i) = rangecells5.item(i).value
'        ImmovableAssetAddress7(i) = rangecells6.item(i).value
'        ImmovableAssetAddress8(i) = rangecells7.item(i).value
'        ImmovableAssetAddress9(i) = rangecells8.item(i).value
'
'        'Flat
'        If ImmovableAssetAddress1(i) = "" Then
'            msgboxAL = msgboxAL + "Flat/Door/Block No. Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        If Len(ImmovableAssetAddress1(i)) > 50 Then
'            msgboxAL = msgboxAL + "Flat/Door/Block No. Address of immovable asset at Sr.No " & i & "  cannot exceed 50 characters in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        'Villages
'        If ImmovableAssetAddress2(i) = "" Then
'            'msgboxAL = msgboxAL + "Name of Premises/Building/Village Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            'ValidateImmovableAssetAddress_AL = False
'            'Exit Function
'        End If
'
'        If Len(ImmovableAssetAddress2(i)) > 50 Then
'            msgboxAL = msgboxAL + "Name of Premises/Building/Village Address of immovable asset at Sr.No " & i & "  cannot exceed 50 characters in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        'Road
'        If ImmovableAssetAddress3(i) = "" Then
'            'msgboxAL = msgboxAL + "Road/Street/Post Office Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            'ValidateImmovableAssetAddress_AL = False
'            'Exit Function
'        End If
'
'        If Len(ImmovableAssetAddress3(i)) > 50 Then
'            msgboxAL = msgboxAL + "Road/Street/Post Office Address of immovable asset at Sr.No " & i & "  cannot exceed 50 characters in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'
'        'Area
'        If ImmovableAssetAddress4(i) = "" Then
'            msgboxAL = msgboxAL + "Area/locality Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        If Len(ImmovableAssetAddress4(i)) > 50 Then
'            msgboxAL = msgboxAL + "Area/locality Address of immovable asset at Sr.No " & i & " cannot exceed 50 characters  in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        'Town
'        If ImmovableAssetAddress5(i) = "" Then
'            msgboxAL = msgboxAL + "Town/City/District Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        If Len(ImmovableAssetAddress5(i)) > 50 Then
'            msgboxAL = msgboxAL + "Town/City/District Address of immovable asset at Sr.No " & i & " cannot exceed 50 characters in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        'State
'        If ImmovableAssetAddress6(i) = "" Or UCase(ImmovableAssetAddress6(i)) = "(SELECT)" Then
'            msgboxAL = msgboxAL + "State Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        'Country
'        If ImmovableAssetAddress7(i) = "" Or UCase(ImmovableAssetAddress7(i)) = "(SELECT)" Then
'            msgboxAL = msgboxAL + "Country Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        'Pincode
'        If Mid(ImmovableAssetAddress6(i), 1, 2) <> "99" And UCase(ImmovableAssetAddress6(i)) <> "(SELECT)" Then
'        If ImmovableAssetAddress8(i) = "" Then
'            msgboxAL = msgboxAL + "Pin Code Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'        End If
'
'        If Len(ImmovableAssetAddress8(i)) > 6 Then
'            msgboxAL = msgboxAL + "PinCode Address of immovable asset  at Sr.No " & i & "  cannot exceed 6 digits in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'        If Not IsNumeric(ImmovableAssetAddress8(i)) Then
'            msgboxAL = msgboxAL + "PinCode Address of immovable asset at Sr.No " & i & "  is Invalid  in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'
'        'Zipcode
'
'        If Mid(ImmovableAssetAddress6(i), 1, 2) = "99" And UCase(ImmovableAssetAddress6(i)) <> "(SELECT)" Then
'        If ImmovableAssetAddress9(i) = "" Then
'            msgboxAL = msgboxAL + "Zip Code Address of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'        End If
'
'        If Len(ImmovableAssetAddress9(i)) > 8 Then
'            msgboxAL = msgboxAL + "ZipCode Address of immovable asset  at Sr.No " & i & "  cannot exceed 6 digits in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'
'
'        'State & Country Check
'        Dim StateCode As Variant
'        Dim CountryCode As Variant
'
'        StateCode = UCase(Mid(Trim(ImmovableAssetAddress6(i)), 1, InStr(1, Trim(ImmovableAssetAddress6(i)), "-") - 1))
'        CountryCode = UCase(Mid(Trim(ImmovableAssetAddress7(i)), 1, InStr(1, Trim(ImmovableAssetAddress7(i)), "-") - 1))
'
'        If (StateCode <> "99" And CountryCode <> "91") Then
'            msgboxAL = msgboxAL + "Country and State mapping of Address for immovable asset at Sr.No " & i & "  is Invalid  in Sheet AL" & Chr(13)
'            ValidateImmovableAssetAddress_AL = False
'            Exit Function
'        End If
'
'    Next
'End Function
'
'
'Function ValidateImmovableAssetAmount_AL() As Boolean
'ValidateImmovableAssetAmount_AL = True
'    Dim rangecells, rangecells_code, rangecells2 As Range
'    Dim a As Long
'    Dim i, j As Long
'
'    setTableInfo_AL1
'
'    Set rangecells = Sheet13.Range("SchAL.A.Amount").Cells
'    ReDim ImmovableAssetAmount(count_AL1)
'
'    For i = 1 To count_AL1
'        ImmovableAssetAmount(i) = rangecells.item(i).value
'            If ImmovableAssetAmount(i) = "" Then
'                msgboxAL = msgboxAL + "Amount(cost) of immovable asset is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateImmovableAssetAmount_AL = False
'                Exit Function
'            End If
'
'        If Trim(ImmovableAssetAmount(i)) <> "" Then
'            For j = 1 To Len(ImmovableAssetAmount(i))
'                If Not IsNumeric(Mid(ImmovableAssetAmount(i), j, 1)) Then
'                    msgboxAL = msgboxAL + "Amount(cost) of immovable asset at Sr.No " & i & " in AL must contain only digits from 0 to 9" & Chr(13)
'                    ValidateImmovableAssetAmount_AL = False
'                    Exit Function
'                End If
'            Next
'        End If
'    Next
'End Function
'
'
'
'Sub setTableInfo_AL2_1()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Name").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Name").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_1 = ccount
'    range_AL2_1 = "SchAL.C.Name;SchAL.A.Address_Flat;SchAL.A.Address_Village;SchAL.A.Address_Road;SchAL.A.Address_Area;SchAL.A.Address_City;SchAL.A.Address_State;SchAL.A.Address_Country;SchAL.A.Address_Pin;SchAL.C.PAN;SchAL.C.Investment;"
'End Sub
'
'Sub setTableInfo_AL2_2()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_Flat").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_Flat").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_2 = ccount
'End Sub
'
'
'Sub setTableInfo_AL2_3()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.PAN").Count
'    Set rangecells = Sheet13.Range("SchAL.C.PAN").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_3 = ccount
'End Sub
'
'Sub setTableInfo_AL2_4()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Investment").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Investment").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_4 = ccount
'End Sub
'
'
'Sub setTableInfo_AL2_5()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_Area").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_Area").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_5 = ccount
'End Sub
'
'Sub setTableInfo_AL2_6()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_City").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_City").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_6 = ccount
'End Sub
'
'Sub setTableInfo_AL2_7()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_Pin").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_Pin").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_7 = ccount
'End Sub
'
'
'Sub setTableInfo_AL2_8()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_State").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_State").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_8 = ccount
'End Sub
'
'
'Sub setTableInfo_AL2_9()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_Country").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_Country").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_9 = ccount
'End Sub
'
'Sub setTableInfo_AL2_10()
'
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'    mIntCells = Sheet13.Range("SchAL.C.Address_Zip").Count
'    Set rangecells = Sheet13.Range("SchAL.C.Address_Zip").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'
'    count_AL2_10 = ccount
'End Sub
'
'
'Function ValidateInterestHeldNameOfFirm_AL() As Boolean
'ValidateInterestHeldNameOfFirm_AL = True
'    Dim rangecells, rangecells_code, rangecells2 As Range
'    Dim a As Long
'    Dim i As Long
'
'    setTableInfo_AL2_1
'
'    Set rangecells = Sheet13.Range("SchAL.C.Name").Cells
'    ReDim InterestHeldNameOfFirm(count_AL2_1)
'
'    For i = 1 To count_AL2_1
'        InterestHeldNameOfFirm(i) = rangecells.item(i).value
'
'            If InterestHeldNameOfFirm(i) = "" Then
'                msgboxAL = msgboxAL + "Name of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldNameOfFirm_AL = False
'                Exit Function
'            End If
'    Next
'End Function
'
'
'Function ValidateInterestHeldAddressOfFirm_AL() As Boolean
'ValidateInterestHeldAddressOfFirm_AL = True
'    Dim rangecells, rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8 As Range
'    Dim a As Long
'    Dim i As Long
'
'    setTableInfo_AL2_1
'
'    'Set rangecells = Sheet13.Range("SchAL.C.Address").Cells
'
'    Set rangecells = Sheet13.Range("SchAL.C.Address_Flat").Cells
'    Set rangecells1 = Sheet13.Range("SchAL.C.Address_Village").Cells
'    Set rangecells2 = Sheet13.Range("SchAL.C.Address_Road").Cells
'    Set rangecells3 = Sheet13.Range("SchAL.C.Address_Area").Cells
'    Set rangecells4 = Sheet13.Range("SchAL.C.Address_City").Cells
'    Set rangecells5 = Sheet13.Range("SchAL.C.Address_State").Cells
'    Set rangecells6 = Sheet13.Range("SchAL.C.Address_Country").Cells
'    Set rangecells7 = Sheet13.Range("SchAL.C.Address_Pin").Cells
'    Set rangecells8 = Sheet13.Range("SchAL.C.Address_Zip").Cells
'
'    ReDim InterestHeldAddressOfFirm(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm1(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm2(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm3(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm4(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm5(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm6(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm7(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm8(count_AL2_1)
'    ReDim InterestHeldAddressOfFirm9(count_AL2_1)
'
'    For i = 1 To count_AL2_1
'            InterestHeldAddressOfFirm1(i) = rangecells.item(i).value
'            InterestHeldAddressOfFirm2(i) = rangecells1.item(i).value
'            InterestHeldAddressOfFirm3(i) = rangecells2.item(i).value
'            InterestHeldAddressOfFirm4(i) = rangecells3.item(i).value
'            InterestHeldAddressOfFirm5(i) = rangecells4.item(i).value
'            InterestHeldAddressOfFirm6(i) = rangecells5.item(i).value
'            InterestHeldAddressOfFirm7(i) = rangecells6.item(i).value
'            InterestHeldAddressOfFirm8(i) = rangecells7.item(i).value
'            InterestHeldAddressOfFirm9(i) = rangecells8.item(i).value
'
'            'Flat/Door/Block
'            If InterestHeldAddressOfFirm1(i) = "" Then
'                msgboxAL = msgboxAL + "Flat/Door/Block No. Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            If Len(InterestHeldAddressOfFirm1(i)) > 50 Then
'                msgboxAL = msgboxAL + "Flat/Door/Block No. Address of firm at Sr.No " & i & "  cannot exceed 50 characters in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'Villages
'            If InterestHeldAddressOfFirm2(i) = "" Then
'                'msgboxAL = msgboxAL + "Name of Premises/Building/Village Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                'ValidateInterestHeldAddressOfFirm_AL = False
'                'Exit Function
'            End If
'
'            If Len(InterestHeldAddressOfFirm2(i)) > 50 Then
'                msgboxAL = msgboxAL + "Name of Premises/Building/Village Address of firm at Sr.No " & i & "   cannot exceed 50 characters in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'
'            'Road
'            If InterestHeldAddressOfFirm3(i) = "" Then
'                'msgboxAL = msgboxAL + "Road/Street/Post Office Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                'ValidateInterestHeldAddressOfFirm_AL = False
'                'Exit Function
'            End If
'
'            If Len(InterestHeldAddressOfFirm3(i)) > 50 Then
'                msgboxAL = msgboxAL + "Road/Street/Post Office Address of firm at Sr.No " & i & "  cannot exceed 50 characters in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'Area
'            If InterestHeldAddressOfFirm4(i) = "" Then
'                msgboxAL = msgboxAL + "Area/locality Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'
'            If Len(InterestHeldAddressOfFirm4(i)) > 50 Then
'                msgboxAL = msgboxAL + "Area/locality Address of firm at Sr.No " & i & "   cannot exceed 50 characters in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'Town
'            If InterestHeldAddressOfFirm5(i) = "" Then
'                msgboxAL = msgboxAL + "Town/City/District Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            If Len(InterestHeldAddressOfFirm5(i)) > 50 Then
'                msgboxAL = msgboxAL + "Town/City/District Address of firm at Sr.No " & i & " cannot exceed 50 characters in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'State
'            If InterestHeldAddressOfFirm6(i) = "" Or UCase(InterestHeldAddressOfFirm6(i)) = "(SELECT)" Then
'                msgboxAL = msgboxAL + "State Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'Country
'            If InterestHeldAddressOfFirm7(i) = "" Or UCase(InterestHeldAddressOfFirm7(i)) = "(SELECT)" Then
'                msgboxAL = msgboxAL + "Country Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'Pincode
'            If Mid(InterestHeldAddressOfFirm6(i), 1, 2) <> "99" And UCase(InterestHeldAddressOfFirm6(i)) <> "(SELECT)" Then
'                If InterestHeldAddressOfFirm8(i) = "" Then
'                    msgboxAL = msgboxAL + "PinCode Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                    ValidateInterestHeldAddressOfFirm_AL = False
'                    Exit Function
'                End If
'            End If
'
'            If Len(InterestHeldAddressOfFirm8(i)) > 6 Then
'                msgboxAL = msgboxAL + "PinCode Address of firm  at Sr.No " & i & "  cannot exceed 6 digits in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            If Not IsNumeric(InterestHeldAddressOfFirm8(i)) Then
'                msgboxAL = msgboxAL + "PinCode Address of firm at Sr.No " & i & "  is Invalid  in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'            'Zipcode
'            If Mid(InterestHeldAddressOfFirm6(i), 1, 2) = "99" And UCase(InterestHeldAddressOfFirm6(i)) <> "(SELECT)" Then
'                If InterestHeldAddressOfFirm9(i) = "" Then
'                    msgboxAL = msgboxAL + "ZipCode Address of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                    ValidateInterestHeldAddressOfFirm_AL = False
'                    Exit Function
'                End If
'            End If
'
'            If Len(InterestHeldAddressOfFirm9(i)) > 8 Then
'                msgboxAL = msgboxAL + "ZipCode Address of firm  at Sr.No " & i & "  cannot exceed 6 digits in Sheet AL" & Chr(13)
'                ValidateInterestHeldAddressOfFirm_AL = False
'                Exit Function
'            End If
'
'
'        'State & Country Check
'        Dim StateCode As Variant
'        Dim CountryCode As Variant
'
'        StateCode = UCase(Mid(Trim(InterestHeldAddressOfFirm6(i)), 1, InStr(1, Trim(InterestHeldAddressOfFirm6(i)), "-") - 1))
'        CountryCode = UCase(Mid(Trim(InterestHeldAddressOfFirm7(i)), 1, InStr(1, Trim(InterestHeldAddressOfFirm7(i)), "-") - 1))
'
'        If (StateCode <> "99" And CountryCode <> "91") Then
'            msgboxAL = msgboxAL + "Country and State mapping of Address for firm at Sr.No " & i & "  is Invalid  in Sheet AL" & Chr(13)
'            ValidateInterestHeldAddressOfFirm_AL = False
'            Exit Function
'        End If
'    Next
'End Function
'
'
'Function ValidateInterestHeldPANOfFirm_AL() As Boolean
'ValidateInterestHeldPANOfFirm_AL = True
'    Dim rangecells, rangecells_code, rangecells2 As Range
'    Dim a As Long
'    Dim i As Long
'
'    setTableInfo_AL2_1
'
'    Set rangecells = Sheet13.Range("SchAL.C.PAN").Cells
'    ReDim InterestHeldPANOfFirm(count_AL2_1)
'
'    For i = 1 To count_AL2_1
'        InterestHeldPANOfFirm(i) = rangecells.item(i).value
'
'            If InterestHeldPANOfFirm(i) = "" Then
'                msgboxAL = msgboxAL + "PAN of firm is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldPANOfFirm_AL = False
'                Exit Function
'            End If
'
'            If Not CheckPAN(UCase(InterestHeldPANOfFirm(i))) Then
'                msgboxAL = msgboxAL + "PAN of firm is not Valid at Sr.No " & i & " in Sheet AL. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
'                ValidateInterestHeldPANOfFirm_AL = False
'                Exit Function
'            End If
'
'            'If (Mid(UCase(InterestHeldPANOfFirm(i)), 4, 1) <> "F" And Mid(UCase(InterestHeldPANOfFirm(i)), 4, 1) <> "A" And Mid(UCase(InterestHeldPANOfFirm(i)), 4, 1) <> "B" And Mid(UCase(InterestHeldPANOfFirm(i)), 4, 1) <> "T") Then
'            '    msgboxAL = msgboxAL + "PAN of firm is not Valid at Sr.No " & i & " in Sheet AL. The fourth character of PAN should either be 'F'or 'A/B/T' only" & Chr(13)
'            '    ValidateInterestHeldPANOfFirm_AL = False
'            '    Exit Function
'            'End If
'    Next
'End Function
'
'Function ValidateInterestHeldInvestmentOfFirm_AL() As Boolean
'ValidateInterestHeldInvestmentOfFirm_AL = True
'    Dim rangecells, rangecells_code, rangecells2 As Range
'    Dim a As Long
'    Dim i, j As Long
'
'    setTableInfo_AL2_1
'
'    Set rangecells = Sheet13.Range("SchAL.C.Investment").Cells
'    ReDim InterestHeldInvestmentOfFirm(count_AL2_1)
'
'    For i = 1 To count_AL2_1
'        InterestHeldInvestmentOfFirm(i) = rangecells.item(i).value
'
'            If InterestHeldInvestmentOfFirm(i) = "" Then
'                msgboxAL = msgboxAL + "Assessee's investment is Mandatory at Sr.No " & i & " in Sheet AL" & Chr(13)
'                ValidateInterestHeldInvestmentOfFirm_AL = False
'                Exit Function
'            End If
'
'        If Trim(InterestHeldInvestmentOfFirm(i)) <> "" Then
'            For j = 1 To Len(InterestHeldInvestmentOfFirm(i))
'                If Not IsNumeric(Mid(InterestHeldInvestmentOfFirm(i), j, 1)) Then
'                    msgboxAL = msgboxAL + "Assessee's investment at Sr.No " & i & " in Sheet AL must contain only digits from 0 to 9" & Chr(13)
'                    ValidateInterestHeldInvestmentOfFirm_AL = False
'                    Exit Function
'                End If
'            Next
'        End If
'    Next
'End Function
'
'
'Sub Al_Prev_Click()
'    Sheet5.Activate
'End Sub
'
'Sub Al_Next_Click()
'    Sheet6.Activate
'End Sub
'
'
'
'
'Sub UNLOCKAL()
'    Application.ScreenUpdating = False
'    Sheet13.Unprotect getmsgstate
'
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "SchAL.A.Description||SchAL.A.Address_Flat||SchAL.A.Address_Village||SchAL.A.Address_Road||SchAL.A.Address_Area||SchAL.A.Address_City||SchAL.A.Address_State||SchAL.A.Address_Country||SchAL.A.Address_Pin||SchAL.A.Amount"
'    UNLOCKRANGEAL (rngname_5_AL)
'
'    Sheet13.Protect getmsgstate
'    Application.ScreenUpdating = True
'End Sub
'
'
'Sub UNLOCKAL1()
'    Application.ScreenUpdating = False
'    Sheet13.Unprotect getmsgstate
'
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "SchAL.A.Description||SchAL.A.Address_Flat||SchAL.A.Address_Village||SchAL.A.Address_Road||SchAL.A.Address_Area||SchAL.A.Address_City||SchAL.A.Address_State||SchAL.A.Address_Country||SchAL.A.Address_Pin||SchAL.A.Amount"
'    UNLOCKRANGEAL1 (rngname_5_AL)
'
'    Sheet13.Protect getmsgstate
'    Application.ScreenUpdating = True
'End Sub
'
'
'Sub LOCKAL()
'    Application.ScreenUpdating = False
'
'    Sheet13.Unprotect getmsgstate
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "SchAL.A.Description||SchAL.A.Address_Flat||SchAL.A.Address_Village||SchAL.A.Address_Road||SchAL.A.Address_Area||SchAL.A.Address_City||SchAL.A.Address_State||SchAL.A.Address_Country||SchAL.A.Address_Pin||SchAL.A.Address_Zip||SchAL.A.Amount"
'    LOCKRANGEAL (rngname_5_AL)
'
'   Sheet13.Protect getmsgstate
'
'    Application.ScreenUpdating = True
'End Sub
'
'
'Sub UNLOCKAL_1()
'    Application.ScreenUpdating = False
'    Sheet13.Unprotect getmsgstate
'
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "SchAL.C.Name||SchAL.C.Address_Flat||SchAL.C.Address_Village||SchAL.C.Address_Road||SchAL.C.Address_Area||SchAL.C.Address_City||SchAL.C.Address_State||SchAL.C.Address_Country||SchAL.C.Address_Pin||SchAL.C.Address_Zip||SchAL.C.PAN||SchAL.C.Investment"
'    UNLOCKRANGEAL (rngname_5_AL)
'
'    Sheet13.Protect getmsgstate
'    Application.ScreenUpdating = True
'End Sub
'
'
'
'
'Sub LOCKAL_1()
'    Application.ScreenUpdating = False
'
'    Sheet13.Unprotect getmsgstate
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "SchAL.C.Name||SchAL.C.Address_Flat||SchAL.C.Address_Village||SchAL.C.Address_Road||SchAL.C.Address_Area||SchAL.C.Address_City||SchAL.C.Address_State||SchAL.C.Address_Country||SchAL.C.Address_Pin||SchAL.C.Address_Zip||SchAL.C.PAN||SchAL.C.Investment"
'    LOCKRANGEAL (rngname_5_AL)
'
'   Sheet13.Protect getmsgstate
'
'    Application.ScreenUpdating = True
'End Sub
'
'
'
'Sub LOCKAL_2()
'    Application.ScreenUpdating = False
'
'    Sheet13.Unprotect getmsgstate
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "IsImmovableAsset||IsInterestHeld||SchAL.B.Jwellery||SchAL.B.Art||SchAL.B.Vehicles||SchAL.B.Bank||SchAL.B.Share||SchAL.B.Insurance||SchAL.B.Loans||SchAL.B.Cash||SchAL.Liability"
'    LOCKRANGEAL (rngname_5_AL)
'
'   Sheet13.Protect getmsgstate
'
'    Application.ScreenUpdating = True
'End Sub
'
'
'Sub UNLOCKAL_2()
'    Application.ScreenUpdating = False
'    Sheet13.Unprotect getmsgstate
'
'
'    Dim rngname_5_AL As Variant
'    rngname_5_AL = "IsImmovableAsset||IsInterestHeld||SchAL.B.Jwellery||SchAL.B.Art||SchAL.B.Vehicles||SchAL.B.Bank||SchAL.B.Share||SchAL.B.Insurance||SchAL.B.Loans||SchAL.B.Cash||SchAL.Liability"
'    UNLOCKRANGEAL (rngname_5_AL)
'
'    Sheet13.Protect getmsgstate
'    Application.ScreenUpdating = True
'End Sub
'
'Sub LOCKRANGEAL(rangenamestring As Variant)
'Dim i As Long
'Dim chCell As Range
'    rangenamestring = Split(rangenamestring, "||")
'        For i = 0 To UBound(rangenamestring)
'            For Each chCell In Sheet13.Range(rangenamestring(i))
'            chCell.Locked = True
'            chCell.Interior.Color = (&HD8D8D8)
'            chCell.ClearContents
'            Next
'        Next
'End Sub
'
'Sub UNLOCKRANGEAL(rangenamestring As Variant)
'Dim i As Long
'Dim chCell As Range
'    rangenamestring = Split(rangenamestring, "||")
'        For i = 0 To UBound(rangenamestring)
'            For Each chCell In Sheet13.Range(rangenamestring(i))
'            chCell.Locked = False
'            chCell.Interior.Color = (&HCCFFCC)
'            chCell.ClearContents
'            Next
'        Next
'End Sub
'
'Sub UNLOCKRANGEAL1(rangenamestring As Variant)
'Dim i As Long
'Dim chCell As Range
'    rangenamestring = Split(rangenamestring, "||")
'        For i = 0 To UBound(rangenamestring)
'            For Each chCell In Sheet13.Range(rangenamestring(i))
'            chCell.Locked = False
'            chCell.Interior.Color = (&HCCFFCC)
'            'chCell.ClearContents
'            Next
'        Next
'End Sub
'
'
