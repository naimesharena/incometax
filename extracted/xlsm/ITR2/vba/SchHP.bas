Attribute VB_Name = "SchHP"
Option Explicit
Dim sPassword As Variant

Public rngname_hpco As Variant
Public rngname_hpco2 As Variant
Public end_hpco As Variant
Public end_hpco2 As Variant

'--Bindu
Public end_hpco_24b As Variant
Public rngname_hpco_24b As Variant
'Shrutika(SIT-93666)03/06/25
Public end_hpco_24c As Variant
Public end_hpco_24e As Variant
Public end_hpco_24f As Variant
Public end_hpco_24k As Variant
Public end_hpco_24s As Variant
Public end_hpco_24r As Variant
'-----------------------------

Public AddrDetail_HP As Variant
Public CityOrTownOrDistrict_HP As Variant
Public CountryCode_HP As Variant
Public OwnerProperty_HP As Variant
Public OwnerPropertyDescription_HP As Variant
Public StateCode_HP As Variant
Public ZipCode_HP As Variant
Public PinCode_HP As Variant
Public ifLetOut_HP As Variant
Public NameofTenant_HP As Variant
Public PANofTenant_HP As Variant
Public TANofTenant_HP As Variant
Public CoName_HP As Variant
Public CoPAN_HP As Variant
Public CoAadhar_HP As Variant
Public CoShare_HP As Variant

Public AnnualLetableValue_HP As Variant
Public RentNotRealized_HP As Variant
Public LocalTaxes_HP As Variant
Public TotalUnrealizedAndTax_HP As Variant
Public BalanceALV_HP As Variant
Public ThirtyPercentOfBalance_HP As Variant
Public IntOnBorwCap_HP As Variant
Public TotalDeduct_HP As Variant
Public Arrears_HP As Variant
Public IncomeOfHP_HP As Variant
Public IncomeOfHPInOwnHand_HP As Variant
Public RentOfEarlierYrSec25AandAA_HP As Variant
Public RentArearsSec25BAfter30pcDeduct_HP As Variant
Public TotalIncomeChargeableUnHP_HP As Variant

Public CoOwnerAadhar_HP As Variant
Public AadharofTenant_HP As Variant


Public CoOwnerName1_HP As Variant
Public CoOwnerName2_HP As Variant
Public CoOwnerName3_HP As Variant
Public CoOwnerName4_HP As Variant
Public CoOwnerName5_HP As Variant

Public CoOwnerPAN1_HP As Variant
Public CoOwnerPAN2_HP As Variant
Public CoOwnerPAN3_HP As Variant
Public CoOwnerPAN4_HP As Variant
Public CoOwnerPAN5_HP As Variant

Public CoOwnerSharePer1_HP As Variant
Public CoOwnerSharePer2_HP As Variant
Public CoOwnerSharePer3_HP As Variant
Public CoOwnerSharePer4_HP As Variant
Public CoOwnerSharePer5_HP As Variant

Public CoOwnedYN_HP As Variant
Public CoOwnedShare_HP As Variant

Public frmsize_hprptfrm As Variant
Public end_hprptfrm As Variant
Public cntrRng_hprptfrm As Variant
Public frmRngname_hprptfrm As Variant
Public rngname_hprptfrm As Variant

Dim msgValidateSheethprptfrm As String
Dim msgValidateSheetHP As String

'added by Chetan C M for AY 2025-26
'Public MsgBox_24b As Variant
Public end_24b As Variant
Public end_24bLoanfrm As Variant
'Public end_24bIFSC As Variant
Public end_24bbankName As Variant
Public end_24bPAN As Variant
Public end_24bAccntNum As Variant
Public end_24bLoanDate As Variant
Public end_24bLoanAmt As Variant
Public end_24bLoanOutstanding As Variant
Public end_24bIntrst As Variant

Public BankorInst_24b As Variant
Public BankName_24B As Variant
Public AccntNum_24b As Variant
Public LoanDate_24B As Variant
Public LoanAmt_24b As Variant
Public LoanOutStanding_24B As Variant
Public Intrst_24B As Variant


                



Function GetIncomeOfHP() As Variant
    If ValidateIncomeOfHP_HP Then
        GetIncomeOfHP = IncomeOfHP_HP
    End If
End Function

Sub setTblinfo_hprptfrm()
    Dim te As Long
    Dim ccount As Long
    
    frmsize_hprptfrm = Sheet8.Range("NumRowsPropertyBlock").value
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value
        If Not Sheet8.Range("HP.AddrDetail" & te).value = "" Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    end_hprptfrm = ccount
    cntrRng_hprptfrm = "PropertySectionCOunt"
    frmRngname_hprptfrm = "hprptfrm"
    
'changed by Chetan C M for AY 2025-26
'start--
'    rngname_hprptfrm = "HP.AddrDetail1;HP.CityOrTownOrDistrict1;HP.CountryCode1;HP.StateCode1;HP.ZipCode1;HP.PinCode1;HP.OwnerProperty1;HP.OwnerPropertyDescription1;HP.CoOwnedYN1;HP.SharePercent1;HP.Co.Srno1;HP.Co.Name1;HP.Co.PAN1;HP.Co.Aadhaar1;HP.Co.Share1;HP.ifLetOut1;HP.NameofTenant1;HP.PANofTenant1;HP.AadharofTenant1;HP.TANofTenant1;HP.AnnualLetableValue1;HP.RentNotRealized1;HP.LocalTaxes1;HP.TotalUnrealizedAndTax1;HP.BalanceALV1;HP.ThirtyPercentOfBalance1;HP.IntOnBorwCap1;HP.TotalDeduct1;HP.RentOfEarlierYrSec_AandAA1;HP.IncomeOfHP1;HP.IncomeOfHPInOwnHand1;Co_OwnnerBlock1;"
'rngname_hprptfrm = "HP.AddrDetail1;HP.CityOrTownOrDistrict1;HP.CountryCode1;HP.StateCode1;HP.ZipCode1;HP.PinCode1;HP.OwnerProperty1;HP.OwnerPropertyDescription1;HP.CoOwnedYN1;HP.SharePercent1;HP.Co.Srno1;HP.Co.Name1;HP.Co.PAN1;HP.Co.Aadhaar1;HP.Co.Share1;HP.ifLetOut1;HP.NameofTenant1;HP.PANofTenant1;HP.AadharofTenant1;HP.TANofTenant1;HP.AnnualLetableValue1;HP.RentNotRealized1;HP.LocalTaxes1;HP.TotalUnrealizedAndTax1;HP.BalanceALV1;HP.ThirtyPercentOfBalance1;HP.IntOnBorwCap1;HP.TotalDeduct1;HP.RentOfEarlierYrSec_AandAA1;HP.IncomeOfHP1;HP.IncomeOfHPInOwnHand1;Co_OwnnerBlock1;LoanfrmBankOrInstitute.24b1;bankName.24b1;loanAccNum.24b1;loanDate.24b1;loanAmt.24b1;loanOutstanding.24b1;Intrst.24b1;TotAmt.24b1"
rngname_hprptfrm = "HP.AddrDetail1;HP.CityOrTownOrDistrict1;HP.CountryCode1;HP.StateCode1;HP.ZipCode1;HP.PinCode1;HP.OwnerProperty1;HP.OwnerPropertyDescription1;HP.CoOwnedYN1;HP.SharePercent1;HP.Co.Srno1;HP.Co.Name1;HP.Co.PAN1;HP.Co.Aadhaar1;HP.Co.Share1;HP.ifLetOut1;HP.NameofTenant1;HP.PANofTenant1;HP.AadharofTenant1;HP.TANofTenant1;HP.AnnualLetableValue1;HP.RentNotRealized1;HP.LocalTaxes1;HP.TotalUnrealizedAndTax1;HP.BalanceALV1;HP.ThirtyPercentOfBalance1;HP.IntOnBorwCap1;HP.TotalDeduct1;HP.RentOfEarlierYrSec_AandAA1;HP.IncomeOfHP1;HP.IncomeOfHPInOwnHand1;Co_OwnnerBlock1;LoanfrmBankOrInstitute.24b1;bankName.24b1;loanAccNum.24b1;loanDate.24b1;loanAmt.24b1;loanOutstanding.24b1;Intrst.24b1;TotAmt.24b1;Combination_24B1"


'--end

 End Sub

Function msgbox_hprptfrm(strmsg As String) As String
     msgValidateSheetHP = msgValidateSheetHP & strmsg & Chr(13)
End Function

Function msgbox_HP(strmsg As String) As String
     msgValidateSheetHP = msgValidateSheetHP & strmsg & Chr(13)
End Function

Sub ValidateSheetHPClick()
ValidateSheetHouseProperty
If msgValidateSheetHP = "" Then fmsgboxStatus "Sheet House Property (HP) is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheetHouseProperty()
subProcCaption = "Validating HP"
    If Not ValidatesheetHP Then
        Sheet8.Activate
        fmsgbox msgValidateSheetHP ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub

Function ValidatesheetHP() As Boolean
    ValidatesheetHP = True
    
    AddTotIncomeUnderHouseProperty_1
    If Not ValidateAddrDetail_HP() Then ValidatesheetHP = False
    If Not ValidateCoName_HP() Then ValidatesheetHP = False
    'If Not Validate_24b() Then ValidatesheetHP = False 'added by Chetan C M for 2025-26
    If Not Validate24B_Table() Then ValidatesheetHP = False
    
    
    Dim te As Long
    Dim de As Long
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value
        If Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & te).value > 0 Then
          If Sheet8.Range("HP.AddrDetail" & te).value = "" Then
          'msgbox_hprptfrm ("* Address Details" & te & " in Sheet HOUSE_PROPERTY  is mandatory")
          
          'Shrutika
          'msgbox_hprptfrm ("* Address of property is mandatory  in schedule HP.")
          msgbox_hprptfrm ("* ""Address of property is mandatory  in schedule HP.""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
    Next
    
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value
        If Sheet8.Range("HP.TotalUnrealizedAndTax" & te).value > 0 Or Sheet8.Range("HP.BalanceALV" & te).value > 0 Or Sheet8.Range("HP.IncomeOfHPInOwnHand" & te).value > 0 Or Sheet8.Range("HP.TotalDeduct" & te).value > 0 Or Sheet8.Range("HP.IncomeOfHP" & te).value > 0 Then
          If Sheet8.Range("HP.AddrDetail" & te).value = "" Then
          'msgbox_hprptfrm ("* Address Details" & te & " in Sheet HOUSE_PROPERTY  is mandatory")
          'Shrutika
          'msgbox_hprptfrm ("*Address of property is mandatory  in schedule HP.")
          'Malli uncomented 23/10/2024
          msgbox_hprptfrm ("* ""Address of property is mandatory  in schedule HP.""")
          
            ValidatesheetHP = False
            Exit Function
        End If
        End If
        
       
        
    Next
    
    'Added by Shrutika(15/09/2025)Tc's
     For de = 1 To Sheet8.Range("PropertySectionCOunt").value
       
        If Sheet8.Range("HP.TotalUnrealizedAndTax" & de) > 99999999999999# Then
            msgbox_hprptfrm "* Total (1b + 1c) in schedule House Property cannot exceed 14 digits" & Chr(13)
            ValidatesheetHP = False
            Exit Function
        End If
        Next
        '------------------------------
        
    'Added by Shrutika---------------
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value
    If Not isdropdownblank(Sheet8.Range("HP.StateCode" & te).value) _
    Or Not isdropdownblank(Sheet8.Range("HP.CountryCode" & te).value) _
    Or Not isdropdownblank(Sheet8.Range("HP.OwnerProperty" & te).value) _
    Or Not isdropdownblank(Sheet8.Range("HP.CoOwnedYN" & te).value) _
    Or Sheet8.Range("HP.AddrDetail" & te).value <> "" Then

         If Sheet8.Range("HP.AddrDetail" & te).value = "" Then
            msgbox_hprptfrm ("* ""Address of property is mandatory  in schedule HP.""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
    Next
'Added by Shrutika -----------------------
    If (Len(Sheet8.Range("HP.AddrDetail1")) > 0) Then
        If Not ValidateCityOrTownOrDistrict_HP() Then ValidatesheetHP = False
        If Not ValidateCountryCode_HP() Then ValidatesheetHP = False
        If Not ValidateStateCode_HP() Then ValidatesheetHP = False
        If Not ValidateZipCode_HP() Then ValidatesheetHP = False
        If Not ValidatePinCode_HP() Then ValidatesheetHP = False
        If Not ValidateOwnerProperty_HP() Then ValidatesheetHP = False
        If Not ValidateCountryStateCode_HP() Then ValidatesheetHP = False
        If Not ValidateifLetOut_HP() Then ValidatesheetHP = False
        If Not ValidateNameofTenant_HP() Then ValidatesheetHP = False
        If Not ValidatePANofTenant_HP() Then ValidatesheetHP = False
        If Not ValidateAadharofTenant_HP() Then ValidatesheetHP = False
        If Not ValidateTANofTenant_HP() Then ValidatesheetHP = False
        If Not ValidateAnnualLetableValue_HP() Then ValidatesheetHP = False
        If Not ValidateRentNotRealized_HP() Then ValidatesheetHP = False
        If Not ValidateLocalTaxes_HP() Then ValidatesheetHP = False
        If Not ValidateTotalUnrealizedAndTax_HP() Then ValidatesheetHP = False
        If Not ValidateBalanceALV_HP() Then ValidatesheetHP = False
        If Not ValidateThirtyPercentOfBalance_HP() Then ValidatesheetHP = False
        If Not ValidateIntOnBorwCap_HP() Then ValidatesheetHP = False
        If Not ValidateTotalDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateArrears_HP() Then ValidatesheetHP = False
        If Not ValidateIncomeOfHP_HP() Then ValidatesheetHP = False
        'If Not ValidateRentOfEarlierYrSec25AandAA_HP() Then ValidatesheetHP = False
        If Not ValidateRentArearsSec25BAfter30pcDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateTotalIncomeChargeableUnHP_HP() Then ValidatesheetHP = False
        
        If Not ValidateCoownerName1_HP() Then ValidatesheetHP = False

        If Not ValidatePercenShare1_HP() Then ValidatesheetHP = False

        If Not ValidateCoOwnerPAN1_HP() Then ValidatesheetHP = False

        
        If Not ValidateCoOwnerAadhar_HP() Then ValidatesheetHP = False
        
        If Not ValidateSharePercent_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnedYN_HP() Then ValidatesheetHP = False
        If Not ValidateCoownerRules_HP() Then ValidatesheetHP = False
        '22/08/2024 Chetan C M changes start
        If Not ValidatesharesRules_HP() Then ValidatesheetHP = False
        '22/08/2024 Chetan C M changes end
    
        If ValidatesheetHP Then
            Dim noofsop, temp1, hpi, SPI, ind As Long
            Dim letoutarr As Variant
            
            letoutarr = GetLetOut
            noofsop = 0
            
            For hpi = 1 To UBound(letoutarr)
                If Mid(Sheet8.Range("HP.ifLetOut" & hpi).value, 1, 1) = "S" Then
                 noofsop = noofsop + 1
                    If (Sheet8.Range("HP.AnnualLetableValue" & hpi).value > 0) Then
                    '    Sheet8.Range("HP.ifLetOut" & hpi).value = "Yes"
                    Sheet8.Activate
                        fmsgboxStatus "Property " & hpi & " is deemed to be letout as its A.L.V. > 0" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                    
                     If Mid(Sheet8.Range("HP.ifLetOut" & hpi).value, 1, 1) = "S" Then
                         'Temp = Temp + 1
                          If noofsop > 2 Then
                               msgbox_hprptfrm "* Self occupied property cannot be selected more than twice from house property is declared in Schedule HP." & Chr(13)
                                ValidatesheetHP = False
                          End If
                          temp1 = Sheet8.Range("HP.IntOnBorwCap" & hpi) + temp1
                          If temp1 > 200000 Then
                            msgbox_hprptfrm "* F for category self-occupied in Schedule HP." & Chr(13)
                                ValidatesheetHP = False
                          End If
                        
                        End If
                    
                    
                    
                Else
                    If letoutarr(hpi) = "" Then
                        If (Sheet8.Range("HP.AnnualLetableValue" & hpi).value = 0) Then
                           ' Sheet8.Range("HP.ifLetOut" & hpi).value = "No"
                            noofsop = noofsop + 1
                        Else
                           ' Sheet8.Range("HP.ifLetOut" & hpi).value = "Yes"
                        End If
                    Else
                       ' If letoutarr(hpi) = "S" Then
                       '
                       ' End If
                    End If
                End If
                
'                If noofsop > 2 Then
'                    noofsop = 2
'                    'Sheet8.Range("HP.ifLetOut" & hpi).value = "Y"
'                    Sheet8.Activate
'                    fmsgboxStatus "Please check the selection in Type of HP since there cannot be more than two self occupied Property" ', vbOKOnly, "Error!"
'                    CloseMsg
'                End If
            Next
    
            letoutarr = GetLetOut
            For SPI = 1 To UBound(letoutarr)
                If (letoutarr(SPI) = "S" Or letoutarr(SPI) = "" Or _
                    Sheet8.Range("HP.AnnualLetableValue" & SPI).value = 0) Then
                    ind = SPI
                    Dim count As Variant
                    Dim tempval As Range
                    count = 0
                If letoutarr(SPI) = "S" Then
                  For Each tempval In Sheet8.Range("HP.NameofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                        'Range(tempval.address).value = ""
                    End If
                  Next
                  If count > 0 Then
                        msgbox_hprptfrm ("* If Self Occupied then Name of Tenant in Sheet HOUSE_PROPERTY at Block " & ind & " is not required")
                        ValidatesheetHP = False
                    End If
                    
                  count = 0
                  For Each tempval In Sheet8.Range("HP.PANofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                       ' Range(tempval.address).value = ""
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then PAN of Tenant in Sheet HOUSE_PROPERTY at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        
                         count = 0
                  For Each tempval In Sheet8.Range("HP.AadharofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                       ' Range(tempval.address).value = ""
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then Aadhaar of Tenant in Sheet HOUSE_PROPERTY at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        
                        
                    count = 0
                  For Each tempval In Sheet8.Range("HP.TANofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                       ' Range(tempval.address).value = ""
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then TAN of Tenant in Sheet HOUSE_PROPERTY at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        count = 0
                
                    End If
                    If val(Sheet8.Range("HP.IntOnBorwCap" & ind)) > 200000 Then
                        If Mid(Sheet8.Range("HP.ifLetOut" & ind), 1, 1) = "D" Or Mid(Sheet8.Range("HP.ifLetOut" & ind), 1, 1) = "L" Then
                        
                        Else
                            Sheet8.Activate
                            fmsgboxStatus "If Not Letout (OR ALV=0) then Interest Payable on borrowed capital should not exceed 2 Lacs" ' , vbOKOnly, "Error!"
                            CloseMsg
                        End If
                    End If
                
                    'Sheet8.Range("HP.IntOnBorwCap" & ind).value = WorksheetFunction.Min(200000, Sheet8.Range("HP.IntOnBorwCap" & ind).value)
                    If Sheet8.Range("HP.LocalTaxes" & ind).value > 0 Then
                    Sheet8.Activate
                        fmsgboxStatus "Tax paid to local authorities can be claimed only if income from house property is declared" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                    'Sheet8.Range("HP.LocalTaxes" & ind).value = 0
                
                    If val(Sheet8.Range("HP.RentNotRealized" & ind)) > 0 Then
                    Sheet8.Activate
                        fmsgboxStatus "Rent not realized cannot exceed Gross rent received or receivable or letable value in schedule HP" ', vbOKOnly, "Error!"
                        
                        CloseMsg
                    End If
'                    Sheet8.Range("HP.RentNotRealized" & ind).value = 0
                End If
                            
                If (letoutarr(SPI) <> "S") Then
                    ind = SPI
                    If val(Sheet8.Range("HP.RentNotRealized" & ind)) > Sheet8.Range("HP.AnnualLetableValue" & ind).value Then
                    Sheet8.Activate
                        fmsgboxStatus "Rent not realized cannot exceed Gross Rent received or receivable or letable value in schedule HP" ', vbOKOnly, "Error!"
                        CloseMsg 'Sheet8.Range("HP.RentNotRealized" & ind).value = Sheet8.Range("HP.AnnualLetableValue" & ind).value
                    End If
                    If letoutarr(SPI) = "L" Then
'                        If Sheet8.Range("HP.AnnualLetableValue" & ind) = "" Then
'                        Sheet8.Activate
'                            MsgBox "Please enter Gross Rent received or receivable or letable value  in Property Block " & SPI, vbOKOnly, "Error!"
'                            CloseMsg
'                        End If
                        If Sheet8.Range("HP.AnnualLetableValue" & ind) <= 0 Then
                        Sheet8.Activate
                        'Commented by Shrutika
                            'fmsgboxStatus "Since the property is let out, field 1a should be greater than zero for Property " & SPI ', vbOKOnly, "Error!"
                            'Added by Shrutika
                            fmsgboxStatus "If Type of House Property is Let Out/ Deemed Let Out then Gross Rent received or receivable or lettable value field cannot be zero/ blank"
                            
                            CloseMsg
                        End If
                    End If
                    
                    If letoutarr(ind) = "D" Then
                        For Each tempval In Sheet8.Range("HP.NameofTenant" & ind).Cells
                            If Not tempval.value = "" Then
                            count = count + 1
                            'Range(tempval.address).value = ""
                            End If
                        Next
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then Name of Tenant in Sheet HOUSE_PROPERTY at Block " & ind & " is not required")
                            ValidatesheetHP = False
                            End If
                        
                        count = 0
                        For Each tempval In Sheet8.Range("HP.PANofTenant" & ind).Cells
                            If Not tempval.value = "" Then
                            count = count + 1
                            ' Range(tempval.address).value = ""
                            End If
                        Next
                        
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then PAN of Tenant in Sheet HOUSE_PROPERTY at Block " & ind & " is not required")
                            ValidatesheetHP = False
                        End If
                        count = 0
                        For Each tempval In Sheet8.Range("HP.TANofTenant" & ind).Cells
                            If Not tempval.value = "" Then
                            count = count + 1
                            ' Range(tempval.address).value = ""
                            End If
                        Next
                        
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then TAN of Tenant in Sheet HOUSE_PROPERTY  at Block " & ind & " is not required")
                            ValidatesheetHP = False
                        End If
                        count = 0
                        End If
                    End If
            Next
            
        End If
    Else
        'If Not ValidateRentOfEarlierYrSec25AandAA_HP() Then ValidatesheetHP = False
        If Not ValidateRentArearsSec25BAfter30pcDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateTotalIncomeChargeableUnHP_HP() Then ValidatesheetHP = False
    End If

End Function

Function ValidateAddrDetail_HP() As Boolean
    Dim i As Long
    ValidateAddrDetail_HP = True
    setTblinfo_hprptfrm
    
    ReDim AddrDetail_HP(end_hprptfrm)
    noOfProcessSub = end_hprptfrm
    For i = 1 To end_hprptfrm
        AddrDetail_HP(i) = Sheet8.Range("HP.AddrDetail" & i).value

        If Not checkfieldspecialcharacter(AddrDetail_HP(i)) Then
            msgbox_hprptfrm ("* AddrDetail" & i & "   characters < > & ' " & Chr(34) & " are not allowed Sheet HOUSE_PROPERTY ")
            ValidateAddrDetail_HP = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateCoName_HP() As Boolean
    Dim i As Long
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    
    Dim state_1 As Variant
    Dim sharepercent As Variant
    Dim cosharepercent As Double
    
    ValidateCoName_HP = True
    

    sharepercent = 0
    cosharepercent = 0
    setTblinfo_hprptfrm
    For i = 1 To end_hprptfrm
    'If UCase(Mid(Sheet8.Range("HP.CoOwnedYN" & i).value, 1, 1)) <> "Y" Then
    'GoTo endlin1
    'Else
        sharepercent = Sheet8.Range("HP.SharePercent" & i).text
        Set rangecells1 = Sheet8.Range("HP.Co.Name" & i).Cells
        Set rangecells2 = Sheet8.Range("HP.Co.PAN" & i).Cells
        Set rangecells3 = Sheet8.Range("HP.Co.Share" & i).Cells
        Set rangecells4 = Sheet8.Range("HP.StateCode" & i).Cells
        Set rangecells5 = Sheet8.Range("HP.Co.Aadhaar" & i).Cells
        
        
        setTblinfo_hpcoindex (i)
        'setTblinfo_24bankname (i)   'added by Chetan C M for AY 2025-26
        
        ReDim CoName_HP(end_hprptfrm, end_hpco)
        ReDim CoPAN_HP(end_hprptfrm, end_hpco)
        ReDim CoAadhar_HP(end_hprptfrm, end_hpco)
        ReDim CoShare_HP(end_hprptfrm, end_hpco)
        'ReDim state_1(end_hprptfrm)
        
        state_1 = rangecells4.value
        state_1 = Mid(state_1, 1, 2)
        
        cosharepercent = 0
        For j = 1 To end_hpco
            CoName_HP(i, j) = rangecells1.item(j).value
            CoPAN_HP(i, j) = rangecells2.item(j).value
            CoAadhar_HP(i, j) = rangecells5.item(j).value
            CoShare_HP(i, j) = rangecells3.item(j).text
            'new change
            If CoShare_HP(i, j) <> "" Then
                CoShare_HP(i, j) = Round(CoShare_HP(i, j), 2)
            
                cosharepercent = cosharepercent + IIf(CoShare_HP(i, j) = "", 0, CoShare_HP(i, j))
            End If
            'new change
            
            '22/08/2024 Chetan C M changes start
'            If Not chkCompulsory(CoShare_HP(i, j)) Then
'                msgbox_hprptfrm ("* Share of CO Owner " & j & "  in Property " & i & " is Mandatory Sheet HOUSE_PROPERTY ")
'                ValidateCoName_HP = False
'                Exit Function
'            End If
             '22/08/2024 Chetan C M changes end
             
            ' Updated by Shrutika(08/09/2025)
'            If Not checkfieldspecialcharacter(CoName_HP(i, j)) Then
'                msgbox_hprptfrm ("* Name of CO Owner " & j & "  in Property " & i & "  cannot contain special characters Sheet HOUSE_PROPERTY ")
'                ValidateCoName_HP = False
'                Exit Function
'            End If
            '----------------------------------------------------
            If Not ValidatePAN(CStr(CoPAN_HP(i, j))) Then
            'changed by Chetan C M on 13/08/2025
            'start--
                'msgbox_hprptfrm ("* Enter Valid PAN of CO Owner " & j & "  in Property " & i & "  with 1st 5 alphabets, next 4 digits and last alphabet Sheet HOUSE_PROPERTY ")
                msgbox_hprptfrm ("* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP")
            '--end
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidateAadhaar(CStr(CoAadhar_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid Aadhaar of CO Owner " & j & "  in Property " & i & "  with 12 digits Sheet HOUSE_PROPERTY ")
                ValidateCoName_HP = False
                Exit Function
            End If
        Next
        
         If sharepercent < 0.01 Then
         'Ankita(23-Jan)
           ' msgbox_hprptfrm ("*Share Percentage of Owner and CO Owner must add to 100 in Property " & i & " Sheet HOUSE_PROPERTY")
            msgbox_hprptfrm ("* ""Your percentage of share in co-owned property cannot be zero""") & Chr(13)
            ValidateCoName_HP = False
            Exit Function
        End If
  
'        If sharepercent < 0.01 Then
'            msgbox_hprptfrm ("* Share Percentage of Owner cannot be less than .01 percent")
'            ValidateCoName_HP = False
'            Exit Function
'        End If
        
'        If (IIf(sharepercent = "", 0, sharepercent) + cosharepercent < 99.89) Or (IIf(sharepercent = "", 0, sharepercent) + cosharepercent > 100#) Then
'                msgbox_hprptfrm ("* Share Percentage of Owner and CO Owner must add to 100 in Property " & i & " Sheet ")
'                ValidateCoName_HP = False
'                Exit Function
'        End If
        'End If
    Next
endlin1:
End Function

Function ValidateCoName_HPold1() As Boolean
    Dim i As Long
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    
    Dim state_1 As Variant
    Dim sharepercent As Variant
    Dim cosharepercent As Double
    
    ValidateCoName_HP = True
    

    sharepercent = 0
    cosharepercent = 0
    setTblinfo_hprptfrm
    For i = 1 To end_hprptfrm
    'If UCase(Mid(Sheet8.Range("HP.CoOwnedYN" & i).value, 1, 1)) <> "Y" Then
    'GoTo endlin1
    'Else
        sharepercent = Sheet8.Range("HP.SharePercent" & i).text
        Set rangecells1 = Sheet8.Range("HP.Co.Name" & i).Cells
        Set rangecells2 = Sheet8.Range("HP.Co.PAN" & i).Cells
        Set rangecells3 = Sheet8.Range("HP.Co.Share" & i).Cells
        Set rangecells4 = Sheet8.Range("HP.StateCode" & i).Cells
        Set rangecells5 = Sheet8.Range("HP.Co.Aadhaar" & i).Cells
        
        
        setTblinfo_hpcoindex (i)
        setTblinfo_24bankname (i)   'added by Chetan C M for AY 2025-26
        
        ReDim CoName_HP(end_hprptfrm, end_hpco)
        ReDim CoPAN_HP(end_hprptfrm, end_hpco)
        ReDim CoAadhar_HP(end_hprptfrm, end_hpco)
        ReDim CoShare_HP(end_hprptfrm, end_hpco)
        'ReDim state_1(end_hprptfrm)
        
        state_1 = rangecells4.value
        state_1 = Mid(state_1, 1, 2)
        
        cosharepercent = 0
        For j = 1 To end_hpco
            CoName_HP(i, j) = rangecells1.item(j).value
            CoPAN_HP(i, j) = rangecells2.item(j).value
            CoAadhar_HP(i, j) = rangecells5.item(j).value
            CoShare_HP(i, j) = rangecells3.item(j).text
            
            CoShare_HP(i, j) = Round(CoShare_HP(i, j), 2)
            
            cosharepercent = cosharepercent + IIf(CoShare_HP(i, j) = "", 0, CoShare_HP(i, j))
    
            If Not chkCompulsory(CoShare_HP(i, j)) Then
                msgbox_hprptfrm ("* Share of CO Owner " & j & "  in Property " & i & " is Mandatory Sheet HOUSE_PROPERTY ")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not checkfieldspecialcharacter(CoName_HP(i, j)) Then
                msgbox_hprptfrm ("* Name of CO Owner " & j & "  in Property " & i & "  cannot contain special characters Sheet HOUSE_PROPERTY ")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidatePAN(CStr(CoPAN_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid PAN of CO Owner " & j & "  in Property " & i & "  with 1st 5 alphabets, next 4 digits and last alphabet Sheet HOUSE_PROPERTY ")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidateAadhaar(CStr(CoAadhar_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid Aadhaar of CO Owner " & j & "  in Property " & i & "  with 12 digits Sheet HOUSE_PROPERTY ")
                ValidateCoName_HP = False
                Exit Function
            End If
        Next
        
         If sharepercent < 0.01 Then
            msgbox_hprptfrm ("*Share Percentage of Owner and CO Owner must add to 100 in Property " & i & " Sheet HOUSE_PROPERTY")
            ValidateCoName_HP = False
            Exit Function
        End If
  
'        If sharepercent < 0.01 Then
'            msgbox_hprptfrm ("* Share Percentage of Owner cannot be less than .01 percent")
'            ValidateCoName_HP = False
'            Exit Function
'        End If
        
'        If (IIf(sharepercent = "", 0, sharepercent) + cosharepercent < 99.89) Or (IIf(sharepercent = "", 0, sharepercent) + cosharepercent > 100#) Then
'                msgbox_hprptfrm ("* Share Percentage of Owner and CO Owner must add to 100 in Property " & i & " Sheet ")
'                ValidateCoName_HP = False
'                Exit Function
'        End If
        'End If
    Next
endlin1:
End Function
 
Function ValidateCityOrTownOrDistrict_HP() As Boolean
    Dim i As Long
    
    ValidateCityOrTownOrDistrict_HP = True
    setTblinfo_hprptfrm
    ReDim CityOrTownOrDistrict_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        
        CityOrTownOrDistrict_HP(i) = Sheet8.Range("HP.CityOrTownOrDistrict" & i).value
        If Not chkCompulsory(CityOrTownOrDistrict_HP(i)) Then
            'msgbox_hprptfrm ("* CityOrTownOrDistrict" & i & "  in Sheet HOUSE_PROPERTY  is Mandatory")
            
            'Shrutika
            msgbox_hprptfrm ("* ""Town/City is mandatory in tab: Sl no B2""")
            
            ValidateCityOrTownOrDistrict_HP = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(CityOrTownOrDistrict_HP(i)) Then
            msgbox_hprptfrm ("* CityOrTownOrDistrict" & i & "   characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
            ValidateCityOrTownOrDistrict_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateStateCode_HP() As Boolean
    Dim i As Long
    ValidateStateCode_HP = True
    setTblinfo_hprptfrm
    ReDim StateCode_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
           
        StateCode_HP(i) = Sheet8.Range("HP.StateCode" & i).value
        StateCode_HP(i) = Mid(StateCode_HP(i), 1, 2)
        If isdropdownblank(StateCode_HP(i)) Then
            StateCode_HP(i) = ""
        End If
        
        If Not chkCompulsory(StateCode_HP(i)) Then
            'msgbox_hprptfrm ("* StateCode" & i & " is Mandatory in Sheet HOUSE_PROPERTY")
            'Shrutika
            
            'msgbox_hprptfrm ("* State of property  is mandatory in schedule HP.")
            msgbox_hprptfrm ("* ""State of property  is mandatory in schedule HP.""")
            
            ValidateStateCode_HP = False
            Exit Function
        End If
        
'    'Malli-------------------------17/09/2024
'     If StateCode_HP(i) <> "" Then
'          Dim PIN_targetadd, state_targetadd As String
'          state_targetadd = Replace(Sheet8.Range("HP.StateCode" & i).Address, "$", "")
'          PIN_targetadd = Replace(state_targetadd, "H", "J")
'        Dim ws1 As Worksheet
'        Set ws1 = Worksheets("House Property")  '23/09/24
'
'         If state_Validation(ws1, PIN_targetadd, state_targetadd) = False Then Sheet8.Range(Replace(state_targetadd, "H", "J")).value = ""
'
'    End If
'    '-----------------------------------------
    Next
End Function

Function ValidateCountryCode_HP() As Boolean
    Dim i As Long
    Dim Country1 As Variant
    ValidateCountryCode_HP = True
    setTblinfo_hprptfrm
    ReDim CountryCode_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
           
        CountryCode_HP(i) = Sheet8.Range("HP.CountryCode" & i).value
        Country1 = Sheet8.Range("HP.CountryCode" & i).value
        If isdropdownblank(CountryCode_HP(i)) Then
            CountryCode_HP(i) = ""
        End If
        
        If Not chkCompulsory(CountryCode_HP(i)) Then
        'Ankita(23-Jan)
            ' msgbox_hprptfrm ("* CountryCode" & i & "is mandatory in Sheet HOUSE_PROPERTY")
            msgbox_hprptfrm ("* ""Country of property  is mandatory in schedule HP.""") & Chr(13)
            ValidateCountryCode_HP = False
            Exit Function
        End If
        CountryCode_HP(i) = Mid(Country1, 1, WorksheetFunction.Search("-", Country1) - 1)
                
    Next
End Function


Function ValidateCountryStateCode_HP() As Boolean
    Dim i As Long
    Dim Country1, state1 As Variant
    ValidateCountryStateCode_HP = True
    setTblinfo_hprptfrm
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If

        Country1 = Sheet8.Range("HP.CountryCode" & i).value
        state1 = Sheet8.Range("HP.StateCode" & i).value
        
        If isdropdownblank(Country1) Then
            Country1 = ""
        End If
        
        If isdropdownblank(state1) Then
            state1 = ""
        End If
        
        If chkCompulsory(Country1) And chkCompulsory(state1) Then
            Country1 = Mid(Country1, 1, WorksheetFunction.Search("-", Country1) - 1)
            state1 = Mid(state1, 1, WorksheetFunction.Search("-", state1) - 1)
                        
            If (Country1 <> "91" And state1 <> "99") Then
                msgbox_hprptfrm ("* Country cannot be other than India as you have selected an Indian state " & i & "  in Sheet HOUSE_PROPERTY")
                ValidateCountryStateCode_HP = False
                Exit Function
            End If
            If (Country1 = "91" And state1 = "99") Then
                msgbox_hprptfrm ("* Country cannot be India as you have selected a Foreign state" & i & "  in Sheet HOUSE_PROPERTY")
                ValidateCountryStateCode_HP = False
                Exit Function
            End If
        End If
    Next
End Function

Function ValidateZipCode_HP() As Boolean
    Dim i As Long
    
    ValidateZipCode_HP = True
    setTblinfo_hprptfrm
    ReDim ZipCode_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        ZipCode_HP(i) = Sheet8.Range("HP.ZipCode" & i).value
        If CountryCode_HP(i) <> "91" Then
            If Not chkCompulsory(ZipCode_HP(i)) Then
                'msgbox_hprptfrm ("* ZipCode" & i & " is mandatory.If ZIP code is not available, then enter XXXXXX in Sheet HOUSE_PROPERTY")
                'Shrutika
                msgbox_hprptfrm ("* ""Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX  in schedule HP.""")
                
                ValidateZipCode_HP = False
                Exit Function
            End If
            If Not checkfieldspecialcharacter(ZipCode_HP(i)) Then
                msgbox_hprptfrm ("* ZipCode" & i & " characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
                ValidateZipCode_HP = False
                Exit Function
            End If
        Else
            ZipCode_HP(i) = ""
        End If
    Next
End Function


Function ValidatePinCode_HP() As Boolean
    Dim i As Long
    
    ValidatePinCode_HP = True
    setTblinfo_hprptfrm
    ReDim PinCode_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        PinCode_HP(i) = Sheet8.Range("HP.PinCode" & i).value
        If CountryCode_HP(i) = "91" Then
            If Not chkCompulsory(PinCode_HP(i)) Then
            'Ankita(23-Jan)
                '  msgbox_hprptfrm ("* PinCode" & i & " is Mandatory in Sheet HOUSE_PROPERTY")
               msgbox_hprptfrm ("* ""Pin code of Property is mandatory in schedule HP """) & Chr(13)
                ValidatePinCode_HP = False
                Exit Function
            End If
            If Not chkNumeric(PinCode_HP(i)) Then
                msgbox_hprptfrm ("* PinCode" & i & " only digits 0 to 9 allowed  in Sheet HOUSE_PROPERTY")
                ValidatePinCode_HP = False
                Exit Function
            End If
        Else
            PinCode_HP(i) = ""
        End If
        
        'Malli----------------17/10/2024
        
         If PinCode_HP(i) <> "" Then
         Dim PIN_targetadd1, state_targetadd1 As String
         
     
     PIN_targetadd1 = Replace(Sheet8.Range("HP.PinCode" & i).Address, "$", "")
     state_targetadd1 = Replace(PIN_targetadd1, "J", "H")
      
                             Dim ws1 As Worksheet
                             Set ws1 = Worksheets("House Property")  '23/09/24
      
        ' If state_Validation(ws1, PIN_targetadd, state_targetadd) = False Then Sheet39.Range(Replace(state_targetadd, "H", "K")).value = ""
    
         If Not PINstate_ModualValidation(ws1, PIN_targetadd1, state_targetadd1) Then
                Sheet8.Range("HP.PinCode" & i).value = ""
                msgbox_hprptfrm ("* ""Pin code of Property is mandatory in schedule HP """) & Chr(13)
                ValidatePinCode_HP = False
                Exit Function
         End If

    
    End If
        '-------------------------------
    Next
End Function

Function ValidateOwnerProperty_HP() As Boolean
    Dim i As Long

    ValidateOwnerProperty_HP = True
    setTblinfo_hprptfrm
    ReDim OwnerProperty_HP(end_hprptfrm)
    ReDim OwnerPropertyDescription_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        OwnerProperty_HP(i) = Sheet8.Range("HP.OwnerProperty" & i).value
        OwnerPropertyDescription_HP(i) = Sheet8.Range("HP.OwnerPropertyDescription" & i).value
        
            If isdropdownblank(OwnerProperty_HP(i)) Then
                'msgbox_hprptfrm ("* Owner of the Property" & i & "  in Sheet HOUSE_PROPERTY  is Mandatory")
                'Shrutika
                'msgbox_hprptfrm ("* Please select dropdown from owner of the Property  in schedule HP.")
                msgbox_hprptfrm ("* ""Please select dropdown from owner of the Property  in schedule HP.""")
                
                ValidateOwnerProperty_HP = False
                Exit Function
            End If
            
            
            
            
        If Mid(OwnerProperty_HP(i), 1, 1) = "O" Then
                
                If Not chkCompulsory(OwnerPropertyDescription_HP(i)) Then
                 msgbox_hprptfrm ("* Description for Owner of the  Property" & i & "  is Mandatory in Sheet HOUSE_PROPERTY")
                 ValidateOwnerProperty_HP = False
                 Exit Function
                End If
              
              If Len(OwnerPropertyDescription_HP(i)) > 50 Then
                 msgbox_hprptfrm ("* Description for Owner of the  Property" & i & "  Cannot be more than 50 characters in Sheet HOUSE_PROPERTY")
                 ValidateOwnerProperty_HP = False
                 Exit Function
                End If
            
             If Not checkfieldspecialcharacter(OwnerPropertyDescription_HP(i)) Then
                msgbox_hprptfrm ("* Description for Owner of the  Property" & i & "   characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
                ValidateOwnerProperty_HP = False
                Exit Function
             End If
        End If
        
    Next
End Function

Function ValidateifLetOut_HP() As Boolean
    Dim i As Long
    ValidateifLetOut_HP = True
    
    setTblinfo_hprptfrm
    ReDim ifLetOut_HP(end_hprptfrm)
    'ValidateAddrDetail_HP
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        ifLetOut_HP(i) = Sheet8.Range("HP.ifLetOut" & i).value
        ifLetOut_HP(i) = Mid(ifLetOut_HP(i), 1, 1)
        
        If isdropdownblank(ifLetOut_HP(i)) Then
            ifLetOut_HP(i) = ""
        End If
        
        If Not chkCompulsory(ifLetOut_HP(i)) Then
            msgbox_hprptfrm ("* Please select type of House Property in schedule HP ")
            ValidateifLetOut_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameofTenant_HP() As Boolean
    Dim i, count As Long
    Dim tempPan As String
    Dim rangecells As Range
    ValidateNameofTenant_HP = True
    setTblinfo_hprptfrm
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        
        Set rangecells = Sheet8.Range("HP.NameofTenant" & i).Cells
        ReDim NameofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        count = 0
        For j = 1 To rangecells.Cells.count
        
        NameofTenant_HP(j) = rangecells.item(j).value
        If Len(NameofTenant_HP(j)) > 0 Then
            If Not checkfieldspecialcharacter(NameofTenant_HP(j)) Then
                msgbox_hprptfrm ("* NameofTenant" & j & "  characters < > & ' " & Chr(34) & " are not allowed  in Sheet HOUSE_PROPERTY")
                ValidateNameofTenant_HP = False
                Exit Function
            End If
        End If
        If Len(NameofTenant_HP(j)) = 0 And (Mid(ifLetOut_HP(i), 1, 1) <> "S" And Mid(ifLetOut_HP(i), 1, 1) <> "(") Then
            If (Mid(ifLetOut_HP(i), 1, 1) <> "D") Then
            count = count + 1
            End If
        End If
       
        Next
          If count = UBound(NameofTenant_HP) Then
            'Commented by Shrutika
            'msgbox_hprptfrm ("* NameofTenant  is required for Let out property in property block in Sheet HOUSE_PROPERTY" & i)
            'Added by Shrutika
             msgbox_hprptfrm ("* ""Please enter name of Tenant in schedule HP""")
             
            ValidateNameofTenant_HP = False
            Exit Function
        End If
    Next
End Function
 
Function ValidatePANofTenant_HP() As Boolean
    Dim i As Long
    Dim tempPan As String
    Dim rangecells As Range
    
    ValidatePANofTenant_HP = True
    setTblinfo_hprptfrm
    
    
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        Set rangecells = Sheet8.Range("HP.PANofTenant" & i).Cells
        ReDim PANofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        
        For j = 1 To rangecells.Cells.count
        PANofTenant_HP(j) = rangecells.item(j).value
        
        tempPan = PANofTenant_HP(j)
        If Len(PANofTenant_HP(j)) > 0 Then
            If Not ValidatePAN(tempPan) Then
            'Commented by Shrutika
                'msgbox_hprptfrm ("* PANofTenant" & j & "  is invalid. First 5 alphabets, next 4 digits, then alphabet  in Sheet HOUSE_PROPERTY")
                'msgbox_hprptfrm ("* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP ")
                  'Added by Shrutika
                  msgbox_hprptfrm ("* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP""")
                ValidatePANofTenant_HP = False
                Exit Function
            End If
        End If
        
        Next
        
    Next
End Function
Function ValidateAadharofTenant_HP() As Boolean
    Dim i As Long
    Dim tempPan As String
    Dim rangecells As Range

    ValidateAadharofTenant_HP = True
    setTblinfo_hprptfrm

    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        Set rangecells = Sheet8.Range("HP.AadharofTenant" & i).Cells
        ReDim AadharofTenant_HP(rangecells.Cells.count)
        Dim j As Variant

        For j = 1 To rangecells.Cells.count
        AadharofTenant_HP(j) = rangecells.item(j).value

          If Len(AadharofTenant_HP(j)) > 0 Then
            If Len(AadharofTenant_HP(j)) <> 12 Then
                msgbox_hprptfrm ("*Aadhaar no of other Tenant " & j & " shall not exceed or less than 12 Digits in Sheet HOUSE_PROPERTY")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If AadharofTenant_HP(j) = "000000000000" Then
                msgbox_hprptfrm ("*Invalid Aadhaar of Tenant should be 12 digits " & j & " shall not exceed or less than 12 Digits in  Sheet HOUSE_PROPERTY")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If AadharofTenant_HP(j) = "111111111111" Then
                msgbox_hprptfrm ("*Invalid Aadhaar of Tenant should be 12 digits " & j & " shall not exceed or less than 12 Digits in Sheet HOUSE_PROPERTY")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If Not checkallfieldspecialcharacter(AadharofTenant_HP(j)) Then
                msgbox_hprptfrm ("* Aadhar of Tenant" & j & "   Special characters are not allowed in Sheet HOUSE_PROPERTY")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If
            End If
        Next

    Next
End Function

Function ValidateTANofTenant_HP() As Boolean
    Dim i As Long
    Dim tempTan As String
    Dim rangecells As Range
    
    ValidateTANofTenant_HP = True
    setTblinfo_hprptfrm
    
    
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        Set rangecells = Sheet8.Range("HP.TANofTenant" & i).Cells
        ReDim TANofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        
        For j = 1 To rangecells.Cells.count
        TANofTenant_HP(j) = rangecells.item(j).value
        
        tempTan = TANofTenant_HP(j)
        If Len(TANofTenant_HP(j)) > 0 Then
        If Not (CheckTAN(tempTan) Or ValidateTenantPan(tempTan)) Then
                'msgbox_hprptfrm ("* PAN/TAN  of Tenant(s)" & j & "  is invalid. Please enter in valid format in Sheet HOUSE_PROPERTY")
                msgbox_hprptfrm ("* ""Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet or First 4 Alphabets, next 5 digits, then 1 Alphabet in schedule HP""")
                ValidateTANofTenant_HP = False
                Exit Function
            End If
        End If
        
        Next
        
    Next
End Function

Function ValidateAnnualLetableValue_HP() As Boolean
    Dim i As Long
    ValidateAnnualLetableValue_HP = True
    setTblinfo_hprptfrm
    
    ReDim AnnualLetableValue_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        AnnualLetableValue_HP(i) = Sheet8.Range("HP.AnnualLetableValue" & i).value
        'If AnnualLetableValue_HP(i) = 0 And Mid(ifLetOut_HP(i), 1, 1) = "Y" Then
          If AnnualLetableValue_HP(i) = 0 And (Mid(ifLetOut_HP(i), 1, 1) = "D" Or Mid(ifLetOut_HP(i), 1, 1) = "L") Then
            'msgbox_hprptfrm ("* Gross Rent received or receivable or letable value " & i & "   must be greater than 0 for Let out Property in Sheet HOUSE_PROPERTY ")
            'Malli
            msgbox_hprptfrm ("* If ""Type of House Property"" is ""Let Out/ Deemed Let Out"" then Gross Rent received or receivable or lettable value field cannot be zero/ blank")
            ValidateAnnualLetableValue_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateRentNotRealized_HP() As Boolean
    Dim i As Long
    
    ValidateRentNotRealized_HP = True
    setTblinfo_hprptfrm
    ReDim RentNotRealized_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        RentNotRealized_HP(i) = Sheet8.Range("HP.RentNotRealized" & i).value
    Next
End Function

Function ValidateLocalTaxes_HP() As Boolean
    Dim i As Long
    
    ValidateLocalTaxes_HP = True
    setTblinfo_hprptfrm
    ReDim LocalTaxes_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        LocalTaxes_HP(i) = Sheet8.Range("HP.LocalTaxes" & i).value
    Next
End Function

Function ValidateTotalUnrealizedAndTax_HP() As Boolean
    Dim i As Long
    
    ValidateTotalUnrealizedAndTax_HP = True
    setTblinfo_hprptfrm
    ReDim TotalUnrealizedAndTax_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        TotalUnrealizedAndTax_HP(i) = Sheet8.Range("HP.TotalUnrealizedAndTax" & i).value
    Next
End Function

Function ValidateBalanceALV_HP() As Boolean
    Dim i As Long
    
    ValidateBalanceALV_HP = True
    setTblinfo_hprptfrm
    ReDim BalanceALV_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        BalanceALV_HP(i) = Sheet8.Range("HP.BalanceALV" & i).value
    Next
End Function

Function ValidateThirtyPercentOfBalance_HP() As Boolean
    Dim i As Long
    
    ValidateThirtyPercentOfBalance_HP = True
    setTblinfo_hprptfrm
    ReDim ThirtyPercentOfBalance_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        ThirtyPercentOfBalance_HP(i) = Sheet8.Range("HP.ThirtyPercentOfBalance" & i).value
    Next
End Function

Function ValidateIntOnBorwCap_HP() As Boolean
    Dim i As Long
    
    ValidateIntOnBorwCap_HP = True
    setTblinfo_hprptfrm
    ReDim IntOnBorwCap_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        IntOnBorwCap_HP(i) = Sheet8.Range("HP.IntOnBorwCap" & i).value
    Next
End Function

Function ValidateTotalDeduct_HP() As Boolean
    Dim i As Long
    
    ValidateTotalDeduct_HP = True
    setTblinfo_hprptfrm
    ReDim TotalDeduct_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        TotalDeduct_HP(i) = Sheet8.Range("HP.TotalDeduct" & i).value
    Next
End Function

Function ValidateArrears_HP() As Boolean
    Dim i As Long
    
    ValidateArrears_HP = True
    setTblinfo_hprptfrm
    ReDim Arrears_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        Arrears_HP(i) = Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & i).value
    Next
End Function


Function ValidateIncomeOfHP_HP() As Boolean
    Dim i As Long
    
    ValidateIncomeOfHP_HP = True
    setTblinfo_hprptfrm
    
    ReDim IncomeOfHP_HP(end_hprptfrm)
    ReDim IncomeOfHPInOwnHand_HP(end_hprptfrm)
    'ValidateAddrDetail_HP
    For i = 1 To end_hprptfrm
        IncomeOfHP_HP(i) = Sheet8.Range("HP.IncomeOfHP" & i).value
        IncomeOfHPInOwnHand_HP(i) = Sheet8.Range("HP.IncomeOfHPInOwnHand" & i).value
    Next
End Function

Function ValidateRentArearsSec25BAfter30pcDeduct_HP() As Boolean
    ValidateRentArearsSec25BAfter30pcDeduct_HP = True
'    RentArearsSec25BAfter30pcDeduct_HP = Sheet8.Range("HP.RentArearsSec25BAfter30pcDeduct").value
End Function

Function ValidateTotalIncomeChargeableUnHP_HP() As Boolean
    ValidateTotalIncomeChargeableUnHP_HP = True
    TotalIncomeChargeableUnHP_HP = Sheet8.Range("HP.TotalIncomeChargeableUnHP").value
    
    If Len(TotalIncomeChargeableUnHP_HP) > 14 Then
        msgbox_hprptfrm ("* Total income chargeable cannot be greater than 14 digits")
        ValidateTotalIncomeChargeableUnHP_HP = False
    End If
    
End Function

Function DefaultifLetOut_HP() As String
    DefaultifLetOut_HP = "S"
End Function

Function DefaultAnnualLetableValue_HP() As String
    DefaultAnnualLetableValue_HP = "0"
End Function

Function DefaultRentNotRealized_HP() As String
    DefaultRentNotRealized_HP = "0"
End Function

Function DefaultLocalTaxes_HP() As String
    DefaultLocalTaxes_HP = "0"
End Function

Function DefaultTotalUnrealizedAndTax_HP() As String
    DefaultTotalUnrealizedAndTax_HP = "0"
End Function

Function DefaultBalanceALV_HP() As String
    DefaultBalanceALV_HP = "0"
End Function

Function DefaultThirtyPercentOfBalance_HP() As String
    DefaultThirtyPercentOfBalance_HP = "0"
End Function

Function DefaultIntOnBorwCap_HP() As String
    DefaultIntOnBorwCap_HP = "0"
End Function

Function DefaultTotalDeduct_HP() As String
    DefaultTotalDeduct_HP = "0"
End Function

Function DefaultIncomeOfHP_HP() As String
    DefaultIncomeOfHP_HP = "0"
End Function


Function DefaultRentArearsSec25BAfter30pcDeduct_HP() As String
    DefaultRentArearsSec25BAfter30pcDeduct_HP = "0"
End Function

Function DefaultTotalIncomeChargeableUnHP_HP() As String
    DefaultTotalIncomeChargeableUnHP_HP = "0"
End Function


Function GetLetOut() As Variant
    If ValidateifLetOut_HP Then
        GetLetOut = ifLetOut_HP
    End If
End Function


Sub AddBlockCall_hprptfrm()
    setTblinfo_hprptfrm
    Call addblock(rngname_hprptfrm, frmRngname_hprptfrm, cntrRng_hprptfrm, frmsize_hprptfrm)
End Sub


Function addblock(prevrangename As Variant, blockname As Variant, frmcounter As Variant, totalblocksize As Variant)
    Dim i, j, k, l, commindex As Long
    Dim newnamerefersto As Variant
    Dim gapbtn, dcounter, counter, offset1, offset2 As Long
    Dim lenn, startrow, endrow, endrow1, endrow2 As Long
    Dim destinationrowindex, destinationcolumnindex, lastdestinationrow As String
    Dim rangearr, newranges, precrange, oldrangeaddress, newrangeaddress, newrangeaddress1, newrangeaddress2 As Variant
    Dim ndestinationrowindex, ndestinationcolumnindex As Long
    Dim noofrowsadded As Long
    Dim rangeCount As Long


    
    Application.EnableEvents = False
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    gapbtn = 0
    dcounter = 0
    counter = Range(frmcounter).value
    Dim lastrow As Variant
    lastrow = onlyDigits(Sheet8.Range("HP.LastCell").AddressLocal, "I")
    
    If counter > 0 Then
        noofrowsadded = (totalblocksize - 27) / counter
        offset1 = 27 + (1) * noofrowsadded
        offset2 = 27 + (counter) * noofrowsadded
        offset1 = Range(blockname & 1).Rows.count
        ndestinationrowindex = Range(blockname & counter).row
        ndestinationrowindex = ndestinationrowindex + gapbtn + offset1
        ndestinationrowindex = lastrow
        ndestinationcolumnindex = Range(blockname & counter).Column
        newnamerefersto = Range(blockname & counter).Address
        lastdestinationrow = ndestinationrowindex + offset1 - 1
        Cells(CLng(ndestinationrowindex - 1), CLng(ndestinationcolumnindex)).Select
    
        Call InsertBlockHP(offset1)
    
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
    
        sPassword = EfilingCommon.getmsgstate
        ActiveSheet.Unprotect Password:=sPassword
        
        Range(frmcounter).value = counter
        
        Application.EnableEvents = False
    
        ActiveSheet.Unprotect Password:=sPassword
    
        'Range(blockname & counter - 1).Copy Destination:=Cells(ndestinationrowindex, ndestinationcolumnindex)
        Range(blockname & 1).Copy Destination:=Cells(ndestinationrowindex, ndestinationcolumnindex)
    End If
    
    '' add all range name to block
    
    commindex = 2
    j = 0
    
    prevrangename = Split(prevrangename, ";")
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
    
    For l = 0 To UBound(newrangeaddress)
        If Not newranges(l) = "" Then
            ThisWorkbook.Names.add name:=newranges(l), _
            RefersTo:="=" & newrangeaddress(l), Visible:=True
            If Range(newranges(l)).Interior.ColorIndex = 35 Then
                Range(newranges(l)).ClearContents
            End If
        End If
    Next
    
    LockUnlockCoOwnersDetails (Range(frmcounter).value)
    
    'by sai on 26/05/2025 as it was auto populated field and changed the logic as per DE V0.5 and V0.6
    Range("TotAmt.24b" & Range(frmcounter).value).formula = "=SUM(Intrst.24b" & Range(frmcounter).value & ")"
    Range("HP.IntOnBorwCap" & Range(frmcounter).value).formula = "=MAX(0,IF(MID(HP.ifLetOut" & Range(frmcounter).value & ",1,1)=""S"",MIN(200000,IF(AND(bacValue=1,MID(HP.ifLetOut" & Range(frmcounter).value & ",1,1)=""S""),0,TotAmt.24b" & Range(frmcounter).value & ")),IF(AND(bacValue=1,MID(HP.ifLetOut" & Range(frmcounter).value & ",1,1)=""S""),0,TotAmt.24b" & Range(frmcounter).value & ")))"
    '----------------------------------------------------------------
    ActiveSheet.Protect Password:=sPassword
    Application.EnableEvents = True
End Function

Function InsertBlockHP(vRows1 As Variant)
    Dim x As Long
    Dim vRows As Long
    Dim sht As Worksheet, shts() As String, i As Long
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    Application.EnableEvents = False
    ActiveCell.EntireRow.Select
    
    
    vRows = CLng(vRows1)
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
    ActiveSheet.Protect Password:=sPassword
    Application.EnableEvents = True
End Function

Sub setTblinfo_hpcoindex(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("hp.co.name" & myindex).count
    Set rangecells = Sheet8.Range("hp.co.name" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco = ccount
    rngname_hpco = "HP.Co.Srno;HP.Co.Name;HP.Co.Pan;HP.Co.Aadhaar;HP.Co.Share;"
End Sub

'added by Chetan C M for AY 2025-26
'start--
Sub setTblinfo_24bankname(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("bankName.24b" & myindex).count
    Set rangecells = Sheet8.Range("bankName.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or UCase(scode) = UCase("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24b = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub

Sub setTblinfo_24bBankorInst(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("LoanfrmBankOrInstitute.24b" & myindex).count
    Set rangecells = Sheet8.Range("LoanfrmBankOrInstitute.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
    
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24b = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bLoan(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("loanAccNum.24b" & myindex).count
    Set rangecells = Sheet8.Range("loanAccNum.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24e = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bdate(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("loanDate.24b" & myindex).count
    Set rangecells = Sheet8.Range("loanDate.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24f = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub

Sub setTblinfo_24bAmount(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("loanAmt.24b" & myindex).count
    Set rangecells = Sheet8.Range("loanAmt.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24k = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bloanout(ByVal myindex As Long)
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("loanOutstanding.24b" & myindex).count
    Set rangecells = Sheet8.Range("loanOutstanding.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24s = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"

End Sub
Sub setTblinfo_24bInt(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("Intrst.24b" & myindex).count
    Set rangecells = Sheet8.Range("Intrst.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24r = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"

End Sub

Sub setTblinfo_hpco2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet8.Range("hp.co.name1").count
    Set rangecells = Sheet8.Range("hp.co.name1").Cells
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If isdropdownblank(scode) Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco2 = ccount
    rngname_hpco2 = "HP.Co.Srno;HP.Co.Name;HP.Co.Pan;HP.Co.Aadhaar;HP.Co.Share;"
End Sub

Function ValidateCoownerName1_HP() As Boolean
    Dim i, j As Long
    Dim rangecells As Range
    ValidateCoownerName1_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnerName1_HP(end_hprptfrm)
    
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet8.Range("HP.Co.Name" & j).Cells
    For i = 1 To end_hprptfrm
        CoOwnerName1_HP(i) = rangecells.item(i).value
        'Updated by Shrutika(08/09/2025)
'        If Not checkfieldspecialcharacter(CoOwnerName1_HP(i)) Then
'            msgbox_hprptfrm ("* Coowner NameA" & i & "   characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
'            ValidateCoownerName1_HP = False
'            Exit Function
'        End If
'--------------------------------
    Next
    Next
End Function


Function ValidatePercenShare1_HP() As Boolean
    Dim i, j As Long
    Dim rangecells As Range
    ValidatePercenShare1_HP = True
    setTblinfo_hprptfrm
    
    
    ReDim CoOwnerSharePer1_HP(end_hprptfrm)
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet8.Range("HP.Co.Share" & j).Cells
    For i = 1 To end_hprptfrm
        CoOwnerSharePer1_HP(i) = rangecells.item(i).value
        If Not checkfieldspecialcharacter(CoOwnerSharePer1_HP(i)) Then
            msgbox_hprptfrm ("* Percentage Share in PropertyA" & i & "   characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
            ValidatePercenShare1_HP = False
            Exit Function
        End If
    Next
    Next
End Function

Function ValidateCoOwnerPAN1_HP() As Boolean
    Dim i, j As Long
    Dim tempPan As String
    Dim rangecells As Range
 
    ValidateCoOwnerPAN1_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnerPAN1_HP(end_hprptfrm)
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet8.Range("HP.Co.Pan" & j).Cells
    
    For i = 1 To end_hprptfrm
        CoOwnerPAN1_HP(i) = rangecells.item(i).value
        If Not checkfieldspecialcharacter(CoOwnerPAN1_HP(i)) Then
            msgbox_hprptfrm ("* PAN of CoownerA" & i & "   characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
            ValidateCoOwnerPAN1_HP = False
            Exit Function
        End If
    '22/08/2024 Chetan C M changes start
        If UCase(CoOwnerPAN1_HP(i)) <> "" Then
    '22/08/2024 Chetan C M changes end

        'new change
        If UCase(CoOwnerPAN1_HP(i)) = Sheet1.Range("sheet1.PAN").value Then
                msgbox_hprptfrm ("* Enter Valid PAN of Other co-owner(s) at Sr. No " & i & " in Property " & i & " Sheet HOUSE_PROPERTY  Which should not match pan with Sheet Part A - General" & Chr(13))
                ValidateCoOwnerPAN1_HP = False
            Exit Function
        End If
                        'new change
    '22/08/2024 Chetan C M changes start
        End If
    '22/08/2024 Chetan C M changes end
    Next
    Next
End Function
Function ValidateCoOwnerAadhar_HP() As Boolean
    Dim i, j As Long
    Dim tempPan As String
    Dim rangecells As Range
 
    ValidateCoOwnerAadhar_HP = True
    setTblinfo_hprptfrm
    
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet8.Range("HP.Co.Aadhaar" & j).Cells
    ReDim CoOwnerAadhar_HP(rangecells.count)
    
    For i = 1 To rangecells.count 'end_hprptfrm
        CoOwnerAadhar_HP(i) = rangecells.item(i).value
        
         If Len(CoOwnerAadhar_HP(i)) > 0 Then
                If Len(CoOwnerAadhar_HP(i)) > 12 Then
                'msgbox_hprptfrm ("*Aadhaar no of other co-owner(s) " & j & " shall not exceed or less than 12 Digits in " & i & " Sheet HOUSE_PROPERTY")
                
                'Shrutika (08/09/2025)
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
                
                If CoOwnerAadhar_HP(i) = "000000000000" Then
                'msgbox_hprptfrm ("*Invalid Aadhaar of Co-Owner(s) should be 12 digits " & j & " shall not exceed or less than 12 Digits in " & i & " Sheet HOUSE_PROPERTY")
                        
                'Shrutika (08/09/2025)
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
                
                If CoOwnerAadhar_HP(i) = "111111111111" Then
                'msgbox_hprptfrm ("*Invalid Aadhaar of Co-Owner(s) should be 12 digits " & j & " shall not exceed or less than 12 Digits in " & i & " Sheet HOUSE_PROPERTY")
                
               'Shrutika (08/09/2025)
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
            
    
                If Not checkallfieldspecialcharacter(CoOwnerAadhar_HP(i)) Then
                'msgbox_hprptfrm ("* Aadhar of CoownerA" & i & "   Special characters are not allowed in Sheet HOUSE_PROPERTY")
                
                'Shrutika (08/09/2025)
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
         End If
    Next
    Next
End Function


Function ValidateSharePercent_HP() As Boolean
    Dim i As Long
    
    ValidateSharePercent_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnedShare_HP(end_hprptfrm)
    ReDim CoOwnedYN_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        CoOwnedShare_HP(i) = Sheet8.Range("HP.SharePercent" & i).value
        CoOwnedYN_HP(i) = Sheet8.Range("HP.CoOwnedYN" & i).value
        
        CoOwnedShare_HP(i) = Round(CoOwnedShare_HP(i), 2)
        
        If Not checkfieldspecialcharacter(CoOwnedShare_HP(i)) Then
            msgbox_hprptfrm ("* Your Percentage of Share" & i & "  characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
            ValidateSharePercent_HP = False
            Exit Function
        End If
        
        If Mid(CoOwnedYN_HP(i), 1, 1) = "N" Then
            If CoOwnedShare_HP(i) <> 100 Then
                msgbox_hprptfrm ("* Your Percentage of Share" & i & "   must be 100 if not Co owned in Sheet HOUSE_PROPERTY")
                ValidateSharePercent_HP = False
                Exit Function
            End If
        Else
'            If CoOwnedShare_HP(i) = 100 Then
'                msgbox_hprptfrm ("* Your Percentage of Share" & i & " must NOT be 100 if Co owned in Sheet HOUSE_PROPERTY")
'                ValidateSharePercent_HP = False
'                Exit Function
'            End If
        End If
    Next
End Function

Function ValidateCoOwnedYN_HP() As Boolean
    Dim i As Long
    
    ValidateCoOwnedYN_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnedYN_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        CoOwnedYN_HP(i) = Sheet8.Range("HP.CoOwnedYN" & i).value
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If
        If Not chkCompulsory(CoOwnedYN_HP(i)) Then
        'Ankita(23-Jan)
            ' msgbox_hprptfrm ("* IS Property Co-Owned" & i & "  is Mandatory in Sheet HOUSE_PROPERTY")
           'msgbox_hprptfrm ("* Please select dropdown from Is property co-owned  in schedule HP")
            msgbox_hprptfrm ("* ""Please select dropdown from Is property co-owned  in schedule HP""") & Chr(13)
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* IS Property Co-Owned" & i & "  characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* IS Property Co-Owned" & i & " characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateCoownerRules_HP() As Boolean
    Dim i As Long
    ValidateCoownerRules_HP = True
    setTblinfo_hprptfrm
    For i = 1 To UBound(CoOwnedYN_HP)
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If

        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
          Dim CoOwnerShare As Variant
          CoOwnerShare = CoOwnedShare_HP(i)
          
        End If
  
        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
         Dim counter, name As Variant
         counter = 0
         For Each name In Sheet8.Range("HP.Co.Name" & i).Cells
         If name = "" Then counter = counter + 1
            If counter = Sheet8.Range("HP.Co.Name" & i).Cells.count Then
            'Commented by Shrutika
               ' msgbox_hprptfrm ("* There must be Atleast one Name of Co-ownerA" & i & " in Sheet HOUSE_PROPERTY")
              'Added by Shrutika
              '22/08/2024 Chetan C M changes start
'                msgbox_hprptfrm ("* ""There must be at least one Co-owner  in schedule HP """)
               'Commented by Shruika
               ' msgbox_hprptfrm ("* ""Please enter the name of the Co-owner""")
                msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP""")
                '22/08/2024 Chetan C M changes end
                ValidateCoownerRules_HP = False
                Exit Function
                
             End If
        Next
        
        '21/08/2024 Chetan C M added
         
        Dim j As Long
        Dim coowener As Range
        Dim Pancoowener As Range
        Dim Aacoowener As Range
        Dim percoowener As Range
        For j = 1 To Sheet8.Range("HP.Co.Name" & i).count

            Set coowener = Sheet8.Range("HP.Co.Name" & i).Cells
            Set Pancoowener = Sheet8.Range("HP.Co.Pan" & i).Cells
            Set Aacoowener = Sheet8.Range("HP.Co.Aadhaar" & i).Cells
            Set percoowener = Sheet8.Range("HP.Co.Share" & i).Cells
            If coowener.item(j).value = "" Then
                If Pancoowener.item(j).value <> "" Or Aacoowener.item(j).value <> "" Or percoowener.item(j).value <> "" Then
                  'Commented by Shrutika
                   ' msgbox_hprptfrm ("* ""Please enter name of Co-owner  in schedule HP at S.No """ & j)
                    'msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP at S.No " & j & """")
                    
                    'changed by Chetan C M on 13/10/2025 for SIT-89185
                    'start--
'                    msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP at S.No " & j & """")
                     msgbox_hprptfrm ("* ""Please enter the name of the co-owner""")
                     '--end
                     
                    ValidateCoownerRules_HP = False
                Exit Function
                End If
            End If
        Next j
            
         '21/08/2024 Chetan C M additon ended
         
         End If
        Next
        

End Function
'22/08/2024 addedd by Chetan C M
Function ValidatesharesRules_HP() As Boolean
     Dim i As Long
    ValidatesharesRules_HP = True
    setTblinfo_hprptfrm
    For i = 1 To UBound(CoOwnedYN_HP)
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If

        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
          Dim CoOwnerShare As Variant
          CoOwnerShare = CoOwnedShare_HP(i)
          
        End If
  
        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
         Dim counter, name, name1 As Variant
         counter = 0
         For Each name In Sheet8.Range("HP.Co.Share" & i).Cells
          If name = "" Then counter = counter + 1
                
                    If counter = Sheet8.Range("HP.Co.Share" & i).Cells.count Then
                    
                     msgbox_hprptfrm ("* ""Please enter Percentage Share of Co-Owner in schedule HP """)
                        ValidatesharesRules_HP = False
                        Exit Function
                    End If
         Next
         
           
         
    Dim j As Long
    Dim coowener As Range
    Dim Pancoowener As Range
    Dim Aacoowener As Range
    Dim percoowener As Range
         For j = 1 To Sheet8.Range("HP.Co.Name" & i).count

         Set coowener = Sheet8.Range("HP.Co.Name" & i).Cells
         Set Pancoowener = Sheet8.Range("HP.Co.Pan" & i).Cells
         Set Aacoowener = Sheet8.Range("HP.Co.Aadhaar" & i).Cells
         Set percoowener = Sheet8.Range("HP.Co.Share" & i).Cells
         
         If percoowener.item(j).value = "" Then
         If coowener.item(j).value <> "" Or Pancoowener.item(j).value <> "" Or Aacoowener.item(j).value <> "" Then
         msgbox_hprptfrm ("* ""Please enter Percentage Share of Co-Owner in schedule HP at S.No """ & j)
        ValidatesharesRules_HP = False
         Exit Function
         End If
         End If
         Next j
                 
        End If
    Next
End Function
Sub AddRows_hpco()
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim te As Long
    Dim i As Long
    Dim frmsize_hprptfrmnew As Long
    Dim newfrmsize_hprptfrm As Long
    
    newrngname = rngname_hpco
    For te = 1 To Sheet8.Range("cntr.hprptfrm").value
        Application.EnableEvents = False
        setTblinfo_hpco
        SelectLastRow ("hp.co.srno" & te)
        
        If te > 1 Then
            numberofrows = InsertRowsAndFillFormulasBPA(numberofrows)
        Else
            numberofrows = InsertRowsAndFillFormulasBPA()
        End If
        
        Call ExendRangeNameToTablehp(CLng(numberofrows), rngname_hpco, te)
        Application.EnableEvents = True
    Next

    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    Application.EnableEvents = False
    For i = 1 To Sheet8.Range("cntr.hprptfrm").value
        newfrmsize_hprptfrm = Sheet8.Range("hprptfrm.size").value
        newfrmsize_hprptfrm = newfrmsize_hprptfrm + numberofrows
        Sheet8.Range("hprptfrm.size").value = newfrmsize_hprptfrm
    Next
    Application.EnableEvents = True
    Sheet8.Protect Password:=sPassword
End Sub

Sub setTblinfo_hpco()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    
    mIntCells = Sheet8.Range("hp.co.name1").count
    Set rangecells = Sheet8.Range("hp.co.name1").Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        If isdropdownblank(scode) Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco = ccount
    rngname_hpco = "HP.Co.Srno;HP.Co.Name;HP.Co.Pan;HP.Co.Share;"
End Sub
 

Sub ExendRangeNameToTablehp(numberofrows As Long, rangenamestring As Variant, myindex As Long)
    Dim i As Long
    Dim firstbound, temp, upperbound, x, lastbound, RangeAddress As Variant
    
    rangenamestring = Split(rangenamestring, ";")
    
    For i = 0 To UBound(rangenamestring) - 1
        rangenamestring(i) = rangenamestring(i) & myindex
        
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


Sub ValidateIncmFrmHP(HpropertyCount As Long)
On Error Resume Next

Dim type_HP As String
Dim incmHP As Double

'For HpropertyCount = 1 To Sheet8.Range("HousePropertySectionCount").Value
type_HP = Sheet8.Range("HP.ifLetOut" & HpropertyCount)
incmHP = Sheet8.Range("HP.IntOnBorwCap" & HpropertyCount)

If type_HP = "No" Then
      If incmHP > 200000 Then
'Changed by Riyaz on 23/06/2025
'            fmsgboxStatus ("* Interest cannot be more than 200000 for category, self-occupied.")
            fmsgboxsmall "*""Amount of deduction for interest u/s 24(b) in case of Self Occupied House Property or aggregate of such properties  cannot exceed 2 lakh. Therefore, interest is being restrcited to 2 lakh""" & Chr(13)
            Sheet8.Range("HP.IntOnBorwCap" & HpropertyCount).value = ""
            Sheet8.Range("HP.IntOnBorwCap" & HpropertyCount).Select
         End If
    End If
'End If
End Sub


Function CheckHousePropertyIncome(HP As Variant) As Boolean
On Error Resume Next
 CheckHousePropertyIncome = True
'Comment by Riyaz
 'If Mid(Range("HP.ifLetOut1").value, 1, 1) = "S" Then
 If (HP > 200000) Then
 fmsgboxsmall "* Interest cannot be more than Rs. 2,00,000 for category self-occupied in schedule HP." ', vbOKOnly, "Error(s)!"
 'Range("HP.IntOnBorwCap1").Select
 'Range("HP.IntOnBorwCap1").value = ""
  CheckHousePropertyIncome = False   'Malli 23/10/2024
 Exit Function
' End If
 End If
End Function

Function AddTotIncomeUnderHouseProperty_1()
    Dim TotIncomeRangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, Totsections As Long

    tot = 0
    TotIncomeRangeName = "HP.IncomeOfHP"
    Totsections = Sheet8.Range("PropertySectionCOunt").value
        For i = 1 To Totsections
        newDefinedName = TotIncomeRangeName & i
        tot = tot + Sheet8.Range(newDefinedName)
        Next
        
        
    Sheet8.Unprotect Password:=EfilingCommon.getmsgstate
      Sheet8.Range("HP.TotalIncomeChargeableUnHP").value = tot + Sheet8.Range("HP.PassTroughIncome").value '+ Sheet8.Range("HP.RentArearsSec25BAfter30pcDeduct").value
    Sheet8.Protect Password:=EfilingCommon.getmsgstate


End Function

'added by Chetan C M for AY 2025-26
'start--


'Sub Validate SheetHouse Property_Click()
'Dim vbMessgaeCaption As String
'
'vbMessgaeCaption = "ITR 2: AY: 2025-26"
'Validate24b_All
'
'fmsgboxStatus "House Property is OK"
'End Sub


'Sub Validate24b_All()
'Dim vbMessgaeCaption As String
'vbMessgaeCaption = "Error"
'   Dim sourceSheet As Worksheet
'    Set sourceSheet = ThisWorkbook.Sheets("House Property")
'If Not Validate_24b Then
'    sourceSheet.Activate
'   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
'    fmsgboxStatus MsgBox_24b
'    CloseMsg
'End If
'End Sub

Function Validate_24b()
Validate_24b = True

setTblinfo_24bLoanfrm
'setTblinfo_24bIFSC

'setTblinfo_24bPAN
setTblinfo_24bBankName
setTblinfo_24bAccntNum
setTblinfo_24bLoanDate
setTblinfo_24bLoanAmt
setTblinfo_24bLoanOutstanding
setTblinfo_24bIntrst


end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, , end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)


If Not ValidateLoanfrm_24b Then Validate_24b = False

If Not ValidateBankName_24b Then Validate_24b = False
If Not ValidateAccntNum_24b Then Validate_24b = False
If Not ValidateLoanDate_24b Then Validate_24b = False
If Not ValidateLoanAmt_24b Then Validate_24b = False
If Not ValidateLoanOutstanding_24b Then Validate_24b = False
If Not ValidateIntrst_24b Then Validate_24b = False
If Not Validategreater_24b Then Validate_24b = False


End Function

Sub setTblinfo_24bLoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("LoanfrmBankOrInstitute.24b" & i).count
    Set rangecells = Range("LoanfrmBankOrInstitute.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanfrm = ccount
End Sub



Sub setTblinfo_24bBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("bankName.24b" & i).count
    Set rangecells = Range("bankName.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bbankName = ccount
End Sub


Sub setTblinfo_24bAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("loanAccNum.24b" & i).count
    Set rangecells = Range("loanAccNum.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bAccntNum = ccount
End Sub

Sub setTblinfo_24bLoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("loanDate.24b" & i).count
    Set rangecells = Range("loanDate.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanDate = ccount
End Sub

Sub setTblinfo_24bLoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("loanAmt.24b" & i).count
    Set rangecells = Range("loanAmt.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanAmt = ccount
End Sub

Sub setTblinfo_24bLoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("loanOutstanding.24b" & i).count
    Set rangecells = Range("loanOutstanding.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanOutstanding = ccount
End Sub

Sub setTblinfo_24bIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
    mIntCells = Range("Intrst.24b" & i).count
    Set rangecells = Range("Intrst.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bIntrst = ccount
End Sub



Function ValidateLoanfrm_24b() As Boolean
    ValidateLoanfrm_24b = True

    Dim rangecells As Range
    Dim i, j As Long
For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("LoanfrmBankOrInstitute.24b" & j).Cells
    ReDim Loanfrm_24b(end_24b)
    For i = 1 To end_24b
        Loanfrm_24b(i) = rangecells.item(i).value
        If isdropdownblank(Loanfrm_24b(i)) Then
             msgbox_hprptfrm ("* Please select dropdown from ""Loan taken from"" in schedule 24(b) at Sr. No " & i & "") & Chr(13)
            ValidateLoanfrm_24b = False
            Exit Function
        End If
         If Loanfrm_24b(i) = "(Select)" Then
          msgbox_hprptfrm ("* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 24(b) schedule at Sr. No " & i & "") & Chr(13)
            ValidateLoanfrm_24b = False
            Exit Function
        End If

Next i
Next j
End Function
'Shrutika(02/06/25)
'Function ValidateLoanfrm_24b() As Boolean
'    Dim i As Long
'
'    ValidateLoanfrm_24b = True
'    setTblinfo_hprptfrm
'    ReDim Loanfrm_24b(end_hprptfrm)
'
'    For i = 1 To end_hprptfrm
'        If Len(AddrDetail_HP(i)) = 0 Then
'            Exit For
'        End If
'
'        Loanfrm_24b(i) = Sheet8.Range("LoanfrmBankOrInstitute.24b" & i).value
'        If Not chkCompulsory(Loanfrm_24b(i)) Then
'
'            msgbox_hprptfrm ("* ""Please select dropdown from Loan taken from: Sl no B2""")
'
'            ValidateLoanfrm_24b = False
'            Exit Function
'        End If
'        If Not checkfieldspecialcharacter(Loanfrm_24b(i)) Then
'            msgbox_hprptfrm ("* CityOrTownOrDistrict" & i & "   characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
'            ValidateLoanfrm_24b = False
'            Exit Function
'        End If
'    Next
'End Function


Function ValidateBankName_24b() As Boolean
    ValidateBankName_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("bankName.24b" & j).Cells
    ReDim BankName_24B(end_24b)
    For i = 1 To end_24b
        BankName_24B(i) = rangecells.item(i).value
        If Not chkCompulsory(BankName_24B(i)) Then
            
            msgbox_hprptfrm ("*""Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateBankName_24b = False
            Exit Function
        End If
        
         If Len(BankName_24B(i)) > 125 Then
           msgbox_hprptfrm ("* Name of the Bank/ Institution/Person at Sr. No " & i & " in Sheet 24(b) should be less than or equal to 125 characters.") & Chr(13)
            ValidateBankName_24b = False
            Exit Function
        End If
        
        
        If Not checkfieldSuperSpecialcharacter(BankName_24B(i)) Then
            msgbox_hprptfrm ("* Name of the Bank/ Institution/Person in schedule 24b at Sl.no. " & i & " should not Contain <, >, characters.") & Chr(13)
            ValidateBankName_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function


Function ValidateAccntNum_24b() As Boolean
    ValidateAccntNum_24b = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim i, j As Long
    For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("loanAccNum.24b" & j).Cells
   
    Set rangecells1 = Sheet8.Range("LoanfrmBankOrInstitute.24b" & j).Cells
    ReDim AccntNum_24b(end_24b)
    ReDim BankorInst_24b(end_24b)
    For i = 1 To end_24b
        AccntNum_24b(i) = rangecells.item(i).value
        
        BankorInst_24b(i) = rangecells1.item(i).value
        If Not chkCompulsory(AccntNum_24b(i)) Then
          msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
          ' msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateAccntNum_24b = False
            Exit Function
        End If
        
         If Len(AccntNum_24b(i)) > 20 Then
         msgbox_hprptfrm ("* Loan Account number  at Sr. No " & i & " in schedule 24b less than 20 characters.") & Chr(13)
            ValidateAccntNum_24b = False
            Exit Function
        End If
        
        
        If BankorInst_24b(i) = "Bank" Then
        If Not ValidateBankAccountNumber_24b(AccntNum_24b(i)) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_24b = False
        Exit Function
         End If
        End If
    
    If BankorInst_24b(i) = "Other than bank" Then
        If Not checkfieldspecialcharacter1(AccntNum_24b(i)) Then
'            MsgBox_24b = MsgBox_24b & "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & i & "") & Chr(13)
           ValidateAccntNum_24b = False
            Exit Function
        End If
        End If
         
Next i
Next j
End Function

Function ValidateLoanDate_24b() As Boolean
    ValidateLoanDate_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("loanDate.24b" & j).Cells
    ReDim LoanDate_24B(end_24b)
    For i = 1 To end_24b
        LoanDate_24B(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanDate_24B(i)) Then
             msgbox_hprptfrm ("* ""Please provide Date of sanction of Loan in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateLoanDate_24b = False
            Exit Function
        End If
         If Len(LoanDate_24B(i)) > 10 Then
          msgbox_hprptfrm ("* Date of sanction of Loan  at Sr. No " & i & " in schedule 24b less than 10 characters.") & Chr(13)
            ValidateLoanDate_24b = False
            Exit Function
        End If
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_24B(i)) Then
            msgbox_hprptfrm ("* ""Please enter date in valid format"" at Sr. No " & i & ".") & Chr(13)
            ValidateLoanDate_24b = False
            Exit Function
        End If
        
        If Not ChkMaxDate_24b(Trim(LoanDate_24B(i)), "31-03-2025") Then
            msgbox_hprptfrm ("* Date can not be after 31/03/2025 at Sr. No " & i & ".") & Chr(13)
            ValidateLoanDate_24b = False
          Exit Function
        End If
                 
                                 
Next i
Next j
End Function

Function ValidateLoanAmt_24b() As Boolean
    ValidateLoanAmt_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("loanAmt.24b" & j).Cells
    ReDim LoanAmt_24b(end_24b)
    For i = 1 To end_24b
        LoanAmt_24b(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanAmt_24b(i)) Then
        'Ayush_05/05/2025
'             MsgBox_24b = MsgBox_24b + "* ""Please provide Total Loan taken in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
                msgbox_hprptfrm ("* ""Please provide Total amount of loan in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
'         If Len(LoanAmt_24b(i)) > 14 Then
'          MsgBox_24b = MsgBox_24b + "* Loan Amount  at Sr. No " & i & " in Sheet 24b less than 15 characters." & Chr(13)
'            ValidateLoanAmt_24b = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanAmt_24b(i)) Then
            msgbox_hprptfrm ("* Loan amount at Sr. No  " & i & "  in schedule 24b should be Numeric value") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
        
        If LoanAmt_24b(i) > 99999999999999# Then
            msgbox_hprptfrm ("* Loan amount at Sr. No  " & i & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
        
        If LoanAmt_24b(i) < 0 Or LoanAmt_24b(i) = 0 Then
        'Ayush_05/05/2025
'            MsgBox_24b = MsgBox_24b & "* Total loan taken should be more than 0 in schedule 24(b) at Sr. No  " & i & "" & Chr(13)
           msgbox_hprptfrm ("* Total amount of loan should be more than 0 in schedule 24(b) at Sr. No  " & i & "") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function

Function ValidateLoanOutstanding_24b() As Boolean
    ValidateLoanOutstanding_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("loanOutstanding.24b" & j).Cells
    ReDim LoanOutStanding_24B(end_24b)
    For i = 1 To end_24b
        LoanOutStanding_24B(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanOutStanding_24B(i)) Then
        ' Ayush_05/05/2025
'             MsgBox_24b = MsgBox_24b + "* ""Loan outstanding as on 31-03-2025 is mandatory in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
                msgbox_hprptfrm ("* ""Loan outstanding as on last date of financial year is mandatory in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
'         If Len(LoanOutstanding_24b(i)) > 14 Then
'          MsgBox_24b = MsgBox_24b + "* Loan Outstanding  at Sr. No " & i & " in Sheet 24b less than 15 characters." & Chr(13)
'            ValidateLoanOutstanding_24b = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanOutStanding_24B(i)) Then
            msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & i & "  in schedule 24b should be Numeric value") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
        
        If LoanOutStanding_24B(i) > 99999999999999# Then
            msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & i & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
        
        If LoanOutStanding_24B(i) < 0 Then
        'Ayush_05/05/2025
'            MsgBox_24b = MsgBox_24b & "* Loan outstanding as on 31.03.2025 can't be less than 0 in schedule 24(b) at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
             msgbox_hprptfrm ("* Loan outstanding as on last date of finacial year can't be less than 0 in schedule 24(b) at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment.") & Chr(13)
             
             'msgbox_hprptfrm ("* Loan outstanding as on last date of financial year is mandatory in Table Section 24(b). Please enter 0 in case the entire loan is repaid during the year") & Chr(13)
             
             
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function

Function ValidateIntrst_24b() As Boolean
    ValidateIntrst_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet8.Range("PropertySectionCOunt").count
    Set rangecells = Sheet8.Range("Intrst.24b" & j).Cells
    ReDim Intrst_24B(end_24b)
    For i = 1 To end_24b
        Intrst_24B(i) = rangecells.item(i).value
        If Not chkCompulsory(Intrst_24B(i)) Then
             msgbox_hprptfrm ("* ""Please provide Interest u/s 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateIntrst_24b = False
            Exit Function
        End If
'         If Len(Intrst_24b(i)) > 15 Then
'          MsgBox_24b = MsgBox_24b + "* Interest  at Sr. No " & i & " in Sheet 24b less than 15 characters." & Chr(13)
'            ValidateIntrst_24b = False
'            Exit Function
'        End If

        If Not IsNumeric(Intrst_24B(i)) Then
           msgbox_hprptfrm ("* Interest at Sr. No  " & i & "  in schedule 24b should be Numeric value") & Chr(13)
            
            ValidateIntrst_24b = False
            Exit Function
        End If
        
        If Intrst_24B(i) > 99999999999999# Then
            msgbox_hprptfrm ("* Interest at Sr. No  " & i & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
            ValidateIntrst_24b = False
            Exit Function
        End If
        
        'Interest u/s 24(b) should be more than 0 in schedule 24(b)
        
        If Intrst_24B(i) < 0 Or Intrst_24B(i) = 0 Then
            msgbox_hprptfrm ("* Interest u/s 24(b) should be more than 0 in schedule 24(b) at Sr. No  " & i & "") & Chr(13)
            ValidateIntrst_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function


Function ValidateBankAccountNumber_24b(BankAccountNumber As Variant) As Boolean
    ValidateBankAccountNumber_24b = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB(), ValidateIFSC As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    Dim cc
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
           ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "") & Chr(13) 'updated by Shrutika(SIT-93563)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
           ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b) " & Chr(13)
            msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
       ' MsgBox_24b = MsgBox_24b & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        msgbox_hprptfrm ("* Loan Account number at Sr.No " & cc & " is mandatory in schedule 24(b)") & Chr(13)
       ' MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
 msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
      msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        'MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If

    ReDim myB(Len(BankAccountNumber) - 1)
    For i = 1 To Len(BankAccountNumber)
        myB(i - 1) = Mid(BankAccountNumber, i, 1)
    Next

    For i = LBound(myB) To UBound(myB)
        If IsNumeric(myB(i)) Then
            countnum = countnum + 1
        End If

        If i > LBound(myB) And i < UBound(myB) Then
            If myB(i) = 0 Then
                If myB(i - 1) = 0 Then
                    zeroCount = zeroCount + 1
                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
                Else
                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
                End If
            End If
        End If

    Next

    If BeforeZero <> "" And AfterZero <> "" Then
        If zeroCount > 1 Then
            'MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc + 1 & " is invalid  in schedule 24(b)" & Chr(13)
           msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        'errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 24(b) is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        errmsgVerification = errmsgVerification & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
End Function



Public Function ChkMaxDate_24b(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_24b = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ayush_25-26
'        If Year > 2024 Then
'        If Year > 2025 Then
        If Year > 2026 Then
            ChkMaxDate_24b = False
            Exit Function
        Else
'            If Year = 2024 Then
'                If Year = 2025 Then
                If Year = 2026 Then
                If month > 4 Then
                    ChkMaxDate_24b = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_24b = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_24b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


Function Validategreater_24b() As Boolean
    Validategreater_24b = True
         If (Len(Sheet8.Range("TotAmt.24b1").value) > 14) Then
            msgbox_hprptfrm ("*  Total of interest on borrowed capital u/s 24(b) cannot exceed 14 Digits.") & Chr(13)
            Validategreater_24b = False
            Exit Function
         End If
End Function


'Sub Addrows24b()
'Sheets("House Property").Activate
'     EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
'    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.24b")
'    EfilingCommon.insertRowUnderSectionWithFormula24b
'End Sub
'--end


Function Validate24B_Table() As Boolean
    Dim i As Long
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim rangecells7 As Range
    
    Dim state_1 As Variant
    Dim sharepercent As Variant
    Dim cosharepercent As Double
    Dim end_hpco_24d As Long 'Added by Shrutika(03/06/2025) SIT-93666
    Validate24B_Table = True
    

'    sharepercent = 0
'    cosharepercent = 0
'    setTblinfo_hprptfrm
    
    end_hprptfrm = Sheet8.Range("PropertySectionCOunt").value
    
    For i = 1 To end_hprptfrm
        
        Set rangecells1 = Sheet8.Range("LoanfrmBankOrInstitute.24b" & i).Cells
        Set rangecells2 = Sheet8.Range("bankName.24b" & i).Cells
        Set rangecells3 = Sheet8.Range("loanAccNum.24b" & i).Cells
        Set rangecells4 = Sheet8.Range("loanDate.24b" & i).Cells
        Set rangecells5 = Sheet8.Range("loanAmt.24b" & i).Cells
        Set rangecells6 = Sheet8.Range("loanOutstanding.24b" & i).Cells
        Set rangecells7 = Sheet8.Range("Intrst.24b" & i).Cells
        
        'Added by Shrutika(03/06/2025) SIT-93666
        'setTblinfo_hpcoindex (i)
        setTblinfo_24bBankorInst (i)
        setTblinfo_24bankname (i)   'added by Chetan C M for AY 2025-26
        setTblinfo_24bLoan (i)
        setTblinfo_24bdate (i)
        setTblinfo_24bAmount (i)
        setTblinfo_24bloanout (i)
        setTblinfo_24bInt (i)
        
        'Added by Shrutika(03/06/2025) SIT-93666
       end_hpco_24d = Application.WorksheetFunction.Max(0, end_hpco_24b, end_hpco_24c, end_hpco_24e, end_hpco_24f, end_hpco_24k, end_hpco_24s, end_hpco_24r)
        
'        ReDim BankorInst_24b(end_hprptfrm, end_hpco_24b)
'        ReDim BankName_24B(end_hprptfrm, end_hpco_24b)
'        ReDim AccntNum_24b(end_hprptfrm, end_hpco_24b)
'        ReDim LoanDate_24B(end_hprptfrm, end_hpco_24b)
'        ReDim LoanAmt_24b(end_hprptfrm, end_hpco_24b)
'        ReDim LoanOutStanding_24B(end_hprptfrm, end_hpco_24b)
'        ReDim Intrst_24B(end_hprptfrm, end_hpco_24b)
        

        ReDim BankorInst_24b(end_hprptfrm, end_hpco_24d)
        ReDim BankName_24B(end_hprptfrm, end_hpco_24d)
        ReDim AccntNum_24b(end_hprptfrm, end_hpco_24d)
        ReDim LoanDate_24B(end_hprptfrm, end_hpco_24d)
        ReDim LoanAmt_24b(end_hprptfrm, end_hpco_24d)
        ReDim LoanOutStanding_24B(end_hprptfrm, end_hpco_24d)
        ReDim Intrst_24B(end_hprptfrm, end_hpco_24d)
        
       
        'For j = 1 To end_hpco_24b
         For j = 1 To end_hpco_24d
                BankorInst_24b(i, j) = rangecells1.item(j).value
                BankName_24B(i, j) = rangecells2.item(j).value
                AccntNum_24b(i, j) = rangecells3.item(j).value
                LoanDate_24B(i, j) = rangecells4.item(j).text
                LoanAmt_24b(i, j) = rangecells5.item(j).text
                LoanOutStanding_24B(i, j) = rangecells6.item(j).text
                Intrst_24B(i, j) = rangecells7.item(j).text
                
                'Added by Shrutika(03/06/2025) SIT-93666
                 If (BankorInst_24b(i, j) = "(Select)") Or (BankorInst_24b(i, j) = "") Then
                    msgbox_hprptfrm ("*""Please select dropdown from ""Loan taken from"" in 24(b) schedule."" at Sr. No " & j & " in Table " & i & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                '--------------------------
                
                ''Name of the Bank
                 If Not chkCompulsory(BankorInst_24b(i, j) <> "Select") Then
                    msgbox_hprptfrm ("*""Please select dropdown from ""Loan taken from"" in 24(b) schedule."" at Sr. No " & j & " in table " & i & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Not chkCompulsory(BankName_24B(i, j)) Then
                    msgbox_hprptfrm ("*""Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Len(BankName_24B(i, j)) > 125 Then
                    msgbox_hprptfrm ("* Name of the Bank/ Institution/Person at Sr. No " & j & " in Sheet 24(b) should be less than or equal to 125 characters.") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                
                If Not checkfieldSuperSpecialcharacter(BankName_24B(i, j)) Then
                    msgbox_hprptfrm ("* Name of the Bank/ Institution/Person in schedule 24b at Sl.no. " & j & " should not Contain <, >, characters.") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                '--Acc number
                If Not chkCompulsory(AccntNum_24b(i, j)) Then
                   ' msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr. No " & j & "") & Chr(13)
                    msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13) 'SIT-93563 updated by shrutika
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Len(AccntNum_24b(i, j)) > 20 Then
                    msgbox_hprptfrm ("* Loan Account number  at Sr. No " & j & " in schedule 24b less than 20 characters.") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                'Added by Aavula Naresh SIT-93662
                If UCase(BankorInst_24b(i, j)) = UCase("Bank ") Then
                    If Not ValidateBankAccountNumber_24b(AccntNum_24b(i, j)) Then
'                        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & j & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                    End If
                End If
                
                If UCase(BankorInst_24b(i, j)) = UCase("Other than Bank") Then
                    If Not checkfieldspecialcharacter1(AccntNum_24b(i, j)) Then
                        '            MsgBox_24b = MsgBox_24b & "* Loan Account number of the Bank / Institution at Sr.No " & j & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
                        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & j & " in table " & i & "") & Chr(13)
                        Validate24B_Table = False
                        Exit Function
                    End If
                End If
             ''date
             If Not chkCompulsory(LoanDate_24B(i, j)) Then
                 msgbox_hprptfrm ("* ""Please provide Date of sanction of Loan in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
             If Len(LoanDate_24B(i, j)) > 10 Then
              msgbox_hprptfrm ("* Date of sanction of Loan  at Sr. No " & j & " in schedule 24b less than 10 characters.") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            'CheckDateddmmyyyy
             If Not CheckDateddmmyyyy(LoanDate_24B(i, j)) Then
                msgbox_hprptfrm ("* ""Please enter date in valid format"" at Sr. No " & j & ".") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
'Konda updated on 19-05-2026--SIT-119907
'            If Not ChkMaxDate_24b(Trim(LoanDate_24B(i, j)), "31-03-2025") Then
'                msgbox_hprptfrm ("* Date can not be after 31/03/2025 at Sr. No " & j & ".") & Chr(13)
'                Validate24B_Table = False
'              Exit Function
'            End If
            
            If Not ChkMaxDate_24b(Trim(LoanDate_24B(i, j)), "31-03-2026") Then
                msgbox_hprptfrm ("* Date can not be after 31/03/2026 at Sr. No " & j & ".") & Chr(13)
                Validate24B_Table = False
              Exit Function
            End If
'================================
            '---Loan amount
            If Not chkCompulsory(LoanAmt_24b(i, j)) Then
            'Ayush_05/05/2025
            '             MsgBox_24b = MsgBox_24b + "* ""Please provide Total Loan taken in schedule 24(b)"" at Sr. No " & j & "" & Chr(13)
                    msgbox_hprptfrm ("* ""Please provide Total amount of loan in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            '         If Len(LoanAmt_24b(i,j)) > 14 Then
            '          MsgBox_24b = MsgBox_24b + "* Loan Amount  at Sr. No " & j & " in Sheet 24b less than 15 characters." & Chr(13)
            '            Validate24B_Table = False
            '            Exit Function
            '        End If
            
            If Not IsNumeric(LoanAmt_24b(i, j)) Then
                msgbox_hprptfrm ("* Loan amount at Sr. No  " & j & "  in schedule 24b should be Numeric value") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanAmt_24b(i, j) > 99999999999999# Then
                msgbox_hprptfrm ("* Loan amount at Sr. No  " & j & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanAmt_24b(i, j) < 0 Or LoanAmt_24b(i, j) = 0 Then
            'Ayush_05/05/2025
            '            MsgBox_24b = MsgBox_24b & "* Total loan taken should be more than 0 in schedule 24(b) at Sr. No  " & j & "" & Chr(13)
               msgbox_hprptfrm ("* Total amount of loan should be more than 0 in schedule 24(b) at Sr. No  " & j & " in table " & i & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            '--Outstanding
            If Not chkCompulsory(LoanOutStanding_24B(i, j)) Then
            ' Ayush_05/05/2025
            '             MsgBox_24b = MsgBox_24b + "* ""Loan outstanding as on 31-03-2025 is mandatory in schedule 24(b)"" at Sr. No " & j & "" & Chr(13)
                   ' msgbox_hprptfrm ("* ""Loan outstanding as on last date of financial year is mandatory in schedule 24(b)"" at Sr. No " & j & "") & Chr(13)
                'msgbox_hprptfrm ("* ""Loan outstanding as on last date of finacial year is mandatory in schedule 24(b). Please enter 0 in case the entire loan is repaid during the year"" at Sr. No " & j & " in table " & i & "") & Chr(13) '(SIT - 93574)
                
                
     'SIT-95355 BY SAI ON 26/06/2025
                msgbox_hprptfrm ("* ""Loan outstanding as on last date of financial year is mandatory in Table Section 24(b). Please enter 0 in case the entire loan is repaid during the year"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                
                Validate24B_Table = False
                Exit Function
            End If
            '         If Len(LoanOutstanding_24b(i)) > 14 Then
            '          MsgBox_24b = MsgBox_24b + "* Loan Outstanding  at Sr. No " & j & " in Sheet 24b less than 15 characters." & Chr(13)
            '            Validate24B_Table = False
            '            Exit Function
            '        End If
            
            If Not IsNumeric(LoanOutStanding_24B(i, j)) Then
                msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & j & "  in schedule 24b should be Numeric value") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanOutStanding_24B(i, j) > 99999999999999# Then
                msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & j & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanOutStanding_24B(i, j) < 0 Then
            'Ayush_05/05/2025
            '            MsgBox_24b = MsgBox_24b & "* Loan outstanding as on 31.03.2025 can't be less than 0 in schedule 24(b) at Sr. No  " & j & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
                 msgbox_hprptfrm ("* Loan outstanding as on last date of finacial year can't be less than 0 in schedule 24(b) at Sr. No  " & j & ". You may please enter as 0 if it become negative as result of excess payment.") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
                
            '--Intrest
            If Not chkCompulsory(Intrst_24B(i, j)) Then
                  ' msgbox_hprptfrm ("* ""Please provide Interest u/s 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
             msgbox_hprptfrm ("* ""Please provide Interest on Borrowed capital u/s 24(b)""") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            '         If Len(Intrst_24b(i)) > 15 Then
            '          MsgBox_24b = MsgBox_24b + "* Interest  at Sr. No " & j & " in Sheet 24b less than 15 characters." & Chr(13)
            '            Validate24B_Table = False
            '            Exit Function
            '        End If
            
            If Not IsNumeric(Intrst_24B(i, j)) Then
               msgbox_hprptfrm ("* Interest at Sr. No  " & j & "  in schedule 24b should be Numeric value") & Chr(13)
                
                Validate24B_Table = False
                Exit Function
            End If
            
            If Intrst_24B(i, j) > 99999999999999# Then
                msgbox_hprptfrm ("* Interest at Sr. No  " & j & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            'Interest u/s 24(b) should be more than 0 in schedule 24(b)
            
            If Intrst_24B(i, j) < 0 Or Intrst_24B(i, j) = 0 Then
               ' msgbox_hprptfrm ("* Interest u/s 24(b) should be more than 0 in schedule 24(b) at Sr. No  " & j & " in table " & i & "") & Chr(13)
               ' msgbox_hprptfrm ("* ""Please provide Interest u/s 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13) 'Added by Shrutika(SIT-93575
                msgbox_hprptfrm ("* ""Please provide Interest on Borrowed capital u/s 24(b)""") & Chr(13) 'Added by Shrutika(Ver 0.6)
                Validate24B_Table = False
                Exit Function
            End If
        Next
        
         If (Len(Sheet8.Range("TotAmt.24b" & i).value) > 14) Then
            msgbox_hprptfrm ("*  Total of interest on borrowed capital u/s 24(b) cannot exceed 14 Digits.") & Chr(13)
            Validate24B_Table = False
            Exit Function
         End If
            
        
        
    Next
endlin1:
End Function


