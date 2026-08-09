Attribute VB_Name = "mdHouseProperty"
Option Explicit

Dim sPassword As Variant
 
Public rngname_hpco As Variant
Public rngname_hpco2 As Variant
Public end_hpco As Variant
Public end_hpco2 As Variant
 

Public end_hpco_24b As Variant
Public rngname_hpco_24b As Variant
Public end_hpco_24a As Variant

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
Public CoOwnerPANCHK_HP As Variant
'Ayush_06/02/2026
Public CowOwned1_HP As Variant


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

Sub cmdNext_Click_HP()
Sheet3.Activate
End Sub
Sub cmdPrev_Click_HP()
Sheet1.Activate
End Sub

Sub Cmd_AddCo_Owners_Click()
AddPropertyCoOWners
End Sub

Sub AddPropertyCoOWners(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet22.Activate
'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   Sheet22.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value
    
        
        mdCommon.DefinedgridNameRange = "HP.Co.Srno||HP.Co.Name||HP.Co.Pan||HP.Co.Aadhaar||HP.Co.Share||Co_OwnnerBlock"
        ActiveCellRange = mdCommon.searchLastRow("HP.Co.Srno" & te)
        
        If te > 1 Then
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   Sheet22.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
        newfrmsize = Sheet22.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet22.Range("NumRowsPropertyBlock").value = newfrmsize
    Next

endline:
   
'----------------Lock Password-------------------START---
   Sheet22.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
Sub AddPropertyTenant(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheet22.Activate

'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   Sheet22.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value


        mdCommon.DefinedgridNameRange = "HP.NameofTenant||HP.PANofTenant||HP.AadharofTenant||HP.TANofTenant"
        ActiveCellRange = mdCommon.searchLastRow("HP.NameofTenant" & te)

        If te > 1 Then
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next



'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   Sheet22.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
        newfrmsize = Sheet22.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet22.Range("NumRowsPropertyBlock").value = newfrmsize
    Next


endline:

'----------------Lock Password-------------------START---
   Sheet22.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'--end

Function ValidateTenantPan(tenantPan As String) As Boolean
On Error Resume Next
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
Function ValidateTenantPan1(tenantPan As String) As Boolean
On Error Resume Next
    ValidateTenantPan1 = True
    If Len(tenantPan) > 0 Then
        If Not ChkAlphabet(Mid(tenantPan, 1, 1)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 2, 1)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 3, 1)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If

        If Not ChkAlphabet(Mid(tenantPan, 4, 1)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If

        If Not ChkAlphabet(Mid(tenantPan, 5, 1)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(tenantPan, 6, 4)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 10, 1)) Then
            ValidateTenantPan1 = False
            Exit Function
        End If
    End If
End Function

Sub AddRows_hpco()
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim te As Long
    Dim i As Long
    Dim frmsize_hprptfrmnew As Long
    Dim newfrmsize_hprptfrm As Long
    
    newrngname = rngname_hpco
    For te = 1 To Sheet22.Range("cntr.hprptfrm").value
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

    sPassword = mdCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    Application.EnableEvents = False
    For i = 1 To Sheet22.Range("cntr.hprptfrm").value
        newfrmsize_hprptfrm = Sheet22.Range("hprptfrm.size").value
        newfrmsize_hprptfrm = newfrmsize_hprptfrm + numberofrows
        Sheet22.Range("hprptfrm.size").value = newfrmsize_hprptfrm
    Next
    Application.EnableEvents = True
    Sheet8.Protect Password:=sPassword
End Sub

Sub ValidateSheetHPClick()
ValidateSheetHouseProperty
If msgValidateSheetHP = "" Then fmsgboxStatus "Sheet House Property (HP) is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheetHouseProperty()
subProcCaption = "Validating HP"
    If Not ValidatesheetHP Then
        Sheet22.Activate
        fmsgbox msgValidateSheetHP ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub
Function GetLetOut() As Variant
    If ValidateifLetOut_HP Then
        GetLetOut = ifLetOut_HP
    End If
End Function


Function msgbox_hprptfrm(strmsg As String) As String
     msgValidateSheetHP = msgValidateSheetHP & strmsg & Chr(13)
End Function
Function msgbox_HP(strmsg As String) As String
     msgValidateSheetHP = msgValidateSheetHP & strmsg & Chr(13)
End Function

Function ValidatesheetHP() As Boolean
    ValidatesheetHP = True
    
    
   ' AddTotIncomeUnderHouseProperty_1
    If Not ValidateAddrDetail_HP() Then ValidatesheetHP = False
    If Not ValidateCoName_HP() Then ValidatesheetHP = False
    If Not Validate24B_Table() Then ValidatesheetHP = False
    
    
    Dim te As Long
    Dim de As Long
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value
        If Sheet22.Range("HP.RentOfEarlierYrSec_AandAA" & te).value > 0 Then
          If Sheet22.Range("HP.AddrDetail" & te).value = "" Then
          msgbox_hprptfrm ("* ""Address of property " & te & " is mandatory in schedule HP.""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
    Next
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value
        If Sheet22.Range("HP.TotalUnrealizedAndTax" & te).value > 0 Or Sheet22.Range("HP.BalanceALV" & te).value > 0 Or Sheet22.Range("HP.IncomeOfHPInOwnHand" & te).value > 0 Or Sheet22.Range("HP.TotalDeduct" & te).value > 0 Or Sheet22.Range("HP.IncomeOfHP" & te).value > 0 Then
          If Sheet22.Range("HP.AddrDetail" & te).value = "" Then
          msgbox_hprptfrm ("* ""Address of property " & te & " is mandatory in schedule HP.""")
          
            ValidatesheetHP = False
            Exit Function
        End If
        End If
    Next
     For de = 1 To Sheet22.Range("PropertySectionCOunt").value
       
        If Sheet22.Range("HP.TotalUnrealizedAndTax" & de) > 99999999999999# Then
            msgbox_hprptfrm "* Total (1b + 1c) in schedule House Property cannot exceed 14 digits" & Chr(13)
            ValidatesheetHP = False
            Exit Function
        End If
        Next
   For te = 1 To Sheet22.Range("PropertySectionCOunt").value
    If Not Sheet22.Range("HP.StateCode" & te).value = "" Or Sheet22.Range("HP.StateCode" & te).value = "(Select)" Or Sheet22.Range("HP.OwnerProperty" & te).value = "" Or Sheet22.Range("HP.CoOwnedYN" & te).value = "" Or Sheet22.Range("HP.AddrDetail" & te).value = "" Then
    If Not isdropdownblank(Sheet22.Range("HP.StateCode" & te).value) _
    Or Not isdropdownblank(Sheet22.Range("HP.OwnerProperty" & te).value) _
    Or Not isdropdownblank(Sheet22.Range("HP.CoOwnedYN" & te).value) _
    Or Sheet22.Range("HP.AddrDetail" & te).value <> "" Then

         If Sheet22.Range("HP.AddrDetail" & te).value = "" Then
            msgbox_hprptfrm ("* ""Address of property " & te & " is mandatory in schedule HP.""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
        End If
        
    Next
    If (Len(Sheet22.Range("HP.AddrDetail1")) > 0) Then
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
        If Not ValidateRentArearsSec25BAfter30pcDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateTotalIncomeChargeableUnHP_HP() Then ValidatesheetHP = False
        'Ayush_11/02/2026
        If Not ValidateCoownerName1_HP() Then ValidatesheetHP = False
        If Not ValidatePercenShare1_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnerPAN1_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnerAadhar_HP() Then ValidatesheetHP = False
        If Not ValidateSharePercent_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnedYN_HP() Then ValidatesheetHP = False
        If Not ValidateCoownerRules_HP() Then ValidatesheetHP = False
        If Not ValidatesharesRules_HP() Then ValidatesheetHP = False

            If ValidatesheetHP Then
            Dim noofsop, temp1, hpi, SPI, ind As Long
            Dim letoutarr As Variant
            letoutarr = GetLetOut
            noofsop = 0
            For hpi = 1 To UBound(letoutarr)
                If Mid(Sheet22.Range("HP.ifLetOut" & hpi).value, 1, 1) = "S" Then
                 noofsop = noofsop + 1
                    If (Sheet22.Range("HP.AnnualLetableValue" & hpi).value > 0) Then
                    Sheet22.Activate
                        fmsgboxStatus "Property " & hpi & " is deemed to be letout as its A.L.V. > 0" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                    
                     If Mid(Sheet22.Range("HP.ifLetOut" & hpi).value, 1, 1) = "S" Then
                          If noofsop > 2 Then
                               msgbox_hprptfrm "* Self occupied property cannot be selected more than twice from house property is declared in Schedule HP." & Chr(13)
                                ValidatesheetHP = False
                          End If
                          temp1 = Sheet22.Range("HP.IntOnBorwCap" & hpi) + temp1
                          If temp1 > 200000 Then
                            msgbox_hprptfrm "* F for category self-occupied in Schedule HP." & Chr(13)
                                ValidatesheetHP = False
                          End If
                        
                        End If
                Else
                    If letoutarr(hpi) = "" Then
                        If (Sheet22.Range("HP.AnnualLetableValue" & hpi).value = 0) Then
                            noofsop = noofsop + 1
                        Else
                        End If
                    Else
                    End If
                End If
            Next
    
            letoutarr = GetLetOut
            For SPI = 1 To UBound(letoutarr)
                If (letoutarr(SPI) = "S" Or letoutarr(SPI) = "" Or _
                    Sheet22.Range("HP.AnnualLetableValue" & SPI).value = 0) Then
                    ind = SPI
                    Dim count As Variant
                    Dim tempval As Range
                    count = 0
                If letoutarr(SPI) = "S" Then
                  For Each tempval In Sheet22.Range("HP.NameofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                    End If
                  Next
                  If count > 0 Then
                        msgbox_hprptfrm ("* If Self Occupied then Name of Tenant in schedule HP at Block " & ind & " is not required")
                        ValidatesheetHP = False
                    End If
                    
                  count = 0
                  For Each tempval In Sheet22.Range("HP.PANofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then PAN of Tenant in schedule HP at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        
                         count = 0
                  For Each tempval In Sheet22.Range("HP.AadharofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then Aadhaar of Tenant in schedule HP at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        
                        
                    count = 0
                  For Each tempval In Sheet22.Range("HP.TANofTenant" & ind).Cells
                    If Not tempval.value = "" Then
                        count = count + 1
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then TAN of Tenant in schedule HP at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        count = 0
                
                    End If
                    If val(Sheet22.Range("HP.IntOnBorwCap" & ind)) > 200000 Then
                        If Mid(Sheet22.Range("HP.ifLetOut" & ind), 1, 1) = "D" Or Mid(Sheet22.Range("HP.ifLetOut" & ind), 1, 1) = "L" Then
                        
                        Else
                            Sheet22.Activate
                            fmsgboxStatus "If Not Letout (OR ALV=0) then Interest Payable on borrowed capital should not exceed 2 Lacs" ' , vbOKOnly, "Error!"
                            CloseMsg
                        End If
                    End If
                
                    If Sheet22.Range("HP.LocalTaxes" & ind).value > 0 Then
                    Sheet22.Activate
                        fmsgboxStatus "Tax paid to local authorities can be claimed only if income from house property is declared" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                
                    If val(Sheet22.Range("HP.RentNotRealized" & ind)) > 0 Then
                    Sheet22.Activate
                        fmsgboxStatus "Rent not realized cannot exceed Gross rent received or receivable or letable value in schedule HP" ', vbOKOnly, "Error!"
                        
                        CloseMsg
                    End If
                End If
                            
                If (letoutarr(SPI) <> "S") Then
                    ind = SPI
                    If val(Sheet22.Range("HP.RentNotRealized" & ind)) > Sheet22.Range("HP.AnnualLetableValue" & ind).value Then
                    Sheet22.Activate
                        fmsgboxStatus "Rent not realized cannot exceed Gross Rent received or receivable or letable value in schedule HP" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                    If letoutarr(SPI) = "L" Then
                        If Sheet22.Range("HP.AnnualLetableValue" & ind) <= 0 Then
                        Sheet22.Activate
'                            fmsgboxStatus "If Type of House Property is Let Out/ Deemed Let Out then Gross Rent received or receivable or lettable value field cannot be zero/ blank"
                            fmsgboxStatus " If ""Type of House Property"" is ""Let Out/ Deemed Let Out"" then ""Gross rent received/ receivable/ lettable value during the year should be more than zero""."
                            
                            CloseMsg
                        End If
                    End If
                    
                    If letoutarr(ind) = "D" Then
                        For Each tempval In Sheet22.Range("HP.NameofTenant" & ind).Cells
                            If Not tempval.value = "" Then
                            count = count + 1
                            End If
                        Next
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then Name of Tenant in schedule HP at Block " & ind & " is not required")
                            ValidatesheetHP = False
                            End If
                        
                        count = 0
                        For Each tempval In Sheet22.Range("HP.PANofTenant" & ind).Cells
                            If Not tempval.value = "" Then
                            count = count + 1
                            End If
                        Next
                        
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then PAN of Tenant in schedule HP at Block " & ind & " is not required")
                            ValidatesheetHP = False
                        End If
                        count = 0
                        For Each tempval In Sheet22.Range("HP.TANofTenant" & ind).Cells
                            If Not tempval.value = "" Then
                            count = count + 1
                            End If
                        Next
                        
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then TAN of Tenant in schedule HP at Block " & ind & " is not required")
                            ValidatesheetHP = False
                        End If
                        count = 0
                        End If
                    End If
            Next
        End If
    Else
        If Not ValidateRentArearsSec25BAfter30pcDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateTotalIncomeChargeableUnHP_HP() Then ValidatesheetHP = False
    End If
    
    If Not ValidateMandatorytable1 Then ValidatesheetHP = False
    
    'Ayush_19/03/2026
    If ValidatesheetHP = True Then
    If Sheet22.Range("HP.AddrDetail1").value = "" Or Sheet22.Range("HP.CityOrTownOrDistrict1").value = "" Or (Sheet22.Range("HP.StateCode1").value = "(Select)" Or Sheet22.Range("HP.StateCode1").value = "") Or (Sheet22.Range("HP.CountryCode1").value = "" Or Sheet22.Range("HP.CountryCode1").value = "(Select)") Or Sheet22.Range("HP.ifLetOut1").value = "(Select)" Or Sheet22.Range("HP.ifLetOut1").value = "" Then
'          If Sheet22.Range("HP.AddrDetail1").value = "" Or Sheet22.Range("HP.CityOrTownOrDistrict1").value = "" Or Sheet22.Range("HP.StateCode1").value = "(Select)" Or Sheet22.Range("HP.StateCode1").value = "" Or Sheet22.Range("HP.CountryCode1").value = "91-INDIA" Or Sheet22.Range("HP.CountryCode1").value = "" Or Sheet22.Range("HP.ifLetOut1").value = "(Select)" Or Sheet22.Range("HP.ifLetOut1").value = "" Then
  
        If Sheet22.Range("Block_1B").value < 8 Then
            If Sheet22.Range("Final_1B") = "Data" Then
                msgbox_hprptfrm ("* Enter all mandatory fields of Block 1B in Schedule HP.")
                ValidatesheetHP = False
                Exit Function
            End If
        End If
    End If
End If

End Function

    Sub setTblinfo_hprptfrm()
    Dim te As Long
    Dim ccount As Long
    
    frmsize_hprptfrm = Sheet22.Range("NumRowsPropertyBlock").value
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value
        If Not Sheet22.Range("HP.AddrDetail" & te).value = "" Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    end_hprptfrm = ccount
    cntrRng_hprptfrm = "PropertySectionCOunt"
    frmRngname_hprptfrm = "hprptfrm"
    rngname_hprptfrm = "HP.AddrDetail1;HP.CityOrTownOrDistrict1;HP.CountryCode1;HP.StateCode1;HP.ZipCode1;HP.PinCode1;HP.OwnerProperty1;HP.OwnerPropertyDescription1;HP.CoOwnedYN1;HP.SharePercent1;HP.Co.Srno1;HP.Co.Name1;HP.Co.PAN1;HP.Co.Aadhaar1;HP.Co.Share1;HP.ifLetOut1;HP.NameofTenant1;HP.PANofTenant1;HP.AadharofTenant1;HP.TANofTenant1;HP.AnnualLetableValue1;HP.RentNotRealized1;HP.LocalTaxes1;HP.TotalUnrealizedAndTax1;HP.BalanceALV1;HP.ThirtyPercentOfBalance1;HP.IntOnBorwCap1;HP.TotalDeduct1;HP.RentOfEarlierYrSec_AandAA1;HP.IncomeOfHP1;HP.IncomeOfHPInOwnHand1;Co_OwnnerBlock1;LoanfrmBankOrInstitute.24b1;bankName.24b1;loanAccNum.24b1;loanDate.24b1;loanAmt.24b1;loanOutstanding.24b1;Intrst.24b1;TotAmt.24b1;Combination_24B1"
 End Sub

    
 Function ValidateAddrDetail_HP() As Boolean
    Dim i As Long
    ValidateAddrDetail_HP = True
    setTblinfo_hprptfrm
    
    ReDim AddrDetail_HP(end_hprptfrm)
    noOfProcessSub = end_hprptfrm
    For i = 1 To end_hprptfrm
        AddrDetail_HP(i) = Sheet22.Range("HP.AddrDetail" & i).value

        If Not CheckSpecialCharacter4New(AddrDetail_HP(i)) Then
            msgbox_hprptfrm ("* AddrDetail" & i & "  characters < > are not allowed in schedule HP ")
            ValidateAddrDetail_HP = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Sub setTblinfo_hpcoindex(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet22.Range("hp.co.name" & myindex).count
    Set rangecells = Sheet22.Range("hp.co.name" & myindex).Cells
    
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
        sharepercent = Sheet22.Range("HP.SharePercent" & i).text
        Set rangecells1 = Sheet22.Range("HP.Co.Name" & i).Cells
        Set rangecells2 = Sheet22.Range("HP.Co.PAN" & i).Cells
        Set rangecells3 = Sheet22.Range("HP.Co.Share" & i).Cells
        Set rangecells4 = Sheet22.Range("HP.StateCode" & i).Cells
        Set rangecells5 = Sheet22.Range("HP.Co.Aadhaar" & i).Cells
        setTblinfo_hpcoindex (i)
        
        ReDim CoName_HP(end_hprptfrm, end_hpco)
        ReDim CoPAN_HP(end_hprptfrm, end_hpco)
        ReDim CoAadhar_HP(end_hprptfrm, end_hpco)
        ReDim CoShare_HP(end_hprptfrm, end_hpco)
        
        state_1 = rangecells4.value
        state_1 = Mid(state_1, 1, 2)
        
        cosharepercent = 0
        For j = 1 To end_hpco
            CoName_HP(i, j) = rangecells1.item(j).value
            CoPAN_HP(i, j) = rangecells2.item(j).value
            CoAadhar_HP(i, j) = rangecells5.item(j).value
            CoShare_HP(i, j) = rangecells3.item(j).text
            If CoShare_HP(i, j) <> "" Then
                CoShare_HP(i, j) = Round(CoShare_HP(i, j), 2)
            
                cosharepercent = cosharepercent + IIf(CoShare_HP(i, j) = "", 0, CoShare_HP(i, j))
            End If
            
            If Not ValidatePAN(CStr(CoPAN_HP(i, j))) Then
                msgbox_hprptfrm ("* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule HP") & Chr(13)
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidateAadhaar(CStr(CoAadhar_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid Aadhaar of other co-owner " & j & "  in Property " & i & "  with 12 digits schedule HP") & Chr(13)
                ValidateCoName_HP = False
                Exit Function
            End If
        Next
        
         If sharepercent < 0.01 Then
            msgbox_hprptfrm ("* ""Your percentage of share in co-owned property cannot be zero""") & Chr(13)
            ValidateCoName_HP = False
            Exit Function
        End If
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
        sharepercent = Sheet22.Range("HP.SharePercent" & i).text
        Set rangecells1 = Sheet22.Range("HP.Co.Name" & i).Cells
        Set rangecells2 = Sheet22.Range("HP.Co.PAN" & i).Cells
        Set rangecells3 = Sheet22.Range("HP.Co.Share" & i).Cells
        Set rangecells4 = Sheet22.Range("HP.StateCode" & i).Cells
        Set rangecells5 = Sheet22.Range("HP.Co.Aadhaar" & i).Cells
        
        
        setTblinfo_hpcoindex (i)
        setTblinfo_24bankname (i)
        ReDim CoName_HP(end_hprptfrm, end_hpco)
        ReDim CoPAN_HP(end_hprptfrm, end_hpco)
        ReDim CoAadhar_HP(end_hprptfrm, end_hpco)
        ReDim CoShare_HP(end_hprptfrm, end_hpco)
        
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
                msgbox_hprptfrm ("* Share of CO Owner " & j & "  in Property " & i & " is Mandatory schedule HP") & Chr(13)
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not checkfieldspecialcharacter(CoName_HP(i, j)) Then
                msgbox_hprptfrm ("* Name of CO Owner " & j & "  in Property " & i & "  cannot contain special characters schedule HP") & Chr(13)
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidatePAN(CStr(CoPAN_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid PAN of CO Owner " & j & "  in Property " & i & "  with 1st 5 alphabets, next 4 digits and last alphabet schedule HP") & Chr(13)
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidateAadhaar(CStr(CoAadhar_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid Aadhaar of CO Owner " & j & "  in Property " & i & "  with 12 digits schedule HP") & Chr(13)
                ValidateCoName_HP = False
                Exit Function
            End If
        Next
        
         If sharepercent < 0.01 Then
            msgbox_hprptfrm ("*Share Percentage of Owner and CO Owner must add to 100 in Property " & i & " schedule HP") & Chr(13)
            ValidateCoName_HP = False
            Exit Function
        End If
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
        
        CityOrTownOrDistrict_HP(i) = Sheet22.Range("HP.CityOrTownOrDistrict" & i).value
        If Not chkCompulsory(CityOrTownOrDistrict_HP(i)) Then
            msgbox_hprptfrm ("* ""Town/City is mandatory in tab: Sl no B2""")
            
            ValidateCityOrTownOrDistrict_HP = False
            Exit Function
        End If
        If Not CheckSpecialCharacter4New(CityOrTownOrDistrict_HP(i)) Then
            msgbox_hprptfrm ("* CityOrTownOrDistrict" & i & "   characters < >  are not allowed in schedule HP") & Chr(13)
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
           
        StateCode_HP(i) = Sheet22.Range("HP.StateCode" & i).value
        StateCode_HP(i) = Mid(StateCode_HP(i), 1, 2)
        If isdropdownblank(StateCode_HP(i)) Then
            StateCode_HP(i) = ""
        End If
        
        If Not chkCompulsory(StateCode_HP(i)) Then
            msgbox_hprptfrm ("* ""State of property  is mandatory in schedule HP.""")
            
            ValidateStateCode_HP = False
            Exit Function
        End If
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
           
        CountryCode_HP(i) = Sheet22.Range("HP.CountryCode" & i).value
        Country1 = Sheet22.Range("HP.CountryCode" & i).value
        If isdropdownblank(CountryCode_HP(i)) Then
            CountryCode_HP(i) = ""
        End If
        
        If Not chkCompulsory(CountryCode_HP(i)) Then
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

        Country1 = Sheet22.Range("HP.CountryCode" & i).value
        state1 = Sheet22.Range("HP.StateCode" & i).value
        
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
                msgbox_hprptfrm ("* ""Country cannot be other than India as you have selected an Indian state " & i & " in schedule HP""")
                ValidateCountryStateCode_HP = False
                Exit Function
            End If
            If (Country1 = "91" And state1 = "99") Then
                msgbox_hprptfrm ("* ""Country cannot be India as you have selected a Foreign state " & i & " in schedule HP""")
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
        ZipCode_HP(i) = Sheet22.Range("HP.ZipCode" & i).value
        If CountryCode_HP(i) <> "91" Then
            If Not chkCompulsory(ZipCode_HP(i)) Then
                msgbox_hprptfrm ("* ""Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX  in schedule HP.""")
                
                ValidateZipCode_HP = False
                Exit Function
            End If
            If Not CheckSpecialCharacter4New(ZipCode_HP(i)) Then
                msgbox_hprptfrm ("* ZipCode" & i & " characters < >  are not allowed in schedule HP")
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
        PinCode_HP(i) = Sheet22.Range("HP.PinCode" & i).value
        If CountryCode_HP(i) = "91" Then
            If Not chkCompulsory(PinCode_HP(i)) Then
               msgbox_hprptfrm ("* ""Pin code of Property is mandatory in schedule HP """) & Chr(13)
                ValidatePinCode_HP = False
                Exit Function
            End If
            If Not chkNumeric(PinCode_HP(i)) Then
                msgbox_hprptfrm ("* PinCode" & i & " only digits 0 to 9 allowed  in schedule HP")
                ValidatePinCode_HP = False
                Exit Function
            End If
        Else
            PinCode_HP(i) = ""
        End If
         If PinCode_HP(i) <> "" Then
         Dim PIN_targetadd1, state_targetadd1 As String
         
     
     PIN_targetadd1 = Replace(Sheet22.Range("HP.PinCode" & i).Address, "$", "")
     state_targetadd1 = Replace(PIN_targetadd1, "J", "H")
      
                             Dim ws1 As Worksheet
                             Set ws1 = Worksheets("HP")
         If Not PINstate_ModualValidation(ws1, PIN_targetadd1, state_targetadd1) Then
                Sheet22.Range("HP.PinCode" & i).value = ""
                msgbox_hprptfrm ("* ""Pin code of Property is mandatory in schedule HP """) & Chr(13)
                ValidatePinCode_HP = False
                Exit Function
         End If

    
    End If
    Next
End Function

'Ankita_20/01/2026=

Function PINstate_ModualValidation(sheetname2 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    PINstate_ModualValidation = True
    
    Dim PinCode As Range
    
    sheetname2.Activate
        
            If Range(pin_trgt_adrs).value <> "" Then
                Dim state1 As String
                
                state1 = Application.IsError(Application.VLookup(Range(pin_trgt_adrs).value, Sheet8.Range("All_PinCode_V"), 2, False))
                
                If state1 = True Then
                    PINstate_ModualValidation = False
                Else
                    state1 = Application.VLookup(Range(pin_trgt_adrs).value, Sheet8.Range("All_PinCode_V"), 2, False)
                    If Not (state1 = UCase(Mid(Range(state_trgt_adrs).value, 4)) Or mdCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value)) Then
                    PINstate_ModualValidation = False
                    End If
                 End If
                End If
                   
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
        OwnerProperty_HP(i) = Sheet22.Range("HP.OwnerProperty" & i).value
        OwnerPropertyDescription_HP(i) = Sheet22.Range("HP.OwnerPropertyDescription" & i).value
        
            If isdropdownblank(OwnerProperty_HP(i)) Then
                msgbox_hprptfrm ("* ""Please select dropdown from owner of the Property  in schedule HP.""")
                
                ValidateOwnerProperty_HP = False
                Exit Function
            End If
        If Mid(OwnerProperty_HP(i), 1, 1) = "O" Then
           '     msgbox_hprptfrm ("* ""In case of others, mandatory to enter owner in textbox in schedule HP""")
                If Not chkCompulsory(OwnerPropertyDescription_HP(i)) Then
                'Ayush_SIT-109560
                   msgbox_hprptfrm ("* ""In case of others, mandatory to enter owner in textbox in schedule HP""")
                  ValidateOwnerProperty_HP = False
                 Exit Function
                End If
              
              If Len(OwnerPropertyDescription_HP(i)) > 50 Then
                 msgbox_hprptfrm ("* Description for Owner of the Property" & i & "  Cannot be more than 50 characters in schedule HP")
                 ValidateOwnerProperty_HP = False
                 Exit Function
                End If
            
             If Not checkfieldspecialcharacter(OwnerPropertyDescription_HP(i)) Then
                msgbox_hprptfrm ("* Description for Owner of the Property" & i & "  characters < > & ' " & Chr(34) & " are not allowed in schedule HP") & Chr(13)
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
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        ifLetOut_HP(i) = Sheet22.Range("HP.ifLetOut" & i).value
        ifLetOut_HP(i) = Mid(ifLetOut_HP(i), 1, 1)
        
        If isdropdownblank(ifLetOut_HP(i)) Then
            ifLetOut_HP(i) = ""
        End If
        
        If Not chkCompulsory(ifLetOut_HP(i)) Then
            msgbox_hprptfrm ("* ""Please select type of House Property in schedule HP """) & Chr(13)
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
        
        Set rangecells = Sheet22.Range("HP.NameofTenant" & i).Cells
        ReDim NameofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        count = 0
        For j = 1 To rangecells.Cells.count
        
        NameofTenant_HP(j) = rangecells.item(j).value
        If Len(NameofTenant_HP(j)) > 0 Then
            If Not CheckSpecialCharacter4New(NameofTenant_HP(j)) Then
                msgbox_hprptfrm ("* Name of Tenant" & j & "  characters < >  are not allowed  in schedule HP")
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
             msgbox_hprptfrm ("* ""Please enter name of Tenant in schedule HP""") & Chr(13)
             
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
        Set rangecells = Sheet22.Range("HP.PANofTenant" & i).Cells
        ReDim PANofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        
        For j = 1 To rangecells.Cells.count
        PANofTenant_HP(j) = rangecells.item(j).value
        
        tempPan = PANofTenant_HP(j)
        If Len(PANofTenant_HP(j)) > 0 Then
            If Not ValidatePAN(tempPan) Then
                  msgbox_hprptfrm ("* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP""") & Chr(13)
                ValidatePANofTenant_HP = False
                Exit Function
            End If
        End If
        
        Next
        
    Next
End Function

Function ValidatePAN(panentry As String) As Boolean
    ValidatePAN = True

    If Len(panentry) > 0 Then
        If Not IsNumeric(Mid(panentry, 6, 4)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 1, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 2, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 3, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 4, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 5, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 10, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
    End If
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
        Set rangecells = Sheet22.Range("HP.AadharofTenant" & i).Cells
        ReDim AadharofTenant_HP(rangecells.Cells.count)
        Dim j As Variant

        For j = 1 To rangecells.Cells.count
        AadharofTenant_HP(j) = rangecells.item(j).value

          If Len(AadharofTenant_HP(j)) > 0 Then
            If Len(AadharofTenant_HP(j)) <> 12 Then
                msgbox_hprptfrm ("*Aadhaar no of other Tenant " & j & " shall not exceed or less than 12 Digits in schedule HP") & Chr(13)
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If AadharofTenant_HP(j) = "000000000000" Then
                msgbox_hprptfrm ("*Invalid Aadhaar of Tenant should be 12 digits " & j & " shall not exceed or less than 12 Digits in schedule HP") & Chr(13)
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If AadharofTenant_HP(j) = "111111111111" Then
                msgbox_hprptfrm ("*Invalid Aadhaar of Tenant should be 12 digits " & j & " shall not exceed or less than 12 Digits in schedule HP") & Chr(13)
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If Not checkallfieldspecialcharacter(AadharofTenant_HP(j)) Then
                msgbox_hprptfrm ("* Aadhar of Tenant" & j & "  Special characters are not allowed in schedule HP")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If
            End If
        Next

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

Function ValidateAadhaar(Aadhaar As String) As Boolean
    ValidateAadhaar = True

   If Aadhaar <> "" Then
        If Not IsNumeric(Aadhaar) Then
            ValidateAadhaar = False
            Exit Function
        End If
    
        If Aadhaar = "000000000000" Then
            ValidateAadhaar = False
            Exit Function
        End If
        
        If Aadhaar = "111111111111" Then
            ValidateAadhaar = False
            Exit Function
        End If
            
        If Len(Aadhaar) <> 12 Then
            ValidateAadhaar = False
            Exit Function
        End If
        Dim i, j As Long
        Dim arr As Variant
        arr = Array(".", ",")
        For i = 1 To Len(Aadhaar)
        For j = 0 To UBound(arr)
        If Mid(Aadhaar, i, 1) = arr(j) Then
        ValidateAadhaar = False
            Exit Function
        End If
        Next
        Next
    End If
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
        Set rangecells = Sheet22.Range("HP.TANofTenant" & i).Cells
        ReDim TANofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        
        For j = 1 To rangecells.Cells.count
        TANofTenant_HP(j) = rangecells.item(j).value
        
        tempTan = TANofTenant_HP(j)
        If Len(TANofTenant_HP(j)) > 0 Then
        If Not (CheckTAN(tempTan) Or ValidateTenantPan(tempTan)) Then
                msgbox_hprptfrm ("* ""Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet or First 4 Alphabets, next 5 digits, then 1 Alphabet in schedule HP""")
                'By Bindu on 24th Feb 26
                msgbox_hprptfrm ("* ""Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet in case of TAN or First 4 Alphabets, next 5 digits, then 1 Alphabet in case of PAN in schedule HP""")

                
                ValidateTANofTenant_HP = False
                Exit Function
            End If
        End If
        
        Next
        
    Next
End Function

Function CheckTAN(TAN As Variant) As Boolean
On Error Resume Next


TAN = UCase(TAN)
    CheckTAN = True
    If Len(TAN) > 0 Then
        If Not ChkAlphabet(Mid(TAN, 1, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 2, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 3, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        
      If Not ChkAlphabet(Mid(TAN, 4, 1)) Then
            CheckTAN = False
            Exit Function
        End If
            
        If Not IsNumeric(Mid(TAN, 5, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 6, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 7, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 8, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(TAN, 9, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 10, 1)) Then
            CheckTAN = False
            Exit Function
        End If
    End If
End Function



Function CheckTAN1(TAN As Variant) As Boolean
On Error Resume Next


TAN = UCase(TAN)
    CheckTAN1 = True
    If Len(TAN) > 0 Then
        If Not ChkAlphabet(Mid(TAN, 1, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 2, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 3, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        
      If Not ChkAlphabet(Mid(TAN, 4, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
            
        If Not IsNumeric(Mid(TAN, 5, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 6, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 7, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 8, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(TAN, 9, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 10, 1)) Then
            CheckTAN1 = False
            Exit Function
        End If
    End If
End Function

Function CheckHousePropertyIncome(HP As Variant) As Boolean
On Error Resume Next
 CheckHousePropertyIncome = True
 If (HP > 200000) Then
 fmsgboxsmall "* Interest cannot be more than Rs. 2,00,000 for category self-occupied in schedule HP." ', vbOKOnly, "Error(s)!"
  CheckHousePropertyIncome = False
 Exit Function
 End If
End Function


Function ValidateAnnualLetableValue_HP() As Boolean
    Dim i As Long
    ValidateAnnualLetableValue_HP = True
    setTblinfo_hprptfrm
    
    ReDim AnnualLetableValue_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        AnnualLetableValue_HP(i) = Sheet22.Range("HP.AnnualLetableValue" & i).value
          If AnnualLetableValue_HP(i) = 0 And (Mid(ifLetOut_HP(i), 1, 1) = "D" Or Mid(ifLetOut_HP(i), 1, 1) = "L") Then
            msgbox_hprptfrm ("* If ""Type of House Property"" is ""Let Out/ Deemed Let Out"" then ""Gross rent received/ receivable/ lettable value during the year should be more than zero"".")
                            
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
        RentNotRealized_HP(i) = Sheet22.Range("HP.RentNotRealized" & i).value
    Next
End Function

Function ValidateLocalTaxes_HP() As Boolean
    Dim i As Long
    
    ValidateLocalTaxes_HP = True
    setTblinfo_hprptfrm
    ReDim LocalTaxes_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        LocalTaxes_HP(i) = Sheet22.Range("HP.LocalTaxes" & i).value
    Next
End Function

Function ValidateTotalUnrealizedAndTax_HP() As Boolean
    Dim i As Long
    
    ValidateTotalUnrealizedAndTax_HP = True
    setTblinfo_hprptfrm
    ReDim TotalUnrealizedAndTax_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        TotalUnrealizedAndTax_HP(i) = Sheet22.Range("HP.TotalUnrealizedAndTax" & i).value
    Next
End Function

Function ValidateBalanceALV_HP() As Boolean
    Dim i As Long
    
    ValidateBalanceALV_HP = True
    setTblinfo_hprptfrm
    ReDim BalanceALV_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        BalanceALV_HP(i) = Sheet22.Range("HP.BalanceALV" & i).value
    Next
End Function


Function ValidateThirtyPercentOfBalance_HP() As Boolean
    Dim i As Long
    
    ValidateThirtyPercentOfBalance_HP = True
    setTblinfo_hprptfrm
    ReDim ThirtyPercentOfBalance_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        ThirtyPercentOfBalance_HP(i) = Sheet22.Range("HP.ThirtyPercentOfBalance" & i).value
    Next
End Function

Function ValidateIntOnBorwCap_HP() As Boolean
    Dim i As Long
    
    ValidateIntOnBorwCap_HP = True
    setTblinfo_hprptfrm
    ReDim IntOnBorwCap_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        IntOnBorwCap_HP(i) = Sheet22.Range("HP.IntOnBorwCap" & i).value
    Next
End Function

Function ValidateTotalDeduct_HP() As Boolean
    Dim i As Long
    
    ValidateTotalDeduct_HP = True
    setTblinfo_hprptfrm
    ReDim TotalDeduct_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        TotalDeduct_HP(i) = Sheet22.Range("HP.TotalDeduct" & i).value
    Next
End Function

Function ValidateArrears_HP() As Boolean
    Dim i As Long
    
    ValidateArrears_HP = True
    setTblinfo_hprptfrm
    ReDim Arrears_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        Arrears_HP(i) = Sheet22.Range("HP.RentOfEarlierYrSec_AandAA" & i).value
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
        IncomeOfHP_HP(i) = Sheet22.Range("HP.IncomeOfHP" & i).value
        IncomeOfHPInOwnHand_HP(i) = Sheet22.Range("HP.IncomeOfHPInOwnHand" & i).value
    Next
End Function

Function ValidateRentArearsSec25BAfter30pcDeduct_HP() As Boolean
    ValidateRentArearsSec25BAfter30pcDeduct_HP = True
'    RentArearsSec25BAfter30pcDeduct_HP = Sheet22.Range("HP.RentArearsSec25BAfter30pcDeduct").value
End Function

Function ValidateTotalIncomeChargeableUnHP_HP() As Boolean
    ValidateTotalIncomeChargeableUnHP_HP = True
    TotalIncomeChargeableUnHP_HP = Sheet22.Range("HP.TotalIncomeChargeableUnHP").value
    
    If Len(TotalIncomeChargeableUnHP_HP) > 14 Then
        msgbox_hprptfrm ("* Total income chargeable cannot be greater than 14 digits")
        ValidateTotalIncomeChargeableUnHP_HP = False
    End If
    
End Function

'Function ValidateCoownerName1_HP() As Boolean
'    Dim i, j As Long
'    Dim rangecells As Range
'    ValidateCoownerName1_HP = True
'    setTblinfo_hprptfrm
'    ReDim CoOwnerName1_HP(end_hprptfrm)
'
'    For j = 1 To end_hprptfrm
'    Set rangecells = Sheet22.Range("HP.Co.Name" & j).Cells
'    For i = 1 To end_hprptfrm
'        CoOwnerName1_HP(i) = rangecells.item(i).value
'
'        If Not chkCompulsory(CoOwnerName1_HP(i)) Then
'             msgbox_hprptfrm ("* ""Please enter name of Tenant in schedule HP""") & Chr(13)
'
'            ValidateCoownerName1_HP = False
'            Exit Function
'        End If
'
'    Next
'    Next
'End Function
Function ValidateCoownerName1_HP() As Boolean
    Dim i, j As Long
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    ValidateCoownerName1_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnerName1_HP(end_hprptfrm)
    ReDim CowOwned1_HP(end_hprptfrm)
    
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet22.Range("HP.Co.Name" & j).Cells
    For i = 1 To end_hprptfrm
        CoOwnerName1_HP(i) = rangecells.item(i).value
     
        CowOwned1_HP(i) = Sheet22.Range("HP.CoOwnedYN" & i).value
        
        
        If Len(CoOwnerName1_HP(i)) > 0 Then
            If Not CheckSpecialCharacter4New(CoOwnerName1_HP(i)) Then
                msgbox_hprptfrm ("* Name of other co-owner characters < >  are not allowed  in schedule HP")
                ValidateCoownerName1_HP = False
                Exit Function
            End If
        End If
    
    
        
'        If Mid(CowOwned1_HP(i), 1, 1) = "Y" Then
'        If rangecells.item(i).Offset(0, 1).value <> "" Or rangecells.item(i).Offset(0, 2).value <> "" Or rangecells.item(i).Offset(0, 3).value <> "" Then
'        If Not chkCompulsory(CoOwnerName1_HP(i)) Then
'             msgbox_hprptfrm ("* ""Please enter the name of the co-owner.""") & Chr(13)
'
'            ValidateCoownerName1_HP = False
'            Exit Function
'        End If
'        End If
'        End If
        
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
    Set rangecells = Sheet22.Range("HP.Co.Share" & j).Cells
    For i = 1 To end_hprptfrm
        CoOwnerSharePer1_HP(i) = rangecells.item(i).value
        If Not checkfieldspecialcharacter(CoOwnerSharePer1_HP(i)) Then
            msgbox_hprptfrm ("* Percentage Share in PropertyA" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
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
    Set rangecells = Sheet22.Range("HP.Co.Pan" & j).Cells
    
    For i = 1 To end_hprptfrm
        CoOwnerPAN1_HP(i) = rangecells.item(i).value
        If Not checkfieldspecialcharacter(CoOwnerPAN1_HP(i)) Then
            msgbox_hprptfrm ("* PAN of other co-owner" & i & "   characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCoOwnerPAN1_HP = False
            Exit Function
        End If
        If UCase(CoOwnerPAN1_HP(i)) <> "" Then
        If UCase(CoOwnerPAN1_HP(i)) = Sheet1.Range("sheet1.PAN").value Then
                msgbox_hprptfrm ("* Enter Valid PAN of Other co-owner(s) at Sr. No " & i & " in Property " & i & " schedule HP which should not match pan with Sheet Part A - General" & Chr(13))
                ValidateCoOwnerPAN1_HP = False
            Exit Function
        End If
        End If
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
    Set rangecells = Sheet22.Range("HP.Co.Aadhaar" & j).Cells
    ReDim CoOwnerAadhar_HP(rangecells.count)
   ' ReDim CoOwnerPANCHK_HP(rangecells.count)
    
    For i = 1 To rangecells.count
        CoOwnerAadhar_HP(i) = rangecells.item(i).value
       ' CoOwnerPANCHK_HP(i) = rangecells.item(i).Offset(0, -1).value
        
'        If CoOwnerAadhar_HP(i) = "" And CoOwnerPANCHK_HP(i) = "" Then
'             msgbox_hprptfrm "*""Any one of the field PAN/aadhaar is mandatory"" at Sr. No " & i & """"
'                ValidateCoOwnerAadhar_HP = False
'                Exit Function
'        End If
        
         If Len(CoOwnerAadhar_HP(i)) > 0 Then
                If Len(CoOwnerAadhar_HP(i)) > 12 Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
                
                If CoOwnerAadhar_HP(i) = "000000000000" Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
                
                If CoOwnerAadhar_HP(i) = "111111111111" Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
            
                If Not checkallfieldspecialcharacter(CoOwnerAadhar_HP(i)) Then
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
        CoOwnedShare_HP(i) = Sheet22.Range("HP.SharePercent" & i).value
        CoOwnedYN_HP(i) = Sheet22.Range("HP.CoOwnedYN" & i).value
        
        CoOwnedShare_HP(i) = Round(CoOwnedShare_HP(i), 2)
        
        If Not checkfieldspecialcharacter(CoOwnedShare_HP(i)) Then
            msgbox_hprptfrm ("* Your Percentage of Share" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateSharePercent_HP = False
            Exit Function
        End If
        
        If Mid(CoOwnedYN_HP(i), 1, 1) = "N" Then
            If CoOwnedShare_HP(i) <> 100 Then
                msgbox_hprptfrm ("* Your Percentage of Share" & i & " must be 100 if not Co-owned in schedule HP")
                ValidateSharePercent_HP = False
                Exit Function
            End If
        Else
        End If
    Next
End Function
Function ValidateCoOwnedYN_HP() As Boolean
    Dim i As Long
    
    ValidateCoOwnedYN_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnedYN_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        CoOwnedYN_HP(i) = Sheet22.Range("HP.CoOwnedYN" & i).value
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If
        If Not chkCompulsory(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* ""Please select dropdown from Is property co-owned  in schedule HP""") & Chr(13)
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* IS Property Co-Owned" & i & "  characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* IS Property Co-Owned" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
    Next
End Function


'Function ValidateCoownerRules_HP() As Boolean
'    Dim i As Long
'    ValidateCoownerRules_HP = True
'    setTblinfo_hprptfrm
'    For i = 1 To UBound(CoOwnedYN_HP)
'        If isdropdownblank(CoOwnedYN_HP(i)) Then
'            CoOwnedYN_HP(i) = ""
'        End If
'
'        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
'          Dim CoOwnerShare As Variant
'          CoOwnerShare = CoOwnedShare_HP(i)
'
'        End If
'
'        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
'         Dim counter, name As Variant
'         counter = 0
'         For Each name In Sheet22.Range("HP.Co.Name" & i).Cells
'         If name = "" Then counter = counter + 1
'            If counter = Sheet22.Range("HP.Co.Name" & i).Cells.count Then
'                msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP""")
'                ValidateCoownerRules_HP = False
'                Exit Function
'
'             End If
'        Next
'
'        Dim j As Long
'        Dim coowener As Range
'        Dim Pancoowener As Range
'        Dim Aacoowener As Range
'        Dim percoowener As Range
'        For j = 1 To Sheet22.Range("HP.Co.Name" & i).count
'
'            Set coowener = Sheet22.Range("HP.Co.Name" & i).Cells
'            Set Pancoowener = Sheet22.Range("HP.Co.Pan" & i).Cells
'            Set Aacoowener = Sheet22.Range("HP.Co.Aadhaar" & i).Cells
'            Set percoowener = Sheet22.Range("HP.Co.Share" & i).Cells
'            If coowener.item(j).value = "" Then
'                If Pancoowener.item(j).value <> "" Or Aacoowener.item(j).value <> "" Or percoowener.item(j).value <> "" Then
'                     msgbox_hprptfrm ("* ""Please enter the name of the co-owner""")
'
'                    ValidateCoownerRules_HP = False
'                Exit Function
'                End If
'            End If
' 'Commented_ayush
''           If coowener.item(j).value <> "" Then
''            If Pancoowener.item(j).value = "" And Aacoowener.item(j).value = "" Then
''                msgbox_hprptfrm ("* ""Any one of the field PAN/aadhaar is mandatory"" at Sr. No " & j & "")
''                ValidateCoownerRules_HP = False
''                Exit Function
''            End If
''            End If
'            '------------------
'        Next j
'         End If
'        Next
'
'
'End Function
'Function ValidateCoownerRules_HP() As Boolean
'    Dim i As Long
'    ValidateCoownerRules_HP = True
'    setTblinfo_hprptfrm
'    For i = 1 To UBound(CoOwnedYN_HP)
'        If isdropdownblank(CoOwnedYN_HP(i)) Then
'            CoOwnedYN_HP(i) = ""
'        End If
'
'        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
'          Dim CoOwnerShare As Variant
'          CoOwnerShare = CoOwnedShare_HP(i)
'        End If
'        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
'         Dim counter, name As Variant
'         counter = 0
'         For Each name In Sheet22.Range("HP.Co.Name" & i).Cells
'         If name = "" Then counter = counter + 1
'            If counter = Sheet22.Range("HP.Co.Name" & i).Cells.count Then
'                msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP""")
'                ValidateCoownerRules_HP = False
'                Exit Function
'             End If
'        Next
''        Dim j As Long
''        Dim coowener As Range
''        Dim Pancoowener As Range
''        Dim Aacoowener As Range
''        Dim percoowener As Range
''        For j = 1 To Sheet22.Range("HP.Co.Name" & i).count
''
''            Set coowener = Sheet22.Range("HP.Co.Name" & i).Cells
''            Set Pancoowener = Sheet22.Range("HP.Co.Pan" & i).Cells
''            Set Aacoowener = Sheet22.Range("HP.Co.Aadhaar" & i).Cells
''            Set percoowener = Sheet22.Range("HP.Co.Share" & i).Cells
''            If coowener.item(j).value = "" Then
''                If Pancoowener.item(j).value <> "" Or Aacoowener.item(j).value <> "" Or percoowener.item(j).value <> "" Then
''                     msgbox_hprptfrm ("* ""Please enter the name of the co-owner""")
''
''                    ValidateCoownerRules_HP = False
''                Exit Function
''                End If
''            End If
'' 'Commented_ayush
'''           If coowener.item(j).value <> "" Then
'''            If Pancoowener.item(j).value = "" And Aacoowener.item(j).value = "" Then
'''                msgbox_hprptfrm ("* ""Any one of the field PAN/aadhaar is mandatory"" at Sr. No " & j & "")
'''                ValidateCoownerRules_HP = False
'''                Exit Function
'''            End If
'''            End If
''            '------------------
''        Next j
'
'
'         End If
'        Next
'
'
'
'Dim j As Long
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    'ValidateCoownerName1_HP = True
'    setTblinfo_hprptfrm
'    ReDim CoOwnerName1_HP(end_hprptfrm)
'    ReDim CowOwned1_HP(end_hprptfrm)
'    For j = 1 To end_hprptfrm
'    Set rangecells = Sheet22.Range("HP.Co.Name" & j).Cells
'    For i = 1 To end_hprptfrm
'        CoOwnerName1_HP(i) = rangecells.item(i).value
'        CowOwned1_HP(i) = Sheet22.Range("HP.CoOwnedYN" & i).value
'
'
'        If Mid(CowOwned1_HP(i), 1, 1) = "Y" Then
'        If rangecells.item(i).Offset(0, 1).value <> "" Or rangecells.item(i).Offset(0, 2).value <> "" Or rangecells.item(i).Offset(0, 3).value <> "" Then
'        If Not chkCompulsory(CoOwnerName1_HP(i)) Then
'            msgbox_hprptfrm ("* ""Please enter the name of the co-owner at S.No " & i & """") & Chr(13)
'
'             ValidateCoownerRules_HP = False
'            Exit Function
'        End If
'        End If
'        End If
'    Next
'    Next
'
'
'
'
'End Function

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

         For Each name In Sheet22.Range("HP.Co.Name" & i).Cells

         If name = "" Then counter = counter + 1

            If counter = Sheet22.Range("HP.Co.Name" & i).Cells.count Then

                msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP""")

                ValidateCoownerRules_HP = False

                Exit Function

             End If
 
         Next

        Dim j As Long

        Dim coowener As Range

        Dim Pancoowener As Range

        Dim Aacoowener As Range

        Dim percoowener As Range

        For j = 1 To Sheet22.Range("HP.Co.Name" & i).count
 
            Set coowener = Sheet22.Range("HP.Co.Name" & i).Cells

            Set Pancoowener = Sheet22.Range("HP.Co.Pan" & i).Cells

            Set Aacoowener = Sheet22.Range("HP.Co.Aadhaar" & i).Cells

            Set percoowener = Sheet22.Range("HP.Co.Share" & i).Cells

            If coowener.item(j).value = "" Then

                If Pancoowener.item(j).value <> "" Or Aacoowener.item(j).value <> "" Or percoowener.item(j).value <> "" Then

                     msgbox_hprptfrm ("* ""Please enter the name of the other co-owner at S.No " & j & """")
 
                    ValidateCoownerRules_HP = False

                Exit Function

                End If

            End If
 
        Next j


End If

Next
 
End Function
 
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
'         Dim counter, name, name1 As Variant
'         counter = 0
'         For Each name In Sheet22.Range("HP.Co.Share" & i).Cells
'          If name = "" Then counter = counter + 1
'
'                    If counter = Sheet22.Range("HP.Co.Share" & i).Cells.count Then
'
'                     msgbox_hprptfrm ("* ""Please enter Percentage Share of Co-Owner in schedule HP """)
'                        ValidatesharesRules_HP = False
'                        Exit Function
'                    End If
'         Next
'
           
         
    Dim j As Long
    Dim coowener As Range
    Dim Pancoowener As Range
    Dim Aacoowener As Range
    Dim percoowener As Range
         For j = 1 To Sheet22.Range("HP.Co.Name" & i).count

         Set coowener = Sheet22.Range("HP.Co.Name" & i).Cells
         Set Pancoowener = Sheet22.Range("HP.Co.Pan" & i).Cells
         Set Aacoowener = Sheet22.Range("HP.Co.Aadhaar" & i).Cells
         Set percoowener = Sheet22.Range("HP.Co.Share" & i).Cells
         
         If percoowener.item(j).value = "" Then
         If coowener.item(j).value <> "" Or Pancoowener.item(j).value <> "" Or Aacoowener.item(j).value <> "" Then
         msgbox_hprptfrm ("* ""Please enter Percentage Share of other Co-Owner at S.No " & j & """")
        ValidatesharesRules_HP = False
         Exit Function
         End If
         End If
         Next j
                 
        End If
    Next
End Function

Sub setTblinfo_hpco()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    
    mIntCells = Sheet22.Range("hp.co.name1").count
    Set rangecells = Sheet22.Range("hp.co.name1").Cells
    
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
 
   
    
    
    
    
    





'Sub AddPropertyCoOWners(Optional iRows As Long = 0)
'On Error GoTo endline
'    Dim newrngname As Variant
'    Dim numberofrows As Long
'    Dim newfrmsize, i, te As Long
'
'
'    Application.EnableEvents = False
'    Sheet22.Activate
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet22.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'    numberofrows = iRows
'    For te = 1 To Sheet22.Range("PropertySectionCOunt").value
'
'
'        mdCommon.DefinedgridNameRange = "HP.Co.Srno||HP.Co.Name||HP.Co.Pan||HP.Co.Aadhaar||HP.Co.Share||Co_OwnnerBlock"
'        ActiveCellRange = EfilingCommon.searchLastRow("HP.Co.Srno" & te)
'
'        If te > 1 Then
'            numberofrows = mdCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
'        Else
'            numberofrows = mdCommon.insertRowUnderSectionWithFormula(iRows, True, te)
'        End If
'    Next
'
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet22.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'
'    Application.EnableEvents = False
'    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
'        newfrmsize = Sheet22.Range("NumRowsPropertyBlock").value
'        newfrmsize = newfrmsize + numberofrows
'        Sheet22.Range("NumRowsPropertyBlock").value = newfrmsize
'    Next
'
'endline:
'
''----------------Lock Password-------------------START---
'   Sheet22.Protect Password:=sPassword
''----------------Lock Password-------------------END-----
'  Application.EnableEvents = True
'End Sub
Sub setTblinfo_24bankname(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet22.Range("bankName.24b" & myindex).count
    Set rangecells = Sheet22.Range("bankName.24b" & myindex).Cells
    
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
    mIntCells = Sheet22.Range("LoanfrmBankOrInstitute.24b" & myindex).count
    Set rangecells = Sheet22.Range("LoanfrmBankOrInstitute.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
    
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24a = ccount
    'end_hpco_24b = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bLoan(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet22.Range("loanAccNum.24b" & myindex).count
    Set rangecells = Sheet22.Range("loanAccNum.24b" & myindex).Cells
    
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
    mIntCells = Sheet22.Range("loanDate.24b" & myindex).count
    Set rangecells = Sheet22.Range("loanDate.24b" & myindex).Cells
    
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
    mIntCells = Sheet22.Range("loanAmt.24b" & myindex).count
    Set rangecells = Sheet22.Range("loanAmt.24b" & myindex).Cells
    
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
    mIntCells = Sheet22.Range("loanOutstanding.24b" & myindex).count
    Set rangecells = Sheet22.Range("loanOutstanding.24b" & myindex).Cells
    
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
    mIntCells = Sheet22.Range("Intrst.24b" & myindex).count
    Set rangecells = Sheet22.Range("Intrst.24b" & myindex).Cells
    
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
    Dim end_hpco_24d As Long
    Validate24B_Table = True
    

'    sharepercent = 0
'    cosharepercent = 0
'    setTblinfo_hprptfrm
    
    end_hprptfrm = Sheet22.Range("PropertySectionCOunt").value
    
    For i = 1 To end_hprptfrm
        
        Set rangecells1 = Sheet22.Range("LoanfrmBankOrInstitute.24b" & i).Cells
        Set rangecells2 = Sheet22.Range("bankName.24b" & i).Cells
        Set rangecells3 = Sheet22.Range("loanAccNum.24b" & i).Cells
        Set rangecells4 = Sheet22.Range("loanDate.24b" & i).Cells
        Set rangecells5 = Sheet22.Range("loanAmt.24b" & i).Cells
        Set rangecells6 = Sheet22.Range("loanOutstanding.24b" & i).Cells
        Set rangecells7 = Sheet22.Range("Intrst.24b" & i).Cells
        
       
        'setTblinfo_hpcoindex (i)
        setTblinfo_24bBankorInst (i)
        setTblinfo_24bankname (i)   'added by Chetan C M for AY 2025-26
        setTblinfo_24bLoan (i)
        setTblinfo_24bdate (i)
        setTblinfo_24bAmount (i)
        setTblinfo_24bloanout (i)
        setTblinfo_24bInt (i)
        
       
       end_hpco_24d = Application.WorksheetFunction.Max(0, end_hpco_24a, end_hpco_24b, end_hpco_24c, end_hpco_24e, end_hpco_24f, end_hpco_24k, end_hpco_24s, end_hpco_24r)
        
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
                    msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13) 'SIT-93563 updated by shrutika
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Len(AccntNum_24b(i, j)) > 20 Then
                    msgbox_hprptfrm ("* Loan Account number  at Sr. No " & j & " in schedule 24b less than 20 characters.") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                
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
                        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & j & " in table " & i & "") & Chr(13)
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
            
            If Not ChkMaxDate_24b(Trim(LoanDate_24B(i, j)), "31-03-2026") Then
                msgbox_hprptfrm ("* Date can not be after 31/03/2026 at Sr. No " & j & ".") & Chr(13)
                Validate24B_Table = False
              Exit Function
            End If
            
            '---Loan amount
            If Not chkCompulsory(LoanAmt_24b(i, j)) Then
            
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
            
            '             MsgBox_24b = MsgBox_24b + "* ""Loan outstanding as on 31-03-2025 is mandatory in schedule 24(b)"" at Sr. No " & j & "" & Chr(13)
                   ' msgbox_hprptfrm ("* ""Loan outstanding as on last date of financial year is mandatory in schedule 24(b)"" at Sr. No " & j & "") & Chr(13)
                'msgbox_hprptfrm ("* ""Loan outstanding as on last date of finacial year is mandatory in schedule 24(b). Please enter 0 in case the entire loan is repaid during the year"" at Sr. No " & j & " in table " & i & "") & Chr(13) '(SIT - 93574)
                
                
     
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
             msgbox_hprptfrm ("* ""Please provide Interest u/s 24(b)"" at Sr. No " & j & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            '         If Len(Intrst_24b(i)) > 15 Then
            '          MsgBox_24b = MsgBox_24b + "* Interest  at Sr. No " & j & " in Sheet 24b less than 15 characters." & Chr(13)
            '            Validate24B_Table = False
            '            Exit Function
            '        End If
            
            If Not IsNumeric(Intrst_24B(i, j)) Then
               msgbox_hprptfrm ("* Interest at Sr. No " & j & " in schedule 24b should be Numeric value") & Chr(13)
                
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
               msgbox_hprptfrm ("* ""Interest u/s 24(b) should be more than 0 in schedule 24(b) at Sr. No " & j & """") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
        Next
        
         If (Len(Sheet22.Range("TotAmt.24b" & i).value) > 14) Then
            msgbox_hprptfrm ("*  Total of interest on borrowed capital u/s 24(b) cannot exceed 14 Digits.") & Chr(13)
            Validate24B_Table = False
            Exit Function
         End If
            
        
        
    Next
endlin1:
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
            msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & cc & "") & Chr(13) 'updated by Shrutika(SIT-93563)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
           ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b) " & Chr(13)
            msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
       ' MsgBox_24b = MsgBox_24b & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
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
 msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
      msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        'MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
       ' MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If

    ReDim myB(Len(BankAccountNumber) - 1)
    For i = 1 To Len(BankAccountNumber)
        myB(i - 1) = Mid(BankAccountNumber, i, 1)
    Next i

    ' ? FIXED LOGIC: count ONLY digits 1–9
    For i = LBound(myB) To UBound(myB)
        If IsNumeric(myB(i)) Then
            If CLng(myB(i)) >= 1 And CLng(myB(i)) <= 9 Then
                countnum = countnum + 1
            End If
        End If

        If i > LBound(myB) And i < UBound(myB) Then
            If myB(i) = "0" Then
                If myB(i - 1) = "0" Then
                    zeroCount = zeroCount + 1
                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
                Else
                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
                End If
            End If
        End If
    Next i

    If BeforeZero <> "" And AfterZero <> "" Then
        If zeroCount > 1 Then
            'MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc + 1 & " is invalid  in schedule 24(b)" & Chr(13)
           msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        'errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 24(b) is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
       msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
End Function
Sub Addrows24b_HP(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheet22.Activate

'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   Sheet22.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value


        mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
        ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.24b" & te)

        If te > 1 Then
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next



'----------------Unlock Password-------------------START---
   sPassword = mdCommon.getmsgstate
   Sheet22.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
        newfrmsize = Sheet22.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet22.Range("NumRowsPropertyBlock").value = newfrmsize
    Next


endline:

'----------------Lock Password-------------------START---
   Sheet22.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
Function AddTotIncomeUnderHouseProperty_1()
    Dim TotIncomeRangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, Totsections As Long

    tot = 0
    TotIncomeRangeName = "HP.IncomeOfHP"
    Totsections = Sheet22.Range("PropertySectionCOunt").value
        For i = 1 To Totsections
        newDefinedName = TotIncomeRangeName & i
        tot = tot + Sheet22.Range(newDefinedName)
        Next
        
        
    Sheet22.Unprotect Password:=getmsgstate
      Sheet22.Range("HP.TotalIncomeChargeableUnHP").value = tot '+ Sheet8.Range("HP.PassTroughIncome").value '+ Sheet8.Range("HP.RentArearsSec25BAfter30pcDeduct").value
    Sheet22.Protect Password:=getmsgstate


End Function
Sub setTblinfo_24bLoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("LoanfrmBankOrInstitute.24b" & i).count
    Set rangecells = Range("LoanfrmBankOrInstitute.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanfrm1 = ccount
End Sub

Sub setTblinfo_24bBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("bankName.24b" & i).count
    Set rangecells = Range("bankName.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bbankName1 = ccount
End Sub

Sub setTblinfo_24bAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanAccNum.24b" & i).count
    Set rangecells = Range("loanAccNum.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bAccntNum1 = ccount
End Sub

Sub setTblinfo_24bLoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanDate.24b" & i).count
    Set rangecells = Range("loanDate.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanDate1 = ccount
End Sub

Sub setTblinfo_24bLoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanAmt.24b" & i).count
    Set rangecells = Range("loanAmt.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanAmt1 = ccount
End Sub

Sub setTblinfo_24bLoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanOutstanding.24b" & i).count
    Set rangecells = Range("loanOutstanding.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanOutstanding1 = ccount
End Sub

Sub setTblinfo_24bIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("Intrst.24b" & i).count
    Set rangecells = Range("Intrst.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bIntrst1 = ccount
End Sub
Function CheckHousePropertyIncome_New(HP As Variant) As Boolean
On Error Resume Next
 CheckHousePropertyIncome_New = True
'Comment by Riyaz
 'If Mid(Range("HP.ifLetOut1").value, 1, 1) = "S" Then
 If (HP > 200000) Then
 fmsgboxsmall "* Interest cannot be more than Rs. 2,00,000 for category self-occupied in schedule HP." ', vbOKOnly, "Error(s)!"
 'Range("HP.IntOnBorwCap1").Select
 'Range("HP.IntOnBorwCap1").value = ""
  CheckHousePropertyIncome_New = False   'Malli 23/10/2024
 Exit Function
' End If
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
    
 Function CheckAtoZ(chr1) As Boolean
    CheckAtoZ = True
    If ((asc(chr1) < 65) Or (asc(chr1) > 90)) Then
        CheckAtoZ = False
    End If
End Function
Function CheckSpecialCharacter4New(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArraynew As Variant
    Dim iCharCountNew, iSpecialCharNew As Long
    
    CheckSpecialCharacter4New = True
    specialCharArraynew = Array(">", "<")
    For iCharCountNew = 1 To Len(emailid)
        For iSpecialCharNew = 0 To UBound(specialCharArraynew)
        If Mid(emailid, iCharCountNew, 1) = specialCharArraynew(iSpecialCharNew) Then
            CheckSpecialCharacter4New = False
            Exit Function
        End If
        Next
    Next
End Function
Function ValidateMandatorytable1() As Boolean
ValidateMandatorytable1 = True
Dim i As Long
Dim flag As Boolean

Dim rangecells As Range



flag = True
For i = 1 To Sheet22.Range("LoanfrmBankOrInstitute.24b2").Rows.count

    

If Sheet22.Range("HP.AddrDetail2").value = "" Or Sheet22.Range("HP.CityOrTownOrDistrict2").value = "" Or (Sheet22.Range("HP.StateCode2").value = "(Select)" Or Sheet22.Range("HP.StateCode2").value = "") Or Sheet22.Range("HP.CountryCode2").value = "" Or (Sheet22.Range("HP.ifLetOut2").value = "(Select)" Or Sheet22.Range("HP.ifLetOut2").value = "") Then
        
        If Sheet22.Range("LoanfrmBankOrInstitute.24b2").item(i).value <> "" And Sheet22.Range("bankName.24b2").item(i).value <> "" And Sheet22.Range("loanAccNum.24b2").item(i).value <> "" And Sheet22.Range("loanDate.24b2").item(i).value <> "" And Sheet22.Range("loanAmt.24b2").item(i).value <> "" And Sheet22.Range("loanOutstanding.24b2").item(i).value <> "" And Sheet22.Range("Intrst.24b2").item(i).value <> "" Then
            flag = False
              msgbox_hprptfrm ("* Enter all mandatory fields of Block 1B in Schedule HP.")

        End If
        

        If flag = False Then
            ValidateMandatorytable1 = False
            Exit Function
        End If
        
    End If
Next i
End Function




