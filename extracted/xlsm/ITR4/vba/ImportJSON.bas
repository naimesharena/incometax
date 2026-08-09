Attribute VB_Name = "ImportJSON"
'14/12/2021
Option Explicit

Dim end_Exempt1
Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

Sub Import()
Dim filepath, jsonText, bacVal As String


Unload UserForm3
fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."

With Application.FileDialog(msoFileDialogFilePicker)
    'Makes sure the user can select only one file
    .AllowMultiSelect = False
    .Title = "Please select a Json file."
    'Filter to just the following types of files to narrow down selection options
    .Filters.add "Json File", "*.json", 1
    'Show the dialog box
    If .Show = True Then
        imported = 1
        filepath = .SelectedItems.item(1)
    Else
    imported = 0
        Exit Sub
    End If
    On Error Resume Next
End With

Open filepath For Input As #1
jsonText = Input$(LOF(1), 1)
Close #1

Dim jsonObject As Object

Set jsonObject = ParseJson(jsonText)

If Not jsonObject.exists("ITR") Then
imported = 0
fmsgbox ("Please select a valid Import JSON.")
Exit Sub
End If


ImportPersonalInfo (jsonText)
ImportFilingStatus (jsonText)
If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
    ImportPartA_139_8A (jsonText)
    ImportPartB_ATI (jsonText)
End If
ImportITR4_IncomeDeductions (jsonText)
ImportTaxComputation (jsonText)
ImportTaxPaid (jsonText)
'ImportFilingInfo (jsonText)
ImportTDSonSalary (jsonText)
ImportTDSOthThanSals (jsonText)
ImportScheduleTDS3Dtls (jsonText)
ImportGst (jsonText)
ImportScheduleBP (jsonText)
ImportADschedule (jsonText)
ImportADAschedule (jsonText)
ImportAEschedule (jsonText)
ImportGoodsTableschedule (jsonText)
ImportTaxReturnPreparer (jsonText)
ImportRefund (jsonText)
ImportScheduleIT (jsonText)
ImportVerification (jsonText)
ImportScheduleTCS (jsonText)
ImportTaxExmpIntIncDtls (jsonText)
ImportLTCG112A (jsonText) ' added by Chetan C M AY 2025-26 on 31/01/2025
'Sheet8.Unprotect Password:=getmsgstate


If (Sheet8.Range("bacValue").value) = 2 Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet16.Visible = True
'    Sheet6.Visible = True
'    ThisWorkbook.Protect Password:=getmsgstate
    ImportSchedule80GGC (jsonText) '80GGC_AY_2024_25_Added by Malli
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
        ImportSchedule80D (jsonText)
    End If
    'Malli_Leagcy40_02/06/2025
    
'    ImportSchedule80DD (jsonText) '2024-25 change
'    ImportSchedule80U (jsonText) '2024-25 change
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
        ImportSchedule80DD (jsonText) '2024-25 change
    End If
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
        ImportSchedule80U (jsonText) '2024-25 change
    End If

    ImportSchedule80G100NoAppr (jsonText)
    ImportSchedule80G50NoAppr (jsonText)
    ImportSchedule80G100Appr (jsonText)
    ImportSchedule80G50Appr (jsonText)
    
    'Konda_AY_2025_26  09/04/2025
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
        ImportSchedule_80C (jsonText)
    End If
    
     'AY_25_26_V0.7
   ' ImportSchedule_80CCC (jsonText)
    '------------------
   'Malli_SIT-118203  23/04/2026
   If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
    ImportSchedule_80E (jsonText)
    ImportSchedule_80EE (jsonText)
    ImportSchedule_80EEA (jsonText)
    ImportSchedule_80EEB (jsonText)
   End If
    'ImportSchedule_Int24B (jsonText)
End If
'Malli comented-------------------------------
''Konda---29-04-2025--SIT-92218
'If Sheet8.Range("BacValue").value = 1 And Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "S" And Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "(" Then
'     ImportSchedule_Int24B (jsonText)
'ElseIf Sheet8.Range("BacValue").value = 2 And Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "(" Then
'     ImportSchedule_Int24B (jsonText)
'End If

'Konda commented on 29-01-2026
''Malli------SIT_SIT-93261
'If Not (Sheet8.Range("BacValue").value = 1 And Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1) = "S") Then
'     ImportSchedule_Int24B (jsonText)
'End If
''-------------------SIT-92218

'Konda_08/05/2025_ as per New Schema_v0.7 and V0.7.1
'If Sheet8.Range("BacValue").value = 2 And Mid(Range("sheet1.NatureofEmployment").value, 1, 1) <> "N" Then
If Sheet8.Range("BacValue").value = 2 And Left(Sheet1.Range("sheet1.Status").value, 1) = "I" And Mid(Range("sheet1.NatureofEmployment").value, 1, 1) <> "N" Then
    ImportScheduleEA10_13A (jsonText)

End If
'----------------------------

'Ayush_05/07/2026----------------------------

Sheet1.Range("IncD.Section80DD").Formula = "=IF(BacValue=1,0,AmtDeduction_80DD)"
Sheet1.Range("IncD.Section80U").Formula = "=IF(BacValue=1,0,AmountDeduction_80U)"

Sheet8.Protect Password:=getmsgstate
End Sub
Function ImportPersonalInfo(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
Dim jsonDictionary As Object
Dim PAN, firstName, middleName, LastName, residenceNo, residenceName, roadOrStreet, localityOrArea, cityOrTownOrDistrict, StateCode, CountryCode, PinCode, zipCode, countryCodeMobile, mobileNo, emailAddress, dob, iEmpCat, sEmpCat, aadhaarCardNo, AadhaarEnrolmentIds As Variant
Dim sCountry, iCountry As Variant
Dim sState, iState, EmailAddress2, Mobile2, CCM2, STD, Phone, Status1 As Variant
Dim YYYY, MM, DD, strDate As String

'Konda updated on 23-01-2026
'Dim WantUpdateAdd1
'V0.8
Dim SecondaryAdd1
Dim residenceNo1, residenceName1, roadOrStreet1, localityOrArea1, cityOrTownOrDistrict1, StateCode1, CountryCode1, PinCode1, zipCode1, countryCodeMobile1, mobileNo1, emailAddress1 As Variant
Dim sState1, iState1, sCountry1, iCountry1, countryCodeMobil1e3
'=============
Set jsonObject = ParseJson(jsonText)

firstName = jsonObject("ITR")("ITR4")("PersonalInfo")("AssesseeName")("FirstName")
middleName = jsonObject("ITR")("ITR4")("PersonalInfo")("AssesseeName")("MiddleName")
LastName = jsonObject("ITR")("ITR4")("PersonalInfo")("AssesseeName")("SurNameOrOrgName")
PAN = jsonObject("ITR")("ITR4")("PersonalInfo")("PAN")
residenceNo = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("ResidenceNo")
residenceName = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("ResidenceName")
roadOrStreet = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("RoadOrStreet")
localityOrArea = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("LocalityOrArea")
cityOrTownOrDistrict = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("CityOrTownOrDistrict")
iState = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("StateCode")
sState = Findtext(CStr(iState), "State")
iCountry = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("CountryCode")
sCountry = Findtext(CStr(iCountry), "Country")
countryCodeMobile = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("CountryCodeMobile")
mobileNo = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("MobileNo")
emailAddress = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("EmailAddress")
PinCode = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("PinCode")
zipCode = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("ZipCode")

'Konda updated on 23-01-2026
'WantUpdateAdd1 = jsonObject("ITR")("ITR4")("PersonalInfo")("WantUpdateAdd")
SecondaryAdd1 = jsonObject("ITR")("ITR4")("PersonalInfo")("SecondaryAdd")
residenceNo1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("ResidenceNo")
residenceName1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("ResidenceName")
roadOrStreet1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("RoadOrStreet")
localityOrArea1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("LocalityOrArea")
cityOrTownOrDistrict1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("CityOrTownOrDistrict")
iState1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("StateCode")
sState1 = Findtext(CStr(iState1), "State")
iCountry1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("CountryCode")
sCountry1 = Findtext(CStr(iCountry1), "Country")
PinCode1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("PinCode")
zipCode1 = jsonObject("ITR")("ITR4")("PersonalInfo")("AlternateAddress")("ZipCode")

countryCodeMobile1 = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("CountryCodeMobileNoSec")
mobileNo1 = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("MobileNoSec")
emailAddress1 = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("EmailAddressSec")

'===============


dob = jsonObject("ITR")("ITR4")("PersonalInfo")("DOB")
If dob <> "" Then
    YYYY = Mid(dob, 1, 4)
    MM = Mid(dob, 6, 2)
    DD = Mid(dob, 9, 2)
    strDate = DD & "/" & MM & "/" & YYYY
End If
aadhaarCardNo = jsonObject("ITR")("ITR4")("PersonalInfo")("AadhaarCardNo")
AadhaarEnrolmentIds = jsonObject("ITR")("ITR4")("PersonalInfo")("AadhaarEnrolmentId")
EmailAddress2 = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("EmailAddressSec")
Mobile2 = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("MobileNoSec")
CCM2 = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("CountryCodeMobileNoSec")
STD = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("Phone")("STDcode")
Phone = jsonObject("ITR")("ITR4")("PersonalInfo")("Address")("Phone")("PhoneNo")

iEmpCat = jsonObject("ITR")("ITR4")("PersonalInfo")("EmployerCategory")
    If iEmpCat = "CGOV" Then
    sEmpCat = "Central Government"
'Konda Updated On 27-02-2026--V0.4
''Konda Updated On 12-02-2026--V0.3
'    ElseIf iEmpCat = "SCJ" Then
'    sEmpCat = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958"
''====================
    ElseIf iEmpCat = "SGOV" Then
    sEmpCat = "State Government"
    ElseIf iEmpCat = "PSU" Then
    sEmpCat = "Public Sector Undertaking"
    ElseIf iEmpCat = "OTH" Then
    sEmpCat = "Others"
    ElseIf iEmpCat = "PE" Then
    sEmpCat = "Pensioners-CG"
    ElseIf iEmpCat = "PESG" Then
    sEmpCat = "Pensioners-SG"
    ElseIf iEmpCat = "PEPS" Then
    sEmpCat = "Pensioners-PSU"
    ElseIf iEmpCat = "PEO" Then
    sEmpCat = "Pensioners-Other"
    ElseIf sEmpCat = "NA" Then
    sEmpCat = "Not Applicable (eg. Family pension etc)"
    End If
Status1 = jsonObject("ITR")("ITR4")("PersonalInfo")("Status")
    If Status1 = "I" Then
    Status1 = "I - INDIVIDUAL"
    ElseIf Status1 = "H" Then
    Status1 = "H - HUF"
    ElseIf Status1 = "F" Then
    Status1 = "F - FIRM(Other than LLP)"
    End If
    
   Sheet1.Unprotect Password:=getmsgstate
     If Status1 <> "" Then
        Sheet1.Range("sheet1.Status").value = Status1
        Application.Wait (Now + TimeValue("00:00:60"))
    End If
   Sheet1.Unprotect Password:=getmsgstate
    If firstName <> "" Then
        Sheet1.Range("sheet1.FirstName").value = firstName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If middleName <> "" Then
        Sheet1.Range("sheet1.MiddleName").value = middleName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If LastName <> "" Then
        Sheet1.Range("sheet1.SurNameOrOrgName").value = LastName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    
    If PAN <> "" Then
        Sheet1.Range("sheet1.PAN").value = PAN
    End If
    Sheet1.Unprotect Password:=getmsgstate
 
    If strDate <> "" Then
        Sheet1.Range("sheet1.DOB").value = strDate
    End If
    Sheet1.Protect Password:=getmsgstate
    If residenceNo <> "" And Sheet1.Range("sheet1.ResidenceNo").Locked = False Then
        Sheet1.Range("sheet1.ResidenceNo").value = residenceNo
    End If
    If residenceName <> "" And Sheet1.Range("sheet1.ResidenceName").Locked = False Then
        Sheet1.Range("sheet1.ResidenceName").value = residenceName
    End If
    If roadOrStreet <> "" And Sheet1.Range("sheet1.RoadOrStreet").Locked = False Then
        Sheet1.Range("sheet1.RoadOrStreet").value = roadOrStreet
    End If
    If localityOrArea <> "" And Sheet1.Range("sheet1.LocalityOrArea").Locked = False Then
        Sheet1.Range("sheet1.LocalityOrArea").value = localityOrArea
    End If
    If cityOrTownOrDistrict <> "" And Sheet1.Range("sheet1.CityOrTownOrDistrict").Locked = False Then
        Sheet1.Range("sheet1.CityOrTownOrDistrict").value = cityOrTownOrDistrict
    End If
    If iCountry = 99 Then
        If sCountry <> "" And Sheet1.Range("sheet1.Country").Locked = False Then
            Sheet1.Range("sheet1.CountryCode").value = sCountry
        End If
        If sState <> "" And Sheet1.Range("sheet1.StateCode1").Locked = False Then
            Sheet1.Range("sheet1.StateCode1").value = sState
        End If
    Else
        If sState <> "" And Sheet1.Range("sheet1.StateCode1").Locked = False Then
            Sheet1.Range("sheet1.StateCode1").value = sState
        End If
         If sCountry <> "" And Sheet1.Range("sheet1.Country").Locked = False Then
            Sheet1.Range("sheet1.CountryCode").value = sCountry
        End If
    End If
    
    If PinCode <> "" And Sheet1.Range("sheet1.PinCode").Locked = False Then
        Sheet1.Range("sheet1.PinCode").value = PinCode
    End If
    If zipCode <> "" And Sheet1.Range("sheet1.ZipCode").Locked = False Then
        Sheet1.Range("HASZIP").value = "No"
        Sheet1.Range("sheet1.ZipCode").value = zipCode
    End If
    
    If emailAddress <> "" And Sheet1.Range("sheet1.EmailAddress").Locked = False Then
        Sheet1.Range("sheet1.EmailAddress").value = emailAddress
    End If
    If countryCodeMobile <> "" And Sheet1.Range("sheet1.MobileCountryCode").Locked = False Then
        Sheet1.Range("sheet1.MobileCountryCode").value = countryCodeMobile
    End If
    If mobileNo <> "" And Sheet1.Range("sheet1.Mobileno").Locked = False Then
        Sheet1.Range("sheet1.Mobileno").value = mobileNo
    End If
'Konda updated on 23-01-2026
    If emailAddress1 <> "" And Sheet1.Range("sheet1.EmailAddress2").Locked = False Then
        Sheet1.Range("sheet1.EmailAddress2").value = emailAddress1
    End If
    If countryCodeMobile1 <> "" And Sheet1.Range("sheet1.mobileCountryCode2").Locked = False Then
        Sheet1.Range("sheet1.mobileCountryCode2").value = countryCodeMobile1
    End If
    If mobileNo1 <> "" And Sheet1.Range("sheet1.MobileNoSec").Locked = False Then
        Sheet1.Range("sheet1.MobileNoSec").value = mobileNo1
    End If

'===================
'Konda updated on 17-03-2026--V0.6
'Konda updated on 23-01-2026
'Dim Add_flag
'    If WantUpdateAdd1 <> "" Then
'        If WantUpdateAdd1 = "Y" Then
'            Add_flag = "Yes"
'        Else
'        Add_flag = "No"
'        End If
'        Sheet1.Range("Secondary_Address").value = Add_flag
'    End If
'If Sheet1.Range("Secondary_Address").value = "Yes" Then
'+===
'V0.8
Dim Add_flag
    If SecondaryAdd1 <> "" Then
        If SecondaryAdd1 = "Y" Then
            Add_flag = "Yes"
        Else
        Add_flag = "No"
        End If
        Sheet1.Range("Secondary_Address").value = Add_flag
    End If
'If Sheet1.Range("Secondary_Address").value = "Yes" Then
'+==
'V0.8
 If Sheet1.Range("Secondary_Address").value = "No" Then
'+=

        If residenceNo1 <> "" And Sheet1.Range("sheet1.ResidenceNo2").Locked = False Then
        Sheet1.Range("sheet1.ResidenceNo2").value = residenceNo1
    End If
    If residenceName1 <> "" And Sheet1.Range("sheet1.ResidenceName2").Locked = False Then
        Sheet1.Range("sheet1.ResidenceName2").value = residenceName1
    End If
    If roadOrStreet1 <> "" And Sheet1.Range("sheet1.RoadOrStreet2").Locked = False Then
        Sheet1.Range("sheet1.RoadOrStreet2").value = roadOrStreet1
    End If
    If localityOrArea1 <> "" And Sheet1.Range("sheet1.LocalityOrArea2").Locked = False Then
        Sheet1.Range("sheet1.LocalityOrArea2").value = localityOrArea1
    End If
    If cityOrTownOrDistrict1 <> "" And Sheet1.Range("sheet1.CityOrTownOrDistrict2").Locked = False Then
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").value = cityOrTownOrDistrict1
    End If
    If iCountry = 99 Then
        If sCountry1 <> "" And Sheet1.Range("sheet1.Country2").Locked = False Then
            Sheet1.Range("sheet1.CountryCode2").value = sCountry1
        End If
        If sState1 <> "" And Sheet1.Range("sheet1.StateCode2").Locked = False Then
            Sheet1.Range("sheet1.StateCode2").value = sState1
        End If
    Else
        If sState1 <> "" And Sheet1.Range("sheet1.StateCode2").Locked = False Then
            Sheet1.Range("sheet1.StateCode2").value = sState1
        End If
         If sCountry1 <> "" And Sheet1.Range("sheet1.Country2").Locked = False Then
            Sheet1.Range("sheet1.CountryCode2").value = sCountry1
        End If
    End If
    
    If PinCode1 <> "" And Sheet1.Range("sheet1.PinCode2").Locked = False Then
        Sheet1.Range("sheet1.PinCode2").value = PinCode1
    End If
    If zipCode1 <> "" And Sheet1.Range("sheet1.ZipCode2").Locked = False Then
        Sheet1.Range("HASZIP2").value = "No"
        Sheet1.Range("sheet1.ZipCode2").value = zipCode1
    End If
End If     'Konda updated on 17-03-2026--V0.6 V0.8

'============================

'    If STD <> "" And Sheet1.Range("sheet1.STDcode").Locked = False Then
'        Sheet1.Range("sheet1.STDcode").value = STD
'    End If
'    If Phone <> "" And Sheet1.Range("sheet1.PhoneNo").Locked = False Then
'        Sheet1.Range("sheet1.PhoneNo").value = Phone
'    End If
'    If CCM2 <> "" And Sheet1.Range("sheet1.mobileCountryCode2").Locked = False Then
'        Sheet1.Range("sheet1.mobileCountryCode2").value = CCM2
'    End If
'    If Mobile2 <> "" And Sheet1.Range("sheet1.MobileNoSec").Locked = False Then
'        Sheet1.Range("sheet1.MobileNoSec").value = Mobile2
'    End If
'    If EmailAddress2 <> "" And Sheet1.Range("sheet1.EmailAddress2").Locked = False Then
'        Sheet1.Range("sheet1.EmailAddress2").value = EmailAddress2
'    End If
    If sEmpCat <> "" And Sheet1.Range("sheet1.NatureofEmployment").Locked = False Then
        Sheet1.Range("sheet1.NatureofEmployment").value = sEmpCat
    End If
       If aadhaarCardNo <> "" And Sheet1.Range("Sheet1.Aadhaar").Locked = False Then
        Sheet1.Range("Sheet1.Aadhaar").value = aadhaarCardNo
    End If
    
    'Ayush
    'Ayush_enrol
'    If Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") <= Dformat(Trim("30/09/2024"), "yyyy-mm-dd") Then
'    If CStr(AadhaarEnrolmentIds) <> "" And Sheet1.Range("Sheet1.AadhaarEnrol").Locked = False Then
'        Sheet1.Range("Sheet1.AadhaarEnrol").value = CStr(AadhaarEnrolmentIds)
'    End If
'    End If
End Function
' import json for filingStatus

Function ImportPartA_139_8A(jsonText As String)

On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson(jsonText)
Set jsonObject = jsonObject("ITR")("ITR4")
If Sheet201.Range("U_AadhaarCardNo").Locked = False Then
    Sheet201.Range("U_AadhaarCardNo").value = jsonObject("PartA_139_8A")("AadhaarCardNo")
End If

'Ayush_enrol
'If Sheet201.Range("U_AadhaarEnrolmentId").Locked = False Then
'    Sheet201.Range("U_AadhaarEnrolmentId").value = jsonObject("PartA_139_8A")("AadhaarEnrolmentId")
'End If


If Sheet201.Range("U_PreviouslyFiledForThisAY").Locked = False Then
    If jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY") = "Y" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY").value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY") = "N" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No"
    End If
End If


If Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Locked = False Then
    If jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY_139_8A") = "1" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "139(1)"
        
    ElseIf jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY_139_8A") = "2" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "Other"
    End If
End If


If Sheet201.Range("U_ITRForm").Locked = False Then
    Sheet201.Range("U_ITRForm").value = jsonObject("PartA_139_8A")("Applicable_139_8A")("ITRForm")
End If

If Sheet201.Range("U_AcknowledgementNo").Locked = False Then
    Sheet201.Range("U_AcknowledgementNo").value = jsonObject("PartA_139_8A")("Applicable_139_8A")("AcknowledgementNo")
End If

If Sheet201.Range("U_OrigRetFiledDate").Locked = False Then
    Sheet201.Range("U_OrigRetFiledDate").value = Mid(jsonObject("PartA_139_8A")("Applicable_139_8A")("OrigRetFiledDate"), 9, 2) & "/" & Mid(jsonObject("PartA_139_8A")("Applicable_139_8A")("OrigRetFiledDate"), 6, 2) & "/" & Mid(jsonObject("PartA_139_8A")("Applicable_139_8A")("OrigRetFiledDate"), 1, 4)
End If

If Sheet201.Range("U_LaidOutIn_139_8A").Locked = False Then
    If jsonObject("PartA_139_8A")("LaidOutIn_139_8A") = "Y" Then
        Sheet201.Range("U_LaidOutIn_139_8A").value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("LaidOutIn_139_8A") = "N" Then
        Sheet201.Range("U_LaidOutIn_139_8A").value = "No"
    End If
End If


Dim Nodelist, node, itemp, vRows

If jsonObject("PartA_139_8A").exists("UpdatingInc") Then

    Set Nodelist = jsonObject("PartA_139_8A")("UpdatingInc")("ReasonsForUpdatingIncDtls")
    If Nodelist.count > 0 Then
'        Dim U_ReasonsForUpdatingIncomeCol
'
'        U_ReasonsForUpdatingIncomeCol = Sheet201.Range("U_ReasonsForUpdatingIncome").Column
        

        itemp = 0
        
        If Nodelist.count > Sheet201.Range("U_ReasonsForUpdatingIncome").Rows.count Then
            Sheet201.Activate
            mdCommon.DefinedgridNameRange = ("U_ReasonsForUpdatingIncome")
            ActiveCellRange = mdCommon.searchLastRow("U_ReasonsForUpdatingIncome")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet201.Range("U_ReasonsForUpdatingIncome").Rows.count)
        End If
        
        For Each node In Nodelist
        itemp = itemp + 1
        Select Case node("ReasonsForUpdatingIncome")
        
        Case "1"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Return previously not filed"
        Case "2"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Income not reported correctly"
        Case "3"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Wrong heads of income chosen"
        Case "4"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Reduction of carried forward loss"
        Case "5"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Reduction of unabsorbed depreciation"
        Case "6"
              Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Reduction of tax credit u/s 115JB/115JC"
        Case "7"
              Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Wrong rate of tax"
        Case "OTH"
               Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Others "
        End Select
        
        
        Next node
    
    End If
    
End If

If Sheet201.Range("U_UpdatedReturnDuringPeriod").Locked = False Then
    If jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "1" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Up to 12 months from the end of Relevant Assessment Year"
        
    ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "2" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 12 to 24 Months from the end of Relevant Assessment  Year"
    
    ''Added by Aavula Naresh
         ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "3" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 24 to 36 Months from the end of Relevant Assessment  Year"
    
     ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "4" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 36 to 48 Months from the end of Relevant Assessment  Year"
    End If
    '''''''''
End If


If Sheet201.Range("U_UnabsorbedDepreciation").Locked = False Then
    If jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UnabsorbedDepreciation") = "Y" Then
        Sheet201.Range("U_UnabsorbedDepreciation").value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UnabsorbedDepreciation") = "N" Then
        Sheet201.Range("U_UnabsorbedDepreciation").value = "No"
    End If
End If

Set Nodelist = New Collection


If jsonObject("PartA_139_8A").exists("RetrntoRedCarriedFL") Then
If jsonObject("PartA_139_8A")("RetrntoRedCarriedFL").exists("UDYear") Then
    Set Nodelist = jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UDYear")("UnabsorbedDepreciationYearDtls")
    If Nodelist.count > 0 Then
'        Dim U_ReasonsForUpdatingIncomeCol
'
'        U_ReasonsForUpdatingIncomeCol = Sheet201.Range("U_ReasonsForUpdatingIncome").Column
        
        
        itemp = 0
        
        If Nodelist.count > Sheet201.Range("U_UnabsorbedDepreciationYear").Rows.count Then
            Sheet201.Activate
            mdCommon.DefinedgridNameRange = ("U_UnabsorbedDepreciationYear")
            ActiveCellRange = mdCommon.searchLastRow("U_UnabsorbedDepreciationYear")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet201.Range("U_UnabsorbedDepreciationYear").Rows.count)
        End If
        
        For Each node In Nodelist
        itemp = itemp + 1
            
            Dim YearUD As Variant
            YearUD = node("UnabsorbedDepreciationYear")
'            If YearUD <> "" Then
'                If YearUD = "2023" Then
'
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2023-24"
'                ElseIf YearUD = "2024" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2024-25"
'                End If
'            End If

        'AY 2024-25 Change-----------------------
        'Ayush_UR
'             If YearUD <> "" Then
'                If YearUD = "2024" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2024-25"
'                ElseIf YearUD = "2025" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2025-26"
'                End If
'             End If

'        If YearUD <> "" Then
'                If YearUD = "2026" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2026-27"
'                ElseIf YearUD = "2025" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2025-26"
'                End If
'             End If
        'AY 2024-25 Change-----------------------
        
         If YearUD <> "" Then
                If YearUD = "2027" Then
                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2027-28"
                ElseIf YearUD = "2026" Then
                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2026-27"
                End If
             End If
                        
            
            
            If Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If node("RevisedReturnFile") = "Y" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).value = "Yes"

                 ElseIf node("RevisedReturnFile") = "N" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).value = "No"
                End If
             End If
             
             If Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If node("UpdatedReturnFile") = "Y" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Yes"

                ElseIf node("UpdatedReturnFile") = "N" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "No"
                End If
             End If
             
             
'              If Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Locked = False Then
'                If node("ReturnType") = "1" Then
'                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Original Return 139(1)/139(4)"
'                ElseIf node("ReturnType") = "2" Then
'                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Revised return"
'                ElseIf node("ReturnType") = "3" Then
'                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Updated return"
'                End If
'             End If

'            Dim OriginalReturn, RevisedReturnFile, UpdatedReturnFile As Variant
'            OriginalReturn = Node("OriginalReturn")
'            RevisedReturnFile = Node("RevisedReturnFile")
'            UpdatedReturnFile = Node("UpdatedReturnFile")
'
'            Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Value = "No"
'            If OriginalReturn = "Y" Then
'                Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Orginal Return 139(1)/139(4)"
'                Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Value = "Yes"
'            End If
'            If RevisedReturnFile = "Y" Then
'                Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Revised return"
'                Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Value = "Yes"
'            End If
'            If UpdatedReturnFile = "Y" Then
'                Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Updated return"
'                Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Value = "Yes"
'            End If
                        
            
        
        Next node
    
    End If
End If
End If




End Function

'Function ImportFilingStatus_old(jsonText As String)
'On Error Resume Next
'
'Dim jsonObject As Object
'Dim jsonDictionary As Object
''AY 2023-24 OLD
''Dim ReturnFileSec, AsseseeRepFlg, init, OptingNewTaxRegimeflg, Node, Nodelist, NewTaxRegime, Form10IEDate, Form10IEAckNo, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar
'
''AY 2024-25 Change
'Dim ReturnFileSec, AsseseeRepFlg, init, Node, Nodelist, Form10IEADate, Form10IEAAckNo, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar
'Set jsonObject = ParseJson(jsonText)
'
'    Dim Filingtype As Variant
'    Dim sReturnFile, iReturnFile As Variant
'    Dim sPort5A, iPort5A As Variant
'    Dim DateofOriginalfile As Variant
'    Dim NoticeDateussec As Variant
'    Dim iProvisoFlag, sProvisoFlag As Variant
'    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
'    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
'    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
'
'    ReturnFileSec = jsonObject("ITR")("ITR4")("FilingStatus")("ReturnFileSec")
'    AsseseeRepFlg = jsonObject("ITR")("ITR4")("FilingStatus")("AsseseeRepFlg")
''    OptingNewTaxRegimeflg = jsonObject("ITR")("ITR4")("FilingStatus")("OptingNewTaxRegime")
''    NewTaxRegime = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegime")
''    Form10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEDate") 'AY 2023-24 Old
'
''Import-09.12.2022.103.09A
'
'     'AY 2023-24 OLD Starts--------------------------
'
''    Dim AYOptedIn, NewForm10IEDate, NewForm10IEAckNo, OldTaxRegime, AYOptedOut, OldForm10IEDate, OldForm10IEAckNo
''    AYOptedIn = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("AssessmentYear")
''    NewForm10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("Form10IEDtls")("Form10IEDate")
''    NewForm10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("Form10IEDtls")("Form10IEAckNo")
''    OldTaxRegime = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegime")
''    AYOptedOut = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegimeDtls")("AssessmentYear")
''    OldForm10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegimeDtls")("Form10IEDtls")("Form10IEDate")
''    OldForm10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegimeDtls")("Form10IEDtls")("Form10IEAckNo")
'
'     'AY 2023-24 OLD Ends-----------------------------
'
''---End Change
'
''PAG_C24, PAG_C25 AY 2024-25 Change Starts-------------------------------------------------------
'
'   Form10IEADate = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate")
'        If Form10IEADate <> "" Then
'           Dim YYYY, MM, DD
'           YYYY = Mid(Form10IEADate, 1, 4)
'           MM = Mid(Form10IEADate, 6, 2)
'           DD = Mid(Form10IEADate, 9, 2)
'           Form10IEDate = DD & "/" & MM & "/" & YYYY
'        End If
'
'   Form10IEAAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo")
'
''PAG_C24, PAG_C25 AY 2024-25 Change Ends-------------------------------------------------------
'
''    Form10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAckNo") AY 2023-24 OLD
'    SeventhProvisio139 = jsonObject("ITR")("ITR4")("FilingStatus")("SeventhProvisio139")
'    DepAmtAggAmtExcd1CrPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
'    AmtSeventhProvisio139i = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139i")
'    IncrExpAggAmt2LkTrvFrgnCntryFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
'    AmtSeventhProvisio139ii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139ii")
'    IncrExpAggAmt1LkElctrctyPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt1LkElctrctyPrYrFlg")
'    AmtSeventhProvisio139iii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139iii")
'    ReceiptNo = jsonObject("ITR")("ITR4")("FilingStatus")("ReceiptNo")
'    NoticeNo = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeNo")
'    OrigRetFiledDate = jsonObject("ITR")("ITR4")("FilingStatus")("OrigRetFiledDate")
'    NoticeDateUnderSec = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeDateUnderSec")
'
'    'PAG_C9 AY 2024-25 Schema implementation ------------------------
'    ItrFilingDueDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("ItrFilingDueDate")
'
'    RepName = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepName")
'    RepCapacity = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepCapacity")
'    RepAddress = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepAddress")
'    RepPAN = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepPAN")
'    RepAadhaar = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepAadhaar")
'
'        iReturnFile = ReturnFileSec
'
'        If iReturnFile = "11" Then
'           sReturnFile = "139(1)-On or before due date"
'        ElseIf iReturnFile = "12" Then
'           sReturnFile = "139(4)-After due date"
'        ElseIf iReturnFile = "13" Then
'            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "142(1)"
'        ElseIf iReturnFile = "14" Then
'            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "148"
'        ElseIf iReturnFile = "15" Then
'            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "153A"
'        ElseIf iReturnFile = "16" Then
'        Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'        Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "153C"
'        ElseIf iReturnFile = "17" Then
'           sReturnFile = "139(5)-Revised Return"
'        ElseIf iReturnFile = "18" Then
'        Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'        Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "139(9)"
'        ElseIf iReturnFile = "20" Then
'           sReturnFile = "119(2)(b)- after condonation of delay"
'        ElseIf iReturnFile = "21" Then
'           sReturnFile = "139(8A)"
'        End If
'
'        If Sheet1.Range("sheet1.ReturnFileSec").Locked = False Then
'            Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
'        End If
'
'
'
'    'AY 2024-25 Change Schema implement Starts--------------
'
'       ItrFilingDueDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("ItrFilingDueDate")
'        YYYY = Mid(ItrFilingDueDate_1, 6, 4)
'        MM = Mid(ItrFilingDueDate_1, 1, 1)
'        DD = Mid(ItrFilingDueDate_1, 3, 2)
'        strDate = DD & "/" & MM & "/" & YYYY
'        MsgBox strDate
'        Sheet1.Range("Sheet1.Duedate").value = strDate
'
'      'AY 2024-25 Change Schema implement End--------------
'
'
'
'
'
'
'
'
'        iProvisoFlag = SeventhProvisio139
'        If iProvisoFlag = "Y" Then
'        sProvisoFlag = "Yes"
'        ElseIf iProvisoFlag = "N" Then
'        sProvisoFlag = "No"
'        End If
'
'
'
'
''        If NewTaxRegime = "Y" Then
''            Sheet8.Unprotect Password:=getmsgstate
''             Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
''                Sheet8.Range("bacValue").value = 1
''            Sheet8.Protect Password:=getmsgstate
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacYes
''            Sheet1.Protect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''                Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''            End If
''            If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''                Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''            End If
''        ElseIf NewTaxRegime = "N" Then
''            Sheet8.Unprotect Password:=getmsgstate
''                Sheet8.Range("bacValue").value = 2
''                Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
''            Sheet8.Protect Password:=getmsgstate
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacNo
''            Sheet1.Protect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''                Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''            End If
''            If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''            Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''            End If
''        Else
''            Sheet8.Shapes("BacYes").OLEFormat.Object.value = 0
''            Sheet8.Shapes("BacNO").OLEFormat.Object.value = 0
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacNo
''            Sheet1.Protect Password:=getmsgstate
''        End If
'
''        If OptingNewTaxRegimeflg < 3 Then
''        Sheet1.Range("OptingNewTaxRegime").value = "No"
''        Else
''        Sheet1.Range("OptingNewTaxRegime").value = "Yes"
''        End If
'
'
''Import-09.12.2022.103.09B
'
'If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
'    If NewTaxRegime = "Y" Then
'            Sheet8.Unprotect Password:=getmsgstate
'             Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
'                Sheet8.Range("bacValue").value = 1
'                Sheet8.Range("oldbacValue").value = 1
'            Sheet8.Protect Password:=getmsgstate
'            Sheet1.Unprotect Password:=getmsgstate
'            resetBacYes
'            Sheet1.Protect Password:=getmsgstate
'        ElseIf NewTaxRegime = "N" Then
'            Sheet8.Unprotect Password:=getmsgstate
'                Sheet8.Range("bacValue").value = 2
'                Sheet8.Range("oldbacValue").value = 2
'                Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
'            Sheet8.Protect Password:=getmsgstate
'            Sheet1.Unprotect Password:=getmsgstate
'            resetBacNo
'            Sheet1.Protect Password:=getmsgstate
'        Else
'            Sheet8.Shapes("BacYes").OLEFormat.Object.value = 0
'            Sheet8.Shapes("BacNO").OLEFormat.Object.value = 0
'            Sheet1.Unprotect Password:=getmsgstate
'            resetBacNo
'            Sheet1.Protect Password:=getmsgstate
'    End If
'End If
'
'    Application.EnableEvents = True
'
'    If Sheet1.Range("AYinOpting").Locked = False Then
'            Sheet1.Range("AYinOpting").value = AYOptedIn
'    End If
'
'    If Sheet1.Range("Sheet1.date10IE_Opt").Locked = False Then
'            Sheet1.Range("Sheet1.date10IE_Opt").value = Mid(NewForm10IEDate, 9, 2) & "/" & Mid(NewForm10IEDate, 6, 2) & "/" & Mid(NewForm10IEDate, 1, 4)
'    End If
'
'    If Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Locked = False Then
'            Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = NewForm10IEAckNo
'    End If
'
'    If Sheet1.Range("Opt_Out115BAC").Locked = False Then
'            Application.EnableEvents = True
'            If OldTaxRegime = "Y" Then
'                Sheet1.Range("Opt_Out115BAC").value = "Yes"
'            ElseIf OldTaxRegime = "N" Then
'                Sheet1.Range("Opt_Out115BAC").value = "No"
'            End If
'            'Application.EnableEvents = False
'    End If
'
'    If Sheet1.Range("AYinOptOut").Locked = False Then
'        Sheet1.Range("AYinOptOut").value = AYOptedOut
'    End If
'
'    If Sheet1.Range("Sheet1.date10IE_OptOut").Locked = False Then
'            Sheet1.Range("Sheet1.date10IE_OptOut").value = Mid(OldForm10IEDate, 9, 2) & "/" & Mid(OldForm10IEDate, 6, 2) & "/" & Mid(OldForm10IEDate, 1, 4)
'    End If
'
'    If Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Locked = False Then
'            Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = OldForm10IEAckNo
'    End If
'
'
''End Change
'If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
'If OptingNewTaxRegimeflg = 1 Then
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").value = 1
'    Sheet8.Range("oldbacValue").value = 2
'    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
'    Sheet1.Unprotect Password:=getmsgstate
'    'resetBacYes
'    Sheet1.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "Opting in now"
'    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
'        Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
'    End If
'    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
'        Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
'    End If
'    Sheet8.Protect Password:=getmsgstate
'End If
'
'If OptingNewTaxRegimeflg = 2 Then
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").value = 2
'    Sheet8.Range("oldbacValue").value = 2
'    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
'    Sheet1.Unprotect Password:=getmsgstate
'    'resetBacNo
'    Sheet1.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "Not opting"
'    Sheet8.Protect Password:=getmsgstate
'End If
'
'If OptingNewTaxRegimeflg = 3 Then
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").value = 1
'    Sheet8.Range("oldbacValue").value = 1
'    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
'    Sheet1.Unprotect Password:=getmsgstate
'    'resetBacYes
'    Sheet1.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt"
'    Sheet8.Protect Password:=getmsgstate
'End If
'
'If OptingNewTaxRegimeflg = 4 Then
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").value = 2
'    Sheet8.Range("oldbacValue").value = 1
'    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
'    Sheet1.Unprotect Password:=getmsgstate
'    'resetBacNo
'    Sheet1.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "Opt out"
'    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
'    Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
'    End If
'    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
'    End If
'    Sheet8.Protect Password:=getmsgstate
'End If
'
''Import-09.12.2022.103.16
'
'If OptingNewTaxRegimeflg = 5 Then
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").value = 2
'    'Sheet8.Range("oldbacValue").value = 2
'    'Sheet1.Shapes("BacNo").OLEFormat.Object.value = 2
'    'Sheet1.Unprotect Password:=getmsgstate
'    'resetBacNo
'    'Sheet1.Protect Password:=getmsgstate
'    'Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in"
''    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''        Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''    End If
''    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''        Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''    End If
'    Sheet8.Protect Password:=getmsgstate
'End If
'
''End Change
'End If
'
'         Application.EnableEvents = True
'
'         If Sheet1.Range("sheet1.SeventhProvisoFlag").Locked = False Then
'            Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
'        End If
'        iDepositAmountFlag = DepAmtAggAmtExcd1CrPrYrFlg
'        If iDepositAmountFlag = "Y" Then
'        sDepositAmountFlag = "Yes"
'        ElseIf iDepositAmountFlag = "N" Then
'        sDepositAmountFlag = "No"
'        End If
'        If Sheet1.Range("Sheet1.DepositAmountFlag").Locked = False Then
'            Sheet1.Range("Sheet1.DepositAmountFlag").value = sDepositAmountFlag
'        End If
'        If Sheet1.Range("Sheet1.DepositAmount").Locked = False Then
'            Sheet1.Range("Sheet1.DepositAmount").value = AmtSeventhProvisio139i
'        End If
'        iAggrigateAmountFlag = IncrExpAggAmt2LkTrvFrgnCntryFlg
'        If iAggrigateAmountFlag = "Y" Then
'        sAggrigateAmountFlag = "Yes"
'        ElseIf iAggrigateAmountFlag = "N" Then
'        sAggrigateAmountFlag = "No"
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmountFlag").value = sAggrigateAmountFlag
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmount").value = AmtSeventhProvisio139ii
'        End If
'        iAggrigateAmountFlag1 = IncrExpAggAmt1LkElctrctyPrYrFlg
'        If iAggrigateAmountFlag1 = "Y" Then
'        sAggrigateAmountFlag1 = "Yes"
'        ElseIf iAggrigateAmountFlag1 = "N" Then
'        sAggrigateAmountFlag1 = "No"
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmountFlag1").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = sAggrigateAmountFlag1
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmount1").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmount1").value = AmtSeventhProvisio139iii
'        End If
'
'
'        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1").value
'
'        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
'
'        If ReturnFileSec = "17" Then
'            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
'                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            DateofOriginalfile = OrigRetFiledDate
'            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
'                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
'            End If
'        Else
'            If Sheet1.Range("Sheet1.ReceiptNo").Locked = False Then
'                Sheet1.Range("Sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            DateofOriginalfile = OrigRetFiledDate
'            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
'                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
'            End If
'            If Sheet1.Range("sheet1.NoticeNo").Locked = False Then
'                Sheet1.Range("sheet1.NoticeNo").value = NoticeNo
'            End If
'            NoticeDateussec = NoticeDateUnderSec
'            If NoticeDateussec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
'                Sheet1.Range("sheet1.NoticeDate").value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
'            End If
'        End If
'
'
'
'
'
''            iPort5A = UCase(Node.SelectSingleNode("PortugeseCC5A").Text)
''            If iPort5A = "Y" Then
''            sPort5A = "Yes"
''
''            End If
'
''            If iPort5A = "N" Then
''            sPort5A = "No"
'
''            End If
'
''import 139 clause iv
'
' Set init = jsonObject("ITR")("ITR4")("FilingStatus")
'    If init.exists("clauseiv7provisio139i") Then
'        If init("clauseiv7provisio139i") = "Y" Then
'        Sheet1.Range("clauseiv7provisio139iFlg") = "Yes"
'        Sheet1.Range("clauseiv7provisio139iFlg_1") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_2") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_3") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_4") = "No"
'
'            For Each Node In init("clauseiv7provisio139iDtls")
'                If Node("clauseiv7provisio139iNature") = "1" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_1") = Node("clauseiv7provisio139iAmount")
'                ElseIf Node("clauseiv7provisio139iNature") = "2" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_2") = Node("clauseiv7provisio139iAmount")
'                ElseIf Node("clauseiv7provisio139iNature") = "3" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_3") = Node("clauseiv7provisio139iAmount")
'                ElseIf Node("clauseiv7provisio139iNature") = "4" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_4") = Node("clauseiv7provisio139iAmount")
'                End If
'            Next Node
'
'        Else
'            Sheet1.Range("clauseiv7provisio139iFlg") = "No"
'        End If
'
'
'    End If
'
'
'
'
'
'
'           Dim RepAssessee As Variant
'           Dim sRepAssessee As Variant
'
'           RepAssessee = AsseseeRepFlg
'           If RepAssessee = "Y" Then
'               sRepAssessee = "Yes"
'
'           ElseIf RepAssessee = "N" Then
'               sRepAssessee = "No"
'           End If
'           If Sheet1.Range("Sheet1.RepAssessee").Locked = False Then
'            Sheet1.Range("Sheet1.RepAssessee").value = sRepAssessee
'           End If
'           If Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.NameRepAssessee").value = RepName
'            End If
'           Dim iCapacity As Variant
'           Dim sCapacity As Variant
'
'           iCapacity = RepCapacity
'
'           If iCapacity = "L" Then
'                sCapacity = "Legal Heir"
'
'           ElseIf iCapacity = "M" Then
'                sCapacity = "Manager"
'
'           ElseIf iCapacity = "G" Then
'                sCapacity = "Guardian"
'
'           ElseIf iCapacity = "O" Then
'                sCapacity = "Other"
'
'           End If
'           If Sheet1.Range("sheet1.CapacityRep").Locked = False Then
'            Sheet1.Range("sheet1.CapacityRep").value = sCapacity
'           End If
'           If Sheet1.Range("sheet1.AddressRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.AddressRepAssessee").value = RepAddress
'           End If
'           If Sheet1.Range("sheet1.PANRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.PANRepAssessee").value = RepPAN
'        End If
'        If Sheet1.Range("sheet1.AadhaarRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.AadhaarRepAssessee").value = RepAadhaar
'        End If
'
''          Sheet1.Range("sheet1.PANOFSPOUSE").value = UCase(Node.SelectSingleNode("PANOfSpouse").Text)
'
'End Function
Function ImportPartB_ATI(jsonText As String)

On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson(jsonText)
Set jsonObject = jsonObject("ITR")("ITR4")("PartB-ATI")

If Sheet202.Range("U_Salaries").Locked = False Then
    Sheet202.Range("U_Salaries").value = jsonObject("HeadOfInc")("Salaries")
End If
If Sheet202.Range("U_IncomeFromHP").Locked = False Then
    Sheet202.Range("U_IncomeFromHP").value = jsonObject("HeadOfInc")("IncomeFromHP")
End If
If Sheet202.Range("U_IncomeFromBP").Locked = False Then
    Sheet202.Range("U_IncomeFromBP").value = jsonObject("HeadOfInc")("IncomeFromBP")
End If
'If Sheet202.Range("U_IncomeFromCG").Locked = False Then
'    Sheet202.Range("U_IncomeFromCG").value = jsonObject("HeadOfInc")("IncomeFromCG")
'End If
If Sheet202.Range("U_IncomeFromOS").Locked = False Then
    Sheet202.Range("U_IncomeFromOS").value = jsonObject("HeadOfInc")("IncomeFromOS")
End If




If Sheet202.Range("U_LatestTotInc").Locked = False Then
    Sheet202.Range("U_LatestTotInc").value = jsonObject("LatestTotInc")
End If

If Sheet202.Range("U_LastAmtPayable").Locked = False Then
    Sheet202.Range("U_LastAmtPayable").value = jsonObject("LastAmtPayable")
End If

If Sheet202.Range("U_Refund").Locked = False Then
    Sheet202.Range("U_Refund").value = jsonObject("Refund")
End If

If Sheet202.Range("U_TotRefund").Locked = False Then
    Sheet202.Range("U_TotRefund").value = jsonObject("TotRefund")
End If
If Sheet202.Range("U_RegAssessementTAX").Locked = False Then
    Sheet202.Range("U_RegAssessementTAX").value = jsonObject("RegAssessementTAX")
End If



Dim Nodelist, node, itemp, vRows

If jsonObject.exists("ScheduleIT1") Then
    If jsonObject("ScheduleIT1").exists("TaxPayment1") Then
        Set Nodelist = jsonObject("ScheduleIT1")("TaxPayment1")("ITTaxPayments")
        itemp = 0
        If Nodelist.count > Sheet202.Range("U_BSRCode1").Rows.count Then
            Sheet202.Activate
            mdCommon.DefinedgridNameRange = ("U_slno1||U_BSRCode1||U_DateDep1||U_SrlNoOfChaln1||U_Amt1||U_DateOfAmount1")
            ActiveCellRange = mdCommon.searchLastRow("U_BSRCode1")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet202.Range("U_BSRCode1").Rows.count)
        End If
        
        
        For Each node In Nodelist
            itemp = itemp + 1
            
            Sheet202.Range("U_BSRCode1").Rows(itemp).Cells(1).value = node("BSRCode")
            
            Sheet202.Range("U_DateDep1").Rows(itemp).Cells(1).value = Mid(node("DateDep"), 9, 2) & "/" & Mid(node("DateDep"), 6, 2) & "/" & Mid(node("DateDep"), 1, 4)
            
            Sheet202.Range("U_SrlNoOfChaln1").Rows(itemp).Cells(1).value = node("SrlNoOfChaln")
            
            Sheet202.Range("U_Amt1").Rows(itemp).Cells(1).value = node("Amt")
        
        Next node
    
    End If
    


End If
Nodelist = New Collection


If jsonObject.exists("ScheduleIT2") Then
    If jsonObject("ScheduleIT2").exists("TaxPayment2") Then
        Set Nodelist = jsonObject("ScheduleIT2")("TaxPayment2")("ITTaxPayments")
        itemp = 0
        If Nodelist.count > Sheet202.Range("U_BSRCode2").Rows.count Then
            Sheet202.Activate
            mdCommon.DefinedgridNameRange = ("U_slno2||U_BSRCode2||U_DateDep2||U_SrlNoOfChaln2||U_Amt2")
            ActiveCellRange = mdCommon.searchLastRow("U_BSRCode2")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet202.Range("U_BSRCode2").Rows.count)
        End If
        
        
        For Each node In Nodelist
            itemp = itemp + 1
            
            Sheet202.Range("U_BSRCode2").Rows(itemp).Cells(1).value = node("BSRCode")
            
            Sheet202.Range("U_DateDep2").Rows(itemp).Cells(1).value = Mid(node("DateDep"), 9, 2) & "/" & Mid(node("DateDep"), 6, 2) & "/" & Mid(node("DateDep"), 1, 4)
            
            Sheet202.Range("U_SrlNoOfChaln2").Rows(itemp).Cells(1).value = node("SrlNoOfChaln")
            
            Sheet202.Range("U_Amt2").Rows(itemp).Cells(1).value = node("Amt")
        
        Next node
    
    End If
    


End If



If Sheet202.Range("U_ReleifUS89").Locked = False Then
    Sheet202.Range("U_ReleifUS89").value = jsonObject("ReleifUS89")
End If




End Function
Function ImportITR4_IncomeDeductions(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, node, Nodelist As Object

Dim NatureColNo, DescriptionColNo, AmtColNo, TotalExRow, TotalDiffRow, TotalXMLRow, rowcount, cnt, RecTDS1 As Long
Dim SalNatureDesc, DeductionUs57iia, IncomeFromSal, TotalIncomeOfHP, IncomeOthSrc, GrossTotIncome, Section80C, Section80CCC, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80D, Section80DD, Section80DDB As Variant
Dim InterestPayable, Section80E, Section80EE, Section80G, Section80GG, Section80GGA, Section80GGC, Section80U, Section80TTA, Section80TTB, Section80DDUsrType, section80DDBUsrType, Section80EEA, Section80EEB, Section80UUsrType, AnyOthSec80CCH As Variant
Dim ArrearsUnrealizedRentRcvd, TotalIncome, GrossSalary, Salary, IncomeNotified89A, IncomeNotifiedOther89A, PerquisitesValue, ProfitsInSalary, DeductionUs16, DeductionUs16ia, EntertainmentAlw16ii, ProfessionalTaxUs16iii, TypeOfHP, GrossRentReceived, TaxPaidlocalAuth As Variant
Dim test As Variant
Dim init, Node2 As Variant

'New Schema updated by Konda as AY-2026-27 on 29-12-2025
Dim Sec80CCC_TypeColNo, Sec80CCC_NameColNo, Sec80CCC_AmountColNo, Sec80CCC_1_TypeColNo, Sec80CCC_1_NameColNo, Sec80CCC_1_AmountColNo, Sec80CCC_1b_TypeColNo, Sec80CCC_1b_NameColNo, Sec80CCC_1b_AmountColNo

Set jsonObject = ParseJson(jsonText)

GrossSalary = jsonObject("ITR")("ITR4")("IncomeDeductions")("GrossSalary")
Salary = jsonObject("ITR")("ITR4")("IncomeDeductions")("Salary")
'IncomeNotified89A = jsonObject("ITR")("ITR4")("IncomeNotified89A")("Salary")
'IncomeNotifiedOther89A = jsonObject("ITR")("ITR4")("IncomeDeductions")("IncomeNotifiedOther89A")
IncomeNotified89A = jsonObject("ITR")("ITR4")("IncomeDeductions")("IncomeNotified89A")
PerquisitesValue = jsonObject("ITR")("ITR4")("IncomeDeductions")("PerquisitesValue")
ProfitsInSalary = jsonObject("ITR")("ITR4")("IncomeDeductions")("ProfitsInSalary")
DeductionUs16 = jsonObject("ITR")("ITR4")("IncomeDeductions")("DeductionUs16")
DeductionUs16ia = jsonObject("ITR")("ITR4")("IncomeDeductions")("DeductionUs16ia")
EntertainmentAlw16ii = jsonObject("ITR")("ITR4")("IncomeDeductions")("EntertainmntalwncUs16ii")
ProfessionalTaxUs16iii = jsonObject("ITR")("ITR4")("IncomeDeductions")("ProfessionalTaxUs16iii")
'Commented by Konda on 23-01-2026================
'TypeOfHP = jsonObject("ITR")("ITR4")("IncomeDeductions")("TypeOfHP")
'GrossRentReceived = jsonObject("ITR")("ITR4")("IncomeDeductions")("GrossRentReceived")
'TaxPaidlocalAuth = jsonObject("ITR")("ITR4")("IncomeDeductions")("TaxPaidlocalAuth")
'InterestPayable = jsonObject("ITR")("ITR4")("IncomeDeductions")("InterestPayable")
'ArrearsUnrealizedRentRcvd = jsonObject("ITR")("ITR4")("IncomeDeductions")("ArrearsUnrealizedRentRcvd")
DeductionUs57iia = jsonObject("ITR")("ITR4")("IncomeDeductions")("DeductionUs57iia")
IncomeOthSrc = jsonObject("ITR")("ITR4")("IncomeDeductions")("IncomeOthSrc")
Section80C = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80C")
Section80CCC = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCC")
Section80CCDEmployeeOrSE = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCDEmployeeOrSE")
Section80CCD1B = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCD1B")
Section80CCDEmployer = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCDEmployer")
Section80D = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80D")
'Section80DD = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80DD")
Section80DDB = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDB")
Section80E = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80E")
Section80EE = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80EE")
Section80G = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80G")
Section80GG = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80GG")
Section80GGA = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80GGA")
'Section80GGC = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80GGC")
'Section80U = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80U")
Section80TTA = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80TTA")
Section80TTB = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80TTB")
'PAG_C61 AY 2024-25 Change
AnyOthSec80CCH = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("AnyOthSec80CCH")

'Section80DDUsrType = Findtext(jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDUsrType"), "Selection80DD")
section80DDBUsrType = Findtext(jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDBUsrType"), "Selection80DDB")
Section80EEA = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80EEA")
Section80EEB = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80EEB")
'Section80UUsrType = Findtext(jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Section80UUsrType"), "Selection80U")
AnyOthSec80CCH = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("AnyOthSec80CCH")
    

'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.AnnualValue").Value = jsonObject("ITR")("ITR4")("IncomeDeductions")("AnnualValue")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.StandardDeduction").Value = jsonObject("ITR")("ITR4")("IncomeDeductions")("StandardDeduction")

   ' Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD_dq5div").Locked = False Then
        Sheet1.Range("IncD_dq5div").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersInc")("DividendInc")("DateRange")("Up16Of12To15Of3")
    End If
   ' Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD_dq4div").Locked = False Then
        Sheet1.Range("IncD_dq4div").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersInc")("DividendInc")("DateRange")("Up16Of3To31Of3")
    End If
   ' Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD_dq3div").Locked = False Then
        Sheet1.Range("IncD_dq3div").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersInc")("DividendInc")("DateRange")("Up16Of9To15Of12")
    End If
   ' Sheet1.Unprotect Password:=getmsgstate
   
   'Newly added by Bindu
   'If Sheet1.Range("IncD_dqdiv").Locked = False Then
    '    Sheet1.Range("IncD_dqdiv").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersIncDtlsOth")("OthSrcOthAmount")
    'End If
    
   
    If Sheet1.Range("IncD_dq1div").Locked = False Then
        Sheet1.Range("IncD_dq1div").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersInc")("DividendInc")("DateRange")("Upto15Of6")
    End If
    'Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD_dq2div").Locked = False Then
        Sheet1.Range("IncD_dq2div").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersInc")("DividendInc")("DateRange")("Upto15Of9")
    End If
    'Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD.IncomeFromSal_ii").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_ii").value = Salary
    End If
    If Sheet1.Range("IncD.IncomeFromSal_iii").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_iii").value = PerquisitesValue
    End If
    If Sheet1.Range("IncD.IncomeFromSal_iv").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_iv").value = ProfitsInSalary
    End If
    
    'Change-06.04.2023.103.ID.04
    'PAG_C61 AY 2023-24 OLD
'    If Sheet1.Range("IncD.AnyOther").Locked = False Then
'        Sheet1.Range("IncD.AnyOther").value = "80CCH-Contribution to Agnipath Scheme"
'        Sheet1.Range("IncD.AnyOtherDeductions").value = AnyOthSec80CCH
'    End If
    
    'PAG_C61 AY 2024-25 Change Starts----------------------------------------------------
    If Sheet1.Range("IncD.AnyOtherDeductions").Locked = False Then
        Sheet1.Range("IncD.AnyOtherDeductions").value = AnyOthSec80CCH
    End If
    'PAG_C61 AY 2024-25 Ends----------------------------------------------------
    

    'End Change
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.TotalChapVIADeductions_Input").Value = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("TotalChapVIADeductions")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.TotalChapVIADeductions").Value = jsonObject("ITR")("ITR4")("IncomeDeductions")("DeductUndChapVIA")("TotalChapVIADeductions")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.TotalIncome").Value = jsonObject("ITR")("ITR4")("IncomeDeductions")("TotalIncome")

 Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")
 'Commented by Konda on 23-01-2026
'    If init.exists("IncomeNotified89AType") Then
'        For Each node In init("IncomeNotified89AType")
'            If node("NOT89ACountrycode") = "US" Then
'                Sheet1.Range("IncomeNotified89A_AmountUS").value = node("NOT89AAmount")
'            ElseIf node("NOT89ACountrycode") = "UK" Then
'                Sheet1.Range("IncomeNotified89A_AmountUK").value = node("NOT89AAmount")
'            ElseIf node("NOT89ACountrycode") = "CA" Then
'                Sheet1.Range("IncomeNotified89A_AmountCan").value = node("NOT89AAmount")
'            End If
'        Next node
'    End If
    
'    If Sheet1.Range("IncomeNotified89A").Locked = False Then
'    Sheet1.Range("IncomeNotified89A").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("IncomeNotified89A")
'    End If
    
'    If Sheet1.Range("IncomeNotifiedOther89A").Locked = False Then
'    Sheet1.Range("IncomeNotifiedOther89A").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("IncomeNotifiedOther89A")
'    End If
'================================

    Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")

    If init.exists("AllwncExemptUs10") Then
        Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("AllwncExemptUs10")("AllwncExemptUs10Dtls")

        NatureColNo = Sheet1.Range("Sheet1.Nature").Column
        DescriptionColNo = Sheet1.Range("Sheet1.Description").Column
        AmtColNo = Sheet1.Range("Sheet1.Amount").Column
        
        
        
        TotalExRow = Range("Sheet1.Nature").Rows.count
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
           ' Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("Sheet1.Nature").Locked = False Then
                Sheet1.Range("Sheet1.Nature").ClearContents
            End If
           ' Sheet1.Unprotect Password:=getmsgstate
           If Sheet1.Range("Sheet1.Description").Locked = False Then
                Sheet1.Range("Sheet1.Description").ClearContents
            End If
          '  Sheet1.Unprotect Password:=getmsgstate
          If Sheet1.Range("Sheet1.Amount").Locked = False Then
            Sheet1.Range("Sheet1.Amount").ClearContents
        End If
        End If
    
       If (TotalDiffRow > 0) Then
        AddDiffRows_Exempt1 (TotalDiffRow)
       End If
    
       rowcount = getRowNo(Sheet1.Range("Sheet1.Nature").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
        
    'Konda updated on 20-03-2026--SIT-114385
       If UCase(node("SalNatureDesc")) <> "10(13A)" And UCase(node("SalNatureDesc")) <> "OTH" And UCase(node("SalNatureDesc")) <> "oth" Then 'Newly added by Malli as per DE V4
            
            rowcount = rowcount + 1
            If Sheet1.Cells(rowcount, NatureColNo).Locked = False Then
                If UCase(node("SalNatureDesc")) = "10(10B)(I)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette "
                ElseIf UCase(node("SalNatureDesc")) = "10(10B)(II)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
                
                ElseIf UCase(node("SalNatureDesc")) = "10(14)(I)(115BAC)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(14)(i)-Allowances referred in sub-clauses (a) to (c) of sub-rule (1) in Rule 2BB"
                ElseIf UCase(node("SalNatureDesc")) = "10(14)(II)(115BAC)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee"
                
'                ElseIf UCase(node("SalNatureDesc")) = "OTH" Then
'                    Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
     'Konda updated on 09-03-2026--V0.5
                ElseIf UCase(node("SalNatureDesc")) = "10(17)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(17)-Allowance MP/MLA/MLC"
                    
     '======================
                
                ElseIf UCase(node("SalNatureDesc")) = "EIC" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Exempt income received by a judge covered under the payment of salaries to Supreme Court/High Court judges Act /Rules"
'Change-24.02.2023.101 - Validation framework
                ElseIf UCase(node("SalNatureDesc")) = "10(10)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10)-Death-cum-retirement gratuity received"
                ElseIf UCase(node("SalNatureDesc")) = "10(13A)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent"
'---
                Else
                    test = Findtext("Sec " & (node("SalNatureDesc")), "Sheet1.NOA")
                    Sheet1.Range("I" & rowcount).value = test
                End If
            End If
'Konda updated on 09-03-2026--V0.5
'               If Not node("SalOthNatOfInc") = "" And Sheet1.Range("Z" & rowcount).Locked = False Then
'                Sheet1.Range("Z" & rowcount).value = node("SalOthNatOfInc")
'               End If
'=============================
'            If Sheet1.Range("AO" & rowcount).Locked = False Then
'                Sheet1.Range("AO" & rowcount).value = node("SalOthAmount")
'            End If
            If UCase(node("SalNatureDesc")) <> "" And Sheet1.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet1.Cells(rowcount, AmtColNo).value = node("SalOthAmount")
            End If

            
            cnt = cnt + 1
            End If 'Malli
        Next node
        RecTDS1 = cnt
    End If
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Deductions_16").Value = DeductionUs16
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Deduction16ia").Value = DeductionUs16ia
    'Sheet1.Unprotect Password:=getmsgstate
    'issue of entertainment allounce
'   If Sheet1.Range("IncD.IncomeFromSal_v").Locked = False Then
'      Sheet1.Range("IncD.IncomeFromSal_v").value = EntertainmentAlw16ii
'   End If
    'issue solved
   ' Sheet1.Unprotect Password:=getmsgstate
   If Sheet1.Range("IncD.IncomeFromSal_ivc").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_ivc").value = ProfessionalTaxUs16iii
    End If
'Commented by Konda on 23-01-2026
'    If TypeOfHP <> "" And Sheet1.Range("IncD.TypeOfHP").Locked = False Then
'        If TypeOfHP = "L" Then
'            Sheet1.Range("IncD.TypeOfHP").value = "Let Out"
'        ElseIf TypeOfHP = "S" Then
'            Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied"
'        ElseIf TypeOfHP = "D" Then
'            Sheet1.Range("IncD.TypeOfHP").value = "Deemed Let Out"
'        End If
'    End If
    
     Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")

    If init.exists("OthersInc") Then
        Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("OthersInc")("OthersIncDtlsOthSrc")
    
        NatureColNo = Sheet1.Range("Sheet1.Nature1").Column
        DescriptionColNo = Sheet1.Range("Sheet1.Description1").Column
        AmtColNo = Sheet1.Range("Sheet1.Amount1").Column
        
        Dim i
        
        For Each node In Nodelist
            i = i + 1
            If UCase(node("OthSrcNatureDesc")) = "DIV" Then
            
                'Newly added by Bindu
                
                Sheet1.Range("IncD_dqdiv").value = node("OthSrcOthAmount")
                
                
                Sheet1.Range("IncD_dq4div").value = node("DividendInc")("DateRange")("Up16Of12To15Of3")
                Sheet1.Range("IncD_dq5div").value = node("DividendInc")("DateRange")("Up16Of3To31Of3")
                Sheet1.Range("IncD_dq3div").value = node("DividendInc")("DateRange")("Up16Of9To15Of12")
                Sheet1.Range("IncD_dq1div").value = node("DividendInc")("DateRange")("Upto15Of6")
                Sheet1.Range("IncD_dq2div").value = node("DividendInc")("DateRange")("Upto15Of9")
                Nodelist.Remove i
                i = i - 1
 'Commented by Konda on 23-01-2026
'            ElseIf UCase(node("OthSrcNatureDesc")) = "OTHNOT89A" Then
'                Sheet1.Range("OSIncomeNotifiedOther89A").value = node("OthSrcOthAmount")
'                Nodelist.Remove i
'                i = i - 1
                
            ElseIf UCase(node("OthSrcNatureDesc")) = "NOT89A" Then
                Sheet1.Range("IncD_q4OS1").value = node("NOT89AInc")("DateRange")("Up16Of12To15Of3")
                Sheet1.Range("IncD_q5OS1").value = node("NOT89AInc")("DateRange")("Up16Of3To31Of3")
                Sheet1.Range("IncD_q3OS1").value = node("NOT89AInc")("DateRange")("Up16Of9To15Of12")
                Sheet1.Range("IncD_q1OS1").value = node("NOT89AInc")("DateRange")("Upto15Of6")
                Sheet1.Range("IncD_q2OS1").value = node("NOT89AInc")("DateRange")("Upto15Of9")
'Commented by Konda on 23-01-2026
'                If node.exists("NOT89A") Then
'                    For Each Node2 In node("NOT89A")
'                        If Node2("NOT89ACountrycode") = "US" Then
''                            Sheet1.Range("OSIncomeNotified89A_AmountUS").value = Node2("NOT89AAmount")
'                        ElseIf Node2("NOT89ACountrycode") = "UK" Then
'                            Sheet1.Range("OSIncomeNotified89A_AmountUK").value = Node2("NOT89AAmount")
'                        ElseIf Node2("NOT89ACountrycode") = "CA" Then
'                            Sheet1.Range("OSIncomeNotified89A_AmountCan").value = Node2("NOT89AAmount")
'                        End If
'                    Next Node2
'                End If
                Nodelist.Remove i
                i = i - 1
              End If
        Next node
        
        TotalExRow = Range("Sheet1.Nature1").Rows.count
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
           ' Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("Sheet1.Nature1").Locked = False Then
                Sheet1.Range("Sheet1.Nature1").ClearContents
            End If
           ' Sheet1.Unprotect Password:=getmsgstate
           If Sheet1.Range("Sheet1.Description1").Locked = False Then
                Sheet1.Range("Sheet1.Description1").ClearContents
            End If
           ' Sheet1.Unprotect Password:=getmsgstate
           If Sheet1.Range("Sheet1.Amount1").Locked = False Then
            Sheet1.Range("Sheet1.Amount1").ClearContents
           End If
        End If
    
       If (TotalDiffRow > 0) Then
        AddDiffRows_Exempt2 (TotalDiffRow)
       End If

'        If (TotalDiffRow > 1) Then
'         AddDiffRows_Exempt2 (TotalDiffRow)
'        End If

       rowcount = getRowNo(Sheet1.Range("Sheet1.Nature1").name)
       rowcount = rowcount - 1
       cnt = 0
        
        For Each node In Nodelist
            rowcount = rowcount + 1
                If Sheet1.Cells(rowcount, NatureColNo).Locked = False Then
                'Ayush_05/07/2026
                If Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
                    If UCase(node("OthSrcNatureDesc")) = "SAV" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Saving Bank Account"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "IFD" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "TAX" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Income Tax Refund"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "FAP" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Family pension"
                    
                    ElseIf UCase(node("OthSrcNatureDesc")) = "10(11)(IP)" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "10(11)(IIP)" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "10(12)(IP)" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "10(12)(IIP)" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "OTH" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
                    End If
                Else
                     If UCase(node("OthSrcNatureDesc")) = "SAV" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Saving Bank Account"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "IFD" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "TAX" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Income Tax Refund"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "FAP" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Family pension"
                    ElseIf UCase(node("OthSrcNatureDesc")) = "OTH" Then
                        Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
                    End If
                End If
                 End If
                    If node("OthSrcOthNatOfInc") <> "" And Sheet1.Cells(rowcount, DescriptionColNo).Locked = False Then
                        Sheet1.Cells(rowcount, DescriptionColNo).value = node("OthSrcOthNatOfInc")
                    End If
                    If Sheet1.Cells(rowcount, AmtColNo).Locked = False Then
                        Sheet1.Cells(rowcount, AmtColNo).value = node("OthSrcOthAmount")
                    End If
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
    End If
    
    If Sheet1.Range("IncD.LessDeduction57").Locked = False Then
        Sheet1.Range("IncD.LessDeduction57").value = DeductionUs57iia
    End If
     Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")

    If init.exists("ExemptIncAgriOthUs10") Then
        Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("ExemptIncAgriOthUs10")("ExemptIncAgriOthUs10Dtls")
    
        NatureColNo = Sheet1.Range("Others.NOI").Column
        DescriptionColNo = Sheet1.Range("Nature_Others").Column
        AmtColNo = Sheet1.Range("Others.Amount").Column
        
        TotalExRow = Range("Others.NOI").Rows.count
    
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        If (TotalXMLRow > 0) Then
           ' Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("Others.NOI").Locked = False Then
                Sheet1.Range("Others.NOI").ClearContents
            End If
           ' Sheet1.Unprotect Password:=getmsgstate
           If Sheet1.Range("Nature_Others").Locked = False Then
                Sheet1.Range("Nature_Others").ClearContents
            End If
           ' Sheet1.Unprotect Password:=getmsgstate
           If Sheet1.Range("Others.Amount").Locked = False Then
            Sheet1.Range("Others.Amount").ClearContents
        End If
        End If
    
       If (TotalDiffRow > 0) Then
           AddDiffRows_Exempt (TotalDiffRow)
       End If
    
       rowcount = getRowNo(Sheet1.Range("Others.NOI").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
            rowcount = rowcount + 1
             If Sheet1.Cells(rowcount, NatureColNo).Locked = False Then
                If UCase(node("NatureDesc")) = "OTH" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
                End If
                
                If UCase(node("NatureDesc")) = "10(34)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(34) (Exempted Dividend Income)"
                End If
                
                If UCase(node("NatureDesc")) = "10(26AAA)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(26AAA) Any income as referred to in section 10(26AAA)"
                End If
                
                If UCase(node("NatureDesc")) = "10(26)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(26) Any income as referred to in section 10(26)"
                End If
                
                If UCase(node("NatureDesc")) = "10(19)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(19) Armed Forces Family pension in case of death during operational duty"
                End If
                
                If UCase(node("NatureDesc")) = "10(10D)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10D) Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(1)"
                End If
                
                If UCase(node("NatureDesc")) = "10(11)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(11) Statutory Provident Fund received"
                End If
                
                If UCase(node("NatureDesc")) = "10(12)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(12) Recognized Provident Fund received"
                End If
                
                If UCase(node("NatureDesc")) = "10(13)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(13) Approved superannuation fund received"
                End If
                
                If UCase(node("NatureDesc")) = "10(16)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(16) Scholarships granted to meet the cost of education"
                End If
                
                If UCase(node("NatureDesc")) = "10(17)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(17) Allowance MP/MLA/MLC"
                End If
                
                If UCase(node("NatureDesc")) = "10(17A)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(17A) Award instituted by Government"
                End If
                
                If UCase(node("NatureDesc")) = "10(18)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(18) Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
                End If
                
                If UCase(node("NatureDesc")) = "10(10BC)" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10BC) Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
                End If
                
                If UCase(node("NatureDesc")) = "DMDP" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Defense Medical Disability Pension"
                End If
            
                If UCase(node("NatureDesc")) = "AGRI" Then
                Sheet1.Cells(rowcount, NatureColNo).value = "Agriculture Income (Less than Equal to Rs.5000)"
                End If

                
            End If
              '  Sheet1.Unprotect Password:=getmsgstate
              If Sheet1.Cells(rowcount, DescriptionColNo).Locked = False Then
                Sheet1.Cells(rowcount, DescriptionColNo).value = node("OthNatOfInc")
            End If
            If Sheet1.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet1.Cells(rowcount, AmtColNo).value = node("OthAmount")
            End If
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
    End If
    
    '    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IncomeFromSal").Value = GrossSalary

'    If Sheet1.Range("IncomeNotified89A").Locked = False Then
'        Sheet1.Range("IncomeNotified89A").value = IncomeNotified89A
'    End If
'    If Sheet1.Range("IncomeNotifiedOther89A").Locked = False Then
'        Sheet1.Range("IncomeNotifiedOther89A").value = IncomeNotifiedOther89A
'    End If
    'issue of entertainment allounce
     If Sheet1.Range("IncD.IncomeFromSal_v").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_v").value = EntertainmentAlw16ii
    End If
    'issue solved
    If Sheet1.Range("IncD.InvestmentDropDown").Locked = False Then
        Sheet1.Range("IncD.InvestmentDropDown").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("DepPayInvClmUndDednVIA")
        If Sheet1.Range("IncD.InvestmentDropDown").value = "Y" Then
            Sheet1.Range("IncD.InvestmentDropDown").value = "Yes"
        ElseIf Sheet1.Range("IncD.InvestmentDropDown").value = "N" Then
            Sheet1.Range("IncD.InvestmentDropDown").value = "No"
        End If
    End If
    
    
    
   ' Sheet1.Unprotect Password:=getmsgstate
'    If Sheet1.Range("IncD.IncomeFromHP_i").Locked = False Then
'        Sheet1.Range("IncD.IncomeFromHP_i").value = GrossRentReceived
'    End If
    'Sheet1.Unprotect Password:=getmsgstate
'    If Sheet1.Range("IncD.IncomeFromHP_ii").Locked = False Then
'        Sheet1.Range("IncD.IncomeFromHP_ii").value = TaxPaidlocalAuth
'    End If
    
    'Commenting below lines as that field - Auto-populate changes made by Bindu on 23rd April 25
    'If Sheet1.Range("IncD.IncomeFromHP_v").Locked = False Then
     '   Sheet1.Range("IncD.IncomeFromHP_v").value = InterestPayable
    'End If
    '--ends
    
'    If Sheet1.Range("IncD.IncomeFromHP_vi").Locked = False Then
'        Sheet1.Range("IncD.IncomeFromHP_vi").value = ArrearsUnrealizedRentRcvd
'    End If
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IncomeFromOS").Value = IncomeOthSrc

'Konda update on 21-04-2025 as Vo.6.2
'    If Sheet1.Range("IncD.Section80C").Locked = False Then
'        Sheet1.Range("IncD.Section80C").value = Section80C
'    End If
'Konda update on 08-05-2025 as Vo.7
    If Sheet1.Range("IncD.Section80CCC").Locked = False Then
        Sheet1.Range("IncD.Section80CCC").value = Section80CCC
    End If
'-----------------------------------------
'=================='New Schema updated by Konda as AY-2026-27 on 31-12-2025==================================
If Section80CCC <> "" And Section80CCC > 0 Then

Dim NodeList80CCC
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80CCC = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")
 
If NodeList80CCC.exists("PensionContribution80CCC") Then


Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")

If init.exists("PensionContribution80CCC") Then

        Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("PensionContribution80CCC")

            Sec80CCC_TypeColNo = Sheet1.Range("Sec80CCC_Type").Column
            Sec80CCC_NameColNo = Sheet1.Range("Sec80CCC_Name").Column
            Sec80CCC_AmountColNo = Sheet1.Range("Sec80CCC_Amount").Column
        
            TotalExRow = Range("Sec80CCC_Type").Rows.count
            TotalXMLRow = Nodelist.count
            TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
          
            If Sheet1.Range("Sec80CCC_Type").Locked = False Then
                Sheet1.Range("Sec80CCC_Type").ClearContents
            End If
            
            If Sheet1.Range("Sec80CCC_Name").Locked = False Then
                Sheet1.Range("Sec80CCC_Name").ClearContents
            End If
            
            If Sheet1.Range("Sec80CCC_Amount").Locked = False Then
                Sheet1.Range("Sec80CCC_Amount").ClearContents
            End If
        End If
     
        If (TotalDiffRow > 0) Then
            Dim VDA80CCC_idi As Long
            VDA80CCC_idi = 0
                For VDA80CCC_idi = 1 To TotalDiffRow
                    Sheet1.Activate
                    AddDiffRows_80CCCTable (1)
                Next
        End If
    
       rowcount = getRowNo(Sheet1.Range("Sec80CCC_Type").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
            
            rowcount = rowcount + 1
            
            If Sheet1.Cells(rowcount, Sec80CCC_TypeColNo).Locked = False Then
            
               'Malli updated as per V0.9 17/04/2026
              Dim TypeofIdentifier_80ccc, TypeofIdentifier_80cccomport As Variant
              TypeofIdentifier_80ccc = node("TypeofIdentifier")
            
                    If TypeofIdentifier_80ccc <> "" Then
                            If TypeofIdentifier_80ccc = "PRAN" Then
                               TypeofIdentifier_80cccomport = "PRAN"
                            ElseIf TypeofIdentifier_80ccc = "OTHPRAN" Then
                               TypeofIdentifier_80cccomport = "Other than PRAN"
                            Else
                               TypeofIdentifier_80cccomport = ""
                            End If
                       'Sheet1.Cells(rowcount, Sec80CCC_TypeColNo).value = node("TypeofIdentifier")
                       Sheet1.Cells(rowcount, Sec80CCC_TypeColNo).value = TypeofIdentifier_80cccomport
                    End If
            
            
            
                
            End If
            
            If Sheet1.Cells(rowcount, Sec80CCC_NameColNo).Locked = False Then
                Sheet1.Cells(rowcount, Sec80CCC_NameColNo).value = node("NameofIdentifier")
            End If
            
            If Sheet1.Cells(rowcount, Sec80CCC_AmountColNo).Locked = False Then
                Sheet1.Cells(rowcount, Sec80CCC_AmountColNo).value = node("Amount")
            End If
            
            cnt = cnt + 1
    
        Next node
        
        RecTDS1 = cnt
    End If
End If
End If

'========================================================================
    If Sheet1.Range("IncD.Section80CCD_SE").Locked = False Then
        Sheet1.Range("IncD.Section80CCD_SE").value = Section80CCDEmployeeOrSE
    End If
   
   
'==============='New Schema updated by Konda as AY-2026-27 on 31-12-2025=====================================
'Malli commented as per V0.9  17/04/2026
'If Section80CCDEmployeeOrSE <> "" And Section80CCDEmployeeOrSE > 0 Then
'
'Dim NodeList80CCCD1
'
'Set jsonObject = ParseJson(jsonText)
'Set NodeList80CCCD1 = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")
'
'If NodeList80CCCD1.exists("PensionContribution80CCD1") Then
'
'
'Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")
'
'If init.exists("PensionContribution80CCD1") Then
'        Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("PensionContribution80CCD1")
'
'            Sec80CCC_1_TypeColNo = Sheet1.Range("Sec80CCC_1_Type").Column
'            Sec80CCC_1_NameColNo = Sheet1.Range("Sec80CCC_1_Name").Column
'            Sec80CCC_1_AmountColNo = Sheet1.Range("Sec80CCC_1_Amount").Column
'
'            TotalExRow = Range("Sec80CCC_1_Type").Rows.count
'            TotalXMLRow = Nodelist.count
'            TotalDiffRow = TotalXMLRow - TotalExRow
'
'        If (TotalXMLRow > 0) Then
'
'            If Sheet1.Range("Sec80CCC_1_Type").Locked = False Then
'                Sheet1.Range("Sec80CCC_1_Type").ClearContents
'            End If
'
'            If Sheet1.Range("Sec80CCC_1_Name").Locked = False Then
'                Sheet1.Range("Sec80CCC_1_Name").ClearContents
'            End If
'
'            If Sheet1.Range("Sec80CCC_1_Amount").Locked = False Then
'                Sheet1.Range("Sec80CCC_1_Amount").ClearContents
'            End If
'        End If
'
'       If (TotalDiffRow > 0) Then
'            Dim VDA80CCD_idi As Long
'            VDA80CCD_idi = 0
'                For VDA80CCD_idi = 1 To TotalDiffRow
'                    Sheet1.Activate
'                    AddDiffRows_80CCCTable_1 (1)
'                Next
'        End If
'
'
'       rowcount = getRowNo(Sheet1.Range("Sec80CCC_1_Type").name)
'       rowcount = rowcount - 1
'       cnt = 0
'
'        For Each node In Nodelist
'
'            rowcount = rowcount + 1
'
'            If Sheet1.Cells(rowcount, Sec80CCC_1_TypeColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Sec80CCC_1_TypeColNo).value = node("TypeofIdentifier")
'            End If
'
'            If Sheet1.Cells(rowcount, Sec80CCC_1_NameColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Sec80CCC_1_NameColNo).value = node("NameofIdentifier")
'            End If
'
'            If Sheet1.Cells(rowcount, Sec80CCC_1_AmountColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Sec80CCC_1_AmountColNo).value = node("Amount")
'            End If
'
'            cnt = cnt + 1
'
'        Next node
'
'        RecTDS1 = cnt
'    End If
'End If
'End If
'========================================================================
    
    If Sheet1.Range("IncD.Section80CCD1B_SE").Locked = False Then
        Sheet1.Range("IncD.Section80CCD1B_SE").value = Section80CCD1B
    End If
    
    

'================'New Schema updated by Konda as AY-2026-27 on 31-12-2025====================================
'Malli commented as per V0.9  17/04/2026
'If Section80CCD1B <> "" And Section80CCD1B > 0 Then
'
'Dim NodeList80CCCD1b
'
'Set jsonObject = ParseJson(jsonText)
'Set NodeList80CCCD1b = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")
'
'If NodeList80CCCD1b.exists("PensionContribution80CCD1B") Then
'
'Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")
'
'If init.exists("PensionContribution80CCD1B") Then
'
'Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("PensionContribution80CCD1B")
'
'            Sec80CCC_1b_TypeColNo = Sheet1.Range("Sec80CCC_1b_Type").Column
'            Sec80CCC_1b_NameColNo = Sheet1.Range("Sec80CCC_1b_Name").Column
'            Sec80CCC_1b_AmountColNo = Sheet1.Range("Sec80CCC_1b_Amount").Column
'
'            TotalExRow = Range("Sec80CCC_1b_Type").Rows.count
'            TotalXMLRow = Nodelist.count
'            TotalDiffRow = TotalXMLRow - TotalExRow
'
'        If (TotalXMLRow > 0) Then
'
'            If Sheet1.Range("Sec80CCC_1b_Type").Locked = False Then
'                Sheet1.Range("Sec80CCC_1b_Type").ClearContents
'            End If
'
'            If Sheet1.Range("Sec80CCC_1b_Name").Locked = False Then
'                Sheet1.Range("Sec80CCC_1b_Name").ClearContents
'            End If
'
'            If Sheet1.Range("Sec80CCC_1b_Amount").Locked = False Then
'                Sheet1.Range("Sec80CCC_1b_Amount").ClearContents
'            End If
'        End If
'
'
'        If (TotalDiffRow > 0) Then
'            Dim VDA80CCD1b_idi As Long
'            VDA80CCD1b_idi = 0
'                For VDA80CCD1b_idi = 1 To TotalDiffRow
'                    Sheet1.Activate
'                    AddDiffRows_80CCCTable_1b (1)
'                Next
'        End If
'
'
'       rowcount = getRowNo(Sheet1.Range("Sec80CCC_1b_Type").name)
'       rowcount = rowcount - 1
'       cnt = 0
'
'        For Each node In Nodelist
'
'            rowcount = rowcount + 1
'
'            If Sheet1.Cells(rowcount, Sec80CCC_1b_TypeColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Sec80CCC_1b_TypeColNo).value = node("TypeofIdentifier")
'            End If
'
'            If Sheet1.Cells(rowcount, Sec80CCC_1b_NameColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Sec80CCC_1b_NameColNo).value = node("NameofIdentifier")
'            End If
'
'            If Sheet1.Cells(rowcount, Sec80CCC_1b_AmountColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Sec80CCC_1b_AmountColNo).value = node("Amount")
'            End If
'
'            cnt = cnt + 1
'
'        Next node
'
'        RecTDS1 = cnt
'    End If
'End If
'End If



'========================================================================
    If Sheet1.Range("IncD.Section80CCD").Locked = False Then
        Sheet1.Range("IncD.Section80CCD").value = Section80CCDEmployer
    End If


'New Schema updated by Konda as AY-2026-27 on 29-12-2025
''Konda update V0.6.2 on 21-04-2025
'Malli uncommented as per V0.9  17/04/2026
  Dim PanNum

'  PanNum = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("PRANNum")
'
'  If Sheet1.Range("Pran_Number").Locked = False Then
'
'        Sheet1.Range("Pran_Number").value = UCase(PanNum)
'    End If


Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")
If init.exists("PRANDtls") Then
Dim PRANNumlist As Object
Dim PRANNumlist_TypeColNo

Set PRANNumlist = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("PRANDtls")
   PRANNumlist_TypeColNo = Sheet1.Range("pran_new").Cells(1, 1).Column

            TotalExRow = Sheet1.Range("Pran_Sl").Rows.count
            TotalXMLRow = PRANNumlist.count
            TotalDiffRow = TotalXMLRow - TotalExRow

           If (TotalXMLRow > 0) Then
                    If Sheet1.Range("pran_new").Locked = False Then
                        Sheet1.Range("pran_new").ClearContents
                    End If
           End If
           
           If (TotalDiffRow > 0) Then
                Dim PRANNum_diff As Long
                PRANNum_diff = 0
                
                For PRANNum_diff = 1 To TotalDiffRow
                Sheet1.Activate
                AddDiffRows_PRANNum (1)
                Next
           
           End If
           
           rowcount = getRowNo(Sheet1.Range("Pran_Sl").name)
           rowcount = rowcount - 1
           cnt = 0
           
           For Each node In PRANNumlist
           
             rowcount = rowcount + 1
             
             Dim PRANNum_Import
             PRANNum_Import = node("PRANNum")
             If PRANNum_Import <> "" Then
                If Sheet1.Cells(rowcount, PRANNumlist_TypeColNo).Locked = False Then
                   Sheet1.Cells(rowcount, PRANNumlist_TypeColNo).value = PRANNum_Import
                End If
             End If
           Next node
End If
''-------------------------------
    
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80DValue").Value = Section80D

'    If Sheet1.Range("SELECT80DD").Locked = False Then
'        Sheet1.Range("SELECT80DD").value = Section80DDUsrType
'    End If
    
    'If Sheet1.Range("IncD.Section80DD").Locked = False Then
     '   Sheet1.Range("IncD.Section80DD").value = Section80DD
    'End If
    
    
    If Sheet1.Range("SELECT80DDB").Locked = False Then
        Sheet1.Range("SELECT80DDB").value = section80DDBUsrType
    End If
    'Sheet1.Specified_Disease
    'Konda-----------AY_2025_26--21-04-2025
    Dim NameOfSpecDisease80DDB_G, NameOfSpecDisease80DDB_G_enum As Variant
    
    NameOfSpecDisease80DDB_G = Sheet1.Range("Sheet1.Specified_Disease").value
    NameOfSpecDisease80DDB_G_enum = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("NameOfSpecDisease80DDB")
   
    If NameOfSpecDisease80DDB_G_enum <> "" Then
        

          If NameOfSpecDisease80DDB_G_enum = "a" Then
            NameOfSpecDisease80DDB_G = "(a) Dementia"

          ElseIf NameOfSpecDisease80DDB_G_enum = "b" Then
            NameOfSpecDisease80DDB_G = "(b) Dystonia Musculorum Deformans"

          ElseIf NameOfSpecDisease80DDB_G_enum = "c" Then
          NameOfSpecDisease80DDB_G = "(c) Motor Neuron Disease"
          

          ElseIf NameOfSpecDisease80DDB_G_enum = "d" Then
          NameOfSpecDisease80DDB_G = "(d) Ataxia"

         ElseIf NameOfSpecDisease80DDB_G_enum = "e" Then
          NameOfSpecDisease80DDB_G = "(e) Chorea"

         ElseIf NameOfSpecDisease80DDB_G_enum = "f" Then
          NameOfSpecDisease80DDB_G = "(f) Hemiballismus"

          ElseIf NameOfSpecDisease80DDB_G_enum = "g" Then
          NameOfSpecDisease80DDB_G = "(g) Aphasia"

          ElseIf NameOfSpecDisease80DDB_G_enum = "h" Then
          NameOfSpecDisease80DDB_G = "(h) Parkinsons Disease"

          ElseIf NameOfSpecDisease80DDB_G_enum = "i" Then
          NameOfSpecDisease80DDB_G = "(i) Malignant Cancers"

          ElseIf NameOfSpecDisease80DDB_G_enum = "j" Then
          NameOfSpecDisease80DDB_G = "(j) Full Blown Acquired Immuno-Deficiency Syndrome (AIDS)"

          ElseIf NameOfSpecDisease80DDB_G_enum = "k" Then
          NameOfSpecDisease80DDB_G = "(k) Chronic Renal failure"

         ElseIf NameOfSpecDisease80DDB_G_enum = "l" Then
          NameOfSpecDisease80DDB_G = "(l) Hematological disorders"

          ElseIf NameOfSpecDisease80DDB_G_enum = "m" Then
          NameOfSpecDisease80DDB_G = "(m) Hemophilia"

          ElseIf NameOfSpecDisease80DDB_G_enum = "n" Then
          NameOfSpecDisease80DDB_G = "(n) Thalassaemia"
          Else
          NameOfSpecDisease80DDB_G = ""
          
        End If
    
    If NameOfSpecDisease80DDB_G_enum <> "" Then
    Sheet1.Range("Sheet1.Specified_Disease").value = (Trim(NameOfSpecDisease80DDB_G))
    End If
    
    End If
    '--------------------------------------

    If Sheet1.Range("IncD.Section80DDB").Locked = False Then
        Sheet1.Range("IncD.Section80DDB").value = Section80DDB
    End If
'Konda updated V0.6.2 on 21-04-2025
'    If Sheet1.Range("IncD.Section80E").Locked = False Then
'        Sheet1.Range("IncD.Section80E").value = Section80E
'    End If
'    If Sheet1.Range("IncD.Section80EE").Locked = False Then
'        Sheet1.Range("IncD.Section80EE").value = Section80EE
'    End If
'
'   ' Sheet1.Unprotect Password:=getmsgstate
'    If Sheet1.Range("IncD.Section80EEA").Locked = False Then
'        Sheet1.Range("IncD.Section80EEA").value = Section80EEA
'    End If
'    If Sheet1.Range("IncD.Section80EEB").Locked = False Then
'        Sheet1.Range("IncD.Section80EEB").value = Section80EEB
'    End If
   ' Sheet1.Unprotect Password:=getmsgstate
   If Sheet1.Range("IncD.Section80G").Locked = False Then
        Sheet1.Range("IncD.Section80G").value = Section80G
    End If
   ' Sheet1.Unprotect Password:=getmsgstate
   If Sheet1.Range("IncD.Section80G_Calc").Locked = False Then
        Sheet1.Range("IncD.Section80G_Calc").value = Section80G
    End If
    If Sheet1.Range("IncD.Section80GG").Locked = False Then
        Sheet1.Range("IncD.Section80GG").value = Section80GG
    End If
 'Konda update V0.6.2 on 21-04-2025
 Dim AcmtNum_10BA
      AcmtNum_10BA = jsonObject("ITR")("ITR4")("IncomeDeductions")("UsrDeductUndChapVIA")("Form10BAAckNum")
 
   If Sheet1.Range("Sheet1.AckNum").Locked = False Then
        Sheet1.Range("Sheet1.AckNum").value = UCase(AcmtNum_10BA)
    End If
'----------------------------------------------------
    
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80GGA").Value = Section80GGA

    If Sheet1.Range("IncD.Section80TTA").Locked = False Then
        Sheet1.Range("IncD.Section80TTA").value = Section80TTA
    End If
    
   ' Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD.Section80TTB").Locked = False Then
        Sheet1.Range("IncD.Section80TTB").value = Section80TTB
    End If
    'If Sheet1.Range("IncD.Section80GGC").Locked = False Then
     '   Sheet1.Range("IncD.Section80GGC").value = Section80GGC
    'End If
    
'    If Sheet1.Range("SELECT80U").Locked = False Then
'        Sheet1.Range("SELECT80U").value = Section80UUsrType
'    End If
    
  '  If Sheet1.Range("IncD.Section80U").Locked = False Then
   '     Sheet1.Range("IncD.Section80U").value = Section80U
   ' End If
'Commented by Konda on 23-01-2026
'    Sheet1.Range("Increliefus89A").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("Increliefus89A")
    Sheet1.Range("OSIncreliefus89A").value = jsonObject("ITR")("ITR4")("IncomeDeductions")("Increliefus89AOS")
  
  'Add by Konda on 23-01-2026
  ImportScheduleHP (jsonObject)
End Function
'New Schema updated by Konda as AY-2026-27 on 31-12-2025
 Sub AddDiffRows_80CCCTable(DiffRows As Long)
    setDiffTblinfo_80CCCTable
    Sheet1.Activate
    searchLastRow ("Sec80CCC_Type")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'New Schema updated by Konda as AY-2026-27 on 31-12-2025
Sub setDiffTblinfo_80CCCTable()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sec80CCC_Type").item(0, 1).count
    Set rangecells = Sheet1.Range("Sec80CCC_Type").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Exempt1 = ccount
    DefinedgridNameRange = "Sec80CCC_Type||Sec80CCC_Name||Sec80CCC_Amount"
 End Sub
'New Schema updated by Konda as AY-2026-27 on 31-12-2025
Sub AddDiffRows_80CCCTable_1(DiffRows As Long)
    setDiffTblinfo_80CCCTable_1
    Sheet1.Activate
    searchLastRow ("Sec80CCC_1_Type")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'New Schema updated by Konda as AY-2026-27 on 31-12-2025
Sub setDiffTblinfo_80CCCTable_1()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sec80CCC_1_Type").item(0, 1).count
    Set rangecells = Sheet1.Range("Sec80CCC_1_Type").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Exempt1 = ccount
    DefinedgridNameRange = "Sec80CCC_1_Type||Sec80CCC_1_Name||Sec80CCC_1_Amount"
 End Sub
'New Schema updated by Konda as AY-2026-27 on 31-12-2025
Sub AddDiffRows_80CCCTable_1b(DiffRows As Long)
    setDiffTblinfo_80CCCTable_1b
    Sheet1.Activate
    searchLastRow ("Sec80CCC_1b_Type")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'New Schema updated by Konda as AY-2026-27 on 31-12-2025
Sub setDiffTblinfo_80CCCTable_1b()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sec80CCC_1b_Type").item(0, 1).count
    Set rangecells = Sheet1.Range("Sec80CCC_1b_Type").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Exempt1 = ccount
    DefinedgridNameRange = "Sec80CCC_1b_Type||Sec80CCC_1b_Name||Sec80CCC_1b_Amount"
 End Sub
' start import json methods
Function ImportTaxComputation(jsonText As String)
    On Error Resume Next
    Dim jsonObject, jsonDictionary As Object
    Dim TotalTaxPayable, Rebate87A, TaxPayableOnRebate, EducationCess, GrossTaxLiability, NetTaxLiability, IntrstPayUs234A, IntrstPayUs234B, IntrstPayUs234C, LateFilingFee234F, TotTaxPlusIntrstPay, Section89, Section89A As Variant

    Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Set jsonObject = ParseJson(jsonText)
    
    TotalTaxPayable = jsonObject("ITR")("ITR4")("TaxComputation")("TotalTaxPayable")
    Rebate87A = jsonObject("ITR")("ITR4")("TaxComputation")("Rebate87A")
    TaxPayableOnRebate = jsonObject("ITR")("ITR4")("TaxComputation")("TaxPayableOnRebate")
    EducationCess = jsonObject("ITR")("ITR4")("TaxComputation")("EducationCess")
    GrossTaxLiability = jsonObject("ITR")("ITR4")("TaxComputation")("GrossTaxLiability")
    NetTaxLiability = jsonObject("ITR")("ITR4")("TaxComputation")("NetTaxLiability")
    IntrstPayUs234A = jsonObject("ITR")("ITR4")("TaxComputation")("IntrstPay")("IntrstPayUs234A")
    IntrstPayUs234B = jsonObject("ITR")("ITR4")("TaxComputation")("IntrstPay")("IntrstPayUs234B")
    IntrstPayUs234C = jsonObject("ITR")("ITR4")("TaxComputation")("IntrstPay")("IntrstPayUs234C")
    LateFilingFee234F = jsonObject("ITR")("ITR4")("TaxComputation")("IntrstPay")("LateFilingFee234F")
    TotTaxPlusIntrstPay = jsonObject("ITR")("ITR4")("TaxComputation")("TotTaxPlusIntrstPay")
    Section89 = jsonObject("ITR")("ITR4")("TaxComputation")("Section89")
'    Section89A = jsonObject("ITR")("ITR4")("TaxComputation")("Section89A")
    
    
  '  Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD.TotalTaxPayable").Locked = False Then
        Sheet1.Range("IncD.TotalTaxPayable").value = TotalTaxPayable
    End If
    'Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.Rebate87A").Value = Rebate87A
    'Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.TaxPayableOnRebate").Value = TaxPayableOnRebate
    'Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.EducationCess").Value = EducationCess
    'Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.GrossTaxLiability").Value = GrossTaxLiability
   ' Sheet1.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
    If Sheet1.Range("IncD.Section89").Locked = False Then
        Sheet1.Range("IncD.Section89").value = Section89
    End If
'    If Sheet1.Range("IncD.Section89A").Locked = False Then
'        Sheet1.Range("IncD.Section89A").value = Section89A
'    End If
    Application.EnableEvents = True
    'Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.NetTaxLiability").Value = NetTaxLiability
    'Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.TotalIntrstPay").Value = jsonObject("ITR")("ITR4")("TaxComputation")("TotalIntrstPay")
   Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("IncD.IntrstPayUs234A").Locked = False Then
        Sheet1.Range("IncD.IntrstPayUs234A").value = IntrstPayUs234A
    End If
    Sheet1.Protect Password:=getmsgstate
    Sheet1.Unprotect Password:=getmsgstate
   If Sheet1.Range("IncD.IntrstPayUs234B").Locked = False Then
        Sheet1.Range("IncD.IntrstPayUs234B").value = IntrstPayUs234B
    End If
    Sheet1.Protect Password:=getmsgstate
    Sheet1.Unprotect Password:=getmsgstate
   If Sheet1.Range("IncD.IntrstPayUs234C").Locked = False Then
        Sheet1.Range("IncD.IntrstPayUs234C").value = IntrstPayUs234C
    End If
    Sheet1.Protect Password:=getmsgstate
    Sheet1.Unprotect Password:=getmsgstate
   If Sheet1.Range("IncD.IntrstPayUs234F").Locked = False Then
        Sheet1.Range("IncD.IntrstPayUs234F").value = LateFilingFee234F
    End If
    Sheet1.Protect Password:=getmsgstate
    'Sheet1.Range("IncD.TotTaxPlusIntrstPay").Value = TotTaxPlusIntrstPay

End Function

Function ImportTaxPaid(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim item As Object
Dim BalTaxPayable, AdvanceTax, TDS, TCS, SelfAssessmentTax, TotalTaxesPaid As Variant

Set jsonObject = ParseJson(jsonText)

BalTaxPayable = jsonObject("ITR")("ITR4")("TaxPaid")("BalTaxPayable")
AdvanceTax = jsonObject("ITR")("ITR4")("TaxPaid")("TaxesPaid")("AdvanceTax")
TDS = jsonObject("ITR")("ITR4")("TaxPaid")("TaxesPaid")("TDS")
TCS = jsonObject("ITR")("ITR4")("TaxPaid")("TaxesPaid")("TCS")
SelfAssessmentTax = jsonObject("ITR")("ITR4")("TaxPaid")("TaxesPaid")("SelfAssessmentTax")
TotalTaxesPaid = jsonObject("ITR")("ITR4")("TaxPaid")("TaxesPaid")("TotalTaxesPaid")


'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.AdvanceTax").value) = "" And Sheet3.Range("IncD.AdvanceTax").Locked = False Then
        Sheet3.Range("IncD.AdvanceTax").value = AdvanceTax
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.TDS").value) = "" And Sheet3.Range("IncD.TDS").Locked = False Then
        Sheet3.Range("IncD.TDS").value = TDS
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.TCS").value) = "" And Sheet3.Range("IncD.TCS").Locked = False Then
        Sheet3.Range("IncD.TCS").value = TCS
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.SelfAssessmentTax").value) = "" And Sheet3.Range("IncD.SelfAssessmentTax").Locked = False Then
        Sheet3.Range("IncD.SelfAssessmentTax").value = SelfAssessmentTax
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.TotalTaxesPaid").value) = "" And Sheet3.Range("IncD.TotalTaxesPaid").Locked = False Then
        Sheet3.Range("IncD.TotalTaxesPaid").value = TotalTaxesPaid
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.BalTaxPayable").value) = "" And Sheet3.Range("IncD.BalTaxPayable").Locked = False Then
        Sheet3.Range("IncD.BalTaxPayable").value = BalTaxPayable
    End If
    
End Function
'
'
'' import json for filingStatus
'
'Function ImportFilingInfo(jsonText As String)
'On Error Resume Next
'
'Dim jsonObject As Object
'Dim jsonDictionary As Object
'Dim ReturnFileSec, AsseseeRepFlg, NewTaxRegime, Form10IEDate, Form10IEAckNo, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar
'Dim YYYY, MM, DD, strDate, YYYY1, MM1, DD1, strDate1 As String
'Set jsonObject = ParseJson(jsonText)
'
'    Dim Filingtype As Variant
'    Dim sReturnFile, iReturnFile As Variant
'    Dim sPort5A, iPort5A As Variant
'    Dim DateofOriginalfile As String
'    Dim NoticeDateussec As String
'    Dim iProvisoFlag, sProvisoFlag As Variant
'    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
'    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
'    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
'    Dim Node, NodeList As Object
'
'
'    ReturnFileSec = jsonObject("ITR")("ITR4")("FilingStatus")("ReturnFileSec")
'    AsseseeRepFlg = jsonObject("ITR")("ITR4")("FilingStatus")("AsseseeRepFlg")
'    NewTaxRegime = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegime")
'    Form10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEDate")
'    Form10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAckNo")
'    SeventhProvisio139 = jsonObject("ITR")("ITR4")("FilingStatus")("SeventhProvisio139")
'    DepAmtAggAmtExcd1CrPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
'    AmtSeventhProvisio139i = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139i")
'    IncrExpAggAmt2LkTrvFrgnCntryFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
'    AmtSeventhProvisio139ii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139ii")
'    IncrExpAggAmt1LkElctrctyPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt1LkElctrctyPrYrFlg")
'    AmtSeventhProvisio139iii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139iii")
'    ReceiptNo = jsonObject("ITR")("ITR4")("FilingStatus")("ReceiptNo")
'    NoticeNo = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeNo")
'
'    OrigRetFiledDate = jsonObject("ITR")("ITR4")("FilingStatus")("OrigRetFiledDate")
'    If OrigRetFiledDate <> "" Then
'        DateofOriginalfile = OrigRetFiledDate
'        YYYY = Mid(DateofOriginalfile, 1, 4)
'        MM = Mid(DateofOriginalfile, 6, 2)
'        DD = Mid(DateofOriginalfile, 9, 2)
'        strDate = DD & "/" & MM & "/" & YYYY
'    End If
'
'
'    NoticeDateUnderSec = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeDateUnderSec")
'    If NoticeDateUnderSec <> "" Then
'        NoticeDateussec = NoticeDateUnderSec
'        YYYY1 = Mid(NoticeDateussec, 1, 4)
'        MM1 = Mid(NoticeDateussec, 6, 2)
'        DD1 = Mid(NoticeDateussec, 9, 2)
'        strDate1 = DD1 & "/" & MM1 & "/" & YYYY1
'    End If
'
'    RepName = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepName")
'    RepCapacity = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepCapacity")
'    RepAddress = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepAddress")
'    RepPAN = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepPAN")
'    RepAadhaar = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepAadhaar")
'
'    For Each Node In NodeList
'
'        iReturnFile = ReturnFileSec
'
'        If iReturnFile = "11" Then
'           sReturnFile = "139(1)-On or before due date"
'        ElseIf iReturnFile = "12" Then
'           sReturnFile = "139(4)-After due date"
'        ElseIf iReturnFile = "13" Then
'           sReturnFile = "142(1)"
'        ElseIf iReturnFile = "14" Then
'           sReturnFile = "148"
'        ElseIf iReturnFile = "15" Then
'           sReturnFile = "153A"
'        ElseIf iReturnFile = "16" Then
'           sReturnFile = "153C"
'        ElseIf iReturnFile = "17" Then
'           sReturnFile = "139(5)-Revised Return"
'        ElseIf iReturnFile = "18" Then
'           sReturnFile = "139(9)"
'        ElseIf iReturnFile = "20" Then
'           sReturnFile = "119(2)(b)- after condonation of delay"
'        End If
'
'        Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
'
'
'        iProvisoFlag = SeventhProvisio139
'        If iProvisoFlag = "Y" Then
'        sProvisoFlag = "Yes"
'        ElseIf iProvisoFlag = "N" Then
'        sProvisoFlag = "No"
'        End If
'
'        Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
'
'        iDepositAmountFlag = DepAmtAggAmtExcd1CrPrYrFlg
'        If iDepositAmountFlag = "Y" Then
'        sDepositAmountFlag = "Yes"
'        ElseIf iDepositAmountFlag = "N" Then
'        sDepositAmountFlag = "No"
'        End If
'
'        Sheet1.Range("Sheet1.DepositAmountFlag").value = sDepositAmountFlag
'        Sheet1.Range("Sheet1.DepositAmount").value = AmtSeventhProvisio139i
'
'        iAggrigateAmountFlag = IncrExpAggAmt2LkTrvFrgnCntryFlg
'        If iAggrigateAmountFlag = "Y" Then
'        sAggrigateAmountFlag = "Yes"
'        ElseIf iAggrigateAmountFlag = "N" Then
'        sAggrigateAmountFlag = "No"
'        End If
'
'        Sheet1.Range("Sheet1.AggrigateAmountFlag").value = sAggrigateAmountFlag
'        Sheet1.Range("Sheet1.AggrigateAmount").value = AmtSeventhProvisio139ii
'
'        iAggrigateAmountFlag1 = IncrExpAggAmt1LkElctrctyPrYrFlg
'        If iAggrigateAmountFlag1 = "Y" Then
'        sAggrigateAmountFlag1 = "Yes"
'        ElseIf iAggrigateAmountFlag1 = "N" Then
'        sAggrigateAmountFlag1 = "No"
'        End If
'
'        Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = sAggrigateAmountFlag
'        Sheet1.Range("Sheet1.AggrigateAmount1").value = AmtSeventhProvisio139iii
'
'        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
'
'        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
'
'        If ReturnFileSec = "17" Then
'            Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
'
'            If strDate <> "" Then
'                Sheet1.Range("sheet1.OrigRetFiledDate").value = strDate
'            End If
'
'        Else
'            If ReceiptNo <> "" Then
'                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            If strDate <> "" Then
'                If Sheet1.Range("sheet1.OrigRetFiledDate").value = "" Then
'                    Sheet1.Range("sheet1.OrigRetFiledDate").value = strDate
'                End If
'            End If
'            If NoticeNo <> "" Then
'                If Sheet1.Range("sheet1.NoticeNo").value = "" Then
'                    Sheet1.Range("sheet1.NoticeNo").value = NoticeNo
'                End If
'            End If
'            If strDate1 <> "" Then
'                If Sheet1.Range("sheet1.NoticeDate").value = "" Then
'                    Sheet1.Range("sheet1.NoticeDate").value = strDate1
'                End If
'            End If
'        End If
'
'
''            iPort5A = UCase(Node.SelectSingleNode("PortugeseCC5A").Text)
''            If iPort5A = "Y" Then
''            sPort5A = "Yes"
''
''            End If
'
''            If iPort5A = "N" Then
''            sPort5A = "No"
'
''            End If
'
'
'
'           Dim RepAssessee As Variant
'           Dim sRepAssessee As Variant
'
'           RepAssessee = AsseseeRepFlg
'           If RepAssessee = "Y" Then
'               sRepAssessee = "Yes"
'
'           ElseIf RepAssessee = "N" Then
'               sRepAssessee = "No"
'           End If
'
'           Sheet1.Range("Sheet1.RepAssessee").value = sRepAssessee
'
'           Sheet1.Range("sheet1.NameRepAssessee").value = RepName
'
'           Dim iCapacity As Variant
'           Dim sCapacity As Variant
'
'           iCapacity = RepCapacity
'
'           If iCapacity = "L" Then
'                sCapacity = "Legal Heir"
'
'           ElseIf iCapacity = "M" Then
'                sCapacity = "Manager"
'
'           ElseIf iCapacity = "G" Then
'                sCapacity = "Guardian"
'
'           ElseIf iCapacity = "O" Then
'                sCapacity = "Other"
'
'           End If
'
'           Sheet1.Range("sheet1.CapacityRep").value = sCapacity
'
'           Sheet1.Range("sheet1.AddressRepAssessee").value = RepAddress
'
'           Sheet1.Range("sheet1.PANRepAssessee").value = RepPAN
'
'
''          Sheet1.Range("sheet1.PANOFSPOUSE").value = UCase(Node.SelectSingleNode("PANOfSpouse").Text)
'
'    Next Node
'
'End Function

'import json for TDSonSalaries

Function ImportTDSonSalary(jsonText As String)
On Error Resume Next
    
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim rowcount, cnt As Variant
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim TotalExRow As Long

    Set jsonObject = ParseJson(jsonText)
    
    Set Nodelist = jsonObject("ITR")("ITR4")("TDSonSalaries")("TDSonSalary")
    
    
    TANNoEmployer = Sheet4.Range("TDSal.TAN").Column
    TDSNameOfEmployer = Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet4.Range("TDSal.IncChrgSal").Column
    TDSTotalTax = Sheet4.Range("TDSal.TotalTDSSal").Column
    
    TotalExRow = Range("TDSal.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet4.Range("TDSal.TAN").Locked = False Then
            Sheet4.Range("TDSal.TAN").ClearContents
        End If
        If Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Locked = False Then
            Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").ClearContents
        End If
        If Sheet4.Range("TDSal.IncChrgSal").Locked = False Then
            Sheet4.Range("TDSal.IncChrgSal").ClearContents
        End If
        If Sheet4.Range("TDSal.TotalTDSSal").Locked = False Then
            Sheet4.Range("TDSal.TotalTDSSal").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet4.Range("TDSal.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
            rowcount = rowcount + 1
            If Sheet4.Cells(rowcount, TANNoEmployer).Locked = False Then
                Sheet4.Cells(rowcount, TANNoEmployer).value = node("EmployerOrDeductorOrCollectDetl")("TAN")
            End If
            If Sheet4.Cells(rowcount, TDSNameOfEmployer).Locked = False Then
                Sheet4.Cells(rowcount, TDSNameOfEmployer).value = node("EmployerOrDeductorOrCollectDetl")("EmployerOrDeductorOrCollecterName")
            End If
            If Sheet4.Cells(rowcount, TDSIncomeCharge).Locked = False Then
                Sheet4.Cells(rowcount, TDSIncomeCharge).value = node("IncChrgSal")
            End If
            If Sheet4.Cells(rowcount, TDSTotalTax).Locked = False Then
                Sheet4.Cells(rowcount, TDSTotalTax).value = node("TotalTDSSal")
            End If
        
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

'import json for TDSonOthThanSals

Function ImportTDSOthThanSals(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim TANColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, SectionTDSColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long

Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR4")("TDSonOthThanSals")("TDSonOthThanSalDtls")
   
    TANColNo = Sheet4.Range("TDsOthr.TAN").Column
    FYColNo = Sheet4.Range("TDsOthr.finyear").Column
    TDSOtherColNo = Sheet4.Range("TDsOthr.Tds").Column
    TDSDeductedColNo = Sheet4.Range("TDsOthr.tdsdeducted").Column
    SectionTDSColNo = Sheet4.Range("TDsOthr.SectionTDS").Column 'added by Chetan C M for AY 2025-26 on 21/04/2025
    TDSClaimedColNo = Sheet4.Range("TDsOthr.tdsclaimed").Column
    TDSGrossAmountColNo = Sheet4.Range("TDsOthr.grossamount").Column
    TDSHeadColNo = Sheet4.Range("TDsOthr.headincome").Column
    
    TotalExRow = Range("TDsOthr.TAN").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        If Sheet4.Range("TDsOthr.TAN").Locked = False Then
            Sheet4.Range("TDsOthr.TAN").ClearContents
        End If
        If Sheet4.Range("TDsOthr.finyear").Locked = False Then
            Sheet4.Range("TDsOthr.finyear").ClearContents
        End If
        If Sheet4.Range("TDsOthr.Tds").Locked = False Then
            Sheet4.Range("TDsOthr.Tds").ClearContents
        End If
        If Sheet4.Range("TDsOthr.tdsdeducted").Locked = False Then
            Sheet4.Range("TDsOthr.tdsdeducted").ClearContents
        End If
        
        If Sheet4.Range("TDsOthr.SectionTDS").Locked = False Then
            Sheet4.Range("TDsOthr.SectionTDS").ClearContents
        End If
        If Sheet4.Range("TDsOthr.tdsclaimed").Locked = False Then
            Sheet4.Range("TDsOthr.tdsclaimed").ClearContents
        End If
        If Sheet4.Range("TDsOthr.grossamount").Locked = False Then
            Sheet4.Range("TDsOthr.grossamount").ClearContents
        End If
        If Sheet4.Range("TDsOthr.headincome").Locked = False Then
            Sheet4.Range("TDsOthr.headincome").ClearContents
        End If
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet4.Range("TDsOthr.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet4.Cells(rowcount, TANColNo).Locked = False Then
            Sheet4.Cells(rowcount, TANColNo).value = node("TANOfDeductor")
        End If
'Konda updated on 21-04-2025

            If Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
                    Dim SectionTDS
                        Dim TDSSection_TDS2 As Variant
                        SectionTDS = node("TDSSection")
                        If SectionTDS = "92A" Then
                        TDSSection_TDS2 = "192-Salary-Payment to Government employees other than Indian Government employees"
                        ElseIf SectionTDS = "92B" Then
                        TDSSection_TDS2 = "192-Salary-Payment to employees other than Government employees"
                        ElseIf SectionTDS = "92C" Then
                        TDSSection_TDS2 = "192-Salary-Payment to Indian Government employees"
                        ElseIf SectionTDS = "192A" Then
                        TDSSection_TDS2 = "192A-TDS on PF withdrawal"
                        ElseIf SectionTDS = "193" Then
                        TDSSection_TDS2 = "193-Interest on Securities"
                        ElseIf SectionTDS = "194" Then
                        TDSSection_TDS2 = "194-Dividends"
                        ElseIf SectionTDS = "94A" Then
                        TDSSection_TDS2 = "194A-Interest other than 'Interest on securities'"
                        ElseIf SectionTDS = "94B" Then
                        TDSSection_TDS2 = "194B-Winning from lottery or crossword puzzle"
                        ElseIf SectionTDS = "94BA" Then
                        TDSSection_TDS2 = "194BA-Winnings from online games"
                        ElseIf SectionTDS = "4BB" Then
                        TDSSection_TDS2 = "194BB-Winning from horse race"
                        ElseIf SectionTDS = "94C" Then
                        TDSSection_TDS2 = "194C-Payments to contractors and sub-contractors"
                        ElseIf SectionTDS = "94D" Then
                        TDSSection_TDS2 = "194D-Insurance commission"
                        ElseIf SectionTDS = "4DA" Then
                        TDSSection_TDS2 = "194DA-Payment in respect of life insurance policy"
                        ElseIf SectionTDS = "94E" Then
                        TDSSection_TDS2 = "194E-Payments to non-resident sportsmen or sports associations"
                        ElseIf SectionTDS = "4EE" Then
                        TDSSection_TDS2 = "194EE-Payments in respect of deposits under National Savings"
                        ElseIf SectionTDS = "4F" Then
                        TDSSection_TDS2 = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
                        ElseIf SectionTDS = "4G" Then
                        TDSSection_TDS2 = "194G-Commission, price, etc. on sale of lottery tickets"
                        ElseIf SectionTDS = "4H" Then
                        TDSSection_TDS2 = "194H-Commission or brokerage"
                        ElseIf SectionTDS = "4-IA" Then
                        TDSSection_TDS2 = "194I(a)-Rent on hiring of plant and machinery"
                        ElseIf SectionTDS = "4-IB" Then
                        TDSSection_TDS2 = "194I(b)-Rent on other than plant and machinery"
                        ElseIf SectionTDS = "4IA" Then
                        TDSSection_TDS2 = "194IA-TDS on Sale of immovable property"
                        ElseIf SectionTDS = "4IB" Then
                        TDSSection_TDS2 = "194IB-Payment of rent by certain individuals or Hindu undivided"
                        ElseIf SectionTDS = "4IC" Then
                        TDSSection_TDS2 = "194IC-Payment under specified agreement"
                        ElseIf SectionTDS = "94J-A" Then
                        TDSSection_TDS2 = "194J(a)-Fees for technical services"
                        ElseIf SectionTDS = "94J-B" Then
                        TDSSection_TDS2 = "194J(b)-Fees for professional  services or royalty etc"
                        ElseIf SectionTDS = "94K" Then
                        TDSSection_TDS2 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
                        ElseIf SectionTDS = "4LA" Then
                        TDSSection_TDS2 = "194LA-Payment of compensation on acquisition of certain immovable"
                        ElseIf SectionTDS = "4LB" Then
                        TDSSection_TDS2 = "194LB-Income by way of Interest from Infrastructure Debt fund"
                        ElseIf SectionTDS = "4LC1" Then
                        TDSSection_TDS2 = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
                        ElseIf SectionTDS = "4LC2" Then
                        TDSSection_TDS2 = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
                        ElseIf SectionTDS = "4LC3" Then
                        TDSSection_TDS2 = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
                        ElseIf SectionTDS = "4BA1" Then
                        TDSSection_TDS2 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
                        ElseIf SectionTDS = "4BA2" Then
                        TDSSection_TDS2 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
                        ElseIf SectionTDS = "LBA1" Then
                        TDSSection_TDS2 = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
                        ElseIf SectionTDS = "LBA2" Then
                        TDSSection_TDS2 = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
                        ElseIf SectionTDS = "LBA3" Then
                        TDSSection_TDS2 = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
                        ElseIf SectionTDS = "LBB" Then
                        TDSSection_TDS2 = "194LBB-Income in respect of units of investment fund"
                        ElseIf SectionTDS = "94R" Then
                        TDSSection_TDS2 = "194R-Benefits or perquisites of business or profession"
                        ElseIf SectionTDS = "94S" Then
                        TDSSection_TDS2 = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
                        ElseIf SectionTDS = "94B-P" Then
                        TDSSection_TDS2 = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
                        ElseIf SectionTDS = "94R-P" Then
                        TDSSection_TDS2 = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
                        ElseIf SectionTDS = "94S-P" Then
                        TDSSection_TDS2 = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
                        ElseIf SectionTDS = "LBC" Then
                        TDSSection_TDS2 = "194LBC-Income in respect of investment in securitization trust"
                        ElseIf SectionTDS = "4LD" Then
                        TDSSection_TDS2 = "194LD-TDS on interest on bonds / government securities"
                        ElseIf SectionTDS = "94M" Then
                        TDSSection_TDS2 = "194M-Payment of certain sums by certain individuals or HUF"
                        ElseIf SectionTDS = "94N" Then
                        TDSSection_TDS2 = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
                        ElseIf SectionTDS = "94N-F" Then
                        TDSSection_TDS2 = "194N-First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
                        ElseIf SectionTDS = "94N-C" Then
                        TDSSection_TDS2 = "194N-Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
                        ElseIf SectionTDS = "94N-FT" Then
                        TDSSection_TDS2 = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
                        ElseIf SectionTDS = "94O" Then
                        TDSSection_TDS2 = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
                        ElseIf SectionTDS = "94P" Then
                        TDSSection_TDS2 = "194P-Deduction of tax in case of specified senior citizen"
                        ElseIf SectionTDS = "94Q" Then
                        TDSSection_TDS2 = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
                        ElseIf SectionTDS = "195" Then
                        TDSSection_TDS2 = "195-Other sums payable to a non-resident"
                        ElseIf SectionTDS = "96A" Then
                        TDSSection_TDS2 = "196A-Income in respect of units of non-residents"
                        ElseIf SectionTDS = "96B" Then
                        TDSSection_TDS2 = "196B-Payments in respect of units to an offshore fund"
                        ElseIf SectionTDS = "96C" Then
                        TDSSection_TDS2 = "196C-Income from foreign currency bonds or shares of Indian"
                        ElseIf SectionTDS = "96D" Then
                        TDSSection_TDS2 = "196D-Income of foreign institutional investors from securities"
                        ElseIf SectionTDS = "96DA" Then
                        TDSSection_TDS2 = "196D(1A)-Income of specified fund from securities"
                        ElseIf SectionTDS = "94BA-P" Then
                        TDSSection_TDS2 = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
                        End If
                        
                        Sheet4.Cells(rowcount, SectionTDSColNo).value = TDSSection_TDS2
                        TDSSection_TDS2 = ""
                    End If

'-----------------------------------------------




        If Sheet4.Cells(rowcount, FYColNo).Locked = False Then
        
            Sheet4.Cells(rowcount, FYColNo).value = node("DeductedYr")
            

            Dim DeductedYr_1 As Variant
            DeductedYr_1 = UCase(node("DeductedYr"))
            
    'New Schema updated by Konda as AY-2026-27 on 29-12-2025

            If DeductedYr_1 = "2024" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2024-25"
            
            ElseIf DeductedYr_1 = "2023" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2023-24"

            ElseIf DeductedYr_1 = "2022" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2022-23"
            
            ElseIf DeductedYr_1 = "2021" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2021-22"
            
            ElseIf DeductedYr_1 = "2020" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2020-21"
            
            ElseIf DeductedYr_1 = "2019" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2019-20"
            
            ElseIf DeductedYr_1 = "2018" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2018-19"
            
            ElseIf DeductedYr_1 = "2017" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2017-18"
            
            ElseIf DeductedYr_1 = "2016" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2016-17"
            
            ElseIf DeductedYr_1 = "2015" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2015-16"
            
            ElseIf DeductedYr_1 = "2014" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2014-15"
            
            ElseIf DeductedYr_1 = "2013" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2013-14"
            
            ElseIf DeductedYr_1 = "2012" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2012-13"
            
            ElseIf DeductedYr_1 = "2011" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2011-12"
            
            ElseIf DeductedYr_1 = "2010" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2010-11"
            
            ElseIf DeductedYr_1 = "2009" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2009-10"
            
            ElseIf DeductedYr_1 = "2008" Then
            Sheet4.Cells(rowcount, FYColNo).value = "2008-09"
            
            
            End If
            
End If

        If Sheet4.Cells(rowcount, TDSOtherColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSOtherColNo).value = node("BroughtFwdTDSAmt")
        End If
        If Sheet4.Cells(rowcount, TDSDeductedColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSDeductedColNo).value = node("TDSDeducted")
        End If
        
'added by Chetan C M for AY 2025-26 on 03/02/2025
 
'start
'    Dim SecTDSDeducted
'    If Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
'    Dim SectionTDS As Variant
'    SectionTDS = node("SecTDSDeducted")
'                If SectionTDS = "192A" Then
'                SectionTDS = "192A-TDS on PF withdrawal"
'                ElseIf SectionTDS = "193" Then
'                SectionTDS = "193-Interest on Securities"
'                ElseIf SectionTDS = "194" Then
'                SectionTDS = "194-Dividends"
'                ElseIf SectionTDS = "94A" Then
'                SectionTDS = "194A-Interest other than 'Interest on securities'"
'                ElseIf SectionTDS = "94C" Then
'                SectionTDS = "194C-Payments to contractors and sub-contractors"
'                ElseIf SectionTDS = "4DA" Then
'                SectionTDS = "194DA-Payment in respect of life insurance policy"
'                ElseIf SectionTDS = "4EE" Then
'                SectionTDS = "194EE-Payments in respect of deposits under National Savings "
'                ElseIf SectionTDS = "4H" Then
'                SectionTDS = "194H-Commission or brokerage"
'                ElseIf SectionTDS = "4-IA" Then
'                SectionTDS = "194I(a)-Rent on hiring of plant and machinery"
'                ElseIf SectionTDS = "4-IB" Then
'                SectionTDS = "194I(b)-Rent on other than plant and machinery"
'                ElseIf SectionTDS = "4IB" Then
'                SectionTDS = "194IB-Payment of rent by certain individuals or Hindu undivided "
'                ElseIf SectionTDS = "94J-A" Then
'                SectionTDS = "194J(a)-Fees for technical services"
'                ElseIf SectionTDS = "94J-B" Then
'                SectionTDS = "194J(b)-Fees for professional  services or royalty etc"
'                ElseIf SectionTDS = "94K" Then
'                SectionTDS = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
'                ElseIf SectionTDS = "4LA" Then
'                SectionTDS = "194LA-Payment of compensation on acquisition of certain immovable "
'                ElseIf SectionTDS = "4BA1" Then
'                SectionTDS = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'                ElseIf SectionTDS = "4BA2" Then
'                SectionTDS = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'                ElseIf SectionTDS = "LBB" Then
'                SectionTDS = "194LBB-Income in respect of units of investment fund"
'                ElseIf SectionTDS = "94R" Then
'                SectionTDS = "194R-Benefits or perquisites of business or profession"
'                ElseIf SectionTDS = "94R-P" Then
'                SectionTDS = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
'                ElseIf SectionTDS = "94M" Then
'                SectionTDS = "194M-Payment of certain sums by certain individuals or HUF"
'                ElseIf SectionTDS = "94N" Then
'                SectionTDS = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
'                ElseIf SectionTDS = "94N-F" Then
'                SectionTDS = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
'                ElseIf SectionTDS = "94N-C" Then
'                SectionTDS = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
'                ElseIf SectionTDS = "94O" Then
'                SectionTDS = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
'                ElseIf SectionTDS = "94P" Then
'                SectionTDS = "194P-Deduction of tax in case of specified senior citizen"
'                ElseIf SectionTDS = "94D" Then
'                SectionTDS = "194D-Insurance commission"
'                ElseIf SectionTDS = "4G" Then
'                SectionTDS = "194G-Commission, price, etc. on sale of lottery tickets"
'                ElseIf SectionTDS = "LBC" Then
'                SectionTDS = "194LBC-Income in respect of investment in securitization trust"
'                ElseIf SectionTDS = "94Q" Then
'                SectionTDS = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
'                Else
'                SectionTDS = "(Select)"
'                End If
'                Sheet4.Cells(rowcount, SectionTDSColNo).value = SectionTDS
'    End If
    
 'end
        
        If Sheet4.Cells(rowcount, TDSClaimedColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSClaimedColNo).value = node("TDSClaimed")
        End If
        If Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
            Dim HeadOfIncome As Variant
            HeadOfIncome = node("HeadOfIncome")
            
            If HeadOfIncome = "BP" Then
            HeadOfIncome = "Income from Business & Profession"
            ElseIf HeadOfIncome = "HP" Then
            HeadOfIncome = "Income from House Property"
            ElseIf HeadOfIncome = "OS" Then
            HeadOfIncome = "Income from Other Source"
            ElseIf HeadOfIncome = "EI" Then
            HeadOfIncome = "Exempt Income"
            ElseIf HeadOfIncome = "NA" Then
            HeadOfIncome = "Not applicable (only in case TDS is deducted u/s 194N)"
            End If
            
            
            
            Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
        End If
        If Sheet4.Cells(rowcount, TDSGrossAmountColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSGrossAmountColNo).value = node("GrossAmount")
        End If
            
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function

Sub ImportScheduleTDS3Dtls(jsonText As String)
On Error Resume Next
    Dim jsonObject, jsonDictionary As Object
    Dim TotalXMLRow1 As Long
    Dim TotalDiffRow1 As Long
    Dim RecTDS11, RecTDS21, cnt As Long
    Dim Nodelist
    Dim node
    Dim InnerNode
    Dim TotalExRow1 As Long
    Dim TotalXMLRow, TotalExRow, TotalDiffRow, rowcount As Long
    Dim PANColNo, AadhaarColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, SectionTDSColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant

    Set jsonObject = ParseJson(jsonText)
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleTDS3Dtls")("TDS3Details")
  
    
    'TDSCreditColNo = Sheet4.Range("TDS2.TdsCredit").Column
    PANColNo = Sheet4.Range("TDsOthr2.PAN").Column
    AadhaarColNo = Sheet4.Range("TDsOthr2.AADHAAR").Column
    FYColNo = Sheet4.Range("TDsOthr2.finyear").Column
    TDSOtherColNo = Sheet4.Range("TDsOthr2.Tds").Column
    TDSDeductedColNo = Sheet4.Range("TDsOthr2.tdsdeducted").Column
  '  SectionTDSColNo = Sheet4.Range("TDsOthr.SectionTDS_ii").Column 'added by Chetan C M for AY 2025-26 on 03/02/2025
    'TDsOthr2.SectionTDSDeducted
    SectionTDSColNo = Sheet4.Range("TDsOthr2.SectionTDSDeducted").Column    'Konda updated on 21-04-2024
    TDSClaimedColNo = Sheet4.Range("TDsOthr2.tdsclaimed").Column
    TDSGrossAmountColNo = Sheet4.Range("TDsOthr2.grossamount").Column
    TDSHeadColNo = Sheet4.Range("TDsOthr2.headamount").Column
    
    TotalExRow = Range("TDsOthr2.PAN").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        If Sheet4.Range("TDsOthr2.PAN").Locked = False Then
            Sheet4.Range("TDsOthr2.PAN").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.AADHAAR").Locked = False Then
            Sheet4.Range("TDsOthr2.AADHAAR").ClearContents
        End If
  'Konda updated on 21-04-2024
        If Sheet4.Range("TDsOthr2.SectionTDSDeducted").Locked = False Then
            Sheet4.Range("TDsOthr2.SectionTDSDeducted").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.finyear").Locked = False Then
            Sheet4.Range("TDsOthr2.finyear").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.Tds").Locked = False Then
            Sheet4.Range("TDsOthr2.Tds").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.tdsdeducted").Locked = False Then
            Sheet4.Range("TDsOthr2.tdsdeducted").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.tdsclaimed").Locked = False Then
            Sheet4.Range("TDsOthr2.tdsclaimed").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.grossamount").Locked = False Then
            Sheet4.Range("TDsOthr2.grossamount").ClearContents
        End If
        If Sheet4.Range("TDsOthr2.headamount").Locked = False Then
            Sheet4.Range("TDsOthr2.headamount").ClearContents
        End If
    End If
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth1 (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet4.Range("TDsOthr2.PAN").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet4.Cells(rowcount, PANColNo).Locked = False Then
            Sheet4.Cells(rowcount, PANColNo).value = UCase(node("PANofTenant"))
        End If
        If Sheet4.Cells(rowcount, AadhaarColNo).Locked = False Then
            Sheet4.Cells(rowcount, AadhaarColNo).value = UCase(node("AadhaarofTenant"))
        End If

'Konda updated on 21-04-2025
        If Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
                    Dim SectionTDS
                        Dim TDSSection_TDS3 As Variant
                        SectionTDS = node("TDSSection")
                        If SectionTDS = "92A" Then
                        TDSSection_TDS3 = "192-Salary-Payment to Government employees other than Indian Government employees"
                        ElseIf SectionTDS = "92B" Then
                        TDSSection_TDS3 = "192-Salary-Payment to employees other than Government employees"
                        ElseIf SectionTDS = "92C" Then
                        TDSSection_TDS3 = "192-Salary-Payment to Indian Government employees"
                        ElseIf SectionTDS = "192A" Then
                        TDSSection_TDS3 = "192A-TDS on PF withdrawal"
                        ElseIf SectionTDS = "193" Then
                        TDSSection_TDS3 = "193-Interest on Securities"
                        ElseIf SectionTDS = "194" Then
                        TDSSection_TDS3 = "194-Dividends"
                        ElseIf SectionTDS = "94A" Then
                        TDSSection_TDS3 = "194A-Interest other than 'Interest on securities'"
                        ElseIf SectionTDS = "94B" Then
                        TDSSection_TDS3 = "194B-Winning from lottery or crossword puzzle"
                        ElseIf SectionTDS = "94BA" Then
                        TDSSection_TDS3 = "194BA-Winnings from online games"
                        ElseIf SectionTDS = "4BB" Then
                        TDSSection_TDS3 = "194BB-Winning from horse race"
                        ElseIf SectionTDS = "94C" Then
                        TDSSection_TDS3 = "194C-Payments to contractors and sub-contractors"
                        ElseIf SectionTDS = "94D" Then
                        TDSSection_TDS3 = "194D-Insurance commission"
                        ElseIf SectionTDS = "4DA" Then
                        TDSSection_TDS3 = "194DA-Payment in respect of life insurance policy"
                        ElseIf SectionTDS = "94E" Then
                        TDSSection_TDS3 = "194E-Payments to non-resident sportsmen or sports associations"
                        ElseIf SectionTDS = "4EE" Then
                        TDSSection_TDS3 = "194EE-Payments in respect of deposits under National Savings"
                        ElseIf SectionTDS = "4F" Then
                        TDSSection_TDS3 = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
                        ElseIf SectionTDS = "4G" Then
                        TDSSection_TDS3 = "194G-Commission, price, etc. on sale of lottery tickets"
                        ElseIf SectionTDS = "4H" Then
                        TDSSection_TDS3 = "194H-Commission or brokerage"
                        ElseIf SectionTDS = "4-IA" Then
                        TDSSection_TDS3 = "194I(a)-Rent on hiring of plant and machinery"
                        ElseIf SectionTDS = "4-IB" Then
                        TDSSection_TDS3 = "194I(b)-Rent on other than plant and machinery"
                        ElseIf SectionTDS = "4IA" Then
                        TDSSection_TDS3 = "194IA-TDS on Sale of immovable property"
                        ElseIf SectionTDS = "4IB" Then
                        TDSSection_TDS3 = "194IB-Payment of rent by certain individuals or Hindu undivided"
                        ElseIf SectionTDS = "4IC" Then
                        TDSSection_TDS3 = "194IC-Payment under specified agreement"
                        ElseIf SectionTDS = "94J-A" Then
                        TDSSection_TDS3 = "194J(a)-Fees for technical services"
                        ElseIf SectionTDS = "94J-B" Then
                        TDSSection_TDS3 = "194J(b)-Fees for professional  services or royalty etc"
                        ElseIf SectionTDS = "94K" Then
                        TDSSection_TDS3 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
                        ElseIf SectionTDS = "4LA" Then
                        TDSSection_TDS3 = "194LA-Payment of compensation on acquisition of certain immovable"
                        ElseIf SectionTDS = "4LB" Then
                        TDSSection_TDS3 = "194LB-Income by way of Interest from Infrastructure Debt fund"
                        ElseIf SectionTDS = "4LC1" Then
                        TDSSection_TDS3 = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
                        ElseIf SectionTDS = "4LC2" Then
                        TDSSection_TDS3 = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
                        ElseIf SectionTDS = "4LC3" Then
                        TDSSection_TDS3 = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
                        ElseIf SectionTDS = "4BA1" Then
                        TDSSection_TDS3 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
                        ElseIf SectionTDS = "4BA2" Then
                        TDSSection_TDS3 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
                        ElseIf SectionTDS = "LBA1" Then
                        TDSSection_TDS3 = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
                        ElseIf SectionTDS = "LBA2" Then
                        TDSSection_TDS3 = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
                        ElseIf SectionTDS = "LBA3" Then
                        TDSSection_TDS3 = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
                        ElseIf SectionTDS = "LBB" Then
                        TDSSection_TDS3 = "194LBB-Income in respect of units of investment fund"
                        ElseIf SectionTDS = "94R" Then
                        TDSSection_TDS3 = "194R-Benefits or perquisites of business or profession"
                        ElseIf SectionTDS = "94S" Then
                        TDSSection_TDS3 = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
                        ElseIf SectionTDS = "94B-P" Then
                        TDSSection_TDS3 = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
                        ElseIf SectionTDS = "94R-P" Then
                        TDSSection_TDS3 = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
                        ElseIf SectionTDS = "94S-P" Then
                        TDSSection_TDS3 = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
                        ElseIf SectionTDS = "LBC" Then
                        TDSSection_TDS3 = "194LBC-Income in respect of investment in securitization trust"
                        ElseIf SectionTDS = "4LD" Then
                        TDSSection_TDS3 = "194LD-TDS on interest on bonds / government securities"
                        ElseIf SectionTDS = "94M" Then
                        TDSSection_TDS3 = "194M-Payment of certain sums by certain individuals or HUF"
                        ElseIf SectionTDS = "94N" Then
                        TDSSection_TDS3 = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
                        ElseIf SectionTDS = "94N-F" Then
                        TDSSection_TDS3 = "194N-First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
                        ElseIf SectionTDS = "94N-C" Then
                        TDSSection_TDS3 = "194N-Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
                        ElseIf SectionTDS = "94N-FT" Then
                        TDSSection_TDS3 = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
                        ElseIf SectionTDS = "94O" Then
                        TDSSection_TDS3 = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
                        ElseIf SectionTDS = "94P" Then
                        TDSSection_TDS3 = "194P-Deduction of tax in case of specified senior citizen"
                        ElseIf SectionTDS = "94Q" Then
                        TDSSection_TDS3 = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
                        ElseIf SectionTDS = "195" Then
                        TDSSection_TDS3 = "195-Other sums payable to a non-resident"
                        ElseIf SectionTDS = "96A" Then
                        TDSSection_TDS3 = "196A-Income in respect of units of non-residents"
                        ElseIf SectionTDS = "96B" Then
                        TDSSection_TDS3 = "196B-Payments in respect of units to an offshore fund"
                        ElseIf SectionTDS = "96C" Then
                        TDSSection_TDS3 = "196C-Income from foreign currency bonds or shares of Indian"
                        ElseIf SectionTDS = "96D" Then
                        TDSSection_TDS3 = "196D-Income of foreign institutional investors from securities"
                        ElseIf SectionTDS = "96DA" Then
                        TDSSection_TDS3 = "196D(1A)-Income of specified fund from securities"
                        ElseIf SectionTDS = "94BA-P" Then
                        TDSSection_TDS3 = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
                        End If
                        
                        Sheet4.Cells(rowcount, SectionTDSColNo).value = TDSSection_TDS3
                        TDSSection_TDS3 = ""
                    End If
        '--------------------------------------------------------------





        If Sheet4.Cells(rowcount, FYColNo).Locked = False Then


                Dim DeductedYr_1 As Variant
                DeductedYr_1 = UCase(node("DeductedYr"))
                
      'New Schema updated by Konda as AY-2026-27 on 29-12-2025
                If DeductedYr_1 = "2024" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2024-25"
                
                ElseIf DeductedYr_1 = "2023" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2023-24"
                
                ElseIf DeductedYr_1 = "2022" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2022-23"
             
                ElseIf DeductedYr_1 = "2021" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2021-22"
                
                ElseIf DeductedYr_1 = "2020" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2020-21"
                
                ElseIf DeductedYr_1 = "2019" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2019-20"
            
                ElseIf DeductedYr_1 = "2018" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2018-19"
            
                ElseIf DeductedYr_1 = "2017" Then
                Sheet4.Cells(rowcount, FYColNo).value = "2017-18"
            
            End If
            
            
        End If

        If Sheet4.Cells(rowcount, TDSOtherColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSOtherColNo).value = UCase(node("BroughtFwdTDSAmt"))
        End If
        If Sheet4.Cells(rowcount, TDSDeductedColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSDeductedColNo).value = UCase(node("TDSDeducted"))
        End If
         
'added by Chetan C M for AY 2025-26 on 03/02/2025
        
'start
'    Dim SecTDSDeducted
'    If Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
'    Dim SectionTDS As Variant
'    SectionTDS = node("SecTDSDeducted")
'                If SectionTDS = "192A" Then
'                SectionTDS = "192A-TDS on PF withdrawal"
'                ElseIf SectionTDS = "193" Then
'                SectionTDS = "193-Interest on Securities"
'                ElseIf SectionTDS = "194" Then
'                SectionTDS = "194-Dividends"
'                ElseIf SectionTDS = "94A" Then
'                SectionTDS = "194A-Interest other than 'Interest on securities'"
'                ElseIf SectionTDS = "94C" Then
'                SectionTDS = "194C-Payments to contractors and sub-contractors"
'                ElseIf SectionTDS = "4DA" Then
'                SectionTDS = "194DA-Payment in respect of life insurance policy"
'                ElseIf SectionTDS = "4EE" Then
'                SectionTDS = "194EE-Payments in respect of deposits under National Savings "
'                ElseIf SectionTDS = "4H" Then
'                SectionTDS = "194H-Commission or brokerage"
'                ElseIf SectionTDS = "4-IA" Then
'                SectionTDS = "194I(a)-Rent on hiring of plant and machinery"
'                ElseIf SectionTDS = "4-IB" Then
'                SectionTDS = "194I(b)-Rent on other than plant and machinery"
'                ElseIf SectionTDS = "4IB" Then
'                SectionTDS = "194IB-Payment of rent by certain individuals or Hindu undivided "
'                ElseIf SectionTDS = "94J-A" Then
'                SectionTDS = "194J(a)-Fees for technical services"
'                ElseIf SectionTDS = "94J-B" Then
'                SectionTDS = "194J(b)-Fees for professional  services or royalty etc"
'                ElseIf SectionTDS = "94K" Then
'                SectionTDS = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
'                ElseIf SectionTDS = "4LA" Then
'                SectionTDS = "194LA-Payment of compensation on acquisition of certain immovable "
'                ElseIf SectionTDS = "4BA1" Then
'                SectionTDS = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'                ElseIf SectionTDS = "4BA2" Then
'                SectionTDS = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'                ElseIf SectionTDS = "LBB" Then
'                SectionTDS = "194LBB-Income in respect of units of investment fund"
'                ElseIf SectionTDS = "94R" Then
'                SectionTDS = "194R-Benefits or perquisites of business or profession"
'                ElseIf SectionTDS = "94R-P" Then
'                SectionTDS = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
'                ElseIf SectionTDS = "94M" Then
'                SectionTDS = "194M-Payment of certain sums by certain individuals or HUF"
'                ElseIf SectionTDS = "94N" Then
'                SectionTDS = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
'                ElseIf SectionTDS = "94N-F" Then
'                SectionTDS = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
'                ElseIf SectionTDS = "94N-C" Then
'                SectionTDS = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
'                ElseIf SectionTDS = "94O" Then
'                SectionTDS = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
'                ElseIf SectionTDS = "94P" Then
'                SectionTDS = "194P-Deduction of tax in case of specified senior citizen"
'                ElseIf SectionTDS = "94D" Then
'                SectionTDS = "194D-Insurance commission"
'                ElseIf SectionTDS = "4G" Then
'                SectionTDS = "194G-Commission, price, etc. on sale of lottery tickets"
'                ElseIf SectionTDS = "LBC" Then
'                SectionTDS = "194LBC-Income in respect of investment in securitization trust"
'                ElseIf SectionTDS = "94Q" Then
'                SectionTDS = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
'                Else
'                SectionTDS = "(Select)"
'                End If
'                Sheet4.Cells(rowcount, SectionTDSColNo).value = SectionTDS
'    End If
     
'end
        
        If Sheet4.Cells(rowcount, TDSClaimedColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSClaimedColNo).value = UCase(node("TDSClaimed"))
        End If
        If Sheet4.Cells(rowcount, TDSGrossAmountColNo).Locked = False Then
            Sheet4.Cells(rowcount, TDSGrossAmountColNo).value = UCase(node("GrossAmount"))
        End If
        If Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
            Dim HeadOfIncome As Variant
              HeadOfIncome = UCase(node("HeadOfIncome"))
           If HeadOfIncome = "BP" Then
            HeadOfIncome = "Income from Business & Profession"
            ElseIf HeadOfIncome = "HP" Then
            HeadOfIncome = "Income from House property"
            ElseIf HeadOfIncome = "OS" Then
            HeadOfIncome = "Income from other sources"
            ElseIf HeadOfIncome = "EI" Then
            HeadOfIncome = "Exempt Income"
            ElseIf HeadOfIncome = "NA" Then
            HeadOfIncome = "Not applicable (only in case TDS is deducted u/s 194N)"
            End If
            Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
        End If
        cnt = cnt + 1
    Next node
    RecTDS11 = cnt
End Sub


'import jso for ScheduleBP/TurnoverGrsRcptForGSTIN

Function ImportGst(jsonText As String)

    On Error Resume Next
    
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1 As Variant
    Dim GSTRnoColNo, GSTRAmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount As Variant
        
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleBP")("TurnoverGrsRcptForGSTIN")
    
    GSTRnoColNo = Sheet3.Range("BP.GSTRno").Column
    GSTRAmountColNo = Sheet3.Range("BP.GSTRAmount").Column
    
    TotalExRow = Range("BP.GSTRno").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet3.Range("BP.GSTRno").Locked = False Then
            Sheet3.Range("BP.GSTRno").ClearContents
        End If
        If Sheet3.Range("BP.GSTRAmount").Locked = False Then
            Sheet3.Range("BP.GSTRAmount").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Gst (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet3.Range("BP.GSTRno").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, GSTRnoColNo).Locked = False Then
            Sheet3.Cells(rowcount, GSTRnoColNo).value = node("GSTINNo")
        End If
        If Sheet3.Cells(rowcount, GSTRAmountColNo).Locked = False Then
            Sheet3.Cells(rowcount, GSTRAmountColNo).value = node("AmtTurnGrossRcptGSTIN")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

' import json for /ScheduleBP
Function ImportScheduleBP(jsonText As String)
On Error Resume Next

    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
    
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1   As Variant
        
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleBP")

   ' For Each Node In NodeList
   
    'New change for the AY 2024-25 Schema----------------------
 
         If Sheet3.Range("BP_E1_GR").Locked = False Then
            Sheet3.Range("BP_E1_GR").value = Nodelist("PersumptiveInc44AD")("GrsTotalTrnOver")
        End If
    '----------------------------------------------------------
   
        If Sheet3.Range("BP_E1a").Locked = False Then
            Sheet3.Range("BP_E1a").value = Nodelist("PersumptiveInc44AD")("GrsTrnOverBank")
        End If
        
    'New change for the AY 2024-25 Schema----------------------
         If Sheet3.Range("BP_E1a_E1b").Locked = False Then
            Sheet3.Range("BP_E1a_E1b").value = Nodelist("PersumptiveInc44AD")("GrsTotalTrnOverInCash")
        End If
    '--------------------------------------------------------------------
        
        If Sheet3.Range("BP_E1b").Locked = False Then
            Sheet3.Range("BP_E1b").value = Nodelist("PersumptiveInc44AD")("GrsTrnOverAnyOthMode")
        End If
        If Sheet3.Range("BP_E2a").Locked = False Then
            Sheet3.Range("BP_E2a").value = Nodelist("PersumptiveInc44AD")("PersumptiveInc44AD6Per")
        End If
        If Sheet3.Range("BP_E2b").Locked = False Then
            Sheet3.Range("BP_E2b").value = Nodelist("PersumptiveInc44AD")("PersumptiveInc44AD8Per")
        End If
         If Sheet3.Range("BP_E3_2").Locked = False Then
            Sheet3.Range("BP_E3_2").value = jsonObject("ITR")("ITR4")("ScheduleBP")("PersumptiveInc44AE")("SalInterestByFirm")
        End If
        If Sheet3.Range("BP_E3_GR").Locked = False Then
            Sheet3.Range("BP_E3_GR").value = Nodelist("PersumptiveInc44ADA")("GrsReceipt")
        End If
        
    'New change for the AY 2024-25 Schema----------------------------------------------------------------
        
        If Sheet3.Range("BP_E3_a").Locked = False Then
            Sheet3.Range("BP_E3_a").value = Nodelist("PersumptiveInc44ADA")("GrsTrnOverBank44ADA")
        End If
        
        If Sheet3.Range("BP_E3_b").Locked = False Then
            Sheet3.Range("BP_E3_b").value = Nodelist("PersumptiveInc44ADA")("GrsTotalTrnOverInCash44ADA")
        End If
        
        If Sheet3.Range("BP_E3_c").Locked = False Then
            Sheet3.Range("BP_E3_c").value = Nodelist("PersumptiveInc44ADA")("GrsTrnOverAnyOthMode44ADA")
        End If
        
    '------------------------------------------------------------------------------------------------------

        
        If Sheet3.Range("BP_E4_PI44ADA").Locked = False Then
            Sheet3.Range("BP_E4_PI44ADA").value = Nodelist("PersumptiveInc44ADA")("TotPersumptiveInc44ADA")
        End If
        If Sheet3.Range("BP_E9").Locked = False Then
            Sheet3.Range("BP_E9").value = Nodelist("FinanclPartclrOfBusiness")("PartnerMemberOwnCapital")
        End If
        If Sheet3.Range("BP_E10").Locked = False Then
            Sheet3.Range("BP_E10").value = Nodelist("FinanclPartclrOfBusiness")("SecuredLoans")
        End If
        If Sheet3.Range("BP_E11").Locked = False Then
            Sheet3.Range("BP_E11").value = Nodelist("FinanclPartclrOfBusiness")("UnSecuredLoans")
        End If
        If Sheet3.Range("BP_E12").Locked = False Then
            Sheet3.Range("BP_E12").value = Nodelist("FinanclPartclrOfBusiness")("Advances")
        End If
        If Sheet3.Range("BP_E13").Locked = False Then
            Sheet3.Range("BP_E13").value = Nodelist("FinanclPartclrOfBusiness")("SundryCreditors")
        End If
        If Sheet3.Range("BP_E14").Locked = False Then
            Sheet3.Range("BP_E14").value = Nodelist("FinanclPartclrOfBusiness")("OthrCurrLiab")
        End If
        If Sheet3.Range("BP_E16").Locked = False Then
            Sheet3.Range("BP_E16").value = Nodelist("FinanclPartclrOfBusiness")("FixedAssets")
        End If
        'Konda updated on 28-01-2026
        If Sheet3.Range("BP_E18a").Locked = False Then
            Sheet3.Range("BP_E18a").value = Nodelist("FinanclPartclrOfBusiness")("Investments")
        End If
        '========
        If Sheet3.Range("BP_E17").Locked = False Then
            Sheet3.Range("BP_E17").value = Nodelist("FinanclPartclrOfBusiness")("Inventories")
        End If
        If Sheet3.Range("BP_E18").Locked = False Then
            Sheet3.Range("BP_E18").value = Nodelist("FinanclPartclrOfBusiness")("SundryDebtors")
        End If
        If Sheet3.Range("BP_E19").Locked = False Then
            Sheet3.Range("BP_E19").value = Nodelist("FinanclPartclrOfBusiness")("BalWithBanks")
        End If
        If Sheet3.Range("BP_E20").Locked = False Then
            Sheet3.Range("BP_E20").value = Nodelist("FinanclPartclrOfBusiness")("CashInHand")
        End If
        If Sheet3.Range("BP_E21").Locked = False Then
            Sheet3.Range("BP_E21").value = Nodelist("FinanclPartclrOfBusiness")("LoansAndAdvances")
        End If
        If Sheet3.Range("BP_E22").Locked = False Then
            Sheet3.Range("BP_E22").value = Nodelist("FinanclPartclrOfBusiness")("OtherAssets")
        End If

   
    'Next Node
   
End Function

' import json for ScheduleBP/NatOfBus44AD

Function ImportADschedule(jsonText As String)

   On Error Resume Next
   
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1   As Variant
    Dim Sec44ADName, Sec44ADCode, Sec44ADDesc As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount As Variant
        
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleBP")("NatOfBus44AD")
    
    
    Sec44ADName = Sheet3.Range("Sec44AD.NOBName").Column
    Sec44ADCode = Sheet3.Range("Sec44AD.NOBCode").Column
    Sec44ADDesc = Sheet3.Range("Sec44AD.NOBDescription").Column
    
    TotalExRow = Range("Sec44AD.NOBName").Rows.count
    
    TotalXMLRow = Nodelist.count
   
    If (TotalXMLRow > 0) Then
        If Sheet3.Range("Sec44AD.NOBName").Locked = False Then
            Sheet3.Range("Sec44AD.NOBName").ClearContents
        End If
        If Sheet3.Range("Sec44AD.NOBCode").Locked = False Then
            Sheet3.Range("Sec44AD.NOBCode").ClearContents
        End If
        If Sheet3.Range("Sec44AD.NOBDescription").Locked = False Then
            Sheet3.Range("Sec44AD.NOBDescription").ClearContents
        End If
    End If
    
'Import-09.12.2022.03.02
    Dim vRows
    If Nodelist.count > Sheet3.Range("Sec44AD.NOBName").Rows.count Then
            Sheet3.Activate
            mdCommon.DefinedgridNameRange = ("Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription")
            ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet3.Range("Sec44AD.NOBName").Rows.count)
    End If
'---End Change

    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, Sec44ADName).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADName).value = node("NameOfBusiness")
        End If
            Dim iNOB As Variant
            Dim sNOB As Variant
            
            iNOB = node("CodeAD")
            sNOB = Findtext(iNOB, "NOB44AD")
        If sNOB <> "09029" Then
        If Sheet3.Cells(rowcount, Sec44ADCode).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADCode).value = sNOB
        End If
        End If
        If Sheet3.Cells(rowcount, Sec44ADDesc).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADDesc).value = node("Description")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

' import json for ScheduleBP/NatOfBus44ADA

Function ImportADAschedule(jsonText As String)

   On Error Resume Next
   
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1   As Variant
    Dim Sec44ADAName, Sec44ADACode, Sec44ADADesc As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount As Variant
    
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleBP")("NatOfBus44ADA")
    
    Sec44ADAName = Sheet3.Range("Sec44AD.NOBName1").Column
    Sec44ADACode = Sheet3.Range("Sec44AD.NOBCode1").Column
    Sec44ADADesc = Sheet3.Range("Sec44AD.NOBDescription1").Column
    
    TotalExRow = Range("Sec44AD.NOBName1").Rows.count
    
    TotalXMLRow = Nodelist.count
   
    If (TotalXMLRow > 0) Then
        If Sheet3.Range("Sec44AD.NOBName1").Locked = False Then
            Sheet3.Range("Sec44AD.NOBName1").ClearContents
        End If
        If Sheet3.Range("Sec44AD.NOBCode1").Locked = False Then
            Sheet3.Range("Sec44AD.NOBCode1").ClearContents
        End If
        If Sheet3.Range("Sec44AD.NOBDescription1").Locked = False Then
            Sheet3.Range("Sec44AD.NOBDescription1").ClearContents
        End If
    End If

'Import-09.12.2022.103.03
    Dim vRows
    If Nodelist.count > Sheet3.Range("Sec44AD.NOBName1").Rows.count Then
            Sheet3.Activate
            mdCommon.DefinedgridNameRange = ("Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1")
            ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName1")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet3.Range("Sec44AD.NOBName1").Rows.count)
    End If
'---End Change

    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName1").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, Sec44ADAName).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADAName).value = node("NameOfBusiness")
        End If
            Dim iNOB As Variant
            Dim sNOB As Variant
            
            iNOB = node("CodeADA")
            sNOB = Findtext(iNOB, "NOB44ADA")
        If Sheet3.Cells(rowcount, Sec44ADACode).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADACode).value = sNOB
        End If
        If Sheet3.Cells(rowcount, Sec44ADADesc).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADADesc).value = node("Description")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

' import json for ScheduleBP/NatOfBus44AE

Function ImportAEschedule(jsonText As String)
On Error Resume Next
   
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1   As Variant
    Dim Sec44AEName, Sec44AECode, Sec44AEDesc As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount As Variant
        
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleBP")("NatOfBus44AE")
    
    Sec44AEName = Sheet3.Range("Sec44AD.NOBName2").Column
    Sec44AECode = Sheet3.Range("Sec44AD.NOBCode2").Column
    Sec44AEDesc = Sheet3.Range("Sec44AD.NOBDescription2").Column
    
    TotalExRow = Range("Sec44AD.NOBName2").Rows.count
    
    TotalXMLRow = Nodelist.count
   
    If (TotalXMLRow > 0) Then
        If Sheet3.Range("Sec44AD.NOBName2").Locked = False Then
            Sheet3.Range("Sec44AD.NOBName2").ClearContents
        End If
        If Sheet3.Range("Sec44AD.NOBCode2").Locked = False Then
            Sheet3.Range("Sec44AD.NOBCode2").ClearContents
        End If
        If Sheet3.Range("Sec44AD.NOBDescription2").Locked = False Then
            Sheet3.Range("Sec44AD.NOBDescription2").ClearContents
        End If
    End If

'Import-09.12.2023.103.04
    Dim vRows
    If Nodelist.count > Sheet3.Range("Sec44AD.NOBName2").Rows.count Then
            Sheet3.Activate
            mdCommon.DefinedgridNameRange = ("Sec44AD.NOBName2||Sec44AD.NOBCode2||Sec44AD.NOBDescription2")
            ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName2")
            vRows = mdCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet3.Range("Sec44AD.NOBName2").Rows.count)
    End If
'---End Change
    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName2").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, Sec44AEName).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AEName).value = node("NameOfBusiness")
        End If
            Dim iNOB As Variant
            Dim sNOB As Variant
            iNOB = node("CodeAE")
            sNOB = Findtext(iNOB, "NOB")
        If Sheet3.Cells(rowcount, Sec44AECode).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AECode).MergeArea.value = sNOB
        End If
        If Sheet3.Cells(rowcount, Sec44AEDesc).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AEDesc).value = node("Description")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

' import json for ScheduleBP/GoodsDtlsUs44AE

Function ImportGoodsTableschedule(jsonText As String)
On Error Resume Next

    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1   As Variant
    Dim SEC44AERegNo, Sec44AEOLE, Sec44AETonnageCpct, Sec44AENoOfMonths, Sec44AEPresumptiveIncome As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount As Variant
        
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleBP")("GoodsDtlsUs44AE")
    
    SEC44AERegNo = Sheet3.Range("Sec44AE.RegNo").Column
    Sec44AEOLE = Sheet3.Range("Sec44AE.OLH").Column
    Sec44AETonnageCpct = Sheet3.Range("Sec44AE.TonnageCpct").Column
    Sec44AENoOfMonths = Sheet3.Range("Sheet44AE.NoOfMonths").Column
    Sec44AEPresumptiveIncome = Sheet3.Range("Sheet44AE.PresumptiveIncome").Column
    
    TotalExRow = Range("Sec44AE.RegNo").Rows.count
    
    TotalXMLRow = Nodelist.count
   
    If (TotalXMLRow > 0) Then
        If Sheet3.Range("Sec44AE.RegNo").Locked = False Then
            Sheet3.Range("Sec44AE.RegNo").ClearContents
        End If
        If Sheet3.Range("Sec44AE.OLH").Locked = False Then
            Sheet3.Range("Sec44AE.OLH").ClearContents
        End If
        If Sheet3.Range("Sec44AE.TonnageCpct").Locked = False Then
            Sheet3.Range("Sec44AE.TonnageCpct").ClearContents
        End If
        If Sheet3.Range("Sec44AE.NoOfMonths").Locked = False Then
            Sheet3.Range("Sheet44AE.NoOfMonths").ClearContents
        End If
        If Sheet3.Range("Sec44AE.PresumptiveIncome").Locked = False Then
            Sheet3.Range("Sheet44AE.PresumptiveIncome").ClearContents
        End If
    End If

    
    rowcount = getRowNo(Sheet3.Range("Sec44AE.RegNo").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, SEC44AERegNo).Locked = False Then
            Sheet3.Cells(rowcount, SEC44AERegNo).value = node("RegNumberGoodsCarriage")
        End If
            Dim iAEOLE As Variant
            Dim sAEOLE As Variant
            iAEOLE = node("OwnedLeasedHiredFlag")
              
            If iAEOLE = "OWN" Then
            sAEOLE = "Owned"
            ElseIf iAEOLE = "LEASE" Then
            sAEOLE = "Leased"
            ElseIf iAEOLE = "HIRED" Then
            sAEOLE = "Hired"
            End If
        If Sheet3.Cells(rowcount, Sec44AEOLE).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AEOLE).value = sAEOLE
        End If
            
        
         If Sheet3.Cells(rowcount, Sec44AETonnageCpct).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AETonnageCpct).value = node("TonnageCapacity")
        End If
        If Sheet3.Cells(rowcount, Sec44AENoOfMonths).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AENoOfMonths).value = node("HoldingPeriod")
        End If
        If Sheet3.Cells(rowcount, Sec44AEPresumptiveIncome).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AEPresumptiveIncome).value = node("PresumptiveIncome")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function
Function ImportTaxReturnPreparer(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
Dim jsonDictionary As Object
Dim IdentificationNoOfTRP, NameOfTRP, ReImbFrmGov As Variant
      
Set jsonObject = ParseJson(jsonText)

IdentificationNoOfTRP = jsonObject("ITR")("ITR4")("TaxReturnPreparer")("IdentificationNoOfTRP")
NameOfTRP = jsonObject("ITR")("ITR4")("TaxReturnPreparer")("NameOfTRP")
ReImbFrmGov = jsonObject("ITR")("ITR4")("TaxReturnPreparer")("ReImbFrmGov")

If IdentificationNoOfTRP <> "" And Sheet5.Range("Ver.IdentificationNoOfTRP").Locked = False Then
    Sheet5.Range("Ver.IdentificationNoOfTRP").value = IdentificationNoOfTRP
End If

If NameOfTRP <> "" And Sheet5.Range("Ver.NameOfTRP").Locked = False Then
    Sheet5.Range("Ver.NameOfTRP").value = NameOfTRP
End If

If ReImbFrmGov <> "" And Sheet5.Range("Ver.ReImbFrmGov").Locked = False Then
    Sheet5.Range("Ver.ReImbFrmGov").value = ReImbFrmGov
End If
End Function

'Function ImportSchedule80D(jsonText As String)
'On Error Resume Next
'Dim jsonObject As Object
'Dim jsonDictionary As Object
'Dim temp80, temp802, SeniorCitizenFlag, SelfAndFamily, HealthInsPremSlfFam, PrevHlthChckUpSlfFam, SelfAndFamilySeniorCitizen As Variant
'Dim HlthInsPremSlfFamSrCtzn, PrevHlthChckUpSlfFamSrCtzn, MedicalExpSlfFamSrCtzn, parentsSeniorCitizenFlag, Parents As Variant
'Dim HlthInsPremParents, PrevHlthChckUpParents, ParentsSeniorCitizen, HlthInsPremParentsSrCtzn, PrevHlthChckUpParentsSrCtzn As Variant
'Dim MedicalExpParentsSrCtzn, EligibleAmountOfDedn  As Variant
'
'
'Set jsonObject = ParseJson(jsonText)
'
'
'
'HealthInsPremSlfFam = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("HealthInsPremSlfFam")
'PrevHlthChckUpSlfFam = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpSlfFam")
'HlthInsPremSlfFamSrCtzn = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("HlthInsPremSlfFamSrCtzn")
'PrevHlthChckUpSlfFamSrCtzn = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpSlfFamSrCtzn")
'MedicalExpSlfFamSrCtzn = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("MedicalExpSlfFamSrCtzn")
'
'HlthInsPremParents = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("HlthInsPremParents")
'PrevHlthChckUpParents = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpParents")
'HlthInsPremParentsSrCtzn = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("HlthInsPremParentsSrCtzn")
'PrevHlthChckUpParentsSrCtzn = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpParentsSrCtzn")
'MedicalExpParentsSrCtzn = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("MedicalExpParentsSrCtzn")
'
'
'            temp80 = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("SeniorCitizenFlag")
'
'            If temp80 = "Y" Then
'                If Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = False Then
'                    Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Yes"
'                End If
'                If Sheet16.Range("Health_InsuranceSC_80D").Locked = False Then
'                    Sheet16.Range("Health_InsuranceSC_80D").value = HlthInsPremSlfFamSrCtzn
'                End If
'                If Sheet16.Range("Preventive_Health_SC_80D").Locked = False Then
'                    Sheet16.Range("Preventive_Health_SC_80D").value = PrevHlthChckUpSlfFamSrCtzn
'                End If
'                If Sheet16.Range("Medical_Expenditure_SC_80D").Locked = False Then
'                    Sheet16.Range("Medical_Expenditure_SC_80D").value = MedicalExpSlfFamSrCtzn
'                End If
'            ElseIf temp80 = "N" Then
'                If Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = False Then
'                    Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "No"
'                End If
'                If Sheet16.Range("Health_Insurance_80D").Locked = False Then
'                    Sheet16.Range("Health_Insurance_80D").value = HealthInsPremSlfFam
'                End If
'                If Sheet16.Range("Preventive_Health_80D").Locked = False Then
'                    Sheet16.Range("Preventive_Health_80D").value = PrevHlthChckUpSlfFam
'                End If
'            ElseIf temp80 = "S" Then
'                If Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = False Then
'                    Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
'                End If
'            End If
'
'            temp802 = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("ParentsSeniorCitizenFlag")
'
'            If temp802 = "Y" Then
'                If Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = False Then
'                    Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Yes"
'                End If
'                If Sheet16.Range("Health_Insurance3_80D").Locked = False Then
'                    Sheet16.Range("Health_Insurance3_80D").value = HlthInsPremParentsSrCtzn
'                End If
'                If Sheet16.Range("Preventive_Health3_80D").Locked = False Then
'                    Sheet16.Range("Preventive_Health3_80D").value = PrevHlthChckUpParentsSrCtzn
'                End If
'                If Sheet16.Range("Medical_Expenditure2_80D").Locked = False Then
'                    Sheet16.Range("Medical_Expenditure2_80D").value = MedicalExpParentsSrCtzn
'                End If
'            ElseIf temp802 = "N" Then
'                If Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = False Then
'                    Sheet16.Range("DropDown_ValueOf_SC_80D").value = "No"
'                End If
'                If Sheet16.Range("Health_Insurance2_80D").Locked = False Then
'                    Sheet16.Range("Health_Insurance2_80D").value = HlthInsPremParents
'                End If
'                If Sheet16.Range("Preventive_Health2_80D").Locked = False Then
'                    Sheet16.Range("Preventive_Health2_80D").value = PrevHlthChckUpParents
'                End If
'            ElseIf temp802 = "P" Then
'                If Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = False Then
'                    Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
'                End If
'            End If
'
'End Function
Function ImportSchedule80D(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
'KONDA 25-04-2025-----------------
Dim NodeList80D
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80D = jsonObject("ITR")("ITR4")
 
If NodeList80D.exists("Schedule80D") Then
 '---------------------------------
Set node = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")
    
            If Trim(node("SeniorCitizenFlag")) <> "" Then
                If node("SeniorCitizenFlag") = "Y" Then
                    Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Yes"
                ElseIf node("SeniorCitizenFlag") = "N" Then
                    Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "No"
                ElseIf node("SeniorCitizenFlag") = "S" Then
                    Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
                End If
            End If
                                            
    'Konda--------19/04/2025
               'This field autopopulate AY_2025-26
'            If Trim(node("HealthInsPremSlfFam")) <> "" And Sheet16.Range("Health_Insurance_80D").Locked = False Then
'                Sheet16.Range("Health_Insurance_80D").Value = node("HealthInsPremSlfFam")
'            End If
    '_________________________________
             
            Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, cnt As Long
            Dim node_80DbSFHI, Nodelist_80DbSFHI
'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
            Set Nodelist_80DbSFHI = node("Sec80DSelfFamHIDtls")("Sch80DInsDtls")
                
            Dim InsurerName_80DbSFHI, PolicyNo_80DbSFHI, ReceiptNo_80DbSFHI, HealthInsAmt_80DbSFHI, RecInt80DbSFHI
                
                InsurerName_80DbSFHI = Sheet16.Range("NameInsurerA1.80D").Column
                PolicyNo_80DbSFHI = Sheet16.Range("PolicyNumA1.80D").Column
                'AY_25_26_V0.7
                'ReceiptNo_80DbSFHI = Sheet16.Range("ReceiptNumA1.80D").Column
                '--------------
                HealthInsAmt_80DbSFHI = Sheet16.Range("AmtA1.80D").Column
                
                 TotalExRow = Sheet16.Range("NameInsurerA1.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbSFHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet16.Range("NameInsurerA1.80D").Locked = False Then
                       Sheet16.Range("NameInsurerA1.80D").ClearContents
                    End If
                    
                    If Sheet16.Range("PolicyNumA1.80D").Locked = False Then
                       Sheet16.Range("PolicyNumA1.80D").ClearContents
                    End If
                    'AY_25_26_V0.7
'                    If Sheet16.Range("ReceiptNumA1.80D").Locked = False Then
'                       Sheet16.Range("ReceiptNumA1.80D").ClearContents
'                    End If
                    '--------------------
                    If Sheet16.Range("AmtA1.80D").Locked = False Then
                       Sheet16.Range("AmtA1.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DSelfFamHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet16.Range("NameInsurerA1.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbSFHI In Nodelist_80DbSFHI
                         rowcount = rowcount + 1
                         
                         If Sheet16.Cells(rowcount, InsurerName_80DbSFHI).Locked = False Then
                             Sheet16.Cells(rowcount, InsurerName_80DbSFHI).value = node_80DbSFHI("InsurerName")
                         End If
                         
                         If Sheet16.Cells(rowcount, PolicyNo_80DbSFHI).Locked = False Then
                             Sheet16.Cells(rowcount, PolicyNo_80DbSFHI).value = node_80DbSFHI("PolicyNo")
                         End If
                          'AY_25_26_V0.7
'                         If Sheet16.Cells(rowcount, ReceiptNo_80DbSFHI).Locked = False Then
'                             Sheet16.Cells(rowcount, ReceiptNo_80DbSFHI).value = node_80DbSFHI("ReceiptNo")
'                         End If
                         '------------
                         If Sheet16.Cells(rowcount, HealthInsAmt_80DbSFHI).Locked = False Then
                             Sheet16.Cells(rowcount, HealthInsAmt_80DbSFHI).value = node_80DbSFHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbSFHI
                     
                 RecInt80DbSFHI = cnt
                     
                
                
                '-------------------------
            
            '---------------------------------
            If Trim(node("PrevHlthChckUpSlfFam")) <> "" And Sheet16.Range("Preventive_Health_80D").locekd = False Then
                Sheet16.Range("Preventive_Health_80D").value = node("PrevHlthChckUpSlfFam")
            End If
            '-------------------
            'This field autopopulate AY_2025-26
'            If Trim(node("HlthInsPremSlfFamSrCtzn")) <> "" And Sheet16.Range("Health_InsuranceSC_80D").Locked = False Then
'                Sheet16.Range("Health_InsuranceSC_80D").Value = node("HlthInsPremSlfFamSrCtzn")
'            End If
            '__________________________________
           
                Dim node_80DbSFsrCHI, Nodelist_80DbSFsrCHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbSFsrCHI = node("Sec80DSelfFamSrCtznHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbSFsrCHI, PolicyNo_80DbSFsrCHI, ReceiptNo_80DbSFsrCHI, HealthInsAmt_80DbSFsrCHI, RecInt80DbSFsrCHI
                
                InsurerName_80DbSFsrCHI = Sheet16.Range("NameInsurerB1.80D").Column
                PolicyNo_80DbSFsrCHI = Sheet16.Range("PolicyNumB1.80D").Column
                  'AY_25_26_V0.7
                'ReceiptNo_80DbSFsrCHI = Sheet16.Range("ReceiptNumB1.80D").Column
                '------------------
                
                HealthInsAmt_80DbSFsrCHI = Sheet16.Range("AmtB1.80D").Column
                
                 TotalExRow = Sheet16.Range("NameInsurerB1.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbSFsrCHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet16.Range("NameInsurerB1.80D").Locked = False Then
                       Sheet16.Range("NameInsurerB1.80D").ClearContents
                    End If
                    
                    If Sheet16.Range("PolicyNumB1.80D").Locked = False Then
                       Sheet16.Range("PolicyNumB1.80D").ClearContents
                    End If
                      'AY_25_26_V0.7
'                    If Sheet16.Range("ReceiptNumB1.80D").Locked = False Then
'                       Sheet16.Range("ReceiptNumB1.80D").ClearContents
'                    End If
                    '----------------
                    If Sheet16.Range("AmtB1.80D").Locked = False Then
                       Sheet16.Range("AmtB1.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DSelfFamSrCtznHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet16.Range("NameInsurerB1.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbSFsrCHI In Nodelist_80DbSFsrCHI
                         rowcount = rowcount + 1
                         
                         If Sheet16.Cells(rowcount, InsurerName_80DbSFsrCHI).Locked = False Then
                             Sheet16.Cells(rowcount, InsurerName_80DbSFsrCHI).value = node_80DbSFsrCHI("InsurerName")
                         End If
                         
                         If Sheet16.Cells(rowcount, PolicyNo_80DbSFsrCHI).Locked = False Then
                             Sheet16.Cells(rowcount, PolicyNo_80DbSFsrCHI).value = node_80DbSFsrCHI("PolicyNo")
                         End If
                           'AY_25_26_V0.7
'                         If Sheet16.Cells(rowcount, ReceiptNo_80DbSFsrCHI).Locked = False Then
'                             Sheet16.Cells(rowcount, ReceiptNo_80DbSFsrCHI).value = node_80DbSFsrCHI("ReceiptNo")
'                         End If
                         '------------------
                         
                         If Sheet16.Cells(rowcount, HealthInsAmt_80DbSFsrCHI).Locked = False Then
                             Sheet16.Cells(rowcount, HealthInsAmt_80DbSFsrCHI).value = node_80DbSFsrCHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbSFsrCHI
                     
                 RecInt80DbSFsrCHI = cnt
            
            '----------------------------------
            If Trim(node("PrevHlthChckUpSlfFamSrCtzn")) <> "" And Sheet16.Range("Preventive_Health_SC_80D").Locked = False Then
                Sheet16.Range("Preventive_Health_SC_80D").value = node("PrevHlthChckUpSlfFamSrCtzn")
            End If
            If Trim(node("MedicalExpSlfFamSrCtzn")) <> "" And Sheet16.Range("Medical_Expenditure_SC_80D").Locked = False Then
                Sheet16.Range("Medical_Expenditure_SC_80D").value = node("MedicalExpSlfFamSrCtzn")
            End If
            
            If Trim(node("ParentsSeniorCitizenFlag")) <> "" Then
                If node("ParentsSeniorCitizenFlag") = "Y" Then
                    Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Yes"
                ElseIf node("ParentsSeniorCitizenFlag") = "N" Then
                    Sheet16.Range("DropDown_ValueOf_SC_80D").value = "No"
                ElseIf node("ParentsSeniorCitizenFlag") = "P" Then
                    Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
                End If
            End If
            '-----------------------------
            'This field autopopulate AY_2025-26
'            If Trim(node("HlthInsPremParents")) <> "" And Sheet16.Range("Health_Insurance2_80D").Locked = False Then
'                Sheet16.Range("Health_Insurance2_80D").Value = node("HlthInsPremParents")
'            End If
            '_________________________________
            
            Dim node_80DbPHI, Nodelist_80DbPHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbPHI = node("Sec80DParentsHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbPHI, PolicyNo_80DbPHI, ReceiptNo_80DbPHI, HealthInsAmt_80DbPHI, RecInt80DbPHI
                
                InsurerName_80DbPHI = Sheet16.Range("NameInsurerA2.80D").Column
                PolicyNo_80DbPHI = Sheet16.Range("PolicyNumA2.80D").Column
                ''AY_25_26_V0.7
                'ReceiptNo_80DbPHI = Sheet16.Range("ReceiptNumA2.80D").Column
                '--------------------------------
                HealthInsAmt_80DbPHI = Sheet16.Range("AmtA2.80D").Column
                
                 TotalExRow = Sheet16.Range("NameInsurerA2.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbPHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet16.Range("NameInsurerA2.80D").Locked = False Then
                       Sheet16.Range("NameInsurerA2.80D").ClearContents
                    End If
                    
                    If Sheet16.Range("PolicyNumA2.80D").Locked = False Then
                       Sheet16.Range("PolicyNumA2.80D").ClearContents
                    End If
                    'AY_25_26_V0.7
'                    If Sheet16.Range("ReceiptNumA2.80D").Locked = False Then
'                       Sheet16.Range("ReceiptNumA2.80D").ClearContents
'                    End If
                    '------------------
                    If Sheet16.Range("AmtA2.80D").Locked = False Then
                       Sheet16.Range("AmtA2.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DParentsHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet16.Range("NameInsurerA2.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbPHI In Nodelist_80DbPHI
                         rowcount = rowcount + 1
                         
                         If Sheet16.Cells(rowcount, InsurerName_80DbPHI).Locked = False Then
                             Sheet16.Cells(rowcount, InsurerName_80DbPHI).value = node_80DbPHI("InsurerName")
                         End If
                         
                         If Sheet16.Cells(rowcount, PolicyNo_80DbPHI).Locked = False Then
                             Sheet16.Cells(rowcount, PolicyNo_80DbPHI).value = node_80DbPHI("PolicyNo")
                         End If
                         'AY_25_26_V0.7
'                         If Sheet16.Cells(rowcount, ReceiptNo_80DbPHI).Locked = False Then
'                             Sheet16.Cells(rowcount, ReceiptNo_80DbPHI).value = node_80DbPHI("ReceiptNo")
'                         End If
                         '--------------------
                         If Sheet16.Cells(rowcount, HealthInsAmt_80DbPHI).Locked = False Then
                             Sheet16.Cells(rowcount, HealthInsAmt_80DbPHI).value = node_80DbPHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbPHI
                     
                 RecInt80DbPHI = cnt
            
            '----------------------------------
            '---------------------------------
            If Trim(node("PrevHlthChckUpParents")) <> "" And Sheet16.Range("Preventive_Health2_80D") = False Then
                Sheet16.Range("Preventive_Health2_80D").value = node("PrevHlthChckUpParents")
            End If
            '-------------------------------------
            'This field autopopulate AY_2025-26
'            If Trim(node("HlthInsPremParentsSrCtzn")) <> "" And Sheet16.Range("Health_Insurance3_80D").Locked = False Then
'                Sheet16.Range("Health_Insurance3_80D").Value = node("HlthInsPremParentsSrCtzn")
'            End If
            '_______________________________________
             Dim node_80DbPsrCHI, Nodelist_80DbPsrCHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbPsrCHI = node("Sec80DParentsSrCtznHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbPsrCHI, PolicyNo_80DbPsrCHI, ReceiptNo_80DbPsrCHI, HealthInsAmt_80DbPsrCHI, RecInt80DbPsrCHI
                
                InsurerName_80DbPsrCHI = Sheet16.Range("NameInsurerB2.80D").Column
                PolicyNo_80DbPsrCHI = Sheet16.Range("PolicyNumB2.80D").Column
                'AY_25_26_V0.7
                 'ReceiptNo_80DbPsrCHI = Sheet16.Range("ReceiptNumB2.80D").Column
                '-------------------
                HealthInsAmt_80DbPsrCHI = Sheet16.Range("AmtB2.80D").Column
                
                 TotalExRow = Sheet16.Range("NameInsurerB2.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbPsrCHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet16.Range("NameInsurerB2.80D").Locked = False Then
                       Sheet16.Range("NameInsurerB2.80D").ClearContents
                    End If
                    
                    If Sheet16.Range("PolicyNumB2.80D").Locked = False Then
                       Sheet16.Range("PolicyNumB2.80D").ClearContents
                    End If
                    'AY_25_26_V0.7

'                    If Sheet16.Range("ReceiptNumB2.80D").Locked = False Then
'                       Sheet16.Range("ReceiptNumB2.80D").ClearContents
'                    End If
                    '----------------
                    If Sheet16.Range("AmtB2.80D").Locked = False Then
                       Sheet16.Range("AmtB2.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DParentsSrCtznHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet16.Range("NameInsurerB2.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbPsrCHI In Nodelist_80DbPsrCHI
                         rowcount = rowcount + 1
                         
                         If Sheet16.Cells(rowcount, InsurerName_80DbPsrCHI).Locked = False Then
                             Sheet16.Cells(rowcount, InsurerName_80DbPsrCHI).value = node_80DbPsrCHI("InsurerName")
                         End If
                         
                         If Sheet16.Cells(rowcount, PolicyNo_80DbPsrCHI).Locked = False Then
                             Sheet16.Cells(rowcount, PolicyNo_80DbPsrCHI).value = node_80DbPsrCHI("PolicyNo")
                         End If
                         'AY_25_26_V0.7
'                         If Sheet16.Cells(rowcount, ReceiptNo_80DbPsrCHI).Locked = False Then
'                             Sheet16.Cells(rowcount, ReceiptNo_80DbPsrCHI).value = node_80DbPsrCHI("ReceiptNo")
'                         End If
                         '--------------------
                         If Sheet16.Cells(rowcount, HealthInsAmt_80DbPsrCHI).Locked = False Then
                             Sheet16.Cells(rowcount, HealthInsAmt_80DbPsrCHI).value = node_80DbPsrCHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbPsrCHI
                     
                 RecInt80DbPsrCHI = cnt
            '---------------------------------------
            If Trim(node("PrevHlthChckUpParentsSrCtzn")) <> "" And Sheet16.Range("Preventive_Health3_80D").Locked = False Then
                Sheet16.Range("Preventive_Health3_80D").value = node("PrevHlthChckUpParentsSrCtzn")
            End If
            If Trim(node("MedicalExpParentsSrCtzn")) <> "" And Sheet16.Range("Medical_Expenditure2_80D").Locked = False Then
                Sheet16.Range("Medical_Expenditure2_80D").value = node("MedicalExpParentsSrCtzn")
            End If
End If
End Function

Function ImportRefund(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim IFSCColNo, BankNameColNo, ACCNOColNo, CheckBox, TypeofAccount As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long

Set jsonObject = ParseJson(jsonText)

  
     Set Nodelist = jsonObject("ITR")("ITR4")("Refund")("BankAccountDtls")("AddtnlBankDetails")


    IFSCColNo = Sheet5.Range("SchBA.IFSC").Column
    BankNameColNo = Sheet5.Range("SchBA.BankName").Column
    ACCNOColNo = Sheet5.Range("SchBA.AcntNo").Column
    
    TypeofAccount = Sheet5.Range("SchBA.Type_of_Account").Column ' AY 2024-25 Schema implementation

'uncommented by Chetan C M AY2025-26 21/01/2025
'--start--
    Dim CheckBoxColNo
    CheckBox = Sheet5.Range("tempXML").Column
    CheckBoxColNo = Sheet5.Range("SchBA.CheckBox").Column
'--end--
    TotalExRow = Range("SchBA.IFSC").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    
    If (TotalXMLRow > 0) Then
        If Sheet5.Range("SchBA.IFSC").Locked = False Then
            Sheet5.Range("SchBA.IFSC").ClearContents
        End If
        If Sheet5.Range("SchBA.BankName").Locked = False Then
            Sheet5.Range("SchBA.BankName").ClearContents
        End If
        If Sheet5.Range("SchBA.AcntNo").Locked = False Then
            Sheet5.Range("SchBA.AcntNo").ClearContents
        End If
        
        'AY 2024-25 Change------------
        If Sheet5.Range("SchBA.Type_of_Account").Locked = False Then
            Sheet5.Range("SchBA.Type_of_Account").ClearContents
        End If
        'AY 2024-25 End ------------
        
        If Sheet5.Range("tempxml").Locked = False Then
            Sheet5.Range("tempxml").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     Dim idi As Long
    idi = 0
    For idi = 1 To TotalDiffRow
     AddDiffRows_BANK (1)
     Next
    End If
    
    rowcount = getRowNo(Sheet5.Range("SchBA.IFSC").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet5.Cells(rowcount, IFSCColNo).Locked = False Then
            Sheet5.Cells(rowcount, IFSCColNo).value = node("IFSCCode")
        End If
        If Sheet5.Cells(rowcount, BankNameColNo).Locked = False Then
            Sheet5.Cells(rowcount, BankNameColNo).value = node("BankName")
        End If
        If Sheet5.Cells(rowcount, ACCNOColNo).Locked = False Then
            Sheet5.Cells(rowcount, ACCNOColNo).value = Trim(node("BankAccountNo"))
        End If
        
'--------------------------------------------------------------------------------
'AY 2024-24 Schema Implementation Starts
    If Sheet5.Cells(rowcount, TypeofAccount).Locked = False Then
                                If UCase(node("AccountType")) = "SB" Then
                                        Sheet5.Cells(rowcount, TypeofAccount).value = "Savings Account"
                                    ElseIf UCase(node("AccountType")) = "CA" Then
                                        Sheet5.Cells(rowcount, TypeofAccount).value = "Current Account"
                                    ElseIf UCase(node("AccountType")) = "CC" Then
                                        Sheet5.Cells(rowcount, TypeofAccount).value = "Cash Credit Account"
                                    ElseIf UCase(node("AccountType")) = "OD" Then
                                        Sheet5.Cells(rowcount, TypeofAccount).value = "Over draft account"
                                    ElseIf UCase(node("AccountType")) = "NRO" Then
                                        Sheet5.Cells(rowcount, TypeofAccount).value = "Non Resident Account"
                                    ElseIf UCase(node("AccountType")) = "OTH" Then
                                        Sheet5.Cells(rowcount, TypeofAccount).value = "Other"
                                End If
    
    End If
'AY 2024-24 Schema Implementation End
'--------------------------------------------------------------------------------

'uncommented by Chetan C M AY2025-26 21/01/2025
'--start--
        If Sheet5.Cells(rowcount, CheckBox).Locked = False Then
            Sheet5.Cells(rowcount, CheckBox).value = node("UseForRefund")
        End If
            If CheckBox = True Then
            CheckBox = "true"
            ElseIf CheckBox = False Then
            CheckBox = "false"
            ElseIf CheckBox = "" Then
            CheckBox = "false"
            End If

        LinkCheckBoxes

            If tempXML = True Then
                Sheet5.Cells(rowcount, CheckBoxColNo).value = "True"
            ElseIf tempXML = False Then
                Sheet5.Cells(rowcount, CheckBoxColNo).value = "False"
            ElseIf tempXML = "" Then
                 Sheet5.Cells(rowcount, CheckBoxColNo).value = "False"
            End If
'--end--
           
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
   
End Function

Function ImportScheduleIT(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim strDate As String
Dim YYYY, MM, DD As String


Set jsonObject = ParseJson(jsonText)


   
    Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleIT")("TaxPayment")

    BSRCodeColNo = Sheet18.Range("TaxP.BSRCode").Column
    DateDepColNo = Sheet18.Range("TaxP.DateDep").Column
    SrlNoChallanColNo = Sheet18.Range("TaxP.SrlNoOfChaln").Column
    AmtColNo = Sheet18.Range("TaxP.Amt").Column
    
    TotalExRow = Range("TaxP.BSRCode").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet18.Range("TaxP.BSRCode").Locked = False Then
            Sheet18.Range("TaxP.BSRCode").ClearContents
        End If
        If Sheet18.Range("TaxP.DateDep").Locked = False Then
            Sheet18.Range("TaxP.DateDep").ClearContents
        End If
        If Sheet18.Range("TaxP.SrlNoOfChaln").Locked = False Then
            Sheet18.Range("TaxP.SrlNoOfChaln").ClearContents
        End If
        If Sheet18.Range("TaxP.Amt").Locked = False Then
            Sheet18.Range("TaxP.Amt").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IT (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet18.Range("TaxP.BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet18.Cells(rowcount, BSRCodeColNo).Locked = False Then
            Sheet18.Cells(rowcount, BSRCodeColNo).value = node("BSRCode")
        End If
            strDate = node("DateDep")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
            If strDate <> "" And Sheet18.Cells(rowcount, DateDepColNo).Locked = False Then
                Sheet18.Cells(rowcount, DateDepColNo).value = strDate
                strDate = ""
            End If
            If Sheet18.Cells(rowcount, SrlNoChallanColNo).Locked = False Then
                Sheet18.Cells(rowcount, SrlNoChallanColNo).value = node("SrlNoOfChaln")
            End If
            If Sheet18.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet18.Cells(rowcount, AmtColNo).value = node("Amt")
            End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function

Function ImportVerification(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
Dim jsonDictionary As Object
Dim FatherName, AssesseeVerName, AssesseeVerPAN, Capacity, Place As Variant


Set jsonObject = ParseJson(jsonText)
 
 FatherName = jsonObject("ITR")("ITR4")("Verification")("Declaration")("FatherName")
 AssesseeVerName = jsonObject("ITR")("ITR4")("Verification")("Declaration")("AssesseeVerName")
 AssesseeVerPAN = jsonObject("ITR")("ITR4")("Verification")("Declaration")("AssesseeVerPAN")
 
 Capacity = jsonObject("ITR")("ITR4")("Verification")("Capacity")
 
    If Capacity = "S" Then
        Capacity = "Self"
    ElseIf Capacity = "R" Then
        Capacity = "Representative"
    ElseIf Capacity = "K" Then
        Capacity = "Karta"
    ElseIf Capacity = "P" Then
        Capacity = "Partner"
    Else
        Capacity = ""
    End If
            
Place = jsonObject("ITR")("ITR4")("Verification")("Place")
        If Capacity <> "" And Sheet5.Range("Ver.capacity").Locked = False Then
            Sheet5.Range("Ver.capacity").value = Capacity
        End If
        If FatherName <> "" And Sheet5.Range("Ver.FatherName").Locked = False Then
            Sheet5.Range("Ver.FatherName").value = FatherName
        End If
        
        If AssesseeVerName <> "" And Sheet5.Range("Ver.AssesseeVerName").Locked = False Then
            Sheet5.Range("Ver.AssesseeVerName").value = AssesseeVerName
        End If
        
        If AssesseeVerPAN <> "" And Sheet5.Range("Ver.PAN").Locked = False Then
            Sheet5.Range("Ver.PAN").value = AssesseeVerPAN
        End If
            
            
        
            
        If Place <> "" And Sheet5.Range("Ver.Place").Locked = False Then
            Sheet5.Range("Ver.Place").value = Place
        End If
                                                
End Function

Function ImportScheduleTCS(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Dim TANColNo, DEDNameColNo, Amtfrom26ASColNo, TaxColNo, ClaimColNo As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt, RecTCS  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim strDate As String
Dim YYYY, MM, DD As String
    
 Set jsonObject = ParseJson(jsonText)
  
   Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleTCS")("TCS")
    
    TANColNo = Sheet17.Range("TCS.TAN").Column
    DEDNameColNo = Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    Amtfrom26ASColNo = Sheet17.Range("TCS.AmountPaid").Column
    TaxColNo = Sheet17.Range("TCS.TotalTCS").Column
    ClaimColNo = Sheet17.Range("TCS.AmtTCSClaimedThisYear").Column

    
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("TCS.TAN").Locked = False Then
            Sheet17.Range("TCS.TAN").ClearContents
        End If
        If Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Locked = False Then
            Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
        End If
        If Sheet17.Range("TCS.AmountPaid").Locked = False Then
            Sheet17.Range("TCS.AmountPaid").ClearContents
        End If
        If Sheet17.Range("TCS.TotalTCS").Locked = False Then
            Sheet17.Range("TCS.TotalTCS").ClearContents
        End If
        If Sheet17.Range("TCS.AmtTCSClaimedThisYear").Locked = False Then
            Sheet17.Range("TCS.AmtTCSClaimedThisYear").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TCS (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet17.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet17.Cells(rowcount, TANColNo).Locked = False Then
            Sheet17.Cells(rowcount, TANColNo).value = node("EmployerOrDeductorOrCollectDetl")("TAN")
        End If
        If Sheet17.Cells(rowcount, DEDNameColNo).Locked = False Then
            Sheet17.Cells(rowcount, DEDNameColNo).value = node("EmployerOrDeductorOrCollectDetl")("EmployerOrDeductorOrCollecterName")
        End If
        If Sheet17.Cells(rowcount, Amtfrom26ASColNo).Locked = False Then
            Sheet17.Cells(rowcount, Amtfrom26ASColNo).value = node("Amtfrom26AS")
        End If
        If Sheet17.Cells(rowcount, TaxColNo).Locked = False Then
            Sheet17.Cells(rowcount, TaxColNo).value = node("TotalTCS")
        End If
        If Sheet17.Cells(rowcount, ClaimColNo).Locked = False Then
            Sheet17.Cells(rowcount, ClaimColNo).value = node("AmtTCSClaimedThisYear")
        End If
        cnt = cnt + 1
    Next node
    RecTCS = cnt
End Function

Function ImportTaxExmpIntIncDtls(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Dim NatureColNo, AmountColNo, DescriptionColNo, DescEIColNo As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt, RecTCS, Necnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim text

 Set jsonObject = ParseJson(jsonText)
    If jsonObject("ITR")("ITR4").exists("TaxExmpIntIncDtls") Then
    If jsonObject("ITR")("ITR4")("TaxExmpIntIncDtls").exists("OthersInc") Then
    If jsonObject("ITR")("ITR4")("TaxExmpIntIncDtls")("OthersInc").exists("OthersIncDtls") Then

   Set Nodelist = jsonObject("ITR")("ITR4")("TaxExmpIntIncDtls")("OthersInc")("OthersIncDtls")
    
    NatureColNo = Sheet5.Range("Sheet5.Nature").Column
    DescriptionColNo = Sheet5.Range("Sheet5.Description").Column
    AmountColNo = Sheet5.Range("Sheet5.Amount").Column
    'Malli_AY_2026_27 17/06/2026 newlly added Description field
    DescEIColNo = Sheet5.Range("Sheet5.DescEI").Column
    '---------------------------
    
    TotalExRow = Range("Sheet5.Nature").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet5.Range("Sheet5.Nature").Locked = False Then
            Sheet5.Range("Sheet5.Nature").ClearContents
        End If
        If Sheet5.Range("Sheet5.Description").Locked = False Then
            Sheet5.Range("Sheet5.Description").ClearContents
        End If
        If Sheet5.Range("Sheet5.DescEI").Locked = False Then
            Sheet5.Range("Sheet5.DescEI").ClearContents
        End If
        If Sheet5.Range("Sheet5.Amount").Locked = False Then
            Sheet5.Range("Sheet5.Amount").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Exempt (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet5.Range("Sheet5.Nature").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each node In Nodelist
     rowcount = rowcount + 1
     
''        If UCase(Node("NatureDesc")) = "OTH" Then
''            Sheet5.Cells(rowcount, NatureColNo).value = "Any Other"
''        ElseIf UCase(Node("NatureDesc")) = "DMDP" Then
''            Sheet5.Cells(rowcount, NatureColNo).value = "Defense medical disability pension"
''        ElseIf UCase(Node("NatureDesc")) = "AGRI" Then
''            Sheet5.Cells(rowcount, NatureColNo).value = "Agriculture Income (Less than Equal to Rs.5000)"
''        Else
''            text = Findtext("Sec" & (Node("NatureDesc")), "PART4_Nature")
''            Sheet5.Cells(rowcount, NatureColNo).value = text
''        End If
'        If Sheet5.Cells(rowcount, NatureColNo).Locked = False Then
'            If UCase(node("NatureDesc")) = "OTH" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Any Other"
'            'Import-25.01.2023.103.TAX.05A
'            'ElseIf UCase(Node("NatureDesc")) = "CG1L" Then
'            '    Sheet5.Cells(rowcount, NatureColNo).value = "Any exempt income including LTCG on which tax is not payable"
''            ElseIf UCase(Node("NatureDesc")) = "LTCG" Then
''                Sheet5.Cells(rowcount, NatureColNo).value = "LTCG u/s 112A not exceeding  Rs. 1 Lakh"
'            ElseIf UCase(node("NatureDesc")) = "10(34)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(34)-(Exempted Dividend Income)"
'            ElseIf UCase(node("NatureDesc")) = "10(26AAA)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(26AAA)-Any income as referred to in section 10(26AAA)"
'            ElseIf UCase(node("NatureDesc")) = "10(26)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(26)-Any income as referred to in section 10(26)"
'            ElseIf UCase(node("NatureDesc")) = "10(19)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(19)-Armed Forces Family pension in case of death during operational duty"
'            ElseIf UCase(node("NatureDesc")) = "10(10D)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(10D)-Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(1)"
'            ElseIf UCase(node("NatureDesc")) = "10(11)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(11)-Statutory Provident Fund received"
'            ElseIf UCase(node("NatureDesc")) = "10(12)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(12)-Recognized Provident Fund received"
'            ElseIf UCase(node("NatureDesc")) = "10(13)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(13)-Approved superannuation fund received"
'            ElseIf UCase(node("NatureDesc")) = "10(16)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(16)-Scholarships granted to meet the cost of education"
'            ElseIf UCase(node("NatureDesc")) = "10(17)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(17)-Allowance MP/MLA/MLC"
'            ElseIf UCase(node("NatureDesc")) = "10(17A)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(17A)-Award instituted by Government"
'            ElseIf UCase(node("NatureDesc")) = "10(18)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
'            ElseIf UCase(node("NatureDesc")) = "10(10BC)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(10BC)-Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
'            ElseIf UCase(node("NatureDesc")) = "DMDP" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Defense Medical Disability Pension"
'            ElseIf UCase(node("NatureDesc")) = "AGRI" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Agriculture Income (Less than Equal to Rs.5000)"
'            ElseIf UCase(node("NatureDesc")) = "10(12C)" Then
'                Sheet5.Cells(rowcount, NatureColNo).value = "Sec 10(12C)- Any payment from the Agniveer Corpus Fund to a person enrolled under the Agnipath Scheme, or to his nominee"
'            End If
'        End If
       
'Konda updated on 10-03-2026--V0.5.1
 
    If Sheet5.Cells(rowcount, NatureColNo).Locked = False Then
        
            If UCase(node("Category")) = "AGRI" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Agricultural  & related incomes"
            
            ElseIf UCase(node("Category")) = "GOVC" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Compensation/other sums received by government or other approved entities"
            
            ElseIf UCase(node("Category")) = "ISI" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Income from specified Investments"
            
            ElseIf UCase(node("Category")) = "SSRA" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Specified sums received by armed forces personnel"
            
            ElseIf UCase(node("Category")) = "SRSC" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Sums received by Senior Citizens/Minors "
            
            ElseIf UCase(node("Category")) = "SRST" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Sums received by specified Category of Taxpayers"
            
            ElseIf UCase(node("Category")) = "SRPC" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana"
            
            ElseIf UCase(node("Category")) = "OTH" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Other Incomes"
            
            End If
           
    End If
        
    If Sheet5.Cells(rowcount, DescriptionColNo).Locked = False Then
        
        If UCase(node("SubCategory")) = "10(1)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(1)-Agricultural income(Less than or equal to 5000)"
        
        ElseIf UCase(node("SubCategory")) = "10(30)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(30)-subsidy received from or through the Tea Board"
        
        ElseIf UCase(node("SubCategory")) = "10(31)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(31)-Rubber/Coffee/Tea development accounts/funds"
        
        ElseIf UCase(node("SubCategory")) = "10(10BB)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(10BB)-payments made under the Bhopal Gas Leak Disaster"
        
        ElseIf UCase(node("SubCategory")) = "10(10BC)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(10BC)-Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
        
        ElseIf UCase(node("SubCategory")) = "10(17A)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(17A)-Award instituted by Government"
        
        ElseIf UCase(node("SubCategory")) = "10(12AB)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12AB)-any sum received as lump sum amount as per clause (vi) of paragraph 2 of the notification number FX-1/3/2024-PR"
        
        ElseIf UCase(node("SubCategory")) = "10(15)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(15)-Interest on specified securities/investments"
 'Ayush_23/03/2026
'        ElseIf UCase(node("SubCategory")) = "10(23EA)" Then
'        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(23EA)-Contributions received from recognised stock exchanges"
'
        ElseIf UCase(node("SubCategory")) = "10(23FBB)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(23FBB)-income referred to in section 115UB, accruing or arising to, or received by, a unit holder of an investment fund"
        
        ElseIf UCase(node("SubCategory")) = "10(23FD)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(23FD)Unit holder income from Business Trust (certain parts)"
        
        ElseIf UCase(node("SubCategory")) = "10(35)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(35)-Income from specified Mutual Funds "
        
        ElseIf UCase(node("SubCategory")) = "10(35A)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(35A)-distributed income referred to in section 115TA received from a securitisation trust"
        
        ElseIf UCase(node("SubCategory")) = "10(12C)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12C)-Agniveer Corpus Fund income"
        
        ElseIf UCase(node("SubCategory")) = "10(18)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
        
        ElseIf UCase(node("SubCategory")) = "10(19)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(19)-Armed Forces Family pension in case of death during operational duty"
        
        ElseIf UCase(node("SubCategory")) = "10(23AA)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(23AA)-Sum received by any person on behalf of any Fund established by the armed forces"
        
        ElseIf UCase(node("SubCategory")) = "DMD" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "Defense Medical Disability Pension"
        
        ElseIf UCase(node("SubCategory")) = "10(32)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(32)-Minor child’s income—small exemption"
        
        ElseIf UCase(node("SubCategory")) = "10(43)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(43)-Reverse mortgage—payments to senior citizens"
        
        ElseIf UCase(node("SubCategory")) = "10(19A)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(19A)-Annual value of one palace in occupation of ex-ruler"
        
        ElseIf UCase(node("SubCategory")) = "10(26)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(26)-Any income as referred to in section 10(26)"
        
        ElseIf UCase(node("SubCategory")) = "10(26AAA)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(26AAA)-Any income as referred to in section 10(26AAA)"
        
        ElseIf UCase(node("SubCategory")) = "10(10D)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(10D)-Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(10D)"
        
        ElseIf UCase(node("SubCategory")) = "10(11)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(11)-Statutory Provident Fund received"
        
        ElseIf UCase(node("SubCategory")) = "10(11A)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(11A)-Sum received from an account opened under the Sukanya Samriddhi Yojana"
        
        ElseIf UCase(node("SubCategory")) = "10(12)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12)-Recognized Provident Fund received"
        
        ElseIf UCase(node("SubCategory")) = "10(12A)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12A)-Any payment from the National Pension System Trust to an assessee"
        
        ElseIf UCase(node("SubCategory")) = "10(12AA)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12AA)-any payment from the National Pension System Trust "
        
        ElseIf UCase(node("SubCategory")) = "10(12B)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee"
        
        ElseIf UCase(node("SubCategory")) = "10(12BA)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(12BA)-partial withdrawal made from the National Pension System"
        
        ElseIf UCase(node("SubCategory")) = "10(13)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(13)-Approved superannuation fund received"
        
        ElseIf UCase(node("SubCategory")) = "10(25)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(25)-Sum received by trustees on behalf of approved superannuation, gratuity, or pension funds"
 'Ayush_27/03/2026
'        ElseIf UCase(node("SubCategory")) = "10(25A)" Then
'        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(25A)-any income under Employees' State Insurance Fund"
'
        ElseIf UCase(node("SubCategory")) = "10(44)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(44)-Income received by any person for, or on behalf of, the New Pension System Trust"
        
        ElseIf UCase(node("SubCategory")) = "10(2)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(2)-Member’s share from HUF"
        
        ElseIf UCase(node("SubCategory")) = "10(16)" Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "10(16)-Scholarships for education"
        
        'Malli_AY_2026_27  17/06/2026
        ElseIf UCase(node("SubCategory")) = UCase("Incmexmptcircular") Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "Income exempt as per CBDT Circular"
        
        ElseIf UCase(node("SubCategory")) = UCase("Incmexmptnotification") Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "Income exempt as per CBDT Notification"
        
        ElseIf UCase(node("SubCategory")) = UCase("Receiptnotincme") Then
        Sheet5.Cells(rowcount, DescriptionColNo).value = "Receipts not in the nature of income"
        '----------------------------
     
        End If
    End If
'================================
        
'        If node("OthNatOfInc") <> "" And Sheet5.Cells(rowcount, DescriptionColNo).Locked = False Then
'            Sheet5.Cells(rowcount, DescriptionColNo).value = UCase(node("OthNatOfInc"))
'        End If
        
        If Sheet5.Cells(rowcount, DescEIColNo).Locked = False Then
            Sheet5.Cells(rowcount, DescEIColNo).value = node("Description")
        End If
        
        If Sheet5.Cells(rowcount, AmountColNo).Locked = False Then
            Sheet5.Cells(rowcount, AmountColNo).value = UCase(node("OthAmount"))
        End If
        
        cnt = cnt + 1
    Next node
    RecTCS = cnt
    End If
    End If
    End If
End Function

'added by Chetan C M AY 2025-26 on 31/01/2025
Function ImportLTCG112A(jsonText As String)

On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist  As Object

Dim text

Dim TotalSales, TotalCost

Set jsonObject = ParseJson(jsonText)

TotalSales = jsonObject("ITR")("ITR4")("LTCG112A")("TotSaleCnsdrn")
TotalCost = jsonObject("ITR")("ITR4")("LTCG112A")("TotCstAcqisn")

    If TotalSales <> "" And Sheet5.Range("IncD.Sale_LTCG").Locked = False Then
            Sheet5.Range("IncD.Sale_LTCG").value = TotalSales
    End If
    
    If TotalCost <> "" And Sheet5.Range("IncD.Cost_LTCG").Locked = False Then
            Sheet5.Range("IncD.Cost_LTCG").value = TotalCost
    End If
    
End Function

Function ImportSchedule80G100NoAppr(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo, cashAmountColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

'Konda updated 09-03-2026--V0.5
Dim TransactionRefNum_80GA, IFSCCode_80GA
'==============
Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80G")("Don100Percent")("DoneeWithPan")
    
    NameColNo = Sheet6.Range("Per10080G.DoneeName").Column
    AddressColNo = Sheet6.Range("Per10080G.AddrDetail").Column
    CityColNo = Sheet6.Range("Per10080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("Per10080G.StateCode").Column
    PincodeColNo = Sheet6.Range("Per10080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("Per10080G.DoneePAN").Column
    AmountColNo = Sheet6.Range("Per10080G.DonationAmt").Column
    cashAmountColNo = Sheet6.Range("Per10080G.Donationcash").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
'Konda updated 09-03-2026--V0.5
    TransactionRefNum_80GA = Sheet6.Range("Per10080G.Traref").Column
    IFSCCode_80GA = Sheet6.Range("Per10080G.IFSC").Column
'============================================
    TotalExRow = Range("Per10080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.count
    
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("Per10080G.DoneeName").Locked = False Then
            Sheet6.Range("Per10080G.DoneeName").ClearContents
        End If
        If Sheet6.Range("Per10080G.AddrDetail").Locked = False Then
            Sheet6.Range("Per10080G.AddrDetail").ClearContents
        End If
        If Sheet6.Range("Per10080G.CityOrTownOrDistrict").Locked = False Then
            Sheet6.Range("Per10080G.CityOrTownOrDistrict").ClearContents
        End If
        If Sheet6.Range("Per10080G.StateCode").Locked = False Then
            Sheet6.Range("Per10080G.StateCode").ClearContents
        End If
        If Sheet6.Range("Per10080G.PinCode").Locked = False Then
            Sheet6.Range("Per10080G.PinCode").ClearContents
        End If
        If Sheet6.Range("Per10080G.DoneePAN").Locked = False Then
            Sheet6.Range("Per10080G.DoneePAN").ClearContents
        End If
        If Sheet6.Range("Per10080G.DonationAmt").Locked = False Then
            Sheet6.Range("Per10080G.DonationAmt").ClearContents
        End If
        If Sheet6.Range("Per10080G.Donationcash").Locked = False Then
            Sheet6.Range("Per10080G.Donationcash").ClearContents
        End If
        
       'Konda updated 09-03-2026--V0.5
       If Sheet6.Range("Per10080G.Traref").Locked = False Then
            Sheet6.Range("Per10080G.Traref").ClearContents
        End If
        
        If Sheet6.Range("Per10080G.IFSC").Locked = False Then
            Sheet6.Range("Per10080G.IFSC").ClearContents
        End If
       '====================================
        
        
            'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_A (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet6.Range("Per10080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet6.Cells(rowcount, NameColNo).Locked = False Then
            Sheet6.Cells(rowcount, NameColNo).value = node("DoneeWithPanName")
        End If
        If Sheet6.Cells(rowcount, AddressColNo).Locked = False Then
            Sheet6.Cells(rowcount, AddressColNo).value = node("AddressDetail")("AddrDetail")
        End If
        If Sheet6.Cells(rowcount, CityColNo).Locked = False Then
            Sheet6.Cells(rowcount, CityColNo).value = node("AddressDetail")("CityOrTownOrDistrict")
        End If
        If Sheet6.Cells(rowcount, StateCodeColNo).Locked = False Then
             Dim iState As Variant
            iState = UCase(node("AddressDetail")("StateCode"))
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
        End If
        If Sheet6.Cells(rowcount, PincodeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PincodeColNo).value = node("AddressDetail")("PinCode")
        End If
        If Sheet6.Cells(rowcount, PanofDoneeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node("DoneePAN")
        End If
        If Sheet6.Cells(rowcount, cashAmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, cashAmountColNo).value = node("DonationAmtCash")
        End If
         If Sheet6.Cells(rowcount, AmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, AmountColNo).value = node("DonationAmtOtherMode")
        End If
   'Konda updated 09-03-2026--V0.5
        If Sheet6.Cells(rowcount, TransactionRefNum_80GA).Locked = False Then
            Sheet6.Cells(rowcount, TransactionRefNum_80GA).value = node("TransactionRefNum")
        End If
        
        If Sheet6.Cells(rowcount, IFSCCode_80GA).Locked = False Then
            Sheet6.Cells(rowcount, IFSCCode_80GA).value = node("IFSCCode")
        End If
   
   '===========================
        
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    

End Function
Function ImportSchedule80G50NoAppr(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo, cashAmountColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long

'Konda updated 09-03-2026--V0.5
Dim TransactionRefNum_80GB, IFSCCode_80GB
'==============

Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80G")("Don50PercentNoApprReqd")("DoneeWithPan")
    
    NameColNo = Sheet6.Range("PerNO5080G.DoneeName").Column
    AddressColNo = Sheet6.Range("PerNO5080G.AddrDetail").Column
    CityColNo = Sheet6.Range("PerNO5080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("PerNO5080G.StateCode").Column
    PincodeColNo = Sheet6.Range("PerNO5080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("PerNO5080G.DoneePAN").Column
    AmountColNo = Sheet6.Range("PerNO5080G.DonationAmt").Column
    cashAmountColNo = Sheet6.Range("PerNO5080G.DonationCash").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    'Konda updated 09-03-2026--V0.5
    TransactionRefNum_80GB = Sheet6.Range("PerNO5080G.Traref").Column
    IFSCCode_80GB = Sheet6.Range("PerNO5080G.IFSC").Column
'============================================

    TotalExRow = Range("PerNO5080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("PerNO5080G.DoneeName").Locked = False Then
            Sheet6.Range("PerNO5080G.DoneeName").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.AddrDetail").Locked = False Then
            Sheet6.Range("PerNO5080G.AddrDetail").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.CityOrTownOrDistrict").Locked = False Then
            Sheet6.Range("PerNO5080G.CityOrTownOrDistrict").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.StateCode").Locked = False Then
            Sheet6.Range("PerNO5080G.StateCode").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.PinCode").Locked = False Then
            Sheet6.Range("PerNO5080G.PinCode").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.DoneePAN").Locked = False Then
            Sheet6.Range("PerNO5080G.DoneePAN").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.DonationAmt").Locked = False Then
            Sheet6.Range("PerNO5080G.DonationAmt").ClearContents
        End If
        If Sheet6.Range("PerNO5080G.DonationCash").Locked = False Then
            Sheet6.Range("PerNO5080G.DonationCash").ClearContents
        End If
        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
        
        'Konda updated 09-03-2026--V0.5
       If Sheet6.Range("PerNO5080G.Traref").Locked = False Then
            Sheet6.Range("PerNO5080G.Traref").ClearContents
        End If
        
        If Sheet6.Range("PerNO5080G.IFSC").Locked = False Then
            Sheet6.Range("PerNO5080G.IFSC").ClearContents
        End If
       '====================================
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_B (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet6.Range("PerNO5080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet6.Cells(rowcount, NameColNo).Locked = False Then
            Sheet6.Cells(rowcount, NameColNo).value = node("DoneeWithPanName")
        End If
        If Sheet6.Cells(rowcount, AddressColNo).Locked = False Then
            Sheet6.Cells(rowcount, AddressColNo).value = node("AddressDetail")("AddrDetail")
        End If
        If Sheet6.Cells(rowcount, CityColNo).Locked = False Then
            Sheet6.Cells(rowcount, CityColNo).value = node("AddressDetail")("CityOrTownOrDistrict")
        End If
        If Sheet6.Cells(rowcount, StateCodeColNo).Locked = False Then
            Dim iState As Variant
            iState = UCase(node("AddressDetail")("StateCode"))
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
        End If
        If Sheet6.Cells(rowcount, PincodeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PincodeColNo).value = node("AddressDetail")("PinCode")
        End If
        If Sheet6.Cells(rowcount, PanofDoneeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node("DoneePAN")
        End If
         If Sheet6.Cells(rowcount, cashAmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, cashAmountColNo).value = node("DonationAmtCash")
        End If
        If Sheet6.Cells(rowcount, AmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, AmountColNo).value = node("DonationAmtOtherMode")
        End If
        
        'Konda updated 09-03-2026--V0.5
        If Sheet6.Cells(rowcount, TransactionRefNum_80GB).Locked = False Then
            Sheet6.Cells(rowcount, TransactionRefNum_80GB).value = node("TransactionRefNum")
        End If
        
        If Sheet6.Cells(rowcount, IFSCCode_80GB).Locked = False Then
            Sheet6.Cells(rowcount, IFSCCode_80GB).value = node("IFSCCode")
        End If
   
   '===========================

        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function
Function ImportSchedule80G100Appr(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo, cashAmountColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long


'Konda updated 09-03-2026--V0.5
Dim TransactionRefNum_80GC, IFSCCode_80GC
'==============

Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80G")("Don100PercentApprReqd")("DoneeWithPan")
    
    NameColNo = Sheet6.Range("PerYES10080G.DoneeName").Column
    AddressColNo = Sheet6.Range("PerYES10080G.AddrDetail").Column
    CityColNo = Sheet6.Range("PerYES10080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("PerYES10080G.StateCode").Column
    PincodeColNo = Sheet6.Range("PerYES10080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("PerYES10080G.DoneePAN").Column
    cashAmountColNo = Sheet6.Range("PerYES10080G.DonationCash").Column
    AmountColNo = Sheet6.Range("PerYES10080G.DonationAmt").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
'Konda updated 09-03-2026--V0.5
    TransactionRefNum_80GC = Sheet6.Range("PerYES10080G.Traref").Column
    IFSCCode_80GC = Sheet6.Range("PerYES10080G.IFSC").Column
'============================================

    TotalExRow = Range("PerYES10080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("PerYES10080G.DoneeName").Locked = False Then
            Sheet6.Range("PerYES10080G.DoneeName").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.AddrDetail").Locked = False Then
            Sheet6.Range("PerYES10080G.AddrDetail").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.CityOrTownOrDistrict").Locked = False Then
            Sheet6.Range("PerYES10080G.CityOrTownOrDistrict").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.StateCode").Locked = False Then
            Sheet6.Range("PerYES10080G.StateCode").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.PinCode").Locked = False Then
            Sheet6.Range("PerYES10080G.PinCode").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.DoneePAN").Locked = False Then
            Sheet6.Range("PerYES10080G.DoneePAN").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.DonationCash").Locked = False Then
            Sheet6.Range("PerYES10080G.DonationCash").ClearContents
        End If
        If Sheet6.Range("PerYES10080G.DonationAmt").Locked = False Then
            Sheet6.Range("PerYES10080G.DonationAmt").ClearContents
        End If
        
  'Konda updated 09-03-2026--V0.5
       If Sheet6.Range("PerYES10080G.Traref").Locked = False Then
            Sheet6.Range("PerYES10080G.Traref").ClearContents
        End If
        
        If Sheet6.Range("PerYES10080G.IFSC").Locked = False Then
            Sheet6.Range("PerYES10080G.IFSC").ClearContents
        End If
'====================================

        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_C (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet6.Range("PerYES10080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet6.Cells(rowcount, NameColNo).Locked = False Then
            Sheet6.Cells(rowcount, NameColNo).value = node("DoneeWithPanName")
        End If
        If Sheet6.Cells(rowcount, AddressColNo).Locked = False Then
            Sheet6.Cells(rowcount, AddressColNo).value = node("AddressDetail")("AddrDetail")
        End If
        If Sheet6.Cells(rowcount, CityColNo).Locked = False Then
            Sheet6.Cells(rowcount, CityColNo).value = node("AddressDetail")("CityOrTownOrDistrict")
        End If
        If Sheet6.Cells(rowcount, StateCodeColNo).Locked = False Then
            Dim iState As Variant
            iState = UCase(node("AddressDetail")("StateCode"))
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
        End If
        If Sheet6.Cells(rowcount, PincodeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PincodeColNo).value = node("AddressDetail")("PinCode")
        End If
        If Sheet6.Cells(rowcount, PanofDoneeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node("DoneePAN")
        End If
        If Sheet6.Cells(rowcount, cashAmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, cashAmountColNo).value = node("DonationAmtCash")
        End If
        If Sheet6.Cells(rowcount, AmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, AmountColNo).value = node("DonationAmtOtherMode")
        End If
        
'Konda updated 09-03-2026--V0.5
        If Sheet6.Cells(rowcount, TransactionRefNum_80GC).Locked = False Then
            Sheet6.Cells(rowcount, TransactionRefNum_80GC).value = node("TransactionRefNum")
        End If
        
        If Sheet6.Cells(rowcount, IFSCCode_80GC).Locked = False Then
            Sheet6.Cells(rowcount, IFSCCode_80GC).value = node("IFSCCode")
        End If
   
   '===========================

        cnt = cnt + 1
    Next node
    RecTDS1 = cnt

End Function
Function ImportSchedule80G50Appr(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, init As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, ARNColNo, AmountColNo, DonationColNo, cashAmountColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

'Konda updated 09-03-2026--V0.5
Dim TransactionRefNum_80GD, IFSCCode_80GD
'==============


Set jsonObject = ParseJson(jsonText)
    Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80G")("Don50PercentApprReqd")("DoneeWithPan")
    
    NameColNo = Sheet6.Range("Per5080G.DoneeName").Column
    AddressColNo = Sheet6.Range("Per5080G.AddrDetail").Column
    CityColNo = Sheet6.Range("Per5080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("Per5080G.StateCode").Column
    PincodeColNo = Sheet6.Range("Per5080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("Per5080G.DoneePAN").Column
    'Import-25.01.2023.103.80G.02A
    ARNColNo = Sheet6.Range("Per5080G.DoneeARN").Column
    'End Change
    AmountColNo = Sheet6.Range("Per5080G.DonationAmt").Column
    cashAmountColNo = Sheet6.Range("Per5080G.DonationCash").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
'Konda updated 09-03-2026--V0.5
    TransactionRefNum_80GD = Sheet6.Range("Per5080G.Traref").Column
    IFSCCode_80GD = Sheet6.Range("Per5080G.IFSC").Column
'============================================
    
    TotalExRow = Range("Per5080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("Per5080G.DoneeName").Locked = False Then
            Sheet6.Range("Per5080G.DoneeName").ClearContents
        End If
        If Sheet6.Range("Per5080G.AddrDetail").Locked = False Then
            Sheet6.Range("Per5080G.AddrDetail").ClearContents
        End If
        If Sheet6.Range("Per5080G.CityOrTownOrDistrict").Locked = False Then
            Sheet6.Range("Per5080G.CityOrTownOrDistrict").ClearContents
        End If
        If Sheet6.Range("Per5080G.StateCode").Locked = False Then
            Sheet6.Range("Per5080G.StateCode").ClearContents
        End If
        If Sheet6.Range("Per5080G.PinCode").Locked = False Then
            Sheet6.Range("Per5080G.PinCode").ClearContents
        End If
        If Sheet6.Range("Per5080G.DoneePAN").Locked = False Then
            Sheet6.Range("Per5080G.DoneePAN").ClearContents
        End If
        'Import-25.01.2023.103.80G.02B
        If Sheet6.Range("Per5080G.DoneeARN").Locked = False Then
            Sheet6.Range("Per5080G.DoneeARN").ClearContents
        End If
        'End Change
        If Sheet6.Range("Per5080G.DonationAmt").Locked = False Then
            Sheet6.Range("Per5080G.DonationAmt").ClearContents
        End If
        If Sheet6.Range("Per5080G.DonationCash").Locked = False Then
            Sheet6.Range("Per5080G.DonationCash").ClearContents
        End If
        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
        
    'Konda updated 09-03-2026--V0.5
       If Sheet6.Range("Per5080G.Traref").Locked = False Then
            Sheet6.Range("Per5080G.Traref").ClearContents
        End If
        
        If Sheet6.Range("Per5080G.IFSC").Locked = False Then
            Sheet6.Range("Per5080G.IFSC").ClearContents
        End If
       '====================================
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_D (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet6.Range("Per5080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet6.Cells(rowcount, NameColNo).Locked = False Then
            Sheet6.Cells(rowcount, NameColNo).value = node("DoneeWithPanName")
        End If
        If Sheet6.Cells(rowcount, AddressColNo).Locked = False Then
            Sheet6.Cells(rowcount, AddressColNo).value = node("AddressDetail")("AddrDetail")
        End If
        If Sheet6.Cells(rowcount, CityColNo).Locked = False Then
            Sheet6.Cells(rowcount, CityColNo).value = node("AddressDetail")("CityOrTownOrDistrict")
        End If
        If Sheet6.Cells(rowcount, StateCodeColNo).Locked = False Then
            Dim iState As Variant
            iState = UCase(node("AddressDetail")("StateCode"))
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
        End If
        If Sheet6.Cells(rowcount, PincodeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PincodeColNo).value = node("AddressDetail")("PinCode")
        End If
        If Sheet6.Cells(rowcount, PanofDoneeColNo).Locked = False Then
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node("DoneePAN")
        End If
        'Import-25.01.2023.103.80G.02C
        If Sheet6.Cells(rowcount, ARNColNo).Locked = False Then
            Sheet6.Cells(rowcount, ARNColNo).value = node("ArnNbr")
        End If
        'End Change
        If Sheet6.Cells(rowcount, cashAmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, cashAmountColNo).value = node("DonationAmtCash")
        End If
        If Sheet6.Cells(rowcount, AmountColNo).Locked = False Then
            Sheet6.Cells(rowcount, AmountColNo).value = node("DonationAmtOtherMode")
        End If
        
        'Konda updated 09-03-2026--V0.5
        If Sheet6.Cells(rowcount, TransactionRefNum_80GD).Locked = False Then
            Sheet6.Cells(rowcount, TransactionRefNum_80GD).value = node("TransactionRefNum")
        End If
        
        If Sheet6.Cells(rowcount, IFSCCode_80GD).Locked = False Then
            Sheet6.Cells(rowcount, IFSCCode_80GD).value = node("IFSCCode")
        End If
   
   '===========================

        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function

' ============================================= '
' Public Methods
' ============================================= '

''1
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1

    ' Remove vbCr, vbLf, and vbTab from json_String
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Object
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = CreateObject("Scripting.Dictionary")
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function
Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function
Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    ' Parse key with single or double quotes
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function
Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function
Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function
Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    json_SkipSpaces json_String, json_Index

    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
        Case "\"
            ' Escaped string, \\, or \/
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)

            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function
Private Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function
Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    json_SkipSpaces json_String, json_Index

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.InStr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

'AY 2024-25 New Schedule Schema implement by Mallikarjun
Function ImportSchedule80GGC(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Capacity As Variant
Dim DateofDonation_80GGC1, NatureTransColNo, ChequenoColNo, IFSCColNo, NameofDonorColNo, AccountofDonorColNo, PanofDoneeColNo, DonationCashColNo, DonationOtherColNo, dateDonation As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount, Rec80GGC1 As Long
Dim YYYY, MM, DD, strDate As String

'Konda updated on 17-03-2026--V0.6
Dim PoliticalPartyNameColNo, PoliticalPartyPANColNo
'================
Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80GGC")("Schedule80GGCDetails")

    DateofDonation_80GGC1 = Sheet11.Range("SCH_80GGC_DonationDate").Column
    'NatureTransColNo = Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").Column
    ChequenoColNo = Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Column
   ' IFSCColNo = Sheet11.Range("SCH_80GGC_IFSCCode").Column
   'Konda updated on 20-03-2026
'        IFSCColNo = Sheet11.Range("SCH_80GGC_BankAccountNo").Column  'Malli
           IFSCColNo = Sheet11.Range("SCH_80GGC_IFSCCode").Column
'========================================
   ' NameofDonorColNo = Sheet11.Range("NameofDonor_80GGC").Column
    'AccountofDonorColNo = Sheet11.Range("AccountofDonor_80GGC").Column
    DonationCashColNo = Sheet11.Range("SCH_80GGC_DonationAmtCash").Column
    DonationOtherColNo = Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Column
'Konda updated on 17-03-2026--V0.6
    PoliticalPartyNameColNo = Sheet11.Range("SCH_80GGC_Political_Name").Column
    PoliticalPartyPANColNo = Sheet11.Range("SCH_80GGC_Political_PAN").Column
'======================
    TotalExRow = Range("SCH_80GGC_DonationDate").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet11.Range("SCH_80GGC_DonationDate").ClearContents
        'Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").ClearContents
        Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").ClearContents
        Sheet11.Range("SCH_80GGC_IFSCCode").ClearContents
        'Sheet13.Range("NameofDonor_80GGC").ClearContents
        'Sheet13.Range("AccountofDonor_80GGC").ClearContents
        Sheet11.Range("SCH_80GGC_DonationAmtCash").ClearContents
        Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").ClearContents
        'Malli_Legacy22_02/06/2025
     '   Sheet11.Range("SCH_80GGC_BankAccountNo").ClearContents
'Konda updated on 17-03-2026--V0.6
        Sheet11.Range("SCH_80GGC_Political_Name").ClearContents
        Sheet11.Range("SCH_80GGC_Political_PAN").ClearContents
'==============
        
        
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_80GGC (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet11.Range("SCH_80GGC_DonationDate").name)
    rowcount = rowcount - 1
    cnt = 0
    Dim Temp80GGC As Variant
    For Each node In Nodelist
        rowcount = rowcount + 1
         
        dateDonation = node("DonationDate")
        YYYY = Mid(dateDonation, 1, 4)
        MM = Mid(dateDonation, 6, 2)
        DD = Mid(dateDonation, 9, 2)
        strDate = DD & "/" & MM & "/" & YYYY
        Sheet11.Cells(rowcount, DateofDonation_80GGC1).value = strDate 'Date
        
        If node("DonationAmtCash") <> "" And Sheet11.Cells(rowcount, DonationCashColNo).Locked = False Then 'Donation in Cash
            Sheet11.Cells(rowcount, DonationCashColNo).value = node("DonationAmtCash")
        End If
        
        If node("DonationAmtOtherMode") <> "" And Sheet11.Cells(rowcount, DonationOtherColNo).Locked = False Then 'Donation in Other mode
            Sheet11.Cells(rowcount, DonationOtherColNo).value = node("DonationAmtOtherMode")
        End If

'Konda updated on 17-03-2026--V0.6
        If node("PoliticalPartyName") <> "" And Sheet11.Cells(rowcount, PoliticalPartyNameColNo).Locked = False Then 'Donation in Other mode
            Sheet11.Cells(rowcount, PoliticalPartyNameColNo).value = node("PoliticalPartyName")
        End If
        
        If node("PoliticalPartyPAN") <> "" And Sheet11.Cells(rowcount, PoliticalPartyPANColNo).Locked = False Then 'Donation in Other mode
            Sheet11.Cells(rowcount, PoliticalPartyPANColNo).value = node("PoliticalPartyPAN")
        End If
'=====================
'        Dim naturetrans
'
'        If Node("NatureOfTrans") = "UPI" Then
'           naturetrans = "UPI transfer"
'        ElseIf Node("NatureOfTrans") = "Cheque" Then
'           naturetrans = "Cheque issued"
'        ElseIf Node("NatureOfTrans") = "IMPS" Then
'           naturetrans = "IMPS"
'        ElseIf Node("NatureOfTrans") = "NEFT" Then
'           naturetrans = "NEFT"
'        ElseIf Node("NatureOfTrans") = "RTGS" Then
'           naturetrans = "RTGS"
'        End If
'
'
'
'        If Node("NatureOfTrans") <> "" And Sheet11.Cells(rowcount, NatureTransColNo).Locked = False Then 'Transaction
'            Sheet11.Cells(rowcount, NatureTransColNo).value = naturetrans
'        End If
        
        If node("IFSCCode") <> "" And Sheet11.Cells(rowcount, IFSCColNo).Locked = False Then 'IFSC
            Sheet11.Cells(rowcount, IFSCColNo).value = node("IFSCCode")
        End If
        
'        If Node("BankName") <> "" And Sheet13.Cells(rowcount, NameofDonorColNo).Locked = False Then 'Bank Name
'            Sheet11.Cells(rowcount, NameofDonorColNo).Value = Node("BankName")
'        End If
        
'        If Node("BankAccountNo") <> "" And Sheet13.Cells(rowcount, AccountofDonorColNo).Locked = False Then 'Bank Acc Number
'            Sheet11.Cells(rowcount, AccountofDonorColNo).Value = Node("BankAccountNo")
'        End If
        
        If node("TransactionRefNum") <> "" And Sheet11.Cells(rowcount, ChequenoColNo).Locked = False Then 'Transacation Ref
            Sheet11.Cells(rowcount, ChequenoColNo).value = node("TransactionRefNum")
        End If
    
            
           
        cnt = cnt + 1
    Next node
    Rec80GGC1 = cnt
End Function

Function ImportSCH80GGC_old(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim DateofDonation, DonationinCash, DonationinOtherMode, TransRefnumincaseofUPItransfer, IFSCcodeofBank, BankNameofDonor, BankAccountofDonor, ChequenumberIMPSNEFTRTGS, TotalDonation, EligibleAmountofDonation, totDonationincash, totDonationinothermode, totTotalDonation, totEligibleAmountofDonation As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80GGC, cnt, rowcount As Long
Dim YYYY1, MM1, DD1, strDate_80GGC As String
Dim NoticeDateUnderSec, NoticeDateussec As Variant


Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80GGC")("Schedule80GGCDetails")
   
    DateofDonation = Sheet11.Range("SCH_80GGC_DonationDate").Column
    DonationinCash = Sheet11.Range("SCH_80GGC_DonationAmtCash").Column
    DonationinOtherMode = Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Column
    TransRefnumincaseofUPItransfer = Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").Column
    IFSCcodeofBank = Sheet11.Range("SCH_80GGC_IFSCCode").Column
'    BankNameofDonor = Sheet11.Range("SCH_80GGC_BankName").Column
'    BankAccountofDonor = Sheet11.Range("SCH_80GGC_BankAccountNo").Column
    ChequenumberIMPSNEFTRTGS = Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Column
    TotalDonation = Sheet11.Range("SCH_80GGC_DonationAmt").Column
    EligibleAmountofDonation = Sheet11.Range("SCH_80GGC_EligibleDonationAmt").Column
    
    totDonationincash = Sheet11.Range("SCH_80GGC_TotalDonationAmtOtherMode80GGC").Column
    totDonationinothermode = Sheet11.Range("SCH_80GGC_EligibleDonationAmt").Column
    totTotalDonation = Sheet11.Range("SCH_80GGC_TotalDonationsUs80GGC").Column
    totEligibleAmountofDonation = Sheet11.Range("SCH_80GGC_TotalEligibleDonationAmt80GGC").Column

 
    
    TotalExRow = Range("SCH_80GGC_DonationDate").Rows.count

    TotalXMLRow = Nodelist.count
   
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then

      If Sheet11.Range("SCH_80GGC_DonationDate").Locked = False Then
            Sheet11.Range("SCH_80GGC_DonationDate").ClearContents
        End If

If Sheet11.Range("SCH_80GGC_DonationAmtCash").Locked = False Then
            Sheet11.Range("SCH_80GGC_DonationAmtCash").ClearContents
        End If

If Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Locked = False Then
            Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").ClearContents
        End If
If Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").Locked = False Then
            Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").ClearContents
        End If
If Sheet11.Range("SCH_80GGC_IFSCCode").Locked = False Then
            Sheet11.Range("SCH_80GGC_IFSCCode").ClearContents
        End If
'If Sheet11.Range("SCH_80GGC_BankName").Locked = False Then
'            Sheet11.Range("SCH_80GGC_BankName").ClearContents
'        End If
'If Sheet11.Range("SCH_80GGC_BankAccountNo").Locked = False Then
'            Sheet11.Range("SCH_80GGC_BankAccountNo").ClearContents
'        End If
If Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Locked = False Then
            Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").ClearContents
        End If



End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_80GGC (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet11.Range("SCH_80GGC_DonationDate").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        
        If Sheet11.Cells(rowcount, DateofDonation).Locked = False Then
        
        NoticeDateUnderSec = node("DonationDate")
     
        NoticeDateussec = NoticeDateUnderSec
        YYYY1 = Mid(NoticeDateussec, 1, 4)
        MM1 = Mid(NoticeDateussec, 6, 2)
        DD1 = Mid(NoticeDateussec, 9, 2)
        strDate_80GGC = DD1 & "/" & MM1 & "/" & YYYY1
     
        Sheet11.Cells(rowcount, DateofDonation).value = strDate_80GGC
        
            'Sheet11.Cells(rowcount, DateofDonation).value = Node("DonationDate")
            
        End If
If Sheet11.Cells(rowcount, DonationinCash).Locked = False Then
            Sheet11.Cells(rowcount, DonationinCash).value = node("DonationAmtCash")
        End If
If Sheet11.Cells(rowcount, DonationinOtherMode).Locked = False Then
            Sheet11.Cells(rowcount, DonationinOtherMode).value = node("DonationAmtOtherMode")
        End If
If Sheet11.Cells(rowcount, TransRefnumincaseofUPItransfer).Locked = False Then
            Sheet11.Cells(rowcount, TransRefnumincaseofUPItransfer).value = node("TransactionRefNumUPI")
        End If
If Sheet11.Cells(rowcount, IFSCcodeofBank).Locked = False Then
            Sheet11.Cells(rowcount, IFSCcodeofBank).value = node("IFSCCode")
        End If
'If Sheet11.Cells(rowcount, BankNameofDonor).Locked = False Then
'            Sheet11.Cells(rowcount, BankNameofDonor).value = Node("BankName")
'        End If
'If Sheet11.Cells(rowcount, BankAccountofDonor).Locked = False Then
'            Sheet11.Cells(rowcount, BankAccountofDonor).value = Node("BankAccountNo")
'        End If
If Sheet11.Cells(rowcount, ChequenumberIMPSNEFTRTGS).Locked = False Then
            Sheet11.Cells(rowcount, ChequenumberIMPSNEFTRTGS).value = node("ChnImpsNeftRtgsNum")
End If


cnt = cnt + 1
    Next node
    Rec80GGC = cnt
End Function
Sub testt()
AddDiffRows_80GGC (4)
End Sub
Sub AddDiffRows_80GGC(DiffRows As Long)
    'setTblinfo_SCH_80GGC_DonationDate
    setDiffTblinfo_80GGC
    Sheet11.Activate
    searchLastRow ("SCH_80GGC_DonationDate")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub

'=======================================================================
'Function ImportFilingStatus(jsonText As String)
'On Error Resume Next
'
'Dim jsonObject As Object
'Dim jsonDictionary As Object
'Dim ReturnFileSec, AsseseeRepFlg, init, OptingNewTaxRegimeflg, node, Nodelist, NewTaxRegime, Form10IEDate, Form10IEDate_2024, Form10IEAckNo, Form10IEAckNo_2024, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar
'Set jsonObject = ParseJson(jsonText)
'
'    Dim Filingtype As Variant
'    Dim sReturnFile, iReturnFile As Variant
'    Dim sPort5A, iPort5A As Variant
'    Dim DateofOriginalfile As Variant
'    Dim NoticeDateussec As Variant
'    Dim iProvisoFlag, sProvisoFlag As Variant
'    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
'    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
'    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
'
'    ReturnFileSec = jsonObject("ITR")("ITR4")("FilingStatus")("ReturnFileSec")
'    AsseseeRepFlg = jsonObject("ITR")("ITR4")("FilingStatus")("AsseseeRepFlg")
''    OptingNewTaxRegimeflg = jsonObject("ITR")("ITR4")("FilingStatus")("OptingNewTaxRegime")
''    NewTaxRegime = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegime")
''    Form10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEDate")
'
''Import-09.12.2022.103.09A
'
'    Dim AYOptedIn, NewForm10IEDate_1, NewForm10IEDate_2, NewForm10IEDate, NewForm10IEAckNo, OldTaxRegime, AYOptedOut, OldForm10IEDate, OldForm10IEAckNo, ItrFilingDueDate_1, OptOutNewTaxRegime_1, OptOutNewTaxRegime_2, NewForm10IEAckNo_1, NewForm10IEAckNo_2
'    AYOptedIn = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("AssessmentYear")
''    NewForm10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("Form10IEDtls")("Form10IEDate")
''    NewForm10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("Form10IEDtls")("Form10IEAckNo")
'
''commented by Chetan C M on 28/01/2025
''--start--
'
''''PAG_C24,PAG_C25 AY 2024-25 Change Starts---------------------
''
''    NewForm10IEDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate")
''    NewForm10IEAckNo_1 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo")
''
''    ItrFilingDueDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("ItrFilingDueDate")
''    OptOutNewTaxRegime_1 = jsonObject("ITR")("ITR4")("FilingStatus")("OptOutNewTaxRegime")
''
'''New Changes in AY 2025-26
''
''    NewForm10IEDate_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate_AY24-25")
''    NewForm10IEAckNo_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo_AY24-25")
''    OptOutNewTaxRegime_2 = jsonObject("ITR")("ITR4")("FilingStatus")("OptOutNewTaxRegime_AY24-25")
'
''--end--
'
''added by Chetan C M AY 2025-26 on 28/01/2025
''--start--
'    Dim ContOptOutNewTaxReg_25, ContwithNewTaxReg_25, OptOutNewTaxReg_25
'    NewForm10IEDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate")
'    NewForm10IEAckNo_1 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo")
'
'    ItrFilingDueDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("ItrFilingDueDate")
'
'    OptOutNewTaxRegime_1 = jsonObject("ITR")("ITR4")("FilingStatus")("OptOutNewTaxRegime_Form10IEA_AY24_25")
'    ContOptOutNewTaxReg_25 = jsonObject("ITR")("ITR4")("FilingStatus")("Yes_ContOptOutNewTaxReg")
'    ContwithNewTaxReg_25 = jsonObject("ITR")("ITR4")("FilingStatus")("No_OptOutNewTaxReg")
'    OptOutNewTaxReg_25 = jsonObject("ITR")("ITR4")("FilingStatus")("NA_OptOutNewTaxReg")
'
'
'    NewForm10IEDate_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate_AY24_25")
'    NewForm10IEAckNo_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo_AY24_25")
'
''    NewForm10IEDate_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate_AY24-25")
''    NewForm10IEAckNo_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo_AY24-25")
''
''    NewForm10IEDate_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADate_AY24-25")
''    NewForm10IEAckNo_2 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo_AY24-25")
'
''--end--
'''PAG_C24,PAG_C25 AY 2024-25 Change ends ---------------------------------------
'
''    OldTaxRegime = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegime")
''    AYOptedOut = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegimeDtls")("AssessmentYear")
''    OldForm10IEDate = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegimeDtls")("Form10IEDtls")("Form10IEDate")
''    OldForm10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("OptedOutNewTaxRegimeDtls")("Form10IEDtls")("Form10IEAckNo")
'
''---End Change
'
''Konda updated AY2025-26-----------------------
'If NewForm10IEDate_2 <> "" Then
'    Dim YYYY, MM, DD
'    YYYY = Mid(NewForm10IEDate_2, 1, 4)
'    MM = Mid(NewForm10IEDate_2, 6, 2)
'    DD = Mid(NewForm10IEDate_2, 9, 2)
'    Form10IEDate_2024 = DD & "/" & MM & "/" & YYYY
'End If
''End -----------------------------
'
'If NewForm10IEDate_1 <> "" Then
'    'Dim YYYY, MM, DD
'    YYYY = Mid(NewForm10IEDate_1, 1, 4)
'    MM = Mid(NewForm10IEDate_1, 6, 2)
'    DD = Mid(NewForm10IEDate_1, 9, 2)
'    Form10IEDate = DD & "/" & MM & "/" & YYYY
'End If
'
''commented by Chetan C M on 28/01/2025
''    Form10IEAckNo = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAckNo")
''    Form10IEAckNo_2024 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNo_AY24-25")
'
'    SeventhProvisio139 = jsonObject("ITR")("ITR4")("FilingStatus")("SeventhProvisio139")
'    DepAmtAggAmtExcd1CrPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
'    AmtSeventhProvisio139i = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139i")
'    IncrExpAggAmt2LkTrvFrgnCntryFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
'    AmtSeventhProvisio139ii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139ii")
'    IncrExpAggAmt1LkElctrctyPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt1LkElctrctyPrYrFlg")
'    AmtSeventhProvisio139iii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139iii")
'    ReceiptNo = jsonObject("ITR")("ITR4")("FilingStatus")("ReceiptNo")
'    NoticeNo = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeNo")
'    OrigRetFiledDate = jsonObject("ITR")("ITR4")("FilingStatus")("OrigRetFiledDate")
'    NoticeDateUnderSec = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeDateUnderSec")
'
'    RepName = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepName")
'    RepCapacity = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepCapacity")
'    RepAddress = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepAddress")
'    RepPAN = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepPAN")
'    RepAadhaar = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepAadhaar")
'    '-------------------------------------------------------------------
''    Dim ItrFilingDueDate_1 As Variant
''   ItrFilingDueDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("ItrFilingDueDate")
'
'    '------------------------------------------------------------------------
'
'        iReturnFile = ReturnFileSec
'
'        If iReturnFile = "11" Then
'           sReturnFile = "139(1)-On or before due date"
'        ElseIf iReturnFile = "12" Then
'           sReturnFile = "139(4)-After due date"
'        ElseIf iReturnFile = "13" Then
'            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "142(1)"
'        ElseIf iReturnFile = "14" Then
'            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "148"
'        ElseIf iReturnFile = "15" Then
'            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "153A"
'        ElseIf iReturnFile = "16" Then
'        Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'        Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "153C"
'        ElseIf iReturnFile = "17" Then
'           sReturnFile = "139(5)-Revised Return"
'        ElseIf iReturnFile = "18" Then
'        Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
'        Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
'           sReturnFile = "139(9)"
'        ElseIf iReturnFile = "20" Then
'        'J200924
'           sReturnFile = "119(2)(b)- After condonation of delay"
''           sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
'
'        ElseIf iReturnFile = "21" Then
'           sReturnFile = "139(8A)"
'        End If
'
'        If Sheet1.Range("sheet1.ReturnFileSec").Locked = False Then
'            Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
'        End If
'
''commented by Chetan C M on 28/01/2025
''--start--
''Konda Update Starts----------------
''             If OptOutNewTaxRegime_2 = "N" Then
''            Sheet1.Range("Sheet1.115BAC").value = "No"
''
''             NewBacYesValueChange
''
''            ElseIf OptOutNewTaxRegime_2 = "Y" Then
''            Sheet1.Range("Sheet1.115BAC").value = "Yes"
''
''             NewBacNoValueChange_2425
''
''            Sheet1.Unprotect Password:=getmsgstate
''            If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False Then
''            Sheet1.Range("BAC115.DateOfFiling").value = Form10IEDate_2024
''            End If
''
''            Sheet1.Unprotect Password:=getmsgstate
''            If Sheet1.Range("BAC115.AcknowledgementNo").Locked = False Then
''            Sheet1.Range("BAC115.AcknowledgementNo").value = NewForm10IEAckNo_2
''            End If
''
''            Sheet1.Unprotect Password:=getmsgstate
''            ElseIf OptOutNewTaxRegime_2 = "NA" And (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
''            Sheet1.Range("Sheet1.115BAC").value = "Not Applicable"
''            End If
'
''End update--------------------
'
''
'''Mallikarjun development Starts--------------------------------------------------------------
''
''   'OptOutNewTaxRegime_1 = jsonObject("ITR")("ITR4")("FilingStatus")("OptOutNewTaxRegime")
''        Application.EnableEvents = True
''
''            If OptOutNewTaxRegime_1 = "N" Then
'''            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No"
''            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime"
'''                If OptOutNewTaxRegime_1 = "N" And OptOutNewTaxRegime_2 = "N" Then
'''                    NewBacYesValueChange
'''                ElseIf OptOutNewTaxRegime_1 = "N" And OptOutNewTaxRegime_2 = "Y" Then
'''                    NewBacNoValueChange_2425
'''
'''                End If
''             'Newly added by Bindu As its not working
''
''            If Left(Sheet1.Range("sheet1.Status").value, 1) <> "F" Then
''                If OptOutNewTaxRegime_1 = "N" Then
''                          NewBacYesValueChange
''                ElseIf OptOutNewTaxRegime_1 = "Y" Then
''                          NewBacNoValueChange_2425
''
''                End If
''            End If
'
''    'Konda Updated------------------
''            Sheet1.Unprotect Password:=getmsgstate
''
''            If Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False Then
''            Sheet1.Range("Sheet1.Dateoffiling").value = Form10IEDate
''            End If
''
''            Sheet1.Unprotect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
''            Sheet1.Range("Sheet1.acknowledgementNumber_New").value = NewForm10IEAckNo_1
''            End If
''    'End-------------------------------
''            ElseIf OptOutNewTaxRegime_1 = "Y" Then
'''            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date"
''            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime"
''
''             NewBacNoValueChange_2425
''
''            Sheet1.Unprotect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False Then
''            'Sheet1.Range("Sheet1.Dateoffiling").value = Mid(NewForm10IEDate_1, 9, 2) & "/" & Mid(NewForm10IEDate_1, 6, 2) & "/" & Mid(NewForm10IEDate_1, 1, 4)
''            Sheet1.Unprotect Password:=getmsgstate
''            Sheet1.Range("Sheet1.Dateoffiling").value = Form10IEDate
''            End If
''
''            Sheet1.Unprotect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
''            Sheet1.Range("Sheet1.acknowledgementNumber_New").value = NewForm10IEAckNo_1
''            End If
'
'
''            Sheet1.Unprotect Password:=getmsgstate
'
'''            ElseIf OptOutNewTaxRegime_1 = "B" Then
'''            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date"
'''            NewBacYesValueChange
'''
'''            Sheet1.Unprotect Password:=getmsgstate
'''            If Sheet1.Range("Sheet1.Dateoffiling").Locked = False Then
'''            'Sheet1.Range("Sheet1.Dateoffiling").value = Mid(NewForm10IEDate_1, 9, 2) & "/" & Mid(NewForm10IEDate_1, 6, 2) & "/" & Mid(NewForm10IEDate_1, 1, 4)
'''            Sheet1.Range("Sheet1.Dateoffiling").value = Form10IEDate
'''            End If
'''
'''            Sheet1.Unprotect Password:=getmsgstate
'''            If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
'''            Sheet1.Range("Sheet1.acknowledgementNumber_New").value = NewForm10IEAckNo_1
'''            End If
'''
'''            Sheet1.Unprotect Password:=getmsgstate
''            ElseIf OptOutNewTaxRegime_1 = "NA" Then
''            'Ayush
''            'Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Not Applicable"
''            If (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
''                        Sheet1.Unprotect Password:=getmsgstate
''                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.Delete
''                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current5"
''                        Application.EnableEvents = True
''                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Not Applicable"
''                        Application.EnableEvents = False
''                        Sheet1.Unprotect Password:=getmsgstate
''                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Interior.Color = (&HD8D8D8)
''                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").MergeArea.Locked = True
''
''               Else:
''                  Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = ""
''               End If
''               '-------------
''            End If
''
'' 'Mallikarjun development Ends--------------------------------------------------------------
'
''--end--
'
''added by Chetan C M AY 2025-26 on 28/01/2025
''--start--
'    If OptOutNewTaxRegime_1 = "N" Then
'        Sheet1.Range("Sheet1.115BAC").value = "No"
'
'        NewBacYesValueChange
'
'        If ContwithNewTaxReg_25 = "N" Then
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "No"
'
'        Else
'            If ContwithNewTaxReg_25 = "Y" Then
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes"
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False Then
'                Sheet1.Range("BAC115.DateOfFiling_New_No").value = Form10IEDate '
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False Then
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = NewForm10IEAckNo_1
'            End If
'
'        End If
'
'    ElseIf OptOutNewTaxRegime_1 = "Y" Then
'        Sheet1.Range("Sheet1.115BAC").value = "Yes"
'
'        NewBacNoValueChange_2425
'
'        Sheet1.Unprotect Password:=getmsgstate
'        If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False Then
'            Sheet1.Range("BAC115.DateOfFiling").value = Form10IEDate_2024  'NewForm10IEDate_2
'        End If
'
'        Sheet1.Unprotect Password:=getmsgstate
'        If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False Then
'            Sheet1.Range("BAC115.AcknowledgementNo").value = NewForm10IEAckNo_2
'        End If
'
'        If ContOptOutNewTaxReg_25 = "Y" Then
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "Yes"
'
'
'        ElseIf ContOptOutNewTaxReg_25 = "N" Then
'
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False Then
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = Form10IEDate 'NewForm10IEDate_1
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False Then
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = NewForm10IEAckNo_1
'            End If
'        End If
'
'
'    ElseIf OptOutNewTaxRegime_1 = "NA" Then
'        Sheet1.Range("Sheet1.115BAC").value = "Not Applicable"
'
'        If OptOutNewTaxReg_25 = "N" Then
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "No"
'
'        ElseIf OptOutNewTaxReg_25 = "Y" Then
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False Then
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = Form10IEDate 'NewForm10IEDate_1
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False Then
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = NewForm10IEAckNo_1
'            End If
'        End If
'
'
'    End If
'
'
'
''--end--
'
'        iProvisoFlag = SeventhProvisio139
'        If iProvisoFlag = "Y" Then
'        sProvisoFlag = "Yes"
'        ElseIf iProvisoFlag = "N" Then
'        sProvisoFlag = "No"
'        End If
'
'
'
'
''        If NewTaxRegime = "Y" Then
''            Sheet8.Unprotect Password:=getmsgstate
''             Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
''                Sheet8.Range("bacValue").value = 1
''            Sheet8.Protect Password:=getmsgstate
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacYes
''            Sheet1.Protect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''                Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''            End If
''            If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''                Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''            End If
''        ElseIf NewTaxRegime = "N" Then
''            Sheet8.Unprotect Password:=getmsgstate
''                Sheet8.Range("bacValue").value = 2
''                Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
''            Sheet8.Protect Password:=getmsgstate
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacNo
''            Sheet1.Protect Password:=getmsgstate
''            If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''                Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''            End If
''            If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''            Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''            End If
''        Else
''            Sheet8.Shapes("BacYes").OLEFormat.Object.value = 0
''            Sheet8.Shapes("BacNO").OLEFormat.Object.value = 0
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacNo
''            Sheet1.Protect Password:=getmsgstate
''        End If
'
''        If OptingNewTaxRegimeflg < 3 Then
''        Sheet1.Range("OptingNewTaxRegime").value = "No"
''        Else
''        Sheet1.Range("OptingNewTaxRegime").value = "Yes"
''        End If
'
'
''Import-09.12.2022.103.09B
'
''If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
''    If NewTaxRegime = "Y" Then
''            Sheet8.Unprotect Password:=getmsgstate
''             Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
''                Sheet8.Range("bacValue").value = 1
''                Sheet8.Range("oldbacValue").value = 1
''            Sheet8.Protect Password:=getmsgstate
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacYes
''            Sheet1.Protect Password:=getmsgstate
''        ElseIf NewTaxRegime = "N" Then
''            Sheet8.Unprotect Password:=getmsgstate
''                Sheet8.Range("bacValue").value = 2
''                Sheet8.Range("oldbacValue").value = 2
''                Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
''            Sheet8.Protect Password:=getmsgstate
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacNo
''            Sheet1.Protect Password:=getmsgstate
''        Else
''            Sheet8.Shapes("BacYes").OLEFormat.Object.value = 0
''            Sheet8.Shapes("BacNO").OLEFormat.Object.value = 0
''            Sheet1.Unprotect Password:=getmsgstate
''            resetBacNo
''            Sheet1.Protect Password:=getmsgstate
''    End If
''End If
'
'    Application.EnableEvents = True
'
''    If Sheet1.Range("AYinOpting").Locked = False Then
''            Sheet1.Range("AYinOpting").value = AYOptedIn
''    End If
''
''    If Sheet1.Range("Sheet1.date10IE_Opt").Locked = False Then
''            Sheet1.Range("Sheet1.date10IE_Opt").value = Mid(NewForm10IEDate, 9, 2) & "/" & Mid(NewForm10IEDate, 6, 2) & "/" & Mid(NewForm10IEDate, 1, 4)
''    End If
''
''    If Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Locked = False Then
''            Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = NewForm10IEAckNo
''    End If
''
''    If Sheet1.Range("Opt_Out115BAC").Locked = False Then
''            Application.EnableEvents = True
''            If OldTaxRegime = "Y" Then
''                Sheet1.Range("Opt_Out115BAC").value = "Yes"
''            ElseIf OldTaxRegime = "N" Then
''                Sheet1.Range("Opt_Out115BAC").value = "No"
''            End If
''            'Application.EnableEvents = False
''    End If
'
''    If Sheet1.Range("AYinOptOut").Locked = False Then
''        Sheet1.Range("AYinOptOut").value = AYOptedOut
''    End If
''
''    If Sheet1.Range("Sheet1.date10IE_OptOut").Locked = False Then
''            Sheet1.Range("Sheet1.date10IE_OptOut").value = Mid(OldForm10IEDate, 9, 2) & "/" & Mid(OldForm10IEDate, 6, 2) & "/" & Mid(OldForm10IEDate, 1, 4)
''    End If
''
''    If Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Locked = False Then
''            Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = OldForm10IEAckNo
''    End If
'
'
''End Change
''If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
'' If OptingNewTaxRegimeflg = 1 Then
''    Sheet8.Unprotect Password:=getmsgstate
''    Sheet8.Range("bacValue").value = 1
''    Sheet8.Range("oldbacValue").value = 2
''    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
''    Sheet1.Unprotect Password:=getmsgstate
''    'resetBacYes
''    Sheet1.Protect Password:=getmsgstate
''    Sheet1.Range("OptingNewTaxRegime").value = "Opting in now"
''    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''        Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''    End If
''    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''        Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''    End If
''    Sheet8.Protect Password:=getmsgstate
''End If
'
''If OptingNewTaxRegimeflg = 2 Then
''    Sheet8.Unprotect Password:=getmsgstate
''    Sheet8.Range("bacValue").value = 2
''    Sheet8.Range("oldbacValue").value = 2
''    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
''    Sheet1.Unprotect Password:=getmsgstate
''    'resetBacNo
''    Sheet1.Protect Password:=getmsgstate
''    Sheet1.Range("OptingNewTaxRegime").value = "Not opting"
''    Sheet8.Protect Password:=getmsgstate
''End If
'
''If OptingNewTaxRegimeflg = 3 Then
''    Sheet8.Unprotect Password:=getmsgstate
''    Sheet8.Range("bacValue").value = 1
''    Sheet8.Range("oldbacValue").value = 1
''    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
''    Sheet1.Unprotect Password:=getmsgstate
''    'resetBacYes
''    Sheet1.Protect Password:=getmsgstate
''    Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt"
''    Sheet8.Protect Password:=getmsgstate
''End If
'
''If OptingNewTaxRegimeflg = 4 Then
''    Sheet8.Unprotect Password:=getmsgstate
''    Sheet8.Range("bacValue").value = 2
''    Sheet8.Range("oldbacValue").value = 1
''    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
''    Sheet1.Unprotect Password:=getmsgstate
''    'resetBacNo
''    Sheet1.Protect Password:=getmsgstate
''    Sheet1.Range("OptingNewTaxRegime").value = "Opt out"
''    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
''    Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
''    End If
''    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
''    Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
''    End If
''    Sheet8.Protect Password:=getmsgstate
''End If
'
''Import-09.12.2022.103.16
'
''If OptingNewTaxRegimeflg = 5 Then
''    Sheet8.Unprotect Password:=getmsgstate
''    Sheet8.Range("bacValue").value = 2
''    'Sheet8.Range("oldbacValue").value = 2
''    'Sheet1.Shapes("BacNo").OLEFormat.Object.value = 2
''    'Sheet1.Unprotect Password:=getmsgstate
''    'resetBacNo
''    'Sheet1.Protect Password:=getmsgstate
''    'Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in"
'''    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
'''        Sheet1.Range("Sheet1.date10IE").value = Form10IEDate
'''    End If
'''    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False And Form10IEAckNo <> "" Then
'''        Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
'''    End If
''    Sheet8.Protect Password:=getmsgstate
''End If
'
''End Change
''End If
'
'         Application.EnableEvents = True
'
'         If Sheet1.Range("sheet1.SeventhProvisoFlag").Locked = False Then
'            Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
'        End If
'        iDepositAmountFlag = DepAmtAggAmtExcd1CrPrYrFlg
'        If iDepositAmountFlag = "Y" Then
'        sDepositAmountFlag = "Yes"
'        ElseIf iDepositAmountFlag = "N" Then
'        sDepositAmountFlag = "No"
'        End If
'        If Sheet1.Range("Sheet1.DepositAmountFlag").Locked = False Then
'            Sheet1.Range("Sheet1.DepositAmountFlag").value = sDepositAmountFlag
'        End If
'        If Sheet1.Range("Sheet1.DepositAmount").Locked = False Then
'            Sheet1.Range("Sheet1.DepositAmount").value = AmtSeventhProvisio139i
'        End If
'        iAggrigateAmountFlag = IncrExpAggAmt2LkTrvFrgnCntryFlg
'        If iAggrigateAmountFlag = "Y" Then
'        sAggrigateAmountFlag = "Yes"
'        ElseIf iAggrigateAmountFlag = "N" Then
'        sAggrigateAmountFlag = "No"
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmountFlag").value = sAggrigateAmountFlag
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmount").value = AmtSeventhProvisio139ii
'        End If
'        iAggrigateAmountFlag1 = IncrExpAggAmt1LkElctrctyPrYrFlg
'        If iAggrigateAmountFlag1 = "Y" Then
'        sAggrigateAmountFlag1 = "Yes"
'        ElseIf iAggrigateAmountFlag1 = "N" Then
'        sAggrigateAmountFlag1 = "No"
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmountFlag1").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = sAggrigateAmountFlag1
'        End If
'        If Sheet1.Range("Sheet1.AggrigateAmount1").Locked = False Then
'            Sheet1.Range("Sheet1.AggrigateAmount1").value = AmtSeventhProvisio139iii
'        End If
'
'
'        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1").value
'
'        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
'
'        If ReturnFileSec = "17" Then
'            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
'                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            DateofOriginalfile = OrigRetFiledDate
'            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
'                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
'            End If
'        Else
'            If Sheet1.Range("Sheet1.ReceiptNo").Locked = False Then
'                Sheet1.Range("Sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            DateofOriginalfile = OrigRetFiledDate
'            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
'                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
'            End If
'            If Sheet1.Range("sheet1.NoticeNo").Locked = False Then
'                Sheet1.Range("sheet1.NoticeNo").value = NoticeNo
'            End If
'            NoticeDateussec = NoticeDateUnderSec
'            If NoticeDateussec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
'                Sheet1.Range("sheet1.NoticeDate").value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
'            End If
'        End If
'
'
'
'
'
''            iPort5A = UCase(Node.SelectSingleNode("PortugeseCC5A").Text)
''            If iPort5A = "Y" Then
''            sPort5A = "Yes"
''
''            End If
'
''            If iPort5A = "N" Then
''            sPort5A = "No"
'
''            End If
'
''import 139 clause iv
'
' Set init = jsonObject("ITR")("ITR4")("FilingStatus")
'    If init.exists("clauseiv7provisio139i") Then
'        If init("clauseiv7provisio139i") = "Y" Then
'        Sheet1.Range("clauseiv7provisio139iFlg") = "Yes"
'        Sheet1.Range("clauseiv7provisio139iFlg_1") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_2") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_3") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_4") = "No"
'
'            For Each node In init("clauseiv7provisio139iDtls")
'                If node("clauseiv7provisio139iNature") = "1" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_1") = node("clauseiv7provisio139iAmount")
'                ElseIf node("clauseiv7provisio139iNature") = "2" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_2") = node("clauseiv7provisio139iAmount")
'                ElseIf node("clauseiv7provisio139iNature") = "3" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_3") = node("clauseiv7provisio139iAmount")
'                ElseIf node("clauseiv7provisio139iNature") = "4" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_4") = node("clauseiv7provisio139iAmount")
'                End If
'            Next node
'
'        Else
'            Sheet1.Range("clauseiv7provisio139iFlg") = "No"
'        End If
'
'
'    End If
'
'
'
'
'
'
'           Dim RepAssessee As Variant
'           Dim sRepAssessee As Variant
'
'           RepAssessee = AsseseeRepFlg
'           If RepAssessee = "Y" Then
'               sRepAssessee = "Yes"
'
'           ElseIf RepAssessee = "N" Then
'               sRepAssessee = "No"
'           End If
'           If Sheet1.Range("Sheet1.RepAssessee").Locked = False Then
'            Sheet1.Range("Sheet1.RepAssessee").value = sRepAssessee
'           End If
'           If Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.NameRepAssessee").value = RepName
'            End If
'           Dim iCapacity As Variant
'           Dim sCapacity As Variant
'
'           iCapacity = RepCapacity
'
'           If iCapacity = "L" Then
'                sCapacity = "Legal Heir"
'
'           ElseIf iCapacity = "M" Then
'                sCapacity = "Manager"
'
'           ElseIf iCapacity = "G" Then
'                sCapacity = "Guardian"
'
'           ElseIf iCapacity = "O" Then
'                sCapacity = "Other"
'
'           End If
'           If Sheet1.Range("sheet1.CapacityRep").Locked = False Then
'            Sheet1.Range("sheet1.CapacityRep").value = sCapacity
'           End If
'           If Sheet1.Range("sheet1.AddressRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.AddressRepAssessee").value = RepAddress
'           End If
'           If Sheet1.Range("sheet1.PANRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.PANRepAssessee").value = RepPAN
'        End If
'        If Sheet1.Range("sheet1.AadhaarRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.AadhaarRepAssessee").value = RepAadhaar
'        End If
'
'      '"ItrFilingDueDate" "2024-07-31"
'       'Sheet1.Duedate "31/07/2024"
'
''PAG_C9 AY 2024-25-----------
'
'        If Sheet1.Range("Sheet1.Duedate").Locked = True Then
'        Dim ItrFilingDueDate_2 As Variant
'
'        ItrFilingDueDate_2 = Mid(ItrFilingDueDate_1, 1, 4) & "/" & Mid(ItrFilingDueDate_1, 6, 2) & "/" & Mid(ItrFilingDueDate_1, 9, 2)
'            Sheet1.Range("Sheet1.Duedate").value = ItrFilingDueDate_2
'        End If
'
''PAG_C9 AY 2024-25-----------
'
''          Sheet1.Range("sheet1.PANOFSPOUSE").value = UCase(Node.SelectSingleNode("PANOfSpouse").Text)
'
'End Function
'=====================================================================

Function ImportFilingStatus(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object

Dim ReturnFileSec, AsseseeRepFlg, init, OptingNewTaxRegimeflg, node, Nodelist, NewTaxRegime, Form10IEDate, Form10IEDate_2024, Form10IEDate_New, Form10IEDate_Old, Form10IEAckNo, Form10IEAckNo_2024, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar, ItrFilingDueDate_1
Dim RepName1, RepEmailID1, RepCountryCodeRepMobileNo1, RepMobileNo1
Set jsonObject = ParseJson(jsonText)

    Dim Filingtype As Variant
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim DateofOriginalfile As Variant
    Dim NoticeDateussec As Variant
    Dim iProvisoFlag, sProvisoFlag As Variant
    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
    
    ReturnFileSec = jsonObject("ITR")("ITR4")("FilingStatus")("ReturnFileSec")
    AsseseeRepFlg = jsonObject("ITR")("ITR4")("FilingStatus")("AsseseeRepFlg")
  
'Import-09.12.2022.103.09A

'    Dim AYOptedIn, NewForm10IEDate_1, NewForm10IEDate_2, NewForm10IEDate, NewForm10IEAckNo, OldTaxRegime, AYOptedOut, OldForm10IEDate, OldForm10IEAckNo, ItrFilingDueDate_1, OptOutNewTaxRegime_1, OptOutNewTaxRegime_2, NewForm10IEAckNo_1, NewForm10IEAckNo_2
'    AYOptedIn = jsonObject("ITR")("ITR4")("FilingStatus")("NewTaxRegimeDtls")("AssessmentYear")

ItrFilingDueDate_1 = jsonObject("ITR")("ITR4")("FilingStatus")("ItrFilingDueDate")

ItrFilingDueDate_1 = Mid(ItrFilingDueDate_1, 1, 4) & "/" & Mid(ItrFilingDueDate_1, 6, 2) & "/" & Mid(ItrFilingDueDate_1, 9, 2)
            Sheet1.Range("Sheet1.Duedate").value = ItrFilingDueDate_1
    
''Konda updated on 22-01-2026 AY 2026-27
'--start--
    Dim Form10IEAEarlierAYOldRegime_2627, Form10IEAAssYear_2627, Form10IEAEarlierAYAckOldRegime_2627, F10IEAEarlierAYNewRegime_2627, AssYrF10IEANewTaxReg_2627, Form10IEAEarlierAYAckNewRegime_2627, F10IEACurrAYNewRegime_2627, F10IEACurrAYOldRegime_2627
    Dim F10IEADateCurrAYNewTax_2627, F10IEAAckNoCurrAYNewTax_2627, F10IEADateCurrAYOldTax_2627, F10IEAAckNoCurrAYOldTax_2627
    
    Form10IEAEarlierAYOldRegime_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAEarlierAYOldRegime")
    Form10IEAAssYear_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAssYear")
    Form10IEAEarlierAYAckOldRegime_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAEarlierAYAckOldRegime")
    
    F10IEAEarlierAYNewRegime_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEAEarlierAYNewRegime")
    AssYrF10IEANewTaxReg_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("AssYrF10IEANewTaxReg")
    Form10IEAEarlierAYAckNewRegime_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAEarlierAYAckNewRegime")
    
    
    F10IEACurrAYNewRegime_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEACurrAYNewRegime")
    F10IEACurrAYOldRegime_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEACurrAYOldRegime")
    
    
'    Form10IEADateCurrAY_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEADateCurrAY") 'Date
'    Form10IEAAckNoCurrAY_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("Form10IEAAckNoCurrAY")
    F10IEADateCurrAYNewTax_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEADateCurrAYNewTax") 'Date
    F10IEAAckNoCurrAYNewTax_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEAAckNoCurrAYNewTax")
    
    F10IEADateCurrAYOldTax_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEADateCurrAYOldTax") 'Date
    F10IEAAckNoCurrAYOldTax_2627 = jsonObject("ITR")("ITR4")("FilingStatus")("F10IEAAckNoCurrAYOldTax")

If F10IEADateCurrAYNewTax_2627 <> "" Then
    Dim YYYY, MM, DD
    YYYY = Mid(F10IEADateCurrAYNewTax_2627, 1, 4)
    MM = Mid(F10IEADateCurrAYNewTax_2627, 6, 2)
    DD = Mid(F10IEADateCurrAYNewTax_2627, 9, 2)
    Form10IEDate_New = DD & "/" & MM & "/" & YYYY
End If
'=======================================================

If F10IEADateCurrAYOldTax_2627 <> "" Then
    Dim YYYY1, MM1, DD1
    YYYY1 = Mid(F10IEADateCurrAYOldTax_2627, 1, 4)
    MM1 = Mid(F10IEADateCurrAYOldTax_2627, 6, 2)
    DD1 = Mid(F10IEADateCurrAYOldTax_2627, 9, 2)
    Form10IEDate_Old = DD1 & "/" & MM1 & "/" & YYYY1
End If

    SeventhProvisio139 = jsonObject("ITR")("ITR4")("FilingStatus")("SeventhProvisio139")
    DepAmtAggAmtExcd1CrPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
    AmtSeventhProvisio139i = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139i")
    IncrExpAggAmt2LkTrvFrgnCntryFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
    AmtSeventhProvisio139ii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139ii")
    IncrExpAggAmt1LkElctrctyPrYrFlg = jsonObject("ITR")("ITR4")("FilingStatus")("IncrExpAggAmt1LkElctrctyPrYrFlg")
    AmtSeventhProvisio139iii = jsonObject("ITR")("ITR4")("FilingStatus")("AmtSeventhProvisio139iii")
    ReceiptNo = jsonObject("ITR")("ITR4")("FilingStatus")("ReceiptNo")
    NoticeNo = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeNo")
    OrigRetFiledDate = jsonObject("ITR")("ITR4")("FilingStatus")("OrigRetFiledDate")
    NoticeDateUnderSec = jsonObject("ITR")("ITR4")("FilingStatus")("NoticeDateUnderSec")
    
    RepName1 = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepName")
    RepEmailID1 = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepEmailID")
    RepCountryCodeRepMobileNo1 = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("CountryCodeRepMobileNo")
    RepMobileNo1 = jsonObject("ITR")("ITR4")("FilingStatus")("AssesseeRep")("RepMobileNo")
    '-------------------------------------------------------------------

        iReturnFile = ReturnFileSec
            
        If iReturnFile = "11" Then
           sReturnFile = "139(1)-On or before due date"
        ElseIf iReturnFile = "12" Then
           sReturnFile = "139(4)-After due date"
        ElseIf iReturnFile = "13" Then
            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
           sReturnFile = "142(1)"
        ElseIf iReturnFile = "14" Then
            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
           sReturnFile = "148"
        ElseIf iReturnFile = "15" Then
            Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
            Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
           sReturnFile = "153A"
        ElseIf iReturnFile = "16" Then
        Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
        Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
           sReturnFile = "153C"
        ElseIf iReturnFile = "17" Then
           sReturnFile = "139(5)-Revised Return"
        ElseIf iReturnFile = "18" Then
        Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
        Sheet1.Shapes("option button 929").OLEFormat.Object.value = 0
           sReturnFile = "139(9)"
        ElseIf iReturnFile = "20" Then
        'J200924
           sReturnFile = "119(2)(b)- After condonation of delay"
'           sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
           
        ElseIf iReturnFile = "21" Then
           sReturnFile = "139(8A)"
        End If
        
        If Sheet1.Range("sheet1.ReturnFileSec").Locked = False Then
            Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
        End If

'Konda updated on 22-01-2026 AY 2026-27
'--start--
    If Form10IEAEarlierAYOldRegime_2627 = "N" Then
        Sheet1.Range("Sheet1.115BAC").value = "No"
    
        NewBacYesValueChange
        
        If F10IEACurrAYOldRegime_2627 = "N" Then
            Sheet1.Range("BAC115.A23b").value = "No"
            
        Else
            If F10IEACurrAYOldRegime_2627 = "Y" Then
            Sheet1.Range("BAC115.A23b").value = "Yes"
            End If
            
            Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False Then
                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = Form10IEDate_Old
            End If
            
            Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False Then
                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = F10IEAAckNoCurrAYOldTax_2627
            End If
            
        End If
    
    ElseIf Form10IEAEarlierAYOldRegime_2627 = "Y" Then
        Sheet1.Range("Sheet1.115BAC").value = "Yes"
    
        NewBacNoValueChange_2425
        
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("BAC115.AssessmentYear_2425").MergeArea.Locked = False Then
            Sheet1.Range("BAC115.AssessmentYear_2425").value = Form10IEAAssYear_2627
        End If
            
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False Then
            Sheet1.Range("BAC115.AcknowledgementNo").value = Form10IEAEarlierAYAckOldRegime_2627
        End If
        
        If F10IEAEarlierAYNewRegime_2627 = "Y" Then
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "Yes"
            
            Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = False Then
                Sheet1.Range("BAC115.AssessmentYear_Aii").value = AssYrF10IEANewTaxReg_2627
            End If
            
            Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False Then
                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = Form10IEAEarlierAYAckNewRegime_2627
            End If
        ElseIf F10IEAEarlierAYNewRegime_2627 = "N" Then
            
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No"
            If F10IEACurrAYNewRegime_2627 <> "" Then
            If F10IEACurrAYNewRegime_2627 = "Y" Then
                Sheet1.Range("BAC115.A23ii_No").value = "Yes"
'            End If 'Commented by Konda on 05-02-2026--SIT-109767
            
            Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False Then
                Sheet1.Range("BAC115.DateOfFiling_New_No").value = Form10IEDate_New
            End If
            
            Sheet1.Unprotect Password:=getmsgstate
            If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False Then
                Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = F10IEAAckNoCurrAYNewTax_2627
            End If
            ElseIf F10IEACurrAYNewRegime_2627 = "N" Then
            
                If F10IEACurrAYNewRegime_2627 = "N" Then
                Sheet1.Range("BAC115.A23ii_No").value = "No"
                End If
                End If 'Added by Konda  in SIT-109767 on 05-02-2026
            End If
        End If
    End If
'========================done=======================

        iProvisoFlag = SeventhProvisio139
        If iProvisoFlag = "Y" Then
        sProvisoFlag = "Yes"
        ElseIf iProvisoFlag = "N" Then
        sProvisoFlag = "No"
        End If
       
        
    Application.EnableEvents = True

         Application.EnableEvents = True
         
         If Sheet1.Range("sheet1.SeventhProvisoFlag").Locked = False Then
            Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
        End If
        iDepositAmountFlag = DepAmtAggAmtExcd1CrPrYrFlg
        If iDepositAmountFlag = "Y" Then
        sDepositAmountFlag = "Yes"
        ElseIf iDepositAmountFlag = "N" Then
        sDepositAmountFlag = "No"
        End If
        If Sheet1.Range("Sheet1.DepositAmountFlag").Locked = False Then
            Sheet1.Range("Sheet1.DepositAmountFlag").value = sDepositAmountFlag
        End If
        If Sheet1.Range("Sheet1.DepositAmount").Locked = False Then
            Sheet1.Range("Sheet1.DepositAmount").value = AmtSeventhProvisio139i
        End If
        iAggrigateAmountFlag = IncrExpAggAmt2LkTrvFrgnCntryFlg
        If iAggrigateAmountFlag = "Y" Then
        sAggrigateAmountFlag = "Yes"
        ElseIf iAggrigateAmountFlag = "N" Then
        sAggrigateAmountFlag = "No"
        End If
        If Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False Then
            Sheet1.Range("Sheet1.AggrigateAmountFlag").value = sAggrigateAmountFlag
        End If
        If Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
            Sheet1.Range("Sheet1.AggrigateAmount").value = AmtSeventhProvisio139ii
        End If
        iAggrigateAmountFlag1 = IncrExpAggAmt1LkElctrctyPrYrFlg
        If iAggrigateAmountFlag1 = "Y" Then
        sAggrigateAmountFlag1 = "Yes"
        ElseIf iAggrigateAmountFlag1 = "N" Then
        sAggrigateAmountFlag1 = "No"
        End If
        If Sheet1.Range("Sheet1.AggrigateAmountFlag1").Locked = False Then
            Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = sAggrigateAmountFlag1
        End If
        If Sheet1.Range("Sheet1.AggrigateAmount1").Locked = False Then
            Sheet1.Range("Sheet1.AggrigateAmount1").value = AmtSeventhProvisio139iii
        End If
   
   
        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1").value
                
        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
        
        If ReturnFileSec = "17" Then
            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
            End If
            DateofOriginalfile = OrigRetFiledDate
            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
            End If
        Else
            If Sheet1.Range("Sheet1.ReceiptNo").Locked = False Then
                Sheet1.Range("Sheet1.ReceiptNo").value = ReceiptNo
            End If
            DateofOriginalfile = OrigRetFiledDate
            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
            End If
            If Sheet1.Range("sheet1.NoticeNo").Locked = False Then
                Sheet1.Range("sheet1.NoticeNo").value = NoticeNo
            End If
            NoticeDateussec = NoticeDateUnderSec
            If NoticeDateussec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
                Sheet1.Range("sheet1.NoticeDate").value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
            End If
        End If
            
    
        


'            iPort5A = UCase(Node.SelectSingleNode("PortugeseCC5A").Text)
'            If iPort5A = "Y" Then
'            sPort5A = "Yes"
'
'            End If

'            If iPort5A = "N" Then
'            sPort5A = "No"

'            End If

'import 139 clause iv

 Set init = jsonObject("ITR")("ITR4")("FilingStatus")
    If init.exists("clauseiv7provisio139i") Then
        If init("clauseiv7provisio139i") = "Y" Then
        Sheet1.Range("clauseiv7provisio139iFlg") = "Yes"
        Sheet1.Range("clauseiv7provisio139iFlg_1") = "No"
        Sheet1.Range("clauseiv7provisio139iFlg_2") = "No"
        Sheet1.Range("clauseiv7provisio139iFlg_3") = "No"
        Sheet1.Range("clauseiv7provisio139iFlg_4") = "No"
            
            For Each node In init("clauseiv7provisio139iDtls")
                If node("clauseiv7provisio139iNature") = "1" Then
                    Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes"
                    Sheet1.Range("clauseiv7provisio139iAmount_1") = node("clauseiv7provisio139iAmount")
                ElseIf node("clauseiv7provisio139iNature") = "2" Then
                    Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes"
                    Sheet1.Range("clauseiv7provisio139iAmount_2") = node("clauseiv7provisio139iAmount")
                ElseIf node("clauseiv7provisio139iNature") = "3" Then
                    Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
                    Sheet1.Range("clauseiv7provisio139iAmount_3") = node("clauseiv7provisio139iAmount")
                ElseIf node("clauseiv7provisio139iNature") = "4" Then
                    Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
                    Sheet1.Range("clauseiv7provisio139iAmount_4") = node("clauseiv7provisio139iAmount")
                End If
            Next node
            
        Else
            Sheet1.Range("clauseiv7provisio139iFlg") = "No"
        End If
        
        
    End If

   
    
    


        
   If AsseseeRepFlg <> "" Then
           Dim RepAssessee As Variant
           Dim sRepAssessee As Variant
            
           RepAssessee = AsseseeRepFlg
           If RepAssessee = "Y" Then
               sRepAssessee = "Yes"
            
           ElseIf RepAssessee = "N" Then
               sRepAssessee = "No"
           End If
           If Sheet1.Range("Sheet1.RepAssessee").Locked = False Then
            Sheet1.Range("Sheet1.RepAssessee").value = sRepAssessee
           End If
           If Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.NameRepAssessee").value = RepName1
            End If
            
            If Sheet1.Range("sheet1.EmailRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.EmailRepAssessee").value = RepEmailID1
            End If
           
           If Sheet1.Range("sheet1.CountryCodeRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.CountryCodeRepAssessee").value = RepCountryCodeRepMobileNo1
            End If
            
            If Sheet1.Range("sheet1.MobileRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.MobileRepAssessee").value = RepMobileNo1
            End If
       End If
           
        
      '"ItrFilingDueDate" "2024-07-31"
       'Sheet1.Duedate "31/07/2024"
      
'PAG_C9 AY 2024-25-----------
        
        If Sheet1.Range("Sheet1.Duedate").Locked = True Then
        Dim ItrFilingDueDate_2 As Variant
        
        ItrFilingDueDate_2 = Mid(ItrFilingDueDate_1, 1, 4) & "/" & Mid(ItrFilingDueDate_1, 6, 2) & "/" & Mid(ItrFilingDueDate_1, 9, 2)
            Sheet1.Range("Sheet1.Duedate").value = ItrFilingDueDate_2
        End If
        
'PAG_C9 AY 2024-25-----------
        
'          Sheet1.Range("sheet1.PANOFSPOUSE").value = UCase(Node.SelectSingleNode("PANOfSpouse").Text)
    
End Function
Function ImportSchedule80DD_Old(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Nature As Variant
Dim Date10IA_2 As Variant


Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR4")
 
If node.exists("Schedule80DD") Then
If Trim(node("NatureOfDisability")) <> "" Then

If node("NatureOfDisability") = "1" Then
   Sheet14.Range("NatureDisability_80DD").value = "1-Dependent person with disability"

   ElseIf node("NatureOfDisability") = "2" Then
   Sheet14.Range("NatureDisability_80DD").value = "2-Dependent person with severe disability"

   Else: Sheet14.Range("NatureDisability_80DD").value = "(Select)"

End If
End If
   
  
If Trim(node("DependentType")) <> "" Then

If node("DependentType") = "1" Then
   Sheet14.Range("Dependent_80DD").value = "1-Spouse"
   
   ElseIf node("DependentType") = "2" Then
   Sheet14.Range("Dependent_80DD").value = "2-Son"
   
   ElseIf node("DependentType") = "3" Then
   Sheet14.Range("Dependent_80DD").value = "3-Daughter"
   
   ElseIf node("DependentType") = "4" Then
   Sheet14.Range("Dependent_80DD").value = "4-Father"
   
   ElseIf node("DependentType") = "5" Then
   Sheet14.Range("Dependent_80DD").value = "5-Mother"
   
   ElseIf node("DependentType") = "6" Then
   Sheet14.Range("Dependent_80DD").value = "6-Brother"
   
   ElseIf node("DependentType") = "7" Then
   Sheet14.Range("Dependent_80DD").value = "7-Sister"
   
   ElseIf node("DependentType") = "8" Then
   Sheet14.Range("Dependent_80DD").value = "8-Member of the HUF (in case of HUF)"
   
   Else
   Sheet14.Range("Dependent_80DD").value = "(Select)"
   
   End If
   
End If

If Trim(node("DependentPan")) <> "" Then
 Sheet14.Range("PanDependent_80DD").value = node("DependentPan")
End If

If Trim(node("DependentAadhaar")) <> "" Then
 Sheet14.Range("AadhaarDependent_80DD").value = node("DependentAadhaar")
End If

If Trim(node("Form10IAFilingDate")) <> "" Then
  Date10IA_2 = Mid(node("Form10IAFilingDate"), 9, 2) & "/" & Mid(node("Form10IAFilingDate"), 6, 2) & "/" & Mid(node("Form10IAFilingDate"), 1, 4)
  Sheet14.Range("Form10IA_80DD").value = Date10IA_2
End If

If Trim(node("Form10IAAckNum")) <> "" Then
 Sheet14.Range("AcknowledgeNum_80DD").value = node("Form10IAAckNum")
End If


If Trim(node("UDIDNum")) <> "" Then
 Sheet14.Range("UDIDNumber_80DD").value = node("UDIDNum")
End If
End If

End Function


Function ImportSchedule80U_Old(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Date_New As Variant
 
Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR4")

If node.exists("Schedule80U") Then
If Trim(node("NatureOfDisability")) <> "" Then
 
If node("NatureOfDisability") = "1" Then
   Sheet14.Range("NatureDisability_80U").value = Trim("1-Self with disability")
   ElseIf node("NatureOfDisability") = "2" Then
   Sheet14.Range("NatureDisability_80U").value = Trim("2-Self with severe disability")
   Else: Sheet14.Range("NatureDisability_80U").value = "(Select)"
End If
End If
 
  
If Trim(node("Form10IAFilingDate")) <> "" Then

  Date_New = Mid(node("Form10IAFilingDate"), 9, 2) & "/" & Mid(node("Form10IAFilingDate"), 6, 2) & "/" & Mid(node("Form10IAFilingDate"), 1, 4)
  Sheet14.Range("DateFiling10IA_80U").value = Date_New
End If
 
If Trim(node("Form10IAAckNum")) <> "" Then
Sheet14.Range("AckNumber10IA_80U").value = UCase(node("Form10IAAckNum"))
End If
 
If Trim(node("UDIDNum")) <> "" Then
Sheet14.Range("UDIDNum_80U").value = UCase(node("UDIDNum"))
End If
 End If
 
End Function
'Malli comented
Function ImportSchedule80DD_1(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'Dim node1, nodelist1 As Object
'Dim Nature As Variant
'Dim Date10IA_2 As Variant
'
'
'Set jsonObject = ParseJson(jsonText)
'Set Nodelist = jsonObject("ITR")("ITR4")
'
'If node.exists("Schedule80DD") Then
'
''Set nodelist1 = jsonObject("ITR")("ITR4")("Schedule80DD")
'If Trim(jsonObject("ITR")("ITR4")("Schedule80DD")("NatureOfDisability")) <> "" Then
'
'If jsonObject("ITR")("ITR4")("Schedule80DD")("NatureOfDisability") = "1" Then
'   Sheet14.Range("NatureDisability_80DD").value = "1-Dependent person with disability"
'
'   ElseIf node("NatureOfDisability") = "2" Then
'   Sheet14.Range("NatureDisability_80DD").value = "2-Dependent person with severe disability"
'
'   Else: Sheet14.Range("NatureDisability_80DD").value = "(Select)"
'
'End If
'End If
'
''If node.exists("Schedule80DD") Then
'
'If jsonObject("ITR")("ITR4")("Schedule80DD")("DependentType") <> "" Then
'
'
'Dim dependentType_80DD
'dependentType_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("DependentType")
'If dependentType_80DD = "1" Then
'   Sheet14.Range("Dependent_80DD").value = "1-Spouse"
'
'   ElseIf dependentType_80DD = "2" Then
'   Sheet14.Range("Dependent_80DD").value = "2-Son"
'
'   ElseIf dependentType_80DD = "3" Then
'   Sheet14.Range("Dependent_80DD").value = "3-Daughter"
'
'   ElseIf dependentType_80DD = "4" Then
'   Sheet14.Range("Dependent_80DD").value = "4-Father"
'
'   ElseIf dependentType_80DD = "5" Then
'   Sheet14.Range("Dependent_80DD").value = "5-Mother"
'
'   ElseIf dependentType_80DD = "6" Then
'   Sheet14.Range("Dependent_80DD").value = "6-Brother"
'
'   ElseIf dependentType_80DD = "7" Then
'   Sheet14.Range("Dependent_80DD").value = "7-Sister"
'
'   ElseIf dependentType_80DD = "8" Then
'   Sheet14.Range("Dependent_80DD").value = "8-Member of the HUF (in case of HUF)"
'
'   Else
'   Sheet14.Range("Dependent_80DD").value = "(Select)"
'
'   End If
'
'End If
'
'If Trim(jsonObject("ITR")("ITR4")("Schedule80DD")("DependentPan")) <> "" Then
' Sheet14.Range("PanDependent_80DD").value = jsonObject("ITR")("ITR4")("Schedule80DD")("DependentPan")
'End If
'
'If Trim(node("DependentAadhaar")) <> "" Then
' Sheet14.Range("AadhaarDependent_80DD").value = jsonObject("ITR")("ITR4")("Schedule80DD")("DependentAadhaar")
'End If
'
'Dim filingdate1_80DD
'
' filingdate1_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("Form10IAFilingDate")
'If Trim(filingdate1_80DD) <> "" Then
'  Date10IA_2 = Mid(filingdate1_80DD, 9, 2) & "/" & Mid(filingdate1_80DD, 6, 2) & "/" & Mid(filingdate1_80DD, 1, 4)
'  Sheet14.Range("Form10IA_80DD").value = Date10IA_2
'End If
'
'
'Dim acknum_80DD
'
'acknum_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("Form10IAAckNum")
'
'If Trim(acknum_80DD) <> "" Then
' Sheet14.Range("AcknowledgeNum_80DD").value = Trim(acknum_80DD)
'End If
'
'Dim UDIDnum_80DD
'
'UDIDnum_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("UDIDNum")
'
'If Trim(UDIDnum_80DD) <> "" Then
' Sheet14.Range("UDIDNumber_80DD").value = Trim(UDIDnum_80DD)
'End If
'End If

End Function

'Malli comented
Function ImportSchedule80U_1(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'Dim Date_New As Variant
'
'Set jsonObject = ParseJson(jsonText)
'Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80U")
'
''If node.exists("Schedule80U") Then
'If Trim(jsonObject("ITR")("ITR4")("Schedule80U")("NatureOfDisability")) <> "" Then
'
'If Trim(jsonObject("ITR")("ITR4")("Schedule80U")("NatureOfDisability")) = "1" Then
'     Sheet14.Range("NatureDisability_80U").value = Trim("1-Self with disability")
'   ElseIf Trim(jsonObject("ITR")("ITR4")("Schedule80U")("NatureOfDisability")) = "2" Then
'     Sheet14.Range("NatureDisability_80U").value = Trim("2-Self with severe disability")
'   Else
'     Sheet14.Range("NatureDisability_80U").value = "(Select)"
'End If
'End If
'
'
'Dim filingdate1_80U
' filingdate1_80U = jsonObject("ITR")("ITR4")("Schedule80U")("Form10IAFilingDate")
'
'If Trim(filingdate1_80U) <> "" Then
'
'  Date_New = Mid(filingdate1_80U, 9, 2) & "/" & Mid(filingdate1_80U, 6, 2) & "/" & Mid(filingdate1_80U, 1, 4)
'  Sheet14.Range("DateFiling10IA_80U").value = Date_New
'End If
'
' Dim acknum_80U
'
' acknum_80U = Trim(jsonObject("ITR")("ITR4")("Schedule80U")("Form10IAAckNum"))
'
'
'If Trim(acknum_80U) <> "" Then
'Sheet14.Range("AckNumber10IA_80U").value = UCase(acknum_80U)
'End If
'
'
'Dim UDIDnum_80U
'
'UDIDnum_80U = Trim(jsonObject("ITR")("ITR4")("Schedule80U")("UDIDNum"))
'
'
'If Trim(UDIDnum_80U) <> "" Then
'Sheet14.Range("UDIDNum_80U").value = UCase(UDIDnum_80U)
'End If
'' End If
 
End Function

'Malli
Function ImportSchedule80DD(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Node1, nodeList1 As Object
Dim Nature As Variant
Dim Date10IA_2 As Variant
 
 
Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR4")

If Nodelist.exists("Schedule80DD") Then

Dim NatureOfDisability_80DD
   NatureOfDisability_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("NatureOfDisability")
   If NatureOfDisability_80DD <> "" Then
   If NatureOfDisability_80DD = "1" Then
   Sheet14.Range("NatureDisability_80DD").value = "1-Dependent person with disability"
   ElseIf NatureOfDisability_80DD = "2" Then
   Sheet14.Range("NatureDisability_80DD").value = "2-Dependent person with severe disability"
   Else: Sheet14.Range("NatureDisability_80DD").value = "(Select)"
   End If
   End If
   
'Konda updated on 19-04-2025
'as AY_2025_26 enhancement
Dim TypeOfDisability_80DD
TypeOfDisability_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("TypeOfDisability")
    If TypeOfDisability_80DD <> "" Then

If TypeOfDisability_80DD = "1" Then
   Sheet14.Range("Disability_80DD").value = "(i) autism, cerebral palsy, or multiple disabilities"
   ElseIf TypeOfDisability_80DD = "2" Then
   Sheet14.Range("Disability_80DD").value = "(ii) others"
   
   'AY_25_26_V0.7
'   If TypeOfDisability_80DD = "1" Then
'   Sheet14.Range("Disability_80DD").value = "(i) blindness"
'   ElseIf TypeOfDisability_80DD = "2" Then
'   Sheet14.Range("Disability_80DD").value = "(ii) low vision"
'   ElseIf TypeOfDisability_80DD = "3" Then
'   Sheet14.Range("Disability_80DD").value = "(iii) leprosy-cured"
'   ElseIf TypeOfDisability_80DD = "4" Then
'   Sheet14.Range("Disability_80DD").value = "(iv) hearing impairment"
'   ElseIf TypeOfDisability_80DD = "5" Then
'   Sheet14.Range("Disability_80DD").value = "(v) locomotor disability"
'   ElseIf TypeOfDisability_80DD = "6" Then
'   Sheet14.Range("Disability_80DD").value = "(vi) mental retardation"
'   ElseIf TypeOfDisability_80DD = "7" Then
'   Sheet14.Range("Disability_80DD").value = "(vii) mental illness"
'   ElseIf TypeOfDisability_80DD = "8" Then
'   Sheet14.Range("Disability_80DD").value = "(viii) autism"
'   ElseIf TypeOfDisability_80DD = "9" Then
'   Sheet14.Range("Disability_80DD").value = "(ix) cerebral palsy"
'   ElseIf TypeOfDisability_80DD = "10" Then
'   Sheet14.Range("Disability_80DD").value = "(x) multiple disability"
   Else
   Sheet14.Range("Disability_80DD").value = "(Select)"
   End If
End If

'-----------------------------------

Dim dependentType_80DD
dependentType_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("DependentType")

 
 
 If dependentType_80DD <> "" Then

If dependentType_80DD = "1" Then
   Sheet14.Range("Dependent_80DD").value = "1-Spouse"
   ElseIf dependentType_80DD = "2" Then
   Sheet14.Range("Dependent_80DD").value = "2-Son"
   ElseIf dependentType_80DD = "3" Then
   Sheet14.Range("Dependent_80DD").value = "3-Daughter"
   ElseIf dependentType_80DD = "4" Then
   Sheet14.Range("Dependent_80DD").value = "4-Father"
   ElseIf dependentType_80DD = "5" Then
   Sheet14.Range("Dependent_80DD").value = "5-Mother"
   ElseIf dependentType_80DD = "6" Then
   Sheet14.Range("Dependent_80DD").value = "6-Brother"
   ElseIf dependentType_80DD = "7" Then
   Sheet14.Range("Dependent_80DD").value = "7-Sister"
   ElseIf dependentType_80DD = "8" Then
   Sheet14.Range("Dependent_80DD").value = "8-Member of the HUF (in case of HUF)"
   Else
   Sheet14.Range("Dependent_80DD").value = "(Select)"
   End If
End If
 
 Dim DependentPan_80DD
 DependentPan_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("DependentPan")
If Trim(DependentPan_80DD) <> "" Then
 Sheet14.Range("PanDependent_80DD").value = DependentPan_80DD
End If
 
Dim DependentAadhaar_80DD
DependentAadhaar_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("DependentAadhaar")
 
If Trim(DependentAadhaar_80DD) <> "" Then
Sheet14.Range("AadhaarDependent_80DD").value = DependentAadhaar_80DD
End If
 
' ''AY_25_26_V0.7-------------------
'Dim filingdate1_80DD
'
'filingdate1_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("Form10IAFilingDate")
'If Trim(filingdate1_80DD) <> "" Then
'  Date10IA_2 = Mid(filingdate1_80DD, 9, 2) & "/" & Mid(filingdate1_80DD, 6, 2) & "/" & Mid(filingdate1_80DD, 1, 4)
'  Sheet14.Range("Form10IA_80DD").value = Date10IA_2
'End If
' '------------------------
 
Dim acknum_80DD
 
acknum_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("Form10IAAckNum")
 
If Trim(acknum_80DD) <> "" Then
Sheet14.Range("AcknowledgeNum_80DD").value = Trim(acknum_80DD)
End If


'''AY_25_26_V0.7-------------------
''Konda updated
''as AY_2025_26 enhancement
'Dim AcknlNum11A2_80DD
'
'AcknlNum11A2_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("FormAckNum11A")
'
'If Trim(AcknlNum11A2_80DD) <> "" Then
'Sheet14.Range("AcknowledgeNum11A2_80DD").value = Trim(AcknlNum11A2_80DD)
'End If
''-----------------------

Dim UDIDnum_80DD
 
UDIDnum_80DD = jsonObject("ITR")("ITR4")("Schedule80DD")("UDIDNum")
 
If Trim(UDIDnum_80DD) <> "" Then
Sheet14.Range("UDIDNumber_80DD").value = Trim(UDIDnum_80DD)
End If
End If
 
End Function
 ' T V N Mallikarjun
Function ImportSchedule80U(jsonText As String)

On Error Resume Next

Dim jsonObject, jsonDictionary As Object

Dim node, Nodelist As Object

Dim Date_New As Variant

Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR4")
 
If Nodelist.exists("Schedule80U") Then
 


Dim NatureOfDisability_80U

NatureOfDisability_80U = jsonObject("ITR")("ITR4")("Schedule80U")("NatureOfDisability")
 
If NatureOfDisability_80U <> "" Then

    If Trim(NatureOfDisability_80U) = "1" Then

     Sheet14.Range("NatureDisability_80U").value = Trim("1-Self with disability")

   ElseIf Trim(NatureOfDisability_80U) = "2" Then

     Sheet14.Range("NatureDisability_80U").value = Trim("2-Self with severe disability")

   Else

     Sheet14.Range("NatureDisability_80U").value = "(Select)"

End If

End If
 
'Konda updated on 19-04-2025
'as AY_2025_26 enhancement
Dim TypeOfDisability_80U
TypeOfDisability_80U = jsonObject("ITR")("ITR4")("Schedule80U")("TypeOfDisability")
    If TypeOfDisability_80U <> "" Then

'If TypeOfDisability_80U = "1" Then
'   Sheet14.Range("Disability_80U").value = "(i) blindness"
'   ElseIf TypeOfDisability_80U = "2" Then
'   Sheet14.Range("Disability_80U").value = "(ii) low vision"
'   ElseIf TypeOfDisability_80U = "3" Then
'   Sheet14.Range("Disability_80U").value = "(iii) leprosy-cured"
'   ElseIf TypeOfDisability_80U = "4" Then
'   Sheet14.Range("Disability_80U").value = "(iv) hearing impairment"
'   ElseIf TypeOfDisability_80U = "5" Then
'   Sheet14.Range("Disability_80U").value = "(v) locomotor disability"
'   ElseIf TypeOfDisability_80U = "6" Then
'   Sheet14.Range("Disability_80U").value = "(vi) mental retardation"
'   ElseIf TypeOfDisability_80U = "7" Then
'   Sheet14.Range("Disability_80U").value = "(vii) mental illness"
'   ElseIf TypeOfDisability_80U = "8" Then
'   Sheet14.Range("Disability_80U").value = "(viii) autism"
'   ElseIf TypeOfDisability_80U = "9" Then
'   Sheet14.Range("Disability_80U").value = "(ix) cerebral palsy"
'   ElseIf TypeOfDisability_80U = "10" Then
'   Sheet14.Range("Disability_80U").value = "(x) multiple disability"
   
   If TypeOfDisability_80U = "1" Then
   Sheet14.Range("Disability_80U").value = "(i) autism, cerebral palsy, or multiple disabilities"
   ElseIf TypeOfDisability_80U = "2" Then
   Sheet14.Range("Disability_80U").value = "(ii) others"
   Else
   Sheet14.Range("Disability_80U").value = "(Select)"
   End If
End If

'End-----------------------------------on 19-04-2025
 
' 'AY_25_26_V0.7
'Dim filingdate1_80U
'
'filingdate1_80U = jsonObject("ITR")("ITR4")("Schedule80U")("Form10IAFilingDate")
'
'If Trim(filingdate1_80U) <> "" Then
'
'  Date_New = Mid(filingdate1_80U, 9, 2) & "/" & Mid(filingdate1_80U, 6, 2) & "/" & Mid(filingdate1_80U, 1, 4)
'
'  Sheet14.Range("DateFiling10IA_80U").value = Date_New
'
'End If
'---------------------------


Dim acknum_80U

acknum_80U = Trim(jsonObject("ITR")("ITR4")("Schedule80U")("Form10IAAckNum"))


If Trim(acknum_80U) <> "" Then

Sheet14.Range("AckNumber10IA_80U").value = UCase(acknum_80U)

End If
 
 
''AY_25_26_V0.7
''Konda updated on 19-04-2025
''as AY_2025_26 enhancement
' Dim AcknlNum11A2_80U
'
'AcknlNum11A2_80U = Trim(jsonObject("ITR")("ITR4")("Schedule80U")("FormAckNum11A"))
'
'If Trim(AcknlNum11A2_80U) <> "" Then
'
'Sheet14.Range("AcknowledgeNum11A2_80U").value = UCase(AcknlNum11A2_80U)
'
'End If
''end----------------------on 19-04-2025
 
 
Dim UDIDnum_80U
 
UDIDnum_80U = Trim(jsonObject("ITR")("ITR4")("Schedule80U")("UDIDNum"))


If Trim(UDIDnum_80U) <> "" Then

Sheet14.Range("UDIDNum_80U").value = UCase(UDIDnum_80U)

End If

End If

End Function
'Konda_80C_AY_2025_26-------09/04/2025
Function ImportSchedule_80C(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Dim NaturePayment_80C, Identification_Number_80C, Amount_80C As Variant
Dim TotalXMLRow, rowcount, cnt, Rec80C  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
'konda added on 25-04-2025-----
Dim NodeList80C
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80C = jsonObject("ITR")("ITR4")
 
If NodeList80C.exists("Schedule80C") Then
 '---------------------------------
    
 Set jsonObject = ParseJson(jsonText)
  
   Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80C")("Schedule80CDtls")
     'AY_25_26_V0.7
    'NaturePayment_80C = Sheet19.Range("NaturePayment.80C").Column
    '----------------------
    Identification_Number_80C = Sheet19.Range("Identification_Number.80C").Column
    Amount_80C = Sheet19.Range("Amount.80C").Column
    
   ' TotalExRow = Sheet19.Range("NaturePayment.80C").Rows.count
    TotalExRow = Sheet19.Range("Amount.80C").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
       ''AY_25_26_V0.7
'        If Sheet19.Range("NaturePayment.80C").Locked = False Then
'            Sheet19.Range("NaturePayment.80C").ClearContents
'        End If
        '-------------
        
        If Sheet19.Range("Identification_Number.80C").Locked = False Then
            Sheet19.Range("Identification_Number.80C").ClearContents
        End If
        If Sheet19.Range("Amount.80C").Locked = False Then
            Sheet19.Range("Amount.80C").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80C (TotalDiffRow)
    End If
    
    'rowcount = getRowNo(Sheet19.Range("NaturePayment.80C").name)
    rowcount = getRowNo(Sheet19.Range("Amount.80C").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        'AY_25_26_V0.7
'        If Sheet19.Cells(rowcount, NaturePayment_80C).Locked = False Then
'            Sheet19.Cells(rowcount, NaturePayment_80C).value = node("NatureOfPayment")
'        End If
        '-----------------
        If Sheet19.Cells(rowcount, Amount_80C).Locked = False Then
            Sheet19.Cells(rowcount, Amount_80C).value = node("Amount")
        End If
        
        If Sheet19.Cells(rowcount, Identification_Number_80C).Locked = False Then
            Sheet19.Cells(rowcount, Identification_Number_80C).value = node("IdentificationNo")
        End If
        
        
        cnt = cnt + 1
    Next node
    Rec80C = cnt
End If
End Function
'Konda_80CCC_AY_2025_26-------09/04/2025
Function ImportSchedule_80CCC(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Dim NaturePayment_80CCC, Policy_Document_Number_80CCC, Amount_80CCC As Variant
Dim TotalXMLRow, rowcount, cnt, Rec80CCC  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
'konda added on 25-04-2025-----
Dim NodeList80CCC
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80CCC = jsonObject("ITR")("ITR4")
 
If NodeList80CCC.exists("Schedule80CCC") Then
 '---------------------------------

  
   Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80CCC")("Schedule80CCCDtls")
    
    NaturePayment_80CCC = Sheet19.Range("Name_of_insurer.80CCC").Column
    Policy_Document_Number_80CCC = Sheet19.Range("Policy_Document_Number.80CCC").Column
    Amount_80CCC = Sheet19.Range("Amount.80CCC").Column
    
    
    
    TotalExRow = Sheet19.Range("Name_of_insurer.80CCC").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet19.Range("Name_of_insurer.80CCC").Locked = False Then
            Sheet19.Range("Name_of_insurer.80CCC").ClearContents
        End If
        If Sheet19.Range("Policy_Document_Number.80CCC").Locked = False Then
            Sheet19.Range("Policy_Document_Number.80CCC").ClearContents
        End If
        If Sheet19.Range("Amount.80CCC").Locked = False Then
            Sheet19.Range("Amount.80CCC").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80CCC (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet19.Range("Name_of_insurer.80CCC").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet19.Cells(rowcount, NaturePayment_80CCC).Locked = False Then
            Sheet19.Cells(rowcount, NaturePayment_80CCC).value = node("InsurerName")
        End If
        If Sheet19.Cells(rowcount, Policy_Document_Number_80CCC).Locked = False Then
            Sheet19.Cells(rowcount, Policy_Document_Number_80CCC).value = node("PolicyDocNo")
        End If
        If Sheet19.Cells(rowcount, Amount_80CCC).Locked = False Then
            Sheet19.Cells(rowcount, Amount_80CCC).value = node("Amount")
        End If
        
        cnt = cnt + 1
    Next node
    Rec80CCC = cnt
End If
End Function
'Konda_80E_AY_2025_26-------09/04/2025
Function ImportSchedule_80E(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80E, IFSC_80E, BankName_80E, PANof_80E, LoanAcctNum_80E, LoanDate_80E, TotalLoanAmt_80E, LoanOutStanding_80E, Intrst_80E As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80E, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String
'konda added on 25-04-2025-----
Dim NodeList80E
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80E = jsonObject("ITR")("ITR4")
 
If NodeList80E.exists("Schedule80E") Then
 '---------------------------------
Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80E")("Schedule80EDtls")

    LoanfrmBankOrInstitute_80E = Sheet20.Range("LoanfrmBankOrInstitute.80E").Column
    'Malli--------V0.7
    'IFSC_80E = Sheet20.Range("IFSC.80E").Column
    'PANof_80E = Sheet20.Range("PAN.80E").Column
    '--------------------
    BankName_80E = Sheet20.Range("bankName.80E").Column
    
    LoanAcctNum_80E = Sheet20.Range("loanAccNum.80E").Column
    LoanDate_80E = Sheet20.Range("loanDate.80E").Column
    TotalLoanAmt_80E = Sheet20.Range("loanAmt.80E").Column
    LoanOutStanding_80E = Sheet20.Range("loanOutstanding.80E").Column
    Intrst_80E = Sheet20.Range("Intrst.80E").Column
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80E").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80E").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80E").ClearContents
        End If
        'Malli--------V0.7
'        If Sheet20.Range("IFSC.80E").Locked = False Then
'            Sheet20.Range("IFSC.80E").ClearContents
'        End If
        '-----------------
        If Sheet20.Range("bankName.80E").Locked = False Then
            Sheet20.Range("bankName.80E").ClearContents
        End If
        
        'Malli---V0.7
'        If Sheet20.Range("PAN.80E").Locked = False Then
'            Sheet20.Range("PAN.80E").ClearContents
'        End If
        '--------
        If Sheet20.Range("loanAccNum.80E").Locked = False Then
            Sheet20.Range("loanAccNum.80E").ClearContents
        End If
        If Sheet20.Range("loanDate.80E").Locked = False Then
            Sheet20.Range("loanDate.80E").ClearContents
        End If
        If Sheet20.Range("loanAmt.80E").Locked = False Then
            Sheet20.Range("loanAmt.80E").ClearContents
        End If
        If Sheet20.Range("loanOutstanding.80E").Locked = False Then
            Sheet20.Range("loanOutstanding.80E").ClearContents
        End If
        If Sheet20.Range("Intrst.80E").Locked = False Then
            Sheet20.Range("Intrst.80E").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80E (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet20.Range("LoanfrmBankOrInstitute.80E").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80E).Locked = False Then
            Dim LoanTknFrom_80E
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80E = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80E = "Institution"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80E).value = LoanTknFrom_80E
        End If
        
         'Malli--------V0.7
'        If Sheet20.Cells(rowcount, IFSC_80E).Locked = False Then
'            Sheet20.Cells(rowcount, IFSC_80E).value = node("IFSCCode")
'        End If
        '-----------------
        If Sheet20.Cells(rowcount, BankName_80E).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80E).value = node("BankOrInstnName")
        End If
         'Malli--------V0.7
'        If Sheet20.Cells(rowcount, PANof_80E).Locked = False Then
'            Sheet20.Cells(rowcount, PANof_80E).value = node("PAN")
'        End If
        '--------------------
        If Sheet20.Cells(rowcount, LoanAcctNum_80E).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80E).value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80E).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80E).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80E).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80E).value = node("TotalLoanAmt")
        End If
        If Sheet20.Cells(rowcount, LoanOutStanding_80E).Locked = False Then
            Sheet20.Cells(rowcount, LoanOutStanding_80E).value = node("LoanOutstndngAmt")
        End If
        
        If Sheet20.Cells(rowcount, Intrst_80E).Locked = False Then
            Sheet20.Cells(rowcount, Intrst_80E).value = node("Interest80E")
        End If
        cnt = cnt + 1
    Next node
    Rec80E = cnt
 End If
End Function
'Konda_80EE_AY_2025_26-------09/04/2025
Function ImportSchedule_80EE(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EE, IFSC_80EE, BankName_80EE, PANof_80EE, LoanAcctNum_80EE, LoanDate_80EE, TotalLoanAmt_80EE, LoanOutStanding_80EE, Intrst_80EE As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EE, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

'konda added on 25-04-2025-----
Dim NodeList80EE
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EE = jsonObject("ITR")("ITR4")
 
If NodeList80EE.exists("Schedule80EE") Then
'---------------------------------

''Malli--------V0.7
'    Dim ResHPVal_80EE
'    ResHPVal_80EE = jsonObject("ITR")("ITR4")("Schedule80EE")("ResHPVal")
'
'    If ResHPVal_80EE <> "" Then
'        If Sheet20.Range("ResidentialHP.80E").Locked = False Then
'                Sheet20.Range("ResidentialHP.80E").value = ResHPVal_80EE
'            End If
'    End If
''------------------


Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80EE")("Schedule80EEDtls")

    LoanfrmBankOrInstitute_80EE = Sheet20.Range("LoanfrmBankOrInstitute.80EE").Column
    'Malli--------V0.7
    'IFSC_80EE = Sheet20.Range("IFSC.80EE").Column
    'PANof_80EE = Sheet20.Range("PAN.80EE").Column
    '--------------------
    BankName_80EE = Sheet20.Range("bankName.80EE").Column
    
    LoanAcctNum_80EE = Sheet20.Range("loanAccNum.80EE").Column
    LoanDate_80EE = Sheet20.Range("loanDate.80EE").Column
    TotalLoanAmt_80EE = Sheet20.Range("loanAmt.80EE").Column
    LoanOutStanding_80EE = Sheet20.Range("loanOutstanding.80EE").Column
    
    Intrst_80EE = Sheet20.Range("Intrst.80EE").Column
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80EE").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80EE").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").ClearContents
        End If
        If Sheet20.Range("IFSC.80EE").Locked = False Then
            Sheet20.Range("IFSC.80EE").ClearContents
        End If
        If Sheet20.Range("bankName.80EE").Locked = False Then
            Sheet20.Range("bankName.80EE").ClearContents
        End If
        If Sheet20.Range("PAN.80EE").Locked = False Then
            Sheet20.Range("PAN.80EE").ClearContents
        End If
        If Sheet20.Range("loanAccNum.80EE").Locked = False Then
            Sheet20.Range("loanAccNum.80EE").ClearContents
        End If
        If Sheet20.Range("loanDate.80EE").Locked = False Then
            Sheet20.Range("loanDate.80EE").ClearContents
        End If
        If Sheet20.Range("loanAmt.80EE").Locked = False Then
            Sheet20.Range("loanAmt.80EE").ClearContents
        End If
        If Sheet20.Range("loanOutstanding.80EE").Locked = False Then
            Sheet20.Range("loanOutstanding.80EE").ClearContents
        End If
        
        If Sheet20.Range("Intrst.80EE").Locked = False Then
            Sheet20.Range("Intrst.80EE").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EE (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet20.Range("LoanfrmBankOrInstitute.80EE").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EE).Locked = False Then
            Dim LoanTknFrom_80EE
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80EE = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80EE = "Institution"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EE).value = LoanTknFrom_80EE
        End If
        'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, IFSC_80EE).Locked = False Then
'            Sheet20.Cells(rowcount, IFSC_80EE).value = node("IFSCCode")
'        End If
        '------------------
        If Sheet20.Cells(rowcount, BankName_80EE).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80EE).value = node("BankOrInstnName")
        End If
         'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, PANof_80EE).Locked = False Then
'            Sheet20.Cells(rowcount, PANof_80EE).value = node("PAN")
'        End If
        '--------------
        If Sheet20.Cells(rowcount, LoanAcctNum_80EE).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80EE).value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80EE).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80EE).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80EE).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80EE).value = node("TotalLoanAmt")
        End If
        If Sheet20.Cells(rowcount, LoanOutStanding_80EE).Locked = False Then
            Sheet20.Cells(rowcount, LoanOutStanding_80EE).value = node("LoanOutstndngAmt")
        End If
        
        If Sheet20.Cells(rowcount, Intrst_80EE).Locked = False Then
            Sheet20.Cells(rowcount, Intrst_80EE).value = node("Interest80EE")
        End If
        cnt = cnt + 1
    Next node
    Rec80EE = cnt
End If
End Function
'Konda_80EEA_AY_2025_26-------09/04/2025
Function ImportSchedule_80EEA(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEA, IFSC_80EEA, BankName_80EEA, PANof_80EEA, LoanAcctNum_80EEA, LoanDate_80EEA, TotalLoanAmt_80EEA, LoanOutStanding_80EEA, Intrst_80EEA As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEA, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

'konda added on 25-04-2025-----SIT-92091
Dim NodeList80EEA
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EEA = jsonObject("ITR")("ITR4")
 
If NodeList80EEA.exists("Schedule80EEA") Then
 '---------------------------------
        
Dim PropStmpDtyVal_80EEA
PropStmpDtyVal_80EEA = jsonObject("ITR")("ITR4")("Schedule80EEA")("PropStmpDtyVal")

If PropStmpDtyVal_80EEA <> "" Then
    If Sheet20.Range("Stampduty.80EEA").Locked = False Then
            Sheet20.Range("Stampduty.80EEA").value = PropStmpDtyVal_80EEA
        End If
End If



Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80EEA")("Schedule80EEADtls")

    LoanfrmBankOrInstitute_80EEA = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Column
     'AY_25_26_V0.7
    'IFSC_80EEA = Sheet20.Range("IFSC.80EEA").Column
    'PANof_80EEA = Sheet20.Range("PAN.80EEA").Column
    '-------------------
    BankName_80EEA = Sheet20.Range("bankName.80EEA").Column
    
    LoanAcctNum_80EEA = Sheet20.Range("loanAccNum.80EEA").Column
    LoanDate_80EEA = Sheet20.Range("loanDate.80EEA").Column
    TotalLoanAmt_80EEA = Sheet20.Range("loanAmt.80EEA").Column
    LoanOutStanding_80EEA = Sheet20.Range("loanOutstanding.80EEA").Column
    
    Intrst_80EEA = Sheet20.Range("Intrst.80EEA").Column
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
        End If
        ''AY_25_26_V0.7
'        If Sheet20.Range("IFSC.80EEA").Locked = False Then
'            Sheet20.Range("IFSC.80EEA").ClearContents
'        End If
        '-------------
        If Sheet20.Range("bankName.80EEA").Locked = False Then
            Sheet20.Range("bankName.80EEA").ClearContents
        End If
        'AY_25_26_V0.7
'        If Sheet20.Range("PAN.80EEA").Locked = False Then
'            Sheet20.Range("PAN.80EEA").ClearContents
'        End If
        '-----------------
        If Sheet20.Range("loanAccNum.80EEA").Locked = False Then
            Sheet20.Range("loanAccNum.80EEA").ClearContents
        End If
        If Sheet20.Range("loanDate.80EEA").Locked = False Then
            Sheet20.Range("loanDate.80EEA").ClearContents
        End If
        If Sheet20.Range("loanAmt.80EEA").Locked = False Then
            Sheet20.Range("loanAmt.80EEA").ClearContents
        End If
        If Sheet20.Range("loanOutstanding.80EEA").Locked = False Then
            Sheet20.Range("loanOutstanding.80EEA").ClearContents
        End If
        
        If Sheet20.Range("Intrst.80EEA").Locked = False Then
            Sheet20.Range("Intrst.80EEA").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EEA (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet20.Range("LoanfrmBankOrInstitute.80EEA").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).Locked = False Then
            Dim LoanTknFrom_80EEA
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80EEA = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80EEA = "Institution"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).value = LoanTknFrom_80EEA
        End If
         'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, IFSC_80EEA).Locked = False Then
'            Sheet20.Cells(rowcount, IFSC_80EEA).value = node("IFSCCode")
'        End If
        '-----------
        If Sheet20.Cells(rowcount, BankName_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80EEA).value = node("BankOrInstnName")
        End If
         'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, PANof_80EEA).Locked = False Then
'            Sheet20.Cells(rowcount, PANof_80EEA).value = node("PAN")
'        End If
        '------------
        If Sheet20.Cells(rowcount, LoanAcctNum_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80EEA).value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80EEA).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80EEA).value = node("TotalLoanAmt")
        End If
        If Sheet20.Cells(rowcount, LoanOutStanding_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, LoanOutStanding_80EEA).value = node("LoanOutstndngAmt")
        End If
        
        If Sheet20.Cells(rowcount, Intrst_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, Intrst_80EEA).value = node("Interest80EEA")
        End If
        cnt = cnt + 1
    Next node
    Rec80EEA = cnt
 End If  '-----SIT-92091
End Function
'Konda_80EEB_AY_2025_26-------09/04/2025
Function ImportSchedule_80EEB(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEB, IFSC_80EEB, BankName_80EEB, Vehicle_value_80EEB, VehicleRegNum_80EEB, PANof_80EEB, LoanAcctNum_80EEB, LoanDate_80EEB, TotalLoanAmt_80EEB, LoanOutStanding_80EEB, Intrst_80EEB As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEB, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

'konda added on 25-04-2025-----
Dim NodeList80EEB
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EEB = jsonObject("ITR")("ITR4")
 
If NodeList80EEB.exists("Schedule80EEB") Then
 '---------------------------------
Set Nodelist = jsonObject("ITR")("ITR4")("Schedule80EEB")("Schedule80EEBDtls")

    LoanfrmBankOrInstitute_80EEB = Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Column
     'AY_25_26_V0.7
'    IFSC_80EEB = Sheet20.Range("IFSC.80EEB").Column
'    PANof_80EEB = Sheet20.Range("PAN.80EEB").Column
'    Vehicle_value_80EEB = Sheet20.Range("Vehicle_value.80EEB").Column
    '----------------------
    BankName_80EEB = Sheet20.Range("bankName.80EEB").Column
    
    LoanAcctNum_80EEB = Sheet20.Range("loanAccNum.80EEB").Column
    LoanDate_80EEB = Sheet20.Range("loanDate.80EEB").Column
    TotalLoanAmt_80EEB = Sheet20.Range("loanAmt.80EEB").Column
    LoanOutStanding_80EEB = Sheet20.Range("loanOutstanding.80EEB").Column
   
    VehicleRegNum_80EEB = Sheet20.Range("VehicleRegNum.80EEB").Column
    Intrst_80EEB = Sheet20.Range("Intrst.80EEB").Column
    
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80EEB").ClearContents
        End If
        'AY_25_26_V0.7
'        If Sheet20.Range("IFSC.80EEB").Locked = False Then
'            Sheet20.Range("IFSC.80EEB").ClearContents
'        End If
        '-------------
        If Sheet20.Range("bankName.80EEB").Locked = False Then
            Sheet20.Range("bankName.80EEB").ClearContents
        End If
        ''AY_25_26_V0.7
'        If Sheet20.Range("PAN.80EEB").Locked = False Then
'            Sheet20.Range("PAN.80EEB").ClearContents
'        End If
        '----------
        If Sheet20.Range("loanAccNum.80EEB").Locked = False Then
            Sheet20.Range("loanAccNum.80EEB").ClearContents
        End If
        If Sheet20.Range("loanDate.80EEB").Locked = False Then
            Sheet20.Range("loanDate.80EEB").ClearContents
        End If
        If Sheet20.Range("loanAmt.80EEB").Locked = False Then
            Sheet20.Range("loanAmt.80EEB").ClearContents
        End If
        If Sheet20.Range("loanOutstanding.80EEB").Locked = False Then
            Sheet20.Range("loanOutstanding.80EEB").ClearContents
        End If
        'AY_25_26_V0.7
'        If Sheet20.Range("Vehicle_value.80EEB").Locked = False Then
'            Sheet20.Range("Vehicle_value.80EEB").ClearContents
'        End If
        '-------------
        If Sheet20.Range("VehicleRegNum.80EEB").Locked = False Then
            Sheet20.Range("VehicleRegNum.80EEB").ClearContents
        End If
        If Sheet20.Range("Intrst.80EEB").Locked = False Then
            Sheet20.Range("Intrst.80EEB").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EEB (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet20.Range("LoanfrmBankOrInstitute.80EEB").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).Locked = False Then
            Dim LoanTknFrom_80EEB
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80EEB = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80EEB = "Institution"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).value = LoanTknFrom_80EEB
        End If
         'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, IFSC_80EEB).Locked = False Then
'            Sheet20.Cells(rowcount, IFSC_80EEB).value = node("IFSCCode")
'        End If
        '-------------
        If Sheet20.Cells(rowcount, BankName_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80EEB).value = node("BankOrInstnName")
        End If
         'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, PANof_80EEB).Locked = False Then
'            Sheet20.Cells(rowcount, PANof_80EEB).value = node("PAN")
'        End If
        '------------
        If Sheet20.Cells(rowcount, LoanAcctNum_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80EEB).value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = Trim(DD & "/" & MM & "/" & YYYY)
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80EEB).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80EEB).value = node("TotalLoanAmt")
        End If
        If Sheet20.Cells(rowcount, LoanOutStanding_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, LoanOutStanding_80EEB).value = node("LoanOutstndngAmt")
        End If
         'AY_25_26_V0.7
'        If Sheet20.Cells(rowcount, Vehicle_value_80EEB).Locked = False Then
'            Sheet20.Cells(rowcount, Vehicle_value_80EEB).value = node("VehicleValue")
'        End If
        '-------------------
        If Sheet20.Cells(rowcount, VehicleRegNum_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, VehicleRegNum_80EEB).value = node("VehicleRegNo")
        End If
        If Sheet20.Cells(rowcount, Intrst_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, Intrst_80EEB).value = node("Interest80EEB")
        End If
        cnt = cnt + 1
    Next node
    Rec80EEB = cnt
End If
End Function
'Konda commented on 29-01-2026=========================
''Konda_Int24B_AY_2025_26-------09/04/2025
'Function ImportSchedule_Int24B(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'Dim LoanfrmBankOrInstitute_24B, IFSC_24B, BankName_24B, PANof_24B, LoanAcctNum_24B, LoanDate_24B, TotalLoanAmt_24B, LoanOutStanding_24B, Intrst_24B As Variant
'Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecInt24B, rowcount, cnt As Long
'Dim strDate As String
'Dim YYYY, MM, DD As String
'
''konda added on 25-04-2025-----
'Dim NodeListInt24B
'
'Set jsonObject = ParseJson(jsonText)
'Set NodeListInt24B = jsonObject("ITR")("ITR4")
'
'If NodeListInt24B.exists("ScheduleUs24B") Then
' '---------------------------------
'Set Nodelist = jsonObject("ITR")("ITR4")("ScheduleUs24B")("ScheduleUs24BDtls")
'
'    LoanfrmBankOrInstitute_24B = Sheet21.Range("LoanfrmBankOrInstitute.24b").Column
'    'AY_25_26_V0.7
''    IFSC_24B = Sheet21.Range("IFSC.24b").Column
''     PANof_24B = Sheet21.Range("PAN.24b").Column
'    '-------------------------------
'    BankName_24B = Sheet21.Range("bankName.24b").Column
'
'    LoanAcctNum_24B = Sheet21.Range("loanAccNum.24b").Column
'    LoanDate_24B = Sheet21.Range("loanDate.24b").Column
'    TotalLoanAmt_24B = Sheet21.Range("loanAmt.24b").Column
'    LoanOutStanding_24B = Sheet21.Range("loanOutstanding.24b").Column
'    Intrst_24B = Sheet21.Range("Intrst.24b").Column
'
'    TotalExRow = Sheet21.Range("LoanfrmBankOrInstitute.24b").Rows.count
'
'    TotalXMLRow = Nodelist.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'        If Sheet21.Range("LoanfrmBankOrInstitute.24b").Locked = False Then
'            Sheet21.Range("LoanfrmBankOrInstitute.24b").ClearContents
'        End If
'        'AY_25_26_V0.7
''        If Sheet21.Range("IFSC.24b").Locked = False Then
''            Sheet21.Range("IFSC.24b").ClearContents
''        End If
'        '-----------------------
'        If Sheet21.Range("bankName.24b").Locked = False Then
'            Sheet21.Range("bankName.24b").ClearContents
'        End If
'        ''AY_25_26_V0.7
''        If Sheet21.Range("PAN.24b").Locked = False Then
''            Sheet21.Range("PAN.24b").ClearContents
''        End If
'        '----------------
'        If Sheet21.Range("loanAccNum.24b").Locked = False Then
'            Sheet21.Range("loanAccNum.24b").ClearContents
'        End If
'        If Sheet21.Range("loanDate.24b").Locked = False Then
'            Sheet21.Range("loanDate.24b").ClearContents
'        End If
'        If Sheet21.Range("loanAmt.24b").Locked = False Then
'            Sheet21.Range("loanAmt.24b").ClearContents
'        End If
'        If Sheet21.Range("loanOutstanding.24b").Locked = False Then
'            Sheet21.Range("loanOutstanding.24b").ClearContents
'        End If
'        If Sheet21.Range("Intrst.24b").Locked = False Then
'            Sheet21.Range("Intrst.24b").ClearContents
'        End If
'
'    End If
'
'    If (TotalDiffRow > 0) Then
'     AddDiffRows_Int24B (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet21.Range("LoanfrmBankOrInstitute.24b").name)
'    rowcount = rowcount - 1
'    cnt = 0
'    For Each node In Nodelist
'        rowcount = rowcount + 1
'        If Sheet21.Cells(rowcount, LoanfrmBankOrInstitute_24B).Locked = False Then
'            Dim LoanTknFrom_24B
'            If node("LoanTknFrom") = "B" Then
'                LoanTknFrom_24B = "Bank"
'            ElseIf node("LoanTknFrom") = "I" Then
'                LoanTknFrom_24B = "Other than bank"
'            End If
'            Application.EnableEvents = True
'            Sheet21.Cells(rowcount, LoanfrmBankOrInstitute_24B).value = LoanTknFrom_24B
'        End If
'        'AY_25_26_V0.7
''        If Sheet21.Cells(rowcount, IFSC_24B).Locked = False Then
''            Sheet21.Cells(rowcount, IFSC_24B).value = node("IFSCCode")
''        End If
'        '---------------
'        If Sheet21.Cells(rowcount, BankName_24B).Locked = False Then
'            Sheet21.Cells(rowcount, BankName_24B).value = node("BankOrInstnName")
'        End If
'        'AY_25_26_V0.7
''        If Sheet21.Cells(rowcount, PANof_24B).Locked = False Then
''            Sheet21.Cells(rowcount, PANof_24B).value = node("PAN")
''        End If
'        '-------------------
'        If Sheet21.Cells(rowcount, LoanAcctNum_24B).Locked = False Then
'            Sheet21.Cells(rowcount, LoanAcctNum_24B).value = node("LoanAccNoOfBankOrInstnRefNo")
'        End If
'
'                strDate = node("DateofLoan")
'            If strDate <> "" Then
'                YYYY = Mid(strDate, 1, 4)
'                MM = Mid(strDate, 6, 2)
'                DD = Mid(strDate, 9, 2)
'                strDate = Trim(DD & "/" & MM & "/" & YYYY)
'            End If
'
'        If Sheet21.Cells(rowcount, LoanDate_24B).Locked = False Then
'            Sheet21.Cells(rowcount, LoanDate_24B).value = strDate
'            strDate = ""
'        End If
'        If Sheet21.Cells(rowcount, TotalLoanAmt_24B).Locked = False Then
'            Sheet21.Cells(rowcount, TotalLoanAmt_24B).value = node("TotalLoanAmt")
'        End If
'        If Sheet21.Cells(rowcount, LoanOutStanding_24B).Locked = False Then
'            Sheet21.Cells(rowcount, LoanOutStanding_24B).value = node("LoanOutstndngAmt")
'        End If
'
'        If Sheet21.Cells(rowcount, Intrst_24B).Locked = False Then
'            Sheet21.Cells(rowcount, Intrst_24B).value = node("InterestUs24B")
'        End If
'        cnt = cnt + 1
'    Next node
'    RecInt24B = cnt
' End If
'End Function
'==================================================


'Konda----------10/04/2025
Sub AddDiffRows_80C(DiffRows As Long)
    'AY_25_26_V0.7
    'DefinedgridNameRange = "NaturePayment.80C||Identification_Number.80C||Amount.80C"
    DefinedgridNameRange = "Identification_Number.80C||Amount.80C"
    '-----------------------
    Sheet19.Activate
    'searchLastRow ("NaturePayment.80C")
    searchLastRow ("Amount.80C")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

'Konda----------10/04/2025
Sub AddDiffRows_80CCC(DiffRows As Long)
    DefinedgridNameRange = "Name_of_insurer.80CCC||Policy_Document_Number.80CCC||Amount.80CCC"
    Sheet19.Activate
    searchLastRow ("Name_of_insurer.80CCC")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda----------10/04/2025
Sub AddDiffRows_80E(DiffRows As Long)
    'Malli--------V0.7
   ' DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||IFSC.80E||bankName.80E||PAN.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
    DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||bankName.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
    '------------------------------
    Sheet20.Activate
    searchLastRow ("LoanfrmBankOrInstitute.80E")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub


'Konda----------10/04/2025


Sub AddDiffRows_80EE(DiffRows As Long)

   ' DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||IFSC.80EE||bankName.80EE||PAN.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE"  'Combination_80EE
'AY_25_26_V0.7
     'DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||IFSC.80EE||bankName.80EE||PAN.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
        DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||bankName.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
'--------------------
    Sheet20.Activate

    searchLastRow ("LoanfrmBankOrInstitute.80EE")

    insertRowUnderSectionWithFormula (DiffRows)

End Sub


''Konda----------10/04/2025


Sub AddDiffRows_80EEA(DiffRows As Long)

     'DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||IFSC.80EEA||bankName.80EEA||PAN.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA"
'AY_25_26_V0.7
   ' DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||IFSC.80EEA||bankName.80EEA||PAN.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
      DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||bankName.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
'--------------------------
    Sheet20.Activate

    searchLastRow ("LoanfrmBankOrInstitute.80EEA")

    insertRowUnderSectionWithFormula (DiffRows)

End Sub



'Konda----------10/04/2025
Sub AddDiffRows_80EEB(DiffRows As Long)
    'AY_25_26_V0.7
   ' DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||IFSC.80EEB||bankName.80EEB||PAN.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||Vehicle_value.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||bankName.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    '------------------
    
    Sheet20.Activate
    searchLastRow ("LoanfrmBankOrInstitute.80EEB")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda commented on 29-01-2026
'Sub AddDiffRows_Int24B(DiffRows As Long)
''AY_25_26_V0.7
'    'DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||IFSC.24b||bankName.24b||PAN.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
'    DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
'    '------------------------
'    Sheet21.Activate
'    searchLastRow ("LoanfrmBankOrInstitute.24b")
'    insertRowUnderSectionWithFormula (DiffRows)
'End Sub
'Konda----------19/04/2024
Sub AddDiffRows_Sec80DSelfFamHIDtls(DiffRows As Long)
'AY_25_26_V0.7
    'DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||ReceiptNumA1.80D||AmtA1.80D"
    DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||AmtA1.80D"
    '-----------------------
    Sheet16.Activate
    searchLastRow ("NameInsurerA1.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Sec80DSelfFamSrCtznHIDtls
'Konda----------19/04/2024
Sub AddDiffRows_Sec80DSelfFamSrCtznHIDtls(DiffRows As Long)
   'AY_25_26_V0.7
    'DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||ReceiptNumB1.80D||AmtB1.80D"
     DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||AmtB1.80D"
     '------------------
    Sheet16.Activate
    searchLastRow ("NameInsurerB1.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Sec80DParentsHIDtls
'Konda----------19/04/2024
Sub AddDiffRows_Sec80DParentsHIDtls(DiffRows As Long)
   'AY_25_26_V0.7
    'DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||ReceiptNumA2.80D||AmtA2.80D"
    DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||AmtA2.80D"
    '------------------
    Sheet16.Activate
    searchLastRow ("NameInsurerA2.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Sec80DParentsSrCtznHIDtls
'Konda----------19/04/2024
Sub AddDiffRows_Sec80DParentsSrCtznHIDtls(DiffRows As Long)
''AY_25_26_V0.7
   ' DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||ReceiptNumB2.80D||AmtB2.80D"
     DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||AmtB2.80D"
     '-------------------------
    Sheet16.Activate
    searchLastRow ("NameInsurerB2.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda_EA10_13A__AY_2025_26-------08-05-2025
Function ImportScheduleEA10_13A(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Nodelist As Object

Set jsonObject = ParseJson(jsonText)
 
Set Nodelist = jsonObject("ITR")("ITR4")
 
If Nodelist.exists("ScheduleEA10_13A") Then
 
Dim Placeofwork_EA10

Placeofwork_EA10 = jsonObject("ITR")("ITR4")("ScheduleEA10_13A")("Placeofwork")

If Placeofwork_EA10 <> "" Then


If Trim(Placeofwork_EA10) = "1" Then
   Sheet15.Range("Sch10of13A_PlaceofWrk").value = Trim("1. Metro")
   
   ElseIf Trim(Placeofwork_EA10) = "2" Then
   Sheet15.Range("Sch10of13A_PlaceofWrk").value = Trim("2. Non-Metro")
   
   Else: Sheet15.Range("Sch10of13A_PlaceofWrk").value = "(Select)"
   
End If
End If


Dim ActlHRARecv_EA10

ActlHRARecv_EA10 = Trim(jsonObject("ITR")("ITR4")("ScheduleEA10_13A")("ActlHRARecv"))


If Trim(ActlHRARecv_EA10) <> "" Then

Sheet15.Range("Sch10of13A_ActlHRArecivedA").value = UCase(ActlHRARecv_EA10)

End If


Dim ActlRentPaid_EA10

ActlRentPaid_EA10 = Trim(jsonObject("ITR")("ITR4")("ScheduleEA10_13A")("ActlRentPaid"))


If Trim(ActlRentPaid_EA10) <> "" Then

Sheet15.Range("Sch10of13A_ActlRentpaid").value = UCase(ActlRentPaid_EA10)

End If


Dim BasicSalary_EA10

BasicSalary_EA10 = Trim(jsonObject("ITR")("ITR4")("ScheduleEA10_13A")("BasicSalary"))

If Sheet15.Range("Sch10of13A_BasicSalary").Locked = False Then
Sheet15.Range("Sch10of13A_BasicSalary").value = UCase(BasicSalary_EA10)

End If

Dim DearnessAllwnc_EA10

DearnessAllwnc_EA10 = Trim(jsonObject("ITR")("ITR4")("ScheduleEA10_13A")("DearnessAllwnc"))

If Sheet15.Range("Sch10of13A_DearAllowance").Locked = False Then
Sheet15.Range("Sch10of13A_DearAllowance").value = UCase(DearnessAllwnc_EA10)
End If

 'Malli-----14/05/2025
Dim Sal40Or50Per_EA10
  Sal40Or50Per_EA10 = jsonObject("ITR")("ITR4")("ScheduleEA10_13A")("Sal40Or50Per")
    If Sheet15.Range("Sch10of13A_50Por40Pofsalary").Locked = False And Sal40Or50Per_EA10 <> "" Then
    Sheet15.Range("Sch10of13A_50Por40Pofsalary").value = Sal40Or50Per_EA10
    End If
'----------------------

End If
End Function

'Konda updated on 22-01-2026
Function ImportScheduleHP(ByVal jsonObject As Object)
On Error Resume Next
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2, init As Object
    Dim node, Node1, Node2 As Object
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant

    Dim rowcount, cnt As Long
    Sheet22.Activate
Set init = jsonObject("ITR")("ITR4")("IncomeDeductions")
If init.exists("PropertyDetails") Then

    Set Nodelist = jsonObject("ITR")("ITR4")("IncomeDeductions")("PropertyDetails")

    TotalExRow = Sheet22.Range("PropertySectionCOunt").value
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    

    Dim i As Long
    For i = 1 To TotalDiffRow
        Sheet22.Activate
        'AddBlockCall_hprptfrm
    Next

    rowcount = 0
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        'Co-Owner
        iTotalCoRow = Sheet22.Range("HP.Co.Name" & rowcount).Rows.count
         
        sTotalCoRow = node("CoOwners").count
        TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
        If TotalDiffCoRow > 0 Then
            Sheet22.Activate
            AddPropertyCoOWners (TotalDiffCoRow)
        End If
        
        'Tenants
        iTotalTenRow = Sheet22.Range("HP.NameofTenant" & rowcount).Rows.count
        sTotalTenRow = node("TenantDetails").count
        TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
        If TotalDiffTenRow > 0 Then
            Sheet22.Activate
            AddPropertyTenant (TotalDiffTenRow)
        End If
            
            If Sheet22.Range("HP.AddrDetail" & rowcount).Locked = False Then
                Sheet22.Range("HP.AddrDetail" & rowcount).value = UCase(node("AddressDetailWithZipCode")("AddrDetail"))
            End If
            If Sheet22.Range("HP.CityOrTownOrDistrict" & rowcount).Locked = False Then
                Sheet22.Range("HP.CityOrTownOrDistrict" & rowcount).value = UCase(node("AddressDetailWithZipCode")("CityOrTownOrDistrict"))
            End If
            
            iState = UCase(node("AddressDetailWithZipCode")("StateCode"))
            sState = Findtext(iState, "State")
            Sheet22.Range("HP.StateCode" & rowcount).value = sState
            iCountry = UCase(node("AddressDetailWithZipCode")("CountryCode"))
            sCountry = Findtext(iCountry, "Country")
            
            If Sheet22.Range("HP.CountryCode" & rowcount).Locked = False Then
                Sheet22.Range("HP.CountryCode" & rowcount).value = sCountry
            End If
            If Sheet22.Range("HP.PinCode" & rowcount).Locked = False Then
                Sheet22.Range("HP.PinCode" & rowcount).value = UCase(node("AddressDetailWithZipCode")("PinCode"))
            End If
            If Sheet22.Range("HP.ZipCode" & rowcount).Locked = False Then
                Sheet22.Range("HP.ZipCode" & rowcount).value = UCase(node("AddressDetailWithZipCode")("ZipCode"))
            End If
            
            Dim OwnerProperty_HP
            
            OwnerProperty_HP = UCase(node("PropertyOwner"))
            
            If OwnerProperty_HP = "SE" Then
               OwnerProperty_HP = "Self"
            ElseIf OwnerProperty_HP = "MI" Then
               OwnerProperty_HP = "Minor"
            ElseIf OwnerProperty_HP = "SP" Then
               OwnerProperty_HP = "Spouse"
            ElseIf OwnerProperty_HP = "OT" Then
               OwnerProperty_HP = "Others"
            Else
                
            End If
            
            If Sheet22.Range("HP.OwnerProperty" & rowcount).Locked = False Then
                Sheet22.Range("HP.OwnerProperty" & rowcount).value = OwnerProperty_HP
            End If
            
            If Sheet22.Range("HP.OwnerPropertyDescription" & rowcount).Locked = False Then
                Sheet22.Range("HP.OwnerPropertyDescription" & rowcount).value = node("PropertyOwnerOther")
            End If
            
            Dim Co_Ownedflag As Variant
            
            Co_Ownedflag = UCase(node("PropCoOwnedFlg"))
            If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
            Co_Ownedflag = "Yes"
            ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
            Co_Ownedflag = "No"
            Else
            Co_Ownedflag = "(Select)"
            End If
            
            Sheet22.Range("HP.CoOwnedYN" & rowcount).value = Co_Ownedflag
             
            Dim hpshare As Variant
            hpshare = UCase(node("AsseseeShareProperty"))
            
            If Sheet22.Range("HP.SharePercent" & rowcount).Locked = False Then
                Sheet22.Range("HP.SharePercent" & rowcount).value = hpshare
            End If
            
            If node.exists("CoOwners") Then
       
            Set nodeList1 = Nothing
            Set nodeList1 = node("CoOwners")
            cnt = 0
            cnt = getRowNo(Sheet22.Range("HP.Co.Name" & rowcount).name)
            cnt = cnt - 1
            
            For Each Node1 In nodeList1
                    cnt = cnt + 1

                If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Name" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Name" & rowcount).Column).value = UCase(Node1("NameCoOwner"))
                End If
                If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Pan" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Pan" & rowcount).Column).value = UCase(Node1("PAN_CoOwner"))
                End If
                If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Aadhaar" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Aadhaar" & rowcount).Column).value = UCase(Node1("Aadhaar_CoOwner"))
                End If
                If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Share" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Share" & rowcount).Column).value = UCase(Node1("PercentShareProperty"))
                End If
                    
            Next Node1
            End If
            
            iLetOut = UCase(node("ifLetOut"))
            
            If iLetOut = "L" Then
                sLetOut = "Let Out"
            ElseIf iLetOut = "D" Then
                sLetOut = "Deemed Let Out"
            ElseIf iLetOut = "S" Then
                sLetOut = "Self Occupied"
            Else
                sLetOut = "(Select)"
            End If
            
            Sheet22.Range("HP.ifLetOut" & rowcount).value = sLetOut
    
            If node.exists("TenantDetails") Then
            If Not IsEmpty(node("TenantDetails")) Then
                Set Nodelist2 = node("TenantDetails")
                cnt = 0
                cnt = getRowNo(Sheet22.Range("HP.NameofTenant" & rowcount).name)
                cnt = cnt - 1
                For Each Node2 In Nodelist2
                    cnt = cnt + 1
   
                    If Sheet22.Cells(cnt, Sheet22.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node2("NameofTenant"))
                    End If
                    If Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node2("PANofTenant"))
                    End If
                    If Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node2("PANTANofTenant"))
                    End If
                        If Node2.Exsist("AadhaarofTenant") Then
                            Sheet22.Cells(cnt, Sheet22.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node2("AadhaarofTenant"))
                        End If
                Next Node2
            End If
            End If
            
                If Sheet22.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                    Sheet22.Range("HP.AnnualLetableValue" & rowcount).value = UCase(node("Rentdetails")("AnnualLetableValue"))
                End If
                
                If Sheet22.Range("HP.RentNotRealized" & rowcount).Locked = False Then
                    Sheet22.Range("HP.RentNotRealized" & rowcount).value = UCase(node("Rentdetails")("RentNotRealized"))
                End If
                If Sheet22.Range("HP.LocalTaxes" & rowcount).Locked = False Then
                    Sheet22.Range("HP.LocalTaxes" & rowcount).value = UCase(node("Rentdetails")("LocalTaxes"))
                End If
                
            
 Dim Nodelist3, Node3
If node("Rentdetails").exists("Section24B") Then

       
        Dim iTotalCoRow_24b, sTotalCoRow_24b, TotalDiffCoRow_24b
        iTotalCoRow_24b = Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Rows.count
         
      
        sTotalCoRow_24b = node("Rentdetails")("Section24B")("Section24BDtls").count
    
        TotalDiffCoRow_24b = WorksheetFunction.Max((sTotalCoRow_24b - iTotalCoRow_24b), 0)
        If TotalDiffCoRow_24b > 0 Then
            Sheet22.Activate
            AddSection24b (TotalDiffCoRow_24b)
        End If
        
 Set Nodelist3 = node("Rentdetails")("Section24B")("Section24BDtls")
                 cnt = 0
                cnt = getRowNo(Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).name)
                cnt = cnt - 1
                For Each Node3 In Nodelist3
                    cnt = cnt + 1
                If Node3("LoanTknFrom") <> "" And Sheet22.Cells(cnt, Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Locked = False Then
          
                    Dim LoanTknFrom_24B As Variant
                    LoanTknFrom_24B = UCase(Node3("LoanTknFrom"))
                    If LoanTknFrom_24B = UCase("B") Then
                     Sheet22.Cells(cnt, Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).value = "Bank "
                    
              
                    ElseIf LoanTknFrom_24B = UCase("I") Then
               
                     Sheet22.Cells(cnt, Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).value = "Other than Bank"
                    End If
                End If
                
                If Node3("BankOrInstnName") <> "" And Sheet22.Cells(cnt, Sheet22.Range("bankName.24b" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("bankName.24b" & rowcount).Column).value = UCase(Node3("BankOrInstnName"))
                End If
                
                If Node3("LoanAccNoOfBankOrInstnRefNo") <> "" And Sheet22.Cells(cnt, Sheet22.Range("loanAccNum.24b" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("loanAccNum.24b" & rowcount).Column).value = UCase(Node3("LoanAccNoOfBankOrInstnRefNo"))
                End If
                
                If Node3("DateofLoan") <> "" And Sheet22.Cells(cnt, Sheet22.Range("loanDate.24b" & rowcount).Column).Locked = False Then
                    
                    Sheet22.Cells(cnt, Sheet22.Range("loanDate.24b" & rowcount).Column).value = Mid(Node3("DateofLoan"), 9, 2) & "/" & Mid(Node3("DateofLoan"), 6, 2) & "/" & Mid(Node3("DateofLoan"), 1, 4)
                End If
                
                If Node3("TotalLoanAmt") <> "" And Sheet22.Cells(cnt, Sheet22.Range("loanAmt.24b" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("loanAmt.24b" & rowcount).Column).value = UCase(Node3("TotalLoanAmt"))
                End If
                
                If Node3("LoanOutstndngAmt") <> "" And Sheet22.Cells(cnt, Sheet22.Range("loanOutstanding.24b" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("loanOutstanding.24b" & rowcount).Column).value = UCase(Node3("LoanOutstndngAmt"))
                End If
                
                If Node3("InterestUs24B") <> "" And Sheet22.Cells(cnt, Sheet22.Range("Intrst.24b" & rowcount).Column).Locked = False Then
                    Sheet22.Cells(cnt, Sheet22.Range("Intrst.24b" & rowcount).Column).value = UCase(Node3("InterestUs24B"))
                End If
                     
                Next Node3
             
End If
                
                If Sheet22.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).Locked = False Then
                    Sheet22.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).value = UCase(node("Rentdetails")("ArrearsUnrealizedRentRcvd"))
                End If
            cnt = cnt + 1
        Next node
End If
    Set Nodelist = jsonObject("ScheduleHP")
    rowcount = 0
        
    If ActiveWorkbook.Sheets("House Property").Visible = xlSheetVisible Then
    
        For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet22.Range("HP.PassTroughIncome").Locked = False Then
            Sheet22.Range("HP.PassTroughIncome").value = UCase(node("PassThroghIncome"))
        End If
        Next node
    End If

End Function

'added by Konda for AY 2026-27
'start--
Sub AddSection24b(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheets("HP").Activate

'----------------Unlock Password-------------------START---
   
   Sheet22.Unprotect Password:=getmsgstate
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet22.Range("PropertySectionCOunt").value


        mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
        ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.24b" & te)

        If te > 1 Then
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = mdCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next



'----------------Unlock Password-------------------START---
   
   Sheet22.Unprotect Password:=getmsgstate
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
        newfrmsize = Sheet22.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet22.Range("NumRowsPropertyBlock").value = newfrmsize
    Next


endline:

'----------------Lock Password-------------------START---
   Sheet22.Protect Password:=getmsgstate
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'--end


'Malli_AY_2026_27 17/04/2026
Sub AddDiffRows_PRANNum(DiffRows As Long)
    'setDiffTblinfo_80CCCTable_1
    DefinedgridNameRange = "Pran_Sl||pran_new"
    Sheet1.Activate
    searchLastRow ("Pran_Sl")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub









