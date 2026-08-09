Attribute VB_Name = "Part_A_General"
Option Explicit
Public end_AuditInfo2, end_AuditInfo1, end_AuditInfo As Variant
Public rngname_AuditInfo, rngname_AuditInfo1, rngname_AuditInfo2, rngname_NRI As Variant
Public AuditAct_Info1, AuditAct_Info_Others, Date_Info1, AuditAct_Info, AuditSection_Info, AuditDate_Info, Audit_YN, end_AuditInfo_Others, JuridictionOfResidence, Taxpayer_IdentificationNumber As Variant

Public end_NRI1, end_NRI As Variant

Public LEI_Number
Public LEI_Date
Public Refund_BTI


Public end_CompType As Variant
Public end_DirectorComp As Variant

Public Share_TypeofCompany, Director_TypeofCompany As Variant

Public MsgBx_General, MsgBx_General1 As String
Dim Msg_Bx_44ABAccount As String

Public MsgBx_SectionCode As String

Public STDcode As String
Public PhoneNo As String
Public LastName As String
Public PAN As String
Public Flat As String
Public Area As String
Public RoadOrStreet As String

'by sadineni on 30/01/2026 for AY 2026-27
Public FlatNew As String
Public AreaNew As String
Public RoadOrStreetNew As String
Public CityNew As String
Public StateNew As String 'by sadineni on 09/02/2026
Public sCountryNew As String
Public PinCodeNew As String
Public ZipCodeNew As String
'==========================================


Public City As String
Public State As String
Public PinCode As String
Public ZipCode As String
Public MobileNo_1, CountryCode1, CountryCode2 As String
Public MobileNo_2 As String
Public sCountry As String
Public Email_1 As String
Public dob As String
Public doc As String
Public PrevAYBAC As Variant
Public NRISEP As Variant
'Mandatory Fields
Public sReceiptNo As String
Public DOF As String
Public sNoticeNo As String
Public sNoticeDate As String, RepName As String, Repcapacity As String
Public NotifiedUs94AFlg As String, NotifiedUs94AFlg_1 As String, RepPAN As String
Public AsseseeRepFlg As String, AsseseeRepFlg_1 As String, RepAddress As String
Public errmsgAdhaar As Variant
Public errmsgPassport As Variant
Public firstName, middleName, ResidenceName As String

Public ReturnType, ResidentialStatus, ResidentialStatusCondition, DirectorCompany, PartnerFirm, SharesOwner, Portugese, Transaction94A, Accounts44AA, Accounts44AB As String
Public Gender, Email, Status, MobileNo, empcat, ReturnFileSec As String

Public ProvisoFlag As Variant
Public DepositAmountFlag As Variant
Public DepositAmount As Variant
Public AggrigateAmountFlag As Variant
Public AggrigateAmount As Variant
Public AggrigateAmountFlag1 As Variant
Public AggrigateAmount1 As Variant


Public Investment As Variant
Public Claim115H As Variant

Public end_SecCode, end_SecDate, end_SecCodeYN As Long
Public SecnCode_General, SecnCodeYN_General, SecnDate_General As Variant

'Ayush_13/01/2026_V0.2
Public Flat_Sec As Variant
Public residenceName_Sec As Variant
Public roadOrStreet_Sec As Variant
Public Area_Sec As Variant
Public City_Sec As Variant
Public State_Sec As Variant
Public Country_Sec As Variant
Public PinCode_Sec As Variant
Public zipCode_Sec As Variant
Public HASZIP_2 As Variant
Dim errmsgID As Variant
'-----------------------

Public LiableSec92Eflg As Variant
Public LiableSec92Eflgdia As Variant
Public AuditDateSec92E As Variant

Public IsDirectorCompany As Variant
Public end_Director1, end_Director2, end_Director3, end_Director4 As Variant
 Public end_Firm1, end_Firm2 As Variant
Public rngname_Director1 As Variant

Public IsShareOwner As Variant
Public end_Shares1, end_Shares2, end_Shares3, end_Shares4, end_Shares5, end_Shares6, end_Shares7, end_Shares8, end_Shares9, end_Shares10, end_Shares11, end_Shares12, end_Shares13 As Variant
Public rngname_Shares1 As Variant

Public Director_DIN, Director_SharesStatus, Director_PAN, Director_NameofCompany As Variant
Public FirmName, Firm_PAN As Variant
Public Share_Acquired_Count, Share_Acquired_Date, Share_Acquired_FaceValue, Share_Acquired_Issue, Share_Acquired_Cost, Share_CB_Cost, Share_CB_Count, Share_Transferred_Sale, Share_Transferred_Count, Share_OP_Cost, Share_OP_Count, Share_PAN, Share_NameofCompany As Variant

Private Sub Img_Home_Click()
    Sheet35.Activate
End Sub
Sub CmdAddComp_Click()
Dim vRows As Long
Sheets("PART A - General").Activate
EfilingCommon.DefinedgridNameRange = "Sheet1.CompanyName||Sheet1.CompanyType||Sheet1.CompPAN||Sheet1.Shares||Sheet1.DIN"
ActiveCellRange = EfilingCommon.searchLastRow("Sheet1.CompanyName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub CmdAddFirm_Click()
Dim vRows As Long
Sheets("PART A - General").Activate
EfilingCommon.DefinedgridNameRange = "Sheet1.NameOfFirm||Sheet1.FirmPAN"
ActiveCellRange = EfilingCommon.searchLastRow("Sheet1.NameOfFirm")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub CmdAddshare_Click()
Dim vRows As Long
Sheets("PART A - General").Activate
EfilingCommon.DefinedgridNameRange = "Share_NameofCompany||Share_TypeofCompany||Share_PAN||Share_OP_Count||Share_OP_Cost||Share_Acquired_Count||Share_Acquired_Date||Share_Acquired_FaceValue||Share_Acquired_Issue||Share_Acquired_Cost||Share_Transferred_Count||Share_Transferred_Sale||Share_CB_Count||Share_CB_Cost"
ActiveCellRange = EfilingCommon.searchLastRow("Share_NameofCompany")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub CmdAddSectionCode_Click()
Dim vRows As Long
Sheets("PART A - General").Activate
EfilingCommon.DefinedgridNameRange = "sheet1.AuditedSection||sheet1.AuditYN||sheet1.DateOfAudit"
ActiveCellRange = EfilingCommon.searchLastRow("sheet1.AuditedSection")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub CmdAddNRI_Click()
Dim vRows As Long
Sheets("PART A - General").Activate
EfilingCommon.DefinedgridNameRange = "NRI_Juridiction||NRI_Taxpayer"
ActiveCellRange = EfilingCommon.searchLastRow("NRI_Juridiction")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub AddBlockAuditDetails_Click()
    Dim vRows As Long
    Sheets("PART A - General").Activate
    EfilingCommon.DefinedgridNameRange = "Audit.Act||Audit.Others||Audit.Sections||Audit.ACTYN||Audit.Date"
    ActiveCellRange = EfilingCommon.searchLastRow("Audit.Act")
   vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub ValidateSchGeneral_Click()
    Validate_PartA_General
    If MsgBx_General = "" Then fmsgboxStatus "Sheet Part A - General is OK"
End Sub

'Validation for mandatory fields
Sub Validate_PartA_General()
If shtPartA_General Then
    'Sheet1.Activate
    'MsgBox (MsgBx_General)
End If
subProcCaption = "Validating Sheet A"
If Not ValidateSectionCodeGrid Then
Sheet1.Activate
fmsgbox (MsgBx_SectionCode) ', vbOKOnly, "Error(s)!"
CloseMsg
End If

End Sub
Sub CloseMsg()
End
End Sub

'function for Sheet General VAlidation Begins

Function shtPartA_General() As Boolean
    shtPartA_General = True
    
    Dim vbMessgaeCaption As String
    
    MsgBx_General = ""
    'MsgBx_General = "PART A-GEneral" & Chr(10)
    'vbMessgaeCaption = "ITR 2: AY: 2024-25" 'Chandru
     vbMessgaeCaption = "ITR 2: AY: 2026-27"  'Shrutika(27-02-25)
    
    'MsgBx_General = "Part A General " & Chr(10)
    
Dim dateOfFiling As Variant
Dim furnishingDate As Variant
Dim AuditDate As Variant
Dim AuditDateSec92E As Variant

VerificationDate = Dformat(VerificationDate, "yyyy-mm-dd")
VerificationDate = Sheet52.Range("sheet9.Date").value

dateOfFiling = Dformat(dateOfFiling, "yyyy-mm-dd")
dateOfFiling = Range("sheet1.OrigRetFiledDate").value


If Dformat(dateOfFiling, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
'fmsgboxStatus ("Date of filing of original return cannot be after System Date in sheet Part-A General")
fmsgboxStatus ("* Date of filing of original return shall not be beyond system date.") 'Added by Shrutika(20/06/25)
shtPartA_General = False
CloseMsg
End If




furnishingDate = Dformat(furnishingDate, "yyyy-mm-dd")
'furnishingDate = Range("sheet1.AuditReportFurnishDate").value
'If Dformat(furnishingDate, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
''fmsgboxStatus ("Date of furnishing of audit report  in sheet Audit information should not be greater than system date in sheet Part-A General")
'fmsgboxStatus ("Date of furnishing of audit report  in sheet Audit information should not be greater than system date")
'shtPartA_General = False
'CloseMsg
'End If

'AuditDate = Range("sheet1.AuditDate").value
'If Dformat(AuditDate, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
''fmsgboxStatus ("Date of audit report in sheet Audit information should not be greater than system date in sheet Part-A General")
'fmsgboxStatus ("Invalid date: date cannot be greater than date of filing of the return")
'shtPartA_General = False
'CloseMsg
'End If

'AuditDateSec92E = Range("sheet1.AuditDateSec92E").value
'If Dformat(AuditDateSec92E, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
''fmsgboxStatus ("Date of audit report for Sec92E in sheet Audit information should not be greater than system date in sheet Part-A General")
'fmsgboxStatus ("Invalid date: date cannot be greater than date of filing of the return")
'shtPartA_General = False
'CloseMsg
'End If

noticeDate = Dformat(dateOfFiling, "yyyy-mm-dd")
noticeDate = Range("sheet1.NoticeDate").value

'DateFi92cd = Range("Sheet1.Datefiled92cd").value


UniqueNumber = Range("Sheet1.UniqueNo").value
If Dformat(noticeDate, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
fmsgboxStatus ("Date of Notice/Order cannot be after System Date in sheet Part-A General")
shtPartA_General = False
CloseMsg
End If

'Change-6, ANK-102, 25.05.2022
If Not CheckSpecialCharacterSlash(UniqueNumber) Then
fmsgboxStatus ("Unique number can be alphanumeric with special character as / (slash) only.")
shtPartA_General = False
CloseMsg
End If
'---end change---

If Len(UniqueNumber) > 100 Then
fmsgboxStatus ("Unique number cannot exceed 100 characters in Part-A General")
shtPartA_General = False
CloseMsg
End If
 
  
If Sheet1.Range("sheet1.STDcode").value <> "" Then
If Not ValidateSTDcode Then CloseMsg
End If
If Not ValidatePhoneNo Then CloseMsg

If Sheet1.Range("Countrycode2").value <> "" Or Sheet1.Range("sheet1.MobileNoSec").value <> "" Then
If Not ValidateMobileNoSec Then CloseMsg
End If


If Not ChkSpecialCharinNonMandatoryfield Then CloseMsg

If Not chkLastName Then MsgBx_General = MsgBx_General + "* Last name is mandatory.Please enter Last name in Personal Information in schedule Part A General" & Chr(13)
If Not ChkPAN Then MsgBx_General = MsgBx_General + "* Please enter a PAN in schedule Part A General under Personal Information" & Chr(13)
If Not ChkFlat Then MsgBx_General = MsgBx_General + "* Flat/ Door/ Building in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
'Commented by Shrutika
'If Not ChkStatus Then MsgBx_General = MsgBx_General + "* Status in Sheet Personal Information is mandatory in schedule Part A General" & Chr(13)
'Added by Shrutika
 If Not ChkStatus Then MsgBx_General = MsgBx_General + """ Status in Sheet Personal Information  is mandatory in schedule Part A General under Personal Information"""

If Not ChkArea Then MsgBx_General = MsgBx_General + "* Area/ Locality in Sheet: Personal Information is mandatory in schedule Part A General under Personal Information" & Chr(13)
If Not ChkDOB Then MsgBx_General = MsgBx_General + "* Date of Birth/Formation in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
'If Not ChkDOC Then MsgBx_General = MsgBx_General + "* Date of Commencement of Business in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
If Not ChkCity Then MsgBx_General = MsgBx_General + "* Town/City/District in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
If Not ChkState Then MsgBx_General = MsgBx_General + "* State in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
'If Not ChkCountry Then MsgBx_General = MsgBx_General + "* Country in Sheet: Personal Information is mandatory" & Chr(13)

If Not ChkCountry Then MsgBx_General = MsgBx_General + "* ""Country is mandatory in tab: Part A General Information in Part A General.""" & Chr(13)

If Not ChkPincode Then shtPartA_General = False
If Not ChkZipcode Then shtPartA_General = False
'If Not ChkGender Then MsgBx_General = MsgBx_General + "* Sex in Sheet : Part A-General  is Mandatory" & Chr(13)
'If Not ChkEmail Then MsgBx_General = MsgBx_General + "* Email Address-1 in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
 'Added by Shrutika(23/04/2026)V0.3
If Not ChkEmail Then MsgBx_General = MsgBx_General + "* ""Primary Email ID of the taxpayer in Sheet: Personal Information is mandatory in schedule Part A General""" & Chr(13)
'If Not ChkMobileNo Then MsgBx_General = MsgBx_General + "* Mobile number in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
'by sadineni on 09/02/2026

 If Not ChkMobileNo Then MsgBx_General = MsgBx_General + "* ""Primary Mobile number in Sheet: Personal Information is mandatory in Part A General.""" & Chr(13)
'by sadineni on 29/01/2026 for AY 2026-27
'==================================================================================================================

If (Mid(Sheet1.Range("sheet1.AsseseeRepFlg").value, 1, 1) = "Y") Then
  If Not RepChkMobileNo Then CloseMsg
End If
'==================================================================================================================================


'If Not ChkEmpCategory Then MsgBx_General = MsgBx_General + "* Employer Category in Sheet : Part A-General  is Mandatory" & Chr(13)
'If Not ChkReturnFileSec Then MsgBx_General = MsgBx_General + "* Filed u/s in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
'Added by Shrutika(19-12-24)AY-25
If Not ChkReturnFileSec Then MsgBx_General = MsgBx_General + "* ""Filing section is mandatory""" & Chr(13)
'new change

    If Trim(SheetALL.Range("bacValue").value) = "" Or IsEmpty(SheetALL.Range("bacValue").value) Then

        'MsgBx_General = MsgBx_General + "* Are you opting for new tax regime u/s 115BAC? is mandatory in Sheet in schedule Part A General " & Chr(13)
    'Chandru
    'MsgBx_General = MsgBx_General + "* Please select the option whether Are you OPTING OUT of the default new tax regime u/s 115BAC?" & Chr(13)
    'Added by Shrutika(19-12-24)AY-2025
     MsgBx_General = MsgBx_General + " * ""Please select option for current assessment year.""" & Chr(13)

    End If
    
    
  
    
'Changes done by sadineni on 27/03/2026 as per DE V0.2
 'Added by Shrutika- 24/04/2026 V0.3 AY-26-27
'************************************************************************************************************************
If Sheet1.Range("Secondary_Address").value = "No" Then

If Not ChkAreaAlternate Then MsgBx_General = MsgBx_General + "*""Area/ Locality in Sheet: Personal Information is mandatory in schedule Part A General under Personal Information""" & Chr(13)
If Not ChkCityAlternate Then MsgBx_General = MsgBx_General + "*""Town/City/District in Sheet: Personal Information is mandatory in schedule Part A General""" & Chr(13)
If Not ChkStateAlternate Then MsgBx_General = MsgBx_General + "* ""State in Sheet: Personal Information is mandatory in schedule Part A General""" & Chr(13)
If Not ChkCountryAlternate Then MsgBx_General = MsgBx_General + "* ""Country is mandatory in tab: Part A General Information in Part A General." & Chr(13)
If Not ChkPincodeAlternate Then shtPartA_General = False
If Not ChkZipcodeAlternate Then shtPartA_General = False
If Not ChkFlat2 Then MsgBx_General = MsgBx_General + "*""Flat/ Door/ Building in Sheet: Personal Information is mandatory in schedule Part A General""" & Chr(13)

'***************************************************************************************************************************
End If

     'If SheetALL.Range("bacValue").value = 1 Then
        If Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then
         If Trim(Range("sheet1.capacityassessee").value) = "" Or IsEmpty(Range("sheet1.capacityassessee").value) Then
          MsgBx_General = MsgBx_General + "* Please furnish form 10-IE to opt for new tax regime" & Chr(13)
         End If

          If Trim(Range("acknowledgmentnumber").value) = "" Or IsEmpty(Range("acknowledgmentnumber").value) Then
            MsgBx_General = MsgBx_General + "* Acknowledgement Number of Form 10IE is mandatory  for persons having income other than salary in Schedule Part A General " & Chr(13)
         End If

    End If

'new change



If Not ChkSeventhProvisoFlag Then MsgBx_General = MsgBx_General + " * Selection of either of the options for 'Are you filing return of income under Seventh proviso to Section 139(1)' is mandatory in Part A General Information" & Chr(13)

If Not ChkDepositAmountFlag Then MsgBx_General = MsgBx_General + " * Selection of either of the options for 'Have you deposited amount or aggregate of amounts exceeding Rs. 1 Crore in one or more current account during the previous year?' is mandatory in Part A General Information" & Chr(13)

If Not ChkAggrigateAmountFlag Then MsgBx_General = MsgBx_General + " * Selection of either of the options for 'Have you incurred expenditure of an amount or aggregate of amount exceeding Rs. 2 lakhs for travel to a foreign country for yourself or for any other person' is mandatory in Part A General Information" & Chr(13)

If Not ChkAggrigateAmountFlag1 Then MsgBx_General = MsgBx_General + " * Selection of either of the options for 'Have you incurred expenditure of amount or aggregate of amount exceeding Rs. 1 lakh on consumption of electricity during the previous year' is mandatory in Part A General Information" & Chr(13)


If Not ChkDepositAmount Then MsgBx_General = MsgBx_General + " * Amount entered cannot be less than or equal to Rs 1 Crore in Part A General Information" & Chr(13)

If Not ChkAggrigateAmount Then MsgBx_General = MsgBx_General + "*Selection of either of the options for Amount entered cannot be less than or equal to Rs 2 Lakhs in Part A General Information" & Chr(13)

If Not ChkAggrigateAmount1 Then MsgBx_General = MsgBx_General + "*Amount entered cannot be less than or equal to Rs 1 Lakhs in Part A General Information" & Chr(13)


If Mid(Sheet1.Range("sheet1.SeventhProvisoFlag").value, 1, 1) <> "N" Then

If Mid(Sheet1.Range("Sheet1.DepositAmountFlag").value, 1, 1) = "Y" Then

If Range("Sheet1.DepositAmount").value = "" Then
    MsgBx_General = MsgBx_General + "*Amount can not be blank/negative/non numeric/ decimal/zero/less than or equal to 1 crore" & Chr(13)
Else
If Range("Sheet1.DepositAmount").value < 10000000 Or Range("Sheet1.DepositAmount").value = 10000000 Then
MsgBx_General = MsgBx_General + "*Amount entered cannot be less than or equal to Rs 1 Crore" & Chr(13)
End If

End If

End If

If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "Y" Then

 If Range("Sheet1.AggrigateAmount").value = "" Then
   MsgBx_General = MsgBx_General + "* Amount can not be blank/negative/non numeric/ decimal/zero/less than or equal to 2 lakhs" & Chr(13)
Else
If Range("Sheet1.AggrigateAmount").value < 200000 Or Range("Sheet1.AggrigateAmount").value = 200000 Then
MsgBx_General = MsgBx_General + "*Amount entered cannot be less than or equal to Rs 2 Lakhs" & Chr(13)
End If
End If

End If

If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "Y" Then

If Range("Sheet1.AggrigateAmount1").value = "" Then
    MsgBx_General = MsgBx_General + "*Amount can not be blank/negative/non numeric/ decimal/zero/less than or equal to 1 lakh" & Chr(13)
Else
If Range("Sheet1.AggrigateAmount1").value < 100000 Or Range("Sheet1.AggrigateAmount1").value = 100000 Then
MsgBx_General = MsgBx_General + "*Amount entered cannot be less than or equal to Rs 1 Lakh" & Chr(13)
End If
End If

End If
End If

'If Not ChkReturnType Then MsgBx_General = MsgBx_General + "* Filing Type in Sheet : Part A-General  is Mandatory" & Chr(13)
If Not ChkResidentialStatus Then MsgBx_General = MsgBx_General + "* Residential Status in Sheet: Personal Information is mandatory in schedule Part A General" & Chr(13)
If Not ChkResidentialStatusCondition Then MsgBx_General = MsgBx_General + "* Please select condition for residential status in schedule Part A General" & Chr(13)


'Added by Ayush 44
'If Not ChkDirectorComapny Then MsgBx_General = MsgBx_General + "* Please select the dropdown Whether you are Director in a company? in schedule Part A General" & Chr(13)
If Not ChkDirectorComapny Then MsgBx_General = MsgBx_General + "* Please select the dropdown for field Whether you are Director in a company at any time during the previous year? in Part A General" & Chr(13)
'If Not ChkDirectorComapny Then MsgBx_General = MsgBx_General + "* Please select the dropdown Whether you are Director in a company? in schedule Part A General" & Chr(13)
'If Not ChkPartnerFirm Then MsgBx_General = MsgBx_General + "* Please enter the name of the Firm(s) in which you are a partner in schedule Part A General" & Chr(13)
If Not ChkShares Then MsgBx_General = MsgBx_General + "* Please select the dropdown Whether you have held unlisted equity shares at any time during the previous year? in schedule Part A General" & Chr(13)
'If Not ChkAdhaarFlag Then MsgBx_General = MsgBx_General + "* Adhaar Status in Sheet : Part A-General  is Mandatory" & Chr(13)
'If Not ChkAdhaarNo Then MsgBx_General = MsgBx_General + "* Aadhaar Number format is incorrect .Please enter the valid Aadhaar number in schedule Part A General" + errmsgAdhaar & Chr(13)

If Not ChkAdhaarNo Then MsgBx_General = MsgBx_General + "* ""Invalid Aadhaar in Part-A General Information""" + errmsgAdhaar & Chr(13)

'Commented by Shrutika_EN
'If Not ChkAdhaarEnrollID Then MsgBx_General = MsgBx_General + "* Please enter valid Aadhaar Enrollment ID in schedule Part A General" + errmsgAdhaar & Chr(13)
'If Not ChkPassNo Then MsgBx_General = MsgBx_General + "* Passport Number in Sheet : Part A-General " + errmsgPassport & Chr(13)

'If Not ChkTransaction94A Then MsgBx_General = MsgBx_General + "* Selection of Transaction under 94A Act in Sheet : Part A-General  is Mandatory" & Chr(13)
If Not ChkClaim115H Then shtPartA_General = False
'If Not ChkForeirgnexchange Then MsgBx_General = MsgBx_General + "* Please select whether assessee is located in an International Financial Services Centre and derives income solely in convertible foreign exchange in Part A General in schedule Part A General" & Chr(13)

    MsgBx_General1 = ""
    'MsgBx_General1 = "PART A-General" & Chr(10)
If Mid(Sheet1.Range("sheet1.Status"), 1, 1) <> "H" Then
    If Not ChkGovernedbyPortugese Then MsgBx_General1 = MsgBx_General1 + "* Whether governed by Portuguese Civil code is mandatory in schedule Part A General" & Chr(13)
End If
'If Not ChkPE Then MsgBx_General1 = MsgBx_General1 + "* Governed by Portugese in Sheet : Part A-General  is Mandatory" & Chr(13)
'If Not ChkAccounts44AA Then MsgBx_General1 = MsgBx_General1 + "* Select a Option in Accounts under 44AA in Sheet : Part A-General  is Mandatory" & Chr(13)
'If Not ChkAccounts44AB Then MsgBx_General1 = MsgBx_General1 + "* Select an Option from Audit Under 44AB in Sheet : Part A-General  is Mandatory" & Chr(13)

'check in ITR3
'If Not ValidateReturnFileSection Then MsgBx_General1 = MsgBx_General1 + "* Please enter the date of Notice u/s 139(9)/142(1)/148/119(2)(b)/section 139 read with section 92CD and Date of filing of original return and Receipt Number in schedule Part A General" & Chr(13)
 'Added by Shrutika
 If Not ValidateReturnFileSection Then MsgBx_General1 = MsgBx_General1 + """Notice/order/advance pricing agreement date is mandatory in Part A General""" & Chr(13)
'----
    'If Not ValidateNotifiedUs94AFlg Then MsgBx_General1 = MsgBx_General1 + "* Whether any transaction has been made with a person located in a jurisdiction notified u/s 94A of the Act? in sheet : PART A-GENERAL Details  is Mandatory" & Chr(13)
    If Not ValidateAsseseeRepFlg Then MsgBx_General1 = MsgBx_General1 + "* Please select whether return is filed by Representative in schedule Part A General" & Chr(13)
 
  '  If Not ValidateRepName Then MsgBx_General1 = MsgBx_General1 + "* Please enter the Name of representative in schedule Part A General" & Chr(13)
 'Added by Shrutika- 23/04/2026 V0.3
    If Not ValidateRepName Then MsgBx_General1 = MsgBx_General1 + "* ""Please enter the Name of representative assessee""" & Chr(13)
  
'by sadineni on 29/01/2026 for AY 2026-27
  '  If Not ValidateRepEmail Then MsgBx_General1 = MsgBx_General1 + "* ""Email Address is mandatory.""" & Chr(13)
' Added by Shrutika- 23/04/2026 V0.3 AY-26-27
   If Not ValidateRepEmail Then MsgBx_General1 = MsgBx_General1 + "*""Email id of representative assessee is mandatory in schedule Part A General""" & Chr(13)
    
    
'Commented by sadineni on 29/01/2026 for AY 2026-27
'=============================================================================================
' 'changed by Chetan C M on 25/08/2025
' 'start--
''    If Not ValidateRepCapacity Then MsgBx_General1 = MsgBx_General1 + "* Please select the capacity of representative in schedule Part A General" & Chr(13)
'     If Not ValidateRepCapacity Then MsgBx_General1 = MsgBx_General1 + "* Please select the capacity of representative assessee  in Part-A General Information. However, same will be taken from the e-filing profile" & Chr(13)
' '--end
'    If Not ValidateRepAddress Then MsgBx_General1 = MsgBx_General1 + "* Please enter the Address of representative in schedule Part A General" & Chr(13)
'    'If Not ValidateRepPAN Then MsgBx_General1 = MsgBx_General1 + "* Please enter the PAN of the representative in schedule Part A General" & Chr(13)
''Added by Shrutika(20-12-24)AY-2025
'If Not ValidateRepPAN Then MsgBx_General1 = MsgBx_General1 + "* Please enter the PAN of the representative assessee." & Chr(13)
'================================================================================================

' IN case of Non Resident
    'If Not ValidatePermanebtEstablishment Then MsgBx_General1 = MsgBx_General1 + "* Whether there is any permanent establishment (PE) in India? in schedule Part A General" & Chr(13)
    If Not ValidateSelection44AA Then MsgBx_General1 = MsgBx_General1 + "* Please Select whether you are liable to maintain accounts as per Section 44AA " & Chr(13)
    If Not ValidateSelection44AE Then MsgBx_General1 = MsgBx_General1 + "*Please select Whether assesee is declaring income under section 44AE/44B/44BB/44BBA in schedule Part A General Information" & Chr(13)
    If Not ValidateSec5cflg Then MsgBx_General1 = MsgBx_General1 + "*Please select Whether during the year Total sales/turnover/gross receipts of business exceeds Rs. 1 crore but does not exceeds Rs. 10 Crore?  in schedule Part A General Information" & Chr(13)
    If Not ValidateSecAggrigateflg Then MsgBx_General1 = MsgBx_General1 + "*Please select If yes is selected at a2i,  whether aggregate of all amounts received including amount received for sales, turnover or gross receipts during the previous year, in cash & non-a/c payee cheque/DD, does not exceed five per cent of said amount  in schedule Part A General Information" & Chr(13)
    If Not ValidateSecAggrigateflg1 Then MsgBx_General1 = MsgBx_General1 + "* Please select If yes is selected at a2i , whether aggregate of all payments made including amount incurred for expenditure, in cash & non-a/c payee cheque/DD, during the previous year does not exceed five per cent of the said payment ? in schedule Part A General Information " & Chr(13)
    
    
'    If Not VAlidateSelect44Ab Then MsgBx_General1 = MsgBx_General1 + "* Please select whether you are liable for audit under section 44AB in schedule Part A General" & Chr(13)
    
'    Validate44AB_accountant

    'Validate92ESec
    ReturntypCheck

'    If Not ValidateAct11() Then
'
'    End If
    
    If Len(Sheet1.Range("NRI_TotPeriodofStay_PreviousYear").value) > 3 Then
    MsgBx_General = MsgBx_General + "* Total period of stay in India during the previous year (in days)cannot be more than 3 digits in sheet PART_A_General" & Chr(13)
    End If
    
    If Len(Sheet1.Range("NRI_TotPeriodofStay_PrecedingYear").value) > 4 Then
    MsgBx_General = MsgBx_General + "* Total period of stay in India during the 4 preceding years (in days)cannot be more than 4 digits in sheet PART_A_General" & Chr(13)
    End If
    
    setTblinfo_NRI
    If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1").value, 1, 3) = "NRI") Then
   If Not end_NRI > 0 Then
'changed by Chetan C M on 25/08/2025
'start--
    'MsgBx_General = MsgBx_General + "* At least one Jurisdiction of residence is mandatory as you selected Non-resident Individual in schedule Part A General" & Chr(13)
    MsgBx_General = MsgBx_General + "* At least one country or territory is mandatory in Part A General as you selected Non-resident Individual" & Chr(13)
'--end
   End If
   
    
    
    
Validate_Juridiction
Validate_Taxpayer
setTblinfo_NRI
  setTblinfo_NRI1
    If (end_NRI1 <> end_NRI) Then
    MsgBx_General = MsgBx_General + "* Enter All Mandatory In Table jurisdiction(s) of residence during the previous year in schedule Part A General" & Chr(13)
  End If
  End If
  
  ValidateDirectorTable
'  ValidateFirmTable

'PAG_LEI_Number 2024-25 Bindu
  ValidateLEI

  ValidateSharesTable
  PrevAYBAC = Range("OptingNewTaxRegime").value
    
  'If (PrevAYBAC = "" Or UCase(PrevAYBAC) = "(SELECT)") And SheetALL.Range("oldbacValue").value > 0 Then MsgBx_General = MsgBx_General & "* Selection of either of the options for 'Are you opting for new tax regime u/s 115BAC ?' is mandatory" & Chr(13)
  'If (PrevAYBAC = "" Or UCase(PrevAYBAC) = "(SELECT)") And SheetALL.Range("oldbacValue").value > 0 Then MsgBx_General = MsgBx_General & "* Please select an option at 'Option for current assessment year'" & Chr(13)
  If (SheetALL.Range("oldbacValue").value = "" Or SheetALL.Range("oldbacValue").value = 0) Then
  MsgBx_General = MsgBx_General & "* Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in  AY 2021-22 ?' is mandatory" & Chr(13)
  End If
    
  'NRISEP = Range("sheet1.NRI_SEP").value
  'If NRISEP = "" Or UCase(NRISEP) = "(SELECT)" Then MsgBx_General = MsgBx_General & "* Please select Significant economic presence (SEP)  in India " & Chr(13)
  
'  If Sheet1.Range("sheet1.NRI_SEP").value = "Yes" And Sheet1.Range("sheet1.NRI_SEP_AggrPayment").value = "" Then
'
'    'fmsgbox "*Please enter aggregate of payments arising from the transaction or transactions during the previous year" & Chr(13)
'    fmsgbox "Please enter aggregate of payments arising from the transaction or transactions during the previous year as referred in Explanation 2A(a) to Section 9(1)(i)" & Chr(13)
'    shtPartA_General = False
'    CloseMsg
'  End If
'
'  If Sheet1.Range("sheet1.NRI_SEP").value = "Yes" And Sheet1.Range("sheet1.NRI_SEP_Usr").value = "" Then
'    fmsgbox "*Please enter number of users in India as referred in Explanation 2A(b) to Section 9(1)(i)" & Chr(13)
'    shtPartA_General = False
'    CloseMsg
'  End If
  
  If Sheet1.Range("sheet1.SeventhProvisoFlag").value = "Yes" Then
  If Sheet1.Range("Sheet1.DepositAmountFlag").value = "Yes" Or Sheet1.Range("Sheet1.AggrigateAmountFlag").value = "Yes" Or Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = "Yes" Or Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
  Else
    fmsgbox "*Select Aleast one option Yes at seventh proviso to section 139(1)  " & Chr(13)
    shtPartA_General = False
    CloseMsg
  End If
  End If
  
  If Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
  'If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "Yes" Or Sheet1.Range("clauseiv7provisio139iFlg_2").value = "Yes" Or Sheet1.Range("clauseiv7provisio139iFlg_3").value = "Yes" Or Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" Then
  If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "Yes" Or Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" Then
  Else
    fmsgbox "*Select Aleast one option Yes at other conditions  prescribed under clause (iv) of seventh proviso to section 139(1) " & Chr(13)
    shtPartA_General = False
    CloseMsg
  End If
  End If
  
  If Sheet1.Range("clauseiv7provisio139iFlg").value = "(Select)" Then
    fmsgbox "*Selection of either of the options for 'Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)' is mandatory in Part A General Information in Part A General Information" & Chr(13)
    shtPartA_General = False
    CloseMsg
  End If
  
  
  If Sheet1.Range("Sheet1.FII").value = "(Select)" Or Sheet1.Range("Sheet1.FII").value = "" Then
  'Konda updated on 07-05-2026--SIT-119151
'    fmsgbox "* Please select whether you are an FII" & Chr(13)
    fmsgbox "* Please select whether you are an FPI" & Chr(13)
    
    shtPartA_General = False
    CloseMsg
  End If
  
       'Added by Shrutika- 23/04/2026 V0.3 AY-26-27

If Sheet1.Range("Secondary_Address").value = "(Select)" Or Sheet1.Range("Secondary_Address").value = "" Then
fmsgboxStatus ("* ""Kindly select the appropriate option for secondary address in Schedule Part A General""") & Chr(13)
shtPartA_General = False
CloseMsg
End If
  
  
  If Sheet1.Range("Sheet1.FII").value = "Yes" Then 'dpk1601
  If Sheet1.Range("Sheet1.sebi").value = "" Then
    fmsgbox "* Please enter the SEBI Registration Number" & Chr(13)
    shtPartA_General = False
    CloseMsg
  End If
  End If 'dpk1601
  validatesheetSebi
  
  
  
  
  If Sheet1.Range("clauseiv7provisio139iFlg").value <> "No" And Sheet1.Range("clauseiv7provisio139iFlg").value <> "(Select)" Then
'    If Sheet1.Range("clauseiv7provisio139iFlg_1").value <> "No" And Sheet1.Range("clauseiv7provisio139iFlg_2").value <> "No" And Sheet1.Range("clauseiv7provisio139iFlg_3").value <> "No" And Sheet1.Range("clauseiv7provisio139iFlg_4").value <> "No" Then
'        fmsgbox "* You have selected ""Yes"" for ""Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)"", hence please enter respective amount in Sheet Income details." & Chr(13)
'        shtPartA_General = False
'    CloseMsg
'    End If
'
'    If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_2").value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_3").value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_4").value = "No" Then
'        fmsgbox "* You have selected ""Yes"" for ""Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)"", hence please enter respective amount in Sheet Income details." & Chr(13)
'        shtPartA_General = False
'    CloseMsg
'    End If
    
    If Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
'        If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "(Select)" Then
'            fmsgbox "*Selection of either of the options for 'total sales, turnover or gross receipts, as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or ' is mandatory in Part A General Information" & Chr(13)
'            shtPartA_General = False
'            CloseMsg
'        End If
'        If Sheet1.Range("clauseiv7provisio139iFlg_2").value = "(Select)" Then
'            fmsgbox "*Selection of either of the options for 'the total gross receipts of the person in profession exceeds ten lakh rupees during the previous year; or' is mandatory in Part A General Information" & Chr(13)
'            shtPartA_General = False
'            CloseMsg
'        End If
        If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "(Select)" Then
            fmsgbox "*Selection of either of the options for 'the aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is twenty-five thousand rupees or more; or' is mandatory in Part A General Information" & Chr(13)
            shtPartA_General = False
            CloseMsg
        End If
        If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "(Select)" Then
            fmsgbox "*Selection of either of the options for 'The deposits in one or more savings bank account of the person, in aggregate, is fifty lakh rupees or more, in the previous year' is mandatory in Part A General Information" & Chr(13)
            shtPartA_General = False
            CloseMsg
        End If
    End If
    
'    If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" Then
'        'If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_1").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
'        If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_1").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
'
'            fmsgbox "*Selection of either of the options for 'total sales, turnover or gross receipts, as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or ' is mandatory in Sheet Income Details" & Chr(13)
'            shtPartA_General = False
'            CloseMsg
'        End If
'    End If
'    If (Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_1").value = "") Then
'    fmsgbox "*Please enter amount of Total sales, turnover or gross receipts etc in Income Details"
'    shtPartA_General = False
'    CloseMsg
'    End If
    
'    If Sheet1.Range("clauseiv7provisio139iFlg_2").value <> "No" Then
'        fmsgbox "* You have selected ""Yes"" for ""Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)"", hence please enter respective amount in Sheet Income details." & Chr(13)
'        shtPartA_General = False
'    CloseMsg
'    End If

'    If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" Then
'        If Sheet1.Range("clauseiv7provisio139iFlg_2").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_2").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
'
'            fmsgbox "*Selection of either of the options for 'the total gross receipts of the person in profession exceeds ten lakh rupees during the previous year; or' is mandatory in Sheet Income Details" & Chr(13)
'            shtPartA_General = False
'            CloseMsg
'        End If
'    End If
'
'    If (Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_2").value = "") Then
'    fmsgbox "*Please enter amount of total gross receipts in Income Details"
'    shtPartA_General = False
'    CloseMsg
'    End If
'
'    If Sheet1.Range("clauseiv7provisio139iFlg_3").value <> "No" Then
'        fmsgbox "* You have selected ""Yes"" for ""Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)"", hence please enter respective amount in Sheet Income details." & Chr(13)
'        shtPartA_General = False
'    CloseMsg
'    End If
    
    If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" Then
        If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_3").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
        
            fmsgbox "*Selection of either of the options for 'the aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is twenty-five thousand rupees or more; or' is mandatory in Sheet Income Details" & Chr(13)
            shtPartA_General = False
            CloseMsg
        End If
    End If

    If (Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
    fmsgbox "*Please enter amount of tax deducted at source and tax collected at source in Income Details"
    shtPartA_General = False
    CloseMsg
    End If
    
'    If Sheet1.Range("clauseiv7provisio139iFlg_4").value <> "No" Then
'        fmsgbox "* You have selected ""Yes"" for ""Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)"", hence please enter respective amount in Sheet Income details." & Chr(13)
'        shtPartA_General = False
'    CloseMsg
'    End If
     If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" Then
        If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_4").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
            fmsgbox "*Selection of either of the options for 'if his total deposits in a savings bank account is fifty lakh rupees or more, in the previous year.' is mandatory in Sheet Income Details" & Chr(13)
            shtPartA_General = False
            CloseMsg
        End If
     End If

    If (Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_4").value = "") Then
    fmsgbox "*Please enter amount of total deposits in Income Details"
    shtPartA_General = False
    CloseMsg
    End If
End If

If Mid(Sheet1.Range("clauseiv7provisio139iFlg").value, 1, 1) <> "N" Then

'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "Y" And Range("clauseiv7provisio139iAmount_1").value <> "" Then
'    If Range("clauseiv7provisio139iAmount_1").value < 6000000 Or Range("clauseiv7provisio139iAmount_1").value = 6000000 Then
'    fmsgbox "* Please select drop down as ""No"" if ""total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year " & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "Yes" And Range("clauseiv7provisio139iAmount_1").value = "" Then
'        fmsgbox "*you have selected ""Yes"" for ""total sales, turnover or gross receipts, as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or "", hence please enter the amount in Sheet Income details." & Chr(13)
'    End If
    
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "Y" Then
'    If Range("clauseiv7provisio139iAmount_2").value < 1000000 Or Range("clauseiv7provisio139iAmount_2").value = 1000000 Then
'    fmsgbox "* Please select drop down as ""No"" if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year " & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_2").value = "Yes" And Range("clauseiv7provisio139iAmount_2").value = "" Then
'        fmsgbox "*you have selected ""Yes"" for ""the total gross receipts of the person in profession exceeds ten lakh rupees during the previous year; or"", hence please enter the amount in Sheet Income details." & Chr(13)
'    End If
    
    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y" Then
    If Range("clauseiv7provisio139iAmount_3").value < 25000 Then
    fmsgbox "* Please select drop down as ""No"" if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees " & Chr(13)
    End If
    End If
    
     If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "Yes" And Range("clauseiv7provisio139iAmount_3").value = "" Then
        fmsgbox "*you have selected ""Yes"" for ""the aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is twenty-five thousand rupees or more; or"", hence please enter the amount in Sheet Income details." & Chr(13)
    End If
    
    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "Y" Then
    If Range("clauseiv7provisio139iAmount_4").value < 5000000 Then
    fmsgbox "* Please select drop down as ""No"" if total deposits in a savings bank account is less than fifty lakh rupees in the previous year " & Chr(13)
    End If
    End If
    
    If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" And Range("clauseiv7provisio139iAmount_4").value = "" Then
        fmsgbox "*you have selected ""Yes"" for ""if his total deposits in a savings bank account is fifty lakh rupees or more, in the previous year."", hence please enter the amount in Sheet Income details." & Chr(13)
    End If
End If

  
    
    If MsgBx_General <> "" Then
        Sheet1.Activate
        fmsgbox MsgBx_General ', vbOKOnly, "Error(s)!"
        shtPartA_General = False
        CloseMsg
    End If
    
    If MsgBx_General1 <> "" Then
        Sheet1.Activate
        fmsgbox MsgBx_General1 ', vbOKOnly, "Error(s)!"
        shtPartA_General = False
        CloseMsg
    End If
    
    If Msg_Bx_44ABAccount <> "" Then
        Sheet1.Activate
        fmsgbox Msg_Bx_44ABAccount ', vbOKOnly, "Error(s)!"
        shtPartA_General = False
        CloseMsg
    End If
    




End Function


'Function to check the Last NAme
Function chkLastName() As Boolean
On Error Resume Next
    
    chkLastName = True
    LastName = Sheet1.Range("sheet1.SurNameOrOrgName")
    
    If Trim(LastName) = "" Or IsEmpty(LastName) Then
        chkLastName = False
        Exit Function
    End If

    If Len(Trim(LastName)) > 75 Then
        fmsgboxStatus ("Last Name cannot be more than 75 characters in schedule Part A General under Personal Information")
        chkLastName = False
        Exit Function
    End If
    
    'If Not CheckSpecialCharacter4(LastName) Then
    '    chkLastName = False
    '    MsgBox ("Name in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
    '    Exit Function
    'End If
    
End Function


Function ChkSpecialCharinNonMandatoryfield() As Boolean

    ChkSpecialCharinNonMandatoryfield = True
    Dim firstname_1, middlename_1, nameOfPremises_1, roadStreet_1, IncomeTaxWardCircle_1 As Variant

    firstname_1 = Range("sheet1.FirstName").value
    middlename_1 = Range("sheet1.MiddleName").value
    nameOfPremises_1 = Range("sheet1.ResidenceName").value
    roadStreet_1 = Range("sheet1.RoadOrStreet").value
   ' IncomeTaxWardCircle_1 = RanFge("sheet1.DesigOfficerWardorCircle").value


    'If Not CheckSpecialCharacter4(firstname_1) Then
    '    ChkSpecialCharinNonMandatoryfield = False
    '    MsgBox ("First Name in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
    '    Exit Function
    'End If
    
    'If Not CheckSpecialCharacter4(middlename_1) Then
    '    ChkSpecialCharinNonMandatoryfield = False
    '    MsgBox ("Middle Name in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
    '    Exit Function
    'End If
    
    If Not CheckSpecialCharacter4(nameOfPremises_1) Then
        ChkSpecialCharinNonMandatoryfield = False
        fmsgboxStatus ("Name Of Premises in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
        Exit Function
    End If
    
    If Not CheckSpecialCharacter4(roadStreet_1) Then
        ChkSpecialCharinNonMandatoryfield = False
        fmsgboxStatus ("Road/Street/Post Office in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
        Exit Function
    End If

'    If Not CheckSpecialCharacter4(IncomeTaxWardCircle_1) Then
'        ChkSpecialCharinNonMandatoryfield = False
'        FMSGBOXSTATUS  ("Income Tax Ward/Circle in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
'        Exit Function
'    End If
End Function

'Function to check the PAN field
Function ChkPAN() As Boolean
On Error Resume Next
    ChkPAN = True
    PAN = Sheet1.Range("sheet1.PAN")
    
    If Trim(PAN) = "" Or IsEmpty(PAN) Then
        ChkPAN = False
        Exit Function
    End If

    If Len(Trim(PAN)) > 10 Then
        fmsgboxStatus ("Please enter a valid 10 character PAN in schedule Part A General under Personal Information")
        ChkPAN = False
        Exit Function
    End If
End Function

Function ChkFlat() As Boolean
On Error Resume Next
    
    ChkFlat = True
    Flat = Sheet1.Range("sheet1.ResidenceNo")
    
    If Trim(Flat) = "" Or IsEmpty(Flat) Then
        ChkFlat = False
        Exit Function
    End If

    If Len(Trim(Flat)) > 50 Then
        fmsgboxStatus ("Flat/ Door/ Building cannot be more than 50 characters in schedule Part A General under Personal Information")
        ChkFlat = False
        Exit Function
    End If

    If Not CheckSpecialCharacter4(Flat) Then
        ChkFlat = False
        fmsgboxStatus ("Flat/Door/Building  in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
        Exit Function
    End If
End Function

'Modified by sadineni on 05/02/2026
Function ChkFlat2() As Boolean
On Error Resume Next
    
    ChkFlat2 = True
    FlatNew = Sheet1.Range("sheet1.ResidenceNo2")
    
If AreaNew <> "" Or CityNew <> "" Or (StateNew <> "" And StateNew <> "(Select)") Or (sCountryNew <> "" And sCountryNew <> "(Select)") Or PinCodeNew <> "" Or ZipCodeNew <> "" Then
    If Trim(FlatNew) = "" Or IsEmpty(Flat) Then
    '    fmsgboxStatus ("""Flat/ Door/ Building in Sheet: Personal Information is mandatory in Part A General"".")
        ChkFlat2 = False
        Exit Function
    End If
End If

    If Len(Trim(FlatNew)) > 50 Then
        fmsgboxStatus ("""Flat/Door/Block No cannot be more than 50 characters in Part A General"".")
        ChkFlat2 = False
        Exit Function
    End If

'    If Not CheckSpecialCharacter4(FlatNew) Then
'        ChkFlat2 = False
'        fmsgboxStatus ("Flat/Door/Building  in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
'        Exit Function
'    End If
      
    If Not CheckSpecialCharacter4New(FlatNew) Then
        ChkFlat2 = False
        'SIT-112112 by sadineni on 27/02/2026 for AY 2026-27
        fmsgboxStatus ("Flat/Door/Building  in Sheet : Part A-General cannot contain special charachters <> ")
        Exit Function
    End If

End Function
'Status Validation
Function ChkStatus() As Boolean
On Error Resume Next
    ChkStatus = True
    Status = Sheet1.Range("sheet1.Status")
    'Status = Mid(Status, 1, 1)
    
    If Trim(Status) = "" Or Trim(Status) = "(Select)" Then
        ChkStatus = False
        Exit Function
    End If
End Function

'Area / Locality Validation
Function ChkArea() As Boolean
On Error Resume Next
    ChkArea = True
    Area = Sheet1.Range("sheet1.LocalityOrArea")
    RoadOrStreet = Sheet1.Range("sheet1.RoadOrStreet")
    If Trim(Area) = "" Or IsEmpty(Area) Then
        ChkArea = False
        Exit Function
    End If

    If Len(Trim(Area)) > 50 Then
        fmsgboxStatus ("Area/ Locality cannot be more than 50 characters in schedule Part A General under Personal Information")
        ChkArea = False
        Exit Function
    End If

    If Not CheckSpecialCharacter4(Area) Then
        ChkArea = False
        fmsgboxStatus ("Area / Locality  in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
        Exit Function
    End If
End Function

'Town/City/District Validation
Function ChkCity() As Boolean
On Error Resume Next
    ChkCity = True
    City = Sheet1.Range("sheet1.CityOrTownOrDistrict")
    
    If Trim(City) = "" Or IsEmpty(City) Then
        ChkCity = False
        Exit Function
    End If

    If Len(Trim(City)) > 50 Then
        fmsgboxStatus ("Town/City/District cannot be more than 50 characters in schedule Part A General")
        ChkCity = False
        Exit Function
    End If

    If Not CheckSpecialCharacter4(City) Then
        ChkCity = False
        fmsgboxStatus ("Town/City/District in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
        Exit Function
    End If

End Function


'State Validation
Function ChkState() As Boolean
On Error Resume Next
    ChkState = True
    State = Sheet1.Range("sheet1.StateCode1")
    'State = Mid(State, 1, 2)
    If Trim(State) = "" Or Trim(State) = "(Select)" Or Trim(State) = "(" Then
        ChkState = False
        Exit Function
    End If
    
    'Malli-------------------------17/09/2024
    If Trim(State) <> "" Then
     Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
      
      state_targetadd = Replace(Sheet1.Range("sheet1.StateCode1").Address, "$", "")
      PIN_targetadd = Replace(Sheet1.Range("sheet1.PinCode").Address, "$", "")   'sheet1.PinCode
      
      Dim ws1 As Worksheet
      Set ws1 = Worksheets("PART A - General")  '23/09/24
      
      
      '17/10/2024
      
    If state_Validation(ws1, PIN_targetadd, state_targetadd) = False Then Sheet1.Range("sheet1.PinCode").value = ""
    
    
    
    
    '----------------------------
    
    End If
    '-----------------------------------------
End Function

'by sadineni on 09/02/2026
Function ChkStateAlternate() As Boolean
On Error Resume Next
    ChkStateAlternate = True
    StateNew = Sheet1.Range("sheet1.StateCode2")
    'State = Mid(StateNew, 1, 2)
If FlatNew <> "" Or AreaNew <> "" Or CityNew <> "" Or (sCountryNew <> "" And sCountryNew <> "(Select)") Or PinCodeNew <> "" Or ZipCodeNew <> "" Then
    If Trim(StateNew) = "" Or Trim(StateNew) = "(Select)" Or Trim(StateNew) = "(" Then
        ChkStateAlternate = False
        Exit Function
    End If
End If
End Function

'Pincode Validation
Function ChkPincode() As Boolean
On Error Resume Next
    ChkPincode = True
    PinCode = Sheet1.Range("sheet1.PinCode")
    
    
    If sCountry <> "" And UCase(Mid(sCountry, 1, InStr(1, sCountry, "-") - 1)) = "91" Then
        If Trim(PinCode) = "" Or IsEmpty(PinCode) Then
            MsgBx_General = MsgBx_General + "* Pin code is mandatory in tab: Personal Information in schedule Part A General" & Chr(13)
            ChkPincode = False
            Exit Function
        End If
        
    
        If Len(Trim(PinCode)) > 6 Then
            MsgBx_General = MsgBx_General + "* Pin Code must be exactly 6 digits and should not start with zero in schedule Part A General" & Chr(13)
            ChkPincode = False
            Exit Function
        End If
        
        If Mid(PinCode, 1, 1) = "0" Then
            MsgBx_General = MsgBx_General + "* Pincode in Sheet : Part A-General should not start with 0" & Chr(13)
            ChkPincode = False
            Exit Function
        End If
    Else
        PinCode = ""
    End If
    
End Function

'ZipCode Validation
Function ChkZipcode() As Boolean
On Error Resume Next
    ChkZipcode = True
    ZipCode = Sheet1.Range("sheet1.ZipCode")
    
    
      If sCountry <> "" And UCase(Mid(sCountry, 1, InStr(1, sCountry, "-") - 1)) <> "91" Then
        If Trim(ZipCode) = "" Or IsEmpty(ZipCode) Then
            MsgBx_General = MsgBx_General + "* Zip Code is compulsory, if there is no Zip Code then select ""No ZIP Code in schedule Part A General" & Chr(13)
            ChkZipcode = False
            Exit Function
        End If
        
    
        If Len(Trim(ZipCode)) > 8 Then
            MsgBx_General = MsgBx_General + "* ZipCode in Sheet : Part A-General cannot exceed 8 characters.Minimum 1 and up to 8 Characters." & Chr(13)
            ChkZipcode = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(ZipCode) Then
            MsgBx_General = MsgBx_General + "* ZipCode in Sheet : Part A-General characters < > & ' " & Chr(34) & " are not allowed"" & chr(13)"
            ChkZipcode = False
            Exit Function
        End If
    Else
        ZipCode = ""
    End If
End Function

'Mobile No Validation
Function ChkMobileNo() As Boolean
On Error Resume Next
    Dim i As Long
    ChkMobileNo = True
    CountryCode1 = Sheet1.Range("CountryCode1").value
    MobileNo = Sheet1.Range("sheet1.Mobileno")
    
    If Trim(MobileNo) = "" Or IsEmpty(MobileNo) Then
        ChkMobileNo = False
        Exit Function
    End If
    
    If Len(Trim(MobileNo)) > 10 Then
        fmsgboxStatus ("Please enter a valid 10 digit Mobile Number in schedule Part A General")
        ChkMobileNo = False
        Exit Function
    End If
    
'    If Not (Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "99" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "(S") Then
'    If Not Len(Trim(MobileNo)) = 10 Then
'        FMSGBOXSTATUS  ("Mobile No 1 in Sheet : Part A-General Should be  10 characters ")
'        ChkMobileNo = False
'        Exit Function
'    End If
'    End If
    
    If (Sheet1.Range("CountryCode1").value) <> "91" Then
        If Not (Len(Trim(MobileNo)) <= 10 And Len(Trim(MobileNo)) >= 5) Then
            'fmsgboxStatus ("Please enter a valid Mobile Number1 in schedule Part A General")
            
    'by sadineni on 29/01/2026 for AY 2026-27
             fmsgboxStatus ("""Please enter a valid Primary Mobile Number in Part A General.""")
            ChkMobileNo = False
            Exit Function
        End If
    Else
        If Not Len(Trim(MobileNo)) = 10 Then
'            fmsgboxStatus ("""Please enter a valid Mobile Number1 in schedule Part A General""")
            fmsgboxStatus ("""Please enter a valid 10 digit Mobile Number1 in Part A General.""")
            ChkMobileNo = False
            Exit Function
        End If
    End If
    
    If Trim(MobileNo) <> "" Then
        For i = 1 To Len(MobileNo)
            If Not IsNumeric(Mid(MobileNo, i, 1)) Then
                fmsgboxStatus ("Mobile Number in schedule Part A General must contain only digits from 0 to 9")
                ChkMobileNo = False
            Exit Function
            End If
        Next
    End If
    
    
    If Trim(MobileNo) <> "" Then
        If Trim(CountryCode1) = "" Then
            'fmsgboxStatus ("Country Code in Mobile 1 in Sheet: Personal Information is mandatory in schedule Part A General")
            
    'by sadineni on 29/01/2026 for AY 2026-27
             fmsgboxStatus ("""Country Code in Primary Mobile Number in Sheet: Personal Information is mandatory in Part A General .""")
            ChkMobileNo = False
            Exit Function
        End If
    End If
    

    If Len(CountryCode1) > 0 Then
        If Trim(CountryCode1) = "" Or IsEmpty(CountryCode1) Then
            ChkMobileNo = False
            Exit Function
        End If
    
    If Len(Trim(CountryCode1)) > 5 Then
        fmsgboxStatus ("Country Code for Mobile No cannot exceed 5 characters in schedule Part A General")
        ChkMobileNo = False
        Exit Function
    End If
    
    If Trim(CountryCode1) <> "" Then
        For i = 1 To Len(CountryCode1)
            If Not IsNumeric(Mid(CountryCode1, i, 1)) Then
                fmsgboxStatus ("Country Code for Mobile No must contain only digits from 0 to 9 in schedule Part A General")
                ChkMobileNo = False
            Exit Function
            End If
        Next
    End If
    
    If Len(Trim(CountryCode1)) + Len(Trim(MobileNo)) > 15 Then
                 fmsgboxStatus ("Country Code for Mobile No 1 +  Mobile No 1 in Sheet : Part A-General cannot exceed 15 characters ")
                ChkMobileNo = False
    Exit Function
    End If
    
    End If
    
    If Trim(MobileNo) <> "" Then
        If Mid(MobileNo, 1, 1) = "0" Then
'          fmsgboxStatus ("Mobile No 1 in Sheet : PART A - GENERAL  cannot start with 0")
'by sadineni on 29/01/2026 for AY 2026-27
           fmsgboxStatus ("""Primary Mobile Number cannot begin with '0' in Part A General.""")
           ChkMobileNo = False
        Exit Function
        End If
    End If
    
    If Trim(CountryCode1) <> "" Then
        If CountryCode1 = "0" Or CountryCode1 = "00" Or CountryCode1 = "000" Or CountryCode1 = "0000" Or CountryCode1 = "00000" Then
          fmsgboxStatus ("Please enter valid Country Code for Mobile No in Sheet : PART A - GENERAL")
          ChkMobileNo = False
        Exit Function
        End If
    End If

End Function

'by sadineni on 29/01/2026 for AY 2026-27
Function RepChkMobileNo() As Boolean
On Error Resume Next
    Dim i As Long
    Dim RepCountryCode1, RepMobileNo As String
    
    RepChkMobileNo = True
    RepCountryCode1 = Sheet1.Range("sheet1.Repcountrycode").value
    RepMobileNo = Sheet1.Range("sheet1.Contact_Representative")
    
    If Trim(RepCountryCode1) = "" Or IsEmpty(RepCountryCode1) Then
'        fmsgboxStatus ("""Country Code in Mobile Number 2 in Sheet: Personal Information is mandatory.""")
 'by sadinenion on 26/03/2026 as per DE V0.2
        ' fmsgboxStatus ("""Country Code in contact no. of representative assessee in Sheet: Personal Information is mandatory.""")
     'Added by Shrutika- 23/04/2026 V0.3 AY-26-27
        fmsgboxStatus ("""Country Code for contact number of representative assessee in Part A General  cannot be blank.""")
        RepChkMobileNo = False
        Exit Function
    End If

'Konda updated on 04-05-2026--SIT-118882
 If Trim(RepCountryCode1) <> "" Then
        If Trim(Sheet1.Range("sheet1.Contact_Representative").value) = "" Then
          fmsgboxStatus ("Contact Number of representative assessee is mandatory in schedule Part A General")
          RepChkMobileNo = False
        Exit Function
        End If
    End If

    If Len(Trim(RepMobileNo)) > 10 Then
        fmsgboxStatus ("Please enter a valid 10 digit Mobile Number for Representative contact number")
        RepChkMobileNo = False
        Exit Function
    End If
    
    If (Sheet1.Range("sheet1.Repcountrycode").value) <> "91" Then
        If Not (Len(Trim(RepMobileNo)) <= 10 And Len(Trim(RepMobileNo)) >= 5) Then
            fmsgboxStatus ("""Please enter a valid Contact Number of representative assessee in schedule Part A General.""")
            RepChkMobileNo = False
            Exit Function
        End If
    Else
    
        If Not Len(Trim(RepMobileNo)) = 10 Then
           fmsgboxStatus ("""Please enter a valid 10 digit Contact Number of  representative assessee in schedule Part A General.""")
            RepChkMobileNo = False
            Exit Function
        End If
        
    End If
    
    If Trim(RepMobileNo) <> "" Then
        For i = 1 To Len(RepMobileNo)
            If Not IsNumeric(Mid(RepMobileNo, i, 1)) Then
                fmsgboxStatus ("""Mobile Number for Representative contact number in schedule Part A General must contain only digits from 0 to 9""")
                RepChkMobileNo = False
            Exit Function
            End If
        Next
    End If
    
    
    If Trim(RepMobileNo) <> "" Then
        If Trim(RepCountryCode1) = "" Then
                    fmsgboxStatus ("""Country Code in Mobile Number 2 in Sheet: Personal Information is mandatory.""")
                    RepChkMobileNo = False
                    Exit Function
        End If
    End If
    
'Konda updated on 19-05-2026---'SIT-119904
' If Len(RepCountryCode1) > 0 Then
If Len(RepCountryCode1) <> "" Then
    
    If Trim(RepCountryCode1) = "" Or IsEmpty(RepCountryCode1) Then
        RepChkMobileNo = False
        Exit Function
    End If
    
    If Len(Trim(RepCountryCode1)) > 5 Then
        fmsgboxStatus ("Country Code for Representative contact number cannot exceed 5 characters in schedule Part A General")
        RepChkMobileNo = False
        Exit Function
    End If
    
    If Trim(RepCountryCode1) <> "" Then
        For i = 1 To Len(RepCountryCode1)
            If Not IsNumeric(Mid(RepCountryCode1, i, 1)) Then
                fmsgboxStatus ("Country Code for Representative contact number must contain only digits from 0 to 9 in schedule Part A General")
                RepCountryCode1 = ""
                RepChkMobileNo = False
            Exit Function
            End If
        Next
    End If

    If Len(Trim(RepCountryCode1)) + Len(Trim(RepMobileNo)) > 15 Then
                 fmsgboxStatus ("Country Code Representative Mobile No +  Mobile No in Sheet : Part A-General cannot exceed 15 characters ")
                RepChkMobileNo = False
    Exit Function
    End If
    
End If
    
    If Trim(RepMobileNo) <> "" Then
        If Mid(RepMobileNo, 1, 1) = "0" Then
         fmsgboxStatus ("""Contact Number of  representative assessee cannot begin with '0'.""")
          RepChkMobileNo = False
        Exit Function
        End If
    End If
    
    If Trim(RepCountryCode1) <> "" Then
        If RepCountryCode1 = "0" Or RepCountryCode1 = "00" Or RepCountryCode1 = "000" Or RepCountryCode1 = "0000" Or RepCountryCode1 = "00000" Then
          fmsgboxStatus ("Please enter valid Country Code for Representative Mobile No in Sheet : PART A - GENERAL")
          RepChkMobileNo = False
        Exit Function
        End If
    End If

End Function

'Email Address Validation
Function ChkEmail() As Boolean
On Error Resume Next
    ChkEmail = True
    Email = Sheet1.Range("sheet1.EmailAddress")
    
    If Trim(Email) = "" Or IsEmpty(Email) Then
        ChkEmail = False
        Exit Function
    End If
    
    If Len(Trim(Email)) > 125 Then
       ' fmsgboxStatus ("Email Address-1 cannot be more than 125 characters in schedule Part A General")
        'Added by Shrutika- 23/04/2026 V0.3
        fmsgboxStatus ("""Email ID cannot be more than 125 characters in schedule Part A General""")

        ChkEmail = False
        Exit Function
    End If
End Function

Function ChkDOB() As Boolean
On Error Resume Next

    ChkDOB = True
    
    dob = Sheet1.Range("sheet1.DOB")
        
    If Trim(dob) = "" Or IsEmpty(dob) Or dob = "00/00/0000" Then
        ChkDOB = False
        Exit Function
    End If
    
    If Not CheckDOB(dob) Then
        ChkDOB = False
        Exit Function
    End If
End Function
Function ChkDOC() As Boolean
On Error Resume Next

    ChkDOC = True
    
    doc = Sheet1.Range("sheet1.DOC")
        
    If Trim(doc) = "" Or IsEmpty(doc) Or doc = "00/00/0000" Then
        ChkDOC = False
        Exit Function
    End If
    
    If Not CheckDOC(doc) Then
        ChkDOC = False
        Exit Function
    End If
End Function

'Conutry Validation
Function ChkCountry() As Boolean
On Error Resume Next
    ChkCountry = True
    sCountry = Sheet1.Range("sheet1.CountryCode")
   ' scountry = Mid(scountry, 1, 4)
   
    If isdropdownblank(sCountry) Then
        sCountry = ""
    End If
    
    If Trim(sCountry) = "" Or Trim(sCountry) = "(Select)" Or Trim(sCountry) = "(" Then
        ChkCountry = False
        Exit Function
    End If
End Function

'Gender Validation
'Function ChkGender() As Boolean
'On Error Resume Next
'    ChkGender = True
'    Gender = Sheet1.Range("sheet1.Gender1")
'
'  Gender = Mid(Gender, 1, 1)
'  If Gender = "N" Then
'  Gender = "X"
'  End If
'
'
'    If Trim(Gender) = "" Or Trim(Gender) = "(Select)" Or Trim(Gender) = "(" Then
'        ChkGender = False
'        Exit Function
'    End If
'End Function


Function ChkEmpCategory() As Boolean
On Error Resume Next
    ChkEmpCategory = True
    empcat = Sheet1.Range("sheet1.EmployerCategory1")
    
    If Trim(empcat) = "" Or Trim(empcat) = "(Select)" Then
        ChkEmpCategory = False
        Exit Function
    End If
End Function

Function ChkReturnFileSec() As Boolean
On Error Resume Next

    ChkReturnFileSec = True
    ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec")
    ReturnFileSec = Mid(ReturnFileSec, 1, 2)
    
    If Trim(ReturnFileSec) = "" Or Trim(ReturnFileSec) = "(Select)" Or Trim(ReturnFileSec) = "(S" Then
        ChkReturnFileSec = False
        Exit Function
    End If
End Function
Function ChkSeventhProvisoFlag() As Boolean
On Error Resume Next
    ChkSeventhProvisoFlag = True
    ProvisoFlag = Sheet1.Range("sheet1.SeventhProvisoFlag")
  
    If (Trim(ProvisoFlag) = "" Or Trim(ProvisoFlag) = "(Select)") And (Sheet1.Range("sheet1.SeventhProvisoFlag").Locked = False) Then
        ChkSeventhProvisoFlag = False
        Exit Function
    End If
End Function
Function ChkDepositAmountFlag() As Boolean
On Error Resume Next
    ChkDepositAmountFlag = True
    DepositAmountFlag = Sheet1.Range("Sheet1.DepositAmountFlag")
    DepositAmount = Sheet1.Range("Sheet1.DepositAmount")
    
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Trim(DepositAmountFlag) = "" Or Trim(DepositAmountFlag) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y") And (Sheet1.Range("sheet1.DepositAmountFlag").Locked = False) Then
        ChkDepositAmountFlag = False
        Exit Function
    End If
    
    
     If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Mid(DepositAmountFlag, 1, 1) = "Y") And (Sheet1.Range("sheet1.DepositAmount").value = "") Then
        ChkDepositAmountFlag = False
        Exit Function
    End If
    End If
    
    
    If (Trim(DepositAmountFlag) = "Yes" And DepositAmount = "") And (Sheet1.Range("sheet1.DepositAmountFlag").Locked = False) Then
    fmsgbox "Please enter amount of deposit in one or more current account during the previous year in Income Details"
    ChkDepositAmountFlag = False
        Exit Function
    End If
    End If
    
End Function
Function ChkAggrigateAmountFlag() As Boolean
On Error Resume Next
    ChkAggrigateAmountFlag = True
    AggrigateAmountFlag = Sheet1.Range("Sheet1.AggrigateAmountFlag")
    AggrigateAmount = Sheet1.Range("Sheet1.AggrigateAmount")
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Trim(AggrigateAmountFlag) = "" Or Trim(AggrigateAmountFlag) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y") And (Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False) Then
        ChkAggrigateAmountFlag = False
        Exit Function
    End If
    
    
     If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Mid(AggrigateAmountFlag, 1, 1) = "Y") And Trim(Sheet1.Range("Sheet1.AggrigateAmount").value = "") Then
        ChkAggrigateAmountFlag = False
        Exit Function
    End If
    End If
    
    
    If (Trim(AggrigateAmountFlag) = "Yes" And AggrigateAmount = "") And (Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False) Then
    fmsgbox "Please enter amount of expenditure incurred for travel to a foreign country for yourself or for any other person in Income Details"
    ChkAggrigateAmountFlag = False
        Exit Function
    End If
    End If
    
End Function
Function ChkAggrigateAmountFlag1() As Boolean
On Error Resume Next
    ChkAggrigateAmountFlag1 = True
    AggrigateAmountFlag1 = Sheet1.Range("Sheet1.AggrigateAmountFlag1")
    AggrigateAmount1 = Sheet1.Range("Sheet1.AggrigateAmount1")
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Trim(AggrigateAmountFlag1) = "" Or Trim(AggrigateAmountFlag1) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y") And (Sheet1.Range("Sheet1.AggrigateAmountFlag1").Locked = False) Then
        ChkAggrigateAmountFlag1 = False
        Exit Function
    End If
    
      If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Mid(AggrigateAmountFlag1, 1, 1) = "Y") And Trim(Sheet1.Range("Sheet1.AggrigateAmount1").value = "") Then
        ChkAggrigateAmountFlag1 = False
        Exit Function
    End If
    End If
    
    
    If (Trim(AggrigateAmountFlag1) = "Yes" And AggrigateAmount1 = "") And (Sheet1.Range("Sheet1.AggrigateAmountFlag1").Locked = False) Then
    fmsgbox "Please enter amount of expenditure incurred on consumption of electricity during the previous year in Income Details"
    ChkAggrigateAmountFlag1 = False
        Exit Function
    End If
    End If
End Function




Function ChkDepositAmount() As Boolean
On Error Resume Next
    ChkDepositAmount = True
    DepositAmountFlag = Sheet1.Range("Sheet1.DepositAmountFlag")
    DepositAmount = Sheet1.Range("Sheet1.DepositAmount")
    
 
     If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Mid(DepositAmountFlag, 1, 1) = "Y") And Trim(Sheet1.Range("sheet1.DepositAmount").value) = "" Then
        ChkDepositAmount = False
        Exit Function
    End If
    End If
    
    
 
    
End Function
Function ChkAggrigateAmount() As Boolean
On Error Resume Next
    ChkAggrigateAmount = True
    AggrigateAmountFlag = Sheet1.Range("Sheet1.AggrigateAmountFlag")
    AggrigateAmount = Sheet1.Range("Sheet1.AggrigateAmount")
   
     If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    If (Mid(AggrigateAmountFlag, 1, 1) = "Y") And Trim(Sheet1.Range("Sheet1.AggrigateAmount").value) = "" Then
        ChkAggrigateAmount = False
        Exit Function
    End If
    End If
    
    
End Function
Function ChkAggrigateAmount1() As Boolean
On Error Resume Next
    ChkAggrigateAmount1 = True
    AggrigateAmountFlag1 = Sheet1.Range("Sheet1.AggrigateAmountFlag1")
    AggrigateAmount1 = Sheet1.Range("Sheet1.AggrigateAmount1")
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value <> "No" Then
    
    
    
    If (Mid(AggrigateAmountFlag1, 1, 1) = "Y") And Trim(Sheet1.Range("Sheet1.AggrigateAmount1").value) = "" Then
        ChkAggrigateAmount1 = False
        Exit Function
    End If
    End If
    
    
End Function





'Function ChkReturnType() As Boolean
'On Error Resume Next
'    ChkReturnType = True
'    ReturnType = Sheet1.Range("sheet1.ReturnType1")
'    'ReturnType = Mid(ReturnType, 1, 1)
'
'    If (Mid(Sheet1.Range("sheet1.ReturnType1").value, 1, 3) = "Rec") Then
'    ReturnType = "T"
'    ElseIf (Mid(Sheet1.Range("sheet1.ReturnType1").value, 1, 1) = "I") Then
'    ReturnType = "C"
'    Else
'    ReturnType = Mid(ReturnType, 1, 1)
'    End If
'
'    If Trim(ReturnType) = "" Or Trim(ReturnType) = "(Select)" Or Trim(ReturnType) = "(" Then
'        ChkReturnType = False
'        Exit Function
'    End If
'End Function

Function ChkResidentialStatus() As Boolean
On Error Resume Next

    ChkResidentialStatus = True
    ResidentialStatus = Sheet1.Range("sheet1.ResidentialStatus1")
    ResidentialStatus = Mid(ResidentialStatus, 1, 3)
    If isdropdownblank(ResidentialStatus) Then
        ChkResidentialStatus = False
        Exit Function
    End If
End Function

Function ChkResidentialStatusCondition() As Boolean
On Error Resume Next

    ChkResidentialStatusCondition = True
    ResidentialStatusCondition = Sheet1.Range("sheet1.ResidentialStatus1_Condition")
    
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
    If isdropdownblank(ResidentialStatusCondition) Then
        ChkResidentialStatusCondition = False
        Exit Function
    End If
    End If
End Function

Function ChkDirectorComapny() As Boolean
On Error Resume Next

    ChkDirectorComapny = True
    DirectorCompany = Sheet1.Range("IsDirectorCompany")
    
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
    If isdropdownblank(DirectorCompany) Then
        ChkDirectorComapny = False
        Exit Function
    End If
    End If
End Function

Function ChkPartnerFirm() As Boolean
On Error Resume Next

    ChkPartnerFirm = True
    PartnerFirm = Sheet1.Range("IsPartnerFirm")
    
    
    If isdropdownblank(PartnerFirm) Then
        ChkPartnerFirm = False
        Exit Function
    End If
    
End Function



Function ChkShares() As Boolean
On Error Resume Next

    ChkShares = True
    SharesOwner = Sheet1.Range("IsShareOwner")
    
    
    If isdropdownblank(SharesOwner) Then
        ChkShares = False
        Exit Function
    End If
    
End Function

Function ChkAdhaarFlag() As Boolean
On Error Resume Next
ChkAdhaarFlag = True

If Sheet1.Range("sheet1.adhaarflag") = "" Or Sheet1.Range("sheet1.adhaarflag") = "(Select)" Then
    'ChkAdhaarFlag = False
End If
End Function

Function ChkPassNo() As Boolean
    On Error Resume Next
    ChkPassNo = True
    
    If Sheet1.Range("sheet1.PassportNo").value <> "" Then
        If IsNumeric(Sheet1.Range("sheet1.PassportNo").value) Then
            If Trim(CLng(Sheet1.Range("sheet1.PassportNo").value)) = 0 Then
                errmsgPassport = "is Invalid" & Chr(13)
                ChkPassNo = False
            End If
        End If
    End If
End Function

Function ChkAdhaarNo() As Boolean
On Error Resume Next
    ChkAdhaarNo = True
    Dim AadharNumber As Variant
    Dim AadharEnrol As Variant
    Dim AadharNumberOfRep As Variant
    
    AadharNumber = Trim(Sheet1.Range("sheet1.adhaarno"))
    'Commented by Shrutika_EN
  '  AadharEnrol = Trim(Sheet1.Range("sheet1.adhaarenrol"))
    AadharNumberOfRep = Trim(Sheet1.Range("sheet1.RepAadhar"))
    
    
    If AadharNumber <> "" Then
        If Not IsNumeric(AadharNumber) Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
    
        If AadharNumber = "000000000000" Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
        
        If AadharNumber < 0 Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
        
        If AadharNumber = "111111111111" Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
        
    
        If Len(AadharNumber) <> 12 Then
            errmsgAdhaar = "For Aadhaar Number- 12 digit."
            ChkAdhaarNo = False
            Exit Function
        End If
    End If

    If AadharNumberOfRep <> "" Then
        If Not IsNumeric(AadharNumberOfRep) Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
    
        If AadharNumberOfRep = "000000000000" Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
        
        If AadharNumberOfRep = "111111111111" Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo = False
            Exit Function
        End If
        
    
        If Len(AadharNumberOfRep) <> 12 Then
            errmsgAdhaar = "For Aadhaar Number- 12 digit."
            ChkAdhaarNo = False
            Exit Function
        End If
    End If


'    If AadharEnrol <> "" Then
'        If Not IsNumeric(AadharEnrol) Then
'            'errmsgAdhaar = "is invalid"
'            ChkAdhaarNo = False
'            Exit Function
'        End If
'
'        If AadharEnrol = "0000000000000000000000000000" Then
'            'errmsgAdhaar = "is invalid"
'            ChkAdhaarNo = False
'            Exit Function
'        End If
'
'        If AadharEnrol = "1111111111111111111111111111" Then
'            'errmsgAdhaar = "is invalid"
'            ChkAdhaarNo = False
'            Exit Function
'        End If
'
'        If Len(AadharEnrol) <> 28 Then
'            errmsgAdhaar = "For Aadhaar Enrolment Id- 28 digit."
'            ChkAdhaarNo = False
'            Exit Function
'        End If
'
'        'If Not ValidateAadharEnrol(AadharEnrol) Then
'        '    ChkAdhaarNo = False
'        '    Exit Function
'        'End If
'    End If
End Function
'Commented by Shrutika_EN
'Function ChkAdhaarEnrollID() As Boolean
'On Error Resume Next
'    ChkAdhaarEnrollID = True
'    Dim AadharNumber As Variant
'    Dim AadharEnrol As Variant
'    Dim AadharNumberOfRep As Variant
'
'    AadharEnrol = Trim(Sheet1.Range("sheet1.adhaarenrol"))
'
'    If AadharEnrol <> "" Then
'        If Not IsNumeric(AadharEnrol) Then
'            'errmsgAdhaar = "is invalid"
'            ChkAdhaarEnrollID = False
'            Exit Function
'        End If
'
'        If AadharEnrol = "0000000000000000000000000000" Then
'            'errmsgAdhaar = "is invalid"
'            ChkAdhaarEnrollID = False
'            Exit Function
'        End If
'
'        If AadharEnrol = "1111111111111111111111111111" Then
'            'errmsgAdhaar = "is invalid"
'            ChkAdhaarEnrollID = False
'            Exit Function
'        End If
'
'        If Len(AadharEnrol) <> 28 Then
'            errmsgAdhaar = "For Aadhaar Enrolment Id- 28 digit."
'            ChkAdhaarEnrollID = False
'            Exit Function
'        End If
'
'        'If Not ValidateAadharEnrol(AadharEnrol) Then
'        '    ChkAdhaarNo = False
'        '    Exit Function
'        'End If
'    End If
'End Function
Function ChkGovernedbyPortugese() As Boolean
On Error Resume Next
    ChkGovernedbyPortugese = True
    
    Portugese = Sheet1.Range("sheet1.PortugeseCC5A")
    Portugese = Mid(Portugese, 1, 1)
    If Trim(Portugese) = "" Or Trim(Portugese) = "(Select)" Or Trim(Portugese) = "(" Then
        ChkGovernedbyPortugese = False
        Exit Function
    End If
End Function


'Function ChkTransaction94A() As Boolean
'On Error Resume Next
'    ChkTransaction94A = True
'    Transaction94A = Sheet1.Range("sheet1.NotifiedUs94AFlg")
'    If Trim(Transaction94A) = "" Or Trim(Transaction94A) = "(Select)" Or Trim(Transaction94A) = "(" Then
'        ChkTransaction94A = False
'        Exit Function
'
'    End If
'End Function
Function ChkForeirgnexchange() As Boolean
On Error Resume Next
Dim Foreign_exchange As Variant
ChkForeirgnexchange = True
 Foreign_exchange = Sheet1.Range("sheet1.Foreign_Exchange").value
 Foreign_exchange = Mid(Foreign_exchange, 1, 1)
 If Trim(Foreign_exchange) = "" Or Trim(Foreign_exchange) = "(" Then
        ChkForeirgnexchange = False
        Exit Function
 End If
End Function
Function ChkClaim115H() As Boolean
On Error Resume Next
    ChkClaim115H = True
    Dim res_1 As Variant
    
    res_1 = Sheet1.Range("sheet1.ResidentialStatus1").value
    
    Claim115H = Sheet1.Range("Sheet1.115H")
    Claim115H = Mid(Claim115H, 1, 1)
    
    If isdropdownblank(Claim115H) Then
        Claim115H = ""
    End If
    If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR") Then
        If Trim(Claim115H) = "" Or Trim(Claim115H) = "(Select)" Or Trim(Claim115H) = "(" Then
            MsgBx_General = MsgBx_General + "* Please select whether you want to claim the benefit u/s 115H or not in schedule Part A General" & Chr(13)
            ChkClaim115H = False
            Exit Function
        End If
    End If
End Function

Function ChkAccounts44AA() As Boolean
'On Error Resume Next
'    ChkAccounts44AA = True
'    Accounts44AA = Sheet1.Range("sheet1.NotifiedUs94AFlg")
'    If Trim(Accounts44AA) = "" Or Trim(Accounts44AA) = "(Select)" Or Trim(Accounts44AA) = "(" Then
'        ChkAccounts44AA = False
'        Exit Function
'    End If
'End Function


'Function ChkAccounts44AB() As Boolean
'On Error Resume Next
'    ChkAccounts44AB = True
'    Accounts44AB = Sheet1.Range("sheet1.NotifiedUs94AFlg")
'    If Trim(Accounts44AB) = "" Or Trim(Accounts44AB) = "(Select)" Or Trim(Accounts44AB) = "(" Then
'        ChkAccounts44AB = False
'        Exit Function
'    End If
'
End Function


Function ValidateSectionCodeGrid() As Boolean
ValidateSectionCodeGrid = True

    MsgBx_SectionCode = ""
    'Commented by Shrutika
    'MsgBx_SectionCode = "PART A General" & Chr(13) & "Section Code" & Chr(10)
    
    'If Not ValidateSecn_Code Then ValidateSectionCodeGrid = False
    
    'If Not ValidateSecn_Date Then ValidateSectionCodeGrid = False
    
    If (end_SecCode <> end_SecCodeYN) Then
    
    MsgBx_SectionCode = MsgBx_SectionCode + "* Both Section Code and Have you furnished the audit report is Mandatory "
    ValidateSectionCodeGrid = False
    End If
    
    If Sheet1.Range("sheet1.RepAadhar").value <> "" Then
    'Added by Ayush 43
    If Sheet1.Range("sheet1.adhaarno").value = Sheet1.Range("sheet1.RepAadhar").value Then
   ' MsgBx_SectionCode = MsgBx_SectionCode + "* aadhaar of the represtative cannot be same as assessee Aadhaar in Part A General "
      'MsgBx_SectionCode = "* Aadhar of the representative cannot be same as assessee Aaadhar in Part-A General Information"
      'Added by Shrutika
      MsgBx_SectionCode = "* ""Aadhar of the representative cannot be same as assessee Aaadhar in Part-A General Information."""
      
    ValidateSectionCodeGrid = False
    End If
    End If
    
        
    
 '  Dim SectnCodesRange As Range
 '  Dim i, mIntCtr As Long
 '  Set SectnCodesRange = Sheet1.Range("sheet1.AuditedSection").Cells
 '  mIntCtr = Sheet1.Range("sheet1.AuditedSection").Count
 '  For i = 1 To mIntCtr
 '  If ((SectnCodesRange.Item(i).Value = "") Or (SectnCodesRange.Item(i).Value <> "(Select)")) Then
        
        

End Function

Function ValidateSecn_Code() As Boolean
ValidateSecn_Code = True
    setTableinfo_SecCode
    Dim rangecells As Range
    Set rangecells = Sheet1.Range("sheet1.AuditedSection").Cells
    Dim i As Long
    noOfProcessSub = end_SecCode
    ReDim SecnCode_General(end_SecCode)
    For i = 1 To end_SecCode
    SecnCode_General(i) = rangecells.item(i).value
    UpdateProgressBar
Next
End Function

Function ValidateSecn_Date() As Boolean
ValidateSecn_Date = True
setTableinfo_SecCode
setTableinfo_SecCodeYN
setTableinfo_SecnDate
Dim rangecells As Range
Set rangecells = Sheet1.Range("sheet1.AuditYN").Cells
Dim rangecells1 As Range
Set rangecells1 = Sheet1.Range("sheet1.DateOfAudit").Cells
Dim i As Long
ReDim SecnCodeYN_General(end_SecCodeYN)
ReDim SecnDate_General(end_SecCodeYN)
For i = 1 To end_SecCodeYN
SecnCodeYN_General(i) = rangecells.item(i).value


If Trim(SecnCodeYN_General(i)) = "" Or Trim(SecnCodeYN_General(i)) = "(Select)" Or Trim(SecnCodeYN_General(i)) = "(" Then
    MsgBx_SectionCode = MsgBx_SectionCode + "* Please select whether you have furnished any other audit report at Sr. No  " & i & " in Part A general" & Chr(13)
    ValidateSecn_Date = False
    Exit Function
End If


Next

For i = 1 To end_SecCodeYN

SecnDate_General(i) = rangecells1.item(i).value
If Len(SecnDate_General(i)) = 0 Then
End If

If Mid(SecnCodeYN_General(i), 1, 1) = "Y" Then
If Len(SecnDate_General(i)) = 0 Then
    MsgBx_SectionCode = MsgBx_SectionCode + "* Date of Furnishing the Report is Compulsary at Sr. No  " & i & " in Part A general" & Chr(13)
    ValidateSecn_Date = False
    Exit Function
End If

VerificationDate = Dformat(VerificationDate, "yyyy-mm-dd")
VerificationDate = Sheet52.Range("sheet9.Date").value

If Dformat(SecnDate_General(i), "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
    MsgBx_SectionCode = MsgBx_SectionCode + "* Invalid date: date cannot be greater than date of filing of the return " & Chr(13)
    ValidateSecn_Date = False
    Exit Function
End If
End If
Next

End Function


Sub setTableinfo_SecCode()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet1.Range("sheet1.AuditedSection").count
    Set rangecells = Sheet1.Range("sheet1.AuditedSection").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
                ccount = ccount + 1
                'rangecells.Item(mIntCtr).Select
            End If
    Next
 end_SecCode = ccount
End Sub
Sub setTableinfo_SecCodeYN()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet1.Range("sheet1.AuditYN").count
    Set rangecells = Sheet1.Range("sheet1.AuditYN").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
                ccount = ccount + 1
                'rangecells.Item(mIntCtr).Select
            End If
    Next
 end_SecCodeYN = ccount
End Sub

Sub setTableinfo_SecnDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet1.Range("sheet1.DateOfAudit").count
    Set rangecells = Sheet1.Range("sheet1.DateOfAudit").Cells
     For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
               'rangecells.Item(mIntCtr).Select
           End If
    Next
 end_SecDate = ccount
End Sub


Function CheckState(State As String, ByRef vType As String) As Boolean
On Error Resume Next
'State is Mandatory in tab: Personal Information
    
    CheckState = True
    If Trim(Range("sheet1.StateCode1").value) = "" Or _
       IsEmpty(Range("sheet1.StateCode1").value) Then
        vType = "1"
        CheckState = False
        Exit Function
    End If
    
EXITS:
    If Range("sheet1.StateCode1").value = "(Select)" Then
        Sheet1.Range("sheet1.CountryCode").value = "(Select)"
        
        Sheet1.Unprotect Password:=getmsgstate
           
'           Sheet1.Range("CountryCode1").MergeArea.Locked = False
            Sheet1.Range("CountryCode1").MergeArea.value = ""
'           Sheet1.Range("CountryCode1").MergeArea.Interior.Color = (&HCCFFCC)
'
'           Sheet1.Range("CountryCode2").MergeArea.Locked = False
            Sheet1.Range("CountryCode2").MergeArea.value = ""
'           Sheet1.Range("CountryCode2").MergeArea.Interior.Color = (&HCCFFCC)
           
           
           Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
           Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
           Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
           
           Sheet1.Range("HASZIP").MergeArea.value = ""
           Sheet1.Range("HASZIP").MergeArea.Locked = True
           Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
           
           Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
           Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
           Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
           Sheet1.Protect Password:=getmsgstate
           
    ElseIf UCase(Range("sheet1.StateCode1").value) = "99-STATE OUTSIDE INDIA" Then
        Sheet1.Range("sheet1.CountryCode").value = "(Select)"
        
        Sheet1.Unprotect Password:=getmsgstate
        
    '           Sheet1.Range("CountryCode1").MergeArea.Locked = False
            Sheet1.Range("CountryCode1").MergeArea.value = ""
    '           Sheet1.Range("CountryCode1").MergeArea.Interior.Color = (&HCCFFCC)
    '
    '           Sheet1.Range("CountryCode2").MergeArea.Locked = False
            Sheet1.Range("CountryCode2").MergeArea.value = ""
    '           Sheet1.Range("CountryCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
           Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
           Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
           Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)
           
           Sheet1.Range("HASZIP").MergeArea.value = ""
           Sheet1.Range("HASZIP").MergeArea.Locked = False
           Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)
           
           Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
           Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
           Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
           Sheet1.Protect Password:=getmsgstate
        
        'Malli added------------------------17/09/24
        'Malli added
        
    ElseIf UCase(Range("sheet1.StateCode1").value) = UCase("99-Foreign") Then
        Sheet1.Range("sheet1.CountryCode").value = "(Select)"
        
        Sheet1.Unprotect Password:=getmsgstate
        
'           Sheet1.Range("CountryCode1").MergeArea.Locked = False
            Sheet1.Range("CountryCode1").MergeArea.value = ""
'           Sheet1.Range("CountryCode1").MergeArea.Interior.Color = (&HCCFFCC)
'
'           Sheet1.Range("CountryCode2").MergeArea.Locked = False
            Sheet1.Range("CountryCode2").MergeArea.value = ""
'           Sheet1.Range("CountryCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
           Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
           Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
           Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)
           
           Sheet1.Range("HASZIP").MergeArea.value = ""
           Sheet1.Range("HASZIP").MergeArea.Locked = False
           Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)
           
           Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
           Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
           Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
           Sheet1.Protect Password:=getmsgstate
      
    
        '--------------------------------------------------------
      ElseIf Not UCase(Range("sheet1.StateCode1").value) = "99-STATE OUTSIDE INDIA" Then
        'Application.EnableEvents=False
        Sheet1.Range("sheet1.CountryCode").value = "91-INDIA"

           Sheet1.Range("CountryCode1").MergeArea.value = "91"
           Sheet1.Unprotect Password:=getmsgstate
           
           Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
          'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
           Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

           Sheet1.Range("HASZIP").MergeArea.value = ""
           Sheet1.Range("HASZIP").MergeArea.Locked = True
           Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

           Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
           Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
           Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
           Sheet1.Protect Password:=getmsgstate
           
           
        'Application.EnableEvents=True
        
'    ElseIf Not UCase(Range("sheet1.StateCode1").value) = "99-STATE OUTSIDE INDIA" Or Range("sheet1.StateCode1").value = "(Select)" Then
'        Sheet1.Range("sheet1.CountryCode").value = "91-INDIA"
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'           Sheet1.Range("CountryCode1").MergeArea.value = "91"
''           Sheet1.Range("CountryCode1").MergeArea.Locked = True
''           Sheet1.Range("CountryCode1").MergeArea.Interior.Color = (&HD8D8D8)
''
' '          Sheet1.Range("CountryCode2").MergeArea.value = "91"
''           Sheet1.Range("CountryCode2").MergeArea.Locked = True
''           Sheet1.Range("CountryCode2").MergeArea.Interior.Color = (&HD8D8D8)
'
'           Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
'           Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
'           Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
'
'           Sheet1.Range("HASZIP").MergeArea.value = ""
'           Sheet1.Range("HASZIP").MergeArea.Locked = True
'           Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
'
'           Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
'           Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
'           Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
'        Sheet1.Protect Password:=getmsgstate
    End If
    
End Function


Function CheckPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 3 is for individuals .So,4th character of PAN should be "P" or "H"
'status = Sheet1.Range("sheet1.Status").value

PAN = UCase(PAN)
    CheckPAN = True
   
    If Len(PAN) > 0 Then
   
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
   
            CheckPAN = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
   
            CheckPAN = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
   
            CheckPAN = False
            Exit Function
        End If
        
      If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
     
            CheckPAN = False
            Exit Function
        End If
         
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
     
            CheckPAN = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
     
            CheckPAN = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
     
            CheckPAN = False
            Exit Function
        End If
    End If
End Function

'new
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


'end

Function CheckRepPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckRepPAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckRepPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckRepPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckRepPAN = False
            Exit Function
        End If
        'Commentd by Shrutika(19-12-24)AY-2025
'        If Not Mid(PAN, 4, 1) = "P" Then
'            'CheckRepPAN = False
'            fmsgboxStatus "Pan of the Representative Assesse should have the 4th character as P"
'            Sheet1.Range("sheet1.RepPAN").value = ""
'            Application.EnableEvents = True
'            CloseMsg
'            Exit Function
'        End If
            
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckRepPAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckRepPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckRepPAN = False
            Exit Function
        End If
    End If
End Function

Function ChkAlphabet(Char As String) As Boolean
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

Function ValidatePinCode() As Boolean
On Error Resume Next
    Dim PinCode As String
    Dim i As Long
    
    ValidatePinCode = True
    PinCode = Sheet1.Range("sheet1.PinCode").value
    'If Len(PinCode) > 6 Then
    '    FMSGBOXSTATUS  ("PinCode in Sheet : PART A-GENERAL  should be at 6 digits")
     '   ValidatePinCode = False
      '  Exit Function
   ' End If
    'If PinCode = "" Or IsEmpty(PinCode) Then
     '   FMSGBOXSTATUS  ("PinCode in Sheet : PART A-GENERAL  is Mandatory")
      '   ValidatePinCode = False
       ' Exit Function
    'End If
    
   ' If Mid(PinCode, 1, 1) = "0" Then
    ' FMSGBOXSTATUS  ("PinCode in Sheet : PART A-GENERAL  should not start with 0")
     '   ValidatePinCode = False
      '  Exit Function
   ' End If
    
   ' For i = 1 To Len(PinCode)
    '    If Not IsNumeric(Mid(PinCode, i, 1)) Then
     '        FMSGBOXSTATUS  ("PinCode in Sheet : PART A-GENERAL  must contain only digits from 0 to 9")
      '      ValidatePinCode = False
       '     Exit Function
       ' End If
    'Next
End Function
Function CheckDate10IFBefore(dob As Variant) As Boolean
On Error Resume Next
Dim DOBNew As Variant
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2019, for A.Y. 2018-19

    CheckDate10IFBefore = True
    
    If Len(dob) > 0 Then
    
        If Not FormatNCheckDate(dob) Then
           
           'MsgBox ("Date of filing of Form 10IE : PART A-GENERAL  must be a valid dd/mm/yyyy format")
           MsgBox ("Please enter date in dd/mm/yyyy format in Date of filing of Form 10IE : PART A-GENERAL  ")
           CheckDate10IFBefore = False
            Exit Function
        End If
        
        DOBNew = dob
        Sheet1.Range("sheet1.capacityassessee") = DOBNew
        DOBNew = Dformat(DOBNew, "")
        
        If Not ChkMinInclusiveDate(DOBNew, "2021-04-01") Then
        
        'If CDate("01/04/2021") > CDate(DOB) Then
        
            MsgBox ("Date of filing of Form 10IE cannot be prior to  01/04/2021.")
            CheckDate10IFBefore = False
            Exit Function
        Else
               
       End If
       'new change
       Dim verDate As Variant
        verDate = Sheet52.Range("sheet9.Date").value
        verDate = Dformat(verDate, "yyyy-mm-dd")
        If Dformat(dob, "yyyy-mm-dd") > verDate Then
            MsgBox ("Date of filing of Form 10IE : PART A-GENERAL cannot be after System Date in sheet Part A-General")
            CheckDate10IFBefore = False
            Exit Function
        End If
        'new change
    End If

End Function



Function CheckDOB(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2018, for A.Y. 2015-16
    Dim nstatus As String
    nstatus = Sheet1.Range("sheet1.Status").value
    CheckDOB = True
    nstatus = Mid(nstatus, 1, 1)
    If Trim(dob) <> "" Or Not IsEmpty(dob) Then
        If (dob = "01/01/0001" And nstatus = "H") Then
            CheckDOB = True
            AssesseeDob = dob
            Exit Function
        Else
            If Not FormatNCheckDate(dob) Then
             
                CheckDOB = False
               
                fmsgboxStatus ("Please enter valid date in dd/mm/yyyy format  in schedule Part A General under Personal Information")
                
                Exit Function
            End If
        
            'If Not ChkMaxDOBDate(dob, "31/03/2024") Then 'dpk1601 ' Chandru (Part A Gen DOB)
            'Added by Shrutika(20-12-24)AY-25
          '  If Not ChkMaxDOBDate(dob, "31/03/2025") Then
          
           'Added by Shrutika- 02/01/2026 AY-26-27
            Dim cutoff As Date
             cutoff = CDate(SheetALL.Range("DOB_1").value)
              If Not ChkMaxDOBDate1(dob, SheetALL.Range("DOB_1").value) Then
                 'fmsgboxStatus ("Date should not be after 31/03/2024 for AY 2024-25 in schedule Part A General under Personal Information") 'dpk1601
              '   fmsgboxStatus ("Date should not be after 31/03/2025 for AY 2025-26 in schedule Part A General under Personal Information") 'dpk1601
                fmsgboxStatus ("Date should not be after " & Dformat1(cutoff, "yyyy-mm-dd") & " for AY 2026-27 in schedule Part A General under Personal Information")
                CheckDOB = False
                Exit Function
            Else
                AssesseeDob = dob
            End If
        End If
    End If
 End Function
 Sub c()
 Application.EnableEvents = True
 End Sub
 
 Function CheckDOC(doc As Variant) As Boolean
On Error Resume Next

    CheckDOC = True
            If Not FormatNCheckDate(doc) Then
                CheckDOC = False
                'fmsgboxStatus ("Please enter valid date of commencement in dd/mm/yyyy format in schedule Part A General under Personal Information")
                fmsgboxStatus ("Please enter valid date of commencement in dd/mm/yyyy format.")
                Exit Function
            End If
        
            If Not ChkMaxDOBDate(doc, "31/03/2024") Then 'dpk1601
                 'fmsgboxStatus ("Date of commencement should not be after 31/03/2022 for AY 2022-23 in schedule Part A General under Personal Information")
                 fmsgboxStatus ("Date of commencement should not be after 31/03/2024 for AY 2024-25") 'dpk1601
                CheckDOC = False
                Exit Function
            End If
 End Function
 Function FormatNCheckDate(ByRef dt As Variant, Optional Year As String = "2024") As Boolean 'dpk1601
On Error Resume Next
    FormatNCheckDate = True
    If Len(dt) > 0 Then
        'Format the date in dd/mm/yyyy format
        If Mid(dt, 3, 1) <> "/" Then 'Checking between dd/mm
            If Mid(dt, 3, 1) = "\" Or Mid(dt, 3, 1) = "-" Or Mid(dt, 3, 1) = "." Or Mid(dt, 3, 1) = "," Then
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
            End If
        End If
        
        If Mid(dt, 6, 1) <> "/" Then 'Checking between mm/yyyy
            If Mid(dt, 6, 1) = "\" Or Mid(dt, 6, 1) = "-" Or Mid(dt, 6, 1) = "." Or Mid(dt, 6, 1) = "," Then
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
            End If
        End If
        
        'Checking Date if it is in correct format :Day(dd), Month(mm) & Year(yyyy)
        
        If val(Mid(dt, 1, 2)) < 0 Or val(Mid(dt, 1, 2)) > 31 Then FormatNCheckDate = False
        If val(Mid(dt, 4, 2)) < 0 Or val(Mid(dt, 4, 2)) > 12 Then FormatNCheckDate = False
        If val(Mid(dt, 7, 4)) < 1900 Then FormatNCheckDate = False

        If Not IsDate(dt) Then FormatNCheckDate = False
        
    End If
End Function
'Function ChkMaxDOBDateDef(dob As Variant, maxDefinedDOB As Variant) As Boolean
'On Error Resume Next
'     ChkMaxDOBDateDef = True
'     If Len(dob) > 0 Then
'        If val(Mid(dob, 7, 4)) > 2021 Then
'                    ChkMaxDOBDateDef = False
'                    Exit Function
'        ElseIf val(Mid(dob, 7, 4)) = 2021 Then
'            If val(Mid(dob, 4, 2)) >= 4 Then
'                If val(Mid(dob, 1, 2)) >= 1 Then
'                    ChkMaxDOBDateDef = False
'                    Exit Function
'                End If
'            End If
'        End If
'     End If
'End Function





'Function ChkMaxDOBDate(dob As Variant, maxDefinedDOB As Variant) As Boolean
'On Error Resume Next
'     ChkMaxDOBDate = True
'     If Len(dob) > 0 Then
'        If val(Mid(dob, 7, 4)) > 2022 Then
'                    ChkMaxDOBDate = False
'                    Exit Function
'        ElseIf val(Mid(dob, 7, 4)) = 2022 Then
'            If val(Mid(dob, 4, 2)) >= 4 Then
'                If val(Mid(dob, 1, 2)) >= 1 Then
'                    ChkMaxDOBDate = False
'                    Exit Function
'                End If
'            End If
'        End If
'     End If
'End Function

'Function CheckDateBefore(dob As Variant) As Boolean
'On Error Resume Next
''The DOB should be in DD/MM/YYYY format only.
''DoF should not be after 31/03/2019, for A.Y. 2018-19
'
'    CheckDateBefore = True
'    If Len(dob) > 0 Then
'        If Not FormatNCheckDate(dob) Then
'            CheckDateBefore = False
'           fmsgboxStatus ("Date of Filing in Sheet : PART A-GENERAL  must be a valid dd/mm/yyyy format")
'            Exit Function
'        End If
'
'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2022") Then
'            fmsgboxStatus ("Date of filing in Sheet : PART A-GENERAL  should be on or after 01/04/2022")
'            CheckDateBefore = False
'            Exit Function
'        Else
'                EfilingDate = dob
'       End If
'    End If
' End Function

'Function CheckNoticeDateBefore(dob As Variant) As Boolean
'On Error Resume Next
''The DOB should be in DD/MM/YYYY format only.
''DoF should not be after 31/03/2017, for A.Y. 2017-18
'
'    CheckNoticeDateBefore = True
'    If Len(dob) > 0 Then
'        If Not FormatNCheckDate(dob) Then
'            CheckNoticeDateBefore = False
'           fmsgboxStatus ("Notice Date in Sheet : PART A-GENERAL  must be a valid dd/mm/yyyy format")
'            Exit Function
'        End If
'
'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2022") Then
'            fmsgboxStatus ("Notice Date in Sheet : PART A-GENERAL  should be on or after 01/04/2022")
'            CheckNoticeDateBefore = False
'            Exit Function
'        Else
'                noticeDate = dob
'       End If
'    End If
' End Function


Function CheckEmailAddress(emailAddress As String, ByRef vType As String) As Boolean
On Error Resume Next
    CheckEmailAddress = True
    If Len(emailAddress) > 0 Then
        'Email Address: Special Character Check
        If Not CheckSpecialCharacter(Mid(emailAddress, 1, 1)) Then
            vType = "1"
            CheckEmailAddress = False
            Exit Function
        End If
        
        'Email Address: Length Check
        If Len(emailAddress) > 125 Then
            vType = "2"
            CheckEmailAddress = False
            Exit Function
        End If
                
        If Not IsValidEmail(emailAddress) Then
            vType = "3"
            CheckEmailAddress = False
            Exit Function
        End If
        
        'If Not CheckSpecialCharacter_findconsequtive(emailAddress) Then
      '  If Not CheckSpecialCharacter(emailAddress) Then
      'by sadineni on 30/03/2026 as per DE V0.2
           If Not CheckSpecialCharacter_FindConsecutive_new(emailAddress) Then
            vType = "3"
            CheckEmailAddress = False
            Exit Function
        End If
        
    Else
            vType = "4"
            CheckEmailAddress = False
            Exit Function
    End If
End Function

Function CheckSpecialCharacter4(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacter4 = True
    specialCharArray = Array("&", """", "'", ">", "<")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacter4 = False
            Exit Function
        End If
        Next
    Next
End Function
Function CheckSpecialCharacterSlash(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacterSlash = True
      specialCharArray = Array("*", "!", "-", "&", ".", "#", "~", ";", "?", "@", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacterSlash = False
            Exit Function
        End If
        Next
    Next
End Function


'Change-5, ANK-NA, 25.05.2022
Function CheckSpecialCharacter5(din As Variant) As Boolean
'for each character, check whether the ascii lies in alphabet range
'and number range and /
CheckSpecialCharacter5 = False
Dim i As Long
For i = 1 To Len(din)
    If (asc(Mid(din, i, 1)) >= 47 And asc(Mid(din, i, 1)) <= 57) _
        Or (asc(Mid(din, i, 1)) >= 65 And asc(Mid(din, i, 1)) <= 90) _
        Or (asc(Mid(din, i, 1)) >= 97 And asc(Mid(din, i, 1)) <= 112) Then
        'does not contain special character
        CheckSpecialCharacter5 = False
    Else
        CheckSpecialCharacter5 = True
        Exit Function
    End If
Next i


End Function
'---end change---

Function CheckSpecialCharacter_New(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacter_New = True
    specialCharArray = Array("""", "'", ">", "<")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacter_New = False
            Exit Function
        End If
        Next
    Next
End Function

Function CheckSpecialCharacternew(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 3 is for individuals .So,4th character of PAN should be "P" or "H"
'status = Sheet1.Range("sheet1.Status").value
PAN = UCase(PAN)
Dim i As Long
    CheckSpecialCharacternew = True
    If Len(PAN) > 0 Then
    For i = 1 To Len(PAN)
        If Not (ChkAlphabet(Mid(PAN, i, 1)) Or IsNumeric(Mid(PAN, i, 1))) Then
            CheckSpecialCharacternew = False
            Exit Function
        End If
     Next
    End If
End Function

Function CheckSpecialCharacter(emailid As Variant) As Boolean
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

Function CheckSpecialCharacter_New_1(emailid As Variant) As Boolean 'Newly added by Bindu
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacter_New_1 = True
    'Chandru - In special char added dot "."
    specialCharArray = Array("@", "*", "!", ".", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacter_New_1 = False
            Exit Function
        End If
        Next
    Next
End Function
Function IsValidEmail(strEmail)
On Error Resume Next
    Dim strFieldArray As Variant
    Dim strFieldItem As Variant
    Dim i As Long, c As String, blnIsItValid As Boolean
    blnIsItValid = True
     
    i = Len(strEmail) - Len(Application.Substitute(strEmail, "@", ""))
    If i <> 1 Then IsValidEmail = False: Exit Function
    ReDim strFieldArray(1 To 2)
    strFieldArray(1) = Left(strEmail, InStr(1, strEmail, "@", 1) - 1)
    strFieldArray(2) = Application.Substitute(Right(strEmail, Len(strEmail) - Len(strFieldArray(1))), "@", "")
    For Each strFieldItem In strFieldArray
        If Len(strFieldItem) <= 0 Then
            blnIsItValid = False
            IsValidEmail = blnIsItValid
            Exit Function
        End If
        For i = 1 To Len(strFieldItem)
            c = LCase(Mid(strFieldItem, i, 1))
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                blnIsItValid = False
                IsValidEmail = blnIsItValid
                Exit Function
            End If
        Next i
        If Left(strFieldItem, 1) = "." Or Right(strFieldItem, 1) = "." Then
            blnIsItValid = False
            IsValidEmail = blnIsItValid
            Exit Function
        End If
    Next strFieldItem
    If InStr(strFieldArray(2), ".") <= 0 Then
        blnIsItValid = False
        IsValidEmail = blnIsItValid
        Exit Function
    End If
    
    If InStr(strEmail, "..") > 0 Then
        blnIsItValid = False
        IsValidEmail = blnIsItValid
        Exit Function
    End If
    IsValidEmail = blnIsItValid
End Function


''by sadineni on 29/01/2026 for AY 2026-27
'Function CheckSpecialCharacter_findconsequtive(emailid As Variant) As Boolean
'
'On Error Resume Next
'
'    Dim specialCharArray As Variant
'
'    Dim iCharCount, iSpecialChar As Long
'
'    Dim M_a_L_L_I As Variant
'
'    M_a_L_L_I = 0
'
'    CheckSpecialCharacter_findconsequtive = True
'
'    specialCharArray = Array("-", "/", "_", ".")
'
'   Dim charcount As Boolean
'
'   Dim d As Variant
'
'   charcount = False
'
'    For iCharCount = 1 To Len(emailid)
'
'           If Mid(emailid, iCharCount, 1) = "-" Or Mid(emailid, iCharCount, 1) = "/" Or _
'              Mid(emailid, iCharCount, 1) = "_" Or Mid(emailid, iCharCount, 1) = "." Then
'
'                          M_a_L_L_I = M_a_L_L_I + 1
'
'           Else:
'
'                          M_a_L_L_I = 0
'
'           End If
'
'             If M_a_L_L_I >= 2 And d = iCharCount - 1 Then
'
'                    CheckSpecialCharacter_findconsequtive = False
'
'                    Exit Function
'
'             End If
'
'             d = iCharCount
'
'        Next
'
'End Function

Function ValidateSTDcode() As Boolean
    ValidateSTDcode = True
    Dim STDcode As String
    Dim i As Long
    
    STDcode = Sheet1.Range("sheet1.STDcode").value
    If Len(STDcode) > 7 Then
     fmsgboxStatus ("(STD/ISD code- STD/ISD Code should not be more than 5 digits (do not prefix '0' before STD/ISD code) in schedule Part A General")
    ValidateSTDcode = False
    Exit Function
    End If



    'If Trim(STDcode) = "" Then
    '    If Trim(STDcode) = "" Or IsEmpty(STDcode) Then
    '        FMSGBOXSTATUS  ("STDCode in Sheet : PART A-GENERAL  is Mandatory")
    ''        ValidateSTDcode = False
      '      Exit Function
      '  End If
    'End If
    
    If Trim(Mid(STDcode, 1, 1)) = 0 Then
        fmsgboxStatus ("STD/ISD Code cannot begin with '0' in Part A General.")
        ValidateSTDcode = False
        Exit Function
    End If

    
    If Trim(STDcode) <> "" Then
        For i = 1 To Len(STDcode)
            If Not IsNumeric(Mid(STDcode, i, 1)) Then
                fmsgboxStatus ("Invalid STD/ISD code in schedule Part A General.")
                ValidateSTDcode = False
            Exit Function
            End If
        Next
    End If

End Function

Function ValidatePhoneNo() As Boolean
     Dim i As Long
    
    ValidatePhoneNo = True
    PhoneNo = Sheet1.Range("sheet1.PhoneNo").value
    STDcode = Sheet1.Range("sheet1.STDcode").value

    If Len(PhoneNo) > 10 Then
        fmsgboxStatus ("Phone Number should be at most 10 digits in schedule Part A General")
    ValidatePhoneNo = False
    Exit Function
    End If

    If Trim(STDcode) <> "" Then
        If Trim(PhoneNo) = "" Or IsEmpty(PhoneNo) Then
            fmsgboxStatus ("Please enter Phone Number in schedule Part A General")
            ValidatePhoneNo = False
            Exit Function
        End If
    End If

    If Trim(PhoneNo) <> "" Then
        For i = 1 To Len(PhoneNo)
            If Not IsNumeric(Mid(PhoneNo, i, 1)) Then
              fmsgboxStatus ("Phone Number must contain only digits from 0 to 9 in schedule Part A General")
            ValidatePhoneNo = False
            Exit Function
            End If
        Next
    End If

    If (Trim(PhoneNo) <> "" Or Trim(STDcode) <> "") Then
        If (Len(PhoneNo) + Len(STDcode)) < 10 Then
            'fmsgboxStatus ("STD/ISD Code & Phone Number should be 10 digits and STD/ISD code cannot begin from 0 in schedule Part A General")
            'Chandru
            fmsgboxStatus ("Invalid Phone Number. STD Code + Landline Number should not be less than 10 digits and cannot begin with '0' in Part A General")
                       
            ValidatePhoneNo = False
            Exit Function
        End If
        
        
        If (Len(PhoneNo) + Len(STDcode)) > 12 Then
            'fmsgboxStatus ("STD/ISD Code & Phone Number should be 10 digits and STD/ISD code cannot begin from 0 in schedule Part A General")
            'Chandru
            fmsgboxStatus ("Invalid Phone Number. STD Code + Landline Number should not be more than 12 digits and cannot begin with '0' in Part A General")
                       
            ValidatePhoneNo = False
            Exit Function
        End If
    End If

End Function

Function ValidateMobileNoSec() As Boolean
On Error Resume Next

    Dim i As Long
    ValidateMobileNoSec = True
    MobileNo_2 = Sheet1.Range("sheet1.MobileNoSec")
    CountryCode2 = Sheet1.Range("CountryCode2").value
    
'    If Trim(MobileNo_2) = "" Or IsEmpty(MobileNo_2) Then
'        ValidateMobileNoSec = False
'        Exit Function
'    End If
    
    If Len(Trim(MobileNo_2)) > 10 Then
        fmsgboxStatus ("Please enter a valid 10 digit Mobile Number1 in schedule Part A General")
        ValidateMobileNoSec = False
        Exit Function
    End If
    
    If Len(Trim(MobileNo_2)) > 0 Then
'    If Not (Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "99" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "(S") Then
'    If Not Len(Trim(MobileNo_2)) = 10 Then
'        FMSGBOXSTATUS  ("Mobile No 2 in Sheet : Part A-General Should be  10 characters ")
'        ValidateMobileNoSec = False
'        Exit Function
'    End If
'    End If

     If Len(Trim(CountryCode2)) = 0 Then
     'fmsgboxStatus ("Country Code in Mobile 2 in Sheet: Personal Information is mandatory in schedule Part A General")
     
' by sadineni on 29/01/2026 for AY 2026-27 09/02/2026
      fmsgboxStatus ("""Country Code in Secondary Mobile 2 in Sheet: Personal Information is mandatory in Part A General .""")
        ValidateMobileNoSec = False
        Exit Function
    End If
    
   If Sheet1.Range("CountryCode2").value <> "91" Then
   
         If Not (Len(Trim(MobileNo_2)) <= 10 And Len(Trim(MobileNo_2)) >= 5) Then
    '        fmsgboxStatus ("Please enter a valid Mobile Number2 in schedule Part A General")
    ' by sadineni on 29/01/2026 for AY 2026-27
             fmsgboxStatus ("""Please enter a valid Secondary Mobile Number in Part A General.""")
             ValidateMobileNoSec = False
             Exit Function
         End If
 
    Else
    
            If Not Len(Trim(MobileNo_2)) = 10 Then
        '        fmsgboxStatus ("Please enter a valid 10 digit Mobile Number2 in schedule Part A General")
        ' by sadineni on 29/01/2026 for AY 2026-27
                 fmsgboxStatus ("""Please enter a valid 10 digit Secondary Mobile Number in schedule Part A General""")
                 ValidateMobileNoSec = False
                Exit Function
            End If
    End If

    End If
    
    If Trim(MobileNo_2) <> "" Then
         If Not IsNumeric(Mid(MobileNo_2, 1, 10)) Then
                fmsgboxStatus ("Mobile No 2 must contain only digits from 0 to 9 in schedule Part A General")
                ValidateMobileNoSec = False
            Exit Function
        End If
    End If
    
    If Trim(MobileNo_2) <> "" Then
        If Mid(MobileNo_2, 1, 1) = "0" Then
'          fmsgboxStatus ("Mobile No 2 cannot start with 0 in schedule Part A General")
          fmsgboxStatus ("""Secondary Mobile Number cannot begin with '0' in schedule Part A General""")
        ValidateMobileNoSec = False
        Exit Function
        End If
        
       If Len(Trim(CountryCode2)) > 5 Then
        fmsgboxStatus ("Country Code for Mobile No 2 cannot exceed 5 characters in schedule Part A General")
        ValidateMobileNoSec = False
        Exit Function
    End If
    
    If Trim(CountryCode2) <> "" Then
        For i = 1 To Len(CountryCode2)
            If Not IsNumeric(Mid(CountryCode2, i, 1)) Then
                fmsgboxStatus ("Country Code for Mobile No 2 must contain only digits from 0 to 9 in schedule Part A General")
                ValidateMobileNoSec = False
            Exit Function
            End If
        Next
    End If
    
    If Len(Trim(CountryCode2)) + Len(Trim(MobileNo_2)) > 15 Then
                 fmsgboxStatus ("Country Code for Mobile No 2 +  Mobile No 2 cannot exceed 15 characters in schedule Part A General")
                ValidateMobileNoSec = False
                 Exit Function
    End If
          
End If


    'Updated by sai on 03/03/2026 for SIT-112260
   If Sheet1.Range("CountryCode2").value <> "91" Then
        If Len(Trim(CountryCode2)) > 0 Then
           If Len(Trim(MobileNo_2)) > 10 Then
                'fmsgboxStatus ("Please enter a valid 10 digit Mobile Number2 in schedule Part A General")
              'Added by Shrutika(24/02/2026)SIT-112260
                fmsgboxStatus ("Please enter a valid 10 digit Secondary Mobile Number in schedule Part A General")
                ValidateMobileNoSec = False
                Exit Function
            End If
        End If
 End If
        
        If Trim(CountryCode2) <> "" Then
            If CountryCode2 = "0" Or CountryCode2 = "00" Or CountryCode2 = "000" Or CountryCode2 = "0000" Or CountryCode2 = "00000" Then
         ' fmsgboxStatus ("Please enter valid Country Code for Mobile No 2 in schedule Part A General")
    'Added by Shrutika(24/02/2026)SIT-112260
            fmsgboxStatus ("Please enter valid Country Code for Secondary Mobile No in schedule Part A General")
            ValidateMobileNoSec = False
            Exit Function
        End If
    End If
    

End Function



'Commented by Shrutika_EN
'Sub StatusBasedDefaultValue()
'On Error Resume Next
'
'Sheet1.Unprotect Password:=getmsgstate
'Status = Sheet1.Range("sheet1.Status")
'If (Mid(Status, 1, 1) = "H") Then
''       Range("sheet1.Gender1") = "Not applicable"
''       Sheet1.Range("sheet1.Gender1").MergeArea.Locked = True
''       Sheet1.Range("sheet1.Gender1").MergeArea.Interior.Color = (&HD8D8D8)
'       Range("Sheet1.PortugeseCC5A") = "No"
'       'Range("sheet1.EmployerCategory1") = "NA"
'       'Range("sheet1.adhaarflag") = "NA"
'  If Sheet1.Range("sheet1.PAN").value <> "" Then
'    If ((Mid(Sheet1.Range("sheet1.PAN"), 4, 1) <> "H")) Then
'       Sheet1.Range("sheet1.PAN").value = ""
'       'Sheet6.Range("Ver.PAN").Value = ""
'       Sheet1.Range("sheet1.PAN").Select
'       fmsgboxStatus "Please select valid Status for PAN"
'   End If
' End If
'ElseIf (Mid(Status, 1, 1) = "I") Then
''                Range("sheet1.Gender1") = "(Select)"
''                Sheet1.Range("sheet1.Gender1").MergeArea.Locked = False
''                Sheet1.Range("sheet1.Gender1").MergeArea.Interior.Color = (&HCCFFCC)
''
'      If Sheet1.Range("sheet1.PAN").value <> "" Then
'           If ((Mid(Sheet1.Range("sheet1.PAN"), 4, 1) <> "P") And (Mid(Sheet1.Range("sheet1.PAN"), 4, 1) <> "P")) Then
'                   Sheet1.Range("sheet1.PAN").value = ""
'                   'Sheet6.Range("Ver.PAN").Value = ""
'                   Sheet1.Range("sheet1.PAN").Select
'                   fmsgboxStatus "Please enter a valid 10 character PAN in schedule Part A General under Personal Information for Individual"
'           End If
'    End If
''            If (Mid(Sheet1.Range("sheet1.Gender1"), 1, 1) = "N") Then
''                   Sheet1.Range("sheet1.Gender1").value = "(Select)"
''           End If
'
'  Else
''                  Range("sheet1.Gender1") = "(Select)"
''                  Sheet1.Range("sheet1.Gender1").MergeArea.Locked = False
''                  Sheet1.Range("sheet1.Gender1").MergeArea.Interior.Color = (&HCCFFCC)
'
'
'End If
'
'Sheet1.Protect Password:=getmsgstate
'End Sub
'----------------------------------

Sub RepInfo()
On Error Resume Next

AsseseeRepFlg = Sheet1.Range("sheet1.AsseseeRepFlg")

If (Mid(AsseseeRepFlg, 1, 1) = "N") Then
      If (Not (Range("sheet1.RepName").value = "")) Or (Not isdropdownblank(Range("sheet1.RepCapacity").value)) Or (Not (Range("sheet1.RepAddress").value = "")) Or (Not (Range("sheet1.RepPAN").value = "")) Then
            fmsgboxStatus " If this return is not filed by representative then Representative Name ,Adreess and PAN must not be filled"
            Range("sheet1.RepName").value = ""
'            Range("sheet1.RepCapacity").value = ""
'            Range("sheet1.RepAddress").value = ""
'            Range("sheet1.RepPAN").value = ""
            
       End If
End If
End Sub

Function ValidateReturnFileSection() As Boolean
On Error Resume Next
    ValidateReturnFileSection = True
 Dim vbMessgaeCaption As String
    
'Common.MsgPISheet = ""
  'vbMessgaeCaption = "ITR 3: AY: 2017-18"

    
    'Notice Number Mandatory
    If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "18" Then
      If Not (ValidateDOF And ValidateReceiptNo And ValidateNoticeDate) Then
          ValidateReturnFileSection = False
        Exit Function
        
    End If
   End If
    ' Else
    If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "17" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "19" Then
       
      If Not (ValidateReceiptNo And ValidateDOF) Then
            'Malli comented
            'MsgBx_General1 = MsgBx_General1 + "* Receipt number is mandatory if return is filed as revised/defective/modified" & Chr(13)
            Exit Function
         
         Else
      If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "17" Then
        If (ValidateNoticeDate) Then
        MsgBx_General1 = MsgBx_General1 + "* Notice Date : Income Details are not Mandatory" & Chr(13)
        Exit Function
        End If
        End If
        
        If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "19" Then
        If Not (ValidateNoticeDate) Then
        MsgBx_General1 = MsgBx_General1 + "* Notice Date : Income Details is Mandatory" & Chr(13)
        Exit Function
        End If
        End If
    End If
    End If
        
        'Malli-------------26/09/2024
        'If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "13" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "14" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "15" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "16" Then
        If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "13" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "14" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "15" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "16" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "20" Then
         If Not (ValidateNoticeDate) Then
         'Change-13.04.2023.103.PAG.01
         'MsgBx_General1 = MsgBx_General1 + "* Notice Date is Mandatory for return filed u/s 139(9)/142(1)/148/92CD." & Chr(13)
        ' MsgBx_General1 = MsgBx_General1 + "* Notice Date is Mandatory for return filed u/s 139(9)/142(1)/148/153C/92CD." & Chr(13)
    'Added by Shrutika(23-12-24)AY-25
         MsgBx_General1 = MsgBx_General1 + "* ""Notice/order/advance pricing agreement date is mandatory in Part A General""" & Chr(13)
         Exit Function

         End If
         End If
         
        ' Else
          If (Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "11" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "12") Then
'         If (ValidateDOF Or ValidateReceiptNo Or ValidateNoticeDate) Then
'            MsgBx_General1 = MsgBx_General1 + "*  ReceiptNo Or Date of Filing Original Return  Or NoticeDate : Income Details is Not Mandatory" & Chr(13)
'           Exit Function
'        End If
        End If
        
        If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "13" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "14" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "18" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "20" Then
          If Sheet1.Range("Sheet1.UniqueNo").value = "" Then
        'Change-13.04.2023.103.PAG.02

         'Added by Shrutika(16-05-2025)
'         MsgBx_General1 = MsgBx_General1 + "* ""Unique number is mandatory for filed in response to notice u/s 139(9)/ 142(1)/ 148/ 153A/ 153C"" Or ""Filed u/s 119(2)(b)- After condonation of delay"", enter unique number"" - in Part A General Information""" & Chr(13)
'Changed by sadineni on 27/03/2026 as per DE V0.2
          MsgBx_General1 = MsgBx_General1 + "* ""Unique number is mandatory for filed in response to notice u/s 139(9)/ 142(1)/ 148/ 153A/ 153C"" Or ""Filed u/s 139(9A)-After getting condonation of delay u/s 119(2)(b)"", enter unique number""- in Part A General Information""" & Chr(13)
         
          Exit Function
        End If
        End If

End Function
Function ValidateDOF() As Boolean
On Error Resume Next
    'Dim sDOF
    ValidateDOF = True
    DOF = ""
    
    DOF = Sheet1.Range("sheet1.OrigRetFiledDate")
   If Trim(DOF) = "" Or Trim((DOF)) = "00/00/0000" Then
        ValidateDOF = False
        MsgBx_General1 = MsgBx_General1 + "*Date of filing of original return is mandatory in Part A General" & Chr(13)
      ' Common.MsgPISheet = Common.MsgPISheet + "* Date of filing is  Mandatory for return filed u/s 139(5)/139(9)/92CD" & Chr(13)
        Exit Function
        ValidateDOF
    End If
End Function

Function ValidateReceiptNo() As Boolean
On Error Resume Next
    ValidateReceiptNo = True
    
    sReceiptNo = Sheet1.Range("sheet1.ReceiptNo")
    
    If Trim(sReceiptNo) = "" Or IsEmpty(sReceiptNo) Then
        ValidateReceiptNo = False
        
        'Shrutika modified
        'MsgBx_General1 = MsgBx_General1 + "Receipt number is mandatory if return is filed as revised/defective/modified - in Part A General Information" & Chr(13)
       ' Common.MsgPISheet = Common.MsgPISheet + "* Receipt No is  Mandatory for return filed u/s 139(5)/139(9)/92CD" & Chr(13)
        
        'Malli---
         MsgBx_General1 = MsgBx_General1 + "* ""Receipt number is mandatory if return is filed as revised/defective/modified""- in Part A General Information" & Chr(13)
         'MsgBx_General1 = MsgBx_General1 + "* ""Receipt number is mandatory if return is filed as revised/defective/modified"" - in Part A General Information" & Chr(13)
        Exit Function
    End If
End Function

Function checkfieldspecialcharacterNotice(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacterNotice = True
    Dim arr As Variant
    arr = Array("@", ".", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacterNotice = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateNoticeDate() As Boolean
On Error Resume Next

    ValidateNoticeDate = True
    sNoticeDate = Sheet1.Range("sheet1.NoticeDate")
    
    If Trim(sNoticeDate) = "" Or IsEmpty(sNoticeDate) Then
        ValidateNoticeDate = False
        Exit Function
    End If
End Function

Function ValidateAsseseeRepFlg() As Boolean
 On Error Resume Next

    ValidateAsseseeRepFlg = True
    AsseseeRepFlg = Sheet1.Range("sheet1.AsseseeRepFlg")
    AsseseeRepFlg_1 = Mid(AsseseeRepFlg, 1, 1)
    If Trim(AsseseeRepFlg_1) = "" Or Trim(AsseseeRepFlg_1) = "(Select)" Or Trim(AsseseeRepFlg_1) = "(" Then
        ValidateAsseseeRepFlg = False
        'Sheet1.Range("sheet1.AsseseeRepFlg").Select
        Exit Function
    End If
 End Function
  
  
Function ValidateRepName() As Boolean
    ValidateRepName = True
    RepName = Sheet1.Range("sheet1.RepName")
    If (AsseseeRepFlg_1 = "Y") Then
        If RepName = "" Or IsEmpty(RepName) Then
            ValidateRepName = False
            Exit Function
        End If
    
        If Not CheckSpecialCharacter4(RepName) Then
            ValidateRepName = False
            fmsgboxStatus ("Representative Name in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
            Exit Function
        End If
    End If
End Function

Function ValidateRepCapacity() As Boolean
    ValidateRepCapacity = True
    Repcapacity = Sheet1.Range("sheet1.RepCapacity")
    If (AsseseeRepFlg_1 = "Y") Then
        If isdropdownblank(Repcapacity) Then
            ValidateRepCapacity = False
            Exit Function
        End If
    
        If Not CheckSpecialCharacter4(Repcapacity) Then
            ValidateRepCapacity = False
            fmsgboxStatus ("Representative Capacity in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
            Exit Function
        End If
    End If
End Function


Function ValidateRepAddress() As Boolean
    ValidateRepAddress = True
    RepAddress = Sheet1.Range("sheet1.RepAddress")
    If (AsseseeRepFlg_1 = "Y") Then
        If RepAddress = "" Or IsEmpty(RepAddress) Then
            ValidateRepAddress = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacter4(RepAddress) Then
            ValidateRepAddress = False
            fmsgboxStatus ("Representative Address in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
            Exit Function
        End If
    End If
End Function


Function ValidateRepPAN() As Boolean
ValidateRepPAN = True
 RepPAN = Sheet1.Range("sheet1.RepPAN")
If (AsseseeRepFlg_1 = "Y") Then
    If RepPAN = "" Or IsEmpty(RepPAN) Then
        ValidateRepPAN = False
Exit Function
    End If
End If

End Function

'Newly developed by sadineni on 29/01/2026 for AY 2026-27
Function ValidateRepEmail() As Boolean
    ValidateRepEmail = True
    Dim RepEmail As Variant
    
    RepEmail = Sheet1.Range("sheet1.Email_Representative")
    If (AsseseeRepFlg_1 = "Y") Then
        If RepEmail = "" Or IsEmpty(RepEmail) Then
            ValidateRepEmail = False
            Exit Function
        End If
        
    If Len(Trim(RepEmail)) > 125 Then
       ' fmsgboxStatus ("""Email Address cannot be more than 125 characters""")
       
       'Added by Shrutika- 23/04/2026 V0.3 AY-26-27
      fmsgboxStatus ("""Email id of the representative assessee cannot be more than 125 characters""")
        ValidateRepEmail = False
        Exit Function
    End If

'        If Not CheckSpecialCharacter_findconsequtive(RepEmail) Then
'            ValidateRepEmail = False
'            fmsgboxStatus ("Representative Mail in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
'            Exit Function
'        End If

    End If
End Function


Function ValidatePermanebtEstablishment() As Boolean
ValidatePermanebtEstablishment = True
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
    If ((Sheet1.Range("sheet1.NRI_PE").value = "(Select)") Or (Sheet1.Range("sheet1.NRI_PE").value = "")) Then
      'Sheet1.Range("sheet1.NRI_PE").Select
      ValidatePermanebtEstablishment = False
    End If
End If
End Function


Function ValidateSelection44AA() As Boolean
ValidateSelection44AA = True
'If ((Sheet1.Range("sheet1.LiableSec44AAflg").value = "(Select)") Or (Sheet1.Range("sheet1.LiableSec44AAflg").value = "")) Then
'  '  Sheet1.Range("sheet1.LiableSec44AAflg").Select
'    ValidateSelection44AA = False
'End If
End Function
Function ValidateSelection44AE() As Boolean
ValidateSelection44AE = True
'If ((Sheet1.Range("sheet1.LiableSec44AEflg").value = "(Select)") Or (Sheet1.Range("sheet1.LiableSec44AEflg").value = "")) Then
'    ValidateSelection44AE = False
'End If
End Function
Function ValidateSec5cflg() As Boolean
ValidateSec5cflg = True
'If ((Sheet1.Range("sheet1.LiableSec44AEflg").value = "No") And (Sheet1.Range("sheet1.LiableSec5cflg").value = "" Or Sheet1.Range("sheet1.LiableSec5cflg").value = "(Select)")) Then
'    ValidateSec5cflg = False
'End If
End Function
Function ValidateSecAggrigateflg() As Boolean
ValidateSecAggrigateflg = True
'If ((Sheet1.Range("sheet1.LiableSec5cflg").value = "Yes") And (Sheet1.Range("sheet1.LiableSecAggrigateflg").value = "" Or Sheet1.Range("sheet1.LiableSecAggrigateflg").value = "(Select)")) Then
'    ValidateSecAggrigateflg = False
'End If
End Function
Function ValidateSecAggrigateflg1() As Boolean
ValidateSecAggrigateflg1 = True
'If ((Sheet1.Range("sheet1.LiableSec5cflg").value = "Yes") And (Sheet1.Range("sheet1.LiableSecAggrigate1flg").value = "" Or Sheet1.Range("sheet1.LiableSecAggrigate1flg").value = "(Select)")) Then
'    ValidateSecAggrigateflg1 = False
'End If
End Function

Function VAlidateSelect44Ab() As Boolean
VAlidateSelect44Ab = True
'If ((Sheet1.Range("sheet1.LiableSec44ABflg").value = "(Select)") Or (Sheet1.Range("sheet1.LiableSec44ABflg").value = "")) Then
'    'Sheet1.Range("sheet1.LiableSec44ABflg").Select
'    VAlidateSelect44Ab = False
'End If
End Function


'Public LiableSec92Eflg As Variant
'Public AuditDateSec92E As Variant
Function ReturntypCheck()
ReturntypCheck = True
                     Exit Function

End Function

Function Validate92ESec() As Boolean
    Validate92ESec = True

    LiableSec92Eflg = Sheet1.Range("sheet1.LiableSec92Eflg").value
    LiableSec92Eflg = Mid(LiableSec92Eflg, 1, 1)
    LiableSec92Eflgdia = Sheet1.Range("sheet1.LiableSec92Eflgdia").value
    LiableSec92Eflgdia = Mid(LiableSec92Eflgdia, 1, 1)
    AuditDateSec92E = Sheet1.Range("sheet1.AuditDateSec92E").value
    
    If isdropdownblank(LiableSec92Eflg) Then
        MsgBx_General = MsgBx_General & "Please select whether you are liable for audit u/s 92E  in Sheet : PART A - GENERAL  is Mandatory" & Chr(13)
        Validate92ESec = False
        Exit Function
    End If
    
    If LiableSec92Eflg = "Y" Then
        If LiableSec92Eflgdia = "(" Then
            MsgBx_General = MsgBx_General & "Please select If (di) is Yes, whether the accounts have been audited u/s. 92E? in Sheet : PART A - GENERAL" & Chr(13)
            Validate92ESec = False
            Exit Function
        End If
    End If
    
    
    If LiableSec92Eflgdia = "Y" Then
        If AuditDateSec92E = "" Then
            MsgBx_General = MsgBx_General & "Please select a Date of audit report in Sheet : PART A - GENERAL" & Chr(13)
            Validate92ESec = False
            Exit Function
        End If
                
        If Not CheckDateddmmyyyy(AuditDateSec92E) Then
            Validate92ESec = False
            MsgBx_General = MsgBx_General & "Date of audit report in Sheet : PART A - GENERAL  must be a valid dd/mm/yyyy format" & Chr(13)
            Exit Function
        Else
            AuditDateSec92E = Dformat(AuditDateSec92E, "yyyy-mm-dd")
        End If

        If Not CheckDateMinDDMMYYYY(AuditDateSec92E, 1, 4, 2021, "Date of audit report in sheet PART A - General cannot be less than 01/04/2021") Then
        
            MsgBx_General = MsgBx_General & "Date of audit report in Sheet : PART A - GENERAL  must not be less than 01/04/2021" & Chr(13)
            Validate92ESec = False
            Exit Function
        End If
    Else
        AuditDateSec92E = ""
    End If
End Function

Function Validate44AB_accountant() As Boolean
Validate44AB_accountant = True

'If Mid(Sheet1.Range("sheet1.LiableSec44ABflg").value, 1, 1) = "Y" Then
'    If Mid(Sheet1.Range("sheet1.AuditAccountantFlg").value, 1, 1) = "" Or Mid(Sheet1.Range("sheet1.AuditAccountantFlg").value, 1, 1) = "(" Then
'    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount & "Please select whether accounts have been audited by an accountant in schedule Part A General" & Chr(13)
'    Validate44AB_accountant = False
'    End If
'End If

'If (Mid(Sheet1.Range("sheet1.LiableSec44ABflg").value, 1, 1) = "Y" And Mid(Sheet1.Range("sheet1.AuditAccountantFlg").value, 1, 1) = "Y") Then
    'If Not ValidateDOF_44AB Then Validate44AB_accountant = False
    'If Not Validate_AuditorNAme Then Validate44AB_accountant = False
    'If Not ValidateMemberShipNum Then Validate44AB_accountant = False
    'If Not ValidtaeProprietorship Then Validate44AB_accountant = False
    'If Not ValidatePanAuditor Then Validate44AB_accountant = False
    'If Not ValidateDateofAuditReport Then Validate44AB_accountant = False
    
   ' MsgBox (Msg_Bx_44ABAccount)
   ' CloseMsg
'End If
End Function


Function ValidateDOF_44AB() As Boolean
ValidateDOF_44AB = True
If Sheet1.Range("sheet1.AuditReportFurnishDate").value = "" Then
    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Please fill the Date of furnishing of the audit report in schedule Part A General" & Chr(13)
    ValidateDOF_44AB = False
End If
End Function

Function Validate_AuditorNAme() As Boolean
    Validate_AuditorNAme = True
    If Sheet1.Range("sheet1.AuditorName").value = "" Then
        Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Please enter Name of the Auditor signing the tax audit Report in schedule Part A General" & Chr(13)
        Validate_AuditorNAme = False
    End If

    If Not CheckSpecialCharacter_New(Sheet1.Range("sheet1.AuditorName")) Then
        Validate_AuditorNAme = False
        Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* AuditorName in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' " & "in schedule Part A General" & Chr(13)
        Exit Function
    End If
End Function


Function ValidateMemberShipNum() As Boolean
ValidateMemberShipNum = True
If Sheet1.Range("sheet1.AuditorMemNo").value = "" Then
    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Please enter MemberShip No. of the auditor in schedule Part A General" & Chr(13)
    ValidateMemberShipNum = False
    Exit Function
End If
If Sheet1.Range("sheet1.AuditorMemNo").value = 0 Then
    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* MemberShip No. of the auditor cannot have only zeroes in schedule Part A General" & Chr(13)
    ValidateMemberShipNum = False
    Exit Function
End If
End Function

Function ValidtaeProprietorship() As Boolean
    ValidtaeProprietorship = True


    If Not CheckSpecialCharacter_New(Sheet1.Range("sheet1.AudFrmName")) Then
        ValidtaeProprietorship = False
        Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* AudFrmName in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' " & "in schedule Part A General" & Chr(13)
        Exit Function
    End If

    If Sheet1.Range("sheet1.AudFrmRegNo").value = "00000000" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "0000000" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "000000" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "00000" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "0000" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "000" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "00" Or _
        Sheet1.Range("sheet1.AudFrmRegNo").value = "0" Then

        Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Proprietorship/firm registration number cannot be zero in schedule Part A General" & Chr(13)
        ValidtaeProprietorship = False
    End If

    If Not CheckSpecialCharacternew(Sheet1.Range("sheet1.AudFrmRegNo")) Then
        ValidtaeProprietorship = False
        Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Proprietorship/firm registration number in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' " & "in schedule Part A General" & Chr(13)
        Exit Function
    End If

End Function

Function ValidatePanAuditor() As Boolean
ValidatePanAuditor = True
If Sheet1.Range("sheet1.AudFrmPAN").value = "" Then
    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Please enter PAN of the Auditor in schedule Part A General" & Chr(13)
ValidatePanAuditor = False
End If

If Sheet1.Range("sheet1.AudFrmPAN").value = "" And Sheet1.Range("sheet1.AudFrmAadhaar").value = "" Then
    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Please enter the PAN / Aadhaar of the proprietorship in schedule Part A General" & Chr(13)
ValidatePanAuditor = False
End If

End Function


Function ValidateDateofAuditReport() As Boolean
ValidateDateofAuditReport = True
If Sheet1.Range("sheet1.AuditDate").value = "" Then
    Msg_Bx_44ABAccount = Msg_Bx_44ABAccount + "* Please enter Date of Report of the Audit in schedule Part A General" & Chr(13)
ValidateDateofAuditReport = False
End If
End Function



Function CheckAudFrmPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckAudFrmPAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
            
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckAudFrmPAN = False
            Exit Function
        End If
    End If
End Function
Function ValidateActSection() As Boolean
 ValidateActSection = True
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim cellrange As Variant
    Dim cellrange1 As Variant
    Dim i As Long
    setTblinfo_AuditInfo2
    setTblinfo_AuditInfo1
    setTblinfo_AuditInfo
    Set rangecells1 = Sheet1.Range("Audit.Act").Cells
    ReDim AuditAct_Info1(end_AuditInfo2)
    Set rangecells2 = Sheet1.Range("Audit.Date").Cells
    ReDim Date_Info1(end_AuditInfo2)


    For i = 1 To end_AuditInfo2
    cellrange = GetMergedAddressCell(rangecells1, i)
        AuditAct_Info1(i) = Sheet1.Range(cellrange).value
         cellrange1 = GetMergedAddressCell(rangecells2, i)
        Date_Info1(i) = Sheet1.Range(cellrange1).value
        
        If isdropdownblank(AuditAct_Info1(i)) Then
            MsgBx_General1 = MsgBx_General1 + ("Please enter Act at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            ValidateActSection = False
            Exit Function
        End If
      Next
      
    For i = 1 To end_AuditInfo2
    cellrange1 = GetMergedAddressCell(rangecells2, i)
        Date_Info1(i) = Sheet1.Range(cellrange1).value
        VerificationDate = Sheet52.Range("sheet9.Date").value
        VerificationDate = Dformat(VerificationDate, "yyyy-mm-dd")
        If Dformat(Date_Info1(i), "yyyy-mm-dd") > VerificationDate Then
            MsgBx_General1 = MsgBx_General1 + ("Invalid date: date cannot be greater than date of filing of the return") & Chr(13)
            ValidateActSection = False
            Exit Function
        End If
      Next
      
End Function
Function ValidateAct() As Boolean
    ValidateAct = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    'Dim hstatus As Variant
    Dim cellrange As Variant
    Dim cellrange1 As Variant
    Dim SharesIndeterministic As Boolean
    
    'setTblinfo_AuditInfo
   
    'setTblinfo_AuditInfo1
    
    Set rangecells = Sheet1.Range("Audit.Act").Cells
    ReDim AuditAct_Info(end_AuditInfo)
    
    
    Dim i As Long
    For i = 1 To end_AuditInfo
    
    cellrange = GetMergedAddressCell(rangecells, i)
    AuditAct_Info(i) = Sheet1.Range(cellrange).value
        

   
        If isdropdownblank(AuditAct_Info(i)) Then
            MsgBx_General1 = MsgBx_General1 + ("Please select the Name of Other Act under which you are liable to get audited at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            ValidateAct = False
            Exit Function
        End If
        

    Next
End Function
Function ValidateAct_Others() As Boolean
    ValidateAct_Others = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    'Dim hstatus As Variant
    Dim cellrange As Variant
    Dim cellrange1 As Variant
    Dim SharesIndeterministic As Boolean
    
   ' setTblinfo_AuditInfo
'    setTblinfo_AuditInfoOthers
    

    Set rangecells = Sheet1.Range("Audit.Others").Cells
    ReDim AuditAct_Info_Others(end_AuditInfo)
   ' ReDim AuditAct_Info(end_AuditInfo)

    
    
    Dim i As Long
    For i = 1 To end_AuditInfo
    

    cellrange = GetMergedAddressCell(rangecells, i)
    AuditAct_Info_Others(i) = Sheet1.Range(cellrange).value
         
   
        If AuditAct_Info(i) = "Others" Then
        If Not chkCompulsory(AuditAct_Info_Others(i)) Then
            MsgBx_General1 = MsgBx_General1 + ("Description is mandatory if Others is selected at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            ValidateAct_Others = False
            Exit Function
        End If
        End If

    Next
End Function
Function ValidateAuditSection() As Boolean
    ValidateAuditSection = True
    Dim rangecells As Range
    Dim hstatus, cellrange As Variant
    Dim SharesIndeterministic As Boolean
    Dim i As Long
    
    'setTblinfo_AuditInfo
    Set rangecells = Sheet1.Range("Audit.Sections").Cells
    ReDim AuditSection_Info(end_AuditInfo)
    
    For i = 1 To end_AuditInfo
    cellrange = GetMergedAddressCell(rangecells, i)
        AuditSection_Info(i) = Sheet1.Range(cellrange).value
        'AuditSection_Info(i) = rangecells.Item(i).Value
        If Not chkCompulsory(AuditSection_Info(i)) Then
            MsgBx_General1 = MsgBx_General1 + ("Please enter Section at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            ValidateAuditSection = False
            Exit Function
        End If
        
        If Len(AuditSection_Info(i)) > 30 Then
            MsgBx_General1 = MsgBx_General1 + ("Section at Sr. No cannot exceed 30 characters " & i & "in schedule Part A General") & Chr(13)
            ValidateAuditSection = False
            Exit Function
        End If
        
    Next
End Function

Function ValidateDateOfAudit() As Boolean
    ValidateDateOfAudit = True

    setTblinfo_AuditInfo
    Dim rangecells As Range
    Dim rangecells3 As Range
    Dim rangecells2 As Range
    Dim rangecells4 As Range
    Dim i As Long
    Dim cellrange, cellrange1, cellrange2, cellrange4 As Variant
    Set rangecells = Sheet1.Range("Audit.Date").Cells
    Set rangecells4 = Sheet1.Range("Audit.ACTYN").Cells
    Set rangecells3 = Sheet1.Range("Audit.Act").Cells
    ReDim AuditDate_Info(end_AuditInfo)
    ReDim Audit_YN(end_AuditInfo)
    For i = 1 To end_AuditInfo
    
    cellrange = GetMergedAddressCell(rangecells, i)
    cellrange4 = GetMergedAddressCell(rangecells4, i)
     cellrange2 = GetMergedAddressCell(rangecells3, i)
        Audit_YN(i) = Sheet1.Range(cellrange4).value
        AuditDate_Info(i) = Sheet1.Range(cellrange).value
        
        If Sheet1.Range(cellrange2).value <> "" Then
             If isdropdownblank(Sheet1.Range(cellrange4).value) Then
             MsgBx_General1 = MsgBx_General1 + ("Please Select Have you got audited under the selected Act other than the Income-tax Act?  at Sr. No  " & i & "in schedule Part A General")
            ValidateDateOfAudit = False
            Exit Function
             End If
       End If
       
       If Mid(Audit_YN(i), 1, 1) = "Y" Then
            If AuditDate_Info(i) = "" Then
            MsgBx_General1 = MsgBx_General1 + ("Please enter valid date at Sr. No  " & i & "in schedule Part A General")
            ValidateDateOfAudit = False
            Exit Function
             End If
       End If


        If Not AuditDate_Info(i) = "" Then
        If Not CheckDateddmmyyyy(AuditDate_Info(i)) Then
            ValidateDateOfAudit = False
            MsgBx_General1 = MsgBx_General1 + ("Date of Audit at Sr. No  " & i & "  in Sheet PART-A General is invalid. Please enter in dd/mm/yyyy format in schedule Part A General") & Chr(13)
            Exit Function
        Else
           
            If Not ChkMinInclusiveDate(Dformat(AuditDate_Info(i), ""), "2021-04-01") Then ''
            
               MsgBx_General1 = MsgBx_General1 + ("Date of Audit at Sr. No  " & i & "should not be less than 01/04/2021 in schedule Part A General") & Chr(13)
               ValidateDateOfAudit = False
               Exit Function
            End If
        
        End If
        End If
     Next
     
     For i = 1 To Sheet1.Range("Audit.Date").count
    Set rangecells2 = Sheet1.Range("Audit.Date").Cells
    Set rangecells3 = Sheet1.Range("Audit.Act").Cells
    cellrange1 = GetMergedAddressCell(rangecells2, i)
    cellrange2 = GetMergedAddressCell(rangecells3, i)
         If Sheet1.Range(cellrange1).value <> "" Then
             If isdropdownblank(Sheet1.Range(cellrange2).value) Then
             MsgBx_General1 = MsgBx_General1 + ("Please enter Audit Act at Sr. No  " & i & "in schedule Part A General")
            ValidateDateOfAudit = False
            Exit Function
        End If
       End If
     
     Next
End Function
Sub setTblinfo_AuditInfo()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Audit.Act").count
    Set rangecells = Sheet1.Range("Audit.Act").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_AuditInfo = ccount
    rngname_AuditInfo = "Audit.Act;Audit.Others;Audit.Sections;Audit.ACTYN;Audit.Date;"
End Sub
Sub setTblinfo_AuditInfoOthers()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Audit.Others").count
    Set rangecells = Sheet1.Range("Audit.Others").Cells
    For mIntCtr = 1 To mIntCells
        If Not Trim(rangecells.item(mIntCtr).value) = "" Then
            ccount = ccount + 1
        End If
    Next
    end_AuditInfo_Others = ccount
    rngname_AuditInfo = "Audit.Act;Audit.Others;Audit.Sections;Audit.ACTYN;Audit.Date;"
End Sub
Sub setTblinfo_AuditInfo1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Audit.Sections").count
    Set rangecells = Sheet1.Range("Audit.Sections").Cells
    For mIntCtr = 1 To mIntCells
        If Not Trim(rangecells.item(mIntCtr).value) = "" Then
            ccount = ccount + 1
        End If
    Next
    end_AuditInfo1 = ccount
    rngname_AuditInfo1 = "Audit.Act;Audit.Others;Audit.Sections;Audit.ACTYN;Audit.Date;"
End Sub
Sub setTblinfo_AuditInfo2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim j As Long
     Dim b As Long
    
    ccount = 0
  
    mIntCells = Sheet1.Range("Audit.Date").count
    Set rangecells = Sheet1.Range("Audit.Date").Cells
    For mIntCtr = 1 To mIntCells
      'j = 1
      b = 0
         For j = mIntCtr To mIntCells
         
        If Not (Trim(rangecells.item(mIntCtr).value) = "" Or Not Trim(rangecells.item(j).value) = "") Then ' Or Not Trim(rangecells.Item(mIntCtr + 2).Value) = "" Then
            b = 1
            
          End If
        Next
        If b = 1 Then
        ccount = ccount + 1
        End If
    Next
    end_AuditInfo2 = ccount
    rngname_AuditInfo2 = "Audit.Act;Audit.Others;Audit.Sections;Audit.ACTYN;Audit.Date;"
End Sub
Function ValidateAct11()
    
    
    
    setTblinfo_AuditInfo
    setTblinfo_AuditInfo1
    setTblinfo_AuditInfo2
    
    end_AuditInfo = WorksheetFunction.Max(0, end_AuditInfo, end_AuditInfo1, end_AuditInfo2)
            'If Not ValidateAct() Then ValidateAct11 = False
            'If Not ValidateActSection() Then ValidateAct11 = False
            'If Not ValidateAct_Others() Then ValidateAct11 = False
            'If Not ValidateAuditSection() Then ValidateAct11 = False
            'If Not ValidateDateOfAudit() Then ValidateAct11 = False
            
        
End Function
        
Function Validate_Juridiction() As Boolean
    Validate_Juridiction = True
    Dim rangecells As Range
    'Dim hstatus As Variant
    Dim cellrange As Variant
    Dim SharesIndeterministic As Boolean
    
    setTblinfo_NRI
    setTblinfo_NRI1
    Set rangecells = Sheet1.Range("NRI_Juridiction").Cells
    ReDim JuridictionOfResidence(end_NRI)
    
    
    Dim i As Long
    For i = 1 To end_NRI
    
    cellrange = GetMergedAddressCell(rangecells, i)
        JuridictionOfResidence(i) = Sheet1.Range(cellrange).value
        'AuditAct_Info(i) = rangecells.Item(i).Value
        If isdropdownblank(JuridictionOfResidence(i)) Then
            MsgBx_General = MsgBx_General + ("Please enter Juridiction Of Residence at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            Validate_Juridiction = False
            Exit Function
        End If
 
    Next
End Function

Function Validate_Taxpayer() As Boolean
    Validate_Taxpayer = True
    Dim rangecells As Range
    'Dim hstatus As Variant
    Dim cellrange As Variant
    Dim SharesIndeterministic As Boolean
    
    setTblinfo_NRI
    setTblinfo_NRI1
    Set rangecells = Sheet1.Range("NRI_Taxpayer").Cells
    ReDim Taxpayer_IdentificationNumber(end_NRI)
    
    
    Dim i As Long
    For i = 1 To end_NRI
    
    cellrange = GetMergedAddressCell(rangecells, i)
        Taxpayer_IdentificationNumber(i) = Sheet1.Range(cellrange).value
        'AuditAct_Info(i) = rangecells.Item(i).Value
        If Not chkCompulsory(Taxpayer_IdentificationNumber(i)) Then
        'Added by Ayush 40
           ' MsgBx_General = MsgBx_General + ("Please enter Taxpayer Identification Number(s) at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            MsgBx_General = MsgBx_General + ("At least one Taxpayer Identification Number is mandatory in Part A General as you selected Non-resident Individual") & Chr(13)
          '  MsgBx_General = MsgBx_General + ("Please enter Taxpayer Identification Number(s) at Sr. No  " & i & "in schedule Part A General") & Chr(13)
            Validate_Taxpayer = False
            Exit Function
        End If
        
        If Len(Taxpayer_IdentificationNumber(i)) > 75 Then
            MsgBx_General = MsgBx_General + ("Taxpayer Identification Number(s) at Sr. No  " & i & " cannot exceed 75 characters in schedule Part A General") & Chr(13)
            Validate_Taxpayer = False
            Exit Function
        End If

    Next
End Function

Sub setTblinfo_NRI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("NRI_Juridiction").count
    Set rangecells = Sheet1.Range("NRI_Juridiction").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_NRI = ccount
    rngname_NRI = "NRI_Juridiction;NRI_Taxpayer;"
End Sub

Sub setTblinfo_NRI1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("NRI_Taxpayer").count
    Set rangecells = Sheet1.Range("NRI_Taxpayer").Cells
    For mIntCtr = 1 To mIntCells
        If Not Trim(rangecells.item(mIntCtr).value) = "" Then
            ccount = ccount + 1
        End If
    Next
    end_NRI1 = ccount
    rngname_NRI = "NRI_Juridiction;NRI_Taxpayer;"
End Sub

'--------------------------------------------Director Table----------------------------------------------

Sub setTblinfo_Director1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.CompanyName").count
    Set rangecells = Sheet1.Range("Sheet1.CompanyName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Director1 = ccount
    rngname_Director1 = "Sheet1.CompanyName;Sheet1.CompPAN;Sheet1.Shares;Sheet1.DIN;"
End Sub
Sub setTblinfo_DirectorComp()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.CompanyType").count
    Set rangecells = Sheet1.Range("Sheet1.CompanyType").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_DirectorComp = ccount
End Sub
Sub setTblinfo_Director2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.CompPAN").count
    Set rangecells = Sheet1.Range("Sheet1.CompPAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Director2 = ccount
End Sub

Sub setTblinfo_Director3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.Shares").count
    Set rangecells = Sheet1.Range("Sheet1.Shares").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_Director3 = ccount
End Sub

Sub setTblinfo_Director4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.DIN").count
    Set rangecells = Sheet1.Range("Sheet1.DIN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Director4 = ccount
End Sub

Function ValidateDirectorTable() As Boolean
     ValidateDirectorTable = True
     Dim i As Long
    setTblinfo_Director1
    setTblinfo_DirectorComp
    setTblinfo_Director2
    setTblinfo_Director3
    setTblinfo_Director4
    
    If Not ValidateNameofcompany_Director() Then ValidateDirectorTable = False
    If (Len(Sheet1.Range("Sheet1.CompanyName").item(1).value) > 0) Then
        
        If Not ValidateTypeofcompany_Director() Then ValidateDirectorTable = False
        If Not ValidatePAN_Director() Then ValidateDirectorTable = False
        If Not ValidateSharesStatus_Director() Then ValidateDirectorTable = False
        If Not ValidateDIN_Director() Then ValidateDirectorTable = False
    End If
    
    If UCase(Trim(Sheet1.Range("IsDirectorCompany").value)) = "YES" Then
        If end_Director1 = 0 Then
          '  MsgBx_General = MsgBx_General + "* Please enter Name of company in which you are director at Sr.No " & i & "schedule Part A General" & Chr(13)
          'MsgBx_General = MsgBx_General + "* Please enter Name of Company in Director detail table in Part A General" & Chr(13)
        'Chandru
        MsgBx_General = MsgBx_General + "*""Since 'Yes' is selected for the question ""Whether you were Director in a company at any time during the previous year?"", please furnish the details."" " & Chr(13)
           
            ValidateDirectorTable = False
        End If
    End If
    
    If (end_Director1 <> end_DirectorComp) Or (end_Director1 <> end_Director3) Then
            MsgBx_General = MsgBx_General + "* Enter all Mandatory Details in Table: Whether you are Director in a company? in schedule Part A General"" & Chr(13)"
            ValidateDirectorTable = False
    End If
End Function
Sub setTblinfo_Firm1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.NameOfFirm").count
    Set rangecells = Sheet1.Range("Sheet1.NameOfFirm").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Firm1 = ccount
    'rngname_Director1 = "Sheet1.CompanyName;Sheet1.CompPAN;Sheet1.Shares;Sheet1.DIN;"
End Sub

Sub setTblinfo_Firm2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.FirmPAN").count
    Set rangecells = Sheet1.Range("Sheet1.FirmPAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Firm2 = ccount
End Sub

Function ValidateFirmTable() As Boolean
     ValidateFirmTable = True
     Dim i As Long
    setTblinfo_Firm1
    setTblinfo_Firm2
    
    end_Firm1 = WorksheetFunction.Max(0, end_Firm1, end_Firm2)
    If UCase(Trim(Sheet1.Range("IsPartnerFirm").value)) = "YES" Then
        If end_Firm1 = 0 Then
            MsgBx_General = MsgBx_General + "* Please enter Name of Firm in which you are Partner at Sr.No " & i & " schedule Part A General" & Chr(13)
            ValidateFirmTable = False
        End If
    End If
    
    If Not ValidateNameofFirm() Then ValidateFirmTable = False
    If Not ValidatePAN_Firm() Then ValidateFirmTable = False
     
   
   
End Function

Function ValidateNameofFirm() As Boolean
    ValidateNameofFirm = True
   
    Dim rangecells As Range
    Dim cellrange As Variant
    
    Dim i As Long
   
    Set rangecells = Sheet1.Range("Sheet1.NameOfFirm").Cells
    ReDim FirmName(end_Firm1)
    
    For i = 1 To end_Firm1
        cellrange = GetMergedAddressCell(rangecells, i)
        FirmName(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(FirmName(i)) Then
            MsgBx_General = MsgBx_General + "* Please enter the name of the Firm(s) in which you are a partner at Sr.No " & i & "schedule Part A General" & Chr(13)
            ValidateNameofFirm = False
            Exit Function
        End If
        
        If Len(FirmName(i)) > 125 Then
            MsgBx_General = MsgBx_General + "* Name of the Firm(s) in which you are a partner in Schedule:should not excced 125 characters at Sr.No " & i & "schedule Part A General" & Chr(13)
            ValidateNameofFirm = False
            Exit Function
        End If
    
    UpdateProgressBar
    Next
End Function

Function ValidatePAN_Firm() As Boolean
    ValidatePAN_Firm = True
    'setTblinfo_Director1
    Dim rangecells As Range
    Dim cellrange As Variant
   
    Set rangecells = Sheet1.Range("Sheet1.FirmPAN").Cells
    ReDim Firm_PAN(end_Firm1)
    
    Dim i As Long
    
    For i = 1 To end_Firm1
        cellrange = GetMergedAddressCell(rangecells, i)
        Firm_PAN(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Firm_PAN(i)) Then
            MsgBx_General = MsgBx_General + "* Please enter PAN of the Firm(s) in which you were a partner at Sr.No " & i & "schedule Part A General" & Chr(13)
            ValidatePAN_Firm = False
            Exit Function
        End If
        
        If Len(Firm_PAN(i)) <> 10 Then
            MsgBx_General = MsgBx_General + "* Please enter a valid 10 digit PAN at Sr.No " & i & "schedule Part A General" & Chr(13)
            ValidatePAN_Firm = False
            Exit Function
        End If
        
        If Not CheckPAN(Firm_PAN(i)) Then
            MsgBx_General = MsgBx_General + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr.No " & i & "schedule Part A General" & Chr(13)
            ValidatePAN_Firm = False
            Exit Function
        End If
    
    UpdateProgressBar
    Next
End Function

Function ValidateNameofcompany_Director() As Boolean
    ValidateNameofcompany_Director = True
    setTblinfo_Director1
    Dim rangecells As Range
    Dim cellrange As Variant
    
    Dim i As Long
   
    Set rangecells = Sheet1.Range("Sheet1.CompanyName").Cells
    ReDim Director_NameofCompany(end_Director1)
    
    For i = 1 To end_Director1
        cellrange = GetMergedAddressCell(rangecells, i)
        Director_NameofCompany(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Director_NameofCompany(i)) Then
            'MsgBx_General = MsgBx_General + "* Please enter Name of company in which you are director  at Sr.No " & i & "in schedule Part A General" & Chr(13)
            
            'Shrutika
            MsgBx_General = MsgBx_General + "* ""Please enter Name of Company in Director detail table in Part A General""" & Chr(13)
            
            ValidateNameofcompany_Director = False
            Exit Function
        End If
        
        If Len(Director_NameofCompany(i)) > 125 Then
            MsgBx_General = MsgBx_General + "* Name of company in Schedule: Part A General should not excced 125 characters at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidateNameofcompany_Director = False
            Exit Function
        End If
    
    UpdateProgressBar
    Next
End Function
Function ValidateTypeofcompany_Director() As Boolean
    ValidateTypeofcompany_Director = True
    setTblinfo_Director1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
    Dim i As Long
    
    Set rangecells = Sheet1.Range("Sheet1.CompanyType").Cells
    ReDim Director_TypeofCompany(end_Director1)
    
    For i = 1 To end_Director1
        cellrange = GetMergedAddressCell(rangecells, i)
        Director_TypeofCompany(i) = (Sheet1.Range(cellrange).value)
        
        If isdropdownblank(Director_TypeofCompany(i)) Then
           ' MsgBx_General = MsgBx_General + "* Please select the dropdown whether company is Domestic or Foreign  at Sr.No " & i & "in schedule Part A General" & Chr(13)
           MsgBx_General = MsgBx_General + "* ""Please select the dropdown whether company is Domestic or Foreign in Director detail table in Part A General""" & Chr(13)
            ValidateTypeofcompany_Director = False
            Exit Function
        End If

        
    
    UpdateProgressBar
    Next
End Function

Function ValidatePAN_Director() As Boolean
    ValidatePAN_Director = True
    setTblinfo_Director1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
    Set rangecells = Sheet1.Range("Sheet1.CompPAN").Cells
    ReDim Director_PAN(end_Director1)
    
    Dim i As Long
    
    For i = 1 To end_Director1
        cellrange = GetMergedAddressCell(rangecells, i)
        Director_PAN(i) = UCase(Sheet1.Range(cellrange).value)
        
        If (Director_TypeofCompany(i)) = "Domestic" Then
            If Not chkCompulsory(Director_PAN(i)) Then
            MsgBx_General = MsgBx_General + "* Please enter PAN of company in which you are director  at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidatePAN_Director = False
            Exit Function
        End If
        End If
        
        
       If Director_PAN(i) <> "" Then
        If Len(Director_PAN(i)) <> 10 Then
            MsgBx_General = MsgBx_General + "* Please enter a valid 10 digit PAN  at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidatePAN_Director = False
            Exit Function
        End If
        
        If Not CheckPAN(Director_PAN(i)) Then
'            MsgBx_General = MsgBx_General + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr.No " & i & "in schedule Part A General" & Chr(13)
            MsgBx_General = MsgBx_General + "* ""Invalid PAN in Director detail table in Part A General. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.The fourth digit ( alphabet) should be C""" & Chr(13)

            ValidatePAN_Director = False
            Exit Function
        End If
        End If
    UpdateProgressBar
    Next
End Function

Function ValidateSharesStatus_Director() As Boolean
    ValidateSharesStatus_Director = True
    setTblinfo_Director1
    Dim rangecells As Range
    Dim cellrange As Variant
   
    Set rangecells = Sheet1.Range("Sheet1.Shares").Cells
    ReDim Director_SharesStatus(end_Director1)
     Dim i As Long
'     Dim i As Long
    For i = 1 To end_Director1
        cellrange = GetMergedAddressCell(rangecells, i)
        Director_SharesStatus(i) = UCase(Sheet1.Range(cellrange).value)
        
        If isdropdownblank(Director_SharesStatus(i)) Then
'            MsgBx_General = MsgBx_General + "* Please select the dropdown whether company is listed or unlisted at Sr.No " & i & "in schedule Part A General" & Chr(13)
           MsgBx_General = MsgBx_General + "* ""Please select the dropdown whether company is listed or unlisted in Director detail table in Part A General.""" & Chr(13)

            ValidateSharesStatus_Director = False
            Exit Function
        End If
        
    UpdateProgressBar
    Next
End Function

Function ValidateDIN_Director() As Boolean
    ValidateDIN_Director = True
    setTblinfo_Director1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Sheet1.DIN").Cells
    ReDim Director_DIN(end_Director1)
     Dim i As Long
    For i = 1 To end_Director1
        cellrange = GetMergedAddressCell(rangecells, i)
        Director_DIN(i) = UCase(Sheet1.Range(cellrange).value)
        
        If (Director_TypeofCompany(i)) = "Domestic" Then
        If Not chkCompulsory(Director_DIN(i)) Then
'            MsgBx_General = MsgBx_General + "* Please enter Director Identification Number (DIN) at Sr.No " & i & "in schedule Part A General" & Chr(13)
           MsgBx_General = MsgBx_General + "* ""Please enter Director Identification Number (DIN) in Director detail table in Part A General.""" & Chr(13)
 
            
            ValidateDIN_Director = False
            Exit Function
        End If
        End If
        
        If Director_DIN(i) <> "" Then
           If Not chkNumeric(Director_DIN(i)) Then
            MsgBx_General = MsgBx_General + "* Please enter valid 8 digit DIN at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidateDIN_Director = False
            Exit Function
        End If
        
         If Director_DIN(i) = "00000000" Then
            MsgBx_General = MsgBx_General + "* Please enter valid 8 digit DIN at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidateDIN_Director = False
            Exit Function
        End If
        
        If Len(Director_DIN(i)) <> 8 Then
            MsgBx_General = MsgBx_General + "* Please enter valid 8 digit DIN at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidateDIN_Director = False
            Exit Function
        End If
    End If
    UpdateProgressBar
    Next
End Function
'--------------------------------------------Shares Table----------------------------------------------
Sub setTblinfo_Shares1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_NameofCompany").count
    Set rangecells = Sheet1.Range("Share_NameofCompany").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares1 = ccount
    rngname_Shares1 = "Share_NameofCompany;Share_TypeofCompany;Share_PAN;Share_OP_Count;Share_OP_Cost;Share_Acquired_Count;Share_Acquired_Date;Share_Acquired_FaceValue;Share_Acquired_Issue;Share_Acquired_Cost;Share_Transferred_Count;Share_Transferred_Sale;Share_CB_Count;Share_CB_Cost;"
End Sub
Sub setTblinfo_SharesCompType()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_TypeofCompany").count
    Set rangecells = Sheet1.Range("Share_TypeofCompany").Cells
    For mIntCtr = 1 To mIntCells
        If isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_CompType = ccount
End Sub
Sub setTblinfo_Shares2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_PAN").count
    Set rangecells = Sheet1.Range("Share_PAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares2 = ccount
End Sub

Sub setTblinfo_Shares3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_OP_Count").count
    Set rangecells = Sheet1.Range("Share_OP_Count").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares3 = ccount
End Sub

Sub setTblinfo_Shares4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_OP_Cost").count
    Set rangecells = Sheet1.Range("Share_OP_Cost").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares4 = ccount
End Sub

Sub setTblinfo_Shares5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Acquired_Count").count
    Set rangecells = Sheet1.Range("Share_Acquired_Count").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares5 = ccount
End Sub

Sub setTblinfo_Shares6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Acquired_Date").count
    Set rangecells = Sheet1.Range("Share_Acquired_Date").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares6 = ccount
End Sub

Sub setTblinfo_Shares7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Acquired_FaceValue").count
    Set rangecells = Sheet1.Range("Share_Acquired_FaceValue").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares7 = ccount
End Sub
Sub setTblinfo_Shares8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Acquired_Issue").count
    Set rangecells = Sheet1.Range("Share_Acquired_Issue").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares8 = ccount
End Sub

Sub setTblinfo_Shares9()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Acquired_Cost").count
    Set rangecells = Sheet1.Range("Share_Acquired_Cost").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares9 = ccount
End Sub

Sub setTblinfo_Shares10()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Transferred_Count").count
    Set rangecells = Sheet1.Range("Share_Transferred_Count").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares10 = ccount
End Sub

Sub setTblinfo_Shares11()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_Transferred_Sale").count
    Set rangecells = Sheet1.Range("Share_Transferred_Sale").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares11 = ccount
End Sub

Sub setTblinfo_Shares12()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_CB_Count").count
    Set rangecells = Sheet1.Range("Share_CB_Count").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares12 = ccount
End Sub

Sub setTblinfo_Shares13()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Share_CB_Cost").count
    Set rangecells = Sheet1.Range("Share_CB_Cost").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Shares13 = ccount
End Sub

Function ValidateSharesTable() As Boolean

     ValidateSharesTable = True
     
    setTblinfo_Shares1
    setTblinfo_SharesCompType
    setTblinfo_Shares2
    setTblinfo_Shares3
    setTblinfo_Shares4
    setTblinfo_Shares5
    setTblinfo_Shares6
    setTblinfo_Shares7
    setTblinfo_Shares8
    setTblinfo_Shares9
    setTblinfo_Shares10
    setTblinfo_Shares11
    setTblinfo_Shares12
    setTblinfo_Shares13
    end_Shares1 = WorksheetFunction.Max(0, end_Shares1, end_CompType, end_Shares2, end_Shares3, end_Shares4, end_Shares5, end_Shares6, end_Shares7, end_Shares8, end_Shares9, end_Shares10, end_Shares11, end_Shares12, end_Shares13)
    end_Shares5 = WorksheetFunction.Max(0, end_Shares5, end_Shares6, end_Shares7, end_Shares8, end_Shares9)
    end_Shares10 = WorksheetFunction.Max(0, end_Shares10, end_Shares11)
    
    
    If Not ValidateNameofcompany_Shares() Then ValidateSharesTable = False
    If (Len(Sheet1.Range("Share_NameofCompany").item(1).value) > 0) Then
        If Not ValidateTypeofcompany_Shares() Then ValidateSharesTable = False
        If Not ValidatePAN_Shares() Then ValidateSharesTable = False
        If Not ValidateOP_Count_Shares() Then ValidateSharesTable = False
        If Not ValidateOP_Cost_Shares() Then ValidateSharesTable = False
        If Not ValidateTransferred_Shares() Then ValidateSharesTable = False
        If Not ValidateCB_Count_Shares() Then ValidateSharesTable = False
        If Not ValidateCB_Cost_Shares() Then ValidateSharesTable = False
        If Not ValidateAcquired_Shares() Then ValidateSharesTable = False
    End If
    
    
    If UCase(Trim(Sheet1.Range("IsShareOwner").value)) = "YES" Then
        If end_Shares1 = 0 Then
            MsgBx_General = MsgBx_General + "* Please enter Name of company in which you have held unlisted shares at any time during the previous year" & Chr(13)
            ValidateSharesTable = False
        End If
    End If

End Function

Function ValidateNameofcompany_Shares() As Boolean
    ValidateNameofcompany_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Share_NameofCompany").Cells
    ReDim Share_NameofCompany(end_Shares1)
    Dim i As Long
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_NameofCompany(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Share_NameofCompany(i)) Then
            MsgBx_General = MsgBx_General + "* Please enter Name of company in which you have held unlisted equity shares at any time during the previous year  at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidateNameofcompany_Shares = False
            Exit Function
        End If
        
        If Len(Share_NameofCompany(i)) > 125 Then
            MsgBx_General = MsgBx_General + "* Name of company in which you have held unlisted shares at any time during the previous year should not excced 125 characters at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidateNameofcompany_Shares = False
            Exit Function
        End If
    
    UpdateProgressBar
    Next
End Function
Function ValidateTypeofcompany_Shares() As Boolean
    ValidateTypeofcompany_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Share_TypeofCompany").Cells
    ReDim Share_TypeofCompany(end_Shares1)
    
    
    Dim i As Long
    For i = 1 To end_Shares1
        Share_TypeofCompany(i) = rangecells.item(i).value
        
        
        If isdropdownblank(Share_TypeofCompany(i)) Then
        'Commented by Shrutika
            'MsgBx_General = MsgBx_General + "* Please select the dropdown whether company is Domestic or Foreign at Sr.No " & i & "in schedule Part A General" & Chr(13)
            'MsgBx_General = MsgBx_General + "* Please select the dropdown whether company is Domestic or Foreign in Director detail table in Part A General Sr.No " & i & " "
            'Added by Shrutika
'            MsgBx_General = MsgBx_General + "* Please select the dropdown whether company is Domestic or Foreign in Director detail table in Part A General"
            MsgBx_General = MsgBx_General + "* Please select the dropdown whether company is Domestic or Foreign in Director detail table in Part A General" & Chr(13)

            
            ValidateTypeofcompany_Shares = False
            Exit Function
        End If
       
    Next
End Function

Function ValidatePAN_Shares() As Boolean
    ValidatePAN_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
    Dim rangecells1 As Range
    Dim cellrange1 As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Share_PAN").Cells
    ReDim Share_PAN(end_Shares1)
   
    
    Dim i As Long
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_PAN(i) = UCase(Sheet1.Range(cellrange).value)
        
        'Jyoti_SIT101629
        'Code modified by Shrutika
        If (Share_TypeofCompany(i)) = "Domestic" Then
            If Not chkCompulsory(Share_PAN(i)) Then
            'MsgBx_General = MsgBx_General + "* Please enter PAN of company in which you have held unlisted equity shares at any time during the previous year at Sr.No " & i & "in schedule Part A General" & Chr(13)
            
            MsgBx_General = MsgBx_General + """Please enter PAN of company in which you have held unlisted shares at any time during the previous year in part A General."""

            ValidatePAN_Shares = False
            Exit Function
        End If
        End If
       
       
If Share_PAN(i) <> "" Then
        If Len(Share_PAN(i)) <> 10 Then
            MsgBx_General = MsgBx_General + "* Please enter a valid 10 digit PAN at Sr.No " & i & "in schedule Part A General" & Chr(13)
            ValidatePAN_Shares = False
            Exit Function
        End If

        If Not CheckPAN(Share_PAN(i)) Then
'            MsgBx_General = MsgBx_General + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr.No " & i & "in schedule Part A General" & Chr(13)
            
          MsgBx_General = MsgBx_General + "* ""Invalid PAN in which you have held unlisted shares at any time during the previous year in part A General. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.""" & Chr(13)

            
            ValidatePAN_Shares = False
            Exit Function
        End If
        End If
    
    UpdateProgressBar
    Next
End Function

Function ValidateOP_Count_Shares() As Boolean
    ValidateOP_Count_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
    
   
    Set rangecells = Sheet1.Range("Share_OP_Count").Cells
    ReDim Share_OP_Count(end_Shares1)
    Dim i As Long
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_OP_Count(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Share_OP_Count(i)) Then
        
        'Ankita(23-Jan)
             '  MsgBx_General = MsgBx_General + "* Please enter opening balance of no. of shares at Sr.No " & i & "in schedule Part A General" & Chr(13)
             MsgBx_General = MsgBx_General + "* ""Please enter opening balance of no. of shares in which you have held unlisted shares at any time during the previous year in part A General""" & Chr(13)
            ValidateOP_Count_Shares = False
            Exit Function
        End If

    UpdateProgressBar
    Next
End Function

Function ValidateOP_Cost_Shares() As Boolean
    ValidateOP_Cost_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Share_OP_Cost").Cells
    ReDim Share_OP_Cost(end_Shares1)
    Dim i As Long
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_OP_Cost(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Share_OP_Cost(i)) Then
        ' Ankita(23-Jan)
           ' MsgBx_General = MsgBx_General + "* Please enter cost of acquisition for opening balance of unlisted shares at Sr.No " & i & "in schedule Part A.
            MsgBx_General = MsgBx_General + "* ""Please enter cost of acquisition for opening balance of unlisted shares in which you have held unlisted shares at any time during the previous year in part A General.""" & Chr(13)
            ValidateOP_Cost_Shares = False
            Exit Function
        End If

    UpdateProgressBar
    Next
End Function
Function ValidateTransferred_Shares() As Boolean
    ValidateTransferred_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As Variant
    Dim cellrange1 As Variant
    Dim i As Long
    Dim j As Long
   
    Set rangecells = Sheet1.Range("Share_Transferred_Count").Cells
    Set rangecells1 = Sheet1.Range("Share_Transferred_Sale").Cells
    ReDim Share_Transferred_Count(end_Shares1)
    ReDim Share_Transferred_Sale(end_Shares1)
    
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_Transferred_Count(i) = UCase(Sheet1.Range(cellrange).value)
        
        cellrange1 = GetMergedAddressCell(rangecells1, i)
        
        If Len(UCase(Sheet1.Range(cellrange1).value)) > 0 Then
            Share_Transferred_Sale(i) = Round(UCase(Sheet1.Range(cellrange1).value), 2)
        Else
             Share_Transferred_Sale(i) = UCase(Sheet1.Range(cellrange1).value)
        End If
        
        If chkCompulsory(Share_Transferred_Sale(i)) Then
            If Not chkCompulsory(Share_Transferred_Count(i)) Then
                'MsgBx_General = MsgBx_General + ("Please enter no. of shares transferred during the year in schedule Part A General.")
                
                'Shrutika
'                 MsgBx_General = MsgBx_General + ("* ""Please enter sale consideration of shares transferred during the year in which you have held unlisted shares at any time during the previous year in part A General.""")
                 MsgBx_General = MsgBx_General + ("* ""Please enter no. of shares transferred during the year in which you have held unlisted shares at any time during the previous year in part A General.""" & Chr(13))
                
                ValidateTransferred_Shares = False
                Exit Function
            End If
        End If
        
        
        If chkCompulsory(Share_Transferred_Count(i)) Then
            If Not chkCompulsory(Share_Transferred_Sale(i)) Then
            'changed by Chetan C M on 25/08/2025
            'start--
               'MsgBx_General = MsgBx_General + ("Please enter sale consideration of shares transferred during the year in schedule Part A General.")
                MsgBx_General = MsgBx_General + ("* ""Please enter sale consideration of shares transferred during the year in which you have held unlisted shares at any time during the previous year in part A General.""" & Chr(13))
            '--end
                ValidateTransferred_Shares = False
                Exit Function
            End If
        End If
        
        If Len(Share_Transferred_Count(i)) > 14 Then
            MsgBx_General = MsgBx_General + ("No. of shares transferred during the year in should not excced 14 digits in schedule Part A General.")
            ValidateTransferred_Shares = False
            Exit Function
        End If
        
        For j = 1 To Len(Share_Transferred_Count(i))
            If Not IsNumeric(Mid(Share_Transferred_Count(i), j, 1)) Then
                MsgBx_General = MsgBx_General + ("No. of shares transferred during the year in is Invalid in schedule Part A General")
                ValidateTransferred_Shares = False
                Exit Function
            End If
        Next
        
        If Len(Share_Transferred_Sale(i)) > 14 Then
            MsgBx_General = MsgBx_General + ("Sale consideration of shares transferred during the year should not excced 14 digits in schedule Part A General.")
            ValidateTransferred_Shares = False
            Exit Function
        End If
        
    Next
End Function
Function ValidateCB_Count_Shares() As Boolean
    ValidateCB_Count_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Share_CB_Count").Cells
    ReDim Share_CB_Count(end_Shares1)
    Dim i As Long
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_CB_Count(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Share_CB_Count(i)) Then
            'MsgBx_General = MsgBx_General + "* Please enter closing balance of no. of shares at Sr.No " & i & "in schedule Part A General" & Chr(13)
            
            'Shrutika
            'MsgBx_General = MsgBx_General + "* Please enter closing balance of no. of shares in which you have held unlisted shares at any time during the previous year in part A General" & Chr(13)
            'MsgBx_General = MsgBx_General + "* Please enter closing balance of no. of shares in which you have held unlisted shares at any time during the previous year in part A General Sr.No " & i & " & Chr(13)
             MsgBx_General = MsgBx_General + "* ""Please enter closing balance of no. of shares in which you have held unlisted shares at any time during the previous year in part A General.""" & Chr(13)
            
            ValidateCB_Count_Shares = False
            Exit Function
        End If
 
    UpdateProgressBar
    Next
End Function

Function ValidateCB_Cost_Shares() As Boolean
    ValidateCB_Cost_Shares = True
    setTblinfo_Shares1
    Dim rangecells As Range
    Dim cellrange As Variant
'    Dim i As Long
   
    Set rangecells = Sheet1.Range("Share_CB_Cost").Cells
    ReDim Share_CB_Cost(end_Shares1)
    Dim i As Long
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_CB_Cost(i) = UCase(Sheet1.Range(cellrange).value)
        
        If Not chkCompulsory(Share_CB_Cost(i)) Then
            'MsgBx_General = MsgBx_General + "* Please enter cost of acquisition for closing balance of unlisted shares at Sr.No " & i & "in schedule Part A General" & Chr(13)
            
            'Shrutika
             MsgBx_General = MsgBx_General + "* ""Please enter cost of acquisition for closing balance of unlisted shares in which you have held unlisted shares at any time during the previous year in part A General.""" & Chr(13)
            
            'MsgBx_General = MsgBx_General + "* Please enter cost of acquisition for closing balance of unlisted shares in which you have held unlisted shares at any time during the previous year in part A General Sr.No " & i & " & Chr(13)
            ValidateCB_Cost_Shares = False
            Exit Function
        End If

    UpdateProgressBar
    Next
End Function


Function ValidateAcquired_Shares() As Boolean
    ValidateAcquired_Shares = True
    setTblinfo_Shares1
    setTblinfo_Shares5
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim cellrange, cellrange1, cellrange2, cellRange3, cellrange4 As Variant
    Dim i As Long
    Dim j As Long
   
    Set rangecells = Sheet1.Range("Share_Acquired_Count").Cells
    Set rangecells1 = Sheet1.Range("Share_Acquired_Date").Cells
    Set rangecells2 = Sheet1.Range("Share_Acquired_FaceValue").Cells
    Set rangecells3 = Sheet1.Range("Share_Acquired_Issue").Cells
    Set rangecells4 = Sheet1.Range("Share_Acquired_Cost").Cells
    
    ReDim Share_Acquired_Count(end_Shares1)
    ReDim Share_Acquired_Date(end_Shares1)
    ReDim Share_Acquired_FaceValue(end_Shares1)
    ReDim Share_Acquired_Issue(end_Shares1)
    ReDim Share_Acquired_Cost(end_Shares1)
    
    For i = 1 To end_Shares1
        cellrange = GetMergedAddressCell(rangecells, i)
        Share_Acquired_Count(i) = UCase(Sheet1.Range(cellrange).value)
        
        cellrange1 = GetMergedAddressCell(rangecells1, i)
        Share_Acquired_Date(i) = UCase(Sheet1.Range(cellrange1).value)
        
        cellrange2 = GetMergedAddressCell(rangecells2, i)
        
        If Len(UCase(Sheet1.Range(cellrange2).value)) > 0 Then
            Share_Acquired_FaceValue(i) = Round(UCase(Sheet1.Range(cellrange2).value), 2)
        Else
            Share_Acquired_FaceValue(i) = UCase(Sheet1.Range(cellrange2).value)
        End If
        
        cellRange3 = GetMergedAddressCell(rangecells3, i)
        
        If Len(UCase(Sheet1.Range(cellRange3).value)) > 0 Then
            Share_Acquired_Issue(i) = Round(UCase(Sheet1.Range(cellRange3).value), 2)
        Else
            Share_Acquired_Issue(i) = UCase(Sheet1.Range(cellRange3).value)
        End If
        
        cellrange4 = GetMergedAddressCell(rangecells4, i)
        
        If Len(UCase(Sheet1.Range(cellrange4).value)) > 0 Then
            Share_Acquired_Cost(i) = Round(UCase(Sheet1.Range(cellrange4).value), 2)
        Else
            Share_Acquired_Cost(i) = UCase(Sheet1.Range(cellrange4).value)
        End If
' Change 14.05.2023.102 This change is done according to the Rules and DE because of issue raised in ***IPIP-38948***
'        If Not chkCompulsory(Share_Acquired_Count(i)) Then
'                MsgBx_General = MsgBx_General + ("Please enter no. of shares acquired during the year in which you have held unlisted shares at any time during the previous year in part A General")
'                ValidateAcquired_Shares = False
'        End If

        If chkCompulsory(Share_Acquired_Count(i)) Or chkCompulsory(Share_Acquired_Date(i)) Or chkCompulsory(Share_Acquired_FaceValue(i)) Or chkCompulsory(Share_Acquired_Cost(i)) Or chkCompulsory(Share_Acquired_Issue(i)) Then
            
            If Not chkCompulsory(Share_Acquired_Count(i)) Then
                MsgBx_General = MsgBx_General + ("* Please enter no. of shares acquired during the year in which you have held unlisted shares at any time during the previous year in part A General") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
            If Len(Share_Acquired_Count(i)) > 14 Then
                MsgBx_General = MsgBx_General + ("* No. of shares acquired during the year in Schedule: Part A General should not excced 14 digits. ") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
            For j = 1 To Len(Share_Acquired_Count(i))
                If Not IsNumeric(Mid(Share_Acquired_Count(i), j, 1)) Then
                    MsgBx_General = MsgBx_General + ("* No. of shares acquired during the year is Invalid in schedule Part A General") & Chr(13)
                    ValidateAcquired_Shares = False
                    Exit Function
                End If
            Next
            
            If Not chkCompulsory(Share_Acquired_Date(i)) Then
            'Ankita(23-Jan)
                '  MsgBx_General = MsgBx_General + ("Please enter date of subscription/purchase shares acquired during the year in schedule Part A General.") & Chr(13)
               MsgBx_General = MsgBx_General + ("* ""Please enter date of subscription/purchase in which you have held unlisted shares at any time during the previous year in part A General.""") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
           ' If Not ChkMaxDOBDate(Share_Acquired_Date(i), "31/03/2024") Then 'Chandru PAG Changes
        'Added by Shrutika(19-12-24)AY-2025
            'If Not ChkMaxDOBDate(Share_Acquired_Date(i), "31/03/2025") Then 'Chandru PAG Changes
            '-----------------------------------------'Added by Shrutika- 28/01/2026 AY-26-27
            Dim cutoff As Date
           cutoff = CDate(SheetALL.Range("Date_Filing1").value)
           If Not ChkMaxDOBDate(Share_Acquired_Date(i), SheetALL.Range("Date_Filing1").value) Then
               ' MsgBx_General = MsgBx_General + ("Please enter date of subscription/purchase shares acquired during the year before 31/03/2024 in schedule Part A General. ") & Chr(13) 'dpk1601
              '  MsgBx_General = MsgBx_General + ("* Please enter date of subscription/purchase shares acquired during the year before 31/03/2025 in schedule Part A General. ") & Chr(13) 'dpk1601
                  MsgBx_General = MsgBx_General + ("* Please enter date of subscription/purchase shares acquired during the year before  " & Dformat1(cutoff, "yyyy-mm-dd") & " in schedule Part A General.") & Chr(13) 'dpk1601
  
                ValidateAcquired_Shares = False
            End If
            '--------------------------------'Added by Shrutika- 28/01/2026 AY-26-27
            If Len(Share_Acquired_Date(i)) > 0 Then
                If Len(Share_Acquired_Date(i)) <> 10 Then
                    MsgBx_General = MsgBx_General + ("* Please enter valid date of subscription/purchase shares acquired during the year in schedule Part A General.") & Chr(13)
                    ValidateAcquired_Shares = False
                End If
            
                If Not FormatNCheckDate(Share_Acquired_Date(i)) Then
                    MsgBx_General = MsgBx_General + ("* Date of subscription/purchase shares acquired during the year must be in a valid dd/mm/yyyy format in schedule Part A General.") & Chr(13)
                    ValidateAcquired_Shares = False
                End If
            End If
            
            If Not chkCompulsory(Share_Acquired_FaceValue(i)) Then
            'Ankita(23-Jan)
                'MsgBx_General = MsgBx_General + ("Please enter face value per share in schedule Part A General.") & Chr(13)
                 MsgBx_General = MsgBx_General + ("* ""Please enter face value per share in which you have held unlisted shares at any time during the previous year in part A General.""") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
            If Len(Share_Acquired_FaceValue(i)) > 14 Then
                MsgBx_General = MsgBx_General + ("* Face value per share should not excced 14 digits in schedule Part A General.") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
            If Not chkCompulsory(Share_Acquired_Cost(i)) Then
                If Not chkCompulsory(Share_Acquired_Issue(i)) Then
                
                'sudha
                    'MsgBx_General = MsgBx_General + ("Please enter Issue price per share in schedule Part A General.") & Chr(13)
                     MsgBx_General = MsgBx_General + ("* Please enter Issue price per share in which you have held unlisted shares at any time during the previous year in part A General.") & Chr(13)
  
                'sudha
                    
                    ValidateAcquired_Shares = False
                End If
            End If
            
            If Len(Share_Acquired_Issue(i)) > 14 Then
                MsgBx_General = MsgBx_General + ("* Issue price per share (incase of fresh issue) should not excced 14 digits in schedule Part A General.") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
            If Not chkCompulsory(Share_Acquired_Issue(i)) Then
                If Not chkCompulsory(Share_Acquired_Cost(i)) Then
                
                'sudha
                    'MsgBx_General = MsgBx_General + ("Please enter Purchase price per share in schedule Part A General.") & Chr(13)
                    
                     MsgBx_General = MsgBx_General + ("* Please enter Purchase price per share in which you have held unlisted shares at any time during the previous year in part A General.") & Chr(13)
                'sudha
                
                    ValidateAcquired_Shares = False
                End If
            End If
            
            If Len(Share_Acquired_Cost(i)) > 14 Then
                MsgBx_General = MsgBx_General + ("* Purchase price per share (in case of purchase from existing shareholder) should not excced 14 digits in schedule Part A General.") & Chr(13)
                ValidateAcquired_Shares = False
            End If
            
            If chkCompulsory(Share_Acquired_Cost(i)) Then
                If chkCompulsory(Share_Acquired_Issue(i)) Then
'                    MsgBx_General = MsgBx_General + ("* Issue Price per share and Purchase price per share cannot filled together, Please provide information in different row in schedule Part A General.") & Chr(13)
                     MsgBx_General = MsgBx_General + ("* ""Please enter details of Issue Price & Purchase Price in separate Rows in which you have held unlisted shares at any time during the previous year in part A General""") & Chr(13)
'End Change***IPIP-38948***
                    ValidateAcquired_Shares = False
                End If
            End If
        End If
    Next
End Function

Sub BacYesValueChange()
On Error Resume Next
Dim answer As Integer
Dim msgString As String
msgString = "1) Certain allowances u/s section 10 (LTA, HRA, allowances  granted to meet expenses in performance of duties of office," & Chr(13) & _
            "Allowances granted to meet personal expenses in performance of duties of office, Allowance received by MP/MLA/MLC)" & Chr(13) & _
            "2)Deductions u/s 16 (Standard Deduction ,Entertainment allowance and Professional tax)" & Chr(13) & _
            "3)Interest payable on borrowed capital for self occupied property" & Chr(13) & _
            "4)Standard Deduction in case of family pension" & Chr(13) & _
            "5)Chapter VIA Deduction (life insurance, health insurance premium, pension funds, provident fund, donation etc except Contribution made by employer to notified pension scheme u/s 80CCD(2))" & Chr(13)

answer = MsgBox("Do you want to change value of Are you opting for new tax regime u/s 115BAC?" & Chr(13) & msgString, vbQuestion + vbOKCancel + vbDefaultButton2, "Confirmation")
If answer = vbOK Then
ThisWorkbook.Unprotect Password:=getmsgstate
  Sheet55.Unprotect Password:=getmsgstate
  Sheet55.Visible = xlSheetHidden
  Sheet55.Protect Password:=getmsgstate
'Change-6, ANK, 28.07.2022
  Sheet56.Unprotect Password:=getmsgstate
  Sheet56.Visible = xlSheetHidden
  Sheet56.Protect Password:=getmsgstate
'---end change---
  Sheet20.Unprotect Password:=getmsgstate
  Sheet20.Visible = xlSheetHidden
  Sheet20.Protect Password:=getmsgstate
'  Sheet18.Unprotect Password:=getmsgstate
'  Sheet18.Visible = xlSheetHidden
'  Sheet18.Protect Password:=getmsgstate
  Sheet50.Unprotect Password:=getmsgstate
  Sheet50.Visible = xlSheetHidden
  Sheet50.Protect Password:=getmsgstate
'  Sheet37.Unprotect Password:=getmsgstate
'  Sheet37.Visible = xlSheetHidden
'  Sheet37.Protect Password:=getmsgstate
  Sheet22.Unprotect Password:=getmsgstate
  Sheet22.Visible = xlSheetHidden
  Sheet22.Protect Password:=getmsgstate
  
  
  Sheet59.Unprotect Password:=getmsgstate
  Sheet59.Visible = xlSheetHidden
  Sheet59.Protect Password:=getmsgstate
  
  Sheet58.Unprotect Password:=getmsgstate
  Sheet58.Visible = xlSheetHidden
  Sheet58.Protect Password:=getmsgstate
  
   'Added by Shrutika(15/05/2025)
  Sheet60.Unprotect Password:=getmsgstate
  Sheet60.Visible = xlSheetHidden
  Sheet60.Protect Password:=getmsgstate
  '--------------------------------------
   'Added by Shrutika(24/05/2025)
  Sheet61.Unprotect Password:=getmsgstate
  Sheet61.Visible = xlSheetHidden
  Sheet61.Protect Password:=getmsgstate
  '--------------------------------------
  
  SheetALL.Unprotect Password:=getmsgstate
  SheetALL.Range("bacValue").value = 1
  SheetALL.Protect Password:=getmsgstate
'Change-07.03.2023.101 -> Salary issue
SalaryExemptAllowanceDropdown
'Newly added by Bindu as per V6 DE Sheet on 11th June 25
Application.EnableEvents = False
            SalaryExemptAllowance1013A

'---
  Sheet1.Unprotect Password:=getmsgstate
  resetBacYes
    Sheet14.LockUnlockDeduction57iia
             Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False

             Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate

Sheet1.Activate


  ThisWorkbook.Protect Password:=getmsgstate
Else
Sheet1.Unprotect Password:=getmsgstate
              
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'             Sheet1.Range("sheet1.capacityassessee").value = ""
'
'
             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
'             Sheet1.Range("acknowledgmentnumber").value = ""
Sheet1.Protect Password:=getmsgstate
Sheet1.Activate
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("bacValue").Clear
    SheetALL.Protect Password:=getmsgstate
    ActiveSheet.Shapes("BacYes").OLEFormat.Object.value = 0
     ActiveSheet.Shapes("BacNO").OLEFormat.Object.value = 0
     resetBacNo
     sheetVisibleForNo
     Sheet1.Activate
     Application.EnableEvents = True
  Exit Sub
End If

End Sub
Sub ITR3BacYesValueChange()
On Error Resume Next
Application.EnableEvents = False
Dim answer As Integer
answer = MsgBox("if you have opted for new tax Regime in AY 2021-22, please confirm ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer <> vbYes Then
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("oldbacValue").Clear
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
Else
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("oldbacValue").value = 1
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
End If
Sheet1.Activate


Application.EnableEvents = True
End Sub

Sub NewBacYesValueChange()
On Error Resume Next
Dim answer As Integer

If SheetALL.Range("oldbacValue").value = "" Then
SheetALL.Unprotect Password:=getmsgstate
SheetALL.Range("bacValue").Clear
SheetALL.Protect Password:=getmsgstate
Sheet1.Range("OptingNewTaxRegime").value = "(Select)"

MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2021-22?"
Application.EnableEvents = True
Exit Sub
End If

Dim pmsg As Variant
Dim pnew As Variant
'pmsg = "Since you are opting for New tax Regime u/s 115BAC. You will not be eligible to set-off House property loss and claim following deduction/allowances. Do you want to continue? "
pmsg = "“Since you are 'opting in Now' or 'continue to opt in' for New tax Regime u/s 115BAC. You will not be eligible to set-off House property loss and claim following deduction/allowances. Do you want to continue?"
pnew = Chr(10) + Chr(13)
pmsg = pmsg + pnew + "1. Certain allowances u/s section 10 (LTA, HRA, allowances  granted to meet expenses in performance of duties of office,Allowances granted to meet personal expenses in performance of duties of office, Allowance received by MP/MLA/MLC, Standard deduction in case of Minor child) "
pmsg = pmsg + pnew + "2. Deductions u/s 16 (Standard Deduction ,Entertainment allowance and Professional tax) "
pmsg = pmsg + pnew + "3. Interest payable on borrowed capital for self occupied property"
pmsg = pmsg + pnew + "4. Standard Deduction in case of family pension"
pmsg = pmsg + pnew + "5. Deduction u/s 10AA in respect of newly established Units in Special Economic Zones "
pmsg = pmsg + pnew + "6. Additional depreciation u/s 32(1) (iia)"
pmsg = pmsg + pnew + "7. Investment allowance in notified backward areas in certain States u/s 32AD"
pmsg = pmsg + pnew + "8. Deduction in respect of tea/coffee/rubber business u/s 33AB"
MsgBox pmsg
pmsg = ""
pmsg = pmsg + pnew + "9. Deduction in respect of business of prospecting for, or extraction or production of, petroleum or natural gas u/s 33ABA"
pmsg = pmsg + pnew + "10.Deduction in respect of expenditure on scientific research u/s 35(2AA) and 35(1)(ii)/35(1)(iia)/35(1)(iii) "
pmsg = pmsg + pnew + "11.Deduction in respect of expenditure on specified business u/s 35AD "
pmsg = pmsg + pnew + "12.Deduction in respect of expenditure on agricultural extension project u/s 35CCC "
pmsg = pmsg + pnew + "13.Chapter VIA Deduction (life insurance, health insurance premium, pension funds, provident fund, donation etc) except Contribution made by employer to notified pension scheme u/s 80CCD(2) & Deduction in respect of employment of new employees u/s 80JJAA "
pmsg = pmsg + pnew + "Do you want to proceed ?( Click Yes to Proceed)"

answer = MsgBox(pmsg, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

'answer = MsgBox("Do you want to change value of Are you opting for new tax regime u/s 115BAC?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
ThisWorkbook.Unprotect Password:=getmsgstate
  Sheet55.Unprotect Password:=getmsgstate
  Sheet55.Visible = xlSheetHidden
  Sheet55.Protect Password:=getmsgstate
  Sheet20.Unprotect Password:=getmsgstate
  Sheet20.Visible = xlSheetHidden
  Sheet20.Protect Password:=getmsgstate
  Sheet18.Unprotect Password:=getmsgstate
  Sheet18.Visible = xlSheetHidden
  Sheet18.Protect Password:=getmsgstate
  Sheet50.Unprotect Password:=getmsgstate
  Sheet50.Visible = xlSheetHidden
  Sheet50.Protect Password:=getmsgstate
  Sheet56.Unprotect Password:=getmsgstate
  Sheet56.Visible = xlSheetHidden
  Sheet56.Protect Password:=getmsgstate
  Sheet37.Unprotect Password:=getmsgstate
  Sheet37.Visible = xlSheetHidden
  Sheet37.Protect Password:=getmsgstate
  Sheet22.Unprotect Password:=getmsgstate
  Sheet22.Visible = xlSheetHidden
  Sheet22.Protect Password:=getmsgstate
  
  SheetALL.Unprotect Password:=getmsgstate
  SheetALL.Range("bacValue").value = 1 ' 'Chandru value changed to 2 from 1 (Ignore)
  SheetALL.Protect Password:=getmsgstate
  Sheet1.Unprotect Password:=getmsgstate
  resetBacYes
  SheetALL.Unprotect Password:=getmsgstate
If (SheetALL.Range("OldBacValue").value = 1 And SheetALL.Range("BacValue").value = 2) Or (SheetALL.Range("BacValue").value = 1) Then

             Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False

             Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
Else
             Sheet1.Range("sheet1.capacityassessee").value = ""
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
             
             Sheet1.Range("acknowledgmentnumber").MergeArea.value = ""
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             
End If
SheetALL.Protect Password:=getmsgstate
  
'             Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)  ', "AU"
'             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False
'
'             Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)  ', "AU"
'             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate

Sheet1.Activate


  ThisWorkbook.Protect Password:=getmsgstate
Else
Sheet1.Unprotect Password:=getmsgstate
SheetALL.Unprotect Password:=getmsgstate
If (SheetALL.Range("OldBacValue").value = 1 And SheetALL.Range("BacValue").value = 2) Or (SheetALL.Range("BacValue").value = 1) Then

             Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False

             Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
Else
             Sheet1.Range("sheet1.capacityassessee").value = ""
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
             Sheet1.Range("acknowledgmentnumber").value = ""
             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
             
End If
SheetALL.Protect Password:=getmsgstate

'             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
'             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'
'             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
'             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True

Sheet1.Protect Password:=getmsgstate
Sheet1.Activate
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("bacValue").Clear
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
     Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
     resetBacNo
     sheetVisibleForNo
  Exit Sub
End If
End Sub
Sub sheetVisibleForNo()
ThisWorkbook.Unprotect Password:=getmsgstate
 Sheet55.Unprotect Password:=getmsgstate
  Sheet55.Visible = xlSheetVisible
  Sheet55.Protect Password:=getmsgstate
  
  
  Sheet58.Unprotect Password:=getmsgstate
  Sheet58.Visible = xlSheetVisible
  Sheet58.Protect Password:=getmsgstate
  
  
  Sheet59.Unprotect Password:=getmsgstate
  Sheet59.Visible = xlSheetVisible
  Sheet59.Protect Password:=getmsgstate
  
  
  Sheet20.Unprotect Password:=getmsgstate
  Sheet20.Visible = xlSheetVisible
  Sheet20.Protect Password:=getmsgstate
   Sheet18.Unprotect Password:=getmsgstate
  'Sheet18.Visible = xlSheetVisible
  Sheet18.Protect Password:=getmsgstate
  Sheet50.Unprotect Password:=getmsgstate
  'Sheet50.Visible = xlSheetVisible
  Sheet50.Protect Password:=getmsgstate
  Sheet56.Unprotect Password:=getmsgstate
  Sheet56.Visible = xlSheetVisible
  Sheet56.Protect Password:=getmsgstate
  Sheet37.Unprotect Password:=getmsgstate
  'Sheet37.Visible = xlSheetVisible
  Sheet37.Protect Password:=getmsgstate
  Sheet22.Unprotect Password:=getmsgstate
  Sheet22.Visible = xlSheetVisible
  Sheet22.Protect Password:=getmsgstate
  
  'added by Chetan C M for AY 2025-26
  Sheet61.Unprotect Password:=getmsgstate
  'Sheet61.Visible = xlSheetHidden
  Sheet61.Visible = xlSheetVisible
  Sheet61.Protect Password:=getmsgstate
  
  
   'Added by Shrutika(15/05/2025)
  If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
  Sheet60.Unprotect Password:=getmsgstate
  Sheet60.Visible = xlSheetHidden
  Sheet60.Protect Password:=getmsgstate
  Else
  Sheet60.Unprotect Password:=getmsgstate
  Sheet60.Visible = xlSheetVisible
  Sheet60.Protect Password:=getmsgstate
  '--------------------------------------
  '--------------------------------------
  End If
  
   ThisWorkbook.Protect Password:=getmsgstate
End Sub
Sub sheetVisibleForNoold1()
ThisWorkbook.Unprotect Password:=getmsgstate
 Sheet55.Unprotect Password:=getmsgstate
  Sheet55.Visible = xlSheetVisible
  Sheet55.Protect Password:=getmsgstate
  Sheet20.Unprotect Password:=getmsgstate
  Sheet20.Visible = xlSheetVisible
  Sheet20.Protect Password:=getmsgstate
   Sheet18.Unprotect Password:=getmsgstate
  Sheet18.Visible = xlSheetVisible
  Sheet18.Protect Password:=getmsgstate
  Sheet50.Unprotect Password:=getmsgstate
  Sheet50.Visible = xlSheetVisible
  Sheet50.Protect Password:=getmsgstate
  Sheet56.Unprotect Password:=getmsgstate
  Sheet56.Visible = xlSheetVisible
  Sheet56.Protect Password:=getmsgstate
  Sheet37.Unprotect Password:=getmsgstate
  Sheet37.Visible = xlSheetVisible
  Sheet37.Protect Password:=getmsgstate
  Sheet22.Unprotect Password:=getmsgstate
  Sheet22.Visible = xlSheetVisible
  Sheet22.Protect Password:=getmsgstate
   ThisWorkbook.Protect Password:=getmsgstate
End Sub
Sub BacNoValueChange()
On Error Resume Next
Dim answer As Integer

'Malli---------23/10/2024
 
           Dim vdate_M1 As Variant
           vdate_M1 = Sheet52.Range("sheet9.Date").value
           
       If UCase(Mid(Trim(Sheet1.Range("sheet1.ReturnFileSec").value), 1, 6)) = "139(4)" And ChkMinInclusiveDate(Dformat(vdate_M1, "yyyy-mm-dd"), "2024-08-01") Then
         Sheet1.Shapes("BacNo").OLEFormat.Object.value = 1
         Exit Sub
       End If
 
 '--------------------------

 
answer = MsgBox("Do you want to change value of Are you opting for new tax regime u/s 115BAC?", vbQuestion + vbOKCancel + vbDefaultButton2, "Confirmation")
If answer = vbOK Then
ThisWorkbook.Unprotect Password:=getmsgstate
 sheetVisibleForNo
  SheetALL.Unprotect Password:=getmsgstate
  SheetALL.Range("bacValue").value = 2 'Chandru value changed to 1 from 2 (Ignore)
  SheetALL.Protect Password:=getmsgstate
'Change-07.03.2023.101 -> Salary issue
SalaryExemptAllowanceDropdown
'Newly added by Bindu as per V6 DE Sheet on 11th June 25
            SalaryExemptAllowance1013A

'---
  Sheet1.Unprotect Password:=getmsgstate
  resetBacNo
           Application.EnableEvents = False
                Sheet14.LockUnlockDeduction57iia
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
             
             
             Sheet1.Range("acknowledgmentnumber").value = ""
             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
             
             Sheet1.Range("sheet1.capacityassessee").value = ""
             Application.EnableEvents = True
            Sheet1.Protect Password:=getmsgstate

  Sheet1.Activate
  ThisWorkbook.Protect Password:=getmsgstate
Else
Sheet1.Unprotect Password:=getmsgstate
                Sheet14.LockUnlockDeduction57iia
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'             Sheet1.Range("sheet1.capacityassessee").value = ""

             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
'             Sheet1.Range("acknowledgmentnumber").value = ""
             
Sheet1.Protect Password:=getmsgstate
Sheet1.Activate
SheetALL.Unprotect Password:=getmsgstate
SheetALL.Range("bacValue").Clear
SheetALL.Protect Password:=getmsgstate
    ActiveSheet.Shapes("BacYes").OLEFormat.Object.value = 0
     ActiveSheet.Shapes("BacNO").OLEFormat.Object.value = 0
    resetBacNo
     sheetVisibleForNo
     Sheet1.Activate
  Exit Sub
End If
End Sub

Sub ITR3BacNoValueChange()
On Error Resume Next
Application.EnableEvents = False
Dim answer As Integer
answer = MsgBox("if you have not opted for new tax Regime in AY 2021-22, please confirm ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer <> vbYes Then
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("oldbacValue").Clear
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
    Else
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("oldbacValue").value = 2
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
End If
Sheet1.Activate


Application.EnableEvents = True
End Sub
Sub NewBacNoValueChange()
On Error Resume Next
Dim answer As Integer

If SheetALL.Range("oldbacValue").value = "" Then
SheetALL.Unprotect Password:=getmsgstate
SheetALL.Range("bacValue").Clear
SheetALL.Protect Password:=getmsgstate
Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2021-22?"
Application.EnableEvents = True
Exit Sub
End If
 
answer = MsgBox("Do you want to change value of 'Option for current assessment year'", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
ThisWorkbook.Unprotect Password:=getmsgstate
 sheetVisibleForNo
  SheetALL.Unprotect Password:=getmsgstate
  SheetALL.Range("bacValue").value = 2
  SheetALL.Protect Password:=getmsgstate
  Sheet1.Unprotect Password:=getmsgstate
  resetBacNo
           Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    SheetALL.Unprotect Password:=getmsgstate
If (SheetALL.Range("OldBacValue").value = 1 And SheetALL.Range("BacValue").value = 2) Or (SheetALL.Range("BacValue").value = 1) Then

             Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False

             Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
Else
             Sheet1.Range("sheet1.capacityassessee").value = ""
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
             Sheet1.Range("acknowledgmentnumber").value = ""
             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True


End If
SheetALL.Protect Password:=getmsgstate
Sheet1.Protect Password:=getmsgstate
           
'             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
'             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'
'
'             Sheet1.Range("acknowledgmentnumber").value = ""
'             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
'             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
'
'             Sheet1.Range("sheet1.capacityassessee").value = ""
             Application.EnableEvents = True
            Sheet1.Protect Password:=getmsgstate

  
  ThisWorkbook.Protect Password:=getmsgstate
  Sheet1.Activate
Else
Sheet1.Unprotect Password:=getmsgstate
SheetALL.Unprotect Password:=getmsgstate
If (SheetALL.Range("OldBacValue").value = 1 And SheetALL.Range("BacValue").value = 2) Or (SheetALL.Range("BacValue").value = 1) Then

             Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False

             Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)  ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
Else
             Sheet1.Range("sheet1.capacityassessee").value = ""
             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
             Sheet1.Range("acknowledgmentnumber").value = ""
             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True

End If
SheetALL.Protect Password:=getmsgstate

  
'             Sheet1.Range("sheet1.capacityassessee").Interior.ColorIndex = 15   ', "AU"
'             Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'
'             Sheet1.Range("acknowledgmentnumber").Interior.ColorIndex = 15   ', "AU"
'             Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True

             
Sheet1.Protect Password:=getmsgstate

Sheet1.Activate
SheetALL.Unprotect Password:=getmsgstate
SheetALL.Range("bacValue").Clear
SheetALL.Protect Password:=getmsgstate
    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
     Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
    resetBacNo
     sheetVisibleForNo
  Exit Sub
End If
End Sub
Sub resetBacNo()

resetESRForNo

resetDPMForNo

resetBPForNo

resetVIAForNo

resetCFLForYes
resetAMTCForNo

    Sheet15.Unprotect Password:=getmsgstate
    Sheet15.Range("UD.115BAC").Interior.ColorIndex = 15   ', "AU"
    Sheet15.Range("UD.115BAC").Locked = True
    Sheet15.Range("UD.115BAC").ClearContents
    Sheet15.Protect Password:=getmsgstate
    
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet39.Range("SAL.Entertainment_allowance_1").MergeArea.Locked = False
    
   '  Sheet1.Unprotect Password:=getmsgstate
   Sheet39.Unprotect Password:=getmsgstate
     Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet39.Range("SAL.DeductionUnderSection1").MergeArea.Locked = False
   Sheet39.Protect Password:=getmsgstate
   
   
'Commented by sai on 26/05/2025 as it was auto populated field and changed the logic as per DE V0.5 and V0.6
'***************************************************************************************************************
   
'   Dim HpCount As Long
'
'    For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value
'        Application.EnableEvents = False
'        sPassword = EfilingCommon.getmsgstate
'        Sheet8.Unprotect Password:=sPassword
'
'        If SheetALL.Range("bacValue").value = 1 Then
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
'
'        Else
'        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
'        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
'
'        End If
'        Application.EnableEvents = True
'        Sheet8.Protect Password:=sPassword
'    Next
'**************************************************************************************************************
' Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'   mIntCells = Range("Sheet1.Nature").Count
'    Set rangecells = Range("Sheet1.Nature").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "(Select)"
'    Next
'
'     mIntCells = Range("Sheet1.Amount").Count
'    Set rangecells = Range("Sheet1.Amount").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = ""
'    Next
'
' mIntCells = Range("Nature_Others_1").Count
'    Set rangecells = Range("Nature_Others_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "Not Applicable"
'    Next
    
    Range("SAL.Entertainment_allowance_1").value = 0
    Range("SAL.DeductionUnderSection1").value = 0
     
End Sub
Sub resetBacNoold1()
resetESRForNo
resetDPMForNo
resetBPForNo
resetVIAForNo
resetCFLForYes
    Sheet15.Unprotect Password:=getmsgstate
    Sheet15.Range("UD.115BAC").Interior.ColorIndex = 15   ', "AU"
    Sheet15.Range("UD.115BAC").Locked = True
    Sheet15.Range("UD.115BAC").ClearContents
    Sheet15.Protect Password:=getmsgstate
    
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet39.Range("SAL.Entertainment_allowance_1").MergeArea.Locked = False
    
   '  Sheet1.Unprotect Password:=getmsgstate
   Sheet39.Unprotect Password:=getmsgstate
     Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet39.Range("SAL.DeductionUnderSection1").MergeArea.Locked = False
   Sheet39.Protect Password:=getmsgstate
   
   Dim HpCount As Long

    For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value
        Application.EnableEvents = False
        sPassword = EfilingCommon.getmsgstate
        Sheet8.Unprotect Password:=sPassword
        
        If SheetALL.Range("bacValue").value = 1 Then
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
                
        Else
        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
                
        End If
        Application.EnableEvents = True
        Sheet8.Protect Password:=sPassword
    Next

' Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'   mIntCells = Range("Sheet1.Nature").Count
'    Set rangecells = Range("Sheet1.Nature").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "(Select)"
'    Next
'
'     mIntCells = Range("Sheet1.Amount").Count
'    Set rangecells = Range("Sheet1.Amount").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = ""
'    Next
'
' mIntCells = Range("Nature_Others_1").Count
'    Set rangecells = Range("Nature_Others_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "Not Applicable"
'    Next
    
    Range("SAL.Entertainment_allowance_1").value = 0
    Range("SAL.DeductionUnderSection1").value = 0
     
End Sub
Sub resetBacYes()

 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    resetESRForYes
    
    resetDPMForYes
    
    resetBPForYes
    resetVIAForYes
    resetCFLForNo
    resetAMTCForYes
    
    Sheet15.Unprotect Password:=getmsgstate
    Sheet15.Range("UD.115BAC").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet15.Range("UD.115BAC").Locked = False
     Sheet15.Protect Password:=getmsgstate
    
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.ColorIndex = 15   ', "AU"
    Sheet39.Range("SAL.Entertainment_allowance_1").MergeArea.Locked = True
    Sheet39.Range("SAL.Entertainment_allowance_1").value = 0
    
   '  Sheet1.Unprotect Password:=getmsgstate
   Sheet39.Unprotect Password:=getmsgstate
     Sheet39.Range("SAL.DeductionUnderSection1").Interior.ColorIndex = 15   ', "AU"
    Sheet39.Range("SAL.DeductionUnderSection1").MergeArea.Locked = True
    Sheet39.Range("SAL.DeductionUnderSection1").value = 0
   Sheet39.Protect Password:=getmsgstate

'Commented by sai on 26/05/2025 as it was auto populated field and changed the logic as per DE V0.5 and V0.6
'**********************************************************************************************************************
'   Dim HpCount As Long
'
'    For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value
'        Application.EnableEvents = False
'        sPassword = EfilingCommon.getmsgstate
'        Sheet8.Unprotect Password:=sPassword
'        If SheetALL.Range("bacValue").value = 1 Then
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
'        Else
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
'        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
'        End If
'        Application.EnableEvents = True
'        Sheet8.Protect Password:=sPassword
'    Next
'**********************************************************************************************************************
'    ccount = 0
'   mIntCells = Range("Sheet1.Nature").Count
'    Set rangecells = Range("Sheet1.Nature").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "(Select)"
'    Next
'
'     mIntCells = Range("Sheet1.Amount").Count
'    Set rangecells = Range("Sheet1.Amount").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = ""
'    Next
'
'         mIntCells = Range("Nature_Others_1").Count
'    Set rangecells = Range("Nature_Others_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "Not Applicable"
'    Next
'
    
   
   ' Range("IncD.Deduction16ia").Value = 0
End Sub
Sub resetBacYesold1()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    resetESRForYes
    resetDPMForYes
    resetBPForYes
    resetVIAForYes
    resetCFLForNo
    
    Sheet15.Unprotect Password:=getmsgstate
    Sheet15.Range("UD.115BAC").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet15.Range("UD.115BAC").Locked = False
     Sheet15.Protect Password:=getmsgstate
    
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.ColorIndex = 15   ', "AU"
    Sheet39.Range("SAL.Entertainment_allowance_1").MergeArea.Locked = True
    Sheet39.Range("SAL.Entertainment_allowance_1").value = 0
    
   '  Sheet1.Unprotect Password:=getmsgstate
   Sheet39.Unprotect Password:=getmsgstate
     Sheet39.Range("SAL.DeductionUnderSection1").Interior.ColorIndex = 15   ', "AU"
    Sheet39.Range("SAL.DeductionUnderSection1").MergeArea.Locked = True
    Sheet39.Range("SAL.DeductionUnderSection1").value = 0
   Sheet39.Protect Password:=getmsgstate
   
   Dim HpCount As Long

    For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value
        Application.EnableEvents = False
        sPassword = EfilingCommon.getmsgstate
        Sheet8.Unprotect Password:=sPassword
        If SheetALL.Range("bacValue").value = 1 Then
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
        Else
            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
        End If
        Application.EnableEvents = True
        Sheet8.Protect Password:=sPassword
    Next
   
'    ccount = 0
'   mIntCells = Range("Sheet1.Nature").Count
'    Set rangecells = Range("Sheet1.Nature").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "(Select)"
'    Next
'
'     mIntCells = Range("Sheet1.Amount").Count
'    Set rangecells = Range("Sheet1.Amount").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = ""
'    Next
'
'         mIntCells = Range("Nature_Others_1").Count
'    Set rangecells = Range("Nature_Others_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "Not Applicable"
'    Next
'
    
   
   ' Range("IncD.Deduction16ia").Value = 0
End Sub
Sub resetAMTCForYes()
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized1").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized1").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized1").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized2").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized2").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized2").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized3").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized3").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized3").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized5").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized5").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized5").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized4").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized4").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized4").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized6").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized6").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized6").ClearContents
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized7").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized7").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized7").ClearContents
    Sheet23.Protect Password:=getmsgstate
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized8").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized8").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized8").ClearContents
    Sheet23.Protect Password:=getmsgstate
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized9").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized9").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized9").ClearContents
    Sheet23.Protect Password:=getmsgstate
    
    'AMTC_C28_AY_2024-25 Add 'Malli
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized10").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized10").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized10").ClearContents
    Sheet23.Protect Password:=getmsgstate
    'AMTC.AmtCreditUtilized10
    
    
        'Added by Aavula Naresh New dev
    
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized11").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized11").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized11").ClearContents
    Sheet23.Protect Password:=getmsgstate
    
    ''''''
    'Added by Aavula Naresh for AY 2026-27 on 20/01/2026
    
        Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized12").Interior.ColorIndex = 15   ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized12").MergeArea.Locked = True
    Sheet23.Range("AMTC.AmtCreditUtilized12").ClearContents
    Sheet23.Protect Password:=getmsgstate
    
    
    
End Sub

Sub resetAMTCForNo()
    Sheet23.Unprotect Password:=getmsgstate
    Sheet23.Range("AMTC.AmtCreditUtilized1").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized1").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized2").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized2").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized3").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized3").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized5").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized5").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized4").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized4").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized6").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized6").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized7").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized7").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized8").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized8").Locked = False
    
    Sheet23.Range("AMTC.AmtCreditUtilized9").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized9").Locked = False
    'AMTC_C28_AY_2024-25 Add 'Malli
    Sheet23.Range("AMTC.AmtCreditUtilized10").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized10").Locked = False
    
    'Added by Aavula Naresh New dev
    
    Sheet23.Range("AMTC.AmtCreditUtilized11").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized11").Locked = False
    
    ''''
    'Added by Aavula Naresh for AY 2026-27 on 20/01/2026
    
    Sheet23.Range("AMTC.AmtCreditUtilized12").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet23.Range("AMTC.AmtCreditUtilized12").Locked = False
    ''''''''
    
    Sheet23.Protect Password:=getmsgstate
    
    
End Sub
Sub resetCFLForYes()

    Sheet17.Unprotect Password:=getmsgstate
    Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3b").Interior.ColorIndex = 15   ', "AU"
    Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3b").MergeArea.Locked = True
    Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4b").MergeArea.Locked = True
'    Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5b").MergeArea.Locked = True
'    Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5b").ClearContents
'
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6b").MergeArea.Locked = True
'    Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7b").MergeArea.Locked = True
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9b").MergeArea.Locked = True
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9b").MergeArea.Locked = True
'    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10b").MergeArea.Locked = True
'    Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10b").ClearContents
    
'    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11b").Interior.ColorIndex = 15   ', "AU"
'    Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11b").MergeArea.Locked = True
'    Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11b").ClearContents
    
    Sheet17.Protect Password:=getmsgstate
End Sub
Sub resetCFLForNo()
    Sheet17.Unprotect Password:=getmsgstate
'    Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3b").Locked = False
    
'    Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4b").Locked = False
    
'    Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5b").Locked = False
    
'    Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6b").Locked = False
    
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7b").Locked = False
    
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9b").Locked = False
    
'    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9b").Locked = False
    
'    Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10b").Locked = False
    
'    Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11b").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11b").Locked = False
    
    Sheet17.Protect Password:=getmsgstate
End Sub

Sub resetVIAForYes()
'change 9r 04.05.2022

    On Error Resume Next
    
 'sudha_2025-26
 
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80C").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80C").MergeArea.Locked = True
'    Range("scvia.Section80C").value = 0
    
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80CCC").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80CCC").MergeArea.Locked = True
'    Range("scvia.Section80CCC").value = 0
'
'changed by Chetan C M for AY 2026-27 as per DE V0.3
    'start--
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80CCD_SE").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80CCD_SE").MergeArea.Locked = True
'    Range("scvia.Section80CCD_SE").value = 0
 
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80CCD1B_SE").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80CCD1B_SE").MergeArea.Locked = True
'    Range("scvia.Section80CCD1B_SE").value = 0
 
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80CCD_SE").Interior.Color = (&HD8D8D8)
    Sheet19.Range("scvia.Section80CCD_SE").Locked = True
    Range("scvia.Section80CCD_SE").value = ""
 
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80CCD1B_SE").Interior.Color = (&HD8D8D8)
    Sheet19.Range("scvia.Section80CCD1B_SE").Locked = True
    Range("scvia.Section80CCD1B_SE").value = ""
    '--end

    Sheet19.Unprotect Password:=getmsgstate
'***************************************************************************************
    Sheet19.Range("Type_of_Identifier_80CCC").Interior.Color = (&HD8D8D8)
    Sheet19.Range("Type_of_Identifier_80CCC").Locked = True
    Range("Type_of_Identifier_80CCC").value = ""    'SIT-117077 by sadineni on 07/04/2026
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Name_of_Identifier_80CCC").Interior.Color = (&HD8D8D8)
    Sheet19.Range("Name_of_Identifier_80CCC").Locked = True
    Range("Name_of_Identifier_80CCC").value = ""    'SIT-117077 by sadineni on 07/04/2026
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Amount_80CCC").Interior.Color = (&HD8D8D8)
    Sheet19.Range("Amount_80CCC").Locked = True
    Range("Amount_80CCC").value = "" 'SIT-117077 by sadineni on 07/04/2026
    
    '****************commented by Chetan C M for AY 2026-27(V 0.3)********************
    '--start--
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("Type_of_Identifier_80CCD1").Interior.Color = (&HD8D8D8)
'    Sheet19.Range("Type_of_Identifier_80CCD1").Locked = True
'    Range("Type_of_Identifier_80CCD1").value = "" 'SIT-117077 by sadineni on 07/04/2026
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("Name_of_Identifier_80CCD1").Interior.Color = (&HD8D8D8)
'    Sheet19.Range("Name_of_Identifier_80CCD1").Locked = True
'    Range("Name_of_Identifier_80CCD1").value = "" 'SIT-117077 by sadineni on 07/04/2026
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("Amount_80CCD1").Interior.Color = (&HD8D8D8)
'    Sheet19.Range("Amount_80CCD1").Locked = True
'    Range("Amount_80CCD1").value = "" 'SIT-117077 by sadineni on 07/04/2026
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("Type_of_Identifier_80CCD1B").Interior.Color = (&HD8D8D8)
'    Sheet19.Range("Type_of_Identifier_80CCD1B").Locked = True
'    Range("Type_of_Identifier_80CCD1B").value = "" 'SIT-117077 by sadineni on 07/04/2026
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("Name_of_Identifier_80CCD1B").Interior.Color = (&HD8D8D8)
'    Sheet19.Range("Name_of_Identifier_80CCD1B").Locked = True
'    Range("Name_of_Identifier_80CCD1B").value = "" 'SIT-117077 by sadineni on 07/04/2026
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("Amount_80CCD1B").Interior.Color = (&HD8D8D8)
'    Sheet19.Range("Amount_80CCD1B").Locked = True
'    Range("Amount_80CCD1B").value = "" 'SIT-117077 by sadineni on 07/04/2026
    
    '************************************----end

'*******************************************************************************************
    Sheet19.Unprotect Password:=getmsgstate
    'Sheet19.Range("scvia.Section80DD").Interior.ColorIndex = 15   ', "AU"
    'Sheet19.Range("scvia.Section80DD").MergeArea.Locked = True
    'Range("scvia.Section80DD").value = 0
    
    'Sheet19.Unprotect Password:=getmsgstate
    'Sheet19.Range("SELECT80DD").Interior.ColorIndex = 15   ', "AU"
    'Sheet19.Range("SELECT80DD").MergeArea.Locked = True
    'Range("SELECT80DD").value = "(Select)"
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80DDB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80DDB").MergeArea.Locked = True
    Range("scvia.Section80DDB").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("SELECT80DDB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("SELECT80DDB").MergeArea.Locked = True
    Range("SELECT80DDB").value = "(Select)"
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.Specified_Disease").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Locked = True
    Range("Sheet19.Specified_Disease").value = "(Select)"
    
'Sudha_2025-26

    
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80E").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80E").MergeArea.Locked = True
'    Range("scvia.Section80E").value = 0
    
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80EE").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80EE").MergeArea.Locked = True
'    Range("scvia.Section80EE").value = 0
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80EEA").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80EEA").MergeArea.Locked = True
'    Range("scvia.Section80EEA").value = 0
'
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80EEB").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80EEB").MergeArea.Locked = True
'    Range("scvia.Section80EEB").value = 0
    
 '====================
 
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80GG").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80GG").MergeArea.Locked = True
    Range("scvia.Section80GG").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.AckNum").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("Sheet19.AckNum").MergeArea.Locked = True
    
    
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80GGC").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80GGC").MergeArea.Locked = True
'    Range("scvia.Section80GGC").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate

    Sheet19.Range("scvia.Section80TTA").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = True
    Range("scvia.Section80TTA").value = 0

    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80TTB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
    Range("scvia.Section80TTB").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
  '  Sheet19.Range("scvia.Section80U").Interior.ColorIndex = 15   ', "AU"
   ' Sheet19.Range("scvia.Section80U").MergeArea.Locked = True
   ' Range("scvia.Section80U").value = 0
            
    Sheet19.Unprotect Password:=getmsgstate
   ' Sheet19.Range("SELECT80U").Interior.ColorIndex = 15   ', "AU"
   ' Sheet19.Range("SELECT80U").MergeArea.Locked = True
   ' Range("SELECT80U").value = "(Select)"
            
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80JJA").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80JJA").MergeArea.Locked = True
'    Range("scvia.Section80JJA").value = 0
'
'      Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80IBA").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80IBA").MergeArea.Locked = True
'    Range("scvia.Section80IBA").value = 0
'
'    Sheet19.Range("scvia.Section80IAB").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80IAB").MergeArea.Locked = True
'    Range("scvia.Section80IAB").value = 0
    
     Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80QQB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80QQB").MergeArea.Locked = True
    Range("scvia.Section80QQB").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.AckNum10CCD").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("Sheet19.AckNum10CCD").MergeArea.Locked = True
    
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80RRB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80RRB").MergeArea.Locked = True
    Range("scvia.Section80RRB").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.AckNum.10CCE").Interior.ColorIndex = 15   ',
    Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = True
    
    
'Sudha_2025-26

    
'    Sheet19.Unprotect Password:=getmsgstate
'      Sheet19.Range("scvia.Section80EE").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80EE").MergeArea.Locked = True
End Sub

Sub resetVIAForNo()
'change-10r 04.05.2022

    On Error Resume Next
    'sudha_2025-26
    
      Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80C").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet19.Range("scvia.Section80C").Locked = False

'    Sheet19.Range("scvia.Section80CCC").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet19.Range("scvia.Section80CCC").Locked = False
'
'changed by Chetan C M for AY 2026-27 as per DE V0.3
    'start--
'    Sheet19.Range("scvia.Section80CCD_SE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet19.Range("scvia.Section80CCD_SE").Locked = False
 
'    Sheet19.Range("scvia.Section80CCD1B_SE").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet19.Range("scvia.Section80CCD1B_SE").Locked = False
 
    Sheet19.Range("scvia.Section80CCD_SE").Interior.Color = (&HCCFFCC)
    Sheet19.Range("scvia.Section80CCD_SE").Locked = False
 
    Sheet19.Range("scvia.Section80CCD1B_SE").Interior.Color = (&HCCFFCC)
    Sheet19.Range("scvia.Section80CCD1B_SE").Locked = False
 
    '--end
    
'**********************************************************************************************
        Sheet19.Range("Type_of_Identifier_80CCC").Interior.Color = (&HCCFFCC)
        Sheet19.Range("Type_of_Identifier_80CCC").Locked = False
        
        Sheet19.Range("Name_of_Identifier_80CCC").Interior.Color = (&HCCFFCC)
        Sheet19.Range("Name_of_Identifier_80CCC").Locked = False
        
        Sheet19.Range("Amount_80CCC").Interior.Color = (&HCCFFCC)
        Sheet19.Range("Amount_80CCC").Locked = False
        
    '****************commented by Chetan C M for AY 2026-27(V 0.3)********************
    '--start--
'        Sheet19.Range("Type_of_Identifier_80CCD1").Interior.Color = (&HCCFFCC)
'        Sheet19.Range("Type_of_Identifier_80CCD1").Locked = False
'
'        Sheet19.Range("Name_of_Identifier_80CCD1").Interior.Color = (&HCCFFCC)
'        Sheet19.Range("Name_of_Identifier_80CCD1").Locked = False
'
'        Sheet19.Range("Amount_80CCD1").Interior.Color = (&HCCFFCC)
'        Sheet19.Range("Amount_80CCD1").Locked = False
'
'        Sheet19.Range("Type_of_Identifier_80CCD1B").Interior.Color = (&HCCFFCC)
'        Sheet19.Range("Type_of_Identifier_80CCD1B").Locked = False
'
'        Sheet19.Range("Name_of_Identifier_80CCD1B").Interior.Color = (&HCCFFCC)
'        Sheet19.Range("Name_of_Identifier_80CCD1B").Locked = False
'
'        Sheet19.Range("Amount_80CCD1B").Interior.Color = (&HCCFFCC)
'        Sheet19.Range("Amount_80CCD1B").Locked = False
    '----end************************************
    
'********************************************************************************************
    

    'Sheet19.Range("scvia.Section80DD").Interior.Color = (&HCCFFCC)  ', "AU"
    'Sheet19.Range("scvia.Section80DD").Locked = False
    
    Sheet19.Unprotect Password:=getmsgstate
    'Sheet19.Range("SELECT80DD").MergeArea.Locked = False
    'Sheet19.Range("SELECT80DD").Interior.Color = (&HCCFFCC)  ', "AU"
    'Range("SELECT80DD").value = "(Select)"
    
    Sheet19.Unprotect Password:=getmsgstate
    
    Sheet19.Range("scvia.Section80DDB").MergeArea.Locked = False
    Sheet19.Range("scvia.Section80DDB").Interior.Color = (&HCCFFCC)   ', "AU"
    
    Sheet19.Unprotect Password:=getmsgstate
    
    Sheet19.Range("SELECT80DDB").MergeArea.Locked = False
    Sheet19.Range("SELECT80DDB").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet19.Range("SELECT80DDB").value = "(Select)"
    
    
    Sheet19.Unprotect Password:=getmsgstate
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Locked = False
    Sheet19.Range("Sheet19.Specified_Disease").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet19.Range("Sheet19.Specified_Disease").value = "(Select)"
    
    If Sheet1.Range("sheet1.Status").value = "H-HUF" Then
    
        'Sudha_23/05/2025
        Sheet19.Unprotect Password:=getmsgstate
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.ClearContents
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Locked = True
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("Sheet19.Specified_Disease").value = "(Select)"
        
'sudha_2025-26

'    Sheet19.Range("scvia.Section80E").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80E").MergeArea.Locked = True
'    Range("scvia.Section80E").value = 0
       
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80EE").Interior.ColorIndex = 15   ', "AU"
'    Sheet19.Range("scvia.Section80EE").MergeArea.Locked = True
'    Range("scvia.Section80EE").value = 0
    
    Else
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80E").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet19.Range("scvia.Section80E").Locked = False
'
'    Sheet19.Unprotect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80EE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet19.Range("scvia.Section80EE").Locked = False
'    End If
'    Sheet1.Protect Password:=getmsgstate
'    Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet19.Range("scvia.Section80EEA").Locked = False
'
'    Sheet19.Range("scvia.Section80EEB").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet19.Range("scvia.Section80EEB").Locked = False
 '====================
 
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80GG").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet19.Range("scvia.Section80GG").Locked = False
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet19.Range("Sheet19.AckNum").MergeArea.Locked = False
      Sheet19.Range("Sheet19.AckNum").Font.Color = vbBlack  'Shrutika(17/06/25)
    
    
'    Sheet19.Range("scvia.Section80GGC").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet19.Range("scvia.Section80GGC").Locked = False
    
 'Malli comented-------------SIT_90188
''Change-08.06.2023.101 -> SIT-53860
'    If calculateAge(Trim(Range("sheet1.DOB").value)) < 60 Then
'        Sheet19.Range("scvia.Section80TTA").Interior.Color = (&HCCFFCC)  ', "AU"
'        Sheet19.Range("scvia.Section80TTA").Locked = False
'    End If
''----
'
''Change-08.06.2023.101 -> SIT-53860
'    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 Then
'        Sheet19.Range("scvia.Section80TTB").Interior.Color = (&HCCFFCC) ', "AU"
'        Sheet19.Range("scvia.Section80TTB").Locked = False
'    End If
''---

'Malli----------SIT_90188

If AY23_24Changes.calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 Then
                
               ' If Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then
                 If (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" And Mid(Range("sheet1.Status").value, 1, 1) = "I") Or Mid(Range("sheet1.Status").value, 1, 1) = "H" Then
                    Sheet19.Unprotect Password:=getmsgstate
                    Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = False
                    Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                Else
                    Sheet19.Range("scvia.Section80TTA").value = ""
                    Sheet19.Unprotect Password:=getmsgstate
                    Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = True
                    Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)
                End If
                
                If Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Range("sheet1.Status").value, 1, 1) = "H" Then
                    Sheet19.Range("scvia.Section80TTB").value = ""
                    Sheet19.Unprotect Password:=getmsgstate
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
                Else
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = False
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                End If
                
        ElseIf AY23_24Changes.calculateAge(Trim(Range("sheet1.DOB").value)) <= 60 Then
                    Sheet19.Unprotect Password:=getmsgstate
                    Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = False
                    Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                    Sheet19.Unprotect Password:=getmsgstate
                    Sheet19.Range("scvia.Section80TTB").value = ""
                    Sheet19.Unprotect Password:=getmsgstate
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
                
        End If
        End If
        
'---------------------------------------------------


  '  Sheet19.Range("scvia.Section80U").Interior.Color = (&HCCFFCC) ', "AU"
   ' Sheet19.Range("scvia.Section80U").Locked = False
    
  '  Sheet19.Range("SELECT80U").MergeArea.Locked = False
   ' Sheet19.Range("SELECT80U").Interior.Color = (&HCCFFCC)  ', "AU"
   ' Sheet19.Range("SELECT80U").value = "(Select)"
    
'    Sheet19.Range("scvia.Section80JJA").Interior.Color = (&HCCFFCC) ', "AU"
'    Sheet19.Range("scvia.Section80JJA").Locked = False
    
'     Sheet19.Range("scvia.Section80IBA").Interior.Color = (&HCCFFCC) ', "AU"
'     Sheet19.Range("scvia.Section80IBA").Locked = False
'
'    Sheet19.Range("scvia.Section80IAB").Interior.Color = (&HCCFFCC) ', "AU"
'     Sheet19.Range("scvia.Section80IAB").Locked = False
    
    If Sheet1.Range("sheet1.ResidentialStatus1").value = "NRI - Non Resident" Then
    Sheet19.Unprotect Password:=getmsgstate
    
    Sheet19.Range("scvia.Section80QQB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80QQB").MergeArea.Locked = True
    Range("scvia.Section80QQB").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.AckNum10CCD").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("Sheet19.AckNum10CCD").MergeArea.Locked = True
    
    
      Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("scvia.Section80RRB").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("scvia.Section80RRB").MergeArea.Locked = True
    Range("scvia.Section80RRB").value = 0
    
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Sheet19.AckNum.10CCE").Interior.ColorIndex = 15   ', "AU"
    Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = True
    
    Else
    Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet19.Range("scvia.Section80QQB").Locked = False
    
        Sheet19.Range("Sheet19.AckNum10CCD").Interior.Color = (&HCCFFCC) ', "AU"
       Sheet19.Range("Sheet19.AckNum10CCD").MergeArea.Locked = False
       Sheet19.Range("Sheet19.AckNum10CCD").Font.Color = vbBlack   'Shrutika(16/06/25)
     
    Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet19.Range("scvia.Section80RRB").Locked = False
    
    Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = False
    Sheet19.Range("Sheet19.AckNum.10CCE").Font.Color = vbBlack  'Shrutika(16/06/25)
    
    End If
End Sub
Sub resetBPForYes()

    Sheet9.Unprotect Password:=getmsgstate
    Sheet9.Range("sheet11.DeductUs32AD").Interior.ColorIndex = 15   ', "AU"
    Sheet9.Range("sheet11.DeductUs32AD").MergeArea.Locked = True
    Sheet9.Range("sheet11.DeductUs32AD").value = 0
   
    Sheet9.Unprotect Password:=getmsgstate
    Sheet9.Range("sheet12.DeductUs35AD").Interior.ColorIndex = 15   ', "AU"
    Sheet9.Range("sheet12.DeductUs35AD").MergeArea.Locked = True
    Sheet9.Range("sheet12.DeductUs35AD").value = 0
    
    Sheet9.Unprotect Password:=getmsgstate
    
    Sheet9.Range("EI.Drpdn").Interior.ColorIndex = 15   ', "AU"
   
    Sheet9.Range("EI.Drpdn").Locked = True
   
    Sheet9.Range("EI.Drpdn").value = "(Select)"
    
End Sub
Sub resetBPForNo()

    Sheet9.Unprotect Password:=getmsgstate
'    Sheet9.Range("sheet11.DeductUs32AD").Interior.Color = (&HCCFFCC)    ', "AU"
'    Sheet9.Range("sheet11.DeductUs32AD").MergeArea.Locked = False

        
    Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet9.Range("sheet12.DeductUs35AD").MergeArea.Locked = False
   
    Sheet9.Range("EI.Drpdn").Interior.Color = (&HCCFFCC)    ', "AU"
    
    Sheet9.Range("EI.Drpdn").Locked = False
    
    Sheet9.Range("EI.Drpdn").value = "(Select)"
    
End Sub

Sub resetESRForYes()
    Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR1ii.AmtDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR1ii.AmtDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR1ii.AmtDebPL").value = 0
     Sheet12.Unprotect Password:=getmsgstate
    Sheet12.Range("ESR1ii.AmtUs35Allowable").Interior.ColorIndex = 15   ', "AU"
    Sheet12.Range("ESR1ii.AmtUs35Allowable").MergeArea.Locked = True
    Sheet12.Range("ESR1ii.AmtUs35Allowable").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR1iia.AmtDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR1iia.AmtDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR1iia.AmtDebPL").value = 0
     Sheet12.Unprotect Password:=getmsgstate
    Sheet12.Range("ESR1iia.AmtUs35Allowable").Interior.ColorIndex = 15   ', "AU"
    Sheet12.Range("ESR1iia.AmtUs35Allowable").MergeArea.Locked = True
    Sheet12.Range("ESR1iia.AmtUs35Allowable").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR1iii.AmtDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR1iii.AmtDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR1iii.AmtDebPL").value = 0
     Sheet12.Unprotect Password:=getmsgstate
    Sheet12.Range("ESR1iii.AmtUs35Allowable").Interior.ColorIndex = 15   ', "AU"
    Sheet12.Range("ESR1iii.AmtUs35Allowable").MergeArea.Locked = True
    Sheet12.Range("ESR1iii.AmtUs35Allowable").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR2AA.AmtDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR2AA.AmtDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR2AA.AmtDebPL").value = 0
     Sheet12.Unprotect Password:=getmsgstate
    Sheet12.Range("ESR2AA.AmtUs35Allowable").Interior.ColorIndex = 15   ', "AU"
    Sheet12.Range("ESR2AA.AmtUs35Allowable").MergeArea.Locked = True
    Sheet12.Range("ESR2AA.AmtUs35Allowable").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR35CCC.AmtDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR35CCC.AmtDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR35CCC.AmtDebPL").value = 0
     Sheet12.Unprotect Password:=getmsgstate
    Sheet12.Range("ESR35CCC.AmtUs35Allowable").Interior.ColorIndex = 15   ', "AU"
    Sheet12.Range("ESR35CCC.AmtUs35Allowable").MergeArea.Locked = True
    Sheet12.Range("ESR35CCC.AmtUs35Allowable").value = ""
     Sheet12.Unprotect Password:=getmsgstate
'    Sheet12.Range("ESR35CCC.ExcessAmtOverDebPL").Interior.ColorIndex = 15   ', "AU"
'    Sheet12.Range("ESR35CCC.ExcessAmtOverDebPL").MergeArea.Locked = True
'    Sheet12.Range("ESR35CCC.ExcessAmtOverDebPL").value = ""
    Sheet12.Protect Password:=getmsgstate
End Sub

Sub resetESRForNo()
    Sheet12.Unprotect Password:=getmsgstate
    Sheet12.Range("ESR1ii.AmtDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR1ii.AmtDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR1ii.AmtUs35Allowable").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR1ii.AmtUs35Allowable").MergeArea.Locked = False
'    Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
'    Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR1iia.AmtDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR1iia.AmtDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR1iia.AmtUs35Allowable").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR1iia.AmtUs35Allowable").MergeArea.Locked = False
'    Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
'    Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR1iii.AmtDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR1iii.AmtDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR1iii.AmtUs35Allowable").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR1iii.AmtUs35Allowable").MergeArea.Locked = False
'    Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
'    Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR2AA.AmtDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR2AA.AmtDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR2AA.AmtUs35Allowable").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR2AA.AmtUs35Allowable").MergeArea.Locked = False
'    Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
'    Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR35CCC.AmtDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR35CCC.AmtDebPL").MergeArea.Locked = False
    Sheet12.Range("ESR35CCC.AmtUs35Allowable").Interior.Color = (&HCCFFCC)    ', "AU"
    Sheet12.Range("ESR35CCC.AmtUs35Allowable").MergeArea.Locked = False
'    Sheet12.Range("ESR35CCC.ExcessAmtOverDebPL").Interior.Color = (&HCCFFCC)    ', "AU"
'    Sheet12.Range("ESR35CCC.ExcessAmtOverDebPL").MergeArea.Locked = False
    Sheet12.Protect Password:=getmsgstate
End Sub
Sub resetDPMForYes()
   Application.EnableEvents = False
   Sheet10.Range("DPM45.RATE").Comment.Visible = True
   Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").value = ""
    
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").value = ""
    
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").value = ""
    
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").value = ""
    
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").value = ""
    
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").value = ""
    
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").value = ""
    
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").value = ""
    
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").value = ""
    
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").value = ""
    
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").value = ""
    
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").value = ""
    
    Sheet10.Range("DPM45.AdditionsGrThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsGrThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsGrThan180Days").value = ""
    
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsLessThan180Days").value = ""
    
    Sheet10.Range("DPM15.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM15.WDVFirstDayBAC").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM30.WDVFirstDayBAC").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM40.WDVFirstDayBAC").MergeArea.Locked = False
    
    'Sheet9.Range("DPM45.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM45.WDVFirstDayBAC").MergeArea.Locked = True
    
'    Sheet10.Range("DPM45.WDVFirstDay").Interior.ColorIndex = 15   ', "AU"
'    Sheet10.Range("DPM45.WDVFirstDay").MergeArea.Locked = True
'    Sheet10.Range("DPM45.WDVFirstDay").value = ""
    
    Sheet10.Range("DPM45.WDVFirstDay3").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.WDVFirstDay3").MergeArea.Locked = True
    Sheet10.Range("DPM45.WDVFirstDay3").value = ""
    
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsLessThan180Days").value = ""
    Sheet10.Range("DPM45.AdditionsGrThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsGrThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsGrThan180Days").value = ""
    Sheet10.Range("DPM45.RealizationTotalPeriod").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.RealizationTotalPeriod").MergeArea.Locked = True
    Sheet10.Range("DPM45.RealizationTotalPeriod").value = ""
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsLessThan180Days").value = ""
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").MergeArea.Locked = True
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").value = ""
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").value = ""
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").value = ""
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").value = ""
    Sheet10.Range("DPM45.DepUnderSec38").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.DepUnderSec38").MergeArea.Locked = True
    Sheet10.Range("DPM45.DepUnderSec38").value = ""
    Sheet10.Range("DPM45.ProAgrdep").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.ProAgrdep").MergeArea.Locked = True
    Sheet10.Range("DPM45.ProAgrdep").value = ""
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").MergeArea.Locked = True
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").value = ""
    Sheet10.Range("DPM45.CapGainUs50").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.CapGainUs50").MergeArea.Locked = True
    Sheet10.Range("DPM45.CapGainUs50").value = ""
'    Sheet9.Range("I7:I9").Interior.ColorIndex = 15   ', "AU"
'    Sheet9.Range("I7:I9").Locked = True
'    Sheet9.Range("I11:I12").Interior.ColorIndex = 15   ', "AU"
'    Sheet9.Range("I11:I12").Locked = True
'    Sheet9.Range("I14:I15").Interior.ColorIndex = 15   ', "AU"
'    Sheet9.Range("I14:I15").Locked = True
'     Sheet9.Range("I19:I27").Interior.ColorIndex = 15   ', "AU"
'    Sheet9.Range("I19:I27").Locked = True
    Sheet10.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub
Sub resetDPMForNo()
Sheet10.Range("DPM45.RATE").Comment.Visible = False
            Application.EnableEvents = True
    Sheet10.Unprotect Password:=getmsgstate
'    Sheet9.Range("I7:I9").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet9.Range("I7:I9").Locked = False
'    Sheet9.Range("I11:I12").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet9.Range("I11:I12").Locked = False
'    Sheet9.Range("I14:I15").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet9.Range("I14:I15").Locked = False
'     Sheet9.Range("I19:I27").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet9.Range("I19:I27").Locked = False
'    Sheet10.Range("DPM45.WDVFirstDay").Interior.Color = (&HCCFFCC)     ', "AU"
'    Sheet10.Range("DPM45.WDVFirstDay").MergeArea.Locked = False
    Sheet10.Range("DPM45.WDVFirstDay3").Interior.Color = (&HCCFFCC)     ', "AU"
    Sheet10.Range("DPM45.WDVFirstDay3").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.Color = (&HD8D8D8)  ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsGrThan180Days").Interior.Color = (&HD8D8D8)  ', "AU"
    Sheet10.Range("DPM45.AdditionsGrThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.RealizationTotalPeriod").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.RealizationTotalPeriod").MergeArea.Locked = False
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.Color = (&HD8D8D8)   ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").Interior.Color = (&HD8D8D8)     ', "AU"
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HD8D8D8)   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").Interior.Color = (&HD8D8D8)  ', "AU"
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
'    Sheet10.Range("DPM45.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet10.Range("DPM45.AddlnDeprlessthan180days").MergeArea.Locked = False
    Sheet10.Range("DPM45.DepUnderSec38").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM45.DepUnderSec38").MergeArea.Locked = False
    Sheet10.Range("DPM45.ProAgrdep").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.ProAgrdep").MergeArea.Locked = False
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").MergeArea.Locked = False
    Sheet10.Range("DPM45.CapGainUs50").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.CapGainUs50").MergeArea.Locked = False
    
     
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HD8D8D8)  ', "AU"
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").Interior.Color = (&HD8D8D8)  ', "AU"
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").MergeArea.Locked = False
    
'    Sheet10.Range("DPM45.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet10.Range("DPM45.AddlnDeprlessthan180days").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.AdditionsGrThan180Days").Interior.Color = (&HD8D8D8)  ', "AU"
    Sheet10.Range("DPM45.AdditionsGrThan180Days").MergeArea.Locked = True
    
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.Color = (&HD8D8D8) ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM15.WDVFirstDayBAC").value = ""
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM30.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM30.WDVFirstDayBAC").value = ""
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM40.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM40.WDVFirstDayBAC").value = ""
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM45.WDVFirstDayBAC").value = ""
End Sub


Sub resetDPMForYesold1()
   
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AddlnDeprlessthan180days").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.AdditionsGrThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsGrThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsGrThan180Days").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    Sheet10.Range("DPM45.AdditionsLessThan180Days").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM15.WDVFirstDayBAC").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM30.WDVFirstDayBAC").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.WDVFirstDayBAC").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM40.WDVFirstDayBAC").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.WDVFirstDayBAC").Interior.ColorIndex = 15  ', "AU"
    Sheet10.Range("DPM45.WDVFirstDayBAC").MergeArea.Locked = True
    
     Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.RealizationTotalPeriod").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.RealizationTotalPeriod").MergeArea.Locked = True
    Sheet10.Range("DPM45.RealizationTotalPeriod").value = ""
     
    Sheet10.Range("DPM45.DepUnderSec38").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.DepUnderSec38").MergeArea.Locked = True
    Sheet10.Range("DPM45.DepUnderSec38").value = ""
     
    Sheet10.Range("DPM45.ProAgrdep").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.ProAgrdep").MergeArea.Locked = True
    Sheet10.Range("DPM45.ProAgrdep").value = ""
     
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").MergeArea.Locked = True
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").value = ""
    
    Sheet10.Range("DPM45.CapGainUs50").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.CapGainUs50").MergeArea.Locked = True
    Sheet10.Range("DPM45.CapGainUs50").value = ""
    
    Sheet10.Range("I7:I9").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("I7:I9").Locked = True
    Sheet10.Range("I7:I9").value = ""

    Sheet10.Range("I11:I12").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("I11:I12").Locked = True
    Sheet10.Range("I11:I12").value = ""

    Sheet10.Range("I14:I15").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("I14:I15").Locked = True
    Sheet10.Range("I14:I15").value = ""
     Sheet10.Range("I19:I21").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("I19:I21").Locked = True
    Sheet10.Range("I19:I21").value = ""
     Sheet10.Range("I23").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("I23").Locked = True
    Sheet10.Range("I23").value = ""
    Sheet10.Range("I25:I27").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("I25:I27").Locked = True
    Sheet10.Range("I25:I27").value = ""
    Sheet10.Protect Password:=getmsgstate
End Sub

Sub resetDPMForNoold1()
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("I7:I9").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("I7:I9").Locked = False
    Sheet10.Range("I11:I12").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("I11:I12").Locked = False
    Sheet10.Range("I14:I15").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("I14:I15").Locked = False
     Sheet10.Range("I19:I21").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("I19:I21").Locked = False
     Sheet10.Range("I23").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("I23").Locked = False
    Sheet10.Range("I25:I27").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("I25:I27").Locked = False
    
     Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").Interior.ColorIndex = 15  ', "AU"
    Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").MergeArea.Locked = True
    
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").Interior.ColorIndex = 15  ', "AU"
    Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").MergeArea.Locked = True
    
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM15.AddlnDeprlessthan180days").MergeArea.Locked = False
    
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM30.AddlnDeprlessthan180days").MergeArea.Locked = False
    
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet10.Range("DPM40.AddlnDeprlessthan180days").MergeArea.Locked = False
    
'    Sheet10.Range("DPM45.AddlnDeprlessthan180days").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet10.Range("DPM45.AddlnDeprlessthan180days").MergeArea.Locked = False
    
    Sheet10.Range("DPM45.RealizationTotalPeriod").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.RealizationTotalPeriod").MergeArea.Locked = False
    
     
    Sheet10.Range("DPM45.DepUnderSec38").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.DepUnderSec38").MergeArea.Locked = False
    
     
    Sheet10.Range("DPM45.ProAgrdep").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.ProAgrdep").MergeArea.Locked = False
    
     
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").MergeArea.Locked = False
    
    
    Sheet10.Range("DPM45.CapGainUs50").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet10.Range("DPM45.CapGainUs50").MergeArea.Locked = False
    
    
    Sheet10.Range("DPM45.AdditionsGrThan180Days").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.AdditionsGrThan180Days").MergeArea.Locked = True
    
    Sheet10.Range("DPM45.AdditionsLessThan180Days").Interior.ColorIndex = 15 ', "AU"
    Sheet10.Range("DPM45.AdditionsLessThan180Days").MergeArea.Locked = True
    
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").Interior.ColorIndex = 15 ', "AU"
    Sheet10.Range("DPM45.RealizationPeriodDuringYear").MergeArea.Locked = True
    
    
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM15.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM15.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM15.WDVFirstDayBAC").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM30.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM30.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM30.WDVFirstDayBAC").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM40.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM40.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM40.WDVFirstDayBAC").value = 0
    
    Sheet10.Unprotect Password:=getmsgstate
    Sheet10.Range("DPM45.WDVFirstDayBAC").Interior.ColorIndex = 15   ', "AU"
    Sheet10.Range("DPM45.WDVFirstDayBAC").MergeArea.Locked = True
    Sheet10.Range("DPM45.WDVFirstDayBAC").value = 0
    
End Sub
Function ValidateLEI() As Boolean



'LEI_Number = Sheet1.Range("LEI_Number").value
'LEI_Date = Sheet1.Range("LEI_Date").value
'
'Refund_BTI = Sheet7.Range("Sheet9.RefundDue").value
'ValidateLEI = True
'
'If Refund_BTI > 50000000 Then
'
'        If LEI_Number = "" Then
'          MsgBx_General = "Please, enter LEI number in Part A General Sheet"
'          ValidateLEI = False
'          Exit Function
'        End If
'
'        If Trim(LEI_Date) = "" Or IsEmpty(LEI_Date) Or LEI_Date = "00/00/0000" Then
'         MsgBx_General = "Please, enter LEI Date in Part A General Sheet"
'         ValidateLEI = False
'         Exit Function
'        End If
'End If
'
'
'If Not FormatNCheckDate(LEI_Date) Then
'    ValidateLEI = False
'    MsgBx_General = "Please enter valid LEI date in dd/mm/yyyy format  in schedule Part A General"
'    Exit Function
'End If
'
'
'If Len(LEI_Number) > 15 Then
'    ValidateLEI = False
'    MsgBx_General = "LEI Number cannot exceed 15 Characters in Part A General"
'    Exit Function
'End If
'
'If Not CheckSpecialCharacter_New_1(LEI_Number) Then
'    ValidateLEI = False
'    MsgBx_General = "LEI Number cannot cannot contain special characters in Part A General"
'    Exit Function
'
'End If
'Chandru
LEI_Number = Sheet1.Range("LEI_Number").value
LEI_Date = Sheet1.Range("LEI_Date").value

Refund_BTI = Sheet7.Range("Sheet9.RefundDue").value
ValidateLEI = True

If Refund_BTI > 500000000 Then
    
        If LEI_Number = "" Then
          MsgBx_General = "* Please, enter LEI number in Part A General Sheet" & Chr(13)
          ValidateLEI = False
          Exit Function
        End If
        
        If Trim(LEI_Date) = "" Or IsEmpty(LEI_Date) Or LEI_Date = "00/00/0000" Then
         MsgBx_General = "* Please, enter LEI Date in Part A General Sheet" & Chr(13)
         ValidateLEI = False
         Exit Function
        End If
End If

        
If Not FormatNCheckDate(LEI_Date) Then
    ValidateLEI = False
    MsgBx_General = "*Please enter valid LEI date in dd/mm/yyyy format  in schedule Part A General" & Chr(13)
    Exit Function
End If

'Newly changed 15 to 20 as per DE sheet
If Len(LEI_Number) > 20 Then
    ValidateLEI = False
    MsgBx_General = "* LEI Number cannot exceed 20 Characters in Part A General" & Chr(13)
    Exit Function
End If

If Not CheckSpecialCharacter_New_1(LEI_Number) Then
    ValidateLEI = False
    MsgBx_General = "* LEI Number cannot  contain special characters in Part A General" & Chr(13)
    Exit Function

End If

'If Not IsNumeric(LEI_Number) Then
'    ValidateLEI = False
'    MsgBx_General = "* Please enter valid LEI Number in Part A General" & Chr(13)
'    Exit Function
'
'End If

End Function

Function validatesheetSebi() As Boolean

validatesheetSebi = True
Dim Sebi As Variant
Sebi = Trim(Range("Sheet1.sebi").value)


If Sebi <> "" Then
    'If Not ChkSebi(Sebi) Then MsgBx_General = "Please Enter a valid Sebi Registration Number" & Chr(13)
    'validatesheetSebi = False
If UCase(Left(Sebi, 2)) <> "IN" Then
   MsgBx_General = "Sebi Registration Number starts with IN"
   validatesheetSebi = False
End If

If Not ChkAlphabet(Mid(Sebi, 3, 2)) Then
   MsgBx_General = "Sebi Registration Number is invalid"
   validatesheetSebi = False
End If

If UCase(Mid(Sebi, 5, 2)) <> "FP" Then
   MsgBx_General = "Sebi Registration Number is invalid"
   validatesheetSebi = False
End If

If Not IsNumeric(Mid(Sebi, 7, 6)) Then
   MsgBx_General = "Sebi Registration Number is invalid"
   validatesheetSebi = False
End If

 If Not checkfieldspecialcharacter_Bank(Sebi) Then
   MsgBx_General = "Sebi Registration Number is invalid"
   validatesheetSebi = False
 End If
End If
'If errmsgTRP <> "" Then validatesheetTRP = False

End Function
Function ChkSebi(item As Variant) As Boolean
On Error Resume Next
Dim i As Long
ChkSebi = True

If Len(item) = 12 Then
'
'    If Not (CheckAtoZ(Mid(item, 1, 1))) Then
'        Range("Sheet1.sebi").value = ""
'        ChkSebi = False
'        Exit Function
'    End If
    
    If Not IsNumeric(Mid(item, 1, 12)) Then
        ChkSebi = False
        Range("Sheet1.sebi").value = ""
        Exit Function
    End If
    
    
Else
    ChkSebi = False
    Range("Sheet1.sebi").value = ""
    Exit Function
End If
End Function
Function ChkDPIIT(item As Variant) As Boolean
On Error Resume Next
Dim i As Long
ChkDPIIT = True

If Len(item) <= 9 Then


    If Mid(item, 1, 4) <> "DIPP" Then
       ChkDPIIT = False
        Range("ESOP.DPIIT").value = ""
        Exit Function


    End If
    
    If Not IsNumeric(Mid(item, 5, 9)) Then
        ChkDPIIT = False
        Range("ESOP.DPIIT").value = ""
        Exit Function
    End If
     
Else
    ChkDPIIT = False
    Range("ESOP.DPIIT").value = ""
    Exit Function
End If
End Function

Function CheckPAB_80DD(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 3 is for individuals .So,4th character of PAN should be "P" or "H"
'status = Sheet1.Range("sheet1.Status").value

PAN = UCase(PAN)
    CheckPAB_80DD = True
   
    If Len(PAN) > 0 Then
   
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
   
            CheckPAB_80DD = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
   
            CheckPAB_80DD = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
   
            CheckPAB_80DD = False
            Exit Function
        End If
        
      If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
     
            CheckPAB_80DD = False
            Exit Function
        End If
        
       'Commented by Shrutika(02/06/25)( SIT-93760)
'       If UCase(Mid(PAN, 4, 1)) <> UCase("P") Then
'          CheckPAB_80DD = False
'            Exit Function
'
'       End If
       
         
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
     
            CheckPAB_80DD = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
     
            CheckPAB_80DD = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
     
            CheckPAB_80DD = False
            Exit Function
        End If
    End If
End Function

'Added by Shrutika- 16/01/2026 AY-26-27

Function Dformat1(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("Date_Filing1").value) > 0 Then
 
Year = Mid(SheetALL.Range("Date_Filing1").value, 7, 4)
month = Mid(SheetALL.Range("Date_Filing1").value, 4, 2)
day = Mid(SheetALL.Range("Date_Filing1").value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat1 = formateddate
 
Else
Dformat1 = ""
End If
End Function

'Added by Shrutika- 20/01/2026 AY-26-27

Function CheckName(name As Variant, Length As Variant, Optional Mandatory As Boolean = False) As Boolean
On Error Resume Next

errmsgID = ""
CheckName = True

If Mandatory Then
    If IsEmpty(name) Or name = "" Then
        CheckName = False
        errmsgID = "is Mandatory"
        Exit Function
    End If
Else
    If Len(name) > Length Then
        CheckName = False
        errmsgID = "should be less than " & Length & " Characters" & Chr(13)
        Exit Function
    End If
End If
  
End Function

'by sadineni on 30/01/2026 for AY 2026-27
Function ChkAreaAlternate() As Boolean
On Error Resume Next
    ChkAreaAlternate = True
    FlatNew = Sheet1.Range("sheet1.ResidenceNo2")
    AreaNew = Sheet1.Range("sheet1.LocalityOrArea2")
    RoadOrStreetNew = Sheet1.Range("sheet1.RoadOrStreet2")
    
If FlatNew <> "" Or CityNew <> "" Or (StateNew <> "" And StateNew <> "(Select)") Or (sCountryNew <> "" And sCountryNew <> "(Select)") Or PinCodeNew <> "" Or ZipCodeNew <> "" Then
    If Trim(AreaNew) = "" Or IsEmpty(AreaNew) Then
          ChkAreaAlternate = False
'          fmsgboxStatus ("Area/ Locality in Sheet: Personal Information is mandatory in schedule Part A General under Personal Information")
'        ChkAreaAlternate = False
        Exit Function
    End If
End If

    If Len(Trim(AreaNew)) > 50 Then
        fmsgboxStatus ("Area/ Locality cannot be more than 50 characters in schedule Part A General under Personal Information")
        ChkAreaAlternate = False
        Exit Function
    End If

'    If Not CheckSpecialCharacter4(AreaNew) Then
'        ChkAreaAlternate = False
'        fmsgboxStatus ("Area / Locality  in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
'        Exit Function
'    End If

    If Not CheckSpecialCharacter4New(AreaNew) Then
        ChkAreaAlternate = False
        'SIT-112112 by sadineni on 27/02/2026 for AY 2026-27
        fmsgboxStatus ("Area / Locality  in Sheet : Part A-General cannot contain special charachters <> ")
        Exit Function
    End If

End Function

'by sadineni on 30/01/2026 for AY 2026-27
Function ChkCityAlternate() As Boolean
On Error Resume Next
    ChkCityAlternate = True
    CityNew = Sheet1.Range("sheet1.CityOrTownOrDistrict2")
    
    
If FlatNew <> "" Or AreaNew <> "" Or (StateNew <> "" And StateNew <> "(Select)") Or (sCountryNew <> "" And sCountryNew <> "(Select)") Or PinCodeNew <> "" Or ZipCodeNew <> "" Then
    If Trim(CityNew) = "" Or IsEmpty(CityNew) Then
        ChkCityAlternate = False
        Exit Function
    End If
End If

    If Len(Trim(CityNew)) > 50 Then
        fmsgboxStatus ("Town/City/District cannot be more than 50 characters in schedule Part A General")
        ChkCityAlternate = False
        Exit Function
    End If

'    If Not CheckSpecialCharacter4(CityNew) Then
'        ChkCityAlternate = False
'        fmsgboxStatus ("Town/City/District in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
'        Exit Function
'    End If

    If Not CheckSpecialCharacter4New(CityNew) Then
        ChkCityAlternate = False
        'SIT-112112 by sadineni on 27/02/2026 for AY 2026-27
        'fmsgboxStatus ("Town/City/District in Sheet : Part A-General cannot contain special charachters <> ' '&' "" ' ")
         fmsgboxStatus ("Town/City/District in Sheet : Part A-General cannot contain special charachters <> ")
        Exit Function
    End If

End Function

'by sadineni on 30/01/2026 for AY 2026-27
Function ChkCountryAlternate() As Boolean
On Error Resume Next
    ChkCountryAlternate = True
    sCountryNew = Sheet1.Range("sheet1.CountryCode2")

'by sadineni on 08/04/2026
If FlatNew <> "" Or AreaNew <> "" Or CityNew <> "" Or (sCountryNew <> "" And sCountryNew <> "(Select)") Or PinCodeNew <> "" Or ZipCodeNew <> "" Then

    If isdropdownblank(sCountryNew) Then
        sCountryNew = ""
    End If
    
    If Trim(sCountryNew) = "" Or Trim(sCountryNew) = "(Select)" Or Trim(sCountryNew) = "(" Then
        ChkCountryAlternate = False
        Exit Function
    End If
End If

End Function

'by sadineni on 30/01/2026 for AY 2026-27
Function ChkPincodeAlternate() As Boolean
On Error Resume Next
    ChkPincodeAlternate = True
    PinCodeNew = Sheet1.Range("sheet1.PinCode2")
    
'by sadineni on 08/04/2026
If (sCountryNew <> "" And sCountryNew <> "(Select)") Then
    If sCountryNew <> "" And UCase(Mid(sCountryNew, 1, InStr(1, sCountryNew, "-") - 1)) = "91" Then
        If Trim(PinCodeNew) = "" Or IsEmpty(PinCodeNew) Then
            MsgBx_General = MsgBx_General + "* ""Pin code is mandatory in tab: Personal Information in schedule Part A General""" & Chr(13)
            ChkPincodeAlternate = False
            Exit Function
        End If
        
        If Len(Trim(PinCodeNew)) > 6 Then
            MsgBx_General = MsgBx_General + "*""Pin Code must be exactly 6 digits and should not start with zero in schedule Part A General""" & Chr(13)
            ChkPincodeAlternate = False
            Exit Function
        End If
        
        If Mid(PinCodeNew, 1, 1) = "0" Then
            MsgBx_General = MsgBx_General + "* Pincode in Sheet : Part A-General should not start with 0" & Chr(13)
            ChkPincodeAlternate = False
            Exit Function
        End If
    Else
        PinCodeNew = ""
    End If
End If
End Function

'by sadineni on 30/01/2026 for AY 2026-27
Function ChkZipcodeAlternate() As Boolean
On Error Resume Next
    ChkZipcodeAlternate = True
    ZipCodeNew = Sheet1.Range("sheet1.ZipCode2")
   
If sCountryNew <> "" And sCountryNew <> "(Select)" Then     'Konda updated on 05-05-2026---SIT-118819
If sCountryNew <> "" And UCase(Mid(sCountryNew, 1, InStr(1, sCountryNew, "-") - 1)) <> "91" Then
        If Trim(ZipCodeNew) = "" Or IsEmpty(ZipCodeNew) Then
            MsgBx_General = MsgBx_General + "*""Zip Code is compulsory, if there is no Zip Code then select ""No ZIP Code in schedule Part A General.""" & Chr(13)
            ChkZipcodeAlternate = False
            Exit Function
        End If
        
    
        If Len(Trim(ZipCodeNew)) > 8 Then
            MsgBx_General = MsgBx_General + "* ZipCode in Sheet : Part A-General cannot exceed 8 characters.Minimum 1 and up to 8 Characters." & Chr(13)
            ChkZipcodeAlternate = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(ZipCodeNew) Then
            MsgBx_General = MsgBx_General + "* ZipCode in Sheet : Part A-General characters < > & ' " & Chr(34) & " are not allowed"" & chr(13)"
            ChkZipcodeAlternate = False
            Exit Function
        End If
    Else
        ZipCodeNew = ""
End If
End If  'Konda updated on 05-05-2026---SIT-118819

End Function

'by sadineni on 13/02/2026
Function CheckEmailAddress_rep(emailAddress As Variant, ByRef vType As Variant) As Boolean
On Error Resume Next
    CheckEmailAddress_rep = True
    If Len(emailAddress) > 0 Or Not IsEmpty(emailAddress) Then
    
        If Not CheckSpecialCharacter(Mid(emailAddress, 1, 1)) Then
            vType = "1"
            CheckEmailAddress_rep = False
            Exit Function
        End If
        
        If Len(emailAddress) > 125 Then
            vType = "2"
            CheckEmailAddress_rep = False
            Exit Function
        End If
                
        If Not IsValidEmail(emailAddress) Then
            vType = "3"
            CheckEmailAddress_rep = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacter_FindConsecutive_new(emailAddress) Then
            vType = "3"
            CheckEmailAddress_rep = False
            Exit Function
        End If
    Else
            vType = "4"
            CheckEmailAddress_rep = False
            Exit Function
    End If
End Function

'by sadineni on 13/02/2026
Public Function CheckSpecialCharacter_FindConsecutive_new(ByVal emailid As String) As Boolean
    Dim i As Long
    Dim ch As String * 1
    Dim consecutive As Long
    ' Quick reject for "_@"
    If InStr(1, emailid, "_@", vbTextCompare) > 0 Then
        CheckSpecialCharacter_FindConsecutive_new = False
        Exit Function
    End If
    ' Scan for two or more consecutive of - / _ .
    For i = 1 To Len(emailid)
        ch = Mid$(emailid, i, 1)
        If ch = "-" Or ch = "/" Or ch = "_" Or ch = "." Or ch = "@" Then
            consecutive = consecutive + 1
            If consecutive >= 2 Then
                CheckSpecialCharacter_FindConsecutive_new = False
                Exit Function
            End If
        Else
            consecutive = 0
        End If
    Next i
    CheckSpecialCharacter_FindConsecutive_new = True
End Function

'by sadineni for AY 2026-27 on 27/02/2026
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


