Attribute VB_Name = "GenerateJSON"
'01/12/2021
Option Explicit
Public xmlLinesCount As Variant
Public OptingNewTaxRegime As Variant
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

Function getHashIteration() As String
'Changed by Ayush on 02/01/2025
'    getHashIteration = "1977"
    getHashIteration = "1988" 'Ayush_29/01/2026
End Function

Function getHashKey() As String
'    getHashKey = "zcKkZH7jksLZWLYT"
'     getHashKey = "RoduFpNqMzQlWO9Q"
'Changed by Ayush on 02/01/2025
'getHashKey = "LO3QtH59fGuVaETa"
getHashKey = "HZX4oKH11zARYIb2" 'Ayush_29/01/2026
End Function

Sub Generate()
On Error Resume Next
    Dim JSONFileName As String
    Dim AsseePan
    Dim Status_3 As String
    Status_3 = Sheet1.Range("sheet1.Status").value
    Status_3 = Mid(Status_3, 1, 1)

    InitProgBar

  ProgressFrameCaption = "Generating JSON"
  mainProcCaption = "Validating"
  noOfProcessMain = 15
  UserForm1.Show vbModeless

   UpdateProgressBar
    ValidateSheetIncomeDetails
    

     UpdateProgressBar
    ValidateSheet44AE

    UpdateProgressBar
    ValidateSheetNOBBP

    UpdateProgressBar
    ValidateTDS_TCS_IT

    UpdateProgressBar
    Validate_IT

    UpdateProgressBar
    Validate_TCS
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
        UpdateProgressBar
        ValidateSheet80G
        Validate80GGC
    End If
    
    If (Sheet8.Range("bacValue").value) = 2 Then
        UpdateProgressBar
        ValidateSheet80G
        UpdateProgressBar
        Validate80D_All
        UpdateProgressBar
        Validate80GGC
        'Newly added
        UpdateProgressBar
        Validate80U
        UpdateProgressBar
        Validate80DD
        
    'Malli------------24/04/2025-
        UpdateProgressBar
'        Validate80C_80CCC_All
'        UpdateProgressBar
'        Validate80C_80CCC2_All
         Validate80C__All  '08/05/2025
    If Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
        UpdateProgressBar
        Validate80E_All
        UpdateProgressBar
        Validate80EE_All
        UpdateProgressBar
        Validate80EEA_All
        UpdateProgressBar
        Validate80EEB_All
        UpdateProgressBar
    End If
        'Validate24b_All
        
        
    End If
    
    'Malli--------SIT_SIT-93261
    If Not (Sheet8.Range("BacValue").value = 1 And Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1) = "S") Then
    UpdateProgressBar
    Validate24b_All
    End If
    '--------------------------
    
    'Malli------08/05/2025
    If Sheet8.Range("BacValue").value = 2 And Left(Sheet1.Range("sheet1.Status").value, 1) = "I" And Mid(Range("sheet1.NatureofEmployment").value, 1, 1) <> "N" Then
     ValidateEA10_13A
    End If
    '-------------------

    UpdateProgressBar
    ValidateDI

    UpdateProgressBar
    ValidateSheetTaxPaidVerificaton

    UpdateProgressBar
    ValidateBA
    
'Change 02, Satya, 15.10.2022
If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then

    UpdateProgressBar
    Validate_Gen_1398A
    UpdateProgressBar
    ValidateATI
End If

'End Change

    Dim income_BP As Variant
    income_BP = Sheet3.Range("BP_E2_PI").value + IIf(Sheet3.Range("BP_E4_PI44ADA").value = "", 0, Sheet3.Range("BP_E4_PI44ADA").value) + Sheet3.Range("BP_E3").value

    AsseePan = Sheet1.Range("sheet1.PAN").value

    JSONFileName = ThisWorkbook.Path & "\" & "ITR4_" & AsseePan & ".json"
    
'    Open JsonFileName For Output As #1
'    Print #1, ToJsonFormat()
'    Close #1
   ' MsgBox ToJsonFormat()
   
   Dim str
    str = ToJsonFormat()
    
    ConvertJSONToString2 (str)
   
   ProgressBarHide

If Mid(Sheet1.Range("sheet1.ReturnType1").value, 1, 3) = "Rec" Then
fmsgboxsmall ("* Please upload the return(JSON) on below link" & Chr(13) & Chr(13) & "https://www.incometax.gov.in-->Login-->e-fiie---> Rectification" & " to file return")  ', vbOKOnly, "ITR-4 AY 2019-2020"
ElseIf Mid(Sheet1.Range("sheet1.ReturnType1").value, 1, 1) = "I" Then
fmsgboxsmall ("* Please upload the return(JSON) on below link" & Chr(13) & Chr(13) & "https://www.incometax.gov.in-->Login-->e-fiie---> Response to notice u/s139(9)" & " to file return") ' , vbOKOnly, "ITR-4 AY 2019-2020"
Else
fmsgboxsmall ("* JSON is saved in the path : " & JSONFileName & Chr(10) & Chr(10) & "Please upload the return(JSON) in below link" & Chr(13) & Chr(13) & "To upload the saved JSON, Go to https://www.incometax.gov.in -> Login -> e-File -> Income Tax Return.") ', vbOKOnly, "ITR-4 AY 2019-2020"
End If

End Sub

Public Function ConvertJSONToString2(ByVal str As String)
Dim ts(1) As Object
    Dim count As Variant
    Dim FSO, jsonObject As Object
    Dim strXml, hashcode As String
    Dim SecretKey As String
    Dim iteration As Long
    Dim FileContents As String
    Dim SEARCH_FOR As String
    Dim REPLACE_WITH As String
    Dim arrSearchThis() As Variant
    Dim objTS As Variant
    Dim i, DigestFoundat As Variant
    Dim jsonText As String
    iteration = getHashIteration
    SecretKey = getHashKey
    
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    Open JSONFileName For Input As #1
'        jsonText = Input$(LOF(1), 1)
'    Close #1
    i = 0
    InitProgBar
    ProgressFrameCaption = "JSON Generation"
    mainProcCaption = "Encrypting JSON"
    noOfProcessMain = 3
    UpdateProgressBar
    UpdateProgressBar
    Set jsonObject = ImportJSON.ParseJson(str)
    str = ToJson(jsonObject)
    'hashcode = Base64_HMACSHA256_JSON(iteration, str, SecretKey)
    'hashcode = "V7CvRvFLjQkccRsQckM9xV/LkAEb4vykJuMOv9vsSeU="
    
    
    'Newly added for Digest Code on 13th Dec 24 by Bindu
    hashcode = HMACSHA256A(str, StrConv(SecretKey, vbFromUnicode), iteration)
    
    
    jsonObject("ITR")("ITR4")("CreationInfo")("Digest") = hashcode
    Dim JSONFileName
    Dim AssesseePan
    AssesseePan = Sheet1.Range("sheet1.PAN").value
    AssesseePan = "ITR4_" & AssesseePan
    JSONFileName = ThisWorkbook.Path & "\" & AssesseePan & ".json"
    Open JSONFileName For Output As #2
    subProcCaption = "Publishing JSON"
    Print #2, ToJson(jsonObject);
            UpdateProgressBar
    Close #2
End Function

Public Function Base64_HMACSHA256_JSON(ByVal iteration As Long, sTextToHash As String, ByVal sSharedSecretKey As String)

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    
    Dim tempTextToHash() As Byte
    Dim tempString, CountTextToHash As Variant
    Dim i, j As Variant
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
    ReDim TextToHash(0)
    subProcCaption = "Encrypting JSON"
    TextToHash = StrConv(sTextToHash, vbFromUnicode)
        UpdateProgressBar
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim byteS() As Byte
    byteS = enc.ComputeHash_2((TextToHash))
    
    For i = 1 To iteration
        byteS = enc.ComputeHash_2((byteS))
    Next
     
    Base64_HMACSHA256_JSON = EncodeBase64json(byteS)
    Set asc = Nothing
    Set enc = Nothing

End Function


Function EncodeBase64json(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64json = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing

End Function

Function ToJsonFormat() As String
Dim str As String
Dim init, itr, itr4

Set init = CreateObject("Scripting.Dictionary")
Set itr = CreateObject("Scripting.Dictionary")
Set itr4 = CreateObject("Scripting.Dictionary")

itr4.add "CreationInfo", CreationInfo()
itr4.add "Form_ITR4", Form_ITR4()
itr4.add "PersonalInfo", PersonalInfo()
itr4.add "FilingStatus", FilingStatus()
If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
itr4.add "PartA_139_8A", PartA_139_8A()
itr4.add "PartB-ATI", PartB_ATI()
End If
itr4.add "IncomeDeductions", ITR4_IncomeDeductions()
itr4.add "TaxComputation", TaxComputation()
itr4.add "TaxPaid", TaxPaid()
itr4.add "Refund", Refund()
itr4.add "Verification", Verification()
'If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Or (Sheet8.Range("bacValue").value) = 2 Then
'    itr4.add "Schedule80G", Schedule80G()
'End If

'If (Sheet8.Range("bacValue").value) = 2 Then
'
'   ' itr4.add "Schedule80G", Schedule80G()
'    itr4.add "Schedule80GGC", Schedule80GGC() '2024-25 Added by Malli
'
'
'  If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
'
'    itr4.add "Schedule80DD", Schedule80DD() '2024-25 Change
'    itr4.add "Schedule80U", Schedule80U() '2024-25 change
'    itr4.add "Schedule80D", Schedule80D()
'
'  End If
'
'End If

If (Sheet8.Range("bacValue").value) = 2 Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
    'itr4.add "Schedule80D", Schedule80D()
    itr4.add "Schedule80G", Schedule80G()
    itr4.add "Schedule80GGC", Schedule80GGC() '2024-25 Added by Malli
    
'Malli_legacy40_02/06/2025

'    itr4.add "Schedule80DD", Schedule80DD() '2024-25 Change
'    itr4.add "Schedule80U", Schedule80U() '2024-25 change

    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
    itr4.add "Schedule80DD", Schedule80DD() '2024-25 Change
    End If
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
    itr4.add "Schedule80U", Schedule80U() '2024-25 change
    End If
    
    'Malli_AY_2025_26  09/04/2025------------------
    'Malli_SIT-118203  23/04/2026
   If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
        itr4.add "Schedule80E", Schedule80E()
        itr4.add "Schedule80EE", Schedule80EE()
        itr4.add "Schedule80EEA", Schedule80EEA()
        itr4.add "Schedule80EEB", Schedule80EEB()
   End If
   
   If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
        itr4.add "Schedule80C", Schedule80C()
        itr4.add "Schedule80D", Schedule80D()
    End If
    
    'AY_25_26_V0.7
    'itr4.add "Schedule80CCC", Schedule80CCC()
    '-----------------
'Konda updated on 29-04-2025---SIT-92218
    'itr4.add "ScheduleUs24B", Schedule24B()
  '-----------------------------------------------
'    itr4.add "Schedule80D", Schedule80D()
    
    
End If

'malli comented-------------
''Konda updated on 29-04-2025---SIT-92218
'If Sheet8.Range("BacValue").value = 1 And Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "S" And Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "(" Then
'      itr4.add "ScheduleUs24B", Schedule24B()
'ElseIf Sheet8.Range("BacValue").value = 2 And Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "(" Then
'     itr4.add "ScheduleUs24B", Schedule24B()
'End If
'Konda commented on 29-01-2026
'Malli---------SIT_SIT-93261--------
'If Not (Sheet8.Range("BacValue").value = 1 And Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1) = "S") Then
'   itr4.add "ScheduleUs24B", Schedule24B()
'End If

'----------------------

'Konda AY_2025_26 08-05-2025------------
'If Sheet8.Range("BacValue").value = 2 And Mid(Range("sheet1.NatureofEmployment").value, 1, 1) <> "N" Then
If Sheet8.Range("BacValue").value = 2 And Left(Sheet1.Range("sheet1.Status").value, 1) = "I" And Mid(Range("sheet1.NatureofEmployment").value, 1, 1) <> "N" Then
      itr4.add "ScheduleEA10_13A", ScheduleEA10_13A()

End If
'-------------------
itr4.add "ScheduleBP", ScheduleBP()
itr4.add "TaxExmpIntIncDtls", TaxExmpIntIncDtls()
itr4.add "LTCG112A", LTCG112A_New() 'added by Chetan C M AY 2025-26 on 31/01/2025
itr4.add "TaxReturnPreparer", TaxReturnPreparer()
itr4.add "ScheduleIT", ScheduleIT()
itr4.add "ScheduleTCS", ScheduleTCS()
itr4.add "TDSonSalaries", TDSonSalaries()
itr4.add "TDSonOthThanSals", TDSonOthThanSals()
itr4.add "ScheduleTDS3Dtls", ScheduleTDS3Dtls()

itr.add "ITR4", itr4
init.add "ITR", itr

str = ToJson(init)
ToJsonFormat = str
End Function
Function CreationInfo() As Object
    Dim jsonDictionary
    Dim SWVersionNo, SWCreatedBy, JSONCreatedBy, IntermediaryCity, Digest, JSONCreationDate As Variant

    Set jsonDictionary = CreateObject("Scripting.Dictionary")
    'Malli------12/07/2024----------
    'SWVersionNo = "R2"
'    SWVersionNo = "R3"
   ' SWVersionNo = "R4" 'Changed by sai on 30/07/2024
  '  SWVersionNo = "R5" 'Changed by Ayush on 14/11/20224
   ' SWVersionNo = "R6"
    
    'changed by Malli as per AY_2025_26  Date:13/12/2024
  '  SWVersionNo = "R1"
'     SWVersionNo = "R2"  'Changed by Ayush on 27/06/2025
'     SWVersionNo = "R3"   'Changed by Ayush on 29/07/2025
    ' SWVersionNo = "R4"   'Changed by Ayush on 25/08/2025
    ' SWVersionNo = "R5"   'Changed by Shrutika on 15/09/2025
      
   '   SWVersionNo = "R6"     'Changed by Jyoti on 18/12/2025
'   SWVersionNo = "R1" 'Changed by Ayush_29/01/2026
    SWVersionNo = "R2" 'Changed by Ayush_29/06/2026
    '------------------
'    SWCreatedBy = "SW90002425"
'    JSONCreatedBy = "SW90002425"
'Changed by Ayush on 29/01/2026
     SWCreatedBy = "SW90002627"
    
'Konda updated on 12-02-2026--SIT-110328
'    JSONCreatedBy = "SW90002526"
    JSONCreatedBy = "SW90002627"
    '--------------
    JSONCreationDate = Range("DateOfProcessing").value
    IntermediaryCity = "Delhi"
    Digest = "-"
    
    jsonDictionary("SWVersionNo") = SWVersionNo
    jsonDictionary("SWCreatedBy") = SWCreatedBy
    jsonDictionary("JSONCreatedBy") = JSONCreatedBy
    jsonDictionary("JSONCreationDate") = JSONCreationDate
    jsonDictionary("IntermediaryCity") = IntermediaryCity
    jsonDictionary("Digest") = Digest
    Set CreationInfo = jsonDictionary
    
End Function
Function Form_ITR4() As Object
Dim jsonDictionary
Dim FormName, Description, AssessmentYear, SchemaVer, FormVer As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")

FormName = "ITR-4"
Description = "For Individuals having Income From Presemptive Business"
'AssessmentYear = "2022"
'AssessmentYear = "2024" 'Year Changed from 2023 to 2024
'AssessmentYear = "2025" 'Year Changed from 2024 to 2025
'New Schema updated by Konda as AY-2026-27 on 29-12-2025
AssessmentYear = "2026" 'Year Changed from 2025 to 2026
SchemaVer = "Ver1.0"
FormVer = "Ver1.0"

    jsonDictionary("FormName") = FormName
    jsonDictionary("Description") = Description
    jsonDictionary("AssessmentYear") = AssessmentYear
    jsonDictionary("SchemaVer") = SchemaVer
    jsonDictionary("FormVer") = FormVer
    Set Form_ITR4 = jsonDictionary
End Function
'XML Generation Function For Personal Information
Function PersonalInfo() As Object
subProcCaption = "Personal Info"
noOfProcessSub = 14
Dim jsonDictionary, AssesseeName, Phone, Address, sCountry

'Konda updated on 22-01-2026 schema-V0.2.4
'SecondaryAdd1
'Dim WantUpdateAdd, AlternateAddress, sCountry_Sec
Dim SecondaryAdd, AlternateAddress, sCountry_Sec

Dim aadharEnroll As String
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set AssesseeName = CreateObject("Scripting.Dictionary")
Set Phone = CreateObject("Scripting.Dictionary")
Set Address = CreateObject("Scripting.Dictionary")
'Konda updated on 22-01-2026 schema-V0.2.4
'V0.8
'Set WantUpdateAdd = CreateObject("Scripting.Dictionary")
Set SecondaryAdd = CreateObject("Scripting.Dictionary")
Set AlternateAddress = CreateObject("Scripting.Dictionary")

'Ayush_enrol
'aadharEnroll = CStr(Sheet1.Range("sheet1.AadhaarEnrol").value)
If Mid(NatureEmp, 1, 3) = "Sta" Then
NatureEmp = "SGOV"
ElseIf Mid(NatureEmp, 1, 3) = "Cen" Then
NatureEmp = "CGOV"
''Konda updated on 27-02-2026--v0.4
''Konda updated on 12-02-2026--v0.3
'ElseIf NatureEmp = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958" Then
'NatureEmp = "SCJ"
''=============================
ElseIf Mid(NatureEmp, 1, 3) = "Pub" Then
NatureEmp = "PSU"
'ElseIf Mid(NatureEmp, 1, 3) = "Pen" Then
'NatureEmp = "PE"
ElseIf NatureEmp = "Pensioners-CG" Then
NatureEmp = "PE"
ElseIf NatureEmp = "Pensioners-SG" Then
NatureEmp = "PESG"
ElseIf NatureEmp = "Pensioners-PSU" Then
NatureEmp = "PEPS"
ElseIf NatureEmp = "Pensioners-Other" Then
NatureEmp = "PEO"
ElseIf Mid(NatureEmp, 1, 3) = "Oth" Then
NatureEmp = "OTH"
ElseIf Mid(NatureEmp, 1, 3) = "Not" Then
NatureEmp = "NA"
End If

sCountry = Sheet1.Range("sheet1.CountryCode").value
sCountry_Sec = Sheet1.Range("sheet1.CountryCode2").value
    If firstName <> "" Then
        AssesseeName("FirstName") = UCase(firstName)
    End If
    
    UpdateProgressBar
    If middleName <> "" Then
        AssesseeName("MiddleName") = UCase(middleName)
    End If
    
    UpdateProgressBar
    If LastName <> "" Then
        AssesseeName("SurNameOrOrgName") = UCase(LastName)
    End If
    jsonDictionary.add "AssesseeName", AssesseeName

    UpdateProgressBar
    If PAN <> "" Then
       jsonDictionary("PAN") = UCase(PAN)
    End If

    UpdateProgressBar
    Address("ResidenceNo") = Trim(Sheet1.Range("sheet1.ResidenceNo").value)
 
    UpdateProgressBar
    If residenceName <> "" Then
        Address("ResidenceName") = UCase(residenceName)
    End If

    UpdateProgressBar
    If Sheet1.Range("sheet1.RoadOrStreet").value <> "" Then
        Address("RoadOrStreet") = roadOrStreet
    End If

    UpdateProgressBar
    If Area <> "" Then
        Address("LocalityOrArea") = Sheet1.Range("sheet1.LocalityOrArea").value
    End If

    UpdateProgressBar
    If City <> "" Then
        Address("CityOrTownOrDistrict") = Sheet1.Range("sheet1.CityOrTownOrDistrict").value
    End If

    UpdateProgressBar
    If State <> "" Then
        Address("StateCode") = Mid((State), 1, 2)
    End If
 
    UpdateProgressBar
    If sCountry <> "" Then
        Address("CountryCode") = Mid(sCountry, 1, InStr(sCountry, "-") - 1)
    End If
       
    If (Sheet1.Range("sheet1.PinCode").value) <> "" Then
        Address("PinCode") = UVCase(CDbl(Sheet1.Range("sheet1.PinCode").value))
    End If

    If zipCode <> "" Then
        Address("ZipCode") = UCase(zipCode)
    End If

'Change-05.03.2023.101 -> VF Issue
'    If (Len(stDcode) > 0 Or Len(phoneNo) > 0) Then
    If (Len(stDcode) > 0 And Len(phoneNo) > 0) Then
        Phone("STDcode") = UVCase(CDbl(stDcode))
        Phone("PhoneNo") = phoneNo
         Address.add "Phone", Phone
    End If
   
    
    If (Len(mobileNo) > 0) Then
        Address("CountryCodeMobile") = Sheet1.Range("sheet1.mobileCountryCode").value
        Address("MobileNo") = UVCase(CDbl(Sheet1.Range("sheet1.Mobileno").value))
    End If

    If Email <> "" Then
        Address("EmailAddress") = Sheet1.Range("sheet1.EmailAddress").value
    End If
    
    If (Len(CountrycodeMobileNo2) > 0) Then
        Address("CountryCodeMobileNoSec") = Sheet1.Range("sheet1.mobileCountryCode2").value
    End If
    
    If (Len(PhoneNo2) > 0) Then
        Address("MobileNoSec") = UVCase(CDbl(Sheet1.Range("sheet1.MobileNoSec").value))
    End If
    
    If Email2 <> "" Then
        Address("EmailAddressSec") = Sheet1.Range("sheet1.EmailAddress2").value
    End If
    jsonDictionary.add "Address", Address
    
'Konda updated on 17-03-2026--V0.6
''Konda updated on 22-01-2026 schema-V0.2.5
'If Sheet1.Range("Secondary_Address") <> "" Then
'        If Sheet1.Range("Secondary_Address").value = "Yes" Then
'            jsonDictionary("WantUpdateAdd") = "Y"
'       Else
'            jsonDictionary("WantUpdateAdd") = "N"
'       End If
'End If
'
'If Sheet1.Range("Secondary_Address").value = "Yes" Then

'========================================
'+= V0.8
If Sheet1.Range("Secondary_Address").value <> "" Then
        If Sheet1.Range("Secondary_Address").value = "Yes" Then
            jsonDictionary("SecondaryAdd") = "Y"
       Else
            jsonDictionary("SecondaryAdd") = "N"
       End If
  End If

If Sheet1.Range("Secondary_Address").value = "Yes" Then
'+=
    If Sheet1.Range("sheet1.ResidenceNo2").value <> "" Then

    AlternateAddress("ResidenceNo") = Trim(Sheet1.Range("sheet1.ResidenceNo2").value)
 
    UpdateProgressBar
    If residenceName_Sec <> "" Then
        AlternateAddress("ResidenceName") = UCase(residenceName_Sec)
    End If

    UpdateProgressBar
    If Sheet1.Range("sheet1.RoadOrStreet2").value <> "" Then
        AlternateAddress("RoadOrStreet") = roadOrStreet_Sec
    End If

    UpdateProgressBar
    If Area_Sec <> "" Then
        AlternateAddress("LocalityOrArea") = Sheet1.Range("sheet1.LocalityOrArea2").value
    End If

    UpdateProgressBar
    If City_Sec <> "" Then
        AlternateAddress("CityOrTownOrDistrict") = Sheet1.Range("sheet1.CityOrTownOrDistrict2").value
    End If

    UpdateProgressBar
    If State_Sec <> "" Then
        AlternateAddress("StateCode") = Mid((State_Sec), 1, 2)
    End If
 
    UpdateProgressBar
    'Ayush_23/03/2026
    If sCountry_Sec <> "" And sCountry_Sec <> "(Select)" Then
        AlternateAddress("CountryCode") = Mid(sCountry_Sec, 1, InStr(sCountry_Sec, "-") - 1)
    End If
       
    If (Sheet1.Range("sheet1.PinCode2").value) <> "" Then
        AlternateAddress("PinCode") = UVCase(CDbl(Sheet1.Range("sheet1.PinCode2").value))
    End If
'Konda updated on 12-02-2026--SIT-110096
'    If zipCode_Sec <> "" Then
'        AlternateAddress("ZipCode") = UCase(zipCode_Sec)
'    End If
    
    If (Sheet1.Range("sheet1.ZipCode2").value) <> "" Then
        AlternateAddress("ZipCode") = (Sheet1.Range("sheet1.ZipCode2").value)
    End If
'==========================================
    jsonDictionary.add "AlternateAddress", AlternateAddress
   End If      'Konda updated on 17-03-2026--V0.6
    'End If  'V0.8
    '+==
 ElseIf Sheet1.Range("Secondary_Address").value = "No" Then
 
    If Sheet1.Range("sheet1.ResidenceNo2").value <> "" Then

    AlternateAddress("ResidenceNo") = Trim(Sheet1.Range("sheet1.ResidenceNo2").value)
 
    UpdateProgressBar
    If residenceName_Sec <> "" Then
        AlternateAddress("ResidenceName") = UCase(residenceName_Sec)
    End If

    UpdateProgressBar
    If Sheet1.Range("sheet1.RoadOrStreet2").value <> "" Then
        AlternateAddress("RoadOrStreet") = roadOrStreet_Sec
    End If

    UpdateProgressBar
    If Area_Sec <> "" Then
        AlternateAddress("LocalityOrArea") = Sheet1.Range("sheet1.LocalityOrArea2").value
    End If

    UpdateProgressBar
    If City_Sec <> "" Then
        AlternateAddress("CityOrTownOrDistrict") = Sheet1.Range("sheet1.CityOrTownOrDistrict2").value
    End If

    UpdateProgressBar
    If State_Sec <> "" Then
        AlternateAddress("StateCode") = Mid((State_Sec), 1, 2)
    End If
 
    UpdateProgressBar
    'Ayush_23/03/2026
    If sCountry_Sec <> "" And sCountry_Sec <> "(Select)" Then
        AlternateAddress("CountryCode") = Mid(sCountry_Sec, 1, InStr(sCountry_Sec, "-") - 1)
    End If
       
    If (Sheet1.Range("sheet1.PinCode2").value) <> "" Then
        AlternateAddress("PinCode") = UVCase(CDbl(Sheet1.Range("sheet1.PinCode2").value))
    End If
'Konda updated on 12-02-2026--SIT-110096
'    If zipCode_Sec <> "" Then
'        AlternateAddress("ZipCode") = UCase(zipCode_Sec)
'    End If
    
    If (Sheet1.Range("sheet1.ZipCode2").value) <> "" Then
        AlternateAddress("ZipCode") = (Sheet1.Range("sheet1.ZipCode2").value)
    End If
'==========================================
    jsonDictionary.add "AlternateAddress", AlternateAddress
   End If      'Konda updated on 17-03-2026--V0.6
    End If  'V0.8
'==================
    '+==
'=============================================
    If AssesseeDob <> "" Then
        jsonDictionary("DOB") = Dformat(AssesseeDob, "")
    End If
       
    If NatureEmp <> "" Then
        jsonDictionary("EmployerCategory") = NatureEmp
    End If
    
    If status = "" Or IsEmpty(status) Or Mid(status, 1, 1) = "(" Then
        jsonDictionary("Status") = "I"
    Else
        jsonDictionary("Status") = Trim(UCase(Mid(Trim(status), 1, InStr(status, "-") - 1)))
    End If

    If AadhaarNumber <> "" Then
        jsonDictionary("AadhaarCardNo") = AadhaarNumber
    End If
'
'Ayush_enrol
'    If Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False Then
'         If Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") <= Dformat(Trim("30/09/2024"), "yyyy-mm-dd") Then
'            If aadharEnroll <> "" Then
'                jsonDictionary("AadhaarEnrolmentId") = aadharEnroll
'            End If
'        End If
'    End If
    

Set PersonalInfo = jsonDictionary
End Function
Function FilingStatus_old() As Object

subProcCaption = "Filling Status"
noOfProcessSub = 7

'Dim jsonDictionary, AssesseeRep, NewTaxRegime, date10IE, acknowledgementNumber, clauseiv7provisio139iType, clauseiv7provisio139iDtls AY 2023-24 OLD
Dim jsonDictionary, AssesseeRep, clauseiv7provisio139iType, clauseiv7provisio139iDtls
Dim NewForm10IEADate_1, NewForm10IEAAckNo_1 'AY 2024-25
Dim ItrFilingDueDate_1 As Variant 'AY 2024-25

Set clauseiv7provisio139iDtls = New Collection
Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set AssesseeRep = CreateObject("Scripting.Dictionary")

' For AY 2024-25 we had declared this
Dim TaxRegime_1 As Variant
TaxRegime_1 = Sheet1.Range("Sheet1.OptOutNewTaxRegime").value


'Schema Date > "7/31/2024"
'Utility date > 31/07/2024



'PAG_C9 AY 2024-25 Schema implementation Starts------------------------

'ItrFilingDueDate_1 = Mid((Sheet1.Range("Sheet1.Duedate").value), 4, 2) & "/" & Mid((Sheet1.Range("Sheet1.Duedate").value), 1, 2) & "/" & Mid((Sheet1.Range("Sheet1.Duedate").value), 7, 4) 'AY 2024-25

'Newly changed by Bindu

ItrFilingDueDate_1 = Mid((Sheet1.Range("Sheet1.Duedate").value), 7, 4) & "-" & Mid((Sheet1.Range("Sheet1.Duedate").value), 4, 2) & "-" & Mid((Sheet1.Range("Sheet1.Duedate").value), 1, 2) 'AY 2024-25

'PAG_C9 AY 2024-25 Schema implementation Ends------------------------





 'AY 2023-24 OLD Starts  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
 
  '  NewTaxRegime = Sheet8.Range("oldbacValue").value
  
'    OptingNewTaxRegime = Sheet1.Range("OptingNewTaxRegime").value
'    date10IE = Sheet1.Range("Sheet1.date10IE").value
'    acknowledgementNumber = Sheet1.Range("Sheet1.acknowledgementNumber").value
    
 'AY 2023-24 OLD Ends  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
 
' AY 2023-24 OLD Starts-------------------------------------------------------------------

'    AYOptedIn = Sheet1.Range("AYinOpting").value
'    NewForm10IEDate = Dformat(Sheet1.Range("Sheet1.date10IE_Opt").value, "")
'    NewForm10IEAckNo = Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value
'    OldTaxRegime = Sheet1.Range("Opt_Out115BAC").value
'    AYinOptOut = Sheet1.Range("AYinOptOut").value
'    OldForm10IEDate = Dformat(Sheet1.Range("Sheet1.date10IE_OptOut").value, "")
'    OldForm10IEAckNo = Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value

'AY 2023-24 OLD End Change------------------------------------------------------------------
     
     
     NewForm10IEADate_1 = Dformat(Sheet1.Range("Sheet1.date10IE_Opt").value, "")
     NewForm10IEAAckNo_1 = Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value
     
     
     

UpdateProgressBar

If ReturnFilledSection <> "" Then
    jsonDictionary("ReturnFileSec") = UVCase(CDbl(Trim(UCase(Mid(Trim(ReturnFilledSection), 1, InStr(ReturnFilledSection, "-") - 1)))))
End If
UpdateProgressBar






If ProvisoFlag = "Yes" Then
    ProvisoFlag = "Y"
ElseIf ProvisoFlag = "No" Then
    ProvisoFlag = "N"
End If



If ProvisoFlag <> "" Then
    jsonDictionary("SeventhProvisio139") = UCase(ProvisoFlag)
End If



If DepositAmountFlag = "Yes" Then
    DepositAmountFlag = "Y"
ElseIf DepositAmountFlag = "No" Then
    DepositAmountFlag = "N"
End If

If DepositAmountFlag <> "" Then
    jsonDictionary("DepAmtAggAmtExcd1CrPrYrFlg") = UCase(DepositAmountFlag)
End If

If DepositAmount <> "" Then
    jsonDictionary("AmtSeventhProvisio139i") = (DepositAmount)
End If

If AggrigateAmountFlag = "Yes" Then
    AggrigateAmountFlag = "Y"
ElseIf AggrigateAmountFlag = "No" Then
    AggrigateAmountFlag = "N"
End If

If AggrigateAmountFlag <> "" Then
    jsonDictionary("IncrExpAggAmt2LkTrvFrgnCntryFlg") = UCase(AggrigateAmountFlag)
End If

If AggrigateAmount <> "" Then
    jsonDictionary("AmtSeventhProvisio139ii") = (AggrigateAmount)
End If

If AggrigateAmountFlag1 = "Yes" Then
    AggrigateAmountFlag1 = "Y"
ElseIf AggrigateAmountFlag1 = "No" Then
    AggrigateAmountFlag1 = "N"
End If

If AggrigateAmountFlag1 <> "" Then
    jsonDictionary("IncrExpAggAmt1LkElctrctyPrYrFlg") = UCase(AggrigateAmountFlag1)
End If

If AggrigateAmount1 <> "" Then
    jsonDictionary("AmtSeventhProvisio139iii") = (AggrigateAmount1)
End If




'------------BACVALUE BASED ONE STARTS--------------------------------------
NewTaxRegime_s = Sheet8.Range("oldbacValue").value

If NewTaxRegime_s <> "" Then

           If NewTaxRegime_s = 1 Then
                
                  If TaxRegime_1 = "No" Then
                            jsonDictionary("OptOutNewTaxRegime") = "N"
                            ElseIf TaxRegime_1 = "Yes-beyond the due date" Then
                           jsonDictionary("OptOutNewTaxRegime") = "ybd"
                    End If
           ElseIf NewTaxRegime_s = 2 Then
            If TaxRegime_1 = "Yes-within the due date" Then
            jsonDictionary("OptOutNewTaxRegime") = "ywd"
            End If

           End If
 End If
'------------BACVALUE BASED ONE END--------------------------------------


'-------WITHOUT BACVALUE STARTS /////////////////////////////////////////////////

            If TaxRegime_1 = "No" Then
             jsonDictionary("OptOutNewTaxRegime") = "N"
                            
            ElseIf TaxRegime_1 = "Yes-but beyond the due date" Then
            jsonDictionary("OptOutNewTaxRegime") = "YBD"
                     
            
            ElseIf TaxRegime_1 = "Yes-within the due date" Then
            jsonDictionary("OptOutNewTaxRegime") = "YWD"
            
            ElseIf TaxRegime_1 = "Not applicable" Then
            jsonDictionary("OptOutNewTaxRegime") = "NA"
            End If
            
'-------WITHOUT BACVALUE STARTS ENDS /////////////////////////////////////////////////
            



If Mid(Sheet1.Range("sheet1.Status"), 1, 1) <> "F" Then

If NewTaxRegime <> "" Then
    If NewTaxRegime = 1 Then
        jsonDictionary("NewTaxRegime") = "Y"
    ElseIf NewTaxRegime = 2 Then
        jsonDictionary("NewTaxRegime") = "N"
    End If
End If


''TaxRegime_1 = Sheet1.Range("Sheet1.OptOutNewTaxRegime").value
'
''AY 2024-25 Change Starts Today---------------
'
'If TaxRegime_1 <> "" Then
'    If TaxRegime_1 = "No" Then
'        jsonDictionary("OptOutNewTaxRegime") = "N"
'    ElseIf TaxRegime_1 = "Yes-within the due date" Then
'        jsonDictionary("OptOutNewTaxRegime") = "N"
'End If
'
''AY 2024-25 Change Ends Today---------------



    If (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Then
        OptingNewTaxRegime = 1
    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Not opting") Then
        OptingNewTaxRegime = 2
    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt") Then
        OptingNewTaxRegime = 3
    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
        OptingNewTaxRegime = 4
    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in") Then
        OptingNewTaxRegime = 5
    End If
    
'Generate-08.12.2022.103.16
'If (Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in") Then
'    OptingNewTaxRegime = 5
'End If
'----End Change

    jsonDictionary("OptingNewTaxRegime") = UVCase(CDbl(OptingNewTaxRegime))
    
End If


'Generate-08.12.2022.103.09B

If Sheet8.Range("oldbacValue").value = "1" Then
        Dim NewTaxRegimeDtls, Form10IEDtls
        Set NewTaxRegimeDtls = CreateObject("Scripting.Dictionary")
        Set Form10IEDtls = CreateObject("Scripting.Dictionary")
        NewTaxRegimeDtls("AssessmentYear") = AYOptedIn
        Form10IEDtls("Form10IEDate") = NewForm10IEDate
        Form10IEDtls("Form10IEAckNo") = NewForm10IEAckNo
        NewTaxRegimeDtls.add "Form10IEDtls", Form10IEDtls
        jsonDictionary.add "NewTaxRegimeDtls", NewTaxRegimeDtls
        
        Set Form10IEDtls = Nothing
        Set Form10IEDtls = CreateObject("Scripting.Dictionary")
        Set NewTaxRegimeDtls = Nothing
        
        If Sheet1.Range("Opt_Out115BAC") = "Yes" Then
            Dim OptedOutNewTaxRegimeDtls
            Set OptedOutNewTaxRegimeDtls = CreateObject("Scripting.Dictionary")
            jsonDictionary("OptedOutNewTaxRegime") = "Y"
            OptedOutNewTaxRegimeDtls("AssessmentYear") = AYinOptOut
            Form10IEDtls("Form10IEDate") = OldForm10IEDate
            Form10IEDtls("Form10IEAckNo") = OldForm10IEAckNo
            OptedOutNewTaxRegimeDtls.add "Form10IEDtls", Form10IEDtls
            jsonDictionary.add "OptedOutNewTaxRegimeDtls", OptedOutNewTaxRegimeDtls
            Set OptedOutNewTaxRegimeDtls = Nothing
        ElseIf Sheet1.Range("Opt_Out115BAC") = "No" Then
            jsonDictionary("OptedOutNewTaxRegime") = "N"
        End If
        
End If

'---End Change




If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 18 Then
    jsonDictionary("ReceiptNo") = RecieptNo
    jsonDictionary("NoticeNo") = NoticeNo
    jsonDictionary("OrigRetFiledDate") = Dformat(DateOfFillingRetrun, "")
End If

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 13 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 14 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 15 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 16 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
    jsonDictionary("NoticeNo") = NoticeNo
End If

UpdateProgressBar

If (Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 17) Or (Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 21) Then
    jsonDictionary("ReceiptNo") = RecieptNo
    If Trim(Sheet1.Range("sheet1.OrigRetFiledDate").value) <> "" Then
        jsonDictionary("OrigRetFiledDate") = Dformat(DateOfFillingRetrun, "")
    End If
End If

UpdateProgressBar

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 13 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 14 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 15 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 16 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 18 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
    jsonDictionary("NoticeDateUnderSec") = Dformat(DateOfNotice, "")
End If

If isdropdownblank(Representativeassesseeflg) Then

Else
    jsonDictionary("AsseseeRepFlg") = Mid(Trim(UCase(Representativeassesseeflg)), 1, 1)
End If
If date10IE <> "" Then
    jsonDictionary("Form10IEDate") = Dformat(date10IE, "")
End If
If acknowledgementNumber <> "" Then
    jsonDictionary("Form10IEAckNo") = UVCase(CDbl(acknowledgementNumber))
End If



'139 clause iv

If Sheet1.Range("clauseiv7provisio139iFlg") = "Yes" Then
        jsonDictionary("clauseiv7provisio139i") = "Y"
       If Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "1"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_1")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "2"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_2")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "3"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_3")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "4"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_4")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       jsonDictionary.add "clauseiv7provisio139iDtls", clauseiv7provisio139iDtls
    ElseIf Sheet1.Range("clauseiv7provisio139iFlg") = "No" Then
        jsonDictionary("clauseiv7provisio139i") = "N"
    End If
    
    





If Mid(Trim(UCase(Representativeassesseeflg)), 1, 1) = "Y" Then

    If NameRepAssessee <> "" Then
    AssesseeRep("RepName") = UCase(NameRepAssessee)
    End If
 'Konda updated on 22-01-2026 schema-V0.2.5
'    If Not isdropdownblank(CapacityRep) Then
'    AssesseeRep("RepCapacity") = Mid(Trim(UCase(CapacityRep)), 1, 1)
'    End If
'
'    If AddressRepAssessee <> "" Then
'    AssesseeRep("RepAddress") = UCase(AddressRepAssessee)
'    End If
'
'    If PANRepAssessee <> "" Then
'    AssesseeRep("RepPAN") = UCase(PANRepAssessee)
'    End If
'
'    If AadhaarRepAssessee <> "" Then
'    AssesseeRep("RepAadhaar") = UCase(AadhaarRepAssessee)
'    End If


    If EmailRepAssessee <> "" Then
    AssesseeRep("RepEmailID") = UCase(EmailRepAssessee)
    End If

    If CountryCodeRepAssessee <> "" Then
    AssesseeRep("CountryCodeRepMobileNo") = Int(CountryCodeRepAssessee)
    End If

    If MobileRepAssessee <> "" Then
    AssesseeRep("RepMobileNo") = Int(MobileRepAssessee)
    End If
'==========================================

jsonDictionary.add "AssesseeRep", AssesseeRep
End If

jsonDictionary("ItrFilingDueDate") = ItrFilingDueDate 'PAG_C9 AY 2024-25

UpdateProgressBar

Set FilingStatus = jsonDictionary

End Function
Function PartA_139_8A() As Object
On Error Resume Next
subProcCaption = "PartA_139_8A"
noOfProcessSub = 10
UpdateProgressBar
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

If Sheet201.Range("U_AadhaarCardNo").value <> "" Then
jsonDictionary("AadhaarCardNo") = Sheet201.Range("U_AadhaarCardNo").value
End If
'Ayush_enrol
'If Sheet201.Range("U_AadhaarEnrolmentId").value <> "" Then
'jsonDictionary("AadhaarEnrolmentId") = Sheet201.Range("U_AadhaarEnrolmentId").value
'End If

If Sheet201.Range("U_Name").value <> "" Then
jsonDictionary("Name") = Sheet201.Range("U_Name").value
End If
If Sheet201.Range("U_PAN").value <> "" Then
jsonDictionary("PAN") = Sheet201.Range("U_PAN").value
End If
UpdateProgressBar

'jsonDictionary("AssessmentYear") = "2022"
'jsonDictionary("AssessmentYear") = "2023"

'jsonDictionary("AssessmentYear") = "2024" 'AY 2024-25

jsonDictionary("AssessmentYear") = "2025" 'AY 2025-26

If Sheet201.Range("U_PreviouslyFiledForThisAY").value <> "" Then
jsonDictionary("PreviouslyFiledForThisAY") = Mid(Sheet201.Range("U_PreviouslyFiledForThisAY").value, 1, 1)
End If

If Mid(Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value, 1, 1) = "1" Then
    jsonDictionary("PreviouslyFiledForThisAY_139_8A") = "1"
ElseIf Mid(Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value, 1, 1) = "O" Then
    jsonDictionary("PreviouslyFiledForThisAY_139_8A") = "2"
End If
Dim Applicable_139_8A As Object
Set Applicable_139_8A = CreateObject("Scripting.Dictionary")

If Sheet201.Range("U_ITRForm").value <> "" Then
Applicable_139_8A("ITRForm") = Sheet201.Range("U_ITRForm").value
End If

UpdateProgressBar
If Sheet201.Range("U_AcknowledgementNo").value <> "" Then
Applicable_139_8A("AcknowledgementNo") = Sheet201.Range("U_AcknowledgementNo").value
End If

If Sheet201.Range("U_OrigRetFiledDate").value <> "" Then
Applicable_139_8A("OrigRetFiledDate") = Dformat(Sheet201.Range("U_OrigRetFiledDate").value, "")
End If

If Applicable_139_8A.count > 0 Then
    jsonDictionary.add "Applicable_139_8A", Applicable_139_8A
End If
UpdateProgressBar

If Mid(Sheet201.Range("U_LaidOutIn_139_8A").value, 1, 1) <> "" Then
jsonDictionary("LaidOutIn_139_8A") = Mid(Sheet201.Range("U_LaidOutIn_139_8A").value, 1, 1)
End If


jsonDictionary("ITRFormUpdatingInc") = "ITR4"

Dim row As Range
Dim UpdatingInc As Collection
Dim UpdatingIncDtls, UpdatingIncObj As Object
Set UpdatingIncDtls = CreateObject("Scripting.Dictionary")
Set UpdatingInc = New Collection
Set UpdatingIncObj = CreateObject("Scripting.Dictionary")

For Each row In Sheet201.Range("U_ReasonsForUpdatingIncome").Rows
    If row.Cells(1).value = "" Or row.Cells(1).value = "(Select)" Then Exit For
    Select Case row.Cells(1).value
    
    Case "Return previously not filed"
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "1"
    
    Case "Income not reported correctly"
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "2"
    
    Case "Wrong heads of income chosen"
    
    UpdatingIncDtls("ReasonsForUpdatingIncome") = "3"
    
    Case "Reduction of carried forward loss"
    
    UpdatingIncDtls("ReasonsForUpdatingIncome") = "4"
    
    Case "Reduction of unabsorbed depreciation"
    
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "5"
    
    Case "Reduction of tax credit u/s 115JB/115JC"
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "6"
        
    Case "Wrong rate of tax"
            
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "7"
            
    Case "Others "
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "OTH"
    
    End Select
    
    UpdatingInc.add UpdatingIncDtls
    Set UpdatingIncDtls = CreateObject("Scripting.Dictionary")
Next row
If UpdatingInc.count > 0 Then
    UpdatingIncObj.add "ReasonsForUpdatingIncDtls", UpdatingInc
    jsonDictionary.add "UpdatingInc", UpdatingIncObj
End If
UpdateProgressBar
If Sheet201.Range("U_UpdatedReturnDuringPeriod").value <> "" Then
    Select Case Sheet201.Range("U_UpdatedReturnDuringPeriod").value
    Case "Up to 12 months from the end of Relevant Assessment Year"
        jsonDictionary("UpdatedReturnDuringPeriod") = "1"
    Case "Between 12 to 24 Months from the end of Relevant Assessment  Year"
        jsonDictionary("UpdatedReturnDuringPeriod") = "2"
        
        Case "Between 24 to 36 Months from the end of Relevant Assessment  Year"
        jsonDictionary("UpdatedReturnDuringPeriod") = "3"
        
        Case "Between 36 to 48 Months from the end of Relevant Assessment  Year"
        jsonDictionary("UpdatedReturnDuringPeriod") = "4"
    End Select
    
    
    
End If
UpdateProgressBar
Dim RetrntoRedCarriedFL As Object
Set RetrntoRedCarriedFL = CreateObject("Scripting.Dictionary")
If Sheet201.Range("U_UnabsorbedDepreciation").value <> "" Then
RetrntoRedCarriedFL("UnabsorbedDepreciation") = Mid(Sheet201.Range("U_UnabsorbedDepreciation").value, 1, 1)
End If
Dim UDYear As Collection
Dim UDYearObj As Object
Dim UnabsorbedDepreciationYearDtls As Object
Set UDYear = New Collection
Set UDYearObj = CreateObject("Scripting.Dictionary")
Set UnabsorbedDepreciationYearDtls = CreateObject("Scripting.Dictionary")
UpdateProgressBar

For Each row In Sheet201.Range("U_UnabsorbedDepreciationYear").Rows
    If row.Cells(1).value = "" Or row.Cells(1).value = "(Select)" Then Exit For
    
'    If row.Cells(1).value = "2023-24" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2023"
'    ElseIf row.Cells(1).value = "2024-25" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2024"
'    End If
    'Ayush_UR
'    If row.Cells(1).value = "2024-25" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2024"
'    ElseIf row.Cells(1).value = "2025-26" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2025"
'    End If


'    If row.Cells(1).value = "2026-27" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2026"
'    ElseIf row.Cells(1).value = "2025-26" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2025"
'    End If
'------------
        If row.Cells(1).value = "2027-28" Then
        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2027"
    ElseIf row.Cells(1).value = "2026-27" Then
        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2026"
    End If
    
    
    
    If row.Cells(1).Offset(0, 1).value <> "" Then
'        UnabsorbedDepreciationYearDtls("ReturnFiledEffectFlg") = Mid(row.Cells(1).Offset(0, 1).value, 1, 1)
         UnabsorbedDepreciationYearDtls("RevisedReturnFile") = Mid(row.Cells(1).Offset(0, 1).value, 1, 1)
    End If
    
    If row.Cells(1).Offset(0, 16).value <> "" Then
        UnabsorbedDepreciationYearDtls("UpdatedReturnFile") = Mid(row.Cells(1).Offset(0, 16).value, 1, 1)
    End If
    
    
'    If row.Cells(1).Offset(0, 16).Value <> "" And row.Cells(1).Offset(0, 16).Value <> "(Select)" Then
'        If row.Cells(1).Offset(0, 16).Value = "Orginal Return 139(1)/139(4)" Then
'            UnabsorbedDepreciationYearDtls("OriginalReturn") = "Y"
'        ElseIf row.Cells(1).Offset(0, 16).Value = "Revised return" Then
'            UnabsorbedDepreciationYearDtls("RevisedReturnFile") = "Y"
'        ElseIf row.Cells(1).Offset(0, 16).Value = "Updated return" Then
'            UnabsorbedDepreciationYearDtls("UpdatedReturnFile") = "Y"
'        End If
'
'    End If
'"Original Return 139(1)/139(4)"
'    If row.Cells(1).Offset(0, 16).value <> "" And row.Cells(1).Offset(0, 16).value <> "(Select)" Then
'        If row.Cells(1).Offset(0, 16).value = "Original Return 139(1)/139(4)" Then
'            UnabsorbedDepreciationYearDtls("ReturnType") = "1"
'        ElseIf row.Cells(1).Offset(0, 16).value = "Revised return" Then
'            UnabsorbedDepreciationYearDtls("ReturnType") = "2"
'        ElseIf row.Cells(1).Offset(0, 16).value = "Updated return" Then
'            UnabsorbedDepreciationYearDtls("ReturnType") = "3"
'        End If
'
'    End If
    
    
    
    
    UDYear.add UnabsorbedDepreciationYearDtls
    Set UnabsorbedDepreciationYearDtls = CreateObject("Scripting.Dictionary")
    UpdateProgressBar
Next row

If UDYear.count > 0 Then
    UDYearObj.add "UnabsorbedDepreciationYearDtls", UDYear
    RetrntoRedCarriedFL.add "UDYear", UDYearObj
End If
If RetrntoRedCarriedFL.count > 0 Then
    jsonDictionary.add "RetrntoRedCarriedFL", RetrntoRedCarriedFL
End If
UpdateProgressBar
Set PartA_139_8A = jsonDictionary

End Function
Function PartB_ATI() As Object
subProcCaption = "PartB_ATI"
noOfProcessSub = 10
Dim jsonDictionary, HeadOfInc As Object
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set HeadOfInc = CreateObject("Scripting.Dictionary")

UpdateProgressBar

HeadOfInc("Salaries") = UVCase(Sheet202.Range("U_Salaries").value)
HeadOfInc("IncomeFromHP") = UVCase(Sheet202.Range("U_IncomeFromHP").value)
HeadOfInc("IncomeFromBP") = UVCase(Sheet202.Range("U_IncomeFromBP").value)
'HeadOfInc("IncomeFromCG") = UVCase(Sheet202.Range("U_IncomeFromCG").value)
HeadOfInc("IncomeFromOS") = UVCase(Sheet202.Range("U_IncomeFromOS").value)
UpdateProgressBar
HeadOfInc("Total") = UVCase(Sheet202.Range("U_Total").value)
jsonDictionary.add "HeadOfInc", HeadOfInc

jsonDictionary("LatestTotInc") = UVCase(Sheet202.Range("U_LatestTotInc").value)
jsonDictionary("UpdatedTotInc") = UVCase(Sheet202.Range("U_UpdatedTotInc").value)
UpdateProgressBar
jsonDictionary("AmtPayable") = UVCase(Sheet202.Range("U_AmtPayable").value)
jsonDictionary("AmtRefundable") = UVCase(Sheet202.Range("U_AmtRefundable").value)
jsonDictionary("LastAmtPayable") = UVCase(Sheet202.Range("U_LastAmtPayable").value)
jsonDictionary("Refund") = UVCase(Sheet202.Range("U_Refund").value)
UpdateProgressBar
jsonDictionary("TotRefund") = UVCase(Sheet202.Range("U_TotRefund").value)
jsonDictionary("FeeIncUS234F") = UVCase(Sheet202.Range("U_FeeIncUS234F").value)
jsonDictionary("RegAssessementTAX") = UVCase(Sheet202.Range("U_RegAssessementTAX").value)
jsonDictionary("AggrLiabilityRefund") = UVCase(Sheet202.Range("U_AggrLiabilityRefund").value)
jsonDictionary("AggrLiabilityNoRefund") = UVCase(Sheet202.Range("U_AggrLiabilityNoRefund").value)
jsonDictionary("AddtnlIncTax") = UVCase(Sheet202.Range("U_AddtnlIncTax").value)
jsonDictionary("NetPayable") = UVCase(Sheet202.Range("U_NetPayable").value)
jsonDictionary("TaxUS140B") = UVCase(Sheet202.Range("U_TaxUS140B").value)
UpdateProgressBar
jsonDictionary("TaxDue10_11") = UVCase(Sheet202.Range("U_TaxDue10_11").value)
UpdateProgressBar
Dim cell As Range
Dim TaxPayment1 As Collection
Dim ITTaxPayments, ScheduleIT1 As Object
Dim TaxPayment1Obj As Object
Set TaxPayment1Obj = CreateObject("Scripting.Dictionary")
Set TaxPayment1 = New Collection
Set ITTaxPayments = CreateObject("Scripting.Dictionary")
Set ScheduleIT1 = CreateObject("Scripting.Dictionary")
Dim itemp
itemp = 0
For Each cell In Sheet202.Range("U_BSRCode1").Cells
    itemp = itemp + 1
    If cell.value = "" Or cell.value = "(Select)" Then Exit For
    ITTaxPayments("slno") = itemp
    ITTaxPayments("BSRCode") = cell.value
    ITTaxPayments("DateDep") = Dformat(cell.Offset(0, 1).value, "")
    ITTaxPayments("SrlNoOfChaln") = cell.Offset(0, 2).value
    ITTaxPayments("Amt") = cell.Offset(0, 5).value
    TaxPayment1.add ITTaxPayments
    Set ITTaxPayments = CreateObject("Scripting.Dictionary")
    
Next cell
UpdateProgressBar
If TaxPayment1.count > 0 Then
    
    TaxPayment1Obj.add "ITTaxPayments", TaxPayment1
    ScheduleIT1.add "TaxPayment1", TaxPayment1Obj
    ScheduleIT1("Total") = Sheet202.Range("U_Total1").value
    jsonDictionary.add "ScheduleIT1", ScheduleIT1
End If

Dim TaxPayment2Obj As Object
Set TaxPayment2Obj = CreateObject("Scripting.Dictionary")
Dim TaxPayment2 As Collection
Set TaxPayment2 = New Collection
Dim ScheduleIT2 As Object
Set ScheduleIT2 = CreateObject("Scripting.Dictionary")
itemp = 0
For Each cell In Sheet202.Range("U_BSRCode2").Cells
    If cell.value = "" Or cell.value = "(Select)" Then Exit For
    itemp = itemp + 1
    ITTaxPayments("slno") = itemp
    ITTaxPayments("BSRCode") = cell.value
    ITTaxPayments("DateDep") = Dformat(cell.Offset(0, 1).value, "")
    ITTaxPayments("SrlNoOfChaln") = cell.Offset(0, 2).value
    ITTaxPayments("Amt") = cell.Offset(0, 5).value
    TaxPayment2.add ITTaxPayments
    Set ITTaxPayments = CreateObject("Scripting.Dictionary")
Next cell
UpdateProgressBar
If TaxPayment2.count > 0 Then
    TaxPayment2Obj.add "ITTaxPayments", TaxPayment2
    ScheduleIT2.add "TaxPayment2", TaxPayment2Obj
    ScheduleIT2("Total") = Sheet202.Range("U_Total2").value
    jsonDictionary.add "ScheduleIT2", ScheduleIT2
End If


jsonDictionary("ReleifUS89") = UVCase(Sheet202.Range("U_ReleifUS89").value)


Set PartB_ATI = jsonDictionary

End Function
Function ITR4_IncomeDeductions() As Object
Dim i As Long
Dim jsonDictionary, NOT89AType, NOT89AInc, UsrDeductUndChapVIA, NOT89A, IncomeNotified89AType, DeductUndChapVIA, AllwncExemptUs10, AllwncExemptUs10Dtl, DividendInc, OthersInc, DateRange, ExemptIncAgriOthUs10, OthersIncDtlsOth, ExemptIncAgriOthUs10Dtl
Dim AllwncExemptUs10Dtls, TotalAllwncExemptUs10, OthersIncDtlsOthSrc, ExemptIncAgriOthUs10Dtls, ExemptIncAgriOthUs10Total
Dim Upto15Of6, Upto15Of9, Up16Of9To15Of12, Up16Of12To15Of3, Up16Of3To31Of3, TotDividendInc
Dim str

Set AllwncExemptUs10Dtls = New Collection
Set OthersIncDtlsOthSrc = New Collection
Set ExemptIncAgriOthUs10Dtls = New Collection
Set IncomeNotified89AType = New Collection
Set NOT89A = New Collection

'New Schema updated by Konda as AY-2026-27 on 29-12-2025
Dim PensionContribution80CCC, PensionContributionFund, PensionContributionFundtls, PensionContribution80CCD1, PensionContribution80CCD1B
Set PensionContribution80CCC = New Collection
Set PensionContribution80CCD1 = New Collection
Set PensionContribution80CCD1B = New Collection
Set PensionContributionFund = CreateObject("scripting.dictionary")
Set PensionContributionFundtls = CreateObject("scripting.dictionary")

Set NOT89AInc = CreateObject("scripting.dictionary")
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set UsrDeductUndChapVIA = CreateObject("Scripting.Dictionary")
Set DeductUndChapVIA = CreateObject("Scripting.Dictionary")
Set AllwncExemptUs10 = CreateObject("Scripting.Dictionary")
Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
Set DividendInc = CreateObject("Scripting.Dictionary")
Set OthersInc = CreateObject("Scripting.Dictionary")
Set DateRange = CreateObject("Scripting.Dictionary")
Set ExemptIncAgriOthUs10 = CreateObject("Scripting.Dictionary")
Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
Set ExemptIncAgriOthUs10Dtl = CreateObject("Scripting.Dictionary")

Set NOT89AType = CreateObject("Scripting.Dictionary")

Upto15Of6 = Sheet1.Range("IncD_dq1div").value
Upto15Of9 = Sheet1.Range("IncD_dq2div").value
Up16Of9To15Of12 = Sheet1.Range("IncD_dq3div").value
Up16Of12To15Of3 = Sheet1.Range("IncD_dq4div").value
Up16Of3To31Of3 = Sheet1.Range("IncD_dq5div").value
TotDividendInc = Sheet1.Range("IncD_dqdiv").value

    If IncomeBP <> "" Then
        jsonDictionary("IncomeFromBusinessProf") = IncomeBP
    Else
        jsonDictionary("IncomeFromBusinessProf") = 0
    End If

    If IncomeSAl1 <> "" Then
        jsonDictionary("GrossSalary") = IncomeSAl1
     Else
        jsonDictionary("GrossSalary") = 0
    End If

    If IncomeSAl2 <> "" Then
        jsonDictionary("Salary") = IncomeSAl2
     Else
        jsonDictionary("Salary") = 0
    End If
    
    If IncomeSAl3 <> "" Then
        jsonDictionary("PerquisitesValue") = IncomeSAl3
     Else
        jsonDictionary("PerquisitesValue") = 0
    End If
                 
    If IncomeSAl4 <> "" Then
        jsonDictionary("ProfitsInSalary") = IncomeSAl4
     Else
        jsonDictionary("ProfitsInSalary") = 0
    End If
    
'Commented by  'Konda on 22-01-2026 schema-V0.2.5
'    If Trim(Sheet1.Range("IncomeNotified89A_AmountUS").value) <> "" Then
'        NOT89AType("NOT89ACountrycode") = "US"
'        NOT89AType("NOT89AAmount") = Sheet1.Range("IncomeNotified89A_AmountUS").value
'        IncomeNotified89AType.add NOT89AType
'        Set NOT89AType = Nothing
'        Set NOT89AType = CreateObject("Scripting.Dictionary")
'    End If
'    If Trim(Sheet1.Range("IncomeNotified89A_AmountUK").value) <> "" Then
'        NOT89AType("NOT89ACountrycode") = "UK"
'        NOT89AType("NOT89AAmount") = Sheet1.Range("IncomeNotified89A_AmountUK").value
'        IncomeNotified89AType.add NOT89AType
'        Set NOT89AType = Nothing
'        Set NOT89AType = CreateObject("Scripting.Dictionary")
'    End If
'    If Trim(Sheet1.Range("IncomeNotified89A_AmountCan").value) <> "" Then
'        NOT89AType("NOT89ACountrycode") = "CA"
'        NOT89AType("NOT89AAmount") = Sheet1.Range("IncomeNotified89A_AmountCan").value
'        IncomeNotified89AType.add NOT89AType
'        Set NOT89AType = Nothing
'        Set NOT89AType = CreateObject("Scripting.Dictionary")
'    End If
'    jsonDictionary.add "IncomeNotified89AType", IncomeNotified89AType
'    Sheet1.Unprotect Password:=getmsgstate
'    If Trim(Sheet1.Range("IncomeNotified89A").value) <> "" Then
'        jsonDictionary("IncomeNotified89A") = Sheet1.Range("IncomeNotified89A").value
'     Else
'        jsonDictionary("IncomeNotified89A") = 0
'    End If
'    Sheet1.Protect Password:=getmsgstate
'    If Trim(Sheet1.Range("IncomeNotifiedOther89A").value) <> "" Then
'        jsonDictionary("IncomeNotifiedOther89A") = Sheet1.Range("IncomeNotifiedOther89A").value
'     Else
'        jsonDictionary("IncomeNotifiedOther89A") = 0
'    End If
'    If Trim(Sheet1.Range("Increliefus89A").value) <> "" Then
'        jsonDictionary("Increliefus89A") = Sheet1.Range("Increliefus89A").value
'     Else
'        jsonDictionary("Increliefus89A") = 0
'    End If

'    If Trim(Sheet1.Range("OSIncreliefus89A").value) <> "" Then
'        jsonDictionary("Increliefus89AOS") = Sheet1.Range("OSIncreliefus89A").value
'     Else
'        jsonDictionary("Increliefus89AOS") = 0
'    End If
''===========================================



   'Malli---Sheet1.HRA--------08/05/2025
   
   ' If (Not IsEmpty(Others_IDNOI)) And UBound(Others_IDNOI) > 0 Then
    If (Not IsEmpty(Others_IDNOI)) Or Sheet1.Range("Sheet1.HRA") > 0 Then
     If (UBound(Others_IDNOI) > 0) Or Sheet1.Range("Sheet1.HRA") > 0 Then
    For i = 1 To end_IDOthersNOI
'Change-24.02.2023.101 - Validation Framework
            If (Others_IDNOI(i) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette") Or (Others_IDNOI(i) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette ") Then
                Others_IDNOI(i) = "10(10B)(i)"
'---
            ElseIf (Others_IDNOI(i) = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government") Then
                Others_IDNOI(i) = "10(10B)(ii)"
            ElseIf (Others_IDNOI(i) = "Sec 10(14)(i)-Allowances referred in sub-clauses (a) to (c) of sub-rule (1) in Rule 2BB") Then
                Others_IDNOI(i) = "10(14)(i)(115BAC)"
            ElseIf (Others_IDNOI(i) = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee") Then
                Others_IDNOI(i) = "10(14)(ii)(115BAC)"
'            ElseIf (Others_IDNOI(i) = "Any Other") Then
'                Others_IDNOI(i) = "OTH"
    'Konda updated on 09--3-2026--V0.5
            ElseIf (Others_IDNOI(i) = "Sec 10(17)-Allowance MP/MLA/MLC") Then
                Others_IDNOI(i) = "10(17)"
    '=================================
            ElseIf (Others_IDNOI(i) = "Exempt income received by a judge covered under the payment of salaries to Supreme Court/High Court judges Act /Rules") Then
                Others_IDNOI(i) = "EIC"
            Else
                If InStr(Others_IDNOI(i), "-") > 0 Then
                    Others_IDNOI(i) = Mid(Others_IDNOI(i), 4, (InStr(Others_IDNOI(i), "-") - 4))
                End If
            End If
        Next
        
        'Newly added by Malli as per DE V7
                If Sheet1.Range("Sheet1.HRA") > 0 Then
                   AllwncExemptUs10Dtl("SalNatureDesc") = "10(13A)"
                   AllwncExemptUs10Dtl("SalOthAmount") = Sheet1.Range("Sheet1.HRA")
                   
                   AllwncExemptUs10Dtls.add AllwncExemptUs10Dtl
                   Set AllwncExemptUs10Dtl = Nothing
                   Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
                 End If
                 
'                AllwncExemptUs10Dtls.add AllwncExemptUs10Dtl
'                Set AllwncExemptUs10Dtl = Nothing
'                Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
'                '----------------------------------
 
        For i = 1 To UBound(Others_IDNOI)
                If Others_IDNOI(i) <> "" Then
                    AllwncExemptUs10Dtl("SalNatureDesc") = Trim((Others_IDNOI(i)))
                Else
                    AllwncExemptUs10Dtl("SalNatureDesc") = ""
                End If
'Konda updated on 09-03-2026--V0.5
'                If Others_IDNOI1(i) <> "" And Others_IDNOI1(i) <> "Not Applicable" Then
'                    AllwncExemptUs10Dtl("SalOthNatOfInc") = (Others_IDNOI1(i))
'                End If
'======================================
                If Others_IDAmt(i) <> "" Then
                    AllwncExemptUs10Dtl("SalOthAmount") = Others_IDAmt(i)
                Else
                    AllwncExemptUs10Dtl("SalOthAmount") = ""
                End If
                AllwncExemptUs10Dtls.add AllwncExemptUs10Dtl
                Set AllwncExemptUs10Dtl = Nothing
                Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
        Next
        If Trim(Sheet1.Range("IncD.IncomeFromSal_allow").value) <> "" Then
            AllwncExemptUs10("TotalAllwncExemptUs10") = Sheet1.Range("IncD.IncomeFromSal_allow").value
        Else
            AllwncExemptUs10("TotalAllwncExemptUs10") = 0
        End If
        AllwncExemptUs10.add "AllwncExemptUs10Dtls", AllwncExemptUs10Dtls
        jsonDictionary.add "AllwncExemptUs10", AllwncExemptUs10
        End If  'Malli
    End If





    If IncomeSAlnet <> "" Then
        jsonDictionary("NetSalary") = IncomeSAlnet
     Else
        jsonDictionary("NetSalary") = 0
    End If
    
    If IncomeSAl6 <> "" Then
        jsonDictionary("DeductionUs16") = IncomeSAl6
     Else
        jsonDictionary("DeductionUs16") = 0
    End If
    
    If IncomeSAl5b <> "" Then
        jsonDictionary("DeductionUs16ia") = IncomeSAl5b
     Else
        jsonDictionary("DeductionUs16ia") = 0
    End If
    
    If IncomeSAl5 <> "" Then
        jsonDictionary("EntertainmntalwncUs16ii") = IncomeSAl5
     Else
        jsonDictionary("EntertainmntalwncUs16ii") = 0
    End If
    
    If IncomeSAl5c <> "" Then
        jsonDictionary("ProfessionalTaxUs16iii") = IncomeSAl5c
     Else
        jsonDictionary("ProfessionalTaxUs16iii") = 0
    End If
    
    If IncomeSAl <> "" Then
        jsonDictionary("IncomeFromSal") = IncomeSAl
     Else
        jsonDictionary("IncomeFromSal") = 0
    End If
    
    UpdateProgressBar
    
'Commented by  'Konda on 22-01-2026 schema-V0.2.5
'    If TypeOfHP <> "" And Mid(Trim(TypeOfHP), 1, 1) <> "(" Then
'        jsonDictionary("TypeOfHP") = Trim(UCase(Mid(TypeOfHP, 1, 1)))
'    End If
'
'    If IncomeHP1 <> "" Then
'        jsonDictionary("GrossRentReceived") = IncomeHP1
'    Else
'        jsonDictionary("GrossRentReceived") = 0
'    End If
'
'    If IncomeHP2 <> "" Then
'        jsonDictionary("TaxPaidlocalAuth") = IncomeHP2
'    Else
'        jsonDictionary("TaxPaidlocalAuth") = 0
'    End If
'
'    If IncomeHP3 <> "" Then
'        jsonDictionary("AnnualValue") = IncomeHP3
'    Else
'        jsonDictionary("AnnualValue") = 0
'    End If
'
'    If IncomeHP4 <> "" Then
'        jsonDictionary("AnnualValue30Percent") = IncomeHP4
'    Else
'        jsonDictionary("AnnualValue30Percent") = 0
'    End If
'
'    If IncomeHP5 <> "" Then
'        jsonDictionary("InterestPayable") = IncomeHP5
'    Else
'        jsonDictionary("InterestPayable") = 0
'    End If
'
'    If IncomeHP6 <> "" Then
'        jsonDictionary("ArrearsUnrealizedRentRcvd") = IncomeHP6
'     Else
'        jsonDictionary("ArrearsUnrealizedRentRcvd") = 0
'    End If
 
'    UpdateProgressBar
'    If IncomeHP <> "" Then
'        jsonDictionary("TotalIncomeOfHP") = IncomeHP
'     Else
'        jsonDictionary("TotalIncomeOfHP") = 0
'    End If
'
'    UpdateProgressBar
'=========================================================
'Konda updated HP on 22-01-2026

Dim m As Long
 Dim PropertyDetails, PropertyDetailsobj, AddressDetailWithZipCode, CoOwners, CoOwnersobj, TenantDetails, TenantDetailsobj, Rentdetails
 
 Set PropertyDetailsobj = CreateObject("Scripting.Dictionary")
 Set AddressDetailWithZipCode = CreateObject("Scripting.Dictionary")
 Set CoOwnersobj = CreateObject("Scripting.Dictionary")
 Set TenantDetailsobj = CreateObject("Scripting.Dictionary")
 Set Rentdetails = CreateObject("Scripting.Dictionary")
 Set PropertyDetails = New Collection
 Set CoOwners = New Collection
 Set TenantDetails = New Collection
 
 subProcCaption = "Schedule HP"
   
    If (Not IsEmpty(AddrDetail_HP)) Then
    If (UBound(AddrDetail_HP) > 0) Then
        noOfProcessSub = UBound(AddrDetail_HP)
        For m = 1 To UBound(AddrDetail_HP)
            If Len(AddrDetail_HP(m)) = 0 Then
                Exit For
            End If
            PropertyDetailsobj("HPSNo") = m
            If AddrDetail_HP(m) <> "" Then
                AddressDetailWithZipCode("AddrDetail") = UCase(AddrDetail_HP(m))
            Else
                AddressDetailWithZipCode("AddrDetail") = ""
            End If
            
            If CityOrTownOrDistrict_HP(m) <> "" Then
                AddressDetailWithZipCode("CityOrTownOrDistrict") = UCase(CityOrTownOrDistrict_HP(m))
            Else
                AddressDetailWithZipCode("CityOrTownOrDistrict") = ""
            End If
            
            If StateCode_HP(m) <> "" Then
                AddressDetailWithZipCode("StateCode") = UCase(StateCode_HP(m))
            End If
            
            If CountryCode_HP(m) <> "" Then
                AddressDetailWithZipCode("CountryCode") = UCase(CountryCode_HP(m))
            End If
            
            If CountryCode_HP(m) = "91" Then
                If PinCode_HP(m) <> "" Then
                    AddressDetailWithZipCode("PinCode") = UVCase(CDbl(PinCode_HP(m)))
                End If
            End If
            
            If CountryCode_HP(m) <> "91" Then
                If ZipCode_HP(m) <> "" Then
                    AddressDetailWithZipCode("ZipCode") = UCase(ZipCode_HP(m))
                Else
                    AddressDetailWithZipCode("ZipCode") = " "
                End If
            End If
            PropertyDetailsobj.add "AddressDetailWithZipCode", AddressDetailWithZipCode
            Set AddressDetailWithZipCode = Nothing
            Set AddressDetailWithZipCode = CreateObject("Scripting.Dictionary")

         If OwnerProperty_HP(m) <> "" Then
            If (OwnerProperty_HP(m)) = "Self" Then
               OwnerProperty_HP(m) = "SE"
            ElseIf (OwnerProperty_HP(m)) = "Minor" Then
               OwnerProperty_HP(m) = "MI"
            ElseIf (OwnerProperty_HP(m)) = "Spouse" Then
               OwnerProperty_HP(m) = "SP"
            ElseIf (OwnerProperty_HP(m)) = "Others" Then
               OwnerProperty_HP(m) = "OT"
            End If
           
            If OwnerProperty_HP(m) <> "" Then
            PropertyDetailsobj("PropertyOwner") = UCase(OwnerProperty_HP(m))
            End If
            
            If OwnerPropertyDescription_HP(m) <> "" Then
            PropertyDetailsobj("PropertyOwnerOther") = UCase(OwnerPropertyDescription_HP(m))
            End If
            
            If CoOwnedYN_HP(m) <> "" Then
            PropertyDetailsobj("PropCoOwnedFlg") = UCase(CoOwnedYN_HP(m))
            End If
            
            If CoOwnedShare_HP(m) <> "" Then
            PropertyDetailsobj("AsseseeShareProperty") = CoOwnedShare_HP(m)
            Else
            PropertyDetailsobj("AsseseeShareProperty") = "100"
            End If
            
            If UCase(CoOwnedYN_HP(m)) = "YES" Then
                Dim j As Long
                Dim rangecells1 As Range
                Dim rangecells2 As Range
                Dim rangecells3 As Range
                Dim rangecells4 As Range

                Set rangecells1 = Sheet22.Range("HP.Co.Name" & m).Cells
                Set rangecells2 = Sheet22.Range("HP.Co.PAN" & m).Cells
                Set rangecells3 = Sheet22.Range("HP.Co.Share" & m).Cells
                Set rangecells4 = Sheet22.Range("HP.Co.Aadhaar" & m).Cells
            
                setTblinfo_hpcoindex (m)
            
                ReDim CoName_HP(m, end_hpco)
                ReDim CoPAN_HP(m, end_hpco)
                ReDim CoAadhar_HP(m, end_hpco)
                ReDim CoShare_HP(m, end_hpco)

                For j = 1 To end_hpco
                    CoName_HP(m, j) = rangecells1.item(j).value
                    CoPAN_HP(m, j) = rangecells2.item(j).value
                    CoAadhar_HP(m, j) = rangecells4.item(j).value
                    CoShare_HP(m, j) = rangecells3.item(j).text
                    CoShare_HP(m, j) = Round(CoShare_HP(m, j), 2)
      
                    CoOwnersobj("CoOwnersSNo") = j
                    If CoName_HP(m, j) <> "" Then
                        CoOwnersobj("NameCoOwner") = CoName_HP(m, j)
                    End If
                    If CoPAN_HP(m, j) <> "" Then
                        CoOwnersobj("PAN_CoOwner") = CoPAN_HP(m, j)
                    End If
                    If CoAadhar_HP(m, j) <> "" Then
                        CoOwnersobj("Aadhaar_CoOwner") = UCase(CoAadhar_HP(m, j))
                    End If
                    CoOwnersobj("PercentShareProperty") = CoShare_HP(m, j)
                    CoOwners.add CoOwnersobj
                    Set CoOwnersobj = Nothing
                    Set CoOwnersobj = CreateObject("Scripting.Dictionary")
                Next
                  PropertyDetailsobj.add "CoOwners", CoOwners
                Set CoOwners = Nothing
                Set CoOwners = New Collection
            End If
          
            
            If ifLetOut_HP(m) <> "" Then
                If ifLetOut_HP(m) = "L" Then
                    PropertyDetailsobj("ifLetOut") = "L"
                End If
                If ifLetOut_HP(m) = "S" Then
                    PropertyDetailsobj("ifLetOut") = "S"
                End If
                If ifLetOut_HP(m) = "D" Then
                    PropertyDetailsobj("ifLetOut") = "D"
                End If
            Else
                PropertyDetailsobj("ifLetOut") = UCase(DefaultifLetOut_HP)
            End If
            
            If ifLetOut_HP(m) = "S" Then
            Else

                Set rangecells1 = Sheet22.Range("HP.NameofTenant" & m).Cells
                Set rangecells2 = Sheet22.Range("HP.PANofTenant" & m).Cells
                Set rangecells3 = Sheet22.Range("HP.TANofTenant" & m).Cells
                Set rangecells4 = Sheet22.Range("HP.AadharofTenant" & m).Cells
            
                setTblinfo_hprptfrm
            
                ReDim NameofTenant_HP(rangecells1.Cells.count)
                ReDim PANofTenant_HP(rangecells2.Cells.count)
                ReDim AadharofTenant_HP(rangecells4.Cells.count)
                ReDim TANofTenant_HP(rangecells3.Cells.count)
                If UBound(NameofTenant_HP) > 0 Then
                    Dim setobject
                    setobject = False
                    For j = 1 To UBound(NameofTenant_HP)
                        NameofTenant_HP(j) = rangecells1.item(j).value
                        PANofTenant_HP(j) = rangecells2.item(j).value
                        TANofTenant_HP(j) = rangecells3.item(j).value
                        AadharofTenant_HP(j) = rangecells4.item(j).value
                        If NameofTenant_HP(j) <> "" Then
                            setobject = True
                            TenantDetailsobj("TenantSNo") = j
                            TenantDetailsobj("NameofTenant") = UCase(NameofTenant_HP(j))
                            If PANofTenant_HP(j) <> "" Then
                            TenantDetailsobj("PANofTenant") = UCase(PANofTenant_HP(j))
                            End If
                            If AadharofTenant_HP(j) <> "" Then
                            TenantDetailsobj("AadhaarofTenant") = UCase(AadharofTenant_HP(j))
                            End If
                            If TANofTenant_HP(j) <> "" Then
                            TenantDetailsobj("PANTANofTenant") = TANofTenant_HP(j)
                            End If
                       
                            TenantDetails.add TenantDetailsobj
                            Set TenantDetailsobj = Nothing
                            Set TenantDetailsobj = CreateObject("Scripting.Dictionary")
                         End If
                    Next
                    If setobject Then
                        PropertyDetailsobj.add "TenantDetails", TenantDetails
                        Set TenantDetails = Nothing
                        Set TenantDetails = New Collection
                    End If
                End If
            End If

            If AnnualLetableValue_HP(m) <> "" Then
                Rentdetails("AnnualLetableValue") = UVCase(CDbl(AnnualLetableValue_HP(m)))
            Else
                Rentdetails("AnnualLetableValue") = UVCase(CDbl(DefaultAnnualLetableValue_HP))
            End If

            If RentNotRealized_HP(m) <> "" Then
                Rentdetails("RentNotRealized") = UVCase(CDbl(RentNotRealized_HP(m)))
            Else
                Rentdetails("RentNotRealized") = UVCase(CDbl(DefaultRentNotRealized_HP))
            End If

            If LocalTaxes_HP(m) <> "" Then
                Rentdetails("LocalTaxes") = UVCase(CDbl(LocalTaxes_HP(m)))
            Else
                Rentdetails("LocalTaxes") = UVCase(CDbl(DefaultLocalTaxes_HP))
            End If

            If TotalUnrealizedAndTax_HP(m) <> "" Then
                Rentdetails("TotalUnrealizedAndTax") = UVCase(CDbl(TotalUnrealizedAndTax_HP(m)))
            Else
                Rentdetails("TotalUnrealizedAndTax") = UVCase(CDbl(DefaultTotalUnrealizedAndTax_HP))
            End If

            If BalanceALV_HP(m) <> "" Then
                Rentdetails("BalanceALV") = UVCase(CDbl(BalanceALV_HP(m)))
            Else
                Rentdetails("BalanceALV") = UVCase(CDbl(DefaultBalanceALV_HP))
            End If

            If IncomeOfHPInOwnHand_HP(m) <> "" Then
                Rentdetails("AnnualOfPropOwned") = UVCase(CDbl(IncomeOfHPInOwnHand_HP(m)))
            Else
                Rentdetails("AnnualOfPropOwned") = 0
            End If

            If ThirtyPercentOfBalance_HP(m) <> "" Then
                Rentdetails("ThirtyPercentOfBalance") = UVCase(CDbl(ThirtyPercentOfBalance_HP(m)))
            Else
                Rentdetails("ThirtyPercentOfBalance") = UVCase(CDbl(DefaultThirtyPercentOfBalance_HP))
            End If

            If IntOnBorwCap_HP(m) <> "" Then
                 Rentdetails("IntOnBorwCap") = UVCase(CDbl(IntOnBorwCap_HP(m)))
            Else
                 Rentdetails("IntOnBorwCap") = UVCase(CDbl(DefaultIntOnBorwCap_HP))
            End If
'Konda--------------AY_2026_27_V0.2.5--------------------------------
           Dim ScheduleUs24BDtls_G, ScheduleUs24BDtls, Section24B, Section24BDtls

Set Section24B = CreateObject("Scripting.Dictionary")
Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
Set Section24BDtls = CreateObject("Scripting.Dictionary")

Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
Set ScheduleUs24BDtls = New Collection

subProcCaption = "24B"

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
Dim mIntCells24B, k
If Sheet22.Range("TotAmt.24b" & m).value > 0 Then
 mIntCells24B = Sheet22.Range("LoanfrmBankOrInstitute.24b" & m).count
 
  For k = 1 To mIntCells24B
   If (Sheet22.Range("LoanfrmBankOrInstitute.24b" & m).Cells(k, 1).value <> "" And Sheet22.Range("LoanfrmBankOrInstitute.24b" & m).Cells(k, 1).value <> "(Select)") Then
   
   LoanTknFrom_G = Sheet22.Range("LoanfrmBankOrInstitute.24b" & m).Cells(k, 1).value
   
   BankOrInstnName_G = Sheet22.Range("bankName.24b" & m).Cells(k, 1).value
   
   LoanAccNoOfBankOrInstnRefNo_G = Sheet22.Range("loanAccNum.24b" & m).Cells(k, 1).value
   DateofLoan_G = Sheet22.Range("loanDate.24b" & m).Cells(k, 1).value
   TotalLoanAmt_G = Sheet22.Range("loanAmt.24b" & m).Cells(k, 1).value
   LoanOutstndngAmt_G = Sheet22.Range("loanOutstanding.24b" & m).Cells(k, 1).value
   InterestPaid_G = Sheet22.Range("Intrst.24b" & m).Cells(k, 1).value
   
   
   If LoanTknFrom_G <> "" Then
            
             If LoanTknFrom_G = "Bank " Then
            LoanTknFrom_G = "B"
            
            ElseIf LoanTknFrom_G = "Other than Bank" Then
            LoanTknFrom_G = "I"
            End If
        ScheduleUs24BDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        ScheduleUs24BDtls_G("LoanTknFrom") = ""
   End If
    
   If BankOrInstnName_G <> "" Then
   ScheduleUs24BDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   ScheduleUs24BDtls_G("BankOrInstnName") = ""
   End If
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   ScheduleUs24BDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   ScheduleUs24BDtls_G("DateofLoan") = ""
   End If
   
   ScheduleUs24BDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   
   ScheduleUs24BDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   
   ScheduleUs24BDtls_G("InterestUs24B") = UVCase(InterestPaid_G)
 
                ScheduleUs24BDtls.add ScheduleUs24BDtls_G
                Set ScheduleUs24BDtls_G = Nothing
                Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
  
               
               Section24B.add "Section24BDtls", ScheduleUs24BDtls
                Set ScheduleUs24BDtls = Nothing
                Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
 

 TotalInterestPaid_G = Sheet22.Range("TotAmt.24b" & m).value
 Section24B("TotalInterestUs24B") = UVCase(TotalInterestPaid_G)

                Rentdetails.add "Section24B", Section24B
                Set Section24B = Nothing
                Set Section24B = CreateObject("Scripting.Dictionary")
End If
           
'---------------------------------------------------------------------
         
            If TotalDeduct_HP(m) <> "" Then
                Rentdetails("TotalDeduct") = UVCase(CDbl(TotalDeduct_HP(m)))
            Else
                Rentdetails("TotalDeduct") = UVCase(CDbl(DefaultTotalDeduct_HP))
            End If
            
            If Sheet22.Range("HP.RentOfEarlierYrSec_AandAA" & m).value <> "" Then
                Rentdetails("ArrearsUnrealizedRentRcvd") = UVCase(CDbl(Sheet22.Range("HP.RentOfEarlierYrSec_AandAA" & m).value))
            End If
            
            If IncomeOfHP_HP(m) <> "" Then
                Rentdetails("IncomeOfHP") = UVCase(CDbl(IncomeOfHP_HP(m)))
            Else
                Rentdetails("IncomeOfHP") = UVCase(CDbl(DefaultIncomeOfHP_HP))
            End If
            PropertyDetailsobj.add "Rentdetails", Rentdetails
            Set Rentdetails = Nothing
            Set Rentdetails = CreateObject("Scripting.Dictionary")
            
            PropertyDetails.add PropertyDetailsobj
            Set PropertyDetailsobj = Nothing
            Set PropertyDetailsobj = CreateObject("Scripting.Dictionary")
        End If
        Next
    jsonDictionary.add "PropertyDetails", PropertyDetails
            
'    If Sheet19.Range("HP.PassTroughIncome").Value <> "" Then
'        jsonDictionary("PassThroghIncome") = UVCase(CDbl(Sheet19.Range("HP.PassTroughIncome").Value))
'    End If

End If
End If

If TotalIncomeChargeableUnHP_HP <> "" Then
        jsonDictionary("TotalIncomeChargeableUnHP") = UVCase(CDbl(TotalIncomeChargeableUnHP_HP))
Else
        jsonDictionary("TotalIncomeChargeableUnHP") = UVCase(CDbl(DefaultTotalIncomeChargeableUnHP_HP))
End If

'=========================================Schema V0.2.2 on 20-01-2026




    If IncomeOS <> "" Then
        jsonDictionary("IncomeOthSrc") = IncomeOS
     Else
        jsonDictionary("IncomeOthSrc") = 0
    End If
         
    If Sheet1.Range("IncD.LessDeduction57").value <> "" Then
        jsonDictionary("DeductionUs57iia") = Sheet1.Range("IncD.LessDeduction57").value
    Else
        jsonDictionary("DeductionUs57iia") = 0
    End If

    If IncomeGross <> "" Then
        jsonDictionary("GrossTotIncome") = IncomeGross
     Else
        jsonDictionary("GrossTotIncome") = 0
    End If
    
    'added by Chetan V3.4 on 07/02/2025
    'start
    If GrossTotIncomeIncLTCG112A <> "" Then
        jsonDictionary("GrossTotIncomeIncLTCG112A") = GrossTotIncomeIncLTCG112A
     Else
        jsonDictionary("GrossTotIncomeIncLTCG112A") = 0
    End If
    'end
    
    If Investment = "Yes" Then
        Investment = "Y"
    ElseIf Investment = "No" Then
        Investment = "N"
    End If
    
    If Investment <> "" And Investment <> "(Select)" Then
        jsonDictionary("DepPayInvClmUndDednVIA") = Investment
    End If

    If TotalTaxIncome <> "" Then
        jsonDictionary("TotalIncome") = TotalTaxIncome
    Else
        jsonDictionary("TotalIncome") = 0
    End If

    UpdateProgressBar
    If (Not IsEmpty(Others_OSNOI)) Then
        If (UBound(Others_OSNOI) > 0) Then
            For i = 1 To UBound(Others_OSNOI)
                If (Others_OSNOI(i) = "Interest from Saving Bank Account") Then
                    Others_OSNOI(i) = "SAV"
                ElseIf (Others_OSNOI(i) = "Interest from Deposit (Bank/Post Office/Cooperative Society)") Then
                    Others_OSNOI(i) = "IFD"
                ElseIf (Others_OSNOI(i) = "Interest from Income Tax Refund") Then
                    Others_OSNOI(i) = "TAX"
                ElseIf (Others_OSNOI(i) = "Any Other Interest Income") Then
                    Others_OSNOI(i) = "OII"
                ElseIf (Others_OSNOI(i) = "Family pension") Then
                    Others_OSNOI(i) = "FAP"
                ElseIf (Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)") Then
                    Others_OSNOI(i) = "10(11)(iP)"
                ElseIf (Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)") Then
                    Others_OSNOI(i) = "10(11)(iiP)"
                ElseIf (Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)") Then
                    Others_OSNOI(i) = "10(12)(iP)"
                ElseIf (Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)") Then
                    Others_OSNOI(i) = "10(12)(iiP)"
                ElseIf (Others_OSNOI(i) = "Any Other") Then
                    Others_OSNOI(i) = "OTH"
                End If
            Next i
                    
            For i = 1 To UBound(Others_OSNOI)
                If Others_OSNOI(i) <> "" Then
                    OthersIncDtlsOth("OthSrcNatureDesc") = Others_OSNOI(i)
                End If
                            
                If Others_OSNOI1(i) <> "" And Others_OSNOI1(i) <> "Not Applicable" Then
                    OthersIncDtlsOth("OthSrcOthNatOfInc") = UCase(Others_OSNOI1(i))
                End If
                        
                If Others_OSAmt(i) <> "" Then
                    OthersIncDtlsOth("OthSrcOthAmount") = Others_OSAmt(i)
                End If
                OthersIncDtlsOthSrc.add OthersIncDtlsOth
                Set OthersIncDtlsOth = Nothing
                Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
            Next
            
             End If
        End If
'        If Sheet1.Range("OSIncomeNotifiedOther89A") <> "" Then
'
'                     Set OthersIncDtlsOth = Nothing
'                    Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
'                    OthersIncDtlsOth("OthSrcNatureDesc") = "OTHNOT89A"
'                    OthersIncDtlsOth("OthSrcOthAmount") = Sheet1.Range("OSIncomeNotifiedOther89A")
'                    OthersIncDtlsOthSrc.add OthersIncDtlsOth
'        End If
        If Up16Of3To31Of3 <> "" Or Up16Of12To15Of3 <> "" Or Up16Of9To15Of12 <> "" Or Upto15Of6 <> "" Or Upto15Of9 <> "" Then
                     Set OthersIncDtlsOth = Nothing
                    Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
                    
                    OthersIncDtlsOth("OthSrcNatureDesc") = "DIV"
                    
                    'Malli---------AY_2025_26
                    If TotDividendInc <> "" Then
                    OthersIncDtlsOth("OthSrcOthAmount") = TotDividendInc
                    Else
                    OthersIncDtlsOth("OthSrcOthAmount") = 0
                    End If
                    '--------------------------------
                    DateRange("Up16Of12To15Of3") = Up16Of12To15Of3
                    DateRange("Up16Of3To31Of3") = Up16Of3To31Of3
                    DateRange("Up16Of9To15Of12") = Up16Of9To15Of12
                    DateRange("Upto15Of6") = Upto15Of6
                    DateRange("Upto15Of9") = Upto15Of9
                    DividendInc.add "DateRange", DateRange
                    OthersIncDtlsOth.add "DividendInc", DividendInc
                    OthersIncDtlsOthSrc.add OthersIncDtlsOth
                End If
                
                
'            If Trim(Sheet1.Range("OSIncomeNotified89A_AmountUS").value) <> "" Or Trim(Sheet1.Range("OSIncomeNotified89A_AmountUK").value) <> "" Or Trim(Sheet1.Range("OSIncomeNotified89A_AmountCan").value) <> "" Then
'                    Set OthersIncDtlsOth = Nothing
'                    Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
'                    Set DateRange = Nothing
'                    Set DateRange = CreateObject("Scripting.Dictionary")
'                    OthersIncDtlsOth("OthSrcOthAmount") = Sheet1.Range("OSIncomeNotified89A")
''                    OthersIncDtlsOth("OthSrcNatureDesc") = "NOT89A"
'                    DateRange("Up16Of12To15Of3") = Sheet1.Range("IncD_q4OS1")
'                    DateRange("Up16Of3To31Of3") = Sheet1.Range("IncD_q5OS1")
'                    DateRange("Up16Of9To15Of12") = Sheet1.Range("IncD_q3OS1")
'                    DateRange("Upto15Of6") = Sheet1.Range("IncD_q1OS1")
'                    DateRange("Upto15Of9") = Sheet1.Range("IncD_q2OS1")
'
' 'Commented by  'Konda on 22-01-2026 schema-V0.2.5
''                    If Trim(Sheet1.Range("OSIncomeNotified89A_AmountUS").value) <> "" Then
''                        NOT89AType("NOT89ACountrycode") = "US"
''                        NOT89AType("NOT89AAmount") = Sheet1.Range("OSIncomeNotified89A_AmountUS").value
''                        NOT89A.add NOT89AType
''                        Set NOT89AType = Nothing
''                        Set NOT89AType = CreateObject("Scripting.Dictionary")
''                    End If
''                    If Trim(Sheet1.Range("OSIncomeNotified89A_AmountUK").value) <> "" Then
''                        NOT89AType("NOT89ACountrycode") = "UK"
''                        NOT89AType("NOT89AAmount") = Sheet1.Range("OSIncomeNotified89A_AmountUK").value
''                        NOT89A.add NOT89AType
''                        Set NOT89AType = Nothing
''                        Set NOT89AType = CreateObject("Scripting.Dictionary")
''                    End If
''                    If Trim(Sheet1.Range("OSIncomeNotified89A_AmountCan").value) <> "" Then
''                        NOT89AType("NOT89ACountrycode") = "CA"
''                        NOT89AType("NOT89AAmount") = Sheet1.Range("OSIncomeNotified89A_AmountCan").value
''                        NOT89A.add NOT89AType
''                        Set NOT89AType = Nothing
''                        Set NOT89AType = CreateObject("Scripting.Dictionary")
''                    End If
''                    OthersIncDtlsOth.add "NOT89A", NOT89A
''                    NOT89AInc.add "DateRange", DateRange
''                    OthersIncDtlsOth.add "NOT89AInc", NOT89AInc
''                    OthersIncDtlsOthSrc.add OthersIncDtlsOth
''                End If
            OthersInc.add "OthersIncDtlsOthSrc", OthersIncDtlsOthSrc
'            DateRange("Up16Of12To15Of3") = Up16Of3To31Of3
'            DateRange("Up16Of3To31Of3") = Up16Of12To15Of3
'            DateRange("Up16Of9To15Of12") = Up16Of9To15Of12
'            DateRange("Upto15Of6") = Upto15Of6
'            DateRange("Upto15Of9") = Upto15Of9
'            DividendInc.add "DateRange", DateRange
'            OthersInc.add "DividendInc", DividendInc
'            OthersInc("TotDividendInc") = TotDividendInc
            jsonDictionary.add "OthersInc", OthersInc
       
    
    UpdateProgressBar
    If Ded80C <> "" Then
        UsrDeductUndChapVIA("Section80C") = Ded80C
    Else
        UsrDeductUndChapVIA("Section80C") = 0
    End If
    
    UpdateProgressBar
    If Ded80CCC <> "" Then
        UsrDeductUndChapVIA("Section80CCC") = Ded80CCC
    Else
        UsrDeductUndChapVIA("Section80CCC") = 0
    End If
    
    UpdateProgressBar
    
    'New Schema updated by Konda as AY-2026-27 on 30-12-2025
    If Ded80CCC <> "" And Ded80CCC > 0 Then
    For i = 1 To UBound(Type_80CCC)
                'Malli updated as per V0.9  17/04/2026
'                    If Type_80CCC(i) <> "" Then
'                        PensionContributionFundtls("TypeofIdentifier") = Trim((Type_80CCC(i)))
'                    Else
'                        PensionContributionFundtls("TypeofIdentifier") = ""
'                    End If
                    
                    If Type_80CCC(i) <> "" Then
                            If UCase(Type_80CCC(i)) = UCase("PRAN") Then
                                PensionContributionFundtls("TypeofIdentifier") = CStr("PRAN")
                            ElseIf UCase(Type_80CCC(i)) = UCase("Other than PRAN") Then
                                PensionContributionFundtls("TypeofIdentifier") = CStr("OTHPRAN")
                            End If
                    Else
                         PensionContributionFundtls("TypeofIdentifier") = ""
                    End If
                    
                '-----------------
                
                If Name_80CCC(i) <> "" Then
                    PensionContributionFundtls("NameofIdentifier") = Name_80CCC(i)
                Else
                    PensionContributionFundtls("NameofIdentifier") = ""
                End If
                
                If Amount_80CCC(i) <> "" Then
                    PensionContributionFundtls("Amount") = Amount_80CCC(i)
                Else
                    PensionContributionFundtls("Amount") = ""
                End If
                
                PensionContribution80CCC.add PensionContributionFundtls
                Set PensionContributionFundtls = Nothing
                Set PensionContributionFundtls = CreateObject("Scripting.Dictionary")
        Next
        
        UsrDeductUndChapVIA.add "PensionContribution80CCC", PensionContribution80CCC
    
    End If
   UpdateProgressBar
    If Ded80CCDSelfEmployee <> "" Then
        UsrDeductUndChapVIA("Section80CCDEmployeeOrSE") = Ded80CCDSelfEmployee
    Else
        UsrDeductUndChapVIA("Section80CCDEmployeeOrSE") = 0
    End If
UpdateProgressBar


' 'Malli uncommented as per V0.9  17/04/2026
'New Schema updated by Konda as AY-2026-27 on 30-12-2025
' If Ded80CCDSelfEmployee <> "" And Ded80CCDSelfEmployee > 0 Then
'
'    For i = 1 To UBound(Type_1_80CCC)
'                If Type_1_80CCC(i) <> "" Then
'                    PensionContributionFundtls("TypeofIdentifier") = Trim(Type_1_80CCC(i))
'                Else
'                    PensionContributionFundtls("TypeofIdentifier") = ""
'                End If
'
'                If Name_1_80CCC(i) <> "" Then
'                    PensionContributionFundtls("NameofIdentifier") = Name_1_80CCC(i)
'                Else
'                    PensionContributionFundtls("NameofIdentifier") = ""
'                End If
'
'                If Amount_1_80CCC(i) <> "" Then
'                    PensionContributionFundtls("Amount") = Amount_1_80CCC(i)
'                Else
'                    PensionContributionFundtls("Amount") = ""
'                End If
'
'                PensionContribution80CCD1.add PensionContributionFundtls
'                Set PensionContributionFundtls = Nothing
'                Set PensionContributionFundtls = CreateObject("Scripting.Dictionary")
'        Next
'
'        UsrDeductUndChapVIA.add "PensionContribution80CCD1", PensionContribution80CCD1
'    End If
    
    UpdateProgressBar
    
    If Ded80CCD1B_SE <> "" Then
        UsrDeductUndChapVIA("Section80CCD1B") = Ded80CCD1B_SE
    Else
        UsrDeductUndChapVIA("Section80CCD1B") = 0
    End If
    
    
'      'Malli uncommented as per V0.9  17/04/2026
'    'New Schema updated by Konda as AY-2026-27 on 30-12-2025
'    If Ded80CCD1B_SE <> "" And Ded80CCD1B_SE > 0 Then
'
'     'Konda updated on 27-02-2026--SIT-111751
'    '    For i = 1 To UBound(Type_1_80CCC)
'        For i = 1 To UBound(Type_1b_80CCC)
'                    If Type_1b_80CCC(i) <> "" Then
'                        PensionContributionFundtls("TypeofIdentifier") = Trim(Type_1b_80CCC(i))
'                    Else
'                        PensionContributionFundtls("TypeofIdentifier") = ""
'                    End If
'
'                    If Name_1b_80CCC(i) <> "" Then
'                        PensionContributionFundtls("NameofIdentifier") = Name_1b_80CCC(i)
'                    Else
'                        PensionContributionFundtls("NameofIdentifier") = ""
'                    End If
'
'                    If Amount_1b_80CCC(i) <> "" Then
'                        PensionContributionFundtls("Amount") = Amount_1b_80CCC(i)
'                    Else
'                        PensionContributionFundtls("Amount") = ""
'                    End If
'
'                    PensionContribution80CCD1B.add PensionContributionFundtls
'                    Set PensionContributionFundtls = Nothing
'                    Set PensionContributionFundtls = CreateObject("Scripting.Dictionary")
'            Next
'
'            UsrDeductUndChapVIA.add "PensionContribution80CCD1B", PensionContribution80CCD1B
'        End If
    
    UpdateProgressBar
    If Ded80CCDEmployer <> "" Then
        UsrDeductUndChapVIA("Section80CCDEmployer") = Ded80CCDEmployer
    Else
        UsrDeductUndChapVIA("Section80CCDEmployer") = 0
    End If
    
    
  'Malli uncommented as per V0.9  17/04/2026
    'Malli----------AY_2025_26_21/04/2025
'New Schema updated by Konda as AY-2026-27 on 29-12-2025
'    Dim PRANNum_G As Variant
'
'    PRANNum_G = Sheet1.Range("Pran_Number").value
'
'    If Sheet1.Range("Pran_Number").Locked = False And Trim(PRANNum_G) <> "" Then
'        UsrDeductUndChapVIA("PRANNum") = UCase(Trim(PRANNum_G))
'    End If


If Sheet1.Range("IncD.Section80CCD_SE") > 0 Or Sheet1.Range("IncD.Section80CCD1B_SE") > 0 Then
 Dim mincellsPRANNum As Variant
 
 Dim PRANDtls, PRANDtls_PRANNum, PRANNum_Gen, prn
 
 Set PRANDtls_PRANNum = CreateObject("scripting.dictionary")
 Set PRANDtls = New Collection
 
 mincellsPRANNum = Sheet1.Range("pran_new").Rows.count
 
 
 For prn = 1 To mincellsPRANNum
 
     PRANNum_Gen = Sheet1.Range("pran_new").Cells(prn, 1).value
     If PRANNum_Gen <> "" Then
         PRANDtls_PRANNum("PRANNum") = CStr(PRANNum_Gen)
  
     
                    PRANDtls.add PRANDtls_PRANNum
                    Set PRANDtls_PRANNum = Nothing
                    Set PRANDtls_PRANNum = CreateObject("Scripting.Dictionary")
  End If
 Next
 Debug.Print PRANDtls.count
 If PRANDtls.count > 0 Then
     UsrDeductUndChapVIA.add "PRANDtls", PRANDtls
 End If
 
 End If

'    '--------------------------------------------------
    
    
    
    
    If Sheet1.Range("IncD.Section80DValue").value <> "" Then
        UsrDeductUndChapVIA("Section80D") = Sheet1.Range("IncD.Section80DValue").value
    Else
        UsrDeductUndChapVIA("Section80D") = 0
    End If

    UpdateProgressBar
    If SELECT80DD <> "" Then
        UsrDeductUndChapVIA("Section80DDUsrType") = SELECT80DD
    End If
    If Ded80DD <> "" Then
        UsrDeductUndChapVIA("Section80DD") = Ded80DD
    Else
        UsrDeductUndChapVIA("Section80DD") = 0
    End If
    
    
    
    If SELECT80DDB <> "" Then
        UsrDeductUndChapVIA("Section80DDBUsrType") = SELECT80DDB
    End If
    
    'Malli-----------AY_2025_26--21/04/2025
    Dim NameOfSpecDisease80DDB_G, NameOfSpecDisease80DDB_G_enum As Variant
    
    NameOfSpecDisease80DDB_G = Sheet1.Range("Sheet1.Specified_Disease").value
    
    If Sheet1.Range("Sheet1.Specified_Disease").Locked = False And NameOfSpecDisease80DDB_G <> "" Then
        
        If NameOfSpecDisease80DDB_G = "(a) Dementia" Then
          NameOfSpecDisease80DDB_G_enum = "a"
          
        ElseIf NameOfSpecDisease80DDB_G = "(b) Dystonia Musculorum Deformans" Then
          NameOfSpecDisease80DDB_G_enum = "b"
          
        ElseIf NameOfSpecDisease80DDB_G = "(c) Motor Neuron Disease" Then
          NameOfSpecDisease80DDB_G_enum = "c"
          
        ElseIf NameOfSpecDisease80DDB_G = "(d) Ataxia" Then
          NameOfSpecDisease80DDB_G_enum = "d"
          
        ElseIf NameOfSpecDisease80DDB_G = "(e) Chorea" Then
          NameOfSpecDisease80DDB_G_enum = "e"
          
        ElseIf NameOfSpecDisease80DDB_G = "(f) Hemiballismus" Then
          NameOfSpecDisease80DDB_G_enum = "f"
          
        ElseIf NameOfSpecDisease80DDB_G = "(g) Aphasia" Then
          NameOfSpecDisease80DDB_G_enum = "g"
          
        ElseIf NameOfSpecDisease80DDB_G = "(h) Parkinsons Disease" Then
          NameOfSpecDisease80DDB_G_enum = "h"
          
        ElseIf NameOfSpecDisease80DDB_G = "(i) Malignant Cancers" Then
          NameOfSpecDisease80DDB_G_enum = "i"
          
        ElseIf NameOfSpecDisease80DDB_G = "(j) Full Blown Acquired Immuno-Deficiency Syndrome (AIDS)" Then
          NameOfSpecDisease80DDB_G_enum = "j"
          
        ElseIf NameOfSpecDisease80DDB_G = "(k) Chronic Renal failure" Then
          NameOfSpecDisease80DDB_G_enum = "k"
          
        ElseIf NameOfSpecDisease80DDB_G = "(l) Hematological disorders" Then
          NameOfSpecDisease80DDB_G_enum = "l"
          
        ElseIf NameOfSpecDisease80DDB_G = "(m) Hemophilia" Then
          NameOfSpecDisease80DDB_G_enum = "m"
          
        ElseIf NameOfSpecDisease80DDB_G = "(n) Thalassaemia" Then
          NameOfSpecDisease80DDB_G_enum = "n"
          
          Else
          NameOfSpecDisease80DDB_G_enum = ""
          
       End If
    
    If NameOfSpecDisease80DDB_G_enum <> "" Then
    'UsrDeductUndChapVIA("NameOfSpecDisease80DDB") = UCase(Trim(NameOfSpecDisease80DDB_G_enum))
    UsrDeductUndChapVIA("NameOfSpecDisease80DDB") = Trim(NameOfSpecDisease80DDB_G_enum)
    End If
    
    End If
    '--------------------------------------
    
    
    If Ded80DDB <> "" Then
        UsrDeductUndChapVIA("Section80DDB") = Ded80DDB
    Else
        UsrDeductUndChapVIA("Section80DDB") = 0
    End If

    UpdateProgressBar
    If Ded80E <> "" Then
        UsrDeductUndChapVIA("Section80E") = Ded80E
    Else
        UsrDeductUndChapVIA("Section80E") = 0
    End If
    
    UpdateProgressBar
    If Ded80EE <> "" Then
        UsrDeductUndChapVIA("Section80EE") = Ded80EE
    Else
        UsrDeductUndChapVIA("Section80EE") = 0
    End If
    If Ded80EEA <> "" Then
        UsrDeductUndChapVIA("Section80EEA") = Ded80EEA
    Else
        UsrDeductUndChapVIA("Section80EEA") = 0
    End If
    If Ded80EEB <> "" Then
        UsrDeductUndChapVIA("Section80EEB") = Ded80EEB
    Else
        UsrDeductUndChapVIA("Section80EEB") = 0
    End If
    If Ded80G <> "" Then
        UsrDeductUndChapVIA("Section80G") = Ded80G
    Else
        UsrDeductUndChapVIA("Section80G") = 0
    End If
    If Ded80GG <> "" Then
        UsrDeductUndChapVIA("Section80GG") = Ded80GG
    Else
        UsrDeductUndChapVIA("Section80GG") = 0
    End If
    
    'Malli----------AY_2025_26_21/04/2025
    Dim Form10BAAckNum_G As Variant
              
    Form10BAAckNum_G = Sheet1.Range("Sheet1.AckNum").value
    If Sheet1.Range("Sheet1.AckNum").Locked = False And Trim(Form10BAAckNum_G) <> "" Then
        UsrDeductUndChapVIA("Form10BAAckNum") = UCase(Trim(Form10BAAckNum_G))
    End If
        '------------------------------------
        

    
    
    
'    If Ded80GGA <> "" Then
'        UsrDeductUndChapVIA("Section80GGA") = Ded80GGA
'    Else
'        UsrDeductUndChapVIA("Section80GGA") = 0
'    End If
                     
    UpdateProgressBar
    If Ded80GGC <> "" Then
        UsrDeductUndChapVIA("Section80GGC") = Ded80GGC
    Else
        UsrDeductUndChapVIA("Section80GGC") = 0
    End If
    If SELECT80U <> "" Then
        UsrDeductUndChapVIA("Section80UUsrType") = SELECT80U
    End If
    If Ded80U <> "" Then
        UsrDeductUndChapVIA("Section80U") = Ded80U
    Else
        UsrDeductUndChapVIA("Section80U") = 0
    End If
    If Ded80TTA = "" Then
        UsrDeductUndChapVIA("Section80TTA") = 0
    Else
        UsrDeductUndChapVIA("Section80TTA") = Ded80TTA
    End If
    If Ded80TTB = "" Then
        UsrDeductUndChapVIA("Section80TTB") = 0
    Else
        UsrDeductUndChapVIA("Section80TTB") = Ded80TTB
    End If
    
'PAG_C61 AY 2023-24 Old-----------------
    
'    'Generate - 06.04.2023.103.ID.01
'    If Trim(Sheet1.Range("IncD.AnyOtherDeductions").value) <> "" Then
'            UsrDeductUndChapVIA("AnyOthSec80CCH") = Sheet1.Range("IncD.AnyOtherDeductions").value
'    End If
'    'End Change

'PAG_C61 AY 2024-25 Change Starts    -----------------
    
    If Ded80CCH = "" Then
        UsrDeductUndChapVIA("AnyOthSec80CCH") = 0
    Else
        UsrDeductUndChapVIA("AnyOthSec80CCH") = Ded80CCH
    End If
    
'PAG_C61 AY 2024-25 Change End-----------------
    
   
    If TotalDed = "" Then
        UsrDeductUndChapVIA("TotalChapVIADeductions") = 0
    Else
        UsrDeductUndChapVIA("TotalChapVIADeductions") = TotalDed
    End If
    jsonDictionary.add "UsrDeductUndChapVIA", UsrDeductUndChapVIA
    
    If Ded80C_calc <> "" Then
        DeductUndChapVIA("Section80C") = Ded80C_calc
    Else
        DeductUndChapVIA("Section80C") = 0
    End If
    UpdateProgressBar
    If Ded80CCC_calc <> "" Then
        DeductUndChapVIA("Section80CCC") = Ded80CCC_calc
    Else
        DeductUndChapVIA("Section80CCC") = 0
    End If
                                      
    UpdateProgressBar
    If Ded80CCDSelfEmployee_calc <> "" Then
        DeductUndChapVIA("Section80CCDEmployeeOrSE") = Ded80CCDSelfEmployee_calc
    Else
        DeductUndChapVIA("Section80CCDEmployeeOrSE") = 0
    End If
         
    UpdateProgressBar
    If Ded80CCD1B_SE_calc <> "" Then
        DeductUndChapVIA("Section80CCD1B") = Ded80CCD1B_SE_calc
    Else
        DeductUndChapVIA("Section80CCD1B") = 0
    End If
    If Ded80CCDEmployer_calc <> "" Then
        DeductUndChapVIA("Section80CCDEmployer") = Ded80CCDEmployer_calc
    Else
        DeductUndChapVIA("Section80CCDEmployer") = 0
    End If
    If Ded80D_calc <> "" Then
        DeductUndChapVIA("Section80D") = Ded80D_calc
    Else
        DeductUndChapVIA("Section80D") = 0
    End If
    
    UpdateProgressBar
    If Ded80DD_calc <> "" Then
        DeductUndChapVIA("Section80DD") = Ded80DD_calc
    Else
        DeductUndChapVIA("Section80DD") = 0
    End If
    If Ded80DDB_calc <> "" Then
        DeductUndChapVIA("Section80DDB") = Ded80DDB_calc
    Else
        DeductUndChapVIA("Section80DDB") = 0
    End If
    If Ded80E_calc <> "" Then
        DeductUndChapVIA("Section80E") = Ded80E_calc
    Else
        DeductUndChapVIA("Section80E") = 0
    End If
                     
    UpdateProgressBar
    If Ded80EE_calc <> "" Then
        DeductUndChapVIA("Section80EE") = Ded80EE_calc
    Else
        DeductUndChapVIA("Section80EE") = 0
    End If
    If Ded80EEA_calc <> "" Then
        DeductUndChapVIA("Section80EEA") = Ded80EEA_calc
    Else
        DeductUndChapVIA("Section80EEA") = 0
    End If
    If Ded80EEB_calc <> "" Then
        DeductUndChapVIA("Section80EEB") = Ded80EEB_calc
    Else
        DeductUndChapVIA("Section80EEB") = 0
    End If
    If Ded80G_calc <> "" Then
        DeductUndChapVIA("Section80G") = Ded80G_calc
    Else
        DeductUndChapVIA("Section80G") = 0
    End If
    If Ded80GG_calc <> "" Then
        DeductUndChapVIA("Section80GG") = Ded80GG_calc
    Else
        DeductUndChapVIA("Section80GG") = 0
    End If
'    If Ded80GGA_calc <> "" Then
'        DeductUndChapVIA("Section80GGA") = Ded80GGA_calc
'    Else
'        DeductUndChapVIA("Section80GGA") = 0
'    End If
    If Ded80GGC_calc <> "" Then
        DeductUndChapVIA("Section80GGC") = Ded80GGC_calc
    Else
        DeductUndChapVIA("Section80GGC") = 0
    End If
    If Ded80U_calc <> "" Then
        DeductUndChapVIA("Section80U") = Ded80U_calc
    Else
        DeductUndChapVIA("Section80U") = 0
    End If
    If Ded80TTA_calc <> "" Then
        DeductUndChapVIA("Section80TTA") = Ded80TTA_calc
    Else
        DeductUndChapVIA("Section80TTA") = 0
    End If
    If Ded80TTB_calc <> "" Then
        DeductUndChapVIA("Section80TTB") = Ded80TTB_calc
    Else
        DeductUndChapVIA("Section80TTB") = 0
    End If
    
    
    
 'PAG_E61 AY 2023-24 Old
    
'    'Generate - 06.04.2023.103.ID.01
'    If Sheet1.Range("IncD.AnyOtherDeductions_Calc").value <> "" Then
'        DeductUndChapVIA("AnyOthSec80CCH") = Sheet1.Range("IncD.AnyOtherDeductions_Calc").value
'    End If
'    'End Change
 'PAG_E61 AY 2023-24 Old
 
    
'PAG_C61 AY 2024-25 Change Starts--------------

    If Ded80CCH_calc <> "" Then
        DeductUndChapVIA("AnyOthSec80CCH") = Ded80CCH_calc
    Else
        DeductUndChapVIA("AnyOthSec80CCH") = 0
    End If
    
'PAG_C61 AY 2024-25 Change Ends--------------
    
                     
    
    If TotalDed_calc <> "" Then
        DeductUndChapVIA("TotalChapVIADeductions") = TotalDed_calc
    Else
        DeductUndChapVIA("TotalChapVIADeductions") = 0
    End If
    jsonDictionary.add "DeductUndChapVIA", DeductUndChapVIA
UpdateProgressBar
Set ITR4_IncomeDeductions = jsonDictionary
End Function
Function TaxComputation() As Object
subProcCaption = "Tax Computation"
noOfProcessSub = 10

Dim i As Long
Dim FeeFurnish234I
Dim jsonDictionary, IntrstPay

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TaxComputation = CreateObject("Scripting.Dictionary")
Set IntrstPay = CreateObject("Scripting.Dictionary")

 FeeFurnish234I = Sheet1.Range("IncD.IntrstPayUs234I").value
 
 
If TaxPayableTotalIncome <> "" Then
    jsonDictionary("TotalTaxPayable") = TaxPayableTotalIncome
Else
    jsonDictionary("TotalTaxPayable") = 0
End If

If Rebate87A <> "" Then
    jsonDictionary("Rebate87A") = UVCase(CDbl(Rebate87A))
Else
    jsonDictionary("Rebate87A") = 0
End If

UpdateProgressBar

If TaxafterRebate <> "" Then
    jsonDictionary("TaxPayableOnRebate") = UVCase(CDbl(TaxafterRebate))
Else
    jsonDictionary("TaxPayableOnRebate") = 0
End If

If EducationCess <> "" Then
    jsonDictionary("EducationCess") = UVCase(CDbl(EducationCess))
Else
    jsonDictionary("EducationCess") = 0
 
End If

UpdateProgressBar
            
If TotalTaxSurchargeEdu <> "" Then
    jsonDictionary("GrossTaxLiability") = UVCase(CDbl(TotalTaxSurchargeEdu))
Else
    jsonDictionary("GrossTaxLiability") = 0
End If

If Relief89 <> "" Then
    jsonDictionary("Section89") = UVCase(CDbl(Relief89))
Else
    jsonDictionary("Section89") = 0
End If

If BalanceTax <> "" Then
    jsonDictionary("NetTaxLiability") = UVCase(CDbl(BalanceTax))
Else
    jsonDictionary("NetTaxLiability") = 0
End If

UpdateProgressBar

IntrstPay("IntrstPayUs234A") = UVCase(CDbl(InterestPay234a))
IntrstPay("IntrstPayUs234B") = UVCase(CDbl(InterestPay234b))
IntrstPay("IntrstPayUs234C") = UVCase(CDbl(InterestPay234c))
IntrstPay("LateFilingFee234F") = UVCase(CDbl(InterestPay234f))
'Konda updated on 27-02-2026--V0.4
IntrstPay("FeeFurnish234I") = UVCase(CDbl(FeeFurnish234I))
'=============

jsonDictionary.add "IntrstPay", IntrstPay


If TotalTaxInterest <> "" Then
jsonDictionary("TotTaxPlusIntrstPay") = UVCase(CDbl(TotalTaxInterest))
Else
jsonDictionary("TotTaxPlusIntrstPay") = 0
End If

UpdateProgressBar
Set TaxComputation = jsonDictionary
       
End Function
Function TaxPaid() As Object

Dim jsonDictionary, TaxesPaid

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TaxesPaid = CreateObject("Scripting.Dictionary")

If AdvanceTaxAmnt <> "" Then
    TaxesPaid("AdvanceTax") = UVCase(CDbl(AdvanceTaxAmnt))
Else
    TaxesPaid("AdvanceTax") = 0
End If

If TDSAmnt <> "" Then
    TaxesPaid("TDS") = UVCase(CDbl(TDSAmnt))
Else
    TaxesPaid("TDS") = 0
End If

UpdateProgressBar

If TCSAmnt <> "" Then
    TaxesPaid("TCS") = UVCase(CDbl(TCSAmnt))
Else
    TaxesPaid("TCS") = 0
End If

UpdateProgressBar

If SelfAssmntTax <> "" Then
    TaxesPaid("SelfAssessmentTax") = UVCase(CDbl(SelfAssmntTax))
Else
    TaxesPaid("SelfAssessmentTax") = 0
End If

If TotalTaxPaid <> "" Then
    TaxesPaid("TotalTaxesPaid") = UVCase(CDbl(TotalTaxPaid))
Else
    TaxesPaid("TotalTaxesPaid") = 0
End If

UpdateProgressBar

jsonDictionary.add "TaxesPaid", TaxesPaid

If TotalTaxPayable <> "" Then
    jsonDictionary("BalTaxPayable") = UVCase(CDbl(TotalTaxPayable))
Else
    jsonDictionary("BalTaxPayable") = 0
End If

UpdateProgressBar

Set TaxPaid = jsonDictionary

End Function
Function Refund() As Object
subProcCaption = ""
noOfProcessSub = 10
Dim i As Long
Dim jsonDictionary, BankAccountDtls, AddtnlBank, AddtnlBankDetails

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set BankAccountDtls = CreateObject("Scripting.Dictionary")
Set AddtnlBank = CreateObject("Scripting.Dictionary")

Set AddtnlBankDetails = New Collection
      jsonDictionary("RefundDue") = Sheet5.Range("IncD.RefundDue").value
     
        noOfProcessSub = UBound(BankIFSC_BA)
        If (Not IsEmpty(BankIFSC_BA)) And UBound(BankIFSC_BA) > 0 Then
            For i = 1 To UBound(BankIFSC_BA)
                
                If BankIFSC_BA(i) <> "" Then
                    AddtnlBank("IFSCCode") = UCase(BankIFSC_BA(i))
                Else
                    AddtnlBank("IFSCCode") = ""
                End If
                If BankName_BA(i) <> "" Then
                    AddtnlBank("BankName") = BankName_BA(i)
                Else
                    AddtnlBank("BankName") = ""
                End If
                
                If BankAccntnum_BA(i) <> "" Then
'Change-05.03.2023.101 VF issue
'                    AddtnlBank("BankAccountNo") = UCase(BankAccntnum_BA(i)) & ""
                    AddtnlBank("BankAccountNo") = UCase(BankAccntnum_BA(i))
                Else
                     AddtnlBank("BankAccountNo") = ""
                End If
                
                
'AY 2024-25 Schema implementation Starts----------------------------------

    If Sheet5.Range("SchBA.Type_of_Account").item(i).value <> "" Then
        If Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Savings Account" Then
            AddtnlBank("AccountType") = "SB"
               ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Current Account" Then
               AddtnlBank("AccountType") = "CA"
                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Cash Credit Account" Then
               AddtnlBank("AccountType") = "CC"
                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Over draft account" Then
               AddtnlBank("AccountType") = "OD"
               
                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Non Resident Account" Then
               AddtnlBank("AccountType") = "NRO"
                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Other" Then
               AddtnlBank("AccountType") = "OTH"
               
'               ElseIf UCase(Mid(BankAccnType_BA, 1, 3)) = UCase(Mid("Non Resident Account", 1, 3)) Then
'                AddtnlBank("AccountType") = "NRO"
'             'ElseIf UCase(BankAccnType_BA) = UCase("Other") Then
'               ElseIf UCase(Mid(BankAccnType_BA, 1, 3)) = UCase(Mid("Other", 1, 3)) Then
'                AddtnlBank("AccountType") = "OTH"
           
               
               
               
        End If
    End If
   
'AY 2024-25 Schema implementation End-----------------------------------
                
'uncommented by Chetan C M AY2025-26 21/01/2025
'--start--
                
                If tempXML(i) = True Then
                      tempXML(i) = "true"
                ElseIf tempXML(i) = "" Then
                      tempXML(i) = "false"
                ElseIf tempXML(i) = False Then
                      tempXML(i) = "false"
                End If
    
                If tempXML(i) <> "" Then
                AddtnlBank("UseForRefund") = tempXML(i)
                End If
'--end--
                AddtnlBankDetails.add AddtnlBank
                Set AddtnlBank = Nothing
                Set AddtnlBank = CreateObject("Scripting.Dictionary")
                
                UpdateProgressBar
            Next
        End If
        BankAccountDtls.add "AddtnlBankDetails", AddtnlBankDetails
        jsonDictionary.add "BankAccountDtls", BankAccountDtls
UpdateProgressBar
Set Refund = jsonDictionary
End Function
 
Function Verification()
subProcCaption = "Verification"
noOfProcessSub = 4
Dim jsonDictionary, Declaration

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Declaration = CreateObject("Scripting.Dictionary")

                 If AssesseName <> "" Then
                    Declaration("AssesseeVerName") = UCase(AssesseName)
                 Else
                    Declaration("AssesseeVerName") = ""
                 End If
                 
                 If AssesseFatherName <> "" Then
                    Declaration("FatherName") = UCase(AssesseFatherName)
                 End If
                 
        UpdateProgressBar
        
                 If VerificationPAN <> "" Then
                    Declaration("AssesseeVerPAN") = UCase(VerificationPAN)
                 End If
                 
        UpdateProgressBar
        jsonDictionary.add "Declaration", Declaration
            
            If Capacity <> "" Then
            jsonDictionary("Capacity") = Mid(UCase(Capacity), 1, 1)
            Else
            jsonDictionary("Capacity") = ""
             End If
             
             If VerificationPlace <> "" Then
             jsonDictionary("Place") = UCase(VerificationPlace)
             Else
            jsonDictionary("Place") = ""
             End If

         
      UpdateProgressBar
        Set Verification = jsonDictionary
End Function

'XML Generation Function for 80G
        
Function Schedule80G() As Object
Dim i As Long
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim TotalDonationsUs80GCash, TotalDonationsUs80GOtherMode, TotalDonationsUs80G, TotalEligibleDonationsUs80G As Variant

Dim DoneeWithPan100 As Object
Dim DoneeWithPan50 As Object
Dim DoneeWithPan100Appr As Object
Dim DoneeWithPan50Appr As Object

Dim DoneeWith100, DoneeWith50, DoneeWith100Appr, DoneeWith50Appr

Set DoneeWith100 = CreateObject("Scripting.Dictionary")
Set DoneeWith50 = CreateObject("Scripting.Dictionary")
Set DoneeWith100Appr = CreateObject("Scripting.Dictionary")
Set DoneeWith50Appr = CreateObject("Scripting.Dictionary")

Dim AddressDetail100, AddressDetail50, AddressDetail00Appr, AddressDetai50Appr

Set AddressDetail100 = CreateObject("Scripting.Dictionary")
Set AddressDetail50 = CreateObject("Scripting.Dictionary")
Set AddressDetail00Appr = CreateObject("Scripting.Dictionary")
Set AddressDetai50Appr = CreateObject("Scripting.Dictionary")

Dim Don100Percent, Don50PercentNoApprReqd, Don100PercentApprReqd, Don50PercentApprReqd
Dim TotDon100PercentCash, TotDon100PercentOtherMode, TotDon100Percent, TotEligibleDon100Percent As Variant
Dim TotDon50PercentNoApprReqdCash, TotDon50PercentNoApprReqdOtherMode, TotDon50PercentNoApprReqd, TotEligibleDon50Percent As Variant
Dim TotDon100PercentApprReqdCash, TotDon100PercentApprReqdOtherMode, TotDon100PercentApprReqd, TotEligibleDon100PercentApprReqd As Variant
Dim TotDon50PercentApprReqdCash, TotDon50PercentApprReqdOtherMode, TotDon50PercentApprReqd, TotEligibleDon50PercentApprReqd As Variant

Set Don100Percent = CreateObject("Scripting.Dictionary")
Set Don50PercentNoApprReqd = CreateObject("Scripting.Dictionary")
Set Don100PercentApprReqd = CreateObject("Scripting.Dictionary")
Set Don50PercentApprReqd = CreateObject("Scripting.Dictionary")

Set DoneeWithPan100 = New Collection
Set DoneeWithPan50 = New Collection
Set DoneeWithPan100Appr = New Collection
Set DoneeWithPan50Appr = New Collection
 
If Not IsEmpty(GA80DoneeName) Then
If (UBound(GA80DoneeName) > 0) Then
    subProcCaption = "Schedule 80G"
    noOfProcessSub = UBound(GA80DoneeName)
    For i = 0 To UBound(GA80DoneeName) - 1
        If GA80DoneeName(i) <> "" Then
            DoneeWith100("DoneeWithPanName") = UCase(GA80DoneeName(i))
        End If
        If GA80DoneePAN(i) <> "" Then
            DoneeWith100("DoneePAN") = UCase(GA80DoneePAN(i))
        End If
        If GA80DoneeAdrDet(i) <> "" Then
            AddressDetail100("AddrDetail") = UCase(GA80DoneeAdrDet(i))
        End If
        If GA80DoneeCityTownn(i) <> "" Then
            AddressDetail100("CityOrTownOrDistrict") = UCase(GA80DoneeCityTownn(i))
        End If
        If GA80DoneeState(i) <> "" And GA80DoneeState(i) <> "(Select)" Then
            AddressDetail100("StateCode") = UCase(Mid(Trim(GA80DoneeState(i)), 1, InStr(1, Trim(GA80DoneeState(i)), "-") - 1))
        End If
        If GA80DoneePinCode(i) <> "" Then
            AddressDetail100("PinCode") = UVCase(CDbl(GA80DoneePinCode(i)))
        End If
        
        DoneeWith100.add "AddressDetail", AddressDetail100
        Set AddressDetail100 = Nothing
        Set AddressDetail100 = CreateObject("Scripting.Dictionary")

        If GA80DoneeInCash(i) <> "" Then
            DoneeWith100("DonationAmtCash") = UVCase(CDbl(GA80DoneeInCash(i)))
        Else
            DoneeWith100("DonationAmtCash") = 0
        End If
        
        If GA80DoneeAmnt(i) <> "" Then
            DoneeWith100("DonationAmtOtherMode") = UVCase(CDbl(GA80DoneeAmnt(i)))
        Else
            DoneeWith100("DonationAmtOtherMode") = 0
        End If
'Konda Updated 09-03-2026--V0.5
        If GA80DoneeTraref(i) <> "" Then
            DoneeWith100("TransactionRefNum") = UCase(GA80DoneeTraref(i))
        End If
        
        
        
        If GA80DoneeIFSC(i) <> "" Then
            DoneeWith100("IFSCCode") = UVCase(GA80DoneeIFSC(i))
        End If
'==========================
        If GA80DoneeTotalDonation(i) <> "" Then
            DoneeWith100("DonationAmt") = UVCase(CDbl(GA80DoneeTotalDonation(i)))
        Else
            DoneeWith100("DonationAmt") = 0
        End If
        
        If GA80DoneeEligibleAmnt(i) <> "" Then
            DoneeWith100("EligibleDonationAmt") = UVCase(CDbl(GA80DoneeEligibleAmnt(i)))
        Else
            DoneeWith100("EligibleDonationAmt") = 0
        End If
        
            DoneeWithPan100.add DoneeWith100
            Set DoneeWith100 = Nothing
            Set DoneeWith100 = CreateObject("Scripting.Dictionary")
        UpdateProgressBar
        Next
        Don100Percent.add "DoneeWithPan", DoneeWithPan100

        If GA80DoneeTot_InCash <> "" Then
            Don100Percent("TotDon100PercentCash") = UVCase(CDbl(GA80DoneeTot_InCash))
        Else
            Don100Percent("TotDon100PercentCash") = 0
        End If
        
        If GA80DoneeTotAmnt <> "" Then
            Don100Percent("TotDon100PercentOtherMode") = UVCase(CDbl(GA80DoneeTotAmnt))
        Else
            Don100Percent("TotDon100PercentOtherMode") = 0
        End If
        
        If GA80DoneeTot_TotalDonation <> "" Then
            Don100Percent("TotDon100Percent") = UVCase(CDbl(GA80DoneeTot_TotalDonation))
        Else
            Don100Percent("TotDon100Percent") = 0
        End If
        
        If GA80DoneeTotEligAmnt <> "" Then
            Don100Percent("TotEligibleDon100Percent") = UVCase(CDbl(GA80DoneeTotEligAmnt))
        Else
            Don100Percent("TotEligibleDon100Percent") = 0
        End If
        jsonDictionary.add "Don100Percent", Don100Percent
        End If
        End If
             
        If Not IsEmpty(GB80DoneePAN) Then
        If (UBound(GB80DoneePAN) > 0) Then
            subProcCaption = "80GB"
            noOfProcessSub = UBound(GB80DoneeName)
        
            For i = 0 To UBound(GB80DoneeName) - 1
                If GB80DoneeName(i) <> "" Then
                    DoneeWith50("DoneeWithPanName") = UCase(GB80DoneeName(i))
                End If
                If GB80DoneePAN(i) <> "" Then
                    DoneeWith50("DoneePAN") = UCase(GB80DoneePAN(i))
                End If
                If GB80DoneeAdrDet(i) <> "" Then
                    AddressDetail50("AddrDetail") = UCase(GB80DoneeAdrDet(i))
                End If
                If GB80DoneeCityTownn(i) <> "" Then
                    AddressDetail50("CityOrTownOrDistrict") = UCase(GB80DoneeCityTownn(i))
                End If
                If GB80DoneeState(i) <> "" And GB80DoneeState(i) <> "(Select)" Then
                    AddressDetail50("StateCode") = UCase(Mid(Trim(GB80DoneeState(i)), 1, InStr(1, Trim(GB80DoneeState(i)), "-") - 1))
                End If
                If GB80DoneePinCode(i) <> "" Then
                    AddressDetail50("PinCode") = UVCase(CDbl(GB80DoneePinCode(i)))
                End If
                DoneeWith50.add "AddressDetail", AddressDetail50
                Set AddressDetail50 = Nothing
                Set AddressDetail50 = CreateObject("Scripting.Dictionary")
    
                If GB80DoneeInCash(i) <> "" Then
                    DoneeWith50("DonationAmtCash") = UVCase(CDbl(GB80DoneeInCash(i)))
                Else
                    DoneeWith50("DonationAmtCash") = 0
                End If
                
                If GB80DoneeAmnt(i) <> "" Then
                    DoneeWith50("DonationAmtOtherMode") = UVCase(CDbl(GB80DoneeAmnt(i)))
                Else
                    DoneeWith50("DonationAmtOtherMode") = 0
                End If
   'Konda Updated 09-03-2026--V0.5
                If GB80DoneeTraref(i) <> "" Then
                    DoneeWith50("TransactionRefNum") = UCase(GB80DoneeTraref(i))
                End If
                
                If GB80DoneeIFSC(i) <> "" Then
                    DoneeWith50("IFSCCode") = UVCase(GB80DoneeIFSC(i))
                End If
'==========================
                
                If GB80DoneeTotalDonation(i) <> "" Then
                    DoneeWith50("DonationAmt") = UVCase(CDbl(GB80DoneeTotalDonation(i)))
                Else
                    DoneeWith50("DonationAmt") = 0
                End If
        
                If GB80DoneeEligibleAmnt(i) <> "" Then
                    DoneeWith50("EligibleDonationAmt") = UVCase(CDbl(GB80DoneeEligibleAmnt(i)))
                Else
                    DoneeWith50("EligibleDonationAmt") = 0
                End If
                DoneeWithPan50.add DoneeWith50
                Set DoneeWith50 = Nothing
                Set DoneeWith50 = CreateObject("Scripting.Dictionary")
    
                UpdateProgressBar
            Next
            Don50PercentNoApprReqd.add "DoneeWithPan", DoneeWithPan50
            
            If GB80DoneeTot_InCash <> "" Then
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdCash") = UVCase(CDbl(GB80DoneeTot_InCash))
            Else
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdCash") = 0
            End If
                             
            If GB80DoneeTotAmnt <> "" Then
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdOtherMode") = UVCase(CDbl(GB80DoneeTotAmnt))
            Else
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdOtherMode") = 0
            End If
                                  
            If GB80DoneeTot_TotalDonation <> "" Then
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqd") = UVCase(CDbl(GB80DoneeTot_TotalDonation))
            Else
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqd") = 0
            End If
                          
                     
            If GB80DoneeTotEligAmnt <> "" Then
                Don50PercentNoApprReqd("TotEligibleDon50Percent") = UVCase(CDbl(GB80DoneeTotEligAmnt))
            Else
                Don50PercentNoApprReqd("TotEligibleDon50Percent") = 0
            End If
            jsonDictionary.add "Don50PercentNoApprReqd", Don50PercentNoApprReqd
            End If
            End If
            


            If Not IsEmpty(GC80DoneeName) Then
            If (UBound(GC80DoneeName) > 0) Then
                subProcCaption = "80GC"
                noOfProcessSub = UBound(GC80DoneeName)
            
                For i = 0 To UBound(GC80DoneeName) - 1
                    If GC80DoneeName(i) <> "" Then
                        DoneeWith100Appr("DoneeWithPanName") = UCase(GC80DoneeName(i))
                    End If
                    If GC80DoneePAN(i) <> "" Then
                        DoneeWith100Appr("DoneePAN") = UCase(GC80DoneePAN(i))
                    End If
                   
                    If GC80DoneeAdrDet(i) <> "" Then
                        AddressDetail00Appr("AddrDetail") = UCase(GC80DoneeAdrDet(i))
                    End If
                    If GC80DoneeCityTownn(i) <> "" Then
                        AddressDetail00Appr("CityOrTownOrDistrict") = UCase(GC80DoneeCityTownn(i))
                    End If
                    If GC80DoneeState(i) <> "" And GC80DoneeState(i) <> "(Select)" Then
                        AddressDetail00Appr("StateCode") = UCase(Mid(Trim(GC80DoneeState(i)), 1, InStr(1, Trim(GC80DoneeState(i)), "-") - 1))
                    End If
                    If GC80DoneePinCode(i) <> "" Then
                        AddressDetail00Appr("PinCode") = UVCase(CDbl(GC80DoneePinCode(i)))
                    End If
                    DoneeWith100Appr.add "AddressDetail", AddressDetail00Appr
                    Set AddressDetail00Appr = Nothing
                    Set AddressDetail00Appr = CreateObject("Scripting.Dictionary")
    
                    If GC80DoneeInCash(i) <> "" Then
                        DoneeWith100Appr("DonationAmtCash") = UVCase(CDbl(GC80DoneeInCash(i)))
                    Else
                        DoneeWith100Appr("DonationAmtCash") = 0
                    End If
                    
                    If GC80DoneeAmnt(i) <> "" Then
                        DoneeWith100Appr("DonationAmtOtherMode") = UVCase(CDbl(GC80DoneeAmnt(i)))
                    Else
                        DoneeWith100Appr("DonationAmtOtherMode") = 0
                    End If
                    
        'Konda Updated 09-03-2026--V0.5
                If GC80DoneeTraref(i) <> "" Then
                    DoneeWith100Appr("TransactionRefNum") = UCase(GC80DoneeTraref(i))
                End If
                
                If GC80DoneeIFSC(i) <> "" Then
                    DoneeWith100Appr("IFSCCode") = UVCase(GC80DoneeIFSC(i))
                End If
'==========================
                    
                    
                    If GC80DoneeTotalDonation(i) <> "" Then
                        DoneeWith100Appr("DonationAmt") = UVCase(CDbl(GC80DoneeTotalDonation(i)))
                    Else
                        DoneeWith100Appr("DonationAmt") = 0
                    End If
                    
                    If GC80DoneeEligibleAmnt(i) <> "" Then
                        DoneeWith100Appr("EligibleDonationAmt") = UVCase(CDbl(GC80DoneeEligibleAmnt(i)))
                    Else
                        DoneeWith100Appr("EligibleDonationAmt") = 0
                    End If
                    DoneeWithPan100Appr.add DoneeWith100Appr
                    Set DoneeWith100Appr = Nothing
                    Set DoneeWith100Appr = CreateObject("Scripting.Dictionary")
    
                UpdateProgressBar
                Next
                Don100PercentApprReqd.add "DoneeWithPan", DoneeWithPan100Appr
            
                If GC80DoneeTot_InCash <> "" Then
                    Don100PercentApprReqd("TotDon100PercentApprReqdCash") = UVCase(CDbl(GC80DoneeTot_InCash))
                Else
                    Don100PercentApprReqd("TotDon100PercentApprReqdCash") = 0
                End If
                
                If GC80DoneeTotAmnt <> "" Then
                    Don100PercentApprReqd("TotDon100PercentApprReqdOtherMode") = UVCase(CDbl(GC80DoneeTotAmnt))
                Else
                    Don100PercentApprReqd("TotDon100PercentApprReqdOtherMode") = 0
                End If
                
                If GC80DoneeTot_TotalDonation <> "" Then
                    Don100PercentApprReqd("TotDon100PercentApprReqd") = UVCase(CDbl(GC80DoneeTot_TotalDonation))
                Else
                    Don100PercentApprReqd("TotDon100PercentApprReqd") = 0
                End If
                
                
                If GC80DoneeTotEligAmnt <> "" Then
                    Don100PercentApprReqd("TotEligibleDon100PercentApprReqd") = UVCase(CDbl(GC80DoneeTotEligAmnt))
                Else
                    Don100PercentApprReqd("TotEligibleDon100PercentApprReqd") = 0
                End If
                jsonDictionary.add "Don100PercentApprReqd", Don100PercentApprReqd
                End If
                End If
                


                If Not IsEmpty(GD80DoneeName) Then
                If (UBound(GD80DoneeName) > 0) Then
                    subProcCaption = "80GD"
                    noOfProcessSub = UBound(GD80DoneeName)
                    
                    For i = 0 To UBound(GD80DoneeName) - 1
                    
                        If GD80DoneeName(i) <> "" Then
                            DoneeWith50Appr("DoneeWithPanName") = UCase(GD80DoneeName(i))
                        End If
                        If GD80DoneePAN(i) <> "" Then
                            DoneeWith50Appr("DoneePAN") = UCase(GD80DoneePAN(i))
                        End If
                        If GD80DoneeARN(i) <> "" Then
                            DoneeWith50Appr("ArnNbr") = UCase(GD80DoneeARN(i))
                        End If
                        If GD80DoneeAdrDet(i) <> "" Then
                            AddressDetai50Appr("AddrDetail") = UCase(GD80DoneeAdrDet(i))
                        End If
                        If GD80DoneeCityTownn(i) <> "" Then
                            AddressDetai50Appr("CityOrTownOrDistrict") = UCase(GD80DoneeCityTownn(i))
                        End If
                        If GD80DoneeState(i) <> "" And GD80DoneeState(i) <> "(Select)" Then
                            AddressDetai50Appr("StateCode") = UCase(Mid(Trim(GD80DoneeState(i)), 1, InStr(1, Trim(GD80DoneeState(i)), "-") - 1))
                        End If
                        If GD80DoneePinCode(i) <> "" Then
                            AddressDetai50Appr("PinCode") = UVCase(CDbl(GD80DoneePinCode(i)))
                        End If
                        DoneeWith50Appr.add "AddressDetail", AddressDetai50Appr
                        Set AddressDetai50Appr = Nothing
                        Set AddressDetai50Appr = CreateObject("Scripting.Dictionary")
    
                        If GD80DoneeInCash(i) <> "" Then
                            DoneeWith50Appr("DonationAmtCash") = UVCase(CDbl(GD80DoneeInCash(i)))
                        Else
                            DoneeWith50Appr("DonationAmtCash") = 0
                        End If
                        
                        If GD80DoneeAmnt(i) <> "" Then
                            DoneeWith50Appr("DonationAmtOtherMode") = UVCase(CDbl(GD80DoneeAmnt(i)))
                        Else
                            DoneeWith50Appr("DonationAmtOtherMode") = 0
                        End If
'Konda Updated 09-03-2026--V0.5
                        If GD80DoneeTraref(i) <> "" Then
                            DoneeWith50Appr("TransactionRefNum") = UCase(GD80DoneeTraref(i))
                        End If
                        
                        If GD80DoneeIFSC(i) <> "" Then
                            DoneeWith50Appr("IFSCCode") = UVCase(GD80DoneeIFSC(i))
                        End If
'==========================
                        
                        If GD80DoneeTotalDonation(i) <> "" Then
                            DoneeWith50Appr("DonationAmt") = UVCase(CDbl(GD80DoneeTotalDonation(i)))
                        Else
                            DoneeWith50Appr("DonationAmt") = 0
                        End If
                        
                        If GD80DoneeEligibleAmnt(i) <> "" Then
                            DoneeWith50Appr("EligibleDonationAmt") = UVCase(CDbl(GD80DoneeEligibleAmnt(i)))
                        Else
                            DoneeWith50Appr("EligibleDonationAmt") = 0
                        End If
                        DoneeWithPan50Appr.add DoneeWith50Appr
                        Set DoneeWith50Appr = Nothing
                        Set DoneeWith50Appr = CreateObject("Scripting.Dictionary")
    
                    UpdateProgressBar
                    Next
                    Don50PercentApprReqd.add "DoneeWithPan", DoneeWithPan50Appr
                
                If GD80DoneeTot_InCash <> "" Then
                    Don50PercentApprReqd("TotDon50PercentApprReqdCash") = UVCase(CDbl(GD80DoneeTot_InCash))
                Else
                    Don50PercentApprReqd("TotDon50PercentApprReqdCash") = 0
                End If
                
                If GD80DoneeTotAmnt <> "" Then
                    Don50PercentApprReqd("TotDon50PercentApprReqdOtherMode") = UVCase(CDbl(GD80DoneeTotAmnt))
                Else
                    Don50PercentApprReqd("TotDon50PercentApprReqdOtherMode") = 0
                End If
                
                If GD80DoneeTot_TotalDonation <> "" Then
                    Don50PercentApprReqd("TotDon50PercentApprReqd") = UVCase(CDbl(GD80DoneeTot_TotalDonation))
                Else
                    Don50PercentApprReqd("TotDon50PercentApprReqd") = 0
                End If
                
                If GD80DoneeTotEligAmnt <> "" Then
                    Don50PercentApprReqd("TotEligibleDon50PercentApprReqd") = UVCase(CDbl(GD80DoneeTotEligAmnt))
                Else
                    Don50PercentApprReqd("TotEligibleDon50PercentApprReqd") = 0
                End If
                jsonDictionary.add "Don50PercentApprReqd", Don50PercentApprReqd
                End If
                End If
                

        If G80DoneeTotCash <> "" Then
            jsonDictionary("TotalDonationsUs80GCash") = UVCase(CDbl(G80DoneeTotCash))
        Else
            jsonDictionary("TotalDonationsUs80GCash") = 0
        End If
        
        If G80DoneeTotAmnt <> "" Then
            jsonDictionary("TotalDonationsUs80GOtherMode") = UVCase(CDbl(G80DoneeTotAmnt))
        Else
            jsonDictionary("TotalDonationsUs80GOtherMode") = 0
        End If
        
        If G80DoneeTotTotalDonation <> "" Then
            jsonDictionary("TotalDonationsUs80G") = UVCase(CDbl(G80DoneeTotTotalDonation))
        Else
            jsonDictionary("TotalDonationsUs80G") = 0
        End If
        
        If G80DoneeTotEligAmnt <> "" Then
            jsonDictionary("TotalEligibleDonationsUs80G") = UVCase(CDbl(G80DoneeTotEligAmnt))
        Else
            jsonDictionary("TotalEligibleDonationsUs80G") = 0
        End If
        Set Schedule80G = jsonDictionary
End Function
Function Schedule80D()
On Error Resume Next
subProcCaption = "Schedule 80D"
noOfProcessSub = 10
If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then

Dim jsonDictionary, Sec80DSelfFamSrCtznHealth

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Sec80DSelfFamSrCtznHealth = CreateObject("Scripting.Dictionary")

FamilyMember = Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value
    
    If FamilyMember = "Yes" Then
     FamilyMember = "Y"
     ElseIf FamilyMember = "No" Then
     FamilyMember = "N"
     ElseIf FamilyMember = "Not Claiming for Self/Family" Then
     FamilyMember = "S"
     End If


    If FamilyMember <> "" And FamilyMember <> "(Select)" Then
        Sec80DSelfFamSrCtznHealth("SeniorCitizenFlag") = UCase(FamilyMember)
    Else
        Sec80DSelfFamSrCtznHealth("SeniorCitizenFlag") = ""
    End If
    
    If Sheet16.Range("Self_And_Family_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("SelfAndFamily") = Sheet16.Range("Self_And_Family_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("SelfAndFamily") = 0
    End If
    'Malli-------------------------16/04/2025
    
    If Sheet16.Range("Health_Insurance_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("HealthInsPremSlfFam") = Sheet16.Range("Health_Insurance_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("HealthInsPremSlfFam") = 0
    End If
    
    '_________________________________
    
Dim i As Long
Dim mIntCells80DA As Variant
Dim Sch80DInsDtls, Sch80DInsDtls_G, InsDtls80Da, Sec80DSelfFamHIDtls

'Set InsDtls80Da = CreateObject("Scripting.Dictionary")
Set Sec80DSelfFamHIDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Da = New Collection

'Sec80DSelfFamHIDtls("HealthInsPremSlfFam") = UVCase(Sheet16.Range("Health_Insurance_80D").value)

If Sheet16.Range("TotAmtA1.80D").value > 0 Then

  
mIntCells80DA = Sheet16.Range("NameInsurerA1.80D").Rows.count
 
For i = 1 To mIntCells80DA

If Sheet16.Range("NameInsurerA1.80D").Cells(i, 1).value <> "" Then

'Debug.Print Sheet16.Range("NameInsurerA1.80D").Cells(i, 1).value

'Dim InsurerName, PolicyNo, ReceiptNo, HealthInsAmt,TotalPayments
Dim InsurerName_80DA, PolicyNo_80DA, ReceiptNo_80DA, HealthInsAmt_80DA, TotalPayments



InsurerName_80DA = Sheet16.Range("NameInsurerA1.80D").Cells(i, 1).value
PolicyNo_80DA = Sheet16.Range("PolicyNumA1.80D").Cells(i, 1).value
' 'AY_25_26_V0.7
'ReceiptNo_80DA = Sheet16.Range("ReceiptNumA1.80D").Cells(i, 1).value
'--------------
HealthInsAmt_80DA = Sheet16.Range("AmtA1.80D").Cells(i, 1).value

If InsurerName_80DA <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80DA)
End If

If PolicyNo_80DA <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80DA)
End If

'AY_25_26_V0.7
'If ReceiptNo_80DA <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80DA)
'End If
'---------------

If HealthInsAmt_80DA <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80DA)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
End If
Next
'End If
                
                Sec80DSelfFamHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DSelfFamHIDtls("TotalPayments") = UVCase(Sheet16.Range("TotAmtA1.80D").value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DSelfFamHIDtls", Sec80DSelfFamHIDtls
                Set Sec80DSelfFamHIDtls = Nothing
                Set Sec80DSelfFamHIDtls = CreateObject("Scripting.Dictionary")
End If
'---------------------------------

    If Sheet16.Range("Preventive_Health_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFam") = Sheet16.Range("Preventive_Health_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFam") = 0
    End If
    
    If Sheet16.Range("Senior_Citizen_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("SelfAndFamilySeniorCitizen") = Sheet16.Range("Senior_Citizen_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("SelfAndFamilySeniorCitizen") = 0
    End If
    
    '----------------------------
    If Sheet16.Range("Health_InsuranceSC_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("HlthInsPremSlfFamSrCtzn") = Sheet16.Range("Health_InsuranceSC_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("HlthInsPremSlfFamSrCtzn") = 0
    End If
    '_____________________________
Dim mIntCells80Db As Variant

Dim InsDtls80Db, Sec80DSelfFamSrCtznHIDtls
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Db = CreateObject("Scripting.Dictionary")
Set Sec80DSelfFamSrCtznHIDtls = CreateObject("Scripting.Dictionary")
 
' Sec80DSelfFamSrCtznHIDtls("HlthInsPremSlfFamSrCtzn") = UVCase(Sheet16.Range("Health_InsuranceSC_80D").value)

If Sheet16.Range("TotAmtB1.80D").value > 0 Then
 
 mIntCells80Db = Sheet16.Range("NameInsurerB1.80D").Rows.count
 
For i = 1 To mIntCells80Db

If Sheet16.Range("NameInsurerB1.80D").Cells(i, 1).value <> "" Then

 
Dim InsurerName_80Db, PolicyNo_80Db, ReceiptNo_80Db, HealthInsAmt_80Db



InsurerName_80Db = Sheet16.Range("NameInsurerB1.80D").Cells(i, 1).value
PolicyNo_80Db = Sheet16.Range("PolicyNumB1.80D").Cells(i, 1).value
  'AY_25_26_V0.7
'ReceiptNo_80Db = Sheet16.Range("ReceiptNumB1.80D").Cells(i, 1).value
'-----------------

HealthInsAmt_80Db = Sheet16.Range("AmtB1.80D").Cells(i, 1).value

If InsurerName_80Db <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Db)
End If

If PolicyNo_80Db <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Db)
End If

  'AY_25_26_V0.7
'If ReceiptNo_80Db <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Db)
'End If
'-------------

If HealthInsAmt_80Db <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Db)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
                
End If
Next
'End If
                Sec80DSelfFamSrCtznHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DSelfFamSrCtznHIDtls("TotalPayments") = UVCase(Sheet16.Range("TotAmtB1.80D").value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DSelfFamSrCtznHIDtls", Sec80DSelfFamSrCtznHIDtls
                Set Sec80DSelfFamSrCtznHIDtls = Nothing
                Set Sec80DSelfFamSrCtznHIDtls = CreateObject("Scripting.Dictionary")


End If
    
    '-----------------------------
    If Sheet16.Range("Preventive_Health_SC_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFamSrCtzn") = Sheet16.Range("Preventive_Health_SC_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFamSrCtzn") = 0
    End If
    
    If Sheet16.Range("Medical_Expenditure_SC_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("MedicalExpSlfFamSrCtzn") = Sheet16.Range("Medical_Expenditure_SC_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("MedicalExpSlfFamSrCtzn") = 0
    End If

    
If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then
Dim SeniorCitizen As Variant
SeniorCitizen = Sheet16.Range("DropDown_ValueOf_SC_80D").value
    
    If SeniorCitizen = "Yes" Then
     SeniorCitizen = "Y"
     ElseIf SeniorCitizen = "No" Then
     SeniorCitizen = "N"
     ElseIf SeniorCitizen = "Not claiming for Parents" Then
     SeniorCitizen = "P"
    End If

    If SeniorCitizen <> "" And SeniorCitizen <> "(Select)" Then
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizenFlag") = UCase(SeniorCitizen)
    Else
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizenFlag") = ""
    End If
    
    If Sheet16.Range("Parents_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("Parents") = Sheet16.Range("Parents_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("Parents") = 0
    End If
    '------------------------------------------
    If Sheet16.Range("Health_Insurance2_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("HlthInsPremParents") = Sheet16.Range("Health_Insurance2_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("HlthInsPremParents") = 0
    End If
    '___________________________________________
    Dim mIntCells80Dc As Variant
Dim InsDtls80Dc, Sec80DParentsHIDtls
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Dc = CreateObject("Scripting.Dictionary")
Set Sec80DParentsHIDtls = CreateObject("Scripting.Dictionary")
 
 'Sec80DParentsHIDtls("HlthInsPremParents") = UVCase(Sheet16.Range("Health_Insurance2_80D").value)

If Sheet16.Range("TotAmtA2.80D").value > 0 Then
 
 mIntCells80Dc = Sheet16.Range("NameInsurerA2.80D").Rows.count
 
For i = 1 To mIntCells80Dc

If Sheet16.Range("NameInsurerA2.80D").Cells(i, 1).value <> "" Then

 
Dim InsurerName_80Dc, PolicyNo_80Dc, ReceiptNo_80Dc, HealthInsAmt_80Dc


InsurerName_80Dc = Sheet16.Range("NameInsurerA2.80D").Cells(i, 1).value
PolicyNo_80Dc = Sheet16.Range("PolicyNumA2.80D").Cells(i, 1).value
'AY_25_26_V0.7
'ReceiptNo_80Dc = Sheet16.Range("ReceiptNumA2.80D").Cells(i, 1).value
'-----------------------------
HealthInsAmt_80Dc = Sheet16.Range("AmtA2.80D").Cells(i, 1).value

If InsurerName_80Dc <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Dc)
End If

If PolicyNo_80Dc <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Dc)
End If
 'AY_25_26_V0.7
'If ReceiptNo_80Dc <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Dc)
'End If
'--------------------
If HealthInsAmt_80Dc <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Dc)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
                
End If
Next
'End If

                Sec80DParentsHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DParentsHIDtls("TotalPayments") = UVCase(Sheet16.Range("TotAmtA2.80D").value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DParentsHIDtls", Sec80DParentsHIDtls
                Set Sec80DParentsHIDtls = Nothing
                Set Sec80DParentsHIDtls = CreateObject("Scripting.Dictionary")


End If


    '-------------------------------------------
    If Sheet16.Range("Preventive_Health2_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParents") = Sheet16.Range("Preventive_Health2_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParents") = 0
    End If
    
    If Sheet16.Range("Parents_SC_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizen") = Sheet16.Range("Parents_SC_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizen") = 0
    End If
    '---------------------------------------------
    If Sheet16.Range("Health_Insurance3_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("HlthInsPremParentsSrCtzn") = Sheet16.Range("Health_Insurance3_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("HlthInsPremParentsSrCtzn") = 0
    End If
    '________________________________________________
    Dim mIntCells80Dd As Variant
Dim InsDtls80Dd, Sec80DParentsSrCtznHIDtls
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Dd = CreateObject("Scripting.Dictionary")
Set Sec80DParentsSrCtznHIDtls = CreateObject("Scripting.Dictionary")
 

'Sec80DParentsSrCtznHIDtls("HlthInsPremParentsSrCtzn") = UVCase(Sheet16.Range("Health_Insurance3_80D").value)

If Sheet16.Range("TotAmtB2.80D").value > 0 Then
 
 mIntCells80Dd = Sheet16.Range("NameInsurerB2.80D").Rows.count
 
For i = 1 To mIntCells80Dd

If Sheet16.Range("NameInsurerB2.80D").Cells(i, 1).value <> "" Then

 
Dim InsurerName_80Dd, PolicyNo_80Dd, ReceiptNo_80Dd, HealthInsAmt_80Dd



InsurerName_80Dd = Sheet16.Range("NameInsurerB2.80D").Cells(i, 1).value
PolicyNo_80Dd = Sheet16.Range("PolicyNumB2.80D").Cells(i, 1).value
''AY_25_26_V0.7
'ReceiptNo_80Dd = Sheet16.Range("ReceiptNumB2.80D").Cells(i, 1).value
'--------------
HealthInsAmt_80Dd = Sheet16.Range("AmtB2.80D").Cells(i, 1).value

If InsurerName_80Dd <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Dd)
End If

If PolicyNo_80Dd <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Dd)
End If
'AY_25_26_V0.7
'If ReceiptNo_80Dd <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Dd)
'End If
'------------------

If HealthInsAmt_80Dd <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Dd)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
End If
Next
'End If

                Sec80DParentsSrCtznHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DParentsSrCtznHIDtls("TotalPayments") = UVCase(Sheet16.Range("TotAmtB2.80D").value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DParentsSrCtznHIDtls", Sec80DParentsSrCtznHIDtls
                Set Sec80DParentsSrCtznHIDtls = Nothing
                Set Sec80DParentsSrCtznHIDtls = CreateObject("Scripting.Dictionary")


End If

    '------------------------------------------------
    If Sheet16.Range("Preventive_Health3_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParentsSrCtzn") = Sheet16.Range("Preventive_Health3_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParentsSrCtzn") = 0
    End If
    
    If Sheet16.Range("Medical_Expenditure2_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("MedicalExpParentsSrCtzn") = Sheet16.Range("Medical_Expenditure2_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("MedicalExpParentsSrCtzn") = 0
    End If
End If
    
    If Sheet16.Range("Eligible_Amount_80D").value <> "" Then
        Sec80DSelfFamSrCtznHealth("EligibleAmountOfDedn") = Sheet16.Range("Eligible_Amount_80D").value
    Else
        Sec80DSelfFamSrCtznHealth("EligibleAmountOfDedn") = 0
    End If

    jsonDictionary.add "Sec80DSelfFamSrCtznHealth", Sec80DSelfFamSrCtznHealth





'Malli---------------15/04/2025
'InsDtls80Da

'Dim i As Long
'Dim mIntCells80DA As Variant
'Dim Sch80DInsDtls, Sch80DInsDtls_G, InsDtls80Da
'
'Set InsDtls80Da = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = New Collection
''Set InsDtls80Da = New Collection
'
'
'If Sheet16.Range("TotAmtA1.80D").value > 0 Then
'
' mIntCells80DA = Sheet16.Range("NameInsurerA1.80D").Rows.count
'
'For i = 1 To mIntCells80DA
'
'If Sheet16.Range("NameInsurerA1.80D").Cells(i, 1).value <> "" Then
'
''Dim InsurerName, PolicyNo, ReceiptNo, HealthInsAmt,TotalPayments
'Dim InsurerName_80DA, PolicyNo_80DA, ReceiptNo_80DA, HealthInsAmt_80DA, TotalPayments
'
'
'
'InsurerName_80DA = Sheet16.Range("NameInsurerA1.80D").Cells(i, 1).value
'PolicyNo_80DA = Sheet16.Range("PolicyNumA1.80D").Cells(i, 1).value
'ReceiptNo_80DA = Sheet16.Range("ReceiptNumA1.80D").Cells(i, 1).value
'HealthInsAmt_80DA = Sheet16.Range("AmtA1.80D").Cells(i, 1).value
'
'If InsurerName_80DA <> "" Then
'Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80DA)
'End If
'
'If PolicyNo_80DA <> "" Then
'Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80DA)
'End If
'
'If ReceiptNo_80DA <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80DA)
'End If
'
'If HealthInsAmt_80DA <> "" Then
'Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80DA)
'End If
'
'End If
'
'                Sch80DInsDtls.add Sch80DInsDtls_G
'                Set Sch80DInsDtls_G = Nothing
'                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'
'Next
'
'                InsDtls80Da.add "Sch80DInsDtls", Sch80DInsDtls
'                Set Sch80DInsDtls = Nothing
'                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'
'
'InsDtls80Da("TotalPayments") = UVCase(Sheet16.Range("TotAmtA1.80D").value)
'
'                Sec80DSelfFamSrCtznHealth.add "InsDtls80Da", InsDtls80Da
'                Set InsDtls80Da = Nothing
'                Set InsDtls80Da = CreateObject("Scripting.Dictionary")
'
'
'End If


'InsDtls80Db

 
'Dim mIntCells80Db As Variant
'Dim InsDtls80Db
'Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = New Collection
'Set InsDtls80Db = CreateObject("Scripting.Dictionary")
'
'
'
'If Sheet16.Range("TotAmtB1.80D").value > 0 Then
'
' mIntCells80Db = Sheet16.Range("NameInsurerB1.80D").Rows.count
'
'For i = 1 To mIntCells80Db
'
'If Sheet16.Range("NameInsurerB1.80D").Cells(i, 1).value <> "" Then
'
'
'Dim InsurerName_80Db, PolicyNo_80Db, ReceiptNo_80Db, HealthInsAmt_80Db
'
'
'
'InsurerName_80Db = Sheet16.Range("NameInsurerB1.80D").Cells(i, 1).value
'PolicyNo_80Db = Sheet16.Range("PolicyNumB1.80D").Cells(i, 1).value
'ReceiptNo_80Db = Sheet16.Range("ReceiptNumB1.80D").Cells(i, 1).value
'HealthInsAmt_80Db = Sheet16.Range("AmtB1.80D").Cells(i, 1).value
'
'If InsurerName_80Db <> "" Then
'Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Db)
'End If
'
'If PolicyNo_80Db <> "" Then
'Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Db)
'End If
'
'If ReceiptNo_80Db <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Db)
'End If
'
'If HealthInsAmt_80Db <> "" Then
'Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Db)
'End If
'
'End If
'
'                Sch80DInsDtls.add Sch80DInsDtls_G
'                Set Sch80DInsDtls_G = Nothing
'                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'
'Next
'
'                InsDtls80Db.add "Sch80DInsDtls", Sch80DInsDtls
'                Set Sch80DInsDtls = Nothing
'                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'
'
'InsDtls80Db("TotalPayments") = UVCase(Sheet16.Range("TotAmtB1.80D").value)
'
'                Sec80DSelfFamSrCtznHealth.add "InsDtls80Db", InsDtls80Db
'                Set InsDtls80Db = Nothing
'                Set InsDtls80Db = CreateObject("Scripting.Dictionary")
'
'
'End If

'InsDtls80Dc

'Dim mIntCells80Dc As Variant
'Dim InsDtls80Dc
'Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = New Collection
'Set InsDtls80Dc = CreateObject("Scripting.Dictionary")
'
'
'
'If Sheet16.Range("TotAmtA2.80D").value > 0 Then
'
' mIntCells80Dc = Sheet16.Range("NameInsurerA2.80D").Rows.count
'
'For i = 1 To mIntCells80Dc
'
'If Sheet16.Range("NameInsurerA2.80D").Cells(i, 1).value <> "" Then
'
'
'Dim InsurerName_80Dc, PolicyNo_80Dc, ReceiptNo_80Dc, HealthInsAmt_80Dc
'
'
'InsurerName_80Dc = Sheet16.Range("NameInsurerA2.80D").Cells(i, 1).value
'PolicyNo_80Dc = Sheet16.Range("PolicyNumA2.80D").Cells(i, 1).value
'ReceiptNo_80Dc = Sheet16.Range("ReceiptNumA2.80D").Cells(i, 1).value
'HealthInsAmt_80Dc = Sheet16.Range("AmtA2.80D").Cells(i, 1).value
'
'If InsurerName_80Dc <> "" Then
'Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Dc)
'End If
'
'If PolicyNo_80Dc <> "" Then
'Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Dc)
'End If
'
'If ReceiptNo_80Dc <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Dc)
'End If
'
'If HealthInsAmt_80Dc <> "" Then
'Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Dc)
'End If
'
'End If
'
'                Sch80DInsDtls.add Sch80DInsDtls_G
'                Set Sch80DInsDtls_G = Nothing
'                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'
'Next
'
'                InsDtls80Dc.add "Sch80DInsDtls", Sch80DInsDtls
'                Set Sch80DInsDtls = Nothing
'                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'
'
'InsDtls80Dc("TotalPayments") = UVCase(Sheet16.Range("TotAmtA2.80D").value)
'
'                Sec80DSelfFamSrCtznHealth.add "InsDtls80Dc", InsDtls80Dc
'                Set InsDtls80Dc = Nothing
'                Set InsDtls80Dc = CreateObject("Scripting.Dictionary")
'
'
'End If
'

'InsDtls80Dd
'Dim mIntCells80Dd As Variant
'Dim InsDtls80Dd
'Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'Set Sch80DInsDtls = New Collection
'Set InsDtls80Dd = CreateObject("Scripting.Dictionary")
'
'
'
'If Sheet16.Range("TotAmtB2.80D").value > 0 Then
'
' mIntCells80Dd = Sheet16.Range("NameInsurerB2.80D").Rows.count
'
'For i = 1 To mIntCells80Dd
'
'If Sheet16.Range("NameInsurerB2.80D").Cells(i, 1).value <> "" Then
'
'
'Dim InsurerName_80Dd, PolicyNo_80Dd, ReceiptNo_80Dd, HealthInsAmt_80Dd
'
'
'
'InsurerName_80Dd = Sheet16.Range("NameInsurerB2.80D").Cells(i, 1).value
'PolicyNo_80Dd = Sheet16.Range("PolicyNumB2.80D").Cells(i, 1).value
'ReceiptNo_80Dd = Sheet16.Range("ReceiptNumB2.80D").Cells(i, 1).value
'HealthInsAmt_80Dd = Sheet16.Range("AmtB2.80D").Cells(i, 1).value
'
'If InsurerName_80Dd <> "" Then
'Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Dd)
'End If
'
'If PolicyNo_80Dd <> "" Then
'Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Dd)
'End If
'
'If ReceiptNo_80Dd <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Dd)
'End If
'
'If HealthInsAmt_80Dd <> "" Then
'Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Dd)
'End If
'
'End If
'
'                Sch80DInsDtls.add Sch80DInsDtls_G
'                Set Sch80DInsDtls_G = Nothing
'                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
'
'Next
'
'                InsDtls80Dd.add "Sch80DInsDtls", Sch80DInsDtls
'                Set Sch80DInsDtls = Nothing
'                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
'
'
'InsDtls80Dd("TotalPayments") = UVCase(Sheet16.Range("TotAmtB2.80D").value)
'
'                Sec80DSelfFamSrCtznHealth.add "InsDtls80Dd", InsDtls80Dd
'                Set InsDtls80Dd = Nothing
'                Set InsDtls80Dd = CreateObject("Scripting.Dictionary")
'
'
'End If


 
'--------------------




End If

Set Schedule80D = jsonDictionary
End Function
Function TaxExmpIntIncDtls()
 
Dim i As Long
Dim jsonDictionary, OthersInc, OthersIncDtl, OthersIncDtls

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set OthersInc = CreateObject("Scripting.Dictionary")
Set OthersIncDtl = CreateObject("Scripting.Dictionary")

Set OthersIncDtls = New Collection

             If Trim(Sheet5.Range("Total_Exempt").value) <> "" Then
                    OthersInc("OthersTotalTaxExe") = Sheet5.Range("Total_Exempt").value
                Else
                    OthersInc("OthersTotalTaxExe") = 0
                End If

            If (Not IsEmpty(Others_NOI)) And UBound(Others_NOI) > 0 Then
            
'Commented by Konda on 09-03-2026--V0.5
'                For i = 1 To UBound(Others_NOI)
'
'                    If (Others_NOI(i) = "Any Other") Then
'                                Others_NOI(i) = "OTH"
'                    ElseIf (Others_NOI(i) = "Defense Medical Disability Pension") Then
'                        Others_NOI(i) = "DMDP"
'                    ElseIf (Others_NOI(i) = "Agriculture Income (Less than Equal to Rs.5000)") Then
'                        Others_NOI(i) = "AGRI"
'                    ElseIf (Others_NOI(i) = "Sec 10(34)-(Exempted Dividend Income)") Then
'                        Others_NOI(i) = "10(34)"
'                    ElseIf (Others_NOI(i) = "Sec 10(26AAA)-Any income as referred to in section 10(26AAA)") Then
'                        Others_NOI(i) = "10(26AAA)"
'                    ElseIf (Others_NOI(i) = "Sec 10(10BC)-Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster") Then
'                        Others_NOI(i) = "10(10BC)"
'                    ElseIf (Others_NOI(i) = "Sec 10(10D)-Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(10D)") Then
'                        Others_NOI(i) = "10(10D)"
'                    ElseIf (Others_NOI(i) = "Sec 10(19)-Family pension received") Then
'                        Others_NOI(i) = "10(19)"
'                    ElseIf (Others_NOI(i) = "Sec 10(26)-Any income as referred to in section 10(26)") Then
'                        Others_NOI(i) = "10(26)"
'
'                    ElseIf Mid(Others_NOI(i), 1, 11) = "Sec 10(12C)" Then
'                        Others_NOI(i) = "10(12C)"
'                    'Generate-25.01.2023.103.TAX.05
''                    ElseIf (Others_NOI(i) = "Any exempt income including LTCG on which tax is not payable") Then
''                        Others_NOI(i) = "CG1L"
''                    ElseIf (Others_NOI(i) = "LTCG u/s 112A not exceeding  Rs. 1 Lakh") Then
''                        Others_NOI(i) = "LTCG"
'                    'End change
'                    Else
'                         If InStr(Others_NOI(i), "-") > 0 Then
'                             Others_NOI(i) = Trim(Mid(Others_NOI(i), 4, (InStr(Trim(Others_NOI(i)), "-") - 4)))
'                         End If
'                    End If
'
'                Next
      
'Konda updated on 09-03-2026--V0.5
                For i = 1 To UBound(Others_NOI)
                    
                    If (Others_NOI(i) = "Agricultural  & related incomes") Then
                        Others_NOI(i) = "AGRI"
                    ElseIf (Others_NOI(i) = "Compensation/other sums received by government or other approved entities") Then
                        Others_NOI(i) = "GOVC"
                    ElseIf (Others_NOI(i) = "Income from specified Investments") Then
                        Others_NOI(i) = "ISI"
                    ElseIf (Others_NOI(i) = "Specified sums received by armed forces personnel") Then
                        Others_NOI(i) = "SSRA"
                    ElseIf (Others_NOI(i) = "Sums received by Senior Citizens/Minors ") Then
                        Others_NOI(i) = "SRSC"
                    ElseIf (Others_NOI(i) = "Sums received by specified Category of Taxpayers") Then
                        Others_NOI(i) = "SRST"
                    ElseIf (Others_NOI(i) = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana") Then
                        Others_NOI(i) = "SRPC"
                    ElseIf (Others_NOI(i) = "Other Incomes") Then
                        Others_NOI(i) = "OTH"
                    End If
                            
                Next

'SubCategory

        For i = 1 To UBound(Others_NOI1)

                    If (Others_NOI1(i) = "10(1)-Agricultural income(Less than or equal to 5000)") Then
                        Others_NOI1(i) = "10(1)"
                    ElseIf (Others_NOI1(i) = "10(30)-subsidy received from or through the Tea Board") Then
                        Others_NOI1(i) = "10(30)"
                    ElseIf (Others_NOI1(i) = "10(31)-Rubber/Coffee/Tea development accounts/funds") Then
                        Others_NOI1(i) = "10(31)"
                    ElseIf (Others_NOI1(i) = "10(10BB)-payments made under the Bhopal Gas Leak Disaster") Then
                        Others_NOI1(i) = "10(10BB)"
                    ElseIf (Others_NOI1(i) = "10(10BC)-Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster") Then
                        Others_NOI1(i) = "10(10BC)"
                    ElseIf (Others_NOI1(i) = "10(17A)-Award instituted by Government") Then
                        Others_NOI1(i) = "10(17A)"
                    ElseIf (Others_NOI1(i) = "10(12AB)-any sum received as lump sum amount as per clause (vi) of paragraph 2 of the notification number FX-1/3/2024-PR") Then
                        Others_NOI1(i) = "10(12AB)"
                    ElseIf (Others_NOI1(i) = "10(15)-Interest on specified securities/investments") Then
                        Others_NOI1(i) = "10(15)"
'                    ElseIf (Others_NOI1(i) = "10(23EA)-Contributions received from recognised stock exchanges") Then
'                        Others_NOI1(i) = "10(23EA)"
                    ElseIf (Others_NOI1(i) = "10(23FBB)-income referred to in section 115UB, accruing or arising to, or received by, a unit holder of an investment fund") Then
                        Others_NOI1(i) = "10(23FBB)"
                    ElseIf (Others_NOI1(i) = "10(23FD)Unit holder income from Business Trust (certain parts)") Then
                        Others_NOI1(i) = "10(23FD)"
                    ElseIf (Others_NOI1(i) = "10(35)-Income from specified Mutual Funds ") Then
                        Others_NOI1(i) = "10(35)"
                    ElseIf (Others_NOI1(i) = "10(35A)-distributed income referred to in section 115TA received from a securitisation trust") Then
                        Others_NOI1(i) = "10(35A)"
                    ElseIf (Others_NOI1(i) = "10(12C)-Agniveer Corpus Fund income") Then
                        Others_NOI1(i) = "10(12C)"
                    ElseIf (Others_NOI1(i) = "10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award") Then
                        Others_NOI1(i) = "10(18)"
                    ElseIf (Others_NOI1(i) = "10(19)-Armed Forces Family pension in case of death during operational duty") Then
                        Others_NOI1(i) = "10(19)"
                    ElseIf (Others_NOI1(i) = "10(23AA)-Sum received by any person on behalf of any Fund established by the armed forces") Then
                        Others_NOI1(i) = "10(23AA)"
                    ElseIf (Others_NOI1(i) = "Defense Medical Disability Pension") Then
                        Others_NOI1(i) = "DMD"
                    ElseIf (Others_NOI1(i) = "10(32)-Minor child’s income—small exemption") Then
                        Others_NOI1(i) = "10(32)"
                    ElseIf (Others_NOI1(i) = "10(43)-Reverse mortgage—payments to senior citizens") Then
                        Others_NOI1(i) = "10(43)"
                    ElseIf (Others_NOI1(i) = "10(19A)-Annual value of one palace in occupation of ex-ruler") Then
                        Others_NOI1(i) = "10(19A)"
                    ElseIf (Others_NOI1(i) = "10(26)-Any income as referred to in section 10(26)") Then
                        Others_NOI1(i) = "10(26)"
                    ElseIf (Others_NOI1(i) = "10(26AAA)-Any income as referred to in section 10(26AAA)") Then
                        Others_NOI1(i) = "10(26AAA)"
                    ElseIf (Others_NOI1(i) = "10(10D)-Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(10D)") Then
                        Others_NOI1(i) = "10(10D)"
                    ElseIf (Others_NOI1(i) = "10(11)-Statutory Provident Fund received") Then
                        Others_NOI1(i) = "10(11)"
                    ElseIf (Others_NOI1(i) = "10(11A)-Sum received from an account opened under the Sukanya Samriddhi Yojana") Then
                        Others_NOI1(i) = "10(11A)"
                    ElseIf (Others_NOI1(i) = "10(12)-Recognized Provident Fund received") Then
                        Others_NOI1(i) = "10(12)"
                    ElseIf (Others_NOI1(i) = "10(12A)-Any payment from the National Pension System Trust to an assessee") Then
                        Others_NOI1(i) = "10(12A)"
                    ElseIf (Others_NOI1(i) = "10(12AA)-any payment from the National Pension System Trust ") Then
                        Others_NOI1(i) = "10(12AA)"
                    ElseIf (Others_NOI1(i) = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee") Then
                        Others_NOI1(i) = "10(12B)"
                    ElseIf (Others_NOI1(i) = "10(12BA)-partial withdrawal made from the National Pension System") Then
                        Others_NOI1(i) = "10(12BA)"
                    ElseIf (Others_NOI1(i) = "10(13)-Approved superannuation fund received") Then
                        Others_NOI1(i) = "10(13)"
                    ElseIf (Others_NOI1(i) = "10(25)-Sum received by trustees on behalf of approved superannuation, gratuity, or pension funds") Then
                        Others_NOI1(i) = "10(25)"
                    'Ayush_27/03/2026
'                    ElseIf (Others_NOI1(i) = "10(25A)-any income under Employees' State Insurance Fund") Then
'                        Others_NOI1(i) = "10(25A)"
                    ElseIf (Others_NOI1(i) = "10(44)-Income received by any person for, or on behalf of, the New Pension System Trust") Then
                        Others_NOI1(i) = "10(44)"
                    ElseIf (Others_NOI1(i) = "10(2)-Member’s share from HUF") Then
                        Others_NOI1(i) = "10(2)"
                    ElseIf (Others_NOI1(i) = "10(16)-Scholarships for education") Then
                        Others_NOI1(i) = "10(16)"
                    'Malli_AY_2026_27 17/06/2026
                    ElseIf (Others_NOI1(i) = "Income exempt as per CBDT Circular") Then
                        Others_NOI1(i) = "Incmexmptcircular"
                    ElseIf (Others_NOI1(i) = "Income exempt as per CBDT Notification") Then
                        Others_NOI1(i) = "Incmexmptnotification"
                    ElseIf (Others_NOI1(i) = "Receipts not in the nature of income") Then
                        Others_NOI1(i) = "Receiptnotincme"
                    
                    '---------------------------
                    
                    Else
                         If InStr(Others_NOI1(i), "-") > 0 Then
                             Others_NOI1(i) = Trim(Mid(Others_NOI1(i), 4, (InStr(Trim(Others_NOI1(i)), "-") - 4)))
                         End If
                    End If

                Next

'======================================
                For i = 1 To UBound(Others_NOI)
                    
                        If Others_NOI(i) <> "" Then
                            OthersIncDtl("Category") = (Others_NOI(i))
                        Else
                             OthersIncDtl("Category") = ""
                        End If
                        
                        If Others_NOI1(i) <> "" Then
                            OthersIncDtl("SubCategory") = (Others_NOI1(i))
                        Else
                             OthersIncDtl("SubCategory") = ""
                        End If


'                        If Others_NOI(i) <> "" Then
'                            OthersIncDtl("NatureDesc") = (Others_NOI(i))
'                        Else
'                             OthersIncDtl("NatureDesc") = ""
'                        End If
                        
'                        If Others_NOI1(i) <> "" And Others_NOI1(i) <> "Not Applicable" Then
'                            OthersIncDtl("OthNatOfInc") = (Others_NOI1(i))
'                        End If

                        'Malli_AY_2026_27 17/06/2026
                        If Others_NOI1(i) = "Incmexmptcircular" Or Others_NOI1(i) = "Incmexmptnotification" Or Others_NOI1(i) = "Receiptnotincme" Then
                               If Others_NOI22(i) <> "" Then
                                  OthersIncDtl("Description") = Others_NOI22(i)
                               End If
                        End If
                        '---------------------------
                    
                        If Others_Amt(i) <> "" Then
                            OthersIncDtl("OthAmount") = Others_Amt(i)
                        Else
                            OthersIncDtl("OthAmount") = ""
                        End If
                        
                         OthersIncDtls.add OthersIncDtl
                         Set OthersIncDtl = Nothing
                         Set OthersIncDtl = CreateObject("Scripting.Dictionary")
                         
                Next
            End If
             OthersInc.add "OthersIncDtls", OthersIncDtls
            jsonDictionary.add "OthersInc", OthersInc
Set TaxExmpIntIncDtls = jsonDictionary
End Function

'added by Chetan C M AY 2025-26 on 05/02/2025

Function LTCG112A_New() As Object

Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Dim TotSaleCnsdrn, TotCstAcqisn, LongCap112A As Variant
 
'Dim LTCG112A
'Set TotSaleCnsdrn = CreateObject("Scripting.Dictionary")
'Set TotCstAcqisn = CreateObject("Scripting.Dictionary")
'
'Set LTCG112A_New = CreateObject("Scripting.Dictionary")
 
 
If Trim(Sheet5.Range("IncD.Sale_LTCG").value) <> "" Then
                jsonDictionary("TotSaleCnsdrn") = Sheet5.Range("IncD.Sale_LTCG").value
            Else
                jsonDictionary("TotSaleCnsdrn") = 0
            End If
            If Trim(Sheet5.Range("IncD.Cost_LTCG").value) <> "" Then
                jsonDictionary("TotCstAcqisn") = Sheet5.Range("IncD.Cost_LTCG").value
            Else
                jsonDictionary("TotCstAcqisn") = 0
            End If
            If Trim(Sheet5.Range("IncD.CG_LTCG").value) <> "" Then
                jsonDictionary("LongCap112A") = Sheet5.Range("IncD.CG_LTCG").value
            Else
                jsonDictionary("LongCap112A") = 0
            End If
 
'jsonDictionary.add "LTCG112A", LTCG112A_New
Set LTCG112A_New = jsonDictionary
 
End Function

Function TaxReturnPreparer() As Object
subProcCaption = "Tax Computation"
Dim jsonDictionary

Set jsonDictionary = CreateObject("Scripting.Dictionary")
If TRPID <> "" Then
    If TRPID <> "" Then
       jsonDictionary("IdentificationNoOfTRP") = UCase(TRPID)
    End If
    
    If TRPName <> "" Then
        jsonDictionary("NameOfTRP") = UCase(TRPName)
    End If
    
    If TRPreimbursementAmnt <> "" Then
    
       jsonDictionary("ReImbFrmGov") = CDbl(TRPreimbursementAmnt)
    Else
       jsonDictionary("ReImbFrmGov") = 0
    End If
UpdateProgressBar
Set TaxReturnPreparer = jsonDictionary
End If
End Function
Function ScheduleBP() As Object

Dim i As Long
Dim jsonDictionary, PersumptiveInc44AD, PersumptiveInc44ADA, PersumptiveInc44AE, FinanclPartclrOfBusiness
Dim NatOfBus44AD, NatOfBus44ADA, NatOfBus44AE, GoodsDtlsUs44AE, TurnoverGrsRcptForGSTIN
Dim NatOfBus44AD1, NatOfBus44ADA1, NatOfBus44AE1, GoodsDtlsUs44AE1, TurnoverGrsRcptForGSTIN1

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set PersumptiveInc44AD = CreateObject("Scripting.Dictionary")
Set PersumptiveInc44ADA = CreateObject("Scripting.Dictionary")
Set PersumptiveInc44AE = CreateObject("Scripting.Dictionary")
Set FinanclPartclrOfBusiness = CreateObject("Scripting.Dictionary")

Set NatOfBus44AD = New Collection
Set NatOfBus44ADA = New Collection
Set NatOfBus44AE = New Collection
Set GoodsDtlsUs44AE = New Collection
Set TurnoverGrsRcptForGSTIN = New Collection

Set NatOfBus44AD1 = CreateObject("Scripting.Dictionary")
Set NatOfBus44ADA1 = CreateObject("Scripting.Dictionary")
Set NatOfBus44AE1 = CreateObject("Scripting.Dictionary")
Set GoodsDtlsUs44AE1 = CreateObject("Scripting.Dictionary")
Set TurnoverGrsRcptForGSTIN1 = CreateObject("Scripting.Dictionary")

If UBound(NOBCode) > 0 Then
   subProcCaption = "Nature Of Business"
   noOfProcessSub = UBound(NOBCode)
  
   If (Not IsEmpty(NOBCode)) And UBound(NOBCode) > 0 Then
        For i = 1 To UBound(NOBName)
            If NOBName(i) <> "" Then
                NatOfBus44AD1("NameOfBusiness") = (NOBName(i))
            Else
            End If
            If NOBCode(i) <> "" Then
                NatOfBus44AD1("CodeAD") = (NOBCode(i))
            Else
            End If
            If NOBDescription(i) <> "" Then
                NatOfBus44AD1("Description") = NOBDescription(i)
            Else
            End If
            NatOfBus44AD.add NatOfBus44AD1
            Set NatOfBus44AD1 = Nothing
            Set NatOfBus44AD1 = CreateObject("Scripting.Dictionary")
        Next
   End If
End If
jsonDictionary.add "NatOfBus44AD", NatOfBus44AD

'E1 AY 2024-25 Schema Change-------------------------------------

If GrossTurnover_Receipts <> "" Or Not IsEmpty(GrossTurnover_Receipts) Or GrossTurnover_Receipts <> 0 Then
    PersumptiveInc44AD("GrsTotalTrnOver") = GrossTurnover_Receipts
Else
    PersumptiveInc44AD("GrsTotalTrnOver") = 0
End If
'-----------------------------------------------------------------


'E1(a)
If GrossTurnover_a <> "" Or Not IsEmpty(GrossTurnover_a) Or GrossTurnover_a <> 0 Then
    PersumptiveInc44AD("GrsTrnOverBank") = GrossTurnover_a
Else
    PersumptiveInc44AD("GrsTrnOverBank") = 0
End If

'E1(b)
'AY 2024-25 Schema Change------------------------
 
 If GrossTurnover_a_b <> "" Or Not IsEmpty(GrossTurnover_a_b) Or GrossTurnover_a_b <> 0 Then
    PersumptiveInc44AD("GrsTotalTrnOverInCash") = GrossTurnover_a_b
Else
    PersumptiveInc44AD("GrsTotalTrnOverInCash") = 0
End If
'------------------------------------------------



If GrossTurnover_b <> "" Or Not IsEmpty(GrossTurnover_b) Or GrossTurnover_b <> 0 Then
    PersumptiveInc44AD("GrsTrnOverAnyOthMode") = GrossTurnover_b
Else
    PersumptiveInc44AD("GrsTrnOverAnyOthMode") = 0
End If

UpdateProgressBar
If PresumptiveIncome44AD_a <> "" Or Not IsEmpty(PresumptiveIncome44AD_a) Or PresumptiveIncome44AD_a <> 0 Then
   PersumptiveInc44AD("PersumptiveInc44AD6Per") = PresumptiveIncome44AD_a
Else
   PersumptiveInc44AD("PersumptiveInc44AD6Per") = 0
End If
If PresumptiveIncome44AD_b <> "" Or Not IsEmpty(PresumptiveIncome44AD_b) Or PresumptiveIncome44AD_b <> 0 Then
   PersumptiveInc44AD("PersumptiveInc44AD8Per") = PresumptiveIncome44AD_b
Else
   PersumptiveInc44AD("PersumptiveInc44AD8Per") = 0
End If
If PresumptiveIncome44AD <> "" Or Not IsEmpty(PresumptiveIncome44AD) Or PresumptiveIncome44AD <> 0 Then
   PersumptiveInc44AD("TotPersumptiveInc44AD") = PresumptiveIncome44AD
Else
   PersumptiveInc44AD("TotPersumptiveInc44AD") = 0
End If

UpdateProgressBar
jsonDictionary.add "PersumptiveInc44AD", PersumptiveInc44AD

If (Not IsEmpty(NOBCode1)) And UBound(NOBCode1) > 0 Then
      For i = 1 To UBound(NOBName1)
            If NOBName1(i) <> "" Then
                NatOfBus44ADA1("NameOfBusiness") = (NOBName1(i))
            End If
            If NOBCode1(i) <> "" Then
                NatOfBus44ADA1("CodeADA") = (NOBCode1(i))
            End If
            If NOBDescription1(i) <> "" Then
                NatOfBus44ADA1("Description") = NOBDescription1(i)
            End If
            NatOfBus44ADA.add NatOfBus44ADA1
            Set NatOfBus44ADA1 = Nothing
            Set NatOfBus44ADA1 = CreateObject("Scripting.Dictionary")
      Next
End If
jsonDictionary.add "NatOfBus44ADA", NatOfBus44ADA

If GrossTurnover44ADA <> "" Or Not IsEmpty(GrossTurnover44ADA) Or GrossTurnover44ADA <> 0 Then
    PersumptiveInc44ADA("GrsReceipt") = GrossTurnover44ADA
Else
    PersumptiveInc44ADA("GrsReceipt") = 0
End If

'AY 2024-25 Schema Change--------------------

'E3(a)
If GrossReceipts_E3_a <> "" Or Not IsEmpty(GrossReceipts_E3_a) Or GrossReceipts_E3_a <> 0 Then
    PersumptiveInc44ADA("GrsTrnOverBank44ADA") = GrossReceipts_E3_a
Else
    PersumptiveInc44ADA("GrsTrnOverBank44ADA") = 0
End If

'E3(b)
If GrossReceipts_E3_b <> "" Or Not IsEmpty(GrossReceipts_E3_b) Or GrossReceipts_E3_b <> 0 Then
    PersumptiveInc44ADA("GrsTotalTrnOverInCash44ADA") = GrossReceipts_E3_b
Else
    PersumptiveInc44ADA("GrsTotalTrnOverInCash44ADA") = 0
End If

'E3(c)
If GrossReceipts_E3_c <> "" Or Not IsEmpty(GrossReceipts_E3_c) Or GrossReceipts_E3_c <> 0 Then
    PersumptiveInc44ADA("GrsTrnOverAnyOthMode44ADA") = GrossReceipts_E3_c
Else
    PersumptiveInc44ADA("GrsTrnOverAnyOthMode44ADA") = 0
End If

'-------------------------------------------------
'E4
If PresumptiveIncome44ADa <> "" Or Not IsEmpty(PresumptiveIncome44ADa) Or PresumptiveIncome44ADa <> 0 Then
    PersumptiveInc44ADA("TotPersumptiveInc44ADA") = PresumptiveIncome44ADa
Else
    PersumptiveInc44ADA("TotPersumptiveInc44ADA") = 0
End If
jsonDictionary.add "PersumptiveInc44ADA", PersumptiveInc44ADA


If (Not IsEmpty(NOBCode2)) And UBound(NOBCode2) > 0 Then
    For i = 1 To UBound(NOBName2)
        If NOBName2(i) <> "" Then
            NatOfBus44AE1("NameOfBusiness") = (NOBName2(i))
        End If
        If NOBCode2(i) <> "" Then
            NatOfBus44AE1("CodeAE") = (NOBCode2(i))
        End If
        If NOBDescription2(i) <> "" Then
            NatOfBus44AE1("Description") = NOBDescription2(i)
        Else
        End If
        NatOfBus44AE.add NatOfBus44AE1
        Set NatOfBus44AE1 = Nothing
        Set NatOfBus44AE1 = CreateObject("Scripting.Dictionary")
    Next
End If
jsonDictionary.add "NatOfBus44AE", NatOfBus44AE
   

If UBound(SEC44AEMonth) > 0 Then
   For i = 1 To UBound(SEC44AEMonth)
       GoodsDtlsUs44AE1("RegNumberGoodsCarriage") = Trim(SEC44AERegNo(i))
       GoodsDtlsUs44AE1("OwnedLeasedHiredFlag") = SEC44AEOLH(i)
       GoodsDtlsUs44AE1("TonnageCapacity") = SEC44AETC(i)
       GoodsDtlsUs44AE1("HoldingPeriod") = SEC44AEMonth(i)
       GoodsDtlsUs44AE1("PresumptiveIncome") = SEC44AEPIncome(i)
       GoodsDtlsUs44AE.add GoodsDtlsUs44AE1
       Set GoodsDtlsUs44AE1 = Nothing
       Set GoodsDtlsUs44AE1 = CreateObject("Scripting.Dictionary")
   Next
End If
jsonDictionary.add "GoodsDtlsUs44AE", GoodsDtlsUs44AE

If PresumptiveIncomeHV <> "" Or Not IsEmpty(PresumptiveIncomeHV) Or PresumptiveIncomeHV <> 0 Then
    PersumptiveInc44AE("TotPersumInc44AE") = PresumptiveIncomeHV
Else
    PersumptiveInc44AE("TotPersumInc44AE") = 0
End If
If SalaryInterestPaid44AD <> "" Or Not IsEmpty(SalaryInterestPaid44AD) Or SalaryInterestPaid44AD <> 0 Then
    PersumptiveInc44AE("SalInterestByFirm") = SalaryInterestPaid44AD
Else
    PersumptiveInc44AE("SalInterestByFirm") = 0
End If
If IncomeChargeabeBPUs44AE <> "" Or Not IsEmpty(IncomeChargeabeBPUs44AE) Or IncomeChargeabeBPUs44AE <> 0 Then
    PersumptiveInc44AE("TotalPersumptiveInc") = IncomeChargeabeBPUs44AE
Else
    PersumptiveInc44AE("TotalPersumptiveInc") = 0

End If
If IncomeChargeabeBPUs <> 0 Or IncomeChargeabeBPUs <> "" Or Not IsEmpty(IncomeChargeabeBPUs) Then
    PersumptiveInc44AE("IncChargeableUnderBus") = IncomeChargeabeBPUs
Else
    PersumptiveInc44AE("IncChargeableUnderBus") = 0
End If
jsonDictionary.add "PersumptiveInc44AE", PersumptiveInc44AE

UpdateProgressBar
For i = 1 To UBound(GSTIN_BP)
    If GSTIN_BP(i) <> "" Or Not IsEmpty(GSTIN_BP(i)) Then
        TurnoverGrsRcptForGSTIN1("GSTINNo") = Trim(GSTIN_BP(i))
    End If
    If GSTINAMOUNT_BP(i) <> "" Or Not IsEmpty(GSTINAMOUNT_BP(i)) Then
        TurnoverGrsRcptForGSTIN1("AmtTurnGrossRcptGSTIN") = GSTINAMOUNT_BP(i)
    End If
    TurnoverGrsRcptForGSTIN.add TurnoverGrsRcptForGSTIN1
    Set TurnoverGrsRcptForGSTIN1 = Nothing
    Set TurnoverGrsRcptForGSTIN1 = CreateObject("Scripting.Dictionary")
Next
jsonDictionary.add "TurnoverGrsRcptForGSTIN", TurnoverGrsRcptForGSTIN
jsonDictionary("TotalTurnoverGrsRcptGSTIN") = Sheet3.Range("TotalGST").value
      
If PartnerMemberOwnCapital <> 0 Or PartnerMemberOwnCapital <> "" Or Not IsEmpty(PartnerMemberOwnCapital) Then
    FinanclPartclrOfBusiness("PartnerMemberOwnCapital") = PartnerMemberOwnCapital
Else
    FinanclPartclrOfBusiness("PartnerMemberOwnCapital") = 0
End If

UpdateProgressBar
If TotSecuredLoans <> 0 Or TotSecuredLoans <> "" Or Not IsEmpty(TotSecuredLoans) Then
    FinanclPartclrOfBusiness("SecuredLoans") = TotSecuredLoans
Else
    FinanclPartclrOfBusiness("SecuredLoans") = 0
End If
If TotUnSecuredLoans <> 0 Or TotUnSecuredLoans <> "" Or Not IsEmpty(TotUnSecuredLoans) Then
    FinanclPartclrOfBusiness("UnSecuredLoans") = TotUnSecuredLoans
Else
    FinanclPartclrOfBusiness("UnSecuredLoans") = 0
End If

UpdateProgressBar
If TotAdvances <> 0 Or TotAdvances <> "" Or Not IsEmpty(TotAdvances) Then
    FinanclPartclrOfBusiness("Advances") = TotAdvances
Else
    FinanclPartclrOfBusiness("Advances") = 0
End If
If TotSundryCreditors <> 0 Or TotSundryCreditors <> "" Or Not IsEmpty(TotSundryCreditors) Then
    FinanclPartclrOfBusiness("SundryCreditors") = TotSundryCreditors
Else
    FinanclPartclrOfBusiness("SundryCreditors") = 0
End If
If TotOthrCurrLiab <> 0 Or TotOthrCurrLiab <> "" Or Not IsEmpty(TotOthrCurrLiab) Then
    FinanclPartclrOfBusiness("OthrCurrLiab") = TotOthrCurrLiab
Else
    FinanclPartclrOfBusiness("OthrCurrLiab") = 0
End If

UpdateProgressBar
If TotalBPE15 <> 0 Or TotalBPE15 <> "" Or Not IsEmpty(TotalBPE15) Then
    FinanclPartclrOfBusiness("TotCapLiabilities") = TotalBPE15
Else
    FinanclPartclrOfBusiness("TotCapLiabilities") = 0
End If
If FixedAssets <> 0 Or FixedAssets <> "" Or Not IsEmpty(FixedAssets) Then
    FinanclPartclrOfBusiness("FixedAssets") = FixedAssets
Else
    FinanclPartclrOfBusiness("FixedAssets") = 0
End If
'Konda updated on 28-01-2026
If Sheet3.Range("BP_E18a").value <> 0 Or Sheet3.Range("BP_E18a").value <> "" Or Not IsEmpty(Sheet3.Range("BP_E18a").value) Then
    FinanclPartclrOfBusiness("Investments") = CDbl(Sheet3.Range("BP_E18a").value)
Else
    FinanclPartclrOfBusiness("Investments") = 0
End If
'=========================================
If Inventories <> 0 Or Inventories <> "" Or Not IsEmpty(Inventories) Then
    FinanclPartclrOfBusiness("Inventories") = Inventories
Else
    FinanclPartclrOfBusiness("Inventories") = 0
End If

UpdateProgressBar
If SundryDebtors <> 0 Or SundryDebtors <> "" Or Not IsEmpty(SundryDebtors) Then
    FinanclPartclrOfBusiness("SundryDebtors") = SundryDebtors
Else
    FinanclPartclrOfBusiness("SundryDebtors") = 0
End If
If BalWithBanks <> 0 Or BalWithBanks <> "" Or Not IsEmpty(BalWithBanks) Then
    FinanclPartclrOfBusiness("BalWithBanks") = BalWithBanks
Else
    FinanclPartclrOfBusiness("BalWithBanks") = 0
End If
If TotCashInHand <> 0 Or TotCashInHand <> "" Or Not IsEmpty(TotCashInHand) Then
    FinanclPartclrOfBusiness("CashInHand") = TotCashInHand
Else
    FinanclPartclrOfBusiness("CashInHand") = 0
End If

UpdateProgressBar
If LoansAndAdvances <> 0 Or LoansAndAdvances <> "" Or Not IsEmpty(LoansAndAdvances) Then
    FinanclPartclrOfBusiness("LoansAndAdvances") = LoansAndAdvances
Else
    FinanclPartclrOfBusiness("LoansAndAdvances") = 0
End If
If OtherAssets <> 0 Or OtherAssets <> "" Or Not IsEmpty(OtherAssets) Then
    FinanclPartclrOfBusiness("OtherAssets") = OtherAssets
Else
    FinanclPartclrOfBusiness("OtherAssets") = 0
End If
If TotalBPE23 <> 0 Or TotalBPE23 <> "" Or Not IsEmpty(TotalBPE23) Then
    FinanclPartclrOfBusiness("TotalAssets") = TotalBPE23
Else
    FinanclPartclrOfBusiness("TotalAssets") = 0
End If
jsonDictionary.add "FinanclPartclrOfBusiness", FinanclPartclrOfBusiness

Set ScheduleBP = jsonDictionary
End Function
Function ScheduleIT() As Object
subProcCaption = "IT"
noOfProcessSub = UBound(BSRCode_IT)
Dim i, j As Long
Dim jsonDictionary, Tax, TaxPayment, TotalTaxPayments

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Tax = CreateObject("Scripting.Dictionary")
Set TaxPayment = New Collection

TotalTaxPayments = Sheet18.Range("TaxP.Sum").value
jsonDictionary("TotalTaxPayments") = TotalTaxPayments

If Not IsEmpty(BSRCode_IT) And UBound(BSRCode_IT) > 0 Then

For i = 1 To UBound(BSRCode_IT)

    If BSRCode_IT(i) <> "" Then
        Tax("BSRCode") = UCase(BSRCode_IT(i))
    End If

    If DateDep_IT(i) <> "" Then
        Tax("DateDep") = UCase(DateDep_IT(i))
    End If

    If SrlNoOfChaln_IT(i) <> "" Then
     Tax("SrlNoOfChaln") = UVCase(CDbl(SrlNoOfChaln_IT(i)))
    End If

    If Amt_IT(i) <> "" Then
        Tax("Amt") = UVCase(CDbl(Amt_IT(i)))
     Else
        Tax("Amt") = 0
    End If
    
    TaxPayment.add Tax
    Set Tax = Nothing
    Set Tax = CreateObject("Scripting.Dictionary")
    
    Next
    jsonDictionary.add "TaxPayment", TaxPayment
End If


Set ScheduleIT = jsonDictionary
End Function

Function ScheduleTCS() As Object

subProcCaption = "TCS"
noOfProcessSub = UBound(TAN_TCS)

Dim i, j As Long
Dim jsonDictionary, TCSObj, EmployerOrDeductorOrCollectDetl
Dim TotalSchTCS, TCS, TotalTCS, TAN, EmployerOrDeductorOrCollecterName, Amtfrom26AS, AmtTCSClaimedThisYear As Variant


Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TCS = New Collection
Set TCSObj = CreateObject("Scripting.Dictionary")
Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")

jsonDictionary("TotalSchTCS") = Sheet17.Range("TCS.Sum").value
If UBound(TAN_TCS) > 0 Then
For i = 1 To UBound(TAN_TCS)
    If Not IsEmpty(TAN_TCS) And UBound(TAN_TCS) > 0 Then
        If TAN_TCS(i) <> "" Then
            EmployerOrDeductorOrCollectDetl("TAN") = TAN_TCS(i)
        End If
        If Not IsEmpty(EmployerOrDeductorOrCollecterName_TCS(i)) Then
            EmployerOrDeductorOrCollectDetl("EmployerOrDeductorOrCollecterName") = EmployerOrDeductorOrCollecterName_TCS(i)
        End If
        TCSObj.add "EmployerOrDeductorOrCollectDetl", EmployerOrDeductorOrCollectDetl
 Set EmployerOrDeductorOrCollectDetl = Nothing
    Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")
        If Not IsEmpty(AmtPaid_TCS(i)) Then
            TCSObj("Amtfrom26AS") = UVCase(CDbl(AmtPaid_TCS(i)))
        End If
        If Not IsEmpty(BroughtFwdTCSAmt_TCS(i)) Then
             TCSObj("TotalTCS") = UVCase(CDbl(BroughtFwdTCSAmt_TCS(i)))
        End If
        If Not IsEmpty(AmtClaimedOnOwnHands_TCS(i)) Then
            TCSObj("AmtTCSClaimedThisYear") = UVCase(CDbl(AmtClaimedOnOwnHands_TCS(i)))
        End If
        TCS.add TCSObj
        Set TCSObj = Nothing
        Set TCSObj = CreateObject("Scripting.Dictionary")
    End If
Next
jsonDictionary.add "TCS", TCS
End If
UpdateProgressBar
Set ScheduleTCS = jsonDictionary
End Function

Function TDSonSalaries() As Object

subProcCaption = "TDS1"
noOfProcessSub = UBound(TAN_TDS1)

Dim i As Long
Dim jsonDictionary, TDSonSal, EmployerOrDeductorOrCollectDetl
Dim TotalTDSonSalaries, TDSonSalary, IncChrgSal, TotalTDSSal As Variant
Set TDSonSalary = New Collection

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TDSonSal = CreateObject("Scripting.Dictionary")
Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")

If Not IsEmpty(TAN_TDS1) And UBound(TAN_TDS1) > 0 Then
    For i = 1 To UBound(TAN_TDS1)
    
    If TAN_TDS1(i) <> "" Then
        EmployerOrDeductorOrCollectDetl("TAN") = UCase(TAN_TDS1(i))
    Else
    
    End If
    
    If EmployerOrDeductorOrCollecterName_TDS1(i) <> "" Then
        EmployerOrDeductorOrCollectDetl("EmployerOrDeductorOrCollecterName") = UCase(EmployerOrDeductorOrCollecterName_TDS1(i))
    Else
    End If
    
    TDSonSal.add "EmployerOrDeductorOrCollectDetl", EmployerOrDeductorOrCollectDetl
   Set EmployerOrDeductorOrCollectDetl = Nothing
    Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")
    
    If AmtClaimedOnOwnHands_TDS1(i) <> "" Then
        TDSonSal("IncChrgSal") = UVCase(CDbl(AmtClaimedOnOwnHands_TDS1(i)))
    Else
        TDSonSal("IncChrgSal") = 0
    End If
    
    If AmtCarriedFwd_TDS1(i) <> "" Then
        TDSonSal("TotalTDSSal") = UVCase(CDbl(AmtCarriedFwd_TDS1(i)))
    Else
        TDSonSal("TotalTDSSal") = 0
    End If
    
    TDSonSalary.add TDSonSal
    Set TDSonSal = Nothing
    Set TDSonSal = CreateObject("Scripting.Dictionary")
    UpdateProgressBar
    
    Next

jsonDictionary.add "TDSonSalary", TDSonSalary
jsonDictionary("TotalTDSonSalaries") = Sheet4.Range("TDSal.Sum").value

Set TDSonSalaries = jsonDictionary
End If
End Function


Function TDSonOthThanSals() As Object

subProcCaption = "TDS2i"
noOfProcessSub = UBound(TAN_TDS2_1)

Dim i As Long
Dim jsonDictionary, TDSonOthThanSalDtl, TDSonOthThanSalDtls
Dim TDSonSalary, IncChrgSal, TotalTDSSal As Variant

Set TDSonOthThanSalDtls = New Collection
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TDSonOthThanSalDtl = CreateObject("Scripting.Dictionary")


If Not IsEmpty(TAN_TDS2_1) And UBound(TAN_TDS2_1) > 0 Then

    For i = 1 To UBound(TAN_TDS2_1)

        If TAN_TDS2_1(i) <> "" Then
            TDSonOthThanSalDtl("TANOfDeductor") = UCase(TAN_TDS2_1(i))
        Else
            TDSonOthThanSalDtl("TANOfDeductor") = 0
        End If

        
'TDS_E2_AY_2023_24 OLd-------------------------------------------------------

'        If Not isdropdownblank(FinYear_TDS2_1(i)) Then
'            TDSonOthThanSalDtl("DeductedYr") = UCase(FinYear_TDS2_1(i))
'        End If

'TDS_E2_AY_2023_24 OLd-------------------------------------------------------



'TDS_C2_AY_2024_25 Malli Change Starts------------------------------------------

        If Not isdropdownblank(FinYear_TDS2_1(i)) Then
            TDSonOthThanSalDtl("DeductedYr") = UCase(Left(FinYear_TDS2_1(i), 4))
        End If
        
'TDS_C2_AY_2024_25 Malli Change Ends------------------------------------------

        If Tds_TDS2_1(i) <> "" Then
            TDSonOthThanSalDtl("BroughtFwdTDSAmt") = UVCase(CDbl(Tds_TDS2_1(i)))
        Else
            TDSonOthThanSalDtl("BroughtFwdTDSAmt") = 0
        End If
        
        If Tdsdeducted_TDS2_1(i) <> "" Then
            TDSonOthThanSalDtl("TDSDeducted") = UVCase(CDbl(Tdsdeducted_TDS2_1(i)))
        Else
            TDSonOthThanSalDtl("TDSDeducted") = 0
        End If
        
        
        'Konda---------AY_2025-26_21/04/2025
        
                Dim TDSSection_tds2_enum As Variant
                
                If SectionTDS_TDS2_1(i) <> "" Then
                
                If (SectionTDS_TDS2_1(i)) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
                TDSSection_tds2_enum = "92A"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "192-Salary-Payment to employees other than Government employees" Then
                TDSSection_tds2_enum = "92B"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "192-Salary-Payment to Indian Government employees" Then
                TDSSection_tds2_enum = "92C"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "192A-TDS on PF withdrawal" Then
                TDSSection_tds2_enum = "192A"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "193-Interest on Securities" Then
                 TDSSection_tds2_enum = "193"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194-Dividends" Then
                 TDSSection_tds2_enum = "194"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194A-Interest other than 'Interest on securities'" Then
                TDSSection_tds2_enum = "94A"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194B-Winning from lottery or crossword puzzle" Then
                TDSSection_tds2_enum = "94B"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194BA-Winnings from online games" Then
                TDSSection_tds2_enum = "94BA"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194BB-Winning from horse race" Then
                TDSSection_tds2_enum = "4BB"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194C-Payments to contractors and sub-contractors" Then
               TDSSection_tds2_enum = "94C"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194D-Insurance commission" Then
               TDSSection_tds2_enum = "94D"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194DA-Payment in respect of life insurance policy" Then
               TDSSection_tds2_enum = "4DA"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194E-Payments to non-resident sportsmen or sports associations" Then
                TDSSection_tds2_enum = "94E"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194EE-Payments in respect of deposits under National Savings" Then
                TDSSection_tds2_enum = "4EE"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India" Then
                TDSSection_tds2_enum = "4F"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194G-Commission, price, etc. on sale of lottery tickets" Then
                 TDSSection_tds2_enum = "4G"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194H-Commission or brokerage" Then
                TDSSection_tds2_enum = "4H"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194I(a)-Rent on hiring of plant and machinery" Then
                 TDSSection_tds2_enum = "4-IA"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194I(b)-Rent on other than plant and machinery" Then
                 TDSSection_tds2_enum = "4-IB"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194IA-TDS on Sale of immovable property" Then
                TDSSection_tds2_enum = "4IA"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
                 TDSSection_tds2_enum = "4IB"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194IC-Payment under specified agreement" Then
                 TDSSection_tds2_enum = "4IC"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194J(a)-Fees for technical services" Then
                 TDSSection_tds2_enum = "94J-A"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194J(b)-Fees for professional  services or royalty etc" Then
                 TDSSection_tds2_enum = "94J-B"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
                TDSSection_tds2_enum = "94K"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LA-Payment of compensation on acquisition of certain immovable" Then
                TDSSection_tds2_enum = "4LA"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LB-Income by way of Interest from Infrastructure Debt fund" Then
                 TDSSection_tds2_enum = "4LB"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC" Then
                 TDSSection_tds2_enum = "4LC1"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC" Then
                TDSSection_tds2_enum = "4LC2"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC" Then
                TDSSection_tds2_enum = "4LC3"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
                 TDSSection_tds2_enum = "4BA1"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
                TDSSection_tds2_enum = "4BA2"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR" Then
                TDSSection_tds2_enum = "LBA1"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR" Then
                TDSSection_tds2_enum = "LBA2"
                
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR" Then
                TDSSection_tds2_enum = "LBA3"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBB-Income in respect of units of investment fund" Then
               TDSSection_tds2_enum = "LBB"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194R-Benefits or perquisites of business or profession" Then
                 TDSSection_tds2_enum = "94R"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons" Then
                TDSSection_tds2_enum = "94S"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released" Then
               TDSSection_tds2_enum = "94B-P"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
                TDSSection_tds2_enum = "94R-P"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released" Then
               TDSSection_tds2_enum = "94S-P"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LBC-Income in respect of investment in securitization trust" Then
                TDSSection_tds2_enum = "LBC"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194LD-TDS on interest on bonds / government securities" Then
                TDSSection_tds2_enum = "4LD"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194M-Payment of certain sums by certain individuals or HUF" Then
                TDSSection_tds2_enum = "94M"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
                 TDSSection_tds2_enum = "94N"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194N-First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
                 TDSSection_tds2_enum = "94N-F"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194N-Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
                TDSSection_tds2_enum = "94N-C"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies" Then
               TDSSection_tds2_enum = "94N-FT"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
                 TDSSection_tds2_enum = "94O"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194P-Deduction of tax in case of specified senior citizen" Then
               TDSSection_tds2_enum = "94P"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
                TDSSection_tds2_enum = "94Q"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "195-Other sums payable to a non-resident" Then
                TDSSection_tds2_enum = "195"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "196A-Income in respect of units of non-residents" Then
                TDSSection_tds2_enum = "96A"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "196B-Payments in respect of units to an offshore fund" Then
                 TDSSection_tds2_enum = "96B"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "196C-Income from foreign currency bonds or shares of Indian" Then
                 TDSSection_tds2_enum = "96C"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "196D-Income of foreign institutional investors from securities" Then
                 TDSSection_tds2_enum = "96D"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "196D(1A)-Income of specified fund from securities" Then
                 TDSSection_tds2_enum = "96DA"
                
                ElseIf (SectionTDS_TDS2_1(i)) = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released" Then
                 TDSSection_tds2_enum = "94BA-P"
                 
               
                
                Else
                 TDSSection_tds2_enum = ""
                End If
                
                If TDSSection_tds2_enum <> "" Then
                TDSonOthThanSalDtl("TDSSection") = UCase(Trim(TDSSection_tds2_enum))
                End If
                
        End If
        '-----------------------------------
        
        
        
        
        'added by Chetan C M for AY 2025-26 on 31/01/2025
'--start--'

'        Dim SecTDSDeducted As Variant
'       ' Dim SectionTDS_TDS2_1
'        If SectionTDS_TDS2_1(i) <> "" Then
'              If (SectionTDS_TDS2_1(i)) = "192A-TDS on PF withdrawal" Then
'              SectionTDS_TDS2_1(i) = "192A"
'              ElseIf (SectionTDS_TDS2_1(i)) = "193-Interest on Securities" Then
'              SectionTDS_TDS2_1(i) = "193"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194-Dividends" Then
'              SectionTDS_TDS2_1(i) = "194"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194A-Interest other than 'Interest on securities'" Then
'              SectionTDS_TDS2_1(i) = "94A"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194C-Payments to contractors and sub-contractors" Then
'              SectionTDS_TDS2_1(i) = "94C"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194DA-Payment in respect of life insurance policy" Then
'              SectionTDS_TDS2_1(i) = "4DA"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194EE-Payments in respect of deposits under National Savings " Then
'              SectionTDS_TDS2_1(i) = "4EE"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194H-Commission or brokerage" Then
'              SectionTDS_TDS2_1(i) = "4H"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194I(a)-Rent on hiring of plant and machinery" Then
'              SectionTDS_TDS2_1(i) = "4-IA"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194I(b)-Rent on other than plant and machinery" Then
'              SectionTDS_TDS2_1(i) = "4-IB"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194IB-Payment of rent by certain individuals or Hindu undivided " Then
'              SectionTDS_TDS2_1(i) = "4IB"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194J(a)-Fees for technical services" Then
'              SectionTDS_TDS2_1(i) = "94J-A"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194J(b)-Fees for professional  services or royalty etc" Then
'              SectionTDS_TDS2_1(i) = "94J-B"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
'              SectionTDS_TDS2_1(i) = "94K"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194LA-Payment of compensation on acquisition of certain immovable " Then
'              SectionTDS_TDS2_1(i) = "4LA"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
'              SectionTDS_TDS2_1(i) = "4BA1"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
'              SectionTDS_TDS2_1(i) = "4BA2"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194LBB-Income in respect of units of investment fund" Then
'              SectionTDS_TDS2_1(i) = "LBB"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194R-Benefits or perquisites of business or profession" Then
'              SectionTDS_TDS2_1(i) = "94R"
'              ElseIf (SectionTDS_TDS2_1(i)) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
'              SectionTDS_TDS2_1(i) = "94R-P"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194M-Payment of certain sums by certain individuals or HUF" Then
'              SectionTDS_TDS2_1(i) = "94M"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
'              SectionTDS_TDS2_1(i) = "94N"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
'              SectionTDS_TDS2_1(i) = "94N-F"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
'              SectionTDS_TDS2_1(i) = "94N-C"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
'              SectionTDS_TDS2_1(i) = "94O"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194P-Deduction of tax in case of specified senior citizen" Then
'              SectionTDS_TDS2_1(i) = "94P"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194D-Insurance commission" Then
'              SectionTDS_TDS2_1(i) = "94D"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194G-Commission, price, etc. on sale of lottery tickets" Then
'              SectionTDS_TDS2_1(i) = "4G"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194LBC-Income in respect of investment in securitization trust" Then
'              SectionTDS_TDS2_1(i) = "LBC"
'              ElseIf (SectionTDS_TDS2_1(i)) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
'              SectionTDS_TDS2_1(i) = "94Q"
'              Else
'              SectionTDS_TDS2_1(i) = ""
'              End If
'            TDSonOthThanSalDtl("SecTDSDeducted") = SectionTDS_TDS2_1(i)
'End If
'--end--'
        
        If TdsClaimed_TDS2_1(i) <> "" Then
            TDSonOthThanSalDtl("TDSClaimed") = UVCase(CDbl(TdsClaimed_TDS2_1(i)))
        Else
            TDSonOthThanSalDtl("TDSClaimed") = 0
        End If
        
        If GrossAmount_TDS2_1(i) <> "" Then
            TDSonOthThanSalDtl("GrossAmount") = UVCase(CDbl(GrossAmount_TDS2_1(i)))
        Else
            TDSonOthThanSalDtl("GrossAmount") = 0
        End If
        
        If Headincome_TDS2_1(i) <> "" Then
            If UCase(Headincome_TDS2_1(i)) = "INCOME FROM OTHER SOURCE" Then
              Headincome_TDS2_1(i) = "OS"
            ElseIf UCase(Headincome_TDS2_1(i)) = "INCOME FROM BUSINESS & PROFESSION" Then
              Headincome_TDS2_1(i) = "BP"
            ElseIf UCase(Headincome_TDS2_1(i)) = "INCOME FROM HOUSE PROPERTY" Then
              Headincome_TDS2_1(i) = "HP"
            ElseIf UCase(Headincome_TDS2_1(i)) = "EXEMPT INCOME" Then
              Headincome_TDS2_1(i) = "EI"
            ElseIf UCase(Headincome_TDS2_1(i)) = "NOT APPLICABLE (ONLY IN CASE TDS IS DEDUCTED U/S 194N)" Then
              Headincome_TDS2_1(i) = "NA"
            End If
            TDSonOthThanSalDtl("HeadOfIncome") = UCase(Headincome_TDS2_1(i))
        End If
        
'Change-05.03.2023.101 VF Issue
        If Headincome_TDS2_1(i) = "" Then
            TDSonOthThanSalDtl("HeadOfIncome") = "NA"
        End If
'---

        If Tdscarried_TDS2_1(i) <> "" Then
            TDSonOthThanSalDtl("TDSCreditCarriedFwd") = UVCase(CDbl(Tdscarried_TDS2_1(i)))
        End If
        
        TDSonOthThanSalDtls.add TDSonOthThanSalDtl
        Set TDSonOthThanSalDtl = Nothing
        Set TDSonOthThanSalDtl = CreateObject("Scripting.Dictionary")
        UpdateProgressBar
    Next
    
   jsonDictionary.add "TDSonOthThanSalDtls", TDSonOthThanSalDtls
End If

jsonDictionary("TotalTDSonOthThanSals") = UVCase(CDbl(Sheet4.Range("TDS2i.Sum").value))

Set TDSonOthThanSals = jsonDictionary
End Function

Function ScheduleTDS3Dtls() As Object

subProcCaption = "TDS2ii"
noOfProcessSub = UBound(PAN_TDS2ii_1)

Dim i As Long
Dim jsonDictionary, TDS3Dtl, TDS3Details

Set TDS3Details = New Collection

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TDS3Dtl = CreateObject("Scripting.Dictionary")

If Not IsEmpty(PAN_TDS2ii_1) And UBound(PAN_TDS2ii_1) > 0 Then

    For i = 1 To UBound(PAN_TDS2ii_1)
    
        If PAN_TDS2ii_1(i) <> "" Then
            TDS3Dtl("PANofTenant") = UCase(PAN_TDS2ii_1(i))
        End If
        
        If Aadhaar_TDS2ii_1(i) <> "" Then
            TDS3Dtl("AadhaarofTenant") = UCase(Aadhaar_TDS2ii_1(i))
        End If
'-------------------------------Malli-------------Start Changes
'TDS_E5_AY_2023_24 OLD
'If Not isdropdownblank(FinYear_TDS2ii_1(i)) Then
'            TDS3Dtl("DeductedYr") = UCase(FinYear_TDS2ii_1(i))
'        End If
        
'TDS_C5 AY_2024_25 Malli Change Starts----------------------------------------------------

        If Not isdropdownblank(FinYear_TDS2ii_1(i)) Then
            TDS3Dtl("DeductedYr") = UCase(Left(FinYear_TDS2ii_1(i), 4))
        End If
        
'TDS_C5 AY_2024_25 Change Ends------------------------------------------------------

        If Tds_TDS2ii_1(i) <> "" Then
            TDS3Dtl("BroughtFwdTDSAmt") = UVCase(CDbl(Tds_TDS2ii_1(i)))
        End If
         
        If Tdsdeducted_TDS2ii_1(i) <> "" Then
            TDS3Dtl("TDSDeducted") = UVCase(CDbl(Tdsdeducted_TDS2ii_1(i)))
        Else
            TDS3Dtl("TDSDeducted") = 0
        End If
        
'Konda----AY_2025_26---------21/04/2025---
            
        Dim TDSSection_tds3_enum As Variant
                
                If SectionTDS_TDS2_2ii(i) <> "" Then
                
                If (SectionTDS_TDS2_2ii(i)) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
                TDSSection_tds3_enum = "92A"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "192-Salary-Payment to employees other than Government employees" Then
                TDSSection_tds3_enum = "92B"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "192-Salary-Payment to Indian Government employees" Then
                TDSSection_tds3_enum = "92C"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "192A-TDS on PF withdrawal" Then
                TDSSection_tds3_enum = "192A"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "193-Interest on Securities" Then
                 TDSSection_tds3_enum = "193"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194-Dividends" Then
                 TDSSection_tds3_enum = "194"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194A-Interest other than 'Interest on securities'" Then
                TDSSection_tds3_enum = "94A"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194B-Winning from lottery or crossword puzzle" Then
                TDSSection_tds3_enum = "94B"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194BA-Winnings from online games" Then
                TDSSection_tds3_enum = "94BA"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194BB-Winning from horse race" Then
                TDSSection_tds3_enum = "4BB"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194C-Payments to contractors and sub-contractors" Then
               TDSSection_tds3_enum = "94C"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194D-Insurance commission" Then
               TDSSection_tds3_enum = "94D"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194DA-Payment in respect of life insurance policy" Then
               TDSSection_tds3_enum = "4DA"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194E-Payments to non-resident sportsmen or sports associations" Then
                TDSSection_tds3_enum = "94E"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194EE-Payments in respect of deposits under National Savings" Then
                TDSSection_tds3_enum = "4EE"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India" Then
                TDSSection_tds3_enum = "4F"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194G-Commission, price, etc. on sale of lottery tickets" Then
                 TDSSection_tds3_enum = "4G"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194H-Commission or brokerage" Then
                TDSSection_tds3_enum = "4H"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194I(a)-Rent on hiring of plant and machinery" Then
                 TDSSection_tds3_enum = "4-IA"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194I(b)-Rent on other than plant and machinery" Then
                 TDSSection_tds3_enum = "4-IB"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194IA-TDS on Sale of immovable property" Then
                TDSSection_tds3_enum = "4IA"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
                 TDSSection_tds3_enum = "4IB"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194IC-Payment under specified agreement" Then
                 TDSSection_tds3_enum = "4IC"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194J(a)-Fees for technical services" Then
                 TDSSection_tds3_enum = "94J-A"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194J(b)-Fees for professional  services or royalty etc" Then
                 TDSSection_tds3_enum = "94J-B"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
                TDSSection_tds3_enum = "94K"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LA-Payment of compensation on acquisition of certain immovable" Then
                TDSSection_tds3_enum = "4LA"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LB-Income by way of Interest from Infrastructure Debt fund" Then
                 TDSSection_tds3_enum = "4LB"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC" Then
                 TDSSection_tds3_enum = "4LC1"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC" Then
                TDSSection_tds3_enum = "4LC2"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC" Then
                TDSSection_tds3_enum = "4LC3"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
                 TDSSection_tds3_enum = "4BA1"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
                TDSSection_tds3_enum = "4BA2"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR" Then
                TDSSection_tds3_enum = "LBA1"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR" Then
                TDSSection_tds3_enum = "LBA2"
                
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR" Then
                TDSSection_tds3_enum = "LBA3"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBB-Income in respect of units of investment fund" Then
               TDSSection_tds3_enum = "LBB"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194R-Benefits or perquisites of business or profession" Then
                 TDSSection_tds3_enum = "94R"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons" Then
                TDSSection_tds3_enum = "94S"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released" Then
               TDSSection_tds3_enum = "94B-P"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
                TDSSection_tds3_enum = "94R-P"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released" Then
               TDSSection_tds3_enum = "94S-P"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBC-Income in respect of investment in securitization trust" Then
                TDSSection_tds3_enum = "LBC"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194LD-TDS on interest on bonds / government securities" Then
                TDSSection_tds3_enum = "4LD"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194M-Payment of certain sums by certain individuals or HUF" Then
                TDSSection_tds3_enum = "94M"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
                 TDSSection_tds3_enum = "94N"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194N-First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
                 TDSSection_tds3_enum = "94N-F"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194N-Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
                TDSSection_tds3_enum = "94N-C"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies" Then
               TDSSection_tds3_enum = "94N-FT"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
                 TDSSection_tds3_enum = "94O"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194P-Deduction of tax in case of specified senior citizen" Then
               TDSSection_tds3_enum = "94P"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
                TDSSection_tds3_enum = "94Q"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "195-Other sums payable to a non-resident" Then
                TDSSection_tds3_enum = "195"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "196A-Income in respect of units of non-residents" Then
                TDSSection_tds3_enum = "96A"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "196B-Payments in respect of units to an offshore fund" Then
                 TDSSection_tds3_enum = "96B"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "196C-Income from foreign currency bonds or shares of Indian" Then
                 TDSSection_tds3_enum = "96C"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "196D-Income of foreign institutional investors from securities" Then
                 TDSSection_tds3_enum = "96D"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "196D(1A)-Income of specified fund from securities" Then
                 TDSSection_tds3_enum = "96DA"
                
                ElseIf (SectionTDS_TDS2_2ii(i)) = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released" Then
                 TDSSection_tds3_enum = "94BA-P"
                
                Else
                 TDSSection_tds3_enum = ""
                End If
                
                If TDSSection_tds3_enum <> "" Then
                TDS3Dtl("TDSSection") = UCase(Trim(TDSSection_tds3_enum))
                End If
                
        End If


'------------------------------------------

'added by Chetan C M for AY 2025-26 on 31/01/2025
'start
' Dim SecTDSDeducted As Variant
'       ' Dim SectionTDS_TDS2_2ii
'If SectionTDS_TDS2_2ii(i) <> "" Then
'              If (SectionTDS_TDS2_2ii(i)) = "192A-TDS on PF withdrawal" Then
'              SectionTDS_TDS2_2ii(i) = "192A"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "193-Interest on Securities" Then
'              SectionTDS_TDS2_2ii(i) = "193"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194-Dividends" Then
'              SectionTDS_TDS2_2ii(i) = "194"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194A-Interest other than 'Interest on securities'" Then
'              SectionTDS_TDS2_2ii(i) = "94A"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194C-Payments to contractors and sub-contractors" Then
'              SectionTDS_TDS2_2ii(i) = "94C"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194DA-Payment in respect of life insurance policy" Then
'              SectionTDS_TDS2_2ii(i) = "4DA"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194EE-Payments in respect of deposits under National Savings " Then
'              SectionTDS_TDS2_2ii(i) = "4EE"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194H-Commission or brokerage" Then
'              SectionTDS_TDS2_2ii(i) = "4H"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194I(a)-Rent on hiring of plant and machinery" Then
'              SectionTDS_TDS2_2ii(i) = "4-IA"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194I(b)-Rent on other than plant and machinery" Then
'              SectionTDS_TDS2_2ii(i) = "4-IB"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194IB-Payment of rent by certain individuals or Hindu undivided " Then
'              SectionTDS_TDS2_2ii(i) = "4IB"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194J(a)-Fees for technical services" Then
'              SectionTDS_TDS2_2ii(i) = "94J-A"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194J(b)-Fees for professional  services or royalty etc" Then
'              SectionTDS_TDS2_2ii(i) = "94J-B"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
'              SectionTDS_TDS2_2ii(i) = "94K"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194LA-Payment of compensation on acquisition of certain immovable " Then
'              SectionTDS_TDS2_2ii(i) = "4LA"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
'              SectionTDS_TDS2_2ii(i) = "4BA1"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
'              SectionTDS_TDS2_2ii(i) = "4BA2"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBB-Income in respect of units of investment fund" Then
'              SectionTDS_TDS2_2ii(i) = "LBB"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194R-Benefits or perquisites of business or profession" Then
'              SectionTDS_TDS2_2ii(i) = "94R"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
'              SectionTDS_TDS2_2ii(i) = "94R-P"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194M-Payment of certain sums by certain individuals or HUF" Then
'              SectionTDS_TDS2_2ii(i) = "94M"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
'              SectionTDS_TDS2_2ii(i) = "94N"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
'              SectionTDS_TDS2_2ii(i) = "94N-F"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
'              SectionTDS_TDS2_2ii(i) = "94N-C"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
'              SectionTDS_TDS2_2ii(i) = "94O"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194P-Deduction of tax in case of specified senior citizen" Then
'              SectionTDS_TDS2_2ii(i) = "94P"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194D-Insurance commission" Then
'              SectionTDS_TDS2_2ii(i) = "94D"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194G-Commission, price, etc. on sale of lottery tickets" Then
'              SectionTDS_TDS2_2ii(i) = "4G"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194LBC-Income in respect of investment in securitization trust" Then
'              SectionTDS_TDS2_2ii(i) = "LBC"
'              ElseIf (SectionTDS_TDS2_2ii(i)) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
'              SectionTDS_TDS2_2ii(i) = "94Q"
'              Else
'              SectionTDS_TDS2_2ii(i) = ""
'              End If
'            TDS3Dtl("SecTDSDeducted") = SectionTDS_TDS2_2ii(i)
'End If
'end
        
        If TdsClaimed_TDS2ii_1(i) <> "" Then
            TDS3Dtl("TDSClaimed") = UVCase(CDbl(TdsClaimed_TDS2ii_1(i)))
        Else
            TDS3Dtl("TDSClaimed") = 0
        End If
        
        If GrossAmount_TDS2ii_1(i) <> "" Then
            TDS3Dtl("GrossAmount") = UVCase(CDbl(GrossAmount_TDS2ii_1(i)))
        Else
            TDS3Dtl("GrossAmount") = 0
        End If
         
        If Headincome_TDS2ii_1(i) <> "" Then
           If UCase(Headincome_TDS2ii_1(i)) = "INCOME FROM HOUSE PROPERTY" Then
               Headincome_TDS2ii_1(i) = "HP"
           ElseIf Headincome_TDS2ii_1(i) = "Income from Business & Profession" Then
               Headincome_TDS2ii_1(i) = "BP"
            ElseIf Headincome_TDS2ii_1(i) = "Income from other sources" Then
               Headincome_TDS2ii_1(i) = "OS"
            ElseIf Headincome_TDS2ii_1(i) = "Exempt Income" Then
                Headincome_TDS2ii_1(i) = "EI"
           
                
           End If
           TDS3Dtl("HeadOfIncome") = UCase(Headincome_TDS2ii_1(i))
        End If
        
        If Tdscarried_TDS2ii_1(i) <> "" Then
           TDS3Dtl("TDSCreditCarriedFwd") = UVCase(CDbl(Tdscarried_TDS2ii_1(i)))
        End If
        
        TDS3Details.add TDS3Dtl
        Set TDS3Dtl = Nothing
        Set TDS3Dtl = CreateObject("Scripting.Dictionary")
        
        UpdateProgressBar
    Next
    
   
    jsonDictionary.add "TDS3Details", TDS3Details
End If


 jsonDictionary("TotalTDS3Details") = UVCase(CDbl(Sheet4.Range("TDS2ii.Sum").value))
Set ScheduleTDS3Dtls = jsonDictionary

End Function
''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String
    Dim json_InnerIndentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ToJson = "null"
    Case VBA.vbDate
        ' Date
        json_DateStr = ConvertToIso(VBA.CDate(JsonValue))

        ToJson = """" & json_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ToJson = JsonValue
        Else
            ToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ToJson = "true"
        Else
            ToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
            End If
        End If

        ' Array
        json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength

        On Error Resume Next

        json_LBound = LBound(JsonValue, 1)
        json_UBound = UBound(JsonValue, 1)
        json_LBound2D = LBound(JsonValue, 2)
        json_UBound2D = UBound(JsonValue, 2)

        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_Index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    ' Append comma to previous line
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    ' 2D Array
                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If
                    json_BufferAppend json_Buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength

                    For json_Index2D = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                        End If

                        json_Converted = ToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2)

                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = "" Then
                            ' (nest to only check if converted = "")
                            If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                json_Converted = "null"
                            End If
                        End If

                        If json_PrettyPrint Then
                            json_Converted = vbNewLine & json_InnerIndentation & json_Converted
                        End If

                        json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                    Next json_Index2D

                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If

                    json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                    json_IsFirstItem2D = True
                Else
                    ' 1D Array
                    json_Converted = ToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1)

                    ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                    If json_Converted = "" Then
                        ' (nest to only check if converted = "")
                        If json_IsUndefined(JsonValue(json_Index)) Then
                            json_Converted = "null"
                        End If
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Index
        End If

        On Error GoTo 0

        If json_PrettyPrint Then
            json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
            Else
                json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If
        End If

        json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength

        ToJson = json_BufferToString(json_Buffer, json_BufferPosition)

    ' Dictionary or Collection
    Case VBA.vbObject
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
            End If
        End If

        ' Dictionary
        If VBA.TypeName(JsonValue) = "Dictionary" Then
            json_BufferAppend json_Buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In JsonValue.Keys
                ' For Objects, undefined (Empty/Nothing) is not added to object
                json_Converted = ToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                If json_Converted = "" Then
                    json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                Else
                    json_SkipItem = False
                End If

                If Not json_SkipItem Then
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                    Else
                        json_Converted = """" & json_Key & """:" & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Key

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength

        ' Collection
        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In JsonValue
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                json_Converted = ToJson(json_Value, Whitespace, json_CurrentIndentation + 1)

                ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                If json_Converted = "" Then
                    ' (nest to only check if converted = "")
                    If json_IsUndefined(json_Value) Then
                        json_Converted = "null"
                    End If
                End If

                If json_PrettyPrint Then
                    json_Converted = vbNewLine & json_Indentation & json_Converted
                End If

                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
            Next json_Value

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
        End If

        ToJson = json_BufferToString(json_Buffer, json_BufferPosition)
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ToJson = JsonValue
        On Error GoTo 0
    End Select
End Function
Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 30 And json_Length <= 100 Then
        Dim json_CharCode As String

        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function
Private Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_AscCode
        Case 34
            ' " -> 34 -> \"
            json_Char = "\"""
        Case 92
            ' \ -> 92 -> \\
            json_Char = "\\"
        Case 47
            ' / -> 47 -> \/ (optional)
            If JsonOptions.EscapeSolidus Then
                json_Char = "\/"
            End If
        Case 8
            ' backspace -> 8 -> \b
            json_Char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            json_Char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            json_Char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            json_Char = "\r"
        Case 9
            ' tab -> 9 -> \t
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            ' Non-ascii characters -> convert to 4-digit hex
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index

    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
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
Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
    ' Empty / Nothing -> undefined
    Select Case VBA.VarType(json_Value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_Value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function
Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

Function UVCase(valV1 As Variant) As Variant
    If (valV1 = "") Then
        valV1 = 0
    End If
        UVCase = valV1
End Function

'AY 2024-25 New Schedule Implementation

''Malli
'Function Schedule80GGC() As Object
'On Error Resume Next
'Dim i As Variant
'Dim jsonDictionary, GGCObj, GGCDtls
'
'Set jsonDictionary = CreateObject("Scripting.Dictionary")
'Set GGCDtls = CreateObject("Scripting.Dictionary")
'
'Set GGCObj = New Collection
'
'
'    For i = 1 To Sheet11.Range("SCH_80GGC_DonationDate").count
'
'        If Sheet11.Range("SCH_80GGC_DonationDate").Cells(i, 1).value <> "" Then
'
'
'            If Sheet11.Range("SCH_80GGC_DonationDate").Cells(i, 1).value <> "" Then
'               GGCDtls("DonationDate") = Mid((Sheet11.Range("SCH_80GGC_DonationDate").Cells(i, 1).value), 7, 4) & "-" & Mid((Sheet11.Range("SCH_80GGC_DonationDate").Cells(i, 1).value), 4, 2) & "-" & Mid((Sheet11.Range("SCH_80GGC_DonationDate").Cells(i, 1).value), 1, 2)
'            End If
'
'            If Sheet11.Range("SCH_80GGC_DonationAmtCash").Cells(i, 1).value <> "" Then
'                GGCDtls("DonationAmtCash") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_DonationAmtCash").Cells(i, 1).value))
'               Else:
'               GGCDtls("DonationAmtCash") = 0
'        End If
'
' If Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Cells(i, 1).value <> "" Then
'                GGCDtls("DonationAmtOtherMode") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Cells(i, 1).value))
'            Else:
'            GGCDtls("DonationAmtOtherMode") = 0
'            End If
'
'
' If Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").Cells(i, 1).value <> "" Then
'                GGCDtls("TransactionRefNumUPI") = UCase(Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").Cells(i, 1).value)
'            End If
'
'
'
' If Sheet11.Range("SCH_80GGC_IFSCCode").Cells(i, 1).value <> "" Then
'                GGCDtls("IFSCCode") = UCase(Sheet11.Range("SCH_80GGC_IFSCCode").Cells(i, 1).value)
'            End If
'
' If Sheet11.Range("SCH_80GGC_BankName").Cells(i, 1).value <> "" Then
'                GGCDtls("BankName") = UCase(Sheet11.Range("SCH_80GGC_BankName").Cells(i, 1).value)
'            End If
'
' If Sheet11.Range("SCH_80GGC_BankAccountNo").Cells(i, 1).value <> "" Then
'                GGCDtls("BankAccountNo") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_BankAccountNo").Cells(i, 1).value))
'            End If
'
' If Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Cells(i, 1).value <> "" Then
'                GGCDtls("ChnImpsNeftRtgsNum") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Cells(i, 1).value))
'            End If
'
'
' If Sheet11.Range("SCH_80GGC_DonationAmt").Cells(i, 1).value <> "" Then
'                GGCDtls("DonationAmt") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_DonationAmt").Cells(i, 1).value))
'                Else:
'                       GGCDtls("DonationAmt") = 0
'            End If
'
' If Sheet11.Range("SCH_80GGC_EligibleDonationAmt").Cells(i, 1).value <> "" Then
'                GGCDtls("EligibleDonationAmt") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_EligibleDonationAmt").Cells(i, 1).value))
'                Else:
'                GGCDtls("EligibleDonationAmt") = 0
'            End If
'
'
'
'            ' End Change.IDS.33
'            GGCObj.add GGCDtls
'
'            Set GGCDtls = Nothing
'            Set GGCDtls = CreateObject("Scripting.Dictionary")
'
'            UpdateProgressBar
'        End If
'    Next
'
'
'        jsonDictionary.add "Schedule80GGCDetails", GGCObj
'
'
'
' If Sheet11.Range("SCH_80GGC_TotalDonationAmtCash80GGC").value > 0 Then
'            jsonDictionary("TotalDonationAmtCash80GGC") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_TotalDonationAmtCash80GGC").value))
'        Else
'            jsonDictionary("TotalDonationAmtCash80GGC") = 0
'        End If
'
' If Sheet11.Range("SCH_80GGC_TotalDonationAmtOtherMode80GGC").value > 0 Then
'            jsonDictionary("TotalDonationAmtOtherMode80GGC") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_TotalDonationAmtOtherMode80GGC").value))
'        Else
'            jsonDictionary("TotalDonationAmtOtherMode80GGC") = 0
'        End If
'
' If Sheet11.Range("SCH_80GGC_TotalDonationsUs80GGC").value > 0 Then
'            jsonDictionary("TotalDonationsUs80GGC") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_TotalDonationsUs80GGC").value))
'        Else
'            jsonDictionary("TotalDonationsUs80GGC") = 0
'        End If
'
' If Sheet11.Range("SCH_80GGC_TotalEligibleDonationAmt80GGC").value > 0 Then
'            jsonDictionary("TotalEligibleDonationAmt80GGC") = UVCase(CDbl(Sheet11.Range("SCH_80GGC_TotalEligibleDonationAmt80GGC").value))
'        Else
'            jsonDictionary("TotalEligibleDonationAmt80GGC") = 0
'        End If
'
'
'
'Set Schedule80GGC = jsonDictionary
'End Function
''---end



Function FilingStatus() As Object

subProcCaption = "Filling Status"
noOfProcessSub = 7

Dim jsonDictionary, AssesseeRep, NewTaxRegime, ate10IE, acknowledgementNumber, date10IE_1, acknowledgementNumber_1, clauseiv7provisio139iType, clauseiv7provisio139iDtls
Dim AYOptedIn, NewForm10IEDate, NewForm10IEAckNo, OldTaxRegime, AYinOptOut, OldForm10IEDate, OldForm10IEAckNo

Dim date10IE_2, acknowledgementNumber_2

Set clauseiv7provisio139iDtls = New Collection
Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set AssesseeRep = CreateObject("Scripting.Dictionary")
 
'    NewTaxRegime = Sheet8.Range("oldbacValue").value
'    OptingNewTaxRegime = Sheet1.Range("OptingNewTaxRegime").value
'    date10IE = Sheet1.Range("Sheet1.date10IE").value
'    acknowledgementNumber = Sheet1.Range("Sheet1.acknowledgementNumber").value

    date10IE_1 = Sheet1.Range("Sheet1.Dateoffiling").value
    acknowledgementNumber_1 = Sheet1.Range("Sheet1.acknowledgementNumber_New").value
    
'commented by Chetan C M on 27/01/2025
'--start--
' 'Konda updated ------------
'    date10IE_2 = Sheet1.Range("BAC115.DateOfFiling").value
'    acknowledgementNumber_2 = Sheet1.Range("BAC115.AcknowledgementNo").value
 '--end--
 
'added by Chetan C M AY 2025-26 on 27/01/2025
 '--start
 Dim date10IE, date10IE_3, acknowledgementNumber_3
' date10IE = Sheet1.Range("BAC115.DateOfFiling").value
' acknowledgementNumber = Sheet1.Range("BAC115.AcknowledgementNo").value
'

' date10IE_1 = Sheet1.Range("BAC115.DateOfFiling_New_Yes").value
' acknowledgementNumber_1 = Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value
'
' date10IE_2 = Sheet1.Range("BAC115.DateOfFiling_New_No").value
' acknowledgementNumber_2 = Sheet1.Range("BAC115.AcknowledgementNo_New_No").value
'
' date10IE_3 = Sheet1.Range("BAC115.DateOfFiling_New_NA").value
' acknowledgementNumber_3 = Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value
 '--end

 
'Generate-08.12.2022.103.09A
'    AYOptedIn = Sheet1.Range("AYinOpting").value
'    NewForm10IEDate = Dformat(Sheet1.Range("Sheet1.date10IE_Opt").value, "")
'    NewForm10IEAckNo = Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value
'    OldTaxRegime = Sheet1.Range("Opt_Out115BAC").value
'    AYinOptOut = Sheet1.Range("AYinOptOut").value
'    OldForm10IEDate = Dformat(Sheet1.Range("Sheet1.date10IE_OptOut").value, "")
'    OldForm10IEAckNo = Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value
'---End Change
 
 
UpdateProgressBar

If ReturnFilledSection <> "" Then
    jsonDictionary("ReturnFileSec") = UVCase(CDbl(Trim(UCase(Mid(Trim(ReturnFilledSection), 1, InStr(ReturnFilledSection, "-") - 1)))))
End If
UpdateProgressBar
'----------------------------------------------------------

'commednted by Chetan C M on 27/01/2025
'--start
''Konda updated-----
'Dim TaxRegime_2 As Variant  'Sheet1.115BAC 'BAC115.DateOfFiling 'BAC115.AcknowledgementNo
'
'TaxRegime_2 = Sheet1.Range("Sheet1.115BAC").value
'
'If TaxRegime_2 <> "" Then
'            If TaxRegime_2 = "No" Then
'
'            jsonDictionary("OptOutNewTaxRegime_AY24-25") = "N"
'
'            ElseIf TaxRegime_2 = "Yes" Then
'            jsonDictionary("OptOutNewTaxRegime_AY24-25") = "Y"
'
'                   If date10IE_2 <> "" Then
'                   jsonDictionary("Form10IEADate_AY24-25") = Dformat(date10IE_2, "")
'                   End If
'                   If acknowledgementNumber_2 <> "" Then
'                   jsonDictionary("Form10IEAAckNo_AY24-25") = UVCase(CDbl(acknowledgementNumber_2))
'                   End If
'
'            ElseIf TaxRegime_2 = "Not Applicable" Then
'            jsonDictionary("OptOutNewTaxRegime_AY24-25") = "NA"
'            End If
'End If
'
'
'
'
''End --------------
'Dim TaxRegime_1 As Variant  'Sheet1.OptOutNewTaxRegime 'Sheet1.Dateoffiling 'Sheet1.acknowledgementNumber_New
'
'TaxRegime_1 = Sheet1.Range("Sheet1.OptOutNewTaxRegime").value
'
'If TaxRegime_1 <> "" Then
''            If TaxRegime_1 = "No" Then
'            If TaxRegime_1 = "New Tax Regime" Then
'            jsonDictionary("OptOutNewTaxRegime") = "N"
'               If date10IE_1 <> "" Then
'                   jsonDictionary("Form10IEADate") = Dformat(date10IE_1, "")
'                   End If
'                   If acknowledgementNumber_1 <> "" Then
'                   jsonDictionary("Form10IEAAckNo") = UVCase(CDbl(acknowledgementNumber_1))
'                   End If
''            ElseIf TaxRegime_1 = "Yes-but beyond the due date" Then
''            jsonDictionary("OptOutNewTaxRegime") = "B"
''            'Form10IEADate  'Form10IEAAckNo"
'''                  jsonDictionary("Form10IEADate") = Sheet1.Range("Sheet1.Dateoffiling").value
'''                  jsonDictionary("Form10IEAAckNo") = Sheet1.Range("Sheet1.acknowledgementNumber_New").value
''                   If date10IE_1 <> "" Then
''                   jsonDictionary("Form10IEADate") = Dformat(date10IE_1, "")
''                   End If
''                   If acknowledgementNumber_1 <> "" Then
''                   jsonDictionary("Form10IEAAckNo") = UVCase(CDbl(acknowledgementNumber_1))
''                   End If
''            ElseIf TaxRegime_1 = "Yes-within the due date" Then
'            ElseIf TaxRegime_1 = "Old Tax Regime" Then
'            jsonDictionary("OptOutNewTaxRegime") = "Y"
'                   If date10IE_1 <> "" Then
'                   jsonDictionary("Form10IEADate") = Dformat(date10IE_1, "")
'                   End If
'                   If acknowledgementNumber_1 <> "" Then
'                   jsonDictionary("Form10IEAAckNo") = UVCase(CDbl(acknowledgementNumber_1))
'                   End If
'            ElseIf TaxRegime_1 = "Not Applicable" Then
'            jsonDictionary("OptOutNewTaxRegime") = "NA"
'            End If
'End If

'--end--

'added by Chetan C M AY 2025-26 on 2/01/2025
'-- start--

'Dim TaxRegime_1, TaxRegime_2, TaxRegime_3, TaxRegime_4 As Variant
'
'TaxRegime_1 = Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value
'TaxRegime_2 = Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value
'TaxRegime_3 = Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value
'
'TaxRegime_4 = Sheet1.Range("Sheet1.115BAC").value
'
'If TaxRegime_4 <> "" Then
'        If TaxRegime_4 = "No" Then
'
'            jsonDictionary("OptOutNewTaxRegime_Form10IEA_AY24_25") = "N"
'
'            If TaxRegime_2 <> "" Then
'                If TaxRegime_2 = "No" Then
'                    jsonDictionary("No_OptOutNewTaxReg") = "N"
'
'                    ElseIf TaxRegime_2 = "Yes" Then
'                    jsonDictionary("No_OptOutNewTaxReg") = "Y"
'
'                    If date10IE_2 <> "" Then
'                        jsonDictionary("Form10IEADate") = Dformat(date10IE_2, "")
'                    End If
'                    If acknowledgementNumber_2 <> "" Then
'                        jsonDictionary("Form10IEAAckNo") = UVCase(CDbl(acknowledgementNumber_2))
'                    End If
'
'                End If
'            End If
'
'        ElseIf TaxRegime_4 = "Yes" Then
'        jsonDictionary("OptOutNewTaxRegime_Form10IEA_AY24_25") = "Y"
'
'            If date10IE <> "" Then
'                jsonDictionary("Form10IEADate_AY24_25") = Dformat(date10IE, "")
'            End If
'            If acknowledgementNumber <> "" Then
'                jsonDictionary("Form10IEAAckNo_AY24_25") = UVCase(CDbl(acknowledgementNumber))
'            End If
'
'            If TaxRegime_1 <> "" Then
'                If TaxRegime_1 = "No" Then
'                    jsonDictionary("Yes_ContOptOutNewTaxReg") = "N"
'                    If date10IE_1 <> "" Then
'                        jsonDictionary("Form10IEADate") = Dformat(date10IE_1, "")
'                    End If
'                    If acknowledgementNumber_1 <> "" Then
'                        jsonDictionary("Form10IEAAckNo") = UVCase(CDbl(acknowledgementNumber_1))
'                    End If
'            ElseIf TaxRegime_1 = "Yes" Then
'            jsonDictionary("Yes_ContOptOutNewTaxReg") = "Y"
'        End If
'        End If
'
'
'        ElseIf TaxRegime_4 = "Not Applicable" Or TaxRegime_4 = "NA" Then
'            jsonDictionary("OptOutNewTaxRegime_Form10IEA_AY24_25") = "NA"
'        If TaxRegime_3 <> "" Then
'            If TaxRegime_3 = "Yes" Then
'            jsonDictionary("NA_OptOutNewTaxReg") = "Y"
'               If date10IE_3 <> "" Then
'                   jsonDictionary("Form10IEADate") = Dformat(date10IE_3, "")
'                End If
'                If acknowledgementNumber_3 <> "" Then
'                   jsonDictionary("Form10IEAAckNo") = UVCase(CDbl(acknowledgementNumber_3))
'                End If
'            ElseIf TaxRegime_3 = "No" Then
'            jsonDictionary("NA_OptOutNewTaxReg") = "N"
'            End If
'        End If
'    End If
'End If
'--end--
'Konda updated on 23-01-2026




'Ai
acknowledgementNumber = Sheet1.Range("BAC115.AcknowledgementNo").value
'Aia
acknowledgementNumber_1 = Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value

'date10IE_1 = Sheet1.Range("BAC115.DateOfFiling_New_No").value
'acknowledgementNumber_2 = Sheet1.Range("BAC115.AcknowledgementNo_New_No").value

date10IE_2 = Sheet1.Range("BAC115.DateOfFiling_New_No").value
acknowledgementNumber_2 = Sheet1.Range("BAC115.AcknowledgementNo_New_No").value

date10IE_3 = Sheet1.Range("BAC115.DateOfFiling_New_NA").value
acknowledgementNumber_3 = Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value
 
 
If Sheet1.Range("sheet1.Status").value = "F - FIRM(Other than LLP)" Then
    jsonDictionary("Form10IEAEarlierAYOldRegime") = "NA"
Else

If Sheet1.Range("Sheet1.115BAC").value <> "" Then
    If Sheet1.Range("Sheet1.115BAC").value = "Yes" Then
        jsonDictionary("Form10IEAEarlierAYOldRegime") = "Y"
        
        If Sheet1.Range("BAC115.AssessmentYear_2425").value <> "" Then
        
            If Sheet1.Range("BAC115.AssessmentYear_2425").value = "2024-25" Then
            jsonDictionary("Form10IEAAssYear") = "2024-25"
            jsonDictionary("Form10IEAEarlierAYAckOldRegime") = acknowledgementNumber
            End If
        
            If Sheet1.Range("BAC115.AssessmentYear_2425").value = "2025-26" Then
            jsonDictionary("Form10IEAAssYear") = "2025-26"
            jsonDictionary("Form10IEAEarlierAYAckOldRegime") = acknowledgementNumber
            End If
        End If
        
        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value <> "" Then
            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No" Then
                jsonDictionary("F10IEAEarlierAYNewRegime") = "N"
                
                If Sheet1.Range("BAC115.A23ii_No").value <> "" Then
                If Sheet1.Range("BAC115.A23ii_No").value = "Yes" Then
                jsonDictionary("F10IEACurrAYNewRegime") = "Y"
                If Sheet1.Range("BAC115.DateOfFiling_New_No").value <> "" Then
                jsonDictionary("F10IEADateCurrAYNewTax") = Dformat(date10IE_2, "")
                End If
                If Sheet1.Range("BAC115.AcknowledgementNo_New_No").value <> "" Then
                jsonDictionary("F10IEAAckNoCurrAYNewTax") = acknowledgementNumber_2
                End If
                Else
                If Sheet1.Range("BAC115.A23ii_No").value = "No" Then
                jsonDictionary("F10IEACurrAYNewRegime") = "N"
                End If
                End If
                End If
            ElseIf Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "Yes" Then
                    jsonDictionary("F10IEAEarlierAYNewRegime") = "Y"
                If Sheet1.Range("BAC115.AssessmentYear_Aii").value <> "" Then
                jsonDictionary("AssYrF10IEANewTaxReg") = Sheet1.Range("BAC115.AssessmentYear_Aii").value
                jsonDictionary("Form10IEAEarlierAYAckNewRegime") = acknowledgementNumber_1
                End If
            End If
        End If
       
    Else
    If Sheet1.Range("Sheet1.115BAC").value = "No" Then
        jsonDictionary("Form10IEAEarlierAYOldRegime") = "N"
        
        If Sheet1.Range("BAC115.A23b").value <> "" Then
            If Sheet1.Range("BAC115.A23b").value = "No" Then
                jsonDictionary("F10IEACurrAYOldRegime") = "N"
            
            ElseIf Sheet1.Range("BAC115.A23b").value = "Yes" Then
                    jsonDictionary("F10IEACurrAYOldRegime") = "Y"
                If Sheet1.Range("BAC115.DateOfFiling_New_NA").value <> "" Then
                
                jsonDictionary("F10IEADateCurrAYOldTax") = Dformat(date10IE_3, "")
                jsonDictionary("F10IEAAckNoCurrAYOldTax") = acknowledgementNumber_3
                End If
            End If
        End If
    End If
    End If
 End If
  
End If

'============================================================================
If ProvisoFlag = "Yes" Then
    ProvisoFlag = "Y"
ElseIf ProvisoFlag = "No" Then
    ProvisoFlag = "N"
End If

If ProvisoFlag <> "" Then
    jsonDictionary("SeventhProvisio139") = UCase(ProvisoFlag)
End If

If DepositAmountFlag = "Yes" Then
    DepositAmountFlag = "Y"
ElseIf DepositAmountFlag = "No" Then
    DepositAmountFlag = "N"
End If

If DepositAmountFlag <> "" Then
    jsonDictionary("DepAmtAggAmtExcd1CrPrYrFlg") = UCase(DepositAmountFlag)
End If

If DepositAmount <> "" Then
    jsonDictionary("AmtSeventhProvisio139i") = (DepositAmount)
End If

If AggrigateAmountFlag = "Yes" Then
    AggrigateAmountFlag = "Y"
ElseIf AggrigateAmountFlag = "No" Then
    AggrigateAmountFlag = "N"
End If

If AggrigateAmountFlag <> "" Then
    jsonDictionary("IncrExpAggAmt2LkTrvFrgnCntryFlg") = UCase(AggrigateAmountFlag)
End If

If AggrigateAmount <> "" Then
    jsonDictionary("AmtSeventhProvisio139ii") = (AggrigateAmount)
End If

If AggrigateAmountFlag1 = "Yes" Then
    AggrigateAmountFlag1 = "Y"
ElseIf AggrigateAmountFlag1 = "No" Then
    AggrigateAmountFlag1 = "N"
End If

If AggrigateAmountFlag1 <> "" Then
    jsonDictionary("IncrExpAggAmt1LkElctrctyPrYrFlg") = UCase(AggrigateAmountFlag1)
End If

If AggrigateAmount1 <> "" Then
    jsonDictionary("AmtSeventhProvisio139iii") = (AggrigateAmount1)
End If



'If (Sheet8.Range("bacValue").value = "1" And Sheet1.Range("OptingNewTaxRegime").value = "No") Or (Sheet8.Range("bacValue").value = "1" And Sheet1.Range("OptingNewTaxRegime").value = "") Then
'    OptingNewTaxRegime = 1
'ElseIf (Sheet8.Range("bacValue").value = "2" And Sheet1.Range("OptingNewTaxRegime").value = "No") Or (Sheet8.Range("bacValue").value = "No" And Sheet1.Range("OptingNewTaxRegime").value = "") Then
'    OptingNewTaxRegime = "2"
'ElseIf (Sheet8.Range("bacValue").value = "1" And Sheet1.Range("OptingNewTaxRegime").value = "Yes") Then
'    OptingNewTaxRegime = "3"
'ElseIf (Sheet8.Range("bacValue").value = "2" And Sheet1.Range("OptingNewTaxRegime").value = "Yes") Then
'    OptingNewTaxRegime = "4"
'End If

'If Mid(Sheet1.Range("sheet1.Status"), 1, 1) <> "F" Then
'
'If NewTaxRegime <> "" Then
'    If NewTaxRegime = 1 Then
'        jsonDictionary("NewTaxRegime") = "Y"
'    ElseIf NewTaxRegime = 2 Then
'        jsonDictionary("NewTaxRegime") = "N"
'    End If
'End If
'
''Change-10.01.2023.101.
''If (Sheet8.Range("bacValue").value = "1" And Sheet8.Range("oldbacValue").value = "2") Then
''    OptingNewTaxRegime = 1
''ElseIf (Sheet8.Range("bacValue").value = "2" And Sheet8.Range("oldbacValue").value = "2") Then
''    OptingNewTaxRegime = 2
''ElseIf (Sheet8.Range("bacValue").value = "1" And Sheet8.Range("oldbacValue").value = "1") Then
''    OptingNewTaxRegime = 3
''ElseIf (Sheet8.Range("bacValue").value = "2" And Sheet8.Range("oldbacValue").value = "1") Then
''    OptingNewTaxRegime = 4
''End If
'
'    If (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Then
'        OptingNewTaxRegime = 1
'    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Not opting") Then
'        OptingNewTaxRegime = 2
'    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt") Then
'        OptingNewTaxRegime = 3
'    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
'        OptingNewTaxRegime = 4
'    ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in") Then
'        OptingNewTaxRegime = 5
'    End If
'
''Generate-08.12.2022.103.16
''If (Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in") Then
''    OptingNewTaxRegime = 5
''End If
''----End Change
'
'    jsonDictionary("OptingNewTaxRegime") = UVCase(CDbl(OptingNewTaxRegime))
'
'End If


'Generate-08.12.2022.103.09B

'If Sheet8.Range("oldbacValue").value = "1" Then
'        Dim NewTaxRegimeDtls, Form10IEDtls
'        Set NewTaxRegimeDtls = CreateObject("Scripting.Dictionary")
'        Set Form10IEDtls = CreateObject("Scripting.Dictionary")
'        'NewTaxRegimeDtls("AssessmentYear") = AYOptedIn
'        'Form10IEDtls("Form10IEDate") = NewForm10IEDate
'        'Form10IEDtls("Form10IEAckNo") = NewForm10IEAckNo
'        NewTaxRegimeDtls.add "Form10IEDtls", Form10IEDtls
'        jsonDictionary.add "NewTaxRegimeDtls", NewTaxRegimeDtls
'
'        Set Form10IEDtls = Nothing
'        Set Form10IEDtls = CreateObject("Scripting.Dictionary")
'        Set NewTaxRegimeDtls = Nothing
'
'        If Sheet1.Range("Opt_Out115BAC") = "Yes" Then
'            Dim OptedOutNewTaxRegimeDtls
'            Set OptedOutNewTaxRegimeDtls = CreateObject("Scripting.Dictionary")
'            jsonDictionary("OptedOutNewTaxRegime") = "Y"
'            'OptedOutNewTaxRegimeDtls("AssessmentYear") = AYinOptOut
'            'Form10IEDtls("Form10IEDate") = OldForm10IEDate
'            'Form10IEDtls("Form10IEAckNo") = OldForm10IEAckNo
'            OptedOutNewTaxRegimeDtls.add "Form10IEDtls", Form10IEDtls
'            jsonDictionary.add "OptedOutNewTaxRegimeDtls", OptedOutNewTaxRegimeDtls
'            Set OptedOutNewTaxRegimeDtls = Nothing
'        ElseIf Sheet1.Range("Opt_Out115BAC") = "No" Then
'            jsonDictionary("OptedOutNewTaxRegime") = "N"
'        End If
'
'End If

'---End Change




If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 18 Then
    jsonDictionary("ReceiptNo") = RecieptNo
    jsonDictionary("NoticeNo") = NoticeNo
    jsonDictionary("OrigRetFiledDate") = Dformat(DateOfFillingRetrun, "")
End If

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 13 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 14 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 15 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 16 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
    jsonDictionary("NoticeNo") = NoticeNo
End If

UpdateProgressBar

If (Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 17) Or (Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 21) Then
    jsonDictionary("ReceiptNo") = RecieptNo
    If Trim(Sheet1.Range("sheet1.OrigRetFiledDate").value) <> "" Then
        jsonDictionary("OrigRetFiledDate") = Dformat(DateOfFillingRetrun, "")
    End If
End If

UpdateProgressBar

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 13 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 14 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 15 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 16 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 18 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
    jsonDictionary("NoticeDateUnderSec") = Dformat(DateOfNotice, "")
End If

If isdropdownblank(Representativeassesseeflg) Then

Else
    jsonDictionary("AsseseeRepFlg") = Mid(Trim(UCase(Representativeassesseeflg)), 1, 1)
End If
'If date10IE <> "" Then
'    jsonDictionary("Form10IEDate") = Dformat(date10IE, "")
'End If
'If acknowledgementNumber <> "" Then
'    jsonDictionary("Form10IEAckNo") = UVCase(CDbl(acknowledgementNumber))
'End If



'139 clause iv

If Sheet1.Range("clauseiv7provisio139iFlg") = "Yes" Then
        jsonDictionary("clauseiv7provisio139i") = "Y"
       If Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "1"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_1")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "2"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_2")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "3"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_3")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "4"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_4")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       jsonDictionary.add "clauseiv7provisio139iDtls", clauseiv7provisio139iDtls
    ElseIf Sheet1.Range("clauseiv7provisio139iFlg") = "No" Then
        jsonDictionary("clauseiv7provisio139i") = "N"
    End If
    
    





If Mid(Trim(UCase(Representativeassesseeflg)), 1, 1) = "Y" Then

    If NameRepAssessee <> "" Then
    AssesseeRep("RepName") = UCase(NameRepAssessee)
    End If
   'Konda updated on 22-01-2026 schema-V0.2.5
   
'    If Not isdropdownblank(CapacityRep) Then
'    AssesseeRep("RepCapacity") = Mid(Trim(UCase(CapacityRep)), 1, 1)
'    End If
'
'    If AddressRepAssessee <> "" Then
'    AssesseeRep("RepAddress") = UCase(AddressRepAssessee)
'    End If
'
'    If PANRepAssessee <> "" Then
'    AssesseeRep("RepPAN") = UCase(PANRepAssessee)
'    End If
'
'    If AadhaarRepAssessee <> "" Then
'    AssesseeRep("RepAadhaar") = UCase(AadhaarRepAssessee)
'    End If


    If EmailRepAssessee <> "" Then
    AssesseeRep("RepEmailID") = UCase(EmailRepAssessee)
    End If

    If CountryCodeRepAssessee <> "" Then
    AssesseeRep("CountryCodeRepMobileNo") = CDbl(CountryCodeRepAssessee)
    End If

    If MobileRepAssessee <> "" Then
    AssesseeRep("RepMobileNo") = CDbl(MobileRepAssessee)
    End If
'==========================================
jsonDictionary.add "AssesseeRep", AssesseeRep
End If
'-------------------------------------------------------------

'"ItrFilingDueDate" "2024-07-31"
'Sheet1.Duedate "31/07/2024"

'PAG_C9 AY 2024-25-----------
Dim ItrFilingDueDate_2 As Variant
'ItrFilingDueDate_2 = Mid((Sheet1.Range("Sheet1.Duedate").value), 7, 4) & "/" & Mid((Sheet1.Range("Sheet1.Duedate").value), 4, 2) & "/" & Mid((Sheet1.Range("Sheet1.Duedate").value), 1, 2)
ItrFilingDueDate_2 = Mid((Sheet1.Range("Sheet1.Duedate").value), 7, 4) & "-" & Mid((Sheet1.Range("Sheet1.Duedate").value), 4, 2) & "-" & Mid((Sheet1.Range("Sheet1.Duedate").value), 1, 2) 'AY 2024-25

If Sheet1.Range("Sheet1.Duedate").value <> "" Then
    jsonDictionary("ItrFilingDueDate") = UCase(ItrFilingDueDate_2)
End If

'-----------------
UpdateProgressBar

Set FilingStatus = jsonDictionary

End Function


Function Schedule80GGC() As Object
Dim i As Long
Dim jsonDictionary, Schedule80GGC_1, AddressDetail
Dim Schedule80GGCDetails, TotalDonationAmtCash80GGC, TotalDonationAmtOtherMode80GGC, TotalDonationsUs80GGC, TotalEligibleDonationAmt80GGC As Variant
Set Schedule80GGCDetails = New Collection
Dim dateDonation
Dim dob As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Schedule80GGC_1 = CreateObject("Scripting.Dictionary")
'Set AddressDetail = CreateObject("Scripting.Dictionary")
'Set Schedule80GGCDetails = CreateObject("Scripting.Dictionary")

TotalDonationAmtCash80GGC = UCase(Range("SCH_80GGC_TotalDonationAmtCash80GGC").value)
TotalDonationAmtOtherMode80GGC = UCase(Range("SCH_80GGC_TotalDonationAmtOtherMode80GGC").value)
TotalDonationsUs80GGC = UCase(Range("SCH_80GGC_TotalDonationsUs80GGC").value)
TotalEligibleDonationAmt80GGC = UCase(Range("SCH_80GGC_TotalEligibleDonationAmt80GGC").value)


subProcCaption = "80GGC"
    If Not IsEmpty(DateofDonation_80GGC) Then
    If (UBound(DateofDonation_80GGC) > 0) Then
    
      For i = 1 To UBound(DateofDonation_80GGC)

            dob = Mid(DateofDonation_80GGC(i), 7, 4) & "-" & Mid(DateofDonation_80GGC(i), 4, 2) & "-" & Mid(DateofDonation_80GGC(i), 1, 2)

            If DateofDonation_80GGC(i) <> "" Then
                Schedule80GGC_1("DonationDate") = dob
            End If
            
            If Donation_cash_80GGC(i) <> "" Then
                Schedule80GGC_1("DonationAmtCash") = UVCase(CDbl(Donation_cash_80GGC(i)))
            Else
                Schedule80GGC_1("DonationAmtCash") = 0
            End If
            
            If Donation_other_80GGC(i) <> "" Then
                Schedule80GGC_1("DonationAmtOtherMode") = UVCase(CDbl(Donation_other_80GGC(i)))
            Else
                Schedule80GGC_1("DonationAmtOtherMode") = 0
            End If

            
'            Dim natureoftrans
'
'            If UCase(NatureofTransaction_80GGC(i)) = UCase("UPI transfer") Then
'               natureoftrans = "UPI"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("Cheque issued") Then
'               natureoftrans = "Cheque"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("IMPS") Then
'               natureoftrans = "IMPS"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("NEFT") Then
'               natureoftrans = "NEFT"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("RTGS") Then
'               natureoftrans = "RTGS"
'            End If
'
'
'            If NatureofTransaction_80GGC(i) <> "" Then
'                Schedule80GGC_1("NatureOfTrans") = natureoftrans
'            End If
            
            
            
            If BankIFSC_80GGC(i) <> "" Then
                Schedule80GGC_1("IFSCCode") = UCase(BankIFSC_80GGC(i))
            End If
            
'            If BankName_80GGC(i) <> "" Then
'                Schedule80GGC_1("BankName") = UCase(BankName_80GGC(i))
'            End If
            
'            If BankAccntnum_80GGC(i) <> "" Then
'                Schedule80GGC_1("BankAccountNo") = UVCase(CDbl(BankAccntnum_80GGC(i)))
'            End If
            
             If ChequeNumber_80GGC(i) <> "" Then
                Schedule80GGC_1("TransactionRefNum") = Trim(ChequeNumber_80GGC(i))
             End If
            
            
            If Trim(Range("SCH_80GGC_DonationAmt").Cells.item(i).value) <> "" Then
                Schedule80GGC_1("DonationAmt") = UVCase(CDbl(Range("SCH_80GGC_DonationAmt").Cells.item(i).value))
            Else
                Schedule80GGC_1("DonationAmt") = 0
            End If
            
            If Trim(Range("SCH_80GGC_EligibleDonationAmt").Cells.item(i).value) <> "" Then
                Schedule80GGC_1("EligibleDonationAmt") = UVCase(CDbl(Range("SCH_80GGC_EligibleDonationAmt").Cells.item(i).value))
            Else
                Schedule80GGC_1("EligibleDonationAmt") = 0
            End If
'Konda updated on 17-03-2026--V0.6
            If Trim(Range("SCH_80GGC_Political_Name").Cells.item(i).value) <> "" Then
                Schedule80GGC_1("PoliticalPartyName") = UCase(Range("SCH_80GGC_Political_Name").Cells.item(i).value)
            End If
            
            If Trim(Range("SCH_80GGC_Political_PAN").Cells.item(i).value) <> "" Then
                Schedule80GGC_1("PoliticalPartyPAN") = UCase(Range("SCH_80GGC_Political_PAN").Cells.item(i).value)
            End If
'===================================
        
            Schedule80GGCDetails.add Schedule80GGC_1
            Set Schedule80GGC_1 = Nothing
            Set Schedule80GGC_1 = CreateObject("Scripting.Dictionary")
    
            UpdateProgressBar
            Next
        End If
    End If
    jsonDictionary.add "Schedule80GGCDetails", Schedule80GGCDetails
            
            If Trim(Range("SCH_80GGC_TotalDonationAmtCash80GGC").value) > 0 Then
                jsonDictionary("TotalDonationAmtCash80GGC") = UVCase(CDbl(TotalDonationAmtCash80GGC))
            Else
                jsonDictionary("TotalDonationAmtCash80GGC") = 0
            End If
            
            If Trim(Range("SCH_80GGC_TotalDonationAmtOtherMode80GGC").value) > 0 Then
                jsonDictionary("TotalDonationAmtOtherMode80GGC") = UVCase(CDbl(TotalDonationAmtOtherMode80GGC))
            Else
                jsonDictionary("TotalDonationAmtOtherMode80GGC") = 0
            End If
            
            If Trim(Range("SCH_80GGC_TotalDonationsUs80GGC").value) > 0 Then
                jsonDictionary("TotalDonationsUs80GGC") = UVCase(CDbl(TotalDonationsUs80GGC))
            Else
                jsonDictionary("TotalDonationsUs80GGC") = 0
            End If
  
            If Trim(Range("SCH_80GGC_TotalEligibleDonationAmt80GGC").value) > 0 Then
                jsonDictionary("TotalEligibleDonationAmt80GGC") = UVCase(CDbl(TotalEligibleDonationAmt80GGC))
            Else
                jsonDictionary("TotalEligibleDonationAmt80GGC") = 0
            End If
Set Schedule80GGC = jsonDictionary
End Function

Function Schedule80DD() As Object
subProcCaption = "Schedule 80DD"
noOfProcessSub = 10
Dim jsonDictionary, Schedule80DD_D

Dim NatureOfDisability, DeductionAmount, DependentType, DependentPan, DependentAadhaar, Form10IAFilingDate, Form10IAAckNum, UDIDNum As Variant
Dim NatureOfDisability_1, DeductionAmount_1, DependentType_1, DependentPan_1, DependentAadhaar_1, Form10IAFilingDate_1, Form10IAAckNum_1, UDIDNum_1 As Variant



Set jsonDictionary = CreateObject("Scripting.Dictionary")


NatureOfDisability_1 = Sheet14.Range("NatureDisability_80DD").Cells(1, 1).value
DeductionAmount_1 = Sheet14.Range("AmtDeduction_80DD").value
DependentType_1 = Sheet14.Range("Dependent_80DD").value
DependentPan_1 = Sheet14.Range("PanDependent_80DD").value
DependentAadhaar_1 = Sheet14.Range("AadhaarDependent_80DD").value
'AY_25_26_V0.7
'Form10IAFilingDate_1 = Sheet14.Range("Form10IA_80DD").value
'-------------------------

Form10IAAckNum_1 = Sheet14.Range("AcknowledgeNum_80DD").value
UDIDNum_1 = Sheet14.Range("UDIDNumber_80DD").value
 
 
 'AY_2025_26 enhancement
Dim TypeOfDisability_GG, FormAckNum11A_1, TypeOfDisability_G
TypeOfDisability_G = Sheet14.Range("Disability_80DD").value
'AY_25_26_V0.7
'FormAckNum11A_1 = Sheet14.Range("AcknowledgeNum11A2_80DD").value
'-------------------------
If TypeOfDisability_G <> "" And TypeOfDisability_G <> "(Select)" Then

'Malli----------
If TypeOfDisability_G = "(i) autism, cerebral palsy, or multiple disabilities" Then
TypeOfDisability_GG = "1"

ElseIf TypeOfDisability_G = "(ii) others" Then
TypeOfDisability_GG = "2"
End If


''AY_25_26_V0.7
'If TypeOfDisability_G = "(i) blindness" Then
'TypeOfDisability_GG = "1"
'
'ElseIf TypeOfDisability_G = "(ii) low vision" Then
'TypeOfDisability_GG = "2"
'
'ElseIf TypeOfDisability_G = "(iii) leprosy-cured" Then
'TypeOfDisability_GG = "3"
'
'ElseIf TypeOfDisability_G = "(iv) hearing impairment" Then
'TypeOfDisability_GG = "4"
'
'ElseIf TypeOfDisability_G = "(v) locomotor disability" Then
'TypeOfDisability_GG = "5"
'
'ElseIf TypeOfDisability_G = "(vi) mental retardation" Then
'TypeOfDisability_GG = "6"
'
'ElseIf TypeOfDisability_G = "(vii) mental illness" Then
'TypeOfDisability_GG = "7"
'
'ElseIf TypeOfDisability_G = "(viii) autism" Then
'TypeOfDisability_GG = "8"
'
'ElseIf TypeOfDisability_G = "(ix) cerebral palsy" Then
'TypeOfDisability_GG = "9"
'
'ElseIf TypeOfDisability_G = "(x) multiple disability" Then
'TypeOfDisability_GG = "10"
'
'End If
''----------------------------------------
End If
'----------------------------------------------

 
 
Dim Ndisability80DD As Variant
Ndisability80DD = Mid(NatureOfDisability_1, 1, 1)

If IsNumeric(Ndisability80DD) Then

If Ndisability80DD = "1" Or Ndisability80DD = "2" Then
jsonDictionary("NatureOfDisability") = Ndisability80DD
'Else: jsonDictionary("NatureOfDisability") = ""
End If


'Malli_18/04/2025
'TypeOfDisability
'FormAckNum11A
If TypeOfDisability_GG <> "" Then
jsonDictionary("TypeOfDisability") = UCase(TypeOfDisability_GG)
Else
jsonDictionary("TypeOfDisability") = ""
End If

'----------------


If Sheet14.Range("AmtDeduction_80DD").value <> "" And Sheet14.Range("AmtDeduction_80DD").value <> "0" Then
jsonDictionary("DeductionAmount") = UVCase(CDbl(DeductionAmount_1))
Else
jsonDictionary("DeductionAmount") = 0
End If

Dim Tdept_typ80DD
Tdept_typ80DD = Mid(DependentType_1, 1, 1)
If Tdept_typ80DD <> "" And Tdept_typ80DD <> "(" Then
jsonDictionary("DependentType") = Tdept_typ80DD
'ElseIf Tdept_typ80DD = "" Or Tdept_typ80DD = "(" Then
'jsonDictionary("DependentType") = ""
End If

If Sheet14.Range("PanDependent_80DD").value <> "" Then
jsonDictionary("DependentPan") = UCase(DependentPan_1)
End If

If Sheet14.Range("AadhaarDependent_80DD").value <> "" Then
jsonDictionary("DependentAadhaar") = UCase(DependentAadhaar_1)
End If

'Malli-----------
'AY_25_26_V0.7
'If Sheet14.Range("Form10IA_80DD").value <> "" Then
'Dim dob As Variant
'dob = Mid(Form10IAFilingDate_1, 7, 4) & "-" & Mid(Form10IAFilingDate_1, 4, 2) & "-" & Mid(Form10IAFilingDate_1, 1, 2)
'jsonDictionary("Form10IAFilingDate") = dob
'End If
'-----------------------------


If Sheet14.Range("AcknowledgeNum_80DD").value <> "" Then
jsonDictionary("Form10IAAckNum") = UCase(Form10IAAckNum_1)
End If

'AY_25_26_V0.7
''Malli----------18/04/2025
'If FormAckNum11A_1 <> "" Then
'jsonDictionary("FormAckNum11A") = UCase(FormAckNum11A_1)
'End If
'-----------------------------------------

If Sheet14.Range("UDIDNumber_80DD").value <> "" Then
jsonDictionary("UDIDNum") = UCase(UDIDNum_1)
End If


Set Schedule80DD = jsonDictionary

End If
End Function


Function Schedule80U() As Object
subProcCaption = "Schedule 80U"
noOfProcessSub = 10
Dim jsonDictionary, Schedule80U_D

 Dim NatureOfDisability, DeductionAmount, Form10IAFilingDate, Form10IAAckNum, UDIDNum As Variant
 
 Dim NatureOfDisability_1, DeductionAmount_1, Form10IAFilingDate_1, Form10IAAckNum_1, UDIDNum_1 As Variant
 
 Set jsonDictionary = CreateObject("Scripting.Dictionary")
 
 
 NatureOfDisability_1 = Sheet14.Range("NatureDisability_80U").Cells(1, 1).value
 
 DeductionAmount_1 = Sheet14.Range("AmountDeduction_80U").value
 ''AY_25_26_V0.7
 'Form10IAFilingDate_1 = Sheet14.Range("DateFiling10IA_80U").value
 '--------------------------------
 Form10IAAckNum_1 = Sheet14.Range("AckNumber10IA_80U").value
 UDIDNum_1 = Sheet14.Range("UDIDNum_80U").value



'Malli------------18/04/2025
'TypeOfDisability_80U ,FormAckNum11A_80U,TypeOfDisability_G_80U
Dim TypeOfDisability_80U, FormAckNum11A_80U, TypeOfDisability_G_80U
TypeOfDisability_80U = Sheet14.Range("Disability_80U").value

'AY_25_26_V0.7

'FormAckNum11A_80U = Sheet14.Range("AcknowledgeNum11A2_80U").value

'---------------------------



Dim Ndisability80U As Variant
Ndisability80U = Mid(NatureOfDisability_1, 1, 1)

If IsNumeric(Ndisability80U) Then

If Ndisability80U = "1" Or Ndisability80U = "2" Then
jsonDictionary("NatureOfDisability") = Ndisability80U
'Else: jsonDictionary("NatureOfDisability") = ""

End If


'Malli----------18/04/2025
If TypeOfDisability_80U <> "" And TypeOfDisability_80U <> "(Select)" Then

If TypeOfDisability_80U = "(i) autism, cerebral palsy, or multiple disabilities" Then
TypeOfDisability_G_80U = "1"

ElseIf TypeOfDisability_80U = "(ii) others" Then
TypeOfDisability_G_80U = "2"
End If
'AY_25_26_V0.7
'If TypeOfDisability_80U = "(i) blindness" Then
'TypeOfDisability_G_80U = "1"
'
'ElseIf TypeOfDisability_80U = "(ii) low vision" Then
'TypeOfDisability_G_80U = "2"
'ElseIf TypeOfDisability_80U = "(iii) leprosy-cured" Then
'TypeOfDisability_G_80U = "3"
'
'ElseIf TypeOfDisability_80U = "(iv) hearing impairment" Then
'TypeOfDisability_G_80U = "4"
'
'ElseIf TypeOfDisability_80U = "(v) locomotor disability" Then
'TypeOfDisability_G_80U = "5"
'
'ElseIf TypeOfDisability_80U = "(vi) mental retardation" Then
'TypeOfDisability_G_80U = "6"
'
'ElseIf TypeOfDisability_80U = "(vii) mental illness" Then
'TypeOfDisability_G_80U = "7"
'
'ElseIf TypeOfDisability_80U = "(viii) autism" Then
'TypeOfDisability_G_80U = "8"
'
'ElseIf TypeOfDisability_80U = "(ix) cerebral palsy" Then
'TypeOfDisability_G_80U = "9"
'
'ElseIf TypeOfDisability_80U = "(x) multiple disability" Then
'TypeOfDisability_G_80U = "10"
'end if
'--------------------------------------


End If

If TypeOfDisability_G_80U <> "" Then
jsonDictionary("TypeOfDisability") = UCase(TypeOfDisability_G_80U)
Else
jsonDictionary("TypeOfDisability") = ""
End If
'-------------------------

 
If Sheet14.Range("AmountDeduction_80U").value <> "" And Sheet14.Range("AmountDeduction_80U").value <> "0" Then
jsonDictionary("DeductionAmount") = UVCase(CDbl(DeductionAmount_1))
Else
jsonDictionary("DeductionAmount") = 0
End If

'''AY_25_26_V0.7
'If Sheet14.Range("DateFiling10IA_80U").value <> "" Then
'Dim dob As Variant
'dob = Mid(Form10IAFilingDate_1, 7, 4) & "-" & Mid(Form10IAFilingDate_1, 4, 2) & "-" & Mid(Form10IAFilingDate_1, 1, 2)
'jsonDictionary("Form10IAFilingDate") = dob
'End If
''-------------------

'If Sheet14.Range("DateFiling10IA_80U").value <> "" Then
'AckNumber10IA_80U
If Sheet14.Range("AckNumber10IA_80U").value <> "" Then
jsonDictionary("Form10IAAckNum") = UCase(Form10IAAckNum_1)
End If

'''AY_25_26_V0.7
''Malli-----------18/04/2025
'If FormAckNum11A_80U <> "" Then
'jsonDictionary("FormAckNum11A") = UCase(FormAckNum11A_80U)
'End If
''--------------

If Sheet14.Range("UDIDNum_80U").value <> "" Then
jsonDictionary("UDIDNum") = UCase(UDIDNum_1)
End If

Set Schedule80U = jsonDictionary
End If
End Function

'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80E() As Object
Dim i As Long
Dim mIntCells80E As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EDtls_G, Schedule80EDtls
Set Schedule80EDtls_G = CreateObject("Scripting.Dictionary")
'Set Schedule80EDtls_G = New Collection

Set Schedule80EDtls = CreateObject("Scripting.Dictionary")
Set Schedule80EDtls = New Collection

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80E"

  
  If Sheet20.Range("TotAmt.80E").value > 0 Then
     
     mIntCells80E = Sheet20.Range("LoanfrmBankOrInstitute.80E").count
  
  For i = 1 To mIntCells80E
  
  If (Sheet20.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).value <> "" And Sheet20.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).value <> "(Select)") Then
  ' Debug.Print Sheet20.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).value
   LoanTknFrom_G = Sheet20.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).value
   
   'Malli_V0.7
   'IFSCCode_G = Sheet20.Range("IFSC.80E").Cells(i, 1).value
  ' PAN_G = Sheet20.Range("PAN.80E").Cells(i, 1).value
   '--------------
   BankOrInstnName_G = Sheet20.Range("bankName.80E").Cells(i, 1).value
   
   LoanAccNoOfBankOrInstnRefNo_G = Sheet20.Range("loanAccNum.80E").Cells(i, 1).value
   DateofLoan_G = Sheet20.Range("loanDate.80E").Cells(i, 1).value
   TotalLoanAmt_G = Sheet20.Range("loanAmt.80E").Cells(i, 1).value
   LoanOutstndngAmt_G = Sheet20.Range("loanOutstanding.80E").Cells(i, 1).value
   InterestPaid_G = Sheet20.Range("Intrst.80E").Cells(i, 1).value
   
   'TotalInterestPaid_G = Sheet20.Range("TotAmt.80E").Cells(i, 1).value
   
   If LoanTknFrom_G <> "" And LoanTknFrom_G <> "(Select)" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
            ElseIf LoanTknFrom_G = "Institution" Then
            LoanTknFrom_G = "I"
            End If
        Schedule80EDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EDtls_G("LoanTknFrom") = ""
   End If
   
'Malli---V0.7
'   If IFSCCode_G <> "" Then
'   Schedule80EDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
   
'   If PAN_G <> "" Then
'   Schedule80EDtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
'---------------------
   
   
   If BankOrInstnName_G <> "" Then
   Schedule80EDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EDtls_G("BankOrInstnName") = ""
   End If
   
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EDtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'Schedule80EDtls_G("InterestPaid") = UVCase(InterestPaid_G)
   Schedule80EDtls_G("Interest80E") = UVCase(InterestPaid_G)
  'End If
  
                Schedule80EDtls.add Schedule80EDtls_G
                Set Schedule80EDtls_G = Nothing
                Set Schedule80EDtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
                jsonDictionary.add "Schedule80EDtls", Schedule80EDtls
                Set Schedule80EDtls = Nothing
                Set Schedule80EDtls = CreateObject("Scripting.Dictionary")

  'End If
        
TotalInterestPaid_G = Sheet20.Range("TotAmt.80E").value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80E") = UVCase(TotalInterestPaid_G)

Set Schedule80E = jsonDictionary

End If 'Malli 23/04/2025

End Function

'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80EE() As Object
Dim i As Long
Dim mIntCells80EE As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EEDtls_G, Schedule80EEDtls
Set Schedule80EEDtls_G = CreateObject("Scripting.Dictionary")
'Set Schedule80EDtls_G = New Collection

Set Schedule80EEDtls = CreateObject("Scripting.Dictionary")
Set Schedule80EEDtls = New Collection

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80EE"


' 'Malli--------V0.7
'  'Malli----------23/04/2025
'
'   If Sheet20.Range("ResidentialHP.80E").value > 0 Then
'  '----------------------
'  'Dim ResHPVal_G
'
'  jsonDictionary("ResHPVal") = UVCase(Sheet20.Range("ResidentialHP.80E").value)
'
' '--------------------------------------------------------------------------------
  If Sheet20.Range("TotAmt.80EE").value > 0 Then
     
     mIntCells80EE = Sheet20.Range("LoanfrmBankOrInstitute.80EE").count
  
  For i = 1 To mIntCells80EE
  
  If (Sheet20.Range("LoanfrmBankOrInstitute.80EE").Cells(i, 1).value <> "" And Sheet20.Range("LoanfrmBankOrInstitute.80EE").Cells(i, 1).value <> "(Select)") Then
  
   'Debug.Print Sheet20.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).value
   
   LoanTknFrom_G = Sheet20.Range("LoanfrmBankOrInstitute.80EE").Cells(i, 1).value
    'Malli--------V0.7
   'IFSCCode_G = Sheet20.Range("IFSC.80EE").Cells(i, 1).value
   ' PAN_G = Sheet20.Range("PAN.80EE").Cells(i, 1).value
   '------------------------
   BankOrInstnName_G = Sheet20.Range("bankName.80EE").Cells(i, 1).value
  
   LoanAccNoOfBankOrInstnRefNo_G = Sheet20.Range("loanAccNum.80EE").Cells(i, 1).value
   DateofLoan_G = Sheet20.Range("loanDate.80EE").Cells(i, 1).value
   TotalLoanAmt_G = Sheet20.Range("loanAmt.80EE").Cells(i, 1).value
   LoanOutstndngAmt_G = Sheet20.Range("loanOutstanding.80EE").Cells(i, 1).value
   InterestPaid_G = Sheet20.Range("Intrst.80EE").Cells(i, 1).value
   
  
   
   If LoanTknFrom_G <> "" And LoanTknFrom_G <> "(Select)" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
            ElseIf LoanTknFrom_G = "Institution" Then
            LoanTknFrom_G = "I"
            End If
        Schedule80EEDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EEDtls_G("LoanTknFrom") = ""
   End If
   
'   'Malli--------V0.7
'   If IFSCCode_G <> "" Then
'   Schedule80EEDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
'   If PAN_G <> "" Then
'   Schedule80EEDtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
'   '-------------------------
   
   If BankOrInstnName_G <> "" Then
   Schedule80EEDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EEDtls_G("BankOrInstnName") = ""
   End If
   
  
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EEDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EEDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EEDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EEDtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EEDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EEDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'Schedule80EEDtls_G("InterestPaid") = UVCase(InterestPaid_G)
    Schedule80EEDtls_G("Interest80EE") = UVCase(InterestPaid_G)
   
  'End If
  
                Schedule80EEDtls.add Schedule80EEDtls_G
                Set Schedule80EEDtls_G = Nothing
                Set Schedule80EEDtls_G = CreateObject("Scripting.Dictionary")
  
  End If
  Next
                jsonDictionary.add "Schedule80EEDtls", Schedule80EEDtls
                Set Schedule80EEDtls = Nothing
                Set Schedule80EEDtls = CreateObject("Scripting.Dictionary")

  'End If
        
TotalInterestPaid_G = Sheet20.Range("TotAmt.80EE").value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80EE") = UVCase(TotalInterestPaid_G)

Set Schedule80EE = jsonDictionary

End If  'Malli '23/04/2025
End Function
'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80EEA() As Object
Dim i As Long
Dim mIntCells80EEA As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EEADtls_G, Schedule80EEADtls
Set Schedule80EEADtls_G = CreateObject("Scripting.Dictionary")
 

Set Schedule80EEADtls = CreateObject("Scripting.Dictionary")
Set Schedule80EEADtls = New Collection

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80EEA"

'Malli------23/04/2025
If Sheet20.Range("Stampduty.80EEA").value > 0 Then
'----------------------------------------------------
  
  jsonDictionary("PropStmpDtyVal") = UVCase(Sheet20.Range("Stampduty.80EEA").value)
  
  If Sheet20.Range("TotAmt.80EEA").value > 0 Then
     
     mIntCells80EEA = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").count
  
  For i = 1 To mIntCells80EEA
  
  If (Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Cells(i, 1).value <> "" And Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Cells(i, 1).value <> "(Select)") Then
  
   
   
   LoanTknFrom_G = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Cells(i, 1).value
   'AY_25_26_V0.7
   'IFSCCode_G = Sheet20.Range("IFSC.80EEA").Cells(i, 1).value
   'PAN_G = Sheet20.Range("PAN.80EEA").Cells(i, 1).value
   '----------------------
   BankOrInstnName_G = Sheet20.Range("bankName.80EEA").Cells(i, 1).value
   
   LoanAccNoOfBankOrInstnRefNo_G = Sheet20.Range("loanAccNum.80EEA").Cells(i, 1).value
   DateofLoan_G = Sheet20.Range("loanDate.80EEA").Cells(i, 1).value
   TotalLoanAmt_G = Sheet20.Range("loanAmt.80EEA").Cells(i, 1).value
   LoanOutstndngAmt_G = Sheet20.Range("loanOutstanding.80EEA").Cells(i, 1).value
   InterestPaid_G = Sheet20.Range("Intrst.80EEA").Cells(i, 1).value
   
  
   
   If LoanTknFrom_G <> "" And LoanTknFrom_G <> "(Select)" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
            ElseIf LoanTknFrom_G = "Institution" Then
            LoanTknFrom_G = "I"
            End If
        Schedule80EEADtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EEADtls_G("LoanTknFrom") = ""
   End If
    'AY_25_26_V0.7
'    If IFSCCode_G <> "" Then
'    Schedule80EEADtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'    End If
'
'    If PAN_G <> "" Then
'    Schedule80EEADtls_G("PAN") = UCase(Trim(PAN_G))
'    End If
   '----------------------
   If BankOrInstnName_G <> "" Then
   Schedule80EEADtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EEADtls_G("BankOrInstnName") = ""
   End If
   
   
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EEADtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EEADtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EEADtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EEADtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EEADtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EEADtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'Schedule80EEADtls_G("InterestPaid") = UVCase(InterestPaid_G)
   Schedule80EEADtls_G("Interest80EEA") = UVCase(InterestPaid_G)
   
  'End If
  
                Schedule80EEADtls.add Schedule80EEADtls_G
                Set Schedule80EEADtls_G = Nothing
                Set Schedule80EEADtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
                jsonDictionary.add "Schedule80EEADtls", Schedule80EEADtls
                Set Schedule80EEADtls = Nothing
                Set Schedule80EEADtls = CreateObject("Scripting.Dictionary")
  End If
        

TotalInterestPaid_G = Sheet20.Range("TotAmt.80EEA").value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80EEA") = UVCase(TotalInterestPaid_G)

Set Schedule80EEA = jsonDictionary

End If 'Malli 23/04/2025

End Function
'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80EEB() As Object
Dim i As Long
Dim mIntCells80EEB As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EEBDtls_G, Schedule80EEBDtls
Set Schedule80EEBDtls_G = CreateObject("Scripting.Dictionary")
'Set Schedule80EDtls_G = New Collection

Set Schedule80EEBDtls = CreateObject("Scripting.Dictionary")
Set Schedule80EEBDtls = New Collection

Dim VehicleValue_G, VehicleRegNo_G, LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80EEB"

  
  If Sheet20.Range("TotAmt.80EEB").value > 0 Then
     
     mIntCells80EEB = Sheet20.Range("LoanfrmBankOrInstitute.80EEB").count
  
  For i = 1 To mIntCells80EEB
  
  If (Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Cells(i, 1).value <> "" And Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Cells(i, 1).value <> "(Select)") Then
  
   
   
   LoanTknFrom_G = Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Cells(i, 1).value
   'AY_25_26_V0.7
   'IFSCCode_G = Sheet20.Range("IFSC.80EEB").Cells(i, 1).value
    'PAN_G = Sheet20.Range("PAN.80EEB").Cells(i, 1).value
   '-----------------------
   BankOrInstnName_G = Sheet20.Range("bankName.80EEB").Cells(i, 1).value
  
   LoanAccNoOfBankOrInstnRefNo_G = Sheet20.Range("loanAccNum.80EEB").Cells(i, 1).value
   DateofLoan_G = Sheet20.Range("loanDate.80EEB").Cells(i, 1).value
   TotalLoanAmt_G = Sheet20.Range("loanAmt.80EEB").Cells(i, 1).value
   LoanOutstndngAmt_G = Sheet20.Range("loanOutstanding.80EEB").Cells(i, 1).value
   'AY_25_26_V0.7
   'VehicleValue_G = Sheet20.Range("Vehicle_value.80EEB").Cells(i, 1).value
   '---------------
   VehicleRegNo_G = Sheet20.Range("VehicleRegNum.80EEB").Cells(i, 1).value
   InterestPaid_G = Sheet20.Range("Intrst.80EEB").Cells(i, 1).value
   
  
   
   If LoanTknFrom_G <> "" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
            ElseIf LoanTknFrom_G = "Institution" Then
            LoanTknFrom_G = "I"
            End If
        Schedule80EEBDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EEBDtls_G("LoanTknFrom") = ""
   End If
   'AY_25_26_V0.7
'   If IFSCCode_G <> "" Then
'   Schedule80EEBDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
'
'   If PAN_G <> "" Then
'   Schedule80EEBDtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
   '-------------------
   If BankOrInstnName_G <> "" Then
   Schedule80EEBDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EEBDtls_G("BankOrInstnName") = ""
   End If
   
   
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EEBDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EEBDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EEBDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EEBDtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EEBDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EEBDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'AY_25_26_V0.7
  ' Schedule80EEBDtls_G("VehicleValue") = UVCase(VehicleValue_G)
   '-----------------
   If VehicleRegNo_G <> "" Then
   Schedule80EEBDtls_G("VehicleRegNo") = UCase(Trim(VehicleRegNo_G))
   Else
   Schedule80EEBDtls_G("VehicleRegNo") = ""
   End If
   
  ' Schedule80EEBDtls_G("InterestPaid") = UVCase(InterestPaid_G)
   Schedule80EEBDtls_G("Interest80EEB") = UVCase(InterestPaid_G)
   
  'End If
  
                Schedule80EEBDtls.add Schedule80EEBDtls_G
                Set Schedule80EEBDtls_G = Nothing
                Set Schedule80EEBDtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
  
               jsonDictionary.add "Schedule80EEBDtls", Schedule80EEBDtls
               Set Schedule80EEBDtls = Nothing
               Set Schedule80EEBDtls = CreateObject("Scripting.Dictionary")
  'End If
           

TotalInterestPaid_G = Sheet20.Range("TotAmt.80EEB").value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80EEB") = UVCase(TotalInterestPaid_G)

Set Schedule80EEB = jsonDictionary

 End If  'Malli 23/04/2025
End Function


'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80C() As Object
Dim i As Long
Dim mIntCells80C As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80CDtls_G, Schedule80CDtls
Set Schedule80CDtls_G = CreateObject("Scripting.Dictionary")
 
Set Schedule80CDtls = CreateObject("Scripting.Dictionary")
Set Schedule80CDtls = New Collection

Dim NatureOfPayment_G, IdentificationNo_G, Amount_G, TotalAmt_G
subProcCaption = "80C"

  
If Sheet19.Range("TotAmount.80C").value > 0 Then
   'AY_25_26_V0.7
     'mIntCells80C = Sheet19.Range("NaturePayment.80C").count
     mIntCells80C = Sheet19.Range("Amount.80C").count
  '----------------------
  
  
   For i = 1 To mIntCells80C
   'AY_25_26_V0.7
  ' If (Sheet19.Range("NaturePayment.80C").Cells(i, 1).value <> "" And Sheet19.Range("NaturePayment.80C").Cells(i, 1).value <> "(Select)") Then
   '-------------------------
    If Sheet19.Range("Amount.80C").Cells(i, 1).value <> "" Then  'And Sheet19.Range("NaturePayment.80C").Cells(i, 1).value <> "(Select)") Then
   
   ''AY_25_26_V0.7
   'NatureOfPayment_G = Sheet19.Range("NaturePayment.80C").Cells(i, 1).value
   '----------------------
   IdentificationNo_G = Sheet19.Range("Identification_Number.80C").Cells(i, 1).value
   Amount_G = Sheet19.Range("Amount.80C").Cells(i, 1).value
   
'    'AY_25_26_V0.7
'   If NatureOfPayment_G <> "" Then
'    Schedule80CDtls_G("NatureOfPayment") = UCase(Trim(NatureOfPayment_G))
'   Else
'   Schedule80CDtls_G("NatureOfPayment") = ""
'   End If
'    '---------------
    
     Schedule80CDtls_G("Amount") = UVCase(Amount_G)
    
   If IdentificationNo_G <> "" Then
    Schedule80CDtls_G("IdentificationNo") = UCase(Trim(IdentificationNo_G))
   Else
   Schedule80CDtls_G("IdentificationNo") = ""
   End If
   
   
    
   'End If
                Schedule80CDtls.add Schedule80CDtls_G
                Set Schedule80CDtls_G = Nothing
                Set Schedule80CDtls_G = CreateObject("Scripting.Dictionary")
   
   End If
   Next
   
              jsonDictionary.add "Schedule80CDtls", Schedule80CDtls
              Set Schedule80CDtls = Nothing
              Set Schedule80CDtls = CreateObject("Scripting.Dictionary")
   
   
'End If
              
TotalAmt_G = Sheet19.Range("TotAmount.80C").value
jsonDictionary("TotalAmt") = UVCase(TotalAmt_G)

Set Schedule80C = jsonDictionary
End If  'Malli---23/04/2025
End Function
'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80CCC() As Object
Dim i As Long
Dim mIntCells80CCC As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80CCCDtls_G, Schedule80CCCDtls
Set Schedule80CCCDtls_G = CreateObject("Scripting.Dictionary")
 
Set Schedule80CCCDtls = CreateObject("Scripting.Dictionary")
Set Schedule80CCCDtls = New Collection

Dim InsurerName_G, PolicyDocNo_G, Amount_G, TotalAmt_G
subProcCaption = "80CCC"

  
If Sheet19.Range("TotAmount.80CCC").value > 0 Then
     mIntCells80CCC = Sheet19.Range("Name_of_insurer.80CCC").count
  
   For i = 1 To mIntCells80CCC
   If (Sheet19.Range("Name_of_insurer.80CCC").Cells(i, 1).value <> "" And Sheet19.Range("Name_of_insurer.80CCC").Cells(i, 1).value <> "(Select)") Then
   
   InsurerName_G = Sheet19.Range("Name_of_insurer.80CCC").Cells(i, 1).value
   PolicyDocNo_G = Sheet19.Range("Policy_Document_Number.80CCC").Cells(i, 1).value
   Amount_G = Sheet19.Range("Amount.80CCC").Cells(i, 1).value
   
    
   
   If InsurerName_G <> "" Then
    Schedule80CCCDtls_G("InsurerName") = UCase(Trim(InsurerName_G))
   Else
   Schedule80CCCDtls_G("InsurerName") = ""
   End If
   
   If PolicyDocNo_G <> "" Then
    Schedule80CCCDtls_G("PolicyDocNo") = UCase(Trim(PolicyDocNo_G))
   Else
   Schedule80CCCDtls_G("PolicyDocNo") = ""
   End If
    
      Schedule80CCCDtls_G("Amount") = UVCase(Amount_G)
    
  ' End If
                Schedule80CCCDtls.add Schedule80CCCDtls_G
                Set Schedule80CCCDtls_G = Nothing
                Set Schedule80CCCDtls_G = CreateObject("Scripting.Dictionary")
   
   End If
   Next
   
              jsonDictionary.add "Schedule80CCCDtls", Schedule80CCCDtls
              Set Schedule80CCCDtls = Nothing
              Set Schedule80CCCDtls = CreateObject("Scripting.Dictionary")
   
   
'End If
              
TotalAmt_G = Sheet19.Range("TotAmount.80CCC").value
jsonDictionary("TotalAmt") = UVCase(TotalAmt_G)

Set Schedule80CCC = jsonDictionary
End If  'Malli 23/04/2025
End Function
'Konda commented on 29-01-2026
''Malli_80EE_AY_2025_26-------09/04/2025
'Function Schedule24B() As Object
'Dim i As Long
'Dim mIntCells24B As Variant
'Dim jsonDictionary
'Set jsonDictionary = CreateObject("Scripting.Dictionary")
'
'
'
'Dim ScheduleUs24BDtls_G, ScheduleUs24BDtls
'
'Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
'
'Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
'Set ScheduleUs24BDtls = New Collection
'
'subProcCaption = "24B"
'
'Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
'
'If Sheet21.Range("TotAmt.24b").value > 0 Then
' mIntCells24B = Sheet21.Range("LoanfrmBankOrInstitute.24b").count
'
'  For i = 1 To mIntCells24B
'   If (Sheet21.Range("LoanfrmBankOrInstitute.24b").Cells(i, 1).value <> "" And Sheet21.Range("LoanfrmBankOrInstitute.24b").Cells(i, 1).value <> "(Select)") Then
'
'   LoanTknFrom_G = Sheet21.Range("LoanfrmBankOrInstitute.24b").Cells(i, 1).value
'   'AY_25_26_V0.7
'  ' IFSCCode_G = Sheet21.Range("IFSC.24b").Cells(i, 1).value
'   'PAN_G = Sheet21.Range("PAN.24b").Cells(i, 1).value
'   '---------------------------
'   BankOrInstnName_G = Sheet21.Range("bankName.24b").Cells(i, 1).value
'
'   LoanAccNoOfBankOrInstnRefNo_G = Sheet21.Range("loanAccNum.24b").Cells(i, 1).value
'   DateofLoan_G = Sheet21.Range("loanDate.24b").Cells(i, 1).value
'   TotalLoanAmt_G = Sheet21.Range("loanAmt.24b").Cells(i, 1).value
'   LoanOutstndngAmt_G = Sheet21.Range("loanOutstanding.24b").Cells(i, 1).value
'   InterestPaid_G = Sheet21.Range("Intrst.24b").Cells(i, 1).value
'
'
'   If LoanTknFrom_G <> "" Then
'            If LoanTknFrom_G = "Bank" Then
'            LoanTknFrom_G = "B"
'            ElseIf LoanTknFrom_G = "Other than bank" Then
'            LoanTknFrom_G = "I"
'            End If
'        ScheduleUs24BDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
'   Else
'        ScheduleUs24BDtls_G("LoanTknFrom") = ""
'   End If
'   'AY_25_26_V0.7
''   If IFSCCode_G <> "" Then
''   ScheduleUs24BDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
''   End If
''
''   If PAN_G <> "" Then
''   ScheduleUs24BDtls_G("PAN") = UCase(Trim(PAN_G))
''   End If
'   '-------------------
'   If BankOrInstnName_G <> "" Then
'   ScheduleUs24BDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
'   Else
'   ScheduleUs24BDtls_G("BankOrInstnName") = ""
'   End If
'
'
'
'   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
'   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
'   Else
'   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
'   End If
'
'   If DateofLoan_G <> "" Then
'   ScheduleUs24BDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
'   Else
'   ScheduleUs24BDtls_G("DateofLoan") = ""
'   End If
'
'
'   ScheduleUs24BDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
'   ScheduleUs24BDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
'   'ScheduleUs24BDtls_G("InterestPaid") = UVCase(InterestPaid_G)
'    ScheduleUs24BDtls_G("InterestUs24B") = UVCase(InterestPaid_G)
''End If
'                ScheduleUs24BDtls.add ScheduleUs24BDtls_G
'                Set ScheduleUs24BDtls_G = Nothing
'                Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
'  End If
'  Next
'              jsonDictionary.add "ScheduleUs24BDtls", ScheduleUs24BDtls
'                Set ScheduleUs24BDtls = Nothing
'                Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
''End If
'
'TotalInterestPaid_G = Sheet21.Range("TotAmt.24b").value
''jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
'jsonDictionary("TotalInterestUs24B") = UVCase(TotalInterestPaid_G)
'
'Set Schedule24B = jsonDictionary
'End If  'Malli 23/04/2025
'
'End Function
'Konda_EA10_13A__AY_2025_26-------08-05-2025
Function ScheduleEA10_13A() As Object
subProcCaption = "ScheduleEA10_13A"
noOfProcessSub = 10
Dim jsonDictionary

Dim Placeofwork_G, ActlHRARecv_G, ActlRentPaid_G, DtlsSalUsSec171_G, BasicSalary_G, DearnessAllwnc_G, ActlRentPaid10Per_G, Sal40Or50Per_G, EligbleExmpAllwncUs13A_G

Set jsonDictionary = CreateObject("Scripting.Dictionary")
   
   Placeofwork_G = Sheet15.Range("Sch10of13A_PlaceofWrk").Cells(1, 1).value

   ActlHRARecv_G = Sheet15.Range("Sch10of13A_ActlHRArecivedA").value

   ActlRentPaid_G = Sheet15.Range("Sch10of13A_ActlRentpaid").value
   DtlsSalUsSec171_G = Sheet15.Range("Sch10of13A_DetlsofSalpersec17of1").value
   BasicSalary_G = Sheet15.Range("Sch10of13A_BasicSalary").value
   DearnessAllwnc_G = Sheet15.Range("Sch10of13A_DearAllowance").value
   ActlRentPaid10Per_G = Sheet15.Range("Sch10of13A_Actlrentpaid10persalaryB").value
   Sal40Or50Per_G = Sheet15.Range("Sch10of13A_50Por40Pofsalary").value
   EligbleExmpAllwncUs13A_G = Sheet15.Range("Sch10of13A_ElgiblExmptAllwnce10of13A").value
  
  
  
Dim Placeofwork_EA As Variant
Placeofwork_EA = Mid(Placeofwork_G, 1, 1)

If IsNumeric(Placeofwork_EA) Then

If Placeofwork_EA = "1" Or Placeofwork_EA = "2" Then
jsonDictionary("Placeofwork") = Placeofwork_EA

End If
   
 
 If ActlHRARecv_G <> "" Then
 jsonDictionary("ActlHRARecv") = UVCase(ActlHRARecv_G)
 End If

 
If ActlRentPaid_G <> "" Then
jsonDictionary("ActlRentPaid") = UVCase(ActlRentPaid_G)

End If

If DtlsSalUsSec171_G <> "" Then
jsonDictionary("DtlsSalUsSec171") = UVCase(DtlsSalUsSec171_G)

End If

If BasicSalary_G <> "" Then
jsonDictionary("BasicSalary") = UVCase(BasicSalary_G)

End If

If DearnessAllwnc_G <> "" Then
jsonDictionary("DearnessAllwnc") = UVCase(DearnessAllwnc_G)

End If

If ActlRentPaid10Per_G <> "" Then
jsonDictionary("ActlRentPaid10Per") = UVCase(ActlRentPaid10Per_G)

End If

If Sal40Or50Per_G <> "" Then
jsonDictionary("Sal40Or50Per") = UVCase(Sal40Or50Per_G)

End If

If EligbleExmpAllwncUs13A_G <> "" Then
jsonDictionary("EligbleExmpAllwncUs13A") = UVCase(EligbleExmpAllwncUs13A_G)

End If

Set ScheduleEA10_13A = jsonDictionary

End If


End Function












