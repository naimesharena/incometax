Attribute VB_Name = "PreFillJson"

Option Explicit
'06/12/2021
'variable issue start
Dim end_Exempt1, end_Exempt2, rDefinedgridNameRange As Variant
'variable issue end

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


'Malli_30/05/2025
Public incDeductionsOthIncCPC_TAX_chk As Boolean
Public incDeductionsOthIncCPC_TAX_RowCnt As Variant
'---------------------------------------------------



Sub ImportPrefill()
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

If jsonObject.exists("ITR") Then
imported = 0
fmsgbox ("Please select a valid prefill json.")
Exit Sub
End If


ImportPersonalInfo_pfl (jsonText)
ImportFilingStatus_pfl (jsonText)
ImportPartA_139_8A (jsonText)
'commented by Chetan C M AY 2025-26 on 21/01/2025
'ImportRefund_pfl (jsonText)
ImportRefund_pfl_old (jsonText)
ImportVerification_pfl (jsonText)

ImportScheduleTCS_pfl (jsonText)
ImportScheduleIT_pfl (jsonText)
'ImportTDSonSalary_pfl (jsonText)
ImportTDSOthThanSals_pfl (jsonText)
ImportScheduleTDS3Dtls_pfl (jsonText)
ImportBPGst_pfl (jsonText)
Import_IncomeDeductions (jsonText)
ImportTDSonSalary_pfl (jsonText)  'Malli_AY_2026-27  SIT-110550

'ImportScheduleBP_pfl (jsonText)



'Malli_AY_2026-27  22/01/2026
ImportScheduleHP_Pfl (jsonText) 'malli_29/01/2026


'Malli_AY_2026-27  03/04/2026
If (Sheet8.Range("bacValue").value) = 2 Then
ImportSchedule80GD_pfl (jsonText)
Importschedule80D_pfl (jsonText)
'ImportSchedule80DD_80U_pfl (jsonText)
ImportSchedule_80E_Pfl (jsonText)
ImportSchedule_80EE_pfl (jsonText)
ImportSchedule_80EEA_pfl (jsonText)
ImportSchedule_80EEB_pfl (jsonText)
 
ImportscheduleEA10_13A_pfl (jsonText)
End If
'----------------------------


'----------------------------

Application.EnableEvents = True
End Sub

Private Function DecodeBase64(ByVal strData As String) As Byte()

 

    Dim objXML As Object
    Dim objNode As Object
   
    ' help from MSXML
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = strData
    DecodeBase64 = objNode.nodeTypedValue
   
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

 

End Function

Function ImportPersonalInfo_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim sCountry, iCountry As Variant
Dim sState, iState, status As Variant
Dim YYYY, MM, DD, strDate As String

Dim DateOFFormOrIncorp, localityOrArea, countryCodeMobileNoSec, mobileNoSec
Dim residenceNo, cityOrTownOrDistrict, countryCodeMobile, mobileNo, emailAddress, emailAddressSec, stDcode, phoneNo, roadOrStreet, PinCode, residenceName, zipCode
Dim residentialStatus, firstName, surNameOrOrgName, middleName, natureOfEmp
Dim dob, aadhaarCardNo, PAN

Set jsonObject = ParseJson(jsonText)

'MsgBox jsonObject
'MsgBox jsonText


'DateOFFormOrIncorp = jsonObject("personalInfo")("orgFirmInfo")("DateOFFormOrIncorp")

localityOrArea = jsonObject("personalInfo")("address")("localityOrArea")
countryCodeMobileNoSec = jsonObject("personalInfo")("address")("countryCodeMobileNoSec")
mobileNoSec = jsonObject("personalInfo")("address")("mobileNoSec")
residenceNo = jsonObject("personalInfo")("address")("residenceNo")

cityOrTownOrDistrict = jsonObject("personalInfo")("address")("cityOrTownOrDistrict")
countryCodeMobile = jsonObject("personalInfo")("address")("countryCodeMobile")
mobileNo = jsonObject("personalInfo")("address")("mobileNo")
emailAddress = jsonObject("personalInfo")("address")("emailAddress")
stDcode = jsonObject("personalInfo")("address")("phone")("stDcode")
phoneNo = jsonObject("personalInfo")("address")("phone")("phoneNo")
iCountry = jsonObject("personalInfo")("address")("countryCode")
sCountry = Findtext(CStr(iCountry), "CountList")
roadOrStreet = jsonObject("personalInfo")("address")("roadOrStreet")
PinCode = jsonObject("personalInfo")("address")("pinCode")
iState = jsonObject("personalInfo")("address")("stateCode")
sState = Findtext(CStr(iState), "State")
zipCode = jsonObject("personalInfo")("address")("zipCode")
residenceName = jsonObject("personalInfo")("address")("residenceName")
emailAddressSec = jsonObject("personalInfo")("address")("emailAddressSecondary")

firstName = jsonObject("personalInfo")("assesseeName")("firstName")
surNameOrOrgName = jsonObject("personalInfo")("assesseeName")("surNameOrOrgName")
middleName = jsonObject("personalInfo")("assesseeName")("middleName")

natureOfEmp = jsonObject("lastFiledITR")("employerCategory")

dob = jsonObject("personalInfo")("dob")
If dob <> "" Then
    YYYY = Mid(dob, 1, 4)
    MM = Mid(dob, 6, 2)
    DD = Mid(dob, 9, 2)
    strDate = DD & "/" & MM & "/" & YYYY
End If

aadhaarCardNo = jsonObject("personalInfo")("aadhaarCardNo")
aadhaarCardNo = StrConv(DecodeBase64(aadhaarCardNo), vbUnicode)
PAN = jsonObject("personalInfo")("pan")

'residentialStatus = jsonObject("personalInfo")("filingStatus")("residentialStatus")
'residentialStatus = Findtext(residentialStatus & " ", "ResStatus")
'
'Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("sheet1.ResidentialStatus1").value = residentialStatus
''Sheet1.Protect Password:=getmsgstate

status = jsonObject("personalInfo")("status")

    If status = "I" Or status = "IND" Then
        status = "I - INDIVIDUAL"
    ElseIf status = "H" Or UCase(status) = "HUF" Then
        status = "H - HUF"
    ElseIf status = "F" Or UCase(status) = "FIRM" Then
        status = "F - FIRM(Other than LLP)"
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

        If surNameOrOrgName <> "" Then
            Sheet1.Range("sheet1.SurNameOrOrgName").value = surNameOrOrgName
        End If
         Sheet1.Unprotect Password:=getmsgstate

        If PAN <> "" Then
            Sheet1.Range("sheet1.PAN").value = PAN
        End If
         Sheet1.Unprotect Password:=getmsgstate

        If strDate <> "" Then
            Sheet1.Range("sheet1.DOB").value = strDate
            strDate = ""
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
        If sCountry <> "" And Sheet1.Range("sheet1.Country").Locked = False Then
            Sheet1.Range("sheet1.Country").value = sCountry
        End If
        If sState <> "" And Sheet1.Range("sheet1.StateCode1").Locked = False Then
            Sheet1.Range("sheet1.StateCode1").value = sState
        End If
        If PinCode <> "" And Sheet1.Range("sheet1.PinCode").Locked = False Then
            Sheet1.Range("sheet1.PinCode").value = PinCode
        End If
        If zipCode <> "" And Sheet1.Range("sheet1.ZipCode").Locked = False Then
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
        Sheet1.Unprotect Password:=getmsgstate
        If status <> "" Then
            Sheet1.Range("sheet1.Status").value = status
        End If
        Sheet1.Protect Password:=getmsgstate
        If stDcode <> "" And Sheet1.Range("sheet1.STDcode").Locked = False Then
            Sheet1.Range("sheet1.STDcode").value = stDcode
        End If
        If phoneNo <> "" And Sheet1.Range("sheet1.PhoneNo").Locked = False Then
            Sheet1.Range("sheet1.PhoneNo").value = phoneNo
        End If
        If countryCodeMobileNoSec <> "" And Sheet1.Range("sheet1.mobileCountryCode2").Locked = False Then
            Sheet1.Range("sheet1.mobileCountryCode2").value = countryCodeMobileNoSec
        End If
        If mobileNoSec <> "" And Sheet1.Range("sheet1.MobileNoSec").Locked = False Then
            Sheet1.Range("sheet1.MobileNoSec").value = mobileNoSec
        End If
        If emailAddressSec <> "" And Sheet1.Range("sheet1.EmailAddress2").Locked = False Then
            Sheet1.Range("sheet1.EmailAddress2").value = emailAddressSec
        End If
        'issue 1 solved
        If aadhaarCardNo <> "" And Sheet1.Range("Sheet1.Aadhaar").Locked = False Then
            Sheet1.Range("Sheet1.Aadhaar").value = aadhaarCardNo
        End If
        'issue 1 solved
        
    'issue 2 solved
     

        If natureOfEmp <> "" And Sheet1.Range("sheet1.NatureofEmployment").Locked = False Then
            If natureOfEmp = "CGOV" Then
                natureOfEmp = "Central Government"
            ElseIf natureOfEmp = "SGOV" Then
                natureOfEmp = "State Government"
            '============
            ElseIf natureOfEmp = "SCJ" Then
                natureOfEmp = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958"
   '===========================
            ElseIf natureOfEmp = "PSU" Then
                natureOfEmp = "Public Sector Undertaking"
            ElseIf natureOfEmp = "OTH" Then
                natureOfEmp = "Others"
            ElseIf natureOfEmp = "PE" Then
                natureOfEmp = "Pensioners-CG"
            ElseIf natureOfEmp = "PESG" Then
                natureOfEmp = "Pensioners-SG"
            ElseIf natureOfEmp = "PEPS" Then
                natureOfEmp = "Pensioners-PSU"
            ElseIf natureOfEmp = "PEO" Then
                natureOfEmp = "Pensioners-Other"
            ElseIf natureOfEmp = "NA" Then
                natureOfEmp = "Not Applicable (eg. Family pension etc)"
            End If
    
            Sheet1.Range("sheet1.NatureofEmployment").value = natureOfEmp
        End If
        

    'issue 2 solved
     
End Function

 

Function ImportVerification_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim FatherName, AssesseeVerName, AssesseeVerPAN, Capacity, Place As Variant


Set jsonObject = ParseJson(jsonText)
 
 FatherName = jsonObject("verification")("declaration")("fatherName")
 AssesseeVerName = jsonObject("verification")("declaration")("assesseeVerName")
 AssesseeVerPAN = jsonObject("verification")("declaration")("assesseeVerPAN")
 
 Capacity = jsonObject("verification")("capacity")
 
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

        If FatherName <> "" And Sheet5.Range("Ver.FatherName").Locked = False Then
            Sheet5.Range("Ver.FatherName").value = FatherName
        End If
        
        If AssesseeVerName <> "" And Sheet5.Range("Ver.AssesseeVerName").Locked = False Then
            Sheet5.Range("Ver.AssesseeVerName").value = AssesseeVerName
        End If
        
        If AssesseeVerPAN <> "" And Sheet5.Range("Ver.PAN").Locked = False Then
            Sheet5.Range("Ver.PAN").value = AssesseeVerPAN
        End If
            
         Application.EnableEvents = False  'Malli_SIT-110552
         
        If Capacity <> "" And Sheet5.Range("Ver.capacity").Locked = False Then
            Sheet5.Range("Ver.capacity").value = Capacity
        End If
        
        Application.EnableEvents = True  'Malli_SIT-110552
End Function
Function ImportRefund_pfl_old(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist
Dim IFSCColNo, BankNameColNo, ACCNOColNo, AccountTypeColNo As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim NodeMain, NodeListMain, NodeTemp

Set jsonObject = ParseJson(jsonText)



Set Nodelist = New Collection

If jsonObject.exists("bankAccountDtls") Then
Set NodeListMain = jsonObject("bankAccountDtls")
Dim flag
If NodeListMain <> Empty Then
For Each NodeMain In NodeListMain
    For Each node In NodeMain("addtnlBankDetails")
        flag = 0
        For Each NodeTemp In Nodelist
            If NodeTemp("bankAccountNo") = node("bankAccountNo") Then
                flag = 1
            End If
        Next NodeTemp
        If flag = 0 Then
            Nodelist.add node
        End If
        
    Next node
Next NodeMain
End If
End If

'AY_2025_26 secoundary prefill source removed
'Malli_AY_2025_26 Comented
'Set NodeListMain = jsonObject("lastFiledITR")("bankAccountDtls")
'If NodeListMain <> Empty Then
'For Each NodeMain In NodeListMain
'    For Each node In NodeMain("addtnlBankDetails")
'        flag = 0
'        For Each NodeTemp In Nodelist
'            If NodeTemp("bankAccountNo") = node("bankAccountNo") Then
'                flag = 1
'            End If
'        Next NodeTemp
'        If flag = 0 Then
'            Nodelist.add node
'        End If
'
'    Next node
'Next NodeMain
'
'End If

'-------------------------------------------------

  
'issue 14 solved

'uncommented by Chetan C M AY 2025-26 on 21/01/2025
'--start
    Set NodeListMain = jsonObject("bankAccountDtls")
    For Each NodeMain In NodeListMain
        Set Nodelist = NodeMain("addtnlBankDetails")
'--end
'        Set NodeListMain2 = jsonObject("lastFiledITR")("bankAccountDtls")
'        For Each NodeList2 In NodeListMain2
'            For Each Node2 In NodeList2
'                For Each Node In NodeList
'                    if Node("")
'                Next Node
'            Next Node2
'        Next NodeMain2
        IFSCColNo = Sheet5.Range("SchBA.IFSC").Column
        BankNameColNo = Sheet5.Range("SchBA.BankName").Column
        ACCNOColNo = Sheet5.Range("SchBA.AcntNo").Column
        'uncommented by Chetan C M AY2025-26 21/01/2025
'--start--
Dim CheckBox
       CheckBox = Sheet5.Range("tempXML").Column
'--end--


        AccountTypeColNo = Sheet5.Range("SchBA.Type_of_Account").Column 'Newly added by Sai
         
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
            
            If Sheet5.Range("SchBA.Type_of_Account").Locked = False Then
                Sheet5.Range("SchBA.Type_of_Account").ClearContents
            End If
            
 'uncommented by Chetan C M AY2025-26 21/01/2025
'--start--
            If Sheet5.Range("tempxml").Locked = False Then
                Sheet5.Range("tempxml").ClearContents
            End If
'--end--
        End If
        
'        If (TotalDiffRow > 0) Then
'         AddDiffRows_BANK (TotalDiffRow)
'        End If
        
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
                
                 If Trim(node("ifsccode")) <> "" And Sheet5.Cells(rowcount, IFSCColNo).Locked = False Then
                    Sheet5.Cells(rowcount, IFSCColNo).value = node("ifsccode")
                End If
                
                 If Trim(node("bankName")) <> "" And Sheet5.Cells(rowcount, BankNameColNo).Locked = False Then
                    Sheet5.Cells(rowcount, BankNameColNo).value = node("bankName")
                End If
                
                 If Trim(node("bankAccountNo")) <> "" And Sheet5.Cells(rowcount, ACCNOColNo).Locked = False Then
                    Sheet5.Cells(rowcount, ACCNOColNo).value = node("bankAccountNo")
                End If
                
                
            '----------------------For Prefill 2024-25 implement
            
'    If Sheet5.Range("SchBA.Type_of_Account").item(i).value <> "" Then
'        If Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Savings Account" Then
'            AddtnlBank("AccountType") = "SB"
'               ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Current Account" Then
'               AddtnlBank("AccountType") = "CA"
'                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Cash Credit Account" Then
'               AddtnlBank("AccountType") = "CC"
'                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Over draft account" Then
'               AddtnlBank("AccountType") = "OD"
'
'                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Non Resident Account" Then
'               AddtnlBank("AccountType") = "NRO"
'                ElseIf Trim(Sheet5.Range("SchBA.Type_of_Account").item(i).value) = "Other" Then
'               AddtnlBank("AccountType") = "OTH"
'
'        End If
'    End If
            
            
            
    If node("AccountType") <> "" Then
        If Trim(node("AccountType")) = "SB" Then
             Sheet5.Cells(rowcount, AccountTypeColNo).value = "Savings Account"
        ElseIf Trim(node("AccountType")) = "CA" Then
                Sheet5.Cells(rowcount, AccountTypeColNo).value = "Current Account"
        ElseIf Trim(node("AccountType")) = "CC" Then
                Sheet5.Cells(rowcount, AccountTypeColNo).value = "Cash Credit Account"
        ElseIf Trim(node("AccountType")) = "OD" Then
                Sheet5.Cells(rowcount, AccountTypeColNo).value = "Over draft account"
               
        ElseIf Trim(node("AccountType")) = "NRO" Then
               Sheet5.Cells(rowcount, AccountTypeColNo).value = "Non Resident Account"
        ElseIf Trim(node("AccountType")) = "OTH" Then
                Sheet5.Cells(rowcount, AccountTypeColNo).value = "Other"
        End If
    End If
            
            'uncommented by Chetan C M AY2025-26 21/01/2025
'--start--
                Dim useForRefund
                useForRefund = node("useForRefund")
                If useForRefund = "" Then
                    Sheet5.Cells(rowcount, CheckBox).value = False
                ElseIf useForRefund = "N" Then
                    Sheet5.Cells(rowcount, CheckBox).value = False
                ElseIf useForRefund = "false" Then
                    Sheet5.Cells(rowcount, CheckBox).value = False
                ElseIf useForRefund = "False" Then
                    Sheet5.Cells(rowcount, CheckBox).value = False
                Else
                    Sheet5.Cells(rowcount, CheckBox).value = True
                End If
'--end--

 'issue 14 solved
 
                If CheckBox = True Then
                    CheckBox = "true"
                ElseIf CheckBox = False Then
                    CheckBox = "false"
                ElseIf CheckBox = "" Then
                    CheckBox = "false"
                End If
                
            LinkCheckBoxes
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
'uncommented by Chetan C M AY2025-26 21/01/2025
   Next NodeMain
End Function

Function ImportScheduleTCS_pfl(jsonText As String)
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
  
   Set Nodelist = jsonObject("form26as")("scheduleTCS")("tcs")
    
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
       
            If Trim(node("employerOrDeductorOrCollectDetl")("tan")) <> "" And Sheet17.Cells(rowcount, TANColNo).Locked = False Then
                Sheet17.Cells(rowcount, TANColNo).value = node("employerOrDeductorOrCollectDetl")("tan")
            End If
            If Trim(node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")) <> "" And Sheet17.Cells(rowcount, DEDNameColNo).Locked = False Then
                Sheet17.Cells(rowcount, DEDNameColNo).value = node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
            If Trim(node("amtfrm26AS")) <> "" And Sheet17.Cells(rowcount, Amtfrom26ASColNo).Locked = False Then
                Sheet17.Cells(rowcount, Amtfrom26ASColNo).value = node("amtfrm26AS")
            End If
            
'issue 12 solved
            If Trim(node("amtTCSClaimedThisYear")) <> "" And Sheet17.Cells(rowcount, TaxColNo).Locked = False Then
                Sheet17.Cells(rowcount, TaxColNo).value = node("amtTCSClaimedThisYear")
            End If


          '  If Trim(node("amtTCSClaimedThisYear")) <> "" And Sheet17.Cells(rowcount, ClaimColNo).Locked = False Then
           '     Sheet17.Cells(rowcount, ClaimColNo).value = node("amtTCSClaimedThisYear")
            'End If

            'SIT-67955 AY_2024_25    'Malli
            
            If Trim(node("amtTCSClaimedThisYear")) <> "" And Sheet17.Cells(rowcount, ClaimColNo).Locked = False Then
                Sheet17.Cells(rowcount, ClaimColNo).value = node("amtTCSClaimedThisYear")
            End If


'issue 12 solved
            
        cnt = cnt + 1
    Next node
    RecTCS = cnt
End Function

Function ImportScheduleIT_pfl(jsonText As String)
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
   
'issue 13 solved
    Set Nodelist = jsonObject("form26as")("taxPayments")("taxPayment")
'issue 13 solved

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
        
            If Trim(node("bsrCode")) <> "" And Sheet18.Cells(rowcount, BSRCodeColNo).Locked = False Then
                Sheet18.Cells(rowcount, BSRCodeColNo).value = node("bsrCode")
            End If
            strDate = node("dateDep")
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
            If Trim(node("srlNoOfChaln")) <> "" And Sheet18.Cells(rowcount, SrlNoChallanColNo).Locked = False Then
                Sheet18.Cells(rowcount, SrlNoChallanColNo).value = node("srlNoOfChaln")
            End If
            If Trim(node("amt")) <> "" And Sheet18.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet18.Cells(rowcount, AmtColNo).value = node("amt")
            End If
      
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function

Function ImportTDSonSalary_pfl(jsonText As String)
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
    
    Set Nodelist = jsonObject("form26as")("tdsOnSalaries")("tdsOnSalary")
    
    
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
        
            If Trim(node("employerOrDeductorOrCollectDetl")("tan")) <> "" And Sheet4.Cells(rowcount, TANNoEmployer).Locked = False Then
                Sheet4.Cells(rowcount, TANNoEmployer).value = node("employerOrDeductorOrCollectDetl")("tan")
            End If
            If Trim(node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")) <> "" And Sheet4.Cells(rowcount, TDSNameOfEmployer).Locked = False Then
                Sheet4.Cells(rowcount, TDSNameOfEmployer).value = node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
            If Trim(node("incChrgSal")) <> "" And Sheet4.Cells(rowcount, TDSIncomeCharge).Locked = False Then
                Sheet4.Cells(rowcount, TDSIncomeCharge).value = node("incChrgSal")
            End If
            If Trim(node("totalTDSSal")) <> "" And Sheet4.Cells(rowcount, TDSTotalTax).Locked = False Then
                Sheet4.Cells(rowcount, TDSTotalTax).value = node("totalTDSSal")
            End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

'import json for TDSonOthThanSals

Function ImportTDSOthThanSals_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist, Nodelist2 As Variant
Dim TANColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long
Dim SectionTDS, TDSSection_TDS2, SectionTDSColNo As Variant '30/04/2025

Set jsonObject = ParseJson(jsonText)

Set Nodelist = New Collection
Set Nodelist = jsonObject("form26as")("tdsOnOthThanSals")("tdSonOthThanSal")
If Nodelist <> Empty Then
Set Nodelist2 = jsonObject("lastFiledITR")("tdsOnOthThanSals")("tdSonOthThanSal")
    For Each node In Nodelist2
        Nodelist.add node
    Next node
Else
    Set Nodelist = jsonObject("lastFiledITR")("tdsOnOthThanSals")("tdSonOthThanSal")
End If
    TANColNo = Sheet4.Range("TDsOthr.TAN").Column
    SectionTDSColNo = Sheet4.Range("TDsOthr.SectionTDS").Column   '30/04/2025
    FYColNo = Sheet4.Range("TDsOthr.finyear").Column
    TDSOtherColNo = Sheet4.Range("TDsOthr.tds").Column
    TDSDeductedColNo = Sheet4.Range("TDsOthr.tdsdeducted").Column
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
        If Sheet4.Range("TDsOthr.SectionTDS").Locked = False Then  '30/04/2025
            Sheet4.Range("TDsOthr.SectionTDS").ClearContents
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
            
            If Trim(node("employerOrDeductorOrCollectDetl")("tan")) <> "" And Sheet4.Cells(rowcount, TANColNo).Locked = False Then
                Sheet4.Cells(rowcount, TANColNo).value = node("employerOrDeductorOrCollectDetl")("tan")
            End If
            

'            If Trim(node("taxDeductCreditDtls")("taxDeductedOwnHands")) <> "" And Sheet4.Cells(rowcount, TDSDeductedColNo).Locked = False Then
'                Sheet4.Cells(rowcount, TDSDeductedColNo).value = node("taxDeductCreditDtls")("taxDeductedOwnHands")
'            End If
            
           'Malli------------30/04/2025-------
            'If Trim(node("sectionCode")) <> "" And Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
                    
                        SectionTDS = node("sectionCode")
                        'Debug.Print node("sectionCode")
                        
                        'If SectionTDS = "92A" Then
                        If UCase(SectionTDS) = UCase("92A") Then
                        TDSSection_TDS2 = "192-Salary-Payment to Government employees other than Indian Government employees"
                        
                        ElseIf UCase(SectionTDS) = UCase("92B") Then
                        TDSSection_TDS2 = "192-Salary-Payment to employees other than Government employees"
                        
                        ElseIf UCase(SectionTDS) = UCase("92C") Then
                        TDSSection_TDS2 = "192-Salary-Payment to Indian Government employees"
                        
                        '13/05/2025
                        'ElseIf SectionTDS = "192A" Then
                         'ElseIf SectionTDS = "2AA" Or SectionTDS = "2aa" Or UCase(SectionTDS) = UCase("2AA") Then
                         ElseIf UCase(SectionTDS) = UCase("2AA") Then
                        TDSSection_TDS2 = "192A-TDS on PF withdrawal"
                        '-------------------------------------
                        ElseIf UCase(SectionTDS) = UCase("193") Then
                        TDSSection_TDS2 = "193-Interest on Securities"
                        
                        ElseIf UCase(SectionTDS) = UCase("194") Then
                        TDSSection_TDS2 = "194-Dividends"
                        
                        ElseIf UCase(SectionTDS) = UCase("94A") Then
                        TDSSection_TDS2 = "194A-Interest other than 'Interest on securities'"
                        
                        ElseIf UCase(SectionTDS) = UCase("94B") Then
                        TDSSection_TDS2 = "194B-Winning from lottery or crossword puzzle"
                        
                        '13/05/2025
                        'ElseIf SectionTDS = "94BA" Then
                           'ElseIf SectionTDS = "9BA" Or SectionTDS = "9ba" Or UCase(SectionTDS) = "9BA" Then
                           ElseIf UCase(SectionTDS) = UCase("9BA") Then
                        TDSSection_TDS2 = "194BA-Winnings from online games"
                        '--------------------------
                        ElseIf UCase(SectionTDS) = UCase("4BB") Then
                        TDSSection_TDS2 = "194BB-Winning from horse race"
                        
                        ElseIf UCase(SectionTDS) = UCase("94C") Then
                        TDSSection_TDS2 = "194C-Payments to contractors and sub-contractors"
                        
                        ElseIf UCase(SectionTDS) = UCase("94D") Then
                        TDSSection_TDS2 = "194D-Insurance commission"
                        
                        ElseIf UCase(SectionTDS) = UCase("4DA") Then
                        TDSSection_TDS2 = "194DA-Payment in respect of life insurance policy"
                        
                        ElseIf UCase(SectionTDS) = UCase("94E") Then
                        TDSSection_TDS2 = "194E-Payments to non-resident sportsmen or sports associations"
                        
                        ElseIf UCase(SectionTDS) = UCase("4EE") Then
                        TDSSection_TDS2 = "194EE-Payments in respect of deposits under National Savings"
                        
                        '13/05/2025---------------------
                       ' ElseIf SectionTDS = "4F" Then
                         'ElseIf SectionTDS = "94F" Or SectionTDS = "94f" Or UCase(SectionTDS) = UCase("94F") Then
                         ElseIf UCase(SectionTDS) = UCase("94F") Then
                        TDSSection_TDS2 = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
                        
                       ' ElseIf SectionTDS = "4G" Then
                         'ElseIf SectionTDS = "94G" Or SectionTDS = "94g" Or UCase(SectionTDS) = "94G" Then
                         ElseIf UCase(SectionTDS) = UCase("94G") Then
                        TDSSection_TDS2 = "194G-Commission, price, etc. on sale of lottery tickets"
                        
                        
                        'ElseIf SectionTDS = "4H" Then
                           'ElseIf SectionTDS = "94H" Or SectionTDS = "94h" Or UCase(SectionTDS) = "94H" Then
                           ElseIf UCase(SectionTDS) = UCase("94H") Then
                        TDSSection_TDS2 = "194H-Commission or brokerage"
                        
                        
                        'ElseIf SectionTDS = "4-IA" Then
                         ' ElseIf SectionTDS = "4IA" Or SectionTDS = "4ia" Or UCase(SectionTDS) = "4IA" Then
                          ElseIf UCase(SectionTDS) = UCase("4IA") Then
                        TDSSection_TDS2 = "194I(a)-Rent on hiring of plant and machinery"
                        
                        'ElseIf SectionTDS = "4-IB" Then
                          'ElseIf SectionTDS = "4IB" Or SectionTDS = "4ib" Or UCase(SectionTDS) = "4IB" Then
                          ElseIf UCase(SectionTDS) = UCase("4IB") Then
                        TDSSection_TDS2 = "194I(b)-Rent on other than plant and machinery"
                        
                         'ElseIf SectionTDS = "4IA" Then
                          'ElseIf SectionTDS = "9IA" Or SectionTDS = "9ia" Or UCase(SectionTDS) = "9IA" Then
                           ElseIf UCase(SectionTDS) = UCase("9IA") Then
                        TDSSection_TDS2 = "194IA-TDS on Sale of immovable property"
                        
                        'ElseIf SectionTDS = "4IB" Then
                          'ElseIf SectionTDS = "9IB" Or SectionTDS = "9ib" Or UCase(SectionTDS) = "9IB" Then
                           ElseIf UCase(SectionTDS) = UCase("9IB") Then
                        TDSSection_TDS2 = "194IB-Payment of rent by certain individuals or Hindu undivided"
                        '---------------------------
                        
                        ElseIf UCase(SectionTDS) = UCase("4IC") Then
                        TDSSection_TDS2 = "194IC-Payment under specified agreement"
                        
                        '13/05/2025
                        'ElseIf SectionTDS = "94J-A" Then
                         'ElseIf SectionTDS = "4JA" Or SectionTDS = "4ja" Or UCase(SectionTDS) = "4JA" Then
                         ElseIf UCase(SectionTDS) = UCase("4JA") Then
                        TDSSection_TDS2 = "194J(a)-Fees for technical services"
                        
                        'ElseIf SectionTDS = "94J-B" Then
                        'ElseIf SectionTDS = "4JB" Or SectionTDS = "4jb" Or UCase(SectionTDS) = "4JB" Then
                         ElseIf UCase(SectionTDS) = UCase("4JB") Then
                        TDSSection_TDS2 = "194J(b)-Fees for professional  services or royalty etc"
                        '-------------
                        
                        ElseIf UCase(SectionTDS) = UCase("94K") Then
                        TDSSection_TDS2 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
                        
                        ElseIf UCase(SectionTDS) = UCase("4LA") Then
                        TDSSection_TDS2 = "194LA-Payment of compensation on acquisition of certain immovable"
                        
                        ElseIf UCase(SectionTDS) = UCase("4LB") Then
                        TDSSection_TDS2 = "194LB-Income by way of Interest from Infrastructure Debt fund"
                        
                        '13/05/2025
                       ' ElseIf SectionTDS = "4LC1" Then
                         'ElseIf SectionTDS = "LC1" Or SectionTDS = "lc1" Or UCase(SectionTDS) = "LC1" Then
                         ElseIf UCase(SectionTDS) = UCase("LC1") Then
                        TDSSection_TDS2 = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
                        
                        'ElseIf SectionTDS = "4LC2" Then
                         'ElseIf SectionTDS = "LC2" Or SectionTDS = "lc2" Or UCase(SectionTDS) = UCase("LC2") Then
                         ElseIf UCase(SectionTDS) = UCase("LC2") Then
                        TDSSection_TDS2 = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
                        
                        'ElseIf SectionTDS = "4LC3" Then
                         'ElseIf SectionTDS = "LC3" Or SectionTDS = "lc3" Or UCase(SectionTDS) = "Lc3" Then
                         ElseIf UCase(SectionTDS) = UCase("Lc3") Then
                        TDSSection_TDS2 = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
                        
'Malli_AY_2026_27  SIT-117189 10/04/2026
'                        ElseIf SectionTDS = "4BA1" Then
                        ElseIf SectionTDS = "BA1" Then
                        TDSSection_TDS2 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'Malli_AY_2026_27  SIT-117189 10/04/2026
'                        ElseIf SectionTDS = "4BA2" Then
                         ElseIf SectionTDS = "BA2" Then
                        TDSSection_TDS2 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"

'                        ElseIf SectionTDS = "LBA1" Then
'                        TDSSection_TDS2 = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
'
'                        ElseIf SectionTDS = "LBA2" Then
'                        TDSSection_TDS2 = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
'
                        'ElseIf SectionTDS = "LBA3" Then
                          'ElseIf SectionTDS = "BA3" Or SectionTDS = "ba3" Or UCase(SectionTDS) = "BA3" Then
                          ElseIf UCase(SectionTDS) = UCase("BA3") Then
                        TDSSection_TDS2 = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
                        '--------------------
                        
                        ElseIf UCase(SectionTDS) = UCase("LBB") Then
                        TDSSection_TDS2 = "194LBB-Income in respect of units of investment fund"
                        
                        ElseIf UCase(SectionTDS) = UCase("94R") Then
                        TDSSection_TDS2 = "194R-Benefits or perquisites of business or profession"
                        
                        ElseIf UCase(SectionTDS) = UCase("94S") Then
                        TDSSection_TDS2 = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
                        
                        '13/05/2025
                       ' ElseIf SectionTDS = "94B-P" Then
                         'ElseIf SectionTDS = "4BP" Or SectionTDS = "4bp" Or UCase(SectionTDS) = "4BP" Then
                         ElseIf UCase(SectionTDS) = UCase("4BP") Then
                        TDSSection_TDS2 = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
                        
                        'ElseIf SectionTDS = "94R-P" Then
                         'ElseIf SectionTDS = "4RP" Or SectionTDS = "4rp" Or UCase(SectionTDS) = "4RP" Then
                         ElseIf UCase(SectionTDS) = UCase("4RP") Then
                        TDSSection_TDS2 = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
                        
                        'ElseIf SectionTDS = "94S-P" Then
                         'ElseIf SectionTDS = "4SP" Or SectionTDS = "4sp" Or UCase(SectionTDS) = "4SP" Then
                         ElseIf UCase(SectionTDS) = UCase("4SP") Then
                        TDSSection_TDS2 = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
                        '----------------
                        
                        ElseIf UCase(SectionTDS) = UCase("LBC") Then
                        TDSSection_TDS2 = "194LBC-Income in respect of investment in securitization trust"
                        
                        ElseIf UCase(SectionTDS) = UCase("4LD") Then
                        TDSSection_TDS2 = "194LD-TDS on interest on bonds / government securities"
                        
                        ElseIf UCase(SectionTDS) = UCase("94M") Then
                        TDSSection_TDS2 = "194M-Payment of certain sums by certain individuals or HUF"
                        
                        ElseIf UCase(SectionTDS) = UCase("94N") Then
                        TDSSection_TDS2 = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
                        
                        '13/05/2025
                       ' ElseIf SectionTDS = "94N-F" Then
                         'ElseIf SectionTDS = "4NF" Or SectionTDS = "4nf" Or UCase(SectionTDS) = "4NF" Then
                         ElseIf UCase(SectionTDS) = UCase("4NF") Then
                        TDSSection_TDS2 = "194N-First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
                        
                        'ElseIf SectionTDS = "94N-C" Then
                         'ElseIf SectionTDS = "4NC" Or SectionTDS = "4nc" Or UCase(SectionTDS) = "4NC" Then
                         ElseIf UCase(SectionTDS) = UCase("4NC") Then
                        TDSSection_TDS2 = "194N-Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
                        
                       ' ElseIf SectionTDS = "94N-FT" Then
                          'ElseIf SectionTDS = "NFT" Or SectionTDS = "nft" Or UCase(SectionTDS) = "NFT" Then
                           ElseIf UCase(SectionTDS) = UCase("NFT") Then
                        TDSSection_TDS2 = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
                        '----------------------
                        
                        ElseIf UCase(SectionTDS) = UCase("94O") Then
                        TDSSection_TDS2 = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
                        
                        ElseIf UCase(SectionTDS) = UCase("94P") Then
                        TDSSection_TDS2 = "194P-Deduction of tax in case of specified senior citizen"
                        
                        ElseIf UCase(SectionTDS) = UCase("94Q") Then
                        TDSSection_TDS2 = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
                        
                        ElseIf UCase(SectionTDS) = UCase("195") Then
                        TDSSection_TDS2 = "195-Other sums payable to a non-resident"
                        
                        ElseIf UCase(SectionTDS) = UCase("96A") Then
                        TDSSection_TDS2 = "196A-Income in respect of units of non-residents"
                        
                        ElseIf UCase(SectionTDS) = UCase("96B") Then
                        TDSSection_TDS2 = "196B-Payments in respect of units to an offshore fund"
                        
                        ElseIf UCase(SectionTDS) = UCase("96C") Then
                        TDSSection_TDS2 = "196C-Income from foreign currency bonds or shares of Indian"
                        
                        ElseIf UCase(SectionTDS) = UCase("96D") Then
                        TDSSection_TDS2 = "196D-Income of foreign institutional investors from securities"
                        
                        '13/05/2025
                        'ElseIf SectionTDS = "96DA" Then
                         'ElseIf SectionTDS = "6DA" Or SectionTDS = "6da" Or UCase(SectionTDS) = "6DA" Then
                          ElseIf UCase(SectionTDS) = UCase("6DA") Then
                         TDSSection_TDS2 = "196D(1A)-Income of specified fund from securities"
                        
                        'ElseIf SectionTDS = "94BA-P" Then
                         'ElseIf SectionTDS = "BAP" Or SectionTDS = "bap" Or UCase(SectionTDS) = "BAP" Then
                          ElseIf UCase(SectionTDS) = UCase("BAP") Then
                        TDSSection_TDS2 = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
                        Else
                        'TDSSection_TDS2 = ""
                        TDSSection_TDS2 = "(Select)"
                        
                        End If
                        
                    If TDSSection_TDS2 <> "" And Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
                        Sheet4.Cells(rowcount, SectionTDSColNo).value = TDSSection_TDS2
                        'TDSSection_TDS2 = ""
                        SectionTDS = ""
                    End If
 
           
           
           
           
           '---------------------------
           
            If Trim(node("taxDeductCreditDtls")("taxDeductedOwnHands")) <> "" And Sheet4.Cells(rowcount, TDSDeductedColNo).Locked = False Then
                Sheet4.Cells(rowcount, TDSDeductedColNo).value = node("taxDeductCreditDtls")("taxDeductedOwnHands")
            End If
            
            If Trim(node("taxDeductCreditDtls")("taxClaimedOwnHands")) <> "" And Sheet4.Cells(rowcount, TDSClaimedColNo).Locked = False Then
                Sheet4.Cells(rowcount, TDSClaimedColNo).value = node("taxDeductCreditDtls")("taxClaimedOwnHands")
            End If
            'issue 15
            
            If Trim(node("deductedYr")) <> "" And Sheet4.Cells(rowcount, FYColNo).Locked = False Then
                Sheet4.Cells(rowcount, FYColNo).value = node("deductedYr")
            End If
            If Trim(node("broughtFwdTDSAmt")) <> "" And Sheet4.Cells(rowcount, TDSOtherColNo).Locked = False Then
                Sheet4.Cells(rowcount, TDSOtherColNo).value = node("broughtFwdTDSAmt")
            End If

            'issue 15
            If Trim(node("grossAmount")) <> "" And Sheet4.Cells(rowcount, TDSGrossAmountColNo).Locked = False Then
                Sheet4.Cells(rowcount, TDSGrossAmountColNo).value = node("grossAmount")
            End If
            
'    'Malli------------------------------
'    'Malli_AY_2025_26
'
'    If node("sectionCode") <> "" Then
'           Dim HeadOfIncome As Variant
'           Dim sectioncode
'           sectioncode = UCase(node("sectionCode"))
'
'           If sectioncode = "94N" Then
'             HeadOfIncome = "Not applicable (only in case TDS is deducted u/s 194N)"
'
'           ElseIf sectioncode = "94I(b)" Or sectioncode = "194IB" Or _
'             sectioncode = "4IB" Or sectioncode = "4Ib" Then
'             HeadOfIncome = "Income from House Property"
'
'           ElseIf sectioncode = "194" Or sectioncode = "94A" Or _
'           sectioncode = "193" Or sectioncode = "94I" Or _
'           sectioncode = "94D" Or sectioncode = "4DA" Or _
'           sectioncode = "94G" Or sectioncode = "94H" Or _
'           sectioncode = "94B" Or sectioncode = "4LB" Or _
'           sectioncode = "4LC" Or sectioncode = "4LD" Or _
'           sectioncode = "LBL" Or sectioncode = "94E" Or _
'           sectioncode = "4BB" Then
'
'               HeadOfIncome = "Income from Other Source"
'
''           ElseIf sectioncode = "94LA" Or sectioncode = "94IA" Then
''           HeadOfIncome = "Income from Capital Gains"
'            Else
'                HeadOfIncome = "(Select)"
'
'           End If
'
'         If HeadOfIncome <> "" And Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
'                Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
'            End If
'
'        End If
    
    
    
    
 'Malli uncomented-----------30/04/2025
   ' -------------------------
       ' Malli comented
       
        
            Dim HeadOfIncome As Variant
            HeadOfIncome = UCase(node("headOfIncome"))

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

            If HeadOfIncome <> "" And Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
                Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
            End If
            
'------------------------------
            
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function
Sub ImportScheduleTDS3Dtls_pfl(jsonText As String)
On Error Resume Next
 
    Dim jsonObject, jsonDictionary As Object
    Dim TotalXMLRow1 As Long
    Dim TotalDiffRow1 As Long
    Dim RecTDS11, RecTDS21, cnt As Long
    Dim Nodelist, Nodelist2 As Collection
    Dim node
    Dim InnerNode
    Dim TotalExRow1 As Long
    Dim TotalXMLRow, TotalExRow, TotalDiffRow, rowcount As Long
    Dim PANColNo, AadhaarColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant
    Dim SectionTDS, TDSSection_TDS3, SectionTDSColNo As Variant '30/04/2025
    
    Set jsonObject = ParseJson(jsonText)
    Set Nodelist = jsonObject("form26as")("scheduleTDS3Dtls")("tds3Details")
If Nodelist <> Empty Then
    Set Nodelist2 = jsonObject("lastFiledITR")("scheduleTDS3Dtls")("tds3Details")
    For Each node In Nodelist2
        Nodelist.add node
    Next node
Else
    Set Nodelist = jsonObject("lastFiledITR")("scheduleTDS3Dtls")("tds3Details")
End If
    PANColNo = Sheet4.Range("TDsOthr2.PAN").Column
    AadhaarColNo = Sheet4.Range("TDsOthr2.AADHAAR").Column
     SectionTDSColNo = Sheet4.Range("TDsOthr2.SectionTDSDeducted").Column   '30/04/2025
    FYColNo = Sheet4.Range("TDsOthr2.finyear").Column
    TDSOtherColNo = Sheet4.Range("TDsOthr2.Tds").Column
    TDSDeductedColNo = Sheet4.Range("TDsOthr2.tdsdeducted").Column
    TDSClaimedColNo = Sheet4.Range("TDsOthr2.tdsclaimed").Column
    TDSGrossAmountColNo = Sheet4.Range("TDsOthr2.grossamount").Column
    'Malli-----AY_2025_26
    'TDSHeadColNo = Sheet4.Range("TDsOthr2.headincome").Column
    TDSHeadColNo = Sheet4.Range("TDsOthr2.headamount").Column
    '-------------------------------
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
        
        If Sheet4.Range("TDsOthr2.SectionTDSDeducted").Locked = False Then   '30/04/2025
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
        'AY_2025_26
'        If Sheet4.Range("TDsOthr2.headincome").Locked = False Then
'            Sheet4.Range("TDsOthr2.headincome").ClearContents
'        End If
        If Sheet4.Range("TDsOthr2.headamount").Locked = False Then
            Sheet4.Range("TDsOthr2.headamount").ClearContents
        End If
        '------------------------------
    End If
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth1 (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet4.Range("TDsOthr2.PAN").name)
    rowcount = rowcount - 1
    cnt = 0
 
    For Each node In Nodelist
    Application.EnableEvents = True
        rowcount = rowcount + 1
            If Trim(node("grossAmount")) <> "" And Sheet4.Cells(rowcount, PANColNo).Locked = False Then
                Sheet4.Cells(rowcount, PANColNo).value = UCase(node("panOfTenant"))
            End If
            'not in json
                If Trim(node("AadhaarofTenant")) <> "" And Sheet4.Cells(rowcount, AadhaarColNo).Locked = False Then
                    Sheet4.Cells(rowcount, AadhaarColNo).value = UCase(node("AadhaarofTenant"))
                End If
 
 
 'Malli------------30/04/2025-------
            'If Trim(node("sectionCode")) <> "" And Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
                    
                        SectionTDS = node("sectionCode")
                        Debug.Print SectionTDS
                        
                        If UCase(SectionTDS) = UCase("92A") Then
                        TDSSection_TDS3 = "192-Salary-Payment to Government employees other than Indian Government employees"
                        
                        ElseIf UCase(SectionTDS) = UCase("92B") Then
                        TDSSection_TDS3 = "192-Salary-Payment to employees other than Government employees"
                        
                        ElseIf UCase(SectionTDS) = UCase("92C") Then
                        TDSSection_TDS3 = "192-Salary-Payment to Indian Government employees"
                        
                        'ElseIf SectionTDS = "192A" Then
                       ' ElseIf SectionTDS = "2AA" Or SectionTDS = "2aa" Or UCase(SectionTDS) = "2AA" Then
                       ElseIf UCase(SectionTDS) = UCase("2AA") Then
                        TDSSection_TDS3 = "192A-TDS on PF withdrawal"
                        
                        ElseIf UCase(SectionTDS) = UCase("193") Then
                        TDSSection_TDS3 = "193-Interest on Securities"
                        
                        ElseIf UCase(SectionTDS) = UCase("194") Then
                        TDSSection_TDS3 = "194-Dividends"
                        
                        ElseIf UCase(SectionTDS) = UCase("94A") Then
                        TDSSection_TDS3 = "194A-Interest other than 'Interest on securities'"
                        
                        ElseIf UCase(SectionTDS) = UCase("94B") Then
                        TDSSection_TDS3 = "194B-Winning from lottery or crossword puzzle"
                        
                       ' ElseIf SectionTDS = "94BA" Then
                          'ElseIf SectionTDS = "9BA" Or SectionTDS = "9ba" Or UCase(SectionTDS) = "9BA" Then
                           ElseIf UCase(SectionTDS) = UCase("9BA") Then
                        TDSSection_TDS3 = "194BA-Winnings from online games"
                        
                        ElseIf UCase(SectionTDS) = UCase("4BB") Then
                        TDSSection_TDS3 = "194BB-Winning from horse race"
                        
                        ElseIf UCase(SectionTDS) = UCase("94C") Then
                        TDSSection_TDS3 = "194C-Payments to contractors and sub-contractors"
                        
                        ElseIf UCase(SectionTDS) = UCase("94D") Then
                        TDSSection_TDS3 = "194D-Insurance commission"
                        
                        ElseIf UCase(SectionTDS) = UCase("4DA") Then
                        TDSSection_TDS3 = "194DA-Payment in respect of life insurance policy"
                        
                        ElseIf UCase(SectionTDS) = UCase("94E") Then
                        TDSSection_TDS3 = "194E-Payments to non-resident sportsmen or sports associations"
                        
                        ElseIf UCase(SectionTDS) = UCase("4EE") Then
                        TDSSection_TDS3 = "194EE-Payments in respect of deposits under National Savings"
                        
                       'ElseIf SectionTDS = "4F" Then
                        ' ElseIf SectionTDS = "94F" Or SectionTDS = "94f" Or UCase(SectionTDS) = "94F" Then
                         ElseIf UCase(SectionTDS) = UCase("94F") Then
                        TDSSection_TDS3 = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
                        
                        'ElseIf SectionTDS = "4G" Then
                         'ElseIf SectionTDS = "94G" Or SectionTDS = "94g" Or UCase(SectionTDS) = "94G" Then
                         ElseIf UCase(SectionTDS) = UCase("94G") Then
                        TDSSection_TDS3 = "194G-Commission, price, etc. on sale of lottery tickets"
                        
                        'ElseIf SectionTDS = "4H" Then
                         'ElseIf SectionTDS = "94H" Or SectionTDS = "94h" Or UCase(SectionTDS) = "94H" Then
                         ElseIf UCase(SectionTDS) = UCase("94H") Then
                        TDSSection_TDS3 = "194H-Commission or brokerage"
                        
                        'ElseIf SectionTDS = "4-IA" Then
                         'ElseIf SectionTDS = "4IA" Or SectionTDS = "4ia" Or UCase(SectionTDS) = "4IA" Then
                         ElseIf UCase(SectionTDS) = UCase("4IA") Then
                        'As per Ashish and Madhuconfirmation dropdown changed  SIT_94076 06/06/2025
                        'TDSSection_TDS3 = "194I(a)-Rent on hiring of plant and machinery"
                        TDSSection_TDS3 = "194IA-TDS on Sale of immovable property"

                        
                        'ElseIf SectionTDS = "4-IB" Then
                         'ElseIf SectionTDS = "4IB" Or SectionTDS = "4ib" Or UCase(SectionTDS) = "4IB" Then
                          ElseIf UCase(SectionTDS) = UCase("4IB") Then
                        'As per Ashish and Madhuconfirmation dropdown changed SIT_94076 06/06/2025
                        'TDSSection_TDS3 = "194I(b)-Rent on other than plant and machinery"
                         TDSSection_TDS3 = "194IB-Payment of rent by certain individuals or Hindu undivided"
                        
                        'ElseIf SectionTDS = "4IA" Then
                         'ElseIf SectionTDS = "9IA" Or SectionTDS = "9ia" Or UCase(SectionTDS) = "9IA" Then
                         ElseIf UCase(SectionTDS) = UCase("9IA") Then
                        TDSSection_TDS3 = "194IA-TDS on Sale of immovable property"
                        
                        'ElseIf SectionTDS = "4IB" Then
                        'ElseIf SectionTDS = "9IB" Or SectionTDS = "9ib" Or UCase(SectionTDS) = "9IB" Then
                        ElseIf UCase(SectionTDS) = UCase("9IB") Then
                        TDSSection_TDS3 = "194IB-Payment of rent by certain individuals or Hindu undivided"
                        
                        ElseIf UCase(SectionTDS) = UCase("4IC") Then
                        TDSSection_TDS3 = "194IC-Payment under specified agreement"
                        
                        'ElseIf SectionTDS = "94J-A" Then
                        'ElseIf SectionTDS = "4JA" Or SectionTDS = "4ja" Or UCase(SectionTDS) = "4JA" Then
                        ElseIf UCase(SectionTDS) = UCase("4JA") Then
                        TDSSection_TDS3 = "194J(a)-Fees for technical services"
                        
                        'ElseIf SectionTDS = "94J-B" Then
                        'ElseIf SectionTDS = "4JB" Or SectionTDS = "4jb" Or UCase(SectionTDS) = "4JB" Then
                        ElseIf UCase(SectionTDS) = UCase("4JB") Then
                        TDSSection_TDS3 = "194J(b)-Fees for professional  services or royalty etc"
                        
                        ElseIf UCase(SectionTDS) = UCase("94K") Then
                        TDSSection_TDS3 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
                        
                        ElseIf UCase(SectionTDS) = UCase("4LA") Then
                        TDSSection_TDS3 = "194LA-Payment of compensation on acquisition of certain immovable"
                        
                        ElseIf UCase(SectionTDS) = UCase("4LB") Then
                        TDSSection_TDS3 = "194LB-Income by way of Interest from Infrastructure Debt fund"
                        
                        'ElseIf SectionTDS = "4LC1" Then
                         'ElseIf SectionTDS = "LC1" Or SectionTDS = "lc1" Or UCase(SectionTDS) = "LC1" Then
                         ElseIf UCase(SectionTDS) = UCase("LC1") Then
                        TDSSection_TDS3 = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
                        
                        'ElseIf SectionTDS = "4LC2" Then
                         'ElseIf SectionTDS = "LC2" Or SectionTDS = "lc2" Or UCase(SectionTDS) = "LC2" Then
                          ElseIf UCase(SectionTDS) = UCase("LC2") Then
                        TDSSection_TDS3 = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
                        
                        'ElseIf SectionTDS = "4LC3" Then
                        ' ElseIf SectionTDS = "LC3" Or SectionTDS = "lc3" Or UCase(SectionTDS) = "Lc3" Then
                         ElseIf UCase(SectionTDS) = UCase("Lc3") Then
                        TDSSection_TDS3 = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
                        
'Malli_AY_2026_27  SIT-117189 10/04/2026
'                        ElseIf SectionTDS = "4BA1" Then
                        ElseIf SectionTDS = "BA1" Then
                        TDSSection_TDS3 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'Malli_AY_2026_27  SIT-117189 10/04/2026
'                        ElseIf SectionTDS = "4BA2" Then
                         ElseIf SectionTDS = "BA2" Then
                        TDSSection_TDS3 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"

'                        ElseIf SectionTDS = "LBA1" Then
'                        TDSSection_TDS3 = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
'
'                        ElseIf SectionTDS = "LBA2" Then
'                        TDSSection_TDS3 = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
'
                        'ElseIf SectionTDS = "LBA3" Then
                        'ElseIf SectionTDS = "BA3" Or SectionTDS = "ba3" Or UCase(SectionTDS) = "BA3" Then
                        ElseIf UCase(SectionTDS) = UCase("BA3") Then
                        TDSSection_TDS3 = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
                        
                        ElseIf UCase(SectionTDS) = UCase("LBB") Then
                        TDSSection_TDS3 = "194LBB-Income in respect of units of investment fund"
                        
                        ElseIf UCase(SectionTDS) = UCase("94R") Then
                        TDSSection_TDS3 = "194R-Benefits or perquisites of business or profession"
                        
                        ElseIf UCase(SectionTDS) = UCase("94S") Then
                        TDSSection_TDS3 = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
                        
                       ' ElseIf SectionTDS = "94B-P" Then
                        'ElseIf SectionTDS = "4BP" Or SectionTDS = "4bp" Or UCase(SectionTDS) = "4BP" Then
                        ElseIf UCase(SectionTDS) = UCase("4BP") Then
                        TDSSection_TDS3 = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
                        
                        'ElseIf SectionTDS = "94R-P" Then
                        ' ElseIf SectionTDS = "4RP" Or SectionTDS = "4rp" Or UCase(SectionTDS) = "4RP" Then
                         ElseIf UCase(SectionTDS) = UCase("4RP") Then
                        TDSSection_TDS3 = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
                        
                        'ElseIf SectionTDS = "94S-P" Then
                       ' ElseIf SectionTDS = "4SP" Or SectionTDS = "4sp" Or UCase(SectionTDS) = "4SP" Then
                        ElseIf UCase(SectionTDS) = UCase("4SP") Then
                        TDSSection_TDS3 = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
                        
                        ElseIf UCase(SectionTDS) = UCase("LBC") Then
                        TDSSection_TDS3 = "194LBC-Income in respect of investment in securitization trust"
                        
                        ElseIf UCase(SectionTDS) = UCase("4LD") Then
                        TDSSection_TDS3 = "194LD-TDS on interest on bonds / government securities"
                        
                        ElseIf UCase(SectionTDS) = UCase("94M") Then
                        TDSSection_TDS3 = "194M-Payment of certain sums by certain individuals or HUF"
                        
                        ElseIf UCase(SectionTDS) = UCase("94N") Then
                        TDSSection_TDS3 = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
                        
                        'ElseIf SectionTDS = "94N-F" Then
                         'ElseIf SectionTDS = "4NF" Or SectionTDS = "4nf" Or UCase(SectionTDS) = "4NF" Then
                          ElseIf UCase(SectionTDS) = UCase("4NF") Then
                        TDSSection_TDS3 = "194N-First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
                        
                        'ElseIf SectionTDS = "94N-C" Then
                         ' ElseIf SectionTDS = "4NC" Or SectionTDS = "4nc" Or UCase(SectionTDS) = "4NC" Then
                           ElseIf UCase(SectionTDS) = UCase("4NC") Then
                        TDSSection_TDS3 = "194N-Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
                        
                        'ElseIf SectionTDS = "94N-FT" Then
                         'ElseIf SectionTDS = "NFT" Or SectionTDS = "nft" Or UCase(SectionTDS) = "NFT" Then
                          ElseIf UCase(SectionTDS) = UCase("NFT") Then
                        TDSSection_TDS3 = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
                        
                        ElseIf UCase(SectionTDS) = UCase("94O") Then
                        TDSSection_TDS3 = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
                        
                        ElseIf UCase(SectionTDS) = UCase("94P") Then
                        TDSSection_TDS3 = "194P-Deduction of tax in case of specified senior citizen"
                        
                        ElseIf UCase(SectionTDS) = UCase("94Q") Then
                        TDSSection_TDS3 = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
                        
                        ElseIf UCase(SectionTDS) = UCase("195") Then
                        TDSSection_TDS3 = "195-Other sums payable to a non-resident"
                        
                        ElseIf UCase(SectionTDS) = UCase("96A") Then
                        TDSSection_TDS3 = "196A-Income in respect of units of non-residents"
                        
                        ElseIf UCase(SectionTDS) = UCase("96B") Then
                        TDSSection_TDS3 = "196B-Payments in respect of units to an offshore fund"
                        
                        ElseIf UCase(SectionTDS) = UCase("96C") Then
                        TDSSection_TDS3 = "196C-Income from foreign currency bonds or shares of Indian"
                        
                        ElseIf UCase(SectionTDS) = UCase("96D") Then
                        TDSSection_TDS3 = "196D-Income of foreign institutional investors from securities"
                        
                        'ElseIf SectionTDS = "96DA" Then
                         'ElseIf SectionTDS = "6DA" Or SectionTDS = "6da" Or UCase(SectionTDS) = "6DA" Then
                         ElseIf UCase(SectionTDS) = UCase("6DA") Then
                        TDSSection_TDS3 = "196D(1A)-Income of specified fund from securities"
                        
                        'ElseIf SectionTDS = "94BA-P" Then
                        'ElseIf SectionTDS = "BAP" Or SectionTDS = "bap" Or UCase(SectionTDS) = "BAP" Then
                        ElseIf UCase(SectionTDS) = UCase("BAP") Then
                        TDSSection_TDS3 = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
                        
                        Else
                        'TDSSection_TDS3 = ""
                        TDSSection_TDS3 = "(Select)"
                        End If
                        
                    If TDSSection_TDS3 <> "" And Sheet4.Cells(rowcount, SectionTDSColNo).Locked = False Then
                        Sheet4.Cells(rowcount, SectionTDSColNo).value = TDSSection_TDS3
                        'TDSSection_TDS3 = ""
                        SectionTDS = ""
                    End If
 
           '---------------------------
 
 
 
 
 
                
'issue 16 solved
                If Trim(node("taxDeductCreditDtls")("taxDeductedOwnHands")) <> "" And Sheet4.Cells(rowcount, TDSDeductedColNo).Locked = False Then
                    Sheet4.Cells(rowcount, TDSDeductedColNo).value = UCase(node("taxDeductCreditDtls")("taxDeductedOwnHands"))
                End If
                If Trim(node("taxDeductCreditDtls")("taxClaimedOwnHands")) <> "" And Sheet4.Cells(rowcount, TDSClaimedColNo).Locked = False Then
                    Sheet4.Cells(rowcount, TDSClaimedColNo).value = UCase(node("taxDeductCreditDtls")("taxClaimedOwnHands"))
                End If
'issue 16 solved
 
                'issue 17
                If Trim(node("deductedYr")) <> "" And Sheet4.Cells(rowcount, FYColNo).Locked = False Then
                    Sheet4.Cells(rowcount, FYColNo).value = UCase(node("deductedYr"))
                End If
                If Trim(node("broughtFwdTDSAmt")) <> "" And Sheet4.Cells(rowcount, TDSOtherColNo).Locked = False Then
                    Sheet4.Cells(rowcount, TDSOtherColNo).value = UCase(node("broughtFwdTDSAmt"))
                End If
'issue 17
                If Trim(node("grossAmount")) <> "" And Sheet4.Cells(rowcount, TDSGrossAmountColNo).Locked = False Then
                    Sheet4.Cells(rowcount, TDSGrossAmountColNo).value = UCase(node("grossAmount"))
                End If
              'Malli------------------------------
'                'Malli_AY_2025_26
'                If node("sectionCode") <> "" Then
'           Dim HeadOfIncome As Variant
'           Dim sectioncode
'           sectioncode = UCase(node("sectionCode"))
'           If sectioncode = "94M" Or sectioncode = "94S" Then
'             HeadOfIncome = "Not applicable (only in case TDS is deducted u/s 194N)"
'           ElseIf sectioncode = "94I(b)" Or sectioncode = "194IB" Or _
'             sectioncode = "4IB" Or sectioncode = "4Ib" Then
'             HeadOfIncome = "Income from House property"
'           Else
'                HeadOfIncome = "(Select)"
'           End If
'         If HeadOfIncome <> "" And Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
'                Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
'            End If
'        End If
               ' -----Malli comented--------------------------------------------
'            Dim HeadOfIncome As Variant
'            'not in json
'              HeadOfIncome = UCase(node("headOfIncome"))
'
'            If HeadOfIncome = "BP" Then
'                HeadOfIncome = "Income from Business & Profession"
'            ElseIf HeadOfIncome = "HP" Then
'                HeadOfIncome = "Income from House Property"
'            ElseIf HeadOfIncome = "OS" Then
'                HeadOfIncome = "Income from Other Source"
'            ElseIf HeadOfIncome = "EI" Then
'                HeadOfIncome = "Exempt Income"
'            ElseIf HeadOfIncome = "NA" Then
'                HeadOfIncome = "Not applicable (only in case TDS is deducted u/s 194N)"
'            End If
'
'
'            If HeadOfIncome <> "" And Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
'                Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
'            End If

Dim HeadOfIncome, HeadOfIncome_tds3 As Variant
            'not in json
            'AY_2024_25 OLD
              'HeadOfIncome = UCase(node("headOfIncome"))
              'AY_2025_26 Change
              HeadOfIncome = UCase(node("HeadOfIncome"))
       'AY_2025-26 Mallikarjun added variable HeadOfIncome_tds3
            If HeadOfIncome = "BP" Then
                HeadOfIncome_tds3 = "Income from Business & Profession"
            ElseIf HeadOfIncome = "HP" Then
                HeadOfIncome_tds3 = "Income from House Property"
            ElseIf HeadOfIncome = "OS" Then
                HeadOfIncome_tds3 = "Income from Other Source"
            ElseIf HeadOfIncome = "EI" Then
               HeadOfIncome_tds3 = "Exempt Income"
               'Commented by Ayush
'            ElseIf HeadOfIncome = "NA" Then
'                HeadOfIncome_tds3 = "Not applicable (only in case TDS is deducted u/s 194N)"
            Else
                 HeadOfIncome_tds3 = "(Select)"
            End If
 

            If HeadOfIncome_tds3 <> "" And Sheet4.Cells(rowcount, TDSHeadColNo).Locked = False Then
                Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome_tds3
            End If
           ' -----------------------------------------------------------
        cnt = cnt + 1
    Next node
    RecTDS11 = cnt
End Sub

'issue 18 solved
'import json for ScheduleBP/TurnoverGrsRcptForGSTIN

Function ImportBPGst_pfl(jsonText As String)
On Error Resume Next
    
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1 As Variant
    Dim GSTRnoColNo, GSTRAmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount As Variant
    Dim Sec44ADName, Sec44ADCode, Sec44ADDesc, grossReceipts, grsTrnOverAnyOthMode_1, grsTrnOverBank44ADA_1 As Variant
    Dim Sec44ADAName, Sec44ADACode, Sec44ADADesc, bankReceipts As Variant
    Dim Sec44AEName, Sec44AECode, Sec44AEDesc As Variant
    Dim SEC44AERegNo, Sec44AEOLE, Sec44AETonnageCpct, Sec44AENoOfMonths As Variant
        
        
    Set jsonObject = ParseJson(jsonText)
    
    Set Nodelist = jsonObject("lastFiledITR")("natOfBus44AD")
    
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

    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
    rowcount = rowcount + 1
    
    If Sheet3.Cells(rowcount, Sec44ADName).Locked = False Then
        Sheet3.Cells(rowcount, Sec44ADName).value = node("nameOfBusiness")
    End If
            Dim iNOB As Variant
            Dim sNOB As Variant
        iNOB = node("codeAD")
        sNOB = Findtext(iNOB, "NOB44AD")
    If Sheet3.Cells(rowcount, Sec44ADCode).Locked = False Then
        Sheet3.Cells(rowcount, Sec44ADCode).value = sNOB
    End If
    If Sheet3.Cells(rowcount, Sec44ADDesc).Locked = False Then
        Sheet3.Cells(rowcount, Sec44ADDesc).value = node("description")
    End If
    cnt = cnt + 1
Next node
RecTDS1 = cnt
  
  
bankReceipts = jsonObject("form26as")("grsTrnOverAnyOthMode")

If bankReceipts <> "" And Sheet3.Range("BP_E1a").Locked = False Then
        Sheet3.Range("BP_E1a").value = bankReceipts
End If
  
Set Nodelist = New Collection

Set Nodelist = jsonObject("lastFiledITR")("natOfBus44ADA")
    
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

    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName1").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, Sec44ADAName).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADAName).value = node("nameOfBusiness")
        End If
        
           
            iNOB = node("codeADA")
            sNOB = Findtext(iNOB, "NOB44ADA")
        If Sheet3.Cells(rowcount, Sec44ADACode).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADACode).value = sNOB
        End If
        If Sheet3.Cells(rowcount, Sec44ADADesc).Locked = False Then
            Sheet3.Cells(rowcount, Sec44ADADesc).value = node("description")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
    
grossReceipts = jsonObject("form26as")("persumptiveInc44ADA")("grsReceipt")

If grossReceipts <> "" And Sheet3.Range("BP_E3_GR").Locked = False Then
        Sheet3.Range("BP_E3_GR").value = grossReceipts
End If

'newly added by Sai for 2024-25 Prefill implement----------------->>>>>>>>>>

'grsTrnOverAnyOthMode_1 = jsonObject("form26as")("grsTrnOverAnyOthMode")

'If grsTrnOverAnyOthMode_1 <> "" And Sheet3.Range("BP_E3_a").Locked = False Then
'        Sheet3.Range("BP_E3_a").value = grsTrnOverAnyOthMode_1
'End If

'SIT-67964  'Malli

grsTrnOverAnyOthMode_1 = jsonObject("form26as")("persumptiveInc44ADA")("grsTrnOverBank44ADA")

If grsTrnOverAnyOthMode_1 <> "" And Sheet3.Range("BP_E3_a").Locked = False Then
        Sheet3.Range("BP_E3_a").value = grsTrnOverAnyOthMode_1
End If
'----------------------------------------------------------------->>>>>>>>>>


Set Nodelist = New Collection
Set Nodelist = jsonObject("lastFiledITR")("natOfBus44AE")
    
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

    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName2").name)
    rowcount = rowcount - 1
    cnt = 0

For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, Sec44AEName).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AEName).value = node("nameOfBusiness")
        End If

            iNOB = node("codeAE")
            sNOB = Findtext(iNOB, "NOB")
        If Sheet3.Cells(rowcount, Sec44AECode).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AECode).MergeArea.value = sNOB
        End If
        If Sheet3.Cells(rowcount, Sec44AEDesc).Locked = False Then
            Sheet3.Cells(rowcount, Sec44AEDesc).value = node("description")
        End If
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
  
Set Nodelist = jsonObject("lastFiledITR")("goodsDtlsUs44AE")
    
    SEC44AERegNo = Sheet3.Range("Sec44AE.RegNo").Column
    Sec44AEOLE = Sheet3.Range("Sec44AE.OLH").Column
    Sec44AETonnageCpct = Sheet3.Range("Sec44AE.TonnageCpct").Column
    Sec44AENoOfMonths = Sheet3.Range("Sheet44AE.noOfMonths").Column
    
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
    End If

    
    rowcount = getRowNo(Sheet3.Range("Sec44AE.RegNo").name)
    rowcount = rowcount - 1
    cnt = 0
    Dim iAEOLE As Variant
    Dim sAEOLE As Variant
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet3.Cells(rowcount, SEC44AERegNo).Locked = False Then
            Sheet3.Cells(rowcount, SEC44AERegNo).value = node("regNumberGoodsCarriage")
        End If
            iAEOLE = node("ownedLeasedHiredFlag")
              
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
            Sheet3.Cells(rowcount, Sec44AETonnageCpct).value = node("tonnageCapacity")
        End If
        If Sheet3.Cells(rowcount, Sec44AENoOfMonths).Locked = False Then
            If node.exists("holdingPeriod") And node("holdingPeriod") <> "" Then
                Sheet3.Cells(rowcount, Sec44AENoOfMonths).value = node("holdingPeriod")
            Else
                Sheet3.Cells(rowcount, Sec44AENoOfMonths).value = 12
            End If
        End If
cnt = cnt + 1
Next node
RecTDS1 = cnt
  
  
Set Nodelist = jsonObject("lastFiledITR")("turnoverGrsRcptForGSTIN")
    
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
       
            If Trim(node("gstinNo")) <> "" And Sheet3.Cells(rowcount, GSTRnoColNo).Locked = False Then
                Sheet3.Cells(rowcount, GSTRnoColNo).value = node("gstinNo")
            End If
            If Trim(node("amtTurnGrossRcptGSTIN")) <> "" And Sheet3.Cells(rowcount, GSTRAmountColNo).Locked = False Then
                Sheet3.Cells(rowcount, GSTRAmountColNo).value = node("amtTurnGrossRcptGSTIN")
            End If
            
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function    'issue 18 solved

'issue 10 solved
Function Import_IncomeDeductions(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary, node, Nodelist As Object
Dim Salary, PerquisitesValue, ProfitsInSalary, TypeOfHP, GrossRentReceived, IncChrgblHeadHP As Variant
Dim init, NatureColNo, DescriptionColNo, AmtColNo, TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, cnt, test
Dim EntertainmentAlw16ii, ProfessionalTaxUs16iii, DeductionUs16ia
Dim Nature1ColNo, Description1ColNo, Amt1ColNo
'Dim Section80C, Section, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80DD, Section80DDB As Variant
Dim TypeOfHP_1
Dim Section80C, Section80CCC, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80DDB As Variant
Dim Section80E, Section80G, Section80TTA, TotalDeductions, ReliefUs89 As Variant

Set jsonObject = ParseJson(jsonText)

Salary = jsonObject("form24q")("incomeDeductions")("salary")
PerquisitesValue = jsonObject("form24q")("incomeDeductions")("perquisitesValue")
ProfitsInSalary = jsonObject("form24q")("incomeDeductions")("profitsInSalary")
DeductionUs16ia = jsonObject("form24q")("incomeDeductions")("deductionUs16Ia")
EntertainmentAlw16ii = jsonObject("form24q")("incomeDeductions")("entertainmentAlw16Ii")
ProfessionalTaxUs16iii = jsonObject("form24q")("incomeDeductions")("professionalTaxUs16Iii")
'TypeOfHP = jsonObject("form26as")("scheduleHP")("propertyDetails")("typeOfHP") 'Newly changed by Bindu

'Malli commented AY_2025-26 27/01/2026
'GrossRentReceived = jsonObject("form26as")("grossRent")
'---------------------------

IncChrgblHeadHP = jsonObject("form24q")("incomeDeductions")("totalIncomeOfHP")

Section80C = jsonObject("form24q")("usrDeductUndChapVIAType")("section80C")
Section80CCC = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCC")
Section80CCDEmployeeOrSE = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployeeOrSE")
Section80CCD1B = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCD1B")
Section80CCDEmployer = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployer")


'Section80DD = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDUsrType")

Section80DDB = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDB")
Section80E = jsonObject("form24q")("usrDeductUndChapVIAType")("section80E")
Section80G = jsonObject("form24q")("usrDeductUndChapVIAType")("section80G")
Section80TTA = jsonObject("form24q")("usrDeductUndChapVIAType")("section80TTA")

'Section80U = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80U")


TotalDeductions = jsonObject("form24q")("usrDeductUndChapVIAType")("totalChapVIADeductions")
'Prefill-30.01.2023.103.ID
'ReliefUs89 = jsonObject("form24q")("taxComputation")("section89")
ReliefUs89 = jsonObject("form10E")("TaxComputation")("section89")
'End Change
    If Salary <> "" And Sheet1.Range("IncD.IncomeFromSal_ii").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_ii").value = Salary
    End If
    
    If PerquisitesValue <> "" And Sheet1.Range("IncD.IncomeFromSal_iii").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_iii").value = PerquisitesValue
    End If
    
    If ProfitsInSalary <> "" And Sheet1.Range("IncD.IncomeFromSal_iv").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_iv").value = ProfitsInSalary
    End If
    
    Set init = jsonObject("form24q")
    If init <> Empty Then
    If init.exists("allwncExemptUs10DtlsType") Then
        Set Nodelist = jsonObject("form24q")("allwncExemptUs10DtlsType")
   
        NatureColNo = Sheet1.Range("Sheet1.Nature").Column
        AmtColNo = Sheet1.Range("Sheet1.Amount").Column
        
        TotalExRow = Range("Sheet1.Nature").Rows.count
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
    
        If (TotalXMLRow > 0) Then
            Sheet1.Range("Sheet1.Nature").ClearContents
            Sheet1.Range("Sheet1.Amount").ClearContents
            
        End If
        
        If (TotalDiffRow > 0) Then
            AddDiffRows_Exempt1 (TotalDiffRow)
        End If
    
       rowcount = getRowNo(Sheet1.Range("Sheet1.Nature").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
            
          If UCase(node("salNatureDesc")) <> "10(13A)" Then 'Newly added by Malli as per DE V4
            rowcount = rowcount + 1
            If Sheet1.Cells(rowcount, NatureColNo).Locked = False Then
                If UCase(node("salNatureDesc")) = "10(10B)(I)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
                ElseIf UCase(node("salNatureDesc")) = "10(10B)(II)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
'Konda updated on 12-03-2026----v0.2
'                ElseIf UCase(node("salNatureDesc")) = "OTH" Then
'                    Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
                ElseIf UCase(node("salNatureDesc")) = "10(17)" Then
                    Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(17)-Allowance MP/MLA/MLC"

                Else
                    test = Findtext("Sec " & (node("salNatureDesc")), "Sheet1.NOA")
                    Sheet1.Range("I" & rowcount).value = test
                
                End If
            End If
'Konda updated on 12-03-2026----V0.2
'            If Sheet1.Range("AO" & rowcount).Locked = False Then
'                Sheet1.Range("AO" & rowcount).value = node("salOthAmount")
'            End If
            If Sheet1.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet1.Cells(rowcount, AmtColNo).value = node("salOthAmount")
            End If
'===========================================
            cnt = cnt + 1
            
            End If 'Malli
        Next node
    End If
    End If
    
    If DeductionUs16ia <> "" And Sheet1.Range("IncD.IncomeFromSal_iva").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_iva").value = DeductionUs16ia
    End If
    
    If EntertainmentAlw16ii <> "" And Sheet1.Range("IncD.IncomeFromSal_v").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_v").value = EntertainmentAlw16ii
    End If
    
    If ProfessionalTaxUs16iii <> "" And Sheet1.Range("IncD.IncomeFromSal_ivc").Locked = False Then
        Sheet1.Range("IncD.IncomeFromSal_ivc").value = ProfessionalTaxUs16iii
    End If
    ' issue-a70 start
    
  'Malli_AY_2025_26----------------------
'     Set init = jsonObject("form26as")
'    If init <> Empty Then
'        If init.exists("scheduleHP") Then
'        Set Nodelist = jsonObject("form26as")("scheduleHP")("propertyDetails")
'         'SIT-67946 AY_2024_25  'Malli
'    For Each node In Nodelist
'        TypeOfHP_1 = node("typeOfHP")
'        If node("typeOfHP") <> "" Then
'            If TypeOfHP_1 = "Y" Then
'                Sheet1.Range("IncD.TypeOfHP").value = "Let Out"
'            ElseIf TypeOfHP_1 = "N" Then
'                Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied"
'            ElseIf TypeOfHP_1 = "D" Then
'                Sheet1.Range("IncD.TypeOfHP").value = "Deemed Let Out"
'            End If
'        End If
'     Next
'      End If
'
'    End If
    
'  'Malli commented AY_2025-26 27/01/2026
'
'    'Malli---------
'    'form26as.typeOfHP 'lastFiledITR.typeOfHP
'If jsonObject("lastFiledITR") <> Null Then
'    TypeOfHP_1 = jsonObject("lastFiledITR")("typeOfHP")
'        If TypeOfHP_1 <> "" And Sheet1.Range("IncD.TypeOfHP").Locked = False Then
'                If TypeOfHP_1 = "Y" Or TypeOfHP_1 = "L" Then
'                    Sheet1.Range("IncD.TypeOfHP").value = "Let Out"
'                    ElseIf TypeOfHP_1 = "N" Or TypeOfHP_1 = "S" Then
'                    Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied"
'                    ElseIf TypeOfHP_1 = "D" Then
'                    Sheet1.Range("IncD.TypeOfHP").value = "Deemed let out"
'                End If
'        End If
'ElseIf jsonObject("form26as") <> Null Then
'    TypeOfHP_1 = jsonObject("form26as")("typeOfHP")
'        If TypeOfHP_1 <> "" And Sheet1.Range("IncD.TypeOfHP").Locked = False Then
'                If TypeOfHP_1 = "Y" Or TypeOfHP_1 = "L" Then
'                    Sheet1.Range("IncD.TypeOfHP").value = "Let Out"
'                    ElseIf TypeOfHP_1 = "N" Or TypeOfHP_1 = "S" Then
'                    Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied"
'                    ElseIf TypeOfHP_1 = "D" Then
'                    Sheet1.Range("IncD.TypeOfHP").value = "Deemed let out"
'                End If
'        End If
'
'End If
'
'
'
'
'    '--------------------------------------------------
'
'    'issue-a70 end
'    Sheet1.Range("IncD.IncomeFromHP_i").value = GrossRentReceived
   '-----------------------------------------------------------------------
    
    
    If IncChrgblHeadHP <> "" And Sheet1.Range("IncD.IncomeFromHP").Locked = False Then
        Sheet1.Range("IncD.IncomeFromHP").value = IncChrgblHeadHP
    End If
    
     'Malli_30/05/2025
        incDeductionsOthIncCPC_TAX_chk = False
        incDeductionsOthIncCPC_TAX_RowCnt = 0
        '------------------------------------------
    
    Set init = jsonObject("form26as")
    If init <> Empty Then
    If init.exists("incomeDeductionsOthersInc") Then
    
   
        Set Nodelist = jsonObject("form26as")("incomeDeductionsOthersInc")
        
        Nature1ColNo = Sheet1.Range("Sheet1.Nature1").Column
        Amt1ColNo = Sheet1.Range("Sheet1.Amount1").Column
        
        TotalExRow = Range("Sheet1.Nature1").Rows.count
        
        Dim i
        i = 1
         For Each node In Nodelist
            
            If node("othSrcNatureDesc") = "DIV" Then
                'Sheet1.Range("IncD_dq1div").value = node("othSrcOthAmount")
                'Newly added by Bindu
                 Sheet1.Range("IncD_dqdiv").value = node("othSrcOthAmount")
                Nodelist.Remove i
            End If
            i = i + 1
        Next node
        TotalXMLRow = Nodelist.count
       
        TotalDiffRow = TotalXMLRow - TotalExRow
        
    
        If (TotalXMLRow > 0) Then
            Sheet1.Range("Sheet1.Nature1").ClearContents
            Sheet1.Range("Sheet1.Amount1").ClearContents
        End If
        
        If (TotalDiffRow > 0) Then
            AddDiffRows_Exempt2 (TotalDiffRow)
        End If
 
        rowcount = getRowNo(Sheet1.Range("Sheet1.Nature1").name)
        rowcount = rowcount - 1
        cnt = 0
    
        For Each node In Nodelist
            rowcount = rowcount + 1
            incDeductionsOthIncCPC_TAX_RowCnt = rowcount   'Malli_30/05/2025
            
            If Sheet1.Cells(rowcount, Nature1ColNo).Locked = False Then
                If UCase(node("othSrcNatureDesc")) = "SAV" Then
                    Sheet1.Cells(rowcount, Nature1ColNo).value = "Interest from Saving Bank Account"
                ElseIf UCase(node("othSrcNatureDesc")) = "IFD" Then
                    Sheet1.Cells(rowcount, Nature1ColNo).value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
                ElseIf UCase(node("othSrcNatureDesc")) = "TAX" Then
                    Sheet1.Cells(rowcount, Nature1ColNo).value = "Interest from Income Tax Refund"
                          incDeductionsOthIncCPC_TAX_chk = True  'Malli_30/05/2025
                ElseIf UCase(node("othSrcNatureDesc")) = "FAP" Then
                    Sheet1.Cells(rowcount, Nature1ColNo).value = "Family pension"
                    
                'Newly added by Bindu
                ElseIf UCase(node("othSrcNatureDesc")) = "DIV" Then
                
                   Sheet1.Range("IncD_dqdiv").value = node("othSrcOthAmount")
                    '--end
            
                
                ElseIf UCase(node("othSrcNatureDesc")) = "OTH" Then
                    Sheet1.Cells(rowcount, Nature1ColNo).value = "Any Other"
                End If
            End If
            
            If UCase(node("othSrcNatureDesc")) <> "DIV" Then
                If Sheet1.Cells(rowcount, Amt1ColNo).Locked = False Then
                    Sheet1.Cells(rowcount, Amt1ColNo).value = node("othSrcOthAmount")
                End If
             End If
            cnt = cnt + 1
        Next node
    End If
    End If
    
    
    'Malli--------------------30/05/2025
    'incDeductionsOthIncCPC_TAX_chk
    'incDeductionsOthIncCPC_TAX_RowCnt
    
If incDeductionsOthIncCPC_TAX_chk <> True Then
           
           Set init = jsonObject("incDeductionsOthIncCPC")
           If init <> Empty Then
               Set Nodelist = jsonObject("incDeductionsOthIncCPC")
               
               Nature1ColNo = Sheet1.Range("Sheet1.Nature1").Column
               Amt1ColNo = Sheet1.Range("Sheet1.Amount1").Column
               TotalExRow = Range("Sheet1.Nature1").Rows.count
               TotalXMLRow = Nodelist.count
               
               TotalDiffRow = TotalXMLRow - TotalExRow
               
            If Not incDeductionsOthIncCPC_TAX_RowCnt > 0 Then
               If (TotalXMLRow > 0) Then
                   Sheet1.Range("Sheet1.Nature1").ClearContents
                   Sheet1.Range("Sheet1.Amount1").ClearContents
               End If
            
               If (TotalDiffRow > 0) Then
                   AddDiffRows_Exempt2 (TotalDiffRow)
               End If
               
               rowcount = getRowNo(Sheet1.Range("Sheet1.Nature1").name)
               rowcount = rowcount - 1
               cnt = 0
            
            End If
               
               
           
               For Each node In Nodelist
               
                   If Not incDeductionsOthIncCPC_TAX_RowCnt > 0 Then
                   rowcount = rowcount + 1
                   End If
                   
                   If Sheet1.Cells(rowcount, Nature1ColNo).Locked = False Then
                       
                       If UCase(node("othSrcNatureDesc")) = "TAX" Then
                       
                            If incDeductionsOthIncCPC_TAX_RowCnt > 0 Then
                                AddDiffRows_Exempt2 (1)
                                 rowcount = rowcount + 1
                            End If
                          
                           
                           Sheet1.Cells(rowcount, Nature1ColNo).value = "Interest from Income Tax Refund"
                           
                           If Sheet1.Cells(rowcount, Amt1ColNo).Locked = False Then
                           Sheet1.Cells(rowcount, Amt1ColNo).value = node("othSrcOthAmount")
                           End If
                           
                       End If
                   End If
                   
                    
                   cnt = cnt + 1
               Next node
           End If
           
    
End If
    
    '------------------------------------
    
    
    
    
     If Section80C <> "" And Sheet1.Range("IncD.Section80C").Locked = False Then
        Sheet1.Range("IncD.Section80C").value = Section80C
    End If
    
    
    If Section80CCC <> "" And Sheet1.Range("IncD.Section80CCC").Locked = False Then
        Sheet1.Range("IncD.Section80CCC").value = Section80CCC
    End If
    If Section80CCDEmployeeOrSE <> "" And Sheet1.Range("IncD.Section80CCD_SE").Locked = False Then
        Sheet1.Range("IncD.Section80CCD_SE").value = Section80CCDEmployeeOrSE
    End If
    If Section80CCD1B <> "" And Sheet1.Range("IncD.Section80CCD1B_SE").Locked = False Then
        Sheet1.Range("IncD.Section80CCD1B_SE").value = Section80CCD1B
    End If
    If Section80CCDEmployer <> "" And Sheet1.Range("IncD.Section80CCD").Locked = False Then
        Sheet1.Range("IncD.Section80CCD").value = Section80CCDEmployer
    End If
    

    Section80DDB = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDB")
    
    
    'Commented as per the 2024-25 Prefill development-----------------------------------------
    
'    Dim Section80DDUsertypes
'     Section80DD = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDUsrType")
'    If Section80DD = "1" Then
'        Section80DDUsertypes = "1- Dependent person with Disability"
'    End If
'    If Section80DD = "2" Then
'        Section80DDUsertypes = "2-Dependent person with severe Disability"
'    End If
    
'    If Sheet1.Range("SELECT80DD").Locked = False Then
'        Sheet1.Range("SELECT80DD").value = Section80DDUsertypes
'    End If

     'Commented as per the 2024-25 Prefill development-----------------------------------------


    Dim section80DDBUsrType
    section80DDBUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDBUsrType")
    If section80DDBUsrType = "1" Then
        section80DDBUsrType = "1-Self or dependent"
    End If
    If section80DDBUsrType = "2" Then
        section80DDBUsrType = "2-Self or Dependent - Senior Citizen"
    End If
    If section80DDBUsrType <> "" And Sheet1.Range("SELECT80DDB").Locked = False Then
        Sheet1.Range("SELECT80DDB").value = section80DDBUsrType
    End If
    
    If Section80DDB <> "" And Sheet1.Range("IncD.Section80DDB").Locked = False Then
        Sheet1.Range("IncD.Section80DDB").value = Section80DDB
    End If
    
    If Section80E <> "" And Sheet1.Range("IncD.Section80E").Locked = False Then
        Sheet1.Range("IncD.Section80E").value = Section80E
    End If
    
   'Malli------AY_2025_26____________30/04/2025
    Dim Form10BAAckNum_Pfl As Variant
    Form10BAAckNum_Pfl = jsonObject("Form10BA")("Form10BAAckNum")
    If Form10BAAckNum_Pfl <> "" Then  '13/06/2025
            If Sheet8.Range("bacValue").value = 1 Then
                    Sheet1.Unprotect Password:=getmsgstate
                    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = False
                    Sheet1.Range("Sheet1.AckNum").MergeArea.Interior.Color = "&HCCFFCC"
                    Sheet1.Range("Sheet1.AckNum").MergeArea.value = Form10BAAckNum_Pfl
                    Sheet1.Unprotect Password:=getmsgstate
                    Sheet1.Range("Sheet1.AckNum").MergeArea.Font.Color = "&HD8D8D8"
                    Sheet1.Range("Sheet1.AckNum").MergeArea.Interior.Color = "&HD8D8D8"
                    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
            Else
                    If Form10BAAckNum_Pfl <> "" And Sheet1.Range("Sheet1.AckNum").Locked = False Then
                        Sheet1.Range("Sheet1.AckNum").value = Form10BAAckNum_Pfl
                    End If
            End If
    End If
    '-------------------------------------------
    
    
    
    If Section80TTA <> "" And Sheet1.Range("IncD.Section80TTA").Locked = False Then
        Sheet1.Range("IncD.Section80TTA").value = Section80TTA
    End If
    
    Dim Section80TTB
    'Malli----------
    'Malli_AY_2025_26_NEW
    Section80TTB = jsonObject("form24q")("usrDeductUndChapVIAType")("section80TTB")
    If Section80TTB <> "" And Sheet1.Range("IncD.Section80TTB").Locked = False Then
        Sheet1.Range("IncD.Section80TTB").value = Section80TTB
    End If
    
    
    
  'Commented as per the Latest Prefill 2024-25--------------------------------------
  'Malli_05/02/2026  SIT-109815
'    Dim Section80UUsrType
'    Section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType")


'   Section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType")
'
'   If Section80UUsrType = "1" Then
'        Section80UUsrType = "1-Self with disability"
'    End If
'
'    If Section80UUsrType = "2" Then
'        Section80UUsrType = "2-Self with severe disability"
'    End If
'
'    If Section80UUsrType <> "" And Sheet1.Range("NatureDisability_80U").Locked = False Then
'        Sheet14.Range("NatureDisability_80U").value = Section80UUsrType
'    End If



'-------------------------------------------
'    If Section80UUsrType = "1" Then
'        Section80UUsrType = "1-Self with disability"
'    End If
'
'    If Section80UUsrType = "2" Then
'        Section80UUsrType = "2-Self with severe disability"
'    End If
'
'    If Section80UUsrType <> "" And Sheet1.Range("SELECT80U").Locked = False Then
'        Sheet1.Range("SELECT80U").value = Section80UUsrType
'    End If
'
'    If Section80U <> "" And Sheet1.Range("IncD.Section80U").Locked = False Then
'        Sheet1.Range("IncD.Section80U").value = Section80U
'    End If
    'Commented as per the Latest Prefill 2024-25---------------------------
    
    
    
    
    If TotalDeductions <> "" And Sheet1.Range("IncD.TotalChapVIADeductions").Locked = False Then
        Sheet1.Range("IncD.TotalChapVIADeductions").value = TotalDeductions
    End If
    If ReliefUs89 <> "" And Sheet1.Range("IncD.Section89").Locked = False Then
        Sheet1.Range("IncD.Section89").value = ReliefUs89
    End If
'issue 10 solved
End Function
Sub AddDiffRows_Exempt2(DiffRows As Long)
    setDiffTblinfo_Exempt2
    Sheet1.Activate
    searchLastRow ("Sheet1.Nature1")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_Exempt2()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.Nature1").count
    Set rangecells = Sheet1.Range("Sheet1.Nature1").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Exempt2 = ccount
    DefinedgridNameRange = "Sheet1.Nature1||Sheet1.Description1||Sheet1.Amount1||Sheet1.Amount11||Sheet1.Amount22||Sheet1.Amount23||Sheet1.Amount24||Sheet1.Amount25"
 End Sub
 
 Sub AddDiffRows_Exempt1(DiffRows As Long)
    setDiffTblinfo_Exempt1
    Sheet1.Activate
    searchLastRow ("Sheet1.Nature")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_Exempt1()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Sheet1.Nature").count
    Set rangecells = Sheet1.Range("Sheet1.Nature").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Exempt1 = ccount
'Konda updated on 12-03-2026---V0.5.2
'    DefinedgridNameRange = "Sheet1.Nature||Sheet1.Description||Sheet1.Amount"
    DefinedgridNameRange = "Sheet1.Nature||Sheet1.Amount"
'=============================
 End Sub
Function Importschedule80D_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim temp80, temp802, SeniorCitizenFlag, SelfAndFamily, HealthInsPremSlfFam, PrevHlthChckUpSlfFam, SelfAndFamilySeniorCitizen As Variant
Dim HlthInsPremSlfFamSrCtzn, PrevHlthChckUpSlfFamSrCtzn, MedicalExpSlfFamSrCtzn, parentsSeniorCitizenFlag, Parents As Variant
Dim HlthInsPremParents, PrevHlthChckUpParents, ParentsSeniorCitizen, HlthInsPremParentsSrCtzn, PrevHlthChckUpParentsSrCtzn As Variant
Dim MedicalExpParentsSrCtzn, EligibleAmountOfDedn  As Variant

      
Set jsonObject = ParseJson(jsonText)

'issue 11 solved

HealthInsPremSlfFam = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("HealthInsPremSlfFam")
PrevHlthChckUpSlfFam = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpSlfFam")
HlthInsPremSlfFamSrCtzn = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("HlthInsPremSlfFamSrCtzn")
PrevHlthChckUpSlfFamSrCtzn = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpSlfFamSrCtzn")
MedicalExpSlfFamSrCtzn = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("MedicalExpSlfFamSrCtzn")

HlthInsPremParents = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("HlthInsPremParents")
PrevHlthChckUpParents = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpParents")
HlthInsPremParentsSrCtzn = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("HlthInsPremParentsSrCtzn")
PrevHlthChckUpParentsSrCtzn = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("PrevHlthChckUpParentsSrCtzn")
MedicalExpParentsSrCtzn = jsonObject("lastFiledITR")("f")("Sec80DSelfFamSrCtznHealth")("MedicalExpParentsSrCtzn")


            temp80 = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("SeniorCitizenFlag")
        If temp80 <> Empty Then
            If temp80 = "Y" Then
                Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Yes"
                Sheet16.Range("Health_InsuranceSC_80D").value = HlthInsPremSlfFamSrCtzn
                Sheet16.Range("Preventive_Health_SC_80D").value = PrevHlthChckUpSlfFamSrCtzn
                Sheet16.Range("Medical_Expenditure_SC_80D").value = MedicalExpSlfFamSrCtzn
            ElseIf temp80 = "N" Then
                Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "No"
                Sheet16.Range("Health_Insurance_80D").value = HealthInsPremSlfFam
                Sheet16.Range("Preventive_Health_80D").value = PrevHlthChckUpSlfFam
            ElseIf temp80 = "S" Then
                Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
            End If
        End If
            temp802 = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")("ParentSeniorCitizenFlag")
            
 'issue 11 solved
        If temp802 <> Empty Then
            If temp802 = "Y" Then
                Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Yes"
                Sheet16.Range("Health_Insurance3_80D").value = HlthInsPremParentsSrCtzn
                Sheet16.Range("Preventive_Health3_80D").value = PrevHlthChckUpParentsSrCtzn
                Sheet16.Range("Medical_Expenditure2_80D").value = MedicalExpParentsSrCtzn
            ElseIf temp802 = "N" Then
                Sheet16.Range("DropDown_ValueOf_SC_80D").value = "No"
                Sheet16.Range("Health_Insurance2_80D").value = HlthInsPremParents
                Sheet16.Range("Preventive_Health2_80D").value = PrevHlthChckUpParents
            ElseIf temp802 = "P" Then
                Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
            End If
        End If
                    
End Function
Function ImportScheduleBP_pfl(jsonText As String)
On Error Resume Next

    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
    
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim RecTDS1   As Variant
        
    Set jsonObject = ParseJson(jsonText)
  
    Set Nodelist = jsonObject("form3CD")("ScheduleBP")

   ' For Each Node In NodeList
           
        If Trim(Nodelist("PersumptiveInc44AD")("GrsTrnOverBank")) <> "" And Sheet3.Range("BP_E1a").Locked = False Then
            Sheet3.Range("BP_E1a").value = Nodelist("PersumptiveInc44AD")("GrsTrnOverBank")
        End If
        If Trim(Nodelist("PersumptiveInc44AD")("GrsTrnOverAnyOthMode")) <> "" And Sheet3.Range("BP_E1b").Locked = False Then
            Sheet3.Range("BP_E1b").value = Nodelist("PersumptiveInc44AD")("GrsTrnOverAnyOthMode")
        End If
        If Trim(Nodelist("PersumptiveInc44AD")("PersumptiveInc44AD6Per")) <> "" And Sheet3.Range("BP_E2a").Locked = False Then
            Sheet3.Range("BP_E2a").value = Nodelist("PersumptiveInc44AD")("PersumptiveInc44AD6Per")
        End If
        If Trim(Nodelist("PersumptiveInc44AD")("PersumptiveInc44AD8Per")) <> "" And Sheet3.Range("BP_E2b").Locked = False Then
            Sheet3.Range("BP_E2b").value = Nodelist("PersumptiveInc44AD")("PersumptiveInc44AD8Per")
        End If
        If Trim(Nodelist("PersumptiveInc44ADA")("GrsReceipt")) <> "" And Sheet3.Range("BP_E3_GR").Locked = False Then
            Sheet3.Range("BP_E3_GR").value = Nodelist("PersumptiveInc44ADA")("GrsReceipt")
        End If
        
        '2024-25 Change added by sai
        If Trim(Nodelist("PersumptiveInc44ADA")("grsTrnOverBank44ADA")) <> "" And Sheet3.Range("BP_E3_a").Locked = False Then
            Sheet3.Range("BP_E3_a").value = Nodelist("PersumptiveInc44ADA")("grsTrnOverBank44ADA")
        End If
        
        
        
        
        If Trim(Nodelist("PersumptiveInc44ADA")("TotPersumptiveInc44ADA")) <> "" And Sheet3.Range("BP_E4_PI44ADA").Locked = False Then
            Sheet3.Range("BP_E4_PI44ADA").value = Nodelist("PersumptiveInc44ADA")("TotPersumptiveInc44ADA")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("PartnerMemberOwnCapital")) <> "" And Sheet3.Range("BP_E9").Locked = False Then
            Sheet3.Range("BP_E9").value = Nodelist("FinanclPartclrOfBusiness")("PartnerMemberOwnCapital")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("SecuredLoans")) <> "" And Sheet3.Range("BP_E10").Locked = False Then
            Sheet3.Range("BP_E10").value = Nodelist("FinanclPartclrOfBusiness")("SecuredLoans")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("UnSecuredLoans")) <> "" And Sheet3.Range("BP_E11").Locked = False Then
            Sheet3.Range("BP_E11").value = Nodelist("FinanclPartclrOfBusiness")("UnSecuredLoans")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("Advances")) <> "" And Sheet3.Range("BP_E12").Locked = False Then
            Sheet3.Range("BP_E12").value = Nodelist("FinanclPartclrOfBusiness")("Advances")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("SundryCreditors")) <> "" And Sheet3.Range("BP_E13").Locked = False Then
            Sheet3.Range("BP_E13").value = Nodelist("FinanclPartclrOfBusiness")("SundryCreditors")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("OthrCurrLiab")) <> "" And Sheet3.Range("BP_E14").Locked = False Then
            Sheet3.Range("BP_E14").value = Nodelist("FinanclPartclrOfBusiness")("OthrCurrLiab")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("FixedAssets")) <> "" And Sheet3.Range("BP_E16").Locked = False Then
            Sheet3.Range("BP_E16").value = Nodelist("FinanclPartclrOfBusiness")("FixedAssets")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("Inventories")) <> "" And Sheet3.Range("BP_E17").Locked = False Then
            Sheet3.Range("BP_E17").value = Nodelist("FinanclPartclrOfBusiness")("Inventories")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("SundryDebtors")) <> "" And Sheet3.Range("BP_E18").Locked = False Then
            Sheet3.Range("BP_E18").value = Nodelist("FinanclPartclrOfBusiness")("SundryDebtors")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("BalWithBanks")) <> "" And Sheet3.Range("BP_E19").Locked = False Then
            Sheet3.Range("BP_E19").value = Nodelist("FinanclPartclrOfBusiness")("BalWithBanks")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("CashInHand")) <> "" And Sheet3.Range("BP_E20").Locked = False Then
            Sheet3.Range("BP_E20").value = Nodelist("FinanclPartclrOfBusiness")("CashInHand")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("LoansAndAdvances")) <> "" And Sheet3.Range("BP_E21").Locked = False Then
            Sheet3.Range("BP_E21").value = Nodelist("FinanclPartclrOfBusiness")("LoansAndAdvances")
        End If
        If Trim(Nodelist("FinanclPartclrOfBusiness")("OtherAssets")) <> "" And Sheet3.Range("BP_E22").Locked = False Then
            Sheet3.Range("BP_E22").value = Nodelist("FinanclPartclrOfBusiness")("OtherAssets")
        End If
  
    'Next Node
   
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







Function ImportPartA_139_8A(jsonText As String)

On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson(jsonText)
If Sheet201.Range("U_AadhaarCardNo").Locked = False Then
    Sheet201.Range("U_AadhaarCardNo").value = jsonObject("personalInfo")("aadhaarCardNo")
End If

'Ayush_enrol
'If Sheet201.Range("U_AadhaarEnrolmentId").Locked = False Then
'    Sheet201.Range("U_AadhaarEnrolmentId").value = jsonObject("personalInfo")("aadhaarCardNo")
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
''Malli-----AY_2025-26------SIT-85919
  If jsonObject("filingStatus")("origRetFiledDate") <> "" Then
    'Sheet201.Range("U_OrigRetFiledDate").value = jsonObject("filingStatus")("origRetFiledDate")
    Sheet201.Range("U_OrigRetFiledDate").value = Mid(jsonObject("filingStatus")("origRetFiledDate"), 9, 2) & "/" & Mid(jsonObject("filingStatus")("origRetFiledDate"), 6, 2) & "/" & Mid(jsonObject("filingStatus")("origRetFiledDate"), 1, 4)
    End If
   ' Sheet201.Range("U_OrigRetFiledDate").value = jsonObject("filingStatus")("origRetFiledDate")
   
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
   ' End If
    
     ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "3" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 24 to 36 Months from the end of Relevant Assessment  Year"
    'End If
    
     ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "4" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 36 to 48 Months from the end of Relevant Assessment  Year"
    End If
    
    
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
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2023-24"
'                ElseIf YearUD = "2024" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2024-25"
'                End If
'            End If
            
            'AY 2024-25-----------------
            'Ayush_UR
'             If YearUD <> "" Then
'                If YearUD = "2024" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2024-25"
'                ElseIf YearUD = "2025" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2025-26"
'                End If
'            End If
'            If YearUD <> "" Then
'                If YearUD = "2026" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2026-27"
'                ElseIf YearUD = "2025" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2025-26"
'                End If
'            End If
            'AY 2024-25--------------------
            
            
              If YearUD <> "" Then
                If YearUD = "2027" Then
                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2027-28"
                ElseIf YearUD = "2026" Then
                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2026-27"
                End If
            End If
            
            
            
            
            If Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If node("ReturnFiledEffectFlg") = "Y" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).value = "Yes"

                ElseIf node("ReturnFiledEffectFlg") = "N" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).value = "No"
                End If
             End If
             
              If Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If node("ReturnType") = "1" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Original Return 139(1)/139(4)"
                ElseIf node("ReturnType") = "2" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Revised return"
                ElseIf node("ReturnType") = "3" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Updated return"
                End If
             End If
        
        Next node
    
    End If
End If
End If

End Function
'Prefill-30.01.2023.103.ID
Sub ImportSchedule80GD_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Nodelist, node As Object
Dim Rect80G11, Rect80G12, cnt As Long
Dim TotalXMLRow, TotalExRow, TotalDiffRow, rowcount As Long
Dim Sch80GDoneeWithPanName, Sch80GAddrDetail, Sch80GcityOrTownOrDistrict, Sch80GStateCode, Sch80GpinCode, Sch80GDoneePAN, Sch80GDoneeARN, Sch80GDonationAmtCash, Sch80GDonationAmtOtherMode As Variant

        Set jsonObject = ParseJson(jsonText)
        Set Nodelist = jsonObject("Schedule80G")("Don50PercentApprReqd")("DoneeWithPan")
        
        Sch80GDoneeWithPanName = Sheet6.Range("Per5080G.DoneeName").Column
        Sch80GAddrDetail = Sheet6.Range("Per5080G.AddrDetail").Column
        Sch80GcityOrTownOrDistrict = Sheet6.Range("Per5080G.CityOrTownOrDistrict").Column
        Sch80GStateCode = Sheet6.Range("Per5080G.StateCode").Column
        Sch80GpinCode = Sheet6.Range("Per5080G.PinCode").Column
        Sch80GDoneePAN = Sheet6.Range("Per5080G.DoneePAN").Column
        Sch80GDoneeARN = Sheet6.Range("Per5080G.DoneeARN").Column
        Sch80GDonationAmtCash = Sheet6.Range("Per5080G.DonationCash").Column
        Sch80GDonationAmtOtherMode = Sheet6.Range("Per5080G.DonationAmt").Column


TotalExRow = Sheet6.Range("Per5080G.DoneeName").Rows.count
TotalXMLRow = Nodelist.count
TotalDiffRow = TotalXMLRow - TotalExRow

        If (TotalXMLRow > 0) Then
            Sheet6.Range("Per5080G.DoneeName").ClearContents
            Sheet6.Range("Per5080G.AddrDetail").ClearContents
            Sheet6.Range("Per5080G.CityOrTownOrDistrict").ClearContents
            Sheet6.Range("Per5080G.StateCode").ClearContents
            Sheet6.Range("Per5080G.PinCode").ClearContents
            Sheet6.Range("Per5080G.DoneePAN").ClearContents
            Sheet6.Range("Per5080G.DoneeARN").ClearContents
            Sheet6.Range("Per5080G.DonationCash").ClearContents
            Sheet6.Range("Per5080G.DonationAmt").ClearContents
        End If
        If (TotalDiffRow > 0) Then
            AddDiffRows_80G_D (TotalDiffRow)
        End If

rowcount = getRowNo(Sheet6.Range("Per5080G.DoneeName").name)
rowcount = rowcount - 1
        cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        
            If Sheet6.Cells(rowcount, Sch80GDoneeWithPanName).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GDoneeWithPanName).value = UCase(node("DoneeWithPanName"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GAddrDetail).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GAddrDetail).value = UCase(node("AddressDetail")("AddrDetail"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GcityOrTownOrDistrict).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GcityOrTownOrDistrict).value = UCase(node("AddressDetail")("CityOrTownOrDistrict"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GStateCode).Locked = False Then
                Dim iState As Variant
                iState = node("AddressDetail")("StateCode")
                If Len(iState) = "1" Then
                iState = "0" & iState
                End If
                If iState = "99" Then
                iState = ""
                End If
                Sheet6.Cells(rowcount, Sch80GStateCode).value = Findtext(iState, "StateList")
            End If
            
            If Sheet6.Cells(rowcount, Sch80GpinCode).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GpinCode).value = UCase(node("AddressDetail")("PinCode"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GDoneePAN).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GDoneePAN).value = UCase(node("DoneePAN"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GDoneeARN).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GDoneeARN).value = UCase(node("DoneeARN"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GDonationAmtCash).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GDonationAmtCash).value = (node("DonationAmtCash"))
            End If
            
            If Sheet6.Cells(rowcount, Sch80GDonationAmtOtherMode).Locked = False Then
                Sheet6.Cells(rowcount, Sch80GDonationAmtOtherMode).value = (node("DonationAmtOtherMode"))
            End If

        cnt = cnt + 1
    Next node
Rect80G11 = cnt
End Sub
'End Change

'Newly added by Sai
'Function ImportSchedule80U_pfl(jsonText As String)
'On Error Resume Next
'Dim init, jsonObject As Object
'Set jsonObject = ParseJson(jsonText)
'
'Dim dateOfFiling_New, acknNumber_New As Variant
'
'dateOfFiling_New = jsonObject("Form10IA")("dateOfFiling")
'acknNumber_New = jsonObject("Form10IA")("acknNumber")
'
'Section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType")
'
'If dateOfFiling <> "" Then
'Sheet14.Range("DatefilingFm10IA_80U").value = dateOfFiling_New  'DatefilingFm10IA_80DD
'End If
'If acknNumber <> "" Then
'        Sheet14.Range("AckNoFm10IAfiled_80U").value = acknNumber_New   'AckNoFm10IAfiled_80DD
'End If
'End Function
 
'Newly modified by Sai

Function ImportRefund_pfl(jsonText As String)
On Error Resume Next
 
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist, NodeTemp
Dim IFSCColNo, BankNameColNo, ACCNOColNo, CheckBox, AccountType_New As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim NodeMain, NodeListMain
 
Set jsonObject = ParseJson(jsonText)
 
 
Set Nodelist = New Collection
If jsonObject.exists("bankAccountDtls") Then
Set NodeListMain = jsonObject("bankAccountDtls")
If NodeListMain <> Empty Then
Dim flag
For Each NodeMain In NodeListMain
    For Each node In NodeMain("addtnlBankDetails")
        flag = 0
        For Each NodeTemp In Nodelist
            If NodeTemp("bankAccountNo") = node("bankAccountNo") Then
                flag = 1
            End If
        Next NodeTemp
        If flag = 0 Then
            Nodelist.add node
        End If
    Next node
Next NodeMain
End If
End If
 
 
'Malli_AY_2025-26
'comented by Mallikarjun as per AY_2025_26 removed 2nd prefill source
'Set NodeListMain = jsonObject("lastFiledITR")("bankAccountDtls")
'If NodeListMain <> Empty Then
'For Each NodeMain In NodeListMain
'    If NodeMain.exists("addtnlBankDetails") Then
'    For Each node In NodeMain("addtnlBankDetails")
'        flag = 0
'        For Each NodeTemp In Nodelist
'            If NodeTemp("bankAccountNo") = node("bankAccountNo") Then
'                flag = 1
'            End If
'        Next NodeTemp
'        If flag = 0 Then
'            Nodelist.add node
'        End If
'    Next node
'    End If
'Next NodeMain
'End If
 
 '---------------------------------
 
 
 'Malli_AY_2025_26 'un comented by Malli
    'issue resolved
    Set NodeListMain = jsonObject("bankAccountDtls")
    'issue resolved
    For Each NodeMain In NodeListMain

        Set Nodelist = NodeMain("addtnlBankDetails")

        IFSCColNo = Sheet5.Range("SchBA.IFSC").Column
        BankNameColNo = Sheet5.Range("SchBA.BankName").Column
        ACCNOColNo = Sheet5.Range("SchBA.AcntNo").Column

        AccountType_New = Sheet5.Range("SchBA.Type_of_Account").Column
        TotalExRow = Range("SchBA.IFSC").Rows.count
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow

        If (TotalXMLRow > 0) Then
            Sheet5.Range("SchBA.IFSC").ClearContents
            Sheet5.Range("SchBA.BankName").ClearContents
            Sheet5.Range("SchBA.AcntNo").ClearContents
            
            Sheet5.Range("SchBA.Type_of_Account").ClearContents 'Newly added by Sai
        End If
        If (TotalDiffRow > 0) Then
         AddDiffRows_BANK (TotalDiffRow)
        End If
        rowcount = getRowNo(Sheet5.Range("SchBA.IFSC").name)
        rowcount = rowcount - 1
        cnt = 0
        For Each node In Nodelist
            rowcount = rowcount + 1
                If Trim(node("ifsccode")) <> "" And Sheet5.Cells(rowcount, IFSCColNo).Locked = False Then
                    Sheet5.Cells(rowcount, IFSCColNo).value = node("ifsccode")
                End If
                If Trim(node("bankName")) <> "" And Sheet5.Cells(rowcount, BankNameColNo).Locked = False Then
                    Sheet5.Cells(rowcount, BankNameColNo).value = node("bankName")
                End If
                If Trim(node("bankAccountNo")) <> "" And Sheet5.Cells(rowcount, ACCNOColNo).Locked = False Then
                    Sheet5.Cells(rowcount, ACCNOColNo).value = node("bankAccountNo")
                End If

                Dim AcType As Variant
                AcType = node("AccountType")
                If AcType = "SB" Then
                    Sheet5.Cells(rowcount, AccountType_New).value = "Savings Account"
                    ElseIf AcType = "CA" Then
                    Sheet5.Cells(rowcount, AccountType_New).value = "Current Account"
                    ElseIf AcType = "CC" Then
                    Sheet5.Cells(rowcount, AccountType_New).value = "Cash Credit Account"
                    ElseIf AcType = "OD" Then
                    Sheet5.Cells(rowcount, AccountType_New).value = "Over draft account"
                    ElseIf AcType = "NRO" Then
                    Sheet5.Cells(rowcount, AccountType_New).value = "Non Resident Account"
                    ElseIf AcType = "OTH" Then
                    Sheet5.Cells(rowcount, AccountType_New).value = "Other"
                End If


            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
   Next NodeMain
End Function

Function ImportSchedule80DD_80U_pfl(jsonText As String)
On Error Resume Next
Dim init, jsonObject As Object
Set jsonObject = ParseJson(jsonText)

Dim schedule80DDObject As Object
Dim dob As Variant
 
 'Malli_AY_2025_26
'Dim dateOfFiling, acknNumber As Object
Dim dateOfFiling, acknNumber As Variant
 
'Malli_commented_AY_2025 -26
'dateOfFiling = jsonObject("Form10IA")("dateOfFiling")
acknNumber = jsonObject("Form10IA")("acknNumber")
 
 
'Malli_commented_AY_2025 -26
 
'If dateOfFiling <> "" Then
'
'dob = Mid(dateOfFiling, 9, 2) & "/" & Mid(dateOfFiling, 6, 2) & "/" & Mid(dateOfFiling, 1, 4)
'If Sheet14.Range("Form10IA_80DD").Locked = False Then
'   Sheet14.Range("Form10IA_80DD").value = dob      'DatefilingFm10IA_80DD
'End If
'
'If Sheet14.Range("DateFiling10IA_80U").Locked = False Then
'   Sheet14.Range("DateFiling10IA_80U").value = dob     'DatefilingFm10IA_80U
'End If
'
'End If

'--------------------



''Malli_C12_80DD_AY_2026-27
'If acknNumber <> "" And Sheet14.Range("AcknowledgeNum_80DD").Locked = False Then
'        Sheet14.Range("AcknowledgeNum_80DD").value = acknNumber   'AckNoFm10IAfiled_80DD
'End If
''--------------------
'
''Malli_C9_80U_AY_2026-27
'If acknNumber <> "" And Sheet14.Range("AckNumber10IA_80U").Locked = False Then
'Sheet14.Range("AckNumber10IA_80U").value = acknNumber            'AckNoFm10IAfiled_80U
'End If
''---------------------
End Function




'Malli_AY_2025_26
Sub lockwithfirm_A23abcfields()
On Error Resume Next

        If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
            Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.DateOfFiling").value = ""
         End If
         
         If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
            Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.AcknowledgementNo").value = ""
         End If
         'TableA
         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
         End If
         If Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
            Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
         End If
         If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
            Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
         End If
         'TableB
         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
         End If
         If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
            Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
         End If
         If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
            Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
         End If
         
         'TableC
         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
         End If
         If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
            Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
         End If
         If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked <> True Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
            Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
            Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
         End If
         
End Sub




Function ImportFilingStatus_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim ReturnFileSec, AsseseeRepFlg, node, Nodelist, NewTaxRegime, Form10IEDate, Form10IEAckNo, section139, depAmtAggAmt, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar
Dim uniqueNo, acknNumber, dateOfFiling, YYYY, MM, DD
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
    
'    Dim AYOptedIn, AYOptedOut, NewForm10IEDate, NewForm10IEAckNo, OptingNewTaxRegime, OldTaxRegime, OldForm10IEDate, OldForm10IEAckNo
    Dim clauseiv7provisio139i, clauseiv7provisio139iNature, clauseiv7provisio139iAmount
    
    
    
    
    ReturnFileSec = jsonObject("filingStatus")("returnFileSec")
    
    ReceiptNo = jsonObject("filingStatus")("receiptNo")
    
    OrigRetFiledDate = jsonObject("filingStatus")("origRetFiledDate")
    NoticeDateUnderSec = jsonObject("filingStatus")("noticeDateUnderSec")
    
    section139 = jsonObject("filingStatus")("SeventhProvisio139")
    '11/09/2024 Chetan C M changes start
    'depAmtAggAmt = jsonObject("filingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
    depAmtAggAmt = jsonObject("insights")("DepAmtAggAmtExcd1CrPrYrFlg")
    '11/09/2024 Chetan C M changes end
    uniqueNo = jsonObject("filingStatus")("uniqueNo")
    
    'NewTaxRegime = jsonObject("filingReturn")("OptedFor115BAC")
    'AcknNumber = jsonObject("form10IE")("AcknNumber")
    'DateOfFiling = jsonObject("form10IE")("DateOfFiling")
    
    'Prefill-30.01.2023.103.ID
    
    
    
    
    
'    filingStatus.OptOutNewTaxRegime

    'For 2024-25 Prefill Development Change added by Sai--------------------------------------------------------------
    
     Dim OptOutNewTaxRegime_New, Form10IEADate_New, Form10IEAckNo_New
    
     OptOutNewTaxRegime_New = jsonObject("filingStatus")("OptOutNewTaxRegime")
     
    'Form10IEA.dateOfFiling
     Form10IEADate_New = jsonObject("Form10IEA")("dateOfFiling")
     
     'Form10IEA.acknNumber
     Form10IEAckNo_New = jsonObject("Form10IEA")("acknNumber")
     
     
If Form10IEADate_New <> "" Then
    Dim YYYY1, MM1, DD1
    YYYY1 = Mid(Form10IEADate_New, 1, 4)
    MM1 = Mid(Form10IEADate_New, 6, 2)
    DD1 = Mid(Form10IEADate_New, 9, 2)
    dateOfFiling = DD1 & "/" & MM1 & "/" & YYYY1
End If
'
'    Dim Date10IEA As varaiant
'    Date10IEA = Mid(Form10IEADate_New, 9, 2) & "/" & Mid(Form10IEADate_New, 6, 2) & "/" & Mid(Form10IEADate_New, 1, 4)
     
     '-----------------------------Old One---------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'    NewTaxRegime = jsonObject("filingStatus")("NewTaxRegime")
'    AYOptedIn = jsonObject("filingStatus")("AYOptedIn")
'    AYOptedOut = jsonObject("filingStatus")("AYOptedOut")

'    NewForm10IEDate = jsonObject("filingStatus")("DateOfFiling")
'    NewForm10IEAckNo = jsonObject("filingStatus")("NewForm10IEAckNo")

'    OldTaxRegime = jsonObject("filingStatus")("OldTaxRegime")
'    OldForm10IEDate = jsonObject("filingStatus")("OldForm10IEDate")
'    OldForm10IEAckNo = jsonObject("filingStatus")("OldForm10IEAckNo")
'    OptingNewTaxRegime = jsonObject("filingStatus")("OptingNewTaxRegime")
'    Form10IEDate = jsonObject("filingStatus")("Form10IEDate")
'    Form10IEAckNo = jsonObject("filingStatus")("Form10IEAckNo")

     '-----------------------------Old One---------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
     'commented by Chetan C M AY 2024-25 old SIT-86604
     '--start--
'    clauseiv7provisio139i = jsonObject("filingStatus")("clauseiv7provisio139i")
'    clauseiv7provisio139iNature = jsonObject("filingStatus")("clauseiv7provisio139iDtls")("clauseiv7provisio139iNature")
'    clauseiv7provisio139iAmount = jsonObject("filingStatus")("clauseiv7provisio139iDtls")("clauseiv7provisio139iAmount")
    '--end--
    'End Change
    
    Dim repEmailID, repMobileNo As Variant
    
    If jsonObject.exists("assesseeRep") Then
        RepName = jsonObject("assesseeRep")("repName")
        
        'Malli commented PAG_C107to110_AY_2026-27
'        RepCapacity = jsonObject("assesseeRep")("repCapacity")
'        RepAddress = jsonObject("assesseeRep")("repAddress")
'        RepPAN = jsonObject("assesseeRep")("repPAN")
'         RepAadhaar = jsonObject("assesseeRep")("aadhaarCardNo")
        '-------------------------------------------
        'Malli PAG_C111&112_AY_2026-27
        repEmailID = jsonObject("assesseeRep")("repEmailID")
        repMobileNo = jsonObject("assesseeRep")("repMobileNo")
        '----------------------------
        If RepName <> "" Then
            Sheet1.Range("sheet1.RepAssessee") = "Yes"
        Else
            Sheet1.Range("sheet1.RepAssessee") = "No"
        End If
    Else
        Sheet1.Range("sheet1.RepAssessee") = "No"
        
    End If
    
    
    
    
    
        iReturnFile = ReturnFileSec
Dim Rdbtn As Variant
'SIT-67940 AY_2024_25  Malli------------------------------
            
            
            'Rdbtn = Sheet1.Range("Radiobuttoncheck").value
            
           'Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
           'Sheet1.Shapes("option button 929").OLEFormat.Object.value = 1
'         If Rdbtn = 1 Then
'                  If iReturnFile = "11" Then
'                  sReturnFile = "139(1)-On or before due date"
'                  ElseIf iReturnFile = "12" Then
'                  sReturnFile = "139(4)-After due date"
'                  ElseIf iReturnFile = "17" Then
'                  sReturnFile = "139(5)-Revised Return"
'                  ElseIf iReturnFile = "20" Then
'                  sReturnFile = "119(2)(b)- after condonation of delay"
'                  ElseIf iReturnFile = "21" Then
'                  sReturnFile = "139(8A)"
'                  End If
'
'            ElseIf Rdbtn = 2 Then
'
'                  If iReturnFile = "18" Then
'                  sReturnFile = "139(9)"
'                  ElseIf iReturnFile = "13" Then
'                  sReturnFile = "142(1)"
'                  ElseIf iReturnFile = "14" Then
'                  sReturnFile = "148"
'                  ElseIf iReturnFile = "16" Then
'                  sReturnFile = "153C"
'                   End If
'
'            End If
             
            
         '--------------Jan29th
         'iReturnFile = 21
         
If iReturnFile = "11" Or iReturnFile = "12" Or iReturnFile = "17" Or iReturnFile = "20" Or iReturnFile = "21" Then
         Sheet1.Shapes("option button 929").OLEFormat.Object.value = 1
         'Dim Rdbtn As Variant
         Rdbtn = Sheet1.Range("Radiobuttoncheck").value
         If Rdbtn = 1 Then
                  If iReturnFile = "11" Then
                  sReturnFile = "139(1)-On or before due date"
                  ElseIf iReturnFile = "12" Then
                  sReturnFile = "139(4)-After due date"
                  ElseIf iReturnFile = "17" Then
                  sReturnFile = "139(5)-Revised Return"
                  ElseIf iReturnFile = "20" Then
                  'Ayush J200924
                  sReturnFile = "119(2)(b)- After condonation of delay"
'                    sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
                  ElseIf iReturnFile = "21" Then
                  sReturnFile = "139(8A)"
                  End If
          End If
 End If
            
  If iReturnFile = "18" Or iReturnFile = "13" Or iReturnFile = "14" Or iReturnFile = "16" Then
             Sheet1.Shapes("option button 930").OLEFormat.Object.value = 1
             'Dim Rdbtn As Variant
             Rdbtn = Sheet1.Range("Radiobuttoncheck").value
             
            If Rdbtn = 2 Then
            
                  If iReturnFile = "18" Then
                  sReturnFile = "139(9)"
                  ElseIf iReturnFile = "13" Then
                  sReturnFile = "142(1)"
                  ElseIf iReturnFile = "14" Then
                  sReturnFile = "148"
                  ElseIf iReturnFile = "16" Then
                  sReturnFile = "153C"
                  End If
               
             End If
    End If
         
         '------------------
            
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
'        ElseIf iReturnFile = "21" Then
'           sReturnFile = "139(8A)"
'        End If
     'Malli----------SIT_93692
        If sReturnFile <> "" And Sheet1.Range("sheet1.ReturnFileSec").Locked = False Then
            Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
        End If
        
   '----------------------------------------------------------------------------------
        iProvisoFlag = section139
        If iProvisoFlag = "Y" Then
            sProvisoFlag = "Yes"
        ElseIf iProvisoFlag = "N" Then
            sProvisoFlag = "No"
        End If
        If Sheet1.Range("sheet1.SeventhProvisoFlag").Locked = False Then
            Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
        End If
        
        If depAmtAggAmt = "Y" Then
            depAmtAggAmt = "Yes"
        ElseIf depAmtAggAmt = "N" Then
            depAmtAggAmt = "No"
        End If
        If Sheet1.Range("Sheet1.DepositAmountFlag").Locked = False Then
            Sheet1.Range("Sheet1.DepositAmountFlag").value = depAmtAggAmt
        End If
        'Konda updated SIT-85916---------10-01-2025
        AmtSeventhProvisio139i = jsonObject("insights")("AmtSeventhProvisio139i")
        If Sheet1.Range("Sheet1.DepositAmount").Locked = False Then
            Sheet1.Range("Sheet1.DepositAmount").value = AmtSeventhProvisio139i
        End If
        'end update--------SIT-85916-----10-01-2025

        
        IncrExpAggAmt2LkTrvFrgnCntryFlg = jsonObject("form26as")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
        AmtSeventhProvisio139ii = jsonObject("form26as")("AmtSeventhProvisio139ii")
        
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
        

    'issue 3 solved
'-----------------------------------------------------------------------Old commented by sai

'If NewTaxRegime = "Y" Then
'            Sheet8.Unprotect Password:=getmsgstate
'                Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
'                Sheet8.Range("bacValue").value = 1
'            Sheet8.Protect Password:=getmsgstate
'            Sheet1.Unprotect Password:=getmsgstate
'                resetBacYes
'            Sheet1.Protect Password:=getmsgstate
'
'     ElseIf NewTaxRegime = "N" Then
'            Sheet8.Unprotect Password:=getmsgstate
'                Sheet8.Range("bacValue").value = 2
'                Sheet1.Shapes("BacNO").OLEFormat.Object.value = 1
'            Sheet8.Protect Password:=getmsgstate
'            Sheet1.Unprotect Password:=getmsgstate
'                resetBacNo
'            Sheet1.Protect Password:=getmsgstate
'     Else
'            Sheet8.Shapes("BacYes").OLEFormat.Object.value = 0
'            Sheet8.Shapes("BacNO").OLEFormat.Object.value = 0
'            Sheet1.Unprotect Password:=getmsgstate
'            resetBacNo
'            Sheet1.Protect Password:=getmsgstate
'End If

'-------------------------------------------------------------------------
        
        
    'Malli_AY_2025_26
   
    'PAG_30 New change
   
   'Form10IEA.OptOutNewTaxRegime_AY24_25
   'Sheet1.115BAC
' If jsonObject("Form10IEA") <> Null Then
'
'     Dim OptOutNewTaxRegime_AY24_25, Form10IEADate_AY24_25, Form10IEAAckNo_AY24_25
'
'     OptOutNewTaxRegime_AY24_25 = jsonObject("Form10IEA")("OptOutNewTaxRegime_AY24_25")
'     Form10IEADate_AY24_25 = jsonObject("Form10IEA")("Form10IEADate_AY24_25")
'     Form10IEAAckNo_AY24_25 = jsonObject("Form10IEA")("Form10IEAAckNo_AY24_25")
'
'     If OptOutNewTaxRegime_AY24_25 <> "" Then
'
'          If OptOutNewTaxRegime_AY24_25 = "N" Then
'                If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'                  Sheet1.Range("Sheet1.115BAC").value = "No"
'                End If
'
'
'
'          ElseIf OptOutNewTaxRegime_AY24_25 = "Y" Then
'                If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'                  Sheet1.Range("Sheet1.115BAC").value = "Yes"
'                End If
'                If Form10IEADate_AY24_25 <> "" And Sheet1.Range("BAC115.DateOfFiling").Locked = False Then
'                   Dim Form10IEADate_AY24_25_D As Variant
'                       Form10IEADate_AY24_25_D = Mid(Form10IEADate_AY24_25, 9, 2) & "/" & Mid(Form10IEADate_AY24_25, 6, 2) & "/" & Mid(Form10IEADate_AY24_25, 1, 4)
'                       Sheet1.Range("BAC115.DateOfFiling").value = Form10IEADate_AY24_25_D
'                End If
'                If Form10IEAAckNo_AY24_25 <> "" And Sheet1.Range("BAC115.AcknowledgementNo").Locked = False Then
'                Sheet1.Range("BAC115.AcknowledgementNo").value = Form10IEAAckNo_AY24_25
'                End If
'          ElseIf OptOutNewTaxRegime_AY24_25 = "NA" And (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
'               If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("Sheet1.115BAC").Validation.Delete
'                                Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NA"
'                                Application.EnableEvents = True
'                                Sheet1.Range("Sheet1.115BAC").value = "Not Applicable"
'                                Application.EnableEvents = False
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = True
'
'               End If
'          End If
'     End If
'End If


'Malli-----------

''comented by Malli version V0.2 change 12/02/2024
'
'If jsonObject("Form10IEA") <> Null Then
'
'     Dim OptOutNewTaxRegime_AY24_25, Form10IEADate_AY24_25, Form10IEAAckNo_AY24_25
'
'     OptOutNewTaxRegime_AY24_25 = jsonObject("Form10IEA")("OptOutNewTaxRegime_AY24_25")
'     Form10IEADate_AY24_25 = jsonObject("Form10IEA")("Form10IEADate_AY24_25")
'     Form10IEAAckNo_AY24_25 = jsonObject("Form10IEA")("Form10IEAAckNo_AY24_25")
'
'     If OptOutNewTaxRegime_AY24_25 <> "" Then
'
'          If OptOutNewTaxRegime_AY24_25 = "N" Then
'                If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'                  Sheet1.Range("Sheet1.115BAC").value = "No"
'                End If
'
'          ElseIf OptOutNewTaxRegime_AY24_25 = "Y" Then
'                If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'                  Sheet1.Range("Sheet1.115BAC").value = "Yes"
'                End If
'
'
'          ElseIf OptOutNewTaxRegime_AY24_25 = "NA" And (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
'               If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("Sheet1.115BAC").Validation.Delete
'                                Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NA"
'                                Application.EnableEvents = True
'                                Sheet1.Range("Sheet1.115BAC").value = "Not Applicable"
'                                Application.EnableEvents = False
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = True
'
'               End If
'          End If
'     End If
'
'   If OptOutNewTaxRegime_AY24_25 = "Y" And Sheet1.Range("Sheet1.115BAC").value = "Yes" Then
'              If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
'                 Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("BAC115.DateOfFiling").value = ""
'              End If
'              If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
'                 Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("BAC115.AcknowledgementNo").value = ""
'              End If
'
'                If Form10IEADate_AY24_25 <> "" And Sheet1.Range("BAC115.DateOfFiling").Locked = False Then
'                   Dim Form10IEADate_AY24_25_D As Variant
'                       Form10IEADate_AY24_25_D = Mid(Form10IEADate_AY24_25, 9, 2) & "/" & Mid(Form10IEADate_AY24_25, 6, 2) & "/" & Mid(Form10IEADate_AY24_25, 1, 4)
'                       Sheet1.Range("BAC115.DateOfFiling").value = Form10IEADate_AY24_25_D
'                End If
'
'                If Form10IEAAckNo_AY24_25 <> "" And Sheet1.Range("BAC115.AcknowledgementNo").Locked = False Then
'                Sheet1.Range("BAC115.AcknowledgementNo").value = Form10IEAAckNo_AY24_25
'                End If
'
'     Else:
'         If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked <> True Then
'             Sheet1.Unprotect Password:=getmsgstate
'             Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
'             Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
'             Sheet1.Range("BAC115.DateOfFiling").value = ""
'         End If
'
'         If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked <> True Then
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
'            Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
'            Sheet1.Range("BAC115.AcknowledgementNo").value = ""
'         End If
'
'     End If
'
'End If



'Malli commented AY_2026_27 03/04/2026
'--------Version2 changes:----------

'dateOfFiling ,acknNumber ,filedForm10IEA ,OptOutNewTaxRegime_AY24_25,Form10IEADate_AY24_25,Form10IEAAckNo_AY24_25
'Form10IEA

Dim Form10IEA_dateOfFiling, Form10IEA_acknNumber
Dim filedForm10IEA, OptOutNewTaxRegime_AY24_25, Form10IEADate_AY24_25, Form10IEAAckNo_AY24_25
''Mid(, 9, 2) & "/" & Mid(, 6, 2) & "/" & Mid(, 1, 4)
'  = jsonObject("")("")
'OptOutNewTaxRegime_AY24_25 = jsonObject("Form10IEA")("OptOutNewTaxRegime_AY24_25")
'Form10IEADate_AY24_25 = jsonObject("Form10IEA")("Form10IEADate_AY24_25")
'Form10IEAAckNo_AY24_25 = jsonObject("Form10IEA")("Form10IEAAckNo_AY24_25")
'
'filedForm10IEA = jsonObject("Form10IEA")("filedForm10IEA")
'Form10IEA_dateOfFiling = jsonObject("Form10IEA")("dateOfFiling")
'Form10IEA_acknNumber = jsonObject("Form10IEA")("acknNumber")
'
'
'Dim Form10IEADateAY24_25Dfmt, Form10IEA_dateOfFiling_Dfmt As Variant
'
'Form10IEADateAY24_25Dfmt = Mid(Form10IEADate_AY24_25, 9, 2) & "/" & Mid(Form10IEADate_AY24_25, 6, 2) & "/" & Mid(Form10IEADate_AY24_25, 1, 4)
'Form10IEA_dateOfFiling_Dfmt = Mid(Form10IEA_dateOfFiling, 9, 2) & "/" & Mid(Form10IEA_dateOfFiling, 6, 2) & "/" & Mid(Form10IEA_dateOfFiling, 1, 4)
'
'If Left(Sheet1.Range("sheet1.Status").value, 1) <> "F" Then
'If OptOutNewTaxRegime_AY24_25 <> "" Then
'
'If OptOutNewTaxRegime_AY24_25 = "Y" Or UCase(OptOutNewTaxRegime_AY24_25) = "YES" Then
''MsgBox UCase(OptOutNewTaxRegime_AY24_25)
'            If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'              Sheet1.Range("Sheet1.115BAC").value = "Yes"
'            End If
'
'                If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
'                 Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("BAC115.DateOfFiling").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
'                 Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("BAC115.AcknowledgementNo").value = ""
'                End If
'
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False
'                 Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
'                 Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked <> False Then
'                 Sheet1.Unprotect Password:=getmsgstate
'                 Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
'                 Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
'                 Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
'                End If
'
'                 If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'                End If
'
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
'                End If
'
'         If Form10IEADate_AY24_25 <> "" Then
'                    If Sheet1.Range("BAC115.DateOfFiling").Locked = False Then
'                        Sheet1.Range("BAC115.DateOfFiling").value = Form10IEADateAY24_25Dfmt
'                    End If
'         End If
'
'         If Form10IEAAckNo_AY24_25 <> "" Then
'                    If Sheet1.Range("BAC115.AcknowledgementNo").Locked = False Then
'                            Sheet1.Range("BAC115.AcknowledgementNo").value = Form10IEAAckNo_AY24_25
'                    End If
'         End If
'
'         If filedForm10IEA <> "" Then
'
'                    If filedForm10IEA = "Y" Or UCase(filedForm10IEA) = "YES" Then
'                        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Locked = False Then
'                                 Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "Yes"
'                        End If
'                            If Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked <> True Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'                                Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'                            End If
'                            If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked <> True Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
'                            End If
'                                    OldTaxRegime_New
'
'                    ElseIf filedForm10IEA = "N" Or UCase(filedForm10IEA) = "NO" Then
'
'                            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Locked = False Then
'                                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No"
'                            End If
'                            If Sheet1.Range("BAC115.DateOfFiling_New_Yes").Locked = False Then
'                                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = Form10IEA_dateOfFiling_Dfmt
'                            End If
'                            If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Locked = False Then
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = Form10IEA_acknNumber
'                            End If
'                                    NewTaxRegime_New
'                    End If
'         End If
'
'ElseIf OptOutNewTaxRegime_AY24_25 = "N" Or UCase(OptOutNewTaxRegime_AY24_25) = "NO" Then
'
'            If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'              Sheet1.Range("Sheet1.115BAC").value = "No"
'            End If
'                'Unlock
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked <> False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
'                End If
'                If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked <> False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
'                Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
'                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked <> False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'                End If
'                'Lock
'
'                If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo").value = ""
'                End If
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
'                End If
'                If Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
'                End If
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
'                End If
'                If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
'                End If
'
'              If filedForm10IEA <> "" Then
'
'                    If filedForm10IEA = "Y" Or UCase(filedForm10IEA) = "YES" Then
'                        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Locked = False Then
'                                 Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes"
'                        End If
'                        If Sheet1.Range("BAC115.DateOfFiling_New_No").Locked = False Then
'                            Sheet1.Range("BAC115.DateOfFiling_New_No").value = Form10IEA_dateOfFiling_Dfmt
'                        End If
'                        If Sheet1.Range("BAC115.AcknowledgementNo_New_No").Locked = False Then
'                            Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = Form10IEA_acknNumber
'                        End If
'                               ' NewTaxRegime_New
'                                 OldTaxRegime_New
'                    ElseIf filedForm10IEA = "N" Or UCase(filedForm10IEA) = "NO" Then
'
'                            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Locked = False Then
'                                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "No"
'                            End If
'
'                            If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked <> True Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
'                                Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'                            End If
'                            If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked <> True Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'                            End If
'                                     NewTaxRegime_New
'                    End If
'         End If
'
'ElseIf OptOutNewTaxRegime_AY24_25 = "NA" Or UCase(OptOutNewTaxRegime_AY24_25) = "NOT APPLICABLE" Then
'          If Sheet1.Range("Sheet1.115BAC").Locked = False Then
'            Sheet1.Range("Sheet1.115BAC").value = "Not Applicable"
'          End If
'                'unlock
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked <> False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HCCFFCC)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked <> False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
'                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked <> False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
'                End If
'                'Lock
'                If Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling").value = ""
'                End If
'
'                If Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo").value = ""
'                End If
'
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
'                End If
'                If Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
'                End If
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
'                End If
'                If Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
'                Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'                End If
'                If Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked <> True Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'                End If
'
'              If filedForm10IEA <> "" Then
'
'                    If filedForm10IEA = "Y" Or UCase(filedForm10IEA) = "YES" Then
'                        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Locked = False Then
'                                 Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes"
'                        End If
'                        If Sheet1.Range("BAC115.DateOfFiling_New_NA").Locked = False Then
'                            Sheet1.Range("BAC115.DateOfFiling_New_NA").value = Form10IEA_dateOfFiling_Dfmt
'                        End If
'                        If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Locked = False Then
'                            Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = Form10IEA_acknNumber
'                        End If
'
'                                       OldTaxRegime_New
'                    ElseIf filedForm10IEA = "N" Or UCase(filedForm10IEA) = "NO" Then
'
'                            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Locked = False Then
'                                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "No"
'                            End If
'
'                            If Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked <> True Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
'                                Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
'                            End If
'                            If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked <> True Then
'                                Sheet1.Unprotect Password:=getmsgstate
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
'                                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
'                            End If
'                                     NewTaxRegime_New
'                    End If
'              End If
'End If
'End If
'
'ElseIf Left(Sheet1.Range("sheet1.Status").value, 1) = "F" Then
'            Sheet1.Unprotect Password:=getmsgstate
'            Application.EnableEvents = False
'            Sheet1.Range("Sheet1.115BAC").Validation.Delete
'            Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NA"
'            Application.EnableEvents = True
'            Sheet1.Range("Sheet1.115BAC").value = "NA"
'            Application.EnableEvents = False
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HD8D8D8)
'            Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = True
'
'            'unlock
'            Call lockwithfirm_A23abcfields
'
'End If
'
'
'
''--------------------------------------




'------------------




    '---------------------
    
    'Malli-------------------Malli_AY_2025_26
        
 'For 115BAC(6) Prefill 2024-25 added by Sai------------------
 
          'OptOutNewTaxRegime_1 = jsonObject("ITR")("ITR4")("FilingStatus")("OptOutNewTaxRegime")
          'OptOutNewTaxRegime_New = jsonObject("filingStatus")("OptOutNewTaxRegime")
          'Date10IEA
   
   
   'comented by Mallikarjun as per AY_2025-26 development
   'start---------------------------
'            If OptOutNewTaxRegime_New = "N" Then
'              Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No"
'              NewBacYesValueChange
'
'            ElseIf OptOutNewTaxRegime_New = "Y" Then
'               Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date"
'               NewBacYesValueChange
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False Then
'               Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("Sheet1.Dateoffiling").value = dateOfFiling
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
'             Sheet1.Range("Sheet1.acknowledgementNumber_New").value = Form10IEAckNo_New
'            End If
'
'
'            Sheet1.Unprotect Password:=getmsgstate
'            ElseIf OptOutNewTaxRegime_New = "B" Then
'             Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date"
'             NewBacNoValueChange_2425
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("Sheet1.Dateoffiling").Locked = False Then
'                Sheet1.Range("Sheet1.Dateoffiling").value = dateOfFiling
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").value = Form10IEAckNo_New
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'            ElseIf OptOutNewTaxRegime_New = "NA" Then
'            'Ayush
'          '  Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Not Applicable"
'          If (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
'                        Sheet1.Unprotect Password:=getmsgstate
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.Delete
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current5"
'                        Application.EnableEvents = True
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Not Applicable"
'                        Application.EnableEvents = False
'                        Sheet1.Unprotect Password:=getmsgstate
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Interior.Color = (&HD8D8D8)
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").MergeArea.Locked = True
'
'               Else:
'                  Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = ""
'               End If
'               '-----------
'            End If
    '-----------------------------------end
            
'Malli_________________________________

'"description": "Y :Old Tax regime ; N :New Tax Regime ; NA:Not Applicable ",
                    
'comented by Mallikarjun version Vo.2 changes 12/02/2025
' If OptOutNewTaxRegime_New = "N" Then
'
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime"
'                NewBacYesValueChange
'
'                If Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.Dateoffiling").value = dateOfFiling
'                End If
'
'                Sheet1.Unprotect Password:=getmsgstate
'                If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
'                Sheet1.Range("Sheet1.acknowledgementNumber_New").value = Form10IEAckNo_New
'                End If
'
' ElseIf OptOutNewTaxRegime_New = "Y" Then
'                Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime"
'                NewBacNoValueChange_2425
'
'                If Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("Sheet1.Dateoffiling").value = dateOfFiling
'                End If
'
'                Sheet1.Unprotect Password:=getmsgstate
'                If Sheet1.Range("Sheet1.acknowledgementNumber_New").Locked = False Then
'                Sheet1.Range("Sheet1.acknowledgementNumber_New").value = Form10IEAckNo_New
'                End If
'
' ElseIf OptOutNewTaxRegime_New = "NA" Then
'                If (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
'                        Sheet1.Unprotect Password:=getmsgstate
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.Delete
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current5"
'                        Application.EnableEvents = True
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Not Applicable"
'                        Application.EnableEvents = False
'                        Sheet1.Unprotect Password:=getmsgstate
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").Interior.Color = (&HD8D8D8)
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").MergeArea.Locked = True
'
'                Else:
'                        Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = ""
'                End If
'                '-----------
'  End If
'
  
    
           
         
            
            
            
'______________________________________
'----------------------------------------Malli
            
 'Change end for prefill 115BAC --------------------------------
 
'Prefill-30.01.2023.103.ID
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
'            Application.EnableEvents = False
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
'    If Sheet1.Range("OptingNewTaxRegime").Locked = False Then
'            If OptingNewTaxRegime = "1" Then
'                Sheet1.Range("OptingNewTaxRegime").value = "Opting in now"
'            ElseIf OptingNewTaxRegime = "2" Then
'                Sheet1.Range("OptingNewTaxRegime").value = "Not opting"
'            ElseIf OptingNewTaxRegime = "3" Then
'                Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt"
'            ElseIf OptingNewTaxRegime = "4" Then
'                Sheet1.Range("OptingNewTaxRegime").value = "Opt out"
'            ElseIf OptingNewTaxRegime = "5" Then
'                Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in"
'            Else
'                Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'            End If
'    End If
'
'    If Sheet1.Range("Sheet1.date10IE").Locked = False Then
'        Sheet1.Range("Sheet1.date10IE").value = Mid(Form10IEDate, 9, 2) & "/" & Mid(Form10IEDate, 6, 2) & "/" & Mid(Form10IEDate, 1, 4)
'    End If
'
'    If Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False Then
'        Sheet1.Range("Sheet1.acknowledgementNumber").value = UCase(Form10IEAckNo)
'    End If





'//* Malli_AY_2026-27 03/04/2026->->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim form10IEAEarlierAYOldRegime_pfl, form10IEAAssYear_pfl, form10IEAEarlierAYAckOldRegime_pfl, f10IEAEarlierAYNewRegime_pfl, assYrF10IEANewTaxReg_pfl, form10IEAEarlierAYAckNewRegime_pfl, f10IEACurrAYNewRegime_pfl, f10IEADateCurrAYNewTax_pfl, f10IEAAckNoCurrAYNewTax_pfl, f10IEACurrAYOldRegime_pfl, f10IEADateCurrAYOldTax_pfl, f10IEAAckNoCurrAYOldTax_pfl As Variant

  form10IEAEarlierAYOldRegime_pfl = jsonObject("Form10IEA")("form10IEAEarlierAYOldRegime")
  form10IEAAssYear_pfl = jsonObject("Form10IEA")("form10IEAAssYear")
  form10IEAEarlierAYAckOldRegime_pfl = jsonObject("Form10IEA")("form10IEAEarlierAYAckOldRegime")
  f10IEAEarlierAYNewRegime_pfl = jsonObject("Form10IEA")("f10IEAEarlierAYNewRegime")
  assYrF10IEANewTaxReg_pfl = jsonObject("Form10IEA")("assYrF10IEANewTaxReg")
  form10IEAEarlierAYAckNewRegime_pfl = jsonObject("Form10IEA")("form10IEAEarlierAYAckNewRegime")
  f10IEACurrAYNewRegime_pfl = jsonObject("Form10IEA")("f10IEACurrAYNewRegime")
  f10IEADateCurrAYNewTax_pfl = jsonObject("Form10IEA")("f10IEADateCurrAYNewTax")
  f10IEAAckNoCurrAYNewTax_pfl = jsonObject("Form10IEA")("f10IEAAckNoCurrAYNewTax")
  f10IEACurrAYOldRegime_pfl = jsonObject("Form10IEA")("f10IEACurrAYOldRegime")
  f10IEADateCurrAYOldTax_pfl = jsonObject("Form10IEA")("f10IEADateCurrAYOldTax")
  f10IEAAckNoCurrAYOldTax_pfl = jsonObject("Form10IEA")("f10IEAAckNoCurrAYOldTax")
  
  
  
  If UCase(Left(Sheet1.Range("sheet1.Status").value, 1)) = UCase("F") Then
      Taxregime_25_26_Status
  
  Else
       
       If form10IEAEarlierAYOldRegime_pfl <> "" Then
       
           If UCase(form10IEAEarlierAYOldRegime_pfl) = UCase("Y") Or UCase(form10IEAEarlierAYOldRegime_pfl) = UCase("YES") Then
                If Sheet1.Range("Sheet1.115BAC").Locked = False Then
                  Sheet1.Range("Sheet1.115BAC").value = "Yes"
                  Taxregime_25_26
                End If
                
                If form10IEAAssYear_pfl <> "" And Sheet1.Range("BAC115.AssessmentYear_2425").Locked = False Then
                   Sheet1.Range("BAC115.AssessmentYear_2425").value = form10IEAAssYear_pfl
                End If
                
                If form10IEAEarlierAYAckOldRegime_pfl <> "" And Sheet1.Range("BAC115.AcknowledgementNo").Locked = False Then
                Sheet1.Range("BAC115.AcknowledgementNo").value = form10IEAEarlierAYAckOldRegime_pfl
                End If
                
                If f10IEAEarlierAYNewRegime_pfl <> "" Then
                
                        If UCase(f10IEAEarlierAYNewRegime_pfl) = UCase("Y") Or UCase(f10IEAEarlierAYNewRegime_pfl) = UCase("YES") Then
                                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Locked = False Then
                                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "Yes"
                                Taxregime_25_26_A
                                End If
                                
                                If assYrF10IEANewTaxReg_pfl <> "" And Sheet1.Range("BAC115.AssessmentYear_Aii").Locked = False Then
                                   Sheet1.Range("BAC115.AssessmentYear_Aii").value = assYrF10IEANewTaxReg_pfl
                                End If
                                
                                If form10IEAEarlierAYAckNewRegime_pfl <> "" And Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Locked = False Then
                                   Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = form10IEAEarlierAYAckNewRegime_pfl
                                End If
                                
                        ElseIf UCase(f10IEAEarlierAYNewRegime_pfl) = UCase("N") Or UCase(f10IEAEarlierAYNewRegime_pfl) = UCase("NO") Then
                                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Locked = False Then
                                Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No"
                                Taxregime_25_26_A
                                End If
                                
                                If f10IEACurrAYNewRegime_pfl <> "" Then
                                     If UCase(f10IEACurrAYNewRegime_pfl) = UCase("Y") Or UCase(f10IEACurrAYNewRegime_pfl) = UCase("YES") Then
                                            If Sheet1.Range("BAC115.A23ii_No").Locked = False Then
                                            Sheet1.Range("BAC115.A23ii_No").value = "Yes"
                                            Taxregime_26_27_A23ii_b
                                            End If
                                            
                                            If f10IEADateCurrAYNewTax_pfl <> "" Then
                                               Dim dfmt_f10IEADateCurrAYNewTax_pfl As Variant
                                               dfmt_f10IEADateCurrAYNewTax_pfl = Mid(f10IEADateCurrAYNewTax_pfl, 9, 2) & "/" & Mid(f10IEADateCurrAYNewTax_pfl, 6, 2) & "/" & Mid(f10IEADateCurrAYNewTax_pfl, 1, 4)
                                                    If Sheet1.Range("BAC115.DateOfFiling_New_No").Locked = False Then
                                                    Sheet1.Range("BAC115.DateOfFiling_New_No").value = dfmt_f10IEADateCurrAYNewTax_pfl
                                                    End If
                                            End If
                                            
                                            If f10IEAAckNoCurrAYNewTax_pfl <> "" And Sheet1.Range("BAC115.AcknowledgementNo_New_No").Locked = False Then
                                               Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = f10IEAAckNoCurrAYNewTax_pfl
                                            End If
                                            
                                     ElseIf UCase(f10IEACurrAYNewRegime_pfl) = UCase("N") Or UCase(f10IEACurrAYNewRegime_pfl) = UCase("NO") Then
                                            If Sheet1.Range("BAC115.A23ii_No").Locked = False Then
                                            Sheet1.Range("BAC115.A23ii_No").value = "No"
                                            Taxregime_26_27_A23ii_b
                                            End If
                                     End If

                                End If
                                
                                
                        End If
                
                End If
                
           ElseIf UCase(form10IEAEarlierAYOldRegime_pfl) = UCase("N") Or UCase(form10IEAEarlierAYOldRegime_pfl) = UCase("NO") Then
                If Sheet1.Range("Sheet1.115BAC").Locked = False Then
                  Sheet1.Range("Sheet1.115BAC").value = "No"
                  Taxregime_25_26
                End If
                
                
                If f10IEACurrAYOldRegime_pfl <> "" Then
                
                   If UCase(f10IEACurrAYOldRegime_pfl) = UCase("Y") Or UCase(f10IEACurrAYOldRegime_pfl) = UCase("YES") Then
                            If Sheet1.Range("BAC115.A23b").Locked = False Then
                               Sheet1.Range("BAC115.A23b").value = "Yes"
                               Taxregime_25_26_C
                            End If
                            
                            If f10IEADateCurrAYOldTax_pfl <> "" Then
                               Dim dfmt_f10IEADateCurrAYOldTax_pfl As Variant
                                   dfmt_f10IEADateCurrAYOldTax_pfl = Mid(f10IEADateCurrAYOldTax_pfl, 9, 2) & "/" & Mid(f10IEADateCurrAYOldTax_pfl, 6, 2) & "/" & Mid(f10IEADateCurrAYOldTax_pfl, 1, 4)
                                   If Sheet1.Range("BAC115.DateOfFiling_New_NA").Locked = False Then
                                        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = dfmt_f10IEADateCurrAYOldTax_pfl
                                   End If
                            End If
                            
                            If f10IEAAckNoCurrAYOldTax_pfl <> "" And Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Locked = False Then
                                Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = f10IEAAckNoCurrAYOldTax_pfl
                            End If
                   
                   
                   ElseIf UCase(f10IEACurrAYOldRegime_pfl) = UCase("N") Or UCase(f10IEACurrAYOldRegime_pfl) = UCase("NO") Then
                            If Sheet1.Range("BAC115.A23b").Locked = False Then
                               Sheet1.Range("BAC115.A23b").value = "No"
                               Taxregime_25_26_C
                            End If
                   
                   End If
                   
                End If
                
           End If
           
       End If

  End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<-<-*//





'added by Chetan C M  AY 2025-26 SIT-86604
'--start--
        
    clauseiv7provisio139i = jsonObject("filingStatus")("clauseiv7provisio139i")
        
    If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
            If clauseiv7provisio139i = "Y" Then
                Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes"
            ElseIf clauseiv7provisio139i = "N" Then
                Sheet1.Range("clauseiv7provisio139iFlg").value = "No"
            Else
                Sheet1.Range("clauseiv7provisio139iFlg").value = "(Select)"
            End If
     End If
         
    Dim Nodecounttt As Variant
          Nodecounttt = 0
 
    For Each node In jsonObject("filingStatus")("clauseiv7provisio139iDtls")
    Nodecounttt = Nodecounttt + 1
    'AY_2024_25 OLD
    'If Sheet1.Range("clauseiv7provisio139iNature") = 1 Then
    If Nodecounttt = 1 Then
        If UCase(node("clauseiv7provisio139iNature")) = UCase("Yes") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_1").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_1").value = "Yes"
                If Sheet1.Range("clauseiv7provisio139iAmount_1").Locked = False Then
                 'Sheet1.Range("clauseiv7provisio139iAmount_1").value = clauseiv7provisio139iAmount
                Sheet1.Range("clauseiv7provisio139iAmount_1").value = node("clauseiv7provisio139iAmount")
                End If
                End If
        ElseIf UCase(node("clauseiv7provisio139iNature")) = UCase("NO") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_1").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_1").value = "No"
                End If
        End If
        'AY_2024_25 OLD
    'ElseIf Sheet1.Range("clauseiv7provisio139iNature") = 2 Then
     ElseIf Nodecounttt = 2 Then
        If UCase(node("clauseiv7provisio139iNature")) = UCase("Yes") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_2").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_2").value = "Yes"
                If Sheet1.Range("clauseiv7provisio139iAmount_2").Locked = False Then
                 'Sheet1.Range("clauseiv7provisio139iAmount_2").value = clauseiv7provisio139iAmount
                Sheet1.Range("clauseiv7provisio139iAmount_2").value = node("clauseiv7provisio139iAmount")
                End If
                End If
        ElseIf UCase(node("clauseiv7provisio139iNature")) = UCase("NO") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_2").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_2").value = "No"
                End If
        End If
        'AY_2024_25 OLD
    'ElseIf Sheet1.Range("clauseiv7provisio139iNature") = 3 Then
    ElseIf Nodecounttt = 3 Then
        If UCase(node("clauseiv7provisio139iNature")) = UCase("Yes") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_3").value = "Yes"
                If Sheet1.Range("clauseiv7provisio139iAmount_3").Locked = False Then
                 'Sheet1.Range("clauseiv7provisio139iAmount_3").value = clauseiv7provisio139iAmount
                Sheet1.Range("clauseiv7provisio139iAmount_3").value = node("clauseiv7provisio139iAmount")
                End If
                End If
        ElseIf UCase(node("clauseiv7provisio139iNature")) = UCase("NO") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_3").value = "No"
                End If
        End If
        'AY_2024_25 OLD
   ' ElseIf Sheet1.Range("clauseiv7provisio139iNature") = 4 Then
   ElseIf Nodecounttt = 4 Then
        If UCase(node("clauseiv7provisio139iNature")) = UCase("Yes") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_4").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes"
                If Sheet1.Range("clauseiv7provisio139iAmount_4").Locked = False Then
                 'Sheet1.Range("clauseiv7provisio139iAmount_4").value = clauseiv7provisio139iAmount
                Sheet1.Range("clauseiv7provisio139iAmount_4").value = node("clauseiv7provisio139iAmount")
                End If
                End If
        ElseIf UCase(node("clauseiv7provisio139iNature")) = UCase("NO") Then
                If Sheet1.Range("clauseiv7provisio139iFlg_4").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_4").value = "No"
                End If
        End If

End If
    Next
    
    
    
'-- Change end--
        
        
        
        
        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
        
        If ReturnFileSec = "17" Then
            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
            End If
            
            DateofOriginalfile = OrigRetFiledDate
            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
            'issue 4 solved
                YYYY = Mid(DateofOriginalfile, 1, 4)
                MM = Mid(DateofOriginalfile, 6, 2)
                DD = Mid(DateofOriginalfile, 9, 2)
                DateofOriginalfile = DD & "/" & MM & "/" & YYYY
                Sheet1.Range("sheet1.OrigRetFiledDate").value = DateofOriginalfile
            'issue 4 solved
            End If
        Else
            Sheet1.Range("Sheet1.ReceiptNo").value = ReceiptNo
            
            DateofOriginalfile = OrigRetFiledDate
            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
            'issue 5 solved
                YYYY = Mid(DateofOriginalfile, 1, 4)
                MM = Mid(DateofOriginalfile, 6, 2)
                DD = Mid(DateofOriginalfile, 9, 2)
                DateofOriginalfile = DD & "/" & MM & "/" & YYYY
                Sheet1.Range("sheet1.OrigRetFiledDate").value = DateofOriginalfile
            'issue 5 solved
            End If
            
            'issue 6 solved
            If uniqueNo <> "" And Sheet1.Range("sheet1.NoticeNo").Locked = False Then
                Sheet1.Range("sheet1.NoticeNo").value = uniqueNo
            End If
            'issue 6 solved
            
            'issue 7 solved
            NoticeDateussec = NoticeDateUnderSec
            If NoticeDateussec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
                YYYY = Mid(NoticeDateussec, 1, 4)
                MM = Mid(NoticeDateussec, 6, 2)
                DD = Mid(NoticeDateussec, 9, 2)
                NoticeDateussec = DD & "/" & MM & "/" & YYYY
                Sheet1.Range("sheet1.NoticeDate").value = NoticeDateussec
            End If
            'issue 7 solved
        End If
        
        'issue 8 solved
'        If AcknNumber <> "" And Sheet1.Range("Sheet1.acknowledgementNumber").Locked = False Then
'            Sheet1.Range("Sheet1.acknowledgementNumber").value = AcknNumber
'        End If
        'issue 8 solved
        
        'issue 9 solved
'        If DateOfFiling <> "" And Sheet1.Range("Sheet1.date10IE").Locked = False Then
'            YYYY = Mid(DateOfFiling, 1, 4)
'            MM = Mid(DateOfFiling, 6, 2)
'            DD = Mid(DateOfFiling, 9, 2)
'            DateOfFiling = DD & "/" & MM & "/" & YYYY
'            Sheet1.Range("Sheet1.date10IE").value = DateOfFiling
'        End If
        'issue 9 solved
        
           
        If RepName <> "" And Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.NameRepAssessee").value = RepName
        End If
        
       'Malli commented PAG_C107to110_AY_2026-27
'        If RepCapacity = "L" Then
'             RepCapacity = "Legal Heir"
'        ElseIf RepCapacity = "M" Then
'             RepCapacity = "Manager"
'        ElseIf RepCapacity = "G" Then
'             RepCapacity = "Guardian"
'        ElseIf RepCapacity = "O" Then
'             RepCapacity = "Other"
'        End If
'        If RepCapacity <> "" And Sheet1.Range("sheet1.CapacityRep").Locked = False Then
'            Sheet1.Range("sheet1.CapacityRep").value = RepCapacity
'        End If
'
'
'        If RepAddress <> "" And Sheet1.Range("sheet1.AddressRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.AddressRepAssessee").value = RepAddress
'        End If
'
'        If RepPAN <> "" And Sheet1.Range("sheet1.PANRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.PANRepAssessee").value = RepPAN
'        End If

'        If RepAadhaar <> "" And Sheet1.Range("sheet1.AadhaarRepAssessee").Locked = False Then
'            Sheet1.Range("sheet1.AadhaarRepAssessee").value = RepAadhaar
'        End If
'        ''------------------------------
         'Malli PAG_C111&112_AY_2026-27
        If repEmailID <> "" And Sheet1.Range("sheet1.EmailRepAssessee").Locked = False Then
          Sheet1.Range("sheet1.EmailRepAssessee").value = repEmailID
        End If
        
        '28/01/2026
        Dim countryCodeRepMobileNo_pfl As Variant
        countryCodeRepMobileNo_pfl = jsonObject("assesseeRep")("countryCodeRepMobileNo")
        
        If countryCodeRepMobileNo_pfl <> "" And Sheet1.Range("sheet1.CountryCodeRepAssessee").Locked = False Then
          Sheet1.Range("sheet1.CountryCodeRepAssessee").value = countryCodeRepMobileNo_pfl
        End If
        '-----------------------------
        
        If repMobileNo <> "" And Sheet1.Range("sheet1.MobileRepAssessee").Locked = False Then
          Sheet1.Range("sheet1.MobileRepAssessee").value = repMobileNo
        End If
        '-------------------
End Function
 
'Malli_AY_2026-27_22/01/2026
Function ImportscheduleEA10_13A_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary   As Object
Dim EA10_13A_placeofwork, EA10_13A_placeofwork_pfl As Variant
Dim EA10_13A_actlHRARecv, EA10_13A_actlHRARecv_pfl As Variant

Set jsonObject = ParseJson(jsonText)

'Place of Residence
'lastFiledITR.scheduleEA10_13A.placeofwork
EA10_13A_placeofwork = jsonObject("lastFiledITR")("scheduleEA10_13A")("placeofwork")

If EA10_13A_placeofwork <> "" And Sheet15.Range("Sch10of13A_PlaceofWrk").Locked = False Then
        If EA10_13A_placeofwork = "1" Then
               EA10_13A_placeofwork_pfl = "1. Metro"
        ElseIf EA10_13A_placeofwork = "2" Then
               EA10_13A_placeofwork_pfl = "2. Non-Metro"
        Else
               EA10_13A_placeofwork_pfl = "(Select)"
        End If
    Sheet15.Range("Sch10of13A_PlaceofWrk").value = EA10_13A_placeofwork_pfl
End If

'Actual HRA received (A)
'form24q.ScheduleEA10_13A.actlHRARecv
 EA10_13A_actlHRARecv_pfl = jsonObject("form24q")("scheduleEA10_13A")("actlHRARecv")
If EA10_13A_actlHRARecv_pfl <> "" And Sheet15.Range("Sch10of13A_PlaceofWrk").Locked = False Then
Sheet15.Range("Sch10of13A_ActlHRArecivedA").value = EA10_13A_actlHRARecv_pfl
End If

End Function


'Malli_AY_2026-27_22/01/2026
'lastFiledITR.schedule80E.schedule80EDtls

Function ImportSchedule_80E_Pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80E As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80E, IFSC_80E, BankName_80E, PANof_80E, LoanAcctNum_80E, LoanDate_80E, TotalLoanAmt_80E, LoanOutStanding_80E, Intrst_80E As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80E, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String
 
 
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80E = jsonObject("lastFiledITR")
 
If NodeList80E.exists("schedule80E") Then
 
Set Nodelist = jsonObject("lastFiledITR")("schedule80E")("schedule80EDtls")

    LoanfrmBankOrInstitute_80E = Sheet20.Range("LoanfrmBankOrInstitute.80E").Column
    BankName_80E = Sheet20.Range("bankName.80E").Column
    LoanAcctNum_80E = Sheet20.Range("loanAccNum.80E").Column
    LoanDate_80E = Sheet20.Range("loanDate.80E").Column
    TotalLoanAmt_80E = Sheet20.Range("loanAmt.80E").Column
     
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80E").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80E").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80E").ClearContents
        End If
        
        If Sheet20.Range("bankName.80E").Locked = False Then
            Sheet20.Range("bankName.80E").ClearContents
        End If
        
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
            Dim LoanTknFrom_80E, LoanTknFrom_80E_pfl
            
            LoanTknFrom_80E_pfl = node("loanTknFrom")
            If UCase(LoanTknFrom_80E_pfl) = UCase("B") Then
                LoanTknFrom_80E = "Bank"
            ElseIf UCase(LoanTknFrom_80E_pfl) = UCase("I") Then
                LoanTknFrom_80E = "Institution"
            Else
                LoanTknFrom_80E = "(Select)"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80E).value = LoanTknFrom_80E
        End If
        
        
        If Sheet20.Cells(rowcount, BankName_80E).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80E).value = node("bankOrInstnName")
        End If
         
        If Sheet20.Cells(rowcount, LoanAcctNum_80E).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80E).value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
            strDate = node("dateofLoan")
            If strDate <> "" Then
               strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4)
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80E).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80E).value = strDate
            strDate = ""
        End If
        
        If Sheet20.Cells(rowcount, TotalLoanAmt_80E).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80E).value = node("totalLoanAmt")
        End If
        
         
        
        cnt = cnt + 1
    Next node
    Rec80E = cnt
 End If
End Function

'Malli_AY_2026-27_22/01/2026
'lastFiledITR.Schedule80EE.Schedule80EEDtls
Function ImportSchedule_80EE_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EE As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EE, IFSC_80EE, BankName_80EE, PANof_80EE, LoanAcctNum_80EE, LoanDate_80EE, TotalLoanAmt_80EE, LoanOutStanding_80EE, Intrst_80EE As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EE, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

 
 
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EE = jsonObject("lastFiledITR")
 
If NodeList80EE.exists("schedule80EE") Then
 
Set Nodelist = jsonObject("lastFiledITR")("schedule80EE")("schedule80EEDtls")

    LoanfrmBankOrInstitute_80EE = Sheet20.Range("LoanfrmBankOrInstitute.80EE").Column
    BankName_80EE = Sheet20.Range("bankName.80EE").Column
    LoanAcctNum_80EE = Sheet20.Range("loanAccNum.80EE").Column
    LoanDate_80EE = Sheet20.Range("loanDate.80EE").Column
    TotalLoanAmt_80EE = Sheet20.Range("loanAmt.80EE").Column
     
    
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
            Dim LoanTknFrom_80EE, LoanTknFrom_80EE_pfl
            
            LoanTknFrom_80EE_pfl = node("loanTknFrom")
            If UCase(LoanTknFrom_80EE_pfl) = UCase("B") Then
                LoanTknFrom_80EE = "Bank"
            ElseIf UCase(LoanTknFrom_80EE_pfl) = UCase("I") Then
                LoanTknFrom_80EE = "Institution"
            Else
                LoanTknFrom_80EE = "(Select)"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EE).value = LoanTknFrom_80EE
        End If
         
        If Sheet20.Cells(rowcount, BankName_80EE).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80EE).value = node("bankOrInstnName")
        End If
         
        If Sheet20.Cells(rowcount, LoanAcctNum_80EE).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80EE).value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("dateofLoan")
            If strDate <> "" Then
                strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4)
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80EE).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80EE).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80EE).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80EE).value = node("totalLoanAmt")
        End If
        
         
        cnt = cnt + 1
    Next node
    Rec80EE = cnt
End If
End Function


'Malli_AY_2026-27_22/01/2026
'lastFiledITR.schedule80EEA.schedule80EEADtls

Function ImportSchedule_80EEA_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EEA As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEA, IFSC_80EEA, BankName_80EEA, PANof_80EEA, LoanAcctNum_80EEA, LoanDate_80EEA, TotalLoanAmt_80EEA, LoanOutStanding_80EEA, Intrst_80EEA As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEA, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)
Set NodeList80EEA = jsonObject("lastFiledITR")
 
If NodeList80EEA.exists("schedule80EEA") Then
  
        
Dim PropStmpDtyVal_80EEA
PropStmpDtyVal_80EEA = jsonObject("lastFiledITR")("schedule80EEA")("propStmpDtyVal")

If PropStmpDtyVal_80EEA <> "" Then
    If Sheet20.Range("Stampduty.80EEA").Locked = False Then
            Sheet20.Range("Stampduty.80EEA").value = PropStmpDtyVal_80EEA
        End If
End If



Set Nodelist = jsonObject("lastFiledITR")("schedule80EEA")("schedule80EEADtls")

    LoanfrmBankOrInstitute_80EEA = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Column
    BankName_80EEA = Sheet20.Range("bankName.80EEA").Column
    LoanAcctNum_80EEA = Sheet20.Range("loanAccNum.80EEA").Column
    LoanDate_80EEA = Sheet20.Range("loanDate.80EEA").Column
    TotalLoanAmt_80EEA = Sheet20.Range("loanAmt.80EEA").Column
     
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
        End If
        
        If Sheet20.Range("bankName.80EEA").Locked = False Then
            Sheet20.Range("bankName.80EEA").ClearContents
        End If
        
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
            Dim LoanTknFrom_80EEA, LoanTknFrom_80EEA_pfl
            
            LoanTknFrom_80EEA_pfl = node("loanTknFrom")
            If UCase(LoanTknFrom_80EEA_pfl) = UCase("B") Then
                LoanTknFrom_80EEA = "Bank"
            ElseIf UCase(LoanTknFrom_80EEA_pfl) = UCase("I") Then
                LoanTknFrom_80EEA = "Institution"
            Else
                LoanTknFrom_80EEA = "(Select)"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).value = LoanTknFrom_80EEA
        End If
        
        If Sheet20.Cells(rowcount, BankName_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80EEA).value = node("bankOrInstnName")
        End If
         
        If Sheet20.Cells(rowcount, LoanAcctNum_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80EEA).value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("dateofLoan")
            If strDate <> "" Then
                strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4)
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80EEA).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80EEA).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80EEA).value = node("totalLoanAmt")
        End If
         
        cnt = cnt + 1
    Next node
    Rec80EEA = cnt
 End If
End Function

'Malli_AY_2026-27_22/01/2026
'lastFiledITR.schedule80EEB.schedule80EEBDtls

Function ImportSchedule_80EEB_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EEB As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEB, IFSC_80EEB, BankName_80EEB, Vehicle_value_80EEB, VehicleRegNum_80EEB, PANof_80EEB, LoanAcctNum_80EEB, LoanDate_80EEB, TotalLoanAmt_80EEB, LoanOutStanding_80EEB, Intrst_80EEB As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEB, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

 
 
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EEB = jsonObject("lastFiledITR")
 
If NodeList80EEB.exists("schedule80EEB") Then
  
Set Nodelist = jsonObject("lastFiledITR")("schedule80EEB")("schedule80EEBDtls")

    LoanfrmBankOrInstitute_80EEB = Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Column
    BankName_80EEB = Sheet20.Range("bankName.80EEB").Column
    LoanAcctNum_80EEB = Sheet20.Range("loanAccNum.80EEB").Column
    LoanDate_80EEB = Sheet20.Range("loanDate.80EEB").Column
    TotalLoanAmt_80EEB = Sheet20.Range("loanAmt.80EEB").Column
    VehicleRegNum_80EEB = Sheet20.Range("VehicleRegNum.80EEB").Column
     
    
    
    TotalExRow = Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet20.Range("LoanfrmBankOrInstitute.80EEB").Locked = False Then
            Sheet20.Range("LoanfrmBankOrInstitute.80EEB").ClearContents
        End If
         
        If Sheet20.Range("bankName.80EEB").Locked = False Then
            Sheet20.Range("bankName.80EEB").ClearContents
        End If
        
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
            Dim LoanTknFrom_80EEB, LoanTknFrom_80EEB_pfl
            LoanTknFrom_80EEB_pfl = node("loanTknFrom")
            
            If UCase(LoanTknFrom_80EEB_pfl) = UCase("B") Then
                LoanTknFrom_80EEB = "Bank"
            ElseIf UCase(LoanTknFrom_80EEB_pfl) = UCase("I") Then
                LoanTknFrom_80EEB = "Institution"
            Else
                LoanTknFrom_80EEB = "(Select)"
            End If
            Sheet20.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).value = LoanTknFrom_80EEB
        End If
          
        If Sheet20.Cells(rowcount, BankName_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, BankName_80EEB).value = node("bankOrInstnName")
        End If
         
        If Sheet20.Cells(rowcount, LoanAcctNum_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, LoanAcctNum_80EEB).value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("dateofLoan")
            If strDate <> "" Then
                strDate = Trim(Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4))
            End If
        
        If Sheet20.Cells(rowcount, LoanDate_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, LoanDate_80EEB).value = strDate
            strDate = ""
        End If
        If Sheet20.Cells(rowcount, TotalLoanAmt_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, TotalLoanAmt_80EEB).value = node("totalLoanAmt")
        End If
         
        If Sheet20.Cells(rowcount, VehicleRegNum_80EEB).Locked = False Then
            Sheet20.Cells(rowcount, VehicleRegNum_80EEB).value = node("vehicleRegNo")
        End If
         
        cnt = cnt + 1
    Next node
    Rec80EEB = cnt
End If
End Function


'Malli_AY_2026-27 20/01/2026
Function ImportScheduleHP_Pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2, Nodelist3, init As Object
    Dim node, Node1, Node2, Node3 As Object
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant
    Dim iTotal24BRow, sTotal24BRow, TotalDiff24BRow As Variant
    Dim rowcount, cnt As Long
    
    
    Dim Node_1, Node_1M, Nodelist_2, Node_2, Node_1MM, Node_3, Nodelist_3 As Object
    Dim panOfTenant_count As Integer
 
 
 '30/01/2026
 Dim ais_Tenchk_pfl As Boolean
 ais_Tenchk_pfl = False
 '--------------------
 
Set jsonObject = ParseJson(jsonText)

'Primary source

Dim init_1 As Object
panOfTenant_count = 0

If jsonObject("form26as") <> Null Then
Set init_1 = jsonObject("form26as")("scheduleHP")

If init_1 <> Empty Then


Set Node_1M = init_1("propertyDetails")
     panOfTenant_count = 0
     For Each Node_1 In Node_1M
     sTotalCoRow = 0
  
       Set Node_1MM = Node_1("tenantDetails")
          For Each Node_2 In Node_1MM
            Dim panOfTenant_Nodechk As Variant
            panOfTenant_Nodechk = Node_2("panOfTenant")
                 If panOfTenant_Nodechk <> "" Then
                 panOfTenant_count = panOfTenant_count + 1
                 End If
       Next Node_2
  Next Node_1
End If
End If

If panOfTenant_count > 0 Then

rowcount = 0

'Malli_AY_2026-27  25/03/2026
  'Debug.Print init_1("propertyDetails").count
  If init_1("propertyDetails").count > 2 Then
  MsgBox "* ""Since you have more than two house properties, please proceed with filing ITR-2 or ITR-3.""", vbOKOnly, "Alert:"
  End If
  
For Each Node_1 In Node_1M
       
                Dim typeOfHP_tre As Boolean
                typeOfHP_tre = True
                rowcount = rowcount + 1
                
                
                iTotalTenRow = Sheet22.Range("HP.NameofTenant" & rowcount).Rows.count
                sTotalTenRow = Node_1("tenantDetails").count
                TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)

                        If TotalDiffTenRow > 0 Then
                        Sheet22.Activate
                        AddPropertyTenant (TotalDiffTenRow)
                        End If
               
                cnt = 0
                cnt = getRowNo(Sheet22.Range("HP.NameofTenant" & rowcount).name)
                cnt = cnt - 1
            
            
iLetOut = Node_1("typeOfHP")
                               
        If iLetOut = "Y" Or iLetOut = "L" Then
            sLetOut = "Let Out"
        ElseIf iLetOut = "D" Then
            sLetOut = "Deemed Let Out"
        ElseIf iLetOut = "N" Or iLetOut = "S" Then
            sLetOut = "Self Occupied"
        Else
            sLetOut = "(Select)"
        End If
         
        If Sheet22.Range("HP.ifLetOut" & rowcount).Locked = False Then
            Sheet22.Range("HP.ifLetOut" & rowcount).value = sLetOut
        End If
  
Set Nodelist_2 = Node_1("tenantDetails")

For Each Node_2 In Nodelist_2
              
     If UCase(Node_2("panOfTenant")) <> "" Then
            
            cnt = cnt + 1

                    If Sheet22.Cells(cnt, Sheet22.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node_2("nameOfTenant"))
                    End If
                    If Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node_2("panOfTenant"))
                        ais_Tenchk_pfl = True
                    End If
                    If Sheet22.Cells(cnt, Sheet22.Range("HP.AadharofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node_2("aadhaarofTenant"))
                    End If
 
                    If Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                        Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node_2("panTANofTenant"))
                        ais_Tenchk_pfl = True
                    End If
                    
            
             End If
             
Next Node_2
                
            
                If Sheet22.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                Sheet22.Range("HP.AnnualLetableValue" & rowcount).value = UCase(Node_1("grossRent"))
                End If
                
                
Next Node_1







Else:
        

    If panOfTenant_count <= 0 Then
    'Malli_SIT-109813
    'Secoundary source
    '\\* AY_2026-27 onwards for ITR-1 & ITR-4 HP schedule Type of House property? field has two types of prefill sources under last field ITR which is _
   '1  lastFiledITR.typeOfHP   ( it should prefill only single block)
   '2  lastFiledITR.scheduleHP.PropertyDetails.ifLetOut  ( it should prefill only multiple block)
   '*//
    Dim typeOfHP_1stLFIsorce, typeOfHP_1stLFIsorcepfl As Variant
    
    typeOfHP_1stLFIsorce = jsonObject("lastFiledITR")("typeOfHP")
    
    If typeOfHP_1stLFIsorce <> "" Then
    'LFI Primary source
     
                    If UCase(typeOfHP_1stLFIsorce) = UCase("Y") Or UCase(typeOfHP_1stLFIsorce) = UCase("L") Then
                        typeOfHP_1stLFIsorcepfl = "Let Out"
                    ElseIf UCase(typeOfHP_1stLFIsorce) = UCase("D") Then
                        typeOfHP_1stLFIsorcepfl = "Deemed Let Out"
                    ElseIf UCase(typeOfHP_1stLFIsorce) = UCase("N") Or UCase(typeOfHP_1stLFIsorce) = UCase("S") Then
                        typeOfHP_1stLFIsorcepfl = "Self Occupied"
                    Else
                        typeOfHP_1stLFIsorcepfl = "(Select)"
                    End If
                    
                    
                    If Sheet22.Range("HP.ifLetOut1").Locked = False Then
                    Sheet22.Range("HP.ifLetOut1").value = typeOfHP_1stLFIsorcepfl
                    End If
    
     Else:
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    'LFI Secoundary source
     
         Set init = jsonObject("lastFiledITR")("scheduleHP")
        
        
                    If init <> Empty Then
                    If init.exists("propertyDetails") Then
                    
                        Set Nodelist = jsonObject("lastFiledITR")("scheduleHP")("propertyDetails")
                     
                        rowcount = 0
                        cnt = 0
                         
                    'Malli_AY_2026-27  25/03/2026
                    'Debug.Print Nodelist.count
                    If Nodelist.count > 2 Then
                         MsgBox "* ""Since you have more than two house properties, please proceed with filing ITR-2 or ITR-3.""", vbOKOnly, "Alert:"
                    End If
                    '------------------------------------
                        
                        For Each node In Nodelist
                            rowcount = rowcount + 1
                            
                                If Sheet22.Range("HP.AddrDetail" & rowcount).Locked = False Then
                                    Sheet22.Range("HP.AddrDetail" & rowcount).value = UCase(node("addressDetailWithZipCode")("addrDetail"))
                                End If
                                If Sheet22.Range("HP.CityOrTownOrDistrict" & rowcount).Locked = False Then
                                    Sheet22.Range("HP.CityOrTownOrDistrict" & rowcount).value = UCase(node("addressDetailWithZipCode")("cityOrTownOrDistrict"))
                                End If
                                
                                iState = UCase(node("addressDetailWithZipCode")("stateCode"))
                                sState = Findtext(iState, "State")
                               
                                Sheet22.Range("HP.StateCode" & rowcount).value = sState
                                iCountry = UCase(node("addressDetailWithZipCode")("countryCode"))
                                sCountry = Findtext(iCountry, "Country")
                                
                                If Sheet22.Range("HP.CountryCode" & rowcount).Locked = False Then
                                    Sheet22.Range("HP.CountryCode" & rowcount).value = sCountry
                                End If
                                If Sheet22.Range("HP.PinCode" & rowcount).Locked = False Then
                                    Sheet22.Range("HP.PinCode" & rowcount).value = UCase(node("addressDetailWithZipCode")("pinCode"))
                                End If
                                If Sheet22.Range("HP.ZipCode" & rowcount).Locked = False Then
                                    Sheet22.Range("HP.ZipCode" & rowcount).value = UCase(node("addressDetailWithZipCode")("zipCode"))
                                End If
                                
                                Dim OwnerProperty_HP
                                
                                OwnerProperty_HP = UCase(node("propertyOwner"))
                                
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
                                    Sheet22.Range("HP.OwnerPropertyDescription" & rowcount).value = node("propertyOwnerOther")
                                End If
                                
                                Dim Co_Ownedflag As Variant
                                
                                Co_Ownedflag = UCase(node("propCoOwnedFlg"))
                                If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
                                Co_Ownedflag = "Yes"
                                ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
                                Co_Ownedflag = "No"
                                Else
                                Co_Ownedflag = "(Select)"
                                End If
                                
                                Sheet22.Range("HP.CoOwnedYN" & rowcount).value = Co_Ownedflag
                                 
                                Dim hpshare As Variant
                                hpshare = UCase(node("asseseeShareProperty"))
                                
                                If Sheet22.Range("HP.SharePercent" & rowcount).Locked = False Then
                                    Sheet22.Range("HP.SharePercent" & rowcount).value = hpshare
                                End If
                                
                        
                        
                        If node.exists("coOwners") Then
                        
                        
                            'Co-Owner
                            iTotalCoRow = Sheet22.Range("HP.Co.Name" & rowcount).Rows.count
                             
                            sTotalCoRow = node("coOwners").count
                            
                            TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
                            If TotalDiffCoRow > 0 Then
                                Sheet22.Activate
                                AddPropertyCoOWners (TotalDiffCoRow)
                            End If
                            
                            
                                Set nodeList1 = node("coOwners")
                                cnt = 0
                                cnt = getRowNo(Sheet22.Range("HP.Co.Name" & rowcount).name)
                                cnt = cnt - 1
                                
                                For Each Node1 In nodeList1
                                        cnt = cnt + 1
                                            If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Name" & rowcount).Column).Locked = False Then
                                                Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Name" & rowcount).Column).value = UCase(Node1("nameCoOwner"))
                                            End If
                                            If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Pan" & rowcount).Column).Locked = False Then
                                                Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Pan" & rowcount).Column).value = UCase(Node1("panCoOwner"))
                                            End If
                                            If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Aadhaar" & rowcount).Column).Locked = False Then
                                                Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Aadhaar" & rowcount).Column).value = UCase(Node1("aadhaarCoOwner"))
                                            End If
                                            If Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Share" & rowcount).Column).Locked = False Then
                                                Sheet22.Cells(cnt, Sheet22.Range("HP.Co.Share" & rowcount).Column).value = UCase(Node1("percentShareProperty"))
                                            End If
                                        
                                Next Node1
                        End If
                                
                                iLetOut = UCase(node("ifLetOut"))
                                
                                
                                If iLetOut = "Y" Or iLetOut = "L" Then
                                    sLetOut = "Let Out"
                                ElseIf iLetOut = "D" Then
                                    sLetOut = "Deemed Let Out"
                                ElseIf iLetOut = "N" Or iLetOut = "S" Then
                                    sLetOut = "Self Occupied"
                                Else
                                    sLetOut = "(Select)"
                                End If
                                
                                
                                Sheet22.Range("HP.ifLetOut" & rowcount).value = sLetOut
                        
                    If node.exists("tenantDetails") Then
                    
                                'Tenant
                                iTotalTenRow = Sheet22.Range("HP.NameofTenant" & rowcount).Rows.count
                                sTotalTenRow = node("tenantDetails").count
                                TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
                                
                                If TotalDiffTenRow > 0 Then
                                    Sheet22.Activate
                                    AddPropertyTenant (TotalDiffTenRow)
                                End If
                                
                                Set Nodelist2 = node("tenantDetails")
                                cnt = 0
                                cnt = getRowNo(Sheet22.Range("HP.NameofTenant" & rowcount).name)
                                cnt = cnt - 1
                                For Each Node2 In Nodelist2
                                    cnt = cnt + 1
                                                If Sheet22.Cells(cnt, Sheet22.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node2("nameOfTenant"))
                                                End If
                                                If Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node2("panOfTenant"))
                                                    ais_Tenchk_pfl = True
                                                End If
                                                If Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node2("panTANofTenant"))
                                                    ais_Tenchk_pfl = True
                                                End If
                                                If Sheet22.Cells(cnt, Sheet22.Range("HP.AadharofTenant" & rowcount).Column).Locked = False Then
                                                        If Node2.Exsist("AadhaarofTenant") Then
                                                            Sheet22.Cells(cnt, Sheet22.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node2("AadhaarofTenant"))
                                                        End If
                                                End If
                                Next Node2
                            End If
                                
                          '24B<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                          If node.exists("rentdetails") Then
                          
                            'Schedule24(B)
                            iTotal24BRow = Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Rows.count
                            sTotal24BRow = node("rentdetails")("section24B")("section24BDtls").count
                            TotalDiff24BRow = WorksheetFunction.Max((sTotal24BRow - iTotal24BRow), 0)
                            
                            If TotalDiff24BRow > 0 Then
                                Sheet22.Activate
                                AddSection24b (TotalDiff24BRow)
                            End If
                             
                          
                            Set Nodelist3 = node("rentdetails")("section24B")("section24BDtls")
                                cnt = 0
                                cnt = getRowNo(Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).name)
                                cnt = cnt - 1
                                
                                For Each Node3 In Nodelist3
                                    cnt = cnt + 1
                                    
                                                If Sheet22.Cells(cnt, Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Locked = False Then
                                                    Dim LoanTknFrom_24B, LoanTknFrom_24B_pfl
                                                    
                                                    LoanTknFrom_24B_pfl = Node3("loanTknFrom")
                                                        If UCase(LoanTknFrom_24B_pfl) = UCase("B") Then
                                                                LoanTknFrom_24B = "Bank "
                                                        ElseIf UCase(LoanTknFrom_24B_pfl) = UCase("I") Then
                                                                LoanTknFrom_24B = "Other than Bank"
                                                        Else
                                                                LoanTknFrom_24B = "(Select)"
                                                        End If
                                                    
                                                    Sheet22.Cells(cnt, Sheet22.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).value = LoanTknFrom_24B
                                                End If
                                                
                                                If Sheet22.Cells(cnt, Sheet22.Range("bankName.24b" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("bankName.24b" & rowcount).Column).value = UCase(Node3("bankOrInstnName"))
                                                End If
                                                
                                                If Sheet22.Cells(cnt, Sheet22.Range("loanAccNum.24b" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("loanAccNum.24b" & rowcount).Column).value = UCase(Node3("loanAccNoOfBankOrInstnRefNo"))
                                                End If
                                                
                                                If Sheet22.Cells(cnt, Sheet22.Range("loanDate.24b" & rowcount).Column).Locked = False Then
                                                    Dim str24BDate As Variant
                                                    str24BDate = Node3("dateofLoan")
                                                    If str24BDate <> "" Then
                                                            str24BDate = Mid(str24BDate, 9, 2) & "/" & Mid(str24BDate, 6, 2) & "/" & Mid(str24BDate, 1, 4)
                                                            Sheet22.Cells(cnt, Sheet22.Range("loanDate.24b" & rowcount).Column).value = str24BDate
                                                               str24BDate = ""
                                                    End If
                                                End If
                                                
                                                If Sheet22.Cells(cnt, Sheet22.Range("loanAmt.24b" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("loanAmt.24b" & rowcount).Column).value = UCase(Node3("totalLoanAmt"))
                                                End If
                                                
                                                
                                                 
                                Next Node3
                            End If
                          
                          
                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                          
                          
                          
                     
                     
                     
                                cnt = cnt + 1
                            Next node
                            
                            
                    End If
                    End If
         End If
        End If
End If



'Malli_30/01/2026
Dim init_ais As Object
Dim Node_1ais, Node_ais, Node_2ais, Node_1MMais As Object
Dim iTotalTenRow_ais, sTotalTenRow_ais, TotalDiffTenRow_ais, panOfTenant_counts_ais As Variant

If ais_Tenchk_pfl <> True Then

        If jsonObject("ais") <> Null Then
        Set init_ais = jsonObject("ais")("ScheduleHP")
        
                If init_ais <> Empty Then
                Set Node_ais = init_ais("PropertyDetails")
                     panOfTenant_counts_ais = 0
                     
                  For Each Node_1ais In Node_ais
                     sTotalCoRow = 0
                  
                       Set Node_1MMais = Node_1ais("TenantDetails")
                          For Each Node_2ais In Node_1MMais
                            Dim ais_panOfTenant_Nodechk, ais_PANTANofTenant_Nodechk As Variant
                            
                            ais_panOfTenant_Nodechk = Node_2ais("PANOfTenant")
                            ais_PANTANofTenant_Nodechk = Node_2ais("PANTANofTenant")
                            
                                 If ais_panOfTenant_Nodechk <> "" Or ais_PANTANofTenant_Nodechk <> "" Then
                                 panOfTenant_counts_ais = panOfTenant_counts_ais + 1
                                 End If
                                 
                       Next Node_2ais
                  Next Node_1ais
                  
                End If
                
                
                If panOfTenant_counts_ais > 0 Then
                rowcount = 0
                
                For Each Node_1ais In Node_ais
                rowcount = rowcount + 1
                
                        iTotalTenRow_ais = Sheet22.Range("HP.NameofTenant" & rowcount).Rows.count
                        sTotalTenRow_ais = Node_1ais("TenantDetails").count
                        TotalDiffTenRow_ais = WorksheetFunction.Max((sTotalTenRow_ais - iTotalTenRow_ais), 0)
        
                                If TotalDiffTenRow_ais > 0 Then
                                Sheet22.Activate
                                AddPropertyTenant (TotalDiffTenRow_ais)
                                End If
                       
                        cnt = 0
                        cnt = getRowNo(Sheet22.Range("HP.NameofTenant" & rowcount).name)
                        cnt = cnt - 1
                
                        If Sheet22.Range("HP.ifLetOut" & rowcount).Locked = False Then
                        Sheet22.Range("HP.ifLetOut" & rowcount).value = "Let Out"
                        End If
                
                
                            Set Node_1MMais = Node_1ais("TenantDetails")
                                  For Each Node_2ais In Node_1MMais
                                        cnt = cnt + 1

                                                If Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node_2ais("PANOfTenant"))
                                                End If
                                                
                                                If Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet22.Cells(cnt, Sheet22.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node_2ais("PANTANofTenant"))
                                                End If
 
                                  Next Node_2ais
                
                
                      cnt = cnt + 1
                Next Node_1ais
                
                End If
                
                
        End If
End If
'----------------




End Function



























 





































