Attribute VB_Name = "ImportXML"

Option Explicit
 
Dim dom1, dom As DOMDocument
Dim end_TaxP, rngname_TaxP As Variant
Dim end_TDSoth, rngname_TDSoth As Variant
Dim end_TDS3oth, rngname_TDS3oth As Variant
Dim end_TCS, rngname_TCS As Variant

Dim sPassword As String

Dim rowcount, cnt As Variant
Dim xml As Variant
Public Common_year As String
Sub XMLImport()
On Error GoTo endline
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim rowcount As Long
    Dim cellcount As Long
    Dim rowRange As Range
    Dim cellrange As Range
    
     fmsgboxStatus "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year"
    'Application.EnableEvents = False
    xml = ""
    
     Set dom1 = New DOMDocument
    dom1.Load (OpenXMLFileDialog)
   
    
    Dim XpathOfPI, XpathOfPI1 As String
    
   Dim assesmentYear, assesmentYear1 As String

    
    XpathOfPI = "/ITRETURN:ITR /ITR4FORM:ITR4/ITRForm:Form_ITR4" ' "/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom1.SelectNodes(XpathOfPI)
    
     For Each node In Nodelist
    assesmentYear = node.SelectSingleNode("ITRForm:AssessmentYear").text
    Next node
    
        XpathOfPI1 = "/ns3:ITR/ns2:ITR4/Form_ITR4" ' "/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom1.SelectNodes(XpathOfPI1)


    For Each node In Nodelist
    assesmentYear1 = node.SelectSingleNode("AssessmentYear").text
    Next node
    xml = dom1.xml
    xml = Replace(xml, "ITRForm:", "")
'
   
    Set dom = New DOMDocument
    dom.LoadXML (xml)
     If assesmentYear <> "" Then
    Common_year = "/ITRETURN:ITR/ITR4FORM:ITR4"
    ElseIf assesmentYear1 <> "" Then
    Common_year = "/ns3:ITR/ns2:ITR4"
    Else
    Common_year = "/ns4:ITR/ns2:ITR4"
    End If
    
    assesmentYear = ""
    
        
            If Not (ValidateXML2() Or ValidateXML1()) Then
               fmsgboxStatus "* Invalid XML.Please Retry."
               
           Else
          
            PersonalInfoXMLImport2
            ExemptXMLImport2
            ExemptXMLImport3
            TaxExmpIntIncDtlsXMLImport
'            TaxCompXMLImport2
            verpanXMLImport2
            FilingInfoXMLImport2
            SalaryXMLImport2
            TaxCompXMLImport2
            Sch80DXMLImport
             XMLImport_80G_A
            XMLImport_80G_B
            XMLImport_80G_C
            XMLImport_80G_D
            
            TDSonSalaryXMLImport2
            ITXMLImport2
            TDSOthXMLImport2
            TDSOthXMLImport3
            TCSXMLImport2
            VeriInfoXMLImport2
            RefundInfoXMLImport2
            'NonResidentRefundXMLImport
            AeInfoXMLImport
            GstXMLImport
            'NOBXMLImport2
            BPXMLImport2
            ADXMLImport
           ADAXMLImport
           AEXMLImport
           GoodsTableXMLImport
            IncomeInfoXMLImport2
            ImmovableXMLImport2
            ALInfoXMLImport3
            ALInfoXMLImport4
            ALveriXMLImport2
'            XMLImport_80G_A
'            XMLImport_80G_B
'            XMLImport_80G_C
'            XMLImport_80G_D
            Sheet1.Activate
             fmsgboxStatus "* Filled up TDS on Salary, Others , Tax Payments and Personal Information as per the XML File used for Pre Filling." & vbCr & vbLf & "Please Verify the Data Filled"
        End If
        

endline:
    Application.EnableEvents = True
    
End Sub
Function ValidateXML()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim assesmentYear5 As String
   
    'Dim GenderUpload As String
    ValidateXML = True
    
    XpathOfPI = Common_year & "/Form_ITR4" ' "/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

   
     
    For Each node In Nodelist
    assesmentYear5 = node.SelectSingleNode("AssessmentYear").text
        If assesmentYear5 = "2022" Then
         Exit Function
        End If
    Next node
    
    
     ValidateXML = False
End Function
Function ValidateXML1()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim assesmentYear4 As String
   
    'Dim GenderUpload As String
    ValidateXML1 = True
    
    XpathOfPI = Common_year & "/Form_ITR4S" ' "/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

   
     
    For Each node In Nodelist
    assesmentYear4 = node.SelectSingleNode("AssessmentYear").text
        If assesmentYear4 = "2022" Then
         Exit Function
        End If
    Next node
    
    
     ValidateXML1 = False
End Function

Function ValidateXML2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim assesmentYear3 As String
   
    'Dim GenderUpload As String
    ValidateXML2 = True
    
    XpathOfPI = Common_year & "/Form_ITR4" ' "/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

   
     
    For Each node In Nodelist
    assesmentYear3 = node.SelectSingleNode("AssessmentYear").text
        If assesmentYear3 = "2022" Then
         Exit Function
        End If
    Next node
    
    
     ValidateXML2 = False
End Function

Function Findtext(myinput As Variant, tblrange As Variant) As Variant
On Error Resume Next
    Dim Rng As Range
    Dim found As Boolean
    Dim searchtext As Variant
    found = False
    searchtext = ""
    
    For Each Rng In Range(tblrange)
        If InStr(Rng.value, "-") > 0 Then
            If Mid(Rng.value, 1, InStr(Rng.value, "-") - 1) = myinput Then
                found = True
                searchtext = Rng.value
                Exit For
            End If
        End If
    Next Rng

    If found Then
         Findtext = searchtext
    Else
        Findtext = myinput
    End If
    End Function
Sub PersonalInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sCountry, iCountry As Variant
    Dim sState, iState As Variant
    Dim sEmpCat, iEmpCat As Variant
    
    XpathOfPI = "/ITR/ns2:ITR4/PersonalInfo" '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes
            Sheet1.Range("sheet1.FirstName").value = UCase(node.SelectSingleNode("AssesseeName/FirstName").text)
            Sheet1.Range("sheet1.MiddleName").value = UCase(node.SelectSingleNode("AssesseeName/MiddleName").text)
            Sheet1.Range("sheet1.SurNameOrOrgName").value = UCase(node.SelectSingleNode("AssesseeName/SurNameOrOrgName").text)
            Sheet1.Range("sheet1.PAN").value = UCase(node.SelectSingleNode("PAN").text)
            Sheet1.Range("sheet1.ResidenceNo").value = UCase(node.SelectSingleNode("Address/ResidenceNo").text)
            Sheet1.Range("sheet1.ResidenceName").value = UCase(node.SelectSingleNode("Address/ResidenceName").text)
            Sheet1.Range("sheet1.RoadOrStreet").value = UCase(node.SelectSingleNode("Address/RoadOrStreet").text)
            Sheet1.Range("sheet1.LocalityOrArea").value = UCase(node.SelectSingleNode("Address/LocalityOrArea").text)
            
            Sheet1.Range("sheet1.CityOrTownOrDistrict").value = UCase(node.SelectSingleNode("Address/CityOrTownOrDistrict").text)
            
            
            iState = UCase(node.SelectSingleNode("Address/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet1.Range("sheet1.StateCode1").value = sState
        
            Application.EnableEvents = False
            
            iCountry = UCase(node.SelectSingleNode("Address/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet1.Range("sheet1.CountryCode").value = sCountry
            
            Application.EnableEvents = True
            
    
            Sheet1.Range("sheet1.PinCode").value = UCase(node.SelectSingleNode("Address/PinCode").text)
            Sheet1.Range("sheet1.ZipCode").value = UCase(node.SelectSingleNode("Address/ZipCode").text)
            Sheet1.Range("sheet1.STDcode").value = UCase(node.SelectSingleNode("Address/Phone/STDcode").text)
            Sheet1.Range("sheet1.PhoneNo").value = UCase(node.SelectSingleNode("Address/Phone/PhoneNo").text)
            Sheet1.Range("sheet1.mobileCountryCode").value = UCase(node.SelectSingleNode("Address/mobileCountryCode").text)
            Sheet1.Range("sheet1.Mobileno").value = UCase(node.SelectSingleNode("Address/MobileNo").text)
            Sheet1.Range("sheet1.mobileCountryCode2").value = UCase(node.SelectSingleNode("Address/mobileCountryCode2").text)
            Sheet1.Range("sheet1.MobileNoSec").value = UCase(node.SelectSingleNode("Address/MobileNoSec").text)
            Sheet1.Range("sheet1.EmailAddress").value = UCase(node.SelectSingleNode("Address/EmailAddress").text)
                       
            dob = node.SelectSingleNode("DOB").text
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            dob = DD & "/" & MM & "/" & YYYY
            
            Sheet1.Range("sheet1.DOB").value = dob
            
            'Sheet1.Range("sheet1.EmployerCategory1").value = Node.SelectSingleNode("EmployerCategory").Text

            
            Status1 = UCase(node.SelectSingleNode("Status").text)
            If Status1 = "I" Then
                Status1 = "I - INDIVIDUAL"
            ElseIf Status1 = "H" Then
                Status1 = "H - HUF"
            ElseIf Status1 = "F" Then
                Status1 = "F - FIRM(Other than LLP)"
            End If
            Sheet1.Range("sheet1.Status").value = Status1
            
            iEmpCat = UCase(node.SelectSingleNode("EmployerCategory").text)
            
              
            If iEmpCat = "GOV" Then
            sEmpCat = "Central Government"
            ElseIf iEmpCat = "SGOV" Then
            sEmpCat = "State Government"
            ElseIf iEmpCat = "PSU" Then
            sEmpCat = "Public Sector Undertaking"
            ElseIf iEmpCat = "OTH" Then
            sEmpCat = "Others"
            ElseIf iEmpCat = "PE" Then
            sEmpCat = "Pensioners"
            ElseIf sEmpCat = "NA" Then
            sEmpCat = "Not Applicable (eg. Family pension etc)"
            End If
            
            
            'sEmpCat = Findtext(sEmpCat, "EmpCatList")
            Sheet1.Range("sheet1.NatureofEmployment").value = sEmpCat
            
            
            
'            GenderUpload = UCase(Node.SelectSingleNode("Gender").Text)
'
'               If GenderUpload = "M" Then
'               GenderUpload = "MALE"
'               ElseIf GenderUpload = "F" Then
'               GenderUpload = "FEMALE"
'               ElseIf GenderUpload = "T" Then
'               GenderUpload = "Transgender"
'               ElseIf GenderUpload = "X" Then
'                If UCase(Node.SelectSingleNode("Status").Text) = "H" Then
'                    GenderUpload = "Not applicable"
'                Else
'                    GenderUpload = "Not applicable"
'                End If
'            End If
'            Sheet1.Range("sheet1.Gender1").value = GenderUpload
            
            Sheet1.Range("sheet1.Aadhaar").value = UCase(node.SelectSingleNode("AadhaarCardNo").text)

        Next InnerNode
    Next node
End Sub
Sub PersonalInfoXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sCountry, iCountry As Variant
    Dim sState, iState As Variant
    Dim sEmpCat, iEmpCat As Variant
    
    XpathOfPI = Common_year & "/PersonalInfo" '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
       ' For Each InnerNode In Node.ChildNodes
            Sheet1.Range("sheet1.FirstName").value = UCase(node.SelectSingleNode("AssesseeName/FirstName").text)
            Sheet1.Range("sheet1.MiddleName").value = UCase(node.SelectSingleNode("AssesseeName/MiddleName").text)
            Sheet1.Range("sheet1.SurNameOrOrgName").value = UCase(node.SelectSingleNode("AssesseeName/SurNameOrOrgName").text)
            Sheet1.Range("sheet1.PAN").value = UCase(node.SelectSingleNode("PAN").text)
            Sheet1.Range("sheet1.ResidenceNo").value = UCase(node.SelectSingleNode("Address/ResidenceNo").text)
            Sheet1.Range("sheet1.ResidenceName").value = UCase(node.SelectSingleNode("Address/ResidenceName").text)
            Sheet1.Range("sheet1.RoadOrStreet").value = UCase(node.SelectSingleNode("Address/RoadOrStreet").text)
            Sheet1.Range("sheet1.LocalityOrArea").value = UCase(node.SelectSingleNode("Address/LocalityOrArea").text)
            Sheet1.Range("sheet1.CityOrTownOrDistrict").value = UCase(node.SelectSingleNode("Address/CityOrTownOrDistrict").text)
            
            iState = UCase(node.SelectSingleNode("Address/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet1.Range("sheet1.StateCode1").value = sState
        
            Application.EnableEvents = False
            
            iCountry = UCase(node.SelectSingleNode("Address/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet1.Range("sheet1.CountryCode").value = sCountry
            
            Application.EnableEvents = True
            
    
            Sheet1.Range("sheet1.PinCode").value = UCase(node.SelectSingleNode("Address/PinCode").text)
            Sheet1.Range("sheet1.ZipCode").value = UCase(node.SelectSingleNode("Address/ZipCode").text)
            Sheet1.Range("sheet1.STDcode").value = UCase(node.SelectSingleNode("Address/Phone/STDcode").text)
            Sheet1.Range("sheet1.PhoneNo").value = UCase(node.SelectSingleNode("Address/Phone/PhoneNo").text)
            Sheet1.Range("sheet1.mobileCountryCode").value = UCase(node.SelectSingleNode("Address/mobileCountryCode").text)
            Sheet1.Range("sheet1.Mobileno").value = UCase(node.SelectSingleNode("Address/MobileNo").text)
            Sheet1.Range("sheet1.mobileCountryCode2").value = UCase(node.SelectSingleNode("Address/CountryCodeMobileNoSec").text)
            Sheet1.Range("sheet1.MobileNoSec").value = UCase(node.SelectSingleNode("Address/MobileNoSec").text)
            Sheet1.Range("sheet1.EmailAddress").value = UCase(node.SelectSingleNode("Address/EmailAddress").text)
                       
            dob = node.SelectSingleNode("DOB").text
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            dob = DD & "/" & MM & "/" & YYYY
            
            Sheet1.Range("sheet1.DOB").value = dob
            
            'Sheet1.Range("sheet1.EmployerCategory1").value = Node.SelectSingleNode("EmployerCategory").Text

            
            Status1 = UCase(node.SelectSingleNode("Status").text)
            If Status1 = "I" Then
                Status1 = "I - INDIVIDUAL"
            ElseIf Status1 = "H" Then
                Status1 = "H - HUF"
            ElseIf Status1 = "F" Then
                Status1 = "F - FIRM(Other than LLP)"
            End If
            Sheet1.Range("sheet1.Status").value = Status1
            
            
            iEmpCat = UCase(node.SelectSingleNode("EmployerCategory").text)
            
              
            If iEmpCat = "CGOV" Then
            sEmpCat = "Central Government"
            ElseIf iEmpCat = "SGOV" Then
            sEmpCat = "State Government"
            ElseIf iEmpCat = "PSU" Then
            sEmpCat = "Public Sector Undertaking"
            ElseIf iEmpCat = "OTH" Then
            sEmpCat = "Others"
            ElseIf iEmpCat = "PE" Then
            sEmpCat = "Pensioners"
            ElseIf sEmpCat = "NA" Then
            sEmpCat = "Not Applicable (eg. Family pension etc)"
            End If
            
            
            'sEmpCat = Findtext(sEmpCat, "EmpCatList")
            Sheet1.Range("sheet1.NatureofEmployment").value = sEmpCat
            
            
            
            
'            GenderUpload = UCase(Node.SelectSingleNode("Gender").Text)
'
'               If GenderUpload = "M" Then
'               GenderUpload = "MALE"
'               ElseIf GenderUpload = "F" Then
'               GenderUpload = "FEMALE"
'               ElseIf GenderUpload = "T" Then
'               GenderUpload = "Transgender"
'               ElseIf GenderUpload = "X" Then
'                If UCase(Node.SelectSingleNode("Status").Text) = "H" Then
'                    GenderUpload = "Not applicable"
'                Else
'                    GenderUpload = "Not applicable"
'                End If
'            End If
'            Sheet1.Range("sheet1.Gender1").value = GenderUpload
            
            Sheet1.Range("sheet1.Aadhaar").value = UCase(node.SelectSingleNode("AadhaarCardNo").text)
            'Ayush_enrol
            'Sheet1.Range("sheet1.AadhaarEnrol").value = UCase(node.SelectSingleNode("AadhaarEnrolmentId").text)
       ' Next InnerNode
    Next node
End Sub
Sub TaxExmpIntIncDtlsXMLImport()
On Error Resume Next

    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim RecExempt As Long
    Dim InnerNode As IXMLDOMNode
    Dim NatureColNo, AmountColNo, DescriptionColNo As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
    
    
   
    
    XpathOfPI = Common_year & "/TaxExmpIntIncDtls/OthersInc/OthersIncDtls"
    Set Nodelist = dom.SelectNodes(XpathOfPI)
    


    NatureColNo = Sheet5.Range("Sheet5.Nature").Column
    DescriptionColNo = Sheet5.Range("Sheet5.Description").Column
    AmountColNo = Sheet5.Range("Sheet5.Amount").Column
    
    TotalExRow = Range("Sheet5.Nature").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet5.Range("Sheet5.Nature").ClearContents
        Sheet5.Range("Sheet5.Description").ClearContents
        Sheet5.Range("Sheet5.Amount").ClearContents
  
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Exempt (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet5.Range("Sheet5.Nature").name)
    rowcount = rowcount - 1
    cnt = 0



    For Each node In Nodelist
     rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
           
            Sheet5.Cells(rowcount, NatureColNo).value = Findtext("Sec" & (node.SelectSingleNode("NatureDesc").text), "PART4_Nature")
            If UCase(node.SelectSingleNode("NatureDesc").text) = "OTH" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Any Other"
            End If
            If UCase(node.SelectSingleNode("NatureDesc").text) = "DMDP" Then
            Sheet5.Cells(rowcount, NatureColNo).value = "Defense medical disability pension"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "AGRI" Then
            
            Sheet5.Cells(rowcount, NatureColNo).value = "Agriculture Income (Less than Equal to Rs.5000)"
            
            End If
            
            
            Sheet5.Cells(rowcount, DescriptionColNo).value = UCase(node.SelectSingleNode("OthNatOfInc").text)
            Sheet5.Cells(rowcount, AmountColNo).value = UCase(node.SelectSingleNode("OthAmount").text)


            
    
        Next InnerNode
        cnt = cnt + 1
    Next node
   RecExempt = cnt
End Sub

Sub TaxCompXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
   
    
    XpathOfPI = Common_year
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes
           
            Sheet1.Range("IncD.Section89").value = UCase(node.SelectSingleNode("TaxComputation/Section89").text)
            Sheet5.Range("AIR.TaxExmpIntInc").value = UCase(node.SelectSingleNode("TaxExmpIntIncDtls/AgricultureIncome").text)
            
            

        Next InnerNode
    Next node
End Sub
Sub verpanXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim Capacity As Variant
   
    
    XpathOfPI = Common_year '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
'    For Each InnerNode In Node.ChildNodes
        
        Capacity = UCase(node.SelectSingleNode("Verification/Capacity").text)
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
            
            Sheet5.Range("Ver.capacity").value = Capacity
            
            Sheet5.Range("Ver.Place").value = UCase(node.SelectSingleNode("Verification/Place").text)
'            Sheet5.Range("Ver.Date").value = Mid(UCase(Node.SelectSingleNode("Verification/Date").Text), 9, 2) & "/" & Mid(UCase(Node.SelectSingleNode("Verification/Date").Text), 6, 2) & "/" & Mid(UCase(Node.SelectSingleNode("Verification/Date").Text), 1, 4)

           Sheet5.Range("Ver.IdentificationNoOfTRP").value = UCase(node.SelectSingleNode("TaxReturnPreparer/IdentificationNoOfTRP").text)
           Sheet5.Range("Ver.NameOfTRP").value = UCase(node.SelectSingleNode("TaxReturnPreparer/NameOfTRP").text)
           Sheet5.Range("Ver.ReImbFrmGov").value = UCase(node.SelectSingleNode("TaxReturnPreparer/ReImbFrmGov").text)
           
           
     
        
            
        
           
            
            
            

'        Next InnerNode
    Next node
End Sub
Sub SalaryXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
   
    
    XpathOfPI = Common_year & "/IncomeDeductions" '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
'        For Each InnerNode In Node.ChildNodes
           
            Sheet1.Range("IncD.IncomeFromSal_ii").value = UCase(node.SelectSingleNode("Salary").text)
            Sheet1.Range("IncD.IncomeFromOS").value = UCase(node.SelectSingleNode("IncomeOthSrc").text)
            
            Sheet1.Unprotect Password:=getmsgstate
            'Sheet1.Range("IncD.IncomeFromHP").value = UCase(Node.SelectSingleNode("TotalIncomeOfHP").Text)
            Sheet1.Protect Password:=getmsgstate
            If UCase(node.SelectSingleNode("TypeOfHP").text) = "L" Then
           Sheet1.Range("IncD.TypeOfHP").value = "Let Out"
            End If
            If UCase(node.SelectSingleNode("TypeOfHP").text) = "D" Then
           Sheet1.Range("IncD.TypeOfHP").value = "Deemed let out"
            End If
            If UCase(node.SelectSingleNode("TypeOfHP").text) = "S" Then
           Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied"
            End If
            
            Sheet1.Range("IncD.IncomeFromSal_ii").value = UCase(node.SelectSingleNode("AlwnsNotExempt").text)
            Sheet1.Range("IncD.IncomeFromSal_iii").value = UCase(node.SelectSingleNode("PerquisitesValue").text)
            Sheet1.Range("IncD.IncomeFromSal_iv").value = UCase(node.SelectSingleNode("ProfitsInSalary").text)
            
            Sheet1.Range("IncD.IncomeFromSal_v").value = UCase(node.SelectSingleNode("EntertainmntalwncUs16ii").text)
            Sheet1.Range("IncD.IncomeFromSal_ivc").value = UCase(node.SelectSingleNode("ProfessionalTaxUs16iii").text)
            
            Sheet1.Range("IncD.IncomeFromHP_i").value = UCase(node.SelectSingleNode("GrossRentReceived").text)
            Sheet1.Range("IncD.IncomeFromHP_ii").value = UCase(node.SelectSingleNode("TaxPaidlocalAuth").text)
            
            Dim IncomeFromHP As Variant
            Dim sIncomeFromHP As Variant
            
            IncomeFromHP = UCase(node.SelectSingleNode("InterestPayable").text)
            
            If IncomeFromHP = 0 Then
            sIncomeFromHP = ""
            End If
    
            Sheet1.Range("IncD.IncomeFromHP_v").value = sIncomeFromHP
'            Sheet1.Range("IncD.IncomeFromHP_v").value = UCase(Node.SelectSingleNode("InterestPayable").Text)
            

            
            Sheet1.Range("IncD.Section80C").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80C").text)
            Sheet1.Range("IncD.Section80CCC").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section").text)
            Sheet1.Range("IncD.Section80CCD_SE").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCDEmployeeOrSE").text)
            Sheet1.Range("IncD.Section80CCD1B_SE").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCD1B").text)
            Sheet1.Range("IncD.Section80CCD").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCDEmployer").text)
'            Sheet1.Range("IncD.Section80CCG").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCG").Text)
            
       
            
           Sheet1.Range("SELECT80D").value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/HealthInsurancePremium").text), "Selection80D")
'           Sheet1.Range("IncD.Section80D").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/Sec80DHealthInsurancePremiumUsr").Text)

'           Sheet1.Range("SELECT80DB").value = Findtext(UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/MedicalExpenditure").Text), "Selection80D_B")
'           Sheet1.Range("IncD.Section80DB").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/Sec80DMedicalExpenditureUsr").Text)
'
'
'           Sheet1.Range("SELECT80DC").value = Findtext(UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/PreventiveHealthCheckUp").Text), "Selection80D_C")
'           Sheet1.Range("IncD.Section80DC").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/Sec80DPreventiveHealthCheckUpUsr").Text)
'
'
            Sheet1.Range("SELECT80DD").value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDUsrType").text), "Selection80DD")
            Sheet1.Range("IncD.Section80DD").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DD").text)
            
            Sheet1.Range("SELECT80DDB").value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDBUsrType").text), "Selection80DDB")
            Sheet1.Range("IncD.Section80DDB").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDB").text)

            Sheet1.Range("IncD.Section80E").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80E").text)
            Sheet1.Range("IncD.Section80EE").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80EE").text)
            Sheet1.Range("IncD.Section80EEA").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80EEA").text)
            Sheet1.Range("IncD.Section80EEB").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80EEB").text)
            
            Sheet1.Range("IncD.Section80GG").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80GG").text)
            Sheet1.Range("IncD.Section80GGC").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80GGC").text)
            'Sheet1.Range("IncD.Section80QQB").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80QQB").Text)
            'Sheet1.Range("IncD.Section80RRB").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80RRB").Text)
            Sheet1.Range("IncD.Section80TTA").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80TTA").text)
            Sheet1.Range("IncD.Section80TTB").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80TTB").text)
            
           Sheet1.Range("SELECT80U").value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80UUsrType").text), "Selection80U")
           Sheet1.Range("IncD.Section80U").value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80U").text)
           
'        Next InnerNode
    Next node
End Sub
Sub TDSonSalaryXMLImport2()
On Error Resume Next
    Dim XpathOfTDS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfTDS = Common_year & "/TDSonSalaries/TDSonSalary" '"/ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfTDS)

    TANNoEmployer = Sheet4.Range("TDSal.TAN").Column
    TDSNameOfEmployer = Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet4.Range("TDSal.IncChrgSal").Column
    TDSTotalTax = Sheet4.Range("TDSal.TotalTDSSal").Column
    
    TotalExRow = Range("TDSal.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("TDSal.TAN").ClearContents
        Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet4.Range("TDSal.IncChrgSal").ClearContents
        Sheet4.Range("TDSal.TotalTDSSal").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet4.Range("TDSal.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet4.Cells(rowcount, TANNoEmployer).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet4.Cells(rowcount, TDSNameOfEmployer).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet4.Cells(rowcount, TDSIncomeCharge).value = UCase(node.SelectSingleNode("IncChrgSal").text)
            Sheet4.Cells(rowcount, TDSTotalTax).value = UCase(node.SelectSingleNode("TotalTDSSal").text)
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub
Sub ITXMLImport2()
On Error Resume Next
    Dim XpathOfIT As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Long
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfIT = Common_year & "/ScheduleIT/TaxPayment" '"/ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfIT)

    BSRCodeColNo = Sheet18.Range("TaxP.BSRCode").Column
    DateDepColNo = Sheet18.Range("TaxP.DateDep").Column
    SrlNoChallanColNo = Sheet18.Range("TaxP.SrlNoOfChaln").Column
    AmtColNo = Sheet18.Range("TaxP.Amt").Column
    
    TotalExRow = Range("TaxP.BSRCode").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet18.Range("TaxP.BSRCode").ClearContents
        Sheet18.Range("TaxP.DateDep").ClearContents
        Sheet18.Range("TaxP.SrlNoOfChaln").ClearContents
        Sheet18.Range("TaxP.Amt").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IT (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet18.Range("TaxP.BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet18.Cells(rowcount, BSRCodeColNo).value = UCase(node.SelectSingleNode("BSRCode").text)
            Sheet18.Cells(rowcount, DateDepColNo).value = UCase(node.SelectSingleNode("DateDep").text)
            Sheet18.Cells(rowcount, SrlNoChallanColNo).value = UCase(node.SelectSingleNode("SrlNoOfChaln").text)
            Sheet18.Cells(rowcount, AmtColNo).value = UCase(node.SelectSingleNode("Amt").text)
            
            strDate = node.SelectSingleNode("DateDep").text
            YYYY = Mid(strDate, 1, 4)
            MM = Mid(strDate, 6, 2)
            DD = Mid(strDate, 9, 2)
            strDate = DD & "/" & MM & "/" & YYYY
            
            Sheet18.Cells(rowcount, DateDepColNo).value = strDate
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub
Sub TDSOthXMLImport2()
On Error Resume Next
    Dim XpathOfTDS2 As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim TANColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant

    XpathOfTDS2 = Common_year & "/TDSonOthThanSals/TDSonOthThanSalDtls" '"/ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOfTDS2)
    
    'TDSCreditColNo = Sheet4.Range("TDS2.TdsCredit").Column
    TANColNo = Sheet4.Range("TDsOthr.TAN").Column
    FYColNo = Sheet4.Range("TDsOthr.finyear").Column
    TDSOtherColNo = Sheet4.Range("TDsOthr.Tds").Column
    TDSDeductedColNo = Sheet4.Range("TDsOthr.tdsdeducted").Column
    TDSClaimedColNo = Sheet4.Range("TDsOthr.tdsclaimed").Column
    TDSGrossAmountColNo = Sheet4.Range("TDsOthr.grossamount").Column
    TDSHeadColNo = Sheet4.Range("TDsOthr.headincome").Column
    
    TotalExRow = Range("TDsOthr.TAN").Rows.count

    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
     Sheet4.Range("TDsOthr.TAN").ClearContents
     Sheet4.Range("TDsOthr.finyear").ClearContents
     Sheet4.Range("TDsOthr.Tds").ClearContents
     Sheet4.Range("TDsOthr.tdsdeducted").ClearContents
     Sheet4.Range("TDsOthr.tdsclaimed").ClearContents
     Sheet4.Range("TDsOthr.grossamount").ClearContents
     Sheet4.Range("TDsOthr.headincome").ClearContents
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet4.Range("TDsOthr.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
       
            
            
            Sheet4.Cells(rowcount, TANColNo).value = UCase(node.SelectSingleNode("TANOfDeductor").text)
            Sheet4.Cells(rowcount, FYColNo).value = UCase(node.SelectSingleNode("DeductedYr").text)
            Sheet4.Cells(rowcount, TDSOtherColNo).value = UCase(node.SelectSingleNode("BroughtFwdTDSAmt").text)
            Sheet4.Cells(rowcount, TDSDeductedColNo).value = UCase(node.SelectSingleNode("TDSDeducted").text)
            Sheet4.Cells(rowcount, TDSClaimedColNo).value = UCase(node.SelectSingleNode("TDSClaimed").text)
            Sheet4.Cells(rowcount, TDSGrossAmountColNo).value = UCase(node.SelectSingleNode("GrossAmount").text)
            Dim HeadOfIncome As Variant
              HeadOfIncome = UCase(node.SelectSingleNode("HeadOfIncome").text)
            
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
            
            
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub
Sub TDSOthXMLImport3()
On Error Resume Next
    Dim XpathOfTDS21 As String
    Dim TotalXMLRow1 As Long
    Dim TotalDiffRow1 As Long
    Dim RecTDS11, RecTDS21, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow1 As Long
    Dim TotalXMLRow, TotalExRow, TotalDiffRow As Long
    Dim PANColNo, AadhaarColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant

    XpathOfTDS21 = Common_year & "/ScheduleTDS3Dtls/TDS3Details" '"/ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOfTDS21)
    
    'TDSCreditColNo = Sheet4.Range("TDS2.TdsCredit").Column
    PANColNo = Sheet4.Range("TDsOthr2.PAN").Column
    AadhaarColNo = Sheet4.Range("TDsOthr2.AADHAAR").Column
    FYColNo = Sheet4.Range("TDsOthr2.finyear").Column
    TDSOtherColNo = Sheet4.Range("TDsOthr2.Tds").Column
    TDSDeductedColNo = Sheet4.Range("TDsOthr2.tdsdeducted").Column
    TDSClaimedColNo = Sheet4.Range("TDsOthr2.tdsclaimed").Column
    TDSGrossAmountColNo = Sheet4.Range("TDsOthr2.grossamount").Column
    TDSHeadColNo = Sheet4.Range("TDsOthr2.headincome").Column
    
    TotalExRow = Range("TDsOthr2.PAN").Rows.count

    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
     Sheet4.Range("TDsOthr2.PAN").ClearContents
     Sheet4.Range("TDsOthr2.AADHAAR").ClearContents
     Sheet4.Range("TDsOthr2.finyear").ClearContents
     Sheet4.Range("TDsOthr2.Tds").ClearContents
     Sheet4.Range("TDsOthr2.tdsdeducted").ClearContents
     Sheet4.Range("TDsOthr2.tdsclaimed").ClearContents
     Sheet4.Range("TDsOthr2.grossamount").ClearContents
     Sheet4.Range("TDsOthr2.headincome").ClearContents
    End If
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth1 (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet4.Range("TDsOthr2.PAN").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
       
            Sheet4.Cells(rowcount, PANColNo).value = UCase(node.SelectSingleNode("PANofTenant").text)
            Sheet4.Cells(rowcount, AadhaarColNo).value = UCase(node.SelectSingleNode("AadhaarofTenant").text)
            
            Sheet4.Cells(rowcount, FYColNo).value = UCase(node.SelectSingleNode("DeductedYr").text)
            Sheet4.Cells(rowcount, TDSOtherColNo).value = UCase(node.SelectSingleNode("BroughtFwdTDSAmt").text)
            Sheet4.Cells(rowcount, TDSDeductedColNo).value = UCase(node.SelectSingleNode("TDSDeducted").text)
            Sheet4.Cells(rowcount, TDSClaimedColNo).value = UCase(node.SelectSingleNode("TDSClaimed").text)
            Sheet4.Cells(rowcount, TDSGrossAmountColNo).value = UCase(node.SelectSingleNode("GrossAmount").text)
            Dim HeadOfIncome As Variant
              HeadOfIncome = UCase(node.SelectSingleNode("HeadOfIncome").text)
            
            If HeadOfIncome = "HP" Then
            HeadOfIncome = "Income from House property"
            End If
            
            
            
            Sheet4.Cells(rowcount, TDSHeadColNo).value = HeadOfIncome
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS11 = cnt
End Sub
Sub TCSXMLImport2()
On Error Resume Next
    Dim XpathOfTCS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTCS As Long
    Dim TANColNo, DEDNameColNo, TaxColNo, ClaimColNo, AmntClaimedBySpouseTCS, Amtfrom26AS, CollectedYr As Variant
    Dim TotalExRow, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    XpathOfTCS = Common_year & "/ScheduleTCS/TCS" '"/ITR/ScheduleTCS/TCS"
    Set Nodelist = dom.SelectNodes(XpathOfTCS)
    
    TANColNo = Sheet17.Range("TCS.TAN").Column
    DEDNameColNo = Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    Amtfrom26AS = Sheet17.Range("TCS.AmountPaid").Column
    TaxColNo = Sheet17.Range("TCS.TotalTCS").Column
    ClaimColNo = Sheet17.Range("TCS.AmtTCSClaimedThisYear").Column
'    AmntClaimedBySpouseTCS = Sheet17.Range("TCS.AmtClaimedBySpouse").Column
    
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet17.Range("TCS.TAN").ClearContents
    Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
    Sheet17.Range("TCS.AmountPaid").ClearContents
    Sheet17.Range("TCS.TotalTCS").ClearContents
    Sheet17.Range("TCS.AmtTCSClaimedThisYear").ClearContents
'    Sheet17.Range("TCS.AmtClaimedBySpouse").ClearContents
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TCS (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet17.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet17.Cells(rowcount, TANColNo).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet17.Cells(rowcount, DEDNameColNo).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            ''Sheet4.Cells(rowCount, CollectedYr).value = Node.SelectSingleNode("CollectedYr").Text
            Sheet17.Cells(rowcount, Amtfrom26AS).value = node.SelectSingleNode("Amtfrom26AS").text
            Sheet17.Cells(rowcount, TaxColNo).value = node.SelectSingleNode("TotalTCS").text
            Sheet17.Cells(rowcount, ClaimColNo).value = node.SelectSingleNode("AmtTCSClaimedThisYear").text
            Sheet17.Cells(rowcount, AmntClaimedBySpouseTCS).value = node.SelectSingleNode("AmtClaimedBySpouse").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTCS = cnt
End Sub
Sub FilingInfoXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
     Dim Filingtype As Variant
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A, ReturnFileSec As Variant
    Dim DateofOriginalfile As Variant
    Dim NoticeDateussec As Variant
    Dim residentialStatus
    Dim TaxStatus
    Dim TaxRefundable
    Dim iProvisoFlag, sProvisoFlag As Variant
    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
    
    XpathOfPI = Common_year & "/FilingStatus"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
'        For Each InnerNode In Node.ChildNodes
        
        
        iReturnFile = UCase(node.SelectSingleNode("ReturnFileSec").text)
            
    
        If iReturnFile = "11" Then
           sReturnFile = "139(1)-On or before due date"
        ElseIf iReturnFile = "12" Then
           sReturnFile = "139(4)-After due date"
        ElseIf iReturnFile = "13" Then
           sReturnFile = "142(1)"
        ElseIf iReturnFile = "14" Then
           sReturnFile = "148"
        ElseIf iReturnFile = "15" Then
           sReturnFile = "153A"
        ElseIf iReturnFile = "16" Then
           sReturnFile = "153C"
        ElseIf iReturnFile = "17" Then
           sReturnFile = "139(5)-Revised Return"
        ElseIf iReturnFile = "18" Then
           sReturnFile = "139(9)"
        ElseIf iReturnFile = "20" Then
        'Ayush J200924
           'sReturnFile = "119(2)(b)- after condonation of delay"
           sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
        End If
            
        Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
        
   
        iProvisoFlag = UCase(node.SelectSingleNode("SeventhProvisio139").text)
        If iProvisoFlag = "Y" Then
        sProvisoFlag = "Yes"
        ElseIf iProvisoFlag = "N" Then
        sProvisoFlag = "No"
        End If
            
        Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
        
        iDepositAmountFlag = UCase(node.SelectSingleNode("DepAmtAggAmtExcd1CrPrYrFlg").text)
        If iDepositAmountFlag = "Y" Then
        sDepositAmountFlag = "Yes"
        ElseIf iDepositAmountFlag = "N" Then
        sDepositAmountFlag = "No"
        End If
            
        Sheet1.Range("Sheet1.DepositAmountFlag").value = sDepositAmountFlag
        Sheet1.Range("Sheet1.DepositAmount").value = UCase(node.SelectSingleNode("AmtSeventhProvisio139i").text)
        
        iAggrigateAmountFlag = UCase(node.SelectSingleNode("IncrExpAggAmt2LkTrvFrgnCntryFlg").text)
        If iAggrigateAmountFlag = "Y" Then
        sAggrigateAmountFlag = "Yes"
        ElseIf iAggrigateAmountFlag = "N" Then
        sAggrigateAmountFlag = "No"
        End If
            
        Sheet1.Range("Sheet1.AggrigateAmountFlag").value = sAggrigateAmountFlag
        Sheet1.Range("Sheet1.AggrigateAmount").value = UCase(node.SelectSingleNode("AmtSeventhProvisio139ii").text)
        
        iAggrigateAmountFlag1 = UCase(node.SelectSingleNode("IncrExpAggAmt1LkElctrctyPrYrFlg").text)
        If iAggrigateAmountFlag1 = "Y" Then
        sAggrigateAmountFlag1 = "Yes"
        ElseIf iAggrigateAmountFlag1 = "N" Then
        sAggrigateAmountFlag1 = "No"
        End If
            
        Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = sAggrigateAmountFlag
        Sheet1.Range("Sheet1.AggrigateAmount1").value = UCase(node.SelectSingleNode("AmtSeventhProvisio139iii").text)
       
   
   
   
   
    ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
            
    ReturnFileSec = Mid(ReturnFileSec, 1, 2)
        If ReturnFileSec = "17" Then
Sheet1.Range("sheet1.ReceiptNo").value = node.SelectSingleNode("ReceiptNo").text

DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
        
        Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
     Else
        Sheet1.Range("Sheet1.ReceiptNo").value = node.SelectSingleNode("ReceiptNo").text
        
        DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
        
        Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
        
        Sheet1.Range("sheet1.NoticeNo").value = node.SelectSingleNode("NoticeNo").text
        
        NoticeDateussec = node.SelectSingleNode("NoticeDateUnderSec").text
        
        Sheet1.Range("sheet1.NoticeDate").value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
      End If
        
    
        
'             Sheet1.Range("sheet1.DesigOfficerWardorCircle").value = UCase(Node.SelectSingleNode("DesigOfficerWardorCircle").Text)
                        
            iPort5A = UCase(node.SelectSingleNode("PortugeseCC5A").text)
            If iPort5A = "Y" Then
            sPort5A = "Yes"
'            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A
            End If
            
            If iPort5A = "N" Then
            sPort5A = "No"
'            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A
            End If
'            sPort5A = IIf(iPort5A = "Y", "Yes", IIf(iPort5A = "N", "No", iPort5A))
'            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A

        
        
            Dim RepAssessee As Variant
            Dim sRepAssessee As Variant
            
            RepAssessee = UCase(node.SelectSingleNode("AsseseeRepFlg").text)
            If RepAssessee = "Y" Then
            sRepAssessee = "Yes"
            
            ElseIf RepAssessee = "N" Then
             sRepAssessee = "No"
            End If
            
            Sheet1.Range("Sheet1.RepAssessee").value = sRepAssessee
            
          
          
          
          
          Sheet1.Range("sheet1.NameRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepName").text)
          
          Dim iCapacity As Variant
          Dim sCapacity As Variant
          
          iCapacity = UCase(node.SelectSingleNode("AssesseeRep/RepCapacity").text)
           
           If iCapacity = "L" Then
           sCapacity = "Legal Heir"
           
           ElseIf iCapacity = "M" Then
           sCapacity = "Manager"
           
           ElseIf iCapacity = "G" Then
           sCapacity = "Guardian"
           
           ElseIf iCapacity = "O" Then
           sCapacity = "Other"
           
           End If
           
           Sheet1.Range("sheet1.CapacityRep").value = sCapacity
          
          Sheet1.Range("sheet1.AddressRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepAddress").text)
         
          Sheet1.Range("sheet1.PANRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepPAN").text)
          
          Sheet1.Range("sheet1.PANRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepPAN").text)
           
          Sheet1.Range("sheet1.PANOFSPOUSE").value = UCase(node.SelectSingleNode("PANOfSpouse").text)
'        Next InnerNode
    Next node
End Sub
Sub VeriInfoXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfPI = Common_year & "/Verification"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
'        For Each InnerNode In Node.ChildNodes

   
        'ReturnFile Sec/ Port 5A
        
            Sheet5.Range("Ver.FatherName").value = UCase(node.SelectSingleNode("Declaration/FatherName").text)
            Sheet5.Range("Ver.AssesseeVerName").value = UCase(node.SelectSingleNode("Declaration/AssesseeVerName").text)
            Sheet5.Range("Ver.PAN").value = UCase(node.SelectSingleNode("Declaration/AssesseeVerPAN").text)
                                                
'        Next InnerNode
    Next node
End Sub


Sub RefundInfoXMLImport2()
On Error Resume Next
    Dim XpathOfRefund As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim IFSC, BANKNAME, ACCNO, CheckBox As Variant
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfRefund = Common_year & "/Refund/BankAccountDtls/AddtnlBankDetails"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)

    IFSC = Range("SchBA.IFSC").Column
    BANKNAME = Range("SchBA.BankName").Column
    ACCNO = Range("SchBA.AcntNo").Column
    CheckBox = Range("tempXML").Column
    
    TotalExRow = Range("SchBA.IFSC").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Range("SchBA.IFSC").ClearContents
        Range("SchBA.BankName").ClearContents
        Range("SchBA.AcntNo").ClearContents
        Range("tempxml").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_BANK (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet5.Range("SchBA.IFSC").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        

            
            Sheet5.Cells(rowcount, IFSC).value = UCase(node.SelectSingleNode("IFSCCode").text)
            Sheet5.Cells(rowcount, BANKNAME).value = UCase(node.SelectSingleNode("BankName").text)
            Sheet5.Cells(rowcount, ACCNO).value = UCase(node.SelectSingleNode("BankAccountNo").text)
            Sheet5.Cells(rowcount, CheckBox).value = UCase(node.SelectSingleNode("UseForRefund").text)
            
            If CheckBox = True Then
            CheckBox = "true"
            ElseIf CheckBox = False Then
            CheckBox = "false"
            ElseIf CheckBox = "" Then
            CheckBox = "false"
            End If
            
LinkCheckBoxes
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub AeInfoXMLImport()
On Error Resume Next
    Dim XpathOfAE As String
    Dim Nodelist As IXMLDOMNodeList
    Dim RecAE As Long
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim Holiding, Income As Variant
    Dim TotalExRow As Long
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long



    XpathOfAE = Common_year & "/Us44aeHeavy"
   Set Nodelist = dom.SelectNodes(XpathOfAE)


   Holiding = Range("HV_HoldingPeriod").Column
   Income = Range("HV_Income").Column


    TotalExRow = Range("HV_HoldingPeriod").Rows.count

    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Range("HV_HoldingPeriod").ClearContents
        Range("HV_Income").ClearContents

    End If

rowcount = getRowNo(Sheet2.Range("HV_HoldingPeriod").name)
rowcount = rowcount - 1
cnt = 0


     For Each node In Nodelist
     rowcount = rowcount + 1

        For Each InnerNode In node.ChildNodes
            Sheet2.Cells(rowcount, Holiding).value = UCase(node.SelectSingleNode("HoldingPeriod").text)
            Sheet2.Cells(rowcount, Income).value = UCase(node.SelectSingleNode("IncomePerVehicle").text)
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecAE = cnt
End Sub

'Sub NOBXMLImport2()
'On Error Resume Next
'    Dim XpathOfNOB As String
'    Dim TotalXMLRow As Long
'    Dim TotalDiffRow As Long
'    Dim RecTDS1 As Long
'    Dim NodeList As IXMLDOMNodeList
'    Dim Node As IXMLDOMNode
'    Dim InnerNode As IXMLDOMNode
'    Dim NOB As Variant
'    Dim TradeName1 As Variant
'    Dim TradeName2 As Variant
'    Dim TradeName3 As Variant
'    Dim strDate As String
'    Dim YYYY, MM, DD As String
'    Dim TotalExRow As Long
'    Dim sNOB, iNOB As Variant
'
'    XpathOfNOB = Common_year & "/NatOfBus/NatureOfBusiness" '"/ITR/TDSonSalaries"
'    Set NodeList = dom.SelectNodes(XpathOfNOB)
'
'    NOB = Sheet3.Range("NOB.Code").Column
'    TradeName1 = Sheet3.Range("NOB.TradeNameOrDescription").Column
'    TradeName2 = Sheet3.Range("NOB.TradeNameOrDescription2").Column
'    TradeName3 = Sheet3.Range("NOB.TradeNameOrDescription3").Column
'
'    TotalExRow = Range("NOB.Code").Rows.Count
'
'    TotalXMLRow = NodeList.length
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'        Sheet3.Range("NOB.Code").ClearContents
'        Sheet3.Range("NOB.TradeNameOrDescription").ClearContents
'        Sheet3.Range("NOB.TradeNameOrDescription2").ClearContents
'        Sheet3.Range("NOB.TradeNameOrDescription3").ClearContents
'
'
'    End If
'
'    rowCount = getRowNo(Sheet3.Range("NOB.Code").name)
'    rowCount = rowCount - 1
'    cnt = 0
'
'    For Each Node In NodeList
'        rowCount = rowCount + 1
'        For Each InnerNode In Node.ChildNodes
'            'Sheet3.Cells(rowCount, NOB).value = UCase(Node.SelectSingleNode("Code").Text)
'
'            iNOB = UCase(Node.SelectSingleNode("Code").Text)
'            sNOB = Findtext(iNOB, "NOB")
'            Sheet3.Cells(rowCount, NOB).value = sNOB
'
'            Sheet3.Cells(rowCount, TradeName1).value = UCase(Node.SelectSingleNode("TradeName1").Text)
'            Sheet3.Cells(rowCount, TradeName2).value = UCase(Node.SelectSingleNode("TradeName2").Text)
'            Sheet3.Cells(rowCount, TradeName3).value = UCase(Node.SelectSingleNode("TradeName3").Text)
'
'
'        Next InnerNode
'        cnt = cnt + 1
'    Next Node
'    RecTDS1 = cnt
'End Sub
Sub GstXMLImport()
On Error Resume Next
   On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim RecTDS1   As Variant
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim GSTRnoColNo, GSTRAmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
        
    XpathOfPI = Common_year & "/ScheduleBP/TurnoverGrsRcptForGSTIN"
    Set Nodelist = dom.SelectNodes(XpathOfPI)
    


    GSTRnoColNo = Sheet3.Range("BP.GSTRno").Column
    GSTRAmountColNo = Sheet3.Range("BP.GSTRAmount").Column
    
    TotalExRow = Range("BP.GSTRno").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet3.Range("BP.GSTRno").ClearContents
        Sheet3.Range("BP.GSTRAmount").ClearContents
  
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Gst (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet3.Range("BP.GSTRno").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
             Sheet3.Cells(rowcount, GSTRnoColNo).value = UCase(node.SelectSingleNode("GSTINNo").text)
            Sheet3.Cells(rowcount, GSTRAmountColNo).value = UCase(node.SelectSingleNode("AmtTurnGrossRcptGSTIN").text)
            
            
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub
Sub BPXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
    
    
   
    
    XpathOfPI = Common_year & "/ScheduleBP"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
   
        For Each InnerNode In node.ChildNodes
           
            Sheet3.Range("BP_E1a").value = UCase(node.SelectSingleNode("PersumptiveInc44AD/GrsTrnOverBank").text)
            Sheet3.Range("BP_E1b").value = UCase(node.SelectSingleNode("PersumptiveInc44AD/GrsTrnOverAnyOthMode").text)
            Sheet3.Range("BP_E2a").value = UCase(node.SelectSingleNode("PersumptiveInc44AD/PersumptiveInc44AD6Per").text)
            Sheet3.Range("BP_E2b").value = UCase(node.SelectSingleNode("PersumptiveInc44AD/PersumptiveInc44AD8Per").text)
            
            Sheet3.Range("BP_E3_GR").value = UCase(node.SelectSingleNode("PersumptiveInc44ADA/GrsReceipt").text)
            Sheet3.Range("BP_E4_PI44ADA").value = UCase(node.SelectSingleNode("PersumptiveInc44ADA/TotPersumptiveInc44ADA").text)
            
            
            Sheet3.Range("BP_E9").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/PartnerMemberOwnCapital").text)
            Sheet3.Range("BP_E10").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/SecuredLoans").text)
            Sheet3.Range("BP_E11").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/UnSecuredLoans").text)
            Sheet3.Range("BP_E12").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/Advances").text)
            Sheet3.Range("BP_E13").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/SundryCreditors").text)
            Sheet3.Range("BP_E14").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/OthrCurrLiab").text)
            
            Sheet3.Range("BP_E16").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/FixedAssets").text)
            Sheet3.Range("BP_E17").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/Inventories").text)
            Sheet3.Range("BP_E18").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/SundryDebtors").text)
            Sheet3.Range("BP_E19").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/BalWithBanks").text)
            Sheet3.Range("BP_E20").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/CashInHand").text)
            Sheet3.Range("BP_E21").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/LoansAndAdvances").text)
            Sheet3.Range("BP_E22").value = UCase(node.SelectSingleNode("FinanclPartclrOfBusiness/OtherAssets").text)
        
    Next InnerNode
   
    Next node
   
End Sub
Sub IncomeInfoXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sTypeOfHP, iTypeOfHP As Variant
        
    XpathOfPI = Common_year & "/IncomeDeductions"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

            iTypeOfHP = UCase(node.SelectSingleNode("TypeOfHP").text)
            sTypeOfHP = IIf(iTypeOfHP = "S", "Self Occupied", IIf(iTypeOfHP = "L", "Let Out", IIf(iTypeOfHP = "D", "Deemed Let Out", iTypeOfHP)))
            Sheet1.Range("IncD.TypeOfHP").value = sTypeOfHP
            
            'Sheet1.Range("IncD.IncomeFromHP").value = UCase(Node.SelectSingleNode("TotalIncomeOfHP").Text)
            Sheet1.Range("IncD.IncomeFromOS").value = UCase(node.SelectSingleNode("IncomeOthSrc").text)
        Next InnerNode
    Next node
End Sub
Sub ImmovableXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfPI = Common_year & "/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

   
        'ReturnFile Sec/ Port 5A
        
             Dim str, Str1 As String
        str = UCase(node.SelectSingleNode("ImmovableFlag").text)
           If str = "Y" Then
            Sheet13.Range("IsImmovableAsset").value = "Yes"
            ElseIf str = "N" Then
            Sheet13.Range("IsImmovableAsset").value = "No"
            Else
            Sheet13.Range("IsImmovableAsset").value = "(Select)"
            End If
            
         Str1 = UCase(node.SelectSingleNode("InterstAOPFlag").text)
            If Str1 = "Y" Then
            Sheet13.Range("IsInterestHeld").value = "Yes"
            ElseIf Str1 = "N" Then
            Sheet13.Range("IsInterestHeld").value = "No"
            Else
            Sheet13.Range("IsInterestHeld").value = "(Select)"
            End If
                       
                                                
        Next InnerNode
    Next node
End Sub
Sub ALInfoXMLImport3()
On Error Resume Next
    Dim XpathOfAL As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim DescriptionColNo, FlatColNo, PremisesColNo, RoadColNo, AreaColNo, TownColNo, StateColNo, CountryColNo, PincodeColNo, ZipcodeColNo, AmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, cnt, TotalDiffRow, RecAL As Long
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim iCountry, sCountry, iState, sState As Variant
        
        
    XpathOfAL = Common_year & "/ScheduleAL/ImmovableDetails"
    Set Nodelist = dom.SelectNodes(XpathOfAL)

     
    DescriptionColNo = Sheet13.Range("SchAL.A.Description").Column
    FlatColNo = Sheet13.Range("SchAL.A.Address_Flat").Column
    PremisesColNo = Sheet13.Range("SchAL.A.Address_Village").Column
    RoadColNo = Sheet13.Range("SchAL.A.Address_Road").Column
    AreaColNo = Sheet13.Range("SchAL.A.Address_Area").Column
    TownColNo = Sheet13.Range("SchAL.A.Address_City").Column
    StateColNo = Sheet13.Range("SchAL.A.Address_State").Column
    CountryColNo = Sheet13.Range("SchAL.A.Address_Country").Column
    PincodeColNo = Sheet13.Range("SchAL.A.Address_Pin").Column
    ZipcodeColNo = Sheet13.Range("SchAL.A.Address_Zip").Column
    AmountColNo = Sheet13.Range("SchAL.A.Amount").Column
    
    TotalExRow = Range("SchAL.A.Description").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet13.Range("SchAL.A.Description").ClearContents
    Sheet13.Range("SchAL.A.Address_Flat").ClearContents
    Sheet13.Range("SchAL.A.Address_Village").ClearContents
    Sheet13.Range("SchAL.A.Address_Road").ClearContents
    Sheet13.Range("SchAL.A.Address_Area").ClearContents
    Sheet13.Range("SchAL.A.Address_City").ClearContents
    Sheet13.Range("SchAL.A.Address_State").ClearContents
    Sheet13.Range("SchAL.A.Address_Country").ClearContents
    Sheet13.Range("SchAL.A.Address_Pin").ClearContents
    Sheet13.Range("SchAL.A.Address_Zip").ClearContents
    Sheet13.Range("SchAL.A.Amount").ClearContents
    
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_AL3 (TotalDiffRow)
    End If
    rowcount = 0
    rowcount = getRowNo(Sheet13.Range("SchAL.A.Description").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
    rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        
'        Dim Str As String
'        Str = UCase(Node.SelectSingleNode("ImmovableFlag").Text)
'           If Str = "Y" Then
'            Sheet13.Range("IsImmovableAsset").value = "Yes"
'            End If
            
            
            
            Sheet13.Cells(rowcount, DescriptionColNo).value = UCase(node.SelectSingleNode("Description").text)
            Sheet13.Cells(rowcount, FlatColNo).value = UCase(node.SelectSingleNode("AddressAL/ResidenceNo").text)
            Sheet13.Cells(rowcount, PremisesColNo).value = UCase(node.SelectSingleNode("AddressAL/ResidenceName").text)
            Sheet13.Cells(rowcount, RoadColNo).value = UCase(node.SelectSingleNode("AddressAL/RoadOrStreet").text)
            Sheet13.Cells(rowcount, AreaColNo).value = UCase(node.SelectSingleNode("AddressAL/LocalityOrArea").text)
            Sheet13.Cells(rowcount, TownColNo).value = UCase(node.SelectSingleNode("AddressAL/CityOrTownOrDistrict").text)
            
             iCountry = UCase(node.SelectSingleNode("AddressAL/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet13.Cells(rowcount, CountryColNo).value = sCountry
            
            iState = UCase(node.SelectSingleNode("AddressAL/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet13.Cells(rowcount, StateColNo).value = sState
    
            
            
            
'            Sheet13.Cells(rowCount, StateColNo).value = UCase(Node.SelectSingleNode("ImmovableDetails/AddressAL/StateCode").Text)
'            Sheet13.Cells(rowCount, CountryColNo).value = UCase(Node.SelectSingleNode("ImmovableDetails/AddressAL/CountryCode").Text)
            
            
            
            Sheet13.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("AddressAL/PinCode").text)
            Sheet13.Cells(rowcount, AmountColNo).value = UCase(node.SelectSingleNode("Amount").text)
            
'            Sheet13.Range("SchAL.B.Bank").value = UCase(Node.SelectSingleNode("MovableAsset/DepositsInBank").Text)
'            Sheet13.Range("SchAL.B.Share").value = UCase(Node.SelectSingleNode("MovableAsset/SharesAndSecurities").Text)
'            Sheet13.Range("SchAL.B.Insurance").value = UCase(Node.SelectSingleNode("MovableAsset/InsurancePolicies").Text)
'            Sheet13.Range("SchAL.B.Loans").value = UCase(Node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").Text)
'            Sheet13.Range("SchAL.B.Cash").value = UCase(Node.SelectSingleNode("MovableAsset/CashInHand").Text)
'            Sheet13.Range("SchAL.B.Jwellery").value = UCase(Node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").Text)
'            Sheet13.Range("SchAL.B.Art").value = UCase(Node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").Text)
'            Sheet13.Range("SchAL.B.Vehicles").value = UCase(Node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").Text)
'            'Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
'            Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("LiabilityInRelatAssets").Text)


            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecAL = cnt
    
End Sub
Sub ALInfoXMLImport4()
On Error Resume Next
    Dim XpathOfAL As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim NameOfFirmColNo, FlatColNo, PremisesColNo, RoadColNo, AreaColNo, TownColNo, StateColNo, CountryColNo, PANColNo, PincodeColNo, ZipcodeColNo, AmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, cnt, TotalDiffRow, RecAL2 As Long
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim iCountry, sCountry, iState, sState As Variant
    
        
    XpathOfAL = Common_year & "/ScheduleAL/InterestHeldInaAsset"
    Set Nodelist = dom.SelectNodes(XpathOfAL)

     NameOfFirmColNo = Sheet13.Range("SchAL.C.Name").Column
    FlatColNo = Sheet13.Range("SchAL.C.Address_Flat").Column
    PremisesColNo = Sheet13.Range("SchAL.C.Address_Village").Column
    RoadColNo = Sheet13.Range("SchAL.C.Address_Road").Column
    AreaColNo = Sheet13.Range("SchAL.C.Address_Area").Column
    TownColNo = Sheet13.Range("SchAL.C.Address_City").Column
    StateColNo = Sheet13.Range("SchAL.C.Address_State").Column
    CountryColNo = Sheet13.Range("SchAL.C.Address_Country").Column
    PincodeColNo = Sheet13.Range("SchAL.C.Address_Pin").Column
    ZipcodeColNo = Sheet13.Range("SchAL.C.Address_Zip").Column
    PANColNo = Sheet13.Range("SchAL.C.PAN").Column
    AmountColNo = Sheet13.Range("SchAL.C.Investment").Column
    
    TotalExRow = Range("SchAL.C.Name").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet13.Range("SchAL.C.Name").ClearContents
    Sheet13.Range("SchAL.C.Address_Flat").ClearContents
    Sheet13.Range("SchAL.C.Address_Village").ClearContents
    Sheet13.Range("SchAL.C.Address_Road").ClearContents
    Sheet13.Range("SchAL.C.Address_Area").ClearContents
    Sheet13.Range("SchAL.C.Address_City").ClearContents
    Sheet13.Range("SchAL.C.Address_Pin").ClearContents
    Sheet13.Range("SchAL.C.Address_Zip").ClearContents
    Sheet13.Range("SchAL.C.PAN").ClearContents
    Sheet13.Range("SchAL.C.Investment").ClearContents
    Sheet13.Range("SchAL.C.Address_State").ClearContents
    Sheet13.Range("SchAL.C.Address_Country").ClearContents
    
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_AL2 (TotalDiffRow)
    End If
    rowcount = 0
    rowcount = getRowNo(Sheet13.Range("SchAL.C.Name").name)
    rowcount = rowcount - 1
    cnt = 0
    
    
    
    
    
    For Each node In Nodelist
    rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        
'        Dim Str As String
'        Str = UCase(Node.SelectSingleNode("ImmovableFlag").Text)
'           If Str = "Y" Then
'            Sheet13.Range("IsInterestHeld").value = "Yes"
'            End If
            
            
            
            Sheet13.Cells(rowcount, NameOfFirmColNo).value = UCase(node.SelectSingleNode("NameOfFirm").text)
            Sheet13.Cells(rowcount, FlatColNo).value = UCase(node.SelectSingleNode("AddressAL/ResidenceNo").text)
            Sheet13.Cells(rowcount, PremisesColNo).value = UCase(node.SelectSingleNode("AddressAL/ResidenceName").text)
            Sheet13.Cells(rowcount, RoadColNo).value = UCase(node.SelectSingleNode("AddressAL/RoadOrStreet").text)
            Sheet13.Cells(rowcount, AreaColNo).value = UCase(node.SelectSingleNode("AddressAL/LocalityOrArea").text)
            Sheet13.Cells(rowcount, TownColNo).value = UCase(node.SelectSingleNode("AddressAL/CityOrTownOrDistrict").text)
            
             iCountry = UCase(node.SelectSingleNode("AddressAL/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet13.Cells(rowcount, CountryColNo).value = sCountry
            
            iState = UCase(node.SelectSingleNode("AddressAL/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet13.Cells(rowcount, StateColNo).value = sState
    
            
            
            
            
            
            'Sheet13.Cells(rowCount, StateColNo).value = UCase(Node.SelectSingleNode("ImmovableDetails/AddressAL/StateCode").Text)
            
            Sheet13.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("AddressAL/PinCode").text)
            Sheet13.Cells(rowcount, PANColNo).value = UCase(node.SelectSingleNode("PanOfFirm").text)
            Sheet13.Cells(rowcount, AmountColNo).value = UCase(node.SelectSingleNode("AssesseInvestment").text)
            
'            Sheet13.Range("SchAL.B.Bank").value = UCase(Node.SelectSingleNode("MovableAsset/DepositsInBank").Text)
'            Sheet13.Range("SchAL.B.Share").value = UCase(Node.SelectSingleNode("MovableAsset/SharesAndSecurities").Text)
'            Sheet13.Range("SchAL.B.Insurance").value = UCase(Node.SelectSingleNode("MovableAsset/InsurancePolicies").Text)
'            Sheet13.Range("SchAL.B.Loans").value = UCase(Node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").Text)
'            Sheet13.Range("SchAL.B.Cash").value = UCase(Node.SelectSingleNode("MovableAsset/CashInHand").Text)
'            Sheet13.Range("SchAL.B.Jwellery").value = UCase(Node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").Text)
'            Sheet13.Range("SchAL.B.Art").value = UCase(Node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").Text)
'            Sheet13.Range("SchAL.B.Vehicles").value = UCase(Node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").Text)
'            'Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
'            Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("LiabilityInRelatAssets").Text)
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecAL2 = cnt
    
End Sub
Sub ALveriXMLImport2()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant

    XpathOfPI = Common_year & "/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

   
        'ReturnFile Sec/ Port 5A
        
            Sheet13.Range("SchAL.B.Bank").value = UCase(node.SelectSingleNode("MovableAsset/DepositsInBank").text)
            Sheet13.Range("SchAL.B.Share").value = UCase(node.SelectSingleNode("MovableAsset/SharesAndSecurities").text)
            Sheet13.Range("SchAL.B.Insurance").value = UCase(node.SelectSingleNode("MovableAsset/InsurancePolicies").text)
            Sheet13.Range("SchAL.B.Loans").value = UCase(node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").text)
            Sheet13.Range("SchAL.B.Cash").value = UCase(node.SelectSingleNode("MovableAsset/CashInHand").text)
            Sheet13.Range("SchAL.B.Jwellery").value = UCase(node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").text)
            Sheet13.Range("SchAL.B.Art").value = UCase(node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").text)
            Sheet13.Range("SchAL.B.Vehicles").value = UCase(node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").text)
            'Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
          
             Sheet13.Range("SchAL.Liability").value = UCase(node.SelectSingleNode("LiabilityInRelatAssets").text)
             
           ' Sheet13.Range("IsImmovableAsset").value = UCase(Node.SelectSingleNode("ImmovableFlag").Text)
                                                
        Next InnerNode
    Next node
End Sub
Sub setDiffTblinfo_80G_A()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet6.Range("Per10080G.DoneeName").count
    Set rangecells = Sheet6.Range("Per10080G.DoneeName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
 'Konda updated 09-03-2026--V0.5
    '    DefinedgridNameRange = "Per10080G.DoneeName||Per10080G.AddrDetail||Per10080G.CityOrTownOrDistrict||Per10080G.StateCode||Per10080G.PinCode||Per10080G.DoneePAN||Per10080G.Donationcash||Per10080G.DonationAmt||Per10080G.TotDonation||Per10080G.EligibleAmt"
    DefinedgridNameRange = "Per10080G.DoneeName||Per10080G.AddrDetail||Per10080G.CityOrTownOrDistrict||Per10080G.StateCode||Per10080G.PinCode||Per10080G.DoneePAN||Per10080G.Donationcash||Per10080G.DonationAmt||Per10080G.Traref||Per10080G.IFSC||Per10080G.TotDonation||Per10080G.EligibleAmt||Data_80G_A||comb_80G_A||Comb_donation_80G_A||Comb_PAN_80G_A"
'====================
End Sub

Sub AddDiffRows_80G_A(DiffRows As Long)
    setDiffTblinfo_80G_A
    Sheet6.Activate
    searchLastRow ("Per10080G.DoneeName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub XMLImport_80G_A()
On Error Resume Next
    Dim XpathOf80G_A As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo As Variant
    
    XpathOf80G_A = Common_year & "/Schedule80G/Don100Percent/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_A)
    
    NameColNo = Sheet6.Range("Per10080G.DoneeName").Column
    AddressColNo = Sheet6.Range("Per10080G.AddrDetail").Column
    CityColNo = Sheet6.Range("Per10080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("Per10080G.StateCode").Column
    PincodeColNo = Sheet6.Range("Per10080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("Per10080G.DoneePAN").Column
    AmountColNo = Sheet6.Range("Per10080G.DonationAmt").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    TotalExRow = Range("Per10080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet6.Range("Per10080G.DoneeName").ClearContents
        Sheet6.Range("Per10080G.AddrDetail").ClearContents
        Sheet6.Range("Per10080G.CityOrTownOrDistrict").ClearContents
        Sheet6.Range("Per10080G.StateCode").ClearContents
        Sheet6.Range("Per10080G.PinCode").ClearContents
        Sheet6.Range("Per10080G.DoneePAN").ClearContents
        Sheet6.Range("Per10080G.DonationAmt").ClearContents
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
        For Each InnerNode In node.ChildNodes
            Sheet6.Cells(rowcount, NameColNo).value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet6.Cells(rowcount, AddressColNo).value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet6.Cells(rowcount, CityColNo).value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
             Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
            Sheet6.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node.SelectSingleNode("DoneePAN").text
            Sheet6.Cells(rowcount, AmountColNo).value = node.SelectSingleNode("DonationAmt").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub setDiffTblinfo_80G_B()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet6.Range("PerNO5080G.DoneeName").count
    Set rangecells = Sheet6.Range("PerNO5080G.DoneeName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'Konda updated 09-03-2026--V0.5
'    DefinedgridNameRange = "PerNO5080G.DoneeName||PerNO5080G.AddrDetail||PerNO5080G.CityOrTownOrDistrict||PerNO5080G.StateCode||PerNO5080G.PinCode||PerNO5080G.DoneePAN||PerNO5080G.DonationCash||PerNO5080G.DonationAmt||PerNO5080G.TotDonation||PerNO5080G.EligibleAmt"
    DefinedgridNameRange = "PerNO5080G.DoneeName||PerNO5080G.AddrDetail||PerNO5080G.CityOrTownOrDistrict||PerNO5080G.StateCode||PerNO5080G.PinCode||PerNO5080G.DoneePAN||PerNO5080G.DonationCash||PerNO5080G.DonationAmt||PerNO5080G.Traref||PerNO5080G.IFSC||PerNO5080G.TotDonation||PerNO5080G.EligibleAmt||Data_80G_B||comb_80G_B||Comb_donation_80G_B||Comb_PAN_80G_B"
'=================
End Sub

Sub AddDiffRows_80G_B(DiffRows As Long)
    setDiffTblinfo_80G_B
    Sheet6.Activate
    searchLastRow ("PerNO5080G.DoneeName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub XMLImport_80G_B()
On Error Resume Next
    Dim XpathOf80G_B As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo As Variant
    
    XpathOf80G_B = Common_year & "/Schedule80G/Don50PercentNoApprReqd/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_B)
    
    NameColNo = Sheet6.Range("PerNO5080G.DoneeName").Column
    AddressColNo = Sheet6.Range("PerNO5080G.AddrDetail").Column
    CityColNo = Sheet6.Range("PerNO5080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("PerNO5080G.StateCode").Column
    PincodeColNo = Sheet6.Range("PerNO5080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("PerNO5080G.DoneePAN").Column
    AmountColNo = Sheet6.Range("PerNO5080G.DonationAmt").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    TotalExRow = Range("PerNO5080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet6.Range("PerNO5080G.DoneeName").ClearContents
        Sheet6.Range("PerNO5080G.AddrDetail").ClearContents
        Sheet6.Range("PerNO5080G.CityOrTownOrDistrict").ClearContents
        Sheet6.Range("PerNO5080G.StateCode").ClearContents
        Sheet6.Range("PerNO5080G.PinCode").ClearContents
        Sheet6.Range("PerNO5080G.DoneePAN").ClearContents
        Sheet6.Range("PerNO5080G.DonationAmt").ClearContents
        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_B (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet6.Range("PerNO5080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet6.Cells(rowcount, NameColNo).value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet6.Cells(rowcount, AddressColNo).value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet6.Cells(rowcount, CityColNo).value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
            Sheet6.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node.SelectSingleNode("DoneePAN").text
            Sheet6.Cells(rowcount, AmountColNo).value = node.SelectSingleNode("DonationAmt").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub setDiffTblinfo_80G_C()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet6.Range("PerYES10080G.DoneeName").count
    Set rangecells = Sheet6.Range("PerYES10080G.DoneeName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'Konda updated 09-03-2026--V0.5
'    DefinedgridNameRange = "PerYES10080G.DoneeName||PerYES10080G.AddrDetail||PerYES10080G.CityOrTownOrDistrict||PerYES10080G.StateCode||PerYES10080G.PinCode||PerYES10080G.DoneePAN||PerYES10080G.DonationCash||PerYES10080G.DonationAmt||PerYES10080G.TotDonation||PerYES10080G.EligibleAmt"
    DefinedgridNameRange = "PerYES10080G.DoneeName||PerYES10080G.AddrDetail||PerYES10080G.CityOrTownOrDistrict||PerYES10080G.StateCode||PerYES10080G.PinCode||PerYES10080G.DoneePAN||PerYES10080G.DonationCash||PerYES10080G.DonationAmt||PerYES10080G.Traref||PerYES10080G.IFSC||PerYES10080G.TotDonation||PerYES10080G.EligibleAmt||Data_80G_C||comb_80G_C||Comb_donation_80G_C||Comb_PAN_80G_C"
'==========================
End Sub

Sub AddDiffRows_80G_C(DiffRows As Long)
    setDiffTblinfo_80G_C
    Sheet6.Activate
    searchLastRow ("PerYES10080G.DoneeName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub XMLImport_80G_C()
On Error Resume Next

    Dim XpathOf80G_C As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo As Variant
    
    XpathOf80G_C = Common_year & "/Schedule80G/Don100PercentApprReqd/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_C)
    
    NameColNo = Sheet6.Range("PerYES10080G.DoneeName").Column
    AddressColNo = Sheet6.Range("PerYES10080G.AddrDetail").Column
    CityColNo = Sheet6.Range("PerYES10080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("PerYES10080G.StateCode").Column
    PincodeColNo = Sheet6.Range("PerYES10080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("PerYES10080G.DoneePAN").Column
    AmountColNo = Sheet6.Range("PerYES10080G.DonationAmt").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    TotalExRow = Range("PerYES10080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet6.Range("PerYES10080G.DoneeName").ClearContents
        Sheet6.Range("PerYES10080G.AddrDetail").ClearContents
        Sheet6.Range("PerYES10080G.CityOrTownOrDistrict").ClearContents
        Sheet6.Range("PerYES10080G.StateCode").ClearContents
        Sheet6.Range("PerYES10080G.PinCode").ClearContents
        Sheet6.Range("PerYES10080G.DoneePAN").ClearContents
        Sheet6.Range("PerYES10080G.DonationAmt").ClearContents
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
        For Each InnerNode In node.ChildNodes
            Sheet6.Cells(rowcount, NameColNo).value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet6.Cells(rowcount, AddressColNo).value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet6.Cells(rowcount, CityColNo).value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
            Sheet6.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node.SelectSingleNode("DoneePAN").text
            Sheet6.Cells(rowcount, AmountColNo).value = node.SelectSingleNode("DonationAmt").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub setDiffTblinfo_80G_D()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet6.Range("Per5080G.DoneeName").count
    Set rangecells = Sheet6.Range("Per5080G.DoneeName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'Konda updated 09-03-2026--V0.5
'    DefinedgridNameRange = "Per5080G.DoneeName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DonationCash||Per5080G.DonationAmt||Per5080G.TotDonation||Per5080G.EligibleAmt||Per5080G.DonationCashTemp"
    DefinedgridNameRange = "Per5080G.DoneeName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DonationCash||Per5080G.DonationAmt||Per5080G.Traref||Per5080G.IFSC||Per5080G.TotDonation||Per5080G.EligibleAmt||Per5080G.DonationCashTemp||Data_80G_D||comb_80G_D||Comb_donation_80G_D||Comb_PAN_80G_D"
'======================
End Sub

Sub AddDiffRows_80G_D(DiffRows As Long)
    setDiffTblinfo_80G_D
    Sheet6.Activate
    searchLastRow ("Per5080G.DoneeName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub XMLImport_80G_D()
On Error Resume Next
    Dim XpathOf80G_D As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo As Variant
    
    XpathOf80G_D = Common_year & "/Schedule80G/Don50PercentApprReqd/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_D)
    
    NameColNo = Sheet6.Range("Per5080G.DoneeName").Column
    AddressColNo = Sheet6.Range("Per5080G.AddrDetail").Column
    CityColNo = Sheet6.Range("Per5080G.CityOrTownOrDistrict").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    StateCodeColNo = Sheet6.Range("Per5080G.StateCode").Column
    PincodeColNo = Sheet6.Range("Per5080G.PinCode").Column
    PanofDoneeColNo = Sheet6.Range("Per5080G.DoneePAN").Column
    AmountColNo = Sheet6.Range("Per5080G.DonationAmt").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    TotalExRow = Range("Per5080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet6.Range("Per5080G.DoneeName").ClearContents
        Sheet6.Range("Per5080G.AddrDetail").ClearContents
        Sheet6.Range("Per5080G.CityOrTownOrDistrict").ClearContents
        Sheet6.Range("Per5080G.StateCode").ClearContents
        Sheet6.Range("Per5080G.PinCode").ClearContents
        Sheet6.Range("Per5080G.DoneePAN").ClearContents
        Sheet6.Range("Per5080G.DonationAmt").ClearContents
        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_D (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet6.Range("Per5080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet6.Cells(rowcount, NameColNo).value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet6.Cells(rowcount, AddressColNo).value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet6.Cells(rowcount, CityColNo).value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet6.Cells(rowcount, StateCodeColNo).value = Findtext(iState, "StateWithoutForeign")
            Sheet6.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet6.Cells(rowcount, PanofDoneeColNo).value = node.SelectSingleNode("DoneePAN").text
            Sheet6.Cells(rowcount, AmountColNo).value = node.SelectSingleNode("DonationAmt").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub SalaryXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
   
    
    XpathOfPI = "/ns4:ITR/ns2:ITR4/IncomeDeductions" '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
'        For Each InnerNode In Node.ChildNodes
           
            Sheet1.Range("IncD.IncomeFromSal_i").value = UCase(node.SelectSingleNode("Salary").text)
            Sheet1.Range("IncD.IncomeFromOS").value = UCase(node.SelectSingleNode("IncomeOthSrc").text)
            
            Sheet1.Unprotect Password:=getmsgstate
            'Sheet1.Range("IncD.IncomeFromHP").value = UCase(Node.SelectSingleNode("TotalIncomeOfHP").Text)
            Sheet1.Protect Password:=getmsgstate
            If UCase(node.SelectSingleNode("TypeOfHP").text) = "L" Then
           Sheet1.Range("IncD.TypeOfHP").value = "Let Out"
            End If
            
            If UCase(node.SelectSingleNode("TypeOfHP").text) = "D" Then
           Sheet1.Range("IncD.TypeOfHP").value = "Deemed let out"
            End If
            
            If UCase(node.SelectSingleNode("TypeOfHP").text) = "S" Then
           Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied"
            End If
            
            

'        Next InnerNode
    Next node
End Sub

Sub TaxCompXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
   
    
    XpathOfPI = "/ITR/ns2:ITR4/TaxComputation" '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes
           
            Sheet1.Range("IncD.Section89").value = UCase(node.SelectSingleNode("Section89").text)
            
            
            

        Next InnerNode
    Next node
End Sub
Sub verpanXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
   
    
    XpathOfPI = "/ITR/ns2:ITR4/Verification/Declaration" '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)


    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes
           
            Sheet5.Range("Ver.PAN").value = UCase(node.SelectSingleNode("AssesseeVerPAN").text)
            
            
            

        Next InnerNode
    Next node
End Sub





Sub ITXMLImport()
On Error Resume Next
    Dim XpathOfIT As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Long
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfIT = "/ITR/ns2:ITR4/ScheduleIT/TaxPayment" '"/ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfIT)

    BSRCodeColNo = Sheet18.Range("TaxP.BSRCode").Column
    DateDepColNo = Sheet18.Range("TaxP.DateDep").Column
    SrlNoChallanColNo = Sheet18.Range("TaxP.SrlNoOfChaln").Column
    AmtColNo = Sheet18.Range("TaxP.Amt").Column
    
    TotalExRow = Range("TaxP.BSRCode").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet18.Range("TaxP.BSRCode").ClearContents
        Sheet18.Range("TaxP.DateDep").ClearContents
        Sheet18.Range("TaxP.SrlNoOfChaln").ClearContents
        Sheet18.Range("TaxP.Amt").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IT (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet18.Range("TaxP.BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet18.Cells(rowcount, BSRCodeColNo).value = UCase(node.SelectSingleNode("BSRCode").text)
            strDate = node.SelectSingleNode("DateDep").text
            YYYY = Mid(strDate, 1, 4)
            MM = Mid(strDate, 6, 2)
            DD = Mid(strDate, 9, 2)
            strDate = DD & "/" & MM & "/" & YYYY
            
            Sheet18.Cells(rowcount, DateDepColNo).value = strDate
            Sheet18.Cells(rowcount, SrlNoChallanColNo).value = node.SelectSingleNode("SrlNoOfChaln").text
            Sheet18.Cells(rowcount, AmtColNo).value = node.SelectSingleNode("Amt").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub




Sub TDSonSalaryXMLImport()
On Error Resume Next
    Dim XpathOfTDS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfTDS = "/ITR/ns2:ITR4/TDSonSalaries/TDSonSalary" '"/ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfTDS)

    TANNoEmployer = Sheet4.Range("TDSal.TAN").Column
    TDSNameOfEmployer = Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet4.Range("TDSal.IncChrgSal").Column
    TDSTotalTax = Sheet4.Range("TDSal.TotalTDSSal").Column
    
    TotalExRow = Range("TDSal.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("TDSal.TAN").ClearContents
        Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet4.Range("TDSal.IncChrgSal").ClearContents
        Sheet4.Range("TDSal.TotalTDSSal").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet4.Range("TDSal.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet4.Cells(rowcount, TANNoEmployer).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet4.Cells(rowcount, TDSNameOfEmployer).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet4.Cells(rowcount, TDSIncomeCharge).value = UCase(node.SelectSingleNode("IncChrgSal").text)
            Sheet4.Cells(rowcount, TDSTotalTax).value = UCase(node.SelectSingleNode("TotalTDSSal").text)
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub

Sub TDSOthXMLImport()
On Error Resume Next
    Dim XpathOfTDS2 As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim TDSCreditColNo, TANColNo, DEDNameColNo, DYAmountPaid, UTNColNo, FYColNo, TaxColNo, ClaimColNo, BroughtFwdTDSAmt, TotTDSOnAmtPaid As Variant

    XpathOfTDS2 = "/ITR/ns2:ITR4/TDSonOthThanSals/TDSonOthThanSal" '"/ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOfTDS2)

    TDSCreditColNo = Sheet4.Range("TDS2.TdsCredit").Column
    TANColNo = Sheet4.Range("TDSoth.TAN").Column
    DEDNameColNo = Sheet4.Range("TDSoth.EmployerOrDeductorOrCollecterName").Column
    FYColNo = Sheet4.Range("TDSoth.DeductedYear").Column
    UTNColNo = Sheet4.Range("TDSoth2.TD_OwnHands").Column
    DYAmountPaid = Sheet4.Range("TDSoth.AmountPaid").Column
    BroughtFwdTDSAmt = Sheet4.Range("TDSoth.TD_OwnHands").Column
    TaxColNo = Sheet4.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Column
    'ClaimColNo = Sheet4.Range("TDSoth.AmtClaimedBySpouse").Column
    TotTDSOnAmtPaid = Sheet4.Range("TDSoth.Sum").Column
    TotalExRow = Range("TDSoth.TAN").Rows.count

    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet4.Range("TDS2.TdsCredit").ClearContents
        Sheet4.Range("TDSoth.TAN").ClearContents
        Sheet4.Range("TDSoth.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet4.Range("TDSoth2.TD_OwnHands").ClearContents
        Sheet4.Range("TDSoth.DeductedYear").ClearContents
        Sheet4.Range("TDSoth.AmountPaid").ClearContents
        Sheet4.Range("TDSoth.TD_OwnHands").ClearContents
        Sheet4.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").ClearContents
        'Sheet4.Range("TDSoth.AmtClaimedBySpouse").ClearContents
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet4.Range("TDSoth.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        Dim credit As Variant
        credit = UCase(node.SelectSingleNode("TDSCreditName").text)
        If credit = "S" Then
        credit = "Self"
       End If
            Sheet4.Cells(rowcount, TDSCreditColNo).value = credit
            
            
            Sheet4.Cells(rowcount, TANColNo).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet4.Cells(rowcount, DEDNameColNo).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet4.Cells(rowcount, UTNColNo).value = UCase(node.SelectSingleNode("TaxDeducted26AS").text)
            Sheet4.Cells(rowcount, FYColNo).value = UCase(node.SelectSingleNode("DeductedYr").text)
            Sheet4.Cells(rowcount, DYAmountPaid).value = UCase(node.SelectSingleNode("RcptDtls26AS").text)
            Sheet4.Cells(rowcount, BroughtFwdTDSAmt).value = UCase(node.SelectSingleNode("TaxDeductCreditDtls/TaxDeductedOwnHands").text)
            Sheet4.Cells(rowcount, TaxColNo).value = node.SelectSingleNode("TaxDeductCreditDtls/TaxCreditedOwnHands").text
            
            'Sheet4.Cells(rowCount, ClaimColNo).value = Node.SelectSingleNode("AmtClaimedBySpouse").Text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub




Sub TDSOthXMLImport1()
On Error Resume Next
    Dim XpathOfTDS21 As String
    Dim TotalXMLRow1 As Long
    Dim TotalDiffRow1 As Long
    Dim RecTDS11, RecTDS21, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow1 As Long
    Dim TDSCreditColNo1, TANColNo1, DEDNameColNo1, DYAmountPaid1, UTNColNo1, FYColNo1, BroughtFwdTDSAmt, TaxColNo, receipt26 As Variant

    XpathOfTDS21 = "/ITR/ns2:ITR4/TDSDtls26QC/TDSDetails26QC" '"/ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOfTDS21)
    TDSCreditColNo1 = Sheet4.Range("TDS3.TdsCredit").Column
    TANColNo1 = Sheet4.Range("TDSoth2.PAN").Column
    DEDNameColNo1 = Sheet4.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Column
    FYColNo1 = Sheet4.Range("TDSoth2.DeductedYear").Column
    'UTNColNo1 = Sheet4.Range("TDSoth2.UTN").Column
    'DYAmountPaid1 = Sheet4.Range("TDS3.TaxDed").Column
    BroughtFwdTDSAmt = Sheet4.Range("TDSoth2.TD_OwnHands").Column
    TaxColNo = Sheet4.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Column
    'ClaimColNo = Sheet4.Range("TDSoth.AmtClaimedBySpouse").Column
    receipt26 = Sheet4.Range("TDSoth2.AmountSubjectDed").Column
    TotalExRow1 = Range("TDSoth2.PAN").Rows.count

    TotalXMLRow1 = Nodelist.Length
    TotalDiffRow1 = TotalXMLRow1 - TotalExRow1

    If (TotalXMLRow1 > 0) Then
        Sheet4.Range("TDS3.TdsCredit").ClearContents
        Sheet4.Range("TDSoth2.PAN").ClearContents
        Sheet4.Range("TDSoth2.EmployerOrDeductorOrCollecterName").ClearContents
        'Sheet4.Range("TDSoth2.UTN").ClearContents
        Sheet4.Range("TDSoth2.DeductedYear").ClearContents
        Sheet4.Range("TDSoth2.TD_OwnHands").ClearContents
        Sheet4.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").ClearContents
        Sheet4.Range("TDSoth2.AmountSubjectDed").ClearContents
        'Sheet4.Range("TDSoth.AmtClaimedBySpouse").ClearContents
    End If

    If (TotalDiffRow1 > 0) Then
     AddDiffRows_TDSoth1 (TotalDiffRow1)
    End If


    rowcount = getRowNo(Sheet4.Range("TDSoth2.PAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        Dim credit As Variant
        credit = UCase(node.SelectSingleNode("TDSCreditName").text)
        If credit = "S" Then
        credit = "Self"
       End If
            Sheet4.Cells(rowcount, TDSCreditColNo1).value = credit
            
            Sheet4.Cells(rowcount, TANColNo1).value = UCase(node.SelectSingleNode("PANofTenant").text)
            Sheet4.Cells(rowcount, DEDNameColNo1).value = UCase(node.SelectSingleNode("NameOfTenant").text)
            'Sheet4.Cells(rowCount, UTNColNo1).value = UCase(Node.SelectSingleNode("RcptDtls26AS").Text)
            Sheet4.Cells(rowcount, FYColNo1).value = UCase(node.SelectSingleNode("DeductedYr").text)
            'Sheet4.Cells(rowCount, DYAmountPaid1).value = UCase(Node.SelectSingleNode("TaxDeducted26AS").Text)
            Sheet4.Cells(rowcount, BroughtFwdTDSAmt).value = UCase(node.SelectSingleNode("TaxDeductCreditDtls/TaxDeductedOwnHands").text)
            Sheet4.Cells(rowcount, TaxColNo).value = node.SelectSingleNode("TaxDeductCreditDtls/TaxCreditedOwnHands").text
            Sheet4.Cells(rowcount, receipt26).value = node.SelectSingleNode("RcptDtls26AS").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS11 = cnt
End Sub


Sub TCSXMLImport()
On Error Resume Next
    Dim XpathOfTCS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTCS As Long
    Dim TANColNo, DEDNameColNo, TaxColNo, ClaimColNo, AmntClaimedBySpouseTCS, Amtfrom26AS, CollectedYr As Variant
    Dim TotalExRow, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    XpathOfTCS = "/ITR/ns2:ITR4/ScheduleTCS/TCS" '"/ITR/ScheduleTCS/TCS"
    Set Nodelist = dom.SelectNodes(XpathOfTCS)
    
    TANColNo = Sheet17.Range("TCS.TAN").Column
    DEDNameColNo = Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    Amtfrom26AS = Sheet17.Range("TCS.AmountPaid").Column
    TaxColNo = Sheet17.Range("TCS.TotalTCS").Column
    ClaimColNo = Sheet17.Range("TCS.AmtTCSClaimedThisYear").Column
'    AmntClaimedBySpouseTCS = Sheet17.Range("TCS.AmtClaimedBySpouse").Column
    
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet17.Range("TCS.TAN").ClearContents
    Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
    Sheet17.Range("TCS.AmountPaid").ClearContents
    Sheet17.Range("TCS.TotalTCS").ClearContents
    Sheet17.Range("TCS.AmtTCSClaimedThisYear").ClearContents
'    Sheet17.Range("TCS.AmtClaimedBySpouse").ClearContents
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TCS (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet17.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet17.Cells(rowcount, TANColNo).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet17.Cells(rowcount, DEDNameColNo).value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            ''Sheet4.Cells(rowCount, CollectedYr).value = Node.SelectSingleNode("CollectedYr").Text
            Sheet17.Cells(rowcount, Amtfrom26AS).value = node.SelectSingleNode("Amtfrom26AS").text
            Sheet17.Cells(rowcount, TaxColNo).value = node.SelectSingleNode("TotalTCS").text
            Sheet17.Cells(rowcount, ClaimColNo).value = node.SelectSingleNode("AmtTCSClaimedThisYear").text
            Sheet17.Cells(rowcount, AmntClaimedBySpouseTCS).value = node.SelectSingleNode("AmtClaimedBySpouse").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTCS = cnt
End Sub



Sub FilingInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
     Dim Filingtype As Variant
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A, ReturnFileSec As Variant
    Dim DateofOriginalfile As Variant
    Dim NoticeDateussec As Variant
    Dim residentialStatus
    Dim TaxStatus
    Dim TaxRefundable
    Dim iProvisoFlag, sProvisoFlag As Variant
    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
    
    XpathOfPI = Common_year & "/FilingStatus"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes
        
        
        iReturnFile = UCase(node.SelectSingleNode("ReturnFileSec").text)
            
    
        If iReturnFile = "11" Then
           sReturnFile = "139(1)-On or before due date"
        ElseIf iReturnFile = "12" Then
           sReturnFile = "139(4)-After due date"
        ElseIf iReturnFile = "13" Then
           sReturnFile = "142(1)"
        ElseIf iReturnFile = "14" Then
           sReturnFile = "148"
        ElseIf iReturnFile = "15" Then
           sReturnFile = "153A"
        ElseIf iReturnFile = "16" Then
           sReturnFile = "153C"
        ElseIf iReturnFile = "17" Then
           sReturnFile = "139(5)-Revised Return"
        ElseIf iReturnFile = "18" Then
           sReturnFile = "139(9)"
        ElseIf iReturnFile = "20" Then
        'J200924
          ' sReturnFile = "119(2)(b)- after condonation of delay"
           sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
        End If
            
        Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
        
   
        iProvisoFlag = UCase(node.SelectSingleNode("SeventhProvisio139").text)
        If iProvisoFlag = "Y" Then
        sProvisoFlag = "Yes"
        ElseIf iProvisoFlag = "N" Then
        sProvisoFlag = "No"
        End If
            
        Sheet1.Range("sheet1.SeventhProvisoFlag").value = sProvisoFlag
        
        iDepositAmountFlag = UCase(node.SelectSingleNode("DepAmtAggAmtExcd1CrPrYrFlg").text)
        If iDepositAmountFlag = "Y" Then
        sDepositAmountFlag = "Yes"
        ElseIf iDepositAmountFlag = "N" Then
        sDepositAmountFlag = "No"
        End If
            
        Sheet1.Range("Sheet1.DepositAmountFlag").value = sDepositAmountFlag
        Sheet1.Range("Sheet1.DepositAmount").value = UCase(node.SelectSingleNode("AmtSeventhProvisio139i").text)
        
        iAggrigateAmountFlag = UCase(node.SelectSingleNode("IncrExpAggAmt2LkTrvFrgnCntryFlg").text)
        If iAggrigateAmountFlag = "Y" Then
        sAggrigateAmountFlag = "Yes"
        ElseIf iAggrigateAmountFlag = "N" Then
        sAggrigateAmountFlag = "No"
        End If
            
        Sheet1.Range("Sheet1.AggrigateAmountFlag").value = sAggrigateAmountFlag
        Sheet1.Range("Sheet1.AggrigateAmount").value = UCase(node.SelectSingleNode("AmtSeventhProvisio139ii").text)
        
        iAggrigateAmountFlag1 = UCase(node.SelectSingleNode("IncrExpAggAmt1LkElctrctyPrYrFlg").text)
        If iAggrigateAmountFlag1 = "Y" Then
        sAggrigateAmountFlag1 = "Yes"
        ElseIf iAggrigateAmountFlag1 = "N" Then
        sAggrigateAmountFlag1 = "No"
        End If
            
        Sheet1.Range("Sheet1.AggrigateAmountFlag1").value = sAggrigateAmountFlag
        Sheet1.Range("Sheet1.AggrigateAmount1").value = UCase(node.SelectSingleNode("AmtSeventhProvisio139iii").text)
       
        
        
   
   
    ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
            
    ReturnFileSec = Mid(ReturnFileSec, 1, 2)
        If ReturnFileSec = "17" Then
Sheet1.Range("sheet1.ReceiptNo").value = node.SelectSingleNode("ReceiptNo").text

DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
        
        Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
     Else
        Sheet1.Range("Sheet1.ReceiptNo").value = node.SelectSingleNode("ReceiptNo").text
        
        DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
        
        Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
        
        Sheet1.Range("sheet1.NoticeNo").value = node.SelectSingleNode("NoticeNo").text
        
        NoticeDateussec = node.SelectSingleNode("NoticeDateUnderSec").text
        
        Sheet1.Range("sheet1.NoticeDate").value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
      End If
        
    
        
'             Sheet1.Range("sheet1.DesigOfficerWardorCircle").value = UCase(Node.SelectSingleNode("DesigOfficerWardorCircle").Text)
                        
            iPort5A = UCase(node.SelectSingleNode("PortugeseCC5A").text)
            If iPort5A = "Y" Then
            sPort5A = "Yes"
'            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A
            End If
            
            If iPort5A = "N" Then
            sPort5A = "No"
'            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A
            End If
'            sPort5A = IIf(iPort5A = "Y", "Yes", IIf(iPort5A = "N", "No", iPort5A))
'            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A
            
             Dim RepAssessee As Variant
            Dim sRepAssessee As Variant
            
            RepAssessee = UCase(node.SelectSingleNode("AsseseeRepFlg").text)
            If RepAssessee = "Y" Then
            sRepAssessee = "Yes"
            
            ElseIf RepAssessee = "N" Then
             sRepAssessee = "No"
            End If
            
            Sheet1.Range("Sheet1.RepAssessee").value = sRepAssessee
            
            Sheet1.Range("sheet1.NameRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepName").text)
          
          Dim iCapacity As Variant
          Dim sCapacity As Variant
          
          iCapacity = UCase(node.SelectSingleNode("AssesseeRep/RepCapacity").text)
           
           If iCapacity = "L" Then
           sCapacity = "Legal Heir"
           
           ElseIf iCapacity = "M" Then
           sCapacity = "Manager"
           
           ElseIf iCapacity = "G" Then
           sCapacity = "Guardian"
           
           ElseIf iCapacity = "O" Then
           sCapacity = "Other"
           
           End If
           
           Sheet1.Range("sheet1.CapacityRep").value = sCapacity
          
          Sheet1.Range("sheet1.AddressRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepAddress").text)
         
          Sheet1.Range("sheet1.PANRepAssessee").value = UCase(node.SelectSingleNode("AssesseeRep/RepPAN").text)

                                    
            Sheet1.Range("sheet1.PANOFSPOUSE").value = UCase(node.SelectSingleNode("PANOfSpouse").text)
        Next InnerNode
    Next node
End Sub


Sub IncomeInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sTypeOfHP, iTypeOfHP As Variant
        
    XpathOfPI = "/ITR/ns2:ITR4/ITR4S_IncomeDeductions"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

            iTypeOfHP = UCase(node.SelectSingleNode("TypeOfHP").text)
            sTypeOfHP = IIf(iTypeOfHP = "S", "Self Occupied", IIf(iTypeOfHP = "L", "Let Out", IIf(iTypeOfHP = "D", "Deemed Let Out", iTypeOfHP)))
            Sheet1.Range("IncD.TypeOfHP").value = sTypeOfHP
            
            'Sheet1.Range("IncD.IncomeFromHP").value = UCase(Node.SelectSingleNode("TotalIncomeOfHP").Text)
            Sheet1.Range("IncD.IncomeFromOS").value = UCase(node.SelectSingleNode("IncomeOthSrc").text)
        Next InnerNode
    Next node
End Sub

Sub ALInfoXMLImport()
On Error Resume Next
    Dim XpathOfAL As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim DescriptionColNo, FlatColNo, PremisesColNo, RoadColNo, AreaColNo, TownColNo, StateColNo, CountryColNo, PincodeColNo, ZipcodeColNo, AmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, cnt, TotalDiffRow, RecAL As Long
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim iCountry, sCountry, iState, sState As Variant
        
        
    XpathOfAL = "/ITR/ns2:ITR4/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfAL)

     
    DescriptionColNo = Sheet13.Range("SchAL.A.Description").Column
    FlatColNo = Sheet13.Range("SchAL.A.Address_Flat").Column
    PremisesColNo = Sheet13.Range("SchAL.A.Address_Village").Column
    RoadColNo = Sheet13.Range("SchAL.A.Address_Road").Column
    AreaColNo = Sheet13.Range("SchAL.A.Address_Area").Column
    TownColNo = Sheet13.Range("SchAL.A.Address_City").Column
    StateColNo = Sheet13.Range("SchAL.A.Address_State").Column
    CountryColNo = Sheet13.Range("SchAL.A.Address_Country").Column
    PincodeColNo = Sheet13.Range("SchAL.A.Address_Pin").Column
    ZipcodeColNo = Sheet13.Range("SchAL.A.Address_Zip").Column
    AmountColNo = Sheet13.Range("SchAL.A.Amount").Column
    
    TotalExRow = Range("SchAL.A.Description").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet13.Range("SchAL.A.Description").ClearContents
    Sheet13.Range("SchAL.A.Address_Flat").ClearContents
    Sheet13.Range("SchAL.A.Address_Village").ClearContents
    Sheet13.Range("SchAL.A.Address_Road").ClearContents
    Sheet13.Range("SchAL.A.Address_Area").ClearContents
    Sheet13.Range("SchAL.A.Address_City").ClearContents
    Sheet13.Range("SchAL.A.Address_State").ClearContents
    Sheet13.Range("SchAL.A.Address_Country").ClearContents
    Sheet13.Range("SchAL.A.Address_Pin").ClearContents
    Sheet13.Range("SchAL.A.Address_Zip").ClearContents
    Sheet13.Range("SchAL.A.Amount").ClearContents
    
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_AL1 (TotalDiffRow)
    End If
    rowcount = 0
    rowcount = getRowNo(Sheet13.Range("SchAL.A.Description").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
    rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        
'        Dim Str As String
'        Str = UCase(Node.SelectSingleNode("ImmovableFlag").Text)
'           If Str = "Y" Then
'            Sheet13.Range("IsImmovableAsset").value = "Yes"
'            End If
            
            
            
            Sheet13.Cells(rowcount, DescriptionColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/Description").text)
            Sheet13.Cells(rowcount, FlatColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/ResidenceNo").text)
            Sheet13.Cells(rowcount, PremisesColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/ResidenceName").text)
            Sheet13.Cells(rowcount, RoadColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/RoadOrStreet").text)
            Sheet13.Cells(rowcount, AreaColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/LocalityOrArea").text)
            Sheet13.Cells(rowcount, TownColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/CityOrTownOrDistrict").text)
            
             iCountry = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet13.Cells(rowcount, CountryColNo).value = sCountry
            
            iState = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet13.Cells(rowcount, StateColNo).value = sState
    
            
            
            
'            Sheet13.Cells(rowCount, StateColNo).value = UCase(Node.SelectSingleNode("ImmovableDetails/AddressAL/StateCode").Text)
'            Sheet13.Cells(rowCount, CountryColNo).value = UCase(Node.SelectSingleNode("ImmovableDetails/AddressAL/CountryCode").Text)
            
            
            
            Sheet13.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/AddressAL/PinCode").text)
            Sheet13.Cells(rowcount, AmountColNo).value = UCase(node.SelectSingleNode("ImmovableDetails/Amount").text)
            
'            Sheet13.Range("SchAL.B.Bank").value = UCase(Node.SelectSingleNode("MovableAsset/DepositsInBank").Text)
'            Sheet13.Range("SchAL.B.Share").value = UCase(Node.SelectSingleNode("MovableAsset/SharesAndSecurities").Text)
'            Sheet13.Range("SchAL.B.Insurance").value = UCase(Node.SelectSingleNode("MovableAsset/InsurancePolicies").Text)
'            Sheet13.Range("SchAL.B.Loans").value = UCase(Node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").Text)
'            Sheet13.Range("SchAL.B.Cash").value = UCase(Node.SelectSingleNode("MovableAsset/CashInHand").Text)
'            Sheet13.Range("SchAL.B.Jwellery").value = UCase(Node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").Text)
'            Sheet13.Range("SchAL.B.Art").value = UCase(Node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").Text)
'            Sheet13.Range("SchAL.B.Vehicles").value = UCase(Node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").Text)
'            'Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
'            Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("LiabilityInRelatAssets").Text)


            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecAL = cnt
    
End Sub

Sub ALInfoXMLImport2()
On Error Resume Next
    Dim XpathOfAL As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim NameOfFirmColNo, FlatColNo, PremisesColNo, RoadColNo, AreaColNo, TownColNo, StateColNo, CountryColNo, PANColNo, PincodeColNo, ZipcodeColNo, AmountColNo As Variant
    Dim TotalExRow, TotalXMLRow, cnt, TotalDiffRow, RecAL2 As Long
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim iCountry, sCountry, iState, sState As Variant
    
        
    XpathOfAL = "/ITR/ns2:ITR4/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfAL)

     NameOfFirmColNo = Sheet13.Range("SchAL.C.Name").Column
    FlatColNo = Sheet13.Range("SchAL.C.Address_Flat").Column
    PremisesColNo = Sheet13.Range("SchAL.C.Address_Village").Column
    RoadColNo = Sheet13.Range("SchAL.C.Address_Road").Column
    AreaColNo = Sheet13.Range("SchAL.C.Address_Area").Column
    TownColNo = Sheet13.Range("SchAL.C.Address_City").Column
    StateColNo = Sheet13.Range("SchAL.C.Address_State").Column
    CountryColNo = Sheet13.Range("SchAL.C.Address_Country").Column
    PincodeColNo = Sheet13.Range("SchAL.C.Address_Pin").Column
    ZipcodeColNo = Sheet13.Range("SchAL.C.Address_Zip").Column
    PANColNo = Sheet13.Range("SchAL.C.PAN").Column
    AmountColNo = Sheet13.Range("SchAL.C.Investment").Column
    
    TotalExRow = Range("SchAL.C.Name").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet13.Range("SchAL.C.Name").ClearContents
    Sheet13.Range("SchAL.C.Address_Flat").ClearContents
    Sheet13.Range("SchAL.C.Address_Village").ClearContents
    Sheet13.Range("SchAL.C.Address_Road").ClearContents
    Sheet13.Range("SchAL.C.Address_Area").ClearContents
    Sheet13.Range("SchAL.C.Address_City").ClearContents
    Sheet13.Range("SchAL.C.Address_Pin").ClearContents
    Sheet13.Range("SchAL.C.Address_Zip").ClearContents
    Sheet13.Range("SchAL.C.PAN").ClearContents
    Sheet13.Range("SchAL.C.Investment").ClearContents
    Sheet13.Range("SchAL.C.Address_State").ClearContents
    Sheet13.Range("SchAL.C.Address_Country").ClearContents
    
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_AL2 (TotalDiffRow)
    End If
    rowcount = 0
    rowcount = getRowNo(Sheet13.Range("SchAL.C.Name").name)
    rowcount = rowcount - 1
    cnt = 0
    
    
    
    
    
    For Each node In Nodelist
    rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        
'        Dim Str As String
'        Str = UCase(Node.SelectSingleNode("ImmovableFlag").Text)
'           If Str = "Y" Then
'            Sheet13.Range("IsInterestHeld").value = "Yes"
'            End If
            
            
            
            Sheet13.Cells(rowcount, NameOfFirmColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/NameOfFirm").text)
            Sheet13.Cells(rowcount, FlatColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/ResidenceNo").text)
            Sheet13.Cells(rowcount, PremisesColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/ResidenceName").text)
            Sheet13.Cells(rowcount, RoadColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/RoadOrStreet").text)
            Sheet13.Cells(rowcount, AreaColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/LocalityOrArea").text)
            Sheet13.Cells(rowcount, TownColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/CityOrTownOrDistrict").text)
            
             iCountry = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet13.Cells(rowcount, CountryColNo).value = sCountry
            
            iState = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet13.Cells(rowcount, StateColNo).value = sState
    
            
            
            
            
            
            'Sheet13.Cells(rowCount, StateColNo).value = UCase(Node.SelectSingleNode("ImmovableDetails/AddressAL/StateCode").Text)
            
            Sheet13.Cells(rowcount, PincodeColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AddressAL/PinCode").text)
            Sheet13.Cells(rowcount, PANColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/PanOfFirm").text)
            Sheet13.Cells(rowcount, AmountColNo).value = UCase(node.SelectSingleNode("InterestHeldInaAsset/AssesseInvestment").text)
            
'            Sheet13.Range("SchAL.B.Bank").value = UCase(Node.SelectSingleNode("MovableAsset/DepositsInBank").Text)
'            Sheet13.Range("SchAL.B.Share").value = UCase(Node.SelectSingleNode("MovableAsset/SharesAndSecurities").Text)
'            Sheet13.Range("SchAL.B.Insurance").value = UCase(Node.SelectSingleNode("MovableAsset/InsurancePolicies").Text)
'            Sheet13.Range("SchAL.B.Loans").value = UCase(Node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").Text)
'            Sheet13.Range("SchAL.B.Cash").value = UCase(Node.SelectSingleNode("MovableAsset/CashInHand").Text)
'            Sheet13.Range("SchAL.B.Jwellery").value = UCase(Node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").Text)
'            Sheet13.Range("SchAL.B.Art").value = UCase(Node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").Text)
'            Sheet13.Range("SchAL.B.Vehicles").value = UCase(Node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").Text)
'            'Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
'            Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("LiabilityInRelatAssets").Text)
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecAL2 = cnt
    
End Sub
Sub ImmovableXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfPI = "/ITR/ns2:ITR4/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

   
        'ReturnFile Sec/ Port 5A
        
            Dim str, Str1 As String
        str = UCase(node.SelectSingleNode("ImmovableFlag").text)
           If str = "Y" Then
            Sheet13.Range("IsImmovableAsset").value = "Yes"
            ElseIf str = "N" Then
            Sheet13.Range("IsImmovableAsset").value = "No"
            Else
            Sheet13.Range("IsImmovableAsset").value = "(Select)"
            End If
            
         Str1 = UCase(node.SelectSingleNode("InterstAOPFlag").text)
            If Str1 = "Y" Then
            Sheet13.Range("IsInterestHeld").value = "Yes"
            ElseIf Str1 = "N" Then
            Sheet13.Range("IsInterestHeld").value = "No"
            Else
            Sheet13.Range("IsInterestHeld").value = "(Select)"
            End If
                                                
        Next InnerNode
    Next node
End Sub


Sub ALveriXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfPI = "/ITR/ns2:ITR4/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

   
        'ReturnFile Sec/ Port 5A
        
            Sheet13.Range("SchAL.B.Bank").value = UCase(node.SelectSingleNode("MovableAsset/DepositsInBank").text)
            Sheet13.Range("SchAL.B.Share").value = UCase(node.SelectSingleNode("MovableAsset/SharesAndSecurities").text)
            Sheet13.Range("SchAL.B.Insurance").value = UCase(node.SelectSingleNode("MovableAsset/InsurancePolicies").text)
            Sheet13.Range("SchAL.B.Loans").value = UCase(node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").text)
            Sheet13.Range("SchAL.B.Cash").value = UCase(node.SelectSingleNode("MovableAsset/CashInHand").text)
            Sheet13.Range("SchAL.B.Jwellery").value = UCase(node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").text)
            Sheet13.Range("SchAL.B.Art").value = UCase(node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").text)
            Sheet13.Range("SchAL.B.Vehicles").value = UCase(node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").text)
            'Sheet13.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
            Sheet13.Range("SchAL.Liability").value = UCase(node.SelectSingleNode("LiabilityInRelatAssets").text)
             
           ' Sheet13.Range("IsImmovableAsset").value = UCase(Node.SelectSingleNode("ImmovableFlag").Text)
                                                
        Next InnerNode
    Next node
End Sub





Sub VeriInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    'Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfPI = "/ITR/ns2:ITR4/Verification"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

   
        'ReturnFile Sec/ Port 5A
        
            Sheet5.Range("Ver.FatherName").value = UCase(node.SelectSingleNode("Declaration/FatherName").text)
                                                
        Next InnerNode
    Next node
End Sub

Sub RefundInfoXMLImport()
On Error Resume Next
    Dim XpathOfRefund As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim IFSC, BANKNAME, ACCNO As Variant
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfRefund = "/ITR/ns2:ITR4/Refund/BankAccountDtls/AddtnlBankDetails"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)

    IFSC = Range("SchBA.IFSC").Column
    BANKNAME = Range("SchBA.BankName").Column
    ACCNO = Range("SchBA.AcntNo").Column
    
    TotalExRow = Range("SchBA.IFSC").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Range("SchBA.IFSC").ClearContents
        Range("SchBA.BankName").ClearContents
        Range("SchBA.AcntNo").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_BANK (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet5.Range("SchBA.IFSC").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
        

            
            Sheet5.Cells(rowcount, IFSC).value = UCase(node.SelectSingleNode("IFSCCode").text)
            Sheet5.Cells(rowcount, BANKNAME).value = UCase(node.SelectSingleNode("BankName").text)
            Sheet5.Cells(rowcount, ACCNO).value = UCase(node.SelectSingleNode("BankAccountNo").text)
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
    XpathOfRefund = "/ns4:ITR/ns2:ITR4/Refund/BankAccountDtls/PriBankDetails"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

   
            Sheet5.Range("Refund.IFSC").value = UCase(node.SelectSingleNode("IFSCCode").text)
            Sheet5.Range("Refund.BankName").value = UCase(node.SelectSingleNode("BankName").text)
            Sheet5.Range("Refund.AcntNo").value = UCase(node.SelectSingleNode("BankAccountNo").text)
        Next InnerNode
    Next node
End Sub

'Sub NOBXMLImport()
'On Error Resume Next
'    Dim XpathOfNOB As String
'    Dim TotalXMLRow As Long
'    Dim TotalDiffRow As Long
'    Dim RecTDS1 As Long
'    Dim NodeList As IXMLDOMNodeList
'    Dim Node As IXMLDOMNode
'    Dim InnerNode As IXMLDOMNode
'    Dim NOB As Variant
'    Dim strDate As String
'    Dim YYYY, MM, DD As String
'    Dim TotalExRow As Long
'    Dim sNOB, iNOB As Variant
'
'    XpathOfNOB = "/ITR/ns2:ITR4/NatOfBus/NatureOfBusiness" '"/ITR/TDSonSalaries"
'    Set NodeList = dom.SelectNodes(XpathOfNOB)
'
'    NOB = Sheet3.Range("NOB.Code").Column
'
'    TotalExRow = Range("NOB.Code").Rows.Count
'
'    TotalXMLRow = NodeList.length
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'        Sheet3.Range("NOB.Code").ClearContents
'
'    End If
'
'    rowCount = getRowNo(Sheet3.Range("NOB.Code").name)
'    rowCount = rowCount - 1
'    cnt = 0
'
'    For Each Node In NodeList
'        rowCount = rowCount + 1
'        For Each InnerNode In Node.ChildNodes
'            'Sheet3.Cells(rowCount, NOB).value = UCase(Node.SelectSingleNode("Code").Text)
'
'            iNOB = UCase(Node.SelectSingleNode("Code").Text)
'            sNOB = Findtext(iNOB, "NOB")
'            Sheet3.Cells(rowCount, NOB).value = sNOB
'
'        Next InnerNode
'        cnt = cnt + 1
'    Next Node
'    RecTDS1 = cnt
'End Sub



Sub AddDiffRows_BANK(DiffRows As Long)
    setTableInfo_BANK
    Sheet5.Activate
    searchLastRow ("SchBA.IFSC")
    insertRowUnderSectionWithFormula (DiffRows)
    LinkCheckBoxes
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TaxP)
End Sub

Sub setTableInfo_BANK()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.IFSC").count
    Set rangecells = Range("SchBA.IFSC").Cells
    For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
               ccount = ccount + 1
           End If
    Next
'     DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.CheckBox||tempXML" '||SchBA.DepositedCash"
'     DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.CheckBox||tempXML"
'Ayush_21
      DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.Type_of_Account||SchBA.CheckBox||tempXML"
End Sub

Sub AddDiffRows_TDSoth(DiffRows As Long)
    setDiffTblinfo_TDSoth
    Sheet4.Activate
    searchLastRow ("TDsOthr.TAN")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub AddDiffRows_TDSoth1(DiffRows As Long)
    setDiffTblinfo_TDS3oth
    Sheet4.Activate
    searchLastRow ("TDsOthr2.PAN")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub AddDiffRows_TCS(DiffRows As Long)
    setDiffTblinfo_TCS
    Sheet17.Activate
    searchLastRow ("TCS.TAN")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub AddDiffRows_AL1(DiffRows As Long)
    setDiffTblinfo_AL1
    Sheet13.Activate
    searchLastRow ("SchAL.A.Description")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub AddDiffRows_AL2(DiffRows As Long)
    setDiffTblinfo_AL2
    Sheet13.Activate
    searchLastRow ("SchAL.C.Name")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub AddDiffRows_AL3(DiffRows As Long)
    setDiffTblinfo_AL3
    Sheet13.Activate
    searchLastRow ("SchAL.A.Description")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_AL1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet13.Range("SchAL.A.Description").count
    Set rangecells = Sheet13.Range("SchAL.A.Description").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    mdCommon.DefinedgridNameRange = "SchAL.A.Description||SchAL.A.Address_Flat||SchAL.A.Address_Village||SchAL.A.Address_Road||SchAL.A.Address_Area||SchAL.A.Address_City||SchAL.A.Address_State||SchAL.A.Address_Country||SchAL.A.Address_Pin||SchAL.A.Address_Zip||SchAL.A.Amount"
End Sub
Sub setDiffTblinfo_AL2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet13.Range("SchAL.C.Name").count
    Set rangecells = Sheet13.Range("SchAL.C.Name").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    mdCommon.DefinedgridNameRange = "SchAL.C.Name||SchAL.C.Address_Flat||SchAL.C.Address_Village||SchAL.C.Address_Road||SchAL.C.Address_Area||SchAL.C.Address_City||SchAL.C.Address_State||SchAL.C.Address_Country||SchAL.C.Address_Pin||SchAL.C.Address_Zip||SchAL.C.PAN||SchAL.C.Investment"
End Sub
Sub setDiffTblinfo_AL3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet13.Range("SchAL.A.Description").count
    Set rangecells = Sheet13.Range("SchAL.A.Description").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    mdCommon.DefinedgridNameRange = "SchAL.A.Description||SchAL.A.Address_Flat||SchAL.A.Address_Village||SchAL.A.Address_Road||SchAL.A.Address_Area||SchAL.A.Address_City||SchAL.A.Address_State||SchAL.A.Address_Country||SchAL.A.Address_Pin||SchAL.A.Address_Zip||SchAL.A.Amount"
End Sub



Sub setDiffTblinfo_TDSoth()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("TDsOthr.TAN").count
    Set rangecells = Sheet4.Range("TDsOthr.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    '
    'AY_2024_25 OLD
    'mdCommon.DefinedgridNameRange = "TDsOthr.TAN||TDsOthr.finyear||TDsOthr.tds||TDsOthr.tdsdeducted||TDsOthr.tdsclaimed||TDsOthr.grossamount||TDsOthr.headincome||TDsOthr.tdscarried"
    'AY_2025_26
    mdCommon.DefinedgridNameRange = "TDsOthr.TAN||TDsOthr.SectionTDS||TDsOthr.finyear||TDsOthr.tds||TDsOthr.tdsdeducted||TDsOthr.tdsclaimed||TDsOthr.grossamount||TDsOthr.headincome||TDsOthr.tdscarried"
End Sub

Sub setDiffTblinfo_TDSoth1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("TDSoth2.PAN").count
    Set rangecells = Sheet4.Range("TDSoth2.PAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    mdCommon.DefinedgridNameRange = "TDS3.TdsCredit||TDSoth2.PAN||TDSoth2.EmployerOrDeductorOrCollecterName||TDSoth2.UTN||TDSoth2.DeductedYear||TDSoth2.AmountSubjectDed ||TDSoth2.TD_OwnHands ||TDSoth2.TD_SIncome||TDSoth2.TD_TDS||TDSoth2.TD_PAN||TDSoth2.ClaimOutOfTotTDSOnAmtPaid||TDSoth2.TDS_SIncome||TDSoth2.TDS_TDS||TDSoth2.TDS_PAN"
End Sub

Sub setDiffTblinfo_TDS3oth()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("TDsOthr2.PAN").count
    Set rangecells = Sheet4.Range("TDsOthr2.PAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS2ii_1 = ccount
    mdCommon.DefinedgridNameRange = "TDsOthr2.PAN||TDsOthr2.AADHAAR||TDsOthr2.SectionTDSDeducted||TDsOthr2.finyear||TDsOthr2.tds||TDsOthr2.tdsdeducted||TDsOthr2.tdsclaimed||TDsOthr2.grossamount||TDsOthr2.headamount||TDsOthr2.carried"
End Sub

Sub setDiffTblinfo_TCS()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet17.Range("TCS.TAN").count
    Set rangecells = Sheet17.Range("TCS.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TCS = ccount
    DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.AmountPaid||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
 End Sub

Sub setDiffTblinfo_TDS1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("TDSal.TAN").count
    Set rangecells = Sheet4.Range("TDSal.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    mdCommon.DefinedgridNameRange = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSal||TDSal.TotalTDSSal"
End Sub
 
Sub AddDiffRows_TDS1(DiffRows As Long)
    setDiffTblinfo_TDS1
    Sheet4.Activate
    searchLastRow ("TDSal.TAN")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub

Function OpenXMLFileDialog() As String
On Error GoTo endline
    Dim FileDialogBox As FileDialog
    Dim FileObject As Variant
    Dim filname As String
    Dim intChoice As Long
    
    intChoice = 0
    
    Set FileDialogBox = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialogBox
        .ButtonName = "Select"
        .AllowMultiSelect = False
        .Filters.add "XML Files", "*.xml", 1
        .Title = "Choose XML file to use for Filling TDS/TaxPayment Data in Utility"
        .InitialView = msoFileDialogViewDetails
                
        intChoice = .Show
        If intChoice <> 0 Then
            For Each FileObject In .SelectedItems
                filname = FileObject
                OpenXMLFileDialog = filname
            Next FileObject
        End If
        On Error GoTo 0
    End With
    
    If intChoice = 0 Then
        Application.EnableEvents = True
        Set FileDialogBox = Nothing
        EndProcessing
    End If
    
    Set FileDialogBox = Nothing
    
endline:

End Function

Sub EndProcessing()
    End
End Sub

Function getRowNo(RangeAddress As String) As Long
    Dim ColStart As Long
    Dim ColEnd As Long
    
    ColStart = InStr(1, RangeAddress, "$")
    ColEnd = InStr(ColStart + 1, RangeAddress, "$")
    ColStart = ColEnd
    ColEnd = InStr(ColStart + 1, RangeAddress, "$")
    
    getRowNo = Mid(Mid(RangeAddress, ColStart + 1, ColEnd - ColStart - 1), 1, Len(Mid(RangeAddress, ColStart + 1, ColEnd - ColStart - 1)) - 1)
End Function

Function ValidateImportfields()
    ValidateImportfields = True
    
    msgError = ""
    ValidateSheetIncomeDetails
    ValidateTDS_TCS_IT
    
End Function

Sub setDiffTblinfo_IT()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet18.Range("TaxP.BSRCode").count
    Set rangecells = Sheet18.Range("TaxP.BSRCode").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    '-----------------------------------
   ' mdCommon.DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1"
    mdCommon.DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT"
    '---------------------------
End Sub
 
Sub AddDiffRows_IT(DiffRows As Long)
    setDiffTblinfo_IT
    Sheet18.Activate
    searchLastRow ("TaxP.BSRCode")
    mdCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_Exempt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet5.Range("Sheet5.Nature").count
    Set rangecells = Sheet5.Range("Sheet5.Nature").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'Malli_AY_2026_27 17/06/2026
    'DefinedgridNameRange = "Sheet5.Nature||Sheet5.Description||Sheet5.Amount"
     DefinedgridNameRange = "Sheet5.Nature||Sheet5.Description||Sheet5.DescEI||Sheet5.Amount"
    '-----------------------------
End Sub

Sub AddDiffRows_Exempt(DiffRows As Long)
    setDiffTblinfo_Exempt
    Sheet5.Activate
    searchLastRow ("Sheet5.Nature")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_Gst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet3.Range("BP.GSTRno").count
    Set rangecells = Sheet3.Range("BP.GSTRno").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "BP.GSTRno||Sec44AD.NOBName3||Sec44AD.NOBCode3||BP.GSTRAmount"
End Sub

Sub AddDiffRows_Gst(DiffRows As Long)
    setDiffTblinfo_Gst
    Sheet3.Activate
    searchLastRow ("BP.GSTRno")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub


Sub ExemptXMLImport2()
On Error Resume Next
    Dim XpathOfExempt As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecExempt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NatureColNo, DescriptionColNo, AmtColNo As Long
    
    Dim TotalExRow As Long
    
    XpathOfExempt = Common_year & "/IncomeDeductions/AllwncExemptUs10/AllwncExemptUs10Dtls"  '"/ns2:ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfExempt)

    NatureColNo = Sheet1.Range("Sheet1.Nature").Column
    DescriptionColNo = Sheet1.Range("Sheet1.Description").Column
    AmtColNo = Sheet1.Range("Sheet1.Amount").Column
    
    TotalExRow = Range("Sheet1.Nature").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet1.Range("Sheet1.Nature").ClearContents
        Sheet1.Range("Sheet1.Description").ClearContents
        Sheet1.Range("Sheet1.Amount").ClearContents
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Exempt1XML (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet1.Range("Sheet1.Nature").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            
            Sheet1.Cells(rowcount, NatureColNo).value = Findtext("Sec " & (node.SelectSingleNode("SalNatureDesc").text), "Sheet1.NOA")
            
            
            If (node.SelectSingleNode("SalNatureDesc").text) = "10(10B)(i)" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
            End If
            
            
            If (node.SelectSingleNode("SalNatureDesc").text) = "10(10B)(ii)" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
            End If
            
            If UCase(node.SelectSingleNode("SalNatureDesc").text) = "OTH" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
            End If
            
            Sheet1.Cells(rowcount, DescriptionColNo).value = node.SelectSingleNode("SalOthNatOfInc").text
            Sheet1.Cells(rowcount, AmtColNo).value = node.SelectSingleNode("SalOthAmount").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecExempt = cnt
    
End Sub
Sub AddDiffRows_Exempt1XML(DiffRows As Long)
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
    end_TaxP = ccount
    DefinedgridNameRange = "Sheet1.Nature||Sheet1.Description||Sheet1.Amount"
 End Sub
Sub ExemptXMLImport3()
On Error Resume Next
    Dim XpathOfExempt3 As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim Exempt3 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NatureColNo, DescriptionColNo, AmtColNo As Long
    
    Dim TotalExRow As Long
    
    XpathOfExempt3 = Common_year & "/IncomeDeductions/OthersInc/OthersIncDtlsOthSrc"  '"/ns2:ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfExempt3)

    NatureColNo = Sheet1.Range("Sheet1.Nature1").Column
    DescriptionColNo = Sheet1.Range("Sheet1.Description1").Column
    AmtColNo = Sheet1.Range("Sheet1.Amount1").Column
    
    TotalExRow = Range("Sheet1.Nature1").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet1.Range("Sheet1.Nature1").ClearContents
        Sheet1.Range("Sheet1.Description1").ClearContents
        Sheet1.Range("Sheet1.Amount1").ClearContents
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Exempt2XML (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet1.Range("Sheet1.Nature1").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "SAV" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Savings Account"
            End If
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "IFD" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
            End If
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "TAX" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Interest from Income Tax Refund"
            End If
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "FAP" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Family pension"
            End If
            
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "OTH" Then
            Sheet1.Cells(rowcount, NatureColNo).value = "Any Other"
            End If
            
            
            Sheet1.Cells(rowcount, DescriptionColNo).value = node.SelectSingleNode("OthSrcOthNatOfInc").text
            Sheet1.Cells(rowcount, AmtColNo).value = node.SelectSingleNode("OthSrcOthAmount").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    Exempt3 = cnt
    
End Sub
Sub AddDiffRows_Exempt2XML(DiffRows As Long)
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
    end_TaxP = ccount
    DefinedgridNameRange = "Sheet1.Nature1||Sheet1.Description1||Sheet1.Amount1||Sheet1.Amount11||Sheet1.Amount22||Sheet1.Amount23||Sheet1.Amount24||Sheet1.Amount25"
 End Sub
Sub Sch80DXMLImport()
On Error Resume Next
    Dim XpathOf80D As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode

    XpathOf80D = Common_year & "/Schedule80D/Sec80DSelfFamSrCtznHealth"
    Set Nodelist = dom.SelectNodes(XpathOf80D)

    For Each node In Nodelist
'        For Each InnerNode In Node.ChildNodes
   
            Dim temp80 As Variant
            
            temp80 = UCase(node.SelectSingleNode("SeniorCitizenFlag").text)
            If temp80 = "Y" Then
            Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Yes"
            ElseIf temp80 = "N" Then
            Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "No"
            ElseIf temp80 = "S" Then
            Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
            End If
                                            
            
            Sheet16.Range("Health_Insurance_80D").value = UCase(node.SelectSingleNode("HealthInsPremSlfFam").text)
            Sheet16.Range("Preventive_Health_80D").value = UCase(node.SelectSingleNode("PrevHlthChckUpSlfFam").text)
            Sheet16.Range("Health_InsuranceSC_80D").value = UCase(node.SelectSingleNode("HlthInsPremSlfFamSrCtzn").text)
            Sheet16.Range("Preventive_Health_SC_80D").value = UCase(node.SelectSingleNode("PrevHlthChckUpSlfFamSrCtzn").text)
            Sheet16.Range("Medical_Expenditure_SC_80D").value = UCase(node.SelectSingleNode("MedicalExpSlfFamSrCtzn").text)
            
            Dim temp802 As Variant
            
            temp802 = UCase(node.SelectSingleNode("ParentsSeniorCitizenFlag").text)
            If temp802 = "Y" Then
            Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Yes"
            ElseIf temp802 = "N" Then
            Sheet16.Range("DropDown_ValueOf_SC_80D").value = "No"
            ElseIf temp802 = "P" Then
            Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
            End If
            
            
    
            Sheet16.Range("Health_Insurance2_80D").value = UCase(node.SelectSingleNode("HlthInsPremParents").text)
            Sheet16.Range("Preventive_Health2_80D").value = UCase(node.SelectSingleNode("PrevHlthChckUpParents").text)
            Sheet16.Range("Health_Insurance3_80D").value = UCase(node.SelectSingleNode("HlthInsPremParentsSrCtzn").text)
            Sheet16.Range("Preventive_Health3_80D").value = UCase(node.SelectSingleNode("PrevHlthChckUpParentsSrCtzn").text)
            Sheet16.Range("Medical_Expenditure2_80D").value = UCase(node.SelectSingleNode("MedicalExpParentsSrCtzn").text)
                                             
            
                                            
'        Next InnerNode
    Next node
    
End Sub


Sub ADXMLImport()
On Error Resume Next
   On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim RecTDS1   As Variant
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim Sec44ADName, Sec44ADCode, Sec44ADDesc As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
        
    XpathOfPI = Common_year & "/ScheduleBP/NatOfBus44AD"
    Set Nodelist = dom.SelectNodes(XpathOfPI)
    
    Sec44ADName = Sheet3.Range("Sec44AD.NOBName").Column
    Sec44ADCode = Sheet3.Range("Sec44AD.NOBCode").Column
    Sec44ADDesc = Sheet3.Range("Sec44AD.NOBDescription").Column
    
    TotalExRow = Range("Sec44AD.NOBName").Rows.count
    
    TotalXMLRow = Nodelist.Length
   
    If (TotalXMLRow > 0) Then
        Sheet3.Range("Sec44AD.NOBName").ClearContents
        Sheet3.Range("Sec44AD.NOBCode").ClearContents
        Sheet3.Range("Sec44AD.NOBDescription").ClearContents

    End If

    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
             Sheet3.Cells(rowcount, Sec44ADName).value = UCase(node.SelectSingleNode("NameOfBusiness").text)
            
            
            Dim iNOB As Variant
            Dim sNOB As Variant
            

            
            iNOB = UCase(node.SelectSingleNode("CodeAD").text)
            sNOB = Findtext(iNOB, "NOB44AD")
            Sheet3.Cells(rowcount, Sec44ADCode).value = sNOB
             
            Sheet3.Cells(rowcount, Sec44ADDesc).value = UCase(node.SelectSingleNode("Description").text)

            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub
Sub ADAXMLImport()
On Error Resume Next
   On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim RecTDS1   As Variant
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim Sec44ADAName, Sec44ADACode, Sec44ADADesc As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
        
    XpathOfPI = Common_year & "/ScheduleBP/NatOfBus44ADA"
    Set Nodelist = dom.SelectNodes(XpathOfPI)
    
    Sec44ADAName = Sheet3.Range("Sec44AD.NOBName1").Column
    Sec44ADACode = Sheet3.Range("Sec44AD.NOBCode1").Column
    Sec44ADADesc = Sheet3.Range("Sec44AD.NOBDescription1").Column
    
    TotalExRow = Range("Sec44AD.NOBName1").Rows.count
    
    TotalXMLRow = Nodelist.Length
   
    If (TotalXMLRow > 0) Then
        Sheet3.Range("Sec44AD.NOBName1").ClearContents
        Sheet3.Range("Sec44AD.NOBCode1").ClearContents
        Sheet3.Range("Sec44AD.NOBDescription1").ClearContents

    End If

    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName1").name)
    rowcount = rowcount - 1
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
             Sheet3.Cells(rowcount, Sec44ADAName).value = UCase(node.SelectSingleNode("NameOfBusiness").text)
             
            
            Dim iNOB As Variant
            Dim sNOB As Variant
            iNOB = UCase(node.SelectSingleNode("CodeADA").text)
            sNOB = Findtext(iNOB, "NOB44ADA")
            Sheet3.Cells(rowcount, Sec44ADACode).value = sNOB
             
            Sheet3.Cells(rowcount, Sec44ADADesc).value = UCase(node.SelectSingleNode("Description").text)

            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub

Sub AEXMLImport()
On Error Resume Next
   On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim RecTDS1   As Variant
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim Sec44AEName, Sec44AECode, Sec44AEDesc As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
        
    XpathOfPI = Common_year & "/ScheduleBP/NatOfBus44AE"
    Set Nodelist = dom.SelectNodes(XpathOfPI)
    
    Sec44AEName = Sheet3.Range("Sec44AD.NOBName2").Column
    Sec44AECode = Sheet3.Range("Sec44AD.NOBCode2").Column
    Sec44AEDesc = Sheet3.Range("Sec44AD.NOBDescription2").Column
    
    TotalExRow = Range("Sec44AD.NOBName2").Rows.count
    
    TotalXMLRow = Nodelist.Length
   
    If (TotalXMLRow > 0) Then
        Sheet3.Range("Sec44AD.NOBName2").ClearContents
        Sheet3.Range("Sec44AD.NOBCode2").ClearContents
        Sheet3.Range("Sec44AD.NOBDescription2").ClearContents

    End If

    
    rowcount = getRowNo(Sheet3.Range("Sec44AD.NOBName2").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each node In Nodelist
        rowcount = rowcount + 1
        'For Each InnerNode In Node.ChildNodes
             Sheet3.Cells(rowcount, Sec44AEName).value = UCase(node.SelectSingleNode("NameOfBusiness").text)
             
            Dim iNOB As Variant
            Dim sNOB As Variant
            iNOB = UCase(node.SelectSingleNode("CodeAE").text)
            sNOB = Findtext(iNOB, "NOB")
            Sheet3.Cells(rowcount, Sec44AECode).MergeArea.value = sNOB
             
            Sheet3.Cells(rowcount, Sec44AEDesc).value = UCase(node.SelectSingleNode("Description").text)

            
       ' Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub

Sub GoodsTableXMLImport()
On Error Resume Next
   On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim RecTDS1   As Variant
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim SEC44AERegNo, Sec44AEOLE, Sec44AETonnageCpct, Sec44AENoOfMonths, Sec44AEPresumptiveIncome As Variant
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
        
    XpathOfPI = Common_year & "/ScheduleBP/GoodsDtlsUs44AE"
    Set Nodelist = dom.SelectNodes(XpathOfPI)
    
    SEC44AERegNo = Sheet3.Range("Sec44AE.RegNo").Column
    Sec44AEOLE = Sheet3.Range("Sec44AE.OLH").Column
    Sec44AETonnageCpct = Sheet3.Range("Sec44AE.TonnageCpct").Column
    Sec44AENoOfMonths = Sheet3.Range("Sheet44AE.NoOfMonths").Column
    Sec44AEPresumptiveIncome = Sheet3.Range("Sheet44AE.PresumptiveIncome").Column
    
    TotalExRow = Range("Sec44AE.RegNo").Rows.count
    
    TotalXMLRow = Nodelist.Length
   
    If (TotalXMLRow > 0) Then
    Sheet3.Range("Sec44AE.RegNo").ClearContents
    Sheet3.Range("Sec44AE.OLH").ClearContents
    Sheet3.Range("Sec44AE.TonnageCpct").ClearContents
    Sheet3.Range("Sheet44AE.NoOfMonths").ClearContents
    Sheet3.Range("Sheet44AE.PresumptiveIncome").ClearContents
    End If

    
    rowcount = getRowNo(Sheet3.Range("Sec44AE.RegNo").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
             Sheet3.Cells(rowcount, SEC44AERegNo).value = UCase(node.SelectSingleNode("RegNumberGoodsCarriage").text)
             Sheet3.Cells(rowcount, Sec44AEOLE).value = UCase(node.SelectSingleNode("OwnedLeasedHiredFlag").text)
            
            Dim iAEOLE As Variant
            Dim sAEOLE As Variant
            iAEOLE = UCase(node.SelectSingleNode("GoodsDtlsUs44AE/OwnedLeasedHiredFlag").text)
              
            If iAEOLE = "OWN" Then
            sAEOLE = "Owned"
            ElseIf iAEOLE = "LEASE" Then
            sAEOLE = "Leased"
            ElseIf iAEOLE = "HIRED" Then
            sAEOLE = "Hired"
            End If
        
             
            Sheet3.Cells(rowcount, Sec44AETonnageCpct).value = UCase(node.SelectSingleNode("TonnageCapacity").text)
            Sheet3.Cells(rowcount, Sec44AENoOfMonths).value = UCase(node.SelectSingleNode("HoldingPeriod").text)
            Sheet3.Cells(rowcount, Sec44AEPresumptiveIncome).value = UCase(node.SelectSingleNode("PresumptiveIncome").text)


            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub


'Malli

Sub setDiffTblinfo_80GGC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long

    ccount = 0
    mIntCells = Sheet11.Range("SCH_80GGC_DonationDate").count
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationDate").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
'    DefinedgridNameRange = "SCH_80GGC_DonationDate||SCH_80GGC_DonationAmtCash||SCH_80GGC_DonationAmtOtherMode||SCH_80GGC_DonationAmt||SCH_80GGC_EligibleDonationAmt||SCH_80GGC_ChnImpsNeftRtgsNum||SCH_80GGC_BankAccountNo||"
'    DefinedgridNameRange = "SCH_80GGC_DonationDate||SCH_80GGC_DonationAmtCash||SCH_80GGC_DonationAmtOtherMode||SCH_80GGC_DonationAmt||SCH_80GGC_EligibleDonationAmt||SCH_80GGC_Political_Name||SCH_80GGC_Political_PAN||SCH_80GGC_ChnImpsNeftRtgsNum||SCH_80GGC_BankAccountNo||"
'SCH_80GGC_IFSCCode
    DefinedgridNameRange = "SCH_80GGC_DonationDate||SCH_80GGC_DonationAmtCash||SCH_80GGC_DonationAmtOtherMode||SCH_80GGC_DonationAmt||SCH_80GGC_EligibleDonationAmt||SCH_80GGC_Political_Name||SCH_80GGC_Political_PAN||SCH_80GGC_ChnImpsNeftRtgsNum||SCH_80GGC_IFSCCode||"

End Sub




















