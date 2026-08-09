Attribute VB_Name = "mdImportXML"
Option Explicit

Dim dom1, dom As DOMDocument
Dim end_TaxP, rngname_TaxP As Variant
Dim end_TDSoth, rngname_TDSoth As Variant
Dim end_TDS3oth, rngname_TDS3oth As Variant
Dim end_TCS, rngname_TCS As Variant

Dim sPassword As String

Dim rowcount, cnt As Variant
Dim xml As Variant
Sub XMLimport()
On Error GoTo endline
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim rowcount As Long
    Dim cellcount As Long
    Dim rowRange As Range
    Dim cellrange As Range
    Application.ScreenUpdating = False
    'Application.EnableEvents = False
    xml = ""
    
    Set dom = New DOMDocument
    dom.Load (OpenXMLFileDialog)
    
    
    
    PersonalInfoXMLImport
    TradingAccountXMLImport
    FilingInfoXMLImport
    VeriInfoXMLImport
    TDSonSalaryXMLImport
    ITXMLImport
    TDSOthXMLImport
    TDS3XMLImport
    TCSXMLImport
    
'   OSDTAAXMLImport
    OSXMLImport
VIAXMLImport
 
    RefundInfoXMLImport
    SalaryInfoXMLImport
    SalaryAllowancesImport
    HPInfoXMLImport
    IFInfoXMLImport
    SPIInfoXMLImport
    FA_A1InfoXMLImport
    FA_A2InfoXMLImport
    FA_A3InfoXMLImport
    FA_A4InfoXMLImport
    FA_BInfoXMLImport
    FA_CInfoXMLImport
    FA_DInfoXMLImport
    FA_EInfoXMLImport
    FA_FInfoXMLImport
    FA_GInfoXMLImport
    ALInfoXMLImport
    Sch5AInfoXMLImport
    PTIInfoXMLImport
    NOBXMLImport
    Sheet1.Activate
    fmsgboxStatus "Please check that the particulars of  General Information, Income, Deductions and Tax credit details have been correctly imported. If need be, kindly make necessary changes to report correct details." ', vbOKOnly, "Alert"
endline:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ValidateImportfields
End Sub

Sub PersonalInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim iCountry, sCountry As Variant
    Dim sState, iState As Variant
    
    XpathOfPI = "/ns4:ITR/ns3:ITR3/PartA_GEN1/PersonalInfo"  '"/ITR/PersonalInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes
            Sheet1.Range("sheet1.FirstName").value = UCase(Node.SelectSingleNode("AssesseeName/FirstName").text)
            Sheet1.Range("sheet1.MiddleName").value = UCase(Node.SelectSingleNode("AssesseeName/MiddleName").text)
            Sheet1.Range("sheet1.SurNameOrOrgName").value = UCase(Node.SelectSingleNode("AssesseeName/SurNameOrOrgName").text)
            Sheet1.Range("sheet1.PAN").value = UCase(Node.SelectSingleNode("PAN").text)
            Status1 = UCase(Node.SelectSingleNode("Status").text)
            If Status1 = "I" Then
               Status1 = "I-INDIVIDUAL"
               ElseIf Status1 = "H" Then
               Status1 = "H-HUF"
            End If
            
            Sheet1.Range("sheet1.Status").value = Status1
       
            Sheet1.Range("sheet1.ResidenceNo").value = UCase(Node.SelectSingleNode("Address/ResidenceNo").text)
            Sheet1.Range("sheet1.ResidenceName").value = UCase(Node.SelectSingleNode("Address/ResidenceName").text)
            Sheet1.Range("sheet1.RoadOrStreet").value = UCase(Node.SelectSingleNode("Address/RoadOrStreet").text)
            Sheet1.Range("sheet1.LocalityOrArea").value = UCase(Node.SelectSingleNode("Address/LocalityOrArea").text)
            Sheet1.Range("sheet1.CityOrTownOrDistrict").value = UCase(Node.SelectSingleNode("Address/CityOrTownOrDistrict").text)
            
            iState = UCase(Node.SelectSingleNode("Address/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet1.Range("sheet1.StateCode1").value = sState
            
            Application.EnableEvents = False

            iCountry = UCase(Node.SelectSingleNode("Address/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet1.Range("sheet1.CountryCode").value = sCountry

            Application.EnableEvents = True
            
            If Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "99" Then
            
            Sheet1.Unprotect Password:=EfilingCommon.getmsgstate

               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
               
                Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)
               
             Sheet1.Protect Password:=EfilingCommon.getmsgstate
            Else
            
            Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
            
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
               
               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
               
               Sheet1.Protect Password:=EfilingCommon.getmsgstate
               
            End If
            
            
            Sheet1.Range("sheet1.PinCode").value = UCase(Node.SelectSingleNode("Address/PinCode").text)
            Sheet1.Range("sheet1.ZipCode").value = UCase(Node.SelectSingleNode("Address/ZipCode").text)
            Sheet1.Range("sheet1.STDcode").value = UCase(Node.SelectSingleNode("Address/Phone/STDcode").text)
            Sheet1.Range("sheet1.PhoneNo").value = UCase(Node.SelectSingleNode("Address/Phone/PhoneNo").text)
            
             Application.EnableEvents = False
            Sheet1.Range("sheet1.Mobileno").value = UCase(Node.SelectSingleNode("Address/MobileNo").text)
            Sheet1.Range("sheet1.MobileNoSec").value = UCase(Node.SelectSingleNode("Address/MobileNoSec").text)
             Application.EnableEvents = True
            
            Sheet1.Range("sheet1.EmailAddress").value = UCase(Node.SelectSingleNode("Address/EmailAddress").text)
            Sheet1.Range("sheet1.EmailAddressSec").value = UCase(Node.SelectSingleNode("Address/EmailAddressSec").text)
             Sheet1.Range("sheet1.adhaarno").value = UCase(Node.SelectSingleNode("AadhaarCardNo").text)
            dob = Node.SelectSingleNode("DOB").text
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            dob = DD & "/" & MM & "/" & YYYY
            
            Sheet1.Range("sheet1.DOB").value = dob

        Next InnerNode
    Next Node
End Sub
Sub TradingAccountXMLImport()
On Error Resume Next
    Dim XpathOfTA As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    

    XpathOfTA = "/ns4:ITR/ns3:ITR3/TradingAccount"
    Set Nodelist = dom.SelectNodes(XpathOfTA)

    For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes
            
            Sheet49.Range("TradingAcc_SalesOfGoods").value = UCase(Node.SelectSingleNode("SaleOfGoods").text)
            Sheet49.Range("TradingAcc_SalesOfServices").value = UCase(Node.SelectSingleNode("SaleOfServices").text)
            
                                                
        Next InnerNode
    Next Node
End Sub

Function Findtext(myinput As Variant, tblRange As Variant) As Variant
On Error Resume Next
    Dim Rng As Range
    Dim found As Boolean
    Dim searchtext As Variant
    found = False
    searchtext = ""
    
    For Each Rng In Range(tblRange)
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

Sub FilingInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim sREP, iREP As Variant
    Dim Repcapacity As Variant
    Dim iResidentialStatus As Variant
    Dim sResidentialStatus As Variant
    
    
    XpathOfPI = "/ns4:ITR/ns3:ITR3/PartA_GEN1/FilingStatus"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes
        'ReturnFile Sec/ Port 5A
        
            iReturnFile = UCase(Node.SelectSingleNode("ReturnFileSec").text)
            sReturnFile = Findtext(iReturnFile & " ", "ReturnFileUnderSection")
            
            If iReturnFile = "11" Then sReturnFile = "139(1)- On or Before due date"
            If iReturnFile = "12" Then sReturnFile = "139(4)- After due date"
            If iReturnFile = "13" Then sReturnFile = "142(1)"
            If iReturnFile = "14" Then sReturnFile = "148"
            If iReturnFile = "15" Then sReturnFile = "153A"
            If iReturnFile = "16" Then sReturnFile = "153C"
            If iReturnFile = "17" Then sReturnFile = "139(5)- Revised Return"
            If iReturnFile = "18" Then sReturnFile = "139(9)"
            If iReturnFile = "19" Then sReturnFile = "92CD-Modified return"
            If iReturnFile = "20" Then sReturnFile = "119(2)(b)- after condonation of delay" 'Uncommented Konda in AY2025-26 23-12-2025
            
'            If iReturnFile = "20" Then sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)" 'changed by sai
            
            Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
            

            
            iResidentialStatus = UCase(Node.SelectSingleNode("ResidentialStatus").text)
            sResidentialStatus = Findtext(iResidentialStatus & " ", "ResiStatus")
            Sheet1.Range("sheet1.ResidentialStatus1").value = sResidentialStatus
            
            
            Sheet1.Range("sheet1.ReceiptNo").value = UCase(Node.SelectSingleNode("ReceiptNo").text)
            
            dob = Node.SelectSingleNode("OrigRetFiledDate").text
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            dob = DD & "/" & MM & "/" & YYYY
            
            Sheet1.Range("sheet1.OrigRetFiledDate").value = dob

            iPort5A = UCase(Node.SelectSingleNode("PortugeseCC5A").text)
            sPort5A = IIf(iPort5A = "Y", "Yes", IIf(iPort5A = "N", "No", iPort5A))
            Sheet1.Range("sheet1.PortugeseCC5A").value = sPort5A
                                    
                                    
        
            iREP = UCase(Node.SelectSingleNode("AsseseeRepFlg").text)
            sREP = IIf(iPort5A = "Y", "Yes", IIf(iREP = "N", "No", iREP))
            Sheet1.Range("sheet1.AsseseeRepFlg").value = sREP

            Sheet1.Range("sheet1.RepName").value = UCase(Node.SelectSingleNode("AssesseeRep/RepName").text)
             Repcapacity = UCase(Node.SelectSingleNode("AssesseeRep/RepCapacity").text)
            If Repcapacity = "L" Then
               Repcapacity = "Legal Heir"
               ElseIf Repcapacity = "M" Then
               Repcapacity = "Manager"
               ElseIf Repcapacity = "G" Then
               Repcapacity = "Guardian"
               ElseIf Repcapacity = "O" Then
               Repcapacity = "Other"
            End If


            Sheet1.Range("sheet1.RepCapacity").value = Repcapacity
            Sheet1.Range("sheet1.RepAddress").value = UCase(Node.SelectSingleNode("AssesseeRep/RepAddress").text)
            Sheet1.Range("sheet1.RepPAN").value = UCase(Node.SelectSingleNode("AssesseeRep/RepPAN").text)


        Next InnerNode
    Next Node
End Sub

Sub VeriInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfPI = "/ns4:ITR/ns3:ITR3/Verification"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes
            
            Sheet52.Range("sheet9.AssesseeVerName").value = UCase(Node.SelectSingleNode("Declaration/AssesseeVerName").text)
            Sheet52.Range("sheet9.FatherName").value = UCase(Node.SelectSingleNode("Declaration/FatherName").text)
            Sheet52.Range("sheet9.PAN").value = UCase(Node.SelectSingleNode("Declaration/AssesseeVerPAN").text)
                                                
        Next InnerNode
    Next Node
End Sub
Sub SalaryInfoXMLImport()
On Error Resume Next
    Dim XpathOfSAL As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant

    XpathOfSAL = "/ns4:ITR/ns3:ITR3/ScheduleS/Salaries"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfSAL)

     
    TotalExRow = Sheet39.Range("SalarySectionCount").value
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    

    Dim i As Long
    For i = 1 To TotalDiffRow
      addSalariesBlock
    Next
    
    rowcount = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet39.Range("SAL.NameOfEmployer" & rowcount).value = UCase(Node.SelectSingleNode("NameOfEmployer").text)
            Sheet39.Range("SAL.AddrDetail" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetail/AddrDetail").text)
            
            iState = UCase(Node.SelectSingleNode("AddressDetail/StateCode").text)

            If Len(iState) = 1 Then
            iState = "0" & iState
            End If
            
            sState = Findtext(iState, "State")
            Sheet39.Range("SAL.StateCode" & rowcount).value = sState
            Sheet39.Range("SAL.PinCode" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet39.Range("SAL.PinCode" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetail/PinCode").text)
            
            
           Sheet39.Range("sheet.EmployerCategory" & rowcount).value = UCase(Node.SelectSingleNode("NatureOfEmployment").text)
            Dim NatureOfEmployment As String
            NatureOfEmployment = UCase(Node.SelectSingleNode("NatureOfEmployment").text)
            
            If NatureOfEmployment = "CGOV" Then NatureOfEmployment = "Central Government"
            If NatureOfEmployment = "SGOV" Then NatureOfEmployment = "State Government"
            If NatureOfEmployment = "PE" Then NatureOfEmployment = "Pensioners"
            If NatureOfEmployment = "PSU" Then NatureOfEmployment = "Public Sector Undertaking"
            If NatureOfEmployment = "OTH" Then NatureOfEmployment = "OTHERS"
            
            Sheet39.Range("sheet.EmployerCategory" & rowcount).value = NatureOfEmployment
 
           
            
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub
Sub SalaryAllowancesImport()
On Error Resume Next
    Dim XpathOfSalALL As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    
    Dim NatureColNo, DescriptionColNo, AmountColNo As Variant
    Dim TotalExRow, cnt  As Long
    
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim iNature, sNature As Variant
    
    XpathOfSalALL = "/ns4:ITR/ns3:ITR3/ScheduleS/AllwncExemptUs10/AllwncExemptUs10Dtls"
    Set Nodelist = dom.SelectNodes(XpathOfSalALL)
    
    NatureColNo = Sheet39.Range("Salary.Nature1").Column
    DescriptionColNo = Sheet39.Range("Salary.Description1").Column
    AmountColNo = Sheet39.Range("Salary.Amount1").Column
    
    TotalExRow = Sheet39.Range("Salary.Nature1").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet39.Range("Salary.Nature1").ClearContents
    Sheet39.Range("Salary.Description1").ClearContents
    Sheet39.Range("Salary.Amount1").ClearContents
    End If
    
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_SalaryAllowance (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet39.Range("Salary.Nature1").name)
    rowcount = rowcount - 1
    cnt = 0
'    If ActiveWorkbook.Sheets("Schedule S").Visible = xlSheetVisible Then
    For Each Node In Nodelist
        rowcount = rowcount + 1
            iNature = "Sec " & UCase(Node.SelectSingleNode("SalNatureDesc").text)
            sNature = Findtext(iNature, "Salary.OthersAllowance")
            
            Sheet39.Cells(rowcount, NatureColNo).value = sNature
            Sheet39.Cells(rowcount, DescriptionColNo).value = UCase(Node.SelectSingleNode("SalOthNatOfInc").text)
            Sheet39.Cells(rowcount, AmountColNo).value = Node.SelectSingleNode("SalOthAmount").text
        cnt = cnt + 1
    Next Node
'    End If
    
    XpathOfSalALL = "/ns4:ITR/ns3:ITR3/ScheduleS"
    Set Nodelist = dom.SelectNodes(XpathOfSalALL)
    rowcount = 0
        
    If ActiveWorkbook.Sheets("Schedule S").Visible = xlSheetVisible Then
    
        For Each Node In Nodelist
        rowcount = rowcount + 1
        Sheet39.Range("SAL.DeductionUnder6Section1").value = UCase(Node.SelectSingleNode("DeductionUnderSection16ia").text)
        Sheet39.Range("SAL.Entertainment_allowance_1").value = UCase(Node.SelectSingleNode("EntertainmntalwncUs16ii").text)
        Sheet39.Range("SAL.DeductionUnderSection1").value = UCase(Node.SelectSingleNode("ProfessionalTaxUs16iii").text)
        Next Node
    End If
End Sub

Sub AddDiffRows_SalaryAllowance(DiffRows As Long)
    setTblinfo_OthersNOI
    Sheet39.Activate
    searchLastRow ("Salary.Nature1")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_IncomeOth)
End Sub
Sub HPInfoXMLImport()
On Error Resume Next
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2 As IXMLDOMNodeList
    Dim Node, Node1, Node2 As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim InnerNode1, InnerNode2 As IXMLDOMNode
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant

    XpathOfHP = "/ns4:ITR/ns3:ITR3/ScheduleHP/PropertyDetails"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfHP)

    'Application.EnableEvents = False
    TotalExRow = Sheet8.Range("PropertySectionCOunt").value
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    

    Dim i As Long
    For i = 1 To TotalDiffRow
        Sheet8.Activate
        AddBlockCall_hprptfrm
    Next

    rowcount = 0
    cnt = 0

    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        'Co-Owner
        iTotalCoRow = Sheet8.Range("HP.Co.Name" & rowcount).value
        sTotalCoRow = Node.SelectNodes("CoOwners").Length
        TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
        If TotalDiffCoRow > 0 Then
            Sheet8.Activate
            AddPropertyCoOWners (TotalDiffCoRow)
        End If
        
        'Tenants
        iTotalTenRow = Sheet8.Range("HP.NameofTenant" & rowcount).value
        sTotalTenRow = Node.SelectNodes("TenantDetails").Length
        TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
        If TotalDiffTenRow > 0 Then
            Sheet8.Activate
            AddPropertyTenant (TotalDiffTenRow)
        End If
        
        For Each InnerNode In Node.ChildNodes
            
            Sheet8.Range("HP.AddrDetail" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetailWithZipCode/AddrDetail").text)
            Sheet8.Range("HP.CityOrTownOrDistrict" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetailWithZipCode/CityOrTownOrDistrict").text)
            
            iState = UCase(Node.SelectSingleNode("AddressDetailWithZipCode/StateCode").text)
            sState = Findtext(iState, "State")
            Sheet8.Range("HP.StateCode" & rowcount).value = sState
            iCountry = UCase(Node.SelectSingleNode("AddressDetailWithZipCode/CountryCode").text)
            sCountry = Findtext(iCountry, "Country")
            Sheet8.Range("HP.CountryCode" & rowcount).value = sCountry
            Sheet8.Range("HP.PinCode" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetailWithZipCode/PinCode").text)
            Sheet8.Range("HP.ZipCode" & rowcount).value = UCase(Node.SelectSingleNode("AddressDetailWithZipCode/ZipCode").text)
'            Sheet8.Range("HP.CoOwnedYN" & rowCount).value = UCase(Node.SelectSingleNode("PropCoOwnedFlg").Text)
            
            Dim Co_Ownedflag As Variant
            Co_Ownedflag = UCase(Node.SelectSingleNode("PropCoOwnedFlg").text)
            If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
            Co_Ownedflag = "Yes"
            ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
            Co_Ownedflag = "No"
            Else
            Co_Ownedflag = "(Select)"
            End If
            Sheet8.Range("HP.CoOwnedYN" & rowcount).value = Co_Ownedflag
             F
            Dim hpshare As Variant
            hpshare = UCase(Node.SelectSingleNode("AsseseeShareProperty").text)
            
            Sheet8.Range("HP.SharePercent" & rowcount).value = hpshare
            
            Set nodeList1 = Node.SelectNodes("CoOwners")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.Co.Name" & rowcount).name)
            cnt = cnt - 1
            For Each Node1 In nodeList1
                'Dim NameCoOwner, COPaN As Range
                'NameCoOwner = Sheet8.Range("HP.Co.Name" & rowCount).Column
                'COPaN = Sheet8.Range("HP.Co.Pan" & rowCount).Column
                cnt = cnt + 1
                For Each InnerNode1 In Node1.ChildNodes
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Name" & rowcount).Column).value = UCase(Node1.SelectSingleNode("NameCoOwner").text)
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Pan" & rowcount).Column).value = UCase(Node1.SelectSingleNode("PAN_CoOwner").text)
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Share" & rowcount).Column).value = UCase(Node1.SelectSingleNode("PercentShareProperty").text)
                Next InnerNode1
            Next Node1
            
            
            iLetOut = UCase(Node.SelectSingleNode("ifLetOut").text)
            
            
            If iLetOut = "Y" Then
            sLetOut = "Let Out"
            ElseIf iLetOut = "D" Then
            sLetOut = "Deemed Let Out"
            ElseIf iLetOut = "N" Then
            sLetOut = "Self Occupiedd"
            End If
            'IIf(iLetOut = "Y", "Let Out", IIf(iLetOut = "D", "Deemed Let Out", IIf(iLetOut = "N", "Self Occupiedd", "")))
            Sheet8.Range("HP.ifLetOut" & rowcount).value = sLetOut
            
'            (Select),Self Occupied, Let Out, Deemed Let Out
            
            Set Nodelist2 = Node.SelectNodes("TenantDetails")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.NameofTenant" & rowcount).name)
            cnt = cnt - 1
            For Each Node2 In Nodelist2
                cnt = cnt + 1
                For Each InnerNode2 In Node2.ChildNodes
                    Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node2.SelectSingleNode("NameofTenant").text)
                    Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node2.SelectSingleNode("PANofTenant").text)
                    Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node2.SelectSingleNode("PANTANofTenant").text)
                    'Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowCount).Column).value = UCase(Node2.SelectSingleNode("PANTANofTenant").Text)
                
                
                Next InnerNode2
            Next Node2
            
          Next InnerNode

            
            Sheet8.Range("HP.AnnualLetableValue" & rowcount).value = UCase(Node.SelectSingleNode("Rentdetails/AnnualLetableValue").text)
            
            Sheet8.Range("HP.RentNotRealized" & rowcount).value = UCase(Node.SelectSingleNode("Rentdetails/RentNotRealized").text)
            Sheet8.Range("HP.LocalTaxes" & rowcount).value = UCase(Node.SelectSingleNode("Rentdetails/LocalTaxes").text)
            Sheet8.Range("HP.IntOnBorwCap" & rowcount).value = UCase(Node.SelectSingleNode("Rentdetails/IntOnBorwCap").text)
            Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).value = UCase(Node.SelectSingleNode("Rentdetails/ArrearsUnrealizedRentRcvd").text)
            
        cnt = cnt + 1
            Next Node
        
    XpathOfHP = "/ns4:ITR/ns3:ITR3/ScheduleHP"
    Set Nodelist = dom.SelectNodes(XpathOfHP)
    rowcount = 0
        
    If ActiveWorkbook.Sheets("House Property").Visible = xlSheetVisible Then
    
        For Each Node In Nodelist
        rowcount = rowcount + 1
        Sheet8.Range("HP.PassTroughIncome").value = UCase(Node.SelectSingleNode("PassThroghIncome").text)
        Next Node
    End If


End Sub


Sub IFInfoXMLImport()
On Error Resume Next
    Dim XpathOfIF As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim FirmName, FirmPAN, BPFirmPAN, IsLiableToAudit, Is92EApplicable, ProfitSharePercent As Variant
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow  As Long

    XpathOfIF = "/ns4:ITR/ns3:ITR3/ScheduleIF/PartnerFirmDetails"
    Set Nodelist = dom.SelectNodes(XpathOfIF)

    FirmName = Sheet21.Range("IF.FirmName").Column
    FirmPAN = Sheet21.Range("IF.FirmPAN").Column
    IsLiableToAudit = Sheet21.Range("IF.IsLiableToAudit").Column
    Is92EApplicable = Sheet21.Range("IF.Is92EApplicable").Column
    ProfitSharePercent = Sheet21.Range("IF.ProfitSharePercent").Column
    
    TotalExRow = Sheet21.Range("IF.FirmName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet21.Range("IF.FirmName").ClearContents
        Sheet21.Range("IF.FirmName").ClearContents
        Sheet21.Range("IF.IsLiableToAudit").ClearContents
        Sheet21.Range("IF.Is92EApplicable").ClearContents
        Sheet21.Range("IF.ProfitSharePercent").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IF (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet21.Range("IF.FirmName").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            
            Sheet21.Cells(rowcount, FirmName).value = UCase(Node.SelectSingleNode("FirmName").text)
            Sheet21.Cells(rowcount, FirmPAN).value = UCase(Node.SelectSingleNode("FirmPAN").text)
        
            Dim LiableToAudit As Variant
            LiableToAudit = UCase(Node.SelectSingleNode("IsLiableToAudit").text)
            If UCase(LiableToAudit) = "YES" Or UCase(Mid(LiableToAudit, 1, 1) = "Y") Then
            LiableToAudit = "Y"
            ElseIf UCase(LiableToAudit) = "NO" Or UCase(Mid(LiableToAudit, 1, 1) = "N") Then
            LiableToAudit = "N"
            Else
            LiableToAudit = "(Select)"
            End If
            Sheet21.Cells(rowcount, IsLiableToAudit).value = LiableToAudit
            
            
            Dim EFirmFlag As Variant
            EFirmFlag = UCase(Node.SelectSingleNode("Sec92EFirmFlag").text)
            If UCase(EFirmFlag) = "YES" Or UCase(Mid(EFirmFlag, 1, 1) = "Y") Then
            EFirmFlag = "Y"
            ElseIf UCase(EFirmFlag) = "NO" Or UCase(Mid(EFirmFlag, 1, 1) = "N") Then
            EFirmFlag = "N"
            Else
            EFirmFlag = "(Select)"
            End If
            Sheet21.Cells(rowcount, Is92EApplicable).value = EFirmFlag
            
            
            Sheet21.Cells(rowcount, ProfitSharePercent).value = UCase(Node.SelectSingleNode("ProfitSharePercent").text)

        Next InnerNode
    Next Node
End Sub


Sub SPIInfoXMLImport()
On Error Resume Next
    Dim XpathOfSPI As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim SpecifiedPersonName, PANofSpecPerson, ReltnShip, NatureOfInc, AmtIncluded As Variant
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow  As Long

    XpathOfSPI = "/ns4:ITR/ns3:ITR3/ScheduleSPI/SpecifiedPerson"
    Set Nodelist = dom.SelectNodes(XpathOfSPI)

    SpecifiedPersonName = Sheet21.Range("SPI.SpecifiedPersonName").Column
    PANofSpecPerson = Sheet21.Range("SPI.PANofSpecPerson").Column
    ReltnShip = Sheet21.Range("SPI.ReltnShip").Column
    NatureOfInc = Sheet21.Range("SPI.NatureOfInc").Column
    AmtIncluded = Sheet21.Range("SPI.AmtIncluded").Column
    
    TotalExRow = Sheet21.Range("SPI.SpecifiedPersonName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet21.Range("SPI.SpecifiedPersonName").ClearContents
        Sheet21.Range("SPI.PANofSpecPerson").ClearContents
        Sheet21.Range("SPI.ReltnShip").ClearContents
        Sheet21.Range("SPI.NatureOfInc").ClearContents
        Sheet21.Range("SPI.AmtIncluded").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_SPI (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet21.Range("SPI.SpecifiedPersonName").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet21.Cells(rowcount, SpecifiedPersonName).value = UCase(Node.SelectSingleNode("SpecifiedPersonName").text)
            Sheet21.Cells(rowcount, PANofSpecPerson).value = UCase(Node.SelectSingleNode("PANofSpecPerson").text)
            Sheet21.Cells(rowcount, ReltnShip).value = UCase(Node.SelectSingleNode("ReltnShip").text)
            Sheet21.Cells(rowcount, NatureOfInc).value = UCase(Node.SelectSingleNode("NatureOfInc").text)
            Sheet21.Cells(rowcount, AmtIncluded).value = UCase(Node.SelectSingleNode("AmtIncluded").text)
        Next InnerNode
    Next Node
End Sub


Sub FA_A1InfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, Bankname, AddressOfBank, ZipCode, NameAsInAccount, OwnerStatus, ForeignAccountNumber, AccOpenDate As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsForiegnBank"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_A1_Country").Column
    Bankname = Sheet27.Range("FA_A1_BankName").Column
    AddressOfBank = Sheet27.Range("FA_A1_BankAdd").Column
    ZipCode = Sheet27.Range("FA_A1_ZipCode").Column
    ForeignAccountNumber = Sheet27.Range("FA_A1_ForeignAccountNumber").Column
    OwnerStatus = Sheet27.Range("FA_A1_StatusBeneficiary").Column
    AccOpenDate = Sheet27.Range("FA_A1_AccOpeningDate").Column
    
    TotalExRow = Sheet27.Range("FA_A1_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_A1_Country").ClearContents
        Sheet27.Range("FA_A1_BankName").ClearContents
        Sheet27.Range("FA_A1_BankAdd").ClearContents
        Sheet27.Range("FA_A1_ZipCode").ClearContents
        Sheet27.Range("FA_A1_ForeignAccountNumber").ClearContents
        Sheet27.Range("FA_A1_StatusBeneficiary").ClearContents
        Sheet27.Range("FA_A1_AccOpeningDate").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A1_FA1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A1_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
            
            Sheet27.Cells(rowcount, Bankname).value = UCase(Node.SelectSingleNode("Bankname").text)
            Sheet27.Cells(rowcount, AddressOfBank).value = UCase(Node.SelectSingleNode("AddressOfBank").text)
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            
            
            iStatus = UCase(Node.SelectSingleNode("OwnerStatus").text)
            sStatus = IIf(iStatus = "BENIFICIARY", "Beneficiary", IIf(iStatus = "BENEFICIAL_OWNER", "Beneficial owner", IIf(iStatus = "OWNER", "Owner", "")))
            Sheet27.Cells(rowcount, OwnerStatus).value = sStatus
            Sheet27.Cells(rowcount, ForeignAccountNumber).value = UCase(Node.SelectSingleNode("ForeignAccountNumber").text)
            
            
            Date_1 = UCase(Node.SelectSingleNode("AccOpenDate").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, AccOpenDate).value = Date_1
            
        Next InnerNode
    Next Node
End Sub
Sub FA_A2InfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, Bankname, AddressOfBank, ZipCode, NameAsInAccount, OwnerStatus, ForeignAccountNumber, AccOpenDate As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DtlsForeignCustodialAcc"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_A2_Country").Column
    Bankname = Sheet27.Range("FA_A2_BankName").Column
    AddressOfBank = Sheet27.Range("FA_A2_BankAdd").Column
    ZipCode = Sheet27.Range("FA_A2_ZipCode").Column
    ForeignAccountNumber = Sheet27.Range("FA_A2_ForeignAccountNumber").Column
    OwnerStatus = Sheet27.Range("FA_A2_StatusBeneficiary").Column
    AccOpenDate = Sheet27.Range("FA_A2_AccOpeningDate").Column
    
    TotalExRow = Sheet27.Range("FA_A2_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_A2_Country").ClearContents
        Sheet27.Range("FA_A2_BankName").ClearContents
        Sheet27.Range("FA_A2_BankAdd").ClearContents
        Sheet27.Range("FA_A2_ZipCode").ClearContents
        Sheet27.Range("FA_A2_ForeignAccountNumber").ClearContents
        Sheet27.Range("FA_A2_StatusBeneficiary").ClearContents
        Sheet27.Range("FA_A2_AccOpeningDate").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A2_FA2 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A2_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
            
            Sheet27.Cells(rowcount, Bankname).value = UCase(Node.SelectSingleNode("FinancialInstName").text)
            Sheet27.Cells(rowcount, AddressOfBank).value = UCase(Node.SelectSingleNode("FinancialInstAddress").text)
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            
            
            iStatus = UCase(Node.SelectSingleNode("Status").text)
            sStatus = IIf(iStatus = "BENIFICIARY", "Beneficiary", IIf(iStatus = "BENEFICIAL_OWNER", "Beneficial owner", IIf(iStatus = "OWNER", "Owner", "")))
            Sheet27.Cells(rowcount, OwnerStatus).value = sStatus
            Sheet27.Cells(rowcount, ForeignAccountNumber).value = UCase(Node.SelectSingleNode("AccountNumber").text)
            
            
            Date_1 = UCase(Node.SelectSingleNode("AccOpenDate").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, AccOpenDate).value = Date_1
            
        Next InnerNode
    Next Node
End Sub
Sub FA_A3InfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, Bankname, AddressOfBank, ZipCode, NatureOfEntity, OwnerStatus, Initialvalue, AccOpenDate As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DtlsForeignEquityDebtInterest"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_A3_Country").Column
    Bankname = Sheet27.Range("FA_A3_BankName").Column
    AddressOfBank = Sheet27.Range("FA_A3_BankAdd").Column
    ZipCode = Sheet27.Range("FA_A3_ZipCode").Column
    NatureOfEntity = Sheet27.Range("FA_A3_NatureOfEntity").Column
    Initialvalue = Sheet27.Range("FA_A3_initialvalue").Column
    AccOpenDate = Sheet27.Range("FA_A3_AccOpeningDate").Column
    
    TotalExRow = Sheet27.Range("FA_A3_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_A3_Country").ClearContents
        Sheet27.Range("FA_A3_BankName").ClearContents
        Sheet27.Range("FA_A3_BankAdd").ClearContents
        Sheet27.Range("FA_A3_ZipCode").ClearContents
        Sheet27.Range("FA_A3_NatureOfEntity").ClearContents
        Sheet27.Range("FA_A3_initialvalue").ClearContents
        Sheet27.Range("FA_A3_AccOpeningDate").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A3_FA3 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A3_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
            
            Sheet27.Cells(rowcount, Bankname).value = UCase(Node.SelectSingleNode("NameOfEntity").text)
            Sheet27.Cells(rowcount, AddressOfBank).value = UCase(Node.SelectSingleNode("AddressOfEntity").text)
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            
            Sheet27.Cells(rowcount, NatureOfEntity).value = UCase(Node.SelectSingleNode("NatureOfEntity").text)
            Sheet27.Cells(rowcount, Initialvalue).value = UCase(Node.SelectSingleNode("InitialValOfInvstmnt").text)
            
            Date_1 = UCase(Node.SelectSingleNode("InterestAcquiringDate").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, AccOpenDate).value = Date_1
            
        Next InnerNode
    Next Node
End Sub
Sub FA_A4InfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, Bankname, AddressOfBank, ZipCode, NatureOfEntity, OwnerStatus, Initialvalue, AccOpenDate As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DtlsForeignCashValueInsurance"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_A4_Country").Column
    Bankname = Sheet27.Range("FA_A4_BankName").Column
    AddressOfBank = Sheet27.Range("FA_A4_BankAdd").Column
    ZipCode = Sheet27.Range("FA_A4_ZipCode").Column
    AccOpenDate = Sheet27.Range("FA_A4_AccOpeningDate").Column
    
    TotalExRow = Sheet27.Range("FA_A4_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_A4_Country").ClearContents
        Sheet27.Range("FA_A4_BankName").ClearContents
        Sheet27.Range("FA_A4_BankAdd").ClearContents
        Sheet27.Range("FA_A4_ZipCode").ClearContents
        Sheet27.Range("FA_A4_AccOpeningDate").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A4_FA4 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A4_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
            
            Sheet27.Cells(rowcount, Bankname).value = UCase(Node.SelectSingleNode("FinancialInstName").text)
            Sheet27.Cells(rowcount, AddressOfBank).value = UCase(Node.SelectSingleNode("FinancialInstAddress").text)
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            
            
            Date_1 = UCase(Node.SelectSingleNode("ContractDate").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, AccOpenDate).value = Date_1
            
        Next InnerNode
    Next Node
End Sub
Sub FA_BInfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, NatureOfEntity, NameOfEntity, AddressOfEntity, NatureOfInt, DateHeld, TotalInv, ZipCode As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsFinancialInterest"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_B_Country").Column
    ZipCode = Sheet27.Range("FA_B_ZipCode").Column
    NatureOfEntity = Sheet27.Range("FA_B_EntityNature").Column
    NameOfEntity = Sheet27.Range("FA_B_EntityName").Column
    AddressOfEntity = Sheet27.Range("FA_B_EntityAdd").Column
    NatureOfInt = Sheet27.Range("FA_B_NatureOfInterest").Column
    DateHeld = Sheet27.Range("FA_B_DateSinceHeld").Column
    TotalInv = Sheet27.Range("FA_B_TotalInv").Column
    
    
    TotalExRow = Sheet27.Range("FA_B_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_B_Country").ClearContents
        Sheet27.Range("FA_B_ZipCode").ClearContents
        Sheet27.Range("FA_B_EntityNature").ClearContents
        Sheet27.Range("FA_B_EntityName").ClearContents
        Sheet27.Range("FA_B_EntityAdd").ClearContents
        Sheet27.Range("FA_B_NatureOfInterest").ClearContents
        Sheet27.Range("FA_B_DateSinceHeld").ClearContents
        Sheet27.Range("FA_B_TotalInv").ClearContents

    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_B1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_B_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)

            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            Sheet27.Cells(rowcount, NatureOfEntity).value = UCase(Node.SelectSingleNode("NatureOfEntity").text)
            Sheet27.Cells(rowcount, NameOfEntity).value = UCase(Node.SelectSingleNode("NameOfEntity").text)
            Sheet27.Cells(rowcount, AddressOfEntity).value = UCase(Node.SelectSingleNode("AddressOfEntity").text)
            
            iStatus = UCase(Node.SelectSingleNode("NatureOfInt").text)
            sStatus = IIf(iStatus = "BENIFICIARY", "Beneficiary", IIf(iStatus = "BENEFICIAL_OWNER", "Beneficial owner", IIf(iStatus = "DIRECT", "Direct", "")))
            Sheet27.Cells(rowcount, NatureOfInt).value = sStatus
            
            
            Date_1 = UCase(Node.SelectSingleNode("DateHeld").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, DateHeld).value = Date_1
            
            Sheet27.Cells(rowcount, TotalInv).value = UCase(Node.SelectSingleNode("TotalInvestment").text)
            
        Next InnerNode
    Next Node
End Sub

Sub FA_CInfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, AddressOfProperty, Ownership, DateOfAcq, ZipCode, TotalInv As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsImmovableProperty"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_C_Country").Column
    ZipCode = Sheet27.Range("FA_C_ZipCode").Column
    AddressOfProperty = Sheet27.Range("FA_C_PropertyAdd").Column
    Ownership = Sheet27.Range("FA_C_OwnershipBeneficiary").Column
    DateOfAcq = Sheet27.Range("FA_C_DateOfAcquisition").Column
'    TotalInv = Sheet27.Range("FA_C_TotalInv").Column
    
    
    TotalExRow = Sheet27.Range("FA_C_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_C_Country").ClearContents
        Sheet27.Range("FA_C_ZipCode").ClearContents
        Sheet27.Range("FA_C_PropertyAdd").ClearContents
        Sheet27.Range("FA_C_OwnershipBeneficiary").ClearContents
        Sheet27.Range("FA_C_DateOfAcquisition").ClearContents
'        Sheet27.Range("FA_C_TotalInv").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_C1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_C_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
           
            Sheet27.Cells(rowcount, AddressOfProperty).value = UCase(Node.SelectSingleNode("AddressOfProperty").text)
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            
            iStatus = UCase(Node.SelectSingleNode("Ownership").text)
            sStatus = IIf(iStatus = "BENIFICIARY", "Beneficiary", IIf(iStatus = "BENEFICIAL_OWNER", "Beneficial owner", IIf(iStatus = "DIRECT", "Direct", "")))
            Sheet27.Cells(rowcount, Ownership).value = sStatus
            
            
            Date_1 = UCase(Node.SelectSingleNode("DateOfAcq").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, DateOfAcq).value = Date_1
            
'            Sheet27.Cells(rowCount, TotalInv).value = UCase(Node.SelectSingleNode("TotalInvestment").Text)
            
            
        Next InnerNode
    Next Node
End Sub


Sub FA_DInfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, NatureOfAsset, Ownership, DateOfAcq, ZipCode As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsOthAssets"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_D_Country").Column
    ZipCode = Sheet27.Range("FA_D_ZipCode").Column
    NatureOfAsset = Sheet27.Range("FA_D_AssetNature").Column
    Ownership = Sheet27.Range("FA_D_OwnershipBeneficiary").Column
    DateOfAcq = Sheet27.Range("FA_D_DateOfAcquisition").Column

    
    TotalExRow = Sheet27.Range("FA_D_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_D_Country").ClearContents
        Sheet27.Range("FA_D_AssetNature").ClearContents
        Sheet27.Range("FA_D_OwnershipBeneficiary").ClearContents
        Sheet27.Range("FA_D_DateOfAcquisition").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_D1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_D_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)

            Sheet27.Cells(rowcount, NatureOfAsset).value = UCase(Node.SelectSingleNode("NatureOfAsset").text)
            
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            
            
            iStatus = UCase(Node.SelectSingleNode("Ownership").text)
            sStatus = IIf(iStatus = "BENIFICIARY", "Beneficiary", IIf(iStatus = "BENEFICIAL_OWNER", "Beneficial owner", IIf(iStatus = "DIRECT", "Direct", "")))
            Sheet27.Cells(rowcount, Ownership).value = sStatus
            
            
            Date_1 = UCase(Node.SelectSingleNode("DateOfAcq").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, DateOfAcq).value = Date_1
            
        Next InnerNode
    Next Node
End Sub

Sub FA_EInfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NameOfInstitution, AddressOfInstitution, NameMentionedInAccnt, InstitutionAccountNumber, ZipCode As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsOfAccntsHvngSigningAuth"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    NameOfInstitution = Sheet27.Range("FA_E_InstName").Column
    ZipCode = Sheet27.Range("FA_E_ZipCode").Column
    AddressOfInstitution = Sheet27.Range("FA_E_InstAdd").Column
    NameMentionedInAccnt = Sheet27.Range("FA_E_AccountName").Column
    InstitutionAccountNumber = Sheet27.Range("FA_E_InstitutionAccountNumber").Column

    
    TotalExRow = Sheet27.Range("FA_E_InstName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_E_InstName").ClearContents
        Sheet27.Range("FA_E_ZipCode").ClearContents
        Sheet27.Range("FA_E_InstAdd").ClearContents
        Sheet27.Range("FA_E_AccountName").ClearContents
        Sheet27.Range("FA_E_InstitutionAccountNumber").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_E1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_E_InstName").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, NameOfInstitution).value = UCase(Node.SelectSingleNode("NameOfInstitution").text)
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            Sheet27.Cells(rowcount, AddressOfInstitution).value = UCase(Node.SelectSingleNode("AddressOfInstitution").text)
            Sheet27.Cells(rowcount, NameMentionedInAccnt).value = UCase(Node.SelectSingleNode("NameMentionedInAccnt").text)
            Sheet27.Cells(rowcount, InstitutionAccountNumber).value = UCase(Node.SelectSingleNode("InstitutionAccountNumber").text)
        Next InnerNode
    Next Node
End Sub

Sub FA_FInfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, NameOfTrust, AddressOfTrust, NameOfOtherTrustees, AddressOfOtherTrustees, NameOfSettlor, AddressOfSettlor, NameOfBeneficiaries, AddressOfBeneficiaries, DateHeld As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsOfTrustOutIndiaTrustee"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_F_Country").Column
    ZipCode = Sheet27.Range("FA_F_ZipCode").Column
    NameOfTrust = Sheet27.Range("FA_F_NameOfTrust").Column
    AddressOfTrust = Sheet27.Range("FA_F_AddressOfTrust").Column
    NameOfOtherTrustees = Sheet27.Range("FA_F_NameOfOtherTrustees").Column
    AddressOfOtherTrustees = Sheet27.Range("FA_F_AddressOfOtherTrustees").Column
    NameOfSettlor = Sheet27.Range("FA_F_NameOfSettlor").Column
    AddressOfSettlor = Sheet27.Range("FA_F_AddressOfSettlor").Column
    NameOfBeneficiaries = Sheet27.Range("FA_F_NameOfBeneficiaries").Column
    AddressOfBeneficiaries = Sheet27.Range("FA_F_AddressOfBeneficiaries").Column
    DateHeld = Sheet27.Range("FA_F_DateSinceHeld").Column
    
    TotalExRow = Sheet27.Range("FA_F_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_F_Country").ClearContents
        Sheet27.Range("FA_F_ZipCode").ClearContents
        Sheet27.Range("FA_F_NameOfTrust").ClearContents
        Sheet27.Range("FA_F_AddressOfTrust").ClearContents
        Sheet27.Range("FA_F_NameOfOtherTrustees").ClearContents
        Sheet27.Range("FA_F_AddressOfOtherTrustees").ClearContents
        Sheet27.Range("FA_F_NameOfSettlor").ClearContents
        Sheet27.Range("FA_F_AddressOfSettlor").ClearContents
        Sheet27.Range("FA_F_NameOfBeneficiaries").ClearContents
        Sheet27.Range("FA_F_AddressOfBeneficiaries").ClearContents
        Sheet27.Range("FA_F_DateSinceHeld").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_F1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_F_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
            
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            Sheet27.Cells(rowcount, NameOfTrust).value = UCase(Node.SelectSingleNode("NameOfTrust").text)
            Sheet27.Cells(rowcount, AddressOfTrust).value = UCase(Node.SelectSingleNode("AddressOfTrust").text)
            Sheet27.Cells(rowcount, NameOfOtherTrustees).value = UCase(Node.SelectSingleNode("NameOfOtherTrustees").text)
            Sheet27.Cells(rowcount, AddressOfOtherTrustees).value = UCase(Node.SelectSingleNode("AddressOfOtherTrustees").text)
        
            Sheet27.Cells(rowcount, NameOfSettlor).value = UCase(Node.SelectSingleNode("NameOfSettlor").text)
            Sheet27.Cells(rowcount, AddressOfSettlor).value = UCase(Node.SelectSingleNode("AddressOfSettlor").text)
            
            Sheet27.Cells(rowcount, NameOfBeneficiaries).value = UCase(Node.SelectSingleNode("NameOfBeneficiaries").text)
            Sheet27.Cells(rowcount, AddressOfBeneficiaries).value = UCase(Node.SelectSingleNode("AddressOfBeneficiaries").text)
            
            Date_1 = UCase(Node.SelectSingleNode("DateHeld").text)
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            Sheet27.Cells(rowcount, DateHeld).value = Date_1
        
        Next InnerNode
    Next Node
End Sub

Sub FA_GInfoXMLImport()
On Error Resume Next
    Dim Date_1 As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim XpathOfFA As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim CountryName, NameOfPerson, AddressOfPerson, ZipCode, IncomeDerived, NatureOfIncome As Variant
    Dim strdate As String
    Dim TotalExRow  As Long
    Dim iStatus, sStatus As Variant
    
    XpathOfFA = "/ns4:ITR/ns3:ITR3/ScheduleFA/DetailsOfOthSourcesIncOutsideIndia"
    Set Nodelist = dom.SelectNodes(XpathOfFA)

    CountryName = Sheet27.Range("FA_G_Country").Column
    ZipCode = Sheet27.Range("FA_G_ZipCode").Column
    NameOfPerson = Sheet27.Range("FA_G_NameOfPerson").Column
    AddressOfPerson = Sheet27.Range("FA_G_AddressOfPerson").Column
    IncomeDerived = Sheet27.Range("FA_G_IncomeDerived").Column
    NatureOfIncome = Sheet27.Range("FA_G_NatureOfIncome").Column
  
    TotalExRow = Sheet27.Range("FA_G_Country").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet27.Range("FA_G_Country").ClearContents
        Sheet27.Range("FA_G_ZipCode").ClearContents
        Sheet27.Range("FA_G_NameOfPerson").ClearContents
        Sheet27.Range("FA_G_AddressOfPerson").ClearContents
        Sheet27.Range("FA_G_IncomeDerived").ClearContents
        Sheet27.Range("FA_G_NatureOfIncome").ClearContents
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_G1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_G_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet27.Cells(rowcount, CountryName).value = UCase(Node.SelectSingleNode("CountryCode").text) & "-" & UCase(Node.SelectSingleNode("CountryName").text)
            
            Sheet27.Cells(rowcount, ZipCode).value = UCase(Node.SelectSingleNode("ZipCode").text)
            Sheet27.Cells(rowcount, NameOfPerson).value = UCase(Node.SelectSingleNode("NameOfPerson").text)
            Sheet27.Cells(rowcount, AddressOfPerson).value = UCase(Node.SelectSingleNode("AddressOfPerson").text)
            Sheet27.Cells(rowcount, IncomeDerived).value = UCase(Node.SelectSingleNode("IncDerived").text)
            Sheet27.Cells(rowcount, NatureOfIncome).value = UCase(Node.SelectSingleNode("NatureOfInc").text)
        Next InnerNode
    Next Node
End Sub

Sub ALInfoXMLImport()
On Error Resume Next
    Dim XpathOfAL As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOfAL = "/ns4:ITR/ns3:ITR3/ScheduleAL"
    Set Nodelist = dom.SelectNodes(XpathOfAL)

    For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes
        
            Sheet43.Range("SchAL.B.Bank").value = UCase(Node.SelectSingleNode("MovableAsset/DepositsInBank").text)
            Sheet43.Range("SchAL.B.Share").value = UCase(Node.SelectSingleNode("MovableAsset/SharesAndSecurities").text)
            Sheet43.Range("SchAL.B.Insurance").value = UCase(Node.SelectSingleNode("MovableAsset/InsurancePolicies").text)
            Sheet43.Range("SchAL.B.Loans").value = UCase(Node.SelectSingleNode("MovableAsset/LoansAndAdvancesGiven").text)
            Sheet43.Range("SchAL.B.Cash").value = UCase(Node.SelectSingleNode("MovableAsset/CashInHand").text)
            Sheet43.Range("SchAL.B.Jwellery").value = UCase(Node.SelectSingleNode("MovableAsset/JewelleryBullionEtc").text)
            Sheet43.Range("SchAL.B.Art").value = UCase(Node.SelectSingleNode("MovableAsset/ArchCollDrawPaintSulpArt").text)
            Sheet43.Range("SchAL.B.Vehicles").value = UCase(Node.SelectSingleNode("MovableAsset/VehiclYachtsBoatsAircrafts").text)
            'Sheet43.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("MovableAsset/TotalImmovablMovablAssets").Text)
            Sheet43.Range("SchAL.Liability").value = UCase(Node.SelectSingleNode("LiabilityInRelatAssets").text)
        
        Next InnerNode
    Next Node
End Sub

Sub Sch5AInfoXMLImport()
On Error Resume Next
    Dim XpathOf5A As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
    XpathOf5A = "/ns4:ITR/ns3:ITR3/Schedule5A2014"
    Set Nodelist = dom.SelectNodes(XpathOf5A)

    For Each Node In Nodelist
'        For Each InnerNode In Node.ChildNodes
        
            Sheet29.Range("S5A_NameOfSpouse").value = UCase(Node.SelectSingleNode("NameOfSpouse").text)
            Sheet29.Range("S5A_PANOfSpouse").value = UCase(Node.SelectSingleNode("PANOfSpouse").text)

'        Next InnerNode
    Next Node
End Sub


Sub PTIInfoXMLImport()
On Error Resume Next
    Dim XpathOfPTI As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant

    XpathOfPTI = "/ns4:ITR/ns3:ITR3/SchedulePTI/SchedulePTIDtls"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfPTI)

     
    TotalExRow = Sheet41.Range("cntr.Ptirptfrm").value
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    

    Dim i As Long
    For i = 1 To TotalDiffRow
        Sheet41.Activate
        AddRows_PTI
    Next
    
    rowcount = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet41.Range("PTI_NameOfBusiness" & rowcount).value = UCase(Node.SelectSingleNode("BusinessName").text)
            Sheet41.Range("PTI_PAN" & rowcount).value = UCase(Node.SelectSingleNode("BusinessPAN").text)
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub

Sub NOBXMLImport()
On Error Resume Next
    Dim XpathOfNOB As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NOB, Trade1, Trade2, Trade3 As Variant
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim sNob, iNob As Variant

    XpathOfNOB = "/ns4:ITR/ns3:ITR3/PartA_GEN2/NatOfBus/NatureOfBusiness" '"/ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfNOB)

    NOB = Sheet2.Range("NOB.Code").Column
    Trade1 = Sheet2.Range("NOB.TradeNameOrDescription").Column
    Trade2 = Sheet2.Range("NOB.TradeNameOrDescription2").Column
    Trade3 = Sheet2.Range("NOB.TradeNameOrDescription3").Column
    TotalExRow = Range("NOB.Code").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet2.Range("NOB.Code").ClearContents
        Sheet2.Range("NOB.TradeNameOrDescription").ClearContents
        Sheet2.Range("NOB.TradeNameOrDescription2").ClearContents
        Sheet2.Range("NOB.TradeNameOrDescription3").ClearContents
    End If
    
    rowcount = getRowNo(Sheet2.Range("NOB.Code").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            'Sheet2.Cells(rowCount, NOB).value = UCase(Node.SelectSingleNode("Code").Text)
            
            iNob = UCase(Node.SelectSingleNode("Code").text)
            sNob = Findtext(iNob, "Nature_of_Business")
            Sheet2.Cells(rowcount, NOB).value = sNob
            Sheet2.Cells(rowcount, Trade1).value = UCase(Node.SelectSingleNode("TradeName1").text)
            Sheet2.Cells(rowcount, Trade2).value = UCase(Node.SelectSingleNode("TradeName2").text)
            Sheet2.Cells(rowcount, Trade3).value = UCase(Node.SelectSingleNode("TradeName3").text)
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub


Sub RefundInfoXMLImport()
On Error Resume Next
    Dim XpathOfRefund As String
    Dim XpathOfSection89 As Variant
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim IFSC, Bankname, ACCNO, CheckBox As Variant
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iStatus, sStatus As Variant
    
     XpathOfSection89 = "/ns4:ITR/ns3:ITR3/PartB_TTI/ComputationOfTaxLiability"
    Set Nodelist = dom.SelectNodes(XpathOfSection89)

        For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes

           Sheet7.Range("Sheet9.Section89").value = UCase(Node.SelectSingleNode("TaxRelief/Section89").text)
          
        Next InnerNode
    Next Node

    XpathOfRefund = "/ns4:ITR/ns3:ITR3/PartB_TTI/Refund/BankAccountDtls/AddtnlBankDetails"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)

    IFSC = Range("Sheet9.OthMICRCode").Column
    Bankname = Range("Sheet9.OthBankName").Column
    ACCNO = Range("Sheet9.OthBankAccountNumber").Column
    CheckBox = Range("tempxml").Column
    
    TotalExRow = Range("Sheet9.OthMICRCode").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Range("Sheet9.OthMICRCode").ClearContents
        Range("Sheet9.OthBankName").ClearContents
        Range("Sheet9.OthBankAccountNumber").ClearContents
        Range("tempxml").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_BANK (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet7.Range("Sheet9.OthMICRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
        

            
            Sheet7.Cells(rowcount, IFSC).value = UCase(Node.SelectSingleNode("IFSCCode").text)
            Sheet7.Cells(rowcount, Bankname).value = UCase(Node.SelectSingleNode("BankName").text)
            Sheet7.Cells(rowcount, ACCNO).value = UCase(Node.SelectSingleNode("BankAccountNo").text)
            Sheet7.Cells(rowcount, CheckBox).value = UCase(Node.SelectSingleNode("UseForRefund").text)
            
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
    
    XpathOfRefund = "/ns4:ITR/ns3:ITR3/PartB_TTI"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)
    
    For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes
            iStatus = UCase(Node.SelectSingleNode("AssetOutIndiaFlag").text)
            sStatus = IIf(iStatus = "YES", "Yes", IIf(iStatus = "NO", "No", ""))
            Sheet7.Range("AOIFlag_1").value = sStatus
        Next InnerNode
    Next Node
End Sub

Sub AddDiffRows_BANK(DiffRows As Long)
    setTableInfo_BANK
    Sheet7.Activate
    searchLastRow ("Sheet9.OthMICRCode")
    insertRowUnderSectionWithFormula (DiffRows)
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
    'Malli
    'DefinedgridNameRange = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.OthCashDeposited"
     DefinedgridNameRange = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount"
End Sub

Sub ITXMLImport()
On Error Resume Next
    Dim XpathOfIT As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Long
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfIT = "/ns4:ITR/ns3:ITR3/ScheduleIT/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfIT)

    BSRCodeColNo = Sheet25.Range("IT.BSRCode").Column
    DateDepColNo = Sheet25.Range("IT.DateDep").Column
    SrlNoChallanColNo = Sheet25.Range("IT.SrlNoOfChaln").Column
    AmtColNo = Sheet25.Range("IT.Amt").Column
    
    TotalExRow = Range("IT.BSRCode").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet25.Range("IT.BSRCode").ClearContents
        Sheet25.Range("IT.DateDep").ClearContents
        Sheet25.Range("IT.SrlNoOfChaln").ClearContents
        Sheet25.Range("IT.Amt").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IT (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet25.Range("IT.BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet25.Cells(rowcount, BSRCodeColNo).value = UCase(Node.SelectSingleNode("BSRCode").text)
            strdate = Node.SelectSingleNode("DateDep").text
            YYYY = Mid(strdate, 1, 4)
            MM = Mid(strdate, 6, 2)
            DD = Mid(strdate, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
            
            Sheet25.Cells(rowcount, DateDepColNo).value = strdate
            Sheet25.Cells(rowcount, SrlNoChallanColNo).value = Node.SelectSingleNode("SrlNoOfChaln").text
            Sheet25.Cells(rowcount, AmtColNo).value = Node.SelectSingleNode("Amt").text
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub

Sub TDSonSalaryXMLImport()
On Error Resume Next
    Dim XpathOfTDS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long

    XpathOfTDS = "/ns4:ITR/ns3:ITR3/ScheduleTDS1/TDSonSalary"
    Set Nodelist = dom.SelectNodes(XpathOfTDS)

    TANNoEmployer = Sheet28.Range("TDS1.TAN").Column
    TDSNameOfEmployer = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet28.Range("TDS1.IncChrgSal").Column
    TDSTotalTax = Sheet28.Range("TDS1.TotalTDSSal").Column
    
    TotalExRow = Range("TDS1.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet28.Range("TDS1.TAN").ClearContents
        Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet28.Range("TDS1.IncChrgSal").ClearContents
        Sheet28.Range("TDS1.TotalTDSSal").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TDS1.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet28.Cells(rowcount, TANNoEmployer).value = UCase(Node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet28.Cells(rowcount, TDSNameOfEmployer).value = UCase(Node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet28.Cells(rowcount, TDSIncomeCharge).value = UCase(Node.SelectSingleNode("IncChrgSal").text)
            Sheet28.Cells(rowcount, TDSTotalTax).value = UCase(Node.SelectSingleNode("TotalTDSSal").text)
            
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub

Sub TDSOthXMLImport()
On Error Resume Next
    Dim XpathOfTDS2 As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim TANColNo, DEDNameColNo, UTNColNo, FYColNo, TaxColNo, ClaimColNo, GrossAmountColNo, BroughtFwdTDSAmt, ClaimColNoSpouse, TDS2Credit, TDS2AadharOfOther As Variant
    Dim TDSCredit2 As Variant
    
    XpathOfTDS2 = "/ns4:ITR/ns3:ITR3/ScheduleTDS2/TDSOthThanSalaryDtls"
    Set Nodelist = dom.SelectNodes(XpathOfTDS2)
    
'    TDS2Credit = Sheet28.Range("TDS2.TDSCreditName").Column
'    TDS2AadharOfOther = Sheet28.Range("TDSAadharOfOther_TDS2").Column
    TANColNo = Sheet28.Range("TDS2.TAN").Column
    FYColNo = Sheet28.Range("TDS2.DeductedYear").Column
    BroughtFwdTDSAmt = Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column
    TaxColNo = Sheet28.Range("TDS2.DeductedInOwnHands").Column
    ClaimColNo = Sheet28.Range("TDS2.ClaimedInOwnHands").Column
    GrossAmountColNo = Sheet28.Range("TDS2.GrossAmount").Column
    
    
    
    TotalExRow = Range("TDS2.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
'        Sheet28.Range("TDS2.TDSCreditName").ClearContents
        Sheet28.Range("TDS2.TAN").ClearContents
        'Sheet28.Range("TDS2.EmployerOrDeductorOrCollecterName").ClearContents
        'Sheet28.Range("TDS2.UTN").ClearContents
        Sheet28.Range("TDS2.DeductedYear").ClearContents
        Sheet28.Range("TDS2.BroughtFwdTDSAmt").ClearContents
        Sheet28.Range("TDS2.DeductedInOwnHands").ClearContents
        Sheet28.Range("TDS2.ClaimedInOwnHands").ClearContents
        Sheet28.Range("TDS2.GrossAmount").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet28.Range("TDS2.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes

            Sheet28.Cells(rowcount, TANColNo).value = UCase(Node.SelectSingleNode("TANOfDeductor").text)
'            Sheet28.Cells(rowCount, DEDNameColNo).value = UCase(Node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").Text)
'            Sheet28.Cells(rowCount, UTNColNo).value = UCase(Node.SelectSingleNode("UniqueTDSCerNo").Text)
            Sheet28.Cells(rowcount, FYColNo).value = UCase(Node.SelectSingleNode("DeductedYr").text)
            Sheet28.Cells(rowcount, BroughtFwdTDSAmt).value = UCase(Node.SelectSingleNode("BroughtFwdTDSAmt").text)
            Sheet28.Cells(rowcount, TaxColNo).value = Node.SelectSingleNode("TaxDeductCreditDtls/TaxDeductedOwnHands").text
            Sheet28.Cells(rowcount, ClaimColNo).value = Node.SelectSingleNode("TaxDeductCreditDtls/TaxClaimedOwnHands").text
            
            Sheet28.Cells(rowcount, GrossAmountColNo).value = Node.SelectSingleNode("GrossAmount").text
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub

Sub TCSXMLImport()
On Error Resume Next
    Dim XpathOfTCS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTCS As Long
    Dim TANColNo, DEDNameColNo, TaxColNo, ClaimColNo, DeductedYearColNo As Variant
    Dim TotalExRow, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    XpathOfTCS = "/ns4:ITR/ns3:ITR3/ScheduleTCS/TCS"
    Set Nodelist = dom.SelectNodes(XpathOfTCS)
    
    TANColNo = Sheet28.Range("TCS.TAN").Column
    DEDNameColNo = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    TaxColNo = Sheet28.Range("TCS.TotalTCS").Column
    DeductedYearColNo = Sheet28.Range("TCS.DeductedYear").Column
    ClaimColNo = Sheet28.Range("TCS.AmtTCSClaimedThisYear").Column
    
    
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet28.Range("TCS.TAN").ClearContents
    Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
    Sheet28.Range("TCS.DeductedYear").ClearContents
    Sheet28.Range("TCS.TotalTCS").ClearContents
    Sheet28.Range("TCS.AmtTCSClaimedThisYear").ClearContents
    
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TCS (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
            Sheet28.Cells(rowcount, TANColNo).value = UCase(Node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet28.Cells(rowcount, DEDNameColNo).value = UCase(Node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet28.Cells(rowcount, DeductedYearColNo).value = Node.SelectSingleNode("DeductedYr").text
            Sheet28.Cells(rowcount, TaxColNo).value = Node.SelectSingleNode("TotalTCS").text
            Sheet28.Cells(rowcount, ClaimColNo).value = Node.SelectSingleNode("AmtTCSClaimedThisYear").text
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTCS = cnt
End Sub

Sub TDS3XMLImport()
On Error Resume Next
    Dim XpathOfTDS3 As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TANColNo, DEDNameColNo, UTNColNo, FYColNo, TaxColNo, ClaimColNo, GrossAmountColNo, ClaimSpouse, BroughtFwdTDSAmt, TDS3Credit As Variant
    Dim TotalExRow As Long
    Dim TDSCredit3 As Variant
    
    XpathOfTDS3 = "/ns4:ITR/ns3:ITR3/ScheduleTDS3/TDS3onOthThanSalDtls"
    Set Nodelist = dom.SelectNodes(XpathOfTDS3)
    
'    TDS3Credit = Sheet28.Range("TDS3.TDSCreditName").Column
    TANColNo = Sheet28.Range("TDS3.PAN").Column
    FYColNo = Sheet28.Range("TDS3.DeductedYear").Column
    BroughtFwdTDSAmt = Sheet28.Range("TDS3.BroughtFwdTDSAmt").Column
    TaxColNo = Sheet28.Range("TDS3.DeductedInOwnHands").Column
    ClaimColNo = Sheet28.Range("TDS3.ClaimedInOwnHands").Column
    GrossAmountColNo = Sheet28.Range("TDS3.GrossAmount").Column

    
    TotalExRow = Range("TDS3.PAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
'        Sheet28.Range("TDS3.TDSCreditName").ClearContents
        Sheet28.Range("TDS3.PAN").ClearContents
        
        Sheet28.Range("TDS3.DeductedYear").ClearContents
        Sheet28.Range("TDS3.BroughtFwdTDSAmt").ClearContents
        Sheet28.Range("TDS3.DeductedInOwnHands").ClearContents
        Sheet28.Range("TDS3.ClaimedInOwnHands").ClearContents
        Sheet28.Range("TDS3.GrossAmount").ClearContents
       
      
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDS3oth (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TDS3.PAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes

            Sheet28.Cells(rowcount, TANColNo).value = UCase(Node.SelectSingleNode("PANOfBuyerTenant").text)
'            Sheet28.Cells(rowCount, DEDNameColNo).value = UCase(Node.SelectSingleNode("EmployerOrDeductorOrCollectDetlTds3/EmployerOrDeductorOrCollecterName").Text)
'            Sheet28.Cells(rowCount, UTNColNo).value = UCase(Node.SelectSingleNode("UniqueTDSCerNo").Text)
            Sheet28.Cells(rowcount, FYColNo).value = UCase(Node.SelectSingleNode("DeductedYr").text)
            Sheet28.Cells(rowcount, BroughtFwdTDSAmt).value = UCase(Node.SelectSingleNode("BroughtFwdTDSAmt").text)
            
            Sheet28.Cells(rowcount, TaxColNo).value = Node.SelectSingleNode("TaxDeductCreditDtls/TaxDeductedOwnHands").text
            Sheet28.Cells(rowcount, ClaimColNo).value = Node.SelectSingleNode("TaxDeductCreditDtls/TaxClaimedOwnHands").text
            Sheet28.Cells(rowcount, GrossAmountColNo).value = Node.SelectSingleNode("TaxDeductCreditDtls/GrossAmount").text
            
            
           
        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Sub
Sub VIAXMLImport()
On Error Resume Next
    Dim XpathOfVIA As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim TotalExRow As Long
    
    XpathOfVIA = "/ns4:ITR/ns3:ITR3/ScheduleVIA"
    Set Nodelist = dom.SelectNodes(XpathOfVIA)
    
'    If ActiveWorkbook.Sheets("VI-A").Visible = xlSheetVisible Then
    For Each Node In Nodelist
            
    Sheet19.Range("scvia.Section80C").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80C").text)
    Sheet19.Range("scvia.Section80CCC").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCC").text)
    Sheet19.Range("scvia.Section80CCD_SE").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCDEmployeeOrSE").text)
    Sheet19.Range("scvia.Section80CCD1B_SE").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCD1B").text)
    Sheet19.Range("scvia.Section80CCD").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCDEmployer").text)
    Sheet19.Range("scvia.Section80DD").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DD").text)
    Sheet19.Range("scvia.Section80DDB").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDB").text)
    Sheet19.Range("scvia.Section80E").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80E").text)
    Sheet19.Range("scvia.Section80U").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80U").text)
    Sheet19.Range("scvia.Section80TTA").value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80TTA").text)
    
    Next Node
    
'End If
End Sub
Sub AddDiffRows_TDSoth(DiffRows As Long)
    Sheet28.Activate
    setTblinfo_TDS2
    Sheet28.Activate
    searchLastRow ("TDS2.TAN")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TDSoth)
End Sub

Sub AddDiffRows_TDS3oth(DiffRows As Long)
   Sheet28.Activate
   setTblinfo_TDS3
    Sheet28.Activate
    searchLastRow ("TDS3.PAN")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TDS3oth)
    End Sub

Sub AddDiffRows_TCS(DiffRows As Long)
    Sheet28.Activate
    setTblinfo_TCS
    Sheet28.Activate
    searchLastRow ("TCS.TAN")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TCS)
End Sub

Sub setTblinfo_TDS2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    Sheet28.Activate
    mIntCells = Sheet28.Range("TDS2.TAN").count
    Set rangecells = Sheet28.Range("TDS2.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    'TDS2.TdsDeducted       Konda updated on 28-05-2025
    'EfilingCommon.DefinedgridNameRange = "TDS2.TDSCreditName||TDS2.GrossAmount||TDS2.HeadOfIncome||TDS2.PanOtherPerson||TDSAadharOfOther_TDS2||TDS2.TAN||TDS2.DeductedYear||TDS2.BroughtFwdTDSAmt||TDS2.DeductedInOwnHands||TDS2.DeductedIncome5A||TDS2.DeductedTDS5A||TDS2.ClaimedInOwnHands||TDS2.ClaimedIncome5A||TDS2.ClaimedTDS5A||TDS2.ClaimedPAN5A||ClaimedAadhar5A_TDS2||TDS2.AmtCarriedForward"
    EfilingCommon.DefinedgridNameRange = "TDS2.TDSCreditName||TDS2.GrossAmount||TDS2.HeadOfIncome||TDS2.PanOtherPerson||TDSAadharOfOther_TDS2||TDS2.TAN||TDS2.TdsDeducted||TDS2.DeductedYear||TDS2.BroughtFwdTDSAmt||TDS2.DeductedInOwnHands||TDS2.DeductedIncome5A||TDS2.DeductedTDS5A||TDS2.ClaimedInOwnHands||TDS2.ClaimedIncome5A||TDS2.ClaimedTDS5A||TDS2.ClaimedPAN5A||ClaimedAadhar5A_TDS2||TDS2.AmtCarriedForward"
End Sub

Sub setTblinfo_TDS3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    Sheet28.Activate
    mIntCells = Sheet28.Range("TDS3.PAN").count
    Set rangecells = Sheet28.Range("TDS3.PAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS3oth = ccount
    'TDS3.TdsIsDeducted  Konda updated on 28-05-2025
    'EfilingCommon.DefinedgridNameRange = "TDS3.TDSCreditName||TDS3.GrossAmount||TDS3.HeadOfIncome||TDS3.PanOtherPerson||Aadhar_TDS3||TDS3.PAN||TDS3.Aadhaar||TDS3.DeductedYear||TDS3.BroughtFwdTDSAmt||TDS3.DeductedInOwnHands||TDS3.DeductedIncome5A||TDS3.DeductedTDS5A||TDS3.ClaimedInOwnHands||TDS3.ClaimedIncome5A||TDS3.ClaimedTDS5A||TDS3.ClaimedPAN5A||ClaimedAadhar5A_TDS3||TDS3.AmtCarriedForward"
    EfilingCommon.DefinedgridNameRange = "TDS3.TDSCreditName||TDS3.GrossAmount||TDS3.HeadOfIncome||TDS3.PanOtherPerson||Aadhar_TDS3||TDS3.PAN||TDS3.Aadhaar||TDS3.TdsIsDeducted||TDS3.DeductedYear||TDS3.BroughtFwdTDSAmt||TDS3.DeductedInOwnHands||TDS3.DeductedIncome5A||TDS3.DeductedTDS5A||TDS3.ClaimedInOwnHands||TDS3.ClaimedIncome5A||TDS3.ClaimedTDS5A||TDS3.ClaimedPAN5A||ClaimedAadhar5A_TDS3||TDS3.AmtCarriedForward"
End Sub

Sub setTblinfo_TCS()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    Sheet28.Activate
    mIntCells = Sheet28.Range("TCS.TAN").count
    Set rangecells = Sheet28.Range("TCS.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TCS = ccount
    EfilingCommon.DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.DeductedYear||TCS.BroughtFwdTDSAmt||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear||TCS.AmtCarriedFwd"
 End Sub

Sub setDiffTblinfo_TDS1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    Sheet28.Activate
    mIntCells = Sheet28.Range("TDS1.TAN").count
    Set rangecells = Sheet28.Range("TDS1.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    EfilingCommon.DefinedgridNameRange = "TDS1.TAN||TDS1.EmployerOrDeductorOrCollecterName||TDS1.IncChrgSal||TDS1.TotalTDSSal"
    'rngname_TaxP = "TDSal.TAN;TDSal.EmployerOrDeductorOrCollecterName;TDSal.IncChrgSal;TDSal.TotalTDSSal;"
End Sub
 
Sub AddDiffRows_TDS1(DiffRows As Long)
    Sheet28.Activate
    setDiffTblinfo_TDS1
    Sheet28.Activate
    searchLastRow ("TDS1.TAN")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TaxP)
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
        CloseMsg
    End If
    
    Set FileDialogBox = Nothing
    
endline:

End Function



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

End Function

Sub setDiffTblinfo_IT()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet25.Range("IT.BSRCode").count
    Set rangecells = Sheet25.Range("IT.BSRCode").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    EfilingCommon.DefinedgridNameRange = "IT.BSRCode||IT.DateDep||IT.SrlNoOfChaln||IT.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT||FormulaOfExSAT1||FormulaOfSATNew"
End Sub
 
Sub AddDiffRows_IT(DiffRows As Long)
    setDiffTblinfo_IT
    Sheet25.Activate
    searchLastRow ("IT.BSRCode")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TaxP)
End Sub


Sub AddDiffRows_IF(DiffRows As Long)
    setTblinfo_IF
    Sheet21.Activate
    searchLastRow ("IF.FirmName")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_IF)
End Sub

Sub AddDiffRows_SPI(DiffRows As Long)
    setTblinfoName_SPI
    Sheet21.Activate
    searchLastRow ("SPI.SpecifiedPersonName")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_SPI)
End Sub

Sub AddRows_A1_FA1(DiffRows As Long)
    setTblinfo_A1_FA
    Sheet27.Activate
    SelectLastRow ("FA_A1_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_A1_FA)
End Sub

Sub AddRows_A2_FA2(DiffRows As Long)
    setTblinfo_A2_FA
    Sheet27.Activate
    SelectLastRow ("FA_A2_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_A2_FA)
End Sub

Sub AddRows_A3_FA3(DiffRows As Long)
    setTblinfo_A3_FA
    Sheet27.Activate
    SelectLastRow ("FA_A3_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_A3_FA)
End Sub

Sub AddRows_A4_FA4(DiffRows As Long)
    setTblinfo_A4_FA
    Sheet27.Activate
    SelectLastRow ("FA_A4_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_A4_FA)
End Sub

Sub AddDiffRows_FA_B1(DiffRows As Long)
    setTblinfo_B_FA
    Sheet27.Activate
    searchLastRow ("FA_B_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_B_FA)
End Sub

Sub AddDiffRows_FA_C1(DiffRows As Long)
    setTblinfo_C_FA
    Sheet27.Activate
    searchLastRow ("FA_C_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_C_FA)
End Sub

Sub AddDiffRows_FA_D1(DiffRows As Long)
    setTblinfo_D_FA
    Sheet27.Activate
    searchLastRow ("FA_D_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_D_FA)
End Sub

Sub AddDiffRows_FA_E1(DiffRows As Long)
    setTblinfo_E_FA
    Sheet27.Activate
    searchLastRow ("FA_E_InstName")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_E_FA)
End Sub

Sub AddDiffRows_FA_F1(DiffRows As Long)
    setTblinfo_F_FA
    Sheet27.Activate
    searchLastRow ("FA_F_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_F_FA)
End Sub

Sub AddDiffRows_FA_G1(DiffRows As Long)
    setTblinfo_G_FA
    Sheet27.Activate
    searchLastRow ("FA_G_Country")
    InsertDiffRowsAndFillFormulas (DiffRows)
 Call ExendRangeNameToTable(DiffRows, rngname_G_FA)
End Sub

Sub OSXMLImport()
On Error Resume Next
    Dim XpathOfOS As String
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
   
    XpathOfOS = "/ns4:ITR/ns3:ITR3/ScheduleOS"
    Set Nodelist = dom.SelectNodes(XpathOfOS)

        For Each Node In Nodelist
        For Each InnerNode In Node.ChildNodes

           Sheet14.Range("os.DividendGross").value = UCase(Node.SelectSingleNode("IncOthThanOwnRaceHorse/DividendGross").text)
           Sheet14.Range("os.TermDeposit").value = UCase(Node.SelectSingleNode("IncOthThanOwnRaceHorse/IntrstFrmTermDeposit").text)
           Sheet14.Range("os.IncomeTaxRefund").value = UCase(Node.SelectSingleNode("IncOthThanOwnRaceHorse/IntrstFrmIncmTaxRefund").text)
           Sheet14.Range("os.Others").value = UCase(Node.SelectSingleNode("IncOthThanOwnRaceHorse/IntrstFrmOthers").text)
           Sheet14.Range("os.WinLottRacePuzz").value = UCase(Node.SelectSingleNode("IncOthThanOwnRaceHorse/LtryPzzlChrgblUs115BB").text)
           
           OSDTAAXMLImport
           
           
           Sheet14.Range("os.Receipts").value = UCase(Node.SelectSingleNode("IncFromOwnHorse/Receipts").text)

        Next InnerNode
    Next Node
End Sub

Sub OSDTAAXMLImport()
On Error Resume Next
    Dim XpathOfOSDTAA As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecOSDTAA As Long
    Dim DTAANature, DTAAAmount As Variant
    Dim TotalExRow, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim iDTAANature, sDTAANature As Variant
    
    XpathOfOSDTAA = "/ns4:ITR/ns3:ITR3/ScheduleOS/IncOthThanOwnRaceHorse/OthersGrossDtls"
    Set Nodelist = dom.SelectNodes(XpathOfOSDTAA)
    
    DTAANature = Sheet14.Range("os.SourceDescription").Column
    DTAAAmount = Sheet14.Range("os.SourceAmount").Column

    TotalExRow = Range("os.SourceDescription").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet14.Range("os.SourceDescription").ClearContents
    Sheet14.Range("os.SourceAmount").ClearContents

    End If

    If (TotalDiffRow > 0) Then
     AddRows_os2 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet14.Range("os.SourceDescription").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In Node.ChildNodes
        
        iDTAANature = (Node.SelectSingleNode("SourceDescription").text)
            
            If Mid(Sheet1.Range("Sheet1.115H").value, 1, 1) = "Y" Then

            sDTAANature = Findtext(iDTAANature + "  ", "OS_ALL")

            ElseIf Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
           
            sDTAANature = Findtext(iDTAANature + " ", "OS_NRI")

            Else
            sDTAANature = Findtext(iDTAANature + " ", "OS_RESandNOR")

            End If
 
            Sheet14.Cells(rowcount, DTAANature).value = sDTAANature
            Sheet14.Cells(rowcount, DTAAAmount).value = UCase(Node.SelectSingleNode("SourceAmount").text)

        Next InnerNode
        cnt = cnt + 1
    Next Node
    RecOSDTAA = cnt
End Sub

    Sub AddRows_os2(DiffRows As Long)
    setTblinfo_OS_xml
    Sheet14.Activate
    SelectLastRow ("os.SourceDescription")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    End Sub

'dpk3
Sub setTblinfo_OS_xml()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.othDesc").count
    Set rangecells = Sheet14.Range("os.othDesc").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os = ccount
    rngname_os = "os.othDesc;os.OtherSections;"
End Sub
