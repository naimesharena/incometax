Attribute VB_Name = "Generate_XML"
Option Explicit
Public xmlLinesCount As Variant
Public Drpdwn35AD, end_SalOthNature As Variant
Public NameOfEmployer_SALCount  As Long
Dim i As Long
Dim OthNature_Salary As Variant
Dim OthDescrp_Salary As Variant
Public end_stcgImbcount, end_ltcgImbcount As Long
Public end_SalaryNature, end_PerquisitesNature, end_ProfitsinLieuOfSalary As Long


Dim OthAmount_Salary, end_OE_xml, rngname_35AD_BP_xml, rngname_OE_BP_xml, BPDrp As Variant
Public end_BP_xml As Variant 'newly added


'Public ctCheck As Variant

Sub ValidatAllSchedule()
mainProcCaption = "Validating All Schedules"  'caption for progress bar caption
UserForm1.Show vbModeless
                                      'counter required for calculating the percentage of processes completed
noOfProcessMain = 38                                       'no of functions called in this sub


UpdateProgressBar


Validate_PartA_General

UpdateProgressBar
Validate_Gen_1398A

UpdateProgressBar
ValidateSheetNOB_ALL

UpdateProgressBar
ValidateBS_All

UpdateProgressBar
ValidateSheet_PL

UpdateProgressBar
ValidateSheetPARTA_OI

UpdateProgressBar
ValidateSheetQD

UpdateProgressBar
ValidatePartB

UpdateProgressBar
ValidateVerification1

UpdateProgressBar
validateSchS

UpdateProgressBar
ValidateSheetHouseProperty

UpdateProgressBar
ValidateSchBP

UpdateProgressBar
ValidateSheetDPM_DOA

UpdateProgressBar
ValidateSheetDEP_DCG


ValidateICDS

UpdateProgressBar
ValidateSheetESR

UpdateProgressBar
ValidateSheetCG_All

UpdateProgressBar
Validate112A

UpdateProgressBar
validate115AD

UpdateProgressBar
ValidateSheetOtherSource

UpdateProgressBar
ValidateCFLABFLA

UpdateProgressBar
ValidateCFL

UpdateProgressBar
ValidateSheetUnabsorbedDepreciation

UpdateProgressBar
ValidateSheet10A

UpdateProgressBar
ValidateSheetSchPTI

If (SheetALL.Range("bacValue").value) = 2 Then
    UpdateProgressBar
    ValidateSheet80G
    
    UpdateProgressBar
    Validate80GGA
    
    UpdateProgressBar
    Validate80GGANew
    
    '80GGC
    UpdateProgressBar
    Validate80GGC
    
    '80DD
     UpdateProgressBar
    Validate80DD
    
    '80U
     UpdateProgressBar
    Validate80U
    
    '28/05/2025
UpdateProgressBar
Validate80C_All
If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then
        UpdateProgressBar
        Validate80E_All
        UpdateProgressBar
        Validate80EE_All
        UpdateProgressBar
        Validate80EEA_All
        UpdateProgressBar
        Validate80EEB_All
End If
    
End If



UpdateProgressBar
ValidateSheet_Trading

UpdateProgressBar
ValidateSheet80

UpdateProgressBar
ValidateSheet80_IA

UpdateProgressBar
ValidateSPI_IF

UpdateProgressBar
ValidateSI

UpdateProgressBar
ValidateSchAMT

UpdateProgressBar

UpdateProgressBar
ValidateScheduleAMTC

UpdateProgressBar
ValidateSheetEI

UpdateProgressBar
ValidateSheetSchPTI

UpdateProgressBar
ValidateSheet_IT

UpdateProgressBar
ValidateScheduleFSI

UpdateProgressBar
ValidateSheetSchTDS

UpdateProgressBar
ValidateTPSC

UpdateProgressBar
Validate_80D
UpdateProgressBar
ValidateSheetSchTR_FA

UpdateProgressBar
VAlidateSheetFA

UpdateProgressBar
ValidateSchedule5A

UpdateProgressBar
ValidateSchAL

UpdateProgressBar
ValidateGSTNew
'    UpdateProgressBar
'    ValidateScheduleCT
'
'    UpdateProgressBar
'    ValidateGST
'
'
'    UpdateProgressBar
'    ValidateFD
'
'    UpdateProgressBar
'    ValidateCS

'ValidateBA
'ValidateSheetFT

mainProcCaption = "Calculating"
UpdateProgressBar
CG_Calc.doSetoff
UpdateProgressBar


UpdateProgressBar
CYLACalculations.CYLABFLASetOff

UpdateProgressBar
ValidateATI

End Sub

Sub ValidatAllScheduleTaxCalculate()
mainProcCaption = "Validating All Schedules"  'caption for progress bar caption
UserForm1.Show vbModeless
                                      'counter required for calculating the percentage of processes completed
noOfProcessMain = 38                                       'no of functions called in this sub


UpdateProgressBar


Validate_PartA_General

UpdateProgressBar
Validate_Gen_1398A

UpdateProgressBar
ValidateSheetNOB_ALL

UpdateProgressBar
ValidateBS_All

UpdateProgressBar
ValidateSheet_PL

UpdateProgressBar
ValidateSheetPARTA_OI

UpdateProgressBar
ValidateSheetQD

UpdateProgressBar
ValidatePartB

UpdateProgressBar
ValidateVerification1

UpdateProgressBar
validateSchS

UpdateProgressBar
ValidateSheetHouseProperty

UpdateProgressBar
ValidateSchBP

UpdateProgressBar
ValidateSheetDPM_DOA

UpdateProgressBar
ValidateSheetDEP_DCG


ValidateICDS

UpdateProgressBar
ValidateSheetESR

UpdateProgressBar
ValidateSheetCG_All

UpdateProgressBar
Validate112A

UpdateProgressBar
validate115AD

UpdateProgressBar
ValidateSheetOtherSource

UpdateProgressBar
ValidateCFLABFLA

UpdateProgressBar
ValidateCFL

UpdateProgressBar
ValidateSheetUnabsorbedDepreciation

UpdateProgressBar
ValidateSheet10A

UpdateProgressBar
ValidateSheetSchPTI


If (SheetALL.Range("bacValue").value) = 2 Then
    UpdateProgressBar
    ValidateSheet80G
    
    UpdateProgressBar
    Validate80GGA
    
    UpdateProgressBar
    Validate80GGANew
    
    '80GGC
    UpdateProgressBar
    Validate80GGC
    
    '80DD
     UpdateProgressBar
    Validate80DD
    
    '80U
     UpdateProgressBar
    Validate80U
    
    '28/05/2025
UpdateProgressBar
Validate80C_All
If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then
        UpdateProgressBar
        Validate80E_All
        UpdateProgressBar
        Validate80EE_All
        UpdateProgressBar
        Validate80EEA_All
        UpdateProgressBar
        Validate80EEB_All
End If
End If



UpdateProgressBar
ValidateSheet_Trading

UpdateProgressBar
ValidateSheet80

UpdateProgressBar
ValidateSheet80_IA

UpdateProgressBar
ValidateSPI_IF

UpdateProgressBar
ValidateSI

UpdateProgressBar
ValidateSchAMT

UpdateProgressBar

UpdateProgressBar
ValidateScheduleAMTC

UpdateProgressBar
ValidateSheetEI

UpdateProgressBar
ValidateSheetSchPTI

UpdateProgressBar
ValidateSheet_IT

UpdateProgressBar
ValidateScheduleFSI

UpdateProgressBar
ValidateSheetSchTDS

UpdateProgressBar
ValidateTPSC

UpdateProgressBar
Validate_80D
UpdateProgressBar
ValidateSheetSchTR_FA

UpdateProgressBar
VAlidateSheetFA

UpdateProgressBar
ValidateSchedule5A

UpdateProgressBar
ValidateSchAL

UpdateProgressBar
ValidateGSTNew
'    UpdateProgressBar
'    ValidateScheduleCT
'
'    UpdateProgressBar
'    ValidateGST
'
'
'    UpdateProgressBar
'    ValidateFD
'
'    UpdateProgressBar
'    ValidateCS

'ValidateBA
'ValidateSheetFT

mainProcCaption = "Calculating"
UpdateProgressBar
CG_Calc.doSetoff
UpdateProgressBar


UpdateProgressBar
CYLACalculations.CYLABFLASetOff

End Sub




Sub ValidateAll_XML()

If CalculateTaxCheckFlag = False Then
fmsgboxStatus "Please ensure to click on 'Calculate Tax' button before generating XML." ', vbCritical, "Error"
CloseMsg
End If
'ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
ValidatAllSchedule
'CG_Calc.doSetoff
'CYLACalculations.CYLABFLASetOff
'calculateTax
ProgressBarHide
'If Range("TDS1.IncSum").value > 0 Then
'    If (Range("SAL.TotalGrossSalary").value < (Range("TDS1.IncSum").value - 10)) Then
'        MsgBox "Gross salary disclosed in Salary schedule is less than 90% of salary reported in TDS1.", vbExclamation, "Warning"
'    End If
'End If

If end_NOBCode = 0 And end_NOBName = 0 And end_NOBC1 = 0 And end_NOBC2 = 0 Then
fmsgboxStatus "Atleast one Business Code is mandatory either in Schedule Nature of Business or in at Point-61 or 62 or 63 in Schedule Profit & Loss A/c" ', vbCritical
Sheet4.Activate
End

If (Sheet3.Range("sheet2.TotFundSrc").value <> Sheet3.Range("sheet3.TotFundApply").value) Then
fmsgboxStatus "Sources of funds must be equal to Application of Funds in Schedule Balance Sheet" ', vbCritical
Sheet3.Activate
End If

End If

    Dim IncChrgPG As Variant
    Dim IncTotFundSrc As Variant
    Dim IncTotSundryDbtAmt As Variant
    Dim IncTotSundryCrdAmt As Variant
    Dim IncTotStkInTradAmt As Variant
    Dim IncCashBalAmt As Variant

    IncChrgPG = IIf(Sheet9.Range("sheet12.IncChrgUnHdProftGain").value = "", 0, Sheet9.Range("sheet12.IncChrgUnHdProftGain").value)

    IncTotFundSrc = IIf(Sheet3.Range("sheet2.TotFundSrc").value = "", 0, Sheet3.Range("sheet2.TotFundSrc").value)

    IncTotSundryDbtAmt = IIf(Sheet3.Range("sheet3.TotSundryDbtAmt").value = "", 0, Sheet3.Range("sheet3.TotSundryDbtAmt").value)
    IncTotSundryCrdAmt = IIf(Sheet3.Range("sheet3.TotSundryCrdAmt").value = "", 0, Sheet3.Range("sheet3.TotSundryCrdAmt").value)
    IncTotStkInTradAmt = IIf(Sheet3.Range("sheet3.TotStkInTradAmt").value = "", 0, Sheet3.Range("sheet3.TotStkInTradAmt").value)
    IncCashBalAmt = IIf(Sheet3.Range("sheet3.CashBalAmt").value = "", 0, Sheet3.Range("sheet3.CashBalAmt").value)
    If Mid(Sheet2.Range("Sheet2.Code1").value, 1, 5) <> "00001" And Mid(Sheet2.Range("Sheet2.Code2").value, 1, 5) <> "00001" And Mid(Sheet2.Range("Sheet.Code3").value, 1, 5) <> "00001" Then
    If ((IncTotFundSrc = 0 And (IncTotSundryDbtAmt = 0 And IncTotSundryCrdAmt = 0 And IncTotStkInTradAmt = 0 And IncCashBalAmt = 0)) And _
       (IncChrgPG > 250000)) Then
        fmsgboxStatus ("* Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9).") ', vbExclamation, "Warning"
    End If
    End If



ThisWorkbook.Unprotect Password:=getmsgstate
Sheet36.Visible = xlSheetVisible
ThisWorkbook.Protect Password:=getmsgstate
Sheet36.Unprotect Password:=getmsgstate
Sheet36.Range("PreXML_SCHS").value = noOfSalFilled
setTblinfo_hprptfrm
Sheet36.Range("PreXML_SCHP").value = end_hprptfrm
Sheet36.Protect Password:=getmsgstate

fmsgboxsmall "To compute Setoff of Income, Tax and Interest using this utility, please follow the below" & Chr(13) & "Click on 'Compute Setoff' button in Sheet CYLA - BFLA" & Chr(13) & "Verify the details before saving the XML." ', vbInformation, "Alert"
fmsgboxsmall "In case you are liable for audit, please ensure Schedule OI is filled ." ', vbInformation, "Alert"

    Sheet36.Activate
    fmsgboxStatus "Click on 'Save XML' button to generate the XML and save the XML file. The path in which the XML will be stored is the extracted Ms Excel Utility folder." ', vbInformation, "ALert"
    ProgressBarHide
'ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate

End Sub

Sub Gen_XML()
On Error Resume Next

     Dim XMLFileName As String
     Dim AssesseePan
UserForm1.Show vbModeless

   'calculateTax
   ValidatAllSchedule

ThisWorkbook.Unprotect Password:=getmsgstate
Sheet36.Visible = xlSheetVisible
ThisWorkbook.Protect Password:=getmsgstate
Sheet36.Activate
   ' EfilingCommon.InitializeAllPublicVariables

       'Validate_PartA_General
       'ValidateSheetNOB_ALL
       'ValidateSheet_TDS

     '   ValidateTDS_All
      '   Validate80G_All
     '   ValidatePartBTI_BTTI_Verification

firstName = Sheet1.Range("sheet1.FirstName").value
middleName = Sheet1.Range("sheet1.MiddleName").value
'MiddleName = Sheet1.Range("sheet1.LastName").value
ResidenceName = Sheet1.Range("Sheet1.ResidenceName").value
'DesigOfficerWardorCircle = Sheet1.Range("sheet1.DesigOfficerWardorCircle").Value
'SpousePAN = Sheet1.Range("sheet1.PANOFSPOUSE").Value




'TypeOfHP = Sheet1.Range("IncD.TypeOfHP").Value
'IncomeFromHP = Sheet1.Range("IncD.IncomeFromHP").Value
'IncomeFromOS = Sheet1.Range("IncD.IncomeFromOS").Value
'GrossTotIncome = Sheet1.Range("IncD.GrossTotIncome").Value


'MobileNoSec = Sheet1.Range("sheet1.MobileNoSec").Value
    InitProgBar
    ProgressFrameCaption = "XML Generation"
    mainProcCaption = "Generating XML"





    AssesseePan = Sheet1.Range("sheet1.PAN").value
    XMLFileName = ThisWorkbook.Path & "\ITR3_" & AssesseePan & ".xml"
    noOfProcessMain = 44
    UpdateProgressBar
    Open XMLFileName For Output As #1
                XMLHeader
                UpdateProgressBar

                PersonalInfo
                UpdateProgressBar

                BalanceSheetXML
                UpdateProgressBar

                ManufacturingAccountSheetXML
                UpdateProgressBar

                TradingAccountSheetXML
                UpdateProgressBar

                PARTA_PL
                UpdateProgressBar

                PARTA_OI 'arjun
                UpdateProgressBar

                PARTA_OD
                UpdateProgressBar

                ScheduleSXML
                UpdateProgressBar

                ScheduleHPXML
                UpdateProgressBar

                ScheduleBPXML
                UpdateProgressBar

                ScheduleDPM
                UpdateProgressBar

                ScheduleDOA
                UpdateProgressBar

                ScheduleDEP
                UpdateProgressBar

                ScheduleDCG
                UpdateProgressBar

                ScheduleESR
                UpdateProgressBar

                ScheduleCGFor23
                UpdateProgressBar

                UpdateProgressBar
                Schedule112A

                UpdateProgressBar
                Schedule115AD

                ScheduleOSXML
                UpdateProgressBar

                ScheduleCYLA
                UpdateProgressBar

                ScheduleBFLA
                UpdateProgressBar

                ScheduleCFL
                UpdateProgressBar

                ScheduleUDXML
                UpdateProgressBar

                ScheduleICDS
                UpdateProgressBar

                Schedule10A
                UpdateProgressBar

                Schedule80G
                UpdateProgressBar

                Schedule80D
                UpdateProgressBar

                Schedule80RA
                UpdateProgressBar

                Schedule80_XML
                UpdateProgressBar

                VIAXML
                UpdateProgressBar

                ScheduleAMTXML
                UpdateProgressBar

                ScheduleAMTC
                UpdateProgressBar


                ScheduleSIXML
                UpdateProgressBar

                ScheduleSPIXML
                UpdateProgressBar


                If Not UBound(firmName_IF) = Empty Then
                    ScheduleIFXML
                    UpdateProgressBar
                    Else
                    UpdateProgressBar
                End If

                UpdateProgressBar

                ScheduleEIXML
                UpdateProgressBar

                XMLSchedulePTI
                UpdateProgressBar

                ScheduleTPSA
                UpdateProgressBar

                ScheduleFSI
                UpdateProgressBar

                ScheduleTR1
                UpdateProgressBar

                ScheduleFA
                UpdateProgressBar

                Schedule5AXML
                UpdateProgressBar

                ScheduleALXML
                UpdateProgressBar
                
                ScheduleGST
                UpdateProgressBar

                PARTB_TIXML
                UpdateProgressBar

                PARTB_TTIXML
                UpdateProgressBar

                TRPXML
                UpdateProgressBar

                ScheduleIT
                UpdateProgressBar

                ScheduleTDS
                UpdateProgressBar

                ScheduleTCS
                UpdateProgressBar

                Verification_XML
                UpdateProgressBar


'                 UpdateProgressBar
'                 XMLPART_CT
'
'                 UpdateProgressBar
'                 XMLPART_GST
'
'                 UpdateProgressBar
'                 XMLPART_FD
'
'                 UpdateProgressBar
'                 XMLPART_CS

                XMLFooter
                UpdateProgressBar
        Close #1
        UpdateProgressBar
'InitProgBar
ConvertXmlToString (XMLFileName)

ProgressBarHide
' and print the Acknowledgement.


fmsgboxStatus "File Saved " & XMLFileName ', vbOKOnly, "Alert"

'If (Mid(Range("Sheet1.ReturnType1").value, 1, 3) = "Rec") Then
'    MsgBox "Please upload the return(xml) on below link" & Chr(13) & Chr(13) & _
'    "https://incometaxindiaefiling.gov.in-->Login-->e-fiie---> Rectification " & " to file return", vbOKOnly, "ITR-3 AY XX19-2020"
'ElseIf (Mid(Range("Sheet1.ReturnType1").value, 1, 1) = "I") Then
'    MsgBox "Please upload the return(xml) on below link" & Chr(13) & Chr(13) & _
'    "https://incometaxindiaefiling.gov.in-->Login-->e-fiie---> Response to notice u/s139(9)" & " to file return", vbOKOnly, "ITR-3 AY XX19-2020"
'Else
    'fmsgboxStatus "Please upload the return(xml) on below link" & Chr(13) & Chr(13) & "https://incometaxindiaefiling.gov.in-->Login-->e-file---> Income Tax Return" ', vbOKOnly, "ITR-3 AY XX19-2020"
    fmsgboxStatus "Please upload the return(xml) on below link" & Chr(13) & Chr(13) & "https://www.incometax.gov.in/iec/foportal/-->Login-->e-file---> Income Tax Return" ', vbOKOnly, "ITR-3 AY 2022-2023"
'End If

End Sub

Public Function ConvertXmlToString(ByVal XMLFileName As String)
    Dim ts(1) As Object
    Dim count As Variant
    Dim FSO As Object
    Dim strXml, hashcode As String
    Dim SecretKey As String
    Dim iteration As Long
    Dim FileContents As String
    Dim SEARCH_FOR As String
    Dim REPLACE_WITH As String
    Dim arrSearchThis() As Variant
    Dim objTS As Variant
    Dim i, DigestFoundat As Variant
    iteration = SheetALL.Range("Hash_Iteration").value
    SecretKey = SheetALL.Range("Hash_Key").value

    Set FSO = CreateObject("Scripting.FileSystemObject")
    'strXml = FSO.OpenTextFile(XMLFileName).ReadAll
    'hashcode = Base64_HMACSHA256(iteration, strXml, SecretKey)
    Set objTS = FSO.OpenTextFile(XMLFileName, 1)
    i = 0
    InitProgBar
    ProgressFrameCaption = "XML Generation"
    mainProcCaption = "Encrypting XML"
    noOfProcessMain = 3
    UpdateProgressBar
    SEARCH_FOR = "<ITRForm:Digest>-</ITRForm:Digest>"
    subProcCaption = "Reading XML"
    noOfProcessSub = xmlLinesCount + 2

    Do Until objTS.AtEndOfStream
       ReDim Preserve arrSearchThis(i)
        arrSearchThis(i) = objTS.ReadLine
        If arrSearchThis(i) = SEARCH_FOR Then DigestFoundat = i
        arrSearchThis(i) = arrSearchThis(i) & Chr(13) & Chr(10)
        i = i + 1
        UpdateProgressBar
    Loop
    UpdateProgressBar
'    hashcode = Base64_HMACSHA256_test(iteration, arrSearchThis, SecretKey)
     hashcode = HMACSHA256A(str, StrConv(SecretKey, vbFromUnicode), iteration) 'by sai on 31/03/2025
     
    'UpdateProgressBar
     REPLACE_WITH = "<ITRForm:Digest>" & hashcode & "</ITRForm:Digest>"
    arrSearchThis(DigestFoundat) = Replace(arrSearchThis(DigestFoundat), SEARCH_FOR, REPLACE_WITH)
    Open XMLFileName For Output As #2
    subProcCaption = "Publishing XML"
    noOfProcessSub = UBound(arrSearchThis)
        For count = 0 To UBound(arrSearchThis)
            Print #2, arrSearchThis(count);
            UpdateProgressBar
        Next
    Close #2
End Function

'Sub validateAllSheets()
 '   Validate_PartA_General
  '  ValidateSheetNOB_ALL
   ' ValidateBS_All
    'ValidateSheet_PL
    'Validate_SchOI_All
   ' ValidateSheetQuantAnlysis_ALL


    'ValidatePartB
   ' SchHP.ValidateHP
    'ValidateSheetUD_All
    'SchCFL.ValidateCFL
    'Validate80G_All
    'SPI_SI.ValidateSPI_IF
    'ValidateSheet_IT
    'ValidateScheduleFSI
    'ValidateSheetSchTDS
    'ValidateScheduleTR
    'ValidateSchedule5A
    'SchAL.ValidateSchAL

'End Sub

'''''''''''''''''''''''''''''''''''''''''''''XML Parts


Function XMLHeader()
Dim creationdt As Variant
creationdt = Dformat(Sheet31.Range("Cuuren_Date").value, "")
Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
    'Print #1, "<ITRETURN:ITR xmlns:ITRETURN=" & Chr(34) & "http://incometaxindiaefiling.gov.in/main" & Chr(34) & " xmlns:ITR3FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR3" & Chr(34) & " xmlns:ITRForm=" & Chr(34) & "http://incometaxindiaefiling.gov.in/master" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & ">"
     Print #1, "<ITRETURN:ITR xmlns:ITRETURN=" & Chr(34) & "http://www.incometax.gov.in/iec/foportal/" & Chr(34) & " xmlns:ITR3FORM=" & Chr(34) & "http://www.incometax.gov.in/iec/foportal/" & Chr(34) & " xmlns:ITRForm=" & Chr(34) & "http://www.incometax.gov.in/iec/foportal/" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & ">"
 '   Print #1, "<ITR6FORM:ITR6>"
Print #1, "<ITR3FORM:ITR3>"
Print #1, " <ITRForm:CreationInfo>"
Print #1, "<ITRForm:SWVersionNo>R1</ITRForm:SWVersionNo>"
Print #1, "    <ITRForm:SWCreatedBy>SW90000101</ITRForm:SWCreatedBy>"
Print #1, "    <ITRForm:XMLCreatedBy>SW90000101</ITRForm:XMLCreatedBy>"
Print #1, "    <ITRForm:XMLCreationDate>" & creationdt & "</ITRForm:XMLCreationDate>"
Print #1, "    <ITRForm:IntermediaryCity>Delhi</ITRForm:IntermediaryCity>"
Print #1, "<ITRForm:Digest>-</ITRForm:Digest>"
Print #1, "</ITRForm:CreationInfo>"


Print #1, "<ITRForm:Form_ITR3>"
Print #1, "<ITRForm:FormName>" & "ITR-3" & "</ITRForm:FormName>"
Print #1, "<ITRForm:Description>For indls and HUFs having income from  a proprietory business or profession</ITRForm:Description>"
Print #1, "<ITRForm:AssessmentYear>" & "2023" & "</ITRForm:AssessmentYear>"
Print #1, "<ITRForm:SchemaVer>" & "Ver1.1" & "</ITRForm:SchemaVer>"
Print #1, "<ITRForm:FormVer>" & "Ver1.0" & "</ITRForm:FormVer>"
Print #1, "</ITRForm:Form_ITR3>"
xmlLinesCount = 18
End Function

Function PersonalInfo()

writeXML " <ITRForm:PartA_GEN1>"
'XML Generation Function For Personal Information
writeXML "    <ITRForm:PersonalInfo>"
'First Name
writeXML "         <ITRForm:AssesseeName>"
If firstName = "" Then
'writeXML "              <ITRForm:FirstName/>"
Else
writeXML "              <ITRForm:FirstName>" & UCase(firstName) & "</ITRForm:FirstName>"
End If

'MiddleName
If middleName = "" Then
'writeXML "              <ITRForm:MiddleName/>"
Else
writeXML "              <ITRForm:MiddleName>" & UCase(middleName) & "</ITRForm:MiddleName>"
End If

 'LastName
If LastName = "" Then
'writeXML "              <ITRForm:SurNameOrOrgName/>"
Else
writeXML "              <ITRForm:SurNameOrOrgName>" & UCase(LastName) & "</ITRForm:SurNameOrOrgName>"
End If
writeXML "         </ITRForm:AssesseeName>"

 'PAN
 If PAN = "" Then
writeXML "         <ITRForm:PAN/>"
Else
writeXML "         <ITRForm:PAN>" & UCase(PAN) & "</ITRForm:PAN>"
End If
writeXML "         <ITRForm:Address>"

'Flat / Door / Building
If Flat = "" Then
writeXML "              <ITRForm:ResidenceNo/>"
Else
writeXML "              <ITRForm:ResidenceNo>" & UCase(Flat) & "</ITRForm:ResidenceNo>"
End If
'ResidenceName
If ResidenceName = "" Then
'writeXML "              <ITRForm:ResidenceName/>"
Else
writeXML "              <ITRForm:ResidenceName>" & UCase(ResidenceName) & "</ITRForm:ResidenceName>"
End If
'RoadOrStreet
If RoadOrStreet = "" Then
'writeXML "              <ITRForm:RoadOrStreet/>"
Else
writeXML "              <ITRForm:RoadOrStreet>" & UCase(RoadOrStreet) & "</ITRForm:RoadOrStreet>"
End If
'AreaOrLocality

If Area = "" Then
writeXML "              <ITRForm:LocalityOrArea/>"
Else
writeXML "              <ITRForm:LocalityOrArea>" & UCase(Area) & "</ITRForm:LocalityOrArea>"
End If
'Town/City/District
If City = "" Then
writeXML "              <ITRForm:CityOrTownOrDistrict/>"
Else
writeXML "              <ITRForm:CityOrTownOrDistrict>" & UCase(City) & "</ITRForm:CityOrTownOrDistrict>"
End If
'State
If State = "" Then
writeXML "              <ITRForm:StateCode/>"
Else
writeXML "              <ITRForm:StateCode>" & UCase(Mid(State, 1, InStr(1, State, "-") - 1)) & "</ITRForm:StateCode>"
End If
'Country
If sCountry = "" Then
writeXML "              <ITRForm:CountryCode/>"
Else
writeXML "              <ITRForm:CountryCode>" & UCase(Mid(sCountry, 1, InStr(1, sCountry, "-") - 1)) & "</ITRForm:CountryCode>"
End If
'Pincode
If PinCode = "" Then
'writeXML "              <ITRForm:PinCode/>"
Else
writeXML "              <ITRForm:PinCode>" & UCase(PinCode) & "</ITRForm:PinCode>"
End If

'ZipCode
If ZipCode = "" Then
    'writeXML "              <ITRForm:ZipCode/>"
Else
    writeXML "              <ITRForm:ZipCode>" & UCase(ZipCode) & "</ITRForm:ZipCode>"
End If

'SDTCode and Phone Number
If STDcode <> "" Then
writeXML "              <ITRForm:Phone>"
If (Len(STDcode) > 0 Or Len(PhoneNo) > 0) Then
writeXML "              <ITRForm:STDcode>" & STDcode & "</ITRForm:STDcode>"
writeXML "              <ITRForm:PhoneNo>" & PhoneNo & "</ITRForm:PhoneNo>"
End If
writeXML "              </ITRForm:Phone>"
End If

'Mobile Number:Primary

If (Len(CountryCode1) > 0) Then
writeXML "<ITRForm:CountryCodeMobile>" & CountryCode1 & "</ITRForm:CountryCodeMobile>"
End If
If (Len(MobileNo) > 0) Then
writeXML "<ITRForm:MobileNo>" & MobileNo & "</ITRForm:MobileNo>"
End If
'Mobile Number:Secondary

If (Len(CountryCode2) > 0) Then
writeXML "<ITRForm:CountryCodeMobileNoSec>" & CountryCode2 & "</ITRForm:CountryCodeMobileNoSec>"
End If
If (Len(MobileNo_2) > 0) Then
writeXML "<ITRForm:MobileNoSec>" & MobileNo_2 & "</ITRForm:MobileNoSec>"
End If
'Email
If email = "" Then
writeXML "<ITRForm:EmailAddress/>"
Else
writeXML "<ITRForm:EmailAddress>" & email & "</ITRForm:EmailAddress>"
End If


'SecondaryEmail
If Sheet1.Range("sheet1.EmailAddressSec").value <> "" Then
writeXML "<ITRForm:EmailAddressSec>" & Sheet1.Range("sheet1.EmailAddressSec").value & "</ITRForm:EmailAddressSec>"
End If
writeXML "</ITRForm:Address>"
'Date Of Birth
If dob = "" Then
writeXML "         <ITRForm:DOB/>"
Else
writeXML "         <ITRForm:DOB>" & Dformat(dob, "") & "</ITRForm:DOB>"
End If

'Employee Category

'If empcat = "" Then
'writeXML "         <ITRForm:EmployerCategory>" & "" & "</ITRForm:EmployerCategory>"
'Else
'writeXML "         <ITRForm:EmployerCategory>" & UCase(empcat) & "</ITRForm:EmployerCategory>"
'End If
'Gender
'If Gender = "" Then
'writeXML "<ITRForm:Gender>" & "" & "</ITRForm:Gender>"
'Else
'writeXML "<ITRForm:Gender>" & UCase(Gender) & "</ITRForm:Gender>"
'End If
'Status
 If Status = "" Then
writeXML "<ITRForm:Status>" & "" & "</ITRForm:Status>"
Else
writeXML "<ITRForm:Status>" & Trim(UCase(Mid(Status, 1, InStr(1, UCase(Status), "-") - 1))) & "</ITRForm:Status>"
End If
'If UCase(Trim(Sheet1.Range("sheet1.adhaarflag"))) = "YES" Or _
'    UCase(Trim(Sheet1.Range("sheet1.adhaarflag"))) = "NO" Then
'
'    writeXML "<ITRForm:AadhaarCardFlg>" & Mid(UCase(Trim(Sheet1.Range("sheet1.adhaarflag"))), 1, 1) & "</ITRForm:AadhaarCardFlg>"
'End If

If Mid(UCase(Trim(Sheet1.Range("sheet1.adhaarno"))), 1, 1) <> "" Then
    writeXML "<ITRForm:AadhaarCardNo>" & UCase(Trim(Sheet1.Range("sheet1.adhaarno"))) & "</ITRForm:AadhaarCardNo>"
Else
End If

'Commented by Aavula
'If Mid(UCase(Trim(Sheet1.Range("sheet1.adhaarenrol"))), 1, 1) <> "" Then
'    writeXML "<ITRForm:AadhaarEnrolmentId>" & UCase(Trim(Sheet1.Range("sheet1.adhaarenrol"))) & "</ITRForm:AadhaarEnrolmentId>"
'Else
'End If

'If Trim(Sheet1.Range("sheet1.PassportNo").value) <> "" Then
'    writeXML "<ITRForm:PassportNo>" & UCase(Trim(Sheet1.Range("sheet1.PassportNo").value)) & "</ITRForm:PassportNo>"
'End If

writeXML "</ITRForm:PersonalInfo>"



writeXML "<ITRForm:FilingStatus>"
'Income Tax Ward / Circle
'If Sheet1.Range("sheet1.DesigOfficerWardorCircle").value = "" Then
'
'Else
'writeXML "<ITRForm:DesigOfficerWardorCircle>" & Sheet1.Range("sheet1.DesigOfficerWardorCircle").value & "</ITRForm:DesigOfficerWardorCircle>"
'End If
'------------------------------------------Return File Under Section
'If ReturnFileSec = "" Then
'writeXML "<ITRForm:ReturnFileSec>" & "" & "</ITRForm:ReturnFileSec>"
'Else
'writeXML "<ITRForm:ReturnFileSec>" & UCase(ReturnFileSec) & "</ITRForm:ReturnFileSec>"
'End If
' 'Notice No (Where the original return filed was defective and a notice is issued to the assessee to file a fresh return Sec 139(9)
'
'If (sNoticeNo = "") Then
'writeXML "<ITRForm:NoticeNo>" & "" & "</ITRForm:NoticeNo>"
'Else
'writeXML "<ITRForm:NoticeNo>" & UCase(sNoticeNo) & "</ITRForm:NoticeNo>"
'End If
''Whether original or revised return?
'If ReturnType = "" And ReturnType <> "(Select)" Then
'writeXML "<ITRForm:ReturnType>" & "" & "</ITRForm:ReturnType>"
'Else
'writeXML "<ITRForm:ReturnType>" & UCase(ReturnType) & "</ITRForm:ReturnType>"
'End If''

'If revised, defective, Modified then Enter Receipt No
'If ReturnType = "O" Then
'If sReceiptNo <> "" Then
'writeXML "<ITRForm:AckNoOriginalReturn>" & UCase(sReceiptNo) & "</ITRForm:AckNoOriginalReturn>"
'End If
'End If
'If ReturnType = "R" Then
'    If sReceiptNo <> "" Then
 '   writeXML "<ITRForm:ReceiptNo>" & UCase(sReceiptNo) & "</ITRForm:ReceiptNo>"
  '  End If
'End If

'Orig Ret Filed Date
'If Sheet1.Range("sheet1.OrigRetFiledDate").value <> "" Then
'writeXML "<ITRForm:OrigRetFiledDate>" & Sheet1.Range("sheet1.OrigRetFiledDate").value & "</ITRForm:OrigRetFiledDate>"
'Else
'writeXML "<ITRForm:OrigRetFiledDate />"
'End If
 '-----------------------------------------------------------------------------------------------------------------------
Dim ReturnFileSec1 As Variant
ReturnFileSec1 = Sheet1.Range("sheet1.ReturnFileSec1").value
If ReturnFileSec1 = "" Or ReturnFileSec1 = "(Select)" Then
'writeXML "         <ITRForm:ReturnFileSec>" & DefaultReturnFileSec & "</ITRForm:ReturnFileSec>"
Else
writeXML "         <ITRForm:ReturnFileSec>" & Mid(ReturnFileSec1, 1, 2) & "</ITRForm:ReturnFileSec>"
End If
'    If sNoticeNo <> "" Then
'    writeXML "<ITRForm:NoticeNo>" & sNoticeNo & "</ITRForm:NoticeNo>"
'    End If

'    If ReturnType <> "" Then
'    writeXML "<ITRForm:ReturnType>" & UCase(ReturnType) & "</ITRForm:ReturnType>"
'    End If

If ProvisoFlag = "Yes" Then
ProvisoFlag = "Y"
ElseIf ProvisoFlag = "No" Then
ProvisoFlag = "N"
End If

If ProvisoFlag <> "" Then
writeXML "<ITRForm:SeventhProvisio139>" & UCase(ProvisoFlag) & "</ITRForm:SeventhProvisio139>"
End If

If DepositAmountFlag = "Yes" Then
DepositAmountFlag = "Y"
ElseIf DepositAmountFlag = "No" Then
DepositAmountFlag = "N"
End If

If DepositAmountFlag <> "" Then
writeXML "<ITRForm:DepAmtAggAmtExcd1CrPrYrFlg>" & UCase(DepositAmountFlag) & "</ITRForm:DepAmtAggAmtExcd1CrPrYrFlg>"
End If

If DepositAmount <> "" Then
writeXML "<ITRForm:AmtSeventhProvisio139i>" & (DepositAmount) & "</ITRForm:AmtSeventhProvisio139i>"
End If

If AggrigateAmountFlag = "Yes" Then
AggrigateAmountFlag = "Y"
ElseIf AggrigateAmountFlag = "No" Then
AggrigateAmountFlag = "N"
End If

If AggrigateAmountFlag <> "" Then
writeXML "<ITRForm:IncrExpAggAmt2LkTrvFrgnCntryFlg>" & UCase(AggrigateAmountFlag) & "</ITRForm:IncrExpAggAmt2LkTrvFrgnCntryFlg>"
End If

If AggrigateAmount <> "" Then
writeXML "<ITRForm:AmtSeventhProvisio139ii>" & (AggrigateAmount) & "</ITRForm:AmtSeventhProvisio139ii>"
End If

If AggrigateAmountFlag1 = "Yes" Then
AggrigateAmountFlag1 = "Y"
ElseIf AggrigateAmountFlag1 = "No" Then
AggrigateAmountFlag1 = "N"
End If

If AggrigateAmountFlag1 <> "" Then
writeXML "<ITRForm:IncrExpAggAmt1LkElctrctyPrYrFlg>" & UCase(AggrigateAmountFlag1) & "</ITRForm:IncrExpAggAmt1LkElctrctyPrYrFlg>"
End If

If AggrigateAmount1 <> "" Then
writeXML "<ITRForm:AmtSeventhProvisio139iii>" & (AggrigateAmount1) & "</ITRForm:AmtSeventhProvisio139iii>"
End If

   If Sheet1.Range("Sheet1.UniqueNo").value <> "" Then
    writeXML "<ITRForm:NoticeNo>" & Sheet1.Range("Sheet1.UniqueNo").value & "</ITRForm:NoticeNo> "
    End If
    

    If sNoticeDate <> "" Then
    writeXML "<ITRForm:NoticeDate>" & Dformat(sNoticeDate, "") & "</ITRForm:NoticeDate>"
    End If

    
    If sReceiptNo <> "" Then
    writeXML "<ITRForm:ReceiptNo>" & sReceiptNo & "</ITRForm:ReceiptNo>"
    End If

    If Sheet1.Range("sheet1.OrigRetFiledDate").value <> "" Then
    writeXML "<ITRForm:OrigRetFiledDate>" & Dformat(Sheet1.Range("sheet1.OrigRetFiledDate").value, "") & "</ITRForm:OrigRetFiledDate> "
    End If
    
    


'Whether original or revised return?
'If ReturnType = "" Then
''writeXML "         <ITRForm:ReturnType>" & DefaultReturnType & "</ITRForm:ReturnType>"
'Else

'End If

'If ReturnFileSec = 17 Then
 '   writeXML "<ITRForm:ReceiptNo>" & sReceiptNo & "</ITRForm:ReceiptNo> "
  '  writeXML "<ITRForm:OrigRetFiledDate>" & Dformat(Sheet1.Range("sheet1.OrigRetFiledDate").value, "") & "</ITRForm:OrigRetFiledDate> "
   ' End If





'Residential Status
If ResidentialStatus = "" Then
writeXML "<ITRForm:ResidentialStatus>" & "" & "</ITRForm:ResidentialStatus>"
Else
writeXML "<ITRForm:ResidentialStatus>" & UCase(ResidentialStatus) & "</ITRForm:ResidentialStatus>"
End If


If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
  If Mid(ResidentialStatusCondition, 1, 25) = "You were in India for 182" Then
   ResidentialStatusCondition = 1
ElseIf Mid(ResidentialStatusCondition, 1, 24) = "You were in India for 60" Then
   ResidentialStatusCondition = 2
ElseIf Mid(ResidentialStatusCondition, 1, 28) = "You have been a non-resident" Then
   ResidentialStatusCondition = 3
ElseIf Mid(ResidentialStatusCondition, 1, 22) = "You have been in India" Then
   ResidentialStatusCondition = 4
ElseIf ResidentialStatusCondition = "You were a non-resident during the previous year." Then
   ResidentialStatusCondition = 5
End If

If Not ResidentialStatusCondition = "" Then
writeXML "<ITRForm:ConditionsResStatus>" & ResidentialStatusCondition & "</ITRForm:ConditionsResStatus>"
End If

End If

 If ResidentialStatus = "NRI" Then

Dim i As Long
If Not IsEmpty(JuridictionOfResidence) Then
writeXML "<ITRForm:JurisdictionResPrevYr>"
For i = 1 To UBound(JuridictionOfResidence)
writeXML "<ITRForm:JurisdictionResPrevYrDtls>"
If JuridictionOfResidence(i) <> "" Then
writeXML "<ITRForm:JurisdictionResidence>" & Mid(JuridictionOfResidence(i), 1, InStr(JuridictionOfResidence(i), "-") - 1) & "</ITRForm:JurisdictionResidence>"
End If

If Taxpayer_IdentificationNumber(i) <> "" Then
writeXML "<ITRForm:TIN>" & UCase(Taxpayer_IdentificationNumber(i)) & "</ITRForm:TIN>"
End If

writeXML "</ITRForm:JurisdictionResPrevYrDtls>"
Next
writeXML "</ITRForm:JurisdictionResPrevYr>"
End If

If Sheet1.Range("NRI_TotPeriodofStay_PreviousYear").value <> "" Then
writeXML "<ITRForm:TotalPrStayIndiaPrevYr>" & Sheet1.Range("NRI_TotPeriodofStay_PreviousYear").value & "</ITRForm:TotalPrStayIndiaPrevYr>"
End If

If Sheet1.Range("NRI_TotPeriodofStay_PrecedingYear").value <> "" Then
writeXML "<ITRForm:TotalPrStayIndia4PrecYr>" & Sheet1.Range("NRI_TotPeriodofStay_PrecedingYear").value & "</ITRForm:TotalPrStayIndia4PrecYr>"
End If

End If
'BenefitUs115HFlg Status
If Claim115H = "" Then
Else
writeXML "<ITRForm:BenefitUs115HFlg>" & UCase(Claim115H) & "</ITRForm:BenefitUs115HFlg>"
End If

'If NotifiedUs94AFlg_1 <> "" Then
'writeXML "<ITRForm:NotifiedUs94AFlg>" & NotifiedUs94AFlg_1 & "</ITRForm:NotifiedUs94AFlg>"
'Else
'writeXML "<ITRForm:NotifiedUs94AFlg />"
'End If

'If Sheet7.Range("Sheet9.RefundDue") > 0 Then
'    writeXML "<ITRForm:TaxStatus>TR</ITRForm:TaxStatus>"
'ElseIf Sheet7.Range("Sheet9.BalTaxPayable1") > 0 Then
'    writeXML "<ITRForm:TaxStatus>TP</ITRForm:TaxStatus>"
'Else
'    writeXML "<ITRForm:TaxStatus>NT</ITRForm:TaxStatus>"
'End If




'Are you Governed by Portugese Civil Code under Section 5A?
 If Portugese = "" Then
'writeXML "<ITRForm:PortugeseCC5A>&""&</ITRForm:PortugeseCC5A>"
Else
writeXML "<ITRForm:PortugeseCC5A>" & UCase(Portugese) & "</ITRForm:PortugeseCC5A>"
End If

'Assessee Rep Flag
If Sheet1.Range("sheet1.AsseseeRepFlg").value = "" Then
'writeXML "<ITRForm:AsseseeRepFlg>" & " " & "</ITRForm:AsseseeRepFlg>"
Else
writeXML "<ITRForm:AsseseeRepFlg>" & Trim(Mid(Sheet1.Range("sheet1.AsseseeRepFlg").value, 1, InStr(1, Sheet1.Range("sheet1.AsseseeRepFlg").value, "-") - 1)) & "</ITRForm:AsseseeRepFlg>"
End If


' Assessee Rep
If Mid(UCase(Sheet1.Range("sheet1.AsseseeRepFlg").value), 1, 1) = "Y" Then
    writeXML "<ITRForm:AssesseeRep>"
    If Sheet1.Range("sheet1.AsseseeRepFlg").value = "" And Sheet1.Range("sheet1.AsseseeRepFlg").value = "(Select)" Then

    Else
    writeXML "<ITRForm:RepName>" & Sheet1.Range("sheet1.RepName").value & "</ITRForm:RepName>"
    writeXML "<ITRForm:RepCapacity>" & Mid(Sheet1.Range("sheet1.RepCapacity").value, 1, 1) & "</ITRForm:RepCapacity>"
    writeXML "<ITRForm:RepAddress>" & Sheet1.Range("sheet1.RepAddress").value & "</ITRForm:RepAddress>"
    writeXML "<ITRForm:RepPAN>" & Sheet1.Range("sheet1.RepPAN").value & "</ITRForm:RepPAN>"
    If Sheet1.Range("sheet1.RepAadhar").value <> "" Then
    writeXML "<ITRForm:RepAadhaar>" & Sheet1.Range("sheet1.RepAadhar").value & "</ITRForm:RepAadhaar>"
    End If
    writeXML "</ITRForm:AssesseeRep>"
    End If
End If



 If Sheet1.Range("IsDirectorCompany").value = "" Or Sheet1.Range("IsDirectorCompany").value = "(Select)" Then

 Else
  writeXML "<ITRForm:CompDirectorPrvYrFlg>" & Mid(Sheet1.Range("IsDirectorCompany").value, 1, 1) & "</ITRForm:CompDirectorPrvYrFlg>"
 End If



'Director in company

If UBound(Director_NameofCompany) > 0 Then
writeXML "<ITRForm:CompDirectorPrvYr>"
For i = 1 To UBound(Director_NameofCompany)
writeXML "<ITRForm:CompDirectorPrvYrDtls>"
If Director_NameofCompany(i) <> "" Then
writeXML "<ITRForm:NameOfCompany>" & UCase(Director_NameofCompany(i)) & "</ITRForm:NameOfCompany>"
End If

If Director_TypeofCompany(i) <> "" Then
   If Director_TypeofCompany(i) = "Domestic" Then
      Director_TypeofCompany(i) = "D"
   ElseIf Director_TypeofCompany(i) = "Foreign" Then
      Director_TypeofCompany(i) = "F"
   End If

writeXML "<ITRForm:CompanyType>" & UCase(Director_TypeofCompany(i)) & "</ITRForm:CompanyType>"
End If

If Director_PAN(i) <> "" Then
writeXML "<ITRForm:PAN>" & UCase(Director_PAN(i)) & "</ITRForm:PAN>"
End If

If Director_SharesStatus(i) <> "" Then
writeXML "<ITRForm:SharesTypes>" & UCase(Mid(Director_SharesStatus(i), 1, 1)) & "</ITRForm:SharesTypes>"
End If

If Director_DIN(i) <> "" Then
writeXML "<ITRForm:DIN>" & UCase(Director_DIN(i)) & "</ITRForm:DIN>"
End If
writeXML "</ITRForm:CompDirectorPrvYrDtls>"
Next
writeXML "</ITRForm:CompDirectorPrvYr>"
End If



  If Sheet1.Range("IsPartnerFirm").value = "" Or Sheet1.Range("IsPartnerFirm").value = "(Select)" Then

 Else
  writeXML "<ITRForm:PartnerInFirmFlg>" & Mid(Sheet1.Range("IsPartnerFirm").value, 1, 1) & "</ITRForm:PartnerInFirmFlg>"
 End If


 If UBound(FirmName) > 0 Then
writeXML "<ITRForm:PartnerInFirm>"
For i = 1 To UBound(FirmName)
writeXML "<ITRForm:PartnerInFirmDtls>"
If FirmName(i) <> "" Then
writeXML "<ITRForm:NameOfFirm>" & UCase(FirmName(i)) & "</ITRForm:NameOfFirm>"
End If

If Firm_PAN(i) <> "" Then
writeXML "<ITRForm:PAN>" & UCase(Firm_PAN(i)) & "</ITRForm:PAN>"
End If

writeXML "</ITRForm:PartnerInFirmDtls>"
Next
writeXML "</ITRForm:PartnerInFirm>"
End If


'Unlisted Shares
  If Sheet1.Range("IsShareOwner").value = "" Or Sheet1.Range("IsShareOwner").value = "(Select)" Then

 Else
  writeXML "<ITRForm:HeldUnlistedEqShrPrYrFlg>" & Mid(Sheet1.Range("IsShareOwner").value, 1, 1) & "</ITRForm:HeldUnlistedEqShrPrYrFlg>"
 End If

  If UBound(Share_NameofCompany) > 0 Then
writeXML "<ITRForm:HeldUnlistedEqShrPrYr>"
For i = 1 To UBound(Share_NameofCompany)
writeXML "<ITRForm:HeldUnlistedEqShrPrYrDtls>"

If Share_NameofCompany(i) <> "" Then
writeXML "<ITRForm:NameOfCompany>" & UCase(Share_NameofCompany(i)) & "</ITRForm:NameOfCompany>"
End If

If Share_TypeofCompany(i) <> "" Then
   If Share_TypeofCompany(i) = "Domestic" Then
      Share_TypeofCompany(i) = "D"
   ElseIf Share_TypeofCompany(i) = "Foreign" Then
      Share_TypeofCompany(i) = "F"
   End If

writeXML "<ITRForm:CompanyType>" & UCase(Share_TypeofCompany(i)) & "</ITRForm:CompanyType>"
End If

If Share_PAN(i) <> "" Then
writeXML "<ITRForm:PAN>" & UCase(Share_PAN(i)) & "</ITRForm:PAN>"
End If

If Share_OP_Count(i) <> "" Then
writeXML "<ITRForm:OpngBalNumberOfShares>" & UCase(Share_OP_Count(i)) & "</ITRForm:OpngBalNumberOfShares>"
End If

If Share_OP_Cost(i) <> "" Then
writeXML "<ITRForm:OpngBalCostOfAcquisition>" & Round(UCase(Share_OP_Cost(i)), 2) & "</ITRForm:OpngBalCostOfAcquisition>"
End If

If Share_Acquired_Count(i) <> "" Then
writeXML "<ITRForm:ShrAcqDurYrNumberOfShares>" & UCase(Share_Acquired_Count(i)) & "</ITRForm:ShrAcqDurYrNumberOfShares>"
End If

If Share_Acquired_Date(i) <> "" Then
writeXML "<ITRForm:DateOfSubscrPurchase>" & Dformat(Share_Acquired_Date(i), "") & "</ITRForm:DateOfSubscrPurchase>"
End If

If Share_Acquired_FaceValue(i) <> "" Then
writeXML "<ITRForm:FaceValuePerShare>" & Round(UCase(Share_Acquired_FaceValue(i)), 2) & "</ITRForm:FaceValuePerShare>"
End If

If Share_Acquired_Issue(i) <> "" Then
writeXML "<ITRForm:IssuePricePerShare>" & UCase(Share_Acquired_Issue(i)) & "</ITRForm:IssuePricePerShare>"
End If

If Share_Acquired_Cost(i) <> "" Then
writeXML "<ITRForm:PurchasePricePerShare>" & Round(UCase(Share_Acquired_Cost(i)), 2) & "</ITRForm:PurchasePricePerShare>"
End If

If Share_Transferred_Count(i) <> "" Then
writeXML "<ITRForm:ShrTrnfNumberOfShares>" & UCase(Share_Transferred_Count(i)) & "</ITRForm:ShrTrnfNumberOfShares>"
End If

If Share_Transferred_Sale(i) <> "" Then
writeXML "<ITRForm:ShrTrnfSaleConsideration>" & Round(UCase(Share_Transferred_Sale(i)), 2) & "</ITRForm:ShrTrnfSaleConsideration>"
End If

If Share_CB_Count(i) <> "" Then
writeXML "<ITRForm:ClsngBalNumberOfShares>" & UCase(Share_CB_Count(i)) & "</ITRForm:ClsngBalNumberOfShares>"
End If

If Share_CB_Cost(i) <> "" Then
writeXML "<ITRForm:ClsngBalCostOfAcquisition>" & Round(UCase(Share_CB_Cost(i)), 2) & "</ITRForm:ClsngBalCostOfAcquisition>"
End If

writeXML "</ITRForm:HeldUnlistedEqShrPrYrDtls>"
Next
writeXML "</ITRForm:HeldUnlistedEqShrPrYr>"
End If


If Mid(Sheet1.Range("sheet1.NRI_PE").value, 1, 1) = "(" Or Sheet1.Range("sheet1.NRI_PE").value = "" Then

Else
writeXML "<ITRForm:NriPEinIndia>" & UCase(Mid(Sheet1.Range("sheet1.NRI_PE").value, 1, 1)) & "</ITRForm:NriPEinIndia>"
End If


If Sheet1.Range("sheet1.Foreign_Exchange").value <> "" Then
   writeXML "<ITRForm:ForeignExchangeFlag>" & Mid(Sheet1.Range("sheet1.Foreign_Exchange").value, 1, 1) & "</ITRForm:ForeignExchangeFlag>"
Else
   writeXML "<ITRForm:ForeignExchangeFlag />"
End If




'If filed, in response to a notice u/s 139(9)/142(1)/148/153A/153C enter date of such notice

'If sNoticeDate <> "" Then
'writeXML " <ITRForm:NoticeDateUnderSec>" & Dformat(sNoticeDate, "") & "</ITRForm:NoticeDateUnderSec>"
'Else

'End If



writeXML "</ITRForm:FilingStatus>"
writeXML "</ITRForm:PartA_GEN1>"

'--------------------------------
writeXML "<ITRForm:PartA_GEN2>"
writeXML "<ITRForm:AuditInfo>"

'LiableSec44AAflg
If Sheet1.Range("sheet1.LiableSec44AAflg").value <> "" And Sheet1.Range("sheet1.LiableSec44AAflg").value <> "(Select)" Then
'writeXML "<ITRForm:LiableSec44AAflg>" & "" & "</ITRForm:LiableSec44AAflg>"
'Else
writeXML "<ITRForm:LiableSec44AAflg>" & Mid(UCase(Sheet1.Range("sheet1.LiableSec44AAflg").value), 1, 1) & "</ITRForm:LiableSec44AAflg>"
End If

If Sheet1.Range("sheet1.LiableSec44AEflg").value <> "" And Sheet1.Range("sheet1.LiableSec44AEflg").value <> "(Select)" Then
'writeXML "<ITRForm:IncDclrdUs>" & "" & "</ITRForm:IncDclrdUs>"
'Else
writeXML "<ITRForm:IncDclrdUs>" & Mid(UCase(Sheet1.Range("sheet1.LiableSec44AEflg").value), 1, 1) & "</ITRForm:IncDclrdUs>"
End If

If Sheet1.Range("sheet1.LiableSec5cflg").value <> "" And Sheet1.Range("sheet1.LiableSec5cflg").value <> "(Select)" Then
'writeXML "<ITRForm:TotalSalesExcOneCr>" & "" & "</ITRForm:TotalSalesExcOneCr>"
'Else
writeXML "<ITRForm:TotalSalesExcOneCr>" & Mid(UCase(Sheet1.Range("sheet1.LiableSec5cflg").value), 1, 1) & "</ITRForm:TotalSalesExcOneCr>"
End If

If Sheet1.Range("sheet1.LiableSecAggrigateflg").value <> "" And Sheet1.Range("sheet1.LiableSecAggrigateflg").value <> "(Select)" Then
'writeXML "<ITRForm:AgrOFAllAmtsRcvd>" & "" & "</ITRForm:AgrOFAllAmtsRcvd>"
'Else
writeXML "<ITRForm:AgrOFAllAmtsRcvd>" & Mid(UCase(Sheet1.Range("sheet1.LiableSecAggrigateflg").value), 1, 1) & "</ITRForm:AgrOFAllAmtsRcvd>"
End If

If Sheet1.Range("sheet1.LiableSecAggrigate1flg").value <> "" And Sheet1.Range("sheet1.LiableSecAggrigate1flg").value <> "(Select)" Then
'writeXML "<ITRForm:AgrOFAllPayMade>" & "" & "</ITRForm:AgrOFAllPayMade>"
'Else
writeXML "<ITRForm:AgrOFAllPayMade>" & Mid(UCase(Sheet1.Range("sheet1.LiableSecAggrigate1flg").value), 1, 1) & "</ITRForm:AgrOFAllPayMade>"
End If

If Sheet1.Range("sheet1.LiableSec44ABflg").value <> "" And Sheet1.Range("sheet1.LiableSec44ABflg").value <> "(Select)" Then
writeXML "<ITRForm:LiableSec44ABflg>" & Mid(UCase(Sheet1.Range("sheet1.LiableSec44ABflg").value), 1, 1) & "</ITRForm:LiableSec44ABflg>"
Else
writeXML "<ITRForm:LiableSec44ABflg>" & "N" & "</ITRForm:LiableSec44ABflg>"
End If

'AuditAccountantFlg
If Mid(Sheet1.Range("sheet1.LiableSec44ABflg").value, 1, 1) <> "(" And Mid(Sheet1.Range("sheet1.LiableSec44ABflg").value, 1, 1) <> "N" Then
writeXML "<ITRForm:AuditAccountantFlg>" & Mid(UCase(Sheet1.Range("sheet1.AuditAccountantFlg").value), 1, 1) & "</ITRForm:AuditAccountantFlg>"

If Trim(Mid(UCase(Sheet1.Range("sheet1.AuditAccountantFlg").value), 1, InStr(1, UCase(Sheet1.Range("sheet1.AuditAccountantFlg").value), "-") - 1)) = "Y" Then
'AuditReportFurnishDate
If Sheet1.Range("sheet1.AuditReportFurnishDate").value <> "" Then
'writeXML "<ITRForm:AuditReportFurnishDate>" & "" & "</ITRForm:AuditReportFurnishDate>"
'Else
writeXML "<ITRForm:AuditReportFurnishDate>" & Dformat(UCase(Sheet1.Range("sheet1.AuditReportFurnishDate").value), "") & "</ITRForm:AuditReportFurnishDate>"
End If

'AuditorName
If Sheet1.Range("sheet1.AuditorName").value <> "" Then
'writeXML "<ITRForm:AuditorName>" & "" & "</ITRForm:AuditorName>"
'Else
writeXML "<ITRForm:AuditorName>" & UCase(Sheet1.Range("sheet1.AuditorName").value) & "</ITRForm:AuditorName>"
End If

'AuditorMemNo
If Sheet1.Range("sheet1.AuditorMemNo").value <> "" Then
'writeXML "<ITRForm:AuditorMemNo>" & "" & "</ITRForm:AuditorMemNo>"
'Else
writeXML "<ITRForm:AuditorMemNo>" & UCase(Sheet1.Range("sheet1.AuditorMemNo").value) & "</ITRForm:AuditorMemNo>"
End If

'AudFrmName
If Sheet1.Range("sheet1.AudFrmName").value = "" Then
'writeXML "<ITRForm:AudFrmName>" & "" & "</ITRForm:AudFrmName>"
Else
writeXML "<ITRForm:AudFrmName>" & UCase(Sheet1.Range("sheet1.AudFrmName").value) & "</ITRForm:AudFrmName>"
End If

'AudFrmRegNo
Dim FrmRegNo As Variant

FrmRegNo = Sheet1.Range("sheet1.AudFrmRegNo").value


If Len(FrmRegNo) = 7 Then
    FrmRegNo = "0" & FrmRegNo
End If

If Len(FrmRegNo) = 6 Then
    FrmRegNo = "00" & FrmRegNo
End If

If Len(FrmRegNo) = 5 Then
    FrmRegNo = "000" & FrmRegNo
End If

If Len(FrmRegNo) = 4 Then
    FrmRegNo = "0000" & FrmRegNo
End If

If Len(FrmRegNo) = 3 Then
    FrmRegNo = "00000" & FrmRegNo
End If

If Len(FrmRegNo) = 2 Then
    FrmRegNo = "000000" & FrmRegNo
End If

If Len(FrmRegNo) = 1 Then
    FrmRegNo = "0000000" & FrmRegNo
End If

If FrmRegNo = "" Then
    'writeXML "<ITRForm:AudFrmRegNo>" & "" & "</ITRForm:AudFrmRegNo>"
Else
    writeXML "<ITRForm:AudFrmRegNo>" & UCase(FrmRegNo) & "</ITRForm:AudFrmRegNo>"
End If

'AudFrmPAN
If Sheet1.Range("sheet1.AudFrmPAN").value <> "" Then
'writeXML "<ITRForm:AudFrmPAN>" & "" & "</ITRForm:AudFrmPAN>"
'Else
writeXML "<ITRForm:AudFrmPAN>" & UCase(Sheet1.Range("sheet1.AudFrmPAN").value) & "</ITRForm:AudFrmPAN>"
End If

If Sheet1.Range("sheet1.AudFrmAadhaar").value <> "" Then
'writeXML "<ITRForm:AudFrmPAN>" & "" & "</ITRForm:AudFrmPAN>"
'Else
writeXML "<ITRForm:AudFrmAadhaar>" & UCase(Sheet1.Range("sheet1.AudFrmAadhaar").value) & "</ITRForm:AudFrmAadhaar>"
End If

'AuditDate
If Sheet1.Range("sheet1.AuditDate").value <> "" Then
'writeXML "<ITRForm:AuditDate>" & "" & "</ITRForm:AuditDate>"
'Else
writeXML "<ITRForm:AuditDate>" & Dformat((Sheet1.Range("sheet1.AuditDate").value), "") & "</ITRForm:AuditDate>"
End If
End If
End If


'92E
    If LiableSec92Eflg = "" Then
        writeXML " <ITRForm:LiableSec92Eflg>N</ITRForm:LiableSec92Eflg>"
    Else
        writeXML " <ITRForm:LiableSec92Eflg>" & UCase(LiableSec92Eflg) & "</ITRForm:LiableSec92Eflg>"
    End If
    
 If isdropdownblank(LiableSec92Eflgdia) Then
        'writeXML " <ITRForm:AccountAuditFlag>N</ITRForm:AccountAuditFlag>"
    Else
        writeXML " <ITRForm:AccountAuditFlag>" & Mid(UCase(LiableSec92Eflgdia), 1, 1) & "</ITRForm:AccountAuditFlag>"
    End If

    If LiableSec92Eflgdia = "Y" And AuditDateSec92E <> "" Then
        writeXML "          <ITRForm:AuditDetails92E>"
        writeXML "          <ITRForm:DateOfAudit>" & AuditDateSec92E & "</ITRForm:DateOfAudit>"
        writeXML "          </ITRForm:AuditDetails92E>"
    End If

'Dim i As Long
If Not IsEmpty(SecnCode_General) Then

For i = 1 To UBound(SecnCode_General)
writeXML "<ITRForm:AuditDetails>"
  If SecnCode_General(i) <> "" Then
  writeXML "<ITRForm:AuditedSection>" & (SecnCode_General(i)) & "</ITRForm:AuditedSection>"
'  Else
'  writeXML "<ITRForm:AuditedSection/>"
  End If
  
  If Mid(SecnCodeYN_General(i), 1, 1) <> "" And Mid(SecnCodeYN_General(i), 1, 1) <> "(" Then
          writeXML " <ITRForm:AuditFlag>" & Mid(SecnCodeYN_General(i), 1, 1) & "</ITRForm:AuditFlag>"
    End If
    

'If Mid(SecnCodeYN_General(i), 1, 1) = "Y" Then
  If SecnDate_General(i) <> "" Then
  writeXML "<ITRForm:DateOfAudit>" & Dformat(SecnDate_General(i), "") & "</ITRForm:DateOfAudit>"
'  Else
'  writeXML "<ITRForm:DateOfAudit/>"
  End If
'  End If
  writeXML "</ITRForm:AuditDetails>"
Next
End If


If Not IsEmpty(AuditAct_Info) Then
        If Not IsEmpty(AuditAct_Info) And UBound(AuditAct_Info) > 0 Then
            For i = 1 To UBound(AuditAct_Info)
                If AuditSection_Info(i) <> "" And AuditAct_Info(i) <> "" Then
                    writeXML "<ITRForm:AuditReportDetails>"

                    If AuditAct_Info(i) = "Banking Regulation Act, 1949" Then
                    AuditAct_Info(i) = "1"
                    ElseIf AuditAct_Info(i) = "Central Excise Act,1944" Then
                     AuditAct_Info(i) = "2"
                    ElseIf AuditAct_Info(i) = "Central Sales Tax Act, 1956" Then
                     AuditAct_Info(i) = "3"
                    ElseIf AuditAct_Info(i) = "Central Goods and Services Tax Act, 2017" Then
                     AuditAct_Info(i) = "4"
                    ElseIf AuditAct_Info(i) = "Charitable And Religious Trusts Act, 1920" Then
                     AuditAct_Info(i) = "5"
                     ElseIf AuditAct_Info(i) = "Companies Act, 2013" Then
                     AuditAct_Info(i) = "6"
                     ElseIf AuditAct_Info(i) = "Electricity Act, 2003" Then
                     AuditAct_Info(i) = "7"
                     ElseIf AuditAct_Info(i) = "Employees Provident Fund and Miscellaneous Provisions Act, 1952" Then
                     AuditAct_Info(i) = "8"
                     ElseIf AuditAct_Info(i) = "Foreign Exchange Management Act, 1999" Then
                     AuditAct_Info(i) = "9"
                     ElseIf AuditAct_Info(i) = "Government Superannuation Fund Act, 1956" Then
                     AuditAct_Info(i) = "10"
                     ElseIf AuditAct_Info(i) = "Indian Trusts Act, 1882" Then
                     AuditAct_Info(i) = "11"
                     ElseIf AuditAct_Info(i) = "Integrated Goods and Services Tax Act, 2017" Then
                     AuditAct_Info(i) = "12"
                     ElseIf AuditAct_Info(i) = "Limited Liability Partnership Act, 2008" Then
                     AuditAct_Info(i) = "13"
                     ElseIf AuditAct_Info(i) = "Payment of Gratuity Act, 1972" Then
                     AuditAct_Info(i) = "14"
                     ElseIf AuditAct_Info(i) = "SEBI Act, 1992" Then
                     AuditAct_Info(i) = "15"
                     ElseIf AuditAct_Info(i) = "Securities Contract (Regulation) Act, 1956" Then
                     AuditAct_Info(i) = "16"
                     ElseIf AuditAct_Info(i) = "State Goods and Services Tax Act, 2017" Then
                     AuditAct_Info(i) = "17"
                     ElseIf AuditAct_Info(i) = "Union Territories Goods and Services Tax Act, 2017" Then
                     AuditAct_Info(i) = "18"
                     ElseIf AuditAct_Info(i) = "Others" Then
                     AuditAct_Info(i) = "19"
                     End If

                    writeXML "<ITRForm:AuditReportAct>" & AuditAct_Info(i) & "</ITRForm:AuditReportAct>"

                    If AuditAct_Info_Others(i) <> "" Then
                    writeXML "<ITRForm:AuditReportActOthers>" & AuditAct_Info_Others(i) & "</ITRForm:AuditReportActOthers>"
                    End If


                    writeXML " <ITRForm:AuditedSection>" & AuditSection_Info(i) & "</ITRForm:AuditedSection>"
                    
                    If Mid(Audit_YN(i), 1, 1) <> "" And Mid(Audit_YN(i), 1, 1) <> "(" Then
                    writeXML " <ITRForm:OtherITActFlag>" & Mid(Audit_YN(i), 1, 1) & "</ITRForm:OtherITActFlag>"
                    End If
                    
                    
                    If Mid(Audit_YN(i), 1, 1) = "Y" Then
                    If AuditDate_Info(i) <> "" Then
                    writeXML "<ITRForm:DateOfAudit>" & Dformat(AuditDate_Info(i), "") & "</ITRForm:DateOfAudit>"
                    End If
                    End If
                    writeXML "</ITRForm:AuditReportDetails>"
                End If
            Next
        End If
    End If
    writeXML "</ITRForm:AuditInfo>"






If UBound(Code_NOB) > 0 Then
writeXML "<ITRForm:NatOfBus>"
For i = 1 To UBound(Code_NOB)
writeXML "<ITRForm:NatureOfBusiness>"
If Code_NOB(i) <> "" Then
  writeXML "<ITRForm:Code>" & Mid(UCase(Code_NOB(i)), 1, InStr(1, Code_NOB(i), "-") - 1) & "</ITRForm:Code>"
  Else

  End If

If TradeName_NOB(i) <> "" Or TradeName1_NOB(i) <> "" Or TradeName2_NOB(i) <> "" Then
'writeXML "<ITRForm:Trade>"


If TradeName_NOB(i) <> "" Then
  writeXML "<ITRForm:TradeName1>" & UCase(TradeName_NOB(i)) & "</ITRForm:TradeName1>"
  Else

  End If

If TradeName1_NOB(i) <> "" Then
  writeXML "<ITRForm:TradeName2>" & UCase(TradeName1_NOB(i)) & "</ITRForm:TradeName2>"
  Else

  End If

If TradeName2_NOB(i) <> "" Then
  writeXML "<ITRForm:TradeName3>" & UCase(TradeName2_NOB(i)) & "</ITRForm:TradeName3>"
  Else

  End If
  'writeXML "</ITRForm:Trade>"
  End If
writeXML "</ITRForm:NatureOfBusiness>"
Next
writeXML "</ITRForm:NatOfBus>"
End If



writeXML "</ITRForm:PartA_GEN2>"

End Function

Function BalanceSheetXML()

writeXML "<ITRForm:PARTA_BS>"

    writeXML "<ITRForm:FundSrc>"
      writeXML "<ITRForm:PropFund>"
      writeXML "<ITRForm:PropCap>" & Sheet2_PropCap & "</ITRForm:PropCap>"

        writeXML "<ITRForm:ResrNSurp>"
          writeXML "<ITRForm:RevResr>" & Sheet2_RevResr & "</ITRForm:RevResr>"
          writeXML "<ITRForm:CapResr>" & Sheet2_CapResr & "</ITRForm:CapResr>"
          writeXML "<ITRForm:StatResr>" & Sheet2_StatResr & "</ITRForm:StatResr>"
          writeXML "<ITRForm:OthResr>" & Sheet2_OthResr & "</ITRForm:OthResr>"
          writeXML "<ITRForm:TotResrNSurp>" & Sheet2_TotResrNSurp & "</ITRForm:TotResrNSurp>"
        writeXML "</ITRForm:ResrNSurp>"

        writeXML "<ITRForm:TotPropFund>" & Sheet2_TotPropFund & "</ITRForm:TotPropFund>"
      writeXML "</ITRForm:PropFund>"
      writeXML "<ITRForm:LoanFunds>"
        writeXML "<ITRForm:SecrLoan>"
          writeXML "<ITRForm:ForeignCurrLoan>" & Sheet2_ForeignCurrLoan & "</ITRForm:ForeignCurrLoan>"
          writeXML "<ITRForm:RupeeLoan>"
            writeXML "<ITRForm:FrmBank>" & Sheet2_FrmBank & "</ITRForm:FrmBank>"
            writeXML "<ITRForm:FrmOthrs>" & Sheet2_FrmOthrs & "</ITRForm:FrmOthrs>"
            writeXML "<ITRForm:TotRupeeLoan>" & Sheet2_TotRupeeLoan & "</ITRForm:TotRupeeLoan>"
          writeXML "</ITRForm:RupeeLoan>"
          writeXML "<ITRForm:TotSecrLoan>" & Sheet2_TotSecrLoan & "</ITRForm:TotSecrLoan>"
        writeXML "</ITRForm:SecrLoan>"
        writeXML "<ITRForm:UnsecrLoan>"
          writeXML "<ITRForm:FrmBank>" & Sheet2_FrmBank1 & "</ITRForm:FrmBank>"
          writeXML "<ITRForm:FrmOthrs>" & Sheet2_FrmOthrs1 & "</ITRForm:FrmOthrs>"
          writeXML "<ITRForm:TotUnSecrLoan>" & Sheet2_TotUnSecrLoan & "</ITRForm:TotUnSecrLoan>"
        writeXML "</ITRForm:UnsecrLoan>"
        writeXML "<ITRForm:TotLoanFund>" & Sheet2_TotLoanFund & "</ITRForm:TotLoanFund>"
      writeXML "</ITRForm:LoanFunds>"
      writeXML "<ITRForm:DeferredTax>" & Sheet2_DeferredTax & "</ITRForm:DeferredTax>"
      writeXML "<ITRForm:TotFundSrc>" & Sheet2_TotFundSrc & "</ITRForm:TotFundSrc>"
    writeXML "</ITRForm:FundSrc>"
    writeXML "<ITRForm:FundApply>"
      writeXML "<ITRForm:FixedAsset>"
        writeXML "<ITRForm:GrossBlock>" & Sheet2_GrossBlock & "</ITRForm:GrossBlock>"
        writeXML "<ITRForm:Depreciation>" & Sheet2_Depreciation & "</ITRForm:Depreciation>"
        writeXML "<ITRForm:NetBlock>" & Sheet2_NetBlock & "</ITRForm:NetBlock>"
        writeXML "<ITRForm:CapWrkProg>" & Sheet2_CapWrkProg & "</ITRForm:CapWrkProg>"
        writeXML "<ITRForm:TotFixedAsset>" & Sheet2_TotFixedAsset & "</ITRForm:TotFixedAsset>"
      writeXML "</ITRForm:FixedAsset>"
      writeXML "<ITRForm:Investments>"
        writeXML "<ITRForm:LongTermInv>"
          writeXML "<ITRForm:GovtOthSecQuoted>" & Sheet2_GovtOthSecQuoted & "</ITRForm:GovtOthSecQuoted>"
          writeXML "<ITRForm:GovOthSecUnQoted>" & Sheet2_GovOthSecUnQoted & "</ITRForm:GovOthSecUnQoted>"
          writeXML "<ITRForm:TotLongTermInv>" & Sheet2_TotLongTermInv & "</ITRForm:TotLongTermInv>"
        writeXML "</ITRForm:LongTermInv>"
        writeXML "<ITRForm:TradeInv>"
          writeXML "<ITRForm:EquityShares>" & Sheet2_EquityShares & "</ITRForm:EquityShares>"
          writeXML "<ITRForm:PreferShares>" & Sheet2_PreferShares & "</ITRForm:PreferShares>"
          writeXML "<ITRForm:Debenture>" & Sheet2_Debenture & "</ITRForm:Debenture>"
          writeXML "<ITRForm:TotTradeInv>" & Sheet2_TotTradeInv & "</ITRForm:TotTradeInv>"
        writeXML "</ITRForm:TradeInv>"
        writeXML "<ITRForm:TotInvestments>" & Sheet2_TotInvestments & "</ITRForm:TotInvestments>"
      writeXML "</ITRForm:Investments>"
      writeXML "<ITRForm:CurrAssetLoanAdv>"
        writeXML "<ITRForm:CurrAsset>"
          writeXML "<ITRForm:Inventories>"
            writeXML "<ITRForm:StoresConsumables>" & Sheet3_StoresConsumables & "</ITRForm:StoresConsumables>"
            writeXML "<ITRForm:RawMatl>" & Sheet3_RawMatl & "</ITRForm:RawMatl>"
            writeXML "<ITRForm:StkInProcess>" & Sheet3_StkInProcess & "</ITRForm:StkInProcess>"
            writeXML "<ITRForm:FinOrTradGood>" & Sheet3_FinOrTradGood & "</ITRForm:FinOrTradGood>"
            writeXML "<ITRForm:TotInventries>" & Sheet3_TotInventries & "</ITRForm:TotInventries>"
          writeXML "</ITRForm:Inventories>"
          writeXML "<ITRForm:SndryDebtors>" & Sheet3_SndryDebtors & "</ITRForm:SndryDebtors>"
          writeXML "<ITRForm:CashOrBankBal>"
            writeXML "<ITRForm:CashinHand>" & Sheet3_CashinHand & "</ITRForm:CashinHand>"
            writeXML "<ITRForm:BankBal>" & Sheet3_BankBal & "</ITRForm:BankBal>"
            writeXML "<ITRForm:TotCashOrBankBal>" & Sheet3_TotCashOrBankBal & "</ITRForm:TotCashOrBankBal>"
          writeXML "</ITRForm:CashOrBankBal>"
          writeXML "<ITRForm:OthCurrAsset>" & Sheet3_OthCurrAsset & "</ITRForm:OthCurrAsset>"
          writeXML "<ITRForm:TotCurrAsset>" & Sheet3_TotCurrAsset & "</ITRForm:TotCurrAsset>"
        writeXML "</ITRForm:CurrAsset>"
        writeXML "<ITRForm:LoanAdv>"
          writeXML "<ITRForm:AdvRecoverable>" & Sheet3_AdvRecoverable & "</ITRForm:AdvRecoverable>"
          writeXML "<ITRForm:Deposits>" & Sheet3_Deposits & "</ITRForm:Deposits>"
          writeXML "<ITRForm:BalWithRevAuth>" & Sheet3_BalWithRevAuth & "</ITRForm:BalWithRevAuth>"
          writeXML "<ITRForm:TotLoanAdv>" & Sheet3_TotLoanAdv & "</ITRForm:TotLoanAdv>"
        writeXML "</ITRForm:LoanAdv>"
        writeXML "<ITRForm:TotCurrAssetLoanAdv>" & Sheet3_TotCurrAssetLoanAdv & "</ITRForm:TotCurrAssetLoanAdv>"
        writeXML "<ITRForm:CurrLiabilitiesProv>"
          writeXML "<ITRForm:CurrLiabilities>"
            writeXML "<ITRForm:SundryCred>" & Sheet3_SundryCred & "</ITRForm:SundryCred>"
            writeXML "<ITRForm:LiabForLeasedAsset>" & Sheet3_LiabForLeasedAsset & "</ITRForm:LiabForLeasedAsset>"
            writeXML "<ITRForm:AccrIntonLeasedAsset>" & Sheet3_AccrIntonLeasedAsset & "</ITRForm:AccrIntonLeasedAsset>"
            writeXML "<ITRForm:AccrIntNotDue>" & Sheet3_AccrIntNotDue & "</ITRForm:AccrIntNotDue>"
            writeXML "<ITRForm:TotCurrLiabilities>" & Sheet3_TotCurrLiabilities & "</ITRForm:TotCurrLiabilities>"
          writeXML "</ITRForm:CurrLiabilities>"
          writeXML "<ITRForm:Provisions>"
            writeXML "<ITRForm:ITProvision>" & Sheet3_ITProvision & "</ITRForm:ITProvision>"
            'writeXML "<ITRForm:WTProvision>" & Sheet3_WTProvision & "</ITRForm:WTProvision>"
            writeXML "<ITRForm:ELSuperAnnGratProvision>" & Sheet3_ELSuperAnnGratProvision & "</ITRForm:ELSuperAnnGratProvision>"
            writeXML "<ITRForm:OthProvision>" & Sheet3_OthProvision & "</ITRForm:OthProvision>"
            writeXML "<ITRForm:TotProvisions>" & Sheet3_TotProvisions & "</ITRForm:TotProvisions>"
          writeXML "</ITRForm:Provisions>"
          writeXML "<ITRForm:TotCurrLiabilitiesProvision>" & Sheet3_TotCurrLiabilitiesProvision & "</ITRForm:TotCurrLiabilitiesProvision>"
        writeXML "</ITRForm:CurrLiabilitiesProv>"
        writeXML "<ITRForm:NetCurrAsset>" & Sheet3_NetCurrAsset & "</ITRForm:NetCurrAsset>"
      writeXML "</ITRForm:CurrAssetLoanAdv>"
      writeXML "<ITRForm:MiscAdjust>"
        writeXML "<ITRForm:MiscExpndr>" & Sheet3_MiscExpndr & "</ITRForm:MiscExpndr>"
        writeXML "<ITRForm:DefTaxAsset>" & Sheet3_DefTaxAsset & "</ITRForm:DefTaxAsset>"
        writeXML "<ITRForm:AccumaltedLosses>" & Sheet3_AccumaltedLosses & "</ITRForm:AccumaltedLosses>"
        writeXML "<ITRForm:TotMiscAdjust>" & Sheet3_TotMiscAdjust & "</ITRForm:TotMiscAdjust>"
      writeXML "</ITRForm:MiscAdjust>"
      writeXML "<ITRForm:TotFundApply>" & Sheet3_TotFundApply & "</ITRForm:TotFundApply>"
    writeXML "</ITRForm:FundApply>"
    writeXML "<ITRForm:NoBooksOfAccBS>"
      writeXML "<ITRForm:TotSundryDbtAmt>" & Sheet3_TotSundryDbtAmt & "</ITRForm:TotSundryDbtAmt>"
      writeXML "<ITRForm:TotSundryCrdAmt>" & Sheet3_TotSundryCrdAmt & "</ITRForm:TotSundryCrdAmt>"
      writeXML "<ITRForm:TotStkInTradAmt>" & Sheet3_TotStkInTradAmt & "</ITRForm:TotStkInTradAmt>"
      writeXML "<ITRForm:CashBalAmt>" & Sheet3_CashBalAmt & "</ITRForm:CashBalAmt>"
    writeXML "</ITRForm:NoBooksOfAccBS>"
  writeXML "</ITRForm:PARTA_BS>"

End Function

Function ManufacturingAccountSheetXML()
writeXML "<ITRForm:ManufacturingAccount>"
writeXML "<ITRForm:OpeningInventory>"

If Sheet48.Range("ManuFactureAcc_OpenStockRawMaterial").value <> "" Then
writeXML "<ITRForm:OpngStckRawMat>" & Sheet48.Range("ManuFactureAcc_OpenStockRawMaterial").value & "</ITRForm:OpngStckRawMat>"
End If

If Sheet48.Range("ManuFactureAcc_OpenStockWorkProgress").value <> "" Then
writeXML "<ITRForm:OpngStckWrkinPrgrs>" & Sheet48.Range("ManuFactureAcc_OpenStockWorkProgress").value & "</ITRForm:OpngStckWrkinPrgrs>"
End If

If Sheet48.Range("ManuFactureAcc_TotalOpeningInventary").value <> "" Then
writeXML "<ITRForm:OpngInvntryTotal>" & Sheet48.Range("ManuFactureAcc_TotalOpeningInventary").value & "</ITRForm:OpngInvntryTotal>"
End If

If Sheet48.Range("ManuFactureAcc_Purchases").value <> "" Then
writeXML "<ITRForm:Purchases>" & Sheet48.Range("ManuFactureAcc_Purchases").value & "</ITRForm:Purchases>"
End If

If Sheet48.Range("ManuFactureAcc_DirectWages").value <> "" Then
writeXML "<ITRForm:DirectWages>" & Sheet48.Range("ManuFactureAcc_DirectWages").value & "</ITRForm:DirectWages>"
End If

If Sheet48.Range("ManuFactureAcc_TotalOfDirectExpenses").value <> "" Then
writeXML "<ITRForm:DirectExpenses>" & Sheet48.Range("ManuFactureAcc_TotalOfDirectExpenses").value & "</ITRForm:DirectExpenses>"
End If

If Sheet48.Range("ManuFactureAcc_CarriageInward").value <> "" Then
writeXML "<ITRForm:CarriageInward>" & Sheet48.Range("ManuFactureAcc_CarriageInward").value & "</ITRForm:CarriageInward>"
End If

If Sheet48.Range("ManuFactureAcc_PowerAndFuel").value <> "" Then
writeXML "<ITRForm:PowerAndFuel>" & Sheet48.Range("ManuFactureAcc_PowerAndFuel").value & "</ITRForm:PowerAndFuel>"
End If

If Sheet48.Range("ManuFactureAcc_OtherDirectExpenses").value <> "" Then
writeXML "<ITRForm:OthDirectExpenses>" & Sheet48.Range("ManuFactureAcc_OtherDirectExpenses").value & "</ITRForm:OthDirectExpenses>"
End If



If Sheet48.Range("ManuFactureAcc_IndirectWages").value <> "" Then
writeXML "<ITRForm:IndirectWages>" & Sheet48.Range("ManuFactureAcc_IndirectWages").value & "</ITRForm:IndirectWages>"
End If


If Sheet48.Range("ManuFactureAcc_FactoryRentRates").value <> "" Then
writeXML "<ITRForm:FactoryRentAndRates>" & Sheet48.Range("ManuFactureAcc_FactoryRentRates").value & "</ITRForm:FactoryRentAndRates>"
End If


If Sheet48.Range("ManuFactureAcc_FactoryInsurance").value <> "" Then
writeXML "<ITRForm:FactoryInsurance>" & Sheet48.Range("ManuFactureAcc_FactoryInsurance").value & "</ITRForm:FactoryInsurance>"
End If


If Sheet48.Range("ManuFactureAcc_FactoryFuelAndPower").value <> "" Then
writeXML "<ITRForm:FactoryFuelAndPower>" & Sheet48.Range("ManuFactureAcc_FactoryFuelAndPower").value & "</ITRForm:FactoryFuelAndPower>"
End If


If Sheet48.Range("ManuFactureAcc_FactoryGeneralExpenses").value <> "" Then
writeXML "<ITRForm:FactoryGeneralExpenses>" & Sheet48.Range("ManuFactureAcc_FactoryGeneralExpenses").value & "</ITRForm:FactoryGeneralExpenses>"
End If


If Sheet48.Range("ManuFactureAcc_DepreciationOfFactoryMachinary").value <> "" Then
writeXML "<ITRForm:DeprctnOfFactoryMachinery>" & Sheet48.Range("ManuFactureAcc_DepreciationOfFactoryMachinary").value & "</ITRForm:DeprctnOfFactoryMachinery>"
End If


If Sheet48.Range("ManuFactureAcc_TotalFactoryOverheads").value <> "" Then
writeXML "<ITRForm:TotalFactoryOverheads>" & Sheet48.Range("ManuFactureAcc_TotalFactoryOverheads").value & "</ITRForm:TotalFactoryOverheads>"
End If


If Sheet48.Range("ManuFactureAcc_TotalOfDebitstoManuFacturingAcc").value <> "" Then
writeXML "<ITRForm:TotalDebtsManfctrngAcc>" & Sheet48.Range("ManuFactureAcc_TotalOfDebitstoManuFacturingAcc").value & "</ITRForm:TotalDebtsManfctrngAcc>"
End If
writeXML "</ITRForm:OpeningInventory>"

writeXML "<ITRForm:ClosingStock>"

If Sheet48.Range("ManuFactureAcc_RawMaterial").value <> "" Then
writeXML "<ITRForm:ClsngStckRawMaterial>" & Sheet48.Range("ManuFactureAcc_RawMaterial").value & "</ITRForm:ClsngStckRawMaterial>"
End If

If Sheet48.Range("ManuFactureAcc_WorkInProgress").value <> "" Then
writeXML "<ITRForm:ClsngStckWrkInPrgrs>" & Sheet48.Range("ManuFactureAcc_WorkInProgress").value & "</ITRForm:ClsngStckWrkInPrgrs>"
End If

If Sheet48.Range("ManuFactureAcc_TotalClosingClock").value <> "" Then
writeXML "<ITRForm:ClsngStckTotal>" & Sheet48.Range("ManuFactureAcc_TotalClosingClock").value & "</ITRForm:ClsngStckTotal>"
End If

writeXML "</ITRForm:ClosingStock>"

If Sheet48.Range("ManuFactureAcc_CostOfGoodsProduced").value <> "" Then
writeXML "<ITRForm:CostOfGoodsPrdcd>" & Sheet48.Range("ManuFactureAcc_CostOfGoodsProduced").value & "</ITRForm:CostOfGoodsPrdcd>"
End If

writeXML "</ITRForm:ManufacturingAccount>"
End Function

Function TradingAccountSheetXML()
writeXML "<ITRForm:TradingAccount>"

If Sheet49.Range("TradingAcc_SalesOfGoods").value <> "" Then
writeXML "<ITRForm:SaleOfGoods>" & Sheet49.Range("TradingAcc_SalesOfGoods").value & "</ITRForm:SaleOfGoods>"
End If

If Sheet49.Range("TradingAcc_SalesOfServices").value <> "" Then
writeXML "<ITRForm:SaleOfServices>" & Sheet49.Range("TradingAcc_SalesOfServices").value & "</ITRForm:SaleOfServices>"
End If


If Not IsEmpty(TradingNature1) Then
Dim i As Long
For i = 1 To UBound(TradingNature1)
writeXML "<ITRForm:OtherOperatingRevenueDtls>"
  If TradingNature1(i) <> "" Then
  writeXML "<ITRForm:OperatingRevenueName>" & (TradingNature1(i)) & "</ITRForm:OperatingRevenueName>"
  End If

  If TradingAmount1(i) <> "" Then
  writeXML "<ITRForm:OperatingRevenueAmt>" & (TradingAmount1(i)) & "</ITRForm:OperatingRevenueAmt>"
  End If

  writeXML "</ITRForm:OtherOperatingRevenueDtls>"
Next
End If

If Sheet49.Range("TradingAcc_RevenueTotal").value <> "" Then
writeXML "<ITRForm:OperatingRevenueTotal>" & Sheet49.Range("TradingAcc_RevenueTotal").value & "</ITRForm:OperatingRevenueTotal>"
End If

If Sheet49.Range("TradingAcc_TotalOfSalesOrGross").value <> "" Then
writeXML "<ITRForm:SalesGrossReceiptsTotal>" & Sheet49.Range("TradingAcc_TotalOfSalesOrGross").value & "</ITRForm:SalesGrossReceiptsTotal>"
End If

If Sheet49.Range("TradingAcc_GrossRecieptsFromProfession").value <> "" Then
writeXML "<ITRForm:GrossRcptFromProfession>" & Sheet49.Range("TradingAcc_GrossRecieptsFromProfession").value & "</ITRForm:GrossRcptFromProfession>"
Else
writeXML "<ITRForm:GrossRcptFromProfession>" & 0 & "</ITRForm:GrossRcptFromProfession>"
End If

writeXML "<ITRForm:ExciseCustomsVAT>"
If Sheet49.Range("TradingAcc_UnionExciseDuties").value <> "" Then
writeXML "<ITRForm:UnionExciseDuty>" & Sheet49.Range("TradingAcc_UnionExciseDuties").value & "</ITRForm:UnionExciseDuty>"
Else
writeXML "<ITRForm:UnionExciseDuty>" & 0 & "</ITRForm:UnionExciseDuty>"
End If

If Sheet49.Range("TradingAcc_ServiceTax").value <> "" Then
writeXML "<ITRForm:ServiceTax>" & Sheet49.Range("TradingAcc_ServiceTax").value & "</ITRForm:ServiceTax>"
Else
writeXML "<ITRForm:ServiceTax>" & 0 & "</ITRForm:ServiceTax>"
End If

If Sheet49.Range("TradingAcc_VatOrSalesTax").value <> "" Then
writeXML "<ITRForm:VATorSaleTax>" & Sheet49.Range("TradingAcc_VatOrSalesTax").value & "</ITRForm:VATorSaleTax>"
Else
writeXML "<ITRForm:VATorSaleTax>" & 0 & "</ITRForm:VATorSaleTax>"
End If

If Sheet49.Range("TradingAcc_CGST").value <> "" Then
writeXML "<ITRForm:CentralGoodServiceTax>" & Sheet49.Range("TradingAcc_CGST").value & "</ITRForm:CentralGoodServiceTax>"
Else
writeXML "<ITRForm:CentralGoodServiceTax>" & 0 & "</ITRForm:CentralGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_SGST").value <> "" Then
writeXML "<ITRForm:StateGoodServiceTax>" & Sheet49.Range("TradingAcc_SGST").value & "</ITRForm:StateGoodServiceTax>"
Else
writeXML "<ITRForm:StateGoodServiceTax>" & 0 & "</ITRForm:StateGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_IGST").value <> "" Then
writeXML "<ITRForm:IntegratedGoodServiceTax>" & Sheet49.Range("TradingAcc_IGST").value & "</ITRForm:IntegratedGoodServiceTax>"
Else
writeXML "<ITRForm:IntegratedGoodServiceTax>" & 0 & "</ITRForm:IntegratedGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_UTGST").value <> "" Then
writeXML "<ITRForm:UnionTerrGoodServiceTax>" & Sheet49.Range("TradingAcc_UTGST").value & "</ITRForm:UnionTerrGoodServiceTax>"
Else
writeXML "<ITRForm:UnionTerrGoodServiceTax>" & 0 & "</ITRForm:UnionTerrGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_AnyOtherDuty").value <> "" Then
writeXML "<ITRForm:OthDutyTaxCess>" & Sheet49.Range("TradingAcc_AnyOtherDuty").value & "</ITRForm:OthDutyTaxCess>"
Else
writeXML "<ITRForm:OthDutyTaxCess>" & 0 & "</ITRForm:OthDutyTaxCess>"
End If

If Sheet49.Range("TradingAcc_TotalDutiesTaxesCess").value <> "" Then
writeXML "<ITRForm:TotExciseCustomsVAT>" & Sheet49.Range("TradingAcc_TotalDutiesTaxesCess").value & "</ITRForm:TotExciseCustomsVAT>"
End If

writeXML "</ITRForm:ExciseCustomsVAT>"

If Sheet49.Range("TradingAcc_TotalRevenueFromOperations").value <> "" Then
writeXML "<ITRForm:TotRevenueFrmOperations>" & Sheet49.Range("TradingAcc_TotalRevenueFromOperations").value & "</ITRForm:TotRevenueFrmOperations>"
End If

If Sheet49.Range("TradingAcc_ClosingStockOfFinishedStocks").value <> "" Then
writeXML "<ITRForm:ClsngStckOfFinishedStcks>" & Sheet49.Range("TradingAcc_ClosingStockOfFinishedStocks").value & "</ITRForm:ClsngStckOfFinishedStcks>"
End If

If Sheet49.Range("TradingAcc_TotalOfCreditsToTradingAccount").value <> "" Then
writeXML "<ITRForm:TardingAccTotCred>" & Sheet49.Range("TradingAcc_TotalOfCreditsToTradingAccount").value & "</ITRForm:TardingAccTotCred>"
End If

If Sheet49.Range("TradingAcc_OpeningStockOfFinishedGoods").value <> "" Then
writeXML "<ITRForm:OpngStckOfFinishedStcks>" & Sheet49.Range("TradingAcc_OpeningStockOfFinishedGoods").value & "</ITRForm:OpngStckOfFinishedStcks>"
Else
writeXML "<ITRForm:OpngStckOfFinishedStcks>" & 0 & "</ITRForm:OpngStckOfFinishedStcks>"
End If

If Sheet49.Range("TradingAcc_Purchases").value <> "" Then
writeXML "<ITRForm:Purchases>" & Sheet49.Range("TradingAcc_Purchases").value & "</ITRForm:Purchases>"
Else
writeXML "<ITRForm:Purchases>" & 0 & "</ITRForm:Purchases>"
End If

If Sheet49.Range("TradingAcc_TotalDirectExpenses").value <> "" Then
writeXML "<ITRForm:DirectExpenses>" & Sheet49.Range("TradingAcc_TotalDirectExpenses").value & "</ITRForm:DirectExpenses>"
Else
writeXML "<ITRForm:DirectExpenses>" & 0 & "</ITRForm:DirectExpenses>"
End If


If Sheet49.Range("TradingAcc_CarriageInward").value <> "" Then
writeXML "<ITRForm:CarriageInward>" & Sheet49.Range("TradingAcc_CarriageInward").value & "</ITRForm:CarriageInward>"
Else
writeXML "<ITRForm:CarriageInward>" & 0 & "</ITRForm:CarriageInward>"
End If

If Sheet49.Range("TradingAcc_PowerAndFuel").value <> "" Then
writeXML "<ITRForm:PowerAndFuel>" & Sheet49.Range("TradingAcc_PowerAndFuel").value & "</ITRForm:PowerAndFuel>"
Else
writeXML "<ITRForm:PowerAndFuel>" & 0 & "</ITRForm:PowerAndFuel>"
End If

If Not IsEmpty(TradingNature2) Then

For i = 1 To UBound(TradingNature2)
writeXML "<ITRForm:OtherIncDtls>"
  If TradingNature2(i) <> "" Then
  writeXML "<ITRForm:NatureOfIncome>" & (TradingNature2(i)) & "</ITRForm:NatureOfIncome>"
  End If

  If TradingAmount2(i) <> "" Then
  writeXML "<ITRForm:Amount>" & (TradingAmount2(i)) & "</ITRForm:Amount>"
  End If

  writeXML "</ITRForm:OtherIncDtls>"
Next

End If
writeXML "<ITRForm:DirectExpensesTotal>" & (Sheet49.Range("TD.OthDE").value) & "</ITRForm:DirectExpensesTotal>"

writeXML "<ITRForm:DutyTaxPay>"
writeXML "<ITRForm:ExciseCustomsVAT>"

If Sheet49.Range("TradingAcc_CustomDuty").value <> "" Then
writeXML "<ITRForm:CustomDuty>" & Sheet49.Range("TradingAcc_CustomDuty").value & "</ITRForm:CustomDuty>"
End If

If Sheet49.Range("TradingAcc_CounterVeilingDuty").value <> "" Then
writeXML "<ITRForm:CounterVailDuty>" & Sheet49.Range("TradingAcc_CounterVeilingDuty").value & "</ITRForm:CounterVailDuty>"
End If

If Sheet49.Range("TradingAcc_SpecialAdditionalDuty").value <> "" Then
writeXML "<ITRForm:SplAddDuty>" & Sheet49.Range("TradingAcc_SpecialAdditionalDuty").value & "</ITRForm:SplAddDuty>"
End If

If Sheet49.Range("TradingAcc_UnionExciseDuty").value <> "" Then
writeXML "<ITRForm:UnionExciseDuty>" & Sheet49.Range("TradingAcc_UnionExciseDuty").value & "</ITRForm:UnionExciseDuty>"
End If

If Sheet49.Range("TradingAcc_DutiesServiceTax").value <> "" Then
writeXML "<ITRForm:ServiceTax>" & Sheet49.Range("TradingAcc_DutiesServiceTax").value & "</ITRForm:ServiceTax>"
End If

If Sheet49.Range("TradingAcc_DutiesVatOrSalesTax").value <> "" Then
writeXML "<ITRForm:VATorSaleTax>" & Sheet49.Range("TradingAcc_DutiesVatOrSalesTax").value & "</ITRForm:VATorSaleTax>"
End If

If Sheet49.Range("TradingAcc_DutiesCGST").value <> "" Then
writeXML "<ITRForm:CentralGoodServiceTax>" & Sheet49.Range("TradingAcc_DutiesCGST").value & "</ITRForm:CentralGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_DutiesSGST").value <> "" Then
writeXML "<ITRForm:StateGoodServiceTax>" & Sheet49.Range("TradingAcc_DutiesSGST").value & "</ITRForm:StateGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_DutiesIGST").value <> "" Then
writeXML "<ITRForm:IntegratedGoodServiceTax>" & Sheet49.Range("TradingAcc_DutiesIGST").value & "</ITRForm:IntegratedGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_DutiesUTGST").value <> "" Then
writeXML "<ITRForm:UnionTerrGoodServiceTax>" & Sheet49.Range("TradingAcc_DutiesUTGST").value & "</ITRForm:UnionTerrGoodServiceTax>"
End If

If Sheet49.Range("TradingAcc_DutiesAnyOtherTax").value <> "" Then
writeXML "<ITRForm:OthDutyTaxCess>" & Sheet49.Range("TradingAcc_DutiesAnyOtherTax").value & "</ITRForm:OthDutyTaxCess>"
End If

If Sheet49.Range("TradingAcc_TotalDuties").value <> "" Then
writeXML "<ITRForm:TotExciseCustomsVAT>" & Sheet49.Range("TradingAcc_TotalDuties").value & "</ITRForm:TotExciseCustomsVAT>"
End If

writeXML "</ITRForm:ExciseCustomsVAT>"
writeXML "</ITRForm:DutyTaxPay>"


If Sheet49.Range("TradingAcc_CostOfGoods").value <> "" Then
writeXML "<ITRForm:GoodsCostPrdcdFrmMA>" & Sheet49.Range("TradingAcc_CostOfGoods").value & "</ITRForm:GoodsCostPrdcdFrmMA>"
End If

If Sheet49.Range("TradingAcc_GrossProfitOrLoss").value <> "" Then
writeXML "<ITRForm:GrossProfitFrmBusProf>" & Sheet49.Range("TradingAcc_GrossProfitOrLoss").value & "</ITRForm:GrossProfitFrmBusProf>"
End If
writeXML "</ITRForm:TradingAccount>"
End Function

Function PARTA_PL()

Dim i As Long

    writeXML "      <ITRForm:PARTA_PL>"
    writeXML "          <ITRForm:CreditsToPL>"


    If Sheet4.Range("PL_GrossProfitLoss").value <> "" Then
        writeXML "<ITRForm:GrossProfitTrnsfFrmTrdAcc>" & Sheet4.Range("PL_GrossProfitLoss").value & "</ITRForm:GrossProfitTrnsfFrmTrdAcc>"
    Else
       writeXML "<ITRForm:GrossProfitTrnsfFrmTrdAcc>0</ITRForm:GrossProfitTrnsfFrmTrdAcc>"
    End If


    writeXML "              <ITRForm:OthIncome>"
    If Sheet4.Range("PL.RentInc").value <> "" Then
        writeXML "                   <ITRForm:RentInc>" & UCase(Sheet4.Range("PL.RentInc").value) & "</ITRForm:RentInc>"
    Else
        writeXML "                   <ITRForm:RentInc>0</ITRForm:RentInc>"
    End If

    If Sheet4.Range("PL.Comissions").value <> "" Then
        writeXML "                   <ITRForm:Comissions>" & UCase(Sheet4.Range("PL.Comissions").value) & "</ITRForm:Comissions>"
    Else
        writeXML "                   <ITRForm:Comissions>0</ITRForm:Comissions>"
    End If

    If Sheet4.Range("PL.Dividends").value <> "" Then
        writeXML "                   <ITRForm:Dividends>" & UCase(Sheet4.Range("PL.Dividends").value) & "</ITRForm:Dividends>"
    Else
        writeXML "                   <ITRForm:Dividends>0</ITRForm:Dividends>"
    End If

    If Sheet4.Range("PL.InterestInc").value <> "" Then
        writeXML "                   <ITRForm:InterestInc>" & UCase(Sheet4.Range("PL.InterestInc").value) & "</ITRForm:InterestInc>"
    Else
        writeXML "                   <ITRForm:InterestInc>0</ITRForm:InterestInc>"
    End If

    If Sheet4.Range("PL.ProfitOnSaleFixedAsset").value <> "" Then
        writeXML "                   <ITRForm:ProfitOnSaleFixedAsset>" & UCase(Sheet4.Range("PL.ProfitOnSaleFixedAsset").value) & "</ITRForm:ProfitOnSaleFixedAsset>"
    Else
        writeXML "                   <ITRForm:ProfitOnSaleFixedAsset>0</ITRForm:ProfitOnSaleFixedAsset>"
    End If

    If Sheet4.Range("PL.ProfitOnInvChrSTT").value <> "" Then
        writeXML "                   <ITRForm:ProfitOnInvChrSTT>" & UCase(Sheet4.Range("PL.ProfitOnInvChrSTT").value) & "</ITRForm:ProfitOnInvChrSTT>"
    Else
        writeXML "                   <ITRForm:ProfitOnInvChrSTT>0</ITRForm:ProfitOnInvChrSTT>"
    End If

    If Sheet4.Range("PL.ProfitOnOthInv").value <> "" Then
        writeXML "                   <ITRForm:ProfitOnOthInv>" & UCase(Sheet4.Range("PL.ProfitOnOthInv").value) & "</ITRForm:ProfitOnOthInv>"
    Else
        writeXML "                   <ITRForm:ProfitOnOthInv>0</ITRForm:ProfitOnOthInv>"
    End If

    If Sheet4.Range("PL.ProfitOnCurrFluct").value <> "" Then
        writeXML "                   <ITRForm:ProfitOnCurrFluct>" & UCase(Sheet4.Range("PL.ProfitOnCurrFluct").value) & "</ITRForm:ProfitOnCurrFluct>"
    Else
        writeXML "                   <ITRForm:ProfitOnCurrFluct>0</ITRForm:ProfitOnCurrFluct>"
    End If

        If Sheet4.Range("PL.ProfitOnConvOfInventary").value <> "" Then
        writeXML "                   <ITRForm:ProfitOnCnvInvntryToCapAsst>" & UCase(Sheet4.Range("PL.ProfitOnConvOfInventary").value) & "</ITRForm:ProfitOnCnvInvntryToCapAsst>"
    Else
        writeXML "                   <ITRForm:ProfitOnCnvInvntryToCapAsst>0</ITRForm:ProfitOnCnvInvntryToCapAsst>"
    End If

    If Sheet4.Range("PL.ProfitOnAgriIncome").value <> "" Then
        writeXML "                   <ITRForm:ProfitOnAgriIncome>" & UCase(Sheet4.Range("PL.ProfitOnAgriIncome").value) & "</ITRForm:ProfitOnAgriIncome>"
    Else
        writeXML "                   <ITRForm:ProfitOnAgriIncome>0</ITRForm:ProfitOnAgriIncome>"
    End If

    If (Not IsEmpty(Nature2)) Then
        For i = 1 To UBound(Nature2)
            If amount2(i) <> "" Then
                If amount2(i) > 0 Then
                    writeXML "<ITRForm:OtherIncDtls>"
                    If Nature2(i) <> "" Then
                        writeXML "<ITRForm:NatureOfIncome>" & Nature2(i) & "</ITRForm:NatureOfIncome>"
                    Else
                        writeXML "<ITRForm:NatureOfIncome><ITRForm:NatureOfIncome>"
                    End If
                    writeXML "<ITRForm:Amount>" & amount2(i) & "</ITRForm:Amount>"
                    writeXML "</ITRForm:OtherIncDtls>"
                End If
            End If
        Next
    End If

    If Sheet4.Range("PL.MiscOthIncome").value <> "" Then
        writeXML "                   <ITRForm:MiscOthIncome>" & UCase(Sheet4.Range("PL.MiscOthIncome").value) & "</ITRForm:MiscOthIncome>"
    Else
        writeXML "                   <ITRForm:MiscOthIncome>0</ITRForm:MiscOthIncome>"
    End If

    If Sheet4.Range("PL.TotOthIncome").value <> "" Then
        writeXML "                   <ITRForm:TotOthIncome>" & UCase(Sheet4.Range("PL.TotOthIncome").value) & "</ITRForm:TotOthIncome>"
    Else
        writeXML "                   <ITRForm:TotOthIncome>0</ITRForm:TotOthIncome>"
    End If

    writeXML "              </ITRForm:OthIncome>"

           If Sheet4.Range("PL_TotalOfCreditsToProfit").value <> "" Then
        writeXML "                   <ITRForm:TotCreditsToPL>" & UCase(Sheet4.Range("PL_TotalOfCreditsToProfit").value) & "</ITRForm:TotCreditsToPL>"
    Else
        writeXML "                   <ITRForm:TotCreditsToPL>0</ITRForm:TotCreditsToPL>"
    End If


    writeXML "          </ITRForm:CreditsToPL>"



    writeXML "          <ITRForm:DebitsToPL>"

    If Sheet4.Range("PL.Freight").value <> "" Then
        writeXML "                   <ITRForm:Freight>" & UCase(Sheet4.Range("PL.Freight").value) & "</ITRForm:Freight>"
    Else
        writeXML "                   <ITRForm:Freight>0</ITRForm:Freight>"
    End If

    If Sheet4.Range("PL.ConsumptionOfStores").value <> "" Then
        writeXML "                   <ITRForm:ConsumptionOfStores>" & UCase(Sheet4.Range("PL.ConsumptionOfStores").value) & "</ITRForm:ConsumptionOfStores>"
    Else
        writeXML "                   <ITRForm:ConsumptionOfStores>0</ITRForm:ConsumptionOfStores>"
    End If

    If Sheet4.Range("PL.PowerFuel").value <> "" Then
        writeXML "                   <ITRForm:PowerFuel>" & UCase(Sheet4.Range("PL.PowerFuel").value) & "</ITRForm:PowerFuel>"
    Else
        writeXML "                   <ITRForm:PowerFuel>0</ITRForm:PowerFuel>"
    End If

    If Sheet4.Range("PL.RentExpdr").value <> "" Then
        writeXML "                   <ITRForm:RentExpdr>" & UCase(Sheet4.Range("PL.RentExpdr").value) & "</ITRForm:RentExpdr>"
    Else
        writeXML "                   <ITRForm:RentExpdr>0</ITRForm:RentExpdr>"
    End If

    If Sheet4.Range("PL.RepairsBldg").value <> "" Then
        writeXML "                   <ITRForm:RepairsBldg>" & UCase(Sheet4.Range("PL.RepairsBldg").value) & "</ITRForm:RepairsBldg>"
    Else
        writeXML "                   <ITRForm:RepairsBldg>0</ITRForm:RepairsBldg>"
    End If

    If Sheet4.Range("PL.RepairMach").value <> "" Then
        writeXML "                   <ITRForm:RepairMach>" & UCase(Sheet4.Range("PL.RepairMach").value) & "</ITRForm:RepairMach>"
    Else
        writeXML "                   <ITRForm:RepairMach>0</ITRForm:RepairMach>"
    End If

    writeXML "                  <ITRForm:EmployeeComp>"
    If Sheet4.Range("PL.SalsWages").value <> "" Then
        writeXML "                       <ITRForm:SalsWages>" & UCase(Sheet4.Range("PL.SalsWages").value) & "</ITRForm:SalsWages>"
    Else
        writeXML "                       <ITRForm:SalsWages>0</ITRForm:SalsWages>"
    End If

    If Sheet4.Range("PL.Bonus").value <> "" Then
        writeXML "                       <ITRForm:Bonus>" & UCase(Sheet4.Range("PL.Bonus").value) & "</ITRForm:Bonus>"
    Else
        writeXML "                       <ITRForm:Bonus>0</ITRForm:Bonus>"
    End If

    If Sheet4.Range("PL.MedExpReimb").value <> "" Then
        writeXML "                       <ITRForm:MedExpReimb>" & UCase(Sheet4.Range("PL.MedExpReimb").value) & "</ITRForm:MedExpReimb>"
    Else
        writeXML "                       <ITRForm:MedExpReimb>0</ITRForm:MedExpReimb>"
    End If

    If Sheet4.Range("PL.LeaveEncash").value <> "" Then
        writeXML "                       <ITRForm:LeaveEncash>" & UCase(Sheet4.Range("PL.LeaveEncash").value) & "</ITRForm:LeaveEncash>"
    Else
        writeXML "                       <ITRForm:LeaveEncash>0</ITRForm:LeaveEncash>"
    End If

    If Sheet4.Range("PL.LeaveTravelBenft").value <> "" Then
        writeXML "                       <ITRForm:LeaveTravelBenft>" & UCase(Sheet4.Range("PL.LeaveTravelBenft").value) & "</ITRForm:LeaveTravelBenft>"
    Else
        writeXML "                       <ITRForm:LeaveTravelBenft>0</ITRForm:LeaveTravelBenft>"
    End If

    If Sheet4.Range("PL.ContToSuperAnnFund").value <> "" Then
        writeXML "                       <ITRForm:ContToSuperAnnFund>" & UCase(Sheet4.Range("PL.ContToSuperAnnFund").value) & "</ITRForm:ContToSuperAnnFund>"
    Else
        writeXML "                       <ITRForm:ContToSuperAnnFund>0</ITRForm:ContToSuperAnnFund>"
    End If

    If Sheet4.Range("PL.ContToPF").value <> "" Then
        writeXML "                       <ITRForm:ContToPF>" & UCase(Sheet4.Range("PL.ContToPF").value) & "</ITRForm:ContToPF>"
    Else
        writeXML "                       <ITRForm:ContToPF>0</ITRForm:ContToPF>"
    End If

    If Sheet4.Range("PL.ContToGratFund").value <> "" Then
        writeXML "                       <ITRForm:ContToGratFund>" & UCase(Sheet4.Range("PL.ContToGratFund").value) & "</ITRForm:ContToGratFund>"
    Else
        writeXML "                       <ITRForm:ContToGratFund>0</ITRForm:ContToGratFund>"
    End If

    If Sheet4.Range("PL.ContToOthFund").value <> "" Then
        writeXML "                       <ITRForm:ContToOthFund>" & UCase(Sheet4.Range("PL.ContToOthFund").value) & "</ITRForm:ContToOthFund>"
    Else
        writeXML "                       <ITRForm:ContToOthFund>0</ITRForm:ContToOthFund>"
    End If

    If Sheet4.Range("PL.OthEmpBenftExpdr").value <> "" Then
        writeXML "                       <ITRForm:OthEmpBenftExpdr>" & UCase(Sheet4.Range("PL.OthEmpBenftExpdr").value) & "</ITRForm:OthEmpBenftExpdr>"
    Else
        writeXML "                       <ITRForm:OthEmpBenftExpdr>0</ITRForm:OthEmpBenftExpdr>"
    End If

    If Sheet4.Range("PL.TotEmployeeComp").value <> "" Then
        writeXML "                       <ITRForm:TotEmployeeComp>" & UCase(Sheet4.Range("PL.TotEmployeeComp").value) & "</ITRForm:TotEmployeeComp>"
    Else
        writeXML "                       <ITRForm:TotEmployeeComp>0</ITRForm:TotEmployeeComp>"
    End If

    If Sheet4.Range("PL.AnyCompPaidToNonRes").value <> "" Then
        writeXML "                       <ITRForm:AnyCompPaidToNonRes>" & Mid(UCase(Sheet4.Range("PL.AnyCompPaidToNonRes").value), 1, 1) & "</ITRForm:AnyCompPaidToNonRes>"
    'Else
        'writeXML "                       <ITRForm:AnyCompPaidToNonRes>N</ITRForm:AnyCompPaidToNonRes>"
    End If

    If Sheet4.Range("PL.AmtPaidToNonRes").value <> "" Then
        writeXML "                       <ITRForm:AmtPaidToNonRes>" & UCase(Sheet4.Range("PL.AmtPaidToNonRes").value) & "</ITRForm:AmtPaidToNonRes>"
    Else
        writeXML "                       <ITRForm:AmtPaidToNonRes>0</ITRForm:AmtPaidToNonRes>"
        'writeXML "                       <ITRForm:TotEmployeeComp>" & UCase(DefaultTotEmployeeComp_PL) & "</ITRForm:TotEmployeeComp>"
    End If

    writeXML "                  </ITRForm:EmployeeComp>"



    writeXML "                  <ITRForm:Insurances>"

    If Sheet4.Range("PL.MedInsur").value <> "" Then
        writeXML "                       <ITRForm:MedInsur>" & UCase(Sheet4.Range("PL.MedInsur").value) & "</ITRForm:MedInsur>"
    Else
        writeXML "                       <ITRForm:MedInsur>0</ITRForm:MedInsur>"
    End If

    If Sheet4.Range("PL.LifeInsur").value <> "" Then
        writeXML "                       <ITRForm:LifeInsur>" & UCase(Sheet4.Range("PL.LifeInsur").value) & "</ITRForm:LifeInsur>"
    Else
        writeXML "                       <ITRForm:LifeInsur>0</ITRForm:LifeInsur>"
    End If

    If Sheet4.Range("PL.KeyManInsur").value <> "" Then
        writeXML "                       <ITRForm:KeyManInsur>" & UCase(Sheet4.Range("PL.KeyManInsur").value) & "</ITRForm:KeyManInsur>"
    Else
        writeXML "                       <ITRForm:KeyManInsur>0</ITRForm:KeyManInsur>"
    End If

    If Sheet4.Range("PL.OthInsur").value <> "" Then
        writeXML "                       <ITRForm:OthInsur>" & UCase(Sheet4.Range("PL.OthInsur").value) & "</ITRForm:OthInsur>"
    Else
        writeXML "                       <ITRForm:OthInsur>0</ITRForm:OthInsur>"
    End If

    If Sheet4.Range("PL.TotInsurances").value <> "" Then
        writeXML "                       <ITRForm:TotInsurances>" & UCase(Sheet4.Range("PL.TotInsurances").value) & "</ITRForm:TotInsurances>"
    Else
        writeXML "                       <ITRForm:TotInsurances>0</ITRForm:TotInsurances>"
    End If

    writeXML "                  </ITRForm:Insurances>"

    If Sheet4.Range("PL.StaffWelfareExp").value <> "" Then
        writeXML "                   <ITRForm:StaffWelfareExp>" & UCase(Sheet4.Range("PL.StaffWelfareExp").value) & "</ITRForm:StaffWelfareExp>"
    Else
        writeXML "                   <ITRForm:StaffWelfareExp>0</ITRForm:StaffWelfareExp>"
    End If

    If Sheet4.Range("PL.Entertainment").value <> "" Then
        writeXML "                   <ITRForm:Entertainment>" & UCase(Sheet4.Range("PL.Entertainment").value) & "</ITRForm:Entertainment>"
    Else
        writeXML "                   <ITRForm:Entertainment>0</ITRForm:Entertainment>"
    End If

    If Sheet4.Range("PL.Hospitality").value <> "" Then
        writeXML "                   <ITRForm:Hospitality>" & UCase(Sheet4.Range("PL.Hospitality").value) & "</ITRForm:Hospitality>"
    Else
        writeXML "                   <ITRForm:Hospitality>0</ITRForm:Hospitality>"
    End If

    If Sheet4.Range("PL.Conference").value <> "" Then
        writeXML "                   <ITRForm:Conference>" & UCase(Sheet4.Range("PL.Conference").value) & "</ITRForm:Conference>"
    Else
        writeXML "                   <ITRForm:Conference>0</ITRForm:Conference>"
    End If

    If Sheet4.Range("PL.SalePromoExp").value <> "" Then
        writeXML "                   <ITRForm:SalePromoExp>" & UCase(Sheet4.Range("PL.SalePromoExp").value) & "</ITRForm:SalePromoExp>"
    Else
        writeXML "                   <ITRForm:SalePromoExp>0</ITRForm:SalePromoExp>"
    End If

    If Sheet4.Range("PL.Advertisement").value <> "" Then
        writeXML "                   <ITRForm:Advertisement>" & UCase(Sheet4.Range("PL.Advertisement").value) & "</ITRForm:Advertisement>"
    Else
        writeXML "                   <ITRForm:Advertisement>0</ITRForm:Advertisement>"
    End If



    writeXML "<ITRForm:CommissionExpdrDtls>"
    If Sheet4.Range("PLCE.NonResOtherCompany").value <> "" Then
        writeXML "<ITRForm:NonResOtherCompany>" & Sheet4.Range("PLCE.NonResOtherCompany").value & "</ITRForm:NonResOtherCompany>"
    Else
        writeXML "<ITRForm:NonResOtherCompany>0</ITRForm:NonResOtherCompany>"
    End If

    If Sheet4.Range("PLCE.Others").value <> "" Then
        writeXML "<ITRForm:Others>" & Sheet4.Range("PLCE.Others").value & "</ITRForm:Others>"
    Else
        writeXML "<ITRForm:Others>0</ITRForm:Others>"
    End If

    If Sheet4.Range("PL.CommissionExpdr").value <> "" Then
        writeXML "                   <ITRForm:Total>" & UCase(Sheet4.Range("PL.CommissionExpdr").value) & "</ITRForm:Total>"
    Else
        writeXML "                   <ITRForm:Total>0</ITRForm:Total>"
    End If
    writeXML "</ITRForm:CommissionExpdrDtls>  "

    writeXML "<ITRForm:RoyalityDtls>"


    If Sheet4.Range("PLRY.NonResOtherCompany").value <> "" Then
        writeXML "<ITRForm:NonResOtherCompany>" & Sheet4.Range("PLRY.NonResOtherCompany").value & "</ITRForm:NonResOtherCompany>"
    Else
        writeXML "<ITRForm:NonResOtherCompany>0</ITRForm:NonResOtherCompany>"
    End If

    If Sheet4.Range("PLRY.Others").value <> "" Then
        writeXML "<ITRForm:Others>" & Sheet4.Range("PLRY.Others").value & "</ITRForm:Others>"
    Else
        writeXML "<ITRForm:Others>0</ITRForm:Others>"
    End If

    If Sheet4.Range("PLRY.Total").value <> "" Then
        writeXML "                   <ITRForm:Total>" & UCase(Sheet4.Range("PLRY.Total").value) & "</ITRForm:Total>"
    Else
        writeXML "                   <ITRForm:Total>0</ITRForm:Total>"
    End If

    writeXML "</ITRForm:RoyalityDtls>"

    writeXML "<ITRForm:ProfessionalConstDtls>"
    If Sheet4.Range("PLPC.NonResOtherCompany").value <> "" Then
        writeXML "<ITRForm:NonResOtherCompany>" & Sheet4.Range("PLPC.NonResOtherCompany").value & "</ITRForm:NonResOtherCompany>"
    Else
        writeXML "<ITRForm:NonResOtherCompany>0</ITRForm:NonResOtherCompany>"
    End If

    If Sheet4.Range("PLPC.Others").value <> "" Then
        writeXML "<ITRForm:Others>" & Sheet4.Range("PLPC.Others").value & "</ITRForm:Others>"
    Else
        writeXML "<ITRForm:Others>0</ITRForm:Others>"
    End If

    If Sheet4.Range("PLPC.Total").value <> "" Then
        writeXML "                   <ITRForm:Total>" & UCase(Sheet4.Range("PLPC.Total").value) & "</ITRForm:Total>"
    Else
        writeXML "                   <ITRForm:Total>0</ITRForm:Total>"
    End If
    writeXML "</ITRForm:ProfessionalConstDtls>"


    If Sheet4.Range("PL.HotelBoardLodge").value <> "" Then
        writeXML "                   <ITRForm:HotelBoardLodge>" & UCase(Sheet4.Range("PL.HotelBoardLodge").value) & "</ITRForm:HotelBoardLodge>"
    Else
        writeXML "                   <ITRForm:HotelBoardLodge>0</ITRForm:HotelBoardLodge>"
    End If

    If Sheet4.Range("PL.TravelExp").value <> "" Then
        writeXML "                   <ITRForm:TravelExp>" & UCase(Sheet4.Range("PL.TravelExp").value) & "</ITRForm:TravelExp>"
    Else
        writeXML "                   <ITRForm:TravelExp>0</ITRForm:TravelExp>"
    End If

    If Sheet4.Range("PL.ForeignTravelExp").value <> "" Then
        writeXML "                   <ITRForm:ForeignTravelExp>" & UCase(Sheet4.Range("PL.ForeignTravelExp").value) & "</ITRForm:ForeignTravelExp>"
    Else
        writeXML "                   <ITRForm:ForeignTravelExp>0</ITRForm:ForeignTravelExp>"
    End If

    If Sheet4.Range("PL.ConveyanceExp").value <> "" Then
        writeXML "                   <ITRForm:ConveyanceExp>" & UCase(Sheet4.Range("PL.ConveyanceExp").value) & "</ITRForm:ConveyanceExp>"
    Else
        writeXML "                   <ITRForm:ConveyanceExp>0</ITRForm:ConveyanceExp>"
    End If

    If Sheet4.Range("PL.TelephoneExp").value <> "" Then
        writeXML "                   <ITRForm:TelephoneExp>" & UCase(Sheet4.Range("PL.TelephoneExp").value) & "</ITRForm:TelephoneExp>"
    Else
        writeXML "                   <ITRForm:TelephoneExp>0</ITRForm:TelephoneExp>"
    End If

    If Sheet4.Range("PL.GuestHouseExp").value <> "" Then
        writeXML "                   <ITRForm:GuestHouseExp>" & UCase(Sheet4.Range("PL.GuestHouseExp").value) & "</ITRForm:GuestHouseExp>"
    Else
        writeXML "                   <ITRForm:GuestHouseExp>0</ITRForm:GuestHouseExp>"
    End If

    If Sheet4.Range("PL.ClubExp").value <> "" Then
        writeXML "                   <ITRForm:ClubExp>" & UCase(Sheet4.Range("PL.ClubExp").value) & "</ITRForm:ClubExp>"
    Else
        writeXML "                   <ITRForm:ClubExp>0</ITRForm:ClubExp>"
    End If

    If Sheet4.Range("PL.FestivalCelebExp").value <> "" Then
        writeXML "                   <ITRForm:FestivalCelebExp>" & UCase(Sheet4.Range("PL.FestivalCelebExp").value) & "</ITRForm:FestivalCelebExp>"
    Else
        writeXML "                   <ITRForm:FestivalCelebExp>0</ITRForm:FestivalCelebExp>"
    End If

    If Sheet4.Range("PL.Scholarship").value <> "" Then
        writeXML "                   <ITRForm:Scholarship>" & UCase(Sheet4.Range("PL.Scholarship").value) & "</ITRForm:Scholarship>"
    Else
        writeXML "                   <ITRForm:Scholarship>0</ITRForm:Scholarship>"
    End If

    If Sheet4.Range("PL.Gift").value <> "" Then
        writeXML "                   <ITRForm:Gift>" & UCase(Sheet4.Range("PL.Gift").value) & "</ITRForm:Gift>"
    Else
        writeXML "                   <ITRForm:Gift>0</ITRForm:Gift>"
    End If

    If Sheet4.Range("PL.Donation").value <> "" Then
        writeXML "                   <ITRForm:Donation>" & UCase(Sheet4.Range("PL.Donation").value) & "</ITRForm:Donation>"
    Else
        writeXML "                   <ITRForm:Donation>0</ITRForm:Donation>"
    End If


    writeXML "                  <ITRForm:RatesTaxesPays>"
    writeXML "                      <ITRForm:ExciseCustomsVAT>"

    If Sheet4.Range("PLRateEx.UnionExciseDuty").value <> "" Then
        writeXML "                           <ITRForm:UnionExciseDuty>" & UCase(Sheet4.Range("PLRateEx.UnionExciseDuty").value) & "</ITRForm:UnionExciseDuty>"
    Else
        writeXML "                           <ITRForm:UnionExciseDuty>0</ITRForm:UnionExciseDuty>"
    End If

    If Sheet4.Range("PLRateEx.ServiceTax").value <> "" Then
        writeXML "                           <ITRForm:ServiceTax>" & UCase(Sheet4.Range("PLRateEx.ServiceTax").value) & "</ITRForm:ServiceTax>"
    Else
        writeXML "                           <ITRForm:ServiceTax>0</ITRForm:ServiceTax>"
    End If

    If Sheet4.Range("PLRateEx.VATorSaleTax").value <> "" Then
        writeXML "                           <ITRForm:VATorSaleTax>" & UCase(Sheet4.Range("PLRateEx.VATorSaleTax").value) & "</ITRForm:VATorSaleTax>"
    Else
        writeXML "                           <ITRForm:VATorSaleTax>0</ITRForm:VATorSaleTax>"
    End If

        If Sheet4.Range("PLRateEx.Cess").value <> "" Then
        writeXML "                           <ITRForm:Cess>" & UCase(Sheet4.Range("PLRateEx.Cess").value) & "</ITRForm:Cess>"
    Else
        writeXML "                           <ITRForm:Cess>0</ITRForm:Cess>"
    End If


    If Sheet4.Range("PLRateEx.CGST").value <> "" Then
        writeXML "                   <ITRForm:CentralGoodServiceTax>" & UCase(Sheet4.Range("PLRateEx.CGST").value) & "</ITRForm:CentralGoodServiceTax>"
    Else
        writeXML "                   <ITRForm:CentralGoodServiceTax>0</ITRForm:CentralGoodServiceTax>"
    End If

    If Sheet4.Range("PLRateEx.SGST").value <> "" Then
        writeXML "                   <ITRForm:StateGoodServiceTax>" & UCase(Sheet4.Range("PLRateEx.SGST").value) & "</ITRForm:StateGoodServiceTax>"
    Else
        writeXML "                   <ITRForm:StateGoodServiceTax>0</ITRForm:StateGoodServiceTax>"
    End If

    If Sheet4.Range("PLRateEx.IGST").value <> "" Then
        writeXML "                   <ITRForm:IntegratedGoodServiceTax>" & UCase(Sheet4.Range("PLRateEx.IGST").value) & "</ITRForm:IntegratedGoodServiceTax>"
    Else
        writeXML "                   <ITRForm:IntegratedGoodServiceTax>0</ITRForm:IntegratedGoodServiceTax>"
    End If

    If Sheet4.Range("PLRateEx.UTGST").value <> "" Then
        writeXML "                   <ITRForm:UnionTerrGoodServiceTax>" & UCase(Sheet4.Range("PLRateEx.UTGST").value) & "</ITRForm:UnionTerrGoodServiceTax>"
    Else
        writeXML "                   <ITRForm:UnionTerrGoodServiceTax>0</ITRForm:UnionTerrGoodServiceTax>"
    End If


    If Sheet4.Range("PLRateEx.OthDutyTaxCess").value <> "" Then
        writeXML "                           <ITRForm:OthDutyTaxCess>" & UCase(Sheet4.Range("PLRateEx.OthDutyTaxCess").value) & "</ITRForm:OthDutyTaxCess>"
    Else
        writeXML "                           <ITRForm:OthDutyTaxCess>0</ITRForm:OthDutyTaxCess>"
    End If

    If Sheet4.Range("PLRateEx.TotExciseCustomsVAT").value <> "" Then
        writeXML "                           <ITRForm:TotExciseCustomsVAT>" & UCase(Sheet4.Range("PLRateEx.TotExciseCustomsVAT").value) & "</ITRForm:TotExciseCustomsVAT>"
    Else
        writeXML "                           <ITRForm:TotExciseCustomsVAT>0</ITRForm:TotExciseCustomsVAT>"
    End If
    writeXML "                      </ITRForm:ExciseCustomsVAT>"
    writeXML "                  </ITRForm:RatesTaxesPays>"

    If Sheet4.Range("PL.AuditFee").value <> "" Then
        writeXML "                   <ITRForm:AuditFee>" & UCase(Sheet4.Range("PL.AuditFee").value) & "</ITRForm:AuditFee>"
    Else
        writeXML "                   <ITRForm:AuditFee>0</ITRForm:AuditFee>"
    End If

If (Not IsEmpty(Nature3)) Then
        For i = 1 To UBound(Nature3)
            If amount3(i) <> "" Then
                If amount3(i) > 0 Then
                    writeXML "<ITRForm:OtherExpensesDtls>"
                    If Nature3(i) <> "" Then
                        writeXML "<ITRForm:ExpenseNature>" & Nature3(i) & "</ITRForm:ExpenseNature>"
                    Else
                        writeXML "<ITRForm:ExpenseNature><ITRForm:ExpenseNature>"
                    End If
                    writeXML "<ITRForm:Amount>" & amount3(i) & "</ITRForm:Amount>"
                    writeXML "</ITRForm:OtherExpensesDtls>"
                 End If
             End If
         Next
End If


    If Sheet4.Range("PL.OtherExpenses").value <> "" Then
        writeXML "                   <ITRForm:OtherExpenses>" & UCase(Sheet4.Range("PL.OtherExpenses").value) & "</ITRForm:OtherExpenses>"
    Else
        writeXML "                   <ITRForm:OtherExpenses>0</ITRForm:OtherExpenses>"
    End If

    writeXML "<ITRForm:BadDebtDtls>"

    If (Not IsEmpty(Pan39_PL)) Then
        For i = 1 To UBound(Pan39_PL)
            If Amt39_PL(i) <> "" Then
                If Amt39_PL(i) > 0 Then
                    writeXML "<ITRForm:BadDebtAmtDtls>"
                    If Pan39_PL(i) <> "" Then
                        writeXML "<ITRForm:PAN>" & Pan39_PL(i) & "</ITRForm:PAN>"
                    Else
                        writeXML "<ITRForm:PAN><ITRForm:PAN>"
                    End If
                    If Aadhar39_PL(i) <> "" Then
                        writeXML "<ITRForm:Aadhaar>" & Aadhar39_PL(i) & "</ITRForm:Aadhaar>"
                    End If
                    writeXML "<ITRForm:Amount>" & Amt39_PL(i) & "</ITRForm:Amount>"
                    writeXML "</ITRForm:BadDebtAmtDtls>"
                End If
            End If
        Next
    End If


'    If Sheet4.Range("PL.OtherExpenses").value <> "" Then
       writeXML "                   <ITRForm:BadDebtAmtDtlsTotal>" & UCase(Sheet4.Range("PLBaddebts").value) & "</ITRForm:BadDebtAmtDtlsTotal>"
'    Else
'        writeXML "                   <ITRForm:OthersWherePANNotAvlble>0</ITRForm:OthersWherePANNotAvlble>"
'    End If

        If Not IsEmpty(Pl_OtherPanName) Then

              For i = 1 To UBound(Pl_OtherPanName)
               writeXML "<ITRForm:OthersPANNotAvlblDtl>"
                 If Pl_OtherPanName(i) <> "" Then
                  writeXML "<ITRForm:Name>" & (Pl_OtherPanName(i)) & "</ITRForm:Name>"
                 End If


                 If PL_Address_Flat(i) <> "" Then
                 writeXML "<ITRForm:FlatDoorBlockNumber>" & (PL_Address_Flat(i)) & "</ITRForm:FlatDoorBlockNumber>"
                 End If

                    If PL_Address_Village(i) <> "" Then
                 writeXML "<ITRForm:PremisesBuildingName>" & (PL_Address_Village(i)) & "</ITRForm:PremisesBuildingName>"
                 End If

                    If PL_Address_Road(i) <> "" Then
                 writeXML "<ITRForm:RoadStreetPostOffice>" & (PL_Address_Road(i)) & "</ITRForm:RoadStreetPostOffice>"
                 End If

                    If PL_Address_Area(i) <> "" Then
                 writeXML "<ITRForm:AreaLocality>" & (PL_Address_Area(i)) & "</ITRForm:AreaLocality>"
                 End If

                    If PL_Address_City(i) <> "" Then
                 writeXML "<ITRForm:TownCityDistrict>" & (PL_Address_City(i)) & "</ITRForm:TownCityDistrict>"
                 End If

                    If Not isdropdownblank(PL_Address_State(i)) Then
                 writeXML "<ITRForm:StateCode>" & UCase(Mid(Trim(PL_Address_State(i)), 1, InStr(1, Trim(PL_Address_State(i)), "-") - 1)) & "</ITRForm:StateCode>"
                 End If

                    If Not isdropdownblank(PL_Address_Country(i)) Then
                 writeXML "<ITRForm:CountryCode>" & UCase(Mid(Trim(PL_Address_Country(i)), 1, InStr(1, Trim(PL_Address_Country(i)), "-") - 1)) & "</ITRForm:CountryCode>"
                 End If

                    If PL_Address_Pin(i) <> "" Then
                 writeXML "<ITRForm:PinCode>" & (PL_Address_Pin(i)) & "</ITRForm:PinCode>"
                 End If

                    If PL_Address_Zip(i) <> "" Then
                 writeXML "<ITRForm:ZipCode>" & (PL_Address_Zip(i)) & "</ITRForm:ZipCode>"
                 End If

                 If Pl_OtherPanAmount(i) <> "" Then
                 writeXML "<ITRForm:Amount>" & (Pl_OtherPanAmount(i)) & "</ITRForm:Amount>"
                 End If

              writeXML "</ITRForm:OthersPANNotAvlblDtl>"
            Next
        End If

'    If Sheet4.Range("PL.OthersWherePANNotAvlble").value <> "" Then
        writeXML "<ITRForm:OthersPANNotAvlblDtlTotal>" & Sheet4.Range("PLOthers").value & "</ITRForm:OthersPANNotAvlblDtlTotal>"
'    Else
'        writeXML "<ITRForm:OthersWherePANNotAvlble>0</ITRForm:OthersWherePANNotAvlble>"
'    End If

    If Sheet4.Range("PL.OthersAmtLt1Lakh").value <> "" Then
        writeXML "<ITRForm:OthersAmtLt1Lakh>" & Sheet4.Range("PL.OthersAmtLt1Lakh").value & "</ITRForm:OthersAmtLt1Lakh>"
    Else
        writeXML "<ITRForm:OthersAmtLt1Lakh>0</ITRForm:OthersAmtLt1Lakh>"
    End If

    If Sheet4.Range("PL.BadDebt").value <> "" Then
        writeXML "                   <ITRForm:BadDebt>" & UCase(Sheet4.Range("PL.BadDebt").value) & "</ITRForm:BadDebt>"
    Else
        writeXML "                   <ITRForm:BadDebt>0</ITRForm:BadDebt>"
    End If

    writeXML "</ITRForm:BadDebtDtls>"

    If Sheet4.Range("PL.ProvForBadDoubtDebt").value <> "" Then
        writeXML "                   <ITRForm:ProvForBadDoubtDebt>" & UCase(Sheet4.Range("PL.ProvForBadDoubtDebt").value) & "</ITRForm:ProvForBadDoubtDebt>"
    Else
        writeXML "                   <ITRForm:ProvForBadDoubtDebt>0</ITRForm:ProvForBadDoubtDebt>"
    End If

    If Sheet4.Range("PL.OthProvisionsExpdr").value <> "" Then
     writeXML "                   <ITRForm:OthProvisionsExpdr>" & UCase(Sheet4.Range("PL.OthProvisionsExpdr").value) & "</ITRForm:OthProvisionsExpdr>"
    Else
        writeXML "                   <ITRForm:OthProvisionsExpdr>0</ITRForm:OthProvisionsExpdr>"
    End If

    If Sheet4.Range("PL.PBIDTA").value <> "" Then
        writeXML "                   <ITRForm:PBIDTA>" & UCase(Sheet4.Range("PL.PBIDTA").value) & "</ITRForm:PBIDTA>"
    Else
        writeXML "                   <ITRForm:PBIDTA>0</ITRForm:PBIDTA>"
    End If

    writeXML "<ITRForm:InterestExpdrtDtls>"

    If Sheet4.Range("PLI.NonResOtherCompany").value <> "" Then
        writeXML "<ITRForm:NonResOtherCompany>" & Sheet4.Range("PLI.NonResOtherCompany").value & "</ITRForm:NonResOtherCompany>"
    Else
        writeXML "<ITRForm:NonResOtherCompany>0</ITRForm:NonResOtherCompany>"
    End If

    If Sheet4.Range("PLI.Others").value <> "" Then
        writeXML "<ITRForm:Others>" & Sheet4.Range("PLI.Others").value & "</ITRForm:Others>"
    Else
        writeXML "<ITRForm:Others>0</ITRForm:Others>"
    End If

    If Sheet4.Range("PL.InterestExpdr").value <> "" Then
        writeXML "                   <ITRForm:InterestExpdr>" & UCase(Sheet4.Range("PL.InterestExpdr").value) & "</ITRForm:InterestExpdr>"
    Else
        writeXML "                   <ITRForm:InterestExpdr>0</ITRForm:InterestExpdr>"
    End If

    writeXML "</ITRForm:InterestExpdrtDtls> "

    If Sheet4.Range("PL.DepreciationAmort").value <> "" Then
        writeXML "                   <ITRForm:DepreciationAmort>" & UCase(Sheet4.Range("PL.DepreciationAmort").value) & "</ITRForm:DepreciationAmort>"
    Else
        writeXML "                   <ITRForm:DepreciationAmort>0</ITRForm:DepreciationAmort>"
    End If

    If Sheet4.Range("PL.PBT").value <> "" Then
        writeXML "                   <ITRForm:PBT>" & UCase(Sheet4.Range("PL.PBT").value) & "</ITRForm:PBT>"
    Else
        writeXML "                   <ITRForm:PBT>0</ITRForm:PBT>"
    End If
    writeXML "              </ITRForm:DebitsToPL>"

    writeXML "              <ITRForm:TaxProvAppr>"

    If Sheet4.Range("PL.ProvForCurrTax").value <> "" Then
        writeXML "                   <ITRForm:ProvForCurrTax>" & UCase(Sheet4.Range("PL.ProvForCurrTax").value) & "</ITRForm:ProvForCurrTax>"
    Else
        writeXML "                   <ITRForm:ProvForCurrTax>0</ITRForm:ProvForCurrTax>"
    End If

    If Sheet4.Range("PL.ProvDefTax").value <> "" Then
        writeXML "                   <ITRForm:ProvDefTax>" & UCase(Sheet4.Range("PL.ProvDefTax").value) & "</ITRForm:ProvDefTax>"
    Else
        writeXML "                   <ITRForm:ProvDefTax>0</ITRForm:ProvDefTax>"
    End If

    If Sheet4.Range("PL.ProfitAfterTax").value <> "" Then
        writeXML "                   <ITRForm:ProfitAfterTax>" & UCase(Sheet4.Range("PL.ProfitAfterTax").value) & "</ITRForm:ProfitAfterTax>"
    Else
        writeXML "                   <ITRForm:ProfitAfterTax>0</ITRForm:ProfitAfterTax>"
    End If

    If Sheet4.Range("PL.BalBFPrevYr").value <> "" Then
        writeXML "                   <ITRForm:BalBFPrevYr>" & UCase(Sheet4.Range("PL.BalBFPrevYr").value) & "</ITRForm:BalBFPrevYr>"
    Else
        writeXML "                   <ITRForm:BalBFPrevYr>0</ITRForm:BalBFPrevYr>"
    End If

    If Sheet4.Range("PL.AmtAvlAppr").value <> "" Then
        writeXML "                   <ITRForm:AmtAvlAppr>" & UCase(Sheet4.Range("PL.AmtAvlAppr").value) & "</ITRForm:AmtAvlAppr>"
    Else
        writeXML "                   <ITRForm:AmtAvlAppr>0</ITRForm:AmtAvlAppr>"
    End If

   If Sheet4.Range("PL.TrfToReserves").value <> "" Then
        writeXML "                   <ITRForm:TrfToReserves>" & UCase(Sheet4.Range("PL.TrfToReserves").value) & "</ITRForm:TrfToReserves>"
    Else
        writeXML "                   <ITRForm:TrfToReserves>0</ITRForm:TrfToReserves>"
    End If

     If Sheet4.Range("PL.PartnerAccBalTrf").value <> "" Then
        writeXML " <ITRForm:ProprietorAccBalTrf>" & UCase(Sheet4.Range("PL.PartnerAccBalTrf").value) & "</ITRForm:ProprietorAccBalTrf>"
    Else
        writeXML " <ITRForm:ProprietorAccBalTrf>0</ITRForm:ProprietorAccBalTrf>"
    End If

    writeXML "              </ITRForm:TaxProvAppr>"



If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
    If Not IsEmpty(NOBCode) Then

              For i = 1 To UBound(NOBCode)
               writeXML "<ITRForm:NatOfBus44AD>"
                 If NOBName(i) <> "" Then
                  writeXML "<ITRForm:NameOfBusiness>" & (NOBName(i)) & "</ITRForm:NameOfBusiness>"
                 End If

                 If NOBCode(i) <> "" Then
                 writeXML "<ITRForm:CodeAD>" & (NOBCode(i)) & "</ITRForm:CodeAD>"
                 End If

                 If NOBDescription(i) <> "" Then
                 writeXML "<ITRForm:Description>" & (NOBDescription(i)) & "</ITRForm:Description>"
                 End If

              writeXML "</ITRForm:NatOfBus44AD>"
            Next
  End If

  writeXML "              <ITRForm:PersumptiveInc44AD>"

  If Sheet4.Range("PL_61i").value <> "" Then
        writeXML " <ITRForm:GrsTrnOverOrReceipt>" & UCase(Sheet4.Range("PL_61i").value) & "</ITRForm:GrsTrnOverOrReceipt>"
    Else
        writeXML " <ITRForm:GrsTrnOverOrReceipt>0</ITRForm:GrsTrnOverOrReceipt>"
  End If

  If Sheet4.Range("PL_61a").value <> "" Then
        writeXML " <ITRForm:GrsTrnOverBank>" & UCase(Sheet4.Range("PL_61a").value) & "</ITRForm:GrsTrnOverBank>"
  End If

  If Sheet4.Range("PL_61b").value <> "" Then
        writeXML " <ITRForm:GrsTrnOverAnyOthMode>" & UCase(Sheet4.Range("PL_61b").value) & "</ITRForm:GrsTrnOverAnyOthMode>"
  End If

    If Sheet4.Range("PL_61ii").value <> "" Then
        writeXML " <ITRForm:TotPersumptiveInc44AD>" & UCase(Sheet4.Range("PL_61ii").value) & "</ITRForm:TotPersumptiveInc44AD>"
  End If

    If Sheet4.Range("PL_61iia").value <> "" Then
        writeXML " <ITRForm:PersumptiveInc44AD6Per>" & UCase(Sheet4.Range("PL_61iia").value) & "</ITRForm:PersumptiveInc44AD6Per>"
  End If

    If Sheet4.Range("PL_61iib").value <> "" Then
        writeXML " <ITRForm:PersumptiveInc44AD8Per>" & UCase(Sheet4.Range("PL_61iib").value) & "</ITRForm:PersumptiveInc44AD8Per>"
  End If

  writeXML "              </ITRForm:PersumptiveInc44AD>"



   If (Not IsEmpty(NOBCode1)) And UBound(NOBCode1) > 0 Then



                For i = 1 To UBound(NOBName1)
                   writeXML "<ITRForm:NatOfBus44ADA>"
                        If NOBName1(i) <> "" Then
                            writeXML "<ITRForm:NameOfBusiness>" & (NOBName1(i)) & "</ITRForm:NameOfBusiness>"
                        Else
                            writeXML "<ITRForm:NameOfBusiness/>"
                        End If

                        If NOBCode1(i) <> "" Then
                            writeXML "<ITRForm:CodeADA>" & (NOBCode1(i)) & "</ITRForm:CodeADA>"
                            Else
                            writeXML "<ITRForm:CodeADA/>"
                        End If

                        If NOBDescription1(i) <> "" Then
                            writeXML "<ITRForm:Description>" & NOBDescription1(i) & "</ITRForm:Description>"
                        Else
                            writeXML "<ITRForm:Description/>"
                        End If
                    writeXML "</ITRForm:NatOfBus44ADA>"
                Next

        End If

    writeXML "              <ITRForm:PersumptiveInc44ADA>"
      If Sheet4.Range("PL_62i").value <> "" Then
        writeXML " <ITRForm:GrsReceipt>" & UCase(Sheet4.Range("PL_62i").value) & "</ITRForm:GrsReceipt>"
       Else
        writeXML " <ITRForm:GrsReceipt>0</ITRForm:GrsReceipt>"
     End If

    If Sheet4.Range("PL_62ii").value <> "" Then
        writeXML " <ITRForm:TotPersumptiveInc44ADA>" & UCase(Sheet4.Range("PL_62ii").value) & "</ITRForm:TotPersumptiveInc44ADA>"
    End If
   writeXML "              </ITRForm:PersumptiveInc44ADA>"

 End If
    If (Not IsEmpty(NOBCode2)) And UBound(NOBCode2) > 0 Then



                For i = 1 To UBound(NOBName2)
                  writeXML "<ITRForm:NatOfBus44AE>"
                        If NOBName2(i) <> "" Then
                            writeXML "<ITRForm:NameOfBusiness>" & (NOBName2(i)) & "</ITRForm:NameOfBusiness>"
                        Else
                            writeXML "<ITRForm:NameOfBusiness/>"
                        End If

                        If NOBCode2(i) <> "" Then
                            writeXML "<ITRForm:CodeAE>" & (NOBCode2(i)) & "</ITRForm:CodeAE>"
                            Else
                            writeXML "<ITRForm:CodeAE/>"
                        End If

                        If NOBDescription2(i) <> "" Then
                            writeXML "<ITRForm:Description>" & NOBDescription2(i) & "</ITRForm:Description>"
                        Else
                            writeXML "<ITRForm:Description/>"
                        End If
                    writeXML "</ITRForm:NatOfBus44AE>"
                Next

        End If



    If UBound(SEC44AEMonth) > 0 Then
        For i = 1 To UBound(SEC44AEMonth)
            writeXML "<ITRForm:GoodsDtlsUs44AE>"
            writeXML "<ITRForm:RegNumberGoodsCarriage>" & SEC44AERegNo(i) & "</ITRForm:RegNumberGoodsCarriage> "
            writeXML "<ITRForm:OwnedLeasedHiredFlag>" & SEC44AEOLH(i) & "</ITRForm:OwnedLeasedHiredFlag> "
            writeXML "<ITRForm:TonnageCapacity>" & SEC44AETC(i) & "</ITRForm:TonnageCapacity>"
            writeXML "<ITRForm:HoldingPeriod>" & SEC44AEMonth(i) & "</ITRForm:HoldingPeriod> "
            writeXML "<ITRForm:PresumptiveIncome>" & SEC44AEPIncome(i) & "</ITRForm:PresumptiveIncome>"
            writeXML "</ITRForm:GoodsDtlsUs44AE>"
        Next
   End If

     If Sheet4.Range("PL_TotalMonths").value <> "" Then
        writeXML " <ITRForm:TotalNumOfMonths>" & UCase(Sheet4.Range("PL_TotalMonths").value) & "</ITRForm:TotalNumOfMonths>"
  End If

    If Sheet4.Range("PL_TotalPresumtiveIncome").value <> "" Then
        writeXML " <ITRForm:TotalPrsumptvIncUs44EGoods>" & UCase(Sheet4.Range("PL_TotalPresumtiveIncome").value) & "</ITRForm:TotalPrsumptvIncUs44EGoods>"
  End If

  If Sheet4.Range("PL_63ii").value <> "" Then
        writeXML " <ITRForm:TotalPrsumptvIncUs44E>" & UCase(Sheet4.Range("PL_63ii").value) & "</ITRForm:TotalPrsumptvIncUs44E>"
  End If

  writeXML "<ITRForm:NoBooksOfAccPL>"

    If Sheet4.Range("PL_64GrossTurnOver").value <> "" Then
        writeXML " <ITRForm:GrossReceipt>" & UCase(Sheet4.Range("PL_64GrossTurnOver").value) & "</ITRForm:GrossReceipt>"
  End If

    If Sheet4.Range("PL_64ia1").value <> "" Then
        writeXML " <ITRForm:GrsRcptAccPayeeOrBankMode>" & UCase(Sheet4.Range("PL_64ia1").value) & "</ITRForm:GrsRcptAccPayeeOrBankMode>"
        Else
      writeXML " <ITRForm:GrsRcptAccPayeeOrBankMode>" & 0 & "</ITRForm:GrsRcptAccPayeeOrBankMode>"
  End If

    If Sheet4.Range("PL_64ia2").value <> "" Then
        writeXML " <ITRForm:GrsRcptOtherMode>" & UCase(Sheet4.Range("PL_64ia2").value) & "</ITRForm:GrsRcptOtherMode>"
        Else
        writeXML " <ITRForm:GrsRcptOtherMode>" & 0 & "</ITRForm:GrsRcptOtherMode>"
  End If

    If Sheet4.Range("PL_64ib").value <> "" Then
        writeXML " <ITRForm:GrossProfit>" & UCase(Sheet4.Range("PL_64ib").value) & "</ITRForm:GrossProfit>"
        Else
       writeXML " <ITRForm:GrossProfit>" & 0 & "</ITRForm:GrossProfit>"
  End If

    If Sheet4.Range("PL_64ic").value <> "" Then
        writeXML " <ITRForm:Expenses>" & UCase(Sheet4.Range("PL_64ic").value) & "</ITRForm:Expenses>"
        Else
         writeXML " <ITRForm:Expenses>" & 0 & "</ITRForm:Expenses>"
  End If

    If Sheet4.Range("PL_64id").value <> "" Then
        writeXML " <ITRForm:NetProfit>" & UCase(Sheet4.Range("PL_64id").value) & "</ITRForm:NetProfit>"
        Else
        writeXML " <ITRForm:NetProfit>" & 0 & "</ITRForm:NetProfit>"
  End If

    If Sheet4.Range("PL_64iia").value <> "" Then
        writeXML " <ITRForm:GrossReceiptPrf>" & UCase(Sheet4.Range("PL_64iia").value) & "</ITRForm:GrossReceiptPrf>"
        Else
       writeXML " <ITRForm:GrossReceiptPrf>" & 0 & "</ITRForm:GrossReceiptPrf>"
  End If

    If Sheet4.Range("PL_64iia_i").value <> "" Then
        writeXML " <ITRForm:GrsRcptAccPayeeOrBankModePrf>" & UCase(Sheet4.Range("PL_64iia_i").value) & "</ITRForm:GrsRcptAccPayeeOrBankModePrf>"
        Else
       writeXML " <ITRForm:GrsRcptAccPayeeOrBankModePrf>" & 0 & "</ITRForm:GrsRcptAccPayeeOrBankModePrf>"
  End If

    If Sheet4.Range("PL_64iia_ii").value <> "" Then
        writeXML " <ITRForm:GrsRcptOtherModePrf>" & UCase(Sheet4.Range("PL_64iia_ii").value) & "</ITRForm:GrsRcptOtherModePrf>"
        Else
        writeXML " <ITRForm:GrsRcptOtherModePrf>" & 0 & "</ITRForm:GrsRcptOtherModePrf>"
  End If

    If Sheet4.Range("PL_64iib").value <> "" Then
        writeXML " <ITRForm:GrossProfitPrf>" & UCase(Sheet4.Range("PL_64iib").value) & "</ITRForm:GrossProfitPrf>"
        Else
        writeXML " <ITRForm:GrossProfitPrf>" & 0 & "</ITRForm:GrossProfitPrf>"
  End If

    If Sheet4.Range("PL_64iic").value <> "" Then
        writeXML " <ITRForm:ExpensesPrf>" & UCase(Sheet4.Range("PL_64iic").value) & "</ITRForm:ExpensesPrf>"
        Else
        writeXML " <ITRForm:ExpensesPrf>" & 0 & "</ITRForm:ExpensesPrf>"
  End If

    If Sheet4.Range("PL_64iid").value <> "" Then
        writeXML " <ITRForm:NetProfitPrf>" & UCase(Sheet4.Range("PL_64iid").value) & "</ITRForm:NetProfitPrf>"
        Else
        writeXML " <ITRForm:NetProfitPrf>" & 0 & "</ITRForm:NetProfitPrf>"

  End If

    If Sheet4.Range("PL_64iii").value <> "" Then
        writeXML " <ITRForm:TotBusinessProfession>" & UCase(Sheet4.Range("PL_64iii").value) & "</ITRForm:TotBusinessProfession>"
  End If

  writeXML "</ITRForm:NoBooksOfAccPL>"

  If Sheet4.Range("PL_65i").value <> "" Then
        writeXML " <ITRForm:TurnverFrmSpecActivity>" & UCase(Sheet4.Range("PL_65i").value) & "</ITRForm:TurnverFrmSpecActivity>"
  Else
    writeXML " <ITRForm:TurnverFrmSpecActivity>" & 0 & "</ITRForm:TurnverFrmSpecActivity>"
  End If

     If Sheet4.Range("PL_65ii").value <> "" Then
        writeXML " <ITRForm:GrossProfit>" & UCase(Sheet4.Range("PL_65ii").value) & "</ITRForm:GrossProfit>"
  End If

     If Sheet4.Range("PL_65iii").value <> "" Then
        writeXML " <ITRForm:Expenditure>" & UCase(Sheet4.Range("PL_65iii").value) & "</ITRForm:Expenditure>"
  End If

     If Sheet4.Range("PL_65iv").value <> "" Then
        writeXML " <ITRForm:NetIncomeFrmSpecActivity>" & UCase(Sheet4.Range("PL_65iv").value) & "</ITRForm:NetIncomeFrmSpecActivity>"
    Else
      writeXML " <ITRForm:NetIncomeFrmSpecActivity>" & 0 & "</ITRForm:NetIncomeFrmSpecActivity>"
  End If

    writeXML "      </ITRForm:PARTA_PL>"

End Function

Function PARTA_OI()
writeXML " <ITRForm:PARTA_OI>"

    If MethodOfAcct_5 = "" Then
        writeXML " <ITRForm:MethodOfAcct>" & DefaultMethodOfAcct_5 & "</ITRForm:MethodOfAcct>"
    Else
        writeXML " <ITRForm:MethodOfAcct>" & UCase(MethodOfAcct_5) & "</ITRForm:MethodOfAcct>"
    End If

    If ChangeInAcctMethFlg_5 = "" Then
        writeXML " <ITRForm:ChangeInAcctMethFlg>" & DefaultChangeInAcctMethFlg_5 & "</ITRForm:ChangeInAcctMethFlg>"
    Else
        writeXML " <ITRForm:ChangeInAcctMethFlg>" & UCase(ChangeInAcctMethFlg_5) & "</ITRForm:ChangeInAcctMethFlg>"
    End If

    If ProfDeviatDueAcctMeth_5 = "" Then
        writeXML " <ITRForm:ProfDeviatDueAcctMeth>0</ITRForm:ProfDeviatDueAcctMeth>"
    Else
        writeXML " <ITRForm:ProfDeviatDueAcctMeth>" & UCase(ProfDeviatDueAcctMeth_5) & "</ITRForm:ProfDeviatDueAcctMeth>"
    End If

    If ProfDeviatDueAcctMeth_6 = "" Then
        writeXML " <ITRForm:DecProOrIncLossUs145_2>0</ITRForm:DecProOrIncLossUs145_2>"
    Else
        writeXML " <ITRForm:DecProOrIncLossUs145_2>" & UCase(ProfDeviatDueAcctMeth_6) & "</ITRForm:DecProOrIncLossUs145_2>"
    End If

    If (ValRawMaterial_5 <> "" Or ValFinishedGoods_5 <> "" Or ChngStockValMetFlg_6 <> "" Or _
        EffectOnPL_6 <> "" Or EffectOnPL_6_4e <> "") Then

        writeXML " <ITRForm:MethodOfValClgStk>"
        If ValRawMaterial_5 = "" Or ValRawMaterial_5 = "(Select)" Then
            writeXML " <ITRForm:ValRawMaterial>" & DefaultValRawMaterial_5 & "</ITRForm:ValRawMaterial>"
        Else
            writeXML " <ITRForm:ValRawMaterial>" & Trim(Mid(ValRawMaterial_5, 1, InStr(1, (ValRawMaterial_5), "-") - 1)) & "</ITRForm:ValRawMaterial>"
        End If

        If ValFinishedGoods_5 = "" Or ValFinishedGoods_5 = "(Select)" Then
            writeXML " <ITRForm:ValFinishedGoods>" & DefaultValFinishedGoods_5 & "</ITRForm:ValFinishedGoods>"
        Else
            writeXML " <ITRForm:ValFinishedGoods>" & Trim(Mid(ValFinishedGoods_5, 1, InStr(1, (ValFinishedGoods_5), "-") - 1)) & "</ITRForm:ValFinishedGoods>"
        End If
    End If

    If (ValRawMaterial_5 <> "" Or ValFinishedGoods_5 <> "" Or ChngStockValMetFlg_6 <> "" Or _
        EffectOnPL_6 <> "" Or EffectOnPL_6_4e <> "") Then

        If ChngStockValMetFlg_6 = "" Or ChngStockValMetFlg_6 = "(" Then
            writeXML " <ITRForm:ChngStockValMetFlg>" & DefaultChngStockValMetFlg_6 & "</ITRForm:ChngStockValMetFlg>"
        Else
            writeXML " <ITRForm:ChngStockValMetFlg>" & UCase(ChngStockValMetFlg_6) & "</ITRForm:ChngStockValMetFlg>"
        End If

        If EffectOnPL_6 = "" Then
            writeXML " <ITRForm:EffectOnPL>0</ITRForm:EffectOnPL>"
        Else
            writeXML " <ITRForm:EffectOnPL>" & UCase(EffectOnPL_6) & "</ITRForm:EffectOnPL>"
        End If

        If EffectOnPL_6_4e = "" Then
            writeXML " <ITRForm:DecProOrIncLossUs145_A>0</ITRForm:DecProOrIncLossUs145_A>"
        Else
            writeXML " <ITRForm:DecProOrIncLossUs145_A>" & UCase(EffectOnPL_6_4e) & "</ITRForm:DecProOrIncLossUs145_A>"
        End If

        writeXML " </ITRForm:MethodOfValClgStk>"
    End If

    writeXML " <ITRForm:NoCredToPLAmt>"
    If Section28Items_6 = "" Then
        writeXML " <ITRForm:Section28Items>0</ITRForm:Section28Items>"
    Else
        writeXML " <ITRForm:Section28Items>" & UCase(Section28Items_6) & "</ITRForm:Section28Items>"
    End If

    If ProformaCreditsDue_6 = "" Then
        writeXML " <ITRForm:ProformaCreditsDue>0</ITRForm:ProformaCreditsDue>"
    Else
        writeXML " <ITRForm:ProformaCreditsDue>" & UCase(ProformaCreditsDue_6) & "</ITRForm:ProformaCreditsDue>"
    End If

    If PrevYrEscalClaim_6 = "" Then
        writeXML " <ITRForm:PrevYrEscalClaim>0</ITRForm:PrevYrEscalClaim>"
    Else
        writeXML " <ITRForm:PrevYrEscalClaim>" & UCase(PrevYrEscalClaim_6) & "</ITRForm:PrevYrEscalClaim>"
    End If

    If OthItemInc_6 = "" Then
        writeXML " <ITRForm:OthItemInc>0</ITRForm:OthItemInc>"
    Else
        writeXML " <ITRForm:OthItemInc>" & UCase(OthItemInc_6) & "</ITRForm:OthItemInc>"
    End If

    If CapReceipt_6 = "" Then
        writeXML " <ITRForm:CapReceipt>0</ITRForm:CapReceipt>"
    Else
        writeXML " <ITRForm:CapReceipt>" & UCase(CapReceipt_6) & "</ITRForm:CapReceipt>"
    End If

    If TotNoCredToPLAmt_6 = "" Then
        writeXML " <ITRForm:TotNoCredToPLAmt>0</ITRForm:TotNoCredToPLAmt>"
    Else
        writeXML " <ITRForm:TotNoCredToPLAmt>" & UCase(TotNoCredToPLAmt_6) & "</ITRForm:TotNoCredToPLAmt>"
    End If

    writeXML " </ITRForm:NoCredToPLAmt>"
    writeXML " <ITRForm:AmtDisallUs36>"

    If StkInsurPrem_6 = "" Then
        writeXML " <ITRForm:StkInsurPrem>0</ITRForm:StkInsurPrem>"
    Else
        writeXML " <ITRForm:StkInsurPrem>" & UCase(StkInsurPrem_6) & "</ITRForm:StkInsurPrem>"
    End If

    If EmpHealthInsurPrem_6 = "" Then
        writeXML " <ITRForm:EmpHealthInsurPrem>0</ITRForm:EmpHealthInsurPrem>"
    Else
        writeXML " <ITRForm:EmpHealthInsurPrem>" & UCase(EmpHealthInsurPrem_6) & "</ITRForm:EmpHealthInsurPrem>"
    End If

    If EmpBonusCommSum_6 = "" Then
        writeXML " <ITRForm:EmpBonusCommSum>0</ITRForm:EmpBonusCommSum>"
    Else
        writeXML " <ITRForm:EmpBonusCommSum>" & UCase(EmpBonusCommSum_6) & "</ITRForm:EmpBonusCommSum>"
    End If

    If IntOnBorrCap_6 = "" Then
        writeXML " <ITRForm:IntOnBorrCap>0</ITRForm:IntOnBorrCap>"
    Else
        writeXML " <ITRForm:IntOnBorrCap>" & UCase(IntOnBorrCap_6) & "</ITRForm:IntOnBorrCap>"
    End If

    If ZeroCoupBondDisc_6 = "" Then
        writeXML " <ITRForm:ZeroCoupBondDisc>0</ITRForm:ZeroCoupBondDisc>"
    Else
        writeXML " <ITRForm:ZeroCoupBondDisc>" & UCase(ZeroCoupBondDisc_6) & "</ITRForm:ZeroCoupBondDisc>"
    End If

    If RecogPFContribAmt_6 = "" Then
        writeXML " <ITRForm:RecogPFContribAmt>0</ITRForm:RecogPFContribAmt>"
    Else
        writeXML " <ITRForm:RecogPFContribAmt>" & UCase(RecogPFContribAmt_6) & "</ITRForm:RecogPFContribAmt>"
    End If

    If AppSuperAnnFundAmt_6 = "" Then
        writeXML " <ITRForm:AppSuperAnnFundAmt>0</ITRForm:AppSuperAnnFundAmt>"
    Else
        writeXML " <ITRForm:AppSuperAnnFundAmt>" & UCase(AppSuperAnnFundAmt_6) & "</ITRForm:AppSuperAnnFundAmt>"
    End If
    writeXML " <ITRForm:PensionSchemeSec80CCD>" & UVCase(PensionSchemeSec80CCD_6) & "</ITRForm:PensionSchemeSec80CCD> "

    If AppGratFundAmt_6 = "" Then
        writeXML " <ITRForm:AppGratFundAmt>0</ITRForm:AppGratFundAmt>"
    Else
        writeXML " <ITRForm:AppGratFundAmt>" & UCase(AppGratFundAmt_6) & "</ITRForm:AppGratFundAmt>"
    End If

    If OthFundAmt_6 = "" Then
        writeXML " <ITRForm:OthFundAmt>0</ITRForm:OthFundAmt>"
    Else
        writeXML " <ITRForm:OthFundAmt>" & UCase(OthFundAmt_6) & "</ITRForm:OthFundAmt>"
    End If
    
     If Sheet5.Range("sheet6.EmpContributionCredits").value = "" Then
        writeXML " <ITRForm:EmpContributionCredits>0</ITRForm:EmpContributionCredits>"
    Else
        writeXML " <ITRForm:EmpContributionCredits>" & UCase(Sheet5.Range("sheet6.EmpContributionCredits").value) & "</ITRForm:EmpContributionCredits>"
    End If


    If BadDebtDoubtAmt_6 = "" Then
        writeXML " <ITRForm:BadDebtDoubtAmt>0</ITRForm:BadDebtDoubtAmt>"
    Else
        writeXML " <ITRForm:BadDebtDoubtAmt>" & UCase(BadDebtDoubtAmt_6) & "</ITRForm:BadDebtDoubtAmt>"
    End If

    If BadDebtDoubtProvn_6 = "" Then
        writeXML " <ITRForm:BadDebtDoubtProvn>0</ITRForm:BadDebtDoubtProvn>"
    Else
        writeXML " <ITRForm:BadDebtDoubtProvn>" & UCase(BadDebtDoubtProvn_6) & "</ITRForm:BadDebtDoubtProvn>"
    End If

    If SpecResrvTranfr_6 = "" Then
        writeXML " <ITRForm:SpecResrvTranfr>0</ITRForm:SpecResrvTranfr>"
    Else
        writeXML " <ITRForm:SpecResrvTranfr>" & UCase(SpecResrvTranfr_6) & "</ITRForm:SpecResrvTranfr>"
    End If

    If FamPlanPromoExp_6 = "" Then
        writeXML " <ITRForm:FamPlanPromoExp>0</ITRForm:FamPlanPromoExp>"
    Else
        writeXML " <ITRForm:FamPlanPromoExp>" & UCase(FamPlanPromoExp_6) & "</ITRForm:FamPlanPromoExp>"
    End If

'    If EmpContributionCredits_6 = "" Then
'        writeXML " <ITRForm:EmpContributionCredits>0</ITRForm:EmpContributionCredits>"
'    Else
'        writeXML " <ITRForm:EmpContributionCredits>" & UCase(EmpContributionCredits_6) & "</ITRForm:EmpContributionCredits>"
'    End If

    writeXML " <ITRForm:SecuritiesPaidAmt>" & UVCase(SecuritiesPaidAmt_6) & "</ITRForm:SecuritiesPaidAmt>"

       If Sheet5.Range("sheet6.MarketLoss").value = "" Then
        writeXML " <ITRForm:MrktLossOthExpLossICDS>0</ITRForm:MrktLossOthExpLossICDS>"
    Else
        writeXML " <ITRForm:MrktLossOthExpLossICDS>" & UVCase(Sheet5.Range("sheet6.MarketLoss").value) & "</ITRForm:MrktLossOthExpLossICDS>"
    End If




    If OthDisallowances_6 = "" Then
        writeXML " <ITRForm:OthDisallowances>0</ITRForm:OthDisallowances>"
    Else
        writeXML " <ITRForm:OthDisallowances>" & UCase(OthDisallowances_6) & "</ITRForm:OthDisallowances>"
    End If

    If TotAmtDisallUs36_6 = "" Then
        writeXML " <ITRForm:TotAmtDisallUs36>0</ITRForm:TotAmtDisallUs36>"
    Else
        writeXML " <ITRForm:TotAmtDisallUs36>" & UCase(TotAmtDisallUs36_6) & "</ITRForm:TotAmtDisallUs36>"
    End If

    writeXML " </ITRForm:AmtDisallUs36>"


    writeXML " <ITRForm:AmtDisallUs37>"
    writeXML " <ITRForm:CapitalNatureExp>" & UVCase(CapitalNatureExp_6) & "</ITRForm:CapitalNatureExp>"
    If PersonalExp_6 = "" Then
        writeXML " <ITRForm:PersonalExp>0</ITRForm:PersonalExp>"
    Else
        writeXML " <ITRForm:PersonalExp>" & UCase(PersonalExp_6) & "</ITRForm:PersonalExp>"
    End If

    writeXML " <ITRForm:BusOrProfessnExp>" & UVCase(BusOrProfessnExp) & "</ITRForm:BusOrProfessnExp>"
    If PoliticPartyExp_6 = "" Then
        writeXML " <ITRForm:PoliticPartyExp>0</ITRForm:PoliticPartyExp>"
    Else
        writeXML " <ITRForm:PoliticPartyExp>" & UCase(PoliticPartyExp_6) & "</ITRForm:PoliticPartyExp>"
    End If

    If LawVoilatPenalExp_6 = "" Then
        writeXML " <ITRForm:LawVoilatPenalExp>0</ITRForm:LawVoilatPenalExp>"
    Else
        writeXML " <ITRForm:LawVoilatPenalExp>" & UCase(LawVoilatPenalExp_6) & "</ITRForm:LawVoilatPenalExp>"
    End If

    If OthPenalFineExp_6 = "" Then
        writeXML " <ITRForm:OthPenalFineExp>0</ITRForm:OthPenalFineExp>"
    Else
        writeXML " <ITRForm:OthPenalFineExp>" & UCase(OthPenalFineExp_6) & "</ITRForm:OthPenalFineExp>"
    End If

    If OffenceExp_6 = "" Then
        writeXML " <ITRForm:OffenceExp>0</ITRForm:OffenceExp>"
    Else
        writeXML " <ITRForm:OffenceExp>" & UCase(OffenceExp_6) & "</ITRForm:OffenceExp>"
    End If

    If ContigentLiability_6 = "" Then
        writeXML " <ITRForm:ContigentLiability>0</ITRForm:ContigentLiability>"
    Else
        writeXML " <ITRForm:ContigentLiability>" & UCase(ContigentLiability_6) & "</ITRForm:ContigentLiability>"
    End If

    If OthAmtNotAllowUs37_6 = "" Then
        writeXML " <ITRForm:OthAmtNotAllowUs37>0</ITRForm:OthAmtNotAllowUs37>"
    Else
        writeXML " <ITRForm:OthAmtNotAllowUs37>" & UCase(OthAmtNotAllowUs37_6) & "</ITRForm:OthAmtNotAllowUs37>"
    End If

    If TotAmtDisallUs37_6 = "" Then
        writeXML " <ITRForm:TotAmtDisallUs37>0</ITRForm:TotAmtDisallUs37>"
    Else
        writeXML " <ITRForm:TotAmtDisallUs37>" & UCase(TotAmtDisallUs37_6) & "</ITRForm:TotAmtDisallUs37>"
    End If
    writeXML " </ITRForm:AmtDisallUs37>"

    writeXML " <ITRForm:AmtDisallUs40>"
    If NonCompChapXVIIBAmt_6 = "" Then
        writeXML " <ITRForm:NonCompChapXVIIBAmt>0</ITRForm:NonCompChapXVIIBAmt>"
    Else
        writeXML " <ITRForm:NonCompChapXVIIBAmt>" & UCase(NonCompChapXVIIBAmt_6) & "</ITRForm:NonCompChapXVIIBAmt>"
    End If
    writeXML " <ITRForm:NonComp40aiiChapXVIIBAmt>" & UVCase(NonComp40aiiChapXVIIBAmt_6) & "</ITRForm:NonComp40aiiChapXVIIBAmt>"
    If NonComp40aibChapXVIIBAmt_7 = "" Then
        writeXML " <ITRForm:NonComp40aibChapXVIIBAmt>0</ITRForm:NonComp40aibChapXVIIBAmt>"
    Else
        writeXML " <ITRForm:NonComp40aibChapXVIIBAmt>" & UCase(NonComp40aibChapXVIIBAmt_7) & "</ITRForm:NonComp40aibChapXVIIBAmt>"
    End If
    writeXML " <ITRForm:NonComp40aiiiChapXVIIBAmt>" & UVCase(NonComp40aiiiChapXVIIBAmt_6) & "</ITRForm:NonComp40aiiiChapXVIIBAmt>"

    If TaxAmtOnProfits_6 = "" Then
        writeXML " <ITRForm:TaxAmtOnProfits>0</ITRForm:TaxAmtOnProfits>"
    Else
        writeXML " <ITRForm:TaxAmtOnProfits>" & UCase(TaxAmtOnProfits_6) & "</ITRForm:TaxAmtOnProfits>"
    End If

    If WTAmt_7 = "" Then
        writeXML " <ITRForm:WTAmt>0</ITRForm:WTAmt>"
    Else
        writeXML " <ITRForm:WTAmt>" & UCase(WTAmt_7) & "</ITRForm:WTAmt>"
    End If

    If RolyatyOrServiceFee_6 <> "" Then
        writeXML " <ITRForm:RolyatyOrServiceFee>" & UCase(RolyatyOrServiceFee_6) & "</ITRForm:RolyatyOrServiceFee>"
    Else
        writeXML " <ITRForm:RolyatyOrServiceFee>0</ITRForm:RolyatyOrServiceFee>"
    End If

    If IntSalBonPartner_7 = "" Then
        writeXML " <ITRForm:IntSalBonPartner>0</ITRForm:IntSalBonPartner>"
    Else
        writeXML " <ITRForm:IntSalBonPartner>" & UCase(IntSalBonPartner_7) & "</ITRForm:IntSalBonPartner>"
    End If

    If OthDisallow_7 = "" Then
        writeXML " <ITRForm:OthDisallow>0</ITRForm:OthDisallow>"
    Else
        writeXML " <ITRForm:OthDisallow>" & UCase(OthDisallow_7) & "</ITRForm:OthDisallow>"
    End If

    If TotAmtDisallUs40_7 = "" Then
        writeXML " <ITRForm:TotAmtDisallUs40>0</ITRForm:TotAmtDisallUs40>"
    Else
        writeXML " <ITRForm:TotAmtDisallUs40>" & UCase(TotAmtDisallUs40_7) & "</ITRForm:TotAmtDisallUs40>"
    End If

    If AmtDisallUs40PyNowAll_7 = "" Then
        writeXML " <ITRForm:AmtDisallUs40PyNowAll>0</ITRForm:AmtDisallUs40PyNowAll>"
    Else
        writeXML " <ITRForm:AmtDisallUs40PyNowAll>" & UCase(AmtDisallUs40PyNowAll_7) & "</ITRForm:AmtDisallUs40PyNowAll>"
    End If

    writeXML " </ITRForm:AmtDisallUs40>"

    writeXML " <ITRForm:AmtDisallUs40A>"

    If AmtPaidUs40A2b_7 = "" Then
        writeXML " <ITRForm:AmtPaidUs40A2b>0</ITRForm:AmtPaidUs40A2b>"
    Else
        writeXML " <ITRForm:AmtPaidUs40A2b>" & UCase(AmtPaidUs40A2b_7) & "</ITRForm:AmtPaidUs40A2b>"
    End If

    If AmtGT20kCash_7 = "" Then
        writeXML " <ITRForm:AmtGT20kCash>0</ITRForm:AmtGT20kCash>"
    Else
        writeXML " <ITRForm:AmtGT20kCash>" & UCase(AmtGT20kCash_7) & "</ITRForm:AmtGT20kCash>"
    End If

    If ProvPmtGrat_7 = "" Then
        writeXML " <ITRForm:ProvPmtGrat>0</ITRForm:ProvPmtGrat>"
    Else
        writeXML " <ITRForm:ProvPmtGrat>" & UCase(ProvPmtGrat_7) & "</ITRForm:ProvPmtGrat>"
    End If

    If ContToSetupTrust_7 = "" Then
        writeXML " <ITRForm:ContToSetupTrust>0</ITRForm:ContToSetupTrust>"
    Else
        writeXML " <ITRForm:ContToSetupTrust>" & UCase(ContToSetupTrust_7) & "</ITRForm:ContToSetupTrust>"
    End If

    If OthDisallow1_7 = "" Then
        writeXML " <ITRForm:OthDisallow>0</ITRForm:OthDisallow>"
    Else
        writeXML " <ITRForm:OthDisallow>" & UCase(OthDisallow1_7) & "</ITRForm:OthDisallow>"
    End If

    If TotAmtDisallUs40A_7 = "" Then
        writeXML " <ITRForm:TotAmtDisallUs40A>0</ITRForm:TotAmtDisallUs40A>"
    Else
        writeXML " <ITRForm:TotAmtDisallUs40A>" & UCase(TotAmtDisallUs40A_7) & "</ITRForm:TotAmtDisallUs40A>"
    End If
    writeXML " </ITRForm:AmtDisallUs40A>"

    writeXML " <ITRForm:AmtDisallUs43BPyNowAll>"
    writeXML " <ITRForm:AmtUs43B>"

    If TaxDutyCesAmt_7 = "" Then
        writeXML " <ITRForm:TaxDutyCesAmt>0</ITRForm:TaxDutyCesAmt>"
    Else
        writeXML " <ITRForm:TaxDutyCesAmt>" & UCase(TaxDutyCesAmt_7) & "</ITRForm:TaxDutyCesAmt>"
    End If

    If ContToEmpPFSFGF_7 = "" Then
        writeXML " <ITRForm:ContToEmpPFSFGF>0</ITRForm:ContToEmpPFSFGF>"
    Else
        writeXML " <ITRForm:ContToEmpPFSFGF>" & UCase(ContToEmpPFSFGF_7) & "</ITRForm:ContToEmpPFSFGF>"
    End If

    If EmpBonusComm_7 = "" Then
        writeXML " <ITRForm:EmpBonusComm>0</ITRForm:EmpBonusComm>"
    Else
        writeXML " <ITRForm:EmpBonusComm>" & UCase(EmpBonusComm_7) & "</ITRForm:EmpBonusComm>"
    End If

    If IntPayaleToFI_7 = "" Then
        writeXML " <ITRForm:IntPayaleToFI>0</ITRForm:IntPayaleToFI>"
    Else
        writeXML " <ITRForm:IntPayaleToFI>" & UCase(IntPayaleToFI_7) & "</ITRForm:IntPayaleToFI>"
    End If

    If IntPayaleToFISchBank_7 = "" Then
        writeXML " <ITRForm:IntPayaleToFISchBank>0</ITRForm:IntPayaleToFISchBank>"
    Else
        writeXML " <ITRForm:IntPayaleToFISchBank>" & UCase(IntPayaleToFISchBank_7) & "</ITRForm:IntPayaleToFISchBank>"
    End If

    If LeaveEncashPayable_7 = "" Then
        writeXML " <ITRForm:LeaveEncashPayable>0</ITRForm:LeaveEncashPayable>"
    Else
        writeXML " <ITRForm:LeaveEncashPayable>" & UCase(LeaveEncashPayable_7) & "</ITRForm:LeaveEncashPayable>"
    End If

    If RailwayAssets_10_g = "" Then
        writeXML " <ITRForm:RailwayAssetsPayable>0</ITRForm:RailwayAssetsPayable>"
    Else
        writeXML " <ITRForm:RailwayAssetsPayable>" & UCase(RailwayAssets_10_g) & "</ITRForm:RailwayAssetsPayable>"
    End If

    If TotAmtUs43b_7 = "" Then
        writeXML " <ITRForm:TotAmtUs43b>0</ITRForm:TotAmtUs43b>"
    Else
        writeXML " <ITRForm:TotAmtUs43b>" & UCase(TotAmtUs43b_7) & "</ITRForm:TotAmtUs43b>"
    End If
    writeXML " </ITRForm:AmtUs43B>"
    writeXML " </ITRForm:AmtDisallUs43BPyNowAll>"

    writeXML " <ITRForm:AmtDisall43B>"
    writeXML " <ITRForm:AmtUs43B>"
    If TaxDutyCesAmt1_7 = "" Then
        writeXML " <ITRForm:TaxDutyCesAmt>0</ITRForm:TaxDutyCesAmt>"
    Else
        writeXML " <ITRForm:TaxDutyCesAmt>" & UCase(TaxDutyCesAmt1_7) & "</ITRForm:TaxDutyCesAmt>"
    End If

    If ContToEmpPFSFGF1_7 = "" Then
        writeXML " <ITRForm:ContToEmpPFSFGF>0</ITRForm:ContToEmpPFSFGF>"
    Else
        writeXML " <ITRForm:ContToEmpPFSFGF>" & UCase(ContToEmpPFSFGF1_7) & "</ITRForm:ContToEmpPFSFGF>"
    End If

    If EmpBonusComm1_7 = "" Then
        writeXML " <ITRForm:EmpBonusComm>0</ITRForm:EmpBonusComm>"
    Else
        writeXML " <ITRForm:EmpBonusComm>" & UCase(EmpBonusComm1_7) & "</ITRForm:EmpBonusComm>"
    End If

    If IntPayaleToFI1_7 = "" Then
        writeXML " <ITRForm:IntPayaleToFI>0</ITRForm:IntPayaleToFI>"
    Else
        writeXML " <ITRForm:IntPayaleToFI>" & UCase(IntPayaleToFI1_7) & "</ITRForm:IntPayaleToFI>"
    End If

    If Sheet5.Range("sheet7.IntPayaleToFI1_11da").value = "" Then
        writeXML " <ITRForm:SumPayaleLoanBrToFinComp>0</ITRForm:SumPayaleLoanBrToFinComp>"
    Else
        writeXML " <ITRForm:SumPayaleLoanBrToFinComp>" & Sheet5.Range("sheet7.IntPayaleToFI1_11da").value & "</ITRForm:SumPayaleLoanBrToFinComp>"
    End If

    If IntPayaleToFISchBank1_7 = "" Then
        writeXML " <ITRForm:IntPayaleToFISchBank>0</ITRForm:IntPayaleToFISchBank>"
    Else
        writeXML " <ITRForm:IntPayaleToFISchBank>" & UCase(IntPayaleToFISchBank1_7) & "</ITRForm:IntPayaleToFISchBank>"
    End If


    If LeaveEncashPayable1_7 = "" Then
        writeXML " <ITRForm:LeaveEncashPayable>0</ITRForm:LeaveEncashPayable>"
    Else
        writeXML " <ITRForm:LeaveEncashPayable>" & UCase(LeaveEncashPayable1_7) & "</ITRForm:LeaveEncashPayable>"
    End If

    If RailwayAssetsPayable_7 = "" Then
        writeXML " <ITRForm:RailwayAssetsPayable>0</ITRForm:RailwayAssetsPayable>"
    Else
        writeXML " <ITRForm:RailwayAssetsPayable>" & UCase(RailwayAssetsPayable_7) & "</ITRForm:RailwayAssetsPayable>"
    End If

    If TotAmtUs43b1_7 = "" Then
        writeXML " <ITRForm:TotAmtUs43b>0</ITRForm:TotAmtUs43b>"
    Else
        writeXML " <ITRForm:TotAmtUs43b>" & UCase(TotAmtUs43b1_7) & "</ITRForm:TotAmtUs43b>"
    End If
    writeXML " </ITRForm:AmtUs43B>"
    writeXML " </ITRForm:AmtDisall43B>"

    writeXML " <ITRForm:AmtExciseCustomsVATOutstanding>"
    writeXML " <ITRForm:ExciseCustomsVAT>"

    If UnionExciseDuty_7 = "" Then
        writeXML " <ITRForm:UnionExciseDuty>0</ITRForm:UnionExciseDuty>"
    Else
        writeXML " <ITRForm:UnionExciseDuty>" & UCase(UnionExciseDuty_7) & "</ITRForm:UnionExciseDuty>"
    End If

    If ServiceTax_7 = "" Then
        writeXML " <ITRForm:ServiceTax>0</ITRForm:ServiceTax>"
    Else
        writeXML " <ITRForm:ServiceTax>" & UCase(ServiceTax_7) & "</ITRForm:ServiceTax>"
    End If

    If VATorSaleTax_7 = "" Then
        writeXML " <ITRForm:VATorSaleTax>0</ITRForm:VATorSaleTax>"
    Else
        writeXML " <ITRForm:VATorSaleTax>" & UCase(VATorSaleTax_7) & "</ITRForm:VATorSaleTax>"
    End If

    If CGST_7 = "" Then
        writeXML " <ITRForm:CentralGoodServiceTax>0</ITRForm:CentralGoodServiceTax>"
    Else
        writeXML " <ITRForm:CentralGoodServiceTax>" & UCase(CGST_7) & "</ITRForm:CentralGoodServiceTax>"
    End If

    If SGST_7 = "" Then
        writeXML " <ITRForm:StateGoodServiceTax>0</ITRForm:StateGoodServiceTax>"
    Else
        writeXML " <ITRForm:StateGoodServiceTax>" & UCase(SGST_7) & "</ITRForm:StateGoodServiceTax>"
    End If

    If IGST_7 = "" Then
        writeXML " <ITRForm:IntegratedGoodServiceTax>0</ITRForm:IntegratedGoodServiceTax>"
    Else
        writeXML " <ITRForm:IntegratedGoodServiceTax>" & UCase(IGST_7) & "</ITRForm:IntegratedGoodServiceTax>"
    End If

    If UTGST_7 = "" Then
        writeXML " <ITRForm:UnionTerrGoodServiceTax>0</ITRForm:UnionTerrGoodServiceTax>"
    Else
        writeXML " <ITRForm:UnionTerrGoodServiceTax>" & UCase(UTGST_7) & "</ITRForm:UnionTerrGoodServiceTax>"
    End If


    If OthDutyTaxCess_7 = "" Then
        writeXML " <ITRForm:OthDutyTaxCess>0</ITRForm:OthDutyTaxCess>"
    Else
        writeXML " <ITRForm:OthDutyTaxCess>" & UCase(OthDutyTaxCess_7) & "</ITRForm:OthDutyTaxCess>"
    End If

    If TotExciseCustomsVAT_7 = "" Then
        writeXML " <ITRForm:TotExciseCustomsVAT>0</ITRForm:TotExciseCustomsVAT>"
    Else
        writeXML " <ITRForm:TotExciseCustomsVAT>" & UCase(TotExciseCustomsVAT_7) & "</ITRForm:TotExciseCustomsVAT>"
    End If


    writeXML " </ITRForm:ExciseCustomsVAT>"
    writeXML " </ITRForm:AmtExciseCustomsVATOutstanding>"

    If DeemedProfUs33ABs_7 = "" Then
        writeXML " <ITRForm:DeemedProfUs33ABs>0</ITRForm:DeemedProfUs33ABs>"
    Else
        writeXML " <ITRForm:DeemedProfUs33ABs>" & UCase(DeemedProfUs33ABs_7) & "</ITRForm:DeemedProfUs33ABs>"
    End If

    If DeemedProfUs33AB_7 = "" Then
        writeXML " <ITRForm:DeemedProfUs33AB>0</ITRForm:DeemedProfUs33AB>"
    Else
        writeXML " <ITRForm:DeemedProfUs33AB>" & UCase(DeemedProfUs33AB_7) & "</ITRForm:DeemedProfUs33AB>"
    End If

    If DeemedProfUs33ABA_7 = "" Then
        writeXML " <ITRForm:DeemedProfUs33ABA>0</ITRForm:DeemedProfUs33ABA>"
    Else
        writeXML " <ITRForm:DeemedProfUs33ABA>" & UCase(DeemedProfUs33ABA_7) & "</ITRForm:DeemedProfUs33ABA>"
    End If

    If ProfTaxAmtUs41_7 = "" Then
        writeXML " <ITRForm:ProfTaxAmtUs41>0</ITRForm:ProfTaxAmtUs41>"
    Else
        writeXML " <ITRForm:ProfTaxAmtUs41>" & UCase(ProfTaxAmtUs41_7) & "</ITRForm:ProfTaxAmtUs41>"
    End If

    If PriorAmtIncCrDrPL_7 = "" Then
        writeXML " <ITRForm:PriorAmtIncCrDrPL>0</ITRForm:PriorAmtIncCrDrPL>"
    Else
        writeXML " <ITRForm:PriorAmtIncCrDrPL>" & UCase(PriorAmtIncCrDrPL_7) & "</ITRForm:PriorAmtIncCrDrPL>"
    End If

    If Sheet5.Range("sheet7.AmountOfExpenditure14A").value = "" Then
        writeXML " <ITRForm:AmountOfExpDisAllwUs14A>0</ITRForm:AmountOfExpDisAllwUs14A>"
    Else
        writeXML " <ITRForm:AmountOfExpDisAllwUs14A>" & UVCase(Sheet5.Range("sheet7.AmountOfExpenditure14A").value) & "</ITRForm:AmountOfExpDisAllwUs14A>"
    End If

    Dim Section92CE_Flag As Variant
    Section92CE_Flag = Sheet5.Range("Sheet7.Section92CE_Flag")
    If Section92CE_Flag = "Yes" Then
    Section92CE_Flag = "Y"
    ElseIf Section92CE_Flag = "No" Then
    Section92CE_Flag = "N"
    End If

    If Section92CE_Flag <> "" And Section92CE_Flag <> "(Select)" Then
        writeXML " <ITRForm:ScheduleTPSAFlg>" & Section92CE_Flag & "</ITRForm:ScheduleTPSAFlg>"
        Else
        writeXML " <ITRForm:ScheduleTPSAFlg>" & "N" & "</ITRForm:ScheduleTPSAFlg>"
    End If

    writeXML " </ITRForm:PARTA_OI>"
End Function


Function PARTA_OD()
    Dim i As Long

    If UBound(ItemName_QDTradingConcern) > 0 Or _
         UBound(ItemName_QDRawMaterial) > 0 Or _
         UBound(ItemName_QDFinishrByProd) > 0 Then

            writeXML "      <ITRForm:PARTA_QD>"
            If Not IsEmpty(ItemName_QDTradingConcern) And UBound(ItemName_QDTradingConcern) > 0 Then
            writeXML "          <ITRForm:TradingConcern>"
                For i = 1 To UBound(ItemName_QDTradingConcern)
                    writeXML "               <ITRForm:QuantitDet>"
                    If ItemName_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:ItemName>" & UCase(ItemName_QDTradingConcern(i)) & "</ITRForm:ItemName>"
'                    Else
'                        writeXML "                   <ITRForm:ItemName/>"
                    End If

                    If UnitOfMeasure_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:UnitOfMeasure>" & UCase(UnitOfMeasure_QDTradingConcern(i)) & "</ITRForm:UnitOfMeasure>"
'                    Else
'                        writeXML "                   <ITRForm:UnitOfMeasure/>"
                    End If

                    If OpeningStock_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:OpeningStock>" & UCase(OpeningStock_QDTradingConcern(i)) & "</ITRForm:OpeningStock>"
                    Else
                        writeXML "                   <ITRForm:OpeningStock>" & UCase(DefaultOpeningStock_QDTradingConcern) & "</ITRForm:OpeningStock>"
                    End If

                    If PurchaseQty_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:PurchaseQty>" & UCase(PurchaseQty_QDTradingConcern(i)) & "</ITRForm:PurchaseQty>"
                    Else
                        writeXML "                   <ITRForm:PurchaseQty>" & UCase(DefaultPurchaseQty_QDTradingConcern) & "</ITRForm:PurchaseQty>"
                    End If

                    If SaleQty_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:SaleQty>" & UCase(SaleQty_QDTradingConcern(i)) & "</ITRForm:SaleQty>"
                    Else
                        writeXML "                   <ITRForm:SaleQty>" & UCase(DefaultSaleQty_QDTradingConcern) & "</ITRForm:SaleQty>"
                    End If

                    If ClgStock_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:ClgStock>" & UCase(ClgStock_QDTradingConcern(i)) & "</ITRForm:ClgStock>"
                    Else
                        writeXML "                   <ITRForm:ClgStock>" & UCase(DefaultClgStock_QDTradingConcern) & "</ITRForm:ClgStock>"
                    End If

                    If AnyShortExces_QDTradingConcern(i) <> "" Then
                        writeXML "                   <ITRForm:AnyShortExces>" & UCase(AnyShortExces_QDTradingConcern(i)) & "</ITRForm:AnyShortExces>"
                    Else
                        writeXML "                   <ITRForm:AnyShortExces>" & UCase(DefaultAnyShortExces_QDTradingConcern) & "</ITRForm:AnyShortExces>"
                    End If
                    writeXML "               </ITRForm:QuantitDet>"
                Next
            writeXML "          </ITRForm:TradingConcern>"
            End If

        If Not IsEmpty(ItemName_QDRawMaterial) And UBound(ItemName_QDRawMaterial) > 0 Then
            writeXML "          <ITRForm:ManfactrConcern>"
            writeXML "              <ITRForm:RawMaterial>"

            For i = 1 To UBound(ItemName_QDRawMaterial)
                writeXML "                   <ITRForm:QuantitDet>"
                If ItemName_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:ItemName>" & UCase(ItemName_QDRawMaterial(i)) & "</ITRForm:ItemName>"
                Else
                    writeXML "                       <ITRForm:ItemName>" & UCase(DefaultItemName_QDRawMaterial) & "</ITRForm:ItemName>"
                End If

                If UnitOfMeasure_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:UnitOfMeasure>" & UCase(UnitOfMeasure_QDRawMaterial(i)) & "</ITRForm:UnitOfMeasure>"
                Else
                    writeXML "                       <ITRForm:UnitOfMeasure>" & UCase(DefaultUnitOfMeasure_QDRawMaterial) & "</ITRForm:UnitOfMeasure>"
                End If

                If OpeningStock_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:OpeningStock>" & UCase(OpeningStock_QDRawMaterial(i)) & "</ITRForm:OpeningStock>"
                Else
                    writeXML "                       <ITRForm:OpeningStock>" & UCase(DefaultOpeningStock_QDRawMaterial) & "</ITRForm:OpeningStock>"
                End If

                If PurchaseQty_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:PurchaseQty>" & UCase(PurchaseQty_QDRawMaterial(i)) & "</ITRForm:PurchaseQty>"
                Else
                    writeXML "                       <ITRForm:PurchaseQty>" & UCase(DefaultPurchaseQty_QDRawMaterial) & "</ITRForm:PurchaseQty>"
                End If

                If PrevYrConsum_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:PrevYrConsum>" & UCase(PrevYrConsum_QDRawMaterial(i)) & "</ITRForm:PrevYrConsum>"
                Else
                    writeXML "                       <ITRForm:PrevYrConsum>" & UCase(DefaultPrevYrConsum_QDRawMaterial) & "</ITRForm:PrevYrConsum>"
                End If

                If SaleQty_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:SaleQty>" & UCase(SaleQty_QDRawMaterial(i)) & "</ITRForm:SaleQty>"
                Else
                    writeXML "                       <ITRForm:SaleQty>" & UCase(DefaultSaleQty_QDRawMaterial) & "</ITRForm:SaleQty>"
                End If

                If ClgStock_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:ClgStock>" & UCase(ClgStock_QDRawMaterial(i)) & "</ITRForm:ClgStock>"
                Else
                    writeXML "                       <ITRForm:ClgStock>" & UCase(DefaultClgStock_QDRawMaterial) & "</ITRForm:ClgStock>"
                End If

                If yldFinisProd_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:yldFinisProd>" & UCase(yldFinisProd_QDRawMaterial(i)) & "</ITRForm:yldFinisProd>"
                Else
                    writeXML "                       <ITRForm:yldFinisProd>" & UCase(DefaultyldFinisProd_QDRawMaterial) & "</ITRForm:yldFinisProd>"
                End If

                If PercentYld_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:PercentYld>" & UCase(PercentYld_QDRawMaterial(i)) & "</ITRForm:PercentYld>"
                Else
                    writeXML "                       <ITRForm:PercentYld>" & UCase(DefaultPercentYld_QDRawMaterial) & "</ITRForm:PercentYld>"
                End If

                If AnyShortExces_QDRawMaterial(i) <> "" Then
                    writeXML "                       <ITRForm:AnyShortExces>" & UCase(AnyShortExces_QDRawMaterial(i)) & "</ITRForm:AnyShortExces>"
                Else
                    writeXML "                       <ITRForm:AnyShortExces>" & UCase(DefaultAnyShortExces_QDRawMaterial) & "</ITRForm:AnyShortExces>"
                End If
                writeXML "                   </ITRForm:QuantitDet>"
            Next
            writeXML "              </ITRForm:RawMaterial>"

            writeXML "              <ITRForm:FinishrByProd>"

            For i = 1 To UBound(ItemName_QDFinishrByProd)
                writeXML "                   <ITRForm:QuantitDet>"
                If ItemName_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:ItemName>" & UCase(ItemName_QDFinishrByProd(i)) & "</ITRForm:ItemName>"
                Else
                    writeXML "                       <ITRForm:ItemName>" & UCase(DefaultItemName_QDFinishrByProd) & "</ITRForm:ItemName>"
                End If

                If UnitOfMeasure_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:UnitOfMeasure>" & UCase(UnitOfMeasure_QDFinishrByProd(i)) & "</ITRForm:UnitOfMeasure>"
                Else
                    writeXML "                       <ITRForm:UnitOfMeasure>" & UCase(DefaultUnitOfMeasure_QDFinishrByProd) & "</ITRForm:UnitOfMeasure>"
                End If

                If OpeningStock_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:OpeningStock>" & UCase(OpeningStock_QDFinishrByProd(i)) & "</ITRForm:OpeningStock>"
                Else
                    writeXML "                       <ITRForm:OpeningStock>" & UCase(DefaultOpeningStock_QDFinishrByProd) & "</ITRForm:OpeningStock>"
                End If

                If PurchaseQty_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:PurchaseQty>" & UCase(PurchaseQty_QDFinishrByProd(i)) & "</ITRForm:PurchaseQty>"
                Else
                    writeXML "                       <ITRForm:PurchaseQty>" & UCase(DefaultPurchaseQty_QDFinishrByProd) & "</ITRForm:PurchaseQty>"
                End If

                If PrevYrConsum_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:PrevyrManfact>" & UCase(PrevYrConsum_QDFinishrByProd(i)) & "</ITRForm:PrevyrManfact>"
                Else
                    writeXML "                       <ITRForm:PrevyrManfact>" & UCase(DefaultPrevYrConsum_QDFinishrByProd) & "</ITRForm:PrevyrManfact>"
                End If

                If SaleQty_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:SaleQty>" & UCase(SaleQty_QDFinishrByProd(i)) & "</ITRForm:SaleQty>"
                Else
                    writeXML "                       <ITRForm:SaleQty>" & UCase(DefaultSaleQty_QDFinishrByProd) & "</ITRForm:SaleQty>"
                End If

                If ClgStock_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:ClgStock>" & UCase(ClgStock_QDFinishrByProd(i)) & "</ITRForm:ClgStock>"
                Else
                    writeXML "                       <ITRForm:ClgStock>" & UCase(DefaultClgStock_QDFinishrByProd) & "</ITRForm:ClgStock>"
                End If

                If AnyShortExces_QDFinishrByProd(i) <> "" Then
                    writeXML "                       <ITRForm:AnyShortExces>" & UCase(AnyShortExces_QDFinishrByProd(i)) & "</ITRForm:AnyShortExces>"
                Else
                    writeXML "                       <ITRForm:AnyShortExces>" & UCase(DefaultAnyShortExces_QDFinishrByProd) & "</ITRForm:AnyShortExces>"
                End If
                writeXML "                   </ITRForm:QuantitDet>"
            Next
            writeXML "              </ITRForm:FinishrByProd>"
            writeXML "          </ITRForm:ManfactrConcern>"
        End If
        writeXML "      </ITRForm:PARTA_QD>"
    End If

End Function

Function PARTB_TIXML()
    Dim Deemeds As Variant

    writeXML " <ITRForm:PartB-TI>"

    If IncomeFromSal_8b = "" Then
        writeXML "<ITRForm:Salaries />"
    Else
        writeXML "<ITRForm:Salaries>" & UCase(IncomeFromSal_8b) & "</ITRForm:Salaries>"
    End If

    If IncomeFromHP_8b = "" Then
        writeXML " <ITRForm:IncomeFromHP/>"
    Else
        writeXML " <ITRForm:IncomeFromHP>" & UCase(IncomeFromHP_8b) & "</ITRForm:IncomeFromHP>"
    End If

    writeXML "      <ITRForm:ProfBusGain>"
    If ProfGainNoSpecBus_8b = "" Then
        writeXML " <ITRForm:ProfGainNoSpecBus>0</ITRForm:ProfGainNoSpecBus>"
    Else
        writeXML " <ITRForm:ProfGainNoSpecBus>" & UCase(ProfGainNoSpecBus_8b) & "</ITRForm:ProfGainNoSpecBus>"
    End If

    If ProfGainSpecBus_8b = "" Then
        writeXML " <ITRForm:ProfGainSpecBus>0</ITRForm:ProfGainSpecBus>"
    Else
        writeXML " <ITRForm:ProfGainSpecBus>" & UCase(ProfGainSpecBus_8b) & "</ITRForm:ProfGainSpecBus>"
    End If

    If ProfGainSpecifiedBus_8b = "" Then
        writeXML " <ITRForm:ProfGainSpecifiedBus>0</ITRForm:ProfGainSpecifiedBus>"
    Else
        writeXML " <ITRForm:ProfGainSpecifiedBus>" & UCase(ProfGainSpecifiedBus_8b) & "</ITRForm:ProfGainSpecifiedBus>"
    End If

    If ProfGain115BF_8b = "" Then
        writeXML " <ITRForm:ProfIncome115BBF>0</ITRForm:ProfIncome115BBF>"
    Else
        writeXML " <ITRForm:ProfIncome115BBF>" & UCase(ProfGain115BF_8b) & "</ITRForm:ProfIncome115BBF>"
    End If

'    If ProfGain115BG_8b = "" Then
'        writeXML " <ITRForm:ProfIncome115BBG>0</ITRForm:ProfIncome115BBG>"
'    Else
'        writeXML " <ITRForm:ProfIncome115BBG>" & UCase(ProfGain115BG_8b) & "</ITRForm:ProfIncome115BBG>"
'    End If

    If TotProfBusGain_8b = "" Then
        writeXML " <ITRForm:TotProfBusGain>0</ITRForm:TotProfBusGain>"
    Else
        writeXML " <ITRForm:TotProfBusGain>" & UCase(TotProfBusGain_8b) & "</ITRForm:TotProfBusGain>"
    End If
    writeXML "      </ITRForm:ProfBusGain>"

    writeXML "      <ITRForm:CapGain>"
    writeXML "      <ITRForm:ShortTerm>"

    If ShortTermUs111A_8b = "" Then
        writeXML " <ITRForm:ShortTerm15Per>0</ITRForm:ShortTerm15Per>"
    Else
        writeXML " <ITRForm:ShortTerm15Per>" & UCase(ShortTermUs111A_8b) & "</ITRForm:ShortTerm15Per>"
    End If

    writeXML " <ITRForm:ShortTerm30Per>" & UVCase(ShortTerm30_8b) & "</ITRForm:ShortTerm30Per>"

    If ShortTermOther_8b = "" Then
        writeXML " <ITRForm:ShortTermAppRate>0</ITRForm:ShortTermAppRate>"
    Else
        writeXML " <ITRForm:ShortTermAppRate>" & UCase(ShortTermOther_8b) & "</ITRForm:ShortTermAppRate>"
    End If

     If ShortTermSpslRate_8b = "" Then
        writeXML " <ITRForm:ShortTermSplRateDTAA>0</ITRForm:ShortTermSplRateDTAA>"
    Else
        writeXML " <ITRForm:ShortTermSplRateDTAA>" & UCase(ShortTermSpslRate_8b) & "</ITRForm:ShortTermSplRateDTAA>"
    End If

    If TotalShortTerm_8b = "" Then
        writeXML " <ITRForm:TotalShortTerm>0</ITRForm:TotalShortTerm>"
    Else
        writeXML " <ITRForm:TotalShortTerm>" & UCase(TotalShortTerm_8b) & "</ITRForm:TotalShortTerm>"
    End If

    writeXML "      </ITRForm:ShortTerm>"

    writeXML "        <ITRForm:LongTerm>"
    writeXML "          <ITRForm:LongTerm10Per>" & UVCase(LongTerm_8b_P) & "</ITRForm:LongTerm10Per>"
    writeXML "          <ITRForm:LongTerm20Per>" & UVCase(LongTerm_8b_NP) & "</ITRForm:LongTerm20Per>"
    writeXML "          <ITRForm:LongTermSplRateDTAA>" & UVCase(LongTerm_8b_SR) & "</ITRForm:LongTermSplRateDTAA>"

    writeXML "          <ITRForm:TotalLongTerm>" & UVCase(LongTerm_8b) & "</ITRForm:TotalLongTerm>"
    writeXML "        </ITRForm:LongTerm>"


    If TotalCapGains_8b = "" Then
        writeXML " <ITRForm:TotalCapGains>0</ITRForm:TotalCapGains>"
    Else
        writeXML " <ITRForm:TotalCapGains>" & UCase(TotalCapGains_8b) & "</ITRForm:TotalCapGains>"
    End If

    writeXML "  </ITRForm:CapGain>"

    writeXML "  <ITRForm:IncFromOS>"

    If OtherSrcThanOwnRaceHorse_8b = "" Then
        writeXML " <ITRForm:OtherSrcThanOwnRaceHorse>0</ITRForm:OtherSrcThanOwnRaceHorse>"
    Else
        writeXML " <ITRForm:OtherSrcThanOwnRaceHorse>" & UCase(OtherSrcThanOwnRaceHorse_8b) & "</ITRForm:OtherSrcThanOwnRaceHorse>"
    End If

    If WinLotteriesRacesGambling_8b = "" Then
        writeXML " <ITRForm:IncChargblSplRate>0</ITRForm:IncChargblSplRate>"
    Else
        writeXML " <ITRForm:IncChargblSplRate>" & UCase(WinLotteriesRacesGambling_8b) & "</ITRForm:IncChargblSplRate>"
    End If

    If FromOwnRaceHorse_8b = "" Then
        writeXML " <ITRForm:FromOwnRaceHorse>0</ITRForm:FromOwnRaceHorse>"
    Else
        writeXML " <ITRForm:FromOwnRaceHorse>" & UCase(FromOwnRaceHorse_8b) & "</ITRForm:FromOwnRaceHorse>"
    End If

    If TotIncFromOS_8b = "" Then
        writeXML " <ITRForm:TotIncFromOS>0</ITRForm:TotIncFromOS>"
    Else
        writeXML " <ITRForm:TotIncFromOS>" & UCase(TotIncFromOS_8b) & "</ITRForm:TotIncFromOS>"
    End If
    writeXML "      </ITRForm:IncFromOS>"

    If TotalTI_8b = "" Then
        writeXML " <ITRForm:TotalTI>0</ITRForm:TotalTI>"
    Else
        writeXML " <ITRForm:TotalTI>" & UCase(TotalTI_8b) & "</ITRForm:TotalTI>"
    End If

    If CurrentYearLoss_8b = "" Then
        writeXML " <ITRForm:CurrentYearLoss>0</ITRForm:CurrentYearLoss>"
    Else
        writeXML " <ITRForm:CurrentYearLoss>" & UCase(CurrentYearLoss_8b) & "</ITRForm:CurrentYearLoss>"
    End If

    If BalanceAfterSetoffLosses_8b = "" Then
        writeXML " <ITRForm:BalanceAfterSetoffLosses>0</ITRForm:BalanceAfterSetoffLosses>"
    Else
        writeXML " <ITRForm:BalanceAfterSetoffLosses>" & UCase(BalanceAfterSetoffLosses_8b) & "</ITRForm:BalanceAfterSetoffLosses>"
    End If

    If BroughtFwdLossesSetoff_8b = "" Then
        writeXML " <ITRForm:BroughtFwdLossesSetoff>0</ITRForm:BroughtFwdLossesSetoff>"
    Else
        writeXML " <ITRForm:BroughtFwdLossesSetoff>" & UCase(BroughtFwdLossesSetoff_8b) & "</ITRForm:BroughtFwdLossesSetoff>"
    End If

    If GrossTotalIncome_8b = "" Then
        writeXML " <ITRForm:GrossTotalIncome>0</ITRForm:GrossTotalIncome>"
    Else
        writeXML " <ITRForm:GrossTotalIncome>" & UCase(GrossTotalIncome_8b) & "</ITRForm:GrossTotalIncome>"
    End If

    If IncChargeableTaxSplRates_8b = "" Then
        writeXML " <ITRForm:IncChargeTaxSplRate111A112>0</ITRForm:IncChargeTaxSplRate111A112>"
    Else
        writeXML " <ITRForm:IncChargeTaxSplRate111A112>" & UCase(IncChargeableTaxSplRates_8b) & "</ITRForm:IncChargeTaxSplRate111A112>"
    End If


    writeXML "<ITRForm:DeductionsUndSchVIADtl>"
    writeXML "<ITRForm:PartBchapterVIA>" & UVCase(PartBchapterVIA_8b) & "</ITRForm:PartBchapterVIA>"
    writeXML "<ITRForm:PartCchapterVIA>" & UVCase(PartCchapterVIA_8b) & "</ITRForm:PartCchapterVIA>"
    writeXML "<ITRForm:TotDeductUndSchVIA>" & UVCase(DeductionsUnderScheduleVIA_8b) & "</ITRForm:TotDeductUndSchVIA>"
    writeXML "</ITRForm:DeductionsUndSchVIADtl>"

    writeXML " <ITRForm:DeductionsUnder10Aor10AA>" & UVCase(DeductionsUnder10Aor10AA_8b) & "</ITRForm:DeductionsUnder10Aor10AA>"

    If TotalIncome_8b = "" Then
        writeXML " <ITRForm:TotalIncome>0</ITRForm:TotalIncome>"
    Else
        writeXML " <ITRForm:TotalIncome>" & UCase(TotalIncome_8b) & "</ITRForm:TotalIncome>"
    End If

    If IncChargeTaxSplRate111A112_8b = "" Then
        writeXML " <ITRForm:IncChargeableTaxSplRates>0</ITRForm:IncChargeableTaxSplRates>"
    Else
        writeXML " <ITRForm:IncChargeableTaxSplRates>" & UCase(IncChargeTaxSplRate111A112_8b) & "</ITRForm:IncChargeableTaxSplRates>"
    End If


    If NetAgricultureIncomeOrOtherIncomeForRate_8b = "" Then
        writeXML " <ITRForm:NetAgricultureIncomeOrOtherIncomeForRate>0</ITRForm:NetAgricultureIncomeOrOtherIncomeForRate>"
    Else
        writeXML " <ITRForm:NetAgricultureIncomeOrOtherIncomeForRate>" & UCase(NetAgricultureIncomeOrOtherIncomeForRate_8b) & "</ITRForm:NetAgricultureIncomeOrOtherIncomeForRate>"
    End If

    If AggregateIncome_8b = "" Then
        writeXML " <ITRForm:AggregateIncome>0</ITRForm:AggregateIncome>"
    Else
        writeXML " <ITRForm:AggregateIncome>" & UCase(AggregateIncome_8b) & "</ITRForm:AggregateIncome>"
    End If

    If LossesOfCurrentYearCarriedFwd_8b = "" Then
        writeXML " <ITRForm:LossesOfCurrentYearCarriedFwd>0</ITRForm:LossesOfCurrentYearCarriedFwd>"
    Else
        writeXML " <ITRForm:LossesOfCurrentYearCarriedFwd>" & UCase(LossesOfCurrentYearCarriedFwd_8b) & "</ITRForm:LossesOfCurrentYearCarriedFwd>"
    End If


      If LossesOfCurrentYearCarriedFwd_8b = "" Then
        writeXML " <ITRForm:DeemedIncomeUs115JC>0</ITRForm:DeemedIncomeUs115JC>"
    Else
        writeXML "<ITRForm:DeemedIncomeUs115JC>" & Sheet7.Range("Sheet8b.Deemed115JC").value & "</ITRForm:DeemedIncomeUs115JC>"
    End If

    writeXML "      </ITRForm:PartB-TI>"
End Function


Function PARTB_TTIXML()
Dim i As Variant
writeXML "<ITRForm:PartB_TTI>"

writeXML "<ITRForm:ComputationOfTaxLiability>"

writeXML "  <ITRForm:TaxPayableOnDeemedTI>"

    If Sheet7.Range("sheet9.TaxDeemedTISec115JC").value = "" Then
        writeXML " <ITRForm:TaxDeemedTISec115JC>0</ITRForm:TaxDeemedTISec115JC>"
    Else
        writeXML " <ITRForm:TaxDeemedTISec115JC>" & Sheet7.Range("sheet9.TaxDeemedTISec115JC").value & "</ITRForm:TaxDeemedTISec115JC>"
    End If

    If Sheet7.Range("sheet9.deemeds").value = "" Then
        writeXML " <ITRForm:SurchargeOnAboveCrore>0</ITRForm:SurchargeOnAboveCrore>"
    Else
        writeXML " <ITRForm:SurchargeOnAboveCrore>" & Sheet7.Range("sheet9.deemeds").value & "</ITRForm:SurchargeOnAboveCrore>"
    End If

    If Sheet7.Range("sheet9.EducationCess_DI").value = "" Then
        writeXML " <ITRForm:EducationCess>0</ITRForm:EducationCess>"
    Else
        writeXML " <ITRForm:EducationCess>" & Sheet7.Range("sheet9.EducationCess_DI").value & "</ITRForm:EducationCess>"
    End If

    If Sheet7.Range("sheet9.TotalTax_DI").value = "" Then
        writeXML " <ITRForm:TotalTax>0</ITRForm:TotalTax>"
    Else
        writeXML " <ITRForm:TotalTax>" & Sheet7.Range("sheet9.TotalTax_DI").value & "</ITRForm:TotalTax>"
    End If
    writeXML "</ITRForm:TaxPayableOnDeemedTI>"

   writeXML "<ITRForm:TaxPayableOnTI>"

     If Sheet7.Range("Sheet9.TaxAtNormalRatesOnAggrInc").value = "" Then
        writeXML " <ITRForm:TaxAtNormalRatesOnAggrInc>0</ITRForm:TaxAtNormalRatesOnAggrInc>"
    Else
        writeXML " <ITRForm:TaxAtNormalRatesOnAggrInc>" & Sheet7.Range("Sheet9.TaxAtNormalRatesOnAggrInc").value & "</ITRForm:TaxAtNormalRatesOnAggrInc>"
    End If

    If Sheet7.Range("sheet9.TaxAtSpecialRates").value = "" Then
        writeXML " <ITRForm:TaxAtSpecialRates>0</ITRForm:TaxAtSpecialRates>"
    Else
        writeXML " <ITRForm:TaxAtSpecialRates>" & Sheet7.Range("sheet9.TaxAtSpecialRates").value & "</ITRForm:TaxAtSpecialRates>"
    End If

     If Sheet7.Range("sheet9.RebateOnAgriInc").value = "" Then
        writeXML " <ITRForm:RebateOnAgriInc>0</ITRForm:RebateOnAgriInc>"
    Else
        writeXML " <ITRForm:RebateOnAgriInc>" & Sheet7.Range("sheet9.RebateOnAgriInc").value & "</ITRForm:RebateOnAgriInc>"
    End If

    If Sheet7.Range("sheet9.TaxPayableOnTotInc").value = "" Then
        writeXML " <ITRForm:TaxPayableOnTotInc>0</ITRForm:TaxPayableOnTotInc>"
    Else
        writeXML " <ITRForm:TaxPayableOnTotInc>" & Sheet7.Range("sheet9.TaxPayableOnTotInc").value & "</ITRForm:TaxPayableOnTotInc>"
    End If

        If Sheet7.Range("Sheet9.RebateUs88E").value = "" Then
        writeXML " <ITRForm:Rebate87A>0</ITRForm:Rebate87A>"
    Else
        writeXML " <ITRForm:Rebate87A>" & Sheet7.Range("Sheet9.RebateUs88E").value & "</ITRForm:Rebate87A>"
    End If

   If Sheet7.Range("Sheet9.BalTaxPayable").value = "" Then
        writeXML " <ITRForm:TaxPayableOnRebate>0</ITRForm:TaxPayableOnRebate>"
    Else
        writeXML " <ITRForm:TaxPayableOnRebate>" & Sheet7.Range("Sheet9.BalTaxPayable").value & "</ITRForm:TaxPayableOnRebate>"
    End If

    If Sheet7.Range("Sheet9.Surcharge_i").value = "" Then
        writeXML " <ITRForm:Surcharge25ofSI>0</ITRForm:Surcharge25ofSI>"
    Else
        writeXML " <ITRForm:Surcharge25ofSI>" & Sheet7.Range("Sheet9.Surcharge_i").value & "</ITRForm:Surcharge25ofSI>"
    End If

    If Sheet7.Range("Sheet9.Surcharge_ii").value = "" Then
        writeXML " <ITRForm:SurchargeOnAboveCrore>0</ITRForm:SurchargeOnAboveCrore>"
    Else
        writeXML " <ITRForm:SurchargeOnAboveCrore>" & Sheet7.Range("Sheet9.Surcharge_ii").value & "</ITRForm:SurchargeOnAboveCrore>"
    End If

    If Sheet7.Range("Sheet9.SurchargeOnTaxPayable").value = "" Then
        writeXML " <ITRForm:TotalSurcharge>0</ITRForm:TotalSurcharge>"
    Else
        writeXML " <ITRForm:TotalSurcharge>" & Sheet7.Range("Sheet9.SurchargeOnTaxPayable").value & "</ITRForm:TotalSurcharge>"
    End If

    If Sheet7.Range("Sheet9.EducationCess").value = "" Then
        writeXML " <ITRForm:EducationCess>0</ITRForm:EducationCess>"
    Else
        writeXML " <ITRForm:EducationCess>" & Sheet7.Range("Sheet9.EducationCess").value & "</ITRForm:EducationCess>"
    End If

    If Sheet7.Range("Sheet9.GrossTaxLiability").value = "" Then
        writeXML " <ITRForm:GrossTaxLiability>0</ITRForm:GrossTaxLiability>"
    Else
        writeXML " <ITRForm:GrossTaxLiability>" & Sheet7.Range("Sheet9.GrossTaxLiability").value & "</ITRForm:GrossTaxLiability>"
    End If
    writeXML "</ITRForm:TaxPayableOnTI>"


    If Sheet7.Range("sheet9.GrossTaxPayable").value = "" Then
        writeXML " <ITRForm:GrossTaxPayable>0</ITRForm:GrossTaxPayable>"
    Else
        writeXML " <ITRForm:GrossTaxPayable>" & Sheet7.Range("sheet9.GrossTaxPayable").value & "</ITRForm:GrossTaxPayable>"
    End If

        If Sheet7.Range("sheet9.CreditUS115JD").value = "" Then
        writeXML " <ITRForm:CreditUS115JD>0</ITRForm:CreditUS115JD>"
    Else
        writeXML " <ITRForm:CreditUS115JD>" & Sheet7.Range("sheet9.CreditUS115JD").value & "</ITRForm:CreditUS115JD>"
    End If

    If Sheet7.Range("sheet9.TaxPayAfterCreditUs115JD").value = "" Then
        writeXML " <ITRForm:TaxPayAfterCreditUs115JD>0</ITRForm:TaxPayAfterCreditUs115JD>"
    Else
        writeXML " <ITRForm:TaxPayAfterCreditUs115JD>" & Sheet7.Range("sheet9.TaxPayAfterCreditUs115JD").value & "</ITRForm:TaxPayAfterCreditUs115JD>"
    End If

    writeXML "<ITRForm:TaxRelief>"

    If Sheet7.Range("Sheet9.Section89").value = "" Then
        writeXML " <ITRForm:Section89>0</ITRForm:Section89>"
    Else
        writeXML " <ITRForm:Section89>" & Sheet7.Range("Sheet9.Section89").value & "</ITRForm:Section89>"
    End If

    If Sheet7.Range("Sheet9.Section90").value = "" Then
        writeXML " <ITRForm:Section90>0</ITRForm:Section90>"
    Else
        writeXML " <ITRForm:Section90>" & Sheet7.Range("Sheet9.Section90").value & "</ITRForm:Section90>"
    End If

    If Sheet7.Range("Sheet9.Section91").value = "" Then
        writeXML " <ITRForm:Section91>0</ITRForm:Section91>"
    Else
        writeXML " <ITRForm:Section91>" & Sheet7.Range("Sheet9.Section91").value & "</ITRForm:Section91>"
    End If

    If Sheet7.Range("Sheet9.TotTaxRelief").value = "" Then
        writeXML " <ITRForm:TotTaxRelief>0</ITRForm:TotTaxRelief>"
    Else
        writeXML " <ITRForm:TotTaxRelief>" & Sheet7.Range("Sheet9.TotTaxRelief").value & "</ITRForm:TotTaxRelief>"
    End If

       writeXML "</ITRForm:TaxRelief>"

        If Sheet7.Range("Sheet9.NetTaxLiability").value = "" Then
        writeXML " <ITRForm:NetTaxLiability>0</ITRForm:NetTaxLiability>"
    Else
        writeXML " <ITRForm:NetTaxLiability>" & Sheet7.Range("Sheet9.NetTaxLiability").value & "</ITRForm:NetTaxLiability>"
    End If

     writeXML "<ITRForm:IntrstPay>"

   If Sheet7.Range("Sheet9.IntrstPayUs234A").value = "" Then
        writeXML " <ITRForm:IntrstPayUs234A>0</ITRForm:IntrstPayUs234A>"
    Else
        writeXML " <ITRForm:IntrstPayUs234A>" & Sheet7.Range("Sheet9.IntrstPayUs234A").value & "</ITRForm:IntrstPayUs234A>"
    End If

    If Sheet7.Range("Sheet9.IntrstPayUs234B").value = "" Then
        writeXML " <ITRForm:IntrstPayUs234B>0</ITRForm:IntrstPayUs234B>"
    Else
        writeXML " <ITRForm:IntrstPayUs234B>" & Sheet7.Range("Sheet9.IntrstPayUs234B").value & "</ITRForm:IntrstPayUs234B>"
    End If

    If Sheet7.Range("Sheet9.IntrstPayUs234C").value = "" Then
        writeXML " <ITRForm:IntrstPayUs234C>0</ITRForm:IntrstPayUs234C>"
    Else
        writeXML " <ITRForm:IntrstPayUs234C>" & Sheet7.Range("Sheet9.IntrstPayUs234C").value & "</ITRForm:IntrstPayUs234C>"
    End If

    If Sheet7.Range("Sheet9.IntrstPayUs234F").value = "" Then
        writeXML " <ITRForm:LateFilingFee234F>0</ITRForm:LateFilingFee234F>"
    Else
        writeXML " <ITRForm:LateFilingFee234F>" & Sheet7.Range("Sheet9.IntrstPayUs234F").value & "</ITRForm:LateFilingFee234F>"
    End If

    If Sheet7.Range("Sheet9.TotalIntrstPay").value = "" Then
        writeXML " <ITRForm:TotalIntrstPay>0</ITRForm:TotalIntrstPay>"
    Else
        writeXML " <ITRForm:TotalIntrstPay>" & Sheet7.Range("Sheet9.TotalIntrstPay").value & "</ITRForm:TotalIntrstPay>"
    End If

    writeXML "</ITRForm:IntrstPay>"


    If Sheet7.Range("Sheet9.AggregateTaxInterestLiability").value = "" Then
        writeXML " <ITRForm:AggregateTaxInterestLiability>0</ITRForm:AggregateTaxInterestLiability>"
    Else
        writeXML " <ITRForm:AggregateTaxInterestLiability>" & Sheet7.Range("Sheet9.AggregateTaxInterestLiability").value & "</ITRForm:AggregateTaxInterestLiability>"
    End If

    writeXML "</ITRForm:ComputationOfTaxLiability>"

writeXML "<ITRForm:TaxPaid>"

writeXML "  <ITRForm:TaxesPaid>"

    If Sheet7.Range("Sheet9.AdvanceTax").value = "" Then
        writeXML " <ITRForm:AdvanceTax>0</ITRForm:AdvanceTax>"
    Else
        writeXML " <ITRForm:AdvanceTax>" & Sheet7.Range("Sheet9.AdvanceTax").value & "</ITRForm:AdvanceTax>"
    End If

    If Sheet7.Range("Sheet9.TDS").value = "" Then
        writeXML " <ITRForm:TDS>0</ITRForm:TDS>"
    Else
        writeXML " <ITRForm:TDS>" & Sheet7.Range("Sheet9.TDS").value & "</ITRForm:TDS>"
    End If


        If Sheet7.Range("Sheet9.TCS").value = "" Then
        writeXML " <ITRForm:TCS>0</ITRForm:TCS>"
    Else
        writeXML " <ITRForm:TCS>" & Sheet7.Range("Sheet9.TCS").value & "</ITRForm:TCS>"
    End If

            If Sheet7.Range("Sheet9.SelfAssessmentTax").value = "" Then
        writeXML " <ITRForm:SelfAssessmentTax>0</ITRForm:SelfAssessmentTax>"
    Else
        writeXML " <ITRForm:SelfAssessmentTax>" & Sheet7.Range("Sheet9.SelfAssessmentTax").value & "</ITRForm:SelfAssessmentTax>"
    End If

          If Sheet7.Range("Sheet9.TotalTaxesPaid").value = "" Then
        writeXML " <ITRForm:TotalTaxesPaid>0</ITRForm:TotalTaxesPaid>"
    Else
        writeXML " <ITRForm:TotalTaxesPaid>" & Sheet7.Range("Sheet9.TotalTaxesPaid").value & "</ITRForm:TotalTaxesPaid>"
    End If

    writeXML "</ITRForm:TaxesPaid>"

              If Sheet7.Range("Sheet9.BalTaxPayable1").value = "" Then
        writeXML " <ITRForm:BalTaxPayable>0</ITRForm:BalTaxPayable>"
    Else
        writeXML " <ITRForm:BalTaxPayable>" & Sheet7.Range("Sheet9.BalTaxPayable1").value & "</ITRForm:BalTaxPayable>"
    End If

          writeXML "</ITRForm:TaxPaid>"

writeXML "<ITRForm:Refund>"

If Sheet7.Range("Sheet9.RefundDue").value = "" Then
        writeXML " <ITRForm:RefundDue>0</ITRForm:RefundDue>"
    Else
        writeXML " <ITRForm:RefundDue>" & Sheet7.Range("Sheet9.RefundDue").value & "</ITRForm:RefundDue>"
    End If

        writeXML "<ITRForm:BankAccountDtls>"
'        If IBANSELECT <> "" Then
'            writeXML "<ITRForm:BankDtlsFlag>" & (UCase(IBANSELECT)) & "</ITRForm:BankDtlsFlag>"
''        Else
'            writeXML "<ITRForm:BankDtlsFlag>Y</ITRForm:BankDtlsFlag>"
'        End If

'        If Sheet7.Range("Sheet9.MICRCode").value <> "" Then
'            writeXML "<ITRForm:PriBankDetails>"
'            If Sheet7.Range("Sheet9.MICRCode").value <> "" Then
'               writeXML "<ITRForm:IFSCCode>" & Sheet7.Range("Sheet9.MICRCode").value & "</ITRForm:IFSCCode>"
'            End If
'
'            writeXML " <ITRForm:BankName>" & BankName_9 & " </ITRForm:BankName>"
'
'            writeXML " <ITRForm:BankAccountNo>" & BankAccountNumber_9 & "</ITRForm:BankAccountNo>"
'            If Sheet7.Range("Sheet9.CashDeposited").value <> "" Then
'                writeXML "<ITRForm:CashDeposited>" & (UCase(Sheet7.Range("Sheet9.CashDeposited").value)) & "</ITRForm:CashDeposited>"
'            End If
                'writeXML "<ITRForm:BankAccountType>" & UCase(BankAccountType_9) & "</ITRForm:BankAccountType>"
'            writeXML "</ITRForm:PriBankDetails>"
'        End If





        writeXML "<ITRForm:BankDtlsFlag>" & UCase(IBANSELECT) & "</ITRForm:BankDtlsFlag>"

        If IBANSELECT = "Y" Then
        If Not IsEmpty(BankIFSC_BA) And UBound(BankIFSC_BA) > 0 Then
            For i = 1 To UBound(BankIFSC_BA)
                writeXML "<ITRForm:AddtnlBankDetails>"

                If BankIFSC_BA(i) <> "" Then
                    writeXML "               <ITRForm:IFSCCode>" & UCase(BankIFSC_BA(i)) & "</ITRForm:IFSCCode>"
                Else
                    writeXML "               <ITRForm:IFSCCode/>"
                End If

                If BankName_BA(i) <> "" Then
                    writeXML "               <ITRForm:BankName>" & BankName_BA(i) & "</ITRForm:BankName>"
                End If

                If BankAccntnum_BA(i) <> "" Then
                    writeXML "               <ITRForm:BankAccountNo>" & UCase(BankAccntnum_BA(i)) & "</ITRForm:BankAccountNo>"
                Else
                    writeXML "               <ITRForm:BankAccountNo>" & 0 & "</ITRForm:BankAccountNo>"
                End If

                If tempxml(i) = "" Then
                      tempxml(i) = "false"
                    ElseIf tempxml(i) = True Then
                      tempxml(i) = "true"
                   ElseIf tempxml(i) = False Then
                      tempxml(i) = "false"
                   End If

                If tempxml(i) <> "" Then
                writeXML "<ITRForm:UseForRefund>" & tempxml(i) & "</ITRForm:UseForRefund>"
                End If

                'If BankAccountStatus_BA(i) <> "" Or BankAccountStatus_BA(i) <> "(Select) " Then
                '    writeXML "               <ITRForm:BankAccountType>" & UCase(BankAccountStatus_BA(i)) & "</ITRForm:BankAccountType>"
                'Else
                '    writeXML "               <ITRForm:BankAccountType>" & "" & "</ITRForm:BankAccountType>"
                'End If
'                If CashDeposited_BA(i) <> "" Then
'                    writeXML "<ITRForm:CashDeposited>" & UCase(CashDeposited_BA(i)) & "</ITRForm:CashDeposited>"
'                End If
                writeXML "</ITRForm:AddtnlBankDetails>"
            Next
        End If
        End If

        If (Not IsEmpty(IBAN)) And UBound(IBAN) > 0 Then
            For i = 1 To UBound(IBAN)
                writeXML "<ITRForm:ForeignBankDetails>"
                If IBAN(i) <> "" Then
                    writeXML "<ITRForm:SWIFTCode>" & UCase(IBAN(i)) & "</ITRForm:SWIFTCode>"
                End If

               If IBANNameOfBank(i) <> "" Then
                    writeXML "<ITRForm:BankName>" & IBANNameOfBank(i) & "</ITRForm:BankName>"
                End If

                If IBANCOUNTRY(i) <> "" Then
                    writeXML "<ITRForm:CountryCode>" & UCase(IBANCOUNTRY(i)) & "</ITRForm:CountryCode>"
                End If

                If IBANACCNO(i) <> "" Then
                    writeXML "<ITRForm:IBAN>" & UCase(IBANACCNO(i)) & "</ITRForm:IBAN>"
                End If
                writeXML "</ITRForm:ForeignBankDetails>"
            Next
        End If

    'If Range("Sheet9.NumofBankAccounts").value <> 0 Then
    '    writeXML "<ITRForm:BankAccounts>" & Range("Sheet9.NumofBankAccounts").value & "</ITRForm:BankAccounts>"
    'Else
    '    writeXML "<ITRForm:BankAccounts>" & 1 & "</ITRForm:BankAccounts>"
    'End If

writeXML " </ITRForm:BankAccountDtls>"
writeXML "</ITRForm:Refund>"

        If Sheet7.Range("AOIFlag_1").value = "" Then
        writeXML " <ITRForm:AssetOutIndiaFlag>NO</ITRForm:AssetOutIndiaFlag>"
    Else
        writeXML " <ITRForm:AssetOutIndiaFlag>" & UCase(Sheet7.Range("AOIFlag_1").value) & "</ITRForm:AssetOutIndiaFlag>"
    End If


'    ctCheck = Trim(Range("CTcheck").value)
'
'    If (Mid(ctCheck, 1, 1) <> "" Or Mid(ctCheck, 1, 1) <> "(") Then
' writeXML "<ITRForm:ScheduleCTFlag>" & Mid(ctCheck, 1, 1) & "</ITRForm:ScheduleCTFlag>"
' End If

  writeXML "</ITRForm:PartB_TTI>"

End Function
Function Verification_XML()
 writeXML "    <ITRForm:Verification>"
            writeXML "          <ITRForm:Declaration>"
                 If AssesseeVerName_9 <> "" Then
                writeXML "               <ITRForm:AssesseeVerName>" & UCase(AssesseeVerName_9) & "</ITRForm:AssesseeVerName>"
                 Else
               ' writeXML "               <ITRForm:AssesseeVerName/>"
                 End If
                 If FatherName_9 <> "" Then
                writeXML "               <ITRForm:FatherName>" & UCase(FatherName_9) & "</ITRForm:FatherName>"
                 End If

                 If verPAN <> "" Then
                writeXML "               <ITRForm:AssesseeVerPAN>" & UCase(verPAN) & "</ITRForm:AssesseeVerPAN>"
                 End If
            writeXML "          </ITRForm:Declaration>"

            If verCapacity <> "" Then
            writeXML "           <ITRForm:Capacity>" & Mid(UCase(verCapacity), 1, 1) & "</ITRForm:Capacity>"
             'Else
          '  writeXML "           <ITRForm:Place/>"
             End If

'             If Place_9 <> "" Then
'            writeXML "           <ITRForm:Place>" & UCase(Place_9) & "</ITRForm:Place>"
'             'Else
'          '  writeXML "           <ITRForm:Place/>"
'             End If
'             If Date_9 <> "" Then
'            writeXML "           <ITRForm:Date>" & Date_9 & "</ITRForm:Date>"
'             'Else
'            'writeXML "           <ITRForm:Date/>"
'             End If
        writeXML "    </ITRForm:Verification>"

End Function
Function TRPXML()
 If TRPID <> "" Then
        writeXML "    <ITRForm:TaxReturnPreparer>"

             If TRPID <> "" Then
            writeXML "           <ITRForm:IdentificationNoOfTRP>" & UCase(TRPID) & "</ITRForm:IdentificationNoOfTRP>"
             Else
            'writeXML "           <ITRForm:IdentificationNoOfTRP/>"
             End If

             If TRPName <> "" Then
            writeXML "           <ITRForm:NameOfTRP>" & UCase(TRPName) & "</ITRForm:NameOfTRP>"
             Else
            'writeXML "           <ITRForm:NameOfTRP/>"
             End If

             If TRPreimbursementAmnt <> "" Then
                writeXML "           <ITRForm:ReImbFrmGov>" & UCase(TRPreimbursementAmnt) & "</ITRForm:ReImbFrmGov>"
             Else
            'writeXML "           <ITRForm:ReImbFrmGov/>" ' " & 0 & "</ITRForm:ReImbFrmGov>"
             End If
        writeXML "    </ITRForm:TaxReturnPreparer>"
         End If

End Function
'To be change
Function ScheduleSXML()
Dim count, cntr As Variant

count = Sheet39.Range("SalarySectionCount").value
subProcPerc = "Schedule S"
noOfProcessSub = count
'If SchSTotIncUnderHeadSalaries1 > 0 Then
If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then

If Sheet39.Range("SAL.IncomeFromSalary1").value > 0 Or Sheet39.Range("SAL.TotalGrossSalary").value > 0 Or Sheet39.Range("SAL.ExemptUSectionOth1").value > 0 Or Sheet39.Range("SAL.Netsalary1").value > 0 Or Sheet39.Range("SAL.Deduction1").value > 0 Then
writeXML "<ITRForm:ScheduleS>"
'  SchSValueOfPerquisites1(cntr), SchSProfitsinLieuOfSalary1(cntr), _

If UBound(SchSEmpName) > 0 Then
For cntr = 1 To count

If SchSEmpName(cntr) <> "" Then
writeXML "   <ITRForm:Salaries>"
If SchSEmpName(cntr) <> "" Then
writeXML "      <ITRForm:NameOfEmployer>" & SchSEmpName(cntr) & "</ITRForm:NameOfEmployer>"
End If

 Dim NatureOfEmployment As String
 NatureOfEmployment = UCase(Mid(Sheet39.Range("sheet.EmployerCategory" & cntr).value, 1, 3))

            If NatureOfEmployment = "CEN" Then NatureOfEmployment = "CGOV"
            If NatureOfEmployment = "STA" Then NatureOfEmployment = "SGOV"
            If NatureOfEmployment = "PEN" Then NatureOfEmployment = "PE"
            If NatureOfEmployment = "PUB" Then NatureOfEmployment = "PSU"
            If NatureOfEmployment = "OTH" Then NatureOfEmployment = "OTH"

            writeXML "<ITRForm:NatureOfEmployment>" & NatureOfEmployment & "</ITRForm:NatureOfEmployment>"

If SchSEmpPAN(cntr) <> "" Then
writeXML "      <ITRForm:TANofEmployer>" & SchSEmpPAN(cntr) & "</ITRForm:TANofEmployer>"
End If

If SchSEmpAddr(cntr) <> "" Then
writeXML "      <ITRForm:AddressDetail>"

If SchSEmpAddr(cntr) <> "" Then
writeXML "          <ITRForm:AddrDetail>" & SchSEmpAddr(cntr) & "</ITRForm:AddrDetail>"
End If

If SchSEmpTown_City(cntr) <> "" Then
writeXML "          <ITRForm:CityOrTownOrDistrict>" & SchSEmpTown_City(cntr) & "</ITRForm:CityOrTownOrDistrict>"
End If

If SchSEmpState(cntr) <> "" Then
writeXML "          <ITRForm:StateCode>" & Mid(SchSEmpState(cntr), 1, InStr(1, SchSEmpState(cntr), "-") - 1) & "</ITRForm:StateCode>"
End If

If SchSEmpPinCode(cntr) <> "" Then
writeXML "          <ITRForm:PinCode>" & SchSEmpPinCode(cntr) & "</ITRForm:PinCode>"
End If

If SchSEmpZipCode(cntr) <> "" Then
writeXML "          <ITRForm:ZipCode>" & SchSEmpZipCode(cntr) & "</ITRForm:ZipCode>"
End If

writeXML "      </ITRForm:AddressDetail>"
End If

writeXML "      <ITRForm:Salarys>"


If SchSSalary1(cntr) <> "" Then
writeXML "          <ITRForm:GrossSalary>" & SchSSalary1(cntr) & "</ITRForm:GrossSalary>"
Else
writeXML "          <ITRForm:GrossSalary>0</ITRForm:GrossSalary>"
End If

If SchSAllowancesNotExempt1(cntr) <> "" Then
writeXML "          <ITRForm:Salary>" & SchSAllowancesNotExempt1(cntr) & "</ITRForm:Salary>"
Else
writeXML "          <ITRForm:Salary>0</ITRForm:Salary>"
End If


  GrossSalary (cntr)

If SchSValueOfPerquisites1(cntr) <> "" Then
writeXML "          <ITRForm:ValueOfPerquisites>" & SchSValueOfPerquisites1(cntr) & "</ITRForm:ValueOfPerquisites>"
Else
writeXML "          <ITRForm:ValueOfPerquisites>0</ITRForm:ValueOfPerquisites>"
End If

  ValueOfPerquisites (cntr)

If SchSProfitsinLieuOfSalary1(cntr) <> "" Then
writeXML "          <ITRForm:ProfitsinLieuOfSalary>" & SchSProfitsinLieuOfSalary1(cntr) & "</ITRForm:ProfitsinLieuOfSalary>"
Else
writeXML "          <ITRForm:ProfitsinLieuOfSalary>0</ITRForm:ProfitsinLieuOfSalary>"
End If

 ProfitsinLieuOfSalary (cntr)
 writeXML "</ITRForm:Salarys>"
 writeXML "</ITRForm:Salaries>"
 End If
Next



  If Sheet39.Range("SAL.TotalGrossSalary").value <> "" Then
            writeXML "<ITRForm:TotalGrossSalary>" & Sheet39.Range("SAL.TotalGrossSalary").value & "</ITRForm:TotalGrossSalary>"
        Else
            writeXML "<ITRForm:TotalGrossSalary>0</ITRForm:TotalGrossSalary>"
        End If

        If Sheet39.Range("SAL.ExemptUSectionOth1").value <> "" Then
            writeXML "<ITRForm:AllwncExtentExemptUs10>" & Sheet39.Range("SAL.ExemptUSectionOth1").value & "</ITRForm:AllwncExtentExemptUs10>"
        Else
            writeXML "<ITRForm:AllwncExtentExemptUs10>0</ITRForm:AllwncExtentExemptUs10>"
        End If


  If Sheet39.Range("SAL.ExemptUSectionOth1").value > 0 Then
            writeXML "<ITRForm:AllwncExemptUs10>"

                If (Not IsEmpty(Others_NOI_S)) And UBound(Others_NOI_S) > 0 Then
                    For i = 1 To UBound(Others_NOI_S)
                        writeXML "<ITRForm:AllwncExemptUs10Dtls>"

                        If (Trim(Others_NOI_S(i))) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
                            writeXML "<ITRForm:SalNatureDesc>10(10B)(i)</ITRForm:SalNatureDesc>"

                        ElseIf (Trim(Others_NOI_S(i))) = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
                            writeXML "<ITRForm:SalNatureDesc>10(10B)(ii)</ITRForm:SalNatureDesc>"

                        ElseIf UCase(Trim(Others_NOI_S(i))) = "ANY OTHER" Then
                            writeXML "<ITRForm:SalNatureDesc>OTH</ITRForm:SalNatureDesc>"
                        Else
                            writeXML "<ITRForm:SalNatureDesc>" & Mid(Others_NOI_S(i), 5, InStr(Others_NOI_S(i), "-") - 5) & "</ITRForm:SalNatureDesc>"
                        End If

                        If Others_NOI1_s(i) <> "" Then
                            writeXML "<ITRForm:SalOthNatOfInc>" & Others_NOI1_s(i) & "</ITRForm:SalOthNatOfInc>"
                        End If

                        If Others_Amt_s(i) <> "" Then
                            writeXML "<ITRForm:SalOthAmount>" & UCase(Others_Amt_s(i)) & "</ITRForm:SalOthAmount>"
                        End If

                        writeXML "</ITRForm:AllwncExemptUs10Dtls>"
                   Next
                 End If

            writeXML "</ITRForm:AllwncExemptUs10>"
        End If




        If Sheet39.Range("SAL.Netsalary1").value <> "" Then
            writeXML "<ITRForm:NetSalary>" & Sheet39.Range("SAL.NetSalary1").value & "</ITRForm:NetSalary>"
        Else
            writeXML "<ITRForm:NetSalary>0</ITRForm:NetSalary>"
        End If

        If Sheet39.Range("SAL.Deduction1").value <> "" Then
            writeXML "<ITRForm:DeductionUS16>" & Sheet39.Range("SAL.Deduction1").value & "</ITRForm:DeductionUS16>"
        Else
            writeXML "<ITRForm:DeductionUS16>0</ITRForm:DeductionUS16>"
        End If

        If Sheet39.Range("SAL.DeductionUnder6Section1").value <> "" Then
            writeXML "<ITRForm:DeductionUnderSection16ia>" & Sheet39.Range("SAL.DeductionUnder6Section1").value & "</ITRForm:DeductionUnderSection16ia>"
        Else
            writeXML "<ITRForm:DeductionUnderSection16ia>0</ITRForm:DeductionUnderSection16ia>"
        End If

        If Sheet39.Range("SAL.Entertainment_allowance_1").value <> "" Then
            writeXML "<ITRForm:EntertainmntalwncUs16ii>" & Sheet39.Range("SAL.Entertainment_allowance_1").value & "</ITRForm:EntertainmntalwncUs16ii>"
        Else
            writeXML "<ITRForm:EntertainmntalwncUs16ii>0</ITRForm:EntertainmntalwncUs16ii>"
        End If

        If Sheet39.Range("SAL.DeductionUnderSection1").value <> "" Then
            writeXML "<ITRForm:ProfessionalTaxUs16iii>" & Sheet39.Range("SAL.DeductionUnderSection1").value & "</ITRForm:ProfessionalTaxUs16iii>"
        Else
            writeXML "<ITRForm:ProfessionalTaxUs16iii>0</ITRForm:ProfessionalTaxUs16iii>"
        End If

        If Sheet39.Range("SAL.IncomeFromSalary1").value <> "" Then
            writeXML "<ITRForm:TotIncUnderHeadSalaries>" & Sheet39.Range("SAL.IncomeFromSalary1").value & "</ITRForm:TotIncUnderHeadSalaries>"
        Else
            writeXML "<ITRForm:TotIncUnderHeadSalaries>0</ITRForm:TotIncUnderHeadSalaries>"
        End If


writeXML "</ITRForm:ScheduleS>"
End If
End If
End If
End Function
Function ConstructOtherAllowances_Salary(ByVal myindex As Long) As String
    ConstructOtherAllowances_Salary = ""
    Dim j  As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Set rangecells1 = Sheet39.Range("Salary.Nature" & myindex).Cells
    Set rangecells2 = Sheet39.Range("Salary.Description" & myindex).Cells
    Set rangecells3 = Sheet39.Range("Salary.Amount" & myindex).Cells

    setTblinfo_SalaryOthAllowindex (myindex)
    ReDim OthNature_Salary(myindex, end_SalOthNature)
    ReDim OthDescrp_Salary(myindex, end_SalOthNature)
    ReDim OthAmount_Salary(myindex, end_SalOthNature)

    For j = 1 To end_SalOthNature
        OthNature_Salary(myindex, j) = rangecells1.item(j).value
        OthDescrp_Salary(myindex, j) = rangecells2.item(j).value
        OthAmount_Salary(myindex, j) = rangecells3.item(j).value

        If Mid(OthNature_Salary(myindex, j), 1, 1) = "A" Then
        OthNature_Salary(myindex, j) = "OTH"
        ElseIf Mid(OthNature_Salary(myindex, j), 1, 1) = "D" Then
         OthNature_Salary(myindex, j) = "DMDP"
        Else
        OthNature_Salary(myindex, j) = Mid(OthNature_Salary(myindex, j), 4, WorksheetFunction.Find("-", OthNature_Salary(myindex, j)) - 4)
        End If

        writeXML "<ITRForm:OthersIncDtls>"
        writeXML "<ITRForm:NatureDesc>" & Trim(OthNature_Salary(myindex, j)) & "</ITRForm:NatureDesc>"
        If OthDescrp_Salary(myindex, j) <> "" Then
        writeXML "<ITRForm:OthNatOfInc>" & OthDescrp_Salary(myindex, j) & "</ITRForm:OthNatOfInc>"
        End If
        writeXML "<ITRForm:OthAmount>" & OthAmount_Salary(myindex, j) & "</ITRForm:OthAmount>"
        writeXML "</ITRForm:OthersIncDtls>"
    Next

End Function
Sub setTblinfo_SalaryOthAllowindex(ByVal myindex As Long)
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet39.Range("Salary.Nature" & myindex).count
 Set rangecells = Sheet39.Range("Salary.Nature" & myindex).Cells

Dim scode As Variant
 For mIntCtr = 1 To mIntCells
    scode = rangecells.item(mIntCtr).value
    If isdropdownblank(scode) Then
    scode = ""
    End If

    If Not scode = "" Then
        ccount = ccount + 1
    End If
 Next
 end_SalOthNature = ccount
 End Sub

Function ScheduleHPXML()
 Dim i As Long
subProcCaption = "Schedule HP"
    writeXML "      <ITRForm:ScheduleHP>"
    If (Not IsEmpty(AddrDetail_HP)) And (UBound(AddrDetail_HP) > 0) Then
            noOfProcessSub = UBound(AddrDetail_HP)
        For i = 1 To UBound(AddrDetail_HP)
            If Len(AddrDetail_HP(i)) = 0 Then
                Exit For
            End If

            writeXML "           <ITRForm:PropertyDetails>"
            writeXML "           <ITRForm:HPSNo>" & i & "</ITRForm:HPSNo>"
            writeXML "              <ITRForm:AddressDetailWithZipCode>"

            If AddrDetail_HP(i) <> "" Then
                writeXML "                   <ITRForm:AddrDetail>" & UCase(AddrDetail_HP(i)) & "</ITRForm:AddrDetail>"
            Else
                writeXML "                   <ITRForm:AddrDetail/>"
            End If

            If CityOrTownOrDistrict_HP(i) <> "" Then
                writeXML "                   <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_HP(i)) & "</ITRForm:CityOrTownOrDistrict>"
            Else
                writeXML "                   <ITRForm:CityOrTownOrDistrict/>"
            End If

            If StateCode_HP(i) <> "" Then
                writeXML "                   <ITRForm:StateCode>" & UCase(StateCode_HP(i)) & "</ITRForm:StateCode>"
            Else
                writeXML "                   <ITRForm:StateCode/>"
            End If

            If CountryCode_HP(i) <> "" Then
                writeXML "                   <ITRForm:CountryCode>" & UCase(CountryCode_HP(i)) & "</ITRForm:CountryCode>"
            Else
                writeXML "                   <ITRForm:CountryCode/>"
            End If


            If CountryCode_HP(i) = "91" Then
                If PinCode_HP(i) <> "" Then
                    writeXML "                   <ITRForm:PinCode>" & UCase(PinCode_HP(i)) & "</ITRForm:PinCode>"
                Else
                    'writeXML "                   <ITRForm:PinCode/>"
                End If
            End If

            If CountryCode_HP(i) <> "91" Then
                If ZipCode_HP(i) <> "" Then
                    writeXML "                   <ITRForm:ZipCode>" & UCase(ZipCode_HP(i)) & "</ITRForm:ZipCode>"
                Else
                    'writeXML "                   <ITRForm:ZipCode/>"
                End If
            End If

            writeXML "              </ITRForm:AddressDetailWithZipCode>"



If OwnerProperty_HP(i) <> "" Then

 If (OwnerProperty_HP(i)) = "Self" Then
    OwnerProperty_HP(i) = "SE"
 ElseIf (OwnerProperty_HP(i)) = "Minor" Then
    OwnerProperty_HP(i) = "MI"
 ElseIf (OwnerProperty_HP(i)) = "Spouse" Then
    OwnerProperty_HP(i) = "SP"
 ElseIf (OwnerProperty_HP(i)) = "Others" Then
    OwnerProperty_HP(i) = "OT"
 End If

 writeXML "<ITRForm:PropertyOwner>" & UCase(OwnerProperty_HP(i)) & "</ITRForm:PropertyOwner>"
End If


           If OwnerPropertyDescription_HP(i) <> "" Then
                writeXML "            <ITRForm:PropertyOwnerOther>" & UCase(OwnerPropertyDescription_HP(i)) & "</ITRForm:PropertyOwnerOther>"

            End If

            If CoOwnedYN_HP(i) <> "" Then
                writeXML "            <ITRForm:PropCoOwnedFlg>" & UCase(CoOwnedYN_HP(i)) & "</ITRForm:PropCoOwnedFlg>"
            Else
                writeXML "            <ITRForm:PropCoOwnedFlg/>"
            End If

            If CoOwnedShare_HP(i) <> "" Then
                writeXML "            <ITRForm:AsseseeShareProperty>" & CoOwnedShare_HP(i) & "</ITRForm:AsseseeShareProperty>"
            Else
                writeXML "            <ITRForm:AsseseeShareProperty>100</ITRForm:AsseseeShareProperty>"
            End If

            If UCase(CoOwnedYN_HP(i)) = "YES" Then
            '----------Co-Owner Part----------
            XMLPART_CoName_HP (i)
            '---------------------------------
            End If

            If ifLetOut_HP(i) <> "" Then
                If ifLetOut_HP(i) = "L" Then
                writeXML "               <ITRForm:ifLetOut>Y</ITRForm:ifLetOut>"
                End If
                If ifLetOut_HP(i) = "S" Then
                writeXML "               <ITRForm:ifLetOut>N</ITRForm:ifLetOut>"
                End If
                If ifLetOut_HP(i) = "D" Then
                writeXML "               <ITRForm:ifLetOut>D</ITRForm:ifLetOut>"
                End If


            Else
                writeXML "               <ITRForm:ifLetOut>" & UCase(DefaultifLetOut_HP) & "</ITRForm:ifLetOut>"
            End If

            If ifLetOut_HP(i) = "S" Then
            Else
                XMLPART_Tenant (i)
            End If

            '
            'If NameofTenant_HP(i) <> "" Then
           '     writeXML "               <ITRForm:NameofTenant>" & UCase(NameofTenant_HP(i)) & "</ITRForm:NameofTenant>"
           ' End If
           '
           ' If PANofTenant_HP(i) <> "" Then
           '     writeXML "               <ITRForm:PANofTenant>" & UCase(PANofTenant_HP(i)) & "</ITRForm:PANofTenant>"
           ' End If

            writeXML "<ITRForm:Rentdetails>"

            If AnnualLetableValue_HP(i) <> "" Then
                writeXML "                   <ITRForm:AnnualLetableValue>" & UCase(AnnualLetableValue_HP(i)) & "</ITRForm:AnnualLetableValue>"
            Else
                writeXML "                   <ITRForm:AnnualLetableValue>" & UCase(DefaultAnnualLetableValue_HP) & "</ITRForm:AnnualLetableValue>"
            End If

            If RentNotRealized_HP(i) <> "" Then
                writeXML "                   <ITRForm:RentNotRealized>" & UCase(RentNotRealized_HP(i)) & "</ITRForm:RentNotRealized>"
            Else
                writeXML "                   <ITRForm:RentNotRealized>" & UCase(DefaultRentNotRealized_HP) & "</ITRForm:RentNotRealized>"
            End If

            If LocalTaxes_HP(i) <> "" Then
                writeXML "                   <ITRForm:LocalTaxes>" & UCase(LocalTaxes_HP(i)) & "</ITRForm:LocalTaxes>"
            Else
                writeXML "                   <ITRForm:LocalTaxes>" & UCase(DefaultLocalTaxes_HP) & "</ITRForm:LocalTaxes>"
            End If

            If TotalUnrealizedAndTax_HP(i) <> "" Then
                writeXML "                   <ITRForm:TotalUnrealizedAndTax>" & UCase(TotalUnrealizedAndTax_HP(i)) & "</ITRForm:TotalUnrealizedAndTax>"
            Else
                writeXML "                   <ITRForm:TotalUnrealizedAndTax>" & UCase(DefaultTotalUnrealizedAndTax_HP) & "</ITRForm:TotalUnrealizedAndTax>"
            End If

            If BalanceALV_HP(i) <> "" Then
                writeXML "                   <ITRForm:BalanceALV>" & UCase(BalanceALV_HP(i)) & "</ITRForm:BalanceALV>"
            Else
                writeXML "                   <ITRForm:BalanceALV>" & UCase(DefaultBalanceALV_HP) & "</ITRForm:BalanceALV>"
            End If

            If IncomeOfHPInOwnHand_HP(i) <> "" Then
                writeXML "                   <ITRForm:AnnualOfPropOwned>" & UCase(IncomeOfHPInOwnHand_HP(i)) & "</ITRForm:AnnualOfPropOwned>"
            Else
                writeXML "                   <ITRForm:AnnualOfPropOwned>0</ITRForm:AnnualOfPropOwned>"
            End If

            If ThirtyPercentOfBalance_HP(i) <> "" Then
                writeXML "                   <ITRForm:ThirtyPercentOfBalance>" & UCase(ThirtyPercentOfBalance_HP(i)) & "</ITRForm:ThirtyPercentOfBalance>"
            Else
                writeXML "                   <ITRForm:ThirtyPercentOfBalance>" & UCase(DefaultThirtyPercentOfBalance_HP) & "</ITRForm:ThirtyPercentOfBalance>"
            End If

            If IntOnBorwCap_HP(i) <> "" Then
                writeXML "                   <ITRForm:IntOnBorwCap>" & UCase(IntOnBorwCap_HP(i)) & "</ITRForm:IntOnBorwCap>"
            Else
                writeXML "                   <ITRForm:IntOnBorwCap>" & UCase(DefaultIntOnBorwCap_HP) & "</ITRForm:IntOnBorwCap>"
            End If


            If TotalDeduct_HP(i) <> "" Then
                writeXML "                   <ITRForm:TotalDeduct>" & UCase(TotalDeduct_HP(i)) & "</ITRForm:TotalDeduct>"
            Else
                writeXML "                   <ITRForm:TotalDeduct>" & UCase(DefaultTotalDeduct_HP) & "</ITRForm:TotalDeduct>"
            End If

           If Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & i).value <> "" Then
           writeXML "<ITRForm:ArrearsUnrealizedRentRcvd>" & UCase(Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & i).value) & "</ITRForm:ArrearsUnrealizedRentRcvd>"
           End If



            If IncomeOfHP_HP(i) <> "" Then
                writeXML "                   <ITRForm:IncomeOfHP>" & UCase(IncomeOfHP_HP(i)) & "</ITRForm:IncomeOfHP>"
            Else
                writeXML "                   <ITRForm:IncomeOfHP>" & UCase(DefaultIncomeOfHP_HP) & "</ITRForm:IncomeOfHP>"
            End If

            writeXML "              </ITRForm:Rentdetails>"

            UpdateProgressBar
            writeXML "</ITRForm:PropertyDetails>"
        Next
    End If





      If Sheet8.Range("HP.PassTroughIncome").value <> "" Then
        writeXML "           <ITRForm:PassThroghIncome>" & UCase(Sheet8.Range("HP.PassTroughIncome").value) & "</ITRForm:PassThroghIncome>"
    End If

    If TotalIncomeChargeableUnHP_HP <> "" Then
        writeXML "           <ITRForm:TotalIncomeChargeableUnHP>" & UCase(TotalIncomeChargeableUnHP_HP) & "</ITRForm:TotalIncomeChargeableUnHP>"
    Else
        writeXML "           <ITRForm:TotalIncomeChargeableUnHP>" & UCase(DefaultTotalIncomeChargeableUnHP_HP) & "</ITRForm:TotalIncomeChargeableUnHP>"
    End If
    writeXML "      </ITRForm:ScheduleHP>"
End Function

Function ScheduleCYLA()
Dim TotHPlossCurYr_15 As Variant
    Dim TotBusLoss_15 As Variant
    Dim TotOthSrcLossNoRaceHorse_15 As Variant

    writeXML " <ITRForm:ScheduleCYLA>"

    writeXML " <ITRForm:Salary>"
    writeXML "  <ITRForm:IncCYLA>"

    If Sheet16.Range("salary.IncOfCurYrUnderThatHead1").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("salary.IncOfCurYrUnderThatHead1").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("salary.HPlossCurYrSetoff1").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("salary.HPlossCurYrSetoff1").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("salary.IncOfCurYrAfterSetOff1").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("salary.IncOfCurYrAfterSetOff1").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If

    writeXML "      </ITRForm:IncCYLA>"
    writeXML "      </ITRForm:Salary>"




    writeXML " <ITRForm:HP>"
    writeXML "  <ITRForm:IncCYLA>"
    If Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("hp.BusLossSetoff2").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("hp.BusLossSetoff2").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("hp.IncOfCurYrAfterSetOff2").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("hp.IncOfCurYrAfterSetOff2").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If

    writeXML "      </ITRForm:IncCYLA>"
    writeXML "      </ITRForm:HP>"




    writeXML " <ITRForm:BusProfExclSpecProf>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("busprof.HPlossCurYrSetoff").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("busprof.HPlossCurYrSetoff").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If


    If Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("busprof.IncOfCurYrAfterSetOff").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("busprof.IncOfCurYrAfterSetOff").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:BusProfExclSpecProf>"

    writeXML " <ITRForm:SpeculativeInc>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If


    If Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("busprofspec.IncOfCurYrAfterSetOff0a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("busprofspec.IncOfCurYrAfterSetOff0a").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:SpeculativeInc>"

    writeXML " <ITRForm:SpecifiedInc>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("busprofspecified.IncOfCurYrAfterSetOff0b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("busprofspecified.IncOfCurYrAfterSetOff0b").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:SpecifiedInc>"
    writeXML " <ITRForm:STCG15Per>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("stcg.HPlossCurYrSetoff1").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("stcg.HPlossCurYrSetoff1").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("stcg.BusLossSetoff1").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("stcg.BusLossSetoff1").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOff1").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOff1").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:STCG15Per>"

    writeXML " <ITRForm:STCG30Per>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("stcg.HPlossCurYrSetoff1a").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("stcg.BusLossSetoff1a").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("stcg.BusLossSetoff1a").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOff1a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOff1a").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:STCG30Per>"



    writeXML " <ITRForm:STCGAppRate>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("stcg.HPlossCurYrSetoff1b").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("stcg.HPlossCurYrSetoff1b").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("stcg.BusLossSetoff1b").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("stcg.BusLossSetoff1b").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOff1b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOff1b").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:STCGAppRate>"


 'New

     writeXML " <ITRForm:STCGDTAARate>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("stcg.BusLossSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("stcg.BusLossSetoff_DTAARates").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOff_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOff_DTAARates").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:STCGDTAARate>"

 'end


    writeXML " <ITRForm:LTCG10Per>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("ltcg.HPlossCurYrSetoff2").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("ltcg.HPlossCurYrSetoff2").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("ltcg.BusLossSetoff2").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("ltcg.BusLossSetoff2").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("ltcg.IncOfCurYrAfterSetOff2").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("ltcg.IncOfCurYrAfterSetOff2").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If

    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:LTCG10Per>"

    writeXML " <ITRForm:LTCG20Per>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("ltcg.BusLossSetoff2a").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("ltcg.BusLossSetoff2a").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("ltcg.IncOfCurYrAfterSetOff2a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("ltcg.IncOfCurYrAfterSetOff2a").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If

    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:LTCG20Per>"


     'New

     writeXML " <ITRForm:LTCGDTAARate>"
    writeXML " <ITRForm:IncCYLA>"
    If Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("ltcg.BusLossSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("ltcg.BusLossSetoff_DTAARates").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("ltcg.IncOfCurYrAfterSetOff_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("ltcg.IncOfCurYrAfterSetOff_DTAARates").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:LTCGDTAARate>"

 'end


    writeXML " <ITRForm:OthSrcExclRaceHorse>"
    writeXML " <ITRForm:IncCYLA>"

    If Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("othSecinclnlhrs.BusLossSetoff3").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("othSecinclnlhrs.BusLossSetoff3").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("othSecinclnlhrs.IncOfCurYrAfterSetOff3").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("othSecinclnlhrs.IncOfCurYrAfterSetOff3").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:OthSrcExclRaceHorse>"

    writeXML " <ITRForm:OthSrcRaceHorse>"
    writeXML " <ITRForm:IncCYLA>"

    If Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("rh.HPlossCurYrSetoff4").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("rh.HPlossCurYrSetoff4").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("rh.BusLossSetoff4").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("rh.BusLossSetoff4").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("rh.IncOfCurYrAfterSetOff4").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("rh.IncOfCurYrAfterSetOff4").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:OthSrcRaceHorse>"



  'New
      writeXML " <ITRForm:IncOSDTAA>"
        writeXML " <ITRForm:IncCYLA>"

    If Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>0</ITRForm:IncOfCurYrUnderThatHead>"
    Else
        writeXML " <ITRForm:IncOfCurYrUnderThatHead>" & UCase(Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value) & "</ITRForm:IncOfCurYrUnderThatHead>"
    End If

    If Sheet16.Range("rh.HPlossCurYrSetoff4_DTAARates").value = "" Then
        writeXML " <ITRForm:HPlossCurYrSetoff>0</ITRForm:HPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:HPlossCurYrSetoff>" & UCase(Sheet16.Range("rh.HPlossCurYrSetoff4_DTAARates").value) & "</ITRForm:HPlossCurYrSetoff>"
    End If

    If Sheet16.Range("rh.BusLossSetoff4_DTAARates").value = "" Then
        writeXML " <ITRForm:BusLossSetoff>0</ITRForm:BusLossSetoff>"
    Else
        writeXML " <ITRForm:BusLossSetoff>" & UCase(Sheet16.Range("rh.BusLossSetoff4_DTAARates").value) & "</ITRForm:BusLossSetoff>"
    End If

    If Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value = "" Then
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>0</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:OthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value) & "</ITRForm:OthSrcLossNoRaceHorseSetoff>"
    End If

    If Sheet16.Range("rh.IncOfCurYrAfterSetOff4_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>0</ITRForm:IncOfCurYrAfterSetOff>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOff>" & UCase(Sheet16.Range("rh.IncOfCurYrAfterSetOff4_DTAARates").value) & "</ITRForm:IncOfCurYrAfterSetOff>"
    End If
    writeXML " </ITRForm:IncCYLA>"
    writeXML " </ITRForm:IncOSDTAA>"

  'end

    writeXML "             <ITRForm:TotalCurYr>"

    TotHPlossCurYr_15 = Sheet16.Range("cyla.TotHPlossCurYr").value
    TotBusLoss_15 = Sheet16.Range("cyla.TotBusLoss").value
    TotOthSrcLossNoRaceHorse_15 = Sheet16.Range("cyla.TotOthSrcLossNoRaceHorse").value

    If TotHPlossCurYr_15 = "" Then
        writeXML " <ITRForm:TotHPlossCurYr>0</ITRForm:TotHPlossCurYr>"
    Else
        writeXML " <ITRForm:TotHPlossCurYr>" & UCase(TotHPlossCurYr_15) & "</ITRForm:TotHPlossCurYr>"
    End If

    If TotBusLoss_15 = "" Then
        writeXML " <ITRForm:TotBusLoss>0</ITRForm:TotBusLoss>"
    Else
        writeXML " <ITRForm:TotBusLoss>" & UCase(TotBusLoss_15) & "</ITRForm:TotBusLoss>"
    End If

    If TotOthSrcLossNoRaceHorse_15 = "" Then
        writeXML " <ITRForm:TotOthSrcLossNoRaceHorse>0</ITRForm:TotOthSrcLossNoRaceHorse>"
    Else
        writeXML " <ITRForm:TotOthSrcLossNoRaceHorse>" & UCase(TotOthSrcLossNoRaceHorse_15) & "</ITRForm:TotOthSrcLossNoRaceHorse>"
    End If

    writeXML "             </ITRForm:TotalCurYr>"

    writeXML " <ITRForm:TotalLossSetOff>"
    If Sheet16.Range("sheet16.TotHPlossCurYrSetoff").value = "" Then
        writeXML " <ITRForm:TotHPlossCurYrSetoff>0</ITRForm:TotHPlossCurYrSetoff>"
    Else
        writeXML " <ITRForm:TotHPlossCurYrSetoff>" & UCase(Sheet16.Range("sheet16.TotHPlossCurYrSetoff").value) & "</ITRForm:TotHPlossCurYrSetoff>"
    End If

    If Sheet16.Range("sheet16.TotBusLossSetoff").value = "" Then
        writeXML " <ITRForm:TotBusLossSetoff>0</ITRForm:TotBusLossSetoff>"
    Else
        writeXML " <ITRForm:TotBusLossSetoff>" & UCase(Sheet16.Range("sheet16.TotBusLossSetoff").value) & "</ITRForm:TotBusLossSetoff>"
    End If

    If Sheet16.Range("sheet16.TotOthSrcLossNoRaceHorseSetoff").value = "" Then
        writeXML " <ITRForm:TotOthSrcLossNoRaceHorseSetoff>0</ITRForm:TotOthSrcLossNoRaceHorseSetoff>"
    Else
        writeXML " <ITRForm:TotOthSrcLossNoRaceHorseSetoff>" & UCase(Sheet16.Range("sheet16.TotOthSrcLossNoRaceHorseSetoff").value) & "</ITRForm:TotOthSrcLossNoRaceHorseSetoff>"
    End If
    writeXML " </ITRForm:TotalLossSetOff>"

    writeXML " <ITRForm:LossRemAftSetOff>"
    If Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value = "" Then
        writeXML " <ITRForm:BalHPlossCurYrAftSetoff>0</ITRForm:BalHPlossCurYrAftSetoff>"
    Else
        writeXML " <ITRForm:BalHPlossCurYrAftSetoff>" & UCase(Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value) & "</ITRForm:BalHPlossCurYrAftSetoff>"
    End If

    If Sheet16.Range("sheet16.BalBusLossAftSetoff").value = "" Then
        writeXML " <ITRForm:BalBusLossAftSetoff>0</ITRForm:BalBusLossAftSetoff>"
    Else
        writeXML " <ITRForm:BalBusLossAftSetoff>" & UCase(Sheet16.Range("sheet16.BalBusLossAftSetoff").value) & "</ITRForm:BalBusLossAftSetoff>"
    End If

    If Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value = "" Then
        writeXML " <ITRForm:BalOthSrcLossNoRaceHorseAftSetoff>0</ITRForm:BalOthSrcLossNoRaceHorseAftSetoff>"
    Else
        writeXML " <ITRForm:BalOthSrcLossNoRaceHorseAftSetoff>" & UCase(Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value) & "</ITRForm:BalOthSrcLossNoRaceHorseAftSetoff>"
    End If
    writeXML " </ITRForm:LossRemAftSetOff>"
    writeXML " </ITRForm:ScheduleCYLA>"

End Function
Function ScheduleBFLA()
writeXML " <ITRForm:ScheduleBFLA>"

   writeXML " <ITRForm:Salary>"

   writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("salary.IncOfCurYrUndHeadFromCYLA").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("salary.IncOfCurYrUndHeadFromCYLA").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("salary.IncOfCurYrAfterSetOffBFLosses").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("salary.IncOfCurYrAfterSetOffBFLosses").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"

    writeXML " </ITRForm:Salary>"


    writeXML " <ITRForm:HP>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("hp.IncOfCurYrUndHeadFromCYLA1").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("hp.IncOfCurYrUndHeadFromCYLA1").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("hp.BFAllUs35Cl4Setoff1").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("hp.BFAllUs35Cl4Setoff1").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("hp.IncOfCurYrAfterSetOffBFLosses1").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("hp.IncOfCurYrAfterSetOffBFLosses1").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:HP>"
    writeXML " <ITRForm:BusProfExclSpecProf>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("busipofincl.IncOfCurYrUndHeadFromCYLA2").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("busipofincl.IncOfCurYrUndHeadFromCYLA2").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("busipofincl.BFlossPrevYrUndSameHeadSetoff2").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("busipofincl.BFlossPrevYrUndSameHeadSetoff2").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("busipofincl.BFAllUs35Cl4Setoff2").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("busipofincl.BFAllUs35Cl4Setoff2").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("busipofincl.IncOfCurYrAfterSetOffBFLosses2").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("busipofincl.IncOfCurYrAfterSetOffBFLosses2").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:BusProfExclSpecProf>"

    writeXML " <ITRForm:SpeculativeInc>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("busipofinclspec.IncOfCurYrAfterSetOffBFLosses2a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("busipofinclspec.IncOfCurYrAfterSetOffBFLosses2a").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:SpeculativeInc>"
    writeXML " <ITRForm:SpecifiedInc>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("busipofinclspecified.IncOfCurYrAfterSetOffBFLosses2b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("busipofinclspecified.IncOfCurYrAfterSetOffBFLosses2b").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:SpecifiedInc>"

    writeXML " <ITRForm:STCG15Per>"

    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:STCG15Per>"
    writeXML " <ITRForm:STCG30Per>"
    writeXML " <ITRForm:IncBFLA>"
    If Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3a").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:STCG30Per>"

    writeXML " <ITRForm:STCGAppRate>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:STCGAppRate>"

    'new

        writeXML " <ITRForm:STCGDTAARate>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b_DTAARates").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b_DTAARates").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b_DTAARates").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:STCGDTAARate>"

    'end


    writeXML " <ITRForm:LTCG10Per>"
    writeXML " <ITRForm:IncBFLA>"
    If Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:LTCG10Per>"

    writeXML " <ITRForm:LTCG20Per>"
    writeXML " <ITRForm:IncBFLA>"
    If Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4a").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4a").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:LTCG20Per>"


 'New

     writeXML " <ITRForm:LTCGDTAARate>"
    writeXML " <ITRForm:IncBFLA>"
    If Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("ltcg.BFAllUs35Cl4Setoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("ltcg.BFAllUs35Cl4Setoff_DTAARates").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses_DTAARates").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If

    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:LTCGDTAARate>"

 'end

    writeXML " <ITRForm:OthSrcExclRaceHorse>"
    writeXML " <ITRForm:IncBFLA>"
    If Sheet16.Range("othsrcincl.IncOfCurYrUndHeadFromCYLA5").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("othsrcincl.IncOfCurYrUndHeadFromCYLA5").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If


    If Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:OthSrcExclRaceHorse>"

    writeXML " <ITRForm:OthSrcRaceHorse>"
    writeXML " <ITRForm:IncBFLA>"

    If Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If

    If Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value = "" Then
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>0</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    Else
        writeXML " <ITRForm:BFlossPrevYrUndSameHeadSetoff>" & UCase(Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value) & "</ITRForm:BFlossPrevYrUndSameHeadSetoff>"
    End If

    If Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("rh.IncOfCurYrAfterSetOffBFLosses6").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("rh.IncOfCurYrAfterSetOffBFLosses6").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:OthSrcRaceHorse>"


    'New

        writeXML " <ITRForm:IncOSDTAA>"
    writeXML " <ITRForm:IncBFLA>"
    If Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>0</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    Else
        writeXML " <ITRForm:IncOfCurYrUndHeadFromCYLA>" & UCase(Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA_DTAARates").value) & "</ITRForm:IncOfCurYrUndHeadFromCYLA>"
    End If


    If Sheet16.Range("rh.BFUnabsorbedDeprSetoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>0</ITRForm:BFUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:BFUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("rh.BFUnabsorbedDeprSetoff_DTAARates").value) & "</ITRForm:BFUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("rh.BFAllUs35Cl4Setoff_DTAARates").value = "" Then
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>0</ITRForm:BFAllUs35Cl4Setoff>"
    Else
        writeXML " <ITRForm:BFAllUs35Cl4Setoff>" & UCase(Sheet16.Range("rh.BFAllUs35Cl4Setoff_DTAARates").value) & "</ITRForm:BFAllUs35Cl4Setoff>"
    End If

    If Sheet16.Range("rh.IncOfCurYrAfterSetOffBFLosses_DTAARates").value = "" Then
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>0</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    Else
        writeXML " <ITRForm:IncOfCurYrAfterSetOffBFLosses>" & UCase(Sheet16.Range("rh.IncOfCurYrAfterSetOffBFLosses_DTAARates").value) & "</ITRForm:IncOfCurYrAfterSetOffBFLosses>"
    End If
    writeXML " </ITRForm:IncBFLA>"
    writeXML " </ITRForm:IncOSDTAA>"

    'end

    writeXML " <ITRForm:TotalBFLossSetOff>"

    If Sheet16.Range("sheet16.TotBFLossSetoff").value = "" Then
        writeXML " <ITRForm:TotBFLossSetoff>0</ITRForm:TotBFLossSetoff>"
    Else
        writeXML " <ITRForm:TotBFLossSetoff>" & UCase(Sheet16.Range("sheet16.TotBFLossSetoff").value) & "</ITRForm:TotBFLossSetoff>"
    End If

    If Sheet16.Range("sheet16.TotUnabsorbedDeprSetoff").value = "" Then
        writeXML " <ITRForm:TotUnabsorbedDeprSetoff>0</ITRForm:TotUnabsorbedDeprSetoff>"
    Else
        writeXML " <ITRForm:TotUnabsorbedDeprSetoff>" & UCase(Sheet16.Range("sheet16.TotUnabsorbedDeprSetoff").value) & "</ITRForm:TotUnabsorbedDeprSetoff>"
    End If

    If Sheet16.Range("sheet16.TotAllUs35cl4Setoff").value = "" Then
        writeXML " <ITRForm:TotAllUs35cl4Setoff>0</ITRForm:TotAllUs35cl4Setoff>"
    Else
        writeXML " <ITRForm:TotAllUs35cl4Setoff>" & UCase(Sheet16.Range("sheet16.TotAllUs35cl4Setoff").value) & "</ITRForm:TotAllUs35cl4Setoff>"
    End If
    writeXML " </ITRForm:TotalBFLossSetOff>"

    If Sheet16.Range("sheet16.IncomeOfCurrYrAftCYLABFLA").value = "" Then
        writeXML " <ITRForm:IncomeOfCurrYrAftCYLABFLA>0</ITRForm:IncomeOfCurrYrAftCYLABFLA>"
    Else
        writeXML " <ITRForm:IncomeOfCurrYrAftCYLABFLA>" & UCase(Sheet16.Range("sheet16.IncomeOfCurrYrAftCYLABFLA").value) & "</ITRForm:IncomeOfCurrYrAftCYLABFLA>"
    End If
    writeXML " </ITRForm:ScheduleBFLA>"

End Function
Function ScheduleCFL()

    writeXML " <ITRForm:ScheduleCFL>"
    If Sheet17.Range("yr1999.BusLossSpecBus").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev9thYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"

        If Sheet17.Range("yr1999.DateOfFiling").value <> "" Then
'        writeXML " <ITRForm:DateOfFiling/>"
'        Else
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr1999.DateOfFiling").value), 7, 4) & "-" & Mid((Sheet17.Range("yr1999.DateOfFiling").value), 4, 2) & "-" & Mid((Sheet17.Range("yr1999.DateOfFiling").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("yr1999.BusLossSpecBus").value <> "" Then
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr1999.BusLossSpecBus").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        Else
           writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev9thYearFromAY>"
   End If

    If Sheet17.Range("yr2003.LossFrmSpecifiedBusCF9").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev8thYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        
        If Sheet17.Range("yr2000.DateOfFiling1").value <> "" Then
'        writeXML " <ITRForm:DateOfFiling/>"
'        Else
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2000.DateOfFiling1").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2000.DateOfFiling1").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2000.DateOfFiling1").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If


        If Sheet17.Range("yr2003.LossFrmSpecifiedBusCF9").value = "" Then
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2003.LossFrmSpecifiedBusCF9").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If



        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev8thYearFromAY>"
    End If




    If Sheet17.Range("yr2001.DateOfFiling2").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev7thYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        If Sheet17.Range("yr2001.DateOfFiling2").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2001.DateOfFiling2").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2001.DateOfFiling2").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2001.DateOfFiling2").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2012").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
            writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2012").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2012").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2012").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2001.HPLossCF2").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2001.HPLossCF2").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If


        If Sheet17.Range("yr2001.BusLossOthThanSpecLossCF2").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2001.BusLossOthThanSpecLossCF2").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

        If Sheet17.Range("yr2003.LossFrmSpecifiedBusCF8").value = "" Then
             writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2003.LossFrmSpecifiedBusCF8").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If
'        writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
'
'        writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & UVCase(LossFrmSpecBusCF3_17) & "</ITRForm:LossFrmSpecifiedBusCF>"

        If Sheet17.Range("CFL_STCG_Normal_2012").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2012").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2012").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2012").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2001.STCGLossCF2").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2001.STCGLossCF2").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2012").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2012").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2012").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2012").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2001.LTCGLossCF2").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2001.LTCGLossCF2").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

'        If OthSrcLossRaceHorseCF2_16 = "" Then
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
'        Else
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & UCase(OthSrcLossRaceHorseCF2_16) & "</ITRForm:OthSrcLossRaceHorseCF>"
'        End If
        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev7thYearFromAY>"
    End If

    If Sheet17.Range("yr2002.DateOfFiling3").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev6thYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        If Sheet17.Range("yr2002.DateOfFiling3").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2002.DateOfFiling3").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2002.DateOfFiling3").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2002.DateOfFiling3").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2013").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
            writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2013").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2013").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2013").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2002.HPLossCF3").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2002.HPLossCF3").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If

        If Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

        If Sheet17.Range("yr2003.LossFrmSpecifiedBusCF3").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2003.LossFrmSpecifiedBusCF3").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        'writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & UVCase(LossFrmSpecBusCF3_16) & "</ITRForm:LossFrmSpecifiedBusCF>"


        If Sheet17.Range("CFL_STCG_Normal_2013").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2013").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2013").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2013").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2002.STCGLossCF3").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2002.STCGLossCF3").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2013").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2013").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2013").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2013").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2002.LTCGLossCF3").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2002.LTCGLossCF3").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

'        If OthSrcLossRaceHorseCF3_16 = "" Then
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
'        Else
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & UCase(OthSrcLossRaceHorseCF3_16) & "</ITRForm:OthSrcLossRaceHorseCF>"
'        End If

        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev6thYearFromAY>"
    End If

    If Sheet17.Range("yr2003.DateOfFiling4").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev5thYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        If Sheet17.Range("yr2003.DateOfFiling4").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2003.DateOfFiling4").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2003.DateOfFiling4").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2003.DateOfFiling4").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2014").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
            writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2014").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2014").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2014").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2003.HPLossCF4").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2003.HPLossCF4").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If


        If Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

'        If LossFrmSpecBusCF4_16 = "" Then
'            writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
'        Else
'            writeXML " <ITRForm:LossFrmSpecBusCF>" & UCase(LossFrmSpecBusCF4_16) & "</ITRForm:LossFrmSpecBusCF>"
'        End If

        If Sheet17.Range("yr2003.LossFrmSpecifiedBusCF4").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2003.LossFrmSpecifiedBusCF4").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        'writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & UVCase(LossFrmSpecifiedBusCF4_16) & "</ITRForm:LossFrmSpecifiedBusCF>"

        If Sheet17.Range("CFL_STCG_Normal_2014").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2014").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2014").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2014").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2003.STCGLossCF4").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2003.STCGLossCF4").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2014").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2014").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2014").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2014").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2003.LTCGLossCF4").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2003.LTCGLossCF4").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

'        If OthSrcLossRaceHorseCF4_16 = "" Then
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
'        Else
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & UCase(OthSrcLossRaceHorseCF4_16) & "</ITRForm:OthSrcLossRaceHorseCF>"
'        End If
        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev5thYearFromAY>"
    End If

    If Sheet17.Range("yr2004.DateOfFiling5").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev4thYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"

        If Sheet17.Range("yr2004.DateOfFiling5").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2004.DateOfFiling5").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2004.DateOfFiling5").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2004.DateOfFiling5").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2015").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
            writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2015").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2015").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2015").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2004.HPLossCF5").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2004.HPLossCF5").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If

        If Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

'        If LossFrmSpecBusCF5_16 = "" Then
            writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
'        Else
'            writeXML " <ITRForm:LossFrmSpecBusCF>" & UCase(LossFrmSpecBusCF5_16) & "</ITRForm:LossFrmSpecBusCF>"
'        End If

        If Sheet17.Range("yr2004.LossFrmSpecifiedBusCF5").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2004.LossFrmSpecifiedBusCF5").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        'writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & UVCase(LossFrmSpecifiedBusCF5_16) & "</ITRForm:LossFrmSpecifiedBusCF>"

        If Sheet17.Range("CFL_STCG_Normal_2015").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2015").value & "</ITRForm:STCGLossCF>"
        End If


        If Sheet17.Range("CFL_STCG_PTI_2015").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2015").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2004.STCGLossCF5").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2004.STCGLossCF5").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2015").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2015").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2015").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2015").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2004.LTCGLossCF5").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2004.LTCGLossCF5").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

'        If Sheet17.Range("yr2004.OthSrcLossRaceHorseCF5").value = "" Then
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
'        Else
'            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("yr2004.OthSrcLossRaceHorseCF5").value & "</ITRForm:OthSrcLossRaceHorseCF>"
'        End If
        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev4thYearFromAY>"
    End If

    If Sheet17.Range("yr2005.DateOfFiling6").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev3rdYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        If Sheet17.Range("yr2005.DateOfFiling6").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2005.DateOfFiling6").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2005.DateOfFiling6").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2005.DateOfFiling6").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2016").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
            writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2016").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2016").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2016").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2005.HPLossCF6").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2005.HPLossCF6").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If

        If Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

        If Sheet17.Range("yr2005.LossFrmSpecBusCF6").value = "" Then
            writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("yr2005.LossFrmSpecBusCF6").value & "</ITRForm:LossFrmSpecBusCF>"
        End If

        If Sheet17.Range("yr2005.LossFrmSpecifiedBusCF6").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2005.LossFrmSpecifiedBusCF6").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        If Sheet17.Range("CFL_STCG_Normal_2016").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2016").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2016").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2016").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2005.STCGLossCF6").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2005.STCGLossCF6").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2016").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2016").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2016").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2016").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2005.LTCGLossCF6").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2005.LTCGLossCF6").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2005.OthSrcLossRaceHorseCF6").value = "" Then
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
        Else
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("yr2005.OthSrcLossRaceHorseCF6").value & "</ITRForm:OthSrcLossRaceHorseCF>"
        End If

        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev3rdYearFromAY>"
    End If


      '2017-18
    If Sheet17.Range("yr2006.DateOfFiling7").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrev2ndYearFromAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        If Sheet17.Range("yr2006.DateOfFiling7").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2006.DateOfFiling7").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2006.DateOfFiling7").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2006.DateOfFiling7").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2017").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
        writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2017").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2017").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2017").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2006.HPLossCF7").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2006.HPLossCF7").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If

        If Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

        If Sheet17.Range("yr2006.LossFrmSpecBusCF7").value = "" Then
            writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("yr2006.LossFrmSpecBusCF7").value & "</ITRForm:LossFrmSpecBusCF>"
        End If

        If Sheet17.Range("yr2006.LossFrmSpecifiedBusCF7").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2006.LossFrmSpecifiedBusCF7").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        If Sheet17.Range("CFL_STCG_Normal_2017").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2017").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2017").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2017").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2006.STCGLossCF7").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2006.STCGLossCF7").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2017").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2017").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2017").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2017").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2006.LTCGLossCF7").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2006.LTCGLossCF7").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2006.OthSrcLossRaceHorseCF7").value = "" Then
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
        Else
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("yr2006.OthSrcLossRaceHorseCF7").value & "</ITRForm:OthSrcLossRaceHorseCF>"
        End If

        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrev2ndYearFromAY>"
    End If



    '2018-19

    If Sheet17.Range("yr2006.DateOfFiling9").value <> "" Then
        writeXML " <ITRForm:LossCFFromPrevYrToAY>"
        writeXML " <ITRForm:CarryFwdLossDetail>"
        If Sheet17.Range("yr2006.DateOfFiling9").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2006.DateOfFiling9").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2006.DateOfFiling9").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2006.DateOfFiling9").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2018").value = "" Then
            writeXML " <ITRForm:HPLossCF>0</ITRForm:HPLossCF>"
        Else
        writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2018").value & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2018").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2018").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2006.HPLossCF9").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2006.HPLossCF9").value & "</ITRForm:TotalHPPTILossCF>"
        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If

        If Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

        If Sheet17.Range("yr2006.LossFrmSpecBusCF9").value = "" Then
            writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("yr2006.LossFrmSpecBusCF9").value & "</ITRForm:LossFrmSpecBusCF>"
        End If

        If Sheet17.Range("yr2006.LossFrmSpecifiedBusCF9").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2006.LossFrmSpecifiedBusCF9").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        If Sheet17.Range("CFL_STCG_Normal_2018").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2018").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2018").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2018").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2006.STCGLossCF9").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2006.STCGLossCF9").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2018").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2018").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2018").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2018").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2006.LTCGLossCF9").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2006.LTCGLossCF9").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2006.OthSrcLossRaceHorseCF9").value = "" Then
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
        Else
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("yr2006.OthSrcLossRaceHorseCF9").value & "</ITRForm:OthSrcLossRaceHorseCF>"
        End If

        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFFromPrevYrToAY>"
    End If


    '2019-20

    If Sheet17.Range("yr2007.DateOfFiling9").value <> "" Then
        writeXML " <ITRForm:LossCFCurrentAssmntYear>"
        writeXML " <ITRForm:CarryFwdLossDetail>"

        If Sheet17.Range("yr2007.DateOfFiling9").value <> "" Then
        writeXML "<ITRForm:DateOfFiling>" & Mid((Sheet17.Range("yr2007.DateOfFiling9").value), 7, 4) & "-" & Mid((Sheet17.Range("yr2007.DateOfFiling9").value), 4, 2) & "-" & Mid((Sheet17.Range("yr2007.DateOfFiling9").value), 1, 2) & "</ITRForm:DateOfFiling>"
        End If

        If Sheet17.Range("CFL_HP_Normal_2019").value <> "" Then
        writeXML " <ITRForm:HPLossCF>" & Sheet17.Range("CFL_HP_Normal_2019").value & "</ITRForm:HPLossCF>"
        Else
        writeXML " <ITRForm:HPLossCF>" & 0 & "</ITRForm:HPLossCF>"
        End If

        If Sheet17.Range("CFL_HP_PTI_2019").value <> "" Then
        writeXML " <ITRForm:HPPTILossCF>" & Sheet17.Range("CFL_HP_PTI_2019").value & "</ITRForm:HPPTILossCF>"
        Else
        writeXML " <ITRForm:HPPTILossCF>" & 0 & "</ITRForm:HPPTILossCF>"
        End If

        If Sheet17.Range("yr2007.HPLossCF9").value <> "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2007.HPLossCF9").value & "</ITRForm:TotalHPPTILossCF>"

        Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & 0 & "</ITRForm:TotalHPPTILossCF>"
        End If

        If Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9").value = "" Then
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
        Else
            writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9").value & "</ITRForm:BusLossOthThanSpecLossCF>"
        End If

        If Sheet17.Range("yr2007.LossFrmSpecBusCF9").value = "" Then
            writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("yr2007.LossFrmSpecBusCF9").value & "</ITRForm:LossFrmSpecBusCF>"
        End If

        If Sheet17.Range("yr2007.LossFrmSpecifiedBusCF9").value = "" Then
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
        Else
            writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2007.LossFrmSpecifiedBusCF9").value & "</ITRForm:LossFrmSpecifiedBusCF>"
        End If

        If Sheet17.Range("CFL_STCG_Normal_2019").value = "" Then
            writeXML " <ITRForm:STCGLossCF>0</ITRForm:STCGLossCF>"
        Else
            writeXML " <ITRForm:STCGLossCF>" & Sheet17.Range("CFL_STCG_Normal_2019").value & "</ITRForm:STCGLossCF>"
        End If

        If Sheet17.Range("CFL_STCG_PTI_2019").value <> "" Then
        writeXML " <ITRForm:STCGPTILossCF>" & Sheet17.Range("CFL_STCG_PTI_2019").value & "</ITRForm:STCGPTILossCF>"
        Else
        writeXML " <ITRForm:STCGPTILossCF>" & 0 & "</ITRForm:STCGPTILossCF>"
        End If

        If Sheet17.Range("yr2007.STCGLossCF9").value <> "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2007.STCGLossCF9").value & "</ITRForm:TotalSTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & 0 & "</ITRForm:TotalSTCGPTILossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_Normal_2019").value = "" Then
            writeXML " <ITRForm:LTCGLossCF>0</ITRForm:LTCGLossCF>"
        Else
            writeXML " <ITRForm:LTCGLossCF>" & Sheet17.Range("CFL_LTCG_Normal_2019").value & "</ITRForm:LTCGLossCF>"
        End If

        If Sheet17.Range("CFL_LTCG_PTI_2019").value <> "" Then
        writeXML " <ITRForm:LTCGPTILossCF>" & Sheet17.Range("CFL_LTCG_PTI_2019").value & "</ITRForm:LTCGPTILossCF>"
        Else
        writeXML " <ITRForm:LTCGPTILossCF>" & 0 & "</ITRForm:LTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2007.LTCGLossCF9").value <> "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2007.LTCGLossCF9").value & "</ITRForm:TotalLTCGPTILossCF>"
        Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & 0 & "</ITRForm:TotalLTCGPTILossCF>"
        End If

        If Sheet17.Range("yr2007.OthSrcLossRaceHorseCF9").value = "" Then
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
        Else
            writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("yr2007.OthSrcLossRaceHorseCF9").value & "</ITRForm:OthSrcLossRaceHorseCF>"
        End If

        writeXML " </ITRForm:CarryFwdLossDetail>"
        writeXML " </ITRForm:LossCFCurrentAssmntYear>"
    End If


    writeXML " <ITRForm:TotalOfBFLossesEarlierYrs>"
    writeXML " <ITRForm:LossSummaryDetail>"
    If Sheet17.Range("totofbfloss.HPLossCF8").value = "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>0</ITRForm:TotalHPPTILossCF>"
    Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("totofbfloss.HPLossCF8").value & "</ITRForm:TotalHPPTILossCF>"
    End If

    If Sheet17.Range("totofbfloss.BusLossOthThanSpecLossCF8").value = "" Then
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
    Else
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("totofbfloss.BusLossOthThanSpecLossCF8").value & "</ITRForm:BusLossOthThanSpecLossCF>"
    End If

     If Sheet17.Range("totofbfloss.LossFrmSpecBusCF8").value = "" Then
        writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("totofbfloss.LossFrmSpecBusCF8").value & "</ITRForm:LossFrmSpecBusCF>"
    End If

    If Sheet17.Range("totofbfloss.LossFrmSpecifiedBusCF8").value = "" Then
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("totofbfloss.LossFrmSpecifiedBusCF8").value & "</ITRForm:LossFrmSpecifiedBusCF>"
    End If

    If Sheet17.Range("totofbfloss.STCGLossCF8").value = "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>0</ITRForm:TotalSTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("totofbfloss.STCGLossCF8").value & "</ITRForm:TotalSTCGPTILossCF>"
    End If

    If Sheet17.Range("totofbfloss.LTCGLossCF8").value = "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>0</ITRForm:TotalLTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("totofbfloss.LTCGLossCF8").value & "</ITRForm:TotalLTCGPTILossCF>"
    End If

    If Sheet17.Range("totofbfloss.OthSrcLossRaceHorseCF8").value = "" Then
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
    Else
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("totofbfloss.OthSrcLossRaceHorseCF8").value & "</ITRForm:OthSrcLossRaceHorseCF>"
    End If

    writeXML " </ITRForm:LossSummaryDetail>"
    writeXML " </ITRForm:TotalOfBFLossesEarlierYrs>"
    writeXML " <ITRForm:AdjTotBFLossInBFLA>"
    writeXML " <ITRForm:LossSummaryDetail>"
    If Sheet17.Range("adjtotloss.HPLossCF9").value = "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>0</ITRForm:TotalHPPTILossCF>"
    Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("adjtotloss.HPLossCF9").value & "</ITRForm:TotalHPPTILossCF>"
    End If

    If Sheet17.Range("adjtotloss.BusLossOthThanSpecLossCF9").value = "" Then
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
    Else
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("adjtotloss.BusLossOthThanSpecLossCF9").value & "</ITRForm:BusLossOthThanSpecLossCF>"
    End If

    If Sheet17.Range("adjtotloss.LossFrmSpecBusCF9").value = "" Then
        writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("adjtotloss.LossFrmSpecBusCF9").value & "</ITRForm:LossFrmSpecBusCF>"
    End If

    If Sheet17.Range("adjtotloss.LossFrmSpecifiedBusCF9").value = "" Then
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("adjtotloss.LossFrmSpecifiedBusCF9").value & "</ITRForm:LossFrmSpecifiedBusCF>"
    End If

    If Sheet17.Range("adjtotloss.STCGLossCF9").value = "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>0</ITRForm:TotalSTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("adjtotloss.STCGLossCF9").value & "</ITRForm:TotalSTCGPTILossCF>"
    End If

    If Sheet17.Range("adjtotloss.LTCGLossCF9").value = "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>0</ITRForm:TotalLTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("adjtotloss.LTCGLossCF9").value & "</ITRForm:TotalLTCGPTILossCF>"
    End If

    If Sheet17.Range("adjtotloss.OthSrcLossRaceHorseCF9").value = "" Then
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
    Else
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("adjtotloss.OthSrcLossRaceHorseCF9").value & "</ITRForm:OthSrcLossRaceHorseCF>"
    End If

    writeXML " </ITRForm:LossSummaryDetail>"
    writeXML " </ITRForm:AdjTotBFLossInBFLA>"
    writeXML " <ITRForm:CurrentAYloss>"
    writeXML " <ITRForm:LossSummaryDetail>"
    If Sheet17.Range("yr2007.HPLossCF10").value = "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>0</ITRForm:TotalHPPTILossCF>"
    Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("yr2007.HPLossCF10").value & "</ITRForm:TotalHPPTILossCF>"
    End If

    If Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").value = "" Then
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
    Else
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").value & "</ITRForm:BusLossOthThanSpecLossCF>"
    End If

    If Sheet17.Range("yr2007.LossFrmSpecBusCF10").value = "" Then
        writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("yr2007.LossFrmSpecBusCF10").value & "</ITRForm:LossFrmSpecBusCF>"
    End If

    If Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").value = "" Then
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").value & "</ITRForm:LossFrmSpecifiedBusCF>"
    End If

    If Sheet17.Range("yr2007.STCGLossCF10").value = "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>0</ITRForm:TotalSTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("yr2007.STCGLossCF10").value & "</ITRForm:TotalSTCGPTILossCF>"
    End If

    If Sheet17.Range("yr2007.LTCGLossCF10").value = "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>0</ITRForm:TotalLTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("yr2007.LTCGLossCF10").value & "</ITRForm:TotalLTCGPTILossCF>"
    End If

    If Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").value = "" Then
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
    Else
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").value & "</ITRForm:OthSrcLossRaceHorseCF>"
    End If
    writeXML " </ITRForm:LossSummaryDetail>"
    writeXML " </ITRForm:CurrentAYloss>"
    writeXML " <ITRForm:TotalLossCFSummary>"
    writeXML " <ITRForm:LossSummaryDetail>"
    If Sheet17.Range("totcfloss.HPLossCF11").value = "" Then
        writeXML " <ITRForm:TotalHPPTILossCF>0</ITRForm:TotalHPPTILossCF>"
    Else
        writeXML " <ITRForm:TotalHPPTILossCF>" & Sheet17.Range("totcfloss.HPLossCF11").value & "</ITRForm:TotalHPPTILossCF>"
    End If

    If Sheet17.Range("totcfloss.BusLossOthThanSpecLossCF11").value = "" Then
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>0</ITRForm:BusLossOthThanSpecLossCF>"
    Else
        writeXML " <ITRForm:BusLossOthThanSpecLossCF>" & Sheet17.Range("totcfloss.BusLossOthThanSpecLossCF11").value & "</ITRForm:BusLossOthThanSpecLossCF>"
    End If

    If Sheet17.Range("totcfloss.LossFrmSpecBusCF11").value = "" Then
        writeXML " <ITRForm:LossFrmSpecBusCF>0</ITRForm:LossFrmSpecBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecBusCF>" & Sheet17.Range("totcfloss.LossFrmSpecBusCF11").value & "</ITRForm:LossFrmSpecBusCF>"
    End If

    If Sheet17.Range("totcfloss.LossFrmSpecifiedBusCF11").value = "" Then
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>0</ITRForm:LossFrmSpecifiedBusCF>"
    Else
        writeXML " <ITRForm:LossFrmSpecifiedBusCF>" & Sheet17.Range("totcfloss.LossFrmSpecifiedBusCF11").value & "</ITRForm:LossFrmSpecifiedBusCF>"
    End If

    If Sheet17.Range("totcfloss.STCGLossCF11").value = "" Then
        writeXML " <ITRForm:TotalSTCGPTILossCF>0</ITRForm:TotalSTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalSTCGPTILossCF>" & Sheet17.Range("totcfloss.STCGLossCF11").value & "</ITRForm:TotalSTCGPTILossCF>"
    End If

    If Sheet17.Range("totcfloss.LTCGLossCF11").value = "" Then
        writeXML " <ITRForm:TotalLTCGPTILossCF>0</ITRForm:TotalLTCGPTILossCF>"
    Else
        writeXML " <ITRForm:TotalLTCGPTILossCF>" & Sheet17.Range("totcfloss.LTCGLossCF11").value & "</ITRForm:TotalLTCGPTILossCF>"
    End If

    If Sheet17.Range("totcfloss.OthSrcLossRaceHorseCF11").value = "" Then
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>0</ITRForm:OthSrcLossRaceHorseCF>"
    Else
        writeXML " <ITRForm:OthSrcLossRaceHorseCF>" & Sheet17.Range("totcfloss.OthSrcLossRaceHorseCF11").value & "</ITRForm:OthSrcLossRaceHorseCF>"
    End If

    writeXML " </ITRForm:LossSummaryDetail>"
    writeXML " </ITRForm:TotalLossCFSummary>"
    writeXML " </ITRForm:ScheduleCFL>"
    'When doing in json also do for 2022-23
    
End Function




Function ScheduleUDXML()
Dim i As Long

    If (TotBF_UD > 0 Or TotSetoff_UD > 0 Or TotBalance_UD > 0 Or TotBF_UD2 > 0 Or _
        TotSetoff_UD2 > 0 Or TotBalance_UD2 > 0) Then
        writeXML "<ITRForm:ITR3ScheduleUD>"
        writeXML "<ITRForm:CurrAssYr>2023-23</ITRForm:CurrAssYr>"
        writeXML "<ITRForm:CurBalCFNY>" & UVCase(Balance_UDCY) & "</ITRForm:CurBalCFNY>"
        writeXML "<ITRForm:CurAllowBalCFNY>" & UVCase(Balance_UDCY2) & "</ITRForm:CurAllowBalCFNY>   "
        For i = 1 To UBound(AssYear_UD)
            writeXML "        <ITRForm:ScheduleUD>"
            If AssYear_UD(i) <> "" Then
                writeXML "            <ITRForm:AssYr>" & AssYear_UD(i) & "</ITRForm:AssYr>"
            Else
                writeXML "            <ITRForm:AssYr>0000-00</ITRForm:AssYr>"
            End If

            If BF_UD(i) <> "" Then
                writeXML "            <ITRForm:AmtBFUD>" & BF_UD(i) & "</ITRForm:AmtBFUD>"
            Else
                writeXML "            <ITRForm:AmtBFUD>0</ITRForm:AmtBFUD>"
            End If

            If Setoff_UD(i) <> "" Then
                writeXML "            <ITRForm:AmtDeprSOCY>" & Setoff_UD(i) & "</ITRForm:AmtDeprSOCY>"
            Else
                writeXML "            <ITRForm:AmtDeprSOCY>0</ITRForm:AmtDeprSOCY>"
            End If

            If Balance_UD(i) <> "" Then
                writeXML "            <ITRForm:BalCFNY>" & Balance_UD(i) & "</ITRForm:BalCFNY>"
            Else
                writeXML "            <ITRForm:BalCFNY>0</ITRForm:BalCFNY>"
            End If

            writeXML "            <ITRForm:AmtBFUAllow>" & UVCase(BF_UD2(i)) & "</ITRForm:AmtBFUAllow>"
            writeXML "            <ITRForm:AmtAllowSOCY>" & UVCase(Setoff_UD2(i)) & "</ITRForm:AmtAllowSOCY>"
            writeXML "            <ITRForm:AllowBalCFNY>" & UVCase(Balance_UD2(i)) & "</ITRForm:AllowBalCFNY>"
            writeXML "        </ITRForm:ScheduleUD>"
        Next

        writeXML "<ITRForm:TotBFUDepritAmt>" & UVCase(TotBF_UD) & "</ITRForm:TotBFUDepritAmt>"
        writeXML "<ITRForm:TotCurYrdepritSetoffInc>" & UVCase(TotSetoff_UD) & "</ITRForm:TotCurYrdepritSetoffInc>"
        writeXML "<ITRForm:TotDepritBalCFNY>" & UVCase(TotBalance_UD) & "</ITRForm:TotDepritBalCFNY>"
        writeXML "<ITRForm:TotBFUAllowAmt>" & UVCase(TotBF_UD2) & "</ITRForm:TotBFUAllowAmt>"
        writeXML "<ITRForm:TotCurYrAllowSetoffInc>" & UVCase(TotSetoff_UD2) & "</ITRForm:TotCurYrAllowSetoffInc>"
        writeXML "<ITRForm:TotalBalCFNY>" & UVCase(TotBalance_UD2) & "</ITRForm:TotalBalCFNY>"
        writeXML "    </ITRForm:ITR3ScheduleUD>"
    End If
End Function


Function ScheduleBPXML()

    Dim i As Long
    Dim Bpnsloss As Variant
    Dim bpnslossSetoff As Variant
    Dim Bpnslossleft As Variant

    Dim Bpsincome As Variant
    Dim Bpssetoff As Variant
    Dim bpsincomeaft As Variant
    Dim Bpspincome As Variant
    Dim Bpspsetoff As Variant
    Dim bpspincomeaft As Variant


    writeXML " <ITRForm:ITR3ScheduleBP>"
    writeXML " <ITRForm:BusinessIncOthThanSpec>"

    If Sheet9.Range("sheet10.ProfBfrTaxPL").value = "" Then
        writeXML " <ITRForm:ProfBfrTaxPL>0</ITRForm:ProfBfrTaxPL>"
    Else
        writeXML " <ITRForm:ProfBfrTaxPL>" & UVCase(Sheet9.Range("sheet10.ProfBfrTaxPL").value) & "</ITRForm:ProfBfrTaxPL>"
    End If

    If Sheet9.Range("sheet10.NetPLFromSpecBus").value = "" Then
        writeXML " <ITRForm:NetPLFromSpecBus>0</ITRForm:NetPLFromSpecBus>"
    Else
        writeXML " <ITRForm:NetPLFromSpecBus>" & UVCase(Sheet9.Range("sheet10.NetPLFromSpecBus").value) & "</ITRForm:NetPLFromSpecBus>"
    End If

    If Sheet9.Range("sheet10.NetPLFromSpecifiedBus").value = "" Then
        writeXML " <ITRForm:NetPLFromSpecifiedBus>0</ITRForm:NetPLFromSpecifiedBus>"
    Else
        writeXML " <ITRForm:NetPLFromSpecifiedBus>" & UVCase(Sheet9.Range("sheet10.NetPLFromSpecifiedBus").value) & "</ITRForm:NetPLFromSpecifiedBus>"
    End If

    writeXML "<ITRForm:IncRecCredPLOthHeadDtls>"
    writeXML " <ITRForm:Salary>" & UVCase(Sheet9.Range("sheet10.IncRecCredPLOthHeads1").value) & "</ITRForm:Salary>"
    writeXML " <ITRForm:HouseProperty>" & UVCase(Sheet9.Range("sheet10.IncRecCredPLOthHeads2").value) & "</ITRForm:HouseProperty>"
    writeXML " <ITRForm:CapitalGains>" & UVCase(Sheet9.Range("sheet10.IncRecCredPLOthHeads3").value) & "</ITRForm:CapitalGains>"
    writeXML " <ITRForm:OtherSources>" & UVCase(Sheet9.Range("sheet10.IncRecCredPLOthHeads4").value) & "</ITRForm:OtherSources>"
    writeXML " <ITRForm:Us115BBF>" & UVCase(Sheet9.Range("sheet10.Inc115BBF").value) & "</ITRForm:Us115BBF>"
    writeXML " <ITRForm:Us115BBG>" & UVCase(Sheet9.Range("sheet10.Inc115BBG").value) & "</ITRForm:Us115BBG>"
    writeXML "</ITRForm:IncRecCredPLOthHeadDtls>"

    If Sheet9.Range("sheet10.PLUs44sChapXIIG").value = "" Then
        writeXML " <ITRForm:PLUs44sChapXIIG>0</ITRForm:PLUs44sChapXIIG>"
    Else
        writeXML " <ITRForm:PLUs44sChapXIIG>" & UVCase(Sheet9.Range("sheet10.PLUs44sChapXIIG").value) & "</ITRForm:PLUs44sChapXIIG>"
    End If

    writeXML " <ITRForm:ProfitLossInclRefrdSec>"

    If Sheet9.Range("BP_4_44AD").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44AD>0</ITRForm:ProfitLossUs44AD>"
    Else
        writeXML " <ITRForm:ProfitLossUs44AD>" & UVCase(Sheet9.Range("BP_4_44AD").value) & "</ITRForm:ProfitLossUs44AD>"
    End If

    If Sheet9.Range("BP_4_44ADA").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44ADA>0</ITRForm:ProfitLossUs44ADA>"
    Else
        writeXML " <ITRForm:ProfitLossUs44ADA>" & UVCase(Sheet9.Range("BP_4_44ADA").value) & "</ITRForm:ProfitLossUs44ADA>"
    End If


    If Sheet9.Range("BP_4_44AE").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44AE>0</ITRForm:ProfitLossUs44AE>"
    Else
        writeXML " <ITRForm:ProfitLossUs44AE>" & UVCase(Sheet9.Range("BP_4_44AE").value) & "</ITRForm:ProfitLossUs44AE>"
    End If


    If Sheet9.Range("BP_4_44B").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44B>0</ITRForm:ProfitLossUs44B>"
    Else
        writeXML " <ITRForm:ProfitLossUs44B>" & UVCase(Sheet9.Range("BP_4_44B").value) & "</ITRForm:ProfitLossUs44B>"
    End If


    If Sheet9.Range("BP_4_44BB").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44BB>0</ITRForm:ProfitLossUs44BB>"
    Else
        writeXML " <ITRForm:ProfitLossUs44BB>" & UVCase(Sheet9.Range("BP_4_44BB").value) & "</ITRForm:ProfitLossUs44BB>"
    End If

    If Sheet9.Range("BP_4_44BBA").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44BBA>0</ITRForm:ProfitLossUs44BBA>"
    Else
        writeXML " <ITRForm:ProfitLossUs44BBA>" & UVCase(Sheet9.Range("BP_4_44BBA").value) & "</ITRForm:ProfitLossUs44BBA>"
    End If

    If Sheet9.Range("BP_4_44DA").value = "" Then
        writeXML " <ITRForm:ProfitLossUs44DA>0</ITRForm:ProfitLossUs44DA>"
    Else
        writeXML " <ITRForm:ProfitLossUs44DA>" & UVCase(Sheet9.Range("BP_4_44DA").value) & "</ITRForm:ProfitLossUs44DA>"
    End If

    writeXML " </ITRForm:ProfitLossInclRefrdSec>"


    If Sheet9.Range("BP_4b_TotalProfitFromActivates").value = "" Then
        writeXML " <ITRForm:TotalProfitFrmActCvrd>0</ITRForm:TotalProfitFrmActCvrd>"
    Else
        writeXML " <ITRForm:TotalProfitFrmActCvrd>" & UVCase(Sheet9.Range("BP_4b_TotalProfitFromActivates").value) & "</ITRForm:TotalProfitFrmActCvrd>"
    End If



    writeXML " <ITRForm:ProfitFrmActCvrd>"

    If Sheet9.Range("BP_4b_ProfitRule7").value = "" Then
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7>0</ITRForm:ProfitFrmActCvrdUndrRule7>"
    Else
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7>" & UVCase(Sheet9.Range("BP_4b_ProfitRule7").value) & "</ITRForm:ProfitFrmActCvrdUndrRule7>"
    End If

    If Sheet9.Range("BP_4b_ProfitRule7A").value = "" Then
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7A>0</ITRForm:ProfitFrmActCvrdUndrRule7A>"
    Else
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7A>" & UVCase(Sheet9.Range("BP_4b_ProfitRule7A").value) & "</ITRForm:ProfitFrmActCvrdUndrRule7A>"
    End If

     If Sheet9.Range("BP_4b_ProfitRule7B1").value = "" Then
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7B1>0</ITRForm:ProfitFrmActCvrdUndrRule7B1>"
    Else
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7B1>" & UVCase(Sheet9.Range("BP_4b_ProfitRule7B1").value) & "</ITRForm:ProfitFrmActCvrdUndrRule7B1>"
    End If



     If Sheet9.Range("BP_4b_ProfitRule7B2").value = "" Then
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7B1A>0</ITRForm:ProfitFrmActCvrdUndrRule7B1A>"
    Else
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule7B1A>" & UVCase(Sheet9.Range("BP_4b_ProfitRule7B2").value) & "</ITRForm:ProfitFrmActCvrdUndrRule7B1A>"
    End If

    If Sheet9.Range("BP_4b_ProfitRule8").value = "" Then
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule8>0</ITRForm:ProfitFrmActCvrdUndrRule8>"
    Else
        writeXML " <ITRForm:ProfitFrmActCvrdUndrRule8>" & UVCase(Sheet9.Range("BP_4b_ProfitRule8").value) & "</ITRForm:ProfitFrmActCvrdUndrRule8>"
    End If


    writeXML " </ITRForm:ProfitFrmActCvrd>"


    writeXML " <ITRForm:IncCredPL>"
    If Sheet9.Range("sheet11.FirmShareInc").value = "" Then
        writeXML " <ITRForm:FirmShareInc>0</ITRForm:FirmShareInc>"
    Else
        writeXML " <ITRForm:FirmShareInc>" & UVCase(Sheet9.Range("sheet11.FirmShareInc").value) & "</ITRForm:FirmShareInc>"
    End If

    If Sheet9.Range("sheet11.AOPBOISharInc").value = "" Then
        writeXML " <ITRForm:AOPBOISharInc>0</ITRForm:AOPBOISharInc>"
    Else
        writeXML " <ITRForm:AOPBOISharInc>" & UVCase(Sheet9.Range("sheet11.AOPBOISharInc").value) & "</ITRForm:AOPBOISharInc>"
    End If


' If Sheet9.Range("sheet11.OthExempInca").value > 0 Then
'    If Not IsEmpty(OEName) And UBound(OEName) > 0 Then
'            For i = 1 To UBound(OEName)
'                writeXML "<ITRForm:OtherExmptIncDtls>"
'
'                If OEName(i) <> "" Then
'                    writeXML "<ITRForm:OperatingRevenueName>" & OEName(i) & "</ITRForm:OperatingRevenueName>"
'                End If
'
'                If OEAmt(i) <> "" Then
'                    writeXML "<ITRForm:OperatingRevenueAmt>" & UVCase(OEAmt(i)) & "</ITRForm:OperatingRevenueAmt>"
'                End If
'
'                writeXML "</ITRForm:OtherExmptIncDtls>"
'            Next
'        End If
'    End If

      BPOE

    If Sheet9.Range("sheet11.OthExempInc").value = "" Then
        writeXML " <ITRForm:OthExempInc>0</ITRForm:OthExempInc>"
    Else
        writeXML " <ITRForm:OthExempInc>" & UVCase(Sheet9.Range("sheet11.OthExempInc").value) & "</ITRForm:OthExempInc>"
    End If

    If Sheet9.Range("sheet11.TotExempIncPL").value = "" Then
        writeXML " <ITRForm:TotExempIncPL>0</ITRForm:TotExempIncPL>"
    Else
        writeXML " <ITRForm:TotExempIncPL>" & UVCase(Sheet9.Range("sheet11.TotExempIncPL").value) & "</ITRForm:TotExempIncPL>"
    End If

    writeXML " </ITRForm:IncCredPL>"

    If Sheet9.Range("sheet11.BalancePLOthThanSpecBus").value = "" Then
        writeXML " <ITRForm:BalancePLOthThanSpecBus>0</ITRForm:BalancePLOthThanSpecBus>"
    Else
        writeXML " <ITRForm:BalancePLOthThanSpecBus>" & UVCase(Sheet9.Range("sheet11.BalancePLOthThanSpecBus").value) & "</ITRForm:BalancePLOthThanSpecBus>"
    End If

    writeXML "<ITRForm:ExpDebToPLOthHeadDtls>"

    writeXML "<ITRForm:Salary>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLOthHeads1").value) & "</ITRForm:Salary>"
    writeXML "<ITRForm:HouseProperty>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLOthHeads2").value) & "</ITRForm:HouseProperty>"
    writeXML "<ITRForm:CapitalGains>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLOthHeads3").value) & "</ITRForm:CapitalGains>"
    writeXML "<ITRForm:OtherSources>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLOthHeads4").value) & "</ITRForm:OtherSources>"
    writeXML " <ITRForm:Us115BBF>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLOth115BF").value) & "</ITRForm:Us115BBF>"
    writeXML " <ITRForm:Us115BBG>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLOth115BG").value) & "</ITRForm:Us115BBG>"
    writeXML "</ITRForm:ExpDebToPLOthHeadDtls>"


    If Sheet9.Range("sheet11.ExpDebToPLExemptInc").value = "" Then
        writeXML " <ITRForm:ExpDebToPLExemptInc>0</ITRForm:ExpDebToPLExemptInc>"
    Else
        writeXML " <ITRForm:ExpDebToPLExemptInc>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLExemptInc").value) & "</ITRForm:ExpDebToPLExemptInc>"
    End If

    If Sheet9.Range("sheet11.ExpDebToPLDisAllowed").value = "" Then
        writeXML " <ITRForm:ExpDebToPLExemptIncDisAllwUs14A>0</ITRForm:ExpDebToPLExemptIncDisAllwUs14A>"
    Else
        writeXML " <ITRForm:ExpDebToPLExemptIncDisAllwUs14A>" & UVCase(Sheet9.Range("sheet11.ExpDebToPLDisAllowed").value) & "</ITRForm:ExpDebToPLExemptIncDisAllwUs14A>"
    End If


    If Sheet9.Range("sheet11.TotExpDebPL").value = "" Then
        writeXML " <ITRForm:TotExpDebPL>0</ITRForm:TotExpDebPL>"
    Else
        writeXML " <ITRForm:TotExpDebPL>" & UVCase(Sheet9.Range("sheet11.TotExpDebPL").value) & "</ITRForm:TotExpDebPL>"
    End If

    If Sheet9.Range("sheet11.AdjustedPLOthThanSpecBus").value = "" Then
        writeXML " <ITRForm:AdjustedPLOthThanSpecBus>0</ITRForm:AdjustedPLOthThanSpecBus>"
    Else
        writeXML " <ITRForm:AdjustedPLOthThanSpecBus>" & UVCase(Sheet9.Range("sheet11.AdjustedPLOthThanSpecBus").value) & "</ITRForm:AdjustedPLOthThanSpecBus>"
    End If

    If Sheet9.Range("sheet11.DepreciationDebPLCosAct").value = "" Then
        writeXML " <ITRForm:DepreciationDebPLCosAct>0</ITRForm:DepreciationDebPLCosAct>"
    Else
        writeXML " <ITRForm:DepreciationDebPLCosAct>" & UVCase(Sheet9.Range("sheet11.DepreciationDebPLCosAct").value) & "</ITRForm:DepreciationDebPLCosAct>"
    End If

    writeXML " <ITRForm:DepreciationAllowITAct32>"
    If Sheet9.Range("sheet11.DepreciationAllowUs32_1_ii").value = "" Then
        writeXML " <ITRForm:DepreciationAllowUs32_1_ii>0</ITRForm:DepreciationAllowUs32_1_ii>"
    Else
        writeXML " <ITRForm:DepreciationAllowUs32_1_ii>" & UVCase(Sheet9.Range("sheet11.DepreciationAllowUs32_1_ii").value) & "</ITRForm:DepreciationAllowUs32_1_ii>"
    End If

    If Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").value = "" Then
        writeXML " <ITRForm:DepreciationAllowUs32_1_i>0</ITRForm:DepreciationAllowUs32_1_i>"
    Else
        writeXML " <ITRForm:DepreciationAllowUs32_1_i>" & UVCase(Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").value) & "</ITRForm:DepreciationAllowUs32_1_i>"
    End If

    If Sheet9.Range("sheet11.TotDeprAllowITAct").value = "" Then
        writeXML " <ITRForm:TotDeprAllowITAct>0</ITRForm:TotDeprAllowITAct>"
    Else
        writeXML " <ITRForm:TotDeprAllowITAct>" & UVCase(Sheet9.Range("sheet11.TotDeprAllowITAct").value) & "</ITRForm:TotDeprAllowITAct>"
    End If

    writeXML " </ITRForm:DepreciationAllowITAct32>"

    If Sheet9.Range("sheet11.AdjustPLAfterDeprOthSpecInc").value = "" Then
        writeXML " <ITRForm:AdjustPLAfterDeprOthSpecInc>0</ITRForm:AdjustPLAfterDeprOthSpecInc>"
    Else
        writeXML " <ITRForm:AdjustPLAfterDeprOthSpecInc>" & UVCase(Sheet9.Range("sheet11.AdjustPLAfterDeprOthSpecInc").value) & "</ITRForm:AdjustPLAfterDeprOthSpecInc>"
    End If

    If Sheet9.Range("sheet11.AmtDebPLDisallowUs36").value = "" Then
        writeXML " <ITRForm:AmtDebPLDisallowUs36>0</ITRForm:AmtDebPLDisallowUs36>"
    Else
        writeXML " <ITRForm:AmtDebPLDisallowUs36>" & UVCase(Sheet9.Range("sheet11.AmtDebPLDisallowUs36").value) & "</ITRForm:AmtDebPLDisallowUs36>"
    End If

    If Sheet9.Range("sheet11.AmtDebPLDisallowUs37").value = "" Then
        writeXML " <ITRForm:AmtDebPLDisallowUs37>0</ITRForm:AmtDebPLDisallowUs37>"
    Else
        writeXML " <ITRForm:AmtDebPLDisallowUs37>" & UVCase(Sheet9.Range("sheet11.AmtDebPLDisallowUs37").value) & "</ITRForm:AmtDebPLDisallowUs37>"
    End If

    If Sheet9.Range("sheet11.AmtDebPLDisallowUs40").value = "" Then
        writeXML " <ITRForm:AmtDebPLDisallowUs40>0</ITRForm:AmtDebPLDisallowUs40>"
    Else
        writeXML " <ITRForm:AmtDebPLDisallowUs40>" & UVCase(Sheet9.Range("sheet11.AmtDebPLDisallowUs40").value) & "</ITRForm:AmtDebPLDisallowUs40>"
    End If

    If Sheet9.Range("sheet11.AmtDebPLDisallowUs40A").value = "" Then
        writeXML " <ITRForm:AmtDebPLDisallowUs40A>0</ITRForm:AmtDebPLDisallowUs40A>"
    Else
        writeXML " <ITRForm:AmtDebPLDisallowUs40A>" & UVCase(Sheet9.Range("sheet11.AmtDebPLDisallowUs40A").value) & "</ITRForm:AmtDebPLDisallowUs40A>"
    End If

    If Sheet9.Range("sheet11.AmtDebPLDisallowUs43B").value = "" Then
        writeXML " <ITRForm:AmtDebPLDisallowUs43B>0</ITRForm:AmtDebPLDisallowUs43B>"
    Else
        writeXML " <ITRForm:AmtDebPLDisallowUs43B>" & UVCase(Sheet9.Range("sheet11.AmtDebPLDisallowUs43B").value) & "</ITRForm:AmtDebPLDisallowUs43B>"
    End If

    If Sheet9.Range("sheet11.InterestDisAllowUs23SMEAct").value = "" Then
        writeXML " <ITRForm:InterestDisAllowUs23SMEAct>0</ITRForm:InterestDisAllowUs23SMEAct>"
    Else
        writeXML " <ITRForm:InterestDisAllowUs23SMEAct>" & UVCase(Sheet9.Range("sheet11.InterestDisAllowUs23SMEAct").value) & "</ITRForm:InterestDisAllowUs23SMEAct>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs41").value = "" Then
        writeXML " <ITRForm:DeemIncUs41>0</ITRForm:DeemIncUs41>"
    Else
        writeXML " <ITRForm:DeemIncUs41>" & UVCase(Sheet9.Range("sheet11.DeemIncUs41").value) & "</ITRForm:DeemIncUs41>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs3380HHD80IA").value = "" Then
        writeXML " <ITRForm:DeemIncUs3380HHD80IA>0</ITRForm:DeemIncUs3380HHD80IA>"
    Else
        writeXML " <ITRForm:DeemIncUs3380HHD80IA>" & UVCase(Sheet9.Range("sheet11.DeemIncUs3380HHD80IA").value) & "</ITRForm:DeemIncUs3380HHD80IA>"
    End If


    If Sheet9.Range("sheet11.DeemIncUs32AD").value = "" Then
        writeXML " <ITRForm:DeemIncUs32AD>0</ITRForm:DeemIncUs32AD>"
    Else
        writeXML " <ITRForm:DeemIncUs32AD>" & UVCase(Sheet9.Range("sheet11.DeemIncUs32AD").value) & "</ITRForm:DeemIncUs32AD>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs33AB").value = "" Then
        writeXML " <ITRForm:DeemIncUs33AB>0</ITRForm:DeemIncUs33AB>"
    Else
        writeXML " <ITRForm:DeemIncUs33AB>" & UVCase(Sheet9.Range("sheet11.DeemIncUs33AB").value) & "</ITRForm:DeemIncUs33AB>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs33ABA").value = "" Then
        writeXML " <ITRForm:DeemIncUs33ABA>0</ITRForm:DeemIncUs33ABA>"
    Else
        writeXML " <ITRForm:DeemIncUs33ABA>" & UVCase(Sheet9.Range("sheet11.DeemIncUs33ABA").value) & "</ITRForm:DeemIncUs33ABA>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs35ABA").value = "" Then
        writeXML " <ITRForm:DeemIncUs35ABA>0</ITRForm:DeemIncUs35ABA>"
    Else
        writeXML " <ITRForm:DeemIncUs35ABA>" & UVCase(Sheet9.Range("sheet11.DeemIncUs35ABA").value) & "</ITRForm:DeemIncUs35ABA>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs35ABB").value = "" Then
        writeXML " <ITRForm:DeemIncUs35ABB>0</ITRForm:DeemIncUs35ABB>"
    Else
        writeXML " <ITRForm:DeemIncUs35ABB>" & UVCase(Sheet9.Range("sheet11.DeemIncUs35ABB").value) & "</ITRForm:DeemIncUs35ABB>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs40A").value = "" Then
        writeXML " <ITRForm:DeemIncUs40A3A>0</ITRForm:DeemIncUs40A3A>"
    Else
        writeXML " <ITRForm:DeemIncUs40A3A>" & UVCase(Sheet9.Range("sheet11.DeemIncUs40A").value) & "</ITRForm:DeemIncUs40A3A>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs72A").value = "" Then
        writeXML " <ITRForm:DeemIncUs72A>0</ITRForm:DeemIncUs72A>"
    Else
        writeXML " <ITRForm:DeemIncUs72A>" & UVCase(Sheet9.Range("sheet11.DeemIncUs72A").value) & "</ITRForm:DeemIncUs72A>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs80HHD").value = "" Then
        writeXML " <ITRForm:DeemIncUs80HHD>0</ITRForm:DeemIncUs80HHD>"
    Else
        writeXML " <ITRForm:DeemIncUs80HHD>" & UVCase(Sheet9.Range("sheet11.DeemIncUs80HHD").value) & "</ITRForm:DeemIncUs80HHD>"
    End If

    If Sheet9.Range("sheet11.DeemIncUs80IA").value = "" Then
        writeXML " <ITRForm:DeemIncUs80IA>0</ITRForm:DeemIncUs80IA>"
    Else
        writeXML " <ITRForm:DeemIncUs80IA>" & UVCase(Sheet9.Range("sheet11.DeemIncUs80IA").value) & "</ITRForm:DeemIncUs80IA>"
    End If


    writeXML " <ITRForm:DeemIncUs43CA>" & UVCase(Sheet9.Range("sheet11.DeemIncUs43CA").value) & "</ITRForm:DeemIncUs43CA>"

    If Sheet9.Range("sheet11.OthItemDisallowUs28To44DA").value = "" Then
        writeXML " <ITRForm:OthItemDisallowUs28To44DA>0</ITRForm:OthItemDisallowUs28To44DA>"
    Else
        writeXML " <ITRForm:OthItemDisallowUs28To44DA>" & UVCase(Sheet9.Range("sheet11.OthItemDisallowUs28To44DA").value) & "</ITRForm:OthItemDisallowUs28To44DA>"
    End If

    If Sheet9.Range("sheet11.AnyOthIncNotInclInExpDisallowPL").value = "" Then
        writeXML " <ITRForm:AnyOthIncNotInclInExpDisallowPL>0</ITRForm:AnyOthIncNotInclInExpDisallowPL>"
    Else
        writeXML " <ITRForm:AnyOthIncNotInclInExpDisallowPL>" & UVCase(Sheet9.Range("sheet11.AnyOthIncNotInclInExpDisallowPL").value) & "</ITRForm:AnyOthIncNotInclInExpDisallowPL>"
    End If

    If Sheet9.Range("BP.Salary").value = "" Then
        writeXML " <ITRForm:AnyOthIncNotInclInSalary>0</ITRForm:AnyOthIncNotInclInSalary>"
    Else
        writeXML " <ITRForm:AnyOthIncNotInclInSalary>" & UVCase(Sheet9.Range("BP.Salary").value) & "</ITRForm:AnyOthIncNotInclInSalary>"
    End If


    If Sheet9.Range("BP.Bonus").value = "" Then
        writeXML " <ITRForm:AnyOthIncNotInclInBonus>0</ITRForm:AnyOthIncNotInclInBonus>"
    Else
        writeXML " <ITRForm:AnyOthIncNotInclInBonus>" & UVCase(Sheet9.Range("BP.Bonus").value) & "</ITRForm:AnyOthIncNotInclInBonus>"
    End If

    If Sheet9.Range("BP.Commission").value = "" Then
        writeXML " <ITRForm:AnyOthIncNotInclInCommission>0</ITRForm:AnyOthIncNotInclInCommission>"
    Else
        writeXML " <ITRForm:AnyOthIncNotInclInCommission>" & UVCase(Sheet9.Range("BP.Commission").value) & "</ITRForm:AnyOthIncNotInclInCommission>"
    End If

    If Sheet9.Range("BP.Interest").value = "" Then
        writeXML " <ITRForm:AnyOthIncNotInclInInterest>0</ITRForm:AnyOthIncNotInclInInterest>"
    Else
        writeXML " <ITRForm:AnyOthIncNotInclInInterest>" & UVCase(Sheet9.Range("BP.Interest").value) & "</ITRForm:AnyOthIncNotInclInInterest>"
    End If

    If Sheet9.Range("BP.Others").value = "" Then
        writeXML " <ITRForm:AnyOthIncNotInclInOthers>0</ITRForm:AnyOthIncNotInclInOthers>"
    Else
        writeXML " <ITRForm:AnyOthIncNotInclInOthers>" & UVCase(Sheet9.Range("BP.Others").value) & "</ITRForm:AnyOthIncNotInclInOthers>"
    End If

    If Sheet9.Range("BP.IncDecProfit").value = "" Then
        writeXML " <ITRForm:IncProfDecLossAccICDSAdj>0</ITRForm:IncProfDecLossAccICDSAdj>"
    Else
        writeXML " <ITRForm:IncProfDecLossAccICDSAdj>" & UVCase(Sheet9.Range("BP.IncDecProfit").value) & "</ITRForm:IncProfDecLossAccICDSAdj>"
    End If

    If Sheet9.Range("sheet11.TotAfterAddToPLDeprOthSpecInc").value = "" Then
        writeXML " <ITRForm:TotAfterAddToPLDeprOthSpecInc>0</ITRForm:TotAfterAddToPLDeprOthSpecInc>"
    Else
        writeXML " <ITRForm:TotAfterAddToPLDeprOthSpecInc>" & UVCase(Sheet9.Range("sheet11.TotAfterAddToPLDeprOthSpecInc").value) & "</ITRForm:TotAfterAddToPLDeprOthSpecInc>"
    End If

    If Sheet9.Range("sheet11.DeductUs32_1_iii").value = "" Then
        writeXML " <ITRForm:DeductUs32_1_iii>0</ITRForm:DeductUs32_1_iii>"
    Else
        writeXML " <ITRForm:DeductUs32_1_iii>" & UVCase(Sheet9.Range("sheet11.DeductUs32_1_iii").value) & "</ITRForm:DeductUs32_1_iii>"
    End If

'    If Sheet9.Range("sheet11.DeductUs32AD").value = "" Then
'        writeXML " <ITRForm:DeductUs32_AD>0</ITRForm:DeductUs32_AD>"
'    Else
'        writeXML " <ITRForm:DeductUs32_AD>" & UVCase(Sheet9.Range("sheet11.DeductUs32AD").value) & "</ITRForm:DeductUs32_AD>"
'    End If

    If Sheet9.Range("sheet11.DebPLUs35ExcessAmt").value = "" Then
        writeXML " <ITRForm:DebPLUs35ExcessAmt>0</ITRForm:DebPLUs35ExcessAmt>"
    Else
        writeXML " <ITRForm:DebPLUs35ExcessAmt>" & UVCase(Sheet9.Range("sheet11.DebPLUs35ExcessAmt").value) & "</ITRForm:DebPLUs35ExcessAmt>"
    End If

    If Sheet9.Range("sheet11.AmtDisallUs40NowAllow").value = "" Then
        writeXML " <ITRForm:AmtDisallUs40NowAllow>0</ITRForm:AmtDisallUs40NowAllow>"
    Else
        writeXML " <ITRForm:AmtDisallUs40NowAllow>" & UVCase(Sheet9.Range("sheet11.AmtDisallUs40NowAllow").value) & "</ITRForm:AmtDisallUs40NowAllow>"
    End If

    If Sheet9.Range("sheet11.AmtDisallUs43BNowAllow").value = "" Then
        writeXML " <ITRForm:AmtDisallUs43BNowAllow>0</ITRForm:AmtDisallUs43BNowAllow>"
    Else
        writeXML " <ITRForm:AmtDisallUs43BNowAllow>" & UVCase(Sheet9.Range("sheet11.AmtDisallUs43BNowAllow").value) & "</ITRForm:AmtDisallUs43BNowAllow>"
    End If

'    writeXML " <ITRForm:DeductUs35AC>"
'    If Sheet9.Range("sheet11.DebPL35ACAmt").value = "" Then
'        writeXML " <ITRForm:DebPL35ACAmt>0</ITRForm:DebPL35ACAmt>"
'    Else
'        writeXML " <ITRForm:DebPL35ACAmt>" & UVCase(Sheet9.Range("sheet11.DebPL35ACAmt").value) & "</ITRForm:DebPL35ACAmt>"
'    End If
'
'    If Sheet9.Range("sheet11.AmtAllowUs35ACt").value = "" Then
'        writeXML " <ITRForm:AmtAllowUs35ACt>0</ITRForm:AmtAllowUs35ACt>"
'    Else
'        writeXML " <ITRForm:AmtAllowUs35ACt>" & UVCase(Sheet9.Range("sheet11.AmtAllowUs35ACt").value) & "</ITRForm:AmtAllowUs35ACt>"
'    End If

'    If Sheet9.Range("sheet11.ExcessAmtDeduct35AC").value = "" Then
'        writeXML " <ITRForm:ExcessAmtDeduct35AC>0</ITRForm:ExcessAmtDeduct35AC>"
'    Else
'        writeXML " <ITRForm:ExcessAmtDeduct35AC>" & UVCase(Sheet9.Range("sheet11.ExcessAmtDeduct35AC").value) & "</ITRForm:ExcessAmtDeduct35AC>"
'    End If

   ' writeXML " </ITRForm:DeductUs35AC>"

    If Sheet9.Range("sheet11.AnyOthAmtAllDeduct").value = "" Then
        writeXML " <ITRForm:AnyOthAmtAllDeduct>0</ITRForm:AnyOthAmtAllDeduct>"
    Else
        writeXML " <ITRForm:AnyOthAmtAllDeduct>" & UVCase(Sheet9.Range("sheet11.AnyOthAmtAllDeduct").value) & "</ITRForm:AnyOthAmtAllDeduct>"
    End If

    If Sheet9.Range("BP.DecPIncL").value = "" Then
        writeXML " <ITRForm:DecProfIncLossAccICDSAdj>0</ITRForm:DecProfIncLossAccICDSAdj>"
    Else
        writeXML " <ITRForm:DecProfIncLossAccICDSAdj>" & UVCase(Sheet9.Range("BP.DecPIncL").value) & "</ITRForm:DecProfIncLossAccICDSAdj>"
    End If

    If Sheet9.Range("sheet11.TotDeductionAmts").value = "" Then
        writeXML " <ITRForm:TotDeductionAmts>0</ITRForm:TotDeductionAmts>"
    Else
        writeXML " <ITRForm:TotDeductionAmts>" & UVCase(Sheet9.Range("sheet11.TotDeductionAmts").value) & "</ITRForm:TotDeductionAmts>"
    End If

    If Sheet9.Range("sheet11.PLAftAdjDedBusOthThanSpec").value = "" Then
        writeXML " <ITRForm:PLAftAdjDedBusOthThanSpec>0</ITRForm:PLAftAdjDedBusOthThanSpec>"
    Else
        writeXML " <ITRForm:PLAftAdjDedBusOthThanSpec>" & UVCase(Sheet9.Range("sheet11.PLAftAdjDedBusOthThanSpec").value) & "</ITRForm:PLAftAdjDedBusOthThanSpec>"
    End If
    writeXML " <ITRForm:DeemedProfitBusUs>"

    If Sheet9.Range("sheet11.Section44AD").value = "" Then
        writeXML " <ITRForm:Section44AD>0</ITRForm:Section44AD>"
    Else
        writeXML " <ITRForm:Section44AD>" & UVCase(Sheet9.Range("sheet11.Section44AD").value) & "</ITRForm:Section44AD>"
    End If

    If Sheet9.Range("sheet11.Section44ADA").value = "" Then
        writeXML " <ITRForm:Section44ADA>0</ITRForm:Section44ADA>"
    Else
        writeXML " <ITRForm:Section44ADA>" & UVCase(Sheet9.Range("sheet11.Section44ADA").value) & "</ITRForm:Section44ADA>"
    End If

    If Sheet9.Range("sheet11.Section44AE").value = "" Then
        writeXML " <ITRForm:Section44AE>0</ITRForm:Section44AE>"
    Else
        writeXML " <ITRForm:Section44AE>" & UVCase(Sheet9.Range("sheet11.Section44AE").value) & "</ITRForm:Section44AE>"
    End If

    If Sheet9.Range("sheet11.Section44B").value = "" Then
        writeXML " <ITRForm:Section44B>0</ITRForm:Section44B>"
    Else
        writeXML " <ITRForm:Section44B>" & UVCase(Sheet9.Range("sheet11.Section44B").value) & "</ITRForm:Section44B>"
    End If

    If Sheet9.Range("Sheet11.Section44BB").value = "" Then
        writeXML " <ITRForm:Section44BB>0</ITRForm:Section44BB>"
    Else
        writeXML " <ITRForm:Section44BB>" & UVCase(Sheet9.Range("Sheet11.Section44BB").value) & "</ITRForm:Section44BB>"
    End If

    If Sheet9.Range("sheet11.Section44BBA").value = "" Then
        writeXML " <ITRForm:Section44BBA>0</ITRForm:Section44BBA>"
    Else
        writeXML " <ITRForm:Section44BBA>" & UVCase(Sheet9.Range("sheet11.Section44BBA").value) & "</ITRForm:Section44BBA>"
    End If

'    If Sheet9.Range("sheet12.Section44BBB").value = "" Then
'        writeXML " <ITRForm:Section44BBB>0</ITRForm:Section44BBB>"
'    Else
'        writeXML " <ITRForm:Section44BBB>" & UVCase(Sheet9.Range("sheet12.Section44BBB").value) & "</ITRForm:Section44BBB>"
'    End If

'    If Sheet9.Range("sheet12.Section44D").value = "" Then
'        writeXML " <ITRForm:Section44D>0</ITRForm:Section44D>"
'    Else
'        writeXML " <ITRForm:Section44D>" & UVCase(Sheet9.Range("sheet12.Section44D").value) & "</ITRForm:Section44D>"
'    End If

    If Sheet9.Range("sheet12.Section44DA").value = "" Then
        writeXML " <ITRForm:Section44DA>0</ITRForm:Section44DA>"
    Else
        writeXML " <ITRForm:Section44DA>" & UVCase(Sheet9.Range("sheet12.Section44DA").value) & "</ITRForm:Section44DA>"
    End If



'    If Sheet9.Range("sheet12.FirstSchTAct").value = "" Then
'        writeXML " <ITRForm:FirstSchTAct>0</ITRForm:FirstSchTAct>"
'    Else
'        writeXML " <ITRForm:FirstSchTAct>" & UVCase(Sheet9.Range("sheet12.FirstSchTAct").value) & "</ITRForm:FirstSchTAct>"
'    End If

    If Sheet9.Range("sheet12.TotDeemedProfitBusUs").value = "" Then
        writeXML " <ITRForm:TotDeemedProfitBusUs>0</ITRForm:TotDeemedProfitBusUs>"
    Else
        writeXML " <ITRForm:TotDeemedProfitBusUs>" & UVCase(Sheet9.Range("sheet12.TotDeemedProfitBusUs").value) & "</ITRForm:TotDeemedProfitBusUs>"
    End If
    writeXML " </ITRForm:DeemedProfitBusUs>"

    If Sheet9.Range("sheet12.NetPLAftAdjBusOthThanSpec").value = "" Then
        writeXML " <ITRForm:NetPLAftAdjBusOthThanSpec>0</ITRForm:NetPLAftAdjBusOthThanSpec>"
    Else
        writeXML " <ITRForm:NetPLAftAdjBusOthThanSpec>" & UVCase(Sheet9.Range("sheet12.NetPLAftAdjBusOthThanSpec").value) & "</ITRForm:NetPLAftAdjBusOthThanSpec>"
    End If

'     If Sheet9.Range("sheet12.DENetPLBusOthThanSpec7A7B7C").value = "" Then
'        writeXML " <ITRForm:NetPLBusOthThanSpec7A7B7CI>0</ITRForm:NetPLBusOthThanSpec7A7B7CI>"
'    Else
'        writeXML " <ITRForm:NetPLBusOthThanSpec7A7B7CI>" & UVCase(Sheet9.Range("sheet12.DENetPLBusOthThanSpec7A7B7C").value) & "</ITRForm:NetPLBusOthThanSpec7A7B7CI>"
'    End If
'
'    If Sheet9.Range("sheet12.NetPLBusOthThanSpec7A7B7C").value = "" Then
'        writeXML " <ITRForm:NetPLBusOthThanSpec7A7B7C>0</ITRForm:NetPLBusOthThanSpec7A7B7C>"
'    Else
'        writeXML " <ITRForm:NetPLBusOthThanSpec7A7B7C>" & UVCase(Sheet9.Range("sheet12.NetPLBusOthThanSpec7A7B7C").value) & "</ITRForm:NetPLBusOthThanSpec7A7B7C>"
'    End If

    If Sheet9.Range("sheet12.NetPLBusOthThanSpec7A7B7C").value = "" Then
        writeXML " <ITRForm:NetPLBusOthThanSpec7A7B7C>0</ITRForm:NetPLBusOthThanSpec7A7B7C>"
    Else
        writeXML " <ITRForm:NetPLBusOthThanSpec7A7B7C>" & UVCase(Sheet9.Range("sheet12.NetPLBusOthThanSpec7A7B7C").value) & "</ITRForm:NetPLBusOthThanSpec7A7B7C>"
    End If




    If Sheet9.Range("BP_Chargeable_rule7").value = "" Then
        writeXML " <ITRForm:ChrgblIncUndrRule7>0</ITRForm:ChrgblIncUndrRule7>"
    Else
        writeXML " <ITRForm:ChrgblIncUndrRule7>" & UVCase(Sheet9.Range("BP_Chargeable_rule7").value) & "</ITRForm:ChrgblIncUndrRule7>"
    End If

    If Sheet9.Range("BP_Chargeable_rule7A").value = "" Then
        writeXML " <ITRForm:DeemedChrgblIncUndrRule7A>0</ITRForm:DeemedChrgblIncUndrRule7A>"
    Else
        writeXML " <ITRForm:DeemedChrgblIncUndrRule7A>" & UVCase(Sheet9.Range("BP_Chargeable_rule7A").value) & "</ITRForm:DeemedChrgblIncUndrRule7A>"
    End If

    If Sheet9.Range("BP_Chargeable_rule7B1").value = "" Then
        writeXML " <ITRForm:DeemedChrgblIncUndrRule7B1>0</ITRForm:DeemedChrgblIncUndrRule7B1>"
    Else
        writeXML " <ITRForm:DeemedChrgblIncUndrRule7B1>" & UVCase(Sheet9.Range("BP_Chargeable_rule7B1").value) & "</ITRForm:DeemedChrgblIncUndrRule7B1>"
    End If

    If Sheet9.Range("BP_Chargeable_rule7B1A").value = "" Then
        writeXML " <ITRForm:DeemedChrgblIncUndrRule7B1A>0</ITRForm:DeemedChrgblIncUndrRule7B1A>"
    Else
        writeXML " <ITRForm:DeemedChrgblIncUndrRule7B1A>" & UVCase(Sheet9.Range("BP_Chargeable_rule7B1A").value) & "</ITRForm:DeemedChrgblIncUndrRule7B1A>"
    End If


    If Sheet9.Range("BP_Chargeable_rule8").value = "" Then
        writeXML " <ITRForm:DeemedChrgblIncUndrRule8>0</ITRForm:DeemedChrgblIncUndrRule8>"
    Else
        writeXML " <ITRForm:DeemedChrgblIncUndrRule8>" & UVCase(Sheet9.Range("BP_Chargeable_rule8").value) & "</ITRForm:DeemedChrgblIncUndrRule8>"
    End If


    If Sheet9.Range("BP_Chargeable_ruleOther").value = "" Then
        writeXML " <ITRForm:IncomeOtherThanRule>0</ITRForm:IncomeOtherThanRule>"
    Else
        writeXML " <ITRForm:IncomeOtherThanRule>" & UVCase(Sheet9.Range("BP_Chargeable_ruleOther").value) & "</ITRForm:IncomeOtherThanRule>"
    End If

    If Sheet9.Range("BP_BalanceOfIncome39").value = "" Then
        writeXML " <ITRForm:BalIncDeemedFrmAgri>0</ITRForm:IncomeOtherThanRule>"
    Else
        writeXML " <ITRForm:BalIncDeemedFrmAgri>" & UVCase(Sheet9.Range("BP_BalanceOfIncome39").value) & "</ITRForm:BalIncDeemedFrmAgri>"
    End If


  writeXML " </ITRForm:BusinessIncOthThanSpec>"

    writeXML " <ITRForm:SpecBusinessInc>"

    If Sheet9.Range("sheet12.NetPLFrmSpecBus").value = "" Then
        writeXML " <ITRForm:NetPLFrmSpecBus>0</ITRForm:NetPLFrmSpecBus>"
    Else
        writeXML " <ITRForm:NetPLFrmSpecBus>" & UVCase(Sheet9.Range("sheet12.NetPLFrmSpecBus").value) & "</ITRForm:NetPLFrmSpecBus>"
    End If

    If Sheet9.Range("sheet12.AdditionUs28to44DA").value = "" Then
        writeXML " <ITRForm:AdditionUs28to44DA>0</ITRForm:AdditionUs28to44DA>"
    Else
        writeXML " <ITRForm:AdditionUs28to44DA>" & UVCase(Sheet9.Range("sheet12.AdditionUs28to44DA").value) & "</ITRForm:AdditionUs28to44DA>"
    End If

    If Sheet9.Range("sheet12.DeductUs28to44DA").value = "" Then
        writeXML " <ITRForm:DeductUs28to44DA>0</ITRForm:DeductUs28to44DA>"
    Else
        writeXML " <ITRForm:DeductUs28to44DA>" & UVCase(Sheet9.Range("sheet12.DeductUs28to44DA").value) & "</ITRForm:DeductUs28to44DA>"
    End If

    If Sheet9.Range("sheet12.AdjustedPLFrmSpecuBus").value = "" Then
        writeXML " <ITRForm:AdjustedPLFrmSpecuBus>0</ITRForm:AdjustedPLFrmSpecuBus>"
    Else
        writeXML " <ITRForm:AdjustedPLFrmSpecuBus>" & UVCase(Sheet9.Range("sheet12.AdjustedPLFrmSpecuBus").value) & "</ITRForm:AdjustedPLFrmSpecuBus>"
    End If
    writeXML " </ITRForm:SpecBusinessInc>"
    writeXML " <ITRForm:SpecifiedBusinessInc>"
    writeXML "        <ITRForm:NetPLFrmSpecifiedBus>" & UVCase(Sheet9.Range("sheet12.NetPLFrmSpecifiedBus").value) & "</ITRForm:NetPLFrmSpecifiedBus>"
    writeXML "        <ITRForm:AddSec28to44DA>" & UVCase(Sheet9.Range("sheet12.AddSec2844DA").value) & "</ITRForm:AddSec28to44DA>"
    writeXML "        <ITRForm:DedSec28to44DAOTDedSec35AD>" & UVCase(Sheet9.Range("sheet12.DedSec2844DA").value) & "</ITRForm:DedSec28to44DAOTDedSec35AD>"
    writeXML "        <ITRForm:ProfitLossSpecifiedBusiness>" & UVCase(Sheet9.Range("sheet12.ProfLossFromSpecifiedBus").value) & "</ITRForm:ProfitLossSpecifiedBusiness>"
    writeXML "        <ITRForm:DeductionUs35AD>" & UVCase(Sheet9.Range("sheet12.DeductUs35AD").value) & "</ITRForm:DeductionUs35AD>"
   ' writeXML "        <ITRForm:DedSec35AD1>" & UVCase(Sheet9.Range("sheet12.DeductUs35AD1").value) & "</ITRForm:DedSec35AD1>"
    'writeXML "        <ITRForm:DedSec35AD1A>" & UVCase(Sheet9.Range("sheet12.DeductUs35AD1A").value) & "</ITRForm:DedSec35AD1A>"
    writeXML "        <ITRForm:PLFrmSpecifiedBus>" & UVCase(Sheet9.Range("sheet12.AdjustedPLFrmSpecifiedBus").value) & "</ITRForm:PLFrmSpecifiedBus>"


        BPDrpdwn

    writeXML "    </ITRForm:SpecifiedBusinessInc>"




    If Sheet9.Range("sheet12.IncChrgUnHdProftGain").value = "" Then
        writeXML " <ITRForm:IncChrgUnHdProftGain>0</ITRForm:IncChrgUnHdProftGain>"
    Else
        writeXML " <ITRForm:IncChrgUnHdProftGain>" & UVCase(Sheet9.Range("sheet12.IncChrgUnHdProftGain").value) & "</ITRForm:IncChrgUnHdProftGain>"
    End If

    Bpnsloss = Sheet9.Range("sheet12.LossSetOffOnBusLoss").value
    Bpsincome = Sheet9.Range("sheet12.IncOfCurYrUnderThatHeada").value
    Bpssetoff = Sheet9.Range("sheet12.BusLossSetoffa").value
    bpsincomeaft = Sheet9.Range("sheet12.IncOfCurYrAfterSetOffa").value

    Bpspincome = Sheet9.Range("sheet12.IncOfCurYrUnderThatHeadb").value
    Bpspsetoff = Sheet9.Range("sheet12.BusLossSetoffb").value
    bpspincomeaft = Sheet9.Range("sheet12.IncOfCurYrAfterSetOffb").value

    bpnslossSetoff = Sheet9.Range("sheet12.TotLossSetOffOnBus").value
    Bpnslossleft = Sheet9.Range("sheet12.LossRemainSetOffOnBus").value

    writeXML "<ITRForm:BusSetoffCurrYr>"
    writeXML "<ITRForm:LossSetOffOnBusLoss>" & UVCase(Bpnsloss) & "</ITRForm:LossSetOffOnBusLoss>"
    writeXML " <ITRForm:SpeculativeInc>"
    writeXML "<ITRForm:IncOfCurYrUnderThatHead>" & UVCase(Bpsincome) & "</ITRForm:IncOfCurYrUnderThatHead>"
    writeXML "<ITRForm:BusLossSetoff>" & UVCase(Bpssetoff) & "</ITRForm:BusLossSetoff>"
    writeXML "<ITRForm:IncOfCurYrAfterSetOff>" & UVCase(bpsincomeaft) & "</ITRForm:IncOfCurYrAfterSetOff>"
    writeXML "</ITRForm:SpeculativeInc>"
    writeXML "<ITRForm:SpecifiedInc>"
    writeXML "<ITRForm:IncOfCurYrUnderThatHead>" & UVCase(Bpspincome) & "</ITRForm:IncOfCurYrUnderThatHead>"
    writeXML "<ITRForm:BusLossSetoff>" & UVCase(Bpspsetoff) & "</ITRForm:BusLossSetoff>"
    writeXML "<ITRForm:IncOfCurYrAfterSetOff>" & UVCase(bpspincomeaft) & "</ITRForm:IncOfCurYrAfterSetOff>"
    writeXML "</ITRForm:SpecifiedInc>"
    writeXML "<ITRForm:TotLossSetOffOnBus>" & UVCase(bpnslossSetoff) & "</ITRForm:TotLossSetOffOnBus>"
    writeXML "<ITRForm:LossRemainSetOffOnBus>" & UVCase(Bpnslossleft) & "</ITRForm:LossRemainSetOffOnBus>"
    writeXML "</ITRForm:BusSetoffCurrYr>"
    writeXML " </ITRForm:ITR3ScheduleBP>"

End Function
Function BPDrpdwn() As String
    BPDrpdwn = ""
    Dim i  As Long
    Dim rangecells1 As Range

    Set rangecells1 = Sheet9.Range("EI.Drpdn").Cells

    setTableinfo_EId_xml

    ReDim BPDrp(end_BP_xml)

    If end_BP_xml > 0 Then

    For i = 1 To end_BP_xml
        BPDrp(i) = rangecells1.item(i).value

    BPDrp(i) = Mid((Sheet9.Range("EI.Drpdn").item(i).value), 2, 2)
    If BPDrp(i) = "a)" Then BPDrp(i) = "a"
    If BPDrp(i) = "aa" Then BPDrp(i) = "aa"
    If BPDrp(i) = "ab" Then BPDrp(i) = "ab"
    If BPDrp(i) = "ac" Then BPDrp(i) = "ac"
    If BPDrp(i) = "ad" Then BPDrp(i) = "ad"
    If BPDrp(i) = "ae" Then BPDrp(i) = "ae"
    If BPDrp(i) = "af" Then BPDrp(i) = "af"
    If BPDrp(i) = "ag" Then BPDrp(i) = "ag"
    If BPDrp(i) = "ah" Then BPDrp(i) = "ah"
    If BPDrp(i) = "ai" Then BPDrp(i) = "ai"
    If BPDrp(i) = "aj" Then BPDrp(i) = "aj"
    If BPDrp(i) = "ak" Then BPDrp(i) = "ak"
    If BPDrp(i) = "b)" Then BPDrp(i) = "b"

    If BPDrp(i) <> "" And BPDrp(i) <> "Se" Then
     writeXML " <ITRForm:DedUs35ADSubSec5Dtls>"
     writeXML "    <ITRForm:DedUs35ADSubSec5>" & BPDrp(i) & "</ITRForm:DedUs35ADSubSec5>"
     writeXML " </ITRForm:DedUs35ADSubSec5Dtls>"
     End If

    Next
    End If

End Function
Sub setTableinfo_EId_xml()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet9.Range("EI.Drpdn").count
    Set rangecells = Sheet9.Range("EI.Drpdn").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "")) Then
                ccount = ccount + 1
            End If
    Next
 end_BP_xml = ccount
 rngname_35AD_BP_xml = "EI.Drpdn;"
End Sub
Function BPOE() As String
    BPOE = ""
    Dim i  As Long
    Dim rangecells1, rangecells11 As Range

    Set rangecells1 = Sheet9.Range("Name_1").Cells
    Set rangecells11 = Sheet9.Range("Amt_1").Cells

    setTableinfo_OE_xml

    ReDim BPOEN(end_OE_xml)
    ReDim BPOEA(end_OE_xml)

    If end_OE_xml > 0 Then

    For i = 1 To end_OE_xml
        BPOEN(i) = rangecells1.item(i).value
        BPOEA(i) = rangecells11.item(i).value




     writeXML "<ITRForm:OtherExmptIncDtls>"

                If BPOEN(i) <> "" Then
                    writeXML "<ITRForm:OperatingRevenueName>" & BPOEN(i) & "</ITRForm:OperatingRevenueName>"
                End If

                If BPOEA(i) <> "" Then
                    writeXML "<ITRForm:OperatingRevenueAmt>" & UVCase(BPOEA(i)) & "</ITRForm:OperatingRevenueAmt>"
                End If

                writeXML "</ITRForm:OtherExmptIncDtls>"

    Next
    End If



End Function
Sub setTableinfo_OE_xml()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet9.Range("Name_1").count
    Set rangecells = Sheet9.Range("Name_1").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "")) Then
                ccount = ccount + 1
            End If
    Next
 end_OE_xml = ccount
 rngname_OE_BP_xml = "Name_1;Amt_1;"
End Sub

Function ScheduleDPM()
writeXML "      <ITRForm:ScheduleDPM>"
    writeXML "          <ITRForm:PlantMachinery>"
    writeXML "              <ITRForm:Rate15>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DPM15 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM15) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM15) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DPM15 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM15) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM15) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DPM15 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM15) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM15) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DPM15 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM15) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM15) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DPM15 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM15) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM15) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DPM15 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM15) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM15) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DPM15 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM15) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM15) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DPM15 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM15) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM15) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DPM15 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM15) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM15) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    If AddlnDeprOnGT180DayAdditions_DPM15 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM15) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM15) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    End If

    If AddlnDeprDuringYearAdditions_DPM15 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM15) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM15) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    End If

    If AddlnDeprlessthan180days_DPM15 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM15) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM15) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    End If

    If TotalDepreciation_DPM15 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM15) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM15) & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DPM15 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DPM15) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DPM15 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DPM15) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DPM15 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DPM15) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DPM15 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM15) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM15) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DPM15 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM15) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM15) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DPM15 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM15) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM15) & "</ITRForm:WDVLastDay>"
    End If

    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate15>"

    writeXML "              <ITRForm:Rate30>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DPM30 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM30) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM30) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DPM30 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM30) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM30) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DPM30 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM30) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM30) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DPM30 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM30) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM30) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DPM30 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM30) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM30) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DPM30 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM30) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM30) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DPM30 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM30) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM30) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DPM30 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM30) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM30) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DPM30 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM30) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM30) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    If AddlnDeprOnGT180DayAdditions_DPM30 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM30) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM30) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    End If

    If AddlnDeprDuringYearAdditions_DPM30 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    End If

    If AddlnDeprlessthan180days_DPM30 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    End If

    If TotalDepreciation_DPM30 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM30) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM30) & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DPM30 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DPM30) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DPM30 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DPM30) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DPM30 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DPM30) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DPM30 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM30) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM30) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DPM30 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM30) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM30) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DPM30 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM30) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM30) & "</ITRForm:WDVLastDay>"
    End If
    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate30>"

    writeXML "              <ITRForm:Rate40>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DPM40 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM40) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM40) & "</ITRForm:WDVFirstDay>"
    End If

'    If WDVFirstDay2_DPM40 <> "" Then
'        writeXML "                       <ITRForm:WDVFirstDayOfPrevYr>" & UCase(WDVFirstDay2_DPM40) & "</ITRForm:WDVFirstDayOfPrevYr>"
'    Else
'        writeXML "                       <ITRForm:WDVFirstDayOfPrevYr>" & 0 & "</ITRForm:WDVFirstDayOfPrevYr>"
'    End If

    If AdditionsGrThan180Days_DPM40 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM40) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM40) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DPM40 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM40) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM40) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DPM40 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM40) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM40) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DPM40 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM40) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM40) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DPM40 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM40) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM40) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DPM40 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM40) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM40) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DPM40 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM40) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM40) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DPM40 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM40) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM40) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    If AddlnDeprOnGT180DayAdditions_DPM40 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM40) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM40) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    End If

    If AddlnDeprDuringYearAdditions_DPM40 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM40) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM40) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    End If

    If AddlnDeprlessthan180days_DPM40 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM40) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    End If

    If TotalDepreciation_DPM40 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM40) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM40) & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DPM40 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DPM40) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DPM40 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DPM40) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DPM40 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DPM40) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If


    If ExpdrOnTrforSaleAsset_DPM40 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM40) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM40) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DPM40 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM40) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM40) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DPM40 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM40) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM40) & "</ITRForm:WDVLastDay>"
    End If

    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate40>"

    writeXML "              <ITRForm:Rate45>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DPM45 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & (WDVFirstDay_DPM45) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & 0 & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DPM45 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & (AdditionsGrThan180Days_DPM45) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & 0 & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DPM45 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & (RealizationTotalPeriod_DPM45) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & 0 & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DPM45 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & (FullRateDeprAmt_DPM45) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & 0 & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DPM45 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & (AdditionsLessThan180Days_DPM45) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & 0 & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DPM45 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & (RealizationPeriodDuringYear_DPM45) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & 0 & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DPM45 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & (HalfRateDeprAmt_DPM45) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & 0 & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DPM45 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & (DepreciationAtFullRate_DPM45) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & 0 & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DPM45 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & (DepreciationAtHalfRate_DPM45) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & 0 & "</ITRForm:DepreciationAtHalfRate>"
    End If

    If AddlnDeprOnGT180DayAdditions_DPM45 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & (AddlnDeprOnGT180DayAdditions_DPM45) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & 0 & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    End If

    If AddlnDeprDuringYearAdditions_DPM45 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & (AddlnDeprDuringYearAdditions_DPM45) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & 0 & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    End If

    If AddlnDeprlessthan180days_DPM45 <> "" Then
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & (AddlnDeprlessthan180days_DPM45) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & 0 & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
    End If

    If TotalDepreciation_DPM45 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & (TotalDepreciation_DPM45) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & 0 & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DPM45 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & (Depreciation_Sec38_DPM45) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & 0 & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DPM45 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & (NetaggregateDep_DPM45) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & 0 & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DPM45 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & (PropaggregateDep_DPM45) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & 0 & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DPM45 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & (ExpdrOnTrforSaleAsset_DPM45) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & 0 & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DPM45 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & (CapGainUs50_DPM45) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & 0 & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DPM45 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & (WDVLastDay_DPM45) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & 0 & "</ITRForm:WDVLastDay>"
    End If

    writeXML "                </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate45>"

'    writeXML "              <ITRForm:Rate50>"
'    writeXML "                  <ITRForm:DepreciationDetail>"
'    If WDVFirstDay_DPM50 <> "" Then
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM50) & "</ITRForm:WDVFirstDay>"
'    Else
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM50) & "</ITRForm:WDVFirstDay>"
'    End If
'
'    If AdditionsGrThan180Days_DPM50 <> "" Then
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM50) & "</ITRForm:AdditionsGrThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM50) & "</ITRForm:AdditionsGrThan180Days>"
'    End If
'
'    If RealizationTotalPeriod_DPM50 <> "" Then
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM50) & "</ITRForm:RealizationTotalPeriod>"
'    Else
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM50) & "</ITRForm:RealizationTotalPeriod>"
'    End If
'
'    If FullRateDeprAmt_DPM50 <> "" Then
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM50) & "</ITRForm:FullRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM50) & "</ITRForm:FullRateDeprAmt>"
'    End If
'
'    If AdditionsLessThan180Days_DPM50 <> "" Then
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM50) & "</ITRForm:AdditionsLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM50) & "</ITRForm:AdditionsLessThan180Days>"
'    End If
'
'    If RealizationPeriodDuringYear_DPM50 <> "" Then
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM50) & "</ITRForm:RealizationPeriodLessThan180days>"
'    Else
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM50) & "</ITRForm:RealizationPeriodLessThan180days>"
'    End If
'
'    If HalfRateDeprAmt_DPM50 <> "" Then
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM50) & "</ITRForm:HalfRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM50) & "</ITRForm:HalfRateDeprAmt>"
'    End If
'
'    If DepreciationAtFullRate_DPM50 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM50) & "</ITRForm:DepreciationAtFullRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM50) & "</ITRForm:DepreciationAtFullRate>"
'    End If
'
'    If DepreciationAtHalfRate_DPM50 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM50) & "</ITRForm:DepreciationAtHalfRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM50) & "</ITRForm:DepreciationAtHalfRate>"
'    End If
'
'    If AddlnDeprOnGT180DayAdditions_DPM50 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM50) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM50) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    End If
'
'    If AddlnDeprDuringYearAdditions_DPM50 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM50) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM50) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    End If
'
'    If AddlnDeprlessthan180days_DPM50 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM50) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    End If
'
'    If TotalDepreciation_DPM50 <> "" Then
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM50) & "</ITRForm:TotalDepreciation>"
'    Else
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM50) & "</ITRForm:TotalDepreciation>"
'    End If
'
'    If ExpdrOnTrforSaleAsset_DPM50 <> "" Then
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM50) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    Else
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM50) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    End If
'
'    If CapGainUs50_DPM50 <> "" Then
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM50) & "</ITRForm:CapGainUs50>"
'    Else
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM50) & "</ITRForm:CapGainUs50>"
'    End If
'
'    If WDVLastDay_DPM50 <> "" Then
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM50) & "</ITRForm:WDVLastDay>"
'    Else
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM50) & "</ITRForm:WDVLastDay>"
'    End If
'
'    writeXML "                  </ITRForm:DepreciationDetail>"
'    writeXML "              </ITRForm:Rate50>"
'
'    writeXML "              <ITRForm:Rate60>"
'    writeXML "                  <ITRForm:DepreciationDetail>"
'    If WDVFirstDay_DPM60 <> "" Then
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM60) & "</ITRForm:WDVFirstDay>"
'    Else
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM60) & "</ITRForm:WDVFirstDay>"
'    End If
'
'    If AdditionsGrThan180Days_DPM60 <> "" Then
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM60) & "</ITRForm:AdditionsGrThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM60) & "</ITRForm:AdditionsGrThan180Days>"
'    End If
'
'    If RealizationTotalPeriod_DPM60 <> "" Then
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM60) & "</ITRForm:RealizationTotalPeriod>"
'    Else
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM60) & "</ITRForm:RealizationTotalPeriod>"
'    End If
'
'    If FullRateDeprAmt_DPM60 <> "" Then
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM60) & "</ITRForm:FullRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM60) & "</ITRForm:FullRateDeprAmt>"
'    End If
'
'    If AdditionsLessThan180Days_DPM60 <> "" Then
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM60) & "</ITRForm:AdditionsLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM60) & "</ITRForm:AdditionsLessThan180Days>"
'    End If
'
'    If RealizationPeriodDuringYear_DPM60 <> "" Then
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM60) & "</ITRForm:RealizationPeriodLessThan180days>"
'    Else
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM60) & "</ITRForm:RealizationPeriodLessThan180days>"
'    End If
'
'    If HalfRateDeprAmt_DPM60 <> "" Then
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM60) & "</ITRForm:HalfRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM60) & "</ITRForm:HalfRateDeprAmt>"
'    End If
'
'    If DepreciationAtFullRate_DPM60 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM60) & "</ITRForm:DepreciationAtFullRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM60) & "</ITRForm:DepreciationAtFullRate>"
'    End If
'
'    If DepreciationAtHalfRate_DPM60 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM60) & "</ITRForm:DepreciationAtHalfRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM60) & "</ITRForm:DepreciationAtHalfRate>"
'    End If
'
'    If AddlnDeprOnGT180DayAdditions_DPM60 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM60) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM60) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    End If
'
'    If AddlnDeprDuringYearAdditions_DPM60 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM60) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM60) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    End If
'
'    If AddlnDeprlessthan180days_DPM60 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM60) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    End If
'
'    If TotalDepreciation_DPM60 <> "" Then
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM60) & "</ITRForm:TotalDepreciation>"
'    Else
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM60) & "</ITRForm:TotalDepreciation>"
'    End If
'
'    If ExpdrOnTrforSaleAsset_DPM60 <> "" Then
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM60) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    Else
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM60) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    End If
'
'    If CapGainUs50_DPM60 <> "" Then
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM60) & "</ITRForm:CapGainUs50>"
'    Else
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM60) & "</ITRForm:CapGainUs50>"
'    End If
'
'    If WDVLastDay_DPM60 <> "" Then
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM60) & "</ITRForm:WDVLastDay>"
'    Else
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM60) & "</ITRForm:WDVLastDay>"
'    End If
'
'    writeXML "                  </ITRForm:DepreciationDetail>"
'    writeXML "              </ITRForm:Rate60>"
'
'    writeXML "              <ITRForm:Rate80>"
'    writeXML "                  <ITRForm:DepreciationDetail>"
'    If WDVFirstDay_DPM80 <> "" Then
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM80) & "</ITRForm:WDVFirstDay>"
'    Else
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM80) & "</ITRForm:WDVFirstDay>"
'    End If
'
'    If AdditionsGrThan180Days_DPM80 <> "" Then
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM80) & "</ITRForm:AdditionsGrThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM80) & "</ITRForm:AdditionsGrThan180Days>"
'    End If
'
'    If RealizationTotalPeriod_DPM80 <> "" Then
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM80) & "</ITRForm:RealizationTotalPeriod>"
'    Else
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM80) & "</ITRForm:RealizationTotalPeriod>"
'    End If
'
'    If FullRateDeprAmt_DPM80 <> "" Then
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM80) & "</ITRForm:FullRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM80) & "</ITRForm:FullRateDeprAmt>"
'    End If
'
'    If AdditionsLessThan180Days_DPM80 <> "" Then
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM80) & "</ITRForm:AdditionsLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM80) & "</ITRForm:AdditionsLessThan180Days>"
'    End If
'
'    If RealizationPeriodDuringYear_DPM80 <> "" Then
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM80) & "</ITRForm:RealizationPeriodLessThan180days>"
'    Else
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM80) & "</ITRForm:RealizationPeriodLessThan180days>"
'    End If
'
'    If HalfRateDeprAmt_DPM80 <> "" Then
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM80) & "</ITRForm:HalfRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM80) & "</ITRForm:HalfRateDeprAmt>"
'    End If
'
'    If DepreciationAtFullRate_DPM80 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM80) & "</ITRForm:DepreciationAtFullRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM80) & "</ITRForm:DepreciationAtFullRate>"
'    End If
'
'    If DepreciationAtHalfRate_DPM80 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM80) & "</ITRForm:DepreciationAtHalfRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM80) & "</ITRForm:DepreciationAtHalfRate>"
'    End If
'
'    If AddlnDeprOnGT180DayAdditions_DPM80 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM80) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM80) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    End If
'
'    If AddlnDeprDuringYearAdditions_DPM80 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM80) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM80) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    End If
'
'    If AddlnDeprlessthan180days_DPM80 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM80) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    End If
'
'    If TotalDepreciation_DPM80 <> "" Then
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM80) & "</ITRForm:TotalDepreciation>"
'    Else
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM80) & "</ITRForm:TotalDepreciation>"
'    End If
'
'    If ExpdrOnTrforSaleAsset_DPM80 <> "" Then
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM80) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    Else
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM80) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    End If
'
'    If CapGainUs50_DPM80 <> "" Then
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM80) & "</ITRForm:CapGainUs50>"
'    Else
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM80) & "</ITRForm:CapGainUs50>"
'    End If
'
'    If WDVLastDay_DPM80 <> "" Then
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM80) & "</ITRForm:WDVLastDay>"
'    Else
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM80) & "</ITRForm:WDVLastDay>"
'    End If
'    writeXML "                  </ITRForm:DepreciationDetail>"
'    writeXML "              </ITRForm:Rate80>"
'
'
'    writeXML "              <ITRForm:Rate100>"
'    writeXML "                  <ITRForm:DepreciationDetail>"
'    If WDVFirstDay_DPM100 <> "" Then
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DPM100) & "</ITRForm:WDVFirstDay>"
'    Else
'        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DPM100) & "</ITRForm:WDVFirstDay>"
'    End If
'
'    If AdditionsGrThan180Days_DPM100 <> "" Then
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DPM100) & "</ITRForm:AdditionsGrThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DPM100) & "</ITRForm:AdditionsGrThan180Days>"
'    End If
'
'    If RealizationTotalPeriod_DPM100 <> "" Then
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DPM100) & "</ITRForm:RealizationTotalPeriod>"
'    Else
'        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DPM100) & "</ITRForm:RealizationTotalPeriod>"
'    End If
'
'    If FullRateDeprAmt_DPM100 <> "" Then
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DPM100) & "</ITRForm:FullRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DPM100) & "</ITRForm:FullRateDeprAmt>"
'    End If
'
'    If AdditionsLessThan180Days_DPM100 <> "" Then
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DPM100) & "</ITRForm:AdditionsLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DPM100) & "</ITRForm:AdditionsLessThan180Days>"
'    End If
'
'    If RealizationPeriodDuringYear_DPM100 <> "" Then
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DPM100) & "</ITRForm:RealizationPeriodLessThan180days>"
'    Else
'        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DPM100) & "</ITRForm:RealizationPeriodLessThan180days>"
'    End If
'
'    If HalfRateDeprAmt_DPM100 <> "" Then
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DPM100) & "</ITRForm:HalfRateDeprAmt>"
'    Else
'        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DPM100) & "</ITRForm:HalfRateDeprAmt>"
'    End If
'
'    If DepreciationAtFullRate_DPM100 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DPM100) & "</ITRForm:DepreciationAtFullRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DPM100) & "</ITRForm:DepreciationAtFullRate>"
'    End If
'
'    If DepreciationAtHalfRate_DPM100 <> "" Then
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DPM100) & "</ITRForm:DepreciationAtHalfRate>"
'    Else
'        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DPM100) & "</ITRForm:DepreciationAtHalfRate>"
'    End If
'
'    If AddlnDeprOnGT180DayAdditions_DPM100 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DPM100) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DPM100) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
'    End If
'
'    If AddlnDeprDuringYearAdditions_DPM100 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DPM100) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM100) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
'    End If
'
'    If AddlnDeprlessthan180days_DPM100 <> "" Then
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(AddlnDeprlessthan180days_DPM100) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    Else
'        writeXML "                       <ITRForm:AddlnDeprOnAssetLessThan180Days>" & UCase(DefaultAddlnDeprDuringYearAdditions_DPM30) & "</ITRForm:AddlnDeprOnAssetLessThan180Days>"
'    End If
'
'    If TotalDepreciation_DPM100 <> "" Then
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DPM100) & "</ITRForm:TotalDepreciation>"
'    Else
'        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DPM100) & "</ITRForm:TotalDepreciation>"
'    End If
'
'    If ExpdrOnTrforSaleAsset_DPM100 <> "" Then
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DPM100) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    Else
'        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DPM100) & "</ITRForm:ExpdrOnTrforSaleAsset>"
'    End If
'
'    If CapGainUs50_DPM100 <> "" Then
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DPM100) & "</ITRForm:CapGainUs50>"
'    Else
'        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DPM100) & "</ITRForm:CapGainUs50>"
'    End If
'
'    If WDVLastDay_DPM100 <> "" Then
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DPM100) & "</ITRForm:WDVLastDay>"
'    Else
'        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DPM100) & "</ITRForm:WDVLastDay>"
'    End If
'    writeXML "                  </ITRForm:DepreciationDetail>"
'    writeXML "              </ITRForm:Rate100>"
    writeXML "          </ITRForm:PlantMachinery>"
    writeXML "      </ITRForm:ScheduleDPM>"
End Function


Function ScheduleDOA()
writeXML "      <ITRForm:ScheduleDOA>"
    writeXML "          <ITRForm:Land>"
    writeXML "                  <ITRForm:DepreciationDetail>"

    If WDVFirstDay_DAOL0 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOL0) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOL0) & "</ITRForm:WDVFirstDay>"
    End If

    If WDVLastDay_DAOB5 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOL0) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOL0) & "</ITRForm:WDVLastDay>"
    End If
    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "          </ITRForm:Land>"

    writeXML "          <ITRForm:Building>"
    writeXML "              <ITRForm:Rate5>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DAOB5 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOB5) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOB5) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DAOB5 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DAOB5) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DAOB5) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DAOB5 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DAOB5) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DAOB5) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DAOB5 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DAOB5) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DAOB5) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DAOB5 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DAOB5) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DAOB5) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DAOB5 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DAOB5) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DAOB5) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DAOB5 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DAOB5) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DAOB5) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DAOB5 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DAOB5) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DAOB5) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DAOB5 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DAOB5) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DAOB5) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    'If AddlnDeprOnGT180DayAdditions_DAOB5 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DAOB5) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DAOB5) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'End If

    'If AddlnDeprDuringYearAdditions_DAOB5 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DAOB5) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DAOB5) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'End If

    If TotalDepreciation_DAOB5 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DAOB5) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DAOB5) & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DAOB05 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DAOB05) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DAOB05 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DAOB05) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DAOB05 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DAOB05) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DAOB5 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DAOB5) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DAOB5) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DAOB5 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DAOB5) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DAOB5) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DAOB5 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOB5) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOB5) & "</ITRForm:WDVLastDay>"
    End If
    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate5>"

    writeXML "              <ITRForm:Rate10>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DAOB10 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOB10) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOB10) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DAOB10 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DAOB10) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DAOB10) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DAOB10 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DAOB10) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DAOB10) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DAOB10 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DAOB10) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DAOB10) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DAOB10 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DAOB10) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DAOB10) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DAOB10 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DAOB10) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DAOB10) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DAOB10 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DAOB10) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DAOB10) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DAOB10 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DAOB10) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DAOB10) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DAOB10 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DAOB10) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DAOB10) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    'If AddlnDeprOnGT180DayAdditions_DAOB10 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DAOB10) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DAOB10) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'End If

    'If AddlnDeprDuringYearAdditions_DAOB10 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DAOB10) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DAOB10) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'End If

    If TotalDepreciation_DAOB10 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DAOB10) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DAOB10) & "</ITRForm:TotalDepreciation>"
    End If

     If Depreciation_Sec38_DAOB10 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DAOB10) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DAOB10 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DAOB10) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DAOB10 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DAOB10) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If


    If ExpdrOnTrforSaleAsset_DAOB10 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DAOB10) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DAOB10) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DAOB10 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DAOB10) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DAOB10) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DAOB10 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOB10) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOB10) & "</ITRForm:WDVLastDay>"
    End If

    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate10>"

    writeXML "              <ITRForm:Rate40>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DAOB100 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOB100) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOB100) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DAOB100 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DAOB100) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DAOB100) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DAOB100 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DAOB100) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DAOB100) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DAOB100 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DAOB100) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DAOB100) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DAOB100 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DAOB100) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DAOB100) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DAOB100 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DAOB100) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DAOB100) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DAOB100 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DAOB100) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DAOB100) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DAOB100 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DAOB100) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DAOB100) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DAOB100 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DAOB100) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DAOB100) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    'If AddlnDeprOnGT180DayAdditions_DAOB100 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DAOB100) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DAOB100) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'End If

    'If AddlnDeprDuringYearAdditions_DAOB100 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DAOB100) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DAOB100) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'End If

    If TotalDepreciation_DAOB100 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DAOB100) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DAOB100) & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DAOB40 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DAOB40) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DAOB40 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DAOB40) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DAOB40 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DAOB40) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DAOB100 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DAOB100) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DAOB100) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DAOB100 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DAOB100) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DAOB100) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DAOB100 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOB100) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOB100) & "</ITRForm:WDVLastDay>"
    End If

    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate40>"

    writeXML "          </ITRForm:Building>"

    writeXML "          <ITRForm:FurnitureFittings>"
    writeXML "              <ITRForm:Rate10>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DAOF10 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOF10) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOF10) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DAOF10 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DAOF10) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DAOF10) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DAOF10 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DAOF10) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DAOF10) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DAOF10 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DAOF10) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DAOF10) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DAOF10 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DAOF10) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DAOF10) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DAOF10 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DAOF10) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DAOF10) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DAOF10 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DAOF10) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DAOF10) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DAOF10 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DAOF10) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DAOF10) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DAOF10 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DAOF10) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DAOF10) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    'If AddlnDeprOnGT180DayAdditions_DAOF10 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DAOF10) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'Else
    '    writeXML "                      DAOI25.AddlnDeprOnGT180DayAdditions <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DAOF10) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'End If

    'If AddlnDeprDuringYearAdditions_DAOF10 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DAOF10) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DAOF10) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'End If

    If TotalDepreciation_DAOF10 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DAOF10) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DAOF10) & "</ITRForm:TotalDepreciation>"
    End If


    If Depreciation_Sec38_DAOF10 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DAOF10) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DAOF10 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DAOF10) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DAOF10 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DAOF10) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DAOF10 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DAOF10) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DAOF10) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DAOF10 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DAOF10) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DAOF10) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DAOF10 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOF10) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOF10) & "</ITRForm:WDVLastDay>"
    End If
    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate10>"
    writeXML "          </ITRForm:FurnitureFittings>"

    writeXML "          <ITRForm:IntangibleAssets>"
    writeXML "              <ITRForm:Rate25>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DAOI25 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOI25) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOI25) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DAOI25 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DAOI25) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DAOI25) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DAOI25 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DAOI25) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DAOI25) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DAOI25 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DAOI25) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DAOI25) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DAOI25 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DAOI25) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DAOI25) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DAOI25 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DAOI25) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DAOI25) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DAOI25 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DAOI25) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DAOI25) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DAOI25 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DAOI25) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DAOI25) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DAOI25 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DAOI25) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DAOI25) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    'If AddlnDeprOnGT180DayAdditions_DAOI25 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DAOI25) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DAOI25) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'End If

    'If AddlnDeprDuringYearAdditions_DAOI25 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DAOI25) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DAOI25) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'End If

    If TotalDepreciation_DAOI25 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DAOI25) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DAOI25) & "</ITRForm:TotalDepreciation>"
    End If

    If Depreciation_Sec38_DAOI25 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DAOI25) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DAOI25 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DAOI25) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DAOI25 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DAOI25) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DAOI25 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DAOI25) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DAOI25) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DAOI25 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DAOI25) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DAOI25) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DAOI25 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOI25) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOI25) & "</ITRForm:WDVLastDay>"
    End If
    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate25>"
    writeXML "          </ITRForm:IntangibleAssets>"

    writeXML "          <ITRForm:Ships>"
    writeXML "              <ITRForm:Rate20>"
    writeXML "                  <ITRForm:DepreciationDetail>"
    If WDVFirstDay_DAOS20 <> "" Then
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(WDVFirstDay_DAOS20) & "</ITRForm:WDVFirstDay>"
    Else
        writeXML "                       <ITRForm:WDVFirstDay>" & UCase(DefaultWDVFirstDay_DAOS20) & "</ITRForm:WDVFirstDay>"
    End If

    If AdditionsGrThan180Days_DAOS20 <> "" Then
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(AdditionsGrThan180Days_DAOS20) & "</ITRForm:AdditionsGrThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsGrThan180Days>" & UCase(DefaultAdditionsGrThan180Days_DAOS20) & "</ITRForm:AdditionsGrThan180Days>"
    End If

    If RealizationTotalPeriod_DAOS20 <> "" Then
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(RealizationTotalPeriod_DAOS20) & "</ITRForm:RealizationTotalPeriod>"
    Else
        writeXML "                       <ITRForm:RealizationTotalPeriod>" & UCase(DefaultRealizationTotalPeriod_DAOS20) & "</ITRForm:RealizationTotalPeriod>"
    End If

    If FullRateDeprAmt_DAOS20 <> "" Then
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(FullRateDeprAmt_DAOS20) & "</ITRForm:FullRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:FullRateDeprAmt>" & UCase(DefaultFullRateDeprAmt_DAOS20) & "</ITRForm:FullRateDeprAmt>"
    End If

    If AdditionsLessThan180Days_DAOS20 <> "" Then
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(AdditionsLessThan180Days_DAOS20) & "</ITRForm:AdditionsLessThan180Days>"
    Else
        writeXML "                       <ITRForm:AdditionsLessThan180Days>" & UCase(DefaultAdditionsLessThan180Days_DAOS20) & "</ITRForm:AdditionsLessThan180Days>"
    End If

    If RealizationPeriodDuringYear_DAOS20 <> "" Then
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(RealizationPeriodDuringYear_DAOS20) & "</ITRForm:RealizationPeriodLessThan180days>"
    Else
        writeXML "                       <ITRForm:RealizationPeriodLessThan180days>" & UCase(DefaultRealizationPeriodDuringYear_DAOS20) & "</ITRForm:RealizationPeriodLessThan180days>"
    End If

    If HalfRateDeprAmt_DAOS20 <> "" Then
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(HalfRateDeprAmt_DAOS20) & "</ITRForm:HalfRateDeprAmt>"
    Else
        writeXML "                       <ITRForm:HalfRateDeprAmt>" & UCase(DefaultHalfRateDeprAmt_DAOS20) & "</ITRForm:HalfRateDeprAmt>"
    End If

    If DepreciationAtFullRate_DAOS20 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DepreciationAtFullRate_DAOS20) & "</ITRForm:DepreciationAtFullRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtFullRate>" & UCase(DefaultDepreciationAtFullRate_DAOS20) & "</ITRForm:DepreciationAtFullRate>"
    End If

    If DepreciationAtHalfRate_DAOS20 <> "" Then
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DepreciationAtHalfRate_DAOS20) & "</ITRForm:DepreciationAtHalfRate>"
    Else
        writeXML "                       <ITRForm:DepreciationAtHalfRate>" & UCase(DefaultDepreciationAtHalfRate_DAOS20) & "</ITRForm:DepreciationAtHalfRate>"
    End If

    'If AddlnDeprOnGT180DayAdditions_DAOS20 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(AddlnDeprOnGT180DayAdditions_DAOS20) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnGT180DayAdditions>" & UCase(DefaultAddlnDeprOnGT180DayAdditions_DAOS20) & "</ITRForm:AddlnDeprOnGT180DayAdditions>"
    'End If

    'If AddlnDeprDuringYearAdditions_DAOS20 <> "" Then
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(AddlnDeprDuringYearAdditions_DAOS20) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'Else
    '    writeXML "                       <ITRForm:AddlnDeprOnLessThan180DayAdditions>" & UCase(DefaultAddlnDeprDuringYearAdditions_DAOS20) & "</ITRForm:AddlnDeprOnLessThan180DayAdditions>"
    'End If

    If TotalDepreciation_DAOS20 <> "" Then
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DAOS20) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "                       <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DAOS20) & "</ITRForm:TotalDepreciation>"
    End If


    If Depreciation_Sec38_DAOS20 <> "" Then
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & UCase(Depreciation_Sec38_DAOS20) & "</ITRForm:DepDisAllowUs38_2>"
    Else
        writeXML "                       <ITRForm:DepDisAllowUs38_2>" & "0" & "</ITRForm:DepDisAllowUs38_2>"
    End If

    If NetaggregateDep_DAOS20 <> "" Then
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & UCase(NetaggregateDep_DAOS20) & "</ITRForm:NetAggregateDepreciation>"
    Else
        writeXML "                       <ITRForm:NetAggregateDepreciation>" & "0" & "</ITRForm:NetAggregateDepreciation>"
    End If

    If PropaggregateDep_DAOS20 <> "" Then
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & UCase(PropaggregateDep_DAOS20) & "</ITRForm:ProportionateAggDepreciation>"
    Else
        writeXML "                       <ITRForm:ProportionateAggDepreciation>" & "0" & "</ITRForm:ProportionateAggDepreciation>"
    End If

    If ExpdrOnTrforSaleAsset_DAOS20 <> "" Then
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(ExpdrOnTrforSaleAsset_DAOS20) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    Else
        writeXML "                       <ITRForm:ExpdrOnTrforSaleAsset>" & UCase(DefaultExpdrOnTrforSaleAsset_DAOS20) & "</ITRForm:ExpdrOnTrforSaleAsset>"
    End If

    If CapGainUs50_DAOS20 <> "" Then
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(CapGainUs50_DAOS20) & "</ITRForm:CapGainUs50>"
    Else
        writeXML "                       <ITRForm:CapGainUs50>" & UCase(DefaultCapGainUs50_DAOS20) & "</ITRForm:CapGainUs50>"
    End If

    If WDVLastDay_DAOS20 <> "" Then
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(WDVLastDay_DAOS20) & "</ITRForm:WDVLastDay>"
    Else
        writeXML "                       <ITRForm:WDVLastDay>" & UCase(DefaultWDVLastDay_DAOS20) & "</ITRForm:WDVLastDay>"
    End If
    writeXML "                  </ITRForm:DepreciationDetail>"
    writeXML "              </ITRForm:Rate20>"
    writeXML "          </ITRForm:Ships>"
    writeXML "      </ITRForm:ScheduleDOA>"
End Function

Function ScheduleDEP()
    writeXML "      <ITRForm:ScheduleDEP>"
    writeXML "          <ITRForm:SummaryFromDeprSch>"
    writeXML "              <ITRForm:PlantMachinerySummary>"

    If DeprBlockTot15Percent_DEPP <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot15Percent>" & UCase(DeprBlockTot15Percent_DEPP) & "</ITRForm:DeprBlockTot15Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot15Percent>" & UCase(DefaultDeprBlockTot15Percent_DEPP) & "</ITRForm:DeprBlockTot15Percent>"
    End If

    If DeprBlockTot30Percent_DEPP <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot30Percent>" & UCase(DeprBlockTot30Percent_DEPP) & "</ITRForm:DeprBlockTot30Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot30Percent>" & UCase(DefaultDeprBlockTot30Percent_DEPP) & "</ITRForm:DeprBlockTot30Percent>"
    End If

    If DeprBlockTot40Percent_DEPP <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DeprBlockTot40Percent_DEPP) & "</ITRForm:DeprBlockTot40Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DefaultDeprBlockTot40Percent_DEPP) & "</ITRForm:DeprBlockTot40Percent>"
    End If

    If Sheet11.Range("DEPP.DeprBlockTot45Percent").value <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot45Percent>" & Sheet11.Range("DEPP.DeprBlockTot45Percent").value & "</ITRForm:DeprBlockTot45Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot45Percent>" & 0 & "</ITRForm:DeprBlockTot45Percent>"
    End If

'    If DeprBlockTot50Percent_DEPP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot50Percent>" & UCase(DeprBlockTot50Percent_DEPP) & "</ITRForm:DeprBlockTot50Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot50Percent>" & UCase(DefaultDeprBlockTot50Percent_DEPP) & "</ITRForm:DeprBlockTot50Percent>"
'    End If
'
'    If DeprBlockTot60Percent_DEPP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot60Percent>" & UCase(DeprBlockTot60Percent_DEPP) & "</ITRForm:DeprBlockTot60Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot60Percent>" & UCase(DefaultDeprBlockTot60Percent_DEPP) & "</ITRForm:DeprBlockTot60Percent>"
'    End If
'
'    If DeprBlockTot80Percent_DEPP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot80Percent>" & UCase(DeprBlockTot80Percent_DEPP) & "</ITRForm:DeprBlockTot80Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot80Percent>" & UCase(DefaultDeprBlockTot80Percent_DEPP) & "</ITRForm:DeprBlockTot80Percent>"
'    End If
'
'    If DeprBlockTot100Percent_DEPP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot100Percent>" & UCase(DeprBlockTot100Percent_DEPP) & "</ITRForm:DeprBlockTot100Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot100Percent>" & UCase(DefaultDeprBlockTot100Percent_DEPP) & "</ITRForm:DeprBlockTot100Percent>"
'    End If

    If TotPlntMach_DEPP <> "" Then
        writeXML "                   <ITRForm:TotPlntMach>" & UCase(TotPlntMach_DEPP) & "</ITRForm:TotPlntMach>"
    Else
        writeXML "                   <ITRForm:TotPlntMach>" & UCase(DefaultTotPlntMach_DEPP) & "</ITRForm:TotPlntMach>"
    End If
    writeXML "              </ITRForm:PlantMachinerySummary>"

    writeXML "              <ITRForm:BuildingSummary>"
    If DeprBlockTot5Percent_DEPB <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot5Percent>" & UCase(DeprBlockTot5Percent_DEPB) & "</ITRForm:DeprBlockTot5Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot5Percent>" & UCase(DefaultDeprBlockTot5Percent_DEPB) & "</ITRForm:DeprBlockTot5Percent>"
    End If

    If DeprBlockTot10Percent_DEPB <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot10Percent>" & UCase(DeprBlockTot10Percent_DEPB) & "</ITRForm:DeprBlockTot10Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot10Percent>" & UCase(DefaultDeprBlockTot10Percent_DEPB) & "</ITRForm:DeprBlockTot10Percent>"
    End If

    If DeprBlockTot100Percent_DEPB <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DeprBlockTot100Percent_DEPB) & "</ITRForm:DeprBlockTot40Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DefaultDeprBlockTot100Percent_DEPB) & "</ITRForm:DeprBlockTot40Percent>"
    End If

    If TotBuildng_DEPB <> "" Then
        writeXML "                   <ITRForm:TotBuildng>" & UCase(TotBuildng_DEPB) & "</ITRForm:TotBuildng>"
    Else
        writeXML "                   <ITRForm:TotBuildng>" & UCase(DefaultTotBuildng_DEPB) & "</ITRForm:TotBuildng>"
    End If

    writeXML "              </ITRForm:BuildingSummary>"
    If FurnitureSummary_DEP <> "" Then
        writeXML "               <ITRForm:FurnitureSummary>" & UCase(FurnitureSummary_DEP) & "</ITRForm:FurnitureSummary>"
    Else
        writeXML "               <ITRForm:FurnitureSummary>" & UCase(DefaultFurnitureSummary_DEP) & "</ITRForm:FurnitureSummary>"
    End If

    If IntangibleAssetSummary_DEP <> "" Then
        writeXML "               <ITRForm:IntangibleAssetSummary>" & UCase(IntangibleAssetSummary_DEP) & "</ITRForm:IntangibleAssetSummary>"
    Else
        writeXML "               <ITRForm:IntangibleAssetSummary>" & UCase(DefaultIntangibleAssetSummary_DEP) & "</ITRForm:IntangibleAssetSummary>"
    End If

    If ShipsSummary_DEP <> "" Then
        writeXML "               <ITRForm:ShipsSummary>" & UCase(ShipsSummary_DEP) & "</ITRForm:ShipsSummary>"
    Else
        writeXML "               <ITRForm:ShipsSummary>" & UCase(DefaultShipsSummary_DEP) & "</ITRForm:ShipsSummary>"
    End If

    If TotalDepreciation_DEP <> "" Then
        writeXML "               <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DEP) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "               <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DEP) & "</ITRForm:TotalDepreciation>"
    End If

    writeXML "          </ITRForm:SummaryFromDeprSch>"
    writeXML "      </ITRForm:ScheduleDEP>"
End Function

Function ScheduleDCG()
writeXML "      <ITRForm:ScheduleDCG>"
    writeXML "          <ITRForm:SummaryFromDeprSchCG>"
    writeXML "              <ITRForm:PlantMachinerySummaryCG>"
    If DeprBlockTot15Percent_DCGP <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot15Percent>" & UCase(DeprBlockTot15Percent_DCGP) & "</ITRForm:DeprBlockTot15Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot15Percent>" & UCase(DefaultDeprBlockTot15Percent_DCGP) & "</ITRForm:DeprBlockTot15Percent>"
    End If

    If DeprBlockTot30Percent_DCGP <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot30Percent>" & UCase(DeprBlockTot30Percent_DCGP) & "</ITRForm:DeprBlockTot30Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot30Percent>" & UCase(DefaultDeprBlockTot30Percent_DCGP) & "</ITRForm:DeprBlockTot30Percent>"
    End If

    If DeprBlockTot40Percent_DCGP <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DeprBlockTot40Percent_DCGP) & "</ITRForm:DeprBlockTot40Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DefaultDeprBlockTot40Percent_DCGP) & "</ITRForm:DeprBlockTot40Percent>"
    End If

    If Sheet11.Range("DCGP.DeprBlockTot45Percent").value <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot45Percent>" & Sheet11.Range("DCGP.DeprBlockTot45Percent").value & "</ITRForm:DeprBlockTot45Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot45Percent>" & 0 & "</ITRForm:DeprBlockTot45Percent>"
    End If

'    If DeprBlockTot50Percent_DCGP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot50Percent>" & UCase(DeprBlockTot50Percent_DCGP) & "</ITRForm:DeprBlockTot50Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot50Percent>" & UCase(DefaultDeprBlockTot50Percent_DCGP) & "</ITRForm:DeprBlockTot50Percent>"
'    End If
'
'    If DeprBlockTot60Percent_DCGP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot60Percent>" & UCase(DeprBlockTot60Percent_DCGP) & "</ITRForm:DeprBlockTot60Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot60Percent>" & UCase(DefaultDeprBlockTot60Percent_DCGP) & "</ITRForm:DeprBlockTot60Percent>"
'    End If
'
'    If DeprBlockTot80Percent_DCGP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot80Percent>" & UCase(DeprBlockTot80Percent_DCGP) & "</ITRForm:DeprBlockTot80Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot80Percent>" & UCase(DefaultDeprBlockTot80Percent_DCGP) & "</ITRForm:DeprBlockTot80Percent>"
'    End If
'
'    If DeprBlockTot100Percent_DCGP <> "" Then
'        writeXML "                   <ITRForm:DeprBlockTot100Percent>" & UCase(DeprBlockTot100Percent_DCGP) & "</ITRForm:DeprBlockTot100Percent>"
'    Else
'        writeXML "                   <ITRForm:DeprBlockTot100Percent>" & UCase(DefaultDeprBlockTot100Percent_DCGP) & "</ITRForm:DeprBlockTot100Percent>"
'    End If

    If TotPlntMach_DCGP <> "" Then
        writeXML "                   <ITRForm:TotPlntMach>" & UCase(TotPlntMach_DCGP) & "</ITRForm:TotPlntMach>"
    Else
        writeXML "                   <ITRForm:TotPlntMach>" & UCase(DefaultTotPlntMach_DCGP) & "</ITRForm:TotPlntMach>"
    End If
    writeXML "              </ITRForm:PlantMachinerySummaryCG>"

    writeXML "              <ITRForm:BuildingSummaryCG>"
    If DeprBlockTot5Percent_DCGB <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot5Percent>" & UCase(DeprBlockTot5Percent_DCGB) & "</ITRForm:DeprBlockTot5Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot5Percent>" & UCase(DefaultDeprBlockTot5Percent_DCGB) & "</ITRForm:DeprBlockTot5Percent>"
    End If

    If DeprBlockTot10Percent_DCGB <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot10Percent>" & UCase(DeprBlockTot10Percent_DCGB) & "</ITRForm:DeprBlockTot10Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot10Percent>" & UCase(DefaultDeprBlockTot10Percent_DCGB) & "</ITRForm:DeprBlockTot10Percent>"
    End If

    If DeprBlockTot100Percent_DCGB <> "" Then
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DeprBlockTot100Percent_DCGB) & "</ITRForm:DeprBlockTot40Percent>"
    Else
        writeXML "                   <ITRForm:DeprBlockTot40Percent>" & UCase(DefaultDeprBlockTot100Percent_DCGB) & "</ITRForm:DeprBlockTot40Percent>"
    End If

    If TotBuildng_DCGB <> "" Then
        writeXML "                   <ITRForm:TotBuildng>" & UCase(TotBuildng_DCGB) & "</ITRForm:TotBuildng>"
    Else
        writeXML "                   <ITRForm:TotBuildng>" & UCase(DefaultTotBuildng_DCGB) & "</ITRForm:TotBuildng>"
    End If
    writeXML "              </ITRForm:BuildingSummaryCG>"
    If FurnitureSummary_DCG <> "" Then
        writeXML "               <ITRForm:FurnitureSummary>" & UCase(FurnitureSummary_DCG) & "</ITRForm:FurnitureSummary>"
    Else
        writeXML "               <ITRForm:FurnitureSummary>" & UCase(DefaultFurnitureSummary_DCG) & "</ITRForm:FurnitureSummary>"
    End If

    If IntangibleAssetSummary_DCG <> "" Then
        writeXML "               <ITRForm:IntangibleAssetSummary>" & UCase(IntangibleAssetSummary_DCG) & "</ITRForm:IntangibleAssetSummary>"
    Else
        writeXML "               <ITRForm:IntangibleAssetSummary>" & UCase(DefaultIntangibleAssetSummary_DCG) & "</ITRForm:IntangibleAssetSummary>"
    End If

    If ShipsSummary_DCG <> "" Then
        writeXML "               <ITRForm:ShipsSummary>" & UCase(ShipsSummary_DCG) & "</ITRForm:ShipsSummary>"
    Else
        writeXML "               <ITRForm:ShipsSummary>" & UCase(DefaultShipsSummary_DCG) & "</ITRForm:ShipsSummary>"
    End If

    If TotalDepreciation_DCG <> "" Then
        writeXML "               <ITRForm:TotalDepreciation>" & UCase(TotalDepreciation_DCG) & "</ITRForm:TotalDepreciation>"
    Else
        writeXML "               <ITRForm:TotalDepreciation>" & UCase(DefaultTotalDepreciation_DCG) & "</ITRForm:TotalDepreciation>"
    End If
    writeXML "          </ITRForm:SummaryFromDeprSchCG>"
    writeXML "      </ITRForm:ScheduleDCG>"
End Function

Function ScheduleESR()
    If ExcessAmtOverDebPL_ESRTOT <> -99 Then
        writeXML "      <ITRForm:ScheduleESR>"
        writeXML "          <ITRForm:DeductionUs35>"
        writeXML "              <ITRForm:Section35_1_i>"
        writeXML "                  <ITRForm:DeductUs35>"

        If AmtDebPL_ESR1i <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESR1i) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESR1i) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESR1i <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESR1i) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESR1i) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESR1i <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESR1i) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESR1i) & "</ITRForm:ExcessAmtOverDebPL>"
        End If

        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:Section35_1_i>"
        writeXML "              <ITRForm:Section35_1_ii>"
        writeXML "                  <ITRForm:DeductUs35>"
        If AmtDebPL_ESR1ii <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESR1ii) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESR1ii) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESR1ii <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESR1ii) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESR1ii) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESR1ii <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESR1ii) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESR1ii) & "</ITRForm:ExcessAmtOverDebPL>"
        End If

        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:Section35_1_ii>"
        writeXML "<ITRForm:Section35_1_iia>"
        writeXML "<ITRForm:DeductUs35>"
        writeXML "<ITRForm:AmtDebPL>" & UVCase(AmtDebPL_ESR1iia) & "</ITRForm:AmtDebPL>"
        writeXML "<ITRForm:AmtUs35Allowable>" & UVCase(AmtUs35Allowable_ESR1iia) & "</ITRForm:AmtUs35Allowable>"
        writeXML "<ITRForm:ExcessAmtOverDebPL>" & UVCase(ExcessAmtOverDebPL_ESR1iia) & "</ITRForm:ExcessAmtOverDebPL>"
        writeXML "</ITRForm:DeductUs35>"
        writeXML "</ITRForm:Section35_1_iia>"

        writeXML "              <ITRForm:Section35_1_iii>"
        writeXML "                  <ITRForm:DeductUs35>"
        If AmtDebPL_ESR1iii <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESR1iii) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESR1iii) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESR1iii <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESR1iii) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESR1iii) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESR1iii <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESR1iii) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESR1iii) & "</ITRForm:ExcessAmtOverDebPL>"
        End If
        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:Section35_1_iii>"
        writeXML "              <ITRForm:Section35_1_iv>"
        writeXML "                  <ITRForm:DeductUs35>"

        If AmtDebPL_ESR1iv <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESR1iv) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESR1iv) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESR1iv <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESR1iv) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESR1iv) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESR1iv <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESR1iv) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESR1iv) & "</ITRForm:ExcessAmtOverDebPL>"
        End If
        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:Section35_1_iv>"
        writeXML "              <ITRForm:Section35_2AA>"
        writeXML "                  <ITRForm:DeductUs35>"

        If AmtDebPL_ESR2AA <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESR2AA) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESR2AA) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESR2AA <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESR2AA) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESR2AA) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESR2AA <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESR2AA) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESR2AA) & "</ITRForm:ExcessAmtOverDebPL>"
        End If
        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:Section35_2AA>"
        writeXML "              <ITRForm:Section35_2AB>"
        writeXML "                  <ITRForm:DeductUs35>"

        If AmtDebPL_ESR2AB <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESR2AB) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESR2AB) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESR2AB <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESR2AB) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESR2AB) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESR2AB <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESR2AB) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESR2AB) & "</ITRForm:ExcessAmtOverDebPL>"
        End If

        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:Section35_2AB>"

        writeXML "<ITRForm:Section35_CCC>"
        writeXML "<ITRForm:DeductUs35>"
        writeXML "<ITRForm:AmtDebPL>" & UVCase(AmtDebPL_ESR35CCC) & "</ITRForm:AmtDebPL>"
        writeXML "<ITRForm:AmtUs35Allowable>" & UVCase(AmtUs35Allowable_ESR35CCC) & "</ITRForm:AmtUs35Allowable>"
        writeXML "<ITRForm:ExcessAmtOverDebPL>" & UVCase(ExcessAmtOverDebPL_ESR35CCC) & "</ITRForm:ExcessAmtOverDebPL>"
        writeXML "</ITRForm:DeductUs35>"
        writeXML "</ITRForm:Section35_CCC>"

        writeXML "<ITRForm:Section35_CCD>"
        writeXML "<ITRForm:DeductUs35>"
        writeXML "<ITRForm:AmtDebPL>" & UVCase(AmtDebPL_ESR35CCD) & "</ITRForm:AmtDebPL>"
        writeXML "<ITRForm:AmtUs35Allowable>" & UVCase(AmtUs35Allowable_ESR35CCD) & "</ITRForm:AmtUs35Allowable>"
        writeXML "<ITRForm:ExcessAmtOverDebPL>" & UVCase(ExcessAmtOverDebPL_ESR35CCD) & "</ITRForm:ExcessAmtOverDebPL>"
        writeXML "</ITRForm:DeductUs35>"
        writeXML "</ITRForm:Section35_CCD>"

        writeXML "              <ITRForm:TotUs35>"
        writeXML "                  <ITRForm:DeductUs35>"
        If AmtDebPL_ESRTOT <> "" Then
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(AmtDebPL_ESRTOT) & "</ITRForm:AmtDebPL>"
        Else
            writeXML "                       <ITRForm:AmtDebPL>" & UCase(DefaultAmtDebPL_ESRTOT) & "</ITRForm:AmtDebPL>"
        End If

        If AmtUs35Allowable_ESRTOT <> "" Then
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(AmtUs35Allowable_ESRTOT) & "</ITRForm:AmtUs35Allowable>"
        Else
            writeXML "                       <ITRForm:AmtUs35Allowable>" & UCase(DefaultAmtUs35Allowable_ESRTOT) & "</ITRForm:AmtUs35Allowable>"
        End If

        If ExcessAmtOverDebPL_ESRTOT <> "" Then
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(ExcessAmtOverDebPL_ESRTOT) & "</ITRForm:ExcessAmtOverDebPL>"
        Else
            writeXML "                       <ITRForm:ExcessAmtOverDebPL>" & UCase(DefaultExcessAmtOverDebPL_ESRTOT) & "</ITRForm:ExcessAmtOverDebPL>"
        End If

        writeXML "                  </ITRForm:DeductUs35>"
        writeXML "              </ITRForm:TotUs35>"

        writeXML "          </ITRForm:DeductionUs35>"
        writeXML "      </ITRForm:ScheduleESR>"
    End If

End Function
Function ScheduleCGFor23()
    Dim i As Long
    Dim STCGA1Count As Long
    Dim LTCGA1Count As Long
    Dim num As Long
    STCGA1Count = Sheet13.Range("PropertySectionCount_STCG").value
    LTCGA1Count = Sheet13.Range("PropertySectionCount_LTCG").value
    
    subProcCaption = "Schedule CG"
    writeXML "<ITRForm:ScheduleCGFor23>"
    writeXML "<ITRForm:ShortTermCapGainFor23>"
    writeXML "<ITRForm:SaleofLandBuild>"
    
     For num = 1 To STCGA1Count
     
      writeXML "<ITRForm:SaleofLandBuildDtls>"
      
      If Sheet13.Range("STCG.Aai_FullConsideration" & num).value <> "" Then
        writeXML "<ITRForm:FullConsideration>" & Sheet13.Range("STCG.Aai_FullConsideration" & num).value & "</ITRForm:FullConsideration>"
    Else
        writeXML "<ITRForm:FullConsideration>0</ITRForm:FullConsideration>"
    End If
    
     
      If Sheet13.Range("STCG.Aaii_PropertyValuation" & num).value <> "" Then
        writeXML "<ITRForm:PropertyValuation>" & Sheet13.Range("STCG.Aaii_PropertyValuation" & num).value & "</ITRForm:PropertyValuation>"
    Else
        writeXML "<ITRForm:PropertyValuation>0</ITRForm:PropertyValuation>"
    End If

     If Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC" & num).value <> "" Then
        writeXML "<ITRForm:FullConsideration50C>" & Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC" & num).value & "</ITRForm:FullConsideration50C>"
    Else
        writeXML "<ITRForm:FullConsideration50C>0</ITRForm:FullConsideration50C>"
    End If

    If Sheet13.Range("STCG.Abi_AquisitCost" & num).value <> "" Then
        writeXML "<ITRForm:AquisitCost>" & Sheet13.Range("STCG.Abi_AquisitCost" & num).value & "</ITRForm:AquisitCost>"
    Else
        writeXML "<ITRForm:AquisitCost>0</ITRForm:AquisitCost>"
    End If

    If Sheet13.Range("STCG.Abii_ImproveCost" & num).value <> "" Then
        writeXML "<ITRForm:ImproveCost>" & Sheet13.Range("STCG.Abii_ImproveCost" & num).value & "</ITRForm:ImproveCost>"
    Else
        writeXML "<ITRForm:ImproveCost>0</ITRForm:ImproveCost>"
    End If

    If Sheet13.Range("STCG.Abiii_ExpOnTrans" & num).value <> "" Then
        writeXML "<ITRForm:ExpOnTrans>" & Sheet13.Range("STCG.Abiii_ExpOnTrans" & num).value & "</ITRForm:ExpOnTrans>"
    Else
        writeXML "<ITRForm:ExpOnTrans>0</ITRForm:ExpOnTrans>"
    End If

    If Sheet13.Range("STCG.Abiv_TotalDedn" & num).value <> "" Then
        writeXML "<ITRForm:TotalDedn>" & Sheet13.Range("STCG.Abiv_TotalDedn" & num).value & "</ITRForm:TotalDedn>"
    Else
        writeXML "<ITRForm:TotalDedn>0</ITRForm:TotalDedn>"
    End If

    If Sheet13.Range("STCG.Ac_BalanceCG" & num).value <> "" Then
        writeXML "<ITRForm:Balance>" & Sheet13.Range("STCG.Ac_BalanceCG" & num).value & "</ITRForm:Balance>"
    Else
        writeXML "<ITRForm:Balance>0</ITRForm:Balance>"
    End If

    writeXML "<ITRForm:ExemptionOrDednUs54>"
    If Sheet13.Range("STCG.Adi_ExemptionAmount" & num).value <> "" Then
    writeXML "<ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML "<ITRForm:ExemptionSecCode>54B</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("STCG.Adi_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML "</ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("STCG.Adii_ExemptionAmount" & num).value <> "" Then
    writeXML "<ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML "<ITRForm:ExemptionSecCode>54D</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("STCG.Adii_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML "</ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("STCG.Adiii_ExemptionAmount" & num).value <> "" Then
    writeXML "<ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML "<ITRForm:ExemptionSecCode>54G</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("STCG.Adiii_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML "</ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("STCG.Adiv_ExemptionAmount" & num).value <> "" Then
    writeXML "<ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML "<ITRForm:ExemptionSecCode>54GA</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("STCG.Adiv_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML "</ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("STCG.Ad_ExemptionGrandTotal" & num).value <> "" Then
        writeXML "<ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("STCG.Ad_ExemptionGrandTotal" & num).value) & "</ITRForm:ExemptionGrandTotal>"
    Else
        writeXML "<ITRForm:ExemptionGrandTotal>0</ITRForm:ExemptionGrandTotal>"
    End If
    
      writeXML "</ITRForm:ExemptionOrDednUs54>"
    
        
    writeXML "<ITRForm:CapgainonAssets>" & UVCase(Sheet13.Range("STCG.Ae_CapgainonAssets" & num).value) & "</ITRForm:CapgainonAssets>"
    
    
     STCGImmovablePropertyXML (num)
     writeXML "</ITRForm:SaleofLandBuildDtls>"
    Next

   
    writeXML "</ITRForm:SaleofLandBuild>"
    
    writeXML " <ITRForm:SlumpSaleInStcg>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(STCG_A20ia_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:NetWorthOfDivision>" & UVCase(STCG_A20ib_NetWorthOfDivision) & "</ITRForm:NetWorthOfDivision>"
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(STCG_A20ic_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:SlumpSaleInStcg>"
    
    writeXML "<ITRForm:EquityMFonSTT>"
    writeXML "<ITRForm:MFSectionCode>1A</ITRForm:MFSectionCode>"
    writeXML "<ITRForm:EquityMFonSTTDtls>"
    writeXML "<ITRForm:FullConsideration>" & UVCase(STCG_A2ia_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML "<ITRForm:DeductSec48>"
    writeXML "<ITRForm:AquisitCost>" & UVCase(STCG_A2ibi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML "<ITRForm:ImproveCost>" & UVCase(STCG_A2ibii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML "<ITRForm:ExpOnTrans>" & UVCase(STCG_A2ibiii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML "<ITRForm:TotalDedn>" & UVCase(STCG_A2ibiv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML "</ITRForm:DeductSec48>"
    writeXML "<ITRForm:BalanceCG>" & UVCase(STCG_A2ic_BalanceCG) & "</ITRForm:BalanceCG>"
    writeXML "<ITRForm:LossSec94of7Or94of8>" & UVCase(STCG_A2id_LossSec94of7Or94of8) & "</ITRForm:LossSec94of7Or94of8>"
    writeXML "<ITRForm:CapgainonAssets>" & UVCase(STCG_A2ie_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML "</ITRForm:EquityMFonSTTDtls>"
    writeXML "</ITRForm:EquityMFonSTT>"
    
    writeXML "<ITRForm:EquityMFonSTT>"
    writeXML "<ITRForm:MFSectionCode>5AD1biip</ITRForm:MFSectionCode>"
    writeXML "<ITRForm:EquityMFonSTTDtls>"
    writeXML "<ITRForm:FullConsideration>" & UVCase(STCG_A2iia_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML "<ITRForm:DeductSec48>"
    writeXML "<ITRForm:AquisitCost>" & UVCase(STCG_A2iibi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML "<ITRForm:ImproveCost>" & UVCase(STCG_A2iibii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML "<ITRForm:ExpOnTrans>" & UVCase(STCG_A2iibiii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML "<ITRForm:TotalDedn>" & UVCase(STCG_A2iibiv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML "</ITRForm:DeductSec48>"
    writeXML "<ITRForm:BalanceCG>" & UVCase(STCG_A2iic_BalanceCG) & "</ITRForm:BalanceCG>"
    writeXML "<ITRForm:LossSec94of7Or94of8>" & UVCase(STCG_A2iid_LossSec94of7Or94of8) & "</ITRForm:LossSec94of7Or94of8>"
    writeXML "<ITRForm:CapgainonAssets>" & UVCase(STCG_A2iie_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML "</ITRForm:EquityMFonSTTDtls>"
    writeXML "</ITRForm:EquityMFonSTT>"
    
    writeXML " <ITRForm:NRITransacSec48Dtl>"
    writeXML " <ITRForm:NRItaxSTTPaid>" & UVCase(STCG_A3a_NRItaxSTTPaid) & "</ITRForm:NRItaxSTTPaid>"
    writeXML " <ITRForm:NRItaxSTTNotPaid>" & UVCase(STCG_A3b_NRItaxSTTNotPaid) & "</ITRForm:NRItaxSTTNotPaid>"
    writeXML " </ITRForm:NRITransacSec48Dtl>"
    
    writeXML " <ITRForm:NRISecur115AD>"
    writeXML " <ITRForm:FullValueConsdRecvUnqshr>" & UVCase(STCG_A4a_FullConsideration1) & "</ITRForm:FullValueConsdRecvUnqshr>"
    writeXML " <ITRForm:FairMrktValueUnqshr>" & UVCase(STCG_A4a_FullConsideration2) & "</ITRForm:FairMrktValueUnqshr>"
    writeXML " <ITRForm:FullValueConsdSec50CA>" & UVCase(STCG_A4a_FullConsideration3) & "</ITRForm:FullValueConsdSec50CA>"
    writeXML " <ITRForm:FullValueConsdOthUnqshr>" & UVCase(STCG_A4a_FullConsideration4) & "</ITRForm:FullValueConsdOthUnqshr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(STCG_A4a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(STCG_A4bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(STCG_A4bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(STCG_A4biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(STCG_A4biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(STCG_A4c_BalanceCG) & "</ITRForm:BalanceCG>"
    writeXML " <ITRForm:LossSec94of7Or94of8>" & UVCase(STCG_A4d_LossSec94of7Or94of8) & "</ITRForm:LossSec94of7Or94of8>"
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(STCG_A4e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:NRISecur115AD>"
    
    writeXML " <ITRForm:SaleOnOtherAssets>"
    writeXML " <ITRForm:FullValueConsdRecvUnqshr>" & UVCase(STCG_A5a_FullConsideration1) & "</ITRForm:FullValueConsdRecvUnqshr>"
    writeXML " <ITRForm:FairMrktValueUnqshr>" & UVCase(STCG_A5a_FullConsideration2) & "</ITRForm:FairMrktValueUnqshr>"
    writeXML " <ITRForm:FullValueConsdSec50CA>" & UVCase(STCG_A5a_FullConsideration3) & "</ITRForm:FullValueConsdSec50CA>"
    writeXML " <ITRForm:FullValueConsdOthUnqshr>" & UVCase(STCG_A5a_FullConsideration4) & "</ITRForm:FullValueConsdOthUnqshr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(STCG_A5a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(STCG_A5bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(STCG_A5bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(STCG_A5biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(STCG_A5biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(STCG_A5c_BalanceCG) & "</ITRForm:BalanceCG>"
    writeXML " <ITRForm:LossSec94of7Or94of8>" & UVCase(STCG_A5d_LossSec94of7Or94of8) & "</ITRForm:LossSec94of7Or94of8>"
    writeXML " <ITRForm:DeemedStcgOnAssets>" & UVCase(STCG_A8_DeemedStcgOnAssets) & "</ITRForm:DeemedStcgOnAssets>"

    writeXML " <ITRForm:ExemptionOrDednUs54>"
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54D</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(STCG_A5ei_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54G</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(STCG_A5eii_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54GA</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(STCG_A5eiii_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(STCG_A5e_ExemptionGrandTotal) & "</ITRForm:ExemptionGrandTotal>"
    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(STCG_A5e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:SaleOnOtherAssets>"
    
   ' Dim i As Variant
   If Not STCG_A6_StatusUnutilizedCG <> "" Or STCG_A6_StatusUnutilizedCG <> "(" Then
    writeXML "<ITRForm:UnutilizedStcgFlag>" & STCG_A6_StatusUnutilizedCG & "</ITRForm:UnutilizedStcgFlag>"
    End If
    
    If STCG_A6_StatusUnutilizedCG = "Y" Then
        writeXML "<ITRForm:UnutilizedCg>"
        noOfProcessSub = UBound(STCG_A7_PreviousYer_1)
        For i = 1 To UBound(STCG_A7_PreviousYer_1)
        If Not isdropdownblank(STCG_A7_DeductionSection_1(i)) Then
            writeXML "<ITRForm:UnutilizedCgPrvYrDtls>"
                If STCG_A7_PreviousYer_1(i) <> "" Then writeXML "<ITRForm:PrvYrInWhichAsstTrnsfrd>" & STCG_A7_PreviousYer_1(i) & "</ITRForm:PrvYrInWhichAsstTrnsfrd>"
                If STCG_A7_DeductionSection_1(i) <> "" Then writeXML "<ITRForm:SectionClmd>" & STCG_A7_DeductionSection_1(i) & "</ITRForm:SectionClmd>"
                If STCG_A7_AsstAcquiredYer_1(i) <> "" Then writeXML "<ITRForm:YrInWhichAssetAcq>" & STCG_A7_AsstAcquiredYer_1(i) & "</ITRForm:YrInWhichAssetAcq>"
                If STCG_A7_AmtUtilised_1(i) <> "" Then writeXML "<ITRForm:AmtUtilized>" & STCG_A7_AmtUtilised_1(i) & "</ITRForm:AmtUtilized>"
                If STCG_A7_AmtNotUsed_1(i) <> "" Then writeXML "<ITRForm:AmtUnutilized>" & STCG_A7_AmtNotUsed_1(i) & "</ITRForm:AmtUnutilized>"
            writeXML "</ITRForm:UnutilizedCgPrvYrDtls>"
            UpdateProgressBar
        End If
        Next
        
         writeXML "</ITRForm:UnutilizedCg>"
         
    
    End If
    writeXML "<ITRForm:AmtDeemedStcg>" & STCG_A6_AmountDeemedOth & "</ITRForm:AmtDeemedStcg>"
    writeXML "<ITRForm:TotalAmtDeemedStcg>" & STCG_A7_AmtDeemed & "</ITRForm:TotalAmtDeemedStcg>"
    
    If (Range("STCG.A8_PassThroughIncome").value) <> "" Then
                writeXML "<ITRForm:PassThrIncNatureSTCG>" & UCase(Range("STCG.A8_PassThroughIncome").value) & "</ITRForm:PassThrIncNatureSTCG>"
                Else
                writeXML "<ITRForm:PassThrIncNatureSTCG>0</ITRForm:PassThrIncNatureSTCG>"
                End If
            
    If (Range("STCG.A8_PassThroughIncome15").value) <> "" Then
                writeXML "<ITRForm:PassThrIncNatureSTCG15Per>" & UCase(Range("STCG.A8_PassThroughIncome15").value) & "</ITRForm:PassThrIncNatureSTCG15Per>"
                Else
                writeXML "<ITRForm:PassThrIncNatureSTCG15Per>0</ITRForm:PassThrIncNatureSTCG15Per>"
                End If
                
    If (Range("STCG.A8_PassThroughIncome30").value) <> "" Then
                writeXML "<ITRForm:PassThrIncNatureSTCG30Per>" & UCase(Range("STCG.A8_PassThroughIncome30").value) & "</ITRForm:PassThrIncNatureSTCG30Per>"
                Else
                writeXML "<ITRForm:PassThrIncNatureSTCG30Per>0</ITRForm:PassThrIncNatureSTCG30Per>"
                End If
                
   If (Range("STCG.A8_PassThroughIncome_ApplicableRate").value) <> "" Then
                writeXML "<ITRForm:PassThrIncNatureSTCGAppRate>" & UCase(Range("STCG.A8_PassThroughIncome_ApplicableRate").value) & "</ITRForm:PassThrIncNatureSTCGAppRate>"
                Else
                writeXML "<ITRForm:PassThrIncNatureSTCGAppRate>0</ITRForm:PassThrIncNatureSTCGAppRate>"
                End If

    
    If STCG_A8_StcgDTAA <> 0 Or STCG_A8b_StcgDTAA <> 0 Then
    writeXML "<ITRForm:NRICgDTAA>"
    noOfProcessSub = UBound(AmountSTCG_DTAA)
        For i = 1 To UBound(AmountSTCG_DTAA)
        If AmountSTCG_DTAA(i) <> "" Then
            writeXML " <ITRForm:NRIDTAADtls>"
            If AmountSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:DTAAamt>" & AmountSTCG_DTAA(i) & "</ITRForm:DTAAamt>"
            If SectionSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:ItemNoincl>" & SectionSTCG_DTAA(i) & "</ITRForm:ItemNoincl>"
            If CountryNameSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:CountryName>" & CountryNameSTCG_DTAA(i) & "</ITRForm:CountryName>"
            If CountryCodeSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:CountryCode>" & CountryCodeSTCG_DTAA(i) & "</ITRForm:CountryCode>"
            If ArticleSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:DTAAarticle>" & ArticleSTCG_DTAA(i) & "</ITRForm:DTAAarticle>"
            If RateTreatySTCG_DTAA(i) <> "" Then writeXML "<ITRForm:RateAsPerTreaty>" & Round(RateTreatySTCG_DTAA(i), 2) & "</ITRForm:RateAsPerTreaty>"
            If TRCSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:TaxRescertifiedFlag>" & Mid(TRCSTCG_DTAA(i), 1, 1) & "</ITRForm:TaxRescertifiedFlag>"
            If SectionITActSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:SecITAct>" & SectionITActSTCG_DTAA(i) & "</ITRForm:SecITAct>"
            If RateITActSTCG_DTAA(i) <> "" Then writeXML "<ITRForm:RateAsPerITAct>" & Round(RateITActSTCG_DTAA(i), 2) & "</ITRForm:RateAsPerITAct>"
            If ApplRate_STCG(i) <> "" Then writeXML "<ITRForm:ApplicableRate>" & Round(ApplRate_STCG(i), 2) & "</ITRForm:ApplicableRate>"
            
           writeXML " </ITRForm:NRIDTAADtls>"
           UpdateProgressBar
        End If
        Next
    writeXML "</ITRForm:NRICgDTAA>"
    End If
    
     writeXML "<ITRForm:TotalAmtNotTaxUsDTAAStcg>" & UVCase(STCG_A8_StcgDTAA) & "</ITRForm:TotalAmtNotTaxUsDTAAStcg>"
    writeXML "<ITRForm:TotalAmtTaxUsDTAAStcg>" & UVCase(STCG_A8b_StcgDTAA) & "</ITRForm:TotalAmtTaxUsDTAAStcg>"
    
    
    'writeXML " <ITRForm:DeemedStcgOnAssets>" & UVCase(STCG_A8_DeemedStcgOnAssets) & "</ITRForm:DeemedStcgOnAssets>"
    writeXML " <ITRForm:TotalSTCG>" & UVCase(STCG_A7_TotalSTCG) & "</ITRForm:TotalSTCG>"
    writeXML " </ITRForm:ShortTermCapGainFor23>"
    
            writeXML " <ITRForm:LongTermCapGain23>"
    writeXML " <ITRForm:SaleofLandBuild>"
    For num = 1 To LTCGA1Count
    writeXML " <ITRForm:SaleofLandBuildDtls>"
   
    writeXML " <ITRForm:FullConsideration>" & UVCase(Sheet13.Range("LTCG.Bai_FullConsideration" & num).value) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:PropertyValuation>" & UVCase(Sheet13.Range("LTCG.Baii_PropertyValuation" & num).value) & "</ITRForm:PropertyValuation>"
    writeXML " <ITRForm:FullConsideration50C>" & UVCase(Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & num).value) & "</ITRForm:FullConsideration50C>"
    
    writeXML " <ITRForm:AquisitCost>" & UVCase(Sheet13.Range("LTCG.Bbi_AquisitCost" & num).value) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(Sheet13.Range("LTCG.Bbii_ImproveCost" & num).value) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & num).value) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(Sheet13.Range("LTCG.Bbiv_TotalDedn" & num).value) & "</ITRForm:TotalDedn>"
   writeXML " <ITRForm:Balance>" & UVCase(Sheet13.Range("LTCG.Bc_BalanceCG" & num).value) & "</ITRForm:Balance>"

    writeXML " <ITRForm:ExemptionOrDednUs54>"
    If Sheet13.Range("LTCG.Bdi_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdi_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("LTCG.Bdii_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54B</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdii_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If

    If Sheet13.Range("LTCG.Bdiii_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54D</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdiii_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If

    If Sheet13.Range("LTCG.Bdiv_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54EC</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdiv_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
'    If Sheet13.Range("LTCG.Bdix_ExemptionAmount54EE" & num).value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54EE</ITRForm:ExemptionSecCode>"
'    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdix_ExemptionAmount54EE" & num).value) & "</ITRForm:ExemptionAmount>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
'    End If

    If Sheet13.Range("LTCG.Bdv_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdv_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If

    If Sheet13.Range("LTCG.Bdvi_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54G</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdvi_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("LTCG.Bdvii_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54GA</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdvii_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If
    
    If Sheet13.Range("LTCG.Bdviii_ExemptionAmount" & num).value <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54GB</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(Sheet13.Range("LTCG.Bdviii_ExemptionAmount" & num).value) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If

   writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.Bd_ExemptionGrandTotal" & num).value) & "</ITRForm:ExemptionGrandTotal>"
   writeXML " </ITRForm:ExemptionOrDednUs54>"
  
    writeXML "<ITRForm:CapgainonAssets>" & UVCase(Sheet13.Range("LTCG.Be_CapgainonAssets" & num).value) & "</ITRForm:CapgainonAssets>"
    
    LTCGImmovablePropertyXML (num)
    writeXML " </ITRForm:SaleofLandBuildDtls>"
    Next
    
    
    writeXML " </ITRForm:SaleofLandBuild>"
    
    writeXML " <ITRForm:SlumpSaleInLtcg>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_B02a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:NetWorthOfDivision>" & UVCase(LTCG_B02b_NetWorthOfDivision) & "</ITRForm:NetWorthOfDivision>"
    writeXML " <ITRForm:SlumpBalance>" & UVCase(LTCG_B02c_SlumpBalance) & "</ITRForm:SlumpBalance>"
    
    writeXML " <ITRForm:ExemptionOrDednUs54>"
    
    If LTCG_B02di_ExemptionAmount <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54EC</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(LTCG_B02di_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If
   
    If LTCG_B02dii_ExemptionAmount <> "" Then
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(LTCG_B02dii_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    End If
 
    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B02d_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_B02e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:SlumpSaleInLtcg>"
    
    writeXML " <ITRForm:SaleofBondsDebntr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_B2a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_B2bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_B2bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_B2biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_B2biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_B2c_BalanceCG) & "</ITRForm:BalanceCG>"
    
    
'    writeXML " <ITRForm:ExemptionOrDednUs54>"
    'If Sheet13.Range("LTCG.B03d_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B03d_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
   ' End If
 
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B03d_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_B2e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:SaleofBondsDebntr>"
    
   
    writeXML " <ITRForm:Proviso112Applicable>"
    writeXML " <ITRForm:Proviso112SectionCode>22</ITRForm:Proviso112SectionCode>"
    writeXML " <ITRForm:Proviso112Applicabledtls>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_B3a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_B3bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_B3bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_B3biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_B3biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_B3c_BalanceCG) & "</ITRForm:BalanceCG>"
   
   
'    writeXML " <ITRForm:ExemptionOrDednUs54>"

    'If Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    'End If
 
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
     
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_B3e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:Proviso112Applicabledtls>"
    writeXML " </ITRForm:Proviso112Applicable>"
    
    writeXML " <ITRForm:Proviso112Applicable>"
    writeXML " <ITRForm:Proviso112SectionCode>5ACA1b</ITRForm:Proviso112SectionCode>"
    writeXML " <ITRForm:Proviso112Applicabledtls>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_Bii3a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_Bii3bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_Bii3bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_Bii3biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_Bii3biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_Bii3c_BalanceCG) & "</ITRForm:BalanceCG>"
 
'     writeXML " <ITRForm:ExemptionOrDednUs54>"

    'If Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    'End If
 
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
 

    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_Bii3e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:Proviso112Applicabledtls>"
    writeXML " </ITRForm:Proviso112Applicable>"
    
    writeXML " <ITRForm:SaleOfEquityShareUs112A>"
    
    
  If Sheet13.Range("LTCG.5.LTCG_AfterThreshold").value = "" Then
        writeXML " <ITRForm:BalanceCG>0</ITRForm:BalanceCG>"
    Else
        writeXML " <ITRForm:BalanceCG>" & Round(UVCase(Sheet13.Range("LTCG.5.LTCG_AfterThreshold").value), 0) & "</ITRForm:BalanceCG>"
    End If
'writeXML " <ITRForm:ExemptionOrDednUs54>"
      
   ' If Sheet13.Range("B5_DeductionTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("B5_DeductionTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
   ' End If
 
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("B5_DeductionTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
  If Sheet13.Range("LTCG.5.LTCG_OnSaleCapitalAssets").value = "" Then
        writeXML " <ITRForm:CapgainonAssets>0</ITRForm:CapgainonAssets>"
    Else
        writeXML " <ITRForm:CapgainonAssets>" & Round(UVCase(Sheet13.Range("LTCG.5.LTCG_OnSaleCapitalAssets").value), 2) & "</ITRForm:CapgainonAssets>"
    End If
      
    writeXML " </ITRForm:SaleOfEquityShareUs112A>"
    writeXML " <ITRForm:NRIProvisoSec48>"
    writeXML " <ITRForm:LTCGWithoutBenefit>" & UVCase(LTCG_B4a_LTCGWithoutBenefit) & "</ITRForm:LTCGWithoutBenefit>"
    
'    writeXML " <ITRForm:ExemptionOrDednUs54>"
'    If Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
'    End If
 
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
   
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_B4c_BalanceCG) & "</ITRForm:BalanceCG>"
    'writeXML " <ITRForm:LTCGOnListedSecurity>" & UVCase(LTCG_B4d_LTCGOnListedSecurity) & "</ITRForm:LTCGOnListedSecurity>"
    'writeXML " <ITRForm:LTCGOnUnlistedSecurity>" & UVCase(LTCG_B4e_LTCGOnUnlistedSecurity) & "</ITRForm:LTCGOnUnlistedSecurity>"
    writeXML " </ITRForm:NRIProvisoSec48>"
    
    writeXML " <ITRForm:NRIOnSec112and115>"
    writeXML " <ITRForm:NRIOnSec112and115Dtls>"
    writeXML " <ITRForm:SectionCode>21ciii</ITRForm:SectionCode>"
    writeXML " <ITRForm:FullValueConsdRecvUnqshr>" & UVCase(LTCG_Bi5a_FullConsideration1) & "</ITRForm:FullValueConsdRecvUnqshr>"
    writeXML " <ITRForm:FairMrktValueUnqshr>" & UVCase(LTCG_Bi5a_FullConsideration2) & "</ITRForm:FairMrktValueUnqshr>"
    writeXML " <ITRForm:FullValueConsdSec50CA>" & UVCase(LTCG_Bi5a_FullConsideration3) & "</ITRForm:FullValueConsdSec50CA>"
    writeXML " <ITRForm:FullValueConsdOthUnqshr>" & UVCase(LTCG_Bi5a_FullConsideration4) & "</ITRForm:FullValueConsdOthUnqshr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_Bi5a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_Bi5bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_Bi5bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_Bi5biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_Bi5biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_Bi5c_BalanceCG) & "</ITRForm:BalanceCG>"
    
''   writeXML " <ITRForm:ExemptionOrDednUs54>"
   
'       If Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
''    End If
'     writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_Bi5e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:NRIOnSec112and115Dtls>"
    
    writeXML " <ITRForm:NRIOnSec112and115Dtls>"
    writeXML " <ITRForm:SectionCode>5AC1c</ITRForm:SectionCode>"
    
    writeXML " <ITRForm:FullValueConsdRecvUnqshr>" & UVCase(LTCG_Bii5a_FullConsideration1) & "</ITRForm:FullValueConsdRecvUnqshr>"
    writeXML " <ITRForm:FairMrktValueUnqshr>" & UVCase(LTCG_Bii5a_FullConsideration2) & "</ITRForm:FairMrktValueUnqshr>"
    writeXML " <ITRForm:FullValueConsdSec50CA>" & UVCase(LTCG_Bii5a_FullConsideration3) & "</ITRForm:FullValueConsdSec50CA>"
    writeXML " <ITRForm:FullValueConsdOthUnqshr>" & UVCase(LTCG_Bii5a_FullConsideration4) & "</ITRForm:FullValueConsdOthUnqshr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_Bii5a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_Bii5bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_Bii5bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_Bii5biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_Bii5biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_Bii5c_BalanceCG) & "</ITRForm:BalanceCG>"
    
'   writeXML " <ITRForm:ExemptionOrDednUs54>"
 
'    If Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
'    End If
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
   
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_Bii5e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:NRIOnSec112and115Dtls>"


    writeXML " <ITRForm:NRIOnSec112and115Dtls>"
    writeXML " <ITRForm:SectionCode>5ADiii</ITRForm:SectionCode>"
    
    writeXML " <ITRForm:FullValueConsdRecvUnqshr>" & UVCase(LTCG_Biii5a_FullConsideration1) & "</ITRForm:FullValueConsdRecvUnqshr>"
    writeXML " <ITRForm:FairMrktValueUnqshr>" & UVCase(LTCG_Biii5a_FullConsideration2) & "</ITRForm:FairMrktValueUnqshr>"
    writeXML " <ITRForm:FullValueConsdSec50CA>" & UVCase(LTCG_Biii5a_FullConsideration3) & "</ITRForm:FullValueConsdSec50CA>"
    writeXML " <ITRForm:FullValueConsdOthUnqshr>" & UVCase(LTCG_Biii5a_FullConsideration4) & "</ITRForm:FullValueConsdOthUnqshr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_Biii5a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_Biii5bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_Biii5bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_Biii5biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_Biii5biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_Biii5c_BalanceCG) & "</ITRForm:BalanceCG>"
    
'      writeXML " <ITRForm:ExemptionOrDednUs54>"
   
'    If Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
'    End If
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_Biii5e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:NRIOnSec112and115Dtls>"
    
    writeXML " </ITRForm:NRIOnSec112and115>"
    
    writeXML " <ITRForm:NRISaleOfEquityShareUs112A>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(Sheet13.Range("LTCG.8.LTCG_AfterThreshold").value) & "</ITRForm:BalanceCG>"
    
'writeXML " <ITRForm:ExemptionOrDednUs54>"
'    If Sheet13.Range("LTCG.B08h_ExemptionGrandTotal").value <> "" Then
'    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
'    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:DeductionUs54F>" & UVCase(Sheet13.Range("LTCG.8.DeductionUnder54F").value) & "</ITRForm:DeductionUs54F>"
'    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
'    End If
    
'    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(Sheet13.Range("LTCG.B08h_ExemptionGrandTotal").value) & "</ITRForm:ExemptionGrandTotal>"
'    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    
  If Sheet13.Range("LTCG.5.LTCG_AssetsB8Above").value = "" Then
        writeXML " <ITRForm:CapgainonAssets>0</ITRForm:CapgainonAssets>"
    Else
        writeXML " <ITRForm:CapgainonAssets>" & UVCase(Sheet13.Range("LTCG.5.LTCG_AssetsB8Above").value) & "</ITRForm:CapgainonAssets>"
    End If
   
    writeXML " </ITRForm:NRISaleOfEquityShareUs112A>"

    writeXML " <ITRForm:NRISaleofForeignAsset>"
    writeXML " <ITRForm:SaleonSpecAsset>" & UVCase(LTCG_B6a_SaleonSpecAsset) & "</ITRForm:SaleonSpecAsset>"
    writeXML " <ITRForm:DednSpecAssetus115>" & UVCase(LTCG_B6b_DednSpecAssetus115) & "</ITRForm:DednSpecAssetus115>"
    writeXML " <ITRForm:BalonSpeciAsset>" & UVCase(LTCG_B6c_BalonSpeciAsset) & "</ITRForm:BalonSpeciAsset>"
    writeXML " <ITRForm:SaleOtherSpecAsset>" & UVCase(LTCG_B6d_SaleOtherSpecAsset) & "</ITRForm:SaleOtherSpecAsset>"
    writeXML " <ITRForm:DednOtherSpecAssetus115>" & UVCase(LTCG_B6e_DednOtherSpecAssetus115) & "</ITRForm:DednOtherSpecAssetus115>"
    writeXML " <ITRForm:BalOtherthanSpecAsset>" & UVCase(LTCG_B6f_BalOtherthanSpecAsset) & "</ITRForm:BalOtherthanSpecAsset>"
    writeXML " </ITRForm:NRISaleofForeignAsset>"
    
    writeXML " <ITRForm:SaleofAssetNA>"
    writeXML " <ITRForm:FullValueConsdRecvUnqshr>" & UVCase(LTCG_B7a_FullConsideration1) & "</ITRForm:FullValueConsdRecvUnqshr>"
    writeXML " <ITRForm:FairMrktValueUnqshr>" & UVCase(LTCG_B7a_FullConsideration2) & "</ITRForm:FairMrktValueUnqshr>"
    writeXML " <ITRForm:FullValueConsdSec50CA>" & UVCase(LTCG_B7a_FullConsideration3) & "</ITRForm:FullValueConsdSec50CA>"
     writeXML " <ITRForm:FullValueConsdOthUnqshr>" & UVCase(LTCG_B7a_FullConsideration4) & "</ITRForm:FullValueConsdOthUnqshr>"
    writeXML " <ITRForm:FullConsideration>" & UVCase(LTCG_B7a_FullConsideration) & "</ITRForm:FullConsideration>"
    writeXML " <ITRForm:DeductSec48>"
    writeXML " <ITRForm:AquisitCost>" & UVCase(LTCG_B7bi_AquisitCost) & "</ITRForm:AquisitCost>"
    writeXML " <ITRForm:ImproveCost>" & UVCase(LTCG_B7bii_ImproveCost) & "</ITRForm:ImproveCost>"
    writeXML " <ITRForm:ExpOnTrans>" & UVCase(LTCG_B7biii_ExpOnTrans) & "</ITRForm:ExpOnTrans>"
    writeXML " <ITRForm:TotalDedn>" & UVCase(LTCG_B7biv_TotalDedn) & "</ITRForm:TotalDedn>"
    writeXML " </ITRForm:DeductSec48>"
    writeXML " <ITRForm:BalanceCG>" & UVCase(LTCG_B7c_BalanceCG) & "</ITRForm:BalanceCG>"
    
    writeXML " <ITRForm:ExemptionOrDednUs54>"
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54D</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(LTCG_B7di_ExemptionAmount0) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"

    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54F</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(LTCG_B7dii_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54G</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(LTCG_B7diii_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"
    
    writeXML " <ITRForm:ExemptionOrDednUs54Dtls>"
    writeXML " <ITRForm:ExemptionSecCode>54GA</ITRForm:ExemptionSecCode>"
    writeXML " <ITRForm:ExemptionAmount>" & UVCase(LTCG_B7div_ExemptionAmount) & "</ITRForm:ExemptionAmount>"
    writeXML " </ITRForm:ExemptionOrDednUs54Dtls>"


    writeXML " <ITRForm:ExemptionGrandTotal>" & UVCase(LTCG_B7d_ExemptionGrandTotal) & "</ITRForm:ExemptionGrandTotal>"
    writeXML " </ITRForm:ExemptionOrDednUs54>"
    
    writeXML " <ITRForm:CapgainonAssets>" & UVCase(LTCG_B7e_CapgainonAssets) & "</ITRForm:CapgainonAssets>"
    writeXML " </ITRForm:SaleofAssetNA>"
    
    'Dim i As Variant
    If Not LTCG_B9_StatusUnutilizedCG <> "" Or LTCG_B9_StatusUnutilizedCG <> "(" Then
    writeXML "<ITRForm:UnutilizedLtcgFlag>" & LTCG_B9_StatusUnutilizedCG & "</ITRForm:UnutilizedLtcgFlag>"
    End If
    
    If LTCG_B9_StatusUnutilizedCG = "Y" Then
        writeXML "<ITRForm:UnutilizedCg>"
        For i = 1 To UBound(LTCG_B9_PreviousYer_1)
        If Not isdropdownblank(LTCG_B9_DeductionSection_1(i)) Then
            writeXML "<ITRForm:UnutilizedCgPrvYrDtls>"
                If LTCG_B9_PreviousYer_1(i) <> "" Then writeXML "<ITRForm:PrvYrInWhichAsstTrnsfrd>" & LTCG_B9_PreviousYer_1(i) & "</ITRForm:PrvYrInWhichAsstTrnsfrd>"
                If LTCG_B9_DeductionSection_1(i) <> "" Then writeXML "<ITRForm:SectionClmd>" & LTCG_B9_DeductionSection_1(i) & "</ITRForm:SectionClmd>"
                If LTCG_B9_AsstAcquiredYer_1(i) <> "" Then writeXML "<ITRForm:YrInWhichAssetAcq>" & LTCG_B9_AsstAcquiredYer_1(i) & "</ITRForm:YrInWhichAssetAcq>"
                If LTCG_B9_AmtUtilised_1(i) <> "" Then writeXML "<ITRForm:AmtUtilized>" & LTCG_B9_AmtUtilised_1(i) & "</ITRForm:AmtUtilized>"
                If LTCG_B9_AmtNotUsed_1(i) <> "" Then writeXML "<ITRForm:AmtUnutilized>" & LTCG_B9_AmtNotUsed_1(i) & "</ITRForm:AmtUnutilized>"
            writeXML "</ITRForm:UnutilizedCgPrvYrDtls>"
        End If
        Next
        writeXML "</ITRForm:UnutilizedCg>"

    End If
    writeXML "<ITRForm:AmtDeemedLtcg>" & LTCG_B9_AmountDeemedOth & "</ITRForm:AmtDeemedLtcg>"
    writeXML "<ITRForm:TotalAmtDeemedLtcg>" & LTCG_B9_AmtDeemed & "</ITRForm:TotalAmtDeemedLtcg>"
    
     If Sheet13.Range("LTCG.B12_PassThroughIncome").value = "" Then
        writeXML " <ITRForm:PassThrIncNatureLTCG>0</ITRForm:PassThrIncNatureLTCG>"
    Else
        writeXML " <ITRForm:PassThrIncNatureLTCG>" & UVCase(Sheet13.Range("LTCG.B12_PassThroughIncome").value) & "</ITRForm:PassThrIncNatureLTCG>"
    End If
    
    If Sheet13.Range("LTCG.B12_PassThroughIncome10_112A").value = "" Then
        writeXML " <ITRForm:PassThrIncNatureLTCGUs112A>0</ITRForm:PassThrIncNatureLTCGUs112A>"
    Else
        writeXML " <ITRForm:PassThrIncNatureLTCGUs112A>" & UVCase(Sheet13.Range("LTCG.B12_PassThroughIncome10_112A").value) & "</ITRForm:PassThrIncNatureLTCGUs112A>"
    End If
    
    If Sheet13.Range("LTCG.B12_PassThroughIncome10").value = "" Then
        writeXML " <ITRForm:PassThrIncNatureLTCG10Per>0</ITRForm:PassThrIncNatureLTCG10Per>"
    Else
        writeXML " <ITRForm:PassThrIncNatureLTCG10Per>" & UVCase(Sheet13.Range("LTCG.B12_PassThroughIncome10").value) & "</ITRForm:PassThrIncNatureLTCG10Per>"
    End If
    
    If Sheet13.Range("LTCG.B12_PassThroughIncome20").value = "" Then
        writeXML " <ITRForm:PassThrIncNatureLTCG20Per>0</ITRForm:PassThrIncNatureLTCG20Per>"
    Else
        writeXML " <ITRForm:PassThrIncNatureLTCG20Per>" & UVCase(Sheet13.Range("LTCG.B12_PassThroughIncome20").value) & "</ITRForm:PassThrIncNatureLTCG20Per>"
    End If
    
    
    
    
    If LTCG_B10_ltcgDTAA <> 0 Or LTCG_B10b_ltcgDTAA <> 0 Then
    writeXML "<ITRForm:NRICgDTAA>"
        For i = 1 To UBound(AmountLTCG_DTAA)
        If AmountLTCG_DTAA(i) <> "" Then
            writeXML " <ITRForm:NRIDTAADtls>"

            If AmountLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:DTAAamt>" & AmountLTCG_DTAA(i) & "</ITRForm:DTAAamt>"
            If SectionLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:ItemNoincl>" & SectionLTCG_DTAA(i) & "</ITRForm:ItemNoincl>"
            If CountryNameLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:CountryName>" & CountryNameLTCG_DTAA(i) & "</ITRForm:CountryName>"
            If CountryCodeLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:CountryCode>" & CountryCodeLTCG_DTAA(i) & "</ITRForm:CountryCode>"
            If ArticleLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:DTAAarticle>" & ArticleLTCG_DTAA(i) & "</ITRForm:DTAAarticle>"
            If RateTreatyLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:RateAsPerTreaty>" & Round(RateTreatyLTCG_DTAA(i), 2) & "</ITRForm:RateAsPerTreaty>"
            If TRCLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:TaxRescertifiedFlag>" & Mid(TRCLTCG_DTAA(i), 1, 1) & "</ITRForm:TaxRescertifiedFlag>"
            If SectionITActLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:SecITAct>" & SectionITActLTCG_DTAA(i) & "</ITRForm:SecITAct>"
            If RateITActLTCG_DTAA(i) <> "" Then writeXML "<ITRForm:RateAsPerITAct>" & Round(RateITActLTCG_DTAA(i), 2) & "</ITRForm:RateAsPerITAct>"
            If Applrate_LTCG(i) <> "" Then writeXML "<ITRForm:ApplicableRate>" & Round(Applrate_LTCG(i), 2) & "</ITRForm:ApplicableRate>"
            
            
           writeXML " </ITRForm:NRIDTAADtls>"
        End If
        Next
    writeXML "</ITRForm:NRICgDTAA>"
    End If
    writeXML "<ITRForm:TotalAmtNotTaxUsDTAALtcg>" & UVCase(LTCG_B10_ltcgDTAA) & "</ITRForm:TotalAmtNotTaxUsDTAALtcg>"
     writeXML "<ITRForm:TotalAmtTaxUsDTAALtcg>" & UVCase(LTCG_B10b_ltcgDTAA) & "</ITRForm:TotalAmtTaxUsDTAALtcg>"
    writeXML " <ITRForm:TotalLTCG>" & UVCase(LTCG_B9_TotalLTCG) & "</ITRForm:TotalLTCG>"
    writeXML " </ITRForm:LongTermCapGain23>"
    
     writeXML " <ITRForm:TotScheduleCGFor23>" & UVCase(CG_C_TotScheduleCGFor23) & "</ITRForm:TotScheduleCGFor23>"
    
    
    
    
    writeXML "      <ITRForm:DeducClaimInfo>"
    


      If Not IsEmpty(DateOfTransfer) Then
         If (UBound(DateOfTransfer)) > 0 Then
        For i = 1 To UBound(DateOfTransfer)

        writeXML "<ITRForm:DeducClaimDtlsUs54>"
        'For i = 1 To UBound(DateOfTransfer)
        If DateOfTransfer(i) <> "" Then
        writeXML " <ITRForm:DateofTransfer>" & Dformat(DateOfTransfer(i), "") & "</ITRForm:DateofTransfer>"
        End If
        
        If CostOfHouse(i) <> "" Then
        writeXML " <ITRForm:CostofNewResHouse>" & UVCase(CostOfHouse(i)) & "</ITRForm:CostofNewResHouse>"
        End If
        
        If (DateOfPurchase(i) <> "") Then
        writeXML "       <ITRForm:DateofPurchase>" & Dformat(DateOfPurchase(i), "") & "</ITRForm:DateofPurchase>"
        End If
        
        If (CashDeposited(i) <> "") Then
        writeXML "  <ITRForm:AmtDeposited>" & UVCase(CashDeposited(i)) & "</ITRForm:AmtDeposited>"
        End If
        
        If (AmountClaimed(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimed(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54>"
        Next
        End If
        End If
        
        
        If Not IsEmpty(DateOfTransferB) Then
        If (UBound(DateOfTransferB)) > 0 Then
        For i = 1 To UBound(DateOfTransferB)

        writeXML "<ITRForm:DeducClaimDtlsUs54B>"
        
        If DateOfTransferB(i) <> "" Then
        writeXML " <ITRForm:DateofTransfer>" & Dformat(DateOfTransferB(i), "") & "</ITRForm:DateofTransfer>"
        End If
        
        If CostOfLandB(i) <> "" Then
        writeXML " <ITRForm:CostofNewAgriLand>" & UVCase(CostOfLandB(i)) & "</ITRForm:CostofNewAgriLand>"
        End If
        
        If (DateOfPurchaseB(i) <> "") Then
        writeXML "       <ITRForm:DateofPurchase>" & Dformat(DateOfPurchaseB(i), "") & "</ITRForm:DateofPurchase>"
        End If
        
        If (CashDepositedB(i) <> "") Then
        writeXML "  <ITRForm:AmtDeposited>" & UVCase(CashDepositedB(i)) & "</ITRForm:AmtDeposited>"
        End If
        
        If (AmountClaimedB(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimedB(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54B>"
        Next
        End If
        End If
        
        
        If Not IsEmpty(DateOfTransferD) Then
        If (UBound(DateOfTransferD)) > 0 Then
        For i = 1 To UBound(DateOfTransferD)
        
        writeXML "<ITRForm:DeducClaimDtlsUs54D>"
        
        If (DateOfTransferD(i) <> "") Then
        writeXML " <ITRForm:DateofAcquisition>" & Dformat(DateOfTransferD(i), "") & "</ITRForm:DateofAcquisition>"
        End If
        
        If (CostOfLandD(i) <> "") Then
        writeXML " <ITRForm:CostofNewLandBuilding>" & UVCase(CostOfLandD(i)) & "</ITRForm:CostofNewLandBuilding>"
        End If
        
        If (DateOfPurchaseD(i) <> "") Then
        writeXML "       <ITRForm:DateofPurchase>" & Dformat(DateOfPurchaseD(i), "") & "</ITRForm:DateofPurchase>"
        End If
        
        If (CashDepositedD(i) <> "") Then
        writeXML "  <ITRForm:AmtDeposited>" & UVCase(CashDepositedD(i)) & "</ITRForm:AmtDeposited>"
        End If
        
        If (AmountClaimedD(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimedD(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54D>"
        
        Next
        End If
        End If


        If Not IsEmpty(DateOfTransferE) Then
        If (UBound(DateOfTransferE)) > 0 Then
        For i = 1 To UBound(DateOfTransferE)
        writeXML "<ITRForm:DeducClaimDtlsUs54EC>"
        
        If (DateOfTransferE(i) <> "") Then
        writeXML " <ITRForm:DateofTransfer>" & Dformat(DateOfTransferE(i), "") & "</ITRForm:DateofTransfer>"
        End If
        
        If (AmountInvestedE(i) <> "") Then
        writeXML " <ITRForm:AmtInvested>" & UVCase(AmountInvestedE(i)) & "</ITRForm:AmtInvested>"
        End If
        
        If (DateOfPurchaseE(i) <> "") Then
        writeXML "       <ITRForm:DateofInvestment>" & Dformat(DateOfPurchaseE(i), "") & "</ITRForm:DateofInvestment>"
        End If
        
        If (AmountClaimedE(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimedE(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54EC>"
        Next
        End If
        End If


        
        If Not IsEmpty(DateOfTransferF) Then
        If (UBound(DateOfTransferF)) > 0 Then
        For i = 1 To UBound(DateOfTransferF)
        
        writeXML "<ITRForm:DeducClaimDtlsUs54F>"
        
              If (DateOfTransferF(i) <> "") Then
        writeXML " <ITRForm:DateofTransfer>" & Dformat(DateOfTransferF(i), "") & "</ITRForm:DateofTransfer>"
        End If
        
        If (CostOfHouseF(i) <> "") Then
        writeXML " <ITRForm:CostofNewResHouse>" & UVCase(CostOfHouseF(i)) & "</ITRForm:CostofNewResHouse>"
        End If
        
        If (DateOfPurchaseF(i) <> "") Then
        writeXML "       <ITRForm:DateofPurchase>" & Dformat(DateOfPurchaseF(i), "") & "</ITRForm:DateofPurchase>"
        End If
        
        If (AmountDepositedF(i) <> "") Then
        writeXML "  <ITRForm:AmtDeposited>" & UVCase(AmountDepositedF(i)) & "</ITRForm:AmtDeposited>"
        End If
        
        If (AmountClaimedF(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimedF(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54F>"
        Next
        End If
        End If



        If Not IsEmpty(DateOfTransferG) Then
        If (UBound(DateOfTransferG)) > 0 Then
        For i = 1 To UBound(DateOfTransferG)
        
        writeXML "<ITRForm:DeducClaimDtlsUs54G>"

        If (DateOfTransferG(i) <> "") Then
        writeXML " <ITRForm:DateofTransfer>" & Dformat(DateOfTransferG(i), "") & "</ITRForm:DateofTransfer>"
        End If
        
        If (CostOfAssetG(i) <> "") Then
        writeXML " <ITRForm:CostofNewAsset>" & UVCase(CostOfAssetG(i)) & "</ITRForm:CostofNewAsset>"
        End If
        
        If (DateOfPurchaseG(i) <> "") Then
        writeXML "       <ITRForm:DateofPurchase>" & Dformat(DateOfPurchaseG(i), "") & "</ITRForm:DateofPurchase>"
        End If
        
        If (CashDepositedG(i) <> "") Then
        writeXML "  <ITRForm:AmtDeposited>" & UVCase(CashDepositedG(i)) & "</ITRForm:AmtDeposited>"
        End If
        
        If (AmountClaimedG(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimedG(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54G>"
        Next
        End If
        End If
        
        
         If Not IsEmpty(DateOfTransferGA) Then
        If (UBound(DateOfTransferGA)) > 0 Then
        For i = 1 To UBound(DateOfTransferGA)

        writeXML "<ITRForm:DeducClaimDtlsUs54GA>"
       
       If (DateOfTransferGA(i) <> "") Then
        writeXML " <ITRForm:DateofTransfer>" & Dformat(DateOfTransferGA(i), "") & "</ITRForm:DateofTransfer>"
        End If
        
        If (CostOfAssetGA(i) <> "") Then
        writeXML " <ITRForm:CostofNewAsset>" & UVCase(CostOfAssetGA(i)) & "</ITRForm:CostofNewAsset>"
        End If
        
        
        If (DateOfPurchaseGA(i) <> "") Then
        writeXML "       <ITRForm:DateofPurchase>" & Dformat(DateOfPurchaseGA(i), "") & "</ITRForm:DateofPurchase>"
        End If
        
        If (CashDepositedGA(i) <> "") Then
        writeXML "  <ITRForm:AmtDeposited>" & UVCase(CashDepositedGA(i)) & "</ITRForm:AmtDeposited>"
        End If
        
        If (AmountClaimedGA(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimedGA(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs54GA>"
        
        Next
        End If
        End If


        If Not IsEmpty(DateOfTransfer5F) Then
        If (UBound(DateOfTransfer5F)) > 0 Then
        For i = 1 To UBound(DateOfTransfer5F)
        writeXML "<ITRForm:DeducClaimDtlsUs115F>"
        
        DateOfTransfer5F(i) = Dformat(DateOfTransfer5F(i), "")
        
        If (DateOfTransfer5F(i) <> "") Then
        writeXML " <ITRForm:DateofTransfer>" & UVCase(DateOfTransfer5F(i)) & "</ITRForm:DateofTransfer>"
        End If
        
        If (AmountInvested5F(i) <> "") Then
        writeXML " <ITRForm:AmtInvested>" & UVCase(AmountInvested5F(i)) & "</ITRForm:AmtInvested>"
        End If
        
        If (DateOfPurchase5F(i) <> "") Then
        writeXML "       <ITRForm:DateofInvestment>" & Dformat(DateOfPurchase5F(i), "") & "</ITRForm:DateofInvestment>"
        End If
        
        If (AmountClaimed5F(i) <> "") Then
        writeXML "  <ITRForm:AmtDeducted>" & UVCase(AmountClaimed5F(i)) & "</ITRForm:AmtDeducted>"
        End If
        
        writeXML "</ITRForm:DeducClaimDtlsUs115F>"
        
        Next
        End If
        End If

    writeXML "<ITRForm:TotDeductClaim>" & UVCase(DEDN_1c_TotDeductClaim) & "</ITRForm:TotDeductClaim>"
    
    
    
    If (DEDN_D2_DeductPAN <> "") Then
        writeXML "<ITRForm:DeductPAN>" & UCase(DEDN_D2_DeductPAN) & "</ITRForm:DeductPAN>"
    End If
    
    writeXML "      </ITRForm:DeducClaimInfo>"
'    End If

    writeXML " <ITRForm:CurrYrLosses>"
    writeXML " <ITRForm:InLossSetOff>"
    writeXML " <ITRForm:TypeOfCapGain>LOSS_TO_BE_SET_OFF</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Ei2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Ei3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Ei4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Ei5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:LtclSetOff10Per>" & UVCase(IHLA_Ei5_LtclSetOff10Per) & "</ITRForm:LtclSetOff10Per>"
    writeXML " <ITRForm:LtclSetOff20Per>" & UVCase(IHLA_Ei6_LtclSetOff20Per) & "</ITRForm:LtclSetOff20Per>"
    writeXML " <ITRForm:LtclSetOffDTAARate>" & UVCase(Sheet13.Range("IHLA.Ei8_LtclSetOffDTAARate").value) & "</ITRForm:LtclSetOffDTAARate>"
    writeXML " </ITRForm:InLossSetOff>"
    
    writeXML " <ITRForm:InStcg15Per>"
    writeXML " <ITRForm:TypeOfCapGain>STCG_PER_15</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(IHLA_Eii1_CurrYearIncome) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Eii3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Eii4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(IHLA_Eii7_CurrYrCapGain) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InStcg15Per>"
    
    writeXML " <ITRForm:InStcg30Per>"
    writeXML " <ITRForm:TypeOfCapGain>STCG_PER_30</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(IHLA_Eiii1_CurrYearIncome) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Eiii2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Eiii4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(IHLA_Eiii7_CurrYrCapGain) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InStcg30Per>"
    
    writeXML " <ITRForm:InStcgAppRate>"
    writeXML " <ITRForm:TypeOfCapGain>STCG_PER_APP_RATE</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(IHLA_Eiv1_CurrYearIncome) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Eiv2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Eiv3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(IHLA_Eiv7_CurrYrCapGain) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InStcgAppRate>"
    
        writeXML " <ITRForm:InStcgDTAARate>"
    writeXML " <ITRForm:TypeOfCapGain>STCG_PER_DTAA_RATE</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(Sheet13.Range("IHLA.Ev_DTAArates").value) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(Sheet13.Range("IHLA.Ev2_DTAArates").value) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(Sheet13.Range("IHLA.Ev3_DTAArates").value) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(Sheet13.Range("IHLA.Ev4_DTAArates").value) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(Sheet13.Range("IHLA.Ev7_DTAArates").value) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InStcgDTAARate>"
    
    writeXML " <ITRForm:InLtcg10Per>"
    writeXML " <ITRForm:TypeOfCapGain>LTCG_PER_10</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(IHLA_Ev1_CurrYearIncome) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Ev2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Ev3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Ev4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:LtclSetOff20Per>" & UVCase(IHLA_Ev6_LtclSetOff20Per) & "</ITRForm:LtclSetOff20Per>"
    writeXML " <ITRForm:LtclSetOffDTAARate>" & UVCase(Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").value) & "</ITRForm:LtclSetOffDTAARate>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(IHLA_Ev7_CurrYrCapGain) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InLtcg10Per>"
    
    writeXML " <ITRForm:InLtcg20Per>"
    writeXML " <ITRForm:TypeOfCapGain>LTCG_PER_20</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(IHLA_Evi1_CurrYearIncome) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Evi2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Evi3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Evi4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:LtclSetOff10Per>" & UVCase(IHLA_Evi5_LtclSetOff10Per) & "</ITRForm:LtclSetOff10Per>"
    writeXML " <ITRForm:LtclSetOffDTAARate>" & UVCase(Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").value) & "</ITRForm:LtclSetOffDTAARate>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(IHLA_Evi7_CurrYrCapGain) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InLtcg20Per>"
    
       writeXML " <ITRForm:InLtcgDTAARate>"
    writeXML " <ITRForm:TypeOfCapGain>LTCG_PER_DTAA_RATE</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:CurrYearIncome>" & UVCase(Sheet13.Range("IHLA.E_LTCGv_DTAArates").value) & "</ITRForm:CurrYearIncome>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(Sheet13.Range("IHLA.E_LTCGv15_DTAArates").value) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(Sheet13.Range("IHLA.E_LTCGv30_DTAArates").value) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").value) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:LtclSetOff10Per>" & UVCase(Sheet13.Range("IHLA.E_LTCGv5_DTAArates").value) & "</ITRForm:LtclSetOff10Per>"
    writeXML " <ITRForm:LtclSetOff20Per>" & UVCase(Sheet13.Range("IHLA.E_LTCGv6_DTAArates").value) & "</ITRForm:LtclSetOff20Per>"
    writeXML " <ITRForm:CurrYrCapGain>" & UVCase(Sheet13.Range("IHLA.E_LTCGv7_DTAArates").value) & "</ITRForm:CurrYrCapGain>"
    writeXML " </ITRForm:InLtcgDTAARate>"
    
    writeXML " <ITRForm:TotLossSetOff>"
    writeXML " <ITRForm:TypeOfCapGain>TOT_LOSS_SET_OFF</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Evii2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Evii3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Evii4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Evii5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:LtclSetOff10Per>" & UVCase(IHLA_Evii5_LtclSetOff10Per) & "</ITRForm:LtclSetOff10Per>"
    writeXML " <ITRForm:LtclSetOff20Per>" & UVCase(IHLA_Evii6_LtclSetOff20Per) & "</ITRForm:LtclSetOff20Per>"
    writeXML " <ITRForm:LtclSetOffDTAARate>" & UVCase(Sheet13.Range("IHLA.Evii8_LtclSetOffDTAARate").value) & "</ITRForm:LtclSetOffDTAARate>"
    writeXML " </ITRForm:TotLossSetOff>"
    
    writeXML " <ITRForm:LossRemainSetOff>"
    writeXML " <ITRForm:TypeOfCapGain>LOSS_REMAIN_SET_OFF</ITRForm:TypeOfCapGain>"
    writeXML " <ITRForm:StclSetoff15Per>" & UVCase(IHLA_Eviii2_StclSetoff15Per) & "</ITRForm:StclSetoff15Per>"
    writeXML " <ITRForm:StclSetoff30Per>" & UVCase(IHLA_Eviii3_StclSetoff30Per) & "</ITRForm:StclSetoff30Per>"
    writeXML " <ITRForm:StclSetoffAppRate>" & UVCase(IHLA_Eviii4_StclSetoffAppRate) & "</ITRForm:StclSetoffAppRate>"
    writeXML " <ITRForm:StclSetoffDTAARate>" & UVCase(Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate").value) & "</ITRForm:StclSetoffDTAARate>"
    writeXML " <ITRForm:LtclSetOff10Per>" & UVCase(IHLA_Eviii5_LtclSetOff10Per) & "</ITRForm:LtclSetOff10Per>"
    writeXML " <ITRForm:LtclSetOff20Per>" & UVCase(IHLA_Eviii6_LtclSetOff20Per) & "</ITRForm:LtclSetOff20Per>"
    writeXML " <ITRForm:LtclSetOffDTAARate>" & UVCase(Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate").value) & "</ITRForm:LtclSetOffDTAARate>"
    writeXML " </ITRForm:LossRemainSetOff>"
    writeXML " </ITRForm:CurrYrLosses>"

    writeXML " <ITRForm:AccruOrRecOfCG>"
    writeXML " <ITRForm:ShortTermUnder15Per>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(AccSTCG_Upto15Of6) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(AccSTCG_Upto15Of9) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(ACCSTCG_Up16Of9To15Of12) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(ACCSTCG_Up16Of12To15Of3) & "</ITRForm:Up16Of12To15Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(ACCSTCG_Up16Of3To31Of3) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:ShortTermUnder15Per>"
    
    writeXML " <ITRForm:ShortTermUnder30Per>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(AccSTCG30_Upto15Of6) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(AccSTCG30_Upto15Of9) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(AccSTCG30_Up16Of9To15Of12) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(AccSTCG30_Up16Of12To15Of3) & "</ITRForm:Up16Of12To15Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(AccSTCG30_Up16Of3To31Of3) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:ShortTermUnder30Per>"
    
    writeXML " <ITRForm:ShortTermUnderAppRate>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(ACCSTCGOTH_Upto15Of6) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(ACCSTCGOTH_Upto15Of9) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(ACCSTCGOTH_Up16Of9To15Of12) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(ACCSTCGOTH_Up16Of12To15Of3) & "</ITRForm:Up16Of12To15Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(ACCSTCGOTH_Up16Of3To31Of3) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:ShortTermUnderAppRate>"
    
    writeXML " <ITRForm:ShortTermUnderDTAARate>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(Sheet13.Range("AccSTCGDTAA.Upto15Of6").value) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(Sheet13.Range("AccSTCGDTAA.Upto15Of9").value) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(Sheet13.Range("AccSTCGDTAA.Up16Of9To15Of12").value) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(Sheet13.Range("AccSTCGDTAA.Up16Of12To15Of3").value) & "</ITRForm:Up16Of12To15Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(Sheet13.Range("AccSTCGDTAA.Up16Of3To31Of3").value) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:ShortTermUnderDTAARate>"
    
    writeXML " <ITRForm:LongTermUnder10Per>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(ACCLTCG_Upto15Of6) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(ACCLTCG_Upto15Of9) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(ACCLTCG_Up16Of9To15Of12) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(ACCLTCG_Up16Of12To15Of3) & "</ITRForm:Up16Of12To15Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(ACCLTCG_Up16Of3To31Of3) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:LongTermUnder10Per>"
    
    writeXML " <ITRForm:LongTermUnder20Per>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(ACCLTCGNP_Upto15Of6) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(ACCLTCGNP_Upto15Of9) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(ACCLTCGNP_Up16Of9To15Of12) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(ACCLTCGNP_Up16Of12To15Of3) & "</ITRForm:Up16Of12To15Of3>"
    'writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(ACCLTCGNP_Up16Of3To31Of3) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(ACCLTCGNP_Up16Of3To31Of3) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:LongTermUnder20Per>"
    
        writeXML " <ITRForm:LongTermUnderDTAARate>"
    writeXML " <ITRForm:DateRange>"
    writeXML " <ITRForm:Upto15Of6>" & UVCase(Sheet13.Range("AccLTCGDTAA.Upto15Of6").value) & "</ITRForm:Upto15Of6>"
    writeXML " <ITRForm:Upto15Of9>" & UVCase(Sheet13.Range("AccLTCGDTAA.Upto15Of9").value) & "</ITRForm:Upto15Of9>"
    writeXML " <ITRForm:Up16Of9To15Of12>" & UVCase(Sheet13.Range("AccLTCGDTAA.Up16Of9To15Of12").value) & "</ITRForm:Up16Of9To15Of12>"
    writeXML " <ITRForm:Up16Of12To15Of3>" & UVCase(Sheet13.Range("AccLTCGDTAA.Up16Of12To15Of3").value) & "</ITRForm:Up16Of12To15Of3>"
    writeXML " <ITRForm:Up16Of3To31Of3>" & UVCase(Sheet13.Range("AccLTCGDTAA.Up16Of3To31Of3").value) & "</ITRForm:Up16Of3To31Of3>"
    writeXML " </ITRForm:DateRange>"
    writeXML " </ITRForm:LongTermUnderDTAARate>"
    
    writeXML " </ITRForm:AccruOrRecOfCG>"
    writeXML "</ITRForm:ScheduleCGFor23>"

End Function

Function Schedule112A()

Dim i As Long
If Not IsEmpty(ISINCode_112A) And UBound(ISINCode_112A) > 0 Then

writeXML "<ITRForm:Schedule112A>"

subProcCaption = "Schedule 112A"
noOfProcessSub = UBound(ISINCode_112A)

For i = 1 To UBound(ISINCode_112A)

writeXML "<ITRForm:Schedule112ADtls>"

If ShareAcq_112A(i) <> "" Then
If Trim(ShareAcq_112A(i)) = "On or before 31st January 2018" Then
writeXML "<ITRForm:ShareOnOrBefore>" & "BE" & "</ITRForm:ShareOnOrBefore>"
ElseIf Trim(ShareAcq_112A(i)) = "After 31st January 2018" Then
writeXML "<ITRForm:ShareOnOrBefore>" & "AE" & "</ITRForm:ShareOnOrBefore>"
End If
End If

If ISINCode_112A(i) <> "" Then
writeXML "<ITRForm:ISINCode>" & UCase(ISINCode_112A(i)) & "</ITRForm:ISINCode>"
End If

If NameofShare_112A(i) <> "" Then
writeXML "<ITRForm:ShareUnitName>" & UCase(NameofShare_112A(i)) & "</ITRForm:ShareUnitName>"
End If

If NoofShare_112A(i) <> "" Then
writeXML "<ITRForm:NumSharesUnits>" & Application.WorksheetFunction.Round(NoofShare_112A(i), 4) & "</ITRForm:NumSharesUnits>"
End If

If SalePriceperShare_112A(i) <> "" Then
writeXML "<ITRForm:SalePricePerShareUnit>" & Application.WorksheetFunction.Round(SalePriceperShare_112A(i), 4) & "</ITRForm:SalePricePerShareUnit>"
End If

If TotalSaleValue_112A(i) <> "" Then
writeXML "<ITRForm:TotSaleValue>" & Application.WorksheetFunction.Round(TotalSaleValue_112A(i), 4) & "</ITRForm:TotSaleValue>"
End If

If COAwithoutIndex_112A(i) <> "" Then
writeXML "<ITRForm:CostAcqWithoutIndx>" & Application.WorksheetFunction.Round(COAwithoutIndex_112A(i), 4) & "</ITRForm:CostAcqWithoutIndx>"
End If

If COAwithIndex_112A(i) <> "" Then
writeXML "<ITRForm:AcquisitionCost>" & Application.WorksheetFunction.Round(COAwithIndex_112A(i), 4) & "</ITRForm:AcquisitionCost>"
End If

If LTCGAssetAcquired_112A(i) <> "" Then
writeXML "<ITRForm:LTCGBeforelower6and11>" & Application.WorksheetFunction.Round(LTCGAssetAcquired_112A(i), 4) & "</ITRForm:LTCGBeforelower6and11>"
End If

If FairMarketValue_112A(i) <> "" Then
writeXML "<ITRForm:FairMktValuePerShareunit>" & Application.WorksheetFunction.Round(FairMarketValue_112A(i), 4) & "</ITRForm:FairMktValuePerShareunit>"
End If

If TotalFairMarketValue_112A(i) <> "" Then
writeXML "<ITRForm:TotFairMktValueCapAst>" & Application.WorksheetFunction.Round(TotalFairMarketValue_112A(i), 4) & "</ITRForm:TotFairMktValueCapAst>"
End If

'If FullValueConsideration_112A(i) <> "" Then
'writeXML "<ITRForm:FullValueConsdr>" & Round(FullValueConsideration_112A(i), 2) & "</ITRForm:FullValueConsdr>"
'End If

'If CostofImprovement_112A(i) <> "" Then
'writeXML "<ITRForm:CostImprvWithoutIndx>" & Round(CostofImprovement_112A(i), 2) & "</ITRForm:CostImprvWithoutIndx>"
'End If

If ExpenditureWholly_112A(i) <> "" Then
writeXML "<ITRForm:ExpExclCnctTransfer>" & Application.WorksheetFunction.Round(ExpenditureWholly_112A(i), 4) & "</ITRForm:ExpExclCnctTransfer>"
End If

If TotalDeductions_112A(i) <> "" Then
writeXML "<ITRForm:TotalDeductions>" & Application.WorksheetFunction.Round(TotalDeductions_112A(i), 4) & "</ITRForm:TotalDeductions>"
End If

If Balance_112A(i) <> "" Then
writeXML "<ITRForm:Balance>" & Application.WorksheetFunction.Round(Balance_112A(i), 4) & "</ITRForm:Balance>"
End If

writeXML "</ITRForm:Schedule112ADtls>"

Next

If Total_TotalSaleValue_112A <> "" Then
writeXML "<ITRForm:SaleValue112A>" & Application.WorksheetFunction.Round(Total_TotalSaleValue_112A, 4) & "</ITRForm:SaleValue112A>"
Else
writeXML "<ITRForm:SaleValue112A>0</ITRForm:SaleValue112A>"
End If

If Total_COAwithoutIndex_112A <> "" Then
writeXML "<ITRForm:CostAcqWithoutIndx112A>" & Application.WorksheetFunction.Round(Total_COAwithoutIndex_112A, 4) & "</ITRForm:CostAcqWithoutIndx112A>"
Else
writeXML "<ITRForm:CostAcqWithoutIndx112A>0</ITRForm:CostAcqWithoutIndx112A>"
End If

If Total_COAwithIndex_112A <> "" Then
writeXML "<ITRForm:AcquisitionCost112A>" & Application.WorksheetFunction.Round(Total_COAwithIndex_112A, 4) & "</ITRForm:AcquisitionCost112A>"
Else
writeXML "<ITRForm:AcquisitionCost112A>0</ITRForm:AcquisitionCost112A>"
End If

If Total_LTCGAssetAcquired_112A <> "" Then
writeXML "<ITRForm:LTCGBeforelowerB1B2112A>" & Application.WorksheetFunction.Round(Total_LTCGAssetAcquired_112A, 4) & "</ITRForm:LTCGBeforelowerB1B2112A>"
Else
writeXML "<ITRForm:LTCGBeforelowerB1B2112A>0</ITRForm:LTCGBeforelowerB1B2112A>"
End If

If Total_TotalFairMarketValue_112A <> "" Then
writeXML "<ITRForm:FairMktValueCapAst112A>" & Application.WorksheetFunction.Round(Total_TotalFairMarketValue_112A, 4) & "</ITRForm:FairMktValueCapAst112A>"
Else
writeXML "<ITRForm:FairMktValueCapAst112A>0</ITRForm:FairMktValueCapAst112A>"
End If

'If Total_FullValueConsideration_112A <> "" Then
'writeXML "<ITRForm:FullValueConsdr112A>" & Round(Total_FullValueConsideration_112A, 2) & "</ITRForm:FullValueConsdr112A>"
'Else
'writeXML "<ITRForm:FullValueConsdr112A>0</ITRForm:FullValueConsdr112A>"
'End If
'
'If Total_CostofImprovement_112A <> "" Then
'writeXML "<ITRForm:CostImprvWithoutIndx112A>" & Round(Total_CostofImprovement_112A, 2) & "</ITRForm:CostImprvWithoutIndx112A>"
'Else
'writeXML "<ITRForm:CostImprvWithoutIndx112A>0</ITRForm:CostImprvWithoutIndx112A>"
'End If

If Total_ExpenditureWholly_112A <> "" Then
writeXML "<ITRForm:ExpExclCnctTransfer112A>" & Application.WorksheetFunction.Round(Total_ExpenditureWholly_112A, 4) & "</ITRForm:ExpExclCnctTransfer112A>"
Else
writeXML "<ITRForm:ExpExclCnctTransfer112A>0</ITRForm:ExpExclCnctTransfer112A>"
End If

If Total_TotalDeductions_112A <> "" Then
writeXML "<ITRForm:Deductions112A>" & Application.WorksheetFunction.Round(Total_TotalDeductions_112A, 4) & "</ITRForm:Deductions112A>"
Else
writeXML "<ITRForm:Deductions112A>0</ITRForm:Deductions112A>"
End If

If Total_Balance_112A <> "" Then
writeXML "<ITRForm:Balance112A>" & Application.WorksheetFunction.Round(Total_Balance_112A, 4) & "</ITRForm:Balance112A>"
Else
writeXML "<ITRForm:Balance112A>0</ITRForm:Balance112A>"
End If

writeXML "</ITRForm:Schedule112A>"
End If
End Function

Function Schedule115AD()

Dim i As Long
If Not IsEmpty(ISINCode_115AD) And UBound(ISINCode_115AD) > 0 Then

writeXML "<ITRForm:Schedule115AD>"

subProcCaption = "115AD"
noOfProcessSub = UBound(ISINCode_115AD)

For i = 1 To UBound(ISINCode_115AD)

writeXML "<ITRForm:Schedule115ADDtls>"

If Trim(ShareAcq_115AD_1(i)) <> "" Then
If Trim(ShareAcq_115AD_1(i)) = "On or before 31st January 2018" Then
writeXML "<ITRForm:ShareOnOrBefore>" & "BE" & "</ITRForm:ShareOnOrBefore>"
ElseIf Trim(ShareAcq_115AD_1(i)) = "After 31st January 2018" Then
writeXML "<ITRForm:ShareOnOrBefore>" & "AE" & "</ITRForm:ShareOnOrBefore>"
End If
End If

If ISINCode_115AD(i) <> "" Then
writeXML "<ITRForm:ISINCode>" & UCase(ISINCode_115AD(i)) & "</ITRForm:ISINCode>"
End If

If NameofShare_115AD(i) <> "" Then
writeXML "<ITRForm:ShareUnitName>" & UCase(NameofShare_115AD(i)) & "</ITRForm:ShareUnitName>"
End If

If NoofShare_115AD(i) <> "" Then
writeXML "<ITRForm:NumSharesUnits>" & Application.WorksheetFunction.Round(NoofShare_115AD(i), 4) & "</ITRForm:NumSharesUnits>"
End If

If SalePriceperShare_115AD(i) <> "" Then
writeXML "<ITRForm:SalePricePerShareUnit>" & Application.WorksheetFunction.Round(SalePriceperShare_115AD(i), 4) & "</ITRForm:SalePricePerShareUnit>"
End If

If TotalSaleValue_115AD(i) <> "" Then
writeXML "<ITRForm:TotSaleValue>" & Application.WorksheetFunction.Round(TotalSaleValue_115AD(i), 4) & "</ITRForm:TotSaleValue>"
End If

If COAwithoutIndex_115AD(i) <> "" Then
writeXML "<ITRForm:CostAcqWithoutIndx>" & Application.WorksheetFunction.Round(COAwithoutIndex_115AD(i), 4) & "</ITRForm:CostAcqWithoutIndx>"
End If

If COAwithIndex_115AD(i) <> "" Then
writeXML "<ITRForm:AcquisitionCost>" & Application.WorksheetFunction.Round(COAwithIndex_115AD(i), 4) & "</ITRForm:AcquisitionCost>"
End If

If LTCGAssetAcquired_115AD(i) <> "" Then
writeXML "<ITRForm:LTCGBeforelower6and11>" & Application.WorksheetFunction.Round(LTCGAssetAcquired_115AD(i), 4) & "</ITRForm:LTCGBeforelower6and11>"
End If

If FairMarketValue_115AD(i) <> "" Then
writeXML "<ITRForm:FairMktValuePerShareunit>" & Application.WorksheetFunction.Round(FairMarketValue_115AD(i), 4) & "</ITRForm:FairMktValuePerShareunit>"
End If

If TotalFairMarketValue_115AD(i) <> "" Then
writeXML "<ITRForm:TotFairMktValueCapAst>" & Application.WorksheetFunction.Round(TotalFairMarketValue_115AD(i), 4) & "</ITRForm:TotFairMktValueCapAst>"
End If

'If FullValueConsideration_115AD(i) <> "" Then
'writeXML "<ITRForm:FullValueConsdr>" & Round(FullValueConsideration_115AD(i), 2) & "</ITRForm:FullValueConsdr>"
'End If
'
'If CostofImprovement_115AD(i) <> "" Then
'writeXML "<ITRForm:CostImprvWithoutIndx>" & Round(CostofImprovement_115AD(i), 2) & "</ITRForm:CostImprvWithoutIndx>"
'End If

If ExpenditureWholly_115AD(i) <> "" Then
writeXML "<ITRForm:ExpExclCnctTransfer>" & Application.WorksheetFunction.Round(ExpenditureWholly_115AD(i), 4) & "</ITRForm:ExpExclCnctTransfer>"
End If

If TotalDeductions_115AD(i) <> "" Then
writeXML "<ITRForm:TotalDeductions>" & Application.WorksheetFunction.Round(TotalDeductions_115AD(i), 4) & "</ITRForm:TotalDeductions>"
End If

If Balance_115AD(i) <> "" Then
writeXML "<ITRForm:Balance>" & Application.WorksheetFunction.Round(Balance_115AD(i), 4) & "</ITRForm:Balance>"
End If

writeXML "</ITRForm:Schedule115ADDtls>"

Next

If Total_TotalSaleValue_115AD <> "" Then
writeXML "<ITRForm:SaleValue115AD>" & Application.WorksheetFunction.Round((Total_TotalSaleValue_115AD), 4) & "</ITRForm:SaleValue115AD>"
Else
writeXML "<ITRForm:SaleValue115AD>0</ITRForm:SaleValue115AD>"
End If

If Total_COAwithoutIndex_115AD <> "" Then
writeXML "<ITRForm:CostAcqWithoutIndx115AD>" & Application.WorksheetFunction.Round((Total_COAwithoutIndex_115AD), 4) & "</ITRForm:CostAcqWithoutIndx115AD>"
Else
writeXML "<ITRForm:CostAcqWithoutIndx115AD>0</ITRForm:CostAcqWithoutIndx115AD>"
End If

If Total_COAwithIndex_115AD <> "" Then
writeXML "<ITRForm:AcquisitionCost115AD>" & Application.WorksheetFunction.Round((Total_COAwithIndex_115AD), 4) & "</ITRForm:AcquisitionCost115AD>"
Else
writeXML "<ITRForm:AcquisitionCost115AD>0</ITRForm:AcquisitionCost115AD>"
End If

If Total_LTCGAssetAcquired_115AD <> "" Then
writeXML "<ITRForm:LTCGBeforelowerB1B2115AD>" & Application.WorksheetFunction.Round((Total_LTCGAssetAcquired_115AD), 4) & "</ITRForm:LTCGBeforelowerB1B2115AD>"
Else
writeXML "<ITRForm:LTCGBeforelowerB1B2115AD>0</ITRForm:LTCGBeforelowerB1B2115AD>"
End If

If Total_TotalFairMarketValue_115AD <> "" Then
writeXML "<ITRForm:FairMktValueCapAst115AD>" & Application.WorksheetFunction.Round((Total_TotalFairMarketValue_115AD), 4) & "</ITRForm:FairMktValueCapAst115AD>"
Else
writeXML "<ITRForm:FairMktValueCapAst115AD>0</ITRForm:FairMktValueCapAst115AD>"
End If

'If Total_FullValueConsideration_115AD <> "" Then
'writeXML "<ITRForm:FullValueConsdr115AD>" & Round((Total_FullValueConsideration_115AD), 2) & "</ITRForm:FullValueConsdr115AD>"
'Else
'writeXML "<ITRForm:FullValueConsdr115AD>0</ITRForm:FullValueConsdr115AD>"
'End If
'
'If Total_CostofImprovement_115AD <> "" Then
'writeXML "<ITRForm:CostImprvWithoutIndx115AD>" & Round(Total_CostofImprovement_115AD, 2) & "</ITRForm:CostImprvWithoutIndx115AD>"
'Else
'writeXML "<ITRForm:CostImprvWithoutIndx115AD>0</ITRForm:CostImprvWithoutIndx115AD>"
'End If

If Total_ExpenditureWholly_115AD <> "" Then
writeXML "<ITRForm:ExpExclCnctTransfer115AD>" & Application.WorksheetFunction.Round(Total_ExpenditureWholly_115AD, 4) & "</ITRForm:ExpExclCnctTransfer115AD>"
Else
writeXML "<ITRForm:ExpExclCnctTransfer115AD>0</ITRForm:ExpExclCnctTransfer115AD>"
End If

If Total_TotalDeductions_115AD <> "" Then
writeXML "<ITRForm:Deductions115AD>" & Application.WorksheetFunction.Round(Total_TotalDeductions_115AD, 4) & "</ITRForm:Deductions115AD>"
Else
writeXML "<ITRForm:Deductions115AD>0</ITRForm:Deductions115AD>"
End If

If Total_Balance_115AD <> "" Then
writeXML "<ITRForm:Balance115AD>" & Application.WorksheetFunction.Round(Total_Balance_115AD, 4) & "</ITRForm:Balance115AD>"
Else
writeXML "<ITRForm:Balance115AD>0</ITRForm:Balance115AD>"
End If

writeXML "</ITRForm:Schedule115AD>"
End If
End Function

Function ScheduleOSXML()
    Dim i As Long
    Dim IncChargeableFrmOthSrc As Variant

    writeXML " <ITRForm:ScheduleOS>"
    writeXML " <ITRForm:IncOthThanOwnRaceHorse>"

    If Sheet14.Range("os.GrossIncomeChargeableToTax").value = "" Then
        writeXML " <ITRForm:GrossIncChrgblTaxAtAppRate>0</ITRForm:GrossIncChrgblTaxAtAppRate>"
    Else
        writeXML " <ITRForm:GrossIncChrgblTaxAtAppRate>" & UCase(Sheet14.Range("os.GrossIncomeChargeableToTax").value) & "</ITRForm:GrossIncChrgblTaxAtAppRate>"
    End If

    If DividendGross_15 = "" Then
        writeXML " <ITRForm:DividendGross>0</ITRForm:DividendGross>"
    Else
        writeXML " <ITRForm:DividendGross>" & UCase(DividendGross_15) & "</ITRForm:DividendGross>"
    End If

    If InterestGross_15 = "" Then
        writeXML " <ITRForm:InterestGross>0</ITRForm:InterestGross>"
    Else
        writeXML " <ITRForm:InterestGross>" & UCase(InterestGross_15) & "</ITRForm:InterestGross>"
    End If

    If InterestGross_i = "" Then
        writeXML " <ITRForm:IntrstFrmSavingBank>0</ITRForm:IntrstFrmSavingBank>"
    Else
        writeXML " <ITRForm:IntrstFrmSavingBank>" & UCase(InterestGross_i) & "</ITRForm:IntrstFrmSavingBank>"
    End If

    If InterestGross_ii = "" Then
        writeXML " <ITRForm:IntrstFrmTermDeposit>0</ITRForm:IntrstFrmTermDeposit>"
    Else
        writeXML " <ITRForm:IntrstFrmTermDeposit>" & UCase(InterestGross_ii) & "</ITRForm:IntrstFrmTermDeposit>"
    End If

    If InterestGross_iii = "" Then
        writeXML " <ITRForm:IntrstFrmIncmTaxRefund>0</ITRForm:IntrstFrmIncmTaxRefund>"
    Else
        writeXML " <ITRForm:IntrstFrmIncmTaxRefund>" & UCase(InterestGross_iii) & "</ITRForm:IntrstFrmIncmTaxRefund>"
    End If


     If Sheet14.Range("os.PassThroughIncome").value = "" Then
        writeXML " <ITRForm:NatofPassThrghIncome>0</ITRForm:NatofPassThrghIncome>"
    Else
        writeXML " <ITRForm:NatofPassThrghIncome>" & UVCase(Sheet14.Range("os.PassThroughIncome").value) & "</ITRForm:NatofPassThrghIncome>"
    End If



    If InterestGross_iv = "" Then
        writeXML " <ITRForm:IntrstFrmOthers>0</ITRForm:IntrstFrmOthers>"
    Else
        writeXML " <ITRForm:IntrstFrmOthers>" & UCase(InterestGross_iv) & "</ITRForm:IntrstFrmOthers>"
    End If


    If RentFromMachPlantBldgs_15 = "" Then
        writeXML " <ITRForm:RentFromMachPlantBldgs>0</ITRForm:RentFromMachPlantBldgs>"
    Else
        writeXML " <ITRForm:RentFromMachPlantBldgs>" & UCase(RentFromMachPlantBldgs_15) & "</ITRForm:RentFromMachPlantBldgs>"
    End If


    If Sheet14.Range("os.Totalsec56").value = "" Then
        writeXML " <ITRForm:Tot562x>0</ITRForm:Tot562x>"
    Else
        writeXML " <ITRForm:Tot562x>" & UVCase(Sheet14.Range("os.Totalsec56").value) & "</ITRForm:Tot562x>"
    End If


    If Sheet14.Range("os.sec56Aggr").value = "" Then
        writeXML " <ITRForm:Aggrtvaluewithoutcons562x>0</ITRForm:Aggrtvaluewithoutcons562x>"
    Else
        writeXML " <ITRForm:Aggrtvaluewithoutcons562x>" & UVCase(Sheet14.Range("os.sec56Aggr").value) & "</ITRForm:Aggrtvaluewithoutcons562x>"
    End If


    If Sheet14.Range("os.sec56immov").value = "" Then
        writeXML " <ITRForm:Immovpropwithoutcons562x>0</ITRForm:Immovpropwithoutcons562x>"
    Else
        writeXML " <ITRForm:Immovpropwithoutcons562x>" & UVCase(Sheet14.Range("os.sec56immov").value) & "</ITRForm:Immovpropwithoutcons562x>"
    End If

    If Sheet14.Range("os.sec56immovCons").value = "" Then
        writeXML " <ITRForm:Immovpropinadeqcons562x>0</ITRForm:Immovpropinadeqcons562x>"
    Else
        writeXML " <ITRForm:Immovpropinadeqcons562x>" & UVCase(Sheet14.Range("os.sec56immovCons").value) & "</ITRForm:Immovpropinadeqcons562x>"
    End If


    If Sheet14.Range("os.sec56property").value = "" Then
        writeXML " <ITRForm:Anyotherpropwithoutcons562x>0</ITRForm:Anyotherpropwithoutcons562x>"
    Else
        writeXML " <ITRForm:Anyotherpropwithoutcons562x>" & UVCase(Sheet14.Range("os.sec56property").value) & "</ITRForm:Anyotherpropwithoutcons562x>"
    End If

    If Sheet14.Range("os.sec56propertyCons").value = "" Then
        writeXML " <ITRForm:Anyotherpropinadeqcons562x>0</ITRForm:Anyotherpropinadeqcons562x>"
    Else
        writeXML " <ITRForm:Anyotherpropinadeqcons562x>" & UVCase(Sheet14.Range("os.sec56propertyCons").value) & "</ITRForm:Anyotherpropinadeqcons562x>"
    End If

    If Sheet14.Range("os_FamilyPension").value = "" Then
        writeXML " <ITRForm:FamilyPension>0</ITRForm:FamilyPension>"
    Else
        writeXML " <ITRForm:FamilyPension>" & UVCase(Sheet14.Range("os_FamilyPension").value) & "</ITRForm:FamilyPension>"
    End If

    If Sheet14.Range("os.TotalAnyOtherIncome").value = "" Then
        writeXML " <ITRForm:AnyOtherIncome>0</ITRForm:AnyOtherIncome>"
    Else
        writeXML " <ITRForm:AnyOtherIncome>" & UVCase(Sheet14.Range("os.TotalAnyOtherIncome").value) & "</ITRForm:AnyOtherIncome>"
    End If

    'New Table

If (Not IsEmpty(OthersSourceSection_OS)) And UBound(OthersSourceSection_OS) > 0 Then
                writeXML "<ITRForm:OthersInc>"

            For i = 1 To UBound(OthersSourceSection_OS)

                 writeXML "<ITRForm:OthersIncDtls>"
                        If OthersSourceSection_OS(i) <> "" Then
                            writeXML "<ITRForm:OthNatOfInc>" & OthersSourceSection_OS(i) & "</ITRForm:OthNatOfInc>"
                        Else
                            writeXML "<ITRForm:OthNatOfInc/>"
                        End If

                        If OthersSourceAmount_OS(i) <> "" Then
                            writeXML "<ITRForm:OthAmount>" & UCase(OthersSourceAmount_OS(i)) & "</ITRForm:OthAmount>"

                        End If

                    writeXML "</ITRForm:OthersIncDtls>"
                Next

     writeXML "</ITRForm:OthersInc>"
    End If


    If Sheet14.Range("os.IncomeChargeableSpecialRates").value = "" Then
        writeXML " <ITRForm:IncChargeableSpecialRates>0</ITRForm:IncChargeableSpecialRates>"
    Else
        writeXML " <ITRForm:IncChargeableSpecialRates>" & UVCase(Sheet14.Range("os.IncomeChargeableSpecialRates").value) & "</ITRForm:IncChargeableSpecialRates>"
    End If

    If Sheet14.Range("os.WinLottRacePuzz").value = "" Then
        writeXML " <ITRForm:LtryPzzlChrgblUs115BB>0</ITRForm:LtryPzzlChrgblUs115BB>"
    Else
        writeXML " <ITRForm:LtryPzzlChrgblUs115BB>" & UVCase(Sheet14.Range("os.WinLottRacePuzz").value) & "</ITRForm:LtryPzzlChrgblUs115BB>"
    End If

    If Sheet14.Range("os.Total115BE").value = "" Then
        writeXML " <ITRForm:IncChrgblUs115BBE>0</ITRForm:IncChrgblUs115BBE>"
    Else
        writeXML " <ITRForm:IncChrgblUs115BBE>" & UVCase(Sheet14.Range("os.Total115BE").value) & "</ITRForm:IncChrgblUs115BBE>"
    End If





        If Sheet14.Range("os.us68").value = "" Then
        writeXML " <ITRForm:CashCreditsUs68>0</ITRForm:CashCreditsUs68>"
    Else
        writeXML " <ITRForm:CashCreditsUs68>" & UVCase(Sheet14.Range("os.us68").value) & "</ITRForm:CashCreditsUs68>"
    End If

        If Sheet14.Range("os.us69").value = "" Then
        writeXML " <ITRForm:UnExplndInvstmntsUs69>0</ITRForm:UnExplndInvstmntsUs69>"
    Else
        writeXML " <ITRForm:UnExplndInvstmntsUs69>" & UVCase(Sheet14.Range("os.us69").value) & "</ITRForm:UnExplndInvstmntsUs69>"
    End If

        If Sheet14.Range("os.us69A").value = "" Then
        writeXML " <ITRForm:UnExplndMoneyUs69A>0</ITRForm:UnExplndMoneyUs69A>"
    Else
        writeXML " <ITRForm:UnExplndMoneyUs69A>" & UVCase(Sheet14.Range("os.us69A").value) & "</ITRForm:UnExplndMoneyUs69A>"
    End If

        If Sheet14.Range("os.us69B").value = "" Then
        writeXML " <ITRForm:UnDsclsdInvstmntsUs69B>0</ITRForm:UnDsclsdInvstmntsUs69B>"
    Else
        writeXML " <ITRForm:UnDsclsdInvstmntsUs69B>" & UVCase(Sheet14.Range("os.us69B").value) & "</ITRForm:UnDsclsdInvstmntsUs69B>"
    End If

        If Sheet14.Range("os.us69C").value = "" Then
        writeXML " <ITRForm:UnExplndExpndtrUs69C>0</ITRForm:UnExplndExpndtrUs69C>"
    Else
        writeXML " <ITRForm:UnExplndExpndtrUs69C>" & UVCase(Sheet14.Range("os.us69C").value) & "</ITRForm:UnExplndExpndtrUs69C>"
    End If

        If Sheet14.Range("os.us69D").value = "" Then
        writeXML " <ITRForm:AmtBrwdRepaidOnHundiUs69D>0</ITRForm:AmtBrwdRepaidOnHundiUs69D>"
    Else
        writeXML " <ITRForm:AmtBrwdRepaidOnHundiUs69D>" & UVCase(Sheet14.Range("os.us69D").value) & "</ITRForm:AmtBrwdRepaidOnHundiUs69D>"
    End If




'    OS section
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(1) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(1)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(1) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(1)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(2) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(2)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(2) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(2)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(3) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(3)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(3) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(3)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(4) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(4)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(4) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(4)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(5) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(5)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(5) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(5)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(6) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(6)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(6) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(6)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(7) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(7)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(7) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(7)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(8) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(8)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(8) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(8)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(9) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(9)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(9) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(9)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(10) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(10)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(10) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(10)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(11) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(11)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(11) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(11)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(12) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(12)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(12) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(12)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(13) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(13)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(13) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(13)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'    writeXML "</ITRForm:OthersGrossDtls>"
'
'    writeXML "<ITRForm:OthersGrossDtls>"
'    If SourceDescription_OS(14) <> "" Then
'        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(14)) & "</ITRForm:SourceDescription>"
'    Else
'    End If
'
'    If SourceAmount_OS(14) <> "" Then
'        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(14)) & "</ITRForm:SourceAmount>"
'    Else
'        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'    End If
'
'   writeXML "</ITRForm:OthersGrossDtls>"
''
''    writeXML "<ITRForm:OthersGrossDtls>"
''    If SourceDescription_OS(15) <> "" Then
''        writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(15)) & "</ITRForm:SourceDescription>"
''    Else
''    End If
'
''    If SourceAmount_OS(15) <> "" Then
''        writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(15)) & "</ITRForm:SourceAmount>"
''    Else
''        writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
''    End If
'
''    writeXML "</ITRForm:OthersGrossDtls>"
''
''    writeXML "<ITRForm:OthersGrossDtls>"
''        writeXML "                   <ITRForm:SourceDescription>Others</ITRForm:SourceDescription>"
''        If Sheet14.Range("os.othDesc").value <> "" Then
''            writeXML "                   <ITRForm:OtherSourceDesc>" & UCase(Sheet14.Range("os.othDesc").value) & "</ITRForm:OtherSourceDesc>"
''        Else
''            writeXML "                   <ITRForm:OtherSourceDesc>Others</ITRForm:OtherSourceDesc>"
''        End If
''
''        If SourceAmount_OS(15) <> "" Then
''            writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(15)) & "</ITRForm:SourceAmount>"
''        Else
''            writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
''        End If
''
''     writeXML "</ITRForm:OthersGrossDtls>"
'
'    'other source goes here
'    If Not IsEmpty(OthersSourceSection_OS) Then
'        For i = 1 To UBound(OthersSourceSection_OS)
'            writeXML "<ITRForm:OthersGrossDtls>"
'            writeXML "<ITRForm:SourceDescription>Others</ITRForm:SourceDescription>"
'            writeXML "<ITRForm:OtherSourceDesc>" & OthersSourceSection_OS(i) & "</ITRForm:OtherSourceDesc>"
'            If OthersSourceAmount_OS(i) <> "" Then
'                writeXML "<ITRForm:SourceAmount>" & UCase(OthersSourceAmount_OS(i)) & "</ITRForm:SourceAmount>"
'            Else
'                writeXML "<ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'            End If
'            writeXML "</ITRForm:OthersGrossDtls>"
'        Next
'    End If
'    ' ends here
'
'
'    If UBound(SourceDescription_OS) > 0 Then
'        For i = 15 To UBound(SourceDescription_OS)
'        If SourceDescription_OS(i) <> "" Then
'            writeXML "           <ITRForm:OthersGrossDtls>"
'                If SourceDescription_OS(i) <> "" Then
'                    writeXML "                   <ITRForm:SourceDescription>" & Trim(SourceDescription_OS(i)) & "</ITRForm:SourceDescription>"
'                Else
'                End If
'
'                If SourceAmount_OS(i) <> "" Then
'                    writeXML "               <ITRForm:SourceAmount>" & UCase(SourceAmount_OS(i)) & "</ITRForm:SourceAmount>"
'                Else
'                   writeXML "               <ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
'                End If
'            writeXML "           </ITRForm:OthersGrossDtls>"
'            End If
'        Next
'    End If

    If UBound(os_AssYear) > 0 Then
    writeXML "           <ITRForm:TaxAccumulatedBalRecPF>"
        For i = 1 To UBound(os_AssYear)


                 writeXML "           <ITRForm:TaxAccmltdBalRecPFDtls>"

                If os_AssYear(i) <> "" And Mid(os_AssYear(i), 1, 1) <> "(" Then
                    writeXML "                   <ITRForm:AssessmentYear>" & UCase(os_AssYear(i)) & "</ITRForm:AssessmentYear>"
                Else
                    writeXML "                   <ITRForm:AssessmentYear>" & 0 & "</ITRForm:AssessmentYear>"
                End If

                If os_IncomeBenefit(i) <> "" Then
                    writeXML "               <ITRForm:IncomeBenefit>" & UCase(os_IncomeBenefit(i)) & "</ITRForm:IncomeBenefit>"
                Else
                   writeXML "               <ITRForm:IncomeBenefit>0</ITRForm:IncomeBenefit>"
                End If

                If os_TaxBenefit(i) <> "" Then
                    writeXML "               <ITRForm:TaxBenefit>" & UCase(os_TaxBenefit(i)) & "</ITRForm:TaxBenefit>"
                Else
                   writeXML "               <ITRForm:TaxBenefit>0</ITRForm:TaxBenefit>"
                End If

                writeXML "           </ITRForm:TaxAccmltdBalRecPFDtls>"

                 Next

                If Sheet14.Range("os.IncomeBenefitTotal").value <> "" Then
                    writeXML "               <ITRForm:TotalIncomeBenefit>" & UCase(Sheet14.Range("os.IncomeBenefitTotal").value) & "</ITRForm:TotalIncomeBenefit>"
                End If

                If Sheet14.Range("os.TaxBenefitTotal").value <> "" Then
                    writeXML "               <ITRForm:TotalTaxBenefit>" & UCase(Sheet14.Range("os.TaxBenefitTotal").value) & "</ITRForm:TotalTaxBenefit>"
                End If


     writeXML "           </ITRForm:TaxAccumulatedBalRecPF>"

    End If


    If Sheet14.Range("os.TotalChargeableSpecialRate").value = "" Then
        writeXML "<ITRForm:OthersGross>0</ITRForm:PassThrIncOSChrgblSplRate>"
    Else
        writeXML "<ITRForm:OthersGross>" & UVCase(Sheet14.Range("os.TotalChargeableSpecialRate").value) & "</ITRForm:OthersGross>"
    End If


   If UBound(SourceDescription_OS) > 0 Then
    If Not IsEmpty(SourceDescription_OS) And UBound(SourceDescription_OS) > 0 Then
        For i = 1 To UBound(SourceDescription_OS)
            writeXML "<ITRForm:OthersGrossDtls>"
            If SourceDescription_OS(i) <> "" Then
                writeXML "<ITRForm:SourceDescription>" & Trim(SourceDescription_OS(i)) & "</ITRForm:SourceDescription>"
            End If

            If SourceAmount_OS(i) <> "" Then
                writeXML "<ITRForm:SourceAmount>" & UCase(SourceAmount_OS(i)) & "</ITRForm:SourceAmount>"
            Else
                writeXML "<ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
            End If
            writeXML "</ITRForm:OthersGrossDtls>"
        Next
    End If
    End If




    If Sheet14.Range("os.TotalPassThroughIncome").value = "" Then
        writeXML "<ITRForm:PassThrIncOSChrgblSplRate>0</ITRForm:PassThrIncOSChrgblSplRate>"
    Else
        writeXML "<ITRForm:PassThrIncOSChrgblSplRate>" & UVCase(Sheet14.Range("os.TotalPassThroughIncome").value) & "</ITRForm:PassThrIncOSChrgblSplRate>"
    End If

    If Not IsEmpty(os_PTIDescription) And UBound(os_PTIDescription) > 0 Then
        For i = 1 To UBound(os_PTIDescription)
            writeXML "<ITRForm:PTIOthersGrossDtls>"
            If os_PTIDescription(i) <> "" Then

            If os_PTIDescription(i) = "PTI_5A1b" Then
                os_PTIDescription(i) = "PTI_5A1bA"
            End If

                writeXML "<ITRForm:SourceDescription>" & os_PTIDescription(i) & "</ITRForm:SourceDescription>"
            End If

            If os_PTIAmount(i) <> "" Then
                writeXML "<ITRForm:SourceAmount>" & UCase(os_PTIAmount(i)) & "</ITRForm:SourceAmount>"
            Else
                writeXML "<ITRForm:SourceAmount>0</ITRForm:SourceAmount>"
            End If
            writeXML "</ITRForm:PTIOthersGrossDtls>"
        Next
    End If



'    If OthersGross_15 = "" Then
'        writeXML " <ITRForm:OthersGross>0</ITRForm:OthersGross>"
'    Else
'        writeXML " <ITRForm:OthersGross>" & UCase(OthersGross_15) & "</ITRForm:OthersGross>"
'    End If

'    If TotalOSGross_15 = "" Then
'        writeXML " <ITRForm:TotalOSGross>0</ITRForm:TotalOSGross>"
'    Else
'        writeXML " <ITRForm:TotalOSGross>" & UCase(TotalOSGross_15) & "</ITRForm:TotalOSGross>"
'    End If

    writeXML " <ITRForm:IncChargblSplRateOS>"
    'writeXML " <ITRForm:WinningFrmLotteries>" & UVCase(fi_WinningFrmLotteries_os) & "</ITRForm:WinningFrmLotteries>"
   ' writeXML " <ITRForm:DividendIncome115BBDA>" & UVCase(Sheet14.Range("os.DividendIncome").value) & "</ITRForm:DividendIncome115BBDA>"
   ' writeXML " <ITRForm:DividendIncome115BBE>" & UVCase(Sheet14.Range("os.DeemedIncome").value) & "</ITRForm:DividendIncome115BBE>"
   ' writeXML " <ITRForm:DividendIncome115BBF>" & UVCase(Sheet14.Range("os.IncomeChargeable115BBF").value) & "</ITRForm:DividendIncome115BBF>"
    'writeXML " <ITRForm:DividendIncome115BBG>" & UVCase(Sheet14.Range("os.IncomeChargeable115BBG").value) & "</ITRForm:DividendIncome115BBG>"
  writeXML "<ITRForm:TotalAmtTaxUsDTAASchOs>" & os_DTAAAmount & "</ITRForm:TotalAmtTaxUsDTAASchOs>"

    If UBound(os_Amount_income) > 0 Then
    If os_Amount_income(1) <> "" Then
    writeXML "<ITRForm:NRIOsDTAA>"
   ' Dim i As Variant
    For i = 1 To UBound(os_Amount_income)
    If os_Amount_income(i) <> "" Then
    writeXML "<ITRForm:NRIDTAADtlsSchOS>"
    writeXML "<ITRForm:DTAAamt>" & os_Amount_income(i) & "</ITRForm:DTAAamt>"
    writeXML "<ITRForm:NatureOfIncome>" & os_Nature_income(i) & "</ITRForm:NatureOfIncome>"
    writeXML "<ITRForm:CountryName>" & os_CountryName(i) & "</ITRForm:CountryName>"
    writeXML "<ITRForm:CountryCode>" & os_CountryCode(i) & "</ITRForm:CountryCode>"
    writeXML "<ITRForm:DTAAarticle>" & os_Article_DTAA(i) & "</ITRForm:DTAAarticle>"
    writeXML "<ITRForm:RateAsPerTreaty>" & Round(os_Rate_DTAA(i), 2) & "</ITRForm:RateAsPerTreaty>"
    If os_TRC(i) <> "" Then
    writeXML "<ITRForm:TaxRescertifiedFlag>" & os_TRC(i) & "</ITRForm:TaxRescertifiedFlag>"
    End If
    If os_SectionAct(i) = "PTI_5A1b" Then
        os_SectionAct(i) = "PTI_5A1bA"
    ElseIf os_SectionAct(i) = "5A1b" Then
        os_SectionAct(i) = "5A1bA"

    End If


    writeXML "<ITRForm:ItemNoincl>" & Trim(os_SectionAct(i)) & "</ITRForm:ItemNoincl>"
    writeXML "<ITRForm:RateAsPerITAct>" & Round(os_Rate_ITAct(i), 2) & "</ITRForm:RateAsPerITAct>"
    writeXML "<ITRForm:ApplicableRate>" & Round(ApplRate(i), 2) & "</ITRForm:ApplicableRate>"
    writeXML "</ITRForm:NRIDTAADtlsSchOS>"
    End If
    Next

    writeXML "</ITRForm:NRIOsDTAA>"
    End If
    End If


    'writeXML "<ITRForm:TotalOSGrossChargblSplRate>" & UVCase(fiii_TotalOSGrossChargblSplRate_os) & "</ITRForm:TotalOSGrossChargblSplRate>"
    writeXML "</ITRForm:IncChargblSplRateOS>"
    'writeXML "<ITRForm:GrossAmtChargblNormalRate>" & UVCase(g_GrossAmtChargblNormalRate_os) & "</ITRForm:GrossAmtChargblNormalRate>"


    writeXML " <ITRForm:Deductions>"
    If Expenses_15 = "" Then
        writeXML " <ITRForm:Expenses>0</ITRForm:Expenses>"
    Else
        writeXML " <ITRForm:Expenses>" & UCase(Expenses_15) & "</ITRForm:Expenses>"
    End If

    If Deduction57iia_15 = "" Then
        writeXML " <ITRForm:DeductionUs57iia>0</ITRForm:DeductionUs57iia>"
    Else
        writeXML " <ITRForm:DeductionUs57iia>" & UCase(Deduction57iia_15) & "</ITRForm:DeductionUs57iia>"
    End If


    If Depreciation_15 = "" Then
        writeXML " <ITRForm:Depreciation>0</ITRForm:Depreciation>"
    Else
        writeXML " <ITRForm:Depreciation>" & UCase(Depreciation_15) & "</ITRForm:Depreciation>"
    End If

    If TotDeductions_15 = "" Then
        writeXML " <ITRForm:TotDeductions>0</ITRForm:TotDeductions>"
    Else
        writeXML " <ITRForm:TotDeductions>" & UCase(TotDeductions_15) & "</ITRForm:TotDeductions>"
    End If
    writeXML " </ITRForm:Deductions>"


    If AmtNotDedSec58 = "" Then
        writeXML " <ITRForm:AmtNotDeductibleUs58>0</ITRForm:AmtNotDeductibleUs58>"
    Else
        writeXML " <ITRForm:AmtNotDeductibleUs58>" & UCase(AmtNotDedSec58) & "</ITRForm:AmtNotDeductibleUs58>"
    End If

    If ProfitCharSec59 = "" Then
        writeXML " <ITRForm:ProfitChargTaxUs59>0</ITRForm:ProfitChargTaxUs59>"
    Else
        writeXML " <ITRForm:ProfitChargTaxUs59>" & UCase(ProfitCharSec59) & "</ITRForm:ProfitChargTaxUs59>"
    End If


    If BalanceNoRaceHorse_15 = "" Then
        writeXML " <ITRForm:BalanceNoRaceHorse>0</ITRForm:BalanceNoRaceHorse>"
    Else
        writeXML " <ITRForm:BalanceNoRaceHorse>" & UCase(BalanceNoRaceHorse_15) & "</ITRForm:BalanceNoRaceHorse>"
    End If

    writeXML " </ITRForm:IncOthThanOwnRaceHorse>"


    If TotOthSrcNoRaceHorse_15 = "" Then
        writeXML " <ITRForm:TotOthSrcNoRaceHorse>0</ITRForm:TotOthSrcNoRaceHorse>"
    Else
        writeXML " <ITRForm:TotOthSrcNoRaceHorse>" & UCase(TotOthSrcNoRaceHorse_15) & "</ITRForm:TotOthSrcNoRaceHorse>"
    End If
    writeXML " <ITRForm:IncFromOwnHorse>"

    If Receipts_15 = "" Then
        writeXML " <ITRForm:Receipts>0</ITRForm:Receipts>"
    Else
        writeXML " <ITRForm:Receipts>" & UCase(Receipts_15) & "</ITRForm:Receipts>"
    End If

    If DeductSec57_15 = "" Then
        writeXML " <ITRForm:DeductSec57>0</ITRForm:DeductSec57>"
    Else
        writeXML " <ITRForm:DeductSec57>" & UCase(DeductSec57_15) & "</ITRForm:DeductSec57>"
    End If

    If AmtNotDed = "" Then
        writeXML " <ITRForm:AmtNotDeductibleUs58>0</ITRForm:AmtNotDeductibleUs58>"
    Else
        writeXML " <ITRForm:AmtNotDeductibleUs58>" & UCase(AmtNotDed) & "</ITRForm:AmtNotDeductibleUs58>"
    End If

    If ProfitSec59 = "" Then
        writeXML " <ITRForm:ProfitChargTaxUs59>0</ITRForm:ProfitChargTaxUs59>"
    Else
        writeXML " <ITRForm:ProfitChargTaxUs59>" & UCase(ProfitSec59) & "</ITRForm:ProfitChargTaxUs59>"
    End If

    If BalanceOwnRaceHorse_15 = "" Then
        writeXML " <ITRForm:BalanceOwnRaceHorse>0</ITRForm:BalanceOwnRaceHorse>"
    Else
        writeXML " <ITRForm:BalanceOwnRaceHorse>" & UCase(BalanceOwnRaceHorse_15) & "</ITRForm:BalanceOwnRaceHorse>"
    End If
    writeXML " </ITRForm:IncFromOwnHorse>"

    IncChargeableFrmOthSrc = Range("os.IncChargeable")
    If IncChargeableFrmOthSrc = "" Then
        writeXML " <ITRForm:IncChargeable>0</ITRForm:IncChargeable>"
    Else
        writeXML " <ITRForm:IncChargeable>" & UVCase(IncChargeableFrmOthSrc) & "</ITRForm:IncChargeable>"
    End If

      writeXML " <ITRForm:DividendIncUs115BBDA>"
      writeXML " <ITRForm:DateRange>"

       If Sheet14.Range("BBDA.Upto15Of6").value = "" Then
        writeXML "<ITRForm:Upto15Of6>0</ITRForm:Upto15Of6>"
         Else
        writeXML "<ITRForm:Upto15Of6>" & UVCase(Sheet14.Range("BBDA.Upto15Of6").value) & "</ITRForm:Upto15Of6>"
       End If

       If Sheet14.Range("BBDA.Upto15Of9").value = "" Then
        writeXML "<ITRForm:Up16Of6To15Of9>0</ITRForm:Up16Of6To15Of9>"
         Else
        writeXML "<ITRForm:Up16Of6To15Of9>" & UVCase(Sheet14.Range("BBDA.Upto15Of9").value) & "</ITRForm:Up16Of6To15Of9>"
       End If

       If Sheet14.Range("BBDA.Up16Of9To15Of12").value = "" Then
        writeXML "<ITRForm:Up16Of9To15Of12>0</ITRForm:Up16Of9To15Of12>"
         Else
        writeXML "<ITRForm:Up16Of9To15Of12>" & UVCase(Sheet14.Range("BBDA.Up16Of9To15Of12").value) & "</ITRForm:Up16Of9To15Of12>"
       End If

       If Sheet14.Range("BBDA.Up16Of12To15Of3").value = "" Then
        writeXML "<ITRForm:Up16Of12To15Of3>0</ITRForm:Up16Of12To15Of3>"
         Else
        writeXML "<ITRForm:Up16Of12To15Of3>" & UVCase(Sheet14.Range("BBDA.Up16Of12To15Of3").value) & "</ITRForm:Up16Of12To15Of3>"
       End If

       If Sheet14.Range("BBDA.Up16Of3To31Of3").value = "" Then
        writeXML "<ITRForm:Up16Of3To31Of3>0</ITRForm:Up16Of3To31Of3>"
         Else
        writeXML "<ITRForm:Up16Of3To31Of3>" & UVCase(Sheet14.Range("BBDA.Up16Of3To31Of3").value) & "</ITRForm:Up16Of3To31Of3>"
       End If
      writeXML " </ITRForm:DateRange>"
      writeXML " </ITRForm:DividendIncUs115BBDA>"



 writeXML " <ITRForm:IncFrmLottery>"
      writeXML " <ITRForm:DateRange>"

       If Sheet14.Range("IncomefrmLottery.Upto15Of6").value = "" Then
        writeXML "<ITRForm:Upto15Of6>0</ITRForm:Upto15Of6>"
         Else
        writeXML "<ITRForm:Upto15Of6>" & UVCase(Sheet14.Range("IncomefrmLottery.Upto15Of6").value) & "</ITRForm:Upto15Of6>"
       End If

       If Sheet14.Range("IncomefrmLottery.Upto15Of9").value = "" Then
        writeXML "<ITRForm:Up16Of6To15Of9>0</ITRForm:Up16Of6To15Of9>"
         Else
        writeXML "<ITRForm:Up16Of6To15Of9>" & UVCase(Sheet14.Range("IncomefrmLottery.Upto15Of9").value) & "</ITRForm:Up16Of6To15Of9>"
       End If

       If Sheet14.Range("IncomefrmLottery.Up16Of9To15Of12").value = "" Then
        writeXML "<ITRForm:Up16Of9To15Of12>0</ITRForm:Up16Of9To15Of12>"
         Else
        writeXML "<ITRForm:Up16Of9To15Of12>" & UVCase(Sheet14.Range("IncomefrmLottery.Up16Of9To15Of12").value) & "</ITRForm:Up16Of9To15Of12>"
       End If

       If Sheet14.Range("IncomefrmLottery.Up16Of12To15Of3").value = "" Then
        writeXML "<ITRForm:Up16Of12To15Of3>0</ITRForm:Up16Of12To15Of3>"
         Else
        writeXML "<ITRForm:Up16Of12To15Of3>" & UVCase(Sheet14.Range("IncomefrmLottery.Up16Of12To15Of3").value) & "</ITRForm:Up16Of12To15Of3>"
       End If

       If Sheet14.Range("IncomefrmLottery.Up16Of3To31Of3").value = "" Then
        writeXML "<ITRForm:Up16Of3To31Of3>0</ITRForm:Up16Of3To31Of3>"
         Else
        writeXML "<ITRForm:Up16Of3To31Of3>" & UVCase(Sheet14.Range("IncomefrmLottery.Up16Of3To31Of3").value) & "</ITRForm:Up16Of3To31Of3>"
       End If
      writeXML " </ITRForm:DateRange>"
      writeXML " </ITRForm:IncFrmLottery>"








    writeXML " </ITRForm:ScheduleOS>"
End Function
'Function XMLPART_CT()
'Dim i As Long
'
'If UCase(Sheet7.Range("CTCheck").value) = "YES" Then
'
'writeXML "<ITRForm:ScheduleCT>"
'
'If Not IsEmpty(Name_A_CT) And UBound(Name_A_CT) > 0 Then
'
'subProcCaption = "Schedule CT"
'noOfProcessSub = UBound(Name_A_CT)
'For i = 1 To UBound(Name_A_CT)
'
'writeXML "<ITRForm:LoanOrDepositSec269SS>"
'
'If Name_A_CT(i) <> "" Then
'writeXML "<ITRForm:Name>" & UCase(Name_A_CT(i)) & "</ITRForm:Name>"
' Else
' writeXML "<ITRForm:Name/>"
'End If
'
' writeXML "<ITRForm:AddressCT>"
'
'If Door_A_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceNo>" & UCase(Door_A_CT(i)) & "</ITRForm:ResidenceNo>"
'Else
'writeXML "<ITRForm:ResidenceNo/>"
'End If
'
'If Village_A_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceName>" & UCase(Village_A_CT(i)) & "</ITRForm:ResidenceName>"
'Else
'writeXML "<ITRForm:ResidenceName/>"
'End If
'
'If Road_A_CT(i) <> "" Then
'writeXML "<ITRForm:RoadOrStreet>" & UCase(Road_A_CT(i)) & "</ITRForm:RoadOrStreet>"
'Else
'writeXML "<ITRForm:RoadOrStreet/>"
'End If
'
'If Area_A_CT(i) <> "" Then
'writeXML "<ITRForm:LocalityOrArea>" & UCase(Area_A_CT(i)) & "</ITRForm:LocalityOrArea>"
'Else
'writeXML "<ITRForm:LocalityOrArea/>"
'End If
'
'If Town_A_CT(i) <> "" Then
'writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(Town_A_CT(i)) & "</ITRForm:CityOrTownOrDistrict>"
'Else
'writeXML "<ITRForm:CityOrTownOrDistrict/>"
'End If
'
'
'If StateCode_A_CT(i) <> "" And StateCode_A_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:StateCode>" & UCase(StateCode_A_CT(i)) & "</ITRForm:StateCode>"
'Else
'writeXML "<ITRForm:StateCode/>"
'End If
'
'If CountryCode_A_CT(i) <> "" And CountryCode_A_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:CountryCode>" & UCase(CountryCode_A_CT(i)) & "</ITRForm:CountryCode>"
'Else
'writeXML "<ITRForm:CountryCode/>"
'End If
'
'If PinCode_A_CT(i) <> "" Then
'writeXML "<ITRForm:PinCode>" & UCase(PinCode_A_CT(i)) & "</ITRForm:PinCode>"
'Else
'writeXML "<ITRForm:PinCode/>"
'End If
'
'If ZipCode_A_CT(i) <> "" Then
'writeXML "<ITRForm:ZipCode>" & UCase(ZipCode_A_CT(i)) & "</ITRForm:ZipCode>"
'Else
'writeXML "<ITRForm:ZipCode/>"
'End If
'
'writeXML "</ITRForm:AddressCT>"
'
'If PAN_A_CT(i) <> "" Then
'writeXML "<ITRForm:Pan>" & UCase(PAN_A_CT(i)) & "</ITRForm:Pan>"
'Else
'writeXML "<ITRForm:Pan/>"
'End If
'
'If Nature_A_CT(i) <> "" Then
'writeXML "<ITRForm:NatureOfTrans>" & UCase(Nature_A_CT(i)) & "</ITRForm:NatureOfTrans>"
'Else
'writeXML "<ITRForm:NatureOfTrans/>"
'End If
'
'If Amount_A_CT(i) <> "" Then
'writeXML "<ITRForm:AggAmount>" & UCase(Amount_A_CT(i)) & "</ITRForm:AggAmount>"
'Else
'writeXML "<ITRForm:AggAmount/>"
'End If
'
'
'writeXML "</ITRForm:LoanOrDepositSec269SS>"
'UpdateProgressBar
'
'Next
'
'
'
'End If
'
'If Not IsEmpty(Name_B_CT) And UBound(Name_B_CT) > 0 Then
'
'subProcCaption = "Schedule CTB"
'noOfProcessSub = UBound(Name_B_CT)
'For i = 1 To UBound(Name_B_CT)
'
' writeXML "<ITRForm:ReceiptExcSec269ST>"
'
'If Name_B_CT(i) <> "" Then
'writeXML "<ITRForm:Name>" & UCase(Name_B_CT(i)) & "</ITRForm:Name>"
' Else
' writeXML "<ITRForm:Name/>"
'End If
'
' writeXML "<ITRForm:AddressCT>"
'
'If Door_B_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceNo>" & UCase(Door_B_CT(i)) & "</ITRForm:ResidenceNo>"
'Else
'writeXML "<ITRForm:ResidenceNo/>"
'End If
'
'If Village_B_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceName>" & UCase(Village_B_CT(i)) & "</ITRForm:ResidenceName>"
'Else
'writeXML "<ITRForm:ResidenceName/>"
'End If
'
'If Road_B_CT(i) <> "" Then
'writeXML "<ITRForm:RoadOrStreet>" & UCase(Road_B_CT(i)) & "</ITRForm:RoadOrStreet>"
'Else
'writeXML "<ITRForm:RoadOrStreet/>"
'End If
'
'If Area_B_CT(i) <> "" Then
'writeXML "<ITRForm:LocalityOrArea>" & UCase(Area_B_CT(i)) & "</ITRForm:LocalityOrArea>"
'Else
'writeXML "<ITRForm:LocalityOrArea/>"
'End If
'
'If Town_B_CT(i) <> "" Then
'writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(Town_B_CT(i)) & "</ITRForm:CityOrTownOrDistrict>"
'Else
'writeXML "<ITRForm:CityOrTownOrDistrict/>"
'End If
'
'If StateCode_B_CT(i) <> "" And StateCode_B_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:StateCode>" & UCase(StateCode_B_CT(i)) & "</ITRForm:StateCode>"
'Else
'writeXML "<ITRForm:StateCode/>"
'End If
'
'If CountryCode_B_CT(i) <> "" And CountryCode_B_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:CountryCode>" & UCase(CountryCode_B_CT(i)) & "</ITRForm:CountryCode>"
'Else
'writeXML "<ITRForm:CountryCode/>"
'End If
'
'If PinCode_B_CT(i) <> "" Then
'writeXML "<ITRForm:PinCode>" & UCase(PinCode_B_CT(i)) & "</ITRForm:PinCode>"
'Else
'writeXML "<ITRForm:PinCode/>"
'End If
'
'If ZipCode_B_CT(i) <> "" Then
'writeXML "<ITRForm:ZipCode>" & UCase(ZipCode_B_CT(i)) & "</ITRForm:ZipCode>"
'Else
'writeXML "<ITRForm:ZipCode/>"
'End If
'
'writeXML "</ITRForm:AddressCT>"
'
'If PAN_B_CT(i) <> "" Then
'writeXML "<ITRForm:Pan>" & UCase(PAN_B_CT(i)) & "</ITRForm:Pan>"
'Else
'writeXML "<ITRForm:Pan/>"
'End If
'
'If Nature_B_CT(i) <> "" Then
'writeXML "<ITRForm:NatureOfTrans>" & UCase(Nature_B_CT(i)) & "</ITRForm:NatureOfTrans>"
'Else
'writeXML "<ITRForm:NatureOfTrans/>"
'End If
'
'If Amount_B_CT(i) <> "" Then
'writeXML "<ITRForm:AggAmount>" & UCase(Amount_B_CT(i)) & "</ITRForm:AggAmount>"
'Else
'writeXML "<ITRForm:AggAmount/>"
'End If
'
'writeXML "</ITRForm:ReceiptExcSec269ST>"
'
'UpdateProgressBar
'Next
'
'
'
'End If
'
'If Not IsEmpty(Name_C_CT) And UBound(Name_C_CT) > 0 Then
'
'subProcCaption = "Schedule CTC"
'noOfProcessSub = UBound(Name_C_CT)
'For i = 1 To UBound(Name_C_CT)
'
'  writeXML "<ITRForm:PaymntExcSec269ST>"
'
'If Name_C_CT(i) <> "" Then
'writeXML "<ITRForm:Name>" & UCase(Name_C_CT(i)) & "</ITRForm:Name>"
' Else
' writeXML "<ITRForm:Name/>"
'End If
'
' writeXML "<ITRForm:AddressCT>"
'
'If Door_C_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceNo>" & UCase(Door_C_CT(i)) & "</ITRForm:ResidenceNo>"
'Else
'writeXML "<ITRForm:ResidenceNo/>"
'End If
'
'If Village_C_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceName>" & UCase(Village_C_CT(i)) & "</ITRForm:ResidenceName>"
'Else
'writeXML "<ITRForm:ResidenceName/>"
'End If
'
'If Road_C_CT(i) <> "" Then
'writeXML "<ITRForm:RoadOrStreet>" & UCase(Road_C_CT(i)) & "</ITRForm:RoadOrStreet>"
'Else
'writeXML "<ITRForm:RoadOrStreet/>"
'End If
'
'If Area_C_CT(i) <> "" Then
'writeXML "<ITRForm:LocalityOrArea>" & UCase(Area_C_CT(i)) & "</ITRForm:LocalityOrArea>"
'Else
'writeXML "<ITRForm:LocalityOrArea/>"
'End If
'
'If Town_C_CT(i) <> "" Then
'writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(Town_C_CT(i)) & "</ITRForm:CityOrTownOrDistrict>"
'Else
'writeXML "<ITRForm:CityOrTownOrDistrict/>"
'End If
'
'If StateCode_C_CT(i) <> "" And StateCode_C_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:StateCode>" & UCase(StateCode_C_CT(i)) & "</ITRForm:StateCode>"
'Else
'writeXML "<ITRForm:StateCode/>"
'End If
'
'If CountryCode_C_CT(i) <> "" And CountryCode_C_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:CountryCode>" & UCase(CountryCode_C_CT(i)) & "</ITRForm:CountryCode>"
'Else
'writeXML "<ITRForm:CountryCode/>"
'End If
'
'If PinCode_C_CT(i) <> "" Then
'writeXML "<ITRForm:PinCode>" & UCase(PinCode_C_CT(i)) & "</ITRForm:PinCode>"
'Else
'writeXML "<ITRForm:PinCode/>"
'End If
'
'If ZipCode_C_CT(i) <> "" Then
'writeXML "<ITRForm:ZipCode>" & UCase(ZipCode_C_CT(i)) & "</ITRForm:ZipCode>"
'Else
'writeXML "<ITRForm:ZipCode/>"
'End If
'
'writeXML "</ITRForm:AddressCT>"
'
'If PAN_C_CT(i) <> "" Then
'writeXML "<ITRForm:Pan>" & UCase(PAN_C_CT(i)) & "</ITRForm:Pan>"
'Else
'writeXML "<ITRForm:Pan/>"
'End If
'
'
'If Nature_C_CT(i) <> "" Then
'writeXML "<ITRForm:NatureOfTrans>" & UCase(Nature_C_CT(i)) & "</ITRForm:NatureOfTrans>"
'Else
'writeXML "<ITRForm:NatureOfTrans/>"
'End If
'
'If Amount_C_CT(i) <> "" Then
'writeXML "<ITRForm:AggAmount>" & UCase(Amount_C_CT(i)) & "</ITRForm:AggAmount>"
'Else
'writeXML "<ITRForm:AggAmount/>"
'End If
'
'writeXML "</ITRForm:PaymntExcSec269ST>"
'
'UpdateProgressBar
'Next
'
'
'
'End If
'
'
'If Not IsEmpty(Name_D_CT) And UBound(Name_D_CT) > 0 Then
'
'subProcCaption = "Schedule CTD"
'noOfProcessSub = UBound(Name_D_CT)
'For i = 1 To UBound(Name_D_CT)
'
'  writeXML "<ITRForm:RepaymntMadeSec269T>"
'
'If Name_D_CT(i) <> "" Then
'writeXML "<ITRForm:Name>" & UCase(Name_D_CT(i)) & "</ITRForm:Name>"
' Else
' writeXML "<ITRForm:Name/>"
'End If
'
' writeXML "<ITRForm:AddressCT>"
'
'If Door_D_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceNo>" & UCase(Door_D_CT(i)) & "</ITRForm:ResidenceNo>"
'Else
'writeXML "<ITRForm:ResidenceNo/>"
'End If
'
'If Village_D_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceName>" & UCase(Village_D_CT(i)) & "</ITRForm:ResidenceName>"
'Else
'writeXML "<ITRForm:ResidenceName/>"
'End If
'
'If Road_D_CT(i) <> "" Then
'writeXML "<ITRForm:RoadOrStreet>" & UCase(Road_D_CT(i)) & "</ITRForm:RoadOrStreet>"
'Else
'writeXML "<ITRForm:RoadOrStreet/>"
'End If
'
'If Area_D_CT(i) <> "" Then
'writeXML "<ITRForm:LocalityOrArea>" & UCase(Area_D_CT(i)) & "</ITRForm:LocalityOrArea>"
'Else
'writeXML "<ITRForm:LocalityOrArea/>"
'End If
'
'If Town_D_CT(i) <> "" Then
'writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(Town_D_CT(i)) & "</ITRForm:CityOrTownOrDistrict>"
'Else
'writeXML "<ITRForm:CityOrTownOrDistrict/>"
'End If
'
'If StateCode_D_CT(i) <> "" And StateCode_D_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:StateCode>" & UCase(StateCode_D_CT(i)) & "</ITRForm:StateCode>"
'Else
'writeXML "<ITRForm:StateCode/>"
'End If
'
'If CountryCode_D_CT(i) <> "" And CountryCode_D_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:CountryCode>" & UCase(CountryCode_D_CT(i)) & "</ITRForm:CountryCode>"
'Else
'writeXML "<ITRForm:CountryCode/>"
'End If
'
'If PinCode_D_CT(i) <> "" Then
'writeXML "<ITRForm:PinCode>" & UCase(PinCode_D_CT(i)) & "</ITRForm:PinCode>"
'Else
'writeXML "<ITRForm:PinCode/>"
'End If
'
'If ZipCode_D_CT(i) <> "" Then
'writeXML "<ITRForm:ZipCode>" & UCase(ZipCode_D_CT(i)) & "</ITRForm:ZipCode>"
'Else
'writeXML "<ITRForm:ZipCode/>"
'End If
'
'writeXML "</ITRForm:AddressCT>"
'
'If PAN_D_CT(i) <> "" Then
'writeXML "<ITRForm:Pan>" & UCase(PAN_D_CT(i)) & "</ITRForm:Pan>"
'Else
'writeXML "<ITRForm:Pan/>"
'End If
'
'If Nature_D_CT(i) <> "" Then
'writeXML "<ITRForm:NatureOfTrans>" & UCase(Nature_D_CT(i)) & "</ITRForm:NatureOfTrans>"
'Else
'writeXML "<ITRForm:NatureOfTrans/>"
'End If
'
'If Amount_D_CT(i) <> "" Then
'writeXML "<ITRForm:AggAmount>" & UCase(Amount_D_CT(i)) & "</ITRForm:AggAmount>"
'Else
'writeXML "<ITRForm:AggAmount/>"
'End If
'
'writeXML "</ITRForm:RepaymntMadeSec269T>"
'
'UpdateProgressBar
'Next
'
'End If
'
'If Not IsEmpty(Name_E_CT) And UBound(Name_E_CT) > 0 Then
'
'subProcCaption = "Schedule CTE"
'noOfProcessSub = UBound(Name_E_CT)
'For i = 1 To UBound(Name_E_CT)
'
' writeXML "<ITRForm:RepaymntRcvdSec269T>"
'
'If Name_E_CT(i) <> "" Then
'writeXML "<ITRForm:Name>" & UCase(Name_E_CT(i)) & "</ITRForm:Name>"
' Else
' writeXML "<ITRForm:Name/>"
'End If
'
' writeXML "<ITRForm:AddressCT>"
'
'If Door_E_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceNo>" & UCase(Door_E_CT(i)) & "</ITRForm:ResidenceNo>"
'Else
'writeXML "<ITRForm:ResidenceNo/>"
'End If
'
'If Village_E_CT(i) <> "" Then
'writeXML "<ITRForm:ResidenceName>" & UCase(Village_E_CT(i)) & "</ITRForm:ResidenceName>"
'Else
'writeXML "<ITRForm:ResidenceName/>"
'End If
'
'If Road_E_CT(i) <> "" Then
'writeXML "<ITRForm:RoadOrStreet>" & UCase(Road_E_CT(i)) & "</ITRForm:RoadOrStreet>"
'Else
'writeXML "<ITRForm:RoadOrStreet/>"
'End If
'
'If Area_E_CT(i) <> "" Then
'writeXML "<ITRForm:LocalityOrArea>" & UCase(Area_E_CT(i)) & "</ITRForm:LocalityOrArea>"
'Else
'writeXML "<ITRForm:LocalityOrArea/>"
'End If
'
'If Town_E_CT(i) <> "" Then
'writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(Town_E_CT(i)) & "</ITRForm:CityOrTownOrDistrict>"
'Else
'writeXML "<ITRForm:CityOrTownOrDistrict/>"
'End If
'
'If StateCode_E_CT(i) <> "" And StateCode_E_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:StateCode>" & UCase(StateCode_E_CT(i)) & "</ITRForm:StateCode>"
'Else
'writeXML "<ITRForm:StateCode/>"
'End If
'
'If CountryCode_E_CT(i) <> "" And CountryCode_E_CT(i) <> "(Select)" Then
'writeXML "<ITRForm:CountryCode>" & UCase(CountryCode_E_CT(i)) & "</ITRForm:CountryCode>"
'Else
'writeXML "<ITRForm:CountryCode/>"
'End If
'
'If PinCode_E_CT(i) <> "" Then
'writeXML "<ITRForm:PinCode>" & UCase(PinCode_E_CT(i)) & "</ITRForm:PinCode>"
'Else
'writeXML "<ITRForm:PinCode/>"
'End If
'
'If ZipCode_E_CT(i) <> "" Then
'writeXML "<ITRForm:ZipCode>" & UCase(ZipCode_E_CT(i)) & "</ITRForm:ZipCode>"
'Else
'writeXML "<ITRForm:ZipCode/>"
'End If
'
'writeXML "</ITRForm:AddressCT>"
'
'If PAN_E_CT(i) <> "" Then
'writeXML "<ITRForm:Pan>" & UCase(PAN_E_CT(i)) & "</ITRForm:Pan>"
'Else
'writeXML "<ITRForm:Pan/>"
'End If
'
'If Nature_E_CT(i) <> "" Then
'writeXML "<ITRForm:NatureOfTrans>" & UCase(Nature_E_CT(i)) & "</ITRForm:NatureOfTrans>"
'Else
'writeXML "<ITRForm:NatureOfTrans/>"
'End If
'
'If Amount_E_CT(i) <> "" Then
'writeXML "<ITRForm:AggAmount>" & UCase(Amount_E_CT(i)) & "</ITRForm:AggAmount>"
'Else
'writeXML "<ITRForm:AggAmount/>"
'End If
'
'writeXML "</ITRForm:RepaymntRcvdSec269T>"
'
'UpdateProgressBar
'Next
'
'End If
'writeXML "</ITRForm:ScheduleCT>"
'End If
'
'End Function
'
'Function XMLPART_GST()
'        writeXML "              <ITRForm:ScheduleGST>"
'        If GST_TotalAmount = "" Then
'            writeXML "          <ITRForm:TotAmtExpenditure>0</ITRForm:TotAmtExpenditure>"
'        Else
'            writeXML "          <ITRForm:TotAmtExpenditure>" & UCase(GST_TotalAmount) & "</ITRForm:TotAmtExpenditure>"
'        End If
'
'        If GST_ServiceExempt = "" Then
'            writeXML "          <ITRForm:RelGoodsServExemptFrmGST>0</ITRForm:RelGoodsServExemptFrmGST>"
'        Else
'            writeXML "          <ITRForm:RelGoodsServExemptFrmGST>" & UCase(GST_ServiceExempt) & "</ITRForm:RelGoodsServExemptFrmGST>"
'        End If
'
'        If GST_EntitiesFailed = "" Then
'            writeXML "          <ITRForm:EntitiesFallUndCompScheme>0</ITRForm:EntitiesFallUndCompScheme>"
'        Else
'            writeXML "          <ITRForm:EntitiesFallUndCompScheme>" & UCase(GST_EntitiesFailed) & "</ITRForm:EntitiesFallUndCompScheme>"
'        End If
'
'        If GST_EntitiesRegistered = "" Then
'            writeXML "          <ITRForm:OthRegisteredEntities>0</ITRForm:OthRegisteredEntities>"
'        Else
'            writeXML "          <ITRForm:OthRegisteredEntities>" & UCase(GST_EntitiesRegistered) & "</ITRForm:OthRegisteredEntities>"
'        End If
'
'        If GST_EntitiesPayment = "" Then
'            writeXML "          <ITRForm:TotPaymntRegEntities>0</ITRForm:TotPaymntRegEntities>"
'        Else
'            writeXML "          <ITRForm:TotPaymntRegEntities>" & UCase(GST_EntitiesPayment) & "</ITRForm:TotPaymntRegEntities>"
'        End If
'
'        If GST_NotRegisteredEntities = "" Then
'            writeXML "          <ITRForm:PaymntMadeNotRegUndGST>0</ITRForm:PaymntMadeNotRegUndGST>"
'        Else
'            writeXML "          <ITRForm:PaymntMadeNotRegUndGST>" & UCase(GST_NotRegisteredEntities) & "</ITRForm:PaymntMadeNotRegUndGST>"
'        End If
'
'        writeXML "         </ITRForm:ScheduleGST>"
'End Function
'
'Function XMLPART_FD()
'        writeXML "              <ITRForm:ScheduleFD>"
'        writeXML "<ITRForm:PaymntMadeOnCapitalAcc>"
'
'        If ForeignCapitalPayment = "" Then
'            writeXML "          <ITRForm:InForeignCurrency>0</ITRForm:InForeignCurrency>"
'        Else
'            writeXML "          <ITRForm:InForeignCurrency>" & UCase(ForeignCapitalPayment) & "</ITRForm:InForeignCurrency>"
'        End If
'
'        If TotalCapitalPayment = "" Then
'            writeXML "          <ITRForm:TotalAmount>0</ITRForm:TotalAmount>"
'        Else
'            writeXML "          <ITRForm:TotalAmount>" & UCase(TotalCapitalPayment) & "</ITRForm:TotalAmount>"
'        End If
'
'        writeXML "</ITRForm:PaymntMadeOnCapitalAcc>"
'
'        writeXML "<ITRForm:PaymntMadeOnRevenueAcc>"
'
'        If ForeignRevenuePayment = "" Then
'            writeXML "          <ITRForm:InForeignCurrency>0</ITRForm:InForeignCurrency>"
'        Else
'            writeXML "          <ITRForm:InForeignCurrency>" & UCase(ForeignRevenuePayment) & "</ITRForm:InForeignCurrency>"
'        End If
'
'        If TotalRevenuePayment = "" Then
'            writeXML "          <ITRForm:TotalAmount>0</ITRForm:TotalAmount>"
'        Else
'            writeXML "          <ITRForm:TotalAmount>" & UCase(TotalRevenuePayment) & "</ITRForm:TotalAmount>"
'        End If
'
'        writeXML "</ITRForm:PaymntMadeOnRevenueAcc>"
'
'        writeXML "<ITRForm:ReceiptsOnCapitalAcc>"
'
'        If ForeignCapitalReceipts = "" Then
'            writeXML "          <ITRForm:InForeignCurrency>0</ITRForm:InForeignCurrency>"
'        Else
'            writeXML "          <ITRForm:InForeignCurrency>" & UCase(ForeignCapitalReceipts) & "</ITRForm:InForeignCurrency>"
'        End If
'
'        If TotalCapitalReceipts = "" Then
'            writeXML "          <ITRForm:TotalAmount>0</ITRForm:TotalAmount>"
'        Else
'            writeXML "          <ITRForm:TotalAmount>" & UCase(TotalCapitalReceipts) & "</ITRForm:TotalAmount>"
'        End If
'
'        writeXML "</ITRForm:ReceiptsOnCapitalAcc>"
'
'        writeXML "<ITRForm:ReceiptsOnRevenueAcc>"
'
'        If ForeignRevenueReceipts = "" Then
'            writeXML "          <ITRForm:InForeignCurrency>0</ITRForm:InForeignCurrency>"
'        Else
'            writeXML "          <ITRForm:InForeignCurrency>" & UCase(ForeignRevenueReceipts) & "</ITRForm:InForeignCurrency>"
'        End If
'
'        If TotalRevenueReceipts = "" Then
'            writeXML "          <ITRForm:TotalAmount>0</ITRForm:TotalAmount>"
'        Else
'            writeXML "          <ITRForm:TotalAmount>" & UCase(TotalRevenueReceipts) & "</ITRForm:TotalAmount>"
'        End If
'
'        writeXML "</ITRForm:ReceiptsOnRevenueAcc>"
'
'
'        writeXML "         </ITRForm:ScheduleFD>"
'End Function

'Function XMLPART_CS()
'        writeXML "     <ITRForm:ScheduleCS>"
'        writeXML "<ITRForm:OpeningCashInHand>"
'
'        If OpeningInhandCash = "" Then
'            writeXML "          <ITRForm:Cash>0</ITRForm:Cash>"
'        Else
'            writeXML "          <ITRForm:Cash>" & UCase(OpeningInhandCash) & "</ITRForm:Cash>"
'        End If
'
'        If OpeningInhandBearerCheque = "" Then
'            writeXML "          <ITRForm:BearerCheque>0</ITRForm:BearerCheque>"
'        Else
'            writeXML "          <ITRForm:BearerCheque>" & UCase(OpeningInhandBearerCheque) & "</ITRForm:BearerCheque>"
'        End If
'
'        If OpeningInhandTotal = "" Then
'            writeXML "          <ITRForm:Total>0</ITRForm:Total>"
'        Else
'            writeXML "          <ITRForm:Total>" & UCase(OpeningInhandTotal) & "</ITRForm:Total>"
'        End If
'
'        writeXML "</ITRForm:OpeningCashInHand>"
'
'        writeXML "<ITRForm:CashRcvdDurYearOnRevAcc>"
'
'        If ReceivedRevenueCash = "" Then
'            writeXML "          <ITRForm:Cash>0</ITRForm:Cash>"
'        Else
'            writeXML "          <ITRForm:Cash>" & UCase(ReceivedRevenueCash) & "</ITRForm:Cash>"
'        End If
'
'        If ReceivedRevenueBearerCheque = "" Then
'            writeXML "          <ITRForm:BearerCheque>0</ITRForm:BearerCheque>"
'        Else
'            writeXML "          <ITRForm:BearerCheque>" & UCase(ReceivedRevenueBearerCheque) & "</ITRForm:BearerCheque>"
'        End If
'
'        If ReceivedRevenueTotal = "" Then
'            writeXML "          <ITRForm:Total>0</ITRForm:Total>"
'        Else
'            writeXML "          <ITRForm:Total>" & UCase(ReceivedRevenueTotal) & "</ITRForm:Total>"
'        End If
'
'        writeXML "</ITRForm:CashRcvdDurYearOnRevAcc>"
'
'        writeXML "<ITRForm:CashRcvdDurYearOnCapitalAcc>"
'
'        If ReceivedCapitalCash = "" Then
'            writeXML "          <ITRForm:Cash>0</ITRForm:Cash>"
'        Else
'            writeXML "          <ITRForm:Cash>" & UCase(ReceivedCapitalCash) & "</ITRForm:Cash>"
'        End If
'
'        If ReceivedCapitalBearerCheque = "" Then
'            writeXML "          <ITRForm:BearerCheque>0</ITRForm:BearerCheque>"
'        Else
'            writeXML "          <ITRForm:BearerCheque>" & UCase(ReceivedCapitalBearerCheque) & "</ITRForm:BearerCheque>"
'        End If
'
'        If ReceivedCapitalTotal = "" Then
'            writeXML "          <ITRForm:Total>0</ITRForm:Total>"
'        Else
'            writeXML "          <ITRForm:Total>" & UCase(ReceivedCapitalTotal) & "</ITRForm:Total>"
'        End If
'
'        writeXML "</ITRForm:CashRcvdDurYearOnCapitalAcc>"
'
'        If ReceivedTotal <> "" Then
'            writeXML "          <ITRForm:TotalCashRecvDrngYr>" & UCase(ReceivedTotal) & "</ITRForm:TotalCashRecvDrngYr>"
'        End If
'
'        writeXML "<ITRForm:CashPaidDurYearOnRevAcc>"
'
'        If PaidRevenueCash = "" Then
'            writeXML "          <ITRForm:Cash>0</ITRForm:Cash>"
'        Else
'            writeXML "          <ITRForm:Cash>" & UCase(PaidRevenueCash) & "</ITRForm:Cash>"
'        End If
'
'        If PaidRevenueBearerCheque = "" Then
'            writeXML "          <ITRForm:BearerCheque>0</ITRForm:BearerCheque>"
'        Else
'            writeXML "          <ITRForm:BearerCheque>" & UCase(PaidRevenueBearerCheque) & "</ITRForm:BearerCheque>"
'        End If
'
'        If PaidRevenueTotal = "" Then
'            writeXML "          <ITRForm:Total>0</ITRForm:Total>"
'        Else
'            writeXML "          <ITRForm:Total>" & UCase(PaidRevenueTotal) & "</ITRForm:Total>"
'        End If
'
'        writeXML "</ITRForm:CashPaidDurYearOnRevAcc>"
'
'        writeXML "<ITRForm:CashPaidDurYearOnCapitalAcc>"
'
'        If PaidCapitalCash = "" Then
'            writeXML "          <ITRForm:Cash>0</ITRForm:Cash>"
'        Else
'            writeXML "          <ITRForm:Cash>" & UCase(PaidCapitalCash) & "</ITRForm:Cash>"
'        End If
'
'        If PaidCapitalBearerCheque = "" Then
'            writeXML "          <ITRForm:BearerCheque>0</ITRForm:BearerCheque>"
'        Else
'            writeXML "          <ITRForm:BearerCheque>" & UCase(PaidCapitalBearerCheque) & "</ITRForm:BearerCheque>"
'        End If
'
'        If PaidCapitalTotal = "" Then
'            writeXML "          <ITRForm:Total>0</ITRForm:Total>"
'        Else
'            writeXML "          <ITRForm:Total>" & UCase(PaidCapitalTotal) & "</ITRForm:Total>"
'        End If
'
'        writeXML "</ITRForm:CashPaidDurYearOnCapitalAcc>"
'
'        If PaidTotal <> "" Then
'            writeXML "          <ITRForm:TotalCashPaidDrngYr>" & UCase(PaidTotal) & "</ITRForm:TotalCashPaidDrngYr>"
'        End If
'
'        writeXML "<ITRForm:ClosingCashInHand>"
'
'        If ClosingInhandCash = "" Then
'            writeXML "          <ITRForm:Cash>0</ITRForm:Cash>"
'        Else
'            writeXML "          <ITRForm:Cash>" & UCase(ClosingInhandCash) & "</ITRForm:Cash>"
'        End If
'
'        If ClosingInhandBearerCheque = "" Then
'            writeXML "          <ITRForm:BearerCheque>0</ITRForm:BearerCheque>"
'        Else
'            writeXML "          <ITRForm:BearerCheque>" & UCase(ClosingInhandBearerCheque) & "</ITRForm:BearerCheque>"
'        End If
'
'        If ClosingInhandTotal = "" Then
'            writeXML "          <ITRForm:Total>0</ITRForm:Total>"
'        Else
'            writeXML "          <ITRForm:Total>" & UCase(ClosingInhandTotal) & "</ITRForm:Total>"
'        End If
'
'        writeXML "</ITRForm:ClosingCashInHand>"
'
'        writeXML "         </ITRForm:ScheduleCS>"
'End Function
Function ScheduleICDS()
noOfProcessSub = 4
subProcCaption = "Schedule ICDS"

If Sheet42.Range("ICDS.FilledFlag") Then
writeXML "<ITRForm:ScheduleICDS>"
If ICDS_AccPolicies <> "" Then writeXML "<ITRForm:AccPolicyAmt>" & ICDS_AccPolicies & "</ITRForm:AccPolicyAmt>"
If ICDS_ValuationInv <> "" Then writeXML "<ITRForm:InventoriesValue>" & ICDS_ValuationInv & "</ITRForm:InventoriesValue>"
UpdateProgressBar
If ICDS_ConsContracts <> "" Then writeXML "<ITRForm:ConstContractsAmt>" & ICDS_ConsContracts & "</ITRForm:ConstContractsAmt>"
If ICDS_RevenueRecog <> "" Then writeXML "<ITRForm:RevenueRcgAmt>" & ICDS_RevenueRecog & "</ITRForm:RevenueRcgAmt>"
If ICDS_TangibleFixAssests <> "" Then writeXML "<ITRForm:TangibleFixedAsset>" & ICDS_TangibleFixAssests & "</ITRForm:TangibleFixedAsset>"
UpdateProgressBar
If ICDS_ChngRates <> "" Then writeXML "<ITRForm:ForeignExgRates>" & ICDS_ChngRates & "</ITRForm:ForeignExgRates>"
If ICDS_Govgrants <> "" Then writeXML "<ITRForm:GovtGrants>" & ICDS_Govgrants & "</ITRForm:GovtGrants>"
UpdateProgressBar
If ICDS_Securities <> "" Then writeXML "<ITRForm:Securities>" & ICDS_Securities & "</ITRForm:Securities>"
If ICDS_BorrowingCosts <> "" Then writeXML "<ITRForm:BorrowingCosts>" & ICDS_BorrowingCosts & "</ITRForm:BorrowingCosts>"
If ICDS_ProvLiability <> "" Then writeXML "<ITRForm:ProvAssets>" & ICDS_ProvLiability & "</ITRForm:ProvAssets>"
If ICDS_Total <> "" Then writeXML "<ITRForm:TotalNetAmt>" & ICDS_Total & "</ITRForm:TotalNetAmt>"
If ICDS_Totalb <> "" Then writeXML "<ITRForm:TotalNetAmtNegative>" & ICDS_Totalb & "</ITRForm:TotalNetAmtNegative>"
UpdateProgressBar
writeXML "</ITRForm:ScheduleICDS>"
Else
UpdateProgressBar
UpdateProgressBar
UpdateProgressBar
UpdateProgressBar
End If
End Function

Function Schedule10A()
  Dim i, j As Long
'    subProcCaption = "Schedule 10A"
'    If (TotalDedUs10Sub_SEZA10 > 0) Then
'    setTblinfo_SEZA10_2
'            writeXML "      <ITRForm:Schedule10A>"
'            noOfProcessSub = UBound(DedFromUndertaking_SEZA10)
'            If Not IsEmpty(DedFromUndertaking_SEZA10) And UBound(DedFromUndertaking_SEZA10) > 0 Then
'                writeXML "          <ITRForm:DeductSEZ>"
'                writeXML "              <ITRForm:DedUs10Detail>"
'                writeXML "              <ITRForm:Undertaking>"
'                For j = 1 To end_SEZA10_2
'                    writeXML "<ITRForm:DedFromUndertakingWithAy>"
'                    If AssmtYrUnit_SEZ10(j) <> "" Then
'                        writeXML "                   <ITRForm:AssmtYrUnit>" & Trim(AssmtYrUnit_SEZ10(j)) & "</ITRForm:AssmtYrUnit>"
'                    Else
'                        writeXML "                   <ITRForm:AssmtYrUnit>0000-00</ITRForm:AssmtYrUnit>"
'                    End If
'
'                    If DedFromUndertaking_SEZA10(j) <> "" Then
'                        writeXML "                   <ITRForm:DedUs10Sub>" & UCase(DedFromUndertaking_SEZA10(j)) & "</ITRForm:DedUs10Sub>"
'                    Else
'                        writeXML "                   <ITRForm:DedUs10Sub>" & UCase(DefaultDedFromUndertaking_SEZA10) & "</ITRForm:DedUs10Sub>"
'                    End If
'                    writeXML "</ITRForm:DedFromUndertakingWithAy>"
'                    UpdateProgressBar
'                Next
'                writeXML "              </ITRForm:Undertaking>"
'                If TotalDedUs10Sub_SEZA10 <> "" Then
'                    writeXML "                   <ITRForm:TotalDedUs10Sub>" & UCase(TotalDedUs10Sub_SEZA10) & "</ITRForm:TotalDedUs10Sub>"
'                Else
'                    writeXML "                   <ITRForm:TotalDedUs10Sub>" & UCase(DefaultTotalDedUs10Sub_SEZA10) & "</ITRForm:TotalDedUs10Sub>"
'                End If
'                writeXML "              </ITRForm:DedUs10Detail>"
'                writeXML "          </ITRForm:DeductSEZ>"
'            End If
'            writeXML "      </ITRForm:Schedule10A>"
'    End If

    If TotalDedUs10Sub_AA10 > 0 Then
    subProcCaption = "Schedule 10AA"
    setTblinfo_AA10_2
        writeXML "      <ITRForm:Schedule10AA>"
        If Not IsEmpty(DedFromUndertaking_AA10) And UBound(DedFromUndertaking_AA10) > 0 Then
            writeXML "          <ITRForm:DeductSEZ>"
            writeXML "              <ITRForm:DedUs10Detail>"
            writeXML "              <ITRForm:Undertaking>"
            noOfProcessSub = end_AA10_2
            For j = 1 To end_AA10_2
                writeXML "<ITRForm:DedFromUndertakingWithAy>"
                If AssmtYrUnit_AA10(j) <> "" Then
                    writeXML "                   <ITRForm:AssmtYrUnit>" & Trim(AssmtYrUnit_AA10(j)) & "</ITRForm:AssmtYrUnit>"
                Else
                    writeXML "                   <ITRForm:AssmtYrUnit>0000-00</ITRForm:AssmtYrUnit>"
                End If

                If DedFromUndertaking_AA10(j) <> "" Then
                    writeXML "                   <ITRForm:DedUs10Sub>" & UCase(DedFromUndertaking_AA10(j)) & "</ITRForm:DedUs10Sub>"
                Else
                    writeXML "                   <ITRForm:DedUs10Sub>" & UCase(DefaultDedFromUndertaking_AA10) & "</ITRForm:DedUs10Sub>"
                End If
                writeXML "</ITRForm:DedFromUndertakingWithAy>"
                UpdateProgressBar
            Next
            writeXML "              </ITRForm:Undertaking>"
            If TotalDedUs10Sub_AA10 <> "" Then
                writeXML "                   <ITRForm:TotalDedUs10Sub>" & UCase(TotalDedUs10Sub_AA10) & "</ITRForm:TotalDedUs10Sub>"
            Else
                writeXML "                   <ITRForm:TotalDedUs10Sub>" & UCase(DefaultTotalDedUs10Sub_AA10) & "</ITRForm:TotalDedUs10Sub>"
            End If
            writeXML "              </ITRForm:DedUs10Detail>"
            writeXML "          </ITRForm:DeductSEZ>"
        End If
        writeXML "      </ITRForm:Schedule10AA>"
    End If
End Function

Function Schedule80G()
 Dim i As Long

'If TotalDonationsUs80G_Per5080G <> 0 Then
    subProcCaption = "Schedule 80GA"
        writeXML "      <ITRForm:Schedule80G>"
        noOfProcessSub = UBound(DoneeName_Per10080G)
        If Not IsEmpty(DoneeName_Per10080G) And UBound(DoneeName_Per10080G) > 0 Then
            writeXML "          <ITRForm:Don100Percent>"
            For i = 1 To UBound(DoneeName_Per10080G)
                If DoneeName_Per10080G(i) = "" Then
                    Exit For
                End If
                writeXML "               <ITRForm:DoneeWithPan>"
                If DoneeName_Per10080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneeWithPanName>" & UCase(DoneeName_Per10080G(i)) & "</ITRForm:DoneeWithPanName>"
                Else
                    writeXML "                   <ITRForm:DoneeWithPanName/>"
                End If

                If DoneePAN_Per10080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneePAN>" & UCase(DoneePAN_Per10080G(i)) & "</ITRForm:DoneePAN>"
                Else
                    writeXML "                   <ITRForm:DoneePAN/>"
                End If

                writeXML "                  <ITRForm:AddressDetail>"
                If AddrDetail_Per10080G(i) <> "" Then
                    writeXML "                       <ITRForm:AddrDetail>" & UCase(AddrDetail_Per10080G(i)) & "</ITRForm:AddrDetail>"
                 Else
                    writeXML "                       <ITRForm:AddrDetail/>"
                End If

                If CityOrTownOrDistrict_Per10080G(i) <> "" Then
                    writeXML "                       <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_Per10080G(i)) & "</ITRForm:CityOrTownOrDistrict>"
                Else
                    writeXML "                       <ITRForm:CityOrTownOrDistrict/>"
                End If

                If StateCode_Per10080G(i) <> "" Then
                    writeXML "                       <ITRForm:StateCode>" & UCase(StateCode_Per10080G(i)) & "</ITRForm:StateCode>"
                Else
                    writeXML "                       <ITRForm:StateCode/>"
                End If

                If PinCode_Per10080G(i) <> "" Then
                    writeXML "                       <ITRForm:PinCode>" & UCase(PinCode_Per10080G(i)) & "</ITRForm:PinCode>"
                Else
                    writeXML "                       <ITRForm:PinCode/>"
                End If
                writeXML "                  </ITRForm:AddressDetail>"

                If DonationAmt_Per10080G_InCash(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtCash>" & UVCase(DonationAmt_Per10080G_InCash(i)) & "</ITRForm:DonationAmtCash>"
                Else
                    writeXML "                   <ITRForm:DonationAmtCash>0</ITRForm:DonationAmtCash>"
                End If

                If DonationAmt_Per10080G(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtOtherMode>" & UVCase(DonationAmt_Per10080G(i)) & "</ITRForm:DonationAmtOtherMode>"
                Else
                    writeXML "                   <ITRForm:DonationAmtOtherMode>0</ITRForm:DonationAmtOtherMode>"
                End If

                If Trim(Range("Per10080G.TotalDonationAmt").Cells.item(i).value) <> "" Then
                writeXML "<ITRForm:DonationAmt>" & UCase(Range("Per10080G.TotalDonationAmt").Cells.item(i).value) & "</ITRForm:DonationAmt>"
                Else
                writeXML "<ITRForm:DonationAmt>0</ITRForm:DonationAmt>"
                End If


               If EligibleAmt_Per10080G(i) <> "" Then
                    writeXML "                   <ITRForm:EligibleDonationAmt>" & UCase(EligibleAmt_Per10080G(i)) & "</ITRForm:EligibleDonationAmt>"
                Else
                    writeXML "                   <ITRForm:EligibleDonationAmt>0</ITRForm:EligibleDonationAmt>"
                End If
                writeXML "               </ITRForm:DoneeWithPan>"
                UpdateProgressBar

            Next

            If Trim(Range("Per10080G.TotDon100Percent_InCash").value) <> "" Then
                writeXML "<ITRForm:TotDon100PercentCash>" & UCase(Range("Per10080G.TotDon100Percent_InCash").value) & "</ITRForm:TotDon100PercentCash>"
                Else
                writeXML "<ITRForm:TotDon100PercentCash>0</ITRForm:TotDon100PercentCash>"
            End If


            If Trim(Range("Per10080G.TotDon100Percent").value) <> "" Then
                writeXML "               <ITRForm:TotDon100PercentOtherMode>" & UCase(Trim(Range("Per10080G.TotDon100Percent").value)) & "</ITRForm:TotDon100PercentOtherMode>"
            Else
                writeXML "               <ITRForm:TotDon100PercentOtherMode>0</ITRForm:TotDon100PercentOtherMode>"
            End If


            If (Range("Per10080G.TotDon100Percent_Total").value) <> "" Then
                writeXML "<ITRForm:TotDon100Percent>" & UCase(Range("Per10080G.TotDon100Percent_Total").value) & "</ITRForm:TotDon100Percent>"
                Else
                writeXML "<ITRForm:TotDon100Percent>0</ITRForm:TotDon100Percent>"
            End If



            If Range("Per10080G.TotElig100Percent").value <> "" Then
                writeXML "               <ITRForm:TotEligibleDon100Percent>" & UCase(Range("Per10080G.TotElig100Percent").value) & "</ITRForm:TotEligibleDon100Percent>"
            Else
                writeXML "               <ITRForm:TotEligibleDon100Percent>0</ITRForm:TotEligibleDon100Percent>"
            End If

            writeXML "          </ITRForm:Don100Percent>"
        End If


             subProcCaption = "Schedule 80GB"
             noOfProcessSub = UBound(DoneeName_PerNO5080G)
        If Not IsEmpty(DoneeName_PerNO5080G) And UBound(DoneeName_PerNO5080G) > 0 Then
            writeXML "          <ITRForm:Don50PercentNoApprReqd>"
            For i = 1 To UBound(DoneeName_PerNO5080G)
                If DoneeName_PerNO5080G(i) = "" Then
                    Exit For
                End If
                writeXML "               <ITRForm:DoneeWithPan>"
                If DoneeName_PerNO5080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneeWithPanName>" & UCase(DoneeName_PerNO5080G(i)) & "</ITRForm:DoneeWithPanName>"
                Else
                    writeXML "                   <ITRForm:DoneeWithPanName/>"
                End If

                If DoneePAN_PerNO5080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneePAN>" & UCase(DoneePAN_PerNO5080G(i)) & "</ITRForm:DoneePAN>"
                Else
                    writeXML "                   <ITRForm:DoneePAN/>"
                End If
                writeXML "                  <ITRForm:AddressDetail>"
                If AddrDetail_PerNO5080G(i) <> "" Then
                    writeXML "                       <ITRForm:AddrDetail>" & UCase(AddrDetail_PerNO5080G(i)) & "</ITRForm:AddrDetail>"
                Else
                    writeXML "                       <ITRForm:AddrDetail/>"
                End If

                If CityOrTownOrDistrict_PerNO5080G(i) <> "" Then
                    writeXML "                       <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_PerNO5080G(i)) & "</ITRForm:CityOrTownOrDistrict>"
                Else
                    writeXML "                       <ITRForm:CityOrTownOrDistrict/>"
                End If

                If StateCode_PerNO5080G(i) <> "" Then
                    writeXML "                       <ITRForm:StateCode>" & UCase(StateCode_PerNO5080G(i)) & "</ITRForm:StateCode>"
                Else
                    writeXML "                       <ITRForm:StateCode/>"
                End If

                If PinCode_PerNO5080G(i) <> "" Then
                    writeXML "                       <ITRForm:PinCode>" & UCase(PinCode_PerNO5080G(i)) & "</ITRForm:PinCode>"
                Else
                    writeXML "                       <ITRForm:PinCode/>"
                End If
                writeXML "                  </ITRForm:AddressDetail>"





                If DonationAmt_PerNO5080G_InCash(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtCash>" & UCase(DonationAmt_PerNO5080G_InCash(i)) & "</ITRForm:DonationAmtCash>"
                Else
                    writeXML "                   <ITRForm:DonationAmtCash>0</ITRForm:DonationAmtCash>"
                End If

                If DonationAmt_PerNO5080G(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtOtherMode>" & UCase(DonationAmt_PerNO5080G(i)) & "</ITRForm:DonationAmtOtherMode>"
                Else
                    writeXML "                   <ITRForm:DonationAmtOtherMode>0</ITRForm:DonationAmtOtherMode>"
                End If

                If Trim(Range("PerNO5080G.TotalDonationAmt").Cells.item(i).value) <> "" Then
                writeXML "<ITRForm:DonationAmt>" & UCase(Range("PerNO5080G.TotalDonationAmt").Cells.item(i).value) & "</ITRForm:DonationAmt>"
                Else
                writeXML "<ITRForm:DonationAmt>0</ITRForm:DonationAmt>"
                End If

                If EligibleAmt_PerNO5080G(i) <> "" Then
                    writeXML "                   <ITRForm:EligibleDonationAmt>" & UCase(EligibleAmt_PerNO5080G(i)) & "</ITRForm:EligibleDonationAmt>"
                Else
                    writeXML "                   <ITRForm:EligibleDonationAmt>0</ITRForm:EligibleDonationAmt>"
                End If
                writeXML "               </ITRForm:DoneeWithPan>"
                UpdateProgressBar
            Next


               If Trim(Range("PerNO5080G.TotDon100Percent_InCash").value) <> "" Then
               writeXML "<ITRForm:TotDon50PercentNoApprReqdCash>" & UCase(Range("PerNO5080G.TotDon100Percent_InCash").value) & "</ITRForm:TotDon50PercentNoApprReqdCash>"
               Else
               writeXML "<ITRForm:TotDon50PercentNoApprReqdCash>0</ITRForm:TotDon50PercentNoApprReqdCash>"
               End If


            If Trim(Range("PerNO5080G.TotDon100Percent").value) <> "" Then
                writeXML "<ITRForm:TotDon50PercentNoApprReqdOtherMode>" & Trim(Range("PerNO5080G.TotDon100Percent").value) & "</ITRForm:TotDon50PercentNoApprReqdOtherMode>"
            Else
                writeXML "<ITRForm:TotDon50PercentNoApprReqdOtherMode>0</ITRForm:TotDon50PercentNoApprReqdOtherMode>"
            End If

            If Trim(Range("PerNO5080G.TotDon100Percent_Total").value) <> "" Then
            writeXML "<ITRForm:TotDon50PercentNoApprReqd>" & UCase(Range("PerNO5080G.TotDon100Percent_Total").value) & "</ITRForm:TotDon50PercentNoApprReqd>"
            Else
            writeXML "<ITRForm:TotDon50PercentNoApprReqd>0</ITRForm:TotDon50PercentNoApprReqd>"
            End If



            If Trim(Range("PerNO5080G.TotElig100Percent").value) <> "" Then
                writeXML "               <ITRForm:TotEligibleDon50Percent>" & UCase(Trim(Range("PerNO5080G.TotElig100Percent").value)) & "</ITRForm:TotEligibleDon50Percent>"
            Else
                writeXML "               <ITRForm:TotEligibleDon50Percent>0</ITRForm:TotEligibleDon50Percent>"
            End If

            writeXML "          </ITRForm:Don50PercentNoApprReqd>"
        End If



        If Not IsEmpty(DoneeName_PerYES10080G) And UBound(DoneeName_PerYES10080G) > 0 Then
        subProcCaption = "Schedule 80GC"
        noOfProcessSub = UBound(DoneeName_PerYES10080G)
            writeXML "          <ITRForm:Don100PercentApprReqd>"
            For i = 1 To UBound(DoneeName_PerYES10080G)
                If DoneeName_PerYES10080G(i) = "" Then
                    Exit For
                End If

                writeXML "               <ITRForm:DoneeWithPan>"
                If DoneeName_PerYES10080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneeWithPanName>" & UCase(DoneeName_PerYES10080G(i)) & "</ITRForm:DoneeWithPanName>"
                Else
                    writeXML "                   <ITRForm:DoneeWithPanName/>"
                End If

                If DoneePAN_PerYES10080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneePAN>" & UCase(DoneePAN_PerYES10080G(i)) & "</ITRForm:DoneePAN>"
                Else
                    writeXML "                   <ITRForm:DoneePAN/>"
                End If

                writeXML "                  <ITRForm:AddressDetail>"
                If AddrDetail_PerYES10080G(i) <> "" Then
                    writeXML "                       <ITRForm:AddrDetail>" & UCase(AddrDetail_PerYES10080G(i)) & "</ITRForm:AddrDetail>"
                Else
                    writeXML "                       <ITRForm:AddrDetail/>"
                End If

                If CityOrTownOrDistrict_PerYES10080G(i) <> "" Then
                    writeXML "                       <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_PerYES10080G(i)) & "</ITRForm:CityOrTownOrDistrict>"
                Else
                    writeXML "                       <ITRForm:CityOrTownOrDistrict/>"
                End If

                If StateCode_PerYES10080G(i) <> "" Then
                    writeXML "                       <ITRForm:StateCode>" & UCase(StateCode_PerYES10080G(i)) & "</ITRForm:StateCode>"
                Else
                    writeXML "                       <ITRForm:StateCode/>"
                End If

                If PinCode_PerYES10080G(i) <> "" Then
                    writeXML "                       <ITRForm:PinCode>" & UCase(PinCode_PerYES10080G(i)) & "</ITRForm:PinCode>"
                Else
                    writeXML "                       <ITRForm:PinCode/>"
                End If
                writeXML "                  </ITRForm:AddressDetail>"




                If DonationAmt_PerYES10080G_InCash(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtCash>" & UCase(DonationAmt_PerYES10080G_InCash(i)) & "</ITRForm:DonationAmtCash>"
                Else
                    writeXML "                   <ITRForm:DonationAmtCash>0</ITRForm:DonationAmtCash>"
                End If

                 If DonationAmt_PerYES10080G(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtOtherMode>" & UCase(DonationAmt_PerYES10080G(i)) & "</ITRForm:DonationAmtOtherMode>"
                Else
                    writeXML "                   <ITRForm:DonationAmtOtherMode>0</ITRForm:DonationAmtOtherMode>"
                End If


                If Trim(Range("PerYES10080G.TotalDonationAmt").Cells.item(i).value) <> "" Then
                writeXML "<ITRForm:DonationAmt>" & UCase(Range("PerYES10080G.TotalDonationAmt").Cells.item(i).value) & "</ITRForm:DonationAmt>"
                Else
                writeXML "<ITRForm:DonationAmt>0</ITRForm:DonationAmt>"
                End If


                If EligibleAmt_PerYES10080G(i) <> "" Then
                    writeXML "                   <ITRForm:EligibleDonationAmt>" & UCase(EligibleAmt_PerYES10080G(i)) & "</ITRForm:EligibleDonationAmt>"
                Else
                    writeXML "                   <ITRForm:EligibleDonationAmt>0</ITRForm:EligibleDonationAmt>"
                End If

                writeXML "               </ITRForm:DoneeWithPan>"
                UpdateProgressBar
            Next

            If Trim(Range("PerYES10080G.TotDon100Percent_InCash").value) <> "" Then
            writeXML "<ITRForm:TotDon100PercentApprReqdCash>" & UCase(Range("PerYES10080G.TotDon100Percent_InCash").value) & "</ITRForm:TotDon100PercentApprReqdCash>"
            Else
            writeXML "<ITRForm:TotDon100PercentApprReqdCash>0</ITRForm:TotDon100PercentApprReqdCash>"
            End If



            If Trim(Range("PerYES10080G.TotDon100Percent").value) <> "" Then
                writeXML "               <ITRForm:TotDon100PercentApprReqdOtherMode>" & Trim(Range("PerYES10080G.TotDon100Percent").value) & "</ITRForm:TotDon100PercentApprReqdOtherMode>"
            Else
                writeXML "               <ITRForm:TotDon100PercentApprReqdOtherMode>0</ITRForm:TotDon100PercentApprReqdOtherMode>"
            End If


            If Trim(Range("PerYES10080G.TotDon100Percent_Total").value) <> "" Then
            writeXML "<ITRForm:TotDon100PercentApprReqd>" & UCase(Range("PerYES10080G.TotDon100Percent_Total").value) & "</ITRForm:TotDon100PercentApprReqd>"
            Else
            writeXML "<ITRForm:TotDon100PercentApprReqd>0</ITRForm:TotDon100PercentApprReqd>"
            End If



            If Trim(Range("PerYES10080G.TotElig100Percent").value) <> "" Then
                writeXML "               <ITRForm:TotEligibleDon100PercentApprReqd>" & UCase(Trim(Range("PerYES10080G.TotElig100Percent").value)) & "</ITRForm:TotEligibleDon100PercentApprReqd>"
            Else
                writeXML "               <ITRForm:TotEligibleDon100PercentApprReqd>0</ITRForm:TotEligibleDon100PercentApprReqd>"
            End If


            writeXML "          </ITRForm:Don100PercentApprReqd>"
        End If


        If Not IsEmpty(DoneeName_Per5080G) And UBound(DoneeName_Per5080G) > 0 Then
        subProcCaption = "Schedule 80GD"
        noOfProcessSub = UBound(DoneeName_Per5080G)
            writeXML "          <ITRForm:Don50PercentApprReqd>"
            For i = 1 To UBound(DoneeName_Per5080G)
                If DoneeName_Per5080G(i) = "" Then
                    Exit For
                End If
                writeXML "               <ITRForm:DoneeWithPan>"

                If DoneeName_Per5080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneeWithPanName>" & UCase(DoneeName_Per5080G(i)) & "</ITRForm:DoneeWithPanName>"
                Else
                    writeXML "                   <ITRForm:DoneeWithPanName/>"
                End If

                If DoneePAN_Per5080G(i) <> "" Then
                    writeXML "                   <ITRForm:DoneePAN>" & UCase(DoneePAN_Per5080G(i)) & "</ITRForm:DoneePAN>"
                Else
                    writeXML "                   <ITRForm:DoneePAN/>"
                End If

                writeXML "                  <ITRForm:AddressDetail>"
                If AddrDetail_Per5080G(i) <> "" Then
                    writeXML "                       <ITRForm:AddrDetail>" & UCase(AddrDetail_Per5080G(i)) & "</ITRForm:AddrDetail>"
                Else
                    writeXML "                       <ITRForm:AddrDetail/>"
                End If

                If CityOrTownOrDistrict_Per5080G(i) <> "" Then
                    writeXML "                       <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_Per5080G(i)) & "</ITRForm:CityOrTownOrDistrict>"
                Else
                    writeXML "                       <ITRForm:CityOrTownOrDistrict/>"
                End If

                If StateCode_Per5080G(i) <> "" Then
                    writeXML "                       <ITRForm:StateCode>" & UCase(StateCode_Per5080G(i)) & "</ITRForm:StateCode>"
                Else
                    writeXML "                       <ITRForm:StateCode/>"
                End If

                If PinCode_Per5080G(i) <> "" Then
                    writeXML "                       <ITRForm:PinCode>" & UCase(PinCode_Per5080G(i)) & "</ITRForm:PinCode>"
                Else
                    writeXML "                       <ITRForm:PinCode/>"
                End If
                writeXML "                  </ITRForm:AddressDetail>"




                If DonationAmt_Per5080G_InCash(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtCash>" & UCase(DonationAmt_Per5080G_InCash(i)) & "</ITRForm:DonationAmtCash>"
                Else
                    writeXML "                   <ITRForm:DonationAmtCash>0</ITRForm:DonationAmtCash>"
                End If


                If DonationAmt_Per5080G(i) <> "" Then
                    writeXML "                   <ITRForm:DonationAmtOtherMode>" & UCase(DonationAmt_Per5080G(i)) & "</ITRForm:DonationAmtOtherMode>"
                Else
                    writeXML "                   <ITRForm:DonationAmtOtherMode>0</ITRForm:DonationAmtOtherMode>"
                End If


                If Trim(Range("Per5080G.TotalDonationAmt").Cells.item(i).value) <> "" Then
                writeXML "<ITRForm:DonationAmt>" & UCase(Range("Per5080G.TotalDonationAmt").Cells.item(i).value) & "</ITRForm:DonationAmt>"
                Else
                writeXML "<ITRForm:DonationAmt>0</ITRForm:DonationAmt>"
                End If



                If EligibleAmt_Per5080G(i) <> "" Then
                    writeXML "                   <ITRForm:EligibleDonationAmt>" & UCase(EligibleAmt_Per5080G(i)) & "</ITRForm:EligibleDonationAmt>"
                Else
                    writeXML "                   <ITRForm:EligibleDonationAmt>0</ITRForm:EligibleDonationAmt>"
                End If

                writeXML "               </ITRForm:DoneeWithPan>"
                UpdateProgressBar
            Next


               If Trim(Range("Per5080G.TotDon100Percent_InCash").value) <> "" Then
                writeXML "<ITRForm:TotDon50PercentApprReqdCash>" & UCase(Range("Per5080G.TotDon100Percent_InCash").value) & "</ITRForm:TotDon50PercentApprReqdCash>"
                Else
                writeXML "<ITRForm:TotDon50PercentApprReqdCash>0</ITRForm:TotDon50PercentApprReqdCash>"
                End If


            If TotDon50PercentApprReqd_Per5080G <> "" Then
                writeXML "               <ITRForm:TotDon50PercentApprReqdOtherMode>" & UCase(TotDon50PercentApprReqd_Per5080G) & "</ITRForm:TotDon50PercentApprReqdOtherMode>"
            Else
                writeXML "               <ITRForm:TotDon50PercentApprReqdOtherMode>0</ITRForm:TotDon50PercentApprReqdOtherMode>"
            End If


            If Trim(Range("Per5080G.TotDon100Percent_Total").value) <> "" Then
            writeXML "<ITRForm:TotDon50PercentApprReqd>" & UCase(Range("Per5080G.TotDon100Percent_Total").value) & "</ITRForm:TotDon50PercentApprReqd>"
            Else
            writeXML "<ITRForm:TotDon50PercentApprReqd>0</ITRForm:TotDon50PercentApprReqd>"
            End If

            If TotElig50PercentApprReqd_Per5080G <> "" Then
                writeXML "               <ITRForm:TotEligibleDon50PercentApprReqd>" & UCase(TotElig50PercentApprReqd_Per5080G) & "</ITRForm:TotEligibleDon50PercentApprReqd>"
            Else
                writeXML "               <ITRForm:TotEligibleDon50PercentApprReqd>0</ITRForm:TotEligibleDon50PercentApprReqd>"
            End If


            writeXML "          </ITRForm:Don50PercentApprReqd>"
        End If

        If Trim(Range("Per5080G.TotalDonationsUs80G_InCash").value) <> "" Then
            writeXML "<ITRForm:TotalDonationsUs80GCash>" & UCase(Range("Per5080G.TotalDonationsUs80G_InCash").value) & "</ITRForm:TotalDonationsUs80GCash>"
            Else
            writeXML "<ITRForm:TotalDonationsUs80GCash>0</ITRForm:TotalDonationsUs80GCash>"
            End If


        If TotalDonationsUs80G_Per5080G <> "" Then
            writeXML "           <ITRForm:TotalDonationsUs80GOtherMode>" & UCase(TotalDonationsUs80G_Per5080G) & "</ITRForm:TotalDonationsUs80GOtherMode>"
        Else
            writeXML "           <ITRForm:TotalDonationsUs80GOtherMode>0</ITRForm:TotalDonationsUs80GOtherMode>"
        End If

        If Trim(Range("Per5080G.TotalDonationsUs80G_Total").value) <> "" Then
        writeXML "<ITRForm:TotalDonationsUs80G>" & UCase(Range("Per5080G.TotalDonationsUs80G_Total").value) & "</ITRForm:TotalDonationsUs80G>"
        Else
        writeXML "<ITRForm:TotalDonationsUs80G>0</ITRForm:TotalDonationsUs80G>"
        End If


        If TotalEligibleDonationsUs80G_Per5080G <> "" Then
            writeXML "           <ITRForm:TotalEligibleDonationsUs80G>" & UCase(TotalEligibleDonationsUs80G_Per5080G) & "</ITRForm:TotalEligibleDonationsUs80G>"
        Else
            writeXML "           <ITRForm:TotalEligibleDonationsUs80G>0</ITRForm:TotalEligibleDonationsUs80G>"
        End If


        writeXML "      </ITRForm:Schedule80G>"
    'End If
End Function

Function Schedule80RA()
 Dim i As Long

subProcCaption = "Schedule 80RA"
noOfProcessSub = UBound(Name_of_Donee_80GGA)


 If Not IsEmpty(Name_of_Donee_80GGA) And UBound(Name_of_Donee_80GGA) > 0 Then
  writeXML "<ITRForm:Schedule80RA>"

  For i = 1 To UBound(Name_of_Donee_80GGA)

        writeXML "<ITRForm:DonationDtlsRsrchAssctn>"
        If Name_of_Donee_80GGA(i) <> "" Then

        writeXML "<ITRForm:NameOfDonee>" & Name_of_Donee_80GGA(i) & "</ITRForm:NameOfDonee>"
        Else
        writeXML "<ITRForm:NameOfDonee/>"
        End If

        writeXML "<ITRForm:AddressDetail>"
        If Address_80GGA(i) <> "" Then
        writeXML "<ITRForm:AddrDetail>" & UCase(Address_80GGA(i)) & "</ITRForm:AddrDetail>"

        End If

        If City_Town_District_80GGA(i) <> "" Then
        writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(City_Town_District_80GGA(i)) & "</ITRForm:CityOrTownOrDistrict>"

        End If

        If State_Code_80GGA(i) <> "" Then
        writeXML "<ITRForm:StateCode>" & Mid((State_Code_80GGA(i)), 1, 2) & "</ITRForm:StateCode>"

        End If

        If Pincode_80GGA(i) <> "" Then
        writeXML "<ITRForm:PinCode>" & UCase(Pincode_80GGA(i)) & "</ITRForm:PinCode>"

        End If

        writeXML "</ITRForm:AddressDetail>"

        If PAN_of_donee_80GGA(i) <> "" Then
        writeXML "<ITRForm:DoneePAN>" & UCase(PAN_of_donee_80GGA(i)) & "</ITRForm:DoneePAN>"

        End If




        If Donation_cash_80GGA(i) <> "" Then
        writeXML "<ITRForm:DonationAmtCash>" & UCase(Donation_cash_80GGA(i)) & "</ITRForm:DonationAmtCash>"
        Else
        writeXML "<ITRForm:DonationAmtCash>0</ITRForm:DonationAmtCash>"
        End If


        If Donation_other_80GGA(i) <> "" Then
        writeXML "<ITRForm:DonationAmtOtherMode>" & UCase(Donation_other_80GGA(i)) & "</ITRForm:DonationAmtOtherMode>"
        Else
        writeXML "<ITRForm:DonationAmtOtherMode>0</ITRForm:DonationAmtOtherMode>"
        End If


        If Trim(Range("Donation_total_80GGA").Cells.item(i).value) <> "" Then
        writeXML "<ITRForm:DonationAmt>" & UCase(Range("Donation_total_80GGA").Cells.item(i).value) & "</ITRForm:DonationAmt>"
        Else
        writeXML "<ITRForm:DonationAmt>0</ITRForm:DonationAmt>"
        End If


        If Trim(Range("Donation_Eligible_80GGA").Cells.item(i).value) <> "" Then
        writeXML "<ITRForm:EligibleDonationAmt>" & UCase(Range("Donation_Eligible_80GGA").Cells.item(i).value) & "</ITRForm:EligibleDonationAmt>"
        Else
        writeXML "<ITRForm:EligibleDonationAmt>0</ITRForm:EligibleDonationAmt>"
        End If

        writeXML "</ITRForm:DonationDtlsRsrchAssctn>"


        UpdateProgressBar

        Next

        If Trim(Range("Total_Donation_Cash_80GGA").value) > 0 Then
        writeXML "<ITRForm:TotalDonationAmtCash80RA>" & UCase(Range("Total_Donation_Cash_80GGA").value) & "</ITRForm:TotalDonationAmtCash80RA>"
        Else
        writeXML "<ITRForm:TotalDonationAmtCash80RA>0</ITRForm:TotalDonationAmtCash80RA>"
        End If

        If Trim(Range("Total_Donation_OtherMode_80GGA").value) > 0 Then
        writeXML "<ITRForm:TotalDonationAmtOtherMode80RA>" & UCase(Range("Total_Donation_OtherMode_80GGA").value) & "</ITRForm:TotalDonationAmtOtherMode80RA>"
        Else
        writeXML "<ITRForm:TotalDonationAmtOtherMode80RA>0</ITRForm:TotalDonationAmtOtherMode80RA>"
        End If

        If Trim(Range("Total_Donation_Total_80GGA").value) > 0 Then
        writeXML "<ITRForm:TotalDonationsUs80RA>" & UCase(Range("Total_Donation_Total_80GGA").value) & "</ITRForm:TotalDonationsUs80RA>"
        Else
        writeXML "<ITRForm:TotalDonationsUs80RA>0</ITRForm:TotalDonationsUs80RA>"
        End If


        If Trim(Range("Total_Donation_Eligible_80GGA").value) > 0 Then
        writeXML "<ITRForm:TotalEligibleDonationAmt80RA>" & UCase(Range("Total_Donation_Eligible_80GGA").value) & "</ITRForm:TotalEligibleDonationAmt80RA>"
        Else
        writeXML "<ITRForm:TotalEligibleDonationAmt80RA>0</ITRForm:TotalEligibleDonationAmt80RA>"
        End If
        writeXML "</ITRForm:Schedule80RA>"
End If



End Function



Function Schedule80_XML()

    Dim i As Variant

    If IA80_TotSchedule80_IA <> "" Then
    writeXML "<ITRForm:Schedule80_IA>"
    writeXML "<ITRForm:Sch80SectionCode>80-IA</ITRForm:Sch80SectionCode>"

        subProcCaption = "Schedule 80_I"
    writeXML "<ITRForm:DeductUs80_IA_4_ii>"
        writeXML "<ITRForm:Sch80LocOrDescCode>TEECOM</ITRForm:Sch80LocOrDescCode>"
        If UBound(IA80_DeductProfUs80_IA_4_I) > 0 Then
        noOfProcessSub = UBound(IA80_DeductProfUs80_IA_4_I)
            For i = 1 To UBound(IA80_DeductProfUs80_IA_4_I)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IA80_DeductProfUs80_IA_4_I(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
   writeXML "</ITRForm:DeductUs80_IA_4_ii>"


    writeXML "<ITRForm:DeductUs80_IA_4_iii>"
    writeXML "<ITRForm:Sch80LocOrDescCode>INDSTRL_PARK_SEZ</ITRForm:Sch80LocOrDescCode>"
    noOfProcessSub = UBound(IA80_DeductProfUs80_IA_4_Ii)
    If UBound(IA80_DeductProfUs80_IA_4_Ii) > 0 Then
        For i = 1 To UBound(IA80_DeductProfUs80_IA_4_Ii)
            writeXML "<ITRForm:Sch80DeductAmtDtls>"
            writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IA80_DeductProfUs80_IA_4_Ii(i)) & "</ITRForm:DeductAmountSec80>"
            writeXML "</ITRForm:Sch80DeductAmtDtls>"
            UpdateProgressBar
        Next
    Else
        writeXML "<ITRForm:Sch80DeductAmtDtls>"
        writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
        writeXML "</ITRForm:Sch80DeductAmtDtls>"
    End If
    writeXML "</ITRForm:DeductUs80_IA_4_iii>"

    writeXML "<ITRForm:DeductUs80_IA_4_iv>"
    writeXML "<ITRForm:Sch80LocOrDescCode>POWER</ITRForm:Sch80LocOrDescCode>"
    noOfProcessSub = UBound(IA80_DeductProfUs80_IA_4_iii)
    If UBound(IA80_DeductProfUs80_IA_4_iii) > 0 Then
        For i = 1 To UBound(IA80_DeductProfUs80_IA_4_iii)
            writeXML "<ITRForm:Sch80DeductAmtDtls>"
            writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IA80_DeductProfUs80_IA_4_iii(i)) & "</ITRForm:DeductAmountSec80>"
            writeXML "</ITRForm:Sch80DeductAmtDtls>"
            UpdateProgressBar
        Next
    Else
            writeXML "<ITRForm:Sch80DeductAmtDtls>"
            writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
            writeXML "</ITRForm:Sch80DeductAmtDtls>"
    End If
    writeXML "</ITRForm:DeductUs80_IA_4_iv>"

    writeXML "<ITRForm:DeductUs80_IA_4_v>"
    writeXML "<ITRForm:Sch80LocOrDescCode>REVIVAL_POWER_PLNT</ITRForm:Sch80LocOrDescCode>"
    noOfProcessSub = UBound(IA80_DeductProfUs80_IA_4_iv)
    If UBound(IA80_DeductProfUs80_IA_4_iv) > 0 Then
        For i = 1 To UBound(IA80_DeductProfUs80_IA_4_iv)
            writeXML "<ITRForm:Sch80DeductAmtDtls>"
            writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IA80_DeductProfUs80_IA_4_iv(i)) & "</ITRForm:DeductAmountSec80>"
            writeXML "</ITRForm:Sch80DeductAmtDtls>"
            UpdateProgressBar
        Next
    Else
          writeXML "<ITRForm:Sch80DeductAmtDtls>"
          writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
          writeXML "</ITRForm:Sch80DeductAmtDtls>"
    End If
    writeXML "</ITRForm:DeductUs80_IA_4_v>"

    writeXML "<ITRForm:TotSchedule80_IA>" & UVCase(IA80_TotSchedule80_IA) & "</ITRForm:TotSchedule80_IA>"
    writeXML "</ITRForm:Schedule80_IA>"
    End If

    If IB80_TotSchedule80_IB <> "" Then

        writeXML "<ITRForm:Schedule80_IB>"
        writeXML "<ITRForm:Sch80SectionCode>80-IB</ITRForm:Sch80SectionCode>"

        writeXML "<ITRForm:DeductJKLocUs80_IB_4_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_JK</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductJKLocUs80_IB_4)
        If UBound(IB80_DeductJKLocUs80_IB_4) > 0 Then
            For i = 1 To UBound(IB80_DeductJKLocUs80_IB_4)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductJKLocUs80_IB_4(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
            writeXML "<ITRForm:Sch80DeductAmtDtls>"
            writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
            writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductJKLocUs80_IB_4_Und>"

'        writeXML "<ITRForm:DeductBackStatesUs80_IB_4_Und>"
'        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_BCKWRD_ST</ITRForm:Sch80LocOrDescCode>"
'        If UBound(IB80_DeductBackStatesUs80_IB_4) > 0 Then
'        noOfProcessSub = UBound(IB80_DeductBackStatesUs80_IB_4)
'            For i = 1 To UBound(IB80_DeductBackStatesUs80_IB_4)
'                writeXML "<ITRForm:Sch80DeductAmtDtls>"
'                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductBackStatesUs80_IB_4(i)) & "</ITRForm:DeductAmountSec80>"
'                writeXML "</ITRForm:Sch80DeductAmtDtls>"
'                UpdateProgressBar
'            Next
'        Else
'            writeXML "<ITRForm:Sch80DeductAmtDtls>"
'            writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
'            writeXML "</ITRForm:Sch80DeductAmtDtls>"
'        End If
'        writeXML "</ITRForm:DeductBackStatesUs80_IB_4_Und>"

'        writeXML "<ITRForm:DeductBackDisttUs80_IB_5_Und>"
'        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_BCLWRD_DIST</ITRForm:Sch80LocOrDescCode>"
'        noOfProcessSub = UBound(IB80_DeductBackDisttUs80_IB_5)
'        If UBound(IB80_DeductBackDisttUs80_IB_5) > 0 Then
'            For i = 1 To UBound(IB80_DeductBackDisttUs80_IB_5)
'                writeXML "<ITRForm:Sch80DeductAmtDtls>"
'                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductBackDisttUs80_IB_5(i)) & "</ITRForm:DeductAmountSec80>"
'                writeXML "</ITRForm:Sch80DeductAmtDtls>"
'                UpdateProgressBar
'            Next
'        Else
'                writeXML "<ITRForm:Sch80DeductAmtDtls>"
'                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
'                writeXML "</ITRForm:Sch80DeductAmtDtls>"
'        End If
'        writeXML "</ITRForm:DeductBackDisttUs80_IB_5_Und>"

        writeXML "<ITRForm:DeductMultiplexUs80_IB_7A_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>MULTPLX_THEATRE</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductMultiplexUs80_IB_7A)
        If UBound(IB80_DeductMultiplexUs80_IB_7A) > 0 Then
            For i = 1 To UBound(IB80_DeductMultiplexUs80_IB_7A)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductMultiplexUs80_IB_7A(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductMultiplexUs80_IB_7A_Und>"

        writeXML "<ITRForm:DeductConvCentUs80_IB_7B_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>CONEVNTION_CENTRE</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductConvCentUs80_IB_7B)
        If UBound(IB80_DeductConvCentUs80_IB_7B) > 0 Then
            For i = 1 To UBound(IB80_DeductConvCentUs80_IB_7B)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductConvCentUs80_IB_7B(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductConvCentUs80_IB_7B_Und>"

        writeXML "<ITRForm:DeductMinOilUs80_IB_9_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>COMM_PROD</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductMinOilUs80_IB_9)
        If UBound(IB80_DeductMinOilUs80_IB_9) > 0 Then
            For i = 1 To UBound(IB80_DeductMinOilUs80_IB_9)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductMinOilUs80_IB_9(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductMinOilUs80_IB_9_Und>"

        writeXML "<ITRForm:DeductHousUs80_IB_10_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>HOUSING_PROJECT</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductHousUs80_IB_10)
        If UBound(IB80_DeductHousUs80_IB_10) > 0 Then
            For i = 1 To UBound(IB80_DeductHousUs80_IB_10)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductHousUs80_IB_10(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductHousUs80_IB_10_Und>"

'        writeXML "<ITRForm:DeductColdChainUs80_IB_11_Und>"
'        writeXML "<ITRForm:Sch80LocOrDescCode>CHAIN_FACILITY</ITRForm:Sch80LocOrDescCode>"
'        noOfProcessSub = UBound(IB80_DeductColdChainUs80_IB_11)
'        If UBound(IB80_DeductColdChainUs80_IB_11) > 0 Then
'            For i = 1 To UBound(IB80_DeductColdChainUs80_IB_11)
'                writeXML "<ITRForm:Sch80DeductAmtDtls>"
'                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductColdChainUs80_IB_11(i)) & "</ITRForm:DeductAmountSec80>"
'                writeXML "</ITRForm:Sch80DeductAmtDtls>"
'                UpdateProgressBar
'            Next
'        Else
'                writeXML "<ITRForm:Sch80DeductAmtDtls>"
'                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
'                writeXML "</ITRForm:Sch80DeductAmtDtls>"
'        End If
'        writeXML "</ITRForm:DeductColdChainUs80_IB_11_Und>"

        writeXML "<ITRForm:DeductFruitVegUs80_IB_11A_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>FRIUTS_VEGTBLE</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductFruitVegUs80_IB_11A)
        If UBound(IB80_DeductFruitVegUs80_IB_11A) > 0 Then
            For i = 1 To UBound(IB80_DeductFruitVegUs80_IB_11A)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductFruitVegUs80_IB_11A(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductFruitVegUs80_IB_11A_Und>"

        writeXML "<ITRForm:DeductFoodGrainUs80_IB_11A_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>STOR_TRANS</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductFoodGrainUs80_IB_11A)
        If UBound(IB80_DeductFoodGrainUs80_IB_11A) > 0 Then
            For i = 1 To UBound(IB80_DeductFoodGrainUs80_IB_11A)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductFoodGrainUs80_IB_11A(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductFoodGrainUs80_IB_11A_Und>"


        writeXML "<ITRForm:DeductRurHospUs80_IB_11B_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>RURAL_HOSPTL</ITRForm:Sch80LocOrDescCode>"
        noOfProcessSub = UBound(IB80_DeductRurHospUs80_IB_11B)
        If UBound(IB80_DeductRurHospUs80_IB_11B) > 0 Then
            For i = 1 To UBound(IB80_DeductRurHospUs80_IB_11B)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductRurHospUs80_IB_11B(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
                UpdateProgressBar
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductRurHospUs80_IB_11B_Und>"

        writeXML "<ITRForm:DeductHospAnyAreaUs80IB_11C_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>MAINTAIN_HOSP</ITRForm:Sch80LocOrDescCode>"
        If UBound(IB80_DeductHospAnyAreaUs80IB_11C) > 0 Then
            For i = 1 To UBound(IB80_DeductHospAnyAreaUs80IB_11C)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IB80_DeductHospAnyAreaUs80IB_11C(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductHospAnyAreaUs80IB_11C_Und>"

        writeXML "<ITRForm:TotSchedule80_IB>" & UVCase(IB80_TotSchedule80_IB) & "</ITRForm:TotSchedule80_IB>"
        writeXML "</ITRForm:Schedule80_IB>"

    End If

    If IC80_TotSchedule80_IC <> "" Then
        writeXML "<ITRForm:Schedule80_IC>"
        writeXML "<ITRForm:Sch80SectionCode>80-IC_IE</ITRForm:Sch80SectionCode>"

        writeXML "<ITRForm:DeductInSikkim_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSTRL_SIKKIM</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_DeductInSikkim) > 0 Then
            For i = 1 To UBound(IC80_DeductInSikkim)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_DeductInSikkim(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductInSikkim_Und>"

        writeXML "<ITRForm:DeductInHimachalP_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_HP</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_DeductInHimachalP) > 0 Then
            For i = 1 To UBound(IC80_DeductInHimachalP)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_DeductInHimachalP(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductInHimachalP_Und>"

        writeXML "<ITRForm:DeductInUttaranchal_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_UTTARANCHAL</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_DeductInUttaranchal) > 0 Then
            For i = 1 To UBound(IC80_DeductInUttaranchal)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_DeductInUttaranchal(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:DeductInUttaranchal_Und>"

        writeXML "<ITRForm:DeductInNorthEast>"
        writeXML "<ITRForm:Assam_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_ASSAM</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_Assam) > 0 Then
            For i = 1 To UBound(IC80_Assam)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_Assam(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:Assam_Und>"

        writeXML "<ITRForm:ArunachalPradesh_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_ARUNPRADESH</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_ArunachalPradesh) > 0 Then
            For i = 1 To UBound(IC80_ArunachalPradesh)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_ArunachalPradesh(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:ArunachalPradesh_Und>"

        writeXML "<ITRForm:Manipur_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_MANIPUR</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_Manipur) > 0 Then
            For i = 1 To UBound(IC80_Manipur)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_Manipur(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
            writeXML "<ITRForm:Sch80DeductAmtDtls>"
            writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
            writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:Manipur_Und>"

        writeXML "<ITRForm:Mizoram_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_MIZORAM</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_Mizoram) > 0 Then
            For i = 1 To UBound(IC80_Mizoram)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_Mizoram(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:Mizoram_Und>"

        writeXML "<ITRForm:Meghalaya_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_MEGHALAYA</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_Meghalaya) > 0 Then
            For i = 1 To UBound(IC80_Meghalaya)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_Meghalaya(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:Meghalaya_Und>"

        writeXML "<ITRForm:Nagaland_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_NAGALND</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_Nagaland) > 0 Then
            For i = 1 To UBound(IC80_Nagaland)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_Nagaland(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:Nagaland_Und>"

        writeXML "<ITRForm:Tripura_Und>"
        writeXML "<ITRForm:Sch80LocOrDescCode>INDSRTL_TRIPURA</ITRForm:Sch80LocOrDescCode>"
        If UBound(IC80_Tripura) > 0 Then
            For i = 1 To UBound(IC80_Tripura)
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>" & UVCase(IC80_Tripura(i)) & "</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
            Next
        Else
                writeXML "<ITRForm:Sch80DeductAmtDtls>"
                writeXML "<ITRForm:DeductAmountSec80>0</ITRForm:DeductAmountSec80>"
                writeXML "</ITRForm:Sch80DeductAmtDtls>"
        End If
        writeXML "</ITRForm:Tripura_Und>"

        writeXML "<ITRForm:TotDeductInNorthEast>" & UVCase(IC80_TotDeductInNorthEast) & "</ITRForm:TotDeductInNorthEast>"
        writeXML "</ITRForm:DeductInNorthEast>"

        writeXML "<ITRForm:TotSchedule80_IC>" & UVCase(IC80_TotSchedule80_IC) & "</ITRForm:TotSchedule80_IC>"
        writeXML "</ITRForm:Schedule80_IC>"
    End If
End Function

Function VIAXML()

writeXML "<ITRForm:ScheduleVIA>"



    writeXML "<ITRForm:UsrDeductUndChapVIA>"
      writeXML "<ITRForm:Section80C>" & scviaSection80C & "</ITRForm:Section80C>"
      writeXML "<ITRForm:Section80CCC>" & scviaSection80CCC & "</ITRForm:Section80CCC>"
      writeXML "<ITRForm:Section80CCDEmployeeOrSE>" & scviaSection80CCD_SE & "</ITRForm:Section80CCDEmployeeOrSE>"
      writeXML "<ITRForm:Section80CCD1B>" & scviaSection80CCD1B_SE & "</ITRForm:Section80CCD1B>"
      writeXML "<ITRForm:Section80CCDEmployer>" & scviaSection80CCD & "</ITRForm:Section80CCDEmployer>"
'      writeXML "<ITRForm:Section80CCG>" & scviaSection80CCG & "</ITRForm:Section80CCG>"
'      writeXML "<ITRForm:Section80DHealthInsPremium>"
'
'If SELECT80D <> "" Then
'    writeXML "<ITRForm:HealthInsurancePremium>" & SELECT80D & "</ITRForm:HealthInsurancePremium>"
'End If
'
'If scviaSection80D <> "" Then
'writeXML "<ITRForm:Sec80DHealthInsurancePremiumUsr>" & scviaSection80D & "</ITRForm:Sec80DHealthInsurancePremiumUsr>"
'Else
'writeXML "<ITRForm:Sec80DHealthInsurancePremiumUsr>" & 0 & "</ITRForm:Sec80DHealthInsurancePremiumUsr>"
'End If
'
'
' UpdateProgressBar
'
'
'If SELECT80DB <> "" Then
'    writeXML "<ITRForm:MedicalExpenditure>" & SELECT80DB & "</ITRForm:MedicalExpenditure>"
'End If

'If scviaSection80DB <> "" Then
'writeXML "<ITRForm:Sec80DMedicalExpenditureUsr>" & scviaSection80DB & "</ITRForm:Sec80DMedicalExpenditureUsr>"
'Else
'writeXML "<ITRForm:Sec80DMedicalExpenditureUsr>" & 0 & "</ITRForm:Sec80DMedicalExpenditureUsr>"
'End If
'
'If SELECT80DC <> "" Then
'    writeXML "<ITRForm:PreventiveHealthCheckUp>" & SELECT80DC & "</ITRForm:PreventiveHealthCheckUp>"
'End If
'
'If scviaSection80DC <> "" Then
'writeXML "<ITRForm:Sec80DPreventiveHealthCheckUpUsr>" & scviaSection80DC & "</ITRForm:Sec80DPreventiveHealthCheckUpUsr>"
'Else
'writeXML "<ITRForm:Sec80DPreventiveHealthCheckUpUsr>" & 0 & "</ITRForm:Sec80DPreventiveHealthCheckUpUsr>"
'End If
'
'
'writeXML "</ITRForm:Section80DHealthInsPremium>"
 UpdateProgressBar
 
        If Sheet19.Range("scvia.Section80DValue").value <> "" Then
        writeXML "<ITRForm:Section80D>" & Sheet19.Range("scvia.Section80DValue").value & "</ITRForm:Section80D>"
      End If

        If SELECT80DD <> "" Then
        writeXML "<ITRForm:Section80DDUsrType>" & SELECT80DD & "</ITRForm:Section80DDUsrType>"
      End If

      writeXML "<ITRForm:Section80DD>" & scviaSection80DD & "</ITRForm:Section80DD>"

      If SELECT80DDB <> "" Then
        writeXML "<ITRForm:Section80DDBUsrType>" & SELECT80DDB & "</ITRForm:Section80DDBUsrType>"
      End If

      writeXML "<ITRForm:Section80DDB>" & scviaSection80DDB & "</ITRForm:Section80DDB>"


      writeXML "<ITRForm:Section80E>" & scviaSection80E & "</ITRForm:Section80E>"
      writeXML "<ITRForm:Section80EE>" & scviaSection80EE & "</ITRForm:Section80EE>"
      writeXML "<ITRForm:Section80EEA>" & Sheet19.Range("scvia.Section80EEA").value & "</ITRForm:Section80EEA>"
      writeXML "<ITRForm:Section80EEB>" & Sheet19.Range("scvia.Section80EEB").value & "</ITRForm:Section80EEB>"
      writeXML "<ITRForm:Section80G>" & scviaSection80G & "</ITRForm:Section80G>"
      writeXML "<ITRForm:Section80GG>" & scviaSection80GG & "</ITRForm:Section80GG>"
      writeXML "<ITRForm:Section80GGC>" & scviaSection80GGC & "</ITRForm:Section80GGC>"
      writeXML "<ITRForm:TotPartBchapterVIA>" & scviaTotPartBchapterVIA & "</ITRForm:TotPartBchapterVIA>"
      writeXML "<ITRForm:Section80IA>" & scviaSection80IA & "</ITRForm:Section80IA>"
      writeXML "<ITRForm:Section80IAB>" & scviaSection80IAB & "</ITRForm:Section80IAB>"
      writeXML "<ITRForm:Section80IB>" & scviaSection80IB & "</ITRForm:Section80IB>"
      writeXML "<ITRForm:Section80IBA>" & scviaSection80IBA & "</ITRForm:Section80IBA>"
      writeXML "<ITRForm:Section80IC>" & scviaSection80IC & "</ITRForm:Section80IC>"
      'writeXML "<ITRForm:Section80ID>" & scviaSection80ID & "</ITRForm:Section80ID>"
      writeXML "<ITRForm:Section80JJA>" & scviaSection80JJA & "</ITRForm:Section80JJA>"
      writeXML "<ITRForm:Section80JJAA>" & scviaSection80JJAA & "</ITRForm:Section80JJAA>"
       writeXML "<ITRForm:Section80QQB>" & scviaSection80QQB & "</ITRForm:Section80QQB>"
       writeXML "<ITRForm:Section80RRB>" & scviaSection80RRB & "</ITRForm:Section80RRB>"
       writeXML "<ITRForm:TotPartCchapterVIA>" & scviaTotPartCchapterVIA & "</ITRForm:TotPartCchapterVIA>"
      writeXML "<ITRForm:Section80TTA>" & scviaSection80TTA & "</ITRForm:Section80TTA>"
      writeXML "<ITRForm:Section80TTB>" & scviaSection80TTB & "</ITRForm:Section80TTB>"
      If SELECT80U <> "" Then
        writeXML "<ITRForm:Section80UUsrType>" & SELECT80U & "</ITRForm:Section80UUsrType>"
      End If


      writeXML "<ITRForm:Section80U>" & scviaSection80U & "</ITRForm:Section80U>"
      writeXML "<ITRForm:TotPartCAandDchapterVIA>" & scviaTotPartCAandDchapterVIA & "</ITRForm:TotPartCAandDchapterVIA>"
      writeXML "<ITRForm:TotalChapVIADeductions>" & scviaUsrTotalChapVIADeductions & "</ITRForm:TotalChapVIADeductions>"
    writeXML "</ITRForm:UsrDeductUndChapVIA>"

    writeXML "<ITRForm:DeductUndChapVIA>"
    writeXML "<ITRForm:Section80C>" & scviaSection80C_Calc & "</ITRForm:Section80C>"
      writeXML "<ITRForm:Section80CCC>" & scviaSection80CCC_Calc & "</ITRForm:Section80CCC>"
      writeXML "<ITRForm:Section80CCDEmployeeOrSE>" & scviaSection80CCD_SE_Calc & "</ITRForm:Section80CCDEmployeeOrSE>"
      writeXML "<ITRForm:Section80CCD1B>" & scviaSection80CCD1B_SE_Calc & "</ITRForm:Section80CCD1B>"
      writeXML "<ITRForm:Section80CCDEmployer>" & scviaSection80CCD_Calc & "</ITRForm:Section80CCDEmployer>"
      'writeXML "<ITRForm:Section80CCG>" & scviaSection80CCG_Calc & "</ITRForm:Section80CCG>"
      writeXML "<ITRForm:Section80D>" & Sheet19.Range("scvia.Section80DValue_Calc").value & "</ITRForm:Section80D>"
      writeXML "<ITRForm:Section80DD>" & scviaSection80DD_Calc & "</ITRForm:Section80DD>"
      writeXML "<ITRForm:Section80DDB>" & scviaSection80DDB_Calc & "</ITRForm:Section80DDB>"
      writeXML "<ITRForm:Section80E>" & scviaSection80E_Calc & "</ITRForm:Section80E>"
      writeXML "<ITRForm:Section80EE>" & scviaSection80EE_Calc & "</ITRForm:Section80EE>"
      writeXML "<ITRForm:Section80EEA>" & Sheet19.Range("scvia.Section80EEA_Calc").value & "</ITRForm:Section80EEA>"
      writeXML "<ITRForm:Section80EEB>" & Sheet19.Range("scvia.Section80EEB_Calc").value & "</ITRForm:Section80EEB>"
      writeXML "<ITRForm:Section80G>" & scviaSection80G_Calc & "</ITRForm:Section80G>"
      writeXML "<ITRForm:Section80GG>" & scviaSection80GG_Calc & "</ITRForm:Section80GG>"
      writeXML "<ITRForm:Section80GGC>" & scviaSection80GGC_Calc & "</ITRForm:Section80GGC>"
      writeXML "<ITRForm:TotPartBchapterVIA>" & scviaTotPartBchapterVIA_Calc & "</ITRForm:TotPartBchapterVIA>"
      writeXML "<ITRForm:Section80IA>" & scviaSection80IA_Calc & "</ITRForm:Section80IA>"
      writeXML "<ITRForm:Section80IAB>" & scviaSection80IAB_Calc & "</ITRForm:Section80IAB>"
      writeXML "<ITRForm:Section80IB>" & scviaSection80IB_Calc & "</ITRForm:Section80IB>"
      writeXML "<ITRForm:Section80IBA>" & scviaSection80IBA_Calc & "</ITRForm:Section80IBA>"
      writeXML "<ITRForm:Section80IC>" & scviaSection80IC_Calc & "</ITRForm:Section80IC>"
      'writeXML "<ITRForm:Section80ID>" & scviaSection80ID_Calc & "</ITRForm:Section80ID>"
      writeXML "<ITRForm:Section80JJA>" & scviaSection80JJA_Calc & "</ITRForm:Section80JJA>"
      writeXML "<ITRForm:Section80JJAA>" & scviaSection80JJAA_Calc & "</ITRForm:Section80JJAA>"
       writeXML "<ITRForm:Section80QQB>" & scviaSection80QQB_Calc & "</ITRForm:Section80QQB>"
       writeXML "<ITRForm:Section80RRB>" & scviaSection80RRB_Calc & "</ITRForm:Section80RRB>"
       writeXML "<ITRForm:TotPartCchapterVIA>" & scviaTotPartCchapterVIA_Calc & "</ITRForm:TotPartCchapterVIA>"
      writeXML "<ITRForm:Section80TTA>" & scviaSection80TTA_Calc & "</ITRForm:Section80TTA>"
      writeXML "<ITRForm:Section80TTB>" & scviaSection80TTB_Calc & "</ITRForm:Section80TTB>"
      writeXML "<ITRForm:Section80U>" & scviaSection80U_Calc & "</ITRForm:Section80U>"
      writeXML "<ITRForm:TotPartCAandDchapterVIA>" & scviaTotPartCAandDchapterVIA_Calc & "</ITRForm:TotPartCAandDchapterVIA>"
      writeXML "<ITRForm:TotalChapVIADeductions>" & scviaTotalChapVIADeductions_Calc & "</ITRForm:TotalChapVIADeductions>"
    writeXML "</ITRForm:DeductUndChapVIA>"
  writeXML "</ITRForm:ScheduleVIA>"
End Function

Function ScheduleAMTXML()


    'If Not IsEmpty(TotalIncItem11_AMT) And TotalIncItem11_AMT > 0 Then
        writeXML "  <ITRForm:ScheduleAMT>"

        If TotalIncItem11_AMT <> "" Then
            writeXML "        <ITRForm:TotalIncItem11>" & TotalIncItem11_AMT & "</ITRForm:TotalIncItem11>"
        Else
            writeXML "        <ITRForm:TotalIncItem11>0</ITRForm:TotalIncItem11>"
        End If

        writeXML "     <ITRForm:AdjustmentSec115JC>"

        If DeductClaimSec6A_AMT <> "" Then
            writeXML "        <ITRForm:DeductClaimSec6A>" & DeductClaimSec6A_AMT & "</ITRForm:DeductClaimSec6A>"
        Else
            writeXML "        <ITRForm:DeductClaimSec6A>0</ITRForm:DeductClaimSec6A>"
        End If

        If DeductClaimSec10AA_AMT <> "" Then
            writeXML "        <ITRForm:DeductClaimSec10AA>" & DeductClaimSec10AA_AMT & "</ITRForm:DeductClaimSec10AA>"
        Else
            writeXML "        <ITRForm:DeductClaimSec10AA>0</ITRForm:DeductClaimSec10AA>"
        End If

        If DeductClaimSec35AD_AMT <> "" Then
            writeXML "        <ITRForm:DeductClaimSec35AD>" & DeductClaimSec35AD_AMT & "</ITRForm:DeductClaimSec35AD>"
        Else
            writeXML "        <ITRForm:DeductClaimSec35AD>0</ITRForm:DeductClaimSec35AD>"
        End If

        If Total_AMT <> "" Then
            writeXML "        <ITRForm:Total>" & Total_AMT & "</ITRForm:Total>"
        Else
            writeXML "        <ITRForm:Total>0</ITRForm:Total>"
        End If

        writeXML "     </ITRForm:AdjustmentSec115JC>"

        If AdjustedUnderSec115JC_AMT <> "" Then
            writeXML "<ITRForm:AdjustedUnderSec115JC>" & AdjustedUnderSec115JC_AMT & "</ITRForm:AdjustedUnderSec115JC>"
        Else
            writeXML "<ITRForm:AdjustedUnderSec115JC>0</ITRForm:AdjustedUnderSec115JC>"
        End If

        If TaxPayableUnderSec115JC_AMT <> "" Then
            writeXML "<ITRForm:TaxPayableUnderSec115JC>" & TaxPayableUnderSec115JC_AMT & "</ITRForm:TaxPayableUnderSec115JC>"
        Else
            writeXML "<ITRForm:TaxPayableUnderSec115JC>0</ITRForm:TaxPayableUnderSec115JC>"
        End If

        writeXML "</ITRForm:ScheduleAMT>"
    'End If



End Function

Function ScheduleSPIXML()
Dim i As Variant
If UBound(personName_SPI) > 0 Then
writeXML "<ITRForm:ScheduleSPI>"

For i = 1 To UBound(personName_SPI)
writeXML "  <ITRForm:SpecifiedPerson>"

writeXML "      <ITRForm:SpecifiedPersonName>" & personName_SPI(i) & "</ITRForm:SpecifiedPersonName>"

If Pan_SPI(i) <> "" Then
writeXML "      <ITRForm:PANofSpecPerson>" & Pan_SPI(i) & "</ITRForm:PANofSpecPerson>"
End If

If Aadhaar_SPI(i) <> "" Then
writeXML "      <ITRForm:AaadhaarOfSpecPerson>" & Aadhaar_SPI(i) & "</ITRForm:AaadhaarOfSpecPerson>"
End If

writeXML "      <ITRForm:ReltnShip>" & relation_SPI(i) & "</ITRForm:ReltnShip>"

writeXML "      <ITRForm:AmtIncluded>" & amt_SPI(i) & "</ITRForm:AmtIncluded>"

                If IncmNature_SPI(i) = "Business/Profession" Then
                   IncmNature_SPI(i) = "BP"
                ElseIf IncmNature_SPI(i) = "Other sources" Then
                   IncmNature_SPI(i) = "OS"
                ElseIf IncmNature_SPI(i) = "Exempt Income" Then
                   IncmNature_SPI(i) = "EI"
                ElseIf IncmNature_SPI(i) = "House Property" Then
                   IncmNature_SPI(i) = "HP"
                ElseIf IncmNature_SPI(i) = "Capital Gains" Then
                   IncmNature_SPI(i) = "CG"
                ElseIf IncmNature_SPI(i) = "Salary" Then
                   IncmNature_SPI(i) = "SA"
                End If


writeXML "      <ITRForm:HeadIncIncluded>" & IncmNature_SPI(i) & "</ITRForm:HeadIncIncluded>"



writeXML "  </ITRForm:SpecifiedPerson>"

Next
writeXML "</ITRForm:ScheduleSPI>"
End If
End Function
Function ScheduleSIXML()
    Dim i As Long
    If Not IsEmpty(SecCode_SI1) And UBound(SecCode_SI1) > 0 Then
        writeXML "      <ITRForm:ScheduleSI>"
        For i = 1 To UBound(SecCode_SI1)
            If Trim(SecCode_SI1(i)) <> "PTI_5BB" And Trim(SecCode_SI1(i)) <> "PTI_5BBE" And Trim(SecCode_SI1(i)) <> "xx" And Trim(SecCode_SI1(i)) <> "FA" And Trim(SecCode_SI1(i)) <> "5BBF_XXXX" And Trim(SecCode_SI1(i)) <> "5BBG_XXXX" Then
                writeXML "           <ITRForm:SplCodeRateTax>"

                If Trim(SecCode_SI1(i)) <> "" Then
                    writeXML "               <ITRForm:SecCode>" & Trim(SecCode_SI1(i)) & "</ITRForm:SecCode>"
                Else
                    writeXML "               <ITRForm:SecCode/>"
                End If

                If SplRatePercent_SI1(i) <> "" Then
                    writeXML "               <ITRForm:SplRatePercent>" & UCase(SplRatePercent_SI1(i)) & "</ITRForm:SplRatePercent>"
                Else
                    writeXML "               <ITRForm:SplRatePercent>" & UCase(DefaultSplRatePercent_SI) & "</ITRForm:SplRatePercent>"
                End If

                If SplRateInc_SI1(i) <> "" Then
                    writeXML "               <ITRForm:SplRateInc>" & UCase(SplRateInc_SI1(i)) & "</ITRForm:SplRateInc>"
                Else
                    writeXML "               <ITRForm:SplRateInc>" & UCase(DefaultSplRateInc_SI) & "</ITRForm:SplRateInc>"
                End If

                If SplRateIncTax_SI1(i) <> "" Then
                    writeXML "               <ITRForm:SplRateIncTax>" & UCase(SplRateIncTax_SI1(i)) & "</ITRForm:SplRateIncTax>"
                Else
                    writeXML "               <ITRForm:SplRateIncTax>" & UCase(DefaultSplRateIncTax_SI) & "</ITRForm:SplRateIncTax>"
                End If

                writeXML "           </ITRForm:SplCodeRateTax>"
            End If
        Next
        
        
        If Sheet21.Range("SI.TotSplRateIncome").value <> 0 Then
            writeXML "           <ITRForm:TotSplRateInc>" & Sheet21.Range("SI.TotSplRateIncome").value & "</ITRForm:TotSplRateInc>"
        Else
            writeXML "           <ITRForm:TotSplRateInc>" & 0 & "</ITRForm:TotSplRateInc>"
        End If
        
        If TotSplRateIncTax_SI <> "" Then
            writeXML "           <ITRForm:TotSplRateIncTax>" & UCase(TotSplRateIncTax_SI) & "</ITRForm:TotSplRateIncTax>"
        Else
            writeXML "           <ITRForm:TotSplRateIncTax>" & UCase(DefaultTotSplRateIncTax_SI) & "</ITRForm:TotSplRateIncTax>"
        End If
        writeXML "      </ITRForm:ScheduleSI>"
    End If
End Function

Function ScheduleIFXML()
    writeXML "  <ITRForm:ScheduleIF>"
        Dim i As Variant

        For i = 1 To UBound(firmName_IF)

        writeXML "      <ITRForm:PartnerFirmDetails>"

        If firmName_IF(i) <> "" Then
        writeXML "              <ITRForm:FirmName>" & firmName_IF(i) & "</ITRForm:FirmName>"
        End If

        If firmPan_IF(i) <> "" Then
        writeXML "              <ITRForm:FirmPAN>" & firmPan_IF(i) & "</ITRForm:FirmPAN>"
        End If

        writeXML "              <ITRForm:IsLiableToAudit>" & Mid(firmLiability_IF(i), 1, 1) & "</ITRForm:IsLiableToAudit>"
        writeXML "              <ITRForm:Sec92EFirmFlag>" & Mid(firm92E_IF(i), 1, 1) & "</ITRForm:Sec92EFirmFlag>"

        If firmProfitShare_IF(i) <> "" Then
        writeXML "              <ITRForm:ProfitSharePercent>" & firmProfitShare_IF(i) & "</ITRForm:ProfitSharePercent>"
        Else
        writeXML "              <ITRForm:ProfitSharePercent>0</ITRForm:ProfitSharePercent>"
        End If

        If firmProfitAmt_IF(i) <> "" Then
        writeXML "              <ITRForm:ProfitShareAmt>" & firmProfitAmt_IF(i) & "</ITRForm:ProfitShareAmt>"
        Else
        writeXML "              <ITRForm:ProfitShareAmt>0</ITRForm:ProfitShareAmt>"
        End If

        If firmCapital_IF(i) <> "" Then
        writeXML "              <ITRForm:FirmCapBalOn31Mar>" & firmCapital_IF(i) & "</ITRForm:FirmCapBalOn31Mar>"
        Else
        writeXML "              <ITRForm:FirmCapBalOn31Mar>0</ITRForm:FirmCapBalOn31Mar>"
        End If

        writeXML "      </ITRForm:PartnerFirmDetails>"
        Next

        If Sheet21.Range("IF.TotalProfitShareAmt").value <> "" Then
        writeXML "       <ITRForm:TotalProfitShareAmt>" & Sheet21.Range("IF.TotalProfitShareAmt").value & "</ITRForm:TotalProfitShareAmt>"
        Else
        writeXML "       <ITRForm:TotalProfitShareAmt>0</ITRForm:TotalProfitShareAmt>"
        End If

        If Sheet21.Range("IF.TotalFirmCapBalOn31Mar").value <> "" Then
        writeXML "       <ITRForm:TotalFirmCapBalOn31Mar>" & Sheet21.Range("IF.TotalFirmCapBalOn31Mar").value & "</ITRForm:TotalFirmCapBalOn31Mar>"
        Else
        writeXML "       <ITRForm:TotalFirmCapBalOn31Mar>0</ITRForm:TotalFirmCapBalOn31Mar>"
        End If

        writeXML "  </ITRForm:ScheduleIF>"

End Function

Function ScheduleEIXML()
writeXML "<ITRForm:ScheduleEI>"

Dim i As Variant

If sceiInterestInc <> "" Then
writeXML "      <ITRForm:InterestInc>" & sceiInterestInc & "</ITRForm:InterestInc>"
Else
writeXML "      <ITRForm:InterestInc>0</ITRForm:InterestInc>"
End If

If sceiDividendInc <> "" Then
writeXML "      <ITRForm:DividendInc>" & sceiDividendInc & "</ITRForm:DividendInc>"
Else
writeXML "      <ITRForm:DividendInc>0</ITRForm:DividendInc>"
End If

'If sceiLTCGWhereSTTPaid <> "" Then
'writeXML "      <ITRForm:LTCGWhereSTTPaid>" & sceiLTCGWhereSTTPaid & "</ITRForm:LTCGWhereSTTPaid>"
'Else
'writeXML "      <ITRForm:LTCGWhereSTTPaid>0</ITRForm:LTCGWhereSTTPaid>"
'End If

If sceiNetAgriIncOrOthrIncRule7 <> "" Then
writeXML "<ITRForm:GrossAgriRecpt>" & sceiNetAgriIncOrOthrIncRule7 & "</ITRForm:GrossAgriRecpt>"
Else
writeXML "<ITRForm:GrossAgriRecpt>0</ITRForm:GrossAgriRecpt>"
End If

If sceiExpenditureOnAgriculture <> "" Then
writeXML "<ITRForm:ExpIncAgri>" & sceiExpenditureOnAgriculture & "</ITRForm:ExpIncAgri>"
Else
writeXML "<ITRForm:ExpIncAgri>0</ITRForm:ExpIncAgri>"
End If

If sceiUnabsorbedAgriculturalloss <> "" Then
writeXML "<ITRForm:UnabAgriLossPrev8>" & sceiUnabsorbedAgriculturalloss & "</ITRForm:UnabAgriLossPrev8>"
Else
writeXML "<ITRForm:UnabAgriLossPrev8>0</ITRForm:UnabAgriLossPrev8>"
End If

If sceiAgriInc <> "" Then
writeXML "<ITRForm:AgriIncRule7and8>" & sceiAgriInc & "</ITRForm:AgriIncRule7and8>"
Else
writeXML "<ITRForm:AgriIncRule7and8>0</ITRForm:AgriIncRule7and8>"
End If

If sceiNetAgriculturalIncome <> "" Then
writeXML "      <ITRForm:NetAgriIncOrOthrIncRule7>" & sceiNetAgriculturalIncome & "</ITRForm:NetAgriIncOrOthrIncRule7>"
Else
writeXML "      <ITRForm:NetAgriIncOrOthrIncRule7>0</ITRForm:NetAgriIncOrOthrIncRule7>"
End If


If UBound(NameOfDistrict_EI) > 0 Then
writeXML "  <ITRForm:ExcNetAgriInc>"


For i = 1 To UBound(NameOfDistrict_EI)

writeXML "  <ITRForm:ExcNetAgriIncDtls>"

If NameOfDistrict_EI(i) <> "" Then
writeXML "      <ITRForm:NameOfDistrict>" & NameOfDistrict_EI(i) & "</ITRForm:NameOfDistrict>"
Else
writeXML "      <ITRForm:NameOfDistrict>0</ITRForm:NameOfDistrict>"
End If

If Pincode_EI(i) <> "" Then
writeXML "      <ITRForm:PinCode>" & Pincode_EI(i) & "</ITRForm:PinCode>"
Else
writeXML "      <ITRForm:PinCode>0</ITRForm:PinCode>"
End If

If AgriculturalLand_In_Acre_EI(i) <> "" Then
writeXML "      <ITRForm:MeasurementOfLand>" & Round(AgriculturalLand_In_Acre_EI(i), 2) & "</ITRForm:MeasurementOfLand>"
Else
writeXML "      <ITRForm:MeasurementOfLand>0</ITRForm:MeasurementOfLand>"
End If

If AgriculturalLand_Owned_On_Held_onlease_EI(i) <> "" Then

If (AgriculturalLand_Owned_On_Held_onlease_EI(i)) = "Owned" Then
    AgriculturalLand_Owned_On_Held_onlease_EI(i) = "O"
 ElseIf (AgriculturalLand_Owned_On_Held_onlease_EI(i)) = "Held on Lease" Then
    AgriculturalLand_Owned_On_Held_onlease_EI(i) = "H"

 End If

writeXML "<ITRForm:AgriLandOwnedFlag>" & AgriculturalLand_Owned_On_Held_onlease_EI(i) & "</ITRForm:AgriLandOwnedFlag>"

End If


If AgriculturalLand_irrigated_or_rainfed_EI(i) <> "" Then

 If (AgriculturalLand_irrigated_or_rainfed_EI(i)) = "Irrigated " Then
    AgriculturalLand_irrigated_or_rainfed_EI(i) = "IRG"
 ElseIf (AgriculturalLand_irrigated_or_rainfed_EI(i)) = "Rain-fed" Then
    AgriculturalLand_irrigated_or_rainfed_EI(i) = "RF"

End If

writeXML "<ITRForm:AgriLandIrrigatedFlag>" & AgriculturalLand_irrigated_or_rainfed_EI(i) & "</ITRForm:AgriLandIrrigatedFlag>"

End If


writeXML "</ITRForm:ExcNetAgriIncDtls>"

Next

writeXML "</ITRForm:ExcNetAgriInc>"
End If

If (Not IsEmpty(NI_EI)) And UBound(NI_EI) > 0 Then

            writeXML "<ITRForm:OthersInc>"
            For i = 1 To UBound(NI_EI)

            If NI_EI(i) = "Defense Medical Disability Pension" Then
            NI_EI(i) = "DMDP"

            ElseIf (NI_EI(i) = "Any other") Then
                        NI_EI(i) = "OTH"
            Else
               NI_EI(i) = Mid(NI_EI(i), 5, (InStr(NI_EI(i), "-") - 5))
            End If

            Next


                For i = 1 To UBound(NI_EI)
                    writeXML "<ITRForm:OthersIncDtls>"
                        If NI_EI(i) <> "" Then
                            writeXML "<ITRForm:NatureDesc>" & NI_EI(i) & "</ITRForm:NatureDesc>"
                        Else
                            writeXML "<ITRForm:NatureDesc/>"
                        End If

                        If Description_EI(i) <> "" And Description_EI(i) <> "NA" Then
                        writeXML "<ITRForm:OthNatOfInc>" & UCase(Description_EI(i)) & "</ITRForm:OthNatOfInc>"

                        End If

                        If Amt_EI(i) <> "" Then
                            writeXML "<ITRForm:OthAmount>" & Amt_EI(i) & "</ITRForm:OthAmount>"
                        Else
                            writeXML "<ITRForm:OthAmount/>"
                        End If
                    writeXML "</ITRForm:OthersIncDtls>"
                Next
                writeXML "</ITRForm:OthersInc>"
End If




If sceiOthers <> "" Then
writeXML "<ITRForm:Others>" & sceiOthers & "</ITRForm:Others>"
Else
writeXML "<ITRForm:Others>0</ITRForm:Others>"
End If



writeXML "  <ITRForm:IncNotChrgblAsPerDTAA>"


For i = 1 To UBound(AmountOfIncome_EIDTAA)

writeXML "  <ITRForm:IncNotChrgblAsPerDTAADtls>"

If AmountOfIncome_EIDTAA(i) <> "" Then
writeXML "      <ITRForm:AmountOfIncome>" & AmountOfIncome_EIDTAA(i) & "</ITRForm:AmountOfIncome>"
Else
writeXML "      <ITRForm:AmountOfIncome>0</ITRForm:AmountOfIncome>"
End If

If NatureOfIncome_EIDTAA(i) <> "" Then
writeXML "      <ITRForm:NatureOfIncome>" & NatureOfIncome_EIDTAA(i) & "</ITRForm:NatureOfIncome>"
Else
writeXML "      <ITRForm:NatureOfIncome>0</ITRForm:NatureOfIncome>"
End If


If CountryNameA1_EIDTAA(i) <> "" Then
writeXML "      <ITRForm:CountryName>" & CountryNameA1_EIDTAA(i) & "</ITRForm:CountryName>"
Else
writeXML "      <ITRForm:CountryName>0</ITRForm:CountryName>"
End If


If CountryCodeA1_EIDTAA(i) <> "" Then
writeXML "      <ITRForm:CountryCode>" & CountryCodeA1_EIDTAA(i) & "</ITRForm:CountryCode>"
Else
writeXML "      <ITRForm:CountryCode>0</ITRForm:CountryCode>"
End If

If Article_EIDTAA(i) <> "" Then
writeXML "      <ITRForm:ArticleOfDTAA>" & Article_EIDTAA(i) & "</ITRForm:ArticleOfDTAA>"
Else
writeXML "      <ITRForm:ArticleOfDTAA>0</ITRForm:ArticleOfDTAA>"
End If

If HeadOfIncome_EIDTAA(i) <> "" Then

 If (HeadOfIncome_EIDTAA(i)) = "Salary" Then
    HeadOfIncome_EIDTAA(i) = "SA"
 ElseIf (HeadOfIncome_EIDTAA(i)) = "House Property" Then
    HeadOfIncome_EIDTAA(i) = "HP"
ElseIf (HeadOfIncome_EIDTAA(i)) = "Profits & Gains from Business & Profession" Then
    HeadOfIncome_EIDTAA(i) = "PG"
 ElseIf (HeadOfIncome_EIDTAA(i)) = "Capital Gain" Then
    HeadOfIncome_EIDTAA(i) = "CG"
 ElseIf (HeadOfIncome_EIDTAA(i)) = "Income from Other sources" Then
    HeadOfIncome_EIDTAA(i) = "OS"

End If

writeXML "<ITRForm:HeadOfIncome>" & HeadOfIncome_EIDTAA(i) & "</ITRForm:HeadOfIncome>"

End If


If TRC_EIDTAA(i) <> "" Then

 If UCase(TRC_EIDTAA(i)) = "YES" Then
    TRC_EIDTAA(i) = "Y"
 ElseIf UCase(TRC_EIDTAA(i)) = "NO" Then
    TRC_EIDTAA(i) = "N"

End If

writeXML "<ITRForm:TRCFlag>" & TRC_EIDTAA(i) & "</ITRForm:TRCFlag>"

End If


writeXML "  </ITRForm:IncNotChrgblAsPerDTAADtls>"

Next

writeXML "  </ITRForm:IncNotChrgblAsPerDTAA>"



If Sheet24.Range("SchEI_TotalIncomeDTAA").value <> "" Then
writeXML "      <ITRForm:IncChrgblAsPerDTAA>" & Sheet24.Range("SchEI_TotalIncomeDTAA").value & "</ITRForm:IncChrgblAsPerDTAA>"
Else
writeXML "      <ITRForm:IncChrgblAsPerDTAA>0</ITRForm:IncChrgblAsPerDTAA>"
End If


If Sheet24.Range("SCHEI_Pass_Through_Income").value <> "" Then
writeXML "      <ITRForm:PassThrIncNotChrgblTax>" & Sheet24.Range("SCHEI_Pass_Through_Income").value & "</ITRForm:PassThrIncNotChrgblTax>"
Else
writeXML "      <ITRForm:PassThrIncNotChrgblTax>0</ITRForm:PassThrIncNotChrgblTax>"
End If

If sceiTotalExemptInc <> "" Then
writeXML "      <ITRForm:TotalExemptInc>" & sceiTotalExemptInc & "</ITRForm:TotalExemptInc>"
Else
writeXML "      <ITRForm:TotalExemptInc>0</ITRForm:TotalExemptInc>"
End If

writeXML "</ITRForm:ScheduleEI>"

End Function

Sub XMLSchedulePTI()
subProcCaption = "Schedule PTI'"
noOfProcessSub = UBound(Investment115UAOfBusiness_PTI)

If Not IsEmpty(Investment115UAOfBusiness_PTI) And UBound(Investment115UAOfBusiness_PTI) > 0 Then
    writeXML "<ITRForm:SchedulePTI>"
        For i = 1 To UBound(Investment115UAOfBusiness_PTI)
            writeXML "<ITRForm:SchedulePTIDtls>"
              'Sheet27.Range("PTI_Investment_UA_UB" & i).value
                If Sheet41.Range("PTI_Investment_UA_UB" & i).value = "SECTION 115UA" Then
                   Sheet41.Range("PTI_Investment_UA_UB" & i).value = "A"
                   ElseIf Sheet41.Range("PTI_Investment_UA_UB" & i).value = "SECTION 115UB" Then
                   Sheet41.Range("PTI_Investment_UA_UB" & i).value = "B"
                End If

                If Sheet41.Range("PTI_Investment_UA_UB" & i).value <> "" Then
                    writeXML "<ITRForm:InvstmntCvrdUs115UA115UB>" & Sheet41.Range("PTI_Investment_UA_UB" & i).value & "</ITRForm:InvstmntCvrdUs115UA115UB>"
                End If

                If Sheet41.Range("PTI_NameOfBusiness" & i).value <> "" Then
                    writeXML "<ITRForm:BusinessName>" & Sheet41.Range("PTI_NameOfBusiness" & i).value & "</ITRForm:BusinessName>"
                End If

                If Sheet41.Range("PTI_PAN" & i).value <> "" Then
                    writeXML "<ITRForm:BusinessPAN>" & Sheet41.Range("PTI_PAN" & i).value & "</ITRForm:BusinessPAN>"
                End If
'
               writeXML "<ITRForm:IncFromHP>"
                    If Sheet41.Range("PTI_HPIncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_HPIncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_HP_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_HP_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If

                    If Sheet41.Range("PTI_HP_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_HP_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_HPTDSAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_HPTDSAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If

                writeXML "</ITRForm:IncFromHP>"
'
writeXML "<ITRForm:CapitalGainsPTI>"
        writeXML "<ITRForm:ShortTermCG>"
                   If Sheet41.Range("PTI_CG_STCGIncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_CG_STCGIncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_STCG_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_CG_STCG_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_CG_STCGTDSAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_CG_STCGTDSAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
           writeXML "</ITRForm:ShortTermCG>"
'
'
            writeXML "<ITRForm:STCG_Sec111A>"
                    If Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If
                    
                    
                    If Sheet41.Range("PTI_CG_STCG_SectionA_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_CG_STCG_SectionA_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
            writeXML "</ITRForm:STCG_Sec111A>"
            
            writeXML "<ITRForm:STCG_Others>"
                    If Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_Others_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_CG_STCG_Others_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
            writeXML "</ITRForm:STCG_Others>"
'
    writeXML "<ITRForm:LongTermCG>"
                    If Sheet41.Range("PTI_CG_LTCGIncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_CG_LTCGIncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_CG_LTCG_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCGTDSAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_CG_LTCGTDSAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
            writeXML "</ITRForm:LongTermCG>"
            
            
            writeXML "<ITRForm:LTCG_Sec112A>"
                    If Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_SectionA_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_CG_LTCG_SectionA_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
            writeXML "</ITRForm:LTCG_Sec112A>"
'
         writeXML "<ITRForm:LTCG_Others>"
                    If Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
                    Else
                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_Others_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_CG_LTCG_Others_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
            writeXML "</ITRForm:LTCG_Others>"
               writeXML "</ITRForm:CapitalGainsPTI>"
               
               
               
                writeXML "<ITRForm:IncClmdPTI>"
                
                
                       writeXML "<ITRForm:TotalSec23FBB>"
                    If Sheet41.Range("PTI_Us_iv_IncomeAmta" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_Us_iv_IncomeAmta" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
'                    Else
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_Us_iv_NetAmta" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_Us_iv_NetAmta" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_Us_iv_TdsAmta" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_Us_iv_TdsAmta" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'PTI_Usb1
                   writeXML "</ITRForm:TotalSec23FBB>"
                
                
                
                
                    writeXML "<ITRForm:Sec23FBB>"
                    If Sheet41.Range("PTI_UsIncomeAmta" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_UsIncomeAmta" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
'                    Else
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_Us_NetAmta" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_Us_NetAmta" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_UsTDSAmta" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_UsTDSAmta" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'PTI_Usb1
                   writeXML "</ITRForm:Sec23FBB>"
                   
                   If Sheet41.Range("PTI_Usb" & i).value <> "" Then
                        writeXML "<ITRForm:SecBIncExmptDtl>"
                           If Sheet41.Range("PTI_Usb" & i).value <> "" Then
                               writeXML "<ITRForm:SectionCode>" & Sheet41.Range("PTI_Usb" & i).value & "</ITRForm:SectionCode>"
                           Else
                                writeXML "<ITRForm:SectionCode>0</ITRForm:SectionCode>"
                            End If
                           writeXML "<ITRForm:SecBCIncExmptDtl>"
   
                   If Sheet41.Range("PTI_UsIncomeAmtb" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_UsIncomeAmtb" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_Us_NetAmtb" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_Us_NetAmtb" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_UsTDSAmtb" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_UsTDSAmtb" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
                           writeXML "</ITRForm:SecBCIncExmptDtl>"
                        writeXML "</ITRForm:SecBIncExmptDtl>"
                   End If
'
               If Sheet41.Range("PTI_Usc" & i).value <> "" Then
                        writeXML "<ITRForm:SecCIncExmptDtl>"
                           If Sheet41.Range("PTI_Usc" & i).value <> "" Then
                               writeXML "<ITRForm:SectionCode>" & Sheet41.Range("PTI_Usc" & i).value & "</ITRForm:SectionCode>"
                           Else
                                writeXML "<ITRForm:SectionCode>0</ITRForm:SectionCode>"
                            End If
                           writeXML "<ITRForm:SecBCIncExmptDtl>"
   
                   If Sheet41.Range("PTI_UsIncomeAmtc" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_UsIncomeAmtc" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_Us_NetAmtc" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_Us_NetAmtc" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_UsTDSAmtc" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_UsTDSAmtc" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
                           writeXML "</ITRForm:SecBCIncExmptDtl>"
                        writeXML "</ITRForm:SecCIncExmptDtl>"
                   End If
writeXML "</ITRForm:IncClmdPTI>"


               writeXML "<ITRForm:IncOthSrc>"
                   If Sheet41.Range("PTI_OSIncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_OSIncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
'                    Else
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_OS_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_OS_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_OSTDSAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_OSTDSAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
               writeXML "</ITRForm:IncOthSrc>"
'
               writeXML "<ITRForm:OS_Dividend>"
                   If Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
'                    Else
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_OS_Dividend_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_OS_Dividend_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_OS_Dividend_TdsAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_OS_Dividend_TdsAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
               writeXML "</ITRForm:OS_Dividend>"
               
                writeXML "<ITRForm:OS_Others>"
                   If Sheet41.Range("PTI_OS_Others_IncomeAmt" & i).value <> "" Then
                        writeXML "<ITRForm:AmountOfInc>" & Sheet41.Range("PTI_OS_Others_IncomeAmt" & i).value & "</ITRForm:AmountOfInc>"
                    Else
                        writeXML "<ITRForm:AmountOfInc>0</ITRForm:AmountOfInc>"
                    End If

'                    If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value <> "" Then
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>" & Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value & "</ITRForm:CurrYrLossShareByInvstFund>"
'                    Else
'                        writeXML "<ITRForm:CurrYrLossShareByInvstFund>0</ITRForm:CurrYrLossShareByInvstFund>"
'                    End If

                    If Sheet41.Range("PTI_OS_Others_NetAmt" & i).value <> "" Then
                        writeXML "<ITRForm:NetIncomeLoss>" & Sheet41.Range("PTI_OS_Others_NetAmt" & i).value & "</ITRForm:NetIncomeLoss>"
                    Else
                        writeXML "<ITRForm:NetIncomeLoss>0</ITRForm:NetIncomeLoss>"
                    End If

                    If Sheet41.Range("PTI_OS_Others_TdsAmt" & i).value <> "" Then
                        writeXML "<ITRForm:TDSAmount>" & Sheet41.Range("PTI_OS_Others_TdsAmt" & i).value & "</ITRForm:TDSAmount>"
                    Else
                        writeXML "<ITRForm:TDSAmount>0</ITRForm:TDSAmount>"
                    End If
'
               writeXML "</ITRForm:OS_Others>"
'
           writeXML "</ITRForm:SchedulePTIDtls>"
            UpdateProgressBar
       Next
       writeXML "</ITRForm:SchedulePTI>"
 End If
End Sub

Function ScheduleIT()
 Dim i As Long
 subProcCaption = "Schedule IT"
    noOfProcessSub = UBound(BSR_TDS)
    If Not IsEmpty(BSR_TDS) And UBound(BSR_TDS) > 0 Then
        writeXML "      <ITRForm:ScheduleIT>"
        For i = 1 To UBound(BSR_TDS)
            writeXML "           <ITRForm:TaxPayment>"
            If BSR_TDS(i) <> "" Then
                writeXML "               <ITRForm:BSRCode>" & UCase(BSR_TDS(i)) & "</ITRForm:BSRCode>"
            Else
                writeXML "               <ITRForm:BSRCode/>"
            End If

            If DateCredit_TDS(i) <> "" Then
                writeXML "               <ITRForm:DateDep>" & Dformat(DateCredit_TDS(i), "") & "</ITRForm:DateDep>"
            Else
                writeXML "               <ITRForm:DateDep/>"
            End If

            If SerialNum_TDS(i) <> "" Then
                writeXML "               <ITRForm:SrlNoOfChaln>" & UCase(SerialNum_TDS(i)) & "</ITRForm:SrlNoOfChaln>"
            Else
                writeXML "               <ITRForm:SrlNoOfChaln/>"
            End If

            If TaxPaid3_TDS(i) <> "" Then
                writeXML "               <ITRForm:Amt>" & UCase(TaxPaid3_TDS(i)) & "</ITRForm:Amt>"
            Else
                writeXML "               <ITRForm:Amt>0</ITRForm:Amt>"
            End If
            writeXML "           </ITRForm:TaxPayment>"
            UpdateProgressBar
        Next
        writeXML "       <ITRForm:TotalTaxPayments>" & Sheet25.Range("IT.Sum").value & "</ITRForm:TotalTaxPayments>"
        writeXML "      </ITRForm:ScheduleIT>"
    End If
End Function

Function ScheduleTDS()
Dim i As Variant
If Sheet1.Range("sheet1.Status").value <> "H-HUF" Then
If Not IsEmpty(TAN_TDS1) And UBound(TAN_TDS1) > 0 Then
writeXML "      <ITRForm:ScheduleTDS1>"
subProcCaption = "Schedule TDS"
For i = 1 To UBound(TAN_TDS1)
noOfProcessSub = UBound(TAN_TDS1)
    writeXML "           <ITRForm:TDSonSalary>"
    writeXML "              <ITRForm:EmployerOrDeductorOrCollectDetl>"
    If TAN_TDS1(i) <> "" Then
    writeXML "                   <ITRForm:TAN>" & UCase(TAN_TDS1(i)) & "</ITRForm:TAN>"
    Else
    writeXML "                   <ITRForm:TAN/>"
    End If
    If EmployerOrDeductorOrCollecterName_TDS1(i) <> "" Then
    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName>" & UCase(EmployerOrDeductorOrCollecterName_TDS1(i)) & "</ITRForm:EmployerOrDeductorOrCollecterName>"
    Else
    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName/>"
    End If
    writeXML "              </ITRForm:EmployerOrDeductorOrCollectDetl>"
    If IncChrgSal_TDS1(i) <> "" Then
    writeXML "               <ITRForm:IncChrgSal>" & UCase(IncChrgSal_TDS1(i)) & "</ITRForm:IncChrgSal>"
    Else
    writeXML "               <ITRForm:IncChrgSal> 0</ITRForm:IncChrgSal>"
    End If
    If TotalTDSSal_TDS1(i) <> "" Then
    writeXML "               <ITRForm:TotalTDSSal>" & UCase(TotalTDSSal_TDS1(i)) & "</ITRForm:TotalTDSSal>"
    Else
    writeXML "               <ITRForm:TotalTDSSal> 0 </ITRForm:TotalTDSSal>"
    End If
    writeXML "           </ITRForm:TDSonSalary>"

    UpdateProgressBar
Next
writeXML "<ITRForm:TotalTDSonSalaries>" & Sheet28.Range("TDS1.Sum").value & "</ITRForm:TotalTDSonSalaries>"
writeXML "      </ITRForm:ScheduleTDS1>"
End If
End If

If Not IsEmpty(TAN_TDS2) And UBound(TAN_TDS2) > 0 Then
writeXML "<ITRForm:ScheduleTDS2>"
subProcCaption = "TDS2"
noOfProcessSub = UBound(TAN_TDS2)
For i = 1 To UBound(TAN_TDS2)
    writeXML "<ITRForm:TDSOthThanSalaryDtls>"

    If TDSCreditName_TDS2(i) <> "" Then
    If UCase(TDSCreditName_TDS2(i)) = "SELF" Then TDSCreditName_TDS2(i) = "S"
    If UCase(TDSCreditName_TDS2(i)) = "OTHER PERSON" Then TDSCreditName_TDS2(i) = "O"
    writeXML "<ITRForm:TDSCreditName>" & UCase(TDSCreditName_TDS2(i)) & "</ITRForm:TDSCreditName>"
    End If


    If TDSPanOfOther_TDS2(i) <> "" Then
    writeXML "<ITRForm:PANofOtherPerson>" & UCase(TDSPanOfOther_TDS2(i)) & "</ITRForm:PANofOtherPerson>"
    End If

    If TDSAadharOfOther_TDS2(i) <> "" Then
    writeXML "<ITRForm:AadhaarOfOtherPerson>" & UCase(TDSAadharOfOther_TDS2(i)) & "</ITRForm:AadhaarOfOtherPerson>"
    End If

    If TAN_TDS2(i) <> "" Then
    writeXML "<ITRForm:TANOfDeductor>" & UCase(TAN_TDS2(i)) & "</ITRForm:TANOfDeductor>"
    Else
    writeXML "<ITRForm:TANOfDeductor/>"
    End If


    If DeductedYear_TDS2(i) <> "" Then
    writeXML "<ITRForm:DeductedYr>" & UCase(DeductedYear_TDS2(i)) & "</ITRForm:DeductedYr>"
'    Else
'    writeXML "<ITRForm:DeductedYr>0000</ITRForm:DeductedYr>"
    End If

    If BroughtFwdTDSAmt_TDS2(i) <> "" Then
    writeXML "<ITRForm:BroughtFwdTDSAmt>" & UVCase(BroughtFwdTDSAmt_TDS2(i)) & "</ITRForm:BroughtFwdTDSAmt>"
    End If

    writeXML "<ITRForm:TaxDeductCreditDtls>"

    If DeductedInOwnHands_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxDeductedOwnHands>" & UVCase(DeductedInOwnHands_TDS2(i)) & "</ITRForm:TaxDeductedOwnHands>"
    End If

    If DeductedIncome5A_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxDeductedIncome>" & UVCase(DeductedIncome5A_TDS2(i)) & "</ITRForm:TaxDeductedIncome>"
    End If

    If DeductedTDS5A_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxDeductedTDS>" & UVCase(DeductedTDS5A_TDS2(i)) & "</ITRForm:TaxDeductedTDS>"
    End If


    If TDSClaimedInOwnHands_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedOwnHands>" & UVCase(TDSClaimedInOwnHands_TDS2(i)) & "</ITRForm:TaxClaimedOwnHands>"
    End If

    If ClaimedIncome5A_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedIncome>" & UVCase(ClaimedIncome5A_TDS2(i)) & "</ITRForm:TaxClaimedIncome>"
    End If

    If ClaimedTDS5A_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedTDS>" & UVCase(ClaimedTDS5A_TDS2(i)) & "</ITRForm:TaxClaimedTDS>"
    End If

    If ClaimedPAN5A_TDS2(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedSpouseOthPrsnPAN>" & UCase(ClaimedPAN5A_TDS2(i)) & "</ITRForm:TaxClaimedSpouseOthPrsnPAN>"
    End If

    If ClaimedAadhar5A_TDS2(i) <> "" Then
    writeXML "<ITRForm:SpouseOthPrsnAadhaar>" & UCase(ClaimedAadhar5A_TDS2(i)) & "</ITRForm:SpouseOthPrsnAadhaar>"
    End If

    writeXML "</ITRForm:TaxDeductCreditDtls>"


If GrossAmount_TDS2(i) <> "" Then
    writeXML "<ITRForm:GrossAmount>" & UVCase(GrossAmount_TDS2(i)) & "</ITRForm:GrossAmount>"
End If


 If HeadIncome_TDS2(i) <> "" Then

 If (HeadIncome_TDS2(i)) = "Income from House Property" Then
    HeadIncome_TDS2(i) = "HP"
 ElseIf (HeadIncome_TDS2(i)) = "Income from Business & Profession" Then
    HeadIncome_TDS2(i) = "BP"
 ElseIf (HeadIncome_TDS2(i)) = "Income from Capital Gains" Then
    HeadIncome_TDS2(i) = "CG"
 ElseIf (HeadIncome_TDS2(i)) = "Income from Other Sources" Then
    HeadIncome_TDS2(i) = "OS"
 ElseIf (HeadIncome_TDS2(i)) = "Exempt Income" Then
    HeadIncome_TDS2(i) = "EI"
ElseIf (HeadIncome_TDS2(i)) = "Not applicable ( only in case TDS is deducted u/s 194N)" Then
    HeadIncome_TDS2(i) = "NA"
 End If

 writeXML "<ITRForm:HeadOfIncome>" & UCase(HeadIncome_TDS2(i)) & "</ITRForm:HeadOfIncome>"
 End If


    If AmtCarriedForward_TDS2(i) <> "" Then
    writeXML "<ITRForm:AmtCarriedFwd>" & UVCase(AmtCarriedForward_TDS2(i)) & "</ITRForm:AmtCarriedFwd>"
    End If


    writeXML "</ITRForm:TDSOthThanSalaryDtls>"
      UpdateProgressBar
Next
writeXML "<ITRForm:TotalTDSonOthThanSals>" & Sheet28.Range("TDS2.Total").value & "</ITRForm:TotalTDSonOthThanSals>"
writeXML "</ITRForm:ScheduleTDS2>"
End If


'
'TDS3
If Not IsEmpty(PAN_TDS3) And UBound(PAN_TDS3) > 0 Then
writeXML "<ITRForm:ScheduleTDS3>"
subProcCaption = "TDS3"
noOfProcessSub = UBound(PAN_TDS3)

For i = 1 To UBound(PAN_TDS3)
    writeXML "<ITRForm:TDS3onOthThanSalDtls>"

    If TDSCreditName_TDS3(i) <> "" Then
    If UCase(TDSCreditName_TDS3(i)) = "SELF" Then TDSCreditName_TDS3(i) = "S"
    If UCase(TDSCreditName_TDS3(i)) = "OTHER PERSON" Then TDSCreditName_TDS3(i) = "O"
    writeXML "<ITRForm:TDSCreditName>" & UCase(TDSCreditName_TDS3(i)) & "</ITRForm:TDSCreditName>"
    End If

    If TDSPanOfOther_TDS3(i) <> "" Then
    writeXML "<ITRForm:PANofOtherPerson>" & UCase(TDSPanOfOther_TDS3(i)) & "</ITRForm:PANofOtherPerson>"
    End If
    
    If Aadhar_TDS3(i) <> "" Then
    writeXML "<ITRForm:AadhaarOfOtherPerson>" & UCase(Aadhar_TDS3(i)) & "</ITRForm:AadhaarOfOtherPerson>"
    End If

    If PAN_TDS3(i) <> "" Then
    writeXML "<ITRForm:PANOfBuyerTenant>" & UCase(PAN_TDS3(i)) & "</ITRForm:PANOfBuyerTenant>"
    End If

    If TDS3_Aadhaar(i) <> "" Then
    writeXML "<ITRForm:AadhaarOfBuyerTenant>" & UCase(TDS3_Aadhaar(i)) & "</ITRForm:AadhaarOfBuyerTenant>"
    End If


    If DeductedYear_TDS3(i) <> "" Then
    writeXML "<ITRForm:DeductedYr>" & UCase(DeductedYear_TDS3(i)) & "</ITRForm:DeductedYr>"
'    Else
'    writeXML "<ITRForm:DeductedYr>0000</ITRForm:DeductedYr>"
    End If

    If BroughtFwdTDSAmt_TDS3(i) <> "" Then
    writeXML "<ITRForm:BroughtFwdTDSAmt>" & UVCase(BroughtFwdTDSAmt_TDS3(i)) & "</ITRForm:BroughtFwdTDSAmt>"
    End If

    writeXML "<ITRForm:TaxDeductCreditDtls>"

    If DeductedInOwnHands_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxDeductedOwnHands>" & UVCase(DeductedInOwnHands_TDS3(i)) & "</ITRForm:TaxDeductedOwnHands>"
    End If

    If DeductedTDS5A_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxDeductedIncome>" & UVCase(DeductedIncome5A_TDS3(i)) & "</ITRForm:TaxDeductedIncome>"
    End If

    If DeductedTDS5A_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxDeductedTDS>" & UVCase(DeductedTDS5A_TDS3(i)) & "</ITRForm:TaxDeductedTDS>"
    End If

'    If DeductedPAN5A_TDS3(i) <> "" Then
'    writeXML "<ITRForm:TaxDeductedSpouseOthPrsnPAN>" & UCase(DeductedPAN5A_TDS3(i)) & "</ITRForm:TaxDeductedSpouseOthPrsnPAN>"
'    End If

    If TDSClaimedInOwnHands_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedOwnHands>" & UVCase(TDSClaimedInOwnHands_TDS3(i)) & "</ITRForm:TaxClaimedOwnHands>"
    End If

    If ClaimedIncome5A_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedIncome>" & UVCase(ClaimedIncome5A_TDS3(i)) & "</ITRForm:TaxClaimedIncome>"
    End If

    If ClaimedTDS5A_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedTDS>" & UVCase(ClaimedTDS5A_TDS3(i)) & "</ITRForm:TaxClaimedTDS>"
    End If

    If ClaimedPAN5A_TDS3(i) <> "" Then
    writeXML "<ITRForm:TaxClaimedSpouseOthPrsnPAN>" & UCase(ClaimedPAN5A_TDS3(i)) & "</ITRForm:TaxClaimedSpouseOthPrsnPAN>"
    End If

    If ClaimedAadhar5A_TDS3(i) <> "" Then
    writeXML "<ITRForm:SpouseOthPrsnAadhaar>" & UCase(ClaimedAadhar5A_TDS3(i)) & "</ITRForm:SpouseOthPrsnAadhaar>"
    End If

    writeXML "</ITRForm:TaxDeductCreditDtls>"


    If GrossAmount_TDS3(i) <> "" Then
    writeXML "<ITRForm:GrossAmount>" & UVCase(GrossAmount_TDS3(i)) & "</ITRForm:GrossAmount>"
End If

  If (HeadIncome_TDS3(i)) = "Income from House Property" Then
    HeadIncome_TDS3(i) = "HP"
 ElseIf (HeadIncome_TDS3(i)) = "Income from Business & Profession" Then
    HeadIncome_TDS3(i) = "BP"
 ElseIf (HeadIncome_TDS3(i)) = "Income from Capital Gains" Then
    HeadIncome_TDS3(i) = "CG"
 ElseIf (HeadIncome_TDS3(i)) = "Income from Other Sources" Then
    HeadIncome_TDS3(i) = "OS"
 ElseIf (HeadIncome_TDS3(i)) = "Exempt Income" Then
    HeadIncome_TDS3(i) = "EI"
 End If

 If HeadIncome_TDS3(i) <> "" Then
 writeXML "<ITRForm:HeadOfIncome>" & UCase(HeadIncome_TDS3(i)) & "</ITRForm:HeadOfIncome>"
 End If



    If AmtCarriedForward_TDS3(i) <> "" Then
    writeXML "<ITRForm:AmtCarriedFwd>" & UVCase(AmtCarriedForward_TDS3(i)) & "</ITRForm:AmtCarriedFwd>"
    End If

    writeXML "</ITRForm:TDS3onOthThanSalDtls>"
    UpdateProgressBar
Next
writeXML "<ITRForm:TotalTDS3OnOthThanSal>" & Sheet28.Range("TDS3.Total").value & "</ITRForm:TotalTDS3OnOthThanSal>"
writeXML "</ITRForm:ScheduleTDS3>"
End If

End Function




Function ScheduleTCS()

Dim i As Variant
If Not IsEmpty(TAN_TCS) And UBound(TAN_TCS) > 0 Then
noOfProcessSub = UBound(TAN_TCS)
writeXML "      <ITRForm:ScheduleTCS>"
For i = 1 To UBound(TAN_TCS)
    writeXML "           <ITRForm:TCS>"
    writeXML "              <ITRForm:EmployerOrDeductorOrCollectDetl>"
    If TAN_TCS(i) <> "" Then
    writeXML "                   <ITRForm:TAN>" & UCase(TAN_TCS(i)) & "</ITRForm:TAN>"
    Else
    writeXML "                   <ITRForm:TAN/>"
    End If
    If EmployerOrDeductorOrCollecterName_TCS(i) <> "" Then
    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName>" & UCase(EmployerOrDeductorOrCollecterName_TCS(i)) & "</ITRForm:EmployerOrDeductorOrCollecterName>"
    Else
    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName/>"
    End If
    writeXML "              </ITRForm:EmployerOrDeductorOrCollectDetl>"

    If DeductedYear_TCS(i) <> "" Then
    writeXML "               <ITRForm:DeductedYr>" & UCase(DeductedYear_TCS(i)) & "</ITRForm:DeductedYr>"
    End If

    If BroughtFwdTDSAmt_TCS(i) <> "" Then
    writeXML "               <ITRForm:BroughtFwdTDSAmt>" & UCase(BroughtFwdTDSAmt_TCS(i)) & "</ITRForm:BroughtFwdTDSAmt>"
    Else
    writeXML "               <ITRForm:BroughtFwdTDSAmt>0</ITRForm:BroughtFwdTDSAmt>"
    End If


    If TotalTCS_TCS(i) <> "" Then
    writeXML "               <ITRForm:TotalTCS>" & UCase(TotalTCS_TCS(i)) & "</ITRForm:TotalTCS>"
'    Else
'    writeXML "               <ITRForm:TotalTCS>0</ITRForm:TotalTCS>"
    End If

    If AmtTCSClaimedThisYear_TCS(i) <> "" Then
    writeXML "               <ITRForm:AmtTCSClaimedThisYear>" & UCase(AmtTCSClaimedThisYear_TCS(i)) & "</ITRForm:AmtTCSClaimedThisYear>"
    Else
    writeXML "               <ITRForm:AmtTCSClaimedThisYear>0</ITRForm:AmtTCSClaimedThisYear>"
    End If

    If AmtCarriedFwd_TCS(i) <> "" Then
    writeXML "               <ITRForm:AmtCarriedFwd>" & UCase(AmtCarriedFwd_TCS(i)) & "</ITRForm:AmtCarriedFwd>"
    Else
    writeXML "               <ITRForm:AmtCarriedFwd>0</ITRForm:AmtCarriedFwd>"
    End If
    writeXML "           </ITRForm:TCS>"
    UpdateProgressBar
Next

    writeXML "<ITRForm:TotalSchTCS>" & Sheet28.Range("TCS.Sum").value & "</ITRForm:TotalSchTCS>"

writeXML "      </ITRForm:ScheduleTCS>"
End If

End Function

Function ScheduleFSI()
    Dim i As Long
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
    If Not IsEmpty(Country_FSI) And UBound(Country_FSI) > 0 Then
        writeXML "  <ITRForm:ScheduleFSI>"
        For i = 1 To UBound(Country_FSI)
            writeXML "        <ITRForm:ScheduleFSIDtls>"

            If CountryName_FSI(i) <> "" Then
                writeXML "            <ITRForm:CountryName>" & CountryName_FSI(i) & "</ITRForm:CountryName>"
            Else
            End If

            If Country_FSI(i) <> "" Then
                writeXML "            <ITRForm:CountryCode>" & Country_FSI(i) & "</ITRForm:CountryCode>"
            Else
            End If

            If TaxIdentificationNo_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxIdentificationNo>" & TaxIdentificationNo_FSI(i) & "</ITRForm:TaxIdentificationNo>"
            Else
            End If

            writeXML "            <ITRForm:IncFromSal>"
            If IncFromSal_FSI(i) <> "" Then
                writeXML "            <ITRForm:IncFrmOutsideInd>" & IncFromSal_FSI(i) & "</ITRForm:IncFrmOutsideInd>"
            Else
                writeXML "            <ITRForm:IncFrmOutsideInd>0</ITRForm:IncFrmOutsideInd>"
            End If

            If TXNOSAL_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPaidOutsideInd>" & TXNOSAL_FSI(i) & "</ITRForm:TaxPaidOutsideInd>"
            Else
                writeXML "            <ITRForm:TaxPaidOutsideInd>0</ITRForm:TaxPaidOutsideInd>"
            End If

            If TXNISAL_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPayableinInd>" & TXNISAL_FSI(i) & "</ITRForm:TaxPayableinInd>"
            Else
                writeXML "            <ITRForm:TaxPayableinInd>0</ITRForm:TaxPayableinInd>"
            End If

            If TXRSAL_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxReliefinInd>" & TXRSAL_FSI(i) & "</ITRForm:TaxReliefinInd>"
            Else
                writeXML "            <ITRForm:TaxReliefinInd>0</ITRForm:TaxReliefinInd>"
            End If

            If DTAASAL_FSI(i) <> "" Then
                writeXML "            <ITRForm:DTAAReliefUs90or90A>" & DTAASAL_FSI(i) & "</ITRForm:DTAAReliefUs90or90A>"
            Else
                writeXML "            <ITRForm:DTAAReliefUs90or90A>0</ITRForm:DTAAReliefUs90or90A>"
            End If

            writeXML "            </ITRForm:IncFromSal>"

            writeXML "            <ITRForm:IncFromHP>"
            If IncFromHP_FSI(i) <> "" Then
                writeXML "            <ITRForm:IncFrmOutsideInd>" & IncFromHP_FSI(i) & "</ITRForm:IncFrmOutsideInd>"
            Else
                writeXML "            <ITRForm:IncFrmOutsideInd>0</ITRForm:IncFrmOutsideInd>"
            End If

            If TXNOHP_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPaidOutsideInd>" & TXNOHP_FSI(i) & "</ITRForm:TaxPaidOutsideInd>"
            Else
                writeXML "            <ITRForm:TaxPaidOutsideInd>0</ITRForm:TaxPaidOutsideInd>"
            End If

            If TXNIHP_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPayableinInd>" & TXNIHP_FSI(i) & "</ITRForm:TaxPayableinInd>"
            Else
                writeXML "            <ITRForm:TaxPayableinInd>0</ITRForm:TaxPayableinInd>"
            End If

            If TXRHP_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxReliefinInd>" & TXRHP_FSI(i) & "</ITRForm:TaxReliefinInd>"
            Else
                writeXML "            <ITRForm:TaxReliefinInd>0</ITRForm:TaxReliefinInd>"
            End If

            If DTAAHP_FSI(i) <> "" Then
                writeXML "            <ITRForm:DTAAReliefUs90or90A>" & DTAAHP_FSI(i) & "</ITRForm:DTAAReliefUs90or90A>"
            Else
                writeXML "            <ITRForm:DTAAReliefUs90or90A>0</ITRForm:DTAAReliefUs90or90A>"
            End If

            writeXML "            </ITRForm:IncFromHP>"

            writeXML "            <ITRForm:IncFromBusiness>"
            If IncFromBP_FSI(i) <> "" Then
                writeXML "            <ITRForm:IncFrmOutsideInd>" & IncFromBP_FSI(i) & "</ITRForm:IncFrmOutsideInd>"
            Else
                writeXML "            <ITRForm:IncFrmOutsideInd>0</ITRForm:IncFrmOutsideInd>"
            End If

            If TXNOBP_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPaidOutsideInd>" & TXNOBP_FSI(i) & "</ITRForm:TaxPaidOutsideInd>"
            Else
                writeXML "            <ITRForm:TaxPaidOutsideInd>0</ITRForm:TaxPaidOutsideInd>"
            End If

            If TXNIBP_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPayableinInd>" & TXNIBP_FSI(i) & "</ITRForm:TaxPayableinInd>"
            Else
                writeXML "            <ITRForm:TaxPayableinInd>0</ITRForm:TaxPayableinInd>"
            End If

            If TXRBP_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxReliefinInd>" & TXRBP_FSI(i) & "</ITRForm:TaxReliefinInd>"
            Else
                writeXML "            <ITRForm:TaxReliefinInd>0</ITRForm:TaxReliefinInd>"
            End If

            If DTAABP_FSI(i) <> "" Then
                writeXML "            <ITRForm:DTAAReliefUs90or90A>" & DTAABP_FSI(i) & "</ITRForm:DTAAReliefUs90or90A>"
            Else
                writeXML "            <ITRForm:DTAAReliefUs90or90A>0</ITRForm:DTAAReliefUs90or90A>"
            End If

            writeXML "            </ITRForm:IncFromBusiness>"

            writeXML "            <ITRForm:IncCapGain>"

            If IncCapGain_FSI(i) <> "" Then
                writeXML "            <ITRForm:IncFrmOutsideInd>" & IncCapGain_FSI(i) & "</ITRForm:IncFrmOutsideInd>"
            Else
                writeXML "            <ITRForm:IncFrmOutsideInd>0</ITRForm:IncFrmOutsideInd>"
            End If

            If TXNOCapGain_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPaidOutsideInd>" & TXNOCapGain_FSI(i) & "</ITRForm:TaxPaidOutsideInd>"
            Else
                writeXML "            <ITRForm:TaxPaidOutsideInd>0</ITRForm:TaxPaidOutsideInd>"
            End If

            If TXNICapGain_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPayableinInd>" & TXNICapGain_FSI(i) & "</ITRForm:TaxPayableinInd>"
            Else
                writeXML "            <ITRForm:TaxPayableinInd>0</ITRForm:TaxPayableinInd>"
            End If

            If TXRCapGain_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxReliefinInd>" & TXRCapGain_FSI(i) & "</ITRForm:TaxReliefinInd>"
            Else
                writeXML "            <ITRForm:TaxReliefinInd>0</ITRForm:TaxReliefinInd>"
            End If

            If DTAACapGain_FSI(i) <> "" Then
                writeXML "            <ITRForm:DTAAReliefUs90or90A>" & DTAACapGain_FSI(i) & "</ITRForm:DTAAReliefUs90or90A>"
            Else
                writeXML "            <ITRForm:DTAAReliefUs90or90A>0</ITRForm:DTAAReliefUs90or90A>"
            End If

            writeXML "            </ITRForm:IncCapGain>"

            writeXML "            <ITRForm:IncOthSrc>"
            If IncOthSrc_FSI(i) <> "" Then
                writeXML "            <ITRForm:IncFrmOutsideInd>" & IncOthSrc_FSI(i) & "</ITRForm:IncFrmOutsideInd>"
            Else
                writeXML "            <ITRForm:IncFrmOutsideInd>0</ITRForm:IncFrmOutsideInd>"
            End If

            If TXNOOthSrc_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPaidOutsideInd>" & TXNOOthSrc_FSI(i) & "</ITRForm:TaxPaidOutsideInd>"
            Else
                writeXML "            <ITRForm:TaxPaidOutsideInd>0</ITRForm:TaxPaidOutsideInd>"
            End If

            If TXNIOthSrc_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxPayableinInd>" & TXNIOthSrc_FSI(i) & "</ITRForm:TaxPayableinInd>"
            Else
                writeXML "            <ITRForm:TaxPayableinInd>0</ITRForm:TaxPayableinInd>"
            End If

            If TXROthSrc_FSI(i) <> "" Then
                writeXML "            <ITRForm:TaxReliefinInd>" & TXROthSrc_FSI(i) & "</ITRForm:TaxReliefinInd>"
            Else
                writeXML "            <ITRForm:TaxReliefinInd>0</ITRForm:TaxReliefinInd>"
            End If

            If DTAAOthSrc_FSI(i) <> "" Then
                writeXML "            <ITRForm:DTAAReliefUs90or90A>" & DTAAOthSrc_FSI(i) & "</ITRForm:DTAAReliefUs90or90A>"
            Else
                writeXML "            <ITRForm:DTAAReliefUs90or90A>0</ITRForm:DTAAReliefUs90or90A>"
            End If

            writeXML "            </ITRForm:IncOthSrc>"

            writeXML "<ITRForm:TotalCountryWise>"
            writeXML "<ITRForm:IncFrmOutsideInd>" & IncT_FSI(i) & "</ITRForm:IncFrmOutsideInd>"
            writeXML "<ITRForm:TaxPaidOutsideInd>" & TXNOT_FSI(i) & "</ITRForm:TaxPaidOutsideInd>"
            writeXML "<ITRForm:TaxPayableinInd>" & TXNIT_FSI(i) & "</ITRForm:TaxPayableinInd>"
            writeXML "<ITRForm:TaxReliefinInd>" & TXRT_FSI(i) & "</ITRForm:TaxReliefinInd>"
            writeXML "</ITRForm:TotalCountryWise>"

            writeXML "        </ITRForm:ScheduleFSIDtls>"
        Next
        writeXML "  </ITRForm:ScheduleFSI>"
    End If
End If
End Function

Function ScheduleTR1()
    Dim i As Long
    Dim Fsicount As Long
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
    If TotalTaxPaidOutsideIndia_TR <> 0 Then
        writeXML "  <ITRForm:ScheduleTR1>"
        If Not IsEmpty(Country_FSI) And UBound(Country_FSI) > 0 Then
            Fsicount = 0
            For i = 1 To UBound(CountryName_TR)
                If (TaxPaidOutsideIndia_TR(i) > 0 Or TaxReliefOutsideIndia_TR(i) > 0) Then
                    writeXML "        <ITRForm:ScheduleTR>"
                    If CountryName_TR(i) <> "" Then
                        Fsicount = Fsicount + 1
                        writeXML "            <ITRForm:CountryName>" & CountryName_TR(i) & "</ITRForm:CountryName>"
                    Else
                    End If

                    If Country_TR(i) <> "" Then
                        writeXML "            <ITRForm:CountryCode>" & Country_TR(i) & "</ITRForm:CountryCode>"
                    Else
                    End If

                    If TIN_TR(i) <> "" Then
                        writeXML "            <ITRForm:TaxIdentificationNo>" & TIN_TR(i) & "</ITRForm:TaxIdentificationNo>"
                    Else
                        writeXML "            <ITRForm:TaxIdentificationNo></ITRForm:TaxIdentificationNo>"
                    End If

                    If TaxPaidOutsideIndia_TR(i) <> "" Then
                        writeXML "            <ITRForm:TaxPaidOutsideIndia>" & TaxPaidOutsideIndia_TR(i) & "</ITRForm:TaxPaidOutsideIndia>"
                    Else
                        writeXML "            <ITRForm:TaxPaidOutsideIndia>0</ITRForm:TaxPaidOutsideIndia>"
                    End If

                    If TaxReliefOutsideIndia_TR(i) <> "" Then
                        writeXML "            <ITRForm:TaxReliefOutsideIndia>" & TaxReliefOutsideIndia_TR(i) & "</ITRForm:TaxReliefOutsideIndia>"
                    Else
                        writeXML "            <ITRForm:TaxReliefOutsideIndia>0</ITRForm:TaxReliefOutsideIndia>"
                    End If

                    If ReliefClaimedUsSection_TR(i) <> "" And ReliefClaimedUsSection_TR(i) <> "(Select)" Then
                        writeXML "            <ITRForm:ReliefClaimedUsSection>" & UCase(ReliefClaimedUsSection_TR(i)) & "</ITRForm:ReliefClaimedUsSection>"
                    Else
                        writeXML "            <ITRForm:ReliefClaimedUsSection></ITRForm:ReliefClaimedUsSection>"
                    End If

                    writeXML "        </ITRForm:ScheduleTR>"
                End If
            Next
        End If


        writeXML "<ITRForm:TotalTaxPaidOutsideIndia>" & TotalTaxPaidOutsideIndia_TR & "</ITRForm:TotalTaxPaidOutsideIndia>"
        writeXML "<ITRForm:TotalTaxReliefOutsideIndia>" & Sheet27.Range("TR_TotalTaxReliefOutsideIndia").value & "</ITRForm:TotalTaxReliefOutsideIndia>"
        writeXML "<ITRForm:TaxReliefOutsideIndiaDTAA>" & TaxReliefOutsideIndiaDTAA_TR & "</ITRForm:TaxReliefOutsideIndiaDTAA>"
        writeXML "<ITRForm:TaxReliefOutsideIndiaNotDTAA>" & TaxReliefOutsideIndiaNotDTAA_TR & "</ITRForm:TaxReliefOutsideIndiaNotDTAA>"

        If (TaxPaidOutsideIndFlg_TR <> "") Then
            writeXML "<ITRForm:TaxPaidOutsideIndFlg>" & UCase(TaxPaidOutsideIndFlg_TR) & "</ITRForm:TaxPaidOutsideIndFlg>"
        End If

        If TaxPaidOutsideIndFlg_TR = "YES" Then
            If (AmtTaxRefunded_TR <> "") Then
                If (AmtTaxRefunded_TR <> 0) Then
                    writeXML "<ITRForm:AmtTaxRefunded>" & AmtTaxRefunded_TR & "</ITRForm:AmtTaxRefunded>"
                End If
            End If

            If (AssmtYrTaxRelief_TR) <> "" Then
                If (AssmtYrTaxRelief_TR) <> "0000-00" Then
                    writeXML "<ITRForm:AssmtYrTaxRelief>" & AssmtYrTaxRelief_TR & "</ITRForm:AssmtYrTaxRelief>"
                End If
            End If
        End If
        writeXML "    </ITRForm:ScheduleTR1>"
    End If
End If
End Function

Function ScheduleFA()
 Dim i, count As Long
 count = 0
 If Not IsEmpty(CountryCodeA1_FA) And UBound(CountryCodeA1_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeA2_FA) And UBound(CountryCodeA2_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeA3_FA) And UBound(CountryCodeA3_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeA4_FA) And UBound(CountryCodeA4_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeB_FA) And UBound(CountryCodeB_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeC_FA) And UBound(CountryCodeC_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeD_FA) And UBound(CountryCodeD_FA) > 0 Then count = count + 1
 If Not IsEmpty(InstNameE_FA) And UBound(InstNameE_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeF_FA) And UBound(CountryCodeF_FA) > 0 Then count = count + 1
 If Not IsEmpty(CountryCodeG_FA) And UBound(CountryCodeG_FA) > 0 Then count = count + 1

    If count > 0 Then
        writeXML "<ITRForm:ScheduleFA>"

'              FA_A1

        If Not IsEmpty(CountryCodeA1_FA) And UBound(CountryCodeA1_FA) > 0 Then
            For i = 1 To UBound(CountryNameA1_FA)
                writeXML "        <ITRForm:DetailsForiegnBank>"
                If CountryNameA1_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameA1_FA(i) & "</ITRForm:CountryName>"
                Else
                End If
'
                If CountryCodeA1_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeA1_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If


                If BankNameA1_FA(i) <> "" Then
                    writeXML "            <ITRForm:Bankname>" & BankNameA1_FA(i) & "</ITRForm:Bankname>"
                Else
                    writeXML "            <ITRForm:Bankname></ITRForm:Bankname>"
                End If

                If BankAddA1_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfBank>" & BankAddA1_FA(i) & "</ITRForm:AddressOfBank>"
                Else
                    writeXML "            <ITRForm:AddressOfBank></ITRForm:AddressOfBank>"
                End If


                If ZipCodeA1_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeA1_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If AccountNameA1_FA(i) <> "" Then
                    writeXML "             <ITRForm:ForeignAccountNumber>" & AccountNameA1_FA(i) & "</ITRForm:ForeignAccountNumber>"
                Else
                    writeXML "             <ITRForm:ForeignAccountNumber></ITRForm:ForeignAccountNumber>"
                End If

                If StatusA1_FA(i) <> "" Then
                writeXML "                   <ITRForm:OwnerStatus>" & StatusA1_FA(i) & "</ITRForm:OwnerStatus>"
                Else
                End If


                If AccountOpeningDateA1_FA(i) <> "" Then
                    writeXML "             <ITRForm:AccOpenDate>" & Dformat(AccountOpeningDateA1_FA(i), "") & "</ITRForm:AccOpenDate>"
                End If

                If PeakBalA1_FA(i) <> "" Then
                    writeXML "            <ITRForm:PeakBalanceDuringYear>" & PeakBalA1_FA(i) & "</ITRForm:PeakBalanceDuringYear>"
                Else
                    writeXML "            <ITRForm:PeakBalanceDuringYear>0</ITRForm:PeakBalanceDuringYear>"
                End If



                If CloseBalA1_FA(i) <> "" Then
                    writeXML "             <ITRForm:ClosingBalance>" & CloseBalA1_FA(i) & "</ITRForm:ClosingBalance>"
                Else
                    writeXML "             <ITRForm:ClosingBalance>0</ITRForm:ClosingBalance>"
                End If



                If GrossInterest_A1_FA(i) <> "" Then
                    writeXML "             <ITRForm:IntrstAccured>" & GrossInterest_A1_FA(i) & "</ITRForm:IntrstAccured>"
                Else
                    writeXML "             <ITRForm:IntrstAccured>0</ITRForm:IntrstAccured>"
                End If


                writeXML "        </ITRForm:DetailsForiegnBank>"
            Next
        End If


'        FA_A2
         If Not IsEmpty(CountryCodeA2_FA) And UBound(CountryCodeA2_FA) > 0 Then
            For i = 1 To UBound(CountryNameA2_FA)
                writeXML "        <ITRForm:DtlsForeignCustodialAcc>"
                If CountryNameA2_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameA2_FA(i) & "</ITRForm:CountryName>"
                Else
                End If
'
                If CountryCodeA2_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeA2_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If


                If BankNameA2_FA(i) <> "" Then
                    writeXML "            <ITRForm:FinancialInstName>" & BankNameA2_FA(i) & "</ITRForm:FinancialInstName>"
                Else
                    writeXML "            <ITRForm:FinancialInstName></ITRForm:FinancialInstName>"
                End If

                If BankAddA2_FA(i) <> "" Then
                    writeXML "            <ITRForm:FinancialInstAddress>" & BankAddA2_FA(i) & "</ITRForm:FinancialInstAddress>"
                Else
                    writeXML "            <ITRForm:FinancialInstAddress></ITRForm:FinancialInstAddress>"
                End If


                If ZipCodeA2_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeA2_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If AccountNameA2_FA(i) <> "" Then
                    writeXML "             <ITRForm:AccountNumber>" & AccountNameA2_FA(i) & "</ITRForm:AccountNumber>"
                Else
                    writeXML "             <ITRForm:AccountNumber></ITRForm:AccountNumber>"
                End If

                If StatusA2_FA(i) <> "" Then
                writeXML "                   <ITRForm:Status>" & StatusA2_FA(i) & "</ITRForm:Status>"
                Else
                End If


                If AccountOpeningDateA2_FA(i) <> "" Then
                    writeXML "             <ITRForm:AccOpenDate>" & Dformat(AccountOpeningDateA2_FA(i), "") & "</ITRForm:AccOpenDate>"
                End If

                If PeakBalA2_FA(i) <> "" Then
                    writeXML "            <ITRForm:PeakBalanceDuringPeriod>" & PeakBalA2_FA(i) & "</ITRForm:PeakBalanceDuringPeriod>"
                Else
                    writeXML "            <ITRForm:PeakBalanceDuringPeriod>0</ITRForm:PeakBalanceDuringPeriod>"
                End If



                If CloseBalA2_FA(i) <> "" Then
                    writeXML "             <ITRForm:ClosingBalance>" & CloseBalA2_FA(i) & "</ITRForm:ClosingBalance>"
                Else
                    writeXML "             <ITRForm:ClosingBalance>0</ITRForm:ClosingBalance>"
                End If


                If GrossInterest_A2_FA(i) <> "" Then
                    writeXML "             <ITRForm:GrossAmtPaidCredited>" & GrossInterest_A2_FA(i) & "</ITRForm:GrossAmtPaidCredited>"
                Else
                    writeXML "             <ITRForm:GrossAmtPaidCredited>0</ITRForm:GrossAmtPaidCredited>"
                End If


 If GrossInterestNature_A2_FA(i) <> "" Then

 If (GrossInterestNature_A2_FA(i)) = "Interest" Then
    GrossInterestNature_A2_FA(i) = "I"
 ElseIf (GrossInterestNature_A2_FA(i)) = "Dividend" Then
    GrossInterestNature_A2_FA(i) = "D"
 ElseIf (GrossInterestNature_A2_FA(i)) = "Proceeds from sale or redemption of financial assets" Then
    GrossInterestNature_A2_FA(i) = "S"
 ElseIf (GrossInterestNature_A2_FA(i)) = "Other income" Then
    GrossInterestNature_A2_FA(i) = "O"
 ElseIf (GrossInterestNature_A2_FA(i)) = "No Amount paid/credited" Then
    GrossInterestNature_A2_FA(i) = "N"
 End If



 writeXML "<ITRForm:NatureOfAmount>" & UCase(GrossInterestNature_A2_FA(i)) & "</ITRForm:NatureOfAmount>"
 End If


                writeXML "        </ITRForm:DtlsForeignCustodialAcc>"
            Next
        End If



 '        FA_A3

         If Not IsEmpty(CountryCodeA3_FA) And UBound(CountryCodeA3_FA) > 0 Then
            For i = 1 To UBound(CountryNameA3_FA)
                writeXML "        <ITRForm:DtlsForeignEquityDebtInterest>"
                If CountryNameA3_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameA3_FA(i) & "</ITRForm:CountryName>"
                Else
                End If
'
                If CountryCodeA3_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeA3_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If


                If BankNameA3_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfEntity>" & BankNameA3_FA(i) & "</ITRForm:NameOfEntity>"
                Else
                    writeXML "            <ITRForm:NameOfEntity></ITRForm:NameOfEntity>"
                End If

                If BankAddA3_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfEntity>" & BankAddA3_FA(i) & "</ITRForm:AddressOfEntity>"
                Else
                    writeXML "            <ITRForm:AddressOfEntity></ITRForm:AddressOfEntity>"
                End If


                If ZipCodeA3_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeA3_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If


                If NatureOfEntityA3_FA(i) <> "" Then
                    writeXML "             <ITRForm:NatureOfEntity>" & NatureOfEntityA3_FA(i) & "</ITRForm:NatureOfEntity>"
                Else
                    writeXML "             <ITRForm:NatureOfEntity></ITRForm:NatureOfEntity>"
                End If

                If AccountOpeningDateA3_FA(i) <> "" Then
                writeXML "                   <ITRForm:InterestAcquiringDate>" & Dformat(AccountOpeningDateA3_FA(i), "") & "</ITRForm:InterestAcquiringDate>"
                Else
                End If


                If Initialvalue_Investment(i) <> "" Then
                    writeXML "             <ITRForm:InitialValOfInvstmnt>" & Initialvalue_Investment(i) & "</ITRForm:InitialValOfInvstmnt>"
                End If

                If PeakBalA3_FA(i) <> "" Then
                    writeXML "            <ITRForm:PeakBalanceDuringPeriod>" & PeakBalA3_FA(i) & "</ITRForm:PeakBalanceDuringPeriod>"
                Else
                    writeXML "            <ITRForm:PeakBalanceDuringPeriod>0</ITRForm:PeakBalanceDuringPeriod>"
                End If



                If CloseBalA3_FA(i) <> "" Then
                    writeXML "             <ITRForm:ClosingBalance>" & CloseBalA3_FA(i) & "</ITRForm:ClosingBalance>"
                Else
                    writeXML "             <ITRForm:ClosingBalance></ITRForm:ClosingBalance>"
                End If


                If TotalGrossAmount_A3_FA(i) <> "" Then
                    writeXML "             <ITRForm:TotGrossAmtPaidCredited >" & TotalGrossAmount_A3_FA(i) & "</ITRForm:TotGrossAmtPaidCredited>"
                Else
                    writeXML "             <ITRForm:TotGrossAmtPaidCredited></ITRForm:TotGrossAmtPaidCredited>"
                End If



                If TotalGrossProceeds_A3_FA(i) <> "" Then
                    writeXML "             <ITRForm:TotGrossProceeds >" & TotalGrossProceeds_A3_FA(i) & "</ITRForm:TotGrossProceeds>"
                Else
                    writeXML "             <ITRForm:TotGrossProceeds></ITRForm:TotGrossProceeds>"
                End If

                writeXML "        </ITRForm:DtlsForeignEquityDebtInterest>"
            Next
        End If



        '        FA_A4

         If Not IsEmpty(CountryCodeA4_FA) And UBound(CountryCodeA4_FA) > 0 Then
            For i = 1 To UBound(CountryNameA4_FA)
                writeXML "        <ITRForm:DtlsForeignCashValueInsurance>"
                If CountryNameA4_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameA4_FA(i) & "</ITRForm:CountryName>"
                Else
                End If
'
                If CountryCodeA4_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeA4_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If


                If BankNameA4_FA(i) <> "" Then
                    writeXML "            <ITRForm:FinancialInstName>" & BankNameA4_FA(i) & "</ITRForm:FinancialInstName>"
                Else
                    writeXML "            <ITRForm:FinancialInstName></ITRForm:FinancialInstName>"
                End If

                If BankAddA4_FA(i) <> "" Then
                    writeXML "            <ITRForm:FinancialInstAddress>" & BankAddA4_FA(i) & "</ITRForm:FinancialInstAddress>"
                Else
                    writeXML "            <ITRForm:FinancialInstAddress></ITRForm:FinancialInstAddress>"
                End If


                If ZipCodeA4_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeA4_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If


                If AccountOpeningDateA4_FA(i) <> "" Then
                writeXML "                   <ITRForm:ContractDate>" & Dformat(AccountOpeningDateA4_FA(i), "") & "</ITRForm:ContractDate>"
                Else
                End If


                If TheCashValue_A4_FA(i) <> "" Then
                    writeXML "             <ITRForm:CashValOrSurrenderVal>" & TheCashValue_A4_FA(i) & "</ITRForm:CashValOrSurrenderVal>"
                Else
                    writeXML "             <ITRForm:CashValOrSurrenderVal>0</ITRForm:CashValOrSurrenderVal>"
                End If


                If TotalGrossAmount_A4_FA(i) <> "" Then
                    writeXML "            <ITRForm:TotGrossAmtPaidCredited>" & TotalGrossAmount_A4_FA(i) & "</ITRForm:TotGrossAmtPaidCredited>"
                Else
                    writeXML "            <ITRForm:TotGrossAmtPaidCredited>0</ITRForm:TotGrossAmtPaidCredited>"
                End If

                writeXML "        </ITRForm:DtlsForeignCashValueInsurance>"
            Next
        End If






        If Not IsEmpty(CountryCodeB_FA) And UBound(CountryCodeB_FA) > 0 Then
            For i = 1 To UBound(CountryNameB_FA)
                writeXML "        <ITRForm:DetailsFinancialInterest>"
                If CountryNameB_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameB_FA(i) & "</ITRForm:CountryName>"
                Else
                End If

                If CountryCodeB_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeB_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If

                If ZipCodeB_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeB_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If EntityNatureB_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfEntity>" & EntityNatureB_FA(i) & "</ITRForm:NatureOfEntity>"
                Else
                    writeXML "            <ITRForm:NatureOfEntity></ITRForm:NatureOfEntity>"
                End If

                If EntityNameB_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfEntity>" & EntityNameB_FA(i) & "</ITRForm:NameOfEntity>"
                Else
                    writeXML "            <ITRForm:NameOfEntity></ITRForm:NameOfEntity>"
                End If

                If EntityAddB_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfEntity>" & EntityAddB_FA(i) & "</ITRForm:AddressOfEntity>"
                Else
                    writeXML "            <ITRForm:AddressOfEntity></ITRForm:AddressOfEntity>"
                End If

                If NatureOfInterestB_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfInt>" & NatureOfInterestB_FA(i) & "</ITRForm:NatureOfInt>"
                End If

                 If DateHeldB_FA(i) <> "" Then
                    writeXML "            <ITRForm:DateHeld>" & Dformat(DateHeldB_FA(i), "") & "</ITRForm:DateHeld>"
                End If

                If TotalInvB_FA(i) <> "" Then
                    writeXML "            <ITRForm:TotalInvestment>" & TotalInvB_FA(i) & "</ITRForm:TotalInvestment>"
                Else
                    writeXML "            <ITRForm:TotalInvestment>0</ITRForm:TotalInvestment>"
                End If


                If IncomeIncurredInAccB_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncFromInt>" & IncomeIncurredInAccB_FA(i) & "</ITRForm:IncFromInt>"
                End If

                 If NatureOfIncomeB_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfInc>" & NatureOfIncomeB_FA(i) & "</ITRForm:NatureOfInc>"
                End If

                If AmountB_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxAmt>" & AmountB_FA(i) & "</ITRForm:IncTaxAmt>"
                End If

                If Schedule_B_FA(i) = "Other sources" Then
                   Schedule_B_FA(i) = "OS"
                ElseIf Schedule_B_FA(i) = "Business" Then
                   Schedule_B_FA(i) = "BU"
                ElseIf Schedule_B_FA(i) = "House Property" Then
                   Schedule_B_FA(i) = "HP"
                ElseIf Schedule_B_FA(i) = "Capital Gains" Then
                   Schedule_B_FA(i) = "CG"
                ElseIf Schedule_B_FA(i) = "Salary" Then
                   Schedule_B_FA(i) = "SA"
                ElseIf Schedule_B_FA(i) = "No Income during the year" Then
                   Schedule_B_FA(i) = "NI"
                ElseIf Schedule_B_FA(i) = "Exempt Income" Then
                   Schedule_B_FA(i) = "EI"

                End If

                If Schedule_B_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxSch>" & Schedule_B_FA(i) & "</ITRForm:IncTaxSch>"
                End If

                If ItemNumber_B_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxSchNo>" & ItemNumber_B_FA(i) & "</ITRForm:IncTaxSchNo>"
                End If

                writeXML "        </ITRForm:DetailsFinancialInterest>"
            Next
        End If

        If Not IsEmpty(CountryCodeC_FA) And UBound(CountryCodeC_FA) > 0 Then
            For i = 1 To UBound(CountryNameC_FA)
                writeXML "        <ITRForm:DetailsImmovableProperty>"
                If CountryNameC_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameC_FA(i) & "</ITRForm:CountryName>"
                Else
                End If

                If CountryCodeC_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeC_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If

                If ZipCodeC_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeC_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If PropertyAddC_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfProperty>" & PropertyAddC_FA(i) & "</ITRForm:AddressOfProperty>"
                Else
                    writeXML "            <ITRForm:AddressOfProperty></ITRForm:AddressOfProperty>"
                End If

                If OwnershipC_FA(i) <> "" Then
                    writeXML "            <ITRForm:Ownership>" & OwnershipC_FA(i) & "</ITRForm:Ownership>"
                End If

                If DateOfAcquistionC_FA(i) <> "" Then
                    writeXML "            <ITRForm:DateOfAcq>" & Dformat(DateOfAcquistionC_FA(i), "") & "</ITRForm:DateOfAcq>"
                End If



                If TotalInvC_FA(i) <> "" Then
                    writeXML "            <ITRForm:TotalInvestment>" & TotalInvC_FA(i) & "</ITRForm:TotalInvestment>"
                Else
                    writeXML "            <ITRForm:TotalInvestment>0</ITRForm:TotalInvestment>"
                End If

                If IncomeFromPropC_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncDrvProperty>" & IncomeFromPropC_FA(i) & "</ITRForm:IncDrvProperty>"
                End If

                If NatureOfIncomeC_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfInc>" & NatureOfIncomeC_FA(i) & "</ITRForm:NatureOfInc>"
                End If

                If AmountC_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxAmt>" & AmountC_FA(i) & "</ITRForm:IncTaxAmt>"
                End If

                If ScheduleOfferedC_FA(i) = "Other sources" Then
                   ScheduleOfferedC_FA(i) = "OS"
                ElseIf ScheduleOfferedC_FA(i) = "Business" Then
                   ScheduleOfferedC_FA(i) = "BU"
                ElseIf ScheduleOfferedC_FA(i) = "House Property" Then
                   ScheduleOfferedC_FA(i) = "HP"
                ElseIf ScheduleOfferedC_FA(i) = "Capital Gains" Then
                   ScheduleOfferedC_FA(i) = "CG"
                ElseIf ScheduleOfferedC_FA(i) = "Salary" Then
                   ScheduleOfferedC_FA(i) = "SA"

                ElseIf ScheduleOfferedC_FA(i) = "No Income during the year" Then
                   ScheduleOfferedC_FA(i) = "NI"
                ElseIf ScheduleOfferedC_FA(i) = "Exempt Income" Then
                   ScheduleOfferedC_FA(i) = "EI"
                End If

                If ScheduleOfferedC_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxSch>" & ScheduleOfferedC_FA(i) & "</ITRForm:IncTaxSch>"
                End If

                If ItemNumofScheduleC_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxSchNo>" & ItemNumofScheduleC_FA(i) & "</ITRForm:IncTaxSchNo>"
                End If

                writeXML "        </ITRForm:DetailsImmovableProperty>"
            Next
        End If

        If Not IsEmpty(CountryCodeD_FA) And UBound(CountryCodeD_FA) > 0 Then
            For i = 1 To UBound(CountryNameD_FA)
                writeXML "        <ITRForm:DetailsOthAssets>"
                If CountryNameD_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameD_FA(i) & "</ITRForm:CountryName>"
                Else

                End If
                If CountryCodeD_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeD_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If

                If ZipCodeD_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeD_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If AssetNatureD_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfAsset>" & AssetNatureD_FA(i) & "</ITRForm:NatureOfAsset>"
                Else
                    writeXML "            <ITRForm:NatureOfAsset></ITRForm:NatureOfAsset>"
                End If

                If OwnershipD_FA(i) <> "" Then
                    writeXML "            <ITRForm:Ownership>" & OwnershipD_FA(i) & "</ITRForm:Ownership>"
                End If

                  If DateofAcquistionD_FA(i) <> "" Then
                    writeXML "            <ITRForm:DateOfAcq>" & Dformat(DateofAcquistionD_FA(i), "") & "</ITRForm:DateOfAcq>"
                End If

                If TotalInvD_FA(i) <> "" Then
                    writeXML "            <ITRForm:TotalInvestment>" & TotalInvD_FA(i) & "</ITRForm:TotalInvestment>"
                Else
                    writeXML "            <ITRForm:TotalInvestment>0</ITRForm:TotalInvestment>"
                End If

                If IncomeFrmAsset_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncDrvAsset>" & IncomeFrmAsset_FA(i) & "</ITRForm:IncDrvAsset>"
                End If

                If NatureOfIncome_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfInc>" & NatureOfIncome_FA(i) & "</ITRForm:NatureOfInc>"
                End If

                If AmountD_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxAmt>" & AmountD_FA(i) & "</ITRForm:IncTaxAmt>"
                End If

                If ScheduleofferedD_FA(i) = "Other sources" Then
                   ScheduleofferedD_FA(i) = "OS"
                ElseIf ScheduleofferedD_FA(i) = "Business" Then
                   ScheduleofferedD_FA(i) = "BU"
                ElseIf ScheduleofferedD_FA(i) = "House Property" Then
                   ScheduleofferedD_FA(i) = "HP"
                ElseIf ScheduleofferedD_FA(i) = "Capital Gains" Then
                   ScheduleofferedD_FA(i) = "CG"
                ElseIf ScheduleofferedD_FA(i) = "Salary" Then
                   ScheduleofferedD_FA(i) = "SA"

                 ElseIf ScheduleofferedD_FA(i) = "No Income during the year" Then
                   ScheduleofferedD_FA(i) = "NI"
                ElseIf ScheduleofferedD_FA(i) = "Exempt Income" Then
                   ScheduleofferedD_FA(i) = "EI"
                End If




                If ScheduleofferedD_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxSch>" & ScheduleofferedD_FA(i) & "</ITRForm:IncTaxSch>"
                End If

                If ItemNumSchedule_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncTaxSchNo>" & ItemNumSchedule_FA(i) & "</ITRForm:IncTaxSchNo>"
                End If

                writeXML "        </ITRForm:DetailsOthAssets>"
            Next
        End If

        If Not IsEmpty(InstNameE_FA) And UBound(InstNameE_FA) > 0 Then
            For i = 1 To UBound(InstNameE_FA)
                writeXML "        <ITRForm:DetailsOfAccntsHvngSigningAuth>"

                If InstNameE_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfInstitution>" & InstNameE_FA(i) & "</ITRForm:NameOfInstitution>"
                Else
                    writeXML "            <ITRForm:NameOfInstitution></ITRForm:NameOfInstitution>"
                End If
                If InstAddE_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfInstitution>" & InstAddE_FA(i) & "</ITRForm:AddressOfInstitution>"
                Else
                    writeXML "            <ITRForm:AddressOfInstitution></ITRForm:AddressOfInstitution>"
                End If

                If CountryNameE_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameE_FA(i) & "</ITRForm:CountryName>"
                Else
                End If

                If CountryCodeE_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeE_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If

                If ZipCodeE_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeE_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If AccountNameE_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameMentionedInAccnt>" & AccountNameE_FA(i) & "</ITRForm:NameMentionedInAccnt>"
                Else
                    writeXML "            <ITRForm:NameMentionedInAccnt></ITRForm:NameMentionedInAccnt>"
                End If

                If InstitutionAccountNumberE_FA(i) <> "" Then
                    writeXML "            <ITRForm:InstitutionAccountNumber>" & InstitutionAccountNumberE_FA(i) & "</ITRForm:InstitutionAccountNumber>"
                Else
                    writeXML "            <ITRForm:InstitutionAccountNumber></ITRForm:InstitutionAccountNumber>"
                End If

                If PeakBalInvE_FA(i) <> "" Then
                    writeXML "            <ITRForm:PeakBalanceOrInvestment>" & PeakBalInvE_FA(i) & "</ITRForm:PeakBalanceOrInvestment>"
                Else
                    writeXML "            <ITRForm:PeakBalanceOrInvestment>0</ITRForm:PeakBalanceOrInvestment>"
                End If

                If AccountTaxableE_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncAccuredTaxFlag>" & AccountTaxableE_FA(i) & "</ITRForm:IncAccuredTaxFlag>"
                End If

                 If IncomeAccuredE_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncAccuredInAcc>" & IncomeAccuredE_FA(i) & "</ITRForm:IncAccuredInAcc>"
                End If

                If AmountE_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedAmt>" & AmountE_FA(i) & "</ITRForm:IncOfferedAmt>"
                End If

                If ScheduleOfferedE_FA(i) = "Other sources" Then
                   ScheduleOfferedE_FA(i) = "OS"
                ElseIf ScheduleOfferedE_FA(i) = "Business" Then
                   ScheduleOfferedE_FA(i) = "BU"
                ElseIf ScheduleOfferedE_FA(i) = "House Property" Then
                   ScheduleOfferedE_FA(i) = "HP"
                ElseIf ScheduleOfferedE_FA(i) = "Capital Gains" Then
                   ScheduleOfferedE_FA(i) = "CG"
                ElseIf ScheduleOfferedE_FA(i) = "Salary" Then
                   ScheduleOfferedE_FA(i) = "SA"

                ElseIf ScheduleOfferedE_FA(i) = "No Income during the year" Then
                   ScheduleOfferedE_FA(i) = "NI"
                ElseIf ScheduleOfferedE_FA(i) = "Exempt Income" Then
                   ScheduleOfferedE_FA(i) = "EI"
                End If



                If ScheduleOfferedE_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedSch>" & ScheduleOfferedE_FA(i) & "</ITRForm:IncOfferedSch>"
                End If

                If ItemNumScheduleE_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedSchNo>" & ItemNumScheduleE_FA(i) & "</ITRForm:IncOfferedSchNo>"
                End If


                writeXML "        </ITRForm:DetailsOfAccntsHvngSigningAuth>"
            Next
        End If

        If Not IsEmpty(CountryCodeF_FA) And UBound(CountryCodeF_FA) > 0 Then
            For i = 1 To UBound(CountryNameF_FA)
                writeXML "        <ITRForm:DetailsOfTrustOutIndiaTrustee>"
                If CountryNameF_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameF_FA(i) & "</ITRForm:CountryName>"
                Else
                End If

                If CountryCodeF_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeF_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If

                If ZipCodeF_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeF_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If NameOfTrustF_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfTrust>" & NameOfTrustF_FA(i) & "</ITRForm:NameOfTrust>"
                Else
                    writeXML "            <ITRForm:NameOfTrust></ITRForm:NameOfTrust>"
                End If

                If AddressOfTrustF_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfTrust>" & AddressOfTrustF_FA(i) & "</ITRForm:AddressOfTrust>"
                Else
                    writeXML "            <ITRForm:AddressOfTrust></ITRForm:AddressOfTrust>"
                End If

                If NameOfOtherTrusteesF_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfOtherTrustees>" & NameOfOtherTrusteesF_FA(i) & "</ITRForm:NameOfOtherTrustees>"
                Else
                    writeXML "            <ITRForm:NameOfOtherTrustees></ITRForm:NameOfOtherTrustees>"
                End If

                If AddressOfOtherTrusteesF_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfOtherTrustees>" & AddressOfOtherTrusteesF_FA(i) & "</ITRForm:AddressOfOtherTrustees>"
                Else
                    writeXML "            <ITRForm:AddressOfOtherTrustees></ITRForm:AddressOfOtherTrustees>"
                End If


                If NameOfSettlorF_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfSettlor>" & NameOfSettlorF_FA(i) & "</ITRForm:NameOfSettlor>"
                Else
                    writeXML "            <ITRForm:NameOfSettlor></ITRForm:NameOfSettlor>"
                End If

                If AddressOfSettlorF_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfSettlor>" & AddressOfSettlorF_FA(i) & "</ITRForm:AddressOfSettlor>"
                Else
                    writeXML "            <ITRForm:AddressOfSettlor></ITRForm:AddressOfSettlor>"
                End If

                If NameOfBeneficiariesF_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfBeneficiaries>" & NameOfBeneficiariesF_FA(i) & "</ITRForm:NameOfBeneficiaries>"
                Else
                    writeXML "            <ITRForm:NameOfBeneficiaries></ITRForm:NameOfBeneficiaries>"
                End If

                If AddressOfBeneficiariesF_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfBeneficiaries>" & AddressOfBeneficiariesF_FA(i) & "</ITRForm:AddressOfBeneficiaries>"
                Else
                    writeXML "            <ITRForm:AddressOfBeneficiaries></ITRForm:AddressOfBeneficiaries>"
                End If

                If DateSinceHeldF_FA(i) <> "" Then
                    writeXML "            <ITRForm:DateHeld>" & Dformat(DateSinceHeldF_FA(i), "") & "</ITRForm:DateHeld>"
                End If

                If WheteherTaxableIncomeF_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncDrvTaxFlag>" & WheteherTaxableIncomeF_FA(i) & "</ITRForm:IncDrvTaxFlag>"
                End If

                If IncomeDerivedF_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncDrvFromTrust>" & IncomeDerivedF_FA(i) & "</ITRForm:IncDrvFromTrust>"
                End If

                If AmountF_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedAmt>" & AmountF_FA(i) & "</ITRForm:IncOfferedAmt>"
                End If

                       If ScheduleOfferedF_FA(i) = "Other sources" Then
                   ScheduleOfferedF_FA(i) = "OS"
                ElseIf ScheduleOfferedF_FA(i) = "Business" Then
                   ScheduleOfferedF_FA(i) = "BU"
                ElseIf ScheduleOfferedF_FA(i) = "House Property" Then
                   ScheduleOfferedF_FA(i) = "HP"
                ElseIf ScheduleOfferedF_FA(i) = "Capital Gains" Then
                   ScheduleOfferedF_FA(i) = "CG"
                ElseIf ScheduleOfferedF_FA(i) = "Salary" Then
                   ScheduleOfferedF_FA(i) = "SA"

                 ElseIf ScheduleOfferedF_FA(i) = "No Income during the year" Then
                   ScheduleOfferedF_FA(i) = "NI"
                ElseIf ScheduleOfferedF_FA(i) = "Exempt Income" Then
                   ScheduleOfferedF_FA(i) = "EI"

                End If

                If ScheduleOfferedF_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedSch>" & ScheduleOfferedF_FA(i) & "</ITRForm:IncOfferedSch>"
                End If

                If ItemNumOfScheduleF_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedSchNo>" & ItemNumOfScheduleF_FA(i) & "</ITRForm:IncOfferedSchNo>"
                End If

                writeXML "        </ITRForm:DetailsOfTrustOutIndiaTrustee>"
            Next
        End If

        If Not IsEmpty(CountryCodeG_FA) And UBound(CountryCodeG_FA) > 0 Then
            For i = 1 To UBound(CountryNameG_FA)
                writeXML "        <ITRForm:DetailsOfOthSourcesIncOutsideIndia>"
                If CountryNameG_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryName>" & CountryNameG_FA(i) & "</ITRForm:CountryName>"
                Else
                End If

                If CountryCodeG_FA(i) <> "" Then
                    writeXML "            <ITRForm:CountryCode>" & CountryCodeG_FA(i) & "</ITRForm:CountryCode>"
                Else
                End If

                If ZipCodeG_FA(i) <> "" Then
                    writeXML "            <ITRForm:ZipCode>" & ZipCodeG_FA(i) & "</ITRForm:ZipCode>"
                Else
                End If

                If PersonNameG_FA(i) <> "" Then
                    writeXML "            <ITRForm:NameOfPerson>" & PersonNameG_FA(i) & "</ITRForm:NameOfPerson>"
                End If

                If PersonAddressG_FA(i) <> "" Then
                    writeXML "            <ITRForm:AddressOfPerson>" & PersonAddressG_FA(i) & "</ITRForm:AddressOfPerson>"
                End If

                If IncomeDerivedG_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncDerived>" & IncomeDerivedG_FA(i) & "</ITRForm:IncDerived>"
                End If

                If NatureIncomeG_FA(i) <> "" Then
                    writeXML "            <ITRForm:NatureOfInc>" & NatureIncomeG_FA(i) & "</ITRForm:NatureOfInc>"
                End If


                If WhetherTaxableIncomeG_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncDrvTaxFlag>" & WhetherTaxableIncomeG_FA(i) & "</ITRForm:IncDrvTaxFlag>"
                End If

                If AmountG_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedAmt>" & AmountG_FA(i) & "</ITRForm:IncOfferedAmt>"
                End If


                If ScheduleOfferedG_FA(i) = "Other sources" Then
                   ScheduleOfferedG_FA(i) = "OS"
                ElseIf ScheduleOfferedG_FA(i) = "Business" Then
                   ScheduleOfferedG_FA(i) = "BU"
                ElseIf ScheduleOfferedG_FA(i) = "House Property" Then
                   ScheduleOfferedG_FA(i) = "HP"
                ElseIf ScheduleOfferedG_FA(i) = "Capital Gains" Then
                   ScheduleOfferedG_FA(i) = "CG"
                ElseIf ScheduleOfferedG_FA(i) = "Salary" Then
                   ScheduleOfferedG_FA(i) = "SA"

                  ElseIf ScheduleOfferedG_FA(i) = "No Income during the year" Then
                   ScheduleOfferedG_FA(i) = "NI"
                ElseIf ScheduleOfferedG_FA(i) = "Exempt Income" Then
                   ScheduleOfferedG_FA(i) = "EI"

                End If

                If ScheduleOfferedG_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedSch>" & ScheduleOfferedG_FA(i) & "</ITRForm:IncOfferedSch>"
               End If

                If ItemNumSchG_FA(i) <> "" Then
                    writeXML "            <ITRForm:IncOfferedSchNo>" & ItemNumSchG_FA(i) & "</ITRForm:IncOfferedSchNo>"
                End If


                writeXML "        </ITRForm:DetailsOfOthSourcesIncOutsideIndia>"
            Next
        End If
        writeXML "    </ITRForm:ScheduleFA>"


    End If
End Function

Function Schedule5AXML()

If Sheet1.Range("sheet1.PortugeseCC5A").value = "Yes" Then
writeXML "<ITRForm:Schedule5A2014>"

If NameOfSpouse_S5A <> "" Then writeXML "  <ITRForm:NameOfSpouse>" & NameOfSpouse_S5A & "</ITRForm:NameOfSpouse>"
If PANOfSpouse_S5A <> "" Then writeXML "  <ITRForm:PANOfSpouse>" & PANOfSpouse_S5A & "</ITRForm:PANOfSpouse>"
If AadhaarOfSpouse_S5A <> "" Then writeXML "  <ITRForm:AadhaarOfSpouse>" & AadhaarOfSpouse_S5A & "</ITRForm:AadhaarOfSpouse>"

writeXML "  <ITRForm:HPHeadIncome>"

If Sheet29.Range("S5A_IncRecvdUndHeadHP").value <> "" Then
writeXML "      <ITRForm:IncRecvdUndHead>" & Sheet29.Range("S5A_IncRecvdUndHeadHP").value & "</ITRForm:IncRecvdUndHead>"
Else
writeXML "      <ITRForm:IncRecvdUndHead>0</ITRForm:IncRecvdUndHead>"
End If

If Sheet29.Range("S5A_HPHeadIncome").value <> "" Then
writeXML "      <ITRForm:AmtApprndOfSpouse>" & Sheet29.Range("S5A_HPHeadIncome").value & "</ITRForm:AmtApprndOfSpouse>"
Else
writeXML "      <ITRForm:AmtApprndOfSpouse>0</ITRForm:AmtApprndOfSpouse>"
End If

If Sheet29.Range("S5A_AmtTDSDeductedHP").value <> "" Then
writeXML "      <ITRForm:AmtTDSDeducted>" & Sheet29.Range("S5A_AmtTDSDeductedHP").value & "</ITRForm:AmtTDSDeducted>"
Else
writeXML "      <ITRForm:AmtTDSDeducted>0</ITRForm:AmtTDSDeducted>"
End If

If Sheet29.Range("S5A_TDSApprndOfSpouseHP").value <> "" Then
writeXML "      <ITRForm:TDSApprndOfSpouse>" & Sheet29.Range("S5A_TDSApprndOfSpouseHP").value & "</ITRForm:TDSApprndOfSpouse>"
Else
writeXML "      <ITRForm:TDSApprndOfSpouse>0</ITRForm:TDSApprndOfSpouse>"
End If

writeXML "  </ITRForm:HPHeadIncome>"

writeXML "  <ITRForm:BusHeadIncome>"

If Sheet29.Range("S5A_IncRecvdUndHeadBP").value <> "" Then
writeXML "      <ITRForm:IncRecvdUndHead>" & Sheet29.Range("S5A_IncRecvdUndHeadBP").value & "</ITRForm:IncRecvdUndHead>"
Else
writeXML "      <ITRForm:IncRecvdUndHead>0</ITRForm:IncRecvdUndHead>"
End If

If Sheet29.Range("S5A_BusHeadIncome").value <> "" Then
writeXML "      <ITRForm:AmtApprndOfSpouse>" & Sheet29.Range("S5A_BusHeadIncome").value & "</ITRForm:AmtApprndOfSpouse>"
Else
writeXML "      <ITRForm:AmtApprndOfSpouse>0</ITRForm:AmtApprndOfSpouse>"
End If

If Sheet29.Range("S5A_AmtTDSDeductedBP").value <> "" Then
writeXML "      <ITRForm:AmtTDSDeducted>" & Sheet29.Range("S5A_AmtTDSDeductedBP").value & "</ITRForm:AmtTDSDeducted>"
Else
writeXML "      <ITRForm:AmtTDSDeducted>0</ITRForm:AmtTDSDeducted>"
End If

If Sheet29.Range("S5A_TDSApprndOfSpouseBP").value <> "" Then
writeXML "      <ITRForm:TDSApprndOfSpouse>" & Sheet29.Range("S5A_TDSApprndOfSpouseBP").value & "</ITRForm:TDSApprndOfSpouse>"
Else
writeXML "      <ITRForm:TDSApprndOfSpouse>0</ITRForm:TDSApprndOfSpouse>"
End If

writeXML "  </ITRForm:BusHeadIncome>"

writeXML "  <ITRForm:CapGainHeadIncome>"


If Sheet29.Range("S5A_IncRecvdUndHeadCG").value <> "" Then
writeXML "      <ITRForm:IncRecvdUndHead>" & Sheet29.Range("S5A_IncRecvdUndHeadCG").value & "</ITRForm:IncRecvdUndHead>"
Else
writeXML "      <ITRForm:IncRecvdUndHead>0</ITRForm:IncRecvdUndHead>"
End If

If Sheet29.Range("S5A_CapGainHeadIncome").value <> "" Then
writeXML "      <ITRForm:AmtApprndOfSpouse>" & Sheet29.Range("S5A_CapGainHeadIncome").value & "</ITRForm:AmtApprndOfSpouse>"
Else
writeXML "      <ITRForm:AmtApprndOfSpouse>0</ITRForm:AmtApprndOfSpouse>"
End If

If Sheet29.Range("S5A_AmtTDSDeductedCG").value <> "" Then
writeXML "      <ITRForm:AmtTDSDeducted>" & Sheet29.Range("S5A_AmtTDSDeductedCG").value & "</ITRForm:AmtTDSDeducted>"
Else
writeXML "      <ITRForm:AmtTDSDeducted>0</ITRForm:AmtTDSDeducted>"
End If

If Sheet29.Range("S5A_TDSApprndOfSpouseCG").value <> "" Then
writeXML "      <ITRForm:TDSApprndOfSpouse>" & Sheet29.Range("S5A_TDSApprndOfSpouseCG").value & "</ITRForm:TDSApprndOfSpouse>"
Else
writeXML "      <ITRForm:TDSApprndOfSpouse>0</ITRForm:TDSApprndOfSpouse>"
End If

writeXML "  </ITRForm:CapGainHeadIncome>"

writeXML "  <ITRForm:OtherSourcesHeadIncome>"

If Sheet29.Range("S5A_IncRecvdUndHeadOS").value <> "" Then
writeXML "      <ITRForm:IncRecvdUndHead>" & Sheet29.Range("S5A_IncRecvdUndHeadOS").value & "</ITRForm:IncRecvdUndHead>"
Else
writeXML "      <ITRForm:IncRecvdUndHead>0</ITRForm:IncRecvdUndHead>"
End If

If Sheet29.Range("S5A_OtherSourcesHeadIncome").value <> "" Then
writeXML "      <ITRForm:AmtApprndOfSpouse>" & Sheet29.Range("S5A_OtherSourcesHeadIncome").value & "</ITRForm:AmtApprndOfSpouse>"
Else
writeXML "      <ITRForm:AmtApprndOfSpouse>0</ITRForm:AmtApprndOfSpouse>"
End If

If Sheet29.Range("S5A_AmtTDSDeductedOS").value <> "" Then
writeXML "      <ITRForm:AmtTDSDeducted>" & Sheet29.Range("S5A_AmtTDSDeductedOS").value & "</ITRForm:AmtTDSDeducted>"
Else
writeXML "      <ITRForm:AmtTDSDeducted>0</ITRForm:AmtTDSDeducted>"
End If

If Sheet29.Range("S5A_TDSApprndOfSpouseOS").value <> "" Then
writeXML "      <ITRForm:TDSApprndOfSpouse>" & Sheet29.Range("S5A_TDSApprndOfSpouseOS").value & "</ITRForm:TDSApprndOfSpouse>"
Else
writeXML "      <ITRForm:TDSApprndOfSpouse>0</ITRForm:TDSApprndOfSpouse>"
End If

writeXML "  </ITRForm:OtherSourcesHeadIncome>"

writeXML "  <ITRForm:TotalHeadIncome>"

If Sheet29.Range("S5A_IncRecvdUndHeadTotal").value <> "" Then
writeXML "      <ITRForm:IncRecvdUndHead>" & Sheet29.Range("S5A_IncRecvdUndHeadTotal").value & "</ITRForm:IncRecvdUndHead>"
Else
writeXML "      <ITRForm:IncRecvdUndHead>0</ITRForm:IncRecvdUndHead>"
End If

If Sheet29.Range("S5A_TotalHeadIncome").value <> "" Then
writeXML "      <ITRForm:AmtApprndOfSpouse>" & Sheet29.Range("S5A_TotalHeadIncome").value & "</ITRForm:AmtApprndOfSpouse>"
Else
writeXML "      <ITRForm:AmtApprndOfSpouse>0</ITRForm:AmtApprndOfSpouse>"
End If

If Sheet29.Range("S5A_AmtTDSDeductedTotal").value <> "" Then
writeXML "      <ITRForm:AmtTDSDeducted>" & Sheet29.Range("S5A_AmtTDSDeductedTotal").value & "</ITRForm:AmtTDSDeducted>"
Else
writeXML "      <ITRForm:AmtTDSDeducted>0</ITRForm:AmtTDSDeducted>"
End If

If Sheet29.Range("S5A_TDSApprndOfSpouseTotal").value <> "" Then
writeXML "      <ITRForm:TDSApprndOfSpouse>" & Sheet29.Range("S5A_TDSApprndOfSpouseTotal").value & "</ITRForm:TDSApprndOfSpouse>"
Else
writeXML "      <ITRForm:TDSApprndOfSpouse>0</ITRForm:TDSApprndOfSpouse>"
End If

writeXML "  </ITRForm:TotalHeadIncome>"

writeXML "</ITRForm:Schedule5A2014>"

End If
End Function

Function ScheduleALXML()
   Dim i As Long
    If (DepositsInBank_AL >= 0 And _
          SharesAndSecurities_AL >= 0 And _
          InsurancePolicies_AL >= 0 And _
          LoansAndAdvancesGiven_AL >= 0 And _
          CashInHand_AL >= 0 And _
          JewelleryBullionEtc_AL >= 0 And _
          ArchCollDrawPaintSulpArt_AL >= 0 And _
          VehiclYachtsBoatsAircrafts_AL >= 0) And _
          Not isdropdownblank(IsImmovableAsset) And _
          Not isdropdownblank(IsInterestHeld) And _
          LiabilityInRelatAssets_AL >= 0 Then

            writeXML "<ITRForm:ScheduleAL>"
'            If Not isdropdownblank(IsImmovableAsset) Then
'                writeXML "<ITRForm:ImmovableFlag>" & UCase(Mid(IsImmovableAsset, 1, 1)) & "</ITRForm:ImmovableFlag>"
'            End If

            If UCase(IsImmovableAsset) = "YES" Then
                If (Not IsEmpty(ImmovableAssetDesc)) And UBound(ImmovableAssetDesc) > 0 Then
                    'ImmovableDetails
                    For i = 1 To UBound(ImmovableAssetDesc)
                        writeXML "<ITRForm:ImmovableDetails>"
                        If ImmovableAssetDesc(i) <> "" Then
                            writeXML "<ITRForm:Description>" & UCase(ImmovableAssetDesc(i)) & "</ITRForm:Description>"
                        End If
                        writeXML "<ITRForm:AddressAL>"
                            If ImmovableAssetAddress1(i) <> "" Then
                                writeXML "<ITRForm:ResidenceNo>" & UCase(ImmovableAssetAddress1(i)) & "</ITRForm:ResidenceNo>"
                            End If
                            If ImmovableAssetAddress2(i) <> "" Then
                                writeXML "<ITRForm:ResidenceName>" & UCase(ImmovableAssetAddress2(i)) & "</ITRForm:ResidenceName>"
                            End If

                            If ImmovableAssetAddress3(i) <> "" Then
                                writeXML "<ITRForm:RoadOrStreet>" & UCase(ImmovableAssetAddress3(i)) & "</ITRForm:RoadOrStreet>"
                            End If

                            If ImmovableAssetAddress4(i) <> "" Then
                                writeXML "<ITRForm:LocalityOrArea>" & UCase(ImmovableAssetAddress4(i)) & "</ITRForm:LocalityOrArea>"
                            End If

                            If ImmovableAssetAddress5(i) <> "" Then
                                writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(ImmovableAssetAddress5(i)) & "</ITRForm:CityOrTownOrDistrict>"
                            End If

                            If ImmovableAssetAddress6(i) <> "" Then
                                writeXML "<ITRForm:StateCode>" & UCase(Mid(Trim(ImmovableAssetAddress6(i)), 1, InStr(1, Trim(ImmovableAssetAddress6(i)), "-") - 1)) & "</ITRForm:StateCode>"
                            End If

                            If ImmovableAssetAddress7(i) <> "" Then
                                writeXML "<ITRForm:CountryCode>" & UCase(UCase(Mid(Trim(ImmovableAssetAddress7(i)), 1, InStr(1, Trim(ImmovableAssetAddress7(i)), "-") - 1))) & "</ITRForm:CountryCode>"
                            End If

                            If ImmovableAssetAddress8(i) <> "" Then
                                writeXML "<ITRForm:PinCode>" & UCase(ImmovableAssetAddress8(i)) & "</ITRForm:PinCode>"
                            End If

                            If ImmovableAssetAddress9(i) <> "" Then
                                writeXML "<ITRForm:ZipCode>" & UCase(ImmovableAssetAddress9(i)) & "</ITRForm:ZipCode>"
                            End If
                        writeXML "</ITRForm:AddressAL>"
                        If ImmovableAssetAmount(i) <> "" Then
                            writeXML "<ITRForm:Amount>" & UVCase(ImmovableAssetAmount(i)) & "</ITRForm:Amount>"
                        End If
                        writeXML "</ITRForm:ImmovableDetails>"
                    Next
                End If
            End If

            'MovableAsset
              If (DepositsInBank_AL >= 0 And _
                    SharesAndSecurities_AL >= 0 And _
                    InsurancePolicies_AL >= 0 And _
                    LoansAndAdvancesGiven_AL >= 0 And _
                    CashInHand_AL >= 0 And _
                    JewelleryBullionEtc_AL >= 0 And _
                    ArchCollDrawPaintSulpArt_AL >= 0 And _
                    VehiclYachtsBoatsAircrafts_AL >= 0) Then

                writeXML "<ITRForm:MovableAsset>"
                    If DepositsInBank_AL <> "" Then
                        writeXML "<ITRForm:DepositsInBank>" & UVCase(DepositsInBank_AL) & "</ITRForm:DepositsInBank>"
                    Else
                        'writeXML "<ITRForm:DepositsInBank>0</ITRForm:DepositsInBank>"
                    End If

                    If SharesAndSecurities_AL <> "" Then
                        writeXML "<ITRForm:SharesAndSecurities>" & UVCase(SharesAndSecurities_AL) & "</ITRForm:SharesAndSecurities>"
                    Else
                        'writeXML "<ITRForm:SharesAndSecurities>0</ITRForm:SharesAndSecurities>"
                    End If

                    If InsurancePolicies_AL <> "" Then
                        writeXML "<ITRForm:InsurancePolicies>" & UVCase(InsurancePolicies_AL) & "</ITRForm:InsurancePolicies>"
                    Else
                        'writeXML "<ITRForm:InsurancePolicies>0</ITRForm:InsurancePolicies>"
                    End If

                    If LoansAndAdvancesGiven_AL <> "" Then
                        writeXML "<ITRForm:LoansAndAdvancesGiven>" & UVCase(LoansAndAdvancesGiven_AL) & "</ITRForm:LoansAndAdvancesGiven>"
                    Else
                        'writeXML "<ITRForm:LoansAndAdvancesGiven>0</ITRForm:LoansAndAdvancesGiven>"
                    End If

                    If CashInHand_AL <> "" Then
                        writeXML "<ITRForm:CashInHand>" & UVCase(CashInHand_AL) & "</ITRForm:CashInHand>"
                    Else
                        'writeXML "<ITRForm:CashInHand>0</ITRForm:CashInHand>"
                    End If

                    If JewelleryBullionEtc_AL <> "" Then
                        writeXML "<ITRForm:JewelleryBullionEtc>" & UVCase(JewelleryBullionEtc_AL) & "</ITRForm:JewelleryBullionEtc>"
                    Else
                        'writeXML "<ITRForm:JewelleryBullionEtc>0</ITRForm:JewelleryBullionEtc>"
                    End If

                    If ArchCollDrawPaintSulpArt_AL <> "" Then
                        writeXML "<ITRForm:ArchCollDrawPaintSulpArt>" & UVCase(ArchCollDrawPaintSulpArt_AL) & "</ITRForm:ArchCollDrawPaintSulpArt>"
                    Else
                        'writeXML "<ITRForm:ArchCollDrawPaintSulpArt>0</ITRForm:ArchCollDrawPaintSulpArt>"
                    End If

                    If VehiclYachtsBoatsAircrafts_AL <> "" Then
                        writeXML "<ITRForm:VehiclYachtsBoatsAircrafts>" & UVCase(VehiclYachtsBoatsAircrafts_AL) & "</ITRForm:VehiclYachtsBoatsAircrafts>"
                    Else
                        'writeXML "<ITRForm:VehiclYachtsBoatsAircrafts>0</ITRForm:VehiclYachtsBoatsAircrafts>"
                    End If
                writeXML "</ITRForm:MovableAsset>"
             End If

            'InterestHeldInaAsset
            If Not isdropdownblank(IsInterestHeld) Then
                writeXML "<ITRForm:InterstAOPFlag>" & UCase(Mid(IsInterestHeld, 1, 1)) & "</ITRForm:InterstAOPFlag>"
            End If
            If UCase(IsInterestHeld) = "YES" Then
                If (Not IsEmpty(InterestHeldNameOfFirm)) And UBound(InterestHeldNameOfFirm) > 0 Then
                    For i = 1 To UBound(InterestHeldNameOfFirm)
                        writeXML "<ITRForm:InterestHeldInaAsset>"
                        If InterestHeldNameOfFirm(i) <> "" Then
                            writeXML "<ITRForm:NameOfFirm>" & UCase(InterestHeldNameOfFirm(i)) & "</ITRForm:NameOfFirm>"
                        End If
                        writeXML "<ITRForm:AddressAL>"
                            If InterestHeldAddressOfFirm1(i) <> "" Then
                                writeXML "<ITRForm:ResidenceNo>" & UCase(InterestHeldAddressOfFirm1(i)) & "</ITRForm:ResidenceNo>"
                            End If
                            If InterestHeldAddressOfFirm2(i) <> "" Then
                                writeXML "<ITRForm:ResidenceName>" & UCase(InterestHeldAddressOfFirm2(i)) & "</ITRForm:ResidenceName>"
                            End If

                            If InterestHeldAddressOfFirm3(i) <> "" Then
                                writeXML "<ITRForm:RoadOrStreet>" & UCase(InterestHeldAddressOfFirm3(i)) & "</ITRForm:RoadOrStreet>"
                            End If

                            If InterestHeldAddressOfFirm4(i) <> "" Then
                                writeXML "<ITRForm:LocalityOrArea>" & UCase(InterestHeldAddressOfFirm4(i)) & "</ITRForm:LocalityOrArea>"
                            End If

                            If InterestHeldAddressOfFirm5(i) <> "" Then
                                writeXML "<ITRForm:CityOrTownOrDistrict>" & UCase(InterestHeldAddressOfFirm5(i)) & "</ITRForm:CityOrTownOrDistrict>"
                            End If

                            If InterestHeldAddressOfFirm6(i) <> "" Then
                                writeXML "<ITRForm:StateCode>" & UCase(Mid(Trim(InterestHeldAddressOfFirm6(i)), 1, InStr(1, Trim(InterestHeldAddressOfFirm6(i)), "-") - 1)) & "</ITRForm:StateCode>"
                            End If

                            If InterestHeldAddressOfFirm7(i) <> "" Then
                                writeXML "<ITRForm:CountryCode>" & UCase(Mid(Trim(InterestHeldAddressOfFirm7(i)), 1, InStr(1, Trim(InterestHeldAddressOfFirm7(i)), "-") - 1)) & "</ITRForm:CountryCode>"
                            End If

                            If InterestHeldAddressOfFirm8(i) <> "" Then
                                writeXML "<ITRForm:PinCode>" & UCase(InterestHeldAddressOfFirm8(i)) & "</ITRForm:PinCode>"
                            End If

                            If InterestHeldAddressOfFirm9(i) <> "" Then
                                writeXML "<ITRForm:ZipCode>" & UCase(InterestHeldAddressOfFirm9(i)) & "</ITRForm:ZipCode>"
                            End If

                        writeXML "</ITRForm:AddressAL>"
                        If InterestHeldPANOfFirm(i) <> "" Then
                            writeXML "<ITRForm:PanOfFirm>" & UCase(InterestHeldPANOfFirm(i)) & "</ITRForm:PanOfFirm>"
                        End If
                        If InterestHeldInvestmentOfFirm(i) <> "" Then
                            writeXML "<ITRForm:AssesseInvestment>" & UVCase(InterestHeldInvestmentOfFirm(i)) & "</ITRForm:AssesseInvestment>"
                        End If
                        writeXML "</ITRForm:InterestHeldInaAsset>"
                    Next
                End If
            End If

            If LiabilityInRelatAssets_AL <> "" Then
                writeXML "<ITRForm:LiabilityInRelatAssets>" & UVCase(LiabilityInRelatAssets_AL) & "</ITRForm:LiabilityInRelatAssets>"
            Else
                'writeXML "<ITRForm:LiabilityInRelatAssets>0</ITRForm:LiabilityInRelatAssets>"
            End If
        writeXML "</ITRForm:ScheduleAL>"
    End If
End Function

Function ScheduleAMTC()
'If TaxSection115JC_AMTC <> "" Then
    writeXML "  <ITRForm:ScheduleAMTC>"
    If TaxSection115JC_AMTC <> "" Then
      writeXML "        <ITRForm:TaxSection115JC>" & TaxSection115JC_AMTC & "</ITRForm:TaxSection115JC>"
    Else
      writeXML "        <ITRForm:TaxSection115JC>0</ITRForm:TaxSection115JC>"
    End If

    If TaxOthProvisions_AMTC <> "" Then
      writeXML "        <ITRForm:TaxOthProvisions>" & TaxOthProvisions_AMTC & "</ITRForm:TaxOthProvisions>"
    Else
      writeXML "        <ITRForm:TaxOthProvisions>0</ITRForm:TaxOthProvisions>"
    End If

    If AmtTaxCreditAvailable_AMTC <> "" Then
      writeXML "        <ITRForm:AmtTaxCreditAvailable>" & AmtTaxCreditAvailable_AMTC & "</ITRForm:AmtTaxCreditAvailable>"
    Else
      writeXML "        <ITRForm:AmtTaxCreditAvailable>0</ITRForm:AmtTaxCreditAvailable>"
    End If


'2013-14
   writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & AssYr_AMTC0 & "</ITRForm:AssYr>"

        If Gross4i0 <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i0 & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If SetOff4i0 <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & SetOff4i0 & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Balance4i0 <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Balance4i0 & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If AMTCredit4i0 <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & AMTCredit4i0 & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If BalAmtCreditCarryFwd10 <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & BalAmtCreditCarryFwd10 & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"

'2014-15
    writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & AssYr_AMTC2 & "</ITRForm:AssYr>"

        If Gross4i2 <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i2 & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If SetOff4i2 <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & SetOff4i2 & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Balance4i2 <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Balance4i2 & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If AMTCredit4i2 <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & AMTCredit4i2 & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If BalAmtCreditCarryFwd12 <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & BalAmtCreditCarryFwd12 & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"

'2015-16

    writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & AssYr_AMTC3 & "</ITRForm:AssYr>"

        If Gross4i3 <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i3 & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If SetOff4i3 <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & SetOff4i3 & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Balance4i3 <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Balance4i3 & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If AMTCredit4i3 <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & AMTCredit4i3 & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If BalAmtCreditCarryFwd13 <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & BalAmtCreditCarryFwd13 & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"


'2016-17
    writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & AssYr_AMTC & "</ITRForm:AssYr>"

        If Gross4i <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If SetOff4i <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & SetOff4i & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Balance4i <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Balance4i & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If AMTCredit4i <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & AMTCredit4i & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If BalAmtCreditCarryFwd1 <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & BalAmtCreditCarryFwd1 & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"


 ' 2017-18


   writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & AssYr_AMTC5 & "</ITRForm:AssYr>"

        If Gross4i5 <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i5 & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If SetOff4i5 <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & SetOff4i5 & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Balance4i5 <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Balance4i5 & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If AMTCredit4i5 <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & AMTCredit4i5 & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If BalAmtCreditCarryFwd15 <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & BalAmtCreditCarryFwd15 & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"



'2018-19

    writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & AssYr_AMTC4 & "</ITRForm:AssYr>"

        If Gross4i4 <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i4 & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If SetOff4i4 <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & SetOff4i4 & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Balance4i4 <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Balance4i4 & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If AMTCredit4i4 <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & AMTCredit4i4 & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If BalAmtCreditCarryFwd14 <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & BalAmtCreditCarryFwd14 & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"


'2019-20
        writeXML "     <ITRForm:ScheduleAMTCDtls>"

        writeXML "        <ITRForm:AssYr>" & Sheet23.Range("AMTC.AssYr20_21").value & "</ITRForm:AssYr>"

        If Gross4i6 <> "" Then
          writeXML "        <ITRForm:AmtCreditFwd>" & Gross4i6 & "</ITRForm:AmtCreditFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditFwd>0</ITRForm:AmtCreditFwd>"
        End If

        If Sheet23.Range("AMTC.AmtCreditSetOfEy6").value <> "" Then
          writeXML "        <ITRForm:AmtCreditSetOfEy>" & Sheet23.Range("AMTC.AmtCreditSetOfEy6").value & "</ITRForm:AmtCreditSetOfEy>"
        Else
          writeXML "        <ITRForm:AmtCreditSetOfEy>0</ITRForm:AmtCreditSetOfEy>"
        End If

        If Sheet23.Range("AMTC.AmtCreditBalBroughtFwd6").value <> "" Then
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>" & Sheet23.Range("AMTC.AmtCreditBalBroughtFwd6").value & "</ITRForm:AmtCreditBalBroughtFwd>"
        Else
          writeXML "        <ITRForm:AmtCreditBalBroughtFwd>0</ITRForm:AmtCreditBalBroughtFwd>"
        End If

        If Sheet23.Range("AMTC.AmtCreditUtilized6").value <> "" Then
          writeXML "        <ITRForm:AmtCreditUtilized>" & Sheet23.Range("AMTC.AmtCreditUtilized6").value & "</ITRForm:AmtCreditUtilized>"
        Else
          writeXML "        <ITRForm:AmtCreditUtilized>0</ITRForm:AmtCreditUtilized>"
        End If

        If Sheet23.Range("AMTC.BalAmtCreditCarryFwd6").value <> "" Then
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>" & Sheet23.Range("AMTC.BalAmtCreditCarryFwd6").value & "</ITRForm:BalAmtCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:BalAmtCreditCarryFwd>0</ITRForm:BalAmtCreditCarryFwd>"
        End If

    writeXML "     </ITRForm:ScheduleAMTCDtls>"


        writeXML "        <ITRForm:CurrAssYr>2023-24</ITRForm:CurrAssYr>"

        If Gross4ii <> "" Then
          writeXML "        <ITRForm:CurrYrAmtCreditFwd>" & Gross4ii & "</ITRForm:CurrYrAmtCreditFwd>"
        Else
          writeXML "        <ITRForm:CurrYrAmtCreditFwd>0</ITRForm:CurrYrAmtCreditFwd>"
        End If

        If BalAmtCreditCarryFwd <> "" Then
          writeXML "        <ITRForm:CurrYrCreditCarryFwd>" & BalAmtCreditCarryFwd & "</ITRForm:CurrYrCreditCarryFwd>"
        Else
          writeXML "        <ITRForm:CurrYrCreditCarryFwd>0</ITRForm:CurrYrCreditCarryFwd>"
        End If


    If Gross <> "" Then
      writeXML "<ITRForm:TotAMTGross>" & Gross & "</ITRForm:TotAMTGross>"
    Else
      writeXML "<ITRForm:TotAMTGross>0</ITRForm:TotAMTGross>"
    End If

    If SetOff <> "" Then
      writeXML "<ITRForm:TotSetOffEys>" & SetOff & "</ITRForm:TotSetOffEys>"
    Else
      writeXML "<ITRForm:TotSetOffEys>0</ITRForm:TotSetOffEys>"
    End If

    If Balance <> "" Then
      writeXML "<ITRForm:TotBalBF>" & Balance & "</ITRForm:TotBalBF>"
    Else
      writeXML "<ITRForm:TotBalBF>0</ITRForm:TotBalBF>"
    End If

    If AMTCredit <> "" Then
      writeXML "<ITRForm:TotAmtCreditUtilisedCY>" & AMTCredit & "</ITRForm:TotAmtCreditUtilisedCY>"
    Else
      writeXML "<ITRForm:TotAmtCreditUtilisedCY>0</ITRForm:TotAmtCreditUtilisedCY>"
    End If

    If BalAmtCreditCarryFwd_Total <> "" Then
      writeXML "<ITRForm:TotBalAMTCreditCF>" & BalAmtCreditCarryFwd_Total & "</ITRForm:TotBalAMTCreditCF>"
    Else
      writeXML "<ITRForm:TotBalAMTCreditCF>0</ITRForm:TotBalAMTCreditCF>"
    End If

    If TaxSection115JD_AMTC <> "" Then
      writeXML "<ITRForm:TaxSection115JD>" & TaxSection115JD_AMTC & "</ITRForm:TaxSection115JD>"
    Else
      writeXML "<ITRForm:TaxSection115JD>0</ITRForm:TaxSection115JD>"
    End If

    If AmtLiabilityAvailable_AMTC <> "" Then
      writeXML "<ITRForm:AmtLiabilityAvailable>" & AmtLiabilityAvailable_AMTC & "</ITRForm:AmtLiabilityAvailable>"
    Else
      writeXML "<ITRForm:AmtLiabilityAvailable>0</ITRForm:AmtLiabilityAvailable>"
    End If

    writeXML "</ITRForm:ScheduleAMTC>"
    'End If



End Function


Sub writeXML(Node As String)
    Node = Replace(Node, "&", "-")
    Print #1, Node
    xmlLinesCount = xmlLinesCount + 1
End Sub

Function XMLPART_CoName_HP(ByVal myindex As Long) As String
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range

    XMLPART_CoName_HP = ""

    Set rangecells1 = Sheet8.Range("HP.Co.Name" & myindex).Cells
    Set rangecells2 = Sheet8.Range("HP.Co.PAN" & myindex).Cells
    Set rangecells3 = Sheet8.Range("HP.Co.Share" & myindex).Cells
    Set rangecells4 = Sheet8.Range("HP.Co.Aadhaar" & myindex).Cells

    setTblinfo_hpcoindex (myindex)

    ReDim CoName_HP(myindex, end_hpco)
    ReDim CoPAN_HP(myindex, end_hpco)
    ReDim CoAadhar_HP(myindex, end_hpco)
    ReDim CoShare_HP(myindex, end_hpco)

    For j = 1 To end_hpco
        CoName_HP(myindex, j) = rangecells1.item(j).value
        CoPAN_HP(myindex, j) = rangecells2.item(j).value
        CoAadhar_HP(myindex, j) = rangecells4.item(j).value
        CoShare_HP(myindex, j) = rangecells3.item(j).text
        CoShare_HP(myindex, j) = Round(CoShare_HP(myindex, j), 2)

        writeXML " <ITRForm:CoOwners>"
        writeXML "<ITRForm:CoOwnersSNo>" & j & "</ITRForm:CoOwnersSNo>"
        writeXML "<ITRForm:NameCoOwner>" & CoName_HP(myindex, j) & "</ITRForm:NameCoOwner>"
        If CoPAN_HP(myindex, j) <> "" Then
            writeXML " <ITRForm:PAN_CoOwner>" & CoPAN_HP(myindex, j) & "</ITRForm:PAN_CoOwner>"
        End If
        If CoAadhar_HP(myindex, j) <> "" Then
            writeXML " <ITRForm:Aadhaar_CoOwner>" & CoAadhar_HP(myindex, j) & "</ITRForm:Aadhaar_CoOwner>"
        End If
        writeXML " <ITRForm:PercentShareProperty>" & CoShare_HP(myindex, j) & "</ITRForm:PercentShareProperty>"
        writeXML " </ITRForm:CoOwners>"
    Next
End Function

Function XMLPART_Tenant(ByVal myindex As Long) As String
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range


    XMLPART_Tenant = ""

    Set rangecells1 = Sheet8.Range("HP.NameofTenant" & myindex).Cells
    Set rangecells2 = Sheet8.Range("HP.PANofTenant" & myindex).Cells
    Set rangecells3 = Sheet8.Range("HP.TANofTenant" & myindex).Cells
    Set rangecells4 = Sheet8.Range("HP.AadharofTenant" & myindex).Cells

    setTblinfo_hprptfrm
    'setTblinfo_hpcoindex (myindex)

    ReDim NameofTenant_HP(rangecells1.Cells.count)
    ReDim PANofTenant_HP(rangecells2.Cells.count)
    ReDim AadharofTenant_HP(rangecells4.Cells.count)
    ReDim TANofTenant_HP(rangecells3.Cells.count)

    For j = 1 To UBound(NameofTenant_HP)
        NameofTenant_HP(j) = rangecells1.item(j).value
        PANofTenant_HP(j) = rangecells2.item(j).value
        AadharofTenant_HP(j) = rangecells4.item(j).value
        TANofTenant_HP(j) = rangecells3.item(j).value
        If NameofTenant_HP(j) <> "" Then
            writeXML "<ITRForm:TenantDetails>"
            writeXML "<ITRForm:TenantSNo>" & j & "</ITRForm:TenantSNo>"
            writeXML "<ITRForm:NameofTenant>" & NameofTenant_HP(j) & "</ITRForm:NameofTenant>"
            If PANofTenant_HP(j) <> "" Then
            writeXML "<ITRForm:PANofTenant>" & PANofTenant_HP(j) & "</ITRForm:PANofTenant>"
            End If
            If AadharofTenant_HP(j) <> "" Then
            writeXML "<ITRForm:AadhaarofTenant>" & AadharofTenant_HP(j) & "</ITRForm:AadhaarofTenant>"
            End If
            If TANofTenant_HP(j) <> "" Then
            writeXML "<ITRForm:PANTANofTenant>" & TANofTenant_HP(j) & "</ITRForm:PANTANofTenant>"
            End If
             writeXML "</ITRForm:TenantDetails>"
        End If
    Next
End Function


Function XMLBA()
    Dim i As Long

    writeXML " <ITRForm:ScheduleBA>"
    If Range("NumBankAccountsHeld").value <> 0 Then
        writeXML "<ITRForm:BankAccounts>" & Range("NumBankAccountsHeld").value & "</ITRForm:BankAccounts>"
    Else
        writeXML "<ITRForm:BankAccounts>" & 1 & "</ITRForm:BankAccounts>"
    End If

    If Not IsEmpty(BankIFSC_BA) And UBound(BankIFSC_BA) > 0 Then
        For i = 1 To UBound(BankIFSC_BA)
            writeXML "           <ITRForm:BankDetail>"

            If BankIFSC_BA(i) <> "" Then
                writeXML "               <ITRForm:IFSCCode>" & UCase(BankIFSC_BA(i)) & "</ITRForm:IFSCCode>"
            Else
                writeXML "               <ITRForm:IFSCCode/>"
            End If

            If BankName_BA(i) <> "" Or BankName_BA(i) <> "(Select)" Then


                Else
                    writeXML "               <ITRForm:BankName>" & BankNameCode(i) & "</ITRForm:BankName>"
                    writeXML "               <ITRForm:BankNameOthers>" & Mid(BankName_BA(i), 8, Len(BankName_BA(i)) - 1) & "</ITRForm:BankNameOthers>"
                End If




            If BankAccntnum_BA(i) <> "" Then
                writeXML "               <ITRForm:BankAccountNo>" & UCase(BankAccntnum_BA(i)) & "</ITRForm:BankAccountNo>"
            Else
                writeXML "               <ITRForm:BankAccountNo>" & 0 & "</ITRForm:BankAccountNo>"
            End If

            If BankAccountStatus_BA(i) <> "" Or BankAccountStatus_BA(i) <> "(Select) " Then
                writeXML "               <ITRForm:BankAccountStatus>" & UCase(BankAccountStatus_BA(i)) & "</ITRForm:BankAccountStatus>"
            Else
                writeXML "               <ITRForm:BankAccountStatus>" & "" & "</ITRForm:BankAccountStatus>"
            End If



            writeXML "           </ITRForm:BankDetail>"
        Next
    End If
    writeXML " </ITRForm:ScheduleBA>"
End Function

Function ScheduleGST()
 Dim i As Long


    subProcCaption = "Schedule GST"
        writeXML "<ITRForm:ScheduleGST>"

        noOfProcessSub = UBound(GSTIN_BP)
        If Not IsEmpty(GSTIN_BP) And UBound(GSTIN_BP) > 0 Then
            For i = 1 To UBound(GSTIN_BP)

           writeXML "<ITRForm:TurnoverGrsRcptForGSTIN>"

                If GSTIN_BP(i) <> "" Then
                    writeXML "<ITRForm:GSTINNo>" & UCase(GSTIN_BP(i)) & "</ITRForm:GSTINNo>"
                Else
                    writeXML "<ITRForm:GSTINNo/>"
                End If

                If GSTINAMOUNT_BP(i) <> "" Then
                    writeXML "<ITRForm:AmtTurnGrossRcptGSTIN>" & UCase(GSTINAMOUNT_BP(i)) & "</ITRForm:AmtTurnGrossRcptGSTIN>"
                Else
                    writeXML "<ITRForm:AmtTurnGrossRcptGSTIN/>"
                End If


         writeXML "</ITRForm:TurnoverGrsRcptForGSTIN>"
     Next
     End If
        writeXML "      </ITRForm:ScheduleGST>"
End Function
Function UVCase(valV1 As Variant) As Variant
    If (valV1 = "") Then
        valV1 = 0
    End If
        UVCase = valV1
End Function

Function XMLFooter()

Print #1, "</ITR3FORM:ITR3>"
Print #1, "</ITRETURN:ITR>"
End Function
Function STCGImmovablePropertyXML(ByVal myindex As Long) As String
    STCGImmovablePropertyXML = ""
    Dim j  As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim rangecells7 As Range

    Set rangecells1 = Sheet13.Range("CG_A_NameOfBuyer" & myindex).Cells
    Set rangecells2 = Sheet13.Range("CG_A_PanOfBuyer" & myindex).Cells
    Set rangecells7 = Sheet13.Range("CG_A_AadhaarOfBuyer" & myindex).Cells
    Set rangecells3 = Sheet13.Range("CG_A_PercentageShare" & myindex).Cells
    Set rangecells4 = Sheet13.Range("CG_A_Amount" & myindex).Cells
    Set rangecells5 = Sheet13.Range("CG_A_AddressOfBuyer" & myindex).Cells
    Set rangecells6 = Sheet13.Range("CG_A_PinCodeOfBuyer" & myindex).Cells

    setTblinfo_STCGImmovableProperty (myindex)

    ReDim STCG_NameOfBuyer(myindex, end_stcgImbcount)
    ReDim STCG_PANOfBuyer(myindex, end_stcgImbcount)
    ReDim STCG_AadharOfBuyer(myindex, end_stcgImbcount)
    ReDim STCG_PercentageShare(myindex, end_stcgImbcount)
    ReDim STCG_Amount(myindex, end_stcgImbcount)
    ReDim STCG_AddressOfProperty(myindex, end_stcgImbcount)
    ReDim STCG_PinCode(myindex, end_stcgImbcount)

    If end_stcgImbcount > 0 Then
    writeXML "<ITRForm:TrnsfImmblPrprty>"
        For j = 1 To end_stcgImbcount
        Dim cellrange As Variant
         Dim cellrange1 As Variant
        cellrange = GetMergedAddressCell(rangecells1, j)
    STCG_NameOfBuyer(myindex, j) = Sheet13.Range(cellrange).value
    cellrange1 = GetMergedAddressCell(rangecells2, j)
    STCG_PANOfBuyer(myindex, j) = Sheet13.Range(cellrange1).value
            STCG_AadharOfBuyer(myindex, j) = rangecells7.item(j).value
            STCG_PercentageShare(myindex, j) = rangecells3.item(j).value
            STCG_Amount(myindex, j) = rangecells4.item(j).value
            STCG_AddressOfProperty(myindex, j) = rangecells5.item(j).value
            STCG_PinCode(myindex, j) = rangecells6.item(j).value

            writeXML "<ITRForm:TrnsfImmblPrprtyDtls>"
            writeXML "<ITRForm:NameOfBuyer>" & STCG_NameOfBuyer(myindex, j) & "</ITRForm:NameOfBuyer>"
            If STCG_PANOfBuyer(myindex, j) <> "" Then
            writeXML "<ITRForm:PANofBuyer>" & UCase(STCG_PANOfBuyer(myindex, j)) & "</ITRForm:PANofBuyer>"
            End If

            If STCG_AadharOfBuyer(myindex, j) <> "" Then
            writeXML "<ITRForm:AaadhaarOfBuyer>" & UCase(STCG_AadharOfBuyer(myindex, j)) & "</ITRForm:AaadhaarOfBuyer>"
            End If
            writeXML "<ITRForm:PercentageShare>" & STCG_PercentageShare(myindex, j) & "</ITRForm:PercentageShare>"
            writeXML "<ITRForm:Amount>" & STCG_Amount(myindex, j) & "</ITRForm:Amount>"
            writeXML "<ITRForm:AddressOfProperty>" & STCG_AddressOfProperty(myindex, j) & "</ITRForm:AddressOfProperty>"
            writeXML "<ITRForm:PinCode>" & STCG_PinCode(myindex, j) & "</ITRForm:PinCode>"
            writeXML "</ITRForm:TrnsfImmblPrprtyDtls>"
        Next
    writeXML "</ITRForm:TrnsfImmblPrprty>"
    End If

End Function

Function LTCGImmovablePropertyXML(ByVal myindex As Long) As String
    LTCGImmovablePropertyXML = ""
    Dim j  As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim rangecells7 As Range

    Set rangecells1 = Sheet13.Range("CG_B_NameOfBuyer" & myindex).Cells
    Set rangecells2 = Sheet13.Range("CG_B_PanOfBuyer" & myindex).Cells
    Set rangecells7 = Sheet13.Range("CG_B_AadhaaeOfBuyer" & myindex).Cells

    Set rangecells3 = Sheet13.Range("CG_B_PercentageShare" & myindex).Cells
    Set rangecells4 = Sheet13.Range("CG_B_Amount" & myindex).Cells
    Set rangecells5 = Sheet13.Range("CG_B_AddressOfBuyer" & myindex).Cells
    Set rangecells6 = Sheet13.Range("CG_B_PinCodeOfBuyer" & myindex).Cells

    setTblinfo_LTCGImmovableProperty (myindex)

    ReDim LTCG_NameOfBuyer(myindex, end_ltcgImbcount)
    ReDim LTCG_PANOfBuyer(myindex, end_ltcgImbcount)
    ReDim LTCG_AadharOfBuyer(myindex, end_ltcgImbcount)
    ReDim LTCG_PercentageShare(myindex, end_ltcgImbcount)
    ReDim LTCG_Amount(myindex, end_ltcgImbcount)
    ReDim LTCG_AddressOfProperty(myindex, end_ltcgImbcount)
    ReDim LTCG_PinCode(myindex, end_ltcgImbcount)

    If end_ltcgImbcount > 0 Then
    writeXML "<ITRForm:TrnsfImmblPrprty>"

        For j = 1 To end_ltcgImbcount
           Dim cellrange As Variant
         Dim cellrange1 As Variant
        cellrange = GetMergedAddressCell(rangecells1, j)
        LTCG_NameOfBuyer(myindex, j) = Sheet13.Range(cellrange).value
         cellrange1 = GetMergedAddressCell(rangecells2, j)
         LTCG_PANOfBuyer(myindex, j) = Sheet13.Range(cellrange1).value

             LTCG_AadharOfBuyer(myindex, j) = rangecells7.item(j).value
'            LTCG_NameOfBuyer(myindex, j) = rangecells1.item(j).value
'            LTCG_PANOfBuyer(myindex, j) = rangecells2.item(j).value
            LTCG_PercentageShare(myindex, j) = rangecells3.item(j).value
            LTCG_Amount(myindex, j) = rangecells4.item(j).value
            LTCG_AddressOfProperty(myindex, j) = rangecells5.item(j).value
            LTCG_PinCode(myindex, j) = rangecells6.item(j).value

            writeXML "<ITRForm:TrnsfImmblPrprtyDtls>"
            writeXML "<ITRForm:NameOfBuyer>" & LTCG_NameOfBuyer(myindex, j) & "</ITRForm:NameOfBuyer>"
            If LTCG_PANOfBuyer(myindex, j) <> "" Then
            writeXML "<ITRForm:PANofBuyer>" & UCase(LTCG_PANOfBuyer(myindex, j)) & "</ITRForm:PANofBuyer>"
            End If
            If LTCG_AadharOfBuyer(myindex, j) <> "" Then
            writeXML "<ITRForm:AaadhaarOfBuyer>" & UCase(LTCG_AadharOfBuyer(myindex, j)) & "</ITRForm:AaadhaarOfBuyer>"
            End If
            writeXML "<ITRForm:PercentageShare>" & LTCG_PercentageShare(myindex, j) & "</ITRForm:PercentageShare>"
            writeXML "<ITRForm:Amount>" & LTCG_Amount(myindex, j) & "</ITRForm:Amount>"
            writeXML "<ITRForm:AddressOfProperty>" & LTCG_AddressOfProperty(myindex, j) & "</ITRForm:AddressOfProperty>"
            writeXML "<ITRForm:PinCode>" & LTCG_PinCode(myindex, j) & "</ITRForm:PinCode>"
            writeXML "</ITRForm:TrnsfImmblPrprtyDtls>"
        Next

    writeXML "</ITRForm:TrnsfImmblPrprty>"
    End If

End Function

Sub setTblinfo_STCGImmovableProperty(ByVal myindex As Long)
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("CG_A_NameOfBuyer" & myindex).count
 Set rangecells = Sheet13.Range("CG_A_NameOfBuyer" & myindex).Cells

Dim scode As Variant
 For mIntCtr = 1 To mIntCells
    scode = rangecells.item(mIntCtr).value

    If Not scode = "" Then
        ccount = ccount + 1
    End If
 Next
 end_stcgImbcount = ccount
 End Sub

 Sub setTblinfo_LTCGImmovableProperty(ByVal myindex As Long)
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("CG_B_NameOfBuyer" & myindex).count
 Set rangecells = Sheet13.Range("CG_B_NameOfBuyer" & myindex).Cells

Dim scode As Variant
 For mIntCtr = 1 To mIntCells
    scode = rangecells.item(mIntCtr).value

    If Not scode = "" Then
        ccount = ccount + 1
    End If
 Next
 end_ltcgImbcount = ccount
 End Sub
Function GrossSalary(ByVal myindex As Long) As String
    GrossSalary = ""
    Dim j  As Long

    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range

    Set rangecells1 = Sheet39.Range("Salary.Section" & myindex).Cells
    Set rangecells2 = Sheet39.Range("Salary.salaryDescription" & myindex).Cells
    Set rangecells3 = Sheet39.Range("Salary.SalaryAmount" & myindex).Cells

    setTblinfo_GrossSalary (myindex)
    ReDim SalaryNature_Salary(myindex, end_SalaryNature)
    ReDim SalaryDes_Salary(myindex, end_SalaryNature)
    ReDim SalaryAmount_Salary(myindex, end_SalaryNature)

    If end_SalaryNature > 0 Then
        writeXML "<ITRForm:NatureOfSalary>"

        For j = 1 To end_SalaryNature
            SalaryNature_Salary(myindex, j) = rangecells1.item(j).value
            SalaryDes_Salary(myindex, j) = rangecells2.item(j).value
            SalaryAmount_Salary(myindex, j) = rangecells3.item(j).value

            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "BASIC SALARY" Then SalaryNature_Salary(myindex, j) = "1"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "DEARNESS ALLOWANCE" Then SalaryNature_Salary(myindex, j) = "2"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "CONVEYANCE ALLOWANCE" Then SalaryNature_Salary(myindex, j) = "3"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "HOUSE RENT ALLOWANCE" Then SalaryNature_Salary(myindex, j) = "4"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "LEAVE TRAVEL ALLOWANCE" Then SalaryNature_Salary(myindex, j) = "5"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "CHILDREN EDUCATION ALLOWANCE" Then SalaryNature_Salary(myindex, j) = "6"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "OTHER ALLOWANCE" Then SalaryNature_Salary(myindex, j) = "7"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "THE CONTRIBUTION MADE  BY THE EMPLOYER  TOWARDS  PENSION SCHEME AS REFERRED U/S 80CCD" Then SalaryNature_Salary(myindex, j) = "8"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "AMOUNT DEEMED TO BE INCOME UNDER RULE 11 OF FOURTH SCHEDULE" Then SalaryNature_Salary(myindex, j) = "9"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "AMOUNT DEEMED TO BE INCOME UNDER RULE 6 OF FOURTH SCHEDULE" Then SalaryNature_Salary(myindex, j) = "10"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "ANNUITY OR PENSION" Then SalaryNature_Salary(myindex, j) = "11"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "COMMUTED PENSION" Then SalaryNature_Salary(myindex, j) = "12"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "GRATUITY" Then SalaryNature_Salary(myindex, j) = "13"
            If UCase(Trim(Mid(SalaryNature_Salary(myindex, j), 1, 4))) = "FEES" Then SalaryNature_Salary(myindex, j) = "14"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "ADVANCE OF SALARY" Then SalaryNature_Salary(myindex, j) = "15"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "LEAVE ENCASHMENT" Then SalaryNature_Salary(myindex, j) = "16"
            If UCase(Trim(SalaryNature_Salary(myindex, j))) = "OTHERS" Then SalaryNature_Salary(myindex, j) = "OTH"

            writeXML "<ITRForm:OthersIncDtls>"
            writeXML "<ITRForm:NatureDesc>" & SalaryNature_Salary(myindex, j) & "</ITRForm:NatureDesc>"

            If SalaryDes_Salary(myindex, j) <> "" Then
            writeXML "<ITRForm:OthNatOfInc>" & SalaryDes_Salary(myindex, j) & "</ITRForm:OthNatOfInc>"
            End If

            writeXML "<ITRForm:OthAmount>" & SalaryAmount_Salary(myindex, j) & "</ITRForm:OthAmount>"
            writeXML "</ITRForm:OthersIncDtls>"
        Next

        writeXML "</ITRForm:NatureOfSalary>"
    End If

End Function
Sub setTblinfo_GrossSalary(ByVal myindex As Long)
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet39.Range("Salary.Section" & myindex).count
 Set rangecells = Sheet39.Range("Salary.Section" & myindex).Cells

 Dim scode As Variant
 For mIntCtr = 1 To mIntCells
    scode = rangecells.item(mIntCtr).value
    If isdropdownblank(scode) Then
    scode = ""
    End If

    If Not scode = "" Then
        ccount = ccount + 1
    End If
 Next
 end_SalaryNature = ccount
End Sub

Function ValueOfPerquisites(ByVal myindex As Long) As String
    ValueOfPerquisites = ""
    Dim j  As Long

    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range

    Set rangecells1 = Sheet39.Range("Salary.ValueSection" & myindex).Cells
    Set rangecells2 = Sheet39.Range("Salary.ValueDescription" & myindex).Cells
    Set rangecells3 = Sheet39.Range("Salary.ValueAmount" & myindex).Cells

    setTblinfo_Perquisites (myindex)
    ReDim ValueOfPerquisitesNature_Salary(myindex, end_PerquisitesNature)
    ReDim ValueOfPerquisitesDes_Salary(myindex, end_PerquisitesNature)
    ReDim ValueOfPerquisitesAmount_Salary(myindex, end_PerquisitesNature)

    If end_PerquisitesNature > 0 Then
        writeXML "<ITRForm:NatureOfPerquisites>"

        For j = 1 To end_PerquisitesNature
            ValueOfPerquisitesNature_Salary(myindex, j) = rangecells1.item(j).value
            ValueOfPerquisitesDes_Salary(myindex, j) = rangecells2.item(j).value
            ValueOfPerquisitesAmount_Salary(myindex, j) = rangecells3.item(j).value

            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "ACCOMMODATION" Then ValueOfPerquisitesNature_Salary(myindex, j) = "1"
            If UCase(Trim(Mid(ValueOfPerquisitesNature_Salary(myindex, j), 1, 4))) = "CARS" Then ValueOfPerquisitesNature_Salary(myindex, j) = "2"
            If UCase(Trim(Mid(ValueOfPerquisitesNature_Salary(myindex, j), 1, 7))) = "SWEEPER" Then ValueOfPerquisitesNature_Salary(myindex, j) = "3"
            If UCase(Trim(Mid(ValueOfPerquisitesNature_Salary(myindex, j), 1, 3))) = "GAS" Then ValueOfPerquisitesNature_Salary(myindex, j) = "4"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "INTEREST FREE OR CONCESSIONAL LOANS" Then ValueOfPerquisitesNature_Salary(myindex, j) = "5"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "HOLIDAY EXPENSES" Then ValueOfPerquisitesNature_Salary(myindex, j) = "6"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "FREE OR CONCESSIONAL TRAVEL" Then ValueOfPerquisitesNature_Salary(myindex, j) = "7"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "FREE MEALS" Then ValueOfPerquisitesNature_Salary(myindex, j) = "8"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "FREE EDUCATION" Then ValueOfPerquisitesNature_Salary(myindex, j) = "9"
            If UCase(Trim(Mid(ValueOfPerquisitesNature_Salary(myindex, j), 1, 5))) = "GIFTS" Then ValueOfPerquisitesNature_Salary(myindex, j) = "10"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "CREDIT CARD EXPENSES" Then ValueOfPerquisitesNature_Salary(myindex, j) = "11"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "CLUB EXPENSES" Then ValueOfPerquisitesNature_Salary(myindex, j) = "12"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "USE OF MOVABLE ASSETS BY EMPLOYEES" Then ValueOfPerquisitesNature_Salary(myindex, j) = "13"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "TRANSFER OF ASSETS TO EMPLOYEE" Then ValueOfPerquisitesNature_Salary(myindex, j) = "14"
            If UCase(Trim(Mid(ValueOfPerquisitesNature_Salary(myindex, j), 1, 26))) = "VALUE OF ANY OTHER BENEFIT" Then ValueOfPerquisitesNature_Salary(myindex, j) = "15"
            If UCase(Trim(Mid(ValueOfPerquisitesNature_Salary(myindex, j), 1, 13))) = "STOCK OPTIONS" Then ValueOfPerquisitesNature_Salary(myindex, j) = "16"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "TAX PAID BY EMPLOYER ON NON-MONETARY PERQUISITE" Then ValueOfPerquisitesNature_Salary(myindex, j) = "17"
            If UCase(Trim(ValueOfPerquisitesNature_Salary(myindex, j))) = "OTHER BENEFITS OR AMENITIES" Then ValueOfPerquisitesNature_Salary(myindex, j) = "OTH"


            writeXML "<ITRForm:OthersIncDtls>"
            writeXML "<ITRForm:NatureDesc>" & ValueOfPerquisitesNature_Salary(myindex, j) & "</ITRForm:NatureDesc>"

            If ValueOfPerquisitesDes_Salary(myindex, j) <> "" Then
            writeXML "<ITRForm:OthNatOfInc>" & ValueOfPerquisitesDes_Salary(myindex, j) & "</ITRForm:OthNatOfInc>"
            End If

            writeXML "<ITRForm:OthAmount>" & ValueOfPerquisitesAmount_Salary(myindex, j) & "</ITRForm:OthAmount>"
            writeXML "</ITRForm:OthersIncDtls>"
        Next

        writeXML "</ITRForm:NatureOfPerquisites>"
    End If

End Function

Sub setTblinfo_Perquisites(ByVal myindex As Long)
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet39.Range("Salary.ValueSection" & myindex).count
 Set rangecells = Sheet39.Range("Salary.ValueSection" & myindex).Cells

 Dim scode As Variant
 For mIntCtr = 1 To mIntCells
    scode = rangecells.item(mIntCtr).value
    If isdropdownblank(scode) Then
    scode = ""
    End If

    If Not scode = "" Then
        ccount = ccount + 1
    End If
 Next
 end_PerquisitesNature = ccount
End Sub

Function ProfitsinLieuOfSalary(ByVal myindex As Long) As String
    ProfitsinLieuOfSalary = ""
    Dim j  As Long

    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range

    Set rangecells1 = Sheet39.Range("Salary.ProfitSection" & myindex).Cells
    Set rangecells2 = Sheet39.Range("Salary.ProfitDescription" & myindex).Cells
    Set rangecells3 = Sheet39.Range("Salary.ProfitAmount" & myindex).Cells

    setTblinfo_ProfitsinLieuOfSalary (myindex)
    ReDim ProfitsinLieuOfSalaryNature_Salary(myindex, end_ProfitsinLieuOfSalary)
    ReDim ProfitsinLieuOfSalaryDes_Salary(myindex, end_ProfitsinLieuOfSalary)
    ReDim ProfitsinLieuOfSalaryAmount_Salary(myindex, end_ProfitsinLieuOfSalary)

    If end_ProfitsinLieuOfSalary > 0 Then
        writeXML "<ITRForm:NatureOfProfitInLieuOfSalary>"

        For j = 1 To end_ProfitsinLieuOfSalary
            ProfitsinLieuOfSalaryNature_Salary(myindex, j) = rangecells1.item(j).value
            ProfitsinLieuOfSalaryDes_Salary(myindex, j) = rangecells2.item(j).value
            ProfitsinLieuOfSalaryAmount_Salary(myindex, j) = rangecells3.item(j).value

            If UCase(Trim(Mid(ProfitsinLieuOfSalaryNature_Salary(myindex, j), 1, 12))) = "COMPENSATION" Then ProfitsinLieuOfSalaryNature_Salary(myindex, j) = "1"
            If UCase(Trim(Mid(ProfitsinLieuOfSalaryNature_Salary(myindex, j), 1, 11))) = "ANY PAYMENT" Then ProfitsinLieuOfSalaryNature_Salary(myindex, j) = "2"
            If UCase(Trim(Mid(ProfitsinLieuOfSalaryNature_Salary(myindex, j), 1, 10))) = "ANY AMOUNT" Then ProfitsinLieuOfSalaryNature_Salary(myindex, j) = "3"
            If UCase(Trim(Mid(ProfitsinLieuOfSalaryNature_Salary(myindex, j), 1, 9))) = "ANY OTHER" Then ProfitsinLieuOfSalaryNature_Salary(myindex, j) = "OTH"

            writeXML "<ITRForm:OthersIncDtls>"
            writeXML "<ITRForm:NatureDesc>" & ProfitsinLieuOfSalaryNature_Salary(myindex, j) & "</ITRForm:NatureDesc>"

            If ProfitsinLieuOfSalaryDes_Salary(myindex, j) <> "" Then
            writeXML "<ITRForm:OthNatOfInc>" & ProfitsinLieuOfSalaryDes_Salary(myindex, j) & "</ITRForm:OthNatOfInc>"
            End If

            writeXML "<ITRForm:OthAmount>" & ProfitsinLieuOfSalaryAmount_Salary(myindex, j) & "</ITRForm:OthAmount>"
            writeXML "</ITRForm:OthersIncDtls>"
        Next

        writeXML "</ITRForm:NatureOfProfitInLieuOfSalary>"
    End If

End Function

Sub setTblinfo_ProfitsinLieuOfSalary(ByVal myindex As Long)
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet39.Range("Salary.ProfitSection" & myindex).count
 Set rangecells = Sheet39.Range("Salary.ProfitSection" & myindex).Cells

 Dim scode As Variant
 For mIntCtr = 1 To mIntCells
    scode = rangecells.item(mIntCtr).value
    If isdropdownblank(scode) Then
    scode = ""
    End If

    If Not scode = "" Then
        ccount = ccount + 1
    End If
 Next
 end_ProfitsinLieuOfSalary = ccount
End Sub
'XML Generation Function For 80D
Function Schedule80D()
subProcCaption = "Schedule 80D"
noOfProcessSub = 10

writeXML "<ITRForm:Schedule80D>"
writeXML "<ITRForm:Sec80DSelfFamSrCtznHealth>"


    If FamilyMember = "Yes" Then
     FamilyMember = "Y"
     ElseIf FamilyMember = "No" Then
     FamilyMember = "N"
     ElseIf FamilyMember = "Not Claiming for Self/Family" Then
     FamilyMember = "S"
     End If


    If FamilyMember <> "" And FamilyMember <> "(Select)" Then
        writeXML "<ITRForm:SeniorCitizenFlag>" & UCase(FamilyMember) & "</ITRForm:SeniorCitizenFlag>"
    End If

    If Sheet55.Range("Self_And_Family_80D").value <> "" Then
        writeXML "<ITRForm:SelfAndFamily>" & Sheet55.Range("Self_And_Family_80D").value & "</ITRForm:SelfAndFamily>"
    End If

    If Sheet55.Range("Health_Insurance_80D").value <> "" Then
        writeXML "<ITRForm:HealthInsPremSlfFam>" & Sheet55.Range("Health_Insurance_80D").value & "</ITRForm:HealthInsPremSlfFam>"
    End If

    If Sheet55.Range("Preventive_Health_80D").value <> "" Then
        writeXML "<ITRForm:PrevHlthChckUpSlfFam>" & Sheet55.Range("Preventive_Health_80D").value & "</ITRForm:PrevHlthChckUpSlfFam>"
    End If

    If Sheet55.Range("Senior_Citizen_80D").value <> "" Then
        writeXML "<ITRForm:SelfAndFamilySeniorCitizen>" & Sheet55.Range("Senior_Citizen_80D").value & "</ITRForm:SelfAndFamilySeniorCitizen>"
    End If

    If Sheet55.Range("Health_InsuranceSC_80D").value <> "" Then
        writeXML "<ITRForm:HlthInsPremSlfFamSrCtzn>" & Sheet55.Range("Health_InsuranceSC_80D").value & "</ITRForm:HlthInsPremSlfFamSrCtzn>"
    End If

    If Sheet55.Range("Preventive_Health_SC_80D").value <> "" Then
        writeXML "<ITRForm:PrevHlthChckUpSlfFamSrCtzn>" & Sheet55.Range("Preventive_Health_SC_80D").value & "</ITRForm:PrevHlthChckUpSlfFamSrCtzn>"
    End If

    If Sheet55.Range("Medical_Expenditure_SC_80D").value <> "" Then
        writeXML "<ITRForm:MedicalExpSlfFamSrCtzn>" & Sheet55.Range("Medical_Expenditure_SC_80D").value & "</ITRForm:MedicalExpSlfFamSrCtzn>"
    End If

If Sheet1.Range("sheet1.Status").value <> "H-HUF" Then
    If SeniorCitizen = "Yes" Then
     SeniorCitizen = "Y"
     ElseIf SeniorCitizen = "No" Then
     SeniorCitizen = "N"
     ElseIf SeniorCitizen = "Not claiming for Parents" Then
     SeniorCitizen = "P"
    End If


    If SeniorCitizen <> "" And SeniorCitizen <> "(Select)" Then
    writeXML "<ITRForm:ParentsSeniorCitizenFlag>" & UCase(SeniorCitizen) & "</ITRForm:ParentsSeniorCitizenFlag>"
    End If

    If Sheet55.Range("Parents_80D").value <> "" Then
        writeXML "<ITRForm:Parents>" & Sheet55.Range("Parents_80D").value & "</ITRForm:Parents>"
    End If

    If Sheet55.Range("Health_Insurance2_80D").value <> "" Then
        writeXML "<ITRForm:HlthInsPremParents>" & Sheet55.Range("Health_Insurance2_80D").value & "</ITRForm:HlthInsPremParents>"
    End If

    If Sheet55.Range("Preventive_Health2_80D").value <> "" Then
        writeXML "<ITRForm:PrevHlthChckUpParents>" & Sheet55.Range("Preventive_Health2_80D").value & "</ITRForm:PrevHlthChckUpParents>"
    End If

    If Sheet55.Range("Parents_SC_80D").value <> "" Then
        writeXML "<ITRForm:ParentsSeniorCitizen>" & Sheet55.Range("Parents_SC_80D").value & "</ITRForm:ParentsSeniorCitizen>"
    End If

    If Sheet55.Range("Health_Insurance3_80D").value <> "" Then
        writeXML "<ITRForm:HlthInsPremParentsSrCtzn>" & Sheet55.Range("Health_Insurance3_80D").value & "</ITRForm:HlthInsPremParentsSrCtzn>"
    End If

    If Sheet55.Range("Preventive_Health3_80D").value <> "" Then
        writeXML "<ITRForm:PrevHlthChckUpParentsSrCtzn>" & Sheet55.Range("Preventive_Health3_80D").value & "</ITRForm:PrevHlthChckUpParentsSrCtzn>"
    End If

    If Sheet55.Range("Medical_Expenditure2_80D").value <> "" Then
        writeXML "<ITRForm:MedicalExpParentsSrCtzn>" & Sheet55.Range("Medical_Expenditure2_80D").value & "</ITRForm:MedicalExpParentsSrCtzn>"
    End If
End If

    If Sheet55.Range("Eligible_Amount_80D").value <> "" Then
        writeXML "<ITRForm:EligibleAmountOfDedn>" & Sheet55.Range("Eligible_Amount_80D").value & "</ITRForm:EligibleAmountOfDedn>"
    End If


writeXML "</ITRForm:Sec80DSelfFamSrCtznHealth>"
writeXML "</ITRForm:Schedule80D>"

End Function

'XML Generation Function For Schedule TPSA
Function ScheduleTPSA()
subProcCaption = "Schedule 80D"
noOfProcessSub = 10
Dim i As Long
writeXML "<ITRForm:ScheduleTPSA>"

    If Sheet57.Range("TPSA_92CE_Amount").value <> "" Then
        writeXML "<ITRForm:AmtPrimaryAdjUs92CE_2A>" & Round(Sheet57.Range("TPSA_92CE_Amount").value, 0) & "</ITRForm:AmtPrimaryAdjUs92CE_2A>"
        Else
        writeXML "<ITRForm:AmtPrimaryAdjUs92CE_2A>" & 0 & "</ITRForm:AmtPrimaryAdjUs92CE_2A>"
    End If

    If Sheet57.Range("TPSA_Additional_Income").value <> "" Then
        writeXML "<ITRForm:AdditionalIncTax18PercAbove>" & Round(Sheet57.Range("TPSA_Additional_Income").value, 0) & "</ITRForm:AdditionalIncTax18PercAbove>"
        Else
        writeXML "<ITRForm:AdditionalIncTax18PercAbove>" & 0 & "</ITRForm:AdditionalIncTax18PercAbove>"
    End If

    If Sheet57.Range("TPSA_Surcharge").value <> "" Then
        writeXML "<ITRForm:Surcharge12Perc>" & Round(Sheet57.Range("TPSA_Surcharge").value, 0) & "</ITRForm:Surcharge12Perc>"
        Else
        writeXML "<ITRForm:Surcharge12Perc>" & 0 & "</ITRForm:Surcharge12Perc>"
    End If

    If Sheet57.Range("TPSA_Education_Cess").value <> "" Then
        writeXML "<ITRForm:HealthEducationCess>" & Round(Sheet57.Range("TPSA_Education_Cess").value, 0) & "</ITRForm:HealthEducationCess>"
        Else
        writeXML "<ITRForm:HealthEducationCess>" & 0 & "</ITRForm:HealthEducationCess>"
    End If

    If Sheet57.Range("TPSA_AddtionalTax_Total").value <> "" Then
        writeXML "<ITRForm:TotalAdditionalTax>" & Round(Sheet57.Range("TPSA_AddtionalTax_Total").value, 0) & "</ITRForm:TotalAdditionalTax>"
        Else
        writeXML "<ITRForm:TotalAdditionalTax>" & 0 & "</ITRForm:TotalAdditionalTax>"
    End If

    If Sheet57.Range("TPSA_Taxpaid").value <> "" Then
        writeXML "<ITRForm:TaxesPaid>" & Round(Sheet57.Range("TPSA_Taxpaid").value, 0) & "</ITRForm:TaxesPaid>"
        Else
        writeXML "<ITRForm:TaxesPaid>" & 0 & "</ITRForm:TaxesPaid>"
    End If

    If Sheet57.Range("TPSA_Nettax").value <> "" Then
        writeXML "<ITRForm:NetTaxPayable>" & Round(Sheet57.Range("TPSA_Nettax").value, 0) & "</ITRForm:NetTaxPayable>"
        Else
        writeXML "<ITRForm:NetTaxPayable>" & 0 & "</ITRForm:NetTaxPayable>"
    End If



    If Not IsEmpty(BSR_TPSC) And UBound(BSR_TPSC) > 0 Then
        For i = 1 To UBound(BSR_TPSC)
            writeXML "<ITRForm:DtlsTaxesPaid>"
            If BSR_TPSC(i) <> "" Then
            writeXML "<ITRForm:BSRCode>" & UCase(BSR_TPSC(i)) & "</ITRForm:BSRCode>"
            End If

            If BankName_TPSC(i) <> "" Then
            writeXML "<ITRForm:BankBranchName>" & UCase(BankName_TPSC(i)) & "</ITRForm:BankBranchName>"
            End If

            If Date_TPSC(i) <> "" Then
            writeXML "<ITRForm:DateDep>" & Dformat(Date_TPSC(i), "") & "</ITRForm:DateDep>"
            End If

            If SrlNoOfChaln_TPSC(i) <> "" Then
            writeXML "<ITRForm:SrlNoOfChaln>" & SrlNoOfChaln_TPSC(i) & "</ITRForm:SrlNoOfChaln>"
            End If

            If Amount_TPSC(i) <> "" Then
            writeXML "<ITRForm:Amount>" & Amount_TPSC(i) & "</ITRForm:Amount>"
            End If
            writeXML "</ITRForm:DtlsTaxesPaid>"
         Next
         End If

            writeXML "<ITRForm:TotalAmountDeposited>" & Sheet57.Range("TPSC_Amount_Deposited").value & "</ITRForm:TotalAmountDeposited>"
            writeXML "</ITRForm:ScheduleTPSA>"
End Function















