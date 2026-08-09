Attribute VB_Name = "mdIncomeDetails"
Option Explicit
'Declaring Variables

Dim msgImport As String
'Ayush_26-27
Public end_80CCC As Variant
Public end_80CCCType As Variant
Public end_80CCCName As Variant
Public end_80CCCAmount As Variant
Public Type_80CCC As Variant
Public Name_80CCC As Variant
Public Amount_80CCC As Variant
Public end_pran As Variant
Public pran_80CCC As Variant


Public end_80CCC_1 As Variant
Public end_80CCC_1_Type As Variant
Public end_80CCC_1_Name As Variant
Public end_80CCC_1_Amount As Variant
Public Type_1_80CCC As Variant
Public Name_1_80CCC As Variant
Public Amount_1_80CCC As Variant

Public end_80CCC_1b As Variant
Public end_80CCC_1b_Type As Variant
Public end_80CCC_1b_Name As Variant
Public end_80CCC_1b_Amount As Variant
Public Type_1b_80CCC As Variant
Public Name_1b_80CCC As Variant
Public Amount_1b_80CCC As Variant




Dim validateType As Variant
Dim dedValidateResult As Variant
Dim errmsgfiling As Variant
Dim errmsgID As Variant

Public ProvisoFlag As Variant
Public DepositAmountFlag As Variant
Public DepositAmount As Variant
Public AggrigateAmountFlag As Variant
Public AggrigateAmount As Variant
Public AggrigateAmountFlag1 As Variant
Public AggrigateAmount1 As Variant
Public Investment As Variant


Public SELECT80D, SELECT80DD, SELECT80DDB, SELECT80U, SELECT80DB, SELECT80DC As Variant
'Newly added by Bindu
Public SELECT80DB_Specified_Disease As Variant

Public zipCode As Variant
Public HASZIP As Variant
Public firstName As Variant
Public middleName As Variant
Public LastName As Variant
Public PAN As Variant
Public Flat As Variant
Public residenceName As Variant
Public status As Variant
Public roadOrStreet As Variant
Public Area As Variant
Public AssesseeDob As String
Public NatureEmp As String
Public Gender As Variant
Public City As Variant
Public PrevAYBAC As Variant
Public State As Variant
Public Country As Variant
Public PinCode As Variant
Public Email As Variant
Public Email2 As Variant
Public mobileNo, CountrycodeMobileNo2, CountrycodeMobileNo1 As Variant
Public stDcode As Variant
Public phoneNo As Variant
Public PhoneNo2 As Variant
Public EmpCategory As Variant
Public IncomeTaxWard As Variant
Public AadhaarNumber As Variant
Public AadhaarEnrolNumber As Variant
Public ReturnFilledSection As Variant

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
'-----------------------

Public ReturnFilledSection1 As Variant
Public ReturnFilledSection2 As Variant
'Public GovernedByPortugese As Variant
Public PanOfSpouse As Variant
Public TypeOfReturn As Variant
Public RecieptNo As Variant
Public DateOfFillingRetrun As Variant
Public RecieptNo1 As Variant
Public DateOfFillingRetrun1 As Variant
Public NoticeNo As Variant
'Public NoticeNumber As Variant
Public DateOfNotice As Variant
'Public ResidentialStatus As Variant

'Public AadharFlag As Variant
'Public AadharNumber As Variant
'Public TaxStatusFlag As Variant

Public IncomeBP As Variant
Public IncomeSAl As Variant
Public IncomeSAl1 As Variant
Public IncomeSAl2 As Variant
Public IncomeSAl3 As Variant
Public IncomeSAl4 As Variant
Public IncomeSAl5 As Variant
Public IncomeSAl5a As Variant
Public IncomeSAl6 As Variant
Public IncomeSAl7 As Variant
Public IncomeSAl8 As Variant
Public IncomeHP As Variant
Public IncomeHP1 As Variant
Public IncomeHP2 As Variant
Public IncomeHP3 As Variant
Public IncomeHP4 As Variant
Public IncomeHP5 As Variant
Public IncomeHP6 As Variant
Public IncomeSAlnet As Variant
Public IncomeHP1a As Variant
Public IncomeHP1b As Variant

Public TypeOfHP As Variant
Public IncomeOS As Variant
Public IncomeGross As Variant
Public GrossTotIncomeIncLTCG112A As Variant 'added by Chetan V3.4 on 07/02/2025

Public Ded80C As Variant
Public Ded80CCC As Variant
Public Ded80CCDSelfEmployee As Variant
Public Ded80CCDEmployer As Variant
Public Ded80CCD1B_SE As Variant
'Public Ded80CCG As Variant
Public Ded80D As Variant
Public Ded80DB As Variant
Public Ded80DC As Variant
Public Ded80DD As Variant
Public Ded80DDB As Variant
Public Ded80E As Variant
Public Ded80EE As Variant
Public Ded80EEA As Variant
Public Ded80EEB As Variant
Public Ded80G As Variant
Public Ded80GG As Variant
Public Ded80GGA As Variant
Public Ded80GGC As Variant
Public Ded80QQB, Ded80RRB, Ded80TTA, Ded80TTB, Ded80CCH, Ded80U As Variant
Public TotalDed As Variant

Public Ded80C_calc As Variant
Public Ded80CCC_calc As Variant
Public Ded80CCDSelfEmployee_calc As Variant
Public Ded80CCDEmployer_calc As Variant
Public Ded80CCD1B_SE_calc As Variant
'Public Ded80CCG_calc As Variant
Public Ded80D_calc As Variant
Public Ded80DD_calc As Variant
Public Ded80DDB_calc As Variant
Public Ded80E_calc As Variant
Public Ded80EE_calc As Variant
Public Ded80EEA_calc As Variant
Public Ded80EEB_calc As Variant
Public Ded80G_calc As Variant
Public Ded80GG_calc As Variant
Public Ded80GGA_calc As Variant
Public Ded80GGC_calc As Variant
Public Ded80QQB_calc, Ded80RRB_calc, Ded80TTA_calc, Ded80TTB_calc, Ded80CCH_calc, Ded80U_calc As Variant
Public TotalDed_calc As Variant
Public end_IDOthersAmt, end_IDOthersNOI, Others_IDNOI, Others_IDNOI1, Others_IDAmt As Variant
Public IncomeSAl5b, IncomeSAl5c As Variant

Public end_OSOthersNOI, end_OSOthersAmt, Others_OSNOI, Others_OSNOI1, Others_OSAmt As Variant

Public TotalTaxIncome As Variant
Public TaxPayableTotalIncome As Variant
Public Rebate87A As Variant
Public TaxafterRebate As Variant
Public Surcharge As Variant
Public EducationCess As Variant
Public TotalTaxSurchargeEdu As Variant
Public Relief89 As Variant
Public Relief89A As Variant

Public BalanceTax As Variant
Public InterestPay234a As Variant
Public InterestPay234b As Variant
Public InterestPay234c As Variant
Public InterestPay234f As Variant
'Public TotalInterest234abc As Variant
Public TotalTaxInterest As Variant
Public msg14didgits As String
Public Representativeassesseeflg As Variant
Public NameRepAssessee As Variant
Public CapacityRep As Variant
Public AddressRepAssessee As Variant
Public PANRepAssessee As Variant
Public AadhaarRepAssessee As Variant
'Ayush_12/01/2026_V0.2
Public EmailRepAssessee As Variant
Public CountryCodeRepAssessee As Variant
Public MobileRepAssessee As Variant





Sub cmdValidateButton_Click()
    ValidateSheetIncomeDetails
    fmsgboxStatus "Sheet Income Details is Ok" ', vbOKOnly, vbMessageCaption
End Sub

Sub cmdNextButton_Click()

    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
       Sheet201.Activate
       ElseIf Sheet22.Visible = True Then
        Sheet22.Activate
    ElseIf Sheet15.Visible = True Then
        Sheet15.Activate
    
    
    Else
       Sheet3.Activate
    End If
    
    

End Sub

Sub cmdGenerateXML_Click()
If CalculateTaxCheckFlag = False Then
'MsgBox "Please ensure to click on 'Calculate Tax' button before generating XML.", vbCritical, "Error"
MsgBox "Please ensure to click on 'Calculate Tax' button before generating JSON."
End
End If
GenerateJSON.Generate
End Sub

Sub cmdPrintButton_Click()
printWorkSheet
End Sub

Sub cmdImportButton_Click()
'IMPPrevVersion
End Sub

Sub cmdImportPersonalDetails_Click()
ImportJSON.Import
End Sub

Sub cmdHelpButton_Click()
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheets("Help").Visible = xlSheetVisible
    Sheets("Help").Activate
    ThisWorkbook.Protect Password:=getmsgstate
End Sub
Sub ValidateSheetIncomeDetails()
mdIncomeDetails.AssignValues
 ValidateIncOnImport
 
If Not mdIncomeDetails.ValidateSheetID Then
    Sheet1.Activate
    fmsgbox msgError ', vbOKOnly, "Error(s)!"
    CloseMsg
End If



If Not ValidateMorethan14Digits Then
    Sheet1.Activate
    fmsgboxsmall msg14didgits ', vbOKOnly, "Error(s)!"
    CloseMsg
End If


End Sub

Function ValidateMorethan14Digits() As Boolean
ValidateMorethan14Digits = True

'If Len(Sheet1.Range("IncD.GrossTotIncome").value) > 14 Then
'Newly updated by Bindu as per DE V3 on 4th Feb 2025
If Len(Sheet1.Range("IncD.GrossTotIncome_New").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Gross Total Income cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


If Len(Sheet1.Range("IncD.TotalChapVIADeductions_Input").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Total Deductions cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


If Len(Sheet1.Range("IncD.TotalChapVIADeductions").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Total Deductions (System Calculated) cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If

'If Len(Sheet1.Range("IncD.TotalIncome").value) > 14 Then
'As per DE V3 updated by Bindu on 4th Feb 2025
If Len(Sheet1.Range("IncD.TotalIncome_New").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Taxable Total Income cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


If Len(Sheet1.Range("IncD.TotalTaxPayable").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Tax Payable on Total Income cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


If Len(Sheet1.Range("IncD.Rebate87A").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Rebate u/s 87A cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


If Len(Sheet1.Range("IncD.TaxPayableOnRebate").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Tax payable after Rebate cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


'If Len(Sheet1.Range("IncD.SurchargeOnTaxPayable").value) > 14 Then
'    ValidateMorethan14Digits = False
'    msg14didgits = msg14didgits + "Surcharge if Taxable Income exceeds One crore cannot exceed 14 digits" & Chr(13)
'End If


If Len(Sheet1.Range("IncD.EducationCess").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Education Cess cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


If Len(Sheet1.Range("IncD.GrossTaxLiability").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Total Tax cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If

If Len(Sheet1.Range("IncD.NetTaxLiability").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Balance Tax after Relief cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If
Sheet1.Unprotect Password:=getmsgstate
If Len(Sheet1.Range("IncD.IntrstPayUs234A").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Interest Payable u/s 234 A cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If
Sheet1.Protect Password:=getmsgstate
Sheet1.Unprotect Password:=getmsgstate
If Len(Sheet1.Range("IncD.IntrstPayUs234B").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Interest Payable u/s 234 B cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If
Sheet1.Protect Password:=getmsgstate
Sheet1.Unprotect Password:=getmsgstate
If Len(Sheet1.Range("IncD.IntrstPayUs234C").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Interest Payable u/s 234 C cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If
Sheet1.Protect Password:=getmsgstate

If Len(Sheet1.Range("Fvalue").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Interest Payable u/s 234 F cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If


'If Len(Sheet1.Range("IncD.TotalIntrstPay").value) > 14 Then
'    ValidateMorethan14Digits = False
'    msg14didgits = msg14didgits + "Total Interest u/s 234A,234B,234C,234F cannot exceed 14 digits" & Chr(13)
'End If


If Len(Sheet1.Range("IncD.TotTaxPlusIntrstPay").value) > 14 Then
    ValidateMorethan14Digits = False
    msg14didgits = msg14didgits + "* Total tax and Interest Payable cannot exceed 14 digits in Sheet Income Details" & Chr(13)
End If
End Function


Function ValidateIncOnImport()
On Error Resume Next
Dim sPassword As String
Dim result As Variant
Dim AllfieldsCheck As Variant
'Sheet1.Range("sheet1.PAN").value = UCase(Sheet1.Range("sheet1.PAN").value)

Dim gridRange, chcells As Variant
Dim iCount As Double

AllfieldsCheck = "sheet1.FirstName||sheet1.MiddleName||sheet1.SurNameOrOrgName||sheet1.ResidenceNo||sheet1.ResidenceName||sheet1.RoadOrStreet||sheet1.LocalityOrArea||sheet1.CityOrTownOrDistrict||sheet1.STDcode||sheet1.PhoneNo||sheet1.mobileCountryCode2||sheet1.MobileNoSec||sheet1.ReceiptNo||sheet1.OrigRetFiledDate||sheet1.NoticeNo||sheet1.NoticeDate" '||sheet1.NoticeNumber"

gridRange = Split(AllfieldsCheck, "||")
    For iCount = 0 To UBound(gridRange)
    For Each chcells In Sheet1.Range(gridRange(iCount))
            If Not checkfieldSuperSpecialcharacter(Sheet1.Range(gridRange(iCount))) Then
            fmsgboxStatus ("* The field should not Contain <, >, characters" & Chr(13))
            Sheet1.Range(gridRange(iCount)).Select
            Sheet1.Range(gridRange(iCount)).value = ""
            CloseMsg
            End If
            Next
        Next


'Validation for Date of Birth
Dim HUFReq As Long
    
    If Sheet1.Range("sheet1.status").value = "H - HUF" Then HUFReq = 1 Else HUFReq = 0
        If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
            result = CheckDate(Trim(Range("sheet1.DOB").value), HUFReq)
                    If result = 1 Then
                    'Ayush_10/02
'                        fmsgboxStatus "*Please enter a valid date of birth in dd/mm/yyyy format in Sheet Income Details" & Chr(13) ', vbOKOnly, "Error(s)!"
                        fmsgboxStatus "*""Date of Birth/Formation - Please enter valid date in valid format.""" & Chr(13) ', vbOKOnly, "Error(s)!"
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").value = ""
                        CloseMsg
                    ElseIf result = 2 Then
                    'Ayush_25-26
                        'fmsgboxStatus "Date of birth should not be after 31/03/2022 for A.Y 2022-23 in Sheet Income Details" & Chr(13) ', vbOKOnly, "Error(s)!" dpk1201
'                        fmsgboxStatus "Date of birth should not be after 31/03/2023 for A.Y 2023-24 in Sheet Income Details" & Chr(13) ', vbOKOnly, "Error(s)!"
                        fmsgboxStatus "Date of birth should not be after 31/03/2026 for A.Y 2026-27 in Sheet Income Details" & Chr(13) ', vbOKOnly, "Error(s)!"
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").value = ""
                        CloseMsg
                    End If
        End If


'Date of filling orignal return Validation: On Data Entry

        If Trim(Range("sheet1.OrigRetFiledDate").value) <> "00/00/0000" And Trim(Range("sheet1.OrigRetFiledDate").value) <> "" Then
                If Not FormatNCheckDate(Trim(Range("sheet1.OrigRetFiledDate").value)) Then
                'Ayush_10/02
'                        fmsgboxStatus "* Date Of filing Orignal Return in Sheet Income Details Please enter a valid date in dd/mm/yyyy format." & Chr(13) ', vbOKOnly, "Error(s)!"
                        fmsgboxStatus "* ""Date of filing of original Return - Please enter valid date in valid format.""" & Chr(13) ', vbOKOnly, "Error(s)!"
                        Sheet1.Range("sheet1.OrigRetFiledDate").Select
                        Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                        CloseMsg
                Else
                     
                    'If ChkMaxDate(Trim(Range("sheet1.OrigRetFiledDate").value), "") Then
                    If ChkMaxDate(Trim(Range("sheet1.OrigRetFiledDate").value), "") Then
'                        fmsgboxStatus "* Date Of  filing Orignal Return in Sheet Income Details should not be prior 01/04/2023" & Chr(13) ', vbOKOnly, "Error(s)!"
                             fmsgboxStatus "* Date Of  filing Original Return in Sheet Income Details should not be prior 01/04/2025" & Chr(13) ', vbOKOnly, "Error(s)!"
                        Sheet1.Range("sheet1.OrigRetFiledDate").Select
                        Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                        CloseMsg
                    End If
                    
                    'Malli_27/10/2025
                    If Dformat(Trim(Range("sheet1.OrigRetFiledDate").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                      fmsgboxStatus " ""Date of filing of original return shall not be beyond system date.""" & Chr(13)
                        Sheet1.Range("sheet1.OrigRetFiledDate").Select
                        Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                        CloseMsg
                    End If
                    '----------------
                    
                End If
            End If


 
'Date of filling in response to notice Validation: On Data Entry

                If Trim(Range("sheet1.NoticeDate").value) <> "00/00/0000" And Trim(Range("sheet1.NoticeDate").value) <> "" Then
                    If Not FormatNCheckDate(Range("sheet1.NoticeDate").value) Then
                    'Ayush_10/02
'                        fmsgboxStatus "*Please enter a date in ""Date of notice /order"" field in dd/mm/yyyy format" & Chr(13) ', vbOKOnly, "Error(s)!"
                             fmsgboxStatus "*""Date of such Notice or Order - Please enter valid date in valid format.""" & Chr(13) ', vbOKOnly, "Error(s)!"
                        Sheet1.Range("sheet1.NoticeDate").Select
                        Sheet1.Range("sheet1.NoticeDate").value = ""
                        CloseMsg
                    Else
                    If ChkMaxDate(Range("sheet1.NoticeDate").value, "") Then
                    'Ayush_25-26
                        'fmsgboxsmall "*Date of notice or Order in Sheet Income Details should be after 31/03/2022 for A.Y.2022-23." & Chr(13) ', vbOKOnly, "Error(s)!" dpk1201
'                        fmsgboxStatus "*Date of Notice/Order cannot be prior to 01/04/2024." & Chr(13) ', vbOKOnly, "Error(s)!"
                        fmsgboxStatus "*Date of Notice/Order cannot be prior to 01/04/2025." & Chr(13)
                        Sheet1.Range("sheet1.NoticeDate").Select
                        Sheet1.Range("sheet1.NoticeDate").value = ""
                        CloseMsg
                    End If
                    
                    
                    If Range("sheet1.NoticeDate").Locked = False Then
                       
                       If Range("sheet1.NoticeDate").value = "" Then
                       'Ayush_25-26
'                          fmsgboxsmall "*Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)"" & chr(13)"
                          fmsgboxsmall "*Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)""" & Chr(13)
                          CloseMsg
                       End If
                    
                    End If
                    
                End If
                    
        End If





End Function



Function AssignValues()
On Error Resume Next

firstName = Range("sheet1.FirstName").value
middleName = Range("sheet1.MiddleName").value
LastName = Range("sheet1.SurNameOrOrgName").value
PAN = Range("sheet1.PAN").value
Flat = Range("sheet1.ResidenceNo").value
residenceName = Range("sheet1.ResidenceName").value
status = Range("sheet1.Status").value
roadOrStreet = Range("sheet1.RoadOrStreet").value
Area = Range("sheet1.LocalityOrArea").value
AssesseeDob = Range("sheet1.DOB").value
NatureEmp = Range("Sheet1.NatureofEmployment").value
'Gender = Range("sheet1.Gender1").value
City = Range("sheet1.CityOrTownOrDistrict").value
State = Range("sheet1.StateCode1").value
PrevAYBAC = Range("OptingNewTaxRegime").value
Country = Range("sheet1.CountryCode").value
PinCode = Range("sheet1.PinCode").value
Email = Range("sheet1.EmailAddress").value
Email2 = Range("sheet1.EmailAddress2").value
mobileNo = Range("sheet1.Mobileno").value
CountrycodeMobileNo1 = Range("sheet1.mobileCountryCode").value
stDcode = Range("sheet1.STDcode").value
phoneNo = Range("sheet1.PhoneNo").value
CountrycodeMobileNo2 = Range("sheet1.mobileCountryCode2").value
PhoneNo2 = Range("sheet1.MobileNoSec").value
EmpCategory = Range("sheet1.EmployerCategory1").value
'IncomeTaxWard = Range("sheet1.DesigOfficerWardorCircle").value
AadhaarNumber = Range("sheet1.Aadhaar").value
'Ayush_13/01/2026_V0.2
Flat_Sec = Range("sheet1.ResidenceNo2").value
residenceName_Sec = Range("sheet1.ResidenceName2").value
roadOrStreet_Sec = Range("sheet1.RoadOrStreet2").value
Area_Sec = Range("sheet1.LocalityOrArea2").value
City_Sec = Range("sheet1.CityOrTownOrDistrict2").value
State_Sec = Range("sheet1.StateCode2").value
Country_Sec = Range("sheet1.CountryCode2").value
PinCode_Sec = Range("sheet1.PinCode2").value
'------------------
'Ayush_enrol
'AadhaarEnrolNumber = Range("sheet1.AadhaarEnrol").value
ReturnFilledSection = Range("sheet1.ReturnFileSec1").value

ReturnFilledSection1 = Range("sheet1.ReturnFileSec").value
'ReturnFilledSection2 = Range("sheet1.ReturnFileSec2").value

'GovernedByPortugese = Range("sheet1.PortugeseCC5A").value
PanOfSpouse = Range("sheet1.PANOFSPOUSE").value
'TypeOfReturn = Range("sheet1.ReturnType1").value
RecieptNo = Range("sheet1.ReceiptNo").value
'RecieptNo1 = Range("sheet1.ReceiptNo1").value
DateOfFillingRetrun = Range("sheet1.OrigRetFiledDate").value
'DateOfFillingRetrun1 = Range("sheet1.OrigRetFiledDate1").value
NoticeNo = Range("sheet1.NoticeNo").value
'NoticeNumber = Range("sheet1.NoticeNumber").value
DateOfNotice = Range("sheet1.NoticeDate").value
'ResidentialStatus = Range("sheet1.ResidentialStatus1").value

'TaxStatusFlag = Range("sheet1.TaxStatus").value

IncomeBP = Range("IncD.IncomeFromBP").value

IncomeSAl1 = Range("IncD.IncomeFromSal_i").value
IncomeSAl2 = Range("IncD.IncomeFromSal_ii").value
IncomeSAl3 = Range("IncD.IncomeFromSal_iii").value
IncomeSAl4 = Range("IncD.IncomeFromSal_iv").value
'Ayush_13/01/2026_commented
IncomeSAl7 = Range("IncomeNotified89A").value
'IncomeSAl8 = Range("IncomeNotifiedOther89A").value
IncomeSAl5a = Range("IncD.IncomeFromSal_allow").value



IncomeSAl6 = Range("IncD.IncomeFromSal_va").value 'tot DEd
IncomeSAl5 = Range("IncD.IncomeFromSal_v").value 'other ded
IncomeSAl5b = Range("IncD.IncomeFromSal_iva").value 'ded 16(ia)
IncomeSAl5c = Range("IncD.IncomeFromSal_ivc").value

IncomeSAl = Range("IncD.IncomeFromSal").value
IncomeSAlnet = Range("IncD.IncomeFromSal_net").value

IncomeHP = Range("IncD.IncomeFromHP").value
'IncomeHP1 = Range("IncD.IncomeFromHP_i").value
'IncomeHP1a = Range("IncD.IncomeFromHP_ia").value
'IncomeHP1b = Range("IncD.IncomeFromHP_ib").value
'IncomeHP2 = Range("IncD.IncomeFromHP_ii").value
'IncomeHP3 = Range("IncD.IncomeFromHP_iii").value
'IncomeHP4 = Range("IncD.IncomeFromHP_iv").value
IncomeHP5 = Range("IncD.IncomeFromHP_v").value
'IncomeHP6 = Range("IncD.IncomeFromHP_vi").value

TypeOfHP = Range("IncD.TypeOfHP").value



IncomeOS = Range("IncD.IncomeFromOS").value
'IncomeGross = Range("IncD.GrossTotIncome").value
'Newly updated by Bindu as per DE V3 on 4th Feb 2025
IncomeGross = Range("IncD.GrossTotIncome").value
GrossTotIncomeIncLTCG112A = Range("IncD.GrossTotIncome_New").value    'added by Chetan V3.4 on 07/02/2025


Ded80C = Range("IncD.Section80C").value
Ded80CCC = Range("IncD.Section80CCC").value  '08/05/2025
Ded80CCDSelfEmployee = Range("IncD.Section80CCD_SE").value
Ded80CCDEmployer = Range("IncD.Section80CCD").value
Ded80CCD1B_SE = Range("IncD.Section80CCD1B_SE").value

'Ded80CCG = Range("IncD.Section80CCG").value
Ded80D = Range("IncD.Section80D").value
Ded80DB = Range("IncD.Section80DB").value
Ded80DC = Range("IncD.Section80DC").value
Ded80DD = Range("IncD.Section80DD").value
Ded80DDB = Range("IncD.Section80DDB").value
Ded80E = Range("IncD.Section80E").value
Ded80EE = Range("IncD.Section80EE").value
Ded80EEA = Range("IncD.Section80EEA").value
Ded80EEB = Range("IncD.Section80EEB").value
Ded80G = Range("IncD.Section80G").value
Ded80GG = Range("IncD.Section80GG").value
Ded80GGA = Range("IncD.Section80GGA").value
Ded80GGC = Range("IncD.Section80GGC").value
'Ded80QQB = Range("IncD.Section80QQB").value
'Ded80RRB = Range("IncD.Section80RRB").value
Ded80TTA = Range("IncD.Section80TTA").value
Ded80TTB = Range("IncD.Section80TTB").value
Ded80CCH = Range("IncD.AnyOtherDeductions").value 'PAG_C61 AY 2024-25
Ded80U = Range("IncD.Section80U").value
TotalDed = Range("IncD.TotalChapVIADeductions_Input").value

Ded80C_calc = Range("IncD.Section80C_Calc").value
Ded80CCC_calc = Range("IncD.Section80CCC_Calc").value  'Malli/08/05/2025
Ded80CCDSelfEmployee_calc = Range("IncD.Section80CCD_Calc_SE").value
Ded80CCDEmployer_calc = Range("IncD.Section80CCD_Calc").value
Ded80CCD1B_SE_calc = Range("IncD.Section80CCD1B_Calc_SE").value
'Ded80CCG_calc = Range("IncD.Section80CCG_Calc").value
Ded80D_calc = Range("IncD.Section80D_Calc").value
Ded80DD_calc = Range("IncD.Section80DD_Calc").value
Ded80DDB_calc = Range("IncD.Section80DDB_Calc").value
Ded80E_calc = Range("IncD.Section80E_Calc").value
Ded80EE_calc = Range("IncD.Section80EE_Calc").value
Ded80EEA_calc = Range("IncD.Section80EEA_Calc").value
Ded80EEB_calc = Range("IncD.Section80EEB_Calc").value
Ded80G_calc = Range("IncD.Section80G_Calc").value
Ded80GG_calc = Range("IncD.Section80GG_Calc").value
Ded80GGA_calc = Range("IncD.Section80GGA_Calc").value
Ded80GGC_calc = Range("IncD.Section80GGC_Calc").value
'Ded80QQB_calc = Range("IncD.Section80QQB_Calc").value
'Ded80RRB_calc = Range("IncD.Section80RRB_Calc").value
Ded80TTA_calc = Range("IncD.Section80TTA_Calc").value
Ded80TTB_calc = Range("IncD.Section80TTB_Calc").value
Ded80CCH_calc = Range("IncD.AnyOtherDeductions_Calc").value 'PAG_C61 AY 2024-25
Ded80U_calc = Range("IncD.Section80U_Calc").value
TotalDed_calc = Range("IncD.TotalChapVIADeductions").value

'TotalTaxIncome = Range("IncD.TotalIncome").value
'As per DE V3 updated by Bindu on 4th Feb 2025
TotalTaxIncome = Range("IncD.TotalIncome_New").value

TaxPayableTotalIncome = Range("IncD.TotalTaxPayable").value
Rebate87A = Range("IncD.Rebate87A").value
TaxafterRebate = Range("IncD.TaxPayableOnRebate").value
'Surcharge = Range("IncD.SurchargeOnTaxPayable").value
EducationCess = Range("IncD.EducationCess").value
TotalTaxSurchargeEdu = Range("IncD.GrossTaxLiability").value
Relief89 = Range("IncD.Section89").value
Relief89A = Range("IncD.Section89A").value
BalanceTax = Range("IncD.NetTaxLiability").value
Sheet1.Unprotect getmsgstate
InterestPay234a = Range("IncD.IntrstPayUs234A").value
Sheet1.Unprotect Password:=getmsgstate
InterestPay234b = Range("IncD.IntrstPayUs234B").value
Sheet1.Unprotect Password:=getmsgstate
InterestPay234c = Range("IncD.IntrstPayUs234C").value
Sheet1.Unprotect Password:=getmsgstate
InterestPay234f = Range("IncD.IntrstPayUs234F").value
Sheet1.Protect getmsgstate
'TotalInterest234abc = Range("IncD.TotalIntrstPay").value
TotalTaxInterest = Range("IncD.TotTaxPlusIntrstPay").value

Representativeassesseeflg = Range("sheet1.RepAssessee").value
NameRepAssessee = Range("sheet1.NameRepAssessee").value
'Ayush_13/01_2026_V0.2_commented
'CapacityRep = Range("sheet1.CapacityRep").value
'AddressRepAssessee = Range("sheet1.AddressRepAssessee").value
'PANRepAssessee = Range("sheet1.PANRepAssessee").value
'AadhaarRepAssessee = Range("sheet1.AadhaarRepAssessee").value
'Ayush_12/01/2026_V0.2
EmailRepAssessee = Range("sheet1.EmailRepAssessee").value
CountryCodeRepAssessee = Range("sheet1.CountryCodeRepAssessee").value
MobileRepAssessee = Range("sheet1.MobileRepAssessee").value

'---------------------

End Function

'Final Validation of all values

Function ValidateSheetID() As Boolean
On Error Resume Next
ValidateSheetID = True

subProcCaption = "Validating Sheet General"
noOfProcessSub = 12

Dim result, HUFrequired As Long
Dim DedValueError As Variant
Dim val As Variant
msgError = ""
        
'Change-23.11.2022.103.08
'If Trim(Sheet8.Range("bacValue").value) = "" Or IsEmpty(Sheet8.Range("bacValue").value) Then
If Trim(Sheet8.Range("oldbacValue").value) = "" Or IsEmpty(Sheet8.Range("oldbacValue").value) Then

    IDM_115BACFlag
'    If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
'        msgError = msgError & "* Have you opted  for new tax regime u/s 115BAC and filed Form 10IE in AY 2021-22 ? is mandatory in Sheet : Income Details " & errmsgID
'    End If
End If
    
    
'Change-06.12.2022.103.19

'Validate_CurrentAY23_24 old

 Validate_CurrentAY24_25      'For Table A23
       'Ayush_25-26 For Table A23(a)
  If Not Validate_CurrentAY25_26 Then ValidateSheetID = False
 
' Validate_dateFiling_25_26   'Ayush_25-26 'Commented by sai on 28/01/2025
 
 
 '===============================================================================
 If Not Validate_CurrentAY25_26_A23i Then ValidateSheetID = False
   'For Table A23(a)(i) added by sai on 27/01/2025 for AY 2025-26
   
   
 If Not Validate_CurrentAY25_26_A23Bi Then ValidateSheetID = False 'For Table A23(b)(i) added by sai on 27/01/2025 for AY 2025-26
 If Not Validate_CurrentAY25_26_A23Ci Then ValidateSheetID = False 'For Table A23(c)(i) added by sai on 27/01/2025 for AY 2025-26
'===============================================================================
 If Trim(Range("BAC115.AssessmentYear_2425").value) = "2025-26" Then
    If Trim(Range("BAC115.AssessmentYear_Aii").value) = "2025-26" Then
'       msgError = msgError & "* The assessment year of Form 10IEA filed for re-entering into new tax regime shall be subsequent to assessment year in which first form 10IEA was filed for choosing old tax regime." & Chr(13)
           msgError = msgError & "* The assessment year of Form 10IEA in field A23(A)(ii)(a) filed for re-entering into new tax regime shall be subsequent to assessment year in which first form 10IEA was filed for choosing old tax regime." & Chr(13)

    End If
End If
    
    


' 'Ayush_Legacy26
'    If (Sheet1.Range("Sheet1.HRA").value > 0 And Sheet1.Range("IncD.Section80GG_Calc").value > 0) Then
'
'        msgError = msgError & "* ""Deduction u/s 10(13A) & 80GG cannot be claimed for the same period""" & Chr(13)
'    End If
 
 
' If Len(Sheet1.Range("IncD.IncomeFromSal_i").value) > 14 Then
'        msgError = msgError & "* Gross total income cannot exceed 14 digits" & Chr(13)
'       ' Validate_80E = False
'    End If
 

'    If Sheet8.Range("bacValue").value = 1 Or (Sheet8.Range("bacValue").value = 2 And Sheet8.Range("oldbacValue").value = 1) Then
'         If Trim(Range("Sheet1.date10IE").value) = "" Or IsEmpty(Range("Sheet1.date10IE").value) Then
'           msgError = msgError & "* Please furnish form 10-IE to opt for new tax regime " & Chr(13)
'         End If
'          If Trim(Range("Sheet1.acknowledgementNumber").value) = "" Or IsEmpty(Range("Sheet1.acknowledgementNumber").value) Then
'            msgError = msgError & "* Acknowledgement Number of Form 10IE is mandatory. " & Chr(13)
'         End If
'    End If
    
    If Trim(Range("IncD_dq1div").value) = "" Or IsEmpty(Range("IncD_dq1div").value) Then
       msgError = msgError & "* Upto 15/6 (i) is mandatory in Sheet : Income Details " & Chr(13)
    End If
    
    If Trim(Range("IncD_dq2div").value) = "" Or IsEmpty(Range("IncD_dq2div").value) Then
        msgError = msgError & "* From 16/6 to 15/9 (ii) is mandatory in Sheet : Income Details " & Chr(13)
    End If
    
    If Trim(Range("IncD_dq3div").value) = "" Or IsEmpty(Range("IncD_dq3div").value) Then
        msgError = msgError & "* From 16/9 to 15/12 (iii) is mandatory in Sheet : Income Details " & Chr(13)
    End If
    
   If Trim(Range("IncD_dq4div").value) = "" Or IsEmpty(Range("IncD_dq4div").value) Then
       msgError = msgError & "* From 16/12 to 15/3 (iv) is mandatory in Sheet : Income Details " & Chr(13)
    End If
    
    If Trim(Range("IncD_dq5div").value) = "" Or IsEmpty(Range("IncD_dq5div").value) Then
        msgError = msgError & "* From 16/3 to 31/3 (v) is mandatory in Sheet : Income Details " & Chr(13)
    End If
        
'    If Trim(Range("OSIncreliefus89A").value) > Trim(Range("OSIncomeNotified89A").value) Then
'        msgError = msgError & "* Relied u/s 89A cannot be claimed more than income offered in pension accrued in a pension fund maintained in a notified country u/s 89A" & Chr(13)
'        Sheet1.Range("OSIncreliefus89A").value = 0
'    End If

  
  If (Range("IncD_dq1div").value + Range("IncD_dq2div").value + Range("IncD_dq3div").value + Range("IncD_dq4div").value + Range("IncD_dq5div").value) = 0 Then
  
    If Range("IncD_dqdiv").value = "" Then
       Range("IncD_dqdiv").value = 0
    End If
   
  End If
     
     If Range("IncD_dqdiv").value <> (Range("IncD_dq1div").value + Range("IncD_dq2div").value + Range("IncD_dq3div").value + Range("IncD_dq4div").value + Range("IncD_dq5div").value) Then
       ' fmsgboxsmall "* Quarterly breakup of dividend income shall match with total dividend income " ', vbOKOnly, "Error(s)!"
        msgError = msgError & "*Total of this field should be equal to amount entered in ""Quarterly breakups of Dividend income""" & Chr(13)
     End If

   


'Ayush_13/01/2026_V0.2_commented
'    If Range("Increliefus89A").value > Range("IncomeNotified89A").value Then
'        msgError = msgError & "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A" & Chr(13)
'        Sheet1.Range("Increliefus89A").value = 0
'    End If
    
    
'    If Sheet1.Range("IncD_qOS1").value <> Sheet1.Range("OSIncomeNotified89ATot").value Then
'    fmsgbox "* Total of this field should be equal to amount entered in Income from retirement benefit account maintained in a notified country u/s 89A -  Income claimed for relief from taxation u/s 89A"
'    ValidateSheetID = False
'    CloseMsg
'    End If
        
        
        
        
        
If Not CheckName(firstName, 25) Then msgError = msgError & "* First Name in Sheet Income Details " & errmsgID
If Not CheckName(middleName, 25) Then msgError = msgError & "* Middle Name in Sheet  Income Details " & errmsgID
If Not CheckName(LastName, 75, True) Then msgError = msgError & "* Last Name " & errmsgID & " in Sheet Income Details" & Chr(13)

If PAN = "" Then
    msgError = msgError & "* PAN is Mandatory in Sheet Income Details" & Chr(13)
Else
    If Not CheckPAN(PAN) Then
        msgError = msgError & "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Sheet Income Details " & Chr(13)
        
    End If
End If
UpdateProgressBar

If Not CheckName(Flat, 50, True) Then msgError = msgError & "* Flat/ Door/ Block No " & errmsgID & " under primary address in Sheet Income Details" & Chr(13)
If Not CheckName(residenceName, 50) Then msgError = msgError & "* Name of Premises / Building / Village under primary address in Sheet Income Details " & errmsgID

If UCase(status) = "(SELECT)" Or status = "" Then
    msgError = msgError & "* Status is Mandatory in Sheet Income Details " & Chr(13)
End If

UpdateProgressBar

If Not CheckName(roadOrStreet, 50) Then msgError = msgError & "* Road/Street/Post Office " & errmsgID & " under primary address in Sheet Income Details" & Chr(13)
If Not CheckName(Area, 50, True) Then msgError = msgError & "* Area/Locality " & errmsgID & " under primary address in Sheet Income Details " & Chr(13)


If AssesseeDob = "" Then
'    msgError = msgError & "* Date Of Birth is Mandatory in Sheet Income Details " & Chr(13)

'PAG_C1 AY 2024-25 Changes
    msgError = msgError & "* Please enter Date of Birth / Formation " & Chr(13)
Else
    If status = "H - HUF" Then
        HUFrequired = 1
    Else
        HUFrequired = 0
    End If
        result = CheckDate(AssesseeDob, HUFrequired)
        'Ayush_28/02
        
'    If result = 1 Then msgError = msgError & "* Please enter a valid date in dd/mm/yyyy format in Sheet Income Details" & Chr(13)
If result = 1 Then msgError = msgError & "* ""Date of Birth/Formation - Please enter valid date in valid format.""" & Chr(13)
    'If result = 2 Then msgError = msgError & "* Date should not be after 31/03/2022 for A.Y 2022-23 in Sheet Income Details" & Chr(13) dpk1201
    If result = 2 Then msgError = msgError & "* Date should not be after 31/03/2023 for A.Y 2023-24 in Sheet Income Details" & Chr(13)
End If
    
    UpdateProgressBar
    
'If Gender = "" Or Gender = "(Select)" Then msgError = msgError & "Gender is Mandatory in Sheet  Income Details" & Chr(13)
If Not CheckName(City, 50, True) Then msgError = msgError & "* Town/City/District" & " " & errmsgID & " under primary address in Sheet Income Details" & Chr(13)
If State = "" Or UCase(State) = "(SELECT)" Then msgError = msgError & "* State is Mandatory under primary address in Sheet Income Details" & Chr(13)
'If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" And (PrevAYBAC = "" Or UCase(PrevAYBAC) = "(SELECT)") Then msgError = msgError & "* Please Select the applicable option" & Chr(13)
If Country = "" Or UCase(Country) = "(SELECT)" Then msgError = msgError & "* Country is Mandatory under primary address in Sheet Income Details" & Chr(13)
'If PinCode = "" Then
'    msgError = msgError & "Pincode is Mandatory in Sheet : Income Details" & Chr(13)
'Else
'    If (UCase(Country) <> "91-INDIA" And PinCode <> "999999") Or (UCase(Country) = "91-INDIA" And PinCode = "999999") Then msgError = msgError & "Pincode is Not valid in Sheet : Income Details" & Chr(13)
'End If



If Not ChkPincode Then ValidateSheetID = False
If Not ChkZipcode Then ValidateSheetID = False

'Ayush_13/01/2026_V0.2
If Sheet1.Range("Secondary_Address").value = "No" Then
'If Flat_Sec <> "" Or Area_Sec <> "" Or City_Sec <> "" Or State_Sec <> "" Or Country_Sec <> "" Or PinCode_Sec <> "" Then
'If Flat_Sec <> "" Or Area_Sec <> "" Or City_Sec <> "" Or (State_Sec <> "" And State_Sec <> "(Select)") Or (Country_Sec <> "" And Country_Sec <> "(Select)") Or PinCode_Sec <> "" Or zipCode_Sec <> "" Then
If Not CheckName(Flat_Sec, 50, True) Then msgError = msgError & "* Flat/ Door/ Block No " & errmsgID & " under secondary address in Sheet Income Details" & Chr(13)
If Not CheckName(residenceName_Sec, 50) Then msgError = msgError & "* Name of Premises / Building / Village under secondary address in Sheet Income Details " & errmsgID


If Not CheckName(roadOrStreet_Sec, 50) Then msgError = msgError & "* Road/Street/Post Office " & errmsgID & " under secondary address in Sheet Income Details" & Chr(13)
If Not CheckName(Area_Sec, 50, True) Then msgError = msgError & "* Area/Locality " & errmsgID & " under secondary address in Sheet Income Details " & Chr(13)

    If Not CheckName(City_Sec, 50, True) Then msgError = msgError & "* Town/City/District " & " " & errmsgID & " under secondary address in Sheet Income Details" & Chr(13)
    If State_Sec = "" Or UCase(State_Sec) = "(SELECT)" Then msgError = msgError & "* State is Mandatory under secondary address in Sheet Income Details" & Chr(13)

    If Country_Sec = "" Or UCase(Country_Sec) = "(SELECT)" Then msgError = msgError & "* Country is Mandatory under secondary address in Sheet Income Details" & Chr(13)

    If Sheet1.Range("sheet1.PinCode2").Locked = False Then
        If Not ChkPincode_Sec Then ValidateSheetID = False
    End If
    
    If Sheet1.Range("sheet1.ZipCode2").Locked = False Then
        If Not ChkZipcode_Sec Then ValidateSheetID = False
    End If
End If
'---------------------

'If Not ValidateSelect80D Then ValidateSheetID = False

UpdateProgressBar

If Email = "" Then
'Ayush_26-27_V0.3
    msgError = msgError & "* ""Primary Email ID of the taxpayer is mandatory in tab: Part A General Information.""" & Chr(13)
Else
    If Not CheckEmailAddress(Email, validateType) Then
        If validateType = "1" Then msgError = msgError & "* Primary Email ID of the taxpayer contains special characters in Sheet Income Details" & Chr(13)
        If validateType = "2" Then msgError = msgError & "* ""Primary Email ID of the taxpayer cannot be more than 125 characters""" & Chr(13)
        If validateType = "3" Then msgError = msgError & "* ""Invalid Primary Email ID of the taxpayer Please retry.""" & Chr(13)
        'Ayush_26-27_V0.3
'        If validateType = "4" Then msgError = msgError & "* Primary Email ID of the taxpayer is mandatory in Sheet Income Details" & Chr(13)
        If validateType = "4" Then msgError = msgError & "* ""Primary Email ID of the taxpayer is mandatory in tab: Part A General Information.""" & Chr(13)
    End If
End If

UpdateProgressBar
    
If mobileNo = "" Then
'    msgError = msgError & "* Mobile number 1 is Mandatory in Sheet  Income Details" & Chr(13)
    msgError = msgError & "* ""Primary Mobile No. in Sheet: Personal Information is mandatory.""" & Chr(13)
Else
    If Not CheckMobileNo(mobileNo) Then msgError = msgError & "* Please enter a valid 10 digit Primary Mobile No. in Sheet  Income Details" & Chr(13)
End If

If CountrycodeMobileNo1 = "91" Then
If Len(mobileNo) <> 10 Then
    msgError = msgError & "* Please enter a valid 10 digit Primary Mobile No. in sheet Income Details" & Chr(13)
End If
Else
If Len(mobileNo) < 5 Or Len(mobileNo) > 10 Then
    msgError = msgError & "* Please enter a valid 10 digit Primary Mobile No. in sheet Income Details" & Chr(13)
End If
End If
'Ayush_01
If (Len(phoneNo) + Len(stDcode)) > 12 Then
 msgError = msgError & "* Invalid Phone Number. STD Code + Landline Number should be upto 12 digits and cannot begin with '0'." & Chr(13)
End If

'sudha_2025-26

'     If Range("IncD.TotalChapVIADeductions_Input") > 2000000 Then
'Modified by sai on 24/04/2024

'    If (Sheet1.Range("IncD.TotalChapVIADeductions").value + Sheet1.Range("IncD.IncomeFromHP_v").value + Sheet1.Range("HRA").value) > 2000000 Then
'
'     msgError = msgError & "*If claim of deductions under chapter VIA, HRA u/s (Sec 10(13A) and Interest payable on borrowed capital are more than Rs.20 lakhs, you are required to file return in ITR 3 / ITR 5, as the case may be. Please refer Rule 12 for further details." & Chr(13)
'
'    End If

'Restricted from front end itself
'     If Range("IncD.TotalChapVIADeductions") > 2000000 Then
'
'     msgError = msgError & "*If claim of deductions under chapter VIA is more than Rs.20 lakhs, you are required to file   return in ITR 3 / ITR 5, as the case may be. Please refer Rule 12 for further details." & Chr(13)
'
'    End If

'Ayush_commented_26-27_changes
''Ayush_16/06/2026
'If Sheet1.Range("IncD.Section80CCD_SE").value > 0 Or Sheet1.Range("IncD.Section80CCD1B_SE").value > 0 Then
'If Sheet1.Range("pran_new").value = "" Then
'msgError = msgError & "* ""PRAN no. shall be mandatory if deduction u/s 80CCD(1) or 80CCD(1B) is claimed"""" & Chr(13)"
'
'End If
'End If

'Ayush_SIT-109605
If Sheet1.Range("Secondary_Address").value = "(Select)" Or Sheet1.Range("Secondary_Address").value = "" Then
msgError = msgError & "* Is the secondary address same as primary address? is mandatory." & Chr(13)
End If



'Ayush_26-27_C155
If Sheet1.Range("IncD.Section80GG").value > 0 Then
If Sheet1.Range("Sheet1.AckNum").value = "" Then
'msgError = msgError & "* Please enter Acknowledgement number of Form 10BA filed for AY 2025-26 in VIA deductions" & Chr(13)
msgError = msgError & "* Please enter Acknowledgement number of Form 10BA filed for AY 2026-27 in VIA deductions" & Chr(13)

End If
End If

'Ayush_23/05
If Sheet1.Range("Sheet1.AckNum").value <> "" Then
If Sheet8.Range("bacValue").value = 2 Then
If Sheet1.Range("IncD.Section80GG").value = "" Or Sheet1.Range("IncD.Section80GG").value = 0 Then

msgError = msgError & "* 80GG amount shall be more than 0 when acknowledgement no. of form 10BA is provided" & Chr(13)

End If
End If
End If







If Sheet1.Range("IncD.Section80GG").value > 0 Then
If Sheet1.Range("Sheet1.AckNum").value <> "" Then
If Len(Sheet1.Range("Sheet1.AckNum").value) <> 15 Then
msgError = msgError & "* ""Please enter valid Acknowledgement Number""" & Chr(13)
End If
End If
End If


'If Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
'            If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) > "2026-08-31" Then
'                msgError = msgError & """Selected Filing section 139(1) cannot be selected post 31st August, 2026""" & Chr(13)
'
'            End If
'        End If

If Range("sheet1.ReturnFileSec").value = "139(5)-Revised Return" Then
            If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) > "2027-03-31" Then
                msgError = msgError & """Selected Filing section 139(5) cannot be selected post 31st March, 2027""" & Chr(13)
                
            End If
        End If


'Ayush_16/06/2026
'If Sheet1.Range("IncD.Section80CCD1B_SE").value > 0 Or Sheet1.Range("IncD.Section80CCD_SE").value > 0 Then
'If Not checkspecialcharacterForPinCode(Sheet1.Range("pran_new").value) Then
'msgError = msgError & "* ""Please enter a valid 12 digit PRAN""" & Chr(13)
'End If
'End If

'Ayush_commented_26-27_changes
''Sudha 2025-26
'If Sheet1.Range("IncD.Section80CCD1B_SE").value > 0 Or Sheet1.Range("IncD.Section80CCD_SE").value > 0 Or Sheet1.Range("IncD.Section80CCC").value > 0 Then
'
'   ' If Not checkspecialcharacterForPinCode(Sheet1.Range("Pran_Number").value) Then
'    If Len(Sheet1.Range("Pran_Number").value) > 125 Then
'      msgError = msgError & "* ""Please enter valid Pran Number""" & Chr(13)
'    End If
'
'End If


'============================================


If CountrycodeMobileNo1 = "" Then
'Ayush_26-27_V0.3
'    msgError = msgError & "* Country code for Primary Mobile No. is Mandatory in Sheet Income Details" & Chr(13)
    msgError = msgError & "*  ""Country Code in Primary Mobile No. is mandatory .""" & Chr(13)
End If

If CountrycodeMobileNo1 = "0" Or CountrycodeMobileNo1 = "00" Or CountrycodeMobileNo1 = "000" Or CountrycodeMobileNo1 = "0000" Or CountrycodeMobileNo1 = "00000" Then
     msgError = msgError & "* Please enter valid Country code for Primary Mobile No. in Sheet Income Details " & Chr(13)
End If

UpdateProgressBar

If stDcode <> "" Then
    val = ValidateSTDcode(stDcode)
    If val <> 0 Then msgError = msgError & "* STD/ISD Code " & val & " in Sheet  Income Details" & Chr(13)
Else
        If phoneNo <> "" Then msgError = msgError & "* Please enter STD/ISD code in Sheet Income Details" & Chr(13)
End If

If phoneNo <> "" Then
    val = ValidatePhoneNo(stDcode, phoneNo)
    If val = "should be 12 digits" Then msgError = msgError & "* STD/ISD code and Phone no " & val & " in Sheet  Income Details" & Chr(13)
'Else
'    If stDcode <> "" Then msgError = msgError & "* Please enter  Phone number in Sheet  Income Details" & Chr(13)
End If

UpdateProgressBar
   
If PhoneNo2 <> "" Then
    If Not CheckMobileNo(PhoneNo2) Then msgError = msgError & "* Please enter a valid 10 digit Secondary Mobile No. in Sheet Income Details" & Chr(13)
End If

If PhoneNo2 <> "" Then
If CountrycodeMobileNo2 = "" Then
'    msgError = msgError & "* Please enter Country Code for Mobile number 2 in Sheet Income Details " & Chr(13)
    
'    msgError = msgError & "*""Country Code in Secondary Mobile No. in Sheet: Personal Information is mandatory.""" & Chr(13)
    msgError = msgError & "* ""Country Code in Secondary Mobile No. in Sheet: Personal Information is mandatory.""" & Chr(13)
    
End If
End If

If CountrycodeMobileNo2 <> "" Then
If PhoneNo2 = "" Then
    msgError = msgError & "* Secondary Mobile No. is Mandatory  in Sheet Income Details " & Chr(13)
End If
End If



If CountrycodeMobileNo2 = "91" Then

If Len(PhoneNo2) <> 10 Then
    msgError = msgError & "* Please enter a valid 10 digit Secondary Mobile No. in sheet Income Details" & Chr(13)
End If



Else
If PhoneNo2 <> "" Then
If Len(PhoneNo2) < 5 Or Len(PhoneNo2) > 10 Then
    msgError = msgError & "* Please enter valid Secondary Mobile No. in sheet Income Details" & Chr(13)
End If
End If
End If


If NatureEmp = "" Or NatureEmp = "(Select)" Then
     msgError = msgError & "*Nature of Employment is mandatory" & Chr(13)
    'msgError = msgError & "* Nature of Employment is mandatory in Sheet Income details" & Chr(13)
End If


If Sheet1.Range("Sheet1.HRA").value > Application.WorksheetFunction.RoundUp(Sheet1.Range("IncD.IncomeFromSal_ii").value / 3, 0) Then

msgError = msgError & "*Exemption of HRA u/s 10(13A) claimed cannot be more than limit prescribed " & Chr(13)

End If


If Representativeassesseeflg = "" Or Representativeassesseeflg = "(Select)" Then
msgError = msgError & "* Please select whether return is filed by Representative assessee in Sheet Income details" & Chr(13)
End If

'Ayush_12/01/2026_V0.2
'If Sheet1.Range("sheet1.RepAssessee").value = "Yes" And Sheet5.Range("Ver.Capacity").value <> "Representative" Then
'msgError = msgError & "* Details of representative assessee is mandatory as the return is being filed by representative in Sheet Tax Paid and Verification" & Chr(13)
'End If

If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
    If NameRepAssessee = "" Then
        msgError = msgError & "* Please enter the Name of representative assessee in Sheet Income details" & Chr(13)
    End If
End If


If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
    If EmailRepAssessee = "" Then
        msgError = msgError & "* ""Email Address of representative assessee is mandatory""." & Chr(13)
    End If
End If

If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
    If CountryCodeRepAssessee = "" Then
        msgError = msgError & "* ""Country Code in contact No. of representative assessee in Sheet: Personal Information is mandatory.""" & Chr(13)
    End If
End If

If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
    If MobileRepAssessee = "" Then
        msgError = msgError & "* ""Mobile No. in contact No. of representative assessee in Sheet: Personal Information is mandatory.""" & Chr(13)
    End If
End If





'-------------------------
''Ayush_12/01/2026_V0.2_commented
'If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
'    If isdropdownblank(CapacityRep) Then
'        msgError = msgError & "* Please select the capacity of representative assessee in Sheet Income details" & Chr(13)
'    End If
'End If
'
'If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
'    If AddressRepAssessee = "" Then
'        msgError = msgError & "* Please enter the Address of representative assessee in Sheet Income details" & Chr(13)
'    End If
'End If
'
'If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
'    If PANRepAssessee = "" Then
'        msgError = msgError & "* Please enter the PAN of the representative assessee in Sheet Income details" & Chr(13)
'    End If
'
'    If PANRepAssessee = PAN Then
'        msgError = msgError & "* PAN of the representative cannot be same as assessee PAN in Sheet Income details" & Chr(13)
'    End If
'
'    If AadhaarRepAssessee <> "" Then
'    If AadhaarRepAssessee = AadhaarNumber Then
'        msgError = msgError & "* Aadhaar no. entered in Part A -Gen and Aadhaar no. of the representative cannot be same in Sheet Income details" & Chr(13)
'    End If
'    End If
'
'
'End If
'
'
'If Not CheckPAN(PANRepAssessee) Then
''Ayush_PAN
'   ' msgError = msgError & "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet and Fourth letter should be ""P" & Chr(13)
'    msgError = msgError & "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet & Chr(13)"
'End If

'---------------------------------------------------

If CountrycodeMobileNo2 = "0" Or CountrycodeMobileNo2 = "00" Or CountrycodeMobileNo2 = "000" Or CountrycodeMobileNo2 = "0000" Or CountrycodeMobileNo2 = "00000" Then
     msgError = msgError & "* Please enter the valid Country code for Mobile number 2 in Sheet Income Details " & Chr(13)
End If

UpdateProgressBar

If Not FillingStatusValidations() Then msgError = msgError & errmsgfiling
If Mid(Range("sheet1.Status").value, 1, 1) <> "F" Then
'    If Not MandatoryList(GovernedByPortugese) Then msgError = msgError & "Please select Whether Person governed by Portuguese Civil Code under Sec 5A in Sheet : Income Details" & Chr(13)
End If
'If Not MandatoryList(TypeOfReturn) Then msgError = msgError & "Filing Type is Mandatory in Sheet : Income Details" & Chr(13) Else TypeOfReturn = Mid(TypeOfReturn, 1, 3)
'If MandatoryList(GovernedByPortugese) Then
'    If Mid(GovernedByPortugese, 1, 1) <> "N" Then
'        If PanOfSpouse = "" Then msgError = msgError & "PAN of Spouse is Mandatory in Sheet : Income Details" & Chr(13)
'        If PanOfSpouse <> "" And Mid(PanOfSpouse, 4, 1) <> "P" Or Not CheckPAN(PanOfSpouse) Then msgError = msgError & "PAN of Spouse is invalid in Sheet : Income Details" & Chr(13)
'    Else
'        PanOfSpouse = ""
'    End If
'End If

UpdateProgressBar
     
If Not ValidateAadhaar Then msgError = msgError & "* Please enter the valid Aadhaar number" & errmsgID & " in Sheet Income Details" & Chr(13)
'If Not MandatoryList(TaxStatusFlag) Then msgError = msgError & "Tax Status is Mandatory in Sheet : Income Details" & Chr(13)

If (RecieptNo <> "") Then
   If Not CheckNoticeRecieptNo(RecieptNo) Then msgError = msgError & "* Invalid Receipt Number. Receipt number should be 15 digits in Sheet Income Details" & Chr(13)
End If

'If (RecieptNo1 <> "") Then
'   If Not CheckNoticeRecieptNo(RecieptNo1) Then msgError = msgError & "Reciept No should be 15 digits in Sheet : Income Details" & Chr(13)
'End If
    
    
    
    If Mid(Range("sheet1.Status").value, 1, 1) <> "F" Then
    If Mid(Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
    Else
        If Not ChkSeventhProvisoFlag Then msgError = msgError + "*Selection of either of the options for 'Are you filing return of income under Seventh proviso to section 139(1) but otherwise not required to furnish return of income?' is mandatory in sheet Income Details" & Chr(13)
    
         ChkDepositAmountFlag
    
         ChkAggrigateAmountFlag

         ChkAggrigateAmountFlag1
       
    End If
End If

 
If Mid(Sheet1.Range("sheet1.SeventhProvisoFlag").value, 1, 1) <> "N" Then
    If Mid(Sheet1.Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y" Then
    
    If Sheet1.Range("Sheet1.DepositAmountFlag").value <> "Yes" And Sheet1.Range("Sheet1.AggrigateAmountFlag").value <> "Yes" And Sheet1.Range("Sheet1.AggrigateAmountFlag1").value <> "Yes" And Sheet1.Range("clauseiv7provisio139iFlg").value <> "Yes" Then
        msgError = msgError + "* You have selected ""Yes"" for ""Are you filing return of income under Seventh proviso to section 139(1) but otherwise not required to furnish return of income? "", hence please enter respective amount in Sheet Income details." & Chr(13)
    End If
    If Mid(Sheet1.Range("Sheet1.DepositAmountFlag").value, 1, 1) = "Y" Then
    If Range("Sheet1.DepositAmount").value < 10000000 Or Range("Sheet1.DepositAmount").value = 10000000 Then
    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 1 Crore" & Chr(13)
    End If
    End If
    
     If Trim(DepositAmountFlag) = "Yes" And DepositAmount = "" Then
        msgError = msgError + "*you have selected ""Yes"" for ""Have you deposited amount or aggregate of amounts exceeding Rs. 1 Crore in one or more current account during the previous year?"", hence please enter the amount in Sheet Income details." & Chr(13)
    End If
    
    If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "Y" Then
    If Range("Sheet1.AggrigateAmount").value < 200000 Or Range("Sheet1.AggrigateAmount").value = 200000 Then
        msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 2 Lakhs" & Chr(13)
    End If
    End If
    
    If Trim(AggrigateAmountFlag) = "Yes" And AggrigateAmount = "" Then
        msgError = msgError + "* you have selected ""Yes"" for ""Have you incurred expenditure of an amount or aggregate of amount exceeding Rs. 2 lakhs for travel to a foreign country for yourself or for any other person?"", hence please enter the amount in Sheet Income details." & Chr(13)
    End If
    
    If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "Y" Then
    If Range("Sheet1.AggrigateAmount1").value < 100000 Or Range("Sheet1.AggrigateAmount1").value = 100000 Then
    msgError = msgError + "*Amount entered cannot be less than or equal to Rs. 1 Lakh" & Chr(13)
    End If
    End If
    
    If Trim(AggrigateAmountFlag1) = "Yes" And AggrigateAmount1 = "" Then
        msgError = msgError + "*you have selected ""Yes"" for ""Have you incurred expenditure of amount or aggregate of amount exceeding Rs. 1 lakh on consumption of electricity during the previous year?', hence please enter the amount in Sheet Income details." & Chr(13)
    End If
    
    
End If

End If

'dpk
If Sheet1.Range("clauseiv7provisio139iFlg").value <> "No" And Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
    If Trim(Sheet1.Range("clauseiv7provisio139iFlg").value) = "(Select)" Then
        msgError = msgError + "Selection of either of the options for 'Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)' is mandatory in Sheet Income Details" & Chr(13)
    End If
    
    If Trim(Sheet1.Range("clauseiv7provisio139iFlg").value) = "Yes" Then
    
    If Sheet1.Range("clauseiv7provisio139iFlg_1").value <> "Yes" And Sheet1.Range("clauseiv7provisio139iFlg_2").value <> "Yes" And Sheet1.Range("clauseiv7provisio139iFlg_3").value <> "Yes" And Sheet1.Range("clauseiv7provisio139iFlg_4").value <> "Yes" Then
        msgError = msgError + "* Please select atleeast one of the conditions prescribed under clause (iv) of seventh proviso to section 139(1)" & Chr(13)
        
    End If
    If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_1").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
    msgError = msgError + "*Selection of either of the options for 'total sales, turnover or gross receipts, as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or ' is mandatory in Sheet Income Details" & Chr(13)
    End If

    If (Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_1").value = "") Then
'    msgError = msgError + "*Please enter amount of Total sales, turnover or gross receipts etc in Income Details"
        msgError = msgError + "*Please select 'No' if total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year" & Chr(13)
    End If
    
        If Sheet1.Range("clauseiv7provisio139iFlg_2").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_2").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
    msgError = msgError + "*Selection of either of the options for 'the total gross receipts of the person in profession exceeds ten lakh rupees during the previous year; or' is mandatory in Sheet Income Details" & Chr(13)
    
    End If
'
    If (Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_2").value = "") Then
'    msgError = msgError + "*Please enter amount of total gross receipts in Income Details"
    msgError = msgError + "*Please select 'No' if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year" & Chr(13)
    End If
    
    If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_3").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
    msgError = msgError + "*Selection of either of the options for 'the aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is twenty-five thousand rupees or more(or fifty-thousand for senior citizen) ; or' is mandatory in Sheet Income Details" & Chr(13)
    
    End If


    If (Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
'    msgError = msgError + "*Please enter amount of tax deducted at source and tax collected at source in Income Details"
    msgError = msgError + "*Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(13)
    End If
'
    If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_4").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
    msgError = msgError + "*Selection of either of the options for 'if his total deposits in a savings bank account is fifty lakh rupees or more, in the previous year.' is mandatory in Sheet Income Details" & Chr(13)
    
    End If


    If (Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_4").value = "") Then
'    msgError = msgError + "*Please enter amount of total deposits in Income Details"
    msgError = msgError + "*Please select 'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(13)
    End If
    
    End If
End If
'
'If Mid(Sheet1.Range("clauseiv7provisio139iFlg").value, 1, 1) <> "N" And Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
'
''    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "Y" And Range("clauseiv7provisio139iAmount_1").value <> "" Then
''    If Range("clauseiv7provisio139iAmount_1").value < 6000000 Or Range("clauseiv7provisio139iAmount_1").value = 6000000 Then
''    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 6000000 " & Chr(13)
''    End If
''    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "Yes" And Range("clauseiv7provisio139iAmount_1").value = "" Then
'        msgError = msgError + "*Please select 'No' if total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year" & Chr(13)
'    End If
'
''    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "Y" Then
''    If Range("clauseiv7provisio139iAmount_2").value < 1000000 Or Range("clauseiv7provisio139iAmount_2").value = 1000000 Then
''    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 1000000 " & Chr(13)
''    End If
''    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_2").value = "Yes" And Range("clauseiv7provisio139iAmount_2").value = "" Then
'        msgError = msgError + "*Please select 'No' if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year" & Chr(13)
'    End If
'
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y" Then
''    If Range("clauseiv7provisio139iAmount_3").value < 25000 Then
''    msgError = msgError + "* Amount entered cannot be less than Rs. 25000 " & Chr(13)
''    End If
'       age = mdCommon.calculateAge(Sheet1.Range("sheet1.DOB").value)
'       If age >= 60 Then
'            If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 50000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
'            msgError = msgError + "* Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)
'            End If
'       End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "Yes" And Range("clauseiv7provisio139iAmount_3").value = "" Then
'        msgError = msgError + "*Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(13)
'    End If
'
''    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "Y" Then
''    If Range("clauseiv7provisio139iAmount_4").value < 5000000 Then
''    msgError = msgError + "* Amount entered cannot be less than Rs. 5000000 " & Chr(13)
''    End If
''    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" And Range("clauseiv7provisio139iAmount_4").value = "" Then
'        msgError = msgError + "*Please select 'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(13)
'    End If
'End If

'dpk


'If Mid(Sheet1.Range("clauseiv7provisio139iFlg").value, 1, 1) <> "N" Then
'
'If Mid(Sheet1.Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "Y" Then
'MsgBox "d3"
'    If Range("clauseiv7provisio139iAmount_1").value < 6000000 Or Range("clauseiv7provisio139iAmount_1").value = 6000000 Then
'    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 6000000 " & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "Yes" And Range("clauseiv7provisio139iAmount_1").value = "" Then
'        msgError = msgError + "*you have selected ""Yes"" for ""total sales, turnover or gross receipts as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or "" hence please enter the amount in Sheet Income details." & Chr(13)
'    End If
'
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "Y" Then
'    If Range("clauseiv7provisio139iAmount_2").value < 1000000 Or Range("clauseiv7provisio139iAmount_2").value = 1000000 Then
'    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 1000000 " & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_2").value = "Yes" And Range("clauseiv7provisio139iAmount_2").value = "" Then
'        msgError = msgError + "*you have selected ""Yes"" for ""the total gross receipts of the person in profession exceeds ten lakh rupees during the previous year; or"", hence please enter the amount in Sheet Income details." & Chr(13)
'    End If
'
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y" Then
'    If Range("clauseiv7provisio139iAmount_3").value < 25000 Or Range("clauseiv7provisio139iAmount_3").value = 25000 Then
'    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 25000 " & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_3").value = "Yes" And Range("clauseiv7provisio139iAmount_3").value = "" Then
'        msgError = msgError + "*you have selected ""Yes"" for ""the aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is twenty-five thousand rupees or more; or"", hence please enter the amount in Sheet Income details." & Chr(13)
'    End If
'
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "Y" Then
'    If Range("clauseiv7provisio139iAmount_4").value < 5000000 Or Range("clauseiv7provisio139iAmount_4").value = 5000000 Then
'    msgError = msgError + "* Amount entered cannot be less than or equal to Rs. 5000000 " & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_4").value = "Yes" And Range("clauseiv7provisio139iAmount_4").value = "" Then
'        msgError = msgError + "*you have selected ""Yes"" for ""if his total deposits in a savings bank account is fifty lakh rupees or more, in the previous year."", hence please enter the amount in Sheet Income details." & Chr(13)
'    End If
'
'
'End If


If (DateOfFillingRetrun <> "") Then
    If Not FormatNCheckDate(DateOfFillingRetrun) Then
    'Ayush_28/02
'        msgError = msgError & "* Please enter date in dd/mm/yyyy format in Sheet Income Details" & Chr(13)
        msgError = msgError & "* ""Date of filing of original Return - Please enter valid date in valid format.""" & Chr(13)
    Else
        'If ChkMaxDate(DateOfFillingRetrun, "") Then msgError = msgError & "* Date of filing of original Return cannot be prior to  01/04/2022: Income Details " & Chr(13)
        If ChkMaxDatePrev(DateOfFillingRetrun, "") Then msgError = msgError & "* Date of filing of original Return cannot be prior to  01/04/2022: Income Details " & Chr(13)
   End If
End If

'If (DateOfFillingRetrun1 <> "") Then
'    If Not FormatNCheckDate(DateOfFillingRetrun1) Then
'        msgError = msgError & "* Date of Filing  must be a valid dd/mm/yyyy format in Sheet Income Details" & Chr(13)
'    Else
'        If ChkMaxDate(DateOfFillingRetrun1, "") Then msgError = msgError & "* Date of Filing  should be after 01/04/2018 in Sheet Income Details " & Chr(13)
'   End If
'End If

UpdateProgressBar
   
If (NoticeNo <> "") Then
        If Not ValidateNoticeNo_1(NoticeNo) Then msgError = msgError & errmsgfiling & Chr(13)
End If

'If (NoticeNo <> "") Then
'        If Not ValidateNoticeNo_2(NoticeNumber) Then msgError = msgError & errmsgfiling & Chr(13)
'End If
   
If (DateOfNotice <> "") Then
   If Not FormatNCheckDate(DateOfNotice) Then
'    msgError = msgError & "* Please enter a valid date in dd/mm/yyyy format" & Chr(13)
msgError = msgError & "* ""Date of such Notice or Order - Please enter valid date in valid format.""" & Chr(13)
   Else
   If ChkMaxDate(DateOfNotice, "") Then msgError = msgError & "* Date of notice or Order cannot be prior to 01/04/2024 in Sheet Income Details" & Chr(13)
   End If
End If




'new
'If Sheet1.Range("IncD.IncomeFromSal_allow").value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
'fmsgboxsmall "* Total of allowances shall not exceed salary as per 17(1)in Sheet Income Details ." ', vbExclamation, "Warning"
'End If

'Ayush_pan=====
'If Sheet1.Range("sheet1.status").value = "I - INDIVIDUAL" Then
'If Mid(Range("sheet1.PAN").value, 4, 1) <> "P" Then
'
'msgError = msgError & "* PAN For Individual the fourth letter should be 'P' "
'End If
'End If
'If Sheet1.Range("sheet1.status").value = "H - HUF" Then
'If Mid(Range("sheet1.PAN").value, 4, 1) <> "H" Then
'msgError = msgError & "* PAN for HUF the fourth letter should be 'H' "
'End If
'End If
'
'If Sheet1.Range("sheet1.status").value = "F - FIRM(Other than LLP)" Then
'If Mid(Range("sheet1.PAN").value, 4, 1) <> "F" Then
'msgError = msgError & "* PAN for FIRM(Other than LLP) the fourth letter should be 'F' "
'End If
'End If
'================
'Ayush_commented_27/01
'If ((IncomeHP1 <> "") Or (IncomeHP2 <> "") Or (IncomeHP5 <> "")) And ((IncomeHP1 <> 0) Or (IncomeHP2 <> 0) Or (IncomeHP5 <> 0)) Then
' If (TypeOfHP = "" Or TypeOfHP = "(Select)") Then
'' msgError = msgError & "* Please select the type of House Property in House Property Income Details" & Chr(13)
''Error Message Changed By Ayush for the SIT-56341
'  msgError = msgError & "* Please select the type of House Property " & Chr(13)
'End If
'End If
'Ayush_commented_27/01
'If TypeOfHP = "Let Out" Or TypeOfHP = "Deemed let out" Then
'If (IncomeHP1 = "" Or IncomeHP1 = 0) Then
'
''msgError = msgError & "*Gross rent received/ receivable/ letable value during the year should be more than zero in House Property Income Details" & Chr(13)
''msgError = msgError & "*Please enter amount in field Gross rent received/receivable/ lettable" & Chr(13)
''Error Message Changed By Ayush for the SIT-56341
' msgError = msgError & "*Gross rent received/ receivable/ lettable value during the year should be more than zero" & Chr(13)
'End If


'If (IncomeHP2 > 0) Then
'If (IncomeHP1 = 0 Or IncomeHP1 = "") Then
'msgError = msgError & "*Tax paid to local authorities can be claimed only if income from house property is declared in House Property Income Details" & Chr(13)
'End If
'End If

'End If

'If Mid(TypeOfHP, 1, 1) = "S" Then
'If (IncomeHP5 > 200000) Then
'msgError = msgError & "* In case of Self Occupied Interest payable on borrowed capital cannot exceed 200000 for income from House Property." & Chr(13)
'End If

'End If
'AYush_Legacy35
'If Sheet1.Range("IncD.IncomeFromSal_allow").value > Sheet1.Range("IncD.IncomeFromSal_i").value Then
'msgError = msgError & "* Total of allowances shall not exceed  Gross salary in Schedule Income Details" ', vbOKOnly, "Error"
'End If

'Ayush_Legacy35
If Sheet1.Range("IncD.IncomeFromSal_allow").value > Sheet1.Range("TotAllowances_17abc").value Then
msgError = msgError & "* Total of allowances shall not exceed gross salary at sl. no. ia+ib+ic in Schedule Income details" ', vbOKOnly, "Error"
End If

'If (Sheet4.Range("TDS2ii.Sum").value > 0) Then
'            If Sheet1.Range("IncD.TypeOfHP").value = "Let Out" Then
'                If Not Sheet1.Range("IncD.IncomeFromHP_i").value > 0 Then
'                    MsgBox "If TDS2(ii)(16C) is filled ,then please ensure the Type of house property is Let out and Gross rent received /receivable /letable value during the year is greater than zero", vbCritical
'                    CloseMsg
'                End If
'            Else
'
'                    MsgBox "If TDS2(ii)(16C) is filled ,then please ensure the Type of house property is Let out and Gross rent received /receivable /letable value during the year is greater than zero", vbCritical
'                    CloseMsg
'
'            End If
'End If


'If Sheet1.Range("IncD.TotalIncome").value > 5000000 Then
'Newly updated by Bindu as per DE V3 on 4th Feb 2025
If (Sheet1.Range("IncD.TotalIncome_New").value - Sheet5.Range("IncD.CG_LTCG").value) > 5000000 Then


'Newly changed by Bindu on 3rd Feb 2024 as per V3 DE sheet

'        fmsgboxsmall "* ITR-4 is for Individuals, HUFs and Firms (other than LLP) being a resident having income from business and profession which is computed under sections 44AD, 44ADA or 44AE and having total income upto Rs.50 Lakh, Please file ITR- 3 or 5" ', vbOKOnly, "Error"

'SIT-56341
         fmsgboxsmall "* ITR-4 is for Individuals, HUFs and Firms (other than LLP) being a resident having income from business and profession which is computed under sections 44AD, 44ADA or 44AE and having total income upto Rs.50 Lakh, Please file ITR- 3 or 5 and your taxable total income is more than Rs.50Lakh" ', vbOKOnly, "Error"

        CloseMsg
    End If

 

'If IncomeGross < 0 Then
'Newly updated by Bindu as per DE V3.1 on 11th Feb 25
If GrossTotIncomeIncLTCG112A < 0 Then
fmsgboxsmall "* To avail the benefit of carry forward and set off of loss, please use ITR-3 or ITR-5 ." ', vbExclamation, "Warning"
End If


'If Not ValidateSelect80D Then ValidateSheetID = False
'If Not ValidateSelect80DB Then ValidateSheetID = False
'If Not ValidateSelect80DC Then ValidateSheetID = False
'If Not ValidateSystem80D Then ValidateSheetID = False
'If Not ValidateSelect80DD Then ValidateSheetID = False
If Not ValidateSelect80DDB Then ValidateSheetID = False
'If Not ValidateSelect80U Then ValidateSheetID = False
If Not validateDed80TTB Then ValidateSheetID = False
'Ayush_26_27
If Not Validate_80CCC Then ValidateSheetID = False
'Ayush_16/06/2026
If Sheet1.Range("IncD.Section80CCD1B_SE").value > 0 Or Sheet1.Range("IncD.Section80CCD_SE").value > 0 Then
    If Not Validate_Pran Then ValidateSheetID = False
End If
'If Not Validate_80CCC_1 Then ValidateSheetID = False
'If Not Validate_80CCC_1b Then ValidateSheetID = False

'If Not MandatoryList(ResidentialStatus) Then msgError = msgError & "Residential Status is Mandatory in Sheet : Income Details" & Chr(13)
DedValueError = DeductionValues()
If DedValueError <> "" Then msgError = msgError & DedValueError & Chr(13)


If Not ValidateIDOthersEI Then ValidateSheetID = False
If Not ValidateOSOthersEI Then ValidateSheetID = False



If msgError <> "" Then
    ValidateSheetID = False
    Exit Function
End If



UpdateProgressBar

End Function


Function ChkSeventhProvisoFlag() As Boolean
On Error Resume Next
    ChkSeventhProvisoFlag = True
    ProvisoFlag = Sheet1.Range("sheet1.SeventhProvisoFlag").value
    
    If Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = "False" Then
    If Trim(ProvisoFlag) = "" Or Trim(ProvisoFlag) = "(Select)" Then
        ChkSeventhProvisoFlag = False
        Exit Function
    End If
    End If
End Function
Function ChkDepositAmountFlag() As Boolean
On Error Resume Next
    ChkDepositAmountFlag = True
    DepositAmountFlag = Sheet1.Range("Sheet1.DepositAmountFlag")
    DepositAmount = Sheet1.Range("Sheet1.DepositAmount")
    
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value = "Yes" Then
    
    If Trim(DepositAmountFlag) = "" Or Trim(DepositAmountFlag) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y" Then
    msgError = msgError + "*Selection of either of the options for 'Have you deposited amount or aggregate of amounts exceeding Rs. 1 Crore in one or more current account during the previous year?' is mandatory in Sheet Income Details" & Chr(13)
        ChkDepositAmountFlag = False
        Exit Function
    End If
    
    If Trim(DepositAmountFlag) = "Yes" And DepositAmount = "" Then
    msgError = msgError + "*Please enter amount of deposit in one or more current account during the previous year in Income Details"
    ChkDepositAmountFlag = False
        Exit Function
    End If
    End If
    
'    If Sheet1.Range("clauseiv7provisio139iFlg").value <> "No" Then
'    If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_1").value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").value = "Yes" Then
'    msgError = msgError + "*Selection of either of the options for 'total sales, turnover or gross receipts, as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or ' is mandatory in Sheet Income Details" & Chr(13)
'        ChkDepositAmountFlag = False
'        Exit Function
'    End If
'
'    If Sheet1.Range("clauseiv7provisio139iFlg_1").value = "Yes" And Sheet1.Range("clauseiv7provisio139iAmount_1").value = "" Then
'    msgError = msgError + "*Please enter amount of Total sales, turnover or gross receipts etc in Income Details"
'    ChkDepositAmountFlag = False
'    Exit Function
'    End If
'    End If
    
    
    
End Function
Function ChkAggrigateAmountFlag() As Boolean
On Error Resume Next
    ChkAggrigateAmountFlag = True
    AggrigateAmountFlag = Sheet1.Range("Sheet1.AggrigateAmountFlag")
    AggrigateAmount = Sheet1.Range("Sheet1.AggrigateAmount")
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value = "Yes" Then
    If Trim(AggrigateAmountFlag) = "" Or Trim(AggrigateAmountFlag) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y" Then
    msgError = msgError + "*Selection of either of the options for 'Have you incurred expenditure of an amount or aggregate of amount exceeding Rs. 2 lakhs for travel to a foreign country for yourself or for any other person' in Sheet Income Details" & Chr(13)
        ChkAggrigateAmountFlag = False
        Exit Function
    End If
    
    If Trim(AggrigateAmountFlag) = "Yes" And AggrigateAmount = "" Then
    msgError = msgError + "*Please enter amount of expenditure incurred for travel to a foreign country for yourself or for any other person in Income Details"
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
    If Sheet1.Range("sheet1.SeventhProvisoFlag").value = "Yes" Then
    If Trim(AggrigateAmountFlag1) = "" Or Trim(AggrigateAmountFlag1) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y" Then
    msgError = msgError + "* Selection of either of the options for 'Have you incurred expenditure of amount or aggregate of amount exceeding Rs. 1 lakh on consumption of electricity during the previous year?' in Sheet Income Details" & Chr(13)
        ChkAggrigateAmountFlag1 = False
        Exit Function
    End If
    
    If Trim(AggrigateAmountFlag1) = "Yes" And AggrigateAmount1 = "" Then
    msgError = msgError + "Please enter amount of expenditure incurred on consumption of electricity during the previous year in Income Details"
    ChkAggrigateAmountFlag1 = False
        Exit Function
    End If
    
    
    End If
End Function
Function ValidateIDOthersEI() As Boolean
ValidateIDOthersEI = True

    If Len((Range("Sheet1.Nature").item(1).value) > 0) Then
        If Not ValidateIDNatureOfIncome Then ValidateIDOthersEI = False
        If Not ValidateIDAmount Then ValidateIDOthersEI = False
    End If

   setTblinfo_IDOthersNOI
   setTblinfo_IDOthersAmt
   
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Dim i As Long
    
    Set rangecells = Range("Sheet1.Nature").Cells
 
    ReDim Others_IDNOI(end_IDOthersAmt)

    For i = 1 To end_IDOthersAmt
   
    cellrange = GetMergedAddressCell(rangecells, i)
   
    Others_IDNOI(i) = Sheet1.Range(cellrange).value

        
            If Others_IDNOI(i) = "(Select)" Or Others_IDNOI(i) = "" Then
'             msgError = msgError & "Please select an option from the drop down at Point B2(ii) at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
             'SIT-56341
              msgError = msgError & "Please select an option from the drop down under Nature of Exempt Allowance in Salary table at Point B2(ii) at Sr. No  " & i & "" & Chr(13)

             ValidateIDOthersEI = False
             Exit Function
         End If

   Next


End Function

Sub setTblinfo_IDOthersNOI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet1.Nature").count
    Set rangecells = Range("Sheet1.Nature").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_IDOthersNOI = ccount
 End Sub
 
 Sub setTblinfo_IDOthersAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet1.Amount").count
    Set rangecells = Range("Sheet1.Amount").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_IDOthersAmt = ccount
End Sub

Function ValidateIDNatureOfIncome() As Boolean
ValidateIDNatureOfIncome = True

    setTblinfo_IDOthersNOI
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Range("Sheet1.Nature").Cells
    Set rangecells1 = Range("Sheet1.Description").Cells
    Dim cellrange As String
    Dim cellRange1 As String
    Dim i As Long
    ReDim Others_IDNOI(end_IDOthersNOI)
    ReDim Others_IDNOI1(end_IDOthersNOI)
    
    

    
'    If Sheet4.Range("TDS.IncSum").value > 0 Then
'        If IIf(Sheet1.Range("IncD.IncomeFromSal").value = "", 0, Sheet1.Range("IncD.IncomeFromSal").value) < (Sheet4.Range("TDS.IncSum").value - Round(Sheet4.Range("TDS.IncSum").value * 0.1, 0)) Then
'        'IIf(Sheet1.Range("IncD.IncomeFromSal").value = "", 0, Sheet1.Range("IncD.IncomeFromSal").value) > (Sheet4.Range("TDS.IncSum").value + Round(Sheet4.Range("TDS.IncSum").value * 0.1, 0)) Then
'        If Not end_OthersNOI > 0 Then
'               msgError = msgError & "Since the amount disclosed in Income chargeable under the Head Salaries is less than 90% of Salary reported in TDS1, please ensure to fill details in " & Chr(34) & "Others" & Chr(34) & " in " & """Exempt Income""" & Chr(13)
'               ValidateNatureOfIncome = False
'             Exit Function
'           End If
'         End If
'    End If
    
     
    For i = 1 To end_IDOthersNOI
    cellrange = GetMergedAddressCell(rangecells, i)
    cellRange1 = GetMergedAddressCell(rangecells1, i)
  
        Others_IDNOI(i) = Sheet1.Range(cellrange).value
        Others_IDNOI1(i) = Sheet1.Range(cellRange1).value
            
         If (Others_IDNOI(i) = "(Select)" Or Others_IDNOI(i) = "") Then
             msgError = msgError & "* Please select an option from the drop down in Exempt Income in Sheet Income Details" & Chr(13)
             ValidateIDNatureOfIncome = False
             Exit Function
         End If
            
'         If Not chkCompulsory(Others_NOI(i)) Then
'             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "Others: Nature of Income at Sr. No  " & i & "  in Sheet Taxes Paid and Verification  is mandatory" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
         
         If (Others_IDNOI(i) = "Any Other") Then
         If Others_IDNOI1(i) = "" Then
         'Change.03.03.2023.102.IDS.113
             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
             msgError = msgError & "* Please enter description in Nature of Exempt Allowance in Salary table at  Point B2(ii) in Sheet Income Details" & Chr(13)
         'End Change.IDS113
             ValidateIDNatureOfIncome = False
             Exit Function
         End If
         End If
         
    'PAG_C34 AY 2024-25 Change
    If (Others_IDNOI(i) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent") Then
        'If (IsEmpty(Range("IncD.Section80GG_Calc").value) = False) Then
    If Range("IncD.Section80GG_Calc").value > 0 Then
        msgError = msgError & "* Deduction u/s 10(13A) & 80GG cannot be claimed for the same period"
             ValidateIDNatureOfIncome = False
             Exit Function
             End If
    End If

         
         
         If Len(Others_IDNOI1(i)) > 125 Then
             msgError = msgError & "* Description at Sr. No  " & i & "  in Sheet Income Details cannot exceed 125 characters in Sheet Income Details" & Chr(13)
             ValidateIDNatureOfIncome = False
             Exit Function
         End If
    Next
    
'    If Len(Sheet5.Range("Total_Exempt").value) > 14 Then
'             msgError = msgError & "Total of Excempt Should not be exceed 14 digits" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'    End If
    
End Function


Function ValidateIDAmount() As Boolean
ValidateIDAmount = True

    setTblinfo_IDOthersNOI
    Dim rangecells As Range
'    Dim rangecells_exempallow As Range 'AYUSH
'    Dim cellrangeExALL As String 'AYUSH
'    Dim Others_IDexemptallow As Variant  'AYUSH
    
    Dim cellrange As String
    Set rangecells = Range("Sheet1.Amount").Cells
    'Set rangecells_exempallow = Range("Sheet1.Nature").Cells 'AYUSH
    Dim i As Long
    ReDim Others_IDAmt(end_IDOthersNOI)
    ReDim Others_IDexemptallow(end_IDOthersNOI)
    
    For i = 1 To end_IDOthersNOI
    cellrange = GetMergedAddressCell(rangecells, i)
    'cellrangeExALL = GetMergedAddressCell(rangecells_exempallow, i) 'AYUSH
    
        Others_IDAmt(i) = Sheet1.Range(cellrange).value
     '   Others_IDexemptallow(i) = Sheet1.Range(cellrangeExALL).value  'AYUSH
        
        If Not chkCompulsory(Others_IDAmt(i)) Then
            msgError = msgError & "* Please enter amount for selected Nature of Exempted Allowance in Sheet Income Details" & Chr(13)
            ValidateIDAmount = False
            Exit Function
        End If
        
        
        If Not IsNumeric(Others_IDAmt(i)) Then
            msgError = msgError & "* Amount for selected Nature of Exempted Allowance at Sr. No  " & i & "  in should be Numeric value in Sheet Income Details" & Chr(13)
            ValidateIDAmount = False
            Exit Function
        End If
        
        If Others_IDAmt(i) > 99999999999999# Then
            msgError = msgError & "* Amount at Sr. No  " & i & " cannot exceed 14 digits in Sheet Income Details" & Chr(13)
            ValidateIDAmount = False
            Exit Function
        End If
        
'        'Ayush----------------------
'        If Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then
'        If Others_IDAmt(i) >= 2500000 And Others_IDexemptallow(i) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'            msgError = msgError & "*Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is only upto Rs 25 lakh." & Chr(13)
'            ValidateIDAmount = False
'            Exit Function
'        End If
'        End If
        '-----------------------------
        
        
    Next
End Function
Function ValidateOSOthersEI() As Boolean
ValidateOSOthersEI = True

    If Len((Range("Sheet1.Nature1").item(1).value) > 0) Then
        If Not ValidateOSNatureOfIncome Then ValidateOSOthersEI = False
        If Not ValidateOSAmount Then ValidateOSOthersEI = False
    End If

    setTblinfo_OSOthersNOI
    setTblinfo_OSOthersAmt
   
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Dim i As Long
    
    Set rangecells = Range("Sheet1.Nature1").Cells
 
    ReDim Others_OSNOI(end_OSOthersAmt)

    For i = 1 To end_OSOthersAmt
   
    cellrange = GetMergedAddressCell(rangecells, i)
   
    Others_OSNOI(i) = Sheet1.Range(cellrange).value

        
            If Others_OSNOI(i) = "(Select)" Or Others_OSNOI(i) = "" Then
             msgError = msgError & "Please select an option from the drop down at B3 'Income from other Sources' at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
             ValidateOSOthersEI = False
             Exit Function
         End If
         
         If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "I" Then
          If Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)" Or Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)" Or Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)" Or Others_OSNOI(i) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)" Then
             msgError = msgError & "Please select nature of income from the dropdown for status HUF/Firm at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
             ValidateOSOthersEI = False
             Exit Function
         End If
         End If

   Next
   
     If Sheet1.Range("temp5").value > 0 Then
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
    msgError = msgError & "*  Family pension should not allowed to be selected by a Firm(Other than LLP)in Sheet Income Details" ', vbOKOnly, "Error(s)!"
   ValidateOSOthersEI = False
             Exit Function
    End If
    End If


End Function

Sub setTblinfo_OSOthersNOI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet1.Nature1").count
    Set rangecells = Range("Sheet1.Nature1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OSOthersNOI = ccount
 End Sub
 
 Sub setTblinfo_OSOthersAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet1.Amount1").count
    Set rangecells = Range("Sheet1.Amount1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_OSOthersAmt = ccount
End Sub

Function ValidateOSNatureOfIncome() As Boolean
ValidateOSNatureOfIncome = True

    setTblinfo_OSOthersNOI
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Set rangecells = Range("Sheet1.Nature1").Cells
    Set rangecells1 = Range("Sheet1.Description1").Cells
    Dim i As Long
    ReDim Others_OSNOI(end_OSOthersNOI)
    ReDim Others_OSNOI1(end_OSOthersNOI)

    
'    If Sheet4.Range("TDS.IncSum").value > 0 Then
'        If IIf(Sheet1.Range("IncD.IncomeFromSal").value = "", 0, Sheet1.Range("IncD.IncomeFromSal").value) < (Sheet4.Range("TDS.IncSum").value - Round(Sheet4.Range("TDS.IncSum").value * 0.1, 0)) Then
'        'IIf(Sheet1.Range("IncD.IncomeFromSal").value = "", 0, Sheet1.Range("IncD.IncomeFromSal").value) > (Sheet4.Range("TDS.IncSum").value + Round(Sheet4.Range("TDS.IncSum").value * 0.1, 0)) Then
'        If Not end_OthersNOI > 0 Then
'               msgError = msgError & "Since the amount disclosed in Income chargeable under the Head Salaries is less than 90% of Salary reported in TDS1, please ensure to fill details in " & Chr(34) & "Others" & Chr(34) & " in " & """Exempt Income""" & Chr(13)
'               ValidateNatureOfIncome = False
'             Exit Function
'           End If
'         End If
'    End If
    
     
    For i = 1 To end_OSOthersNOI
    cellrange = GetMergedAddressCell(rangecells, i)
    cellRange1 = GetMergedAddressCell(rangecells1, i)
    
        Others_OSNOI(i) = Sheet1.Range(cellrange).value
        Others_OSNOI1(i) = Sheet1.Range(cellRange1).value
            
         If (Others_OSNOI(i) = "(Select)" Or Others_OSNOI(i) = "") Then
             msgError = msgError & "* Please select an option from the drop down at B4" & Chr(13)
             ValidateOSNatureOfIncome = False
             Exit Function
         End If
            
'         If Not chkCompulsory(Others_NOI(i)) Then
'             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "Others: Nature of Income at Sr. No  " & i & "  in Sheet Taxes Paid and Verification  is mandatory" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
         
         If (Others_OSNOI(i) = "Any Other") Then
         If Others_OSNOI1(i) = "" Then
'             msgError = msgError & "* Please enter description in Nature of Exempt Allowance in income details at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
     'SIT-56341
              msgError = msgError & "* Please enter description in Nature of Income (Income from other Sources) at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
             
             ValidateOSNatureOfIncome = False
             Exit Function
         End If
         End If
         
         
         If Len(Others_OSNOI1(i)) > 125 Then
             msgError = msgError & "* Description at Sr. No  " & i & "  in Sheet Income Details cannot exceed 125 characters" & Chr(13)
             ValidateOSNatureOfIncome = False
             Exit Function
         End If
    Next
    
'    If Len(Sheet5.Range("Total_Exempt").value) > 14 Then
'             msgError = msgError & "Total of Excempt Should not be exceed 14 digits" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'    End If
    
End Function


Function ValidateOSAmount() As Boolean
ValidateOSAmount = True

    setTblinfo_OSOthersNOI
    Dim rangecells As Range
    Dim cellrange As String
    Set rangecells = Range("Sheet1.Amount1").Cells
    Dim i As Long
    ReDim Others_OSAmt(end_OSOthersNOI)
    For i = 1 To end_OSOthersNOI
    
    cellrange = GetMergedAddressCell(rangecells, i)
   

        Others_OSAmt(i) = Sheet1.Range(cellrange).value
        
        If Not chkCompulsory(Others_OSAmt(i)) Then
            msgError = msgError & "Income from other Sources - Please enter Amount in other source Income Details" & Chr(13)
            ValidateOSAmount = False
            Exit Function
        End If
        
        
        If Not IsNumeric(Others_OSAmt(i)) Then
            msgError = msgError & "Nature of Exempted Allowance at Sr. No  " & i & "  in Sheet Schedule Income Details should be Numeric value" & Chr(13)
            ValidateOSAmount = False
            Exit Function
        End If
        
        If Others_OSAmt(i) > 99999999999999# Then
            msgError = msgError & "* Nature of Exempted Allowance at Sr. No  " & i & "  in Sheet Schedule Income Details cannot exceed 14 digits" & Chr(13)
            ValidateOSAmount = False
            Exit Function
        End If
    Next
End Function
Function ValidateAadhaar()
On Error Resume Next
    ValidateAadhaar = True
    
    If AadhaarNumber <> "" Then
        If Not IsNumeric(AadhaarNumber) Then
            'errmsgID = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
        
        If AadhaarNumber = "000000000000" Or AadhaarNumber = "111111111111" Then
            'errmsgID = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
        
        If Len(AadhaarNumber) <> 12 Then
            errmsgID = "* For Aadhaar Number- 12 digit"
            ValidateAadhaar = False
            Exit Function
        End If
    End If
    
    If AadhaarEnrolNumber <> "" Then
        If Not IsNumeric(AadhaarEnrolNumber) Then
            'errmsgID = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
        
        If AadhaarEnrolNumber = "0000000000000000000000000000" Or AadhaarEnrolNumber = "1111111111111111111111111111" Then
            'errmsgID = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
        
        If Len(AadhaarEnrolNumber) <> 28 Then
            errmsgID = "* For Aadhaar Enrolment Id- 28 digit."
            ValidateAadhaar = False
            Exit Function
        End If
        
        'If Not ValidateAadharEnrol(AadhaarEnrolNumber) Then
        '    validateAadhaar = False
        '    Exit Function
        'End If
    End If
End Function

Function FillingStatusValidations() As Boolean
On Error Resume Next
Dim result As Variant
Dim flag As Boolean
flag = True
FillingStatusValidations = True
errmsgfiling = ""


If ReturnFilledSection1 = "" Or ReturnFilledSection1 = "(Select)" Then
        'Newly changed by Bindu
        'errmsgfiling = errmsgfiling & "* Please select ""Filed u/s"" or ""Filed in response to notice u/s"" in Sheet Income Detils" & Chr(13)
        errmsgfiling = errmsgfiling & "* Filing section is mandatory" & Chr(13)
        
        FillingStatusValidations = False
        Exit Function
End If

'If ReturnFilledSection2 <> "" And ReturnFilledSection2 <> "(Select)" Then
'    If ReturnFilledSection1 <> "" And ReturnFilledSection1 <> "(Select)" Then
'        errmsgfiling = errmsgfiling & "Either ""Filed u/s"" or ""Filed in response to notice u/s"" must have to be selected in tab: Part A General Information" & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'    End If
'End If
  




If (ReturnFilledSection = "" Or ReturnFilledSection = "(Select)") Then
    'Newly added by Bindu
    errmsgfiling = errmsgfiling & "* Filing section is mandatory" & Chr(13)
    'errmsgfiling = errmsgfiling & "* Please select ""Filed u/s"" or ""Filed in response to notice u/s"" in Sheet Income Detils" & Chr(13)
    FillingStatusValidations = False
    Exit Function
End If
    
'If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 17 Then
'    errmsgfiling = errmsgfiling & "Return Type cannot be Revised if Return is NOT filed u/s 139(5) in Sheet : Income Details" & Chr(13)
'
'    FillingStatusValidations = False
'    Exit Function
'End If

If Mid(Trim(UCase(TypeOfReturn)), 1, 3) = "REC" Then
If RecieptNo1 = "" Then
    errmsgfiling = errmsgfiling & "* Please enter Reciept Number in Sheet Income Details" & TypeOfReturn & Chr(13)
        FillingStatusValidations = False
        Exit Function
End If

'If DateOfFillingRetrun1 = "" Then
'    errmsgfiling = errmsgfiling & "Please enter Date Of Filing in Sheet Income Details" & TypeOfReturn & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'End If

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 13 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 14 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 15 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 16 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
        If Sheet1.Range("sheet1.NoticeNo").value = "" Then
'        Change.02.03.2023.102.IDS.86
'Ayush_25-26
'                errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148"" or Filed u/s ""119(2)(b)""in Sheet Income Details" & Chr(13)
                'errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)""in Sheet Income Details" & Chr(13)
                errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)""in Sheet Income Details" & Chr(13)
'        End Change IDS.86
                FillingStatusValidations = False
                Exit Function
            End If
            
            
            
            
            If Sheet1.Range("sheet1.NoticeDate").value = "" Then
'            Change.02.03.2023.102.IDS.87
'       errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148"" or Filed u/s ""119(2)(b)- after condonation of delay""in Sheet Income Details" & Chr(13)
          
          'J200924 need to check
          
        ' errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)""" & Chr(13)
          errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" " & Chr(13)
          

'            End Change IDS.87

                FillingStatusValidations = False
                Exit Function
            End If
            
            

End If

Else

'
'If RecieptNo1 <> "" Then
'    If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 17 And Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 18 And Mid(Trim(UCase(TypeOfReturn)), 1, 3) = "Rec" Then
'        errmsgfiling = errmsgfiling & "Please Do not fill Reciept no for " & ReturnFilledSection & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'    End If
'End If
'
'If DateOfFillingRetrun1 <> "" Then
'    If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 17 And Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 18 And Mid(Trim(UCase(TypeOfReturn)), 1, 3) = "Rec" Then
'        errmsgfiling = errmsgfiling & "Please Do not fill Date Of Filing for " & ReturnFilledSection & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'    End If
'End If

'If NoticeNo <> "" Then
'If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 18 And Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) <> 20 Then
'errmsgfiling = errmsgfiling & "Please Do not fill Unique Number for " & ReturnFilledSection & Chr(13)
'FillingStatusValidations = False
'Exit Function
'End If
'End If

'If NoticeNumber <> "" Then
'    If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 11 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 12 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 17 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 18 Then
'        errmsgfiling = errmsgfiling & "Please Do not fill Notice Number for " & ReturnFilledSection & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'End If
'End If
'If NoticeNumber = "" Then

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 13 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 14 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 15 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 16 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
        If Sheet1.Range("sheet1.NoticeNo").value = "" Then
'        Change.02.03.2023.102.IDS.86 A
'                errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148"" or Filed u/s ""119(2)(b)""in Sheet Income Details" & Chr(13)
             '   errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)""in Sheet Income Details" & Chr(13)
             errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)"" in Sheet Income Details" & Chr(13)
'        End Change IDS.86A
                FillingStatusValidations = False
                Exit Function
            End If

End If
    
    



If DateOfNotice <> "" Then
    If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 11 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 12 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 17 Then
        errmsgfiling = errmsgfiling & "* Please Do not fill Date of Notice for " & ReturnFilledSection & Chr(13)
        FillingStatusValidations = False
        Exit Function
        
    End If
Else
    If Not (Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 11 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 12 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 17 Or Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20) Then
        
        If Mid(ReturnFilledSection, 1, 2) = 18 Then
            If Sheet1.Range("sheet1.ReceiptNo").value = "" Then
                errmsgfiling = errmsgfiling & "* Please Enter Receipt Number in Sheet Income Details" & Chr(13)
                flag = False
            End If
            
            If Not chkNumeric(Sheet1.Range("sheet1.ReceiptNo").value) Then
            errmsgfiling = errmsgfiling & "* Invalid Receipt Number. Receipt number should be 15 digits in Sheet Income Details" & ReturnFilledSection & Chr(13)
            flag = False
            End If
            
            
            If Sheet1.Range("sheet1.ReceiptNo").value = "000000000000000" Then
                errmsgfiling = errmsgfiling & "* Invalid Receipt Number. Receipt number should be 15 digits in Sheet Income Details" & ReturnFilledSection & Chr(13)
                flag = False
            End If
            
            If Sheet1.Range("sheet1.OrigRetFiledDate").value = "" Then
                errmsgfiling = errmsgfiling & "* Date of filing of original Return is mandatory in Sheet Income Details" & Chr(13)
                flag = False
            End If
            
            If Sheet1.Range("sheet1.NoticeNo").value = "" Then
                'Change-25.01.2023.103.ID.27
                'errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148 in Sheet Income Details" & Chr(13)
'                errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C or Filed u/s is 119(2)(b) in Sheet Income Details" & Chr(13)
                 errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C or Filed u/s is ""139(9A)"" in Sheet Income Details" & Chr(13)
                flag = False
            End If
            
            If Sheet1.Range("sheet1.NoticeNo").value = "00000000000000000000000" Then
                errmsgfiling = errmsgfiling & "* Please enter valid Unique number/ DIN in Sheet Income Details" & ReturnFilledSection & Chr(13)
                flag = False
            End If
            
            If Len(Sheet1.Range("sheet1.NoticeNo").value) > 23 Then
                errmsgfiling = errmsgfiling & "* Unique number/ DIN cannot exceed 23 characters" & ReturnFilledSection & Chr(13)
                flag = False
            End If
            
        End If
            If Sheet1.Range("sheet1.NoticeDate").Locked = False And Sheet1.Range("sheet1.NoticeDate").value = "" Then
            
'            Change.02.03.2023.IDS.87A
              'PAG_C31 AY 2023-24 Old
'                errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148"" or Filed u/s ""119(2)(b)- after condonation of delay""in Sheet Income Details" & Chr(13)

              'PAG_C31 AY 2024-25 Change
              
              'J200924 need to check
       'errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)""" & Chr(13)
 errmsgfiling = errmsgfiling & "* Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C""  or Filed u/s is ""119(2)(b)"" " & Chr(13)

'         End Change.87A
                flag = False
            End If
            
            If flag = False Then
                FillingStatusValidations = False
                Exit Function
            End If
     End If
End If


If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 18 Then
'    If Mid(Trim(UCase(TypeOfReturn)), 1, 1) <> "I" Then
'        errmsgfiling = errmsgfiling & "Type of Return should be In response to Notice u/s 139(9) in Sheet : Income Details" & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'    Else
        If RecieptNo = "" Then errmsgfiling = errmsgfiling & "* Please enter Receipt Number is mandatory Sheet Income Details" & Chr(13)
'        Change.27.01.2023.102.CDS01
'Ayush_25-26
'        If NoticeNo = "" Then errmsgfiling = errmsgfiling & "* Unique Number is Mandatory in Sheet Income Details" & Chr(13)
       ' If NoticeNo = "" Then errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)""in Sheet Income Details" & Chr(13)
        If NoticeNo = "" Then errmsgfiling = errmsgfiling & "* Unique number/ DIN is mandatory for return ""Filed in response to notice u/s"" is ""139(9)/142(1)/148/153C"" or Filed u/s is ""139(9A)""in Sheet Income Details" & Chr(13)
'       End Change======
        If DateOfFillingRetrun = "" Then
            errmsgfiling = errmsgfiling & "* Date of filing of original Return is mandatory in sheet Income Details" & Chr(13)
        Else
            If Not FormatNCheckDate(DateOfFillingRetrun1) Then errmsgfiling = errmsgfiling & "* Please enter a valid date in dd/mm/yyyy format in Sheet Income Details" & Chr(13)
            If ChkMaxDatePrev(DateOfFillingRetrun, "31/03/2022") Then errmsgfiling = errmsgfiling & "* Date of filing of original Return cannot be prior to  01/04/2022 in Sheet Income Details" & Chr(13)
        End If
            
            
        If DateOfNotice = "" Then
            errmsgfiling = errmsgfiling & "* Notice Date is mandatory in sheet Income Details : Income Details" & Chr(13)
        Else
        'J200924
           ' If Not FormatNCheckDate(DateOfNotice) Then errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)- after condonation of delay in Sheet Income Details""." & Chr(13)
            If Not FormatNCheckDate(DateOfNotice) Then errmsgfiling = errmsgfiling & "* Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" " & Chr(13)
            
            If ChkMaxDate(DateOfNotice, "31/03/2022") Then errmsgfiling = errmsgfiling & "* Date of notice or Order cannot be prior to 01/04/2022 in Sheet Income Details" & Chr(13)
        End If
        
        If errmsgfiling <> "" Then
            FillingStatusValidations = False
            Exit Function
        End If
    End If
End If

If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 17 Then
'    If Mid(Trim(UCase(TypeOfReturn)), 1, 1) <> "R" Then
'        errmsgfiling = errmsgfiling & "Type of Return should be Revised in Sheet : Income Details" & Chr(13)
'        FillingStatusValidations = False
'        Exit Function
'    Else
        NoticeNo = ""
        DateOfNotice = ""
        If RecieptNo = "" Then errmsgfiling = errmsgfiling & "* Receipt Number is mandatory in Sheet Income Details" & Chr(13)
        If DateOfFillingRetrun = "" Then
            errmsgfiling = errmsgfiling & "* Date of filing of original Return is mandatory in sheet Income Details" & Chr(13)
        Else
           ' If Not FormatNCheckDate(DateOfFillingRetrun) Then errmsgfiling = errmsgfiling & "* Please enter a valid date in dd/mm/yyyy format in Sheet Income Details" & Chr(13)
            'If ChkMaxDate(DateOfFillingRetrun, "31/03/2022") Then errmsgfiling = errmsgfiling & "* Date of filing of original Return cannot be prior to  01/04/2022 in Sheet Income Details" & Chr(13)
           ' If ChkMaxDatePrev(DateOfFillingRetrun, "31/03/2022") Then errmsgfiling = errmsgfiling & "* Date of filing of original Return cannot be prior to  01/04/2022 in Sheet Income Details" & Chr(13)
        End If

        If errmsgfiling <> "" Then
            FillingStatusValidations = False
            Exit Function
        End If
'    End If
'End If
End If
'Ayush-05
If Mid(Trim(UCase(ReturnFilledSection)), 1, InStr(Trim(UCase(ReturnFilledSection)), "-") - 1) = 20 Then
'If Mid(Trim(UCase(TypeOfReturn)), 1, 1) = "O" Then
        If DateOfNotice = "" Then

        'J200924
       ' errmsgfiling = errmsgfiling & "Date of notice or Order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s ""119(2)(b)- after condonation of delay in Sheet Income Details" & Chr(13)
         errmsgfiling = errmsgfiling & "* Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)""" & Chr(13)

        FillingStatusValidations = False
        Exit Function
        End If
End If
'End If
        
End Function

Function DeductionValues() As Variant
On Error Resume Next
DeductionValues = ""

If Not CheckDedValue(IncomeSAl, False) Then DeductionValues = DeductionValues & "Income from salary " & dedValidateResult & Chr(13)

'If Not CheckDedValue(IncomeHP, True) Then
'    DeductionValues = DeductionValues & "Income from House Property " & dedValidateResult & Chr(13)
'Else
'    If Not hpValidation(IncomeHP, TypeOfHp) Then DeductionValues = DeductionValues & "" & dedValidateResult & Chr(13)
'End If

If Not CheckDedValue(IncomeOS, True) Then DeductionValues = DeductionValues & "Income from House Property " & dedValidateResult & Chr(13)
If Not CheckDedValue(IncomeOS, True) Then DeductionValues = DeductionValues & "Income from House Property " & dedValidateResult & Chr(13)


End Function
Function CheckDedValue(item As Variant, negative As Boolean) As Boolean
On Error Resume Next
CheckDedValue = True
dedValidateResult = ""
If item <> "" Then
    If Not negative Then
        If item < 0 Or Not IsNumeric(item) Then
        dedValidateResult = "Cannot be negative or non mumeric"
        CheckDedValue = False
        End If
    Else
        If Not IsNumeric(item) Then
        dedValidateResult = "Cannot be non mumeric"
        CheckDedValue = False
        End If
    End If
End If
End Function

Function hpValidation(HP As Variant, HpType As Variant) As Boolean
On Error Resume Next
hpValidation = True
dedValidateResult = ""

If Not (HP = "" Or HP = 0) Then
    If UCase(HpType) = "(SELECT)" Or UCase(HpType) = "" Then
        dedValidateResult = "* Please select Type of House Property from drop down"
        hpValidation = False
    End If
    
    If UCase(HpType) = "SELF OCCUPIED" And HP < -200000 Then
        dedValidateResult = "* Income from House Property :Loss cannot exceed -200000 for Self Occupied"
        hpValidation = False
    End If
    
    If UCase(HpType) = "LET OUT" And HP < -200000 Then
        dedValidateResult = "* Income from House Property :Loss cannot exceed -200000 for Let Out"
        hpValidation = False
    End If

If UCase(HpType) = "DEEMED LET OUT" And HP < -200000 Then
        dedValidateResult = "* Income from House Property :Loss cannot exceed -200000 for Let Out"
        hpValidation = False
    End If
End If

If MandatoryList(HpType) Then
If HP = "" Then
dedValidateResult = "* Income from House Property is Required in Sheet  Income Details"
hpValidation = False
End If
End If

End Function

Function MandatoryList(item As Variant) As Boolean
On Error Resume Next
MandatoryList = True
If IsEmpty(item) Or UCase(item) = "(SELECT)" Then MandatoryList = False
Exit Function
End Function

Function CheckDateOfNotice(DateOfNotice As Variant) As Variant
On Error Resume Next
Dim val As Variant
CheckDateOfNotice = 0
If LCase(ReturnFilledSection) = "18 - u/s 139(9)" Or LCase(ReturnFilledSection) = "13 - u/s 142(1)" Or LCase(ReturnFilledSection) = "14 - u/s 148" Or LCase(ReturnFilledSection) = "15 - u/s 153A" Or LCase(ReturnFilledSection) = "16 - u/s 153C r/w 153A" Then
    
    If IsEmpty(DateOfNotice) Then
        CheckDateOfNotice = "is Mandatory"
        Exit Function
    End If
    
    val = CheckDate(DateOfNotice)
    If val = 1 Then CheckDateOfNotice = "* must be a valid dd/mm/yyyy format"
    If val <> 2 Then CheckDateOfNotice = "* should not be after 31/03/2022"
Else
CheckDateOfNotice = "is not required to fill for the selected Filled Return Section"
Exit Function
End If

End Function

Function CheckNoticeRecieptNo(No As Variant) As Boolean
On Error Resume Next
CheckNoticeRecieptNo = True
If Not (Len(No) = 15 And IsNumeric(No)) Then
CheckNoticeRecieptNo = False
Exit Function
End If
End Function

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
        'Change.03.03.2023.102.IDSNO.109
        'errmsgID = "should be less than " & length & " charachters" & Chr(13)
        errmsgID = "should be less than " & Length & " Characters" & Chr(13)
        'End Change
        Exit Function
    End If
End If

  
End Function

Function CheckPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

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
        
        If Not CheckSpecialCharacterdot(PAN) Then
        CheckPAN = False
            Exit Function
        End If
               
    End If
End Function
Function CheckPAN1(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPAN1 = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
                
        If Not ChkAlphabetP(Mid(PAN, 4, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
            
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacterdot(PAN) Then
        CheckPAN1 = False
            Exit Function
        End If
               
    End If
End Function


'Pincode Validation
Function ChkPincode() As Boolean
On Error Resume Next
    ChkPincode = True
    Dim i As Long
    
    PinCode = Sheet1.Range("sheet1.PinCode")
    If Country <> "" And UCase(Mid(Country, 1, InStr(1, Country, "-") - 1)) = "91" Then
        If Trim(PinCode) = "" Or IsEmpty(PinCode) Then
            msgError = msgError & "* PinCode under primary address is Mandatory in Sheet  Income Details" & Chr(13)
            ChkPincode = False
            Exit Function
        End If
        
    
        If Len(Trim(PinCode)) > 6 Then
            msgError = msgError & "* Pincode under primary address in Sheet Income Details  cannot exceed 6 characters " & Chr(13)
            ChkPincode = False
            Exit Function
        End If
        
         For i = 1 To Len(PinCode)
            If Not IsNumeric(Mid(PinCode, i, 1)) Then
                msgError = msgError & "* PinCode under primary address in Sheet Income Details  must contain only digits from 0 to 9" & Chr(13)
                ChkPincode = False
                Exit Function
            End If
        Next
    Else
        PinCode = ""
    End If
    
End Function


'ZipCode Validation
Function ChkZipcode() As Boolean
On Error Resume Next
    ChkZipcode = True
    zipCode = Sheet1.Range("sheet1.ZipCode")
    
    
      If Country <> "" And UCase(Mid(Country, 1, InStr(1, Country, "-") - 1)) <> "91" Then
        If Trim(zipCode) = "" Or IsEmpty(zipCode) Then
            msgError = msgError & "* ZipCode under primary address in Sheet Income Details  is Mandatory, if there is no ZipCode then select ""No ZIP Code""" & Chr(13)
            ChkZipcode = False
            Exit Function
        End If
        
    
        If Len(Trim(zipCode)) > 8 Then
            msgError = msgError & "* ZipCode under primary address in Sheet Income Details cannot exceed 8 characters.Minimum 1 and up to 8 Characters." & Chr(13)
            ChkZipcode = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter3(zipCode) Then
            msgError = msgError & "* ZipCode under primary address in Sheet Income Details characters < > & ' " & Chr(34) & " are not allowed"" & chr(13)"
            ChkZipcode = False
            Exit Function
        End If
    Else
        zipCode = ""
    End If
End Function
'Ayush_13/01/2026_V0.2
' Secondary Pincode Validation
Function ChkPincode_Sec() As Boolean
On Error Resume Next
    ChkPincode_Sec = True
    Dim i As Long
    
    PinCode_Sec = Sheet1.Range("sheet1.PinCode2")
    If Country_Sec <> "" And UCase(Mid(Country_Sec, 1, InStr(1, Country_Sec, "-") - 1)) = "91" Then
        If Trim(PinCode_Sec) = "" Or IsEmpty(PinCode_Sec) Then
            msgError = msgError & "* PinCode is Mandatory under secondary address in Sheet  Income Details" & Chr(13)
            ChkPincode_Sec = False
            Exit Function
        End If
        
    
        If Len(Trim(PinCode_Sec)) > 6 Then
            msgError = msgError & "* Pincode under secondary address in Sheet Income Details  cannot exceed 6 characters " & Chr(13)
            ChkPincode_Sec = False
            Exit Function
        End If
        
         For i = 1 To Len(PinCode_Sec)
            If Not IsNumeric(Mid(PinCode_Sec, i, 1)) Then
                msgError = msgError & "* PinCode under secondary address in Sheet Income Details  must contain only digits from 0 to 9" & Chr(13)
                ChkPincode_Sec = False
                Exit Function
            End If
        Next
    Else
        PinCode_Sec = ""
    End If
    
End Function


'Secondary ZipCode Validation
Function ChkZipcode_Sec() As Boolean
On Error Resume Next
    ChkZipcode_Sec = True
    zipCode_Sec = Sheet1.Range("sheet1.ZipCode2")
    
    
      If Country_Sec <> "" And UCase(Mid(Country_Sec, 1, InStr(1, Country_Sec, "-") - 1)) <> "91" Then
        If Trim(zipCode_Sec) = "" Or IsEmpty(zipCode_Sec) Then
            msgError = msgError & "*  Zip Code under secondary address is compulsory, if there is no Zip Code then select ""No ZIP Code""." & Chr(13)
            ChkZipcode_Sec = False
            Exit Function
        End If
        
    
        If Len(Trim(zipCode_Sec)) > 8 Then
            msgError = msgError & "* ZipCode under secondary address in Sheet Income Details cannot exceed 8 characters.Minimum 1 and up to 8 Characters." & Chr(13)
            ChkZipcode_Sec = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacter4New(zipCode_Sec) Then
            msgError = msgError & "* ZipCode under secondary address in Sheet Income Details characters < > are not allowed"" & chr(13)"
            ChkZipcode_Sec = False
            Exit Function
        End If
    Else
        zipCode_Sec = ""
    End If
End Function

'--------------------



'Function ValidateSelect80D() As Boolean
'ValidateSelect80D = True
'    Dim Usr80DVal As Variant
'    Dim status_6 As Variant
'    Dim status_6i As Variant
'    Dim DOB_6 As Variant
'
'
'    status_6 = Sheet1.Range("Sheet1.Status").value
'    status_6 = Mid(status_6, 1, 1)
'
'    status_6i = Sheet1.Range("sheet1.ResidentialStatus1").value
'    status_6i = Mid(status_6i, 1, 2)
'
'    DOB_6 = Sheet1.Range("Sheet1.Dob").value
'
'
'    SELECT80D = Sheet1.Range("SELECT80D").value
'    SELECT80D = Mid(SELECT80D, 1, 1)
'    If isdropdownblank(SELECT80D) Then
'        SELECT80D = ""
'    End If
'
'    If ((SELECT80D = "7" Or SELECT80D = "2") And status_6i = "NR") Then
'        msgError = msgError & "*  ""Option 2&7"" would be allowed to select only if Residential status as ""Resident"" " & Chr(13)
'        Sheet1.Range("SELECT80D").value = "(Select)"
'        ValidateSelect80D = False
'       Exit Function
'    End If
'
'    Usr80DVal = Sheet1.Range("IncD.Section80D").value
'
'    If SELECT80D <> "" Then
'        If IsEmpty(Usr80DVal) Or Usr80DVal = 0 Then
'            msgError = msgError & "* Please enter the value for Part A Sec.80D (Health Insurance Premium)   deduction under Chapter VIA" & Chr(13)
'            ValidateSelect80D = False
'            Exit Function
'        End If
'    End If
'
'    Usr80DVal = IIf(Trim(Usr80DVal) = "", 0, Usr80DVal)
'
'    If status_6 = "I" And Usr80DVal > 0 Then
'        If isdropdownblank(SELECT80D) Then
'            msgError = msgError & "* Please select an option at dropdown of Part A Sec.80D  (Health Insurance Premium)  under Chapter VIA" & Chr(13)
'            ValidateSelect80D = False
'            Exit Function
'        End If
'    End If
'
'    If (SELECT80D = "7" And calculateAge(DOB_6) <= 59) Then
'        msgError = msgError & "* Please select a valid option from the dropdown of Part A Sec.80D (Health Insurance Premium)  under Chapter VIA" & Chr(13)
'        ValidateSelect80D = False
'       Exit Function
'    End If
'
''    If (SELECT80D = "2" And calculateAge(DOB_6) <= 59) Then
''        msgError = msgError & "Please select a valid option from the dropdown of Part A Sec.80D under Chapter VIA" & Chr(13)
''        ValidateSelect80D = False
''       Exit Function
''    End If
'
'
'
'End Function
Function validateDed80TTB() As Boolean


validateDed80TTB = True
If Ded80TTB > 0 Then
 If Ded80TTA > 0 Then
 msgError = msgError & "* If assessee claiming a deduction u/s 80TTB then 80TTA is not allowed." & Chr(13)
        validateDed80TTB = False
       Exit Function
    End If
End If


End Function

'Function ValidateSelect80DB() As Boolean
'ValidateSelect80DB = True
'    Dim Usr80DBVal As Variant
'    Dim status_6 As Variant
'    Dim status_6i As Variant
'    Dim DOB_6 As Variant
'    Dim Temp80DB As Variant
'    Dim Temp80DBCode As Variant
'    Dim Temp80DBNew As Variant
'
'   status_6 = Sheet1.Range("Sheet1.Status").value
'    status_6 = Mid(status_6, 1, 1)
'
'
'    status_6i = Sheet1.Range("sheet1.ResidentialStatus1").value
'    status_6i = Mid(status_6i, 1, 2)
'
'    DOB_6 = Sheet1.Range("Sheet1.Dob").value
'
'
'    SELECT80DB = Sheet1.Range("SELECT80DB").value
'    SELECT80DB = Mid(SELECT80DB, 1, 1)
'
'    If isdropdownblank(SELECT80DB) Then
'        SELECT80DB = ""
'    End If
    
'    If ((SELECT80DB = "3" Or SELECT80DB = "1") And status_6i = "NR") Then
'        msgError = msgError & " ""Option 1&3"" would be allowed to select only if Residential status as ""Resident or Resident but not Ordinarily resident"" " & Chr(13)
'        Sheet1.Range("SELECT80DB").value = "(Select)"
'        ValidateSelect80DB = False
'       Exit Function
'    End If
    
'    Usr80DBVal = Sheet1.Range("IncD.Section80DB").value
'
'    If SELECT80DB <> "" Then
'        If IsEmpty(Usr80DBVal) Or Usr80DBVal = 0 Then
'            msgError = msgError & "* Please enter the value for Part B Sec.80D  (Medical Expenditure) deduction under Chapter VIA" & Chr(13)
'            ValidateSelect80DB = False
'            Exit Function
'        End If
'    End If
'
'    Usr80DBVal = IIf(Trim(Usr80DBVal) = "", 0, Usr80DBVal)
'
'    If status_6 = "I" And Usr80DBVal > 0 Then
'        If isdropdownblank(SELECT80DB) Then
'            msgError = msgError & "* Please select an option at dropdown of  Part B Sec.80D  (Medical Expenditure) under Chapter VIA" & Chr(13)
'            ValidateSelect80DB = False
'            Exit Function
'        End If
'    End If
'
    
    
'    If ((SELECT80DB = "1" Or SELECT80DB = "2" Or SELECT80DB = "3") And calculateAge(DOB_6) <= 59) Then
'        msgError = msgError & "Please select a valid option from the dropdown of Part B Sec.80D  (Medical Expenditure) under Chapter VIA" & Chr(13)
'        ValidateSelect80DB = False
'       Exit Function
'    End If

'    If (SELECT80DB = "3" And calculateAge(DOB_6) < 80) Then
'        msgError = msgError & "Please select a valid option from the dropdown of Part B Sec.80D under Chapter VIA" & Chr(13)
'        ValidateSelect80DB = False
'       Exit Function
'    End If
    
    
    
    
'End Function

'Function ValidateSelect80DC() As Boolean
'ValidateSelect80DC = True
'    Dim Usr80DCVal As Variant
'    Dim status_6 As Variant
'    Dim DOB_6 As Variant
'
'
'    status_6 = Sheet1.Range("Sheet1.Status").value
'    status_6 = Mid(status_6, 1, 1)
'
'    DOB_6 = Sheet1.Range("Sheet1.Dob").value
'
'
'    SELECT80DC = Sheet1.Range("SELECT80DC").value
'    SELECT80DC = Mid(SELECT80DC, 1, 1)
'    If isdropdownblank(SELECT80DC) Then
'        SELECT80DC = ""
'    End If
'
'    Usr80DCVal = Sheet1.Range("IncD.Section80DC").value
'
'    If SELECT80DC <> "" Then
'        If IsEmpty(Usr80DCVal) Or Usr80DCVal = 0 Then
'            msgError = msgError & "* Please enter the value for Part C Sec.80D (Preventive Health check-up) deduction under Chapter VIA" & Chr(13)
'            ValidateSelect80DC = False
'            Exit Function
'        End If
'    End If
    
'    Usr80DCVal = IIf(Trim(Usr80DCVal) = "", 0, Usr80DCVal)
'
'    If status_6 = "I" And Usr80DCVal > 0 Then
'        If isdropdownblank(SELECT80DC) Then
'            msgError = msgError & "* Please select an option at dropdown of  Part C Sec.80D (Preventive Health check-up) under Chapter VIA" & Chr(13)
'            ValidateSelect80DC = False
'            Exit Function
'        End If
'    End If
'
'
'End Function



'Function ValidateSelect80DD() As Boolean
'ValidateSelect80DD = True
'    Dim Usr80DDVal As Variant
'    Dim status_6 As Variant
'
'    status_6 = Sheet1.Range("Sheet1.Status").value
'    status_6 = Mid(status_6, 1, 1)
'
'    SELECT80DD = Sheet1.Range("SELECT80DD").value
'    SELECT80DD = Mid(SELECT80DD, 1, 1)
'    If isdropdownblank(SELECT80DD) Then
'        SELECT80DD = ""
'    End If
'
'    Usr80DDVal = Sheet1.Range("IncD.Section80DD").value
'
'    If SELECT80DD <> "" Then
'        If IsEmpty(Usr80DDVal) Or Usr80DDVal = 0 Then
'            msgError = msgError & "* Please enter the value for 80DD deduction under Chapter VIA" & Chr(13)
'            ValidateSelect80DD = False
'            Exit Function
'        End If
'    End If
'
'    Usr80DDVal = IIf(Trim(Usr80DDVal) = "", 0, Usr80DDVal)
'
'    If status_6 = "I" And Usr80DDVal > 0 Then
'        If isdropdownblank(SELECT80DD) Then
'            msgError = msgError & "* Please select an option at dropdown of Sec.80DD under Chapter VIA" & Chr(13)
'            ValidateSelect80DD = False
'            Exit Function
'        End If
'    End If
'
'End Function

Function ValidateSelect80DDB() As Boolean
ValidateSelect80DDB = True
    Dim Usr80DDBVal As Variant
    Dim status_6 As Variant

    status_6 = Sheet1.Range("Sheet1.Status").value
    status_6 = Mid(status_6, 1, 1)

    SELECT80DDB = Sheet1.Range("SELECT80DDB").value
    SELECT80DDB = Mid(SELECT80DDB, 1, 1)
    If isdropdownblank(SELECT80DDB) Then
        SELECT80DDB = ""
    End If
    
    'Newly added by Bindu
    SELECT80DB_Specified_Disease = Trim(Sheet1.Range("Sheet1.Specified_Disease").value)
    If UCase(Trim(Sheet1.Range("Sheet1.Specified_Disease").value)) <> UCase(Trim("(Select)")) Then
        SELECT80DB_Specified_Disease = Mid(SELECT80DB_Specified_Disease, 2, 3)
    Else
        SELECT80DB_Specified_Disease = Mid(SELECT80DB_Specified_Disease, 1, 1)
    End If
            
    If isdropdownblank(SELECT80DB_Specified_Disease) Then
        SELECT80DB_Specified_Disease = ""
    End If
    
    
    Usr80DDBVal = Sheet1.Range("IncD.Section80DDB").value
    
    If SELECT80DDB <> "" Then
        If IsEmpty(Usr80DDBVal) Or Usr80DDBVal = 0 Then
           ' msgError = msgError & "* Please enter the value for 80DDB deduction under Chapter VIA" & Chr(13)
           msgError = msgError & "* ""Please provide mandatory details under 80DDB deduction.""" & Chr(13)
            ValidateSelect80DDB = False
            Exit Function
        End If
    End If
    
    Usr80DDBVal = IIf(Trim(Usr80DDBVal) = "", 0, Usr80DDBVal)
        
    If status_6 = "I" And Usr80DDBVal > 0 Then
        If isdropdownblank(SELECT80DDB) Then
'            msgError = msgError & "* Please select an option at dropdown of Sec.80DDB under Chapter VIA" & Chr(13)
           'Error Message Changed By Ayush for the SIT-56341
           msgError = msgError & "* Please select an option from the drop down under 80DDB deduction." & Chr(13)
         
            ValidateSelect80DDB = False
            Exit Function
        End If
    End If
    
    
    'SIT-77740
    
    If status_6 = "H" And Usr80DDBVal > 0 Then
        If isdropdownblank(SELECT80DDB) Then
'            msgError = msgError & "* Please select an option at dropdown of Sec.80DDB under Chapter VIA" & Chr(13)

           msgError = msgError & "*""Please select an option from the drop down under 80DDB deduction.""" & Chr(13)
       
            ValidateSelect80DDB = False
            Exit Function
        End If
    End If
    
    'Newly added by Bindu
     Usr80DDBVal = Sheet1.Range("IncD.Section80DDB").value
    
    If SELECT80DB_Specified_Disease <> "" Then
        If IsEmpty(Usr80DDBVal) Or Usr80DDBVal = 0 Then
           ' msgError = msgError & "* Please enter the value for 80DDB deduction under Chapter VIA" & Chr(13)
           msgError = msgError & "* ""Please provide mandatory details under 80DDB deduction.""" & Chr(13)
            ValidateSelect80DDB = False
            Exit Function
        End If
    End If
 'Ayush_17/06
 If Not (SELECT80DB_Specified_Disease = "" Or SELECT80DB_Specified_Disease <> "(Select)") And (IsEmpty(Usr80DDBVal) Or Usr80DDBVal = 0) And (isdropdownblank(SELECT80DDB) Or SELECT80DDB = "") Then
       If SELECT80DB_Specified_Disease <> "" Or SELECT80DB_Specified_Disease <> "(Select)" Then
            If isdropdownblank(SELECT80DDB) Or IsEmpty(Usr80DDBVal) Or Usr80DDBVal = 0 Then
                msgError = msgError & "* ""Please provide mandatory details under 80DDB deduction.""" & Chr(13)
                ValidateSelect80DDB = False
            Exit Function
            End If
       End If
       End If
    
    
    Usr80DDBVal = IIf(Trim(Usr80DDBVal) = "", 0, Usr80DDBVal)
        
    If status_6 = "I" And Usr80DDBVal > 0 Then
        If isdropdownblank(SELECT80DB_Specified_Disease) Then
'
           msgError = msgError & "* ""Name of Specified disease is mandatory""" & Chr(13)
            ValidateSelect80DDB = False
            Exit Function
        End If
    End If
    
    
    'SIT-77740
    
    If status_6 = "H" And Usr80DDBVal > 0 Then
        If isdropdownblank(SELECT80DB_Specified_Disease) Then
           msgError = msgError & "* ""Name of Specified disease is mandatory""" & Chr(13)
            ValidateSelect80DDB = False
            Exit Function
        End If
    End If
    
        'sudha_2025-26
    'For SIT-92091 Modified by sai on 25/04/2025
    If Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = False Then
    
        If (Range("SELECT80DDB").value <> "(Select)" And Range("SELECT80DDB").value <> "") Or Sheet1.Range("IncD.Section80DDB").value > 0 Then
                If Sheet1.Range("Sheet1.Specified_Disease") = "(Select)" Or Sheet1.Range("Sheet1.Specified_Disease") = "" Then
                      msgError = msgError & "* ""Name of Specified disease is mandatory""" & Chr(13)
                      ValidateSelect80DDB = False
                      Exit Function
               End If
        End If
        
    End If

    
    
    
    
    
End Function

'Function ValidateSelect80U() As Boolean
'ValidateSelect80U = True
'    Dim Usr80UVal As Variant
'    Dim status_6 As Variant
'
'    status_6 = Sheet1.Range("Sheet1.Status").value
'    status_6 = Mid(status_6, 1, 1)
'
'    SELECT80U = Sheet1.Range("SELECT80U").value
'    SELECT80U = Mid(SELECT80U, 1, 1)
'    If isdropdownblank(SELECT80U) Then
'        SELECT80U = ""
'    End If
'
'    Usr80UVal = Sheet1.Range("IncD.Section80U").value
'
'    If SELECT80U <> "" Then
'        If IsEmpty(Usr80UVal) Or Usr80UVal = 0 Then
'            msgError = msgError & "* Please enter the value for 80U deduction under Chapter VIA" & Chr(13)
'            ValidateSelect80U = False
'            Exit Function
'        End If
'    End If
'
'    Usr80UVal = IIf(Trim(Usr80UVal) = "", 0, Usr80UVal)
'
'    If status_6 = "I" And Usr80UVal > 0 Then
'        If isdropdownblank(SELECT80U) Then
'            msgError = msgError & "* Please select an option at dropdown of Sec.80U under Chapter VIA" & Chr(13)
'            ValidateSelect80U = False
'            Exit Function
'        End If
'    End If
'
'End Function

Function CheckEmailAddress(emailAddress As Variant, ByRef vType As Variant) As Boolean
On Error Resume Next
    CheckEmailAddress = True
    If Len(emailAddress) > 0 Or Not IsEmpty(emailAddress) Then
    
        If Not CheckSpecialCharacter(Mid(emailAddress, 1, 1)) Then
            vType = "1"
            CheckEmailAddress = False
            Exit Function
        End If
        
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
        
        If Not CheckSpecialCharacter_findconsequtive(emailAddress) Then
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

'Ayush_13/02/2026
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
        If ch = "-" Or ch = "/" Or ch = "_" Or ch = "." Then
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
'---------------------------
'--------------

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
    
    'Ayush
'    If InStr(strEmail, "--") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "__") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, ".-") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "-.") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "._") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "_.") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'
'    If InStr(strEmail, "-_") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "_-") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "//") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "./", "/.") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "/-", "-/") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'     If InStr(strEmail, "/_", "_/") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "/_", "_/") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    'Ayush
'    If InStr(strEmail, ".-", "-.") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "._", "_.") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
'
'    If InStr(strEmail, "_-", "-_") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
    
'     If InStr(strEmail, "_@", "@_") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
    
'    If InStr(strEmail, "-@", "@-") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
    
'    If InStr(strEmail, ".@", "@.") > 0 Then
'        blnIsItValid = False
'        IsValidEmail = blnIsItValid
'        Exit Function
'    End If
    
    
    IsValidEmail = blnIsItValid
End Function


'Ayush
Function CheckSpecialCharacter_findconsequtive(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    Dim M_a_L_L_I As Variant
    M_a_L_L_I = 0
    CheckSpecialCharacter_findconsequtive = True
    specialCharArray = Array("-", "/", "_", ".")
   Dim charcount As Boolean
   Dim d As Variant
   charcount = False
    For iCharCount = 1 To Len(emailid)
           If Mid(emailid, iCharCount, 1) = "-" Or Mid(emailid, iCharCount, 1) = "/" Or _
              Mid(emailid, iCharCount, 1) = "_" Or Mid(emailid, iCharCount, 1) = "." Then
                          M_a_L_L_I = M_a_L_L_I + 1
           Else:
                          M_a_L_L_I = 0
           End If
             If M_a_L_L_I >= 2 And d = iCharCount - 1 Then
                    CheckSpecialCharacter_findconsequtive = False
                    Exit Function
             End If
             d = iCharCount
        Next

End Function

Function ValidateSTDcode(stDcode As Variant) As Variant
    ValidateSTDcode = 0
    Dim i As Long
    
    
    If Len(stDcode) > 5 Then
      ValidateSTDcode = "should be at most 5 digits"
   
    Exit Function
    End If
    
    If Trim(Mid(stDcode, 1, 1)) = 0 Then
        ValidateSTDcode = "* Do not prefix '0' before STD/ISD code"
        
        Exit Function
    End If

    
    If Trim(stDcode) <> "" Then
        For i = 1 To Len(stDcode)
            If Not IsNumeric(Mid(stDcode, i, 1)) Then
                ValidateSTDcode = "* must contain only digits from 0 to 9"
                 'ValidateSTDcode = 1
                Exit Function
            End If
        Next
    End If

End Function

Function ValidatePhoneNo(stDcode As Variant, phoneNo As Variant) As Variant
    Dim i As Long
    
ValidatePhoneNo = 0
If phoneNo <> "" Then
    If Len(phoneNo) > 12 Then
        ValidatePhoneNo = "* should be at most 12 digits"
        Exit Function
    End If

    If Not IsNumeric(phoneNo) Then
        ValidatePhoneNo = "is invalid"
        Exit Function
    End If

    If Mid(phoneNo, 1, 1) = 0 Then
        ValidatePhoneNo = "cannot begin with Zero"
        Exit Function
    End If

    If Trim(stDcode) <> "" Then
        If Trim(phoneNo) = "" Or IsEmpty(phoneNo) Then
                ValidatePhoneNo = "is Compulsory"
                Exit Function
        End If
    End If

    If Trim(phoneNo) <> "" Then
        For i = 1 To Len(phoneNo)
            If Not IsNumeric(Mid(phoneNo, i, 1)) Then
                ValidatePhoneNo = "* must contain only digits from 0 to 9"
                Exit Function
            End If
        Next
    End If

    If (Trim(phoneNo) <> "" Or Trim(stDcode) <> "") Then
'        If (Len(phoneNo) + Len(stDcode)) <> 10 Then
'       ValidatePhoneNo = "* Invalid Phone Number. STD/ISD Code + Landline Number should be 10 digits and cannot begin with '0'."

     'PAG_C2 AY 2024-25 Changes
       'SIT-66361
       'Ayush_01
'         If (Len(phoneNo) + Len(stDcode)) > 12 Then
''          ValidatePhoneNo = "* Invalid Phone Number. STD/ISD Code + Landline Number should be 12 digits and cannot begin with '0'."
'       fmsgbox ("* Invalid Phone Number. STD Code + Landline Number should be upto 12 digits and cannot begin with '0'.")
'     ' ValidatePhoneNo = "* Invalid Phone Number. STD Code + Landline Number should be upto 12 digits and cannot begin with '0'."
'            Exit Function
'        End If
    End If
End If

End Function
Function CheckHousePropertyIncome(HP As Variant) As Boolean
On Error Resume Next
 CheckHousePropertyIncome = True
 If Mid(Range("IncD.TypeOfHP").value, 1, 1) = "S" Then
 If (HP > 200000) Then
' fmsgboxsmall "* In case of Self Occupied Interest payable on borrowed capital cannot exceed 200000 for income from House Property." ', vbOKOnly, "Error(s)!"

  fmsgboxsmall "*""The amount cannot exceed Rs. 2 Lakhs""" ', vbOKOnly, "Error(s)!"

 Range("IncD.IncomeFromHP_v").Select
 Range("IncD.IncomeFromHP_v").value = "" ' No need to add/change the code here in Front end formula restricting for 200000 (Comments from Bindu)
 Exit Function
 End If
 End If
 
 'If Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "S" Then
 'If (hp < -200000) Then
 'MsgBox "Let Out loss cannot exceed 200000 for income from House Property.", vbOKOnly, mdCommon.vbMessageCaption
 'Range("IncD.IncomeFromHP").Select
' Range("IncD.IncomeFromHP").value = ""
 'Exit Function
' End If
' End If
 
' If (hp < 0) Then
'    hp = hp * -1
'
'        If (hp > IncSal) Then
'        CheckHousePropertyIncome = False
'        End If
'        Exit Function
' End If
 
End Function

Function CheckMobileNo(num As Variant) As Boolean
CheckMobileNo = True

If IsNumeric(num) Then
    If num < 10000 Or num > 9999999999# Then
        CheckMobileNo = False
    End If
Else
    CheckMobileNo = False
End If

End Function
Function checkfieldspecialcharacterNoticeNo(field As Variant) As Boolean
    checkfieldspecialcharacterNoticeNo = True
    Dim i, j As Long
    Dim arr As Variant
    
    arr = Array("@", "*", ".", "!", "-", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", " ")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacterNoticeNo = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateNoticeNo_1(item As Variant) As Boolean
    ValidateNoticeNo_1 = True
    errmsgfiling = ""
'Change.02.03.2023.102.IDS.76
'    If Len(item) > 30 Then
'        errmsgfiling = "* Notice Number  should not exceed 30 characters"
    If Len(item) > 100 Then
        errmsgfiling = "* Notice Number  should not exceed 100 characters"
'End Change.IDS.76
        ValidateNoticeNo_1 = False
        Exit Function
    End If
    
    If Not checkfieldspecialcharacterNoticeNo(item) Then
        errmsgfiling = "* Unique Number is Invalid in sheet  Income Details. No Special characters other than  ""/"" is allowed"
        ValidateNoticeNo_1 = False
        Exit Function
    End If
End Function

Function ValidateNoticeNo_2(item As Variant) As Boolean
    ValidateNoticeNo_2 = True
    errmsgfiling = ""
    If Len(item) > 23 Then
        errmsgfiling = "* Unique Number should not exceed 23 characters"
        ValidateNoticeNo_2 = False
        Exit Function
    End If
    
    If Not checkfieldspecialcharacterNoticeNo(item) Then
        errmsgfiling = "* Unique Number is Invalid in sheet Income Details. No Special characters other than  ""/"" is allowed"
        ValidateNoticeNo_2 = False
        Exit Function
    End If
End Function

Function CheckPANIncomeDetails_Spouse(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPANIncomeDetails_Spouse = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
                
       If Not (Mid(PAN, 4, 1) = "P") Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacterdot(PAN) Then
            CheckPANIncomeDetails_Spouse = False
            Exit Function
        End If
        
        
               
    End If
End Function

Function CheckPAN_Tenant(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPAN_Tenant = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
                
       If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(PAN, 7, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(PAN, 8, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(PAN, 9, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN_Tenant = False
            Exit Function
        End If
               
    End If
End Function

Function CheckTAN(TAN As Variant) As Boolean
On Error Resume Next


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
Sub UNLOCKPI()
    Application.ScreenUpdating = False
    Sheet1.Unprotect getmsgstate


    Dim rngname_5_AL As Variant
'    rngname_5_AL = "IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||IncomeNotifiedOther89A||Sheet1.Nature||Sheet1.Amount||IncD.Section||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE||IncD.Section80CCD||IncD.Section80GG||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||Increliefus89A||OSIncomeNotified89A_AmountUS||OSIncomeNotified89A_AmountUK||OSIncomeNotified89A_AmountCan||OSIncomeNotifiedOther89A||OSIncreliefus89A||IncD_q1OS1||IncD_q2OS1||IncD_q3OS1||IncD_q4OS1||IncD_q5OS1"
'     rngname_5_AL = "sheet1.FirstName||sheet1.MiddleName||IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||IncomeNotifiedOther89A||Sheet1.Nature||Sheet1.Amount||IncD.Section80CCC||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE||IncD.Section80CCD||IncD.Section80GG||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||Increliefus89A||OSIncomeNotified89A_AmountUS||OSIncomeNotified89A_AmountUK||OSIncomeNotified89A_AmountCan||OSIncomeNotifiedOther89A||OSIncreliefus89A||IncD_q1OS1||IncD_q2OS1||IncD_q3OS1||IncD_q4OS1||IncD_q5OS1"
        rngname_5_AL = "sheet1.FirstName||sheet1.MiddleName||IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||Sheet1.Nature||Sheet1.Amount||IncD.Section80CCD||IncD.Section80GG||OSIncreliefus89A||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE"
    UNLOCKRANGEPI (rngname_5_AL)
    
    'Ayush_17/06
    rngname_5_AL = "Sheet1.AckNum"
    UNLOCKRANGEPI (rngname_5_AL)
    
    Sheet1.Protect getmsgstate
    Application.ScreenUpdating = True
End Sub

Sub LOCKPI()
    Application.ScreenUpdating = False

    Sheet1.Unprotect getmsgstate

    Dim rngname_5_AL As Variant
'      rngname_5_AL = "sheet1.FirstName||sheet1.MiddleName||Sheet1.Description||IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||IncomeNotifiedOther89A||Sheet1.Nature||Sheet1.Amount||IncD.Section80CCC||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE||IncD.Section80CCD||IncD.Section80GG||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||Increliefus89A||OSIncomeNotified89A_AmountUS||OSIncomeNotified89A_AmountUK||OSIncomeNotified89A_AmountCan||OSIncomeNotifiedOther89A||OSIncreliefus89A||IncD_q1OS1||IncD_q2OS1||IncD_q3OS1||IncD_q4OS1||IncD_q5OS1"
'        rngname_5_AL = "sheet1.FirstName||sheet1.MiddleName||Sheet1.Description||IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||IncomeNotifiedOther89A||Sheet1.Nature||Sheet1.Amount||IncD.Section80CCC||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE||IncD.Section80CCD||IncD.Section80GG||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||Increliefus89A||OSIncomeNotified89A_AmountUS||OSIncomeNotified89A_AmountUK||OSIncomeNotified89A_AmountCan||OSIncomeNotifiedOther89A||OSIncreliefus89A||IncD_q1OS1||IncD_q2OS1||IncD_q3OS1||IncD_q4OS1||IncD_q5OS1"
'      rngname_5_AL = "Sheet1.Description||IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||IncomeNotifiedOther89A||Sheet1.Nature||Sheet1.Amount||IncD.Section||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE||IncD.Section80CCD||IncD.Section80GG||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||Increliefus89A||OSIncomeNotified89A_AmountUS||OSIncomeNotified89A_AmountUK||OSIncomeNotified89A_AmountCan||OSIncomeNotifiedOther89A||OSIncreliefus89A||IncD_q1OS1||IncD_q2OS1||IncD_q3OS1||IncD_q4OS1||IncD_q5OS1"
        rngname_5_AL = "sheet1.FirstName||sheet1.MiddleName||Sheet1.Description||IncD.IncomeFromSal_iva||IncD.IncomeFromSal_ivc||IncD.IncomeFromSal_ii||IncD.IncomeFromSal_iii||IncD.IncomeFromSal_v||IncD.IncomeFromSal_iv||Sheet1.Nature||Sheet1.Amount||IncD.Section80CCD||IncD.Section80GG||OSIncreliefus89A||IncD.Section80CCD_SE||IncD.Section80CCD1B_SE"
    LOCKRANGEPI (rngname_5_AL)
    
   'Ayush_17/06
    rngname_5_AL = "Sheet1.AckNum"
    LOCKRANGEPI1 (rngname_5_AL)

   Sheet1.Protect getmsgstate

    Application.ScreenUpdating = True
End Sub
Sub LOCKEE()
    Application.ScreenUpdating = False

    Sheet1.Unprotect getmsgstate

    Dim rngname_5_AL As Variant
    rngname_5_AL = "IncD.Section80E||IncD.Section80EE||IncD.Section80EEB||IncD.Section80EEA"
    LOCKRANGEPI (rngname_5_AL)

   Sheet1.Protect getmsgstate

    Application.ScreenUpdating = True
End Sub

Sub LOCKRANGEPI(rangenamestring As Variant)
Application.EnableEvents = False

Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet1.Range(rangenamestring(i))
            chcell.MergeArea.Locked = True
            chcell.MergeArea.Interior.Color = (&HD8D8D8)
            'chcell.MergeArea.Font.Color = (&HD8D8D8)
            chcell.MergeArea.ClearContents
            Next
        Next
        
Application.EnableEvents = True
End Sub

Sub LOCKRANGEPI1(rangenamestring As Variant)
Application.EnableEvents = False

Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet1.Range(rangenamestring(i))
            chcell.MergeArea.Locked = True
            chcell.MergeArea.Interior.Color = (&HD8D8D8)
            chcell.MergeArea.Font.Color = (&HD8D8D8)
            'chcell.MergeArea.ClearContents
            Next
        Next
        
Application.EnableEvents = True
End Sub

Sub UNLOCKRANGEPI(rangenamestring As Variant)
Application.EnableEvents = False
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet1.Range(rangenamestring(i))
            If rangenamestring(i) = "IncD.IncomeFromSal_iva" Then
            
            'PAG_C35 AY 2024-25 Change
            'chcell.MergeArea.Formula = "=MIN(IncD.IncomeFromSal_net,50000)"
            
           chcell.MergeArea.Formula = "=IF(Return_Dropdown_Regime=1,MIN(IncD.IncomeFromSal_net,75000),IF(Return_Dropdown_Regime=2,MIN(IncD.IncomeFromSal_net,50000),0))"
           
            
            
            ''PAG_C35 AY 2023-24 Old
'            chcell.MergeArea.Formula = "=IF(bacValue=1,0,MIN(IncD.IncomeFromSal_net,50000))"
            
            chcell.MergeArea.Interior.ColorIndex = 2
             chcell.MergeArea.Locked = True
        Else
            chcell.MergeArea.Locked = False
            chcell.MergeArea.Interior.Color = (&HCCFFCC)
            'chCell.MergeArea.ClearContents
            End If
            Next
        Next
 Application.EnableEvents = True
End Sub



Function ValidateSystem80D() As Boolean
ValidateSystem80D = True
    Dim Usr80DVal As Variant
    Dim Usr80DBVal As Variant
    Dim Usr80DCVal As Variant
    Dim SYSTEM80DB As Variant
    Dim HUFStatus As Variant
    
    Usr80DVal = Sheet1.Range("IncD.Section80DValu").value
    Usr80DBVal = Sheet1.Range("IncD.Section80DB").value
    Usr80DCVal = Sheet1.Range("IncD.Section80DC").value
    SYSTEM80DB = Sheet1.Range("IncD.Section80D_Calc").value
    HUFStatus = Mid(Sheet1.Range("sheet1.Status").value, 1, 1)
    
    Usr80DVal = IIf(Trim(Usr80DVal) = "", 0, Usr80DVal)
    Usr80DBVal = IIf(Trim(Usr80DBVal) = "", 0, Usr80DBVal)
    Usr80DCVal = IIf(Trim(Usr80DCVal) = "", 0, Usr80DCVal)
    SYSTEM80DB = IIf(Trim(SYSTEM80DB) = "", 0, SYSTEM80DB)
        
    If SYSTEM80DB > (Usr80DVal + Usr80DBVal + Usr80DCVal) Then
        msgError = msgError & "* Deduction u/s 80D should not be more than sum of amount claimed at 'Health Insurance, Medical Expenditure and Preventive Health Check Up' under Chapter VIA in Income Details" & Chr(13)
        ValidateSystem80D = False
        Exit Function
    End If
    
    If HUFStatus = "H" Then
        If SYSTEM80DB > 50000 Then
            msgError = msgError & "* Deduction u/s 80D should not be more than Rs. 50000 under Chapter VIA in Income Details" & Chr(13)
            ValidateSystem80D = False
            Exit Function
        End If
    End If
    
End Function


'Sub reset10IE()
'On Error Resume Next
'Sheet8.Unprotect Password:=getmsgstate
'Sheet1.Unprotect Password:=getmsgstate
''If (Sheet1.Range("OptingNewTaxRegime").value = "No" And Sheet8.Range("BacValue").value = "1") Or (Sheet1.Range("OptingNewTaxRegime").value = "Yes" And Sheet8.Range("BacValue").value = "2") Or (Sheet1.Range("OptingNewTaxRegime").value = "Return Not filed Previous AY" And Sheet8.Range("BacValue").value = "1") Then
'' If (Sheet8.Range("OldBacValue").value = 1 And Sheet8.Range("BacValue").value = 2) Or (Sheet8.Range("OldBacValue").value = 2 And Sheet8.Range("BacValue").value = 1) Then
''
''    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
''    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
''    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
''    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
''
''
''    Else
''    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
''    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
''    Sheet1.Range("Sheet1.date10IE").value = ""
''    Sheet1.Unprotect Password:=getmsgstate
''    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
''    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
''    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'' End If
'
'If (Sheet8.Range("OldBacValue").value = 2 And Sheet8.Range("BacValue").value = 2) Or Sheet1.Range("sheet1.Status") = "F - FIRM(Other than LLP)" Then
'
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.date10IE").value = ""
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'
'
'    Else
'    Sheet1.Range("Sheet1.date10IE").value = ""
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
'
'
' End If
'
'    Sheet8.Protect Password:=getmsgstate
'    Sheet1.Protect Password:=getmsgstate
'End Sub


'Sub BacYesValueChange()
'On Error Resume Next
'Application.EnableEvents = False
'Dim answer As Integer
''answer = MsgBox("if you have opted for new tax Regime in AY 2021-22, please confirm ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") dpk1201
'answer = MsgBox("if you have opted for new tax Regime in AY 2022-23, please confirm ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
'If answer <> vbYes Then
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("oldbacValue").Clear
'    Sheet8.Protect Password:=getmsgstate
'    ActiveSheet.Shapes("BacYes").OLEFormat.Object.value = 0
'    ActiveSheet.Shapes("BacNO").OLEFormat.Object.value = 0
'    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'Else
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("oldbacValue").value = 1
'    Sheet8.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'End If



'Change-30.11.2022.103.09A
'reset10IE
'reset10IEAy23_24

'Application.EnableEvents = True
'Sheet1.Activate
'End Sub

Sub NewBacYesValueChange()
On Error Resume Next
Application.EnableEvents = False

If Sheet8.Range("oldbacValue").value = "" Then
Sheet8.Unprotect Password:=getmsgstate
Sheet8.Range("bacValue").Clear
Sheet8.Protect Password:=getmsgstate
Sheet1.Range("OptingNewTaxRegime").value = "(Select)"

'MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2021-22?" dpk1201
MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2024-25?"
Application.EnableEvents = True
Exit Sub
End If


Dim answer As Integer
 If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
      Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
      fmsgboxsmall "* In case of Firm, New tax regime u/s 115BAC is not applicable."
      Application.EnableEvents = True
      Exit Sub
 End If
 
'answer = MsgBox("Since you are opting in Now or continue to opt for New tax Regime u/s 115BAC. You will not be eligible to set-off House property loss and claim following deduction/allowances. Do you want to continue?" & vbCrLf & "1) Certain allowances u/s section 10 (LTA, HRA, allowances granted to meet expenses in performance of duties of office, Allowances granted to meet personal expenses in performance of duties of office, Allowance received by MP/MLA/MLC) " & vbCrLf & "2) Deductions u/s 16 (Standard Deduction ,Entertainment allowance and Professional tax) " & vbCrLf & "3) Interest payable on borrowed capital for self occupied property " & vbCrLf & "4) Standard Deduction in case of family pension " & vbCrLf & "5) Chapter VIA Deduction (life insurance, health insurance premium, pension funds, provident fund,donation etc except Contribution made by employer to notified pension scheme u/s 80CCD(2))", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
'answer = MsgBox("Since you are opting New tax Regime u/s 115BAC. You will not be eligible to set-off House property loss and claim following deduction/allowances. Do you want to continue?" & vbCrLf & "1) Certain allowances u/s section 10 (LTA, HRA, allowances granted to meet expenses in performance of duties of office, Allowances granted to meet personal expenses in performance of duties of office, Allowance received by MP/MLA/MLC) " & vbCrLf & "2) Deductions u/s 16 (Standard Deduction ,Entertainment allowance and Professional tax) " & vbCrLf & "3) Interest payable on borrowed capital for self occupied property " & vbCrLf & "4) Standard Deduction in case of family pension " & vbCrLf & "5) Chapter VIA Deduction (life insurance, health insurance premium, pension funds, provident fund,donation etc except Contribution made by employer to notified pension scheme u/s 80CCD(2))", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

answer = MsgBox("Please be noted that, You will not be eligible to set-off House property loss and nor eligible to claim following deduction/allowances under ""default tax regime"". Do you want to continue?" & vbCrLf & "1) Certain allowances u/s section 10 (LTA, HRA, allowances granted to meet expenses in performance of duties of office, Allowances granted to meet personal expenses in performance of duties of office, Allowance received by MP/MLA/MLC) " & vbCrLf & "2) Deductions u/s 16 such as Entertainment allowance and Professional tax " & vbCrLf & "3) Interest payable on borrowed capital for self occupied property " & vbCrLf & "4) Chapter VIA Deduction (life insurance, health insurance premium, pension funds, provident fund,donation etc except Contribution made by employer to notified pension scheme u/s 80CCD(2))")


'If answer = vbYes Then
ThisWorkbook.Unprotect Password:=getmsgstate

   '===============Commented by sai on 22/07/2024 for After due date issue for enable/ disable 80 Schedules

'    If Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
'        'fmsgbox "* Please select ""Have you opted  for new tax regime u/s 115BAC and filed Form 10IE in AY 2021-22 ? ""  as n since you have filed return later."
''    If Sheet8.Range("oldbacValue") <> 1 Then
''        MsgBox "Please select Option for current assessment year - Not opting "
''      Else
''        MsgBox "Please select Option for current assessment year - Opt out "
''    End If
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").Clear
'    Sheet8.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
''    ActiveSheet.Shapes("BacYes").OLEFormat.Object.value = 0
''    ActiveSheet.Shapes("BacNO").OLEFormat.Object.value = 0
'     Application.EnableEvents = True
'         Exit Sub
'    End If

'=====================================================================================================

  Sheet16.Unprotect Password:=getmsgstate
  Sheet16.Visible = xlSheetHidden
  Sheet16.Protect Password:=getmsgstate
  Sheet6.Unprotect Password:=getmsgstate
  Sheet6.Visible = xlSheetHidden
  Sheet6.Protect Password:=getmsgstate
  
  'AY 2024-25 Sheet(80GGC) implementation visible/unvisible Starts------------
  Sheet11.Unprotect Password:=getmsgstate
  Sheet11.Visible = xlSheetHidden
  Sheet11.Protect Password:=getmsgstate
  'AY 2024-25 Sheet(80GGC) implementation visible/unvisible End------------
  
  'New Schedule 80DD-80U Added by Sai
  Sheet14.Unprotect Password:=getmsgstate
  Sheet14.Visible = xlSheetHidden
  Sheet14.Protect Password:=getmsgstate
  
  
  
  Sheet8.Unprotect Password:=getmsgstate
  Sheet8.Range("bacValue").value = 1
  Sheet8.Protect Password:=getmsgstate
  Sheet1.Unprotect Password:=getmsgstate
  resetBacYes
  Sheet1.Activate
  Sheet1.Protect Password:=getmsgstate
  ThisWorkbook.Protect Password:=getmsgstate
'Else
'    Sheet8.Unprotect Password:=getmsgstate
'    Sheet8.Range("bacValue").Clear
'    Sheet8.Protect Password:=getmsgstate
'    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
''    ActiveSheet.Shapes("BacYes").OLEFormat.Object.value = 0
''     ActiveSheet.Shapes("BacNO").OLEFormat.Object.value = 0
'      Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU" sai
'    Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet8.Unprotect Password:=getmsgstate
'
'    reset10IE
    
    
' If (Sheet1.Range("OptingNewTaxRegime").value = "No" And Sheet8.Range("BacValue").value = "1") Or (Sheet1.Range("OptingNewTaxRegime").value = "Yes" And Sheet8.Range("BacValue").value = "2") Or (Sheet1.Range("OptingNewTaxRegime").value = "Return Not filed Previous AY" And Sheet8.Range("BacValue").value = "1") Then
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
'
'
'    Else
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.date10IE").value = ""
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'    End If
'    Sheet1.Protect Password:=getmsgstate
'    Sheet8.Protect Password:=getmsgstate
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    resetIncomeDetails
'     Sheet1.Activate
'     Sheet1.Protect Password:=getmsgstate
'     Application.EnableEvents = True
'  Exit Sub

'End If

'Ayush_02
'If (Sheet1.Range("Sheet1.OptOutNewTaxRegime") = "No") Then
'            Sheet1.Unprotect Password:=getmsgstate
'
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.value = ""
'            Sheet1.Unprotect Password:=getmsgstate
'
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.value = ""
'
'            Sheet1.Protect Password:=getmsgstate
'
'End If

Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True
End Sub
Sub BacNoValueChange()
On Error Resume Next
Application.EnableEvents = False
Dim answer As Integer
'answer = MsgBox("if you have not opted for new tax Regime in AY 2021-22, please confirm ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") dpk1201
answer = MsgBox("if you have not opted for new tax Regime in AY 2022-23, please confirm ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer <> vbYes Then
    Sheet8.Unprotect Password:=getmsgstate
    Sheet8.Range("oldbacValue").Clear
    Sheet8.Protect Password:=getmsgstate
    ActiveSheet.Shapes("BacYes").OLEFormat.Object.value = 0
    ActiveSheet.Shapes("BacNO").OLEFormat.Object.value = 0
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
    Else
    Sheet8.Unprotect Password:=getmsgstate
    Sheet8.Range("oldbacValue").value = 2
    Sheet8.Protect Password:=getmsgstate
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
End If


'Change-30.11.2022.103.09B
'reset10IE
reset10IEAy23_24
'Change-01.12.2022.103.16A
'CurrentYearOptionAY23_24 ' AY 2023-24 Old

Application.EnableEvents = True
Sheet1.Activate
End Sub
'-----------added by sai ---------------------
Sub NewBacNoValueChange_2425()
On Error Resume Next
Application.EnableEvents = False

If Sheet8.Range("oldbacValue").value = "" Then
    Sheet8.Unprotect Password:=getmsgstate
    Sheet8.Range("bacValue").Clear
    Sheet8.Protect Password:=getmsgstate
    Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
    'MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2021-22?" dpk1201
    'MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2022-23?"
    Application.EnableEvents = True
    Exit Sub
End If


ThisWorkbook.Unprotect Password:=getmsgstate
 
  Sheet16.Unprotect Password:=getmsgstate
  Sheet16.Visible = xlSheetVisible
  Sheet16.Protect Password:=getmsgstate
  
  
 
  'AY 2024-25 Sheet(80GGC) implementation visible/unvisible Starts------------------
  Sheet11.Unprotect Password:=getmsgstate
  Sheet11.Visible = xlSheetVisible
  Sheet11.Protect Password:=getmsgstate
  'AY 2024-25 Sheet(80GGC) implementation visible/unvisible End------------------
  
    'New Schedule 80DD-80U Added by Sai
      Sheet14.Unprotect Password:=getmsgstate
      Sheet14.Visible = xlSheetVisible
      Sheet14.Protect Password:=getmsgstate
  
  
  Sheet6.Unprotect Password:=getmsgstate
  Sheet6.Visible = xlSheetVisible
  Sheet6.Protect Password:=getmsgstate
  Sheet8.Unprotect Password:=getmsgstate
  Sheet8.Range("bacValue").value = 2
  Sheet8.Protect Password:=getmsgstate
  Sheet1.Unprotect Password:=getmsgstate
  
   
  
  
'   Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False

  ' Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HD8D8D8)  ', "AU"
   'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = True
   
   

  resetBacNo
  Sheet1.Protect Password:=getmsgstate
 
  Sheet1.Activate
  ThisWorkbook.Protect Password:=getmsgstate

Application.EnableEvents = True
End Sub
'---------------------------------------------------------
Sub NewBacNoValueChange()
On Error Resume Next
Application.EnableEvents = False

If Sheet8.Range("oldbacValue").value = "" Then
Sheet8.Unprotect Password:=getmsgstate
Sheet8.Range("bacValue").Clear
Sheet8.Protect Password:=getmsgstate
Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2021-22?" dpk1201
MsgBox "Please Select the option in Have you opted for new tax regime u/s 115BAC and filed Form 10IE  in AY 2022-23?"
Application.EnableEvents = True
Exit Sub
End If

Dim answer As Integer
'need to remove ism
answer = MsgBox("Do you want to change value of Option for current assessment year? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
ThisWorkbook.Unprotect Password:=getmsgstate

Sheet16.Unprotect Password:=getmsgstate
  Sheet16.Visible = xlSheetVisible
  Sheet16.Protect Password:=getmsgstate
  Sheet6.Unprotect Password:=getmsgstate
  Sheet6.Visible = xlSheetVisible
  Sheet6.Protect Password:=getmsgstate
  Sheet8.Unprotect Password:=getmsgstate
  Sheet8.Range("bacValue").value = 2
  Sheet8.Protect Password:=getmsgstate
  Sheet1.Unprotect Password:=getmsgstate
   Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
  resetBacNo
  Sheet1.Protect Password:=getmsgstate

  Sheet1.Activate
  ThisWorkbook.Protect Password:=getmsgstate
Else
Sheet8.Unprotect Password:=getmsgstate
Sheet8.Range("bacValue").Clear
Sheet8.Protect Password:=getmsgstate
Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
      Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'Sheet1.Unprotect Password:=getmsgstate
' Sheet8.Unprotect Password:=getmsgstate

'reset10IE

' If (Sheet1.Range("OptingNewTaxRegime").value = "No" And Sheet8.Range("BacValue").value = "1") Or (Sheet1.Range("OptingNewTaxRegime").value = "Yes" And Sheet8.Range("BacValue").value = "2") Or (Sheet1.Range("OptingNewTaxRegime").value = "Return Not filed Previous AY" And Sheet8.Range("BacValue").value = "1") Then
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
'
'
'    Else
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'
'    Sheet1.Range("Sheet1.date10IE").value = ""
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'    End If
'    Sheet1.Protect Password:=getmsgstate
'    Sheet8.Protect Password:=getmsgstate
    resetIncomeDetails
     Sheet1.Activate
     Application.EnableEvents = True
  Exit Sub
End If
Application.EnableEvents = True
End Sub
Sub LockUnlock80TTBNew()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate

Dim DOB_6 As Variant
DOB_6 = Sheet1.Range("sheet1.DOB").value
If calculateAge(DOB_6) <= 59 Then

            Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
            Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
            Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HCCFFCC"

ElseIf calculateAge(DOB_6) > 59 And calculateAge(DOB_6) <> 2022 Then

            Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
            Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"


End If

If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Or (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") Then

        Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"

ElseIf (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") Then
        Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"

Else
        Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HCCFFCC"


End If





If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F") Then
        Sheet1.Range("IncD.Section80TTA").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HD8D8D8"

Else
        Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HCCFFCC"


End If



    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then



                 Sheet1.Range("IncD.Section80TTA").value = ""
                 Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
                 Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)


'
'
'                 Else
'                 Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
'
'

                End If



    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then

                      Sheet1.Range("IncD.Section80TTB").value = ""
                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)


                End If

If calculateAge(Trim(Range("sheet1.DOB").value)) <= 59 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then

                      Sheet1.Range("IncD.Section80TTB").value = ""
                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)


                End If



Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
Sub resetIncomeDetails()
Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False

    'Sheet1.Range("IncD.IncomeFromHP_v").Interior.Color = (&HCCFFCC)  ', "AU"
    'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = False ' Commented by Bindu on 23rd April 25 as those fields are auto-populate
            If Mid(Trim(Range("sheet1.Status")), 1, 1) = "I" Then
            Sheet1.Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
            Sheet1.Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = False
            Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HCCFFCC)
            End If
            
'     Sheet1.Range("IncD.Section80C").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80C").MergeArea.Locked = False
    
    Sheet1.Range("IncD.Section80C").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80C").Formula = "=IF(BacValue=1,0,TotAmount.80C)"
    
    Sheet1.Range("IncD.Section80CCC").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80CCC").Formula = "=IF(BacValue=1,0,sum_80CCC)"
    
'Ayush_16/06/2026
'    Sheet1.Range("IncD.Section80CCD_SE").Interior.Color = vbWhite  ', "AU"
'    Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
'    Sheet1.Range("IncD.Section80CCD_SE").Formula = "=IF(BacValue=1,0,sum_80CCD1)"
'
'    Sheet1.Range("IncD.Section80CCD1B_SE").Interior.Color = vbWhite  ', "AU"
'    Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = True
'    Sheet1.Range("IncD.Section80CCD1B_SE").Formula = "=IF(BacValue=1,0,sum_80CCD1b)"
'
   
   'Malli<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
   'Ayush_05/01/2025
    If Mid(Trim(Range("sheet1.Status")), 1, 1) = "I" Then
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Type").Interior.Color = (&HCCFFCC)
     Sheet1.Range("Sec80CCC_Type").ClearContents
    Sheet1.Range("Sec80CCC_Type").Locked = False

    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Name").Interior.Color = (&HCCFFCC)
    Sheet1.Range("Sec80CCC_Name").ClearContents
    Sheet1.Range("Sec80CCC_Name").Locked = False

    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Amount").Interior.Color = (&HCCFFCC)
    Sheet1.Range("Sec80CCC_Amount").ClearContents
    Sheet1.Range("Sec80CCC_Amount").Locked = False
'Ayush_16/04/2026
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Type").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Sec80CCC_1_Type").ClearContents
'    Sheet1.Range("Sec80CCC_1_Type").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Name").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Sec80CCC_1_Name").ClearContents
'    Sheet1.Range("Sec80CCC_1_Name").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Amount").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Sec80CCC_1_Amount").ClearContents
'    Sheet1.Range("Sec80CCC_1_Amount").Locked = False
'
'    Sheet1.Range("Sec80CCC_1b_Type").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Sec80CCC_1b_Type").ClearContents
'    Sheet1.Range("Sec80CCC_1b_Type").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1b_Name").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Sec80CCC_1b_Name").ClearContents
'    Sheet1.Range("Sec80CCC_1b_Name").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1b_Amount").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Sec80CCC_1b_Amount").ClearContents
'    Sheet1.Range("Sec80CCC_1b_Amount").Locked = False
    
    Else
    
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Type").Interior.Color = (&HD8D8D8)
    Sheet1.Range("Sec80CCC_Type").Locked = True

    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Name").Interior.Color = (&HD8D8D8)
    Sheet1.Range("Sec80CCC_Name").Locked = True

    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Amount").Interior.Color = (&HD8D8D8)
    Sheet1.Range("Sec80CCC_Amount").Locked = True
'Ayush_16/04/2026
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Type").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sec80CCC_1_Type").Locked = True
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Name").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sec80CCC_1_Name").Locked = True
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Amount").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sec80CCC_1_Amount").Locked = True
'
'    Sheet1.Range("Sec80CCC_1b_Type").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sec80CCC_1b_Type").Locked = True
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1b_Name").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sec80CCC_1b_Name").Locked = True
'
'    Sheet1.Unprotect Password:=getmsgstate
'
'    Sheet1.Range("Sec80CCC_1b_Amount").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sec80CCC_1b_Amount").Locked = True
    End If
    '----------------
   '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
    
    
'    Sheet1.Range("IncD.Section80CCC").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = False

    
'    Sheet1.Range("IncD.Section80CCC").Interior.Color = vbWhite  ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
'    Sheet1.Range("IncD.Section80CCC").Formula = "=IF(BacValue=1,0,TotAmount.)"
    
    
 'Malli_AY_2026_27 SIT-118203   27/04/2026
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
'Ayush_16/06/2026
     Sheet1.Range("IncD.Section80CCD_SE").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = False
     Sheet1.Range("IncD.Section80CCD1B_SE").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = False
    '----------
 End If
 '-----------------------------------
    
     Sheet1.Range("SELECT80DD").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet1.Range("SELECT80DD").MergeArea.Locked = False
'    Sheet1.Range("SELECT80DD").Value = "(Select)"
     'Sheet1.Range("IncD.Section80DD").Interior.ColorIndex = 2
    'Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
    Sheet1.Range("SELECT80DDB").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet1.Range("SELECT80DDB").MergeArea.Locked = False
    
    
'sudha_2025-26
    Sheet1.Range("Sheet1.Specified_Disease").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = False
    
    '===========================
    
'    Sheet1.Range("SELECT80DDB").Value = "(Select)"
    Sheet1.Range("IncD.Section80DDB").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = False
     
'    Sheet1.Range("IncD.Section80E").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80E").MergeArea.Locked = False
    
    Sheet1.Range("IncD.Section80E").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80E").Formula = "=IF(BacValue=1,0,TotAmt.80E)"
    
    
'     Sheet1.Range("IncD.Section80EE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = False
    
    
    Sheet1.Range("IncD.Section80EE").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80EE").Formula = "=IF(BacValue=1,0,TotAmt.80EE)"
    
    
'     Sheet1.Range("IncD.Section80EEA").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = False
    
    Sheet1.Range("IncD.Section80EEA").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80EEA").Formula = "=IF(BacValue=1,0,TotAmt.80EEA)"
    
    
    
'     Sheet1.Range("IncD.Section80EEB").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = False
    
    
    Sheet1.Range("IncD.Section80EEB").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80EEB").Formula = "=IF(BacValue=1,0,TotAmt.80EEB)"
    
    'Ayush_20/05
    If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
            Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80E").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section80E").value = ""  'Malli_SIT-118203
            
            Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section80EE").value = ""   'Malli_SIT-118203
            
            Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section80EEA").value = ""   'Malli_SIT-118203
            
            Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80EEB").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section80EEB").value = ""   'Malli_SIT-118203
    End If
    '-------------------
    
    
    If Not Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
        Sheet1.Range("IncD.Section80GG").Interior.Color = (&HCCFFCC)  ', "AU"
        Sheet1.Range("IncD.Section80GG").MergeArea.Locked = False
        
        'sudha_2025-25
        Sheet1.Range("Sheet1.AckNum").Interior.Color = (&HCCFFCC)  ', "AU"
        Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AckNum").Font.Color = vbBlack
        
  '=======================
  
    End If
    
    
    'Ayush_23/05
    If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("IncD.Section80GG").Interior.Color = (&HD8D8D8)  ', "AU"
        Sheet1.Range("IncD.Section80GG").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80GG").MergeArea.ClearContents
        
        'Ayush_Legacy_02/06/2025
        
        Sheet1.Unprotect Password:=getmsgstate
       Sheet1.Range("IncD.Section80CCC").Interior.Color = vbWhite  ', "AU"
      Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
      Sheet1.Range("IncD.Section80CCC").Formula = "=IF(BacValue=1,0,sum_80CCC)"
'Ayush_16/06/2026
'     Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("IncD.Section80CCD_SE").Interior.Color = vbWhite  ', "AU"
'      Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
'      Sheet1.Range("IncD.Section80CCD_SE").Formula = "=IF(BacValue=1,0,sum_80CCD1)"
        
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("IncD.Section80CCD_SE").Interior.Color = (&HD8D8D8)  ', "AU"
'        Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
'        Sheet1.Range("IncD.Section80CCD_SE").MergeArea.ClearContents
'Ayush_16/06/2026
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("IncD.Section80CCD1B_SE").Interior.Color = vbWhite  ', "AU"
'        Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = True
'        Sheet1.Range("IncD.Section80CCD1B_SE").Formula = "=IF(BacValue=1,0,sum_80CCD1b)"
'
        
        'Ayush_02/04/2026
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNum").Font.ColorIndex = 15
        Sheet1.Range("Sheet1.AckNum").MergeArea.ClearContents
        Sheet1.Range("Sheet1.AckNum").Interior.Color = (&HD8D8D8)  ', "AU"
        Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
      '  Sheet1.Range("Sheet1.AckNum").FormulaHidden = True
        
      '  Sheet1.Range("Sheet1.AckNum").MergeArea.ClearContents
        
  End If
    
    
    'MsgBox Sheet1.Range("FP_Value").value
    If Sheet1.Range("FP_Value").value > 0 Then
            Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)
            Sheet1.Range("IncD.LessDeduction57").Locked = False
    End If
    
     'Sheet1.Range("IncD.Section80GGC").Interior.Color = (&HCCFFCC)  ', "AU"
    'Sheet1.Range("IncD.Section80GGC").MergeArea.Locked = False
    'Ayush_17/03/2026
If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
     Sheet1.Range("IncD.Section80TTA").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
End If
     Sheet1.Range("IncD.Section80TTB").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
    Sheet1.Range("SELECT80U").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet1.Range("SELECT80U").MergeArea.Locked = False
'    Sheet1.Range("SELECT80U").Value = "(Select)"
    ' Sheet1.Range("IncD.Section80U").Interior.ColorIndex = 2
    'Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
    Sheet1.Protect Password:=getmsgstate
    LockUnlock80TTBNew
     
     Application.EnableEvents = True
    
End Sub

Sub resetBacNo()

    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sheet1.date10IE").value = ""
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
    
'   Sheet1.Unprotect Password:=getmsgstate
' Sheet8.Unprotect Password:=getmsgstate
 

'reset10IE

'reset10IEAy23_24 PAG_E22 AY 2023-24 Old---------------------------------------
 
' If (Sheet1.Range("OptingNewTaxRegime").value = "No" And Sheet8.Range("BacValue").value = "1") Or (Sheet1.Range("OptingNewTaxRegime").value = "Yes" And Sheet8.Range("BacValue").value = "2") Or (Sheet1.Range("OptingNewTaxRegime").value = "Return Not filed Previous AY" And Sheet8.Range("BacValue").value = "1") Then
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
'
'
'    Else
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'
'    Sheet1.Range("Sheet1.date10IE").value = ""
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'    End If
    'Sheet1.Protect Password:=getmsgstate
    'Sheet8.Protect Password:=getmsgstate
'PAG_E22 AY 2023-24 Old----------------------------
 'Ayush_28/11/2024
' Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'
'   mIntCells = Range("Sheet1.Nature").count
'    Set rangecells = Range("Sheet1.Nature").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = "(Select)"
'    Next

    resetNatureexempt

'
'     mIntCells = Range("Sheet1.Amount").count
'    Set rangecells = Range("Sheet1.Amount").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).value = ""
'    Next
'
' mIntCells = Range("Sheet1.Description").count
'    Set rangecells = Range("Sheet1.Description").Cells
'    For mIntCtr = 1 To mIntCells
'    Sheet1.Unprotect Password:=getmsgstate
'    rangecells.item(mIntCtr).value = "Not Applicable"
'    Next
     Sheet1.Unprotect Password:=getmsgstate
     Application.EnableEvents = False
    Range("IncD.IncomeFromSal_v").value = 0
     Sheet1.Unprotect Password:=getmsgstate
    Range("IncD.IncomeFromSal_ivc").value = 0
      resetIncomeDetails
  Sheet1.Protect Password:=getmsgstate
 Sheet8.Protect Password:=getmsgstate
  Application.EnableEvents = True
End Sub



'Ayush_28/11/2024
Sub resetNatureexempt()
   
            Dim nature1
            Dim nature2
            Dim NatureRow
            
           Application.EnableEvents = False
            Set NatureRow = Sheet1.Range("Sheet1.Nature")
            
            For Each nature1 In NatureRow.Rows
                
            If Sheet1.Range("I" & nature1.row).value <> "(Select)" Or Sheet1.Range("I" & nature1.row).value <> "" Or Sheet1.Range("I" & nature1.row).value <> "Any Other" Then
            'Ayush_27/01
'                If ((Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government")) And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime")) Then
                      If ((Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government")) And (Sheet8.Range("bacValue").value = 1) Then
                   nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature1"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                    End If
                   
                   
                 'Ayush_27/01
'                 ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
                    ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And (Sheet8.Range("bacValue").value = 1) Then
                    nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature7"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                    End If
                    '-----------------
                    
                    'Ayush_05/05/2025
                ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And (Sheet8.Range("bacValue").value = 2) Then
                 nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature8"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                End If
                    '-------------
                    
                

''                ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
                'Ayush_11/02/2026_V0.3
                   ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And (Sheet8.Range("bacValue").value = 2) Then
                'By Bindu as per DE V0.4 on 24th Feb 26
                '  ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958") And (Sheet8.Range("bacValue").value = 2) Then
                  '  Sheet1.Unprotect Password:=getmsgstate
                    nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature2"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                    End If
                  'Ayush_27/01
'                ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime")) Then
                ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet8.Range("bacValue").value = 1)) Then
                    nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature3"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                    End If
                    
                 'Ayush_27/01
'                ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
                ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And (Sheet8.Range("bacValue").value = 2) Then
                     nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature4"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                    End If
                    
                'Ayush_27/01
'                ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime")) Then
                    ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet8.Range("bacValue").value = 1)) Then
                     nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature5"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                    End If
                    
                  'Ayush_27/01
'                 ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
                ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And (Sheet8.Range("bacValue").value = 2) Then
                      nature2 = Application.IsError(Application.VLookup(Sheet1.Range("I" & nature1.row).value, Sheet8.Range("Employment_Nature6"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                   Sheet1.Unprotect Password:=getmsgstate
                      nature1.value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "I", "AA")).value = ""
                     
                    End If
                End If
            End If
            
       ID_Allowance_Amount_2 ("AA" & nature1.row)
       Sheet1.Protect Password:=getmsgstate
       'Application.EnableEvents = True
     Next
     
'SIT- 81028 added by Chetan C M
   Dim mIntCells As Long
   Dim rangecells As Range
   Dim mIntCtr As Long
    Dim mIntCtrr
    mIntCells = Range("Sheet1.Nature").Rows.count
    Set rangecells = Range("Sheet1.Nature").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
         mIntCtrr = mIntCtr + 1
            If (rangecells.item(mIntCtr, 1).value = "" Or rangecells.item(mIntCtr, 1).value = "(Select)") Then
                                 
                If Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "H")) < Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "H")) Then
                    Sheet1.Unprotect Password:=getmsgstate
                    rangecells.item(mIntCtr, 1).value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "I"))
                    Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "Z")).value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "Z"))
                    Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "AA")) = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "AA"))
                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "I")) = "(Select)"
                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "Z")) = ""
                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "AA")) = ""
                 
              End If
            End If
    Next mIntCtr
    
    'added the below function to check data in all cells while adding multiple rows
    For mIntCtr = 1 To mIntCells
         mIntCtrr = mIntCtr + 1
            If (rangecells.item(mIntCtr, 1).value = "" Or rangecells.item(mIntCtr, 1).value = "(Select)") Then
                                 
                 If Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "H")) < Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "H")) Then
                 Sheet1.Unprotect Password:=getmsgstate
                 rangecells.item(mIntCtr, 1).value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "I"))
                 Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "Z")).value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "Z"))
                 Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "AA")) = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "AA"))
                 Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "I")) = "(Select)"
                 Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "Z")) = ""
                 Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "I", "AA")) = ""
                 
            End If
            End If
    Next mIntCtr
    
'end Chetan C M
Application.EnableEvents = True

End Sub

Sub resetBacYesold()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
    ' Sheet1.Range("IncD.LessDeduction57").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = True
    
    'Commented by Bindu as its Auto-populate field
'    Sheet1.Range("IncD.IncomeFromHP_v").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'    Sheet1.Range("IncD.IncomeFromHP_v").value = ""
    
     Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.IncomeFromSal_v").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
    Range("IncD.IncomeFromSal_v").value = 0
    
     Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.IncomeFromSal_ivc").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = True
    Range("IncD.IncomeFromSal_ivc").value = 0
    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80C").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
     Range("IncD.Section80C").value = 0
     Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80CCC").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
    Range("IncD.Section80CCC").value = 0
    Sheet1.Unprotect Password:=getmsgstate
      Sheet1.Range("IncD.Section80CCD_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD_SE").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80CCD1B_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD1B_SE").value = 0
    Sheet1.Unprotect Password:=getmsgstate
   ' Sheet1.Range("IncD.Section80DD").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
    'Range("IncD.Section80DD").value = 0
    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80DDB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
    Range("IncD.Section80DDB").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80E").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
    Range("IncD.Section80E").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    
    Sheet1.Range("IncD.Section80EE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
    Range("IncD.Section80EE").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
    Range("IncD.Section80EEA").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
    Range("IncD.Section80EEB").value = 0
    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80GG").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80GG").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80GG").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.Section80GGC").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80GGC").MergeArea.Locked = True
    'Range("IncD.Section80GGC").value = 0
    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80TTA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
    Range("IncD.Section80TTA").value = 0
    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80TTB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
    Range("IncD.Section80TTB").value = 0
    Sheet1.Unprotect Password:=getmsgstate
      
   '   Sheet1.Range("IncD.Section80U").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
    'Range("IncD.Section80U").value = 0
    Sheet1.Unprotect Password:=getmsgstate
    ccount = 0
   mIntCells = Range("Sheet1.Nature").count
    Set rangecells = Range("Sheet1.Nature").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
    rangecells.item(mIntCtr).value = "(Select)"
    Next
    
     mIntCells = Range("Sheet1.Amount").count
    Set rangecells = Range("Sheet1.Amount").Cells
    For mIntCtr = 1 To mIntCells
    rangecells.item(mIntCtr).value = ""
    Next
         
         mIntCells = Range("Nature_Others_1").count
    Set rangecells = Range("Nature_Others_1").Cells
    For mIntCtr = 1 To mIntCells
    rangecells.item(mIntCtr).value = "Not Applicable"
    Next
    
   
   ' Range("IncD.Deduction16ia").Value = 0
End Sub
Sub resetBacYes()
Application.EnableEvents = False
' Dim rangecells As Range
  '  Dim mIntCells As Long
  '  Dim mIntCtr As Long
   ' Dim ccount As Long
    'Sheet1.Unprotect Password:=getmsgstate
 'Sheet8.Unprotect Password:=getmsgstate

'reset10IE

'reset10IEAy23_24 'Commented by Bindu on 25th Jan 25 as Fields in that function are hidden not using in AY 25-26



' If (Sheet1.Range("OptingNewTaxRegime").value = "No" And Sheet8.Range("BacValue").value = "1") Or (Sheet1.Range("OptingNewTaxRegime").value = "Yes" And Sheet8.Range("BacValue").value = "2") Or (Sheet1.Range("OptingNewTaxRegime").value = "Return Not filed Previous AY" And Sheet8.Range("BacValue").value = "1") Then
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
'
'
'    Else
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.date10IE").Interior.ColorIndex = 15   ', "AU"
'
'    Sheet1.Range("Sheet1.date10IE").value = ""
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.ColorIndex = 15   ', "AU"
'
'    Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'    End If
    
    'Sheet8.Protect Password:=getmsgstate
    
'    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
    Sheet1.Unprotect Password:=getmsgstate
    
'    Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = True
'    Sheet1.Range("IncD.LessDeduction57").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.LessDeduction57").value = ""
    
    'Commented by Bindu as this field auto-populate
'    If Range("IncD.TypeOfHP").value = "Self Occupied" Then
'        Sheet1.Range("IncD.IncomeFromHP_v").Interior.ColorIndex = 15   ', "AU"
'        Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'        Sheet1.Range("IncD.IncomeFromHP_v").value = ""
'    End If
    'dpk
   
    
    ' Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.IncomeFromSal_v").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
    Range("IncD.IncomeFromSal_v").value = 0
    
   '  Sheet1.Unprotect Password:=getmsgstate
   
     Sheet1.Range("IncD.IncomeFromSal_ivc").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = True
    Range("IncD.IncomeFromSal_ivc").value = 0
   ' Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80C").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
     Range("IncD.Section80C").value = 0
  '   Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80CCC").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
'    Range("IncD.Section80CCC").value = 0
    
    
    If Sheet1.Range("Sec80CCC_Type").Locked = True Then
    Sheet1.Range("Sec80CCC_Type").ClearNotes
    End If
    
    'Ayush_05/01/2025
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Type").Interior.ColorIndex = 15   ', "AU"
     Sheet1.Range("Sec80CCC_Type").ClearContents
    Sheet1.Range("Sec80CCC_Type").Locked = True
    
    
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Name").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("Sec80CCC_Name").ClearContents
    Sheet1.Range("Sec80CCC_Name").Locked = True
     
    
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Sec80CCC_Amount").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("Sec80CCC_Amount").ClearContents
    Sheet1.Range("Sec80CCC_Amount").Locked = True
    '----------------
    'Ayush_16/06/2026
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80CCD_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD_SE").value = 0
    
    'Ayush_05/01/2025
    'Ayush_16/04/2026
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Type").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sec80CCC_1_Type").ClearContents
'    Sheet1.Range("Sec80CCC_1_Type").Locked = True
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Name").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sec80CCC_1_Name").ClearContents
'    Sheet1.Range("Sec80CCC_1_Name").Locked = True
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1_Amount").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sec80CCC_1_Amount").ClearContents
'    Sheet1.Range("Sec80CCC_1_Amount").Locked = True
    '----------------
    
 'Ayush_commented_26-27_changes
'         'Ayush_16/05
    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("pran_new").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("pran_new").Locked = True

    
    Sheet1.Range("Sheet1.AckNum").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
    Sheet1.Range("Sheet1.AckNum").Font.ColorIndex = 15
    
    
    
    'Sheet1.Range("Sheet1.AckNum").MergeArea.ClearContents
    
    'Range("Sheet1.AckNum").value = 0
    
    '------------------------
    'Ayush_16/06/2026
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80CCD1B_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD1B_SE").value = 0
    'Ayush_05/01/2025
    'Ayush_16/04/2026
'    Sheet1.Range("Sec80CCC_1b_Type").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sec80CCC_1b_Type").ClearContents
'    Sheet1.Range("Sec80CCC_1b_Type").Locked = True
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1b_Name").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sec80CCC_1b_Name").ClearContents
'    Sheet1.Range("Sec80CCC_1b_Name").Locked = True
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Sec80CCC_1b_Amount").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sec80CCC_1b_Amount").ClearContents
'    Sheet1.Range("Sec80CCC_1b_Amount").Locked = True
    '----------------
    
    
    
    
    
    Sheet1.Range("SELECT80DD").Interior.ColorIndex = 15
    Sheet1.Range("SELECT80DD").MergeArea.Locked = True
    Sheet1.Range("SELECT80DD").value = "(Select)"
   ' Sheet1.Unprotect Password:=getmsgstate
  '  Sheet1.Range("IncD.Section80DD").Interior.ColorIndex = 15   ', "AU"
   ' Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
    'Range("IncD.Section80DD").value = 0
    Sheet1.Range("SELECT80DDB").Interior.ColorIndex = 15
    Sheet1.Range("SELECT80DDB").MergeArea.Locked = True
    Sheet1.Range("SELECT80DDB").value = "(Select)"
    
    
       'Sudha_2025-26
   
    Sheet1.Range("Sheet1.Specified_Disease").Interior.ColorIndex = 15
    Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = True
    Sheet1.Range("Sheet1.Specified_Disease").value = "(Select)"
    
  '====================
  
   ' Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80DDB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
    Range("IncD.Section80DDB").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80E").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
    Range("IncD.Section80E").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate

    Sheet1.Range("IncD.Section80EE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
    Range("IncD.Section80EE").value = 0
   ' Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
    Range("IncD.Section80EEA").value = 0
 '   Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
    Range("IncD.Section80EEB").value = 0
 '   Sheet1.Unprotect Password:=getmsgstate
    Range("IncD.Section80GG").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80GG").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80GG").MergeArea.Locked = True
    
 'sudha_2025-26
 
'    Sheet1.Range("Sheet1.AckNum").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
'    Sheet1.Range("Sheet1.AckNum").Font.Color = vbWhite
    
    
'===================

    
  '  Sheet1.Unprotect Password:=getmsgstate
    'Sheet1.Range("IncD.Section80GGC").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80GGC").MergeArea.Locked = True
    'Range("IncD.Section80GGC").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80TTA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
    Range("IncD.Section80TTA").value = 0
'    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80TTB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
    Range("IncD.Section80TTB").value = 0
    Sheet1.Range("SELECT80U").Interior.ColorIndex = 15
    Sheet1.Range("SELECT80U").MergeArea.Locked = True
    Sheet1.Range("SELECT80U").value = "(Select)"
   ' Sheet1.Unprotect Password:=getmsgstate
    '  Sheet1.Range("IncD.Section80U").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
    'Range("IncD.Section80U").value = 0
   ' Sheet1.Unprotect Password:=getmsgstate
   'Ayush_28/11/2024
    'ccount = 0
    resetNatureexempt
    
    
    '_---Not working need to add new function
'     If Range("Sheet1.Nature1").value = "Family pension" Then
'        Sheet1.Range("IncD.IncomeFromHP_v").Interior.ColorIndex = 15   ', "AU"
'        Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'        Sheet1.Range("IncD.IncomeFromHP_v").value = ""
'    End If
    'Ayush_28/11/2024
  ' mIntCells = Range("Sheet1.Nature").count
  '  Set rangecells = Range("Sheet1.Nature").Cells
  '  Dim countrycd As Variant
  '  For mIntCtr = 1 To mIntCells
  '  rangecells.item(mIntCtr).value = "(Select)"
  '  Next

   '  mIntCells = Range("Sheet1.Amount").count
   ' Set rangecells = Range("Sheet1.Amount").Cells
   ' For mIntCtr = 1 To mIntCells
   ' rangecells.item(mIntCtr).value = ""
   ' Next

    '     mIntCells = Range("Sheet1.Description").count
   ' Set rangecells = Range("Sheet1.Description").Cells
   ' For mIntCtr = 1 To mIntCells
   ' rangecells.item(mIntCtr).value = "Not Applicable"
   ' Next

    
    
   Application.EnableEvents = True
   ' Range("IncD.Deduction16ia").Value = 0
End Sub
Sub resetBacYesold1()
Application.EnableEvents = False
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
    Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
    ' Sheet1.Range("IncD.LessDeduction57").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = True
    'Sheet1.Range("IncD.LessDeduction57").value = ""
    
    'Commented by Bindu as its auto-populate field
'    If Range("IncD.TypeOfHP").value = "Self Occupied" Then
'        Sheet1.Range("IncD.IncomeFromHP_v").Interior.ColorIndex = 15   ', "AU"
'        Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'        Sheet1.Range("IncD.IncomeFromHP_v").value = ""
'    End If
    'dpk
'    If Range("Sheet1.Nature1").value = "Family pension" Then
'        Sheet1.Range("IncD.IncomeFromHP_v").Interior.ColorIndex = 15   ', "AU"
'        Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'        Sheet1.Range("IncD.IncomeFromHP_v").value = ""
'    End If
    
    ' Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.IncomeFromSal_v").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
    Range("IncD.IncomeFromSal_v").value = 0
    
   '  Sheet1.Unprotect Password:=getmsgstate
   
     Sheet1.Range("IncD.IncomeFromSal_ivc").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = True
    Range("IncD.IncomeFromSal_ivc").value = 0
   ' Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80C").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
     Range("IncD.Section80C").value = 0
  '   Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80CCC").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
    Range("IncD.Section80CCC").value = 0
    'Sheet1.Unprotect Password:=getmsgstate
      Sheet1.Range("IncD.Section80CCD_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD_SE").value = 0
    'Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80CCD1B_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD1B_SE").value = 0
    Sheet1.Range("SELECT80DD").Interior.ColorIndex = 15
    Sheet1.Range("SELECT80DD").MergeArea.Locked = True
    Sheet1.Range("SELECT80DD").value = "(Select)"
   ' Sheet1.Unprotect Password:=getmsgstate
   ' Sheet1.Range("IncD.Section80DD").Interior.ColorIndex = 15   ', "AU"
   ' Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
    'Range("IncD.Section80DD").value = 0
    Sheet1.Range("SELECT80DDB").Interior.ColorIndex = 15
    Sheet1.Range("SELECT80DDB").MergeArea.Locked = True
    Sheet1.Range("SELECT80DDB").value = "(Select)"
    
   ' Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80DDB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
    Range("IncD.Section80DDB").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80E").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
    Range("IncD.Section80E").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate

    Sheet1.Range("IncD.Section80EE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
    Range("IncD.Section80EE").value = 0
   ' Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
    Range("IncD.Section80EEA").value = 0
 '   Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
    Range("IncD.Section80EEB").value = 0
 '   Sheet1.Unprotect Password:=getmsgstate
    Range("IncD.Section80GG").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80GG").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80GG").MergeArea.Locked = True
    
  '  Sheet1.Unprotect Password:=getmsgstate
   ' Sheet1.Range("IncD.Section80GGC").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80GGC").MergeArea.Locked = True
    'Range("IncD.Section80GGC").value = 0
  '  Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80TTA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
    Range("IncD.Section80TTA").value = 0
'    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("IncD.Section80TTB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
    Range("IncD.Section80TTB").value = 0
   Sheet1.Range("SELECT80U").Interior.ColorIndex = 15
    Sheet1.Range("SELECT80U").MergeArea.Locked = True
    Sheet1.Range("SELECT80U").value = "(Select)"
   ' Sheet1.Unprotect Password:=getmsgstate
   '   Sheet1.Range("IncD.Section80U").Interior.ColorIndex = 15   ', "AU"
    'Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
    'Range("IncD.Section80U").value = 0
   ' Sheet1.Unprotect Password:=getmsgstate
    ccount = 0
   mIntCells = Range("Sheet1.Nature").count
    Set rangecells = Range("Sheet1.Nature").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
    rangecells.item(mIntCtr).value = "(Select)"
    Next
    
     mIntCells = Range("Sheet1.Amount").count
    Set rangecells = Range("Sheet1.Amount").Cells
    For mIntCtr = 1 To mIntCells
    rangecells.item(mIntCtr).value = ""
    Next
         
         mIntCells = Range("Sheet1.Description").count
    Set rangecells = Range("Sheet1.Description").Cells
    For mIntCtr = 1 To mIntCells
    rangecells.item(mIntCtr).value = "Not Applicable"
    Next
    
   Application.EnableEvents = True
   ' Range("IncD.Deduction16ia").Value = 0
End Sub
'Ayush_26-27_C125
Sub AddRows_Sec80CCC()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    
  '  If Sheet1.Range("IncD.Section80CCC").value > 0 Then

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    mdCommon.DefinedgridNameRange = "Sec80CCC_Sl||Sec80CCC_Type||Sec80CCC_Name||Sec80CCC_Amount"
    ActiveCellRange = mdCommon.searchLastRow("Sec80CCC_Type")
    vRows = mdCommon.insertRowUnderSectionWithFormula_IncomeDetails_80CCC
    
'    Dim icnt, i As Variant
'    icnt = Sheet1.Range("Sec80CCC_Sl").Rows.count
'    For i = 1 To icnt
'        Application.EnableEvents = False
'        Sheet1.Unprotect Password:=getmsgstate
'               Sheet1.Range("Sec80CCC_Sl").Cells(i, 1).value = i
'        Application.EnableEvents = True
'    Next

  '  End If
    

End Sub
'Ayush_26-27_C131
'Sub AddRows_Sec80CCC_1()
'    Dim vRows As Long
'    Dim sourceSheet As Worksheet
'
'  '  If Sheet1.Range("IncD.Section80CCD_SE").value > 0 Then
'
'    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
'    sourceSheet.Activate
'    mdCommon.DefinedgridNameRange = "Sec80CCC_1_Sl||Sec80CCC_1_Type||Sec80CCC_1_Name||Sec80CCC_1_Amount"
'    ActiveCellRange = mdCommon.searchLastRow("Sec80CCC_1_Type")
'    vRows = mdCommon.insertRowUnderSectionWithFormula_IncomeDetails_80CCC
'
''    Dim icnt, i As Variant
''    icnt = Sheet1.Range("Sec80CCC_1_Sl").Rows.count
''    For i = 1 To icnt
''        Application.EnableEvents = False
''        Sheet1.Unprotect Password:=getmsgstate
''               Sheet1.Range("Sec80CCC_1_Sl").Cells(i, 1).value = i
''        Application.EnableEvents = True
''    Next
'
' '   End If
'
'
'End Sub

'Ayush_26-27_C137
'Sub AddRows_Sec80CCC_1b()
'    Dim vRows As Long
'    Dim sourceSheet As Worksheet
'
' '   If Sheet1.Range("IncD.Section80CCD1B_SE").value > 0 Then
'
'    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
'    sourceSheet.Activate
'    mdCommon.DefinedgridNameRange = "Sec80CCC_1b_Sl||Sec80CCC_1b_Type||Sec80CCC_1b_Name||Sec80CCC_1b_Amount"
'    ActiveCellRange = mdCommon.searchLastRow("Sec80CCC_1b_Type")
'    vRows = mdCommon.insertRowUnderSectionWithFormula_IncomeDetails_80CCC
'
''    Dim icnt, i As Variant
''    icnt = Sheet1.Range("Sec80CCC_1b_Sl").Rows.count
''    For i = 1 To icnt
''        Application.EnableEvents = False
''        Sheet1.Unprotect Password:=getmsgstate
''               Sheet1.Range("Sec80CCC_1b_Sl").Cells(i, 1).value = i
''        Application.EnableEvents = True
''    Next
' '   End If
'
'
'End Sub
Sub AddRows_Sec80CCC_pran()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    
 

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    mdCommon.DefinedgridNameRange = "pran_new||Pran_Sl"
    ActiveCellRange = mdCommon.searchLastRow("pran_new")
    vRows = mdCommon.insertRowUnderSectionWithFormula_IncomeDetails_80CCC
    

    

End Sub





' ========== DETECTION HELPERS ==========

' Normalize the OnAction (strip workbook or module qualifiers, trim spaces)
Private Function NormalizeOnAction(ByVal act As String) As String
    Dim bangPos As Long, dotPos As Long
    act = Trim$(act)
    If Len(act) = 0 Then
        NormalizeOnAction = ""
        Exit Function
    End If
    ' Strip "Workbook.xlsm!"
    bangPos = InStr(1, act, "!", vbTextCompare)
    If bangPos > 0 Then act = Mid$(act, bangPos + 1)

    ' Some buttons may have module qualification like "Module1.AddRows_Sec80CCC"
    dotPos = InStrRev(act, ".", -1, vbTextCompare)
    If dotPos > 0 Then act = Mid$(act, dotPos + 1)

    NormalizeOnAction = Trim$(act)
End Function

' True only for Form Control buttons (not ActiveX)
Private Function IsFormControlButton(ByVal shp As Shape) As Boolean
    On Error Resume Next
    Dim hasCF As Boolean
    hasCF = Not shp.ControlFormat Is Nothing     ' ControlFormat exists for Form Controls
    If hasCF Then
        ' msoFormControl + type button (type 0=button)
        ' Some versions: shp.FormControlType = xlButtonControl (but not universally available)
        IsFormControlButton = True
    Else
        IsFormControlButton = False
    End If
End Function

' Traverse groups to collect form control buttons assigned to given macro
Public Function FindFormButtonsByMacro(ws As Worksheet, macroName As String) As Collection
    Dim result As New Collection
    Dim shp As Shape
    For Each shp In ws.Shapes
        CollectIfMatch shp, macroName, result
    Next shp

    If result.count = 0 Then
        Set FindFormButtonsByMacro = Nothing
    Else
        Set FindFormButtonsByMacro = result
    End If
End Function

' Recursive worker: handles groups
Private Sub CollectIfMatch(ByVal shp As Shape, ByVal macroName As String, ByRef result As Collection)
    On Error Resume Next

    If shp.Type = msoGroup Then
        ' Check items inside the group
        Dim gi As Shape
        For Each gi In shp.GroupItems
            CollectIfMatch gi, macroName, result
        Next gi
        Exit Sub
    End If

    ' Only consider Form Control buttons
    If Not IsFormControlButton(shp) Then Exit Sub

    Dim act As String
    act = ""
    act = shp.OnAction

    If Len(act) = 0 Then Exit Sub

    act = NormalizeOnAction(act)

    If StrComp(act, macroName, vbTextCompare) = 0 Then
        result.add shp
    End If
End Sub

' === Protection helper (optional) ===
Public Sub SafeProtect(ws As Worksheet, doProtect As Boolean)
    On Error Resume Next
    If doProtect Then
        ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    Else
        ws.Unprotect Password:=""
    End If
End Sub

' ========== DIAGNOSTICS & FIXERS ==========

' List all shapes on Sheet1 with their OnAction (normalized + raw) to Immediate Window
Public Sub InspectSheet1Buttons()
    Dim ws As Worksheet
    Set ws = Sheet1

    Dim shp As Shape, rawAct As String, normAct As String
    Debug.Print "---- Inspecting shapes on Sheet1 ----"
    For Each shp In ws.Shapes
        rawAct = ""
        On Error Resume Next
        rawAct = shp.OnAction
        On Error GoTo 0

        normAct = NormalizeOnAction(rawAct)
        Debug.Print "Name: "; shp.name; _
                    " | Type: "; shp.Type; _
                    " | FormBtn: "; IsFormControlButton(shp); _
                    " | OnAction(raw): "; rawAct; _
                    " | OnAction(norm): "; normAct
    Next shp
    Debug.Print "--------------------------------------"
End Sub

' Fixer: reassign the macro to only the intended buttons based on criteria
' Option A: by current normalized OnAction match
Public Sub ReassignAddRowsMacro_ByOnAction()
    Dim ws As Worksheet: Set ws = Sheet1
    Dim shp As Shape, rawAct As String, normAct As String

    For Each shp In ws.Shapes
        If IsFormControlButton(shp) Then
            rawAct = ""
            On Error Resume Next
            rawAct = shp.OnAction
            On Error GoTo 0

            normAct = NormalizeOnAction(rawAct)

            ' If you find close matches or older names, unify them:
            If InStr(1, normAct, "AddRows_Sec80CCC", vbTextCompare) > 0 Then
                shp.OnAction = "AddRows_Sec80CCC"
            End If
        End If
    Next shp
    MsgBox "Reassignment complete (by OnAction).", vbInformation
End Sub

' Option B: by visible caption text on the button (works if the button's text says "Add Rows")
Public Sub ReassignAddRowsMacro_ByCaption()
    Dim ws As Worksheet: Set ws = Sheet1
    Dim shp As Shape

    For Each shp In ws.Shapes
        If IsFormControlButton(shp) Then
            ' ControlFormat.Characters.Text usually returns caption for Form Controls
            Dim txt As String
            On Error Resume Next
            txt = shp.TextFrame.Characters.text
            If Len(txt) = 0 Then txt = shp.AlternativeText ' fallback
            On Error GoTo 0

            If InStr(1, txt, "Add Rows", vbTextCompare) > 0 Then
                shp.OnAction = "AddRows_Sec80CCC"
            End If
        End If
    Next shp

    MsgBox "Reassignment complete (by caption).", vbInformation
End Sub

' Option C: by shape name pattern (e.g., names containing "AddRows")
Public Sub ReassignAddRowsMacro_ByNamePattern()
    Dim ws As Worksheet: Set ws = Sheet1
    Dim shp As Shape

    For Each shp In ws.Shapes
        If IsFormControlButton(shp) Then
            If InStr(1, shp.name, "AddRows", vbTextCompare) > 0 Then
                shp.OnAction = "AddRows_Sec80CCC"
            End If
        End If
    Next shp

    MsgBox "Reassignment complete (by name pattern).", vbInformation
End Sub

Function Validate_80CCC()
Validate_80CCC = True

'MsgBox_80CCC = ""

setTblinfo_80CCC_Type
setTblinfo_80CCC_Name
setTblinfo_80CCC_Amount


end_80CCC = WorksheetFunction.Max(0, end_80CCCType, end_80CCCName, end_80CCCAmount)

If Not ValidateType_80CCC Then Validate_80CCC = False
'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not ValidateIFSC_80E Then Validate_80E = False
'If Not ValidatePAN_80E Then Validate_80E = False

If Not ValidateName_80CCC Then Validate_80CCC = False
If Not ValidateAmount_80CCC Then Validate_80CCC = False




End Function



Sub setTblinfo_80CCC_Type()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_Type").count
    Set rangecells = Range("Sec80CCC_Type").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_80CCCType = ccount
 End Sub
 
 Sub setTblinfo_80CCC_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_Name").count
    Set rangecells = Range("Sec80CCC_Name").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCCName = ccount
End Sub

Sub setTblinfo_80CCC_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_Amount").count
    Set rangecells = Range("Sec80CCC_Amount").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCCAmount = ccount
End Sub


Function ValidateType_80CCC() As Boolean
    ValidateType_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_Type").Cells
    
    
    
    ReDim Type_80CCC(end_80CCC)
    For i = 1 To end_80CCC
        Type_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Type_80CCC(i)) Then
             msgError = msgError & "* Type of Identifier is mandatory in schedule 80CCC at sr. no. " & i & " " & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
         If Len(Type_80CCC(i)) > 125 Then
          msgError = msgError & "* Type of Identifier of Section 80CCC at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Type_80CCC(i)) Then
             msgError = msgError & "* Type of Identifier of Section 80CCC at Sl.no. " & i & " should not contain special characters." & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
         
Next
End Function


Function ValidateName_80CCC() As Boolean
    ValidateName_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_Name").Cells
    ReDim Name_80CCC(end_80CCC)
    For i = 1 To end_80CCC
        Name_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Name_80CCC(i)) Then
             msgError = msgError & "* Identifier No. is mandatory in schedule 80CCC at sr. no. " & i & "" & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
         If Len(Name_80CCC(i)) > 125 Then
          msgError = msgError & "* Identifier No. of Section 80CCC length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Name_80CCC(i)) Then
             msgError = msgError & "* Identifier No. of Section 80CCC at Sl.no. " & i & " should not contain special characters." & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
         
Next
End Function

Function ValidateAmount_80CCC() As Boolean
    ValidateAmount_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_Amount").Cells
    ReDim Amount_80CCC(end_80CCC)
    For i = 1 To end_80CCC
        Amount_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Amount_80CCC(i)) Then
             msgError = msgError & "* Amount is mandatory in schedule 80CCC at sr. no. " & i & "" & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
         
        
        If Not IsNumeric(Amount_80CCC(i)) Then
            msgError = msgError & "* Amount of Section 80CCC at Sr. No  " & i & "  in sheet Income Details should be Numeric value" & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
        
        If Amount_80CCC(i) > 99999999999999# Then
            msgError = msgError & "* Amount of Section 80CCC at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
         
Next
End Function

Function Validate_80CCC_1()
Validate_80CCC_1 = True

'MsgBox_80CCC = ""

setTblinfo_80CCC_1_Type
setTblinfo_80CCC_1_Name
setTblinfo_80CCC_1_Amount


end_80CCC_1 = WorksheetFunction.Max(0, end_80CCC_1_Type, end_80CCC_1_Name, end_80CCC_1_Amount)

If Not ValidateType_1_80CCC Then Validate_80CCC_1 = False
'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not ValidateIFSC_80E Then Validate_80E = False
'If Not ValidatePAN_80E Then Validate_80E = False

If Not ValidateName_1_80CCC Then Validate_80CCC_1 = False
If Not ValidateAmount_1_80CCC Then Validate_80CCC_1 = False




End Function



Sub setTblinfo_80CCC_1_Type()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_1_Type").count
    Set rangecells = Range("Sec80CCC_1_Type").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_80CCC_1_Type = ccount
 End Sub
 
 Sub setTblinfo_80CCC_1_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_1_Name").count
    Set rangecells = Range("Sec80CCC_1_Name").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCC_1_Name = ccount
End Sub

Sub setTblinfo_80CCC_1_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_1_Amount").count
    Set rangecells = Range("Sec80CCC_1_Amount").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCC_1_Amount = ccount
End Sub


Function ValidateType_1_80CCC() As Boolean
    ValidateType_1_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_1_Type").Cells
    
    
    
    ReDim Type_1_80CCC(end_80CCC_1)
    For i = 1 To end_80CCC_1
        Type_1_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Type_1_80CCC(i)) Then
             msgError = msgError & "* Type of Identifier is mandatory in schedule 80CCD(1) at sr. no. " & i & "" & Chr(13)
            ValidateType_1_80CCC = False
            Exit Function
        End If
         If Len(Type_1_80CCC(i)) > 125 Then
          msgError = msgError & "* Type of Identifier of Section 80CCD(1) at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateType_1_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Type_1_80CCC(i)) Then
             msgError = msgError & "* Type of Identifier of Section 80CCD(1) at Sl.no. " & i & " should not contain special characters." & Chr(13)
            ValidateType_1_80CCC = False
            Exit Function
        End If
         
Next
End Function


Function ValidateName_1_80CCC() As Boolean
    ValidateName_1_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_1_Name").Cells
    ReDim Name_1_80CCC(end_80CCC_1)
    For i = 1 To end_80CCC_1
        Name_1_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Name_1_80CCC(i)) Then
             msgError = msgError & "* Name of Identifier is mandatory in schedule 80CCD(1) at sr. no. " & i & "" & Chr(13)
            ValidateName_1_80CCC = False
            Exit Function
        End If
         If Len(Name_1_80CCC(i)) > 125 Then
          msgError = msgError & "* Name of Identifier of Section 80CCD(1) length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateName_1_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Name_1_80CCC(i)) Then
             msgError = msgError & "* Name of Identifier of Section 80CCD(1) at Sl.no. " & i & " should not contain special characters." & Chr(13)
            ValidateName_1_80CCC = False
            Exit Function
        End If
         
Next
End Function

Function ValidateAmount_1_80CCC() As Boolean
    ValidateAmount_1_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_1_Amount").Cells
    ReDim Amount_1_80CCC(end_80CCC_1)
    For i = 1 To end_80CCC_1
        Amount_1_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Amount_1_80CCC(i)) Then
             msgError = msgError & "* Amount is mandatory in schedule 80CCD(1) at sr. no. " & i & "" & Chr(13)
            ValidateAmount_1_80CCC = False
            Exit Function
        End If
         
        
        If Not IsNumeric(Amount_1_80CCC(i)) Then
            msgError = msgError & "* Amount of Section 80CCD(1) at Sr. No  " & i & "  in sheet Income Details should be Numeric value" & Chr(13)
            ValidateAmount_1_80CCC = False
            Exit Function
        End If
        
        If Amount_1_80CCC(i) > 99999999999999# Then
            msgError = msgError & "* Amount of Section 80CCD(1) at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
            ValidateAmount_1_80CCC = False
            Exit Function
        End If
         
Next
End Function


Function Validate_80CCC_1b()
Validate_80CCC_1b = True

'MsgBox_80CCC = ""

setTblinfo_80CCC_1b_Type
setTblinfo_80CCC_1b_Name
setTblinfo_80CCC_1b_Amount


end_80CCC_1b = WorksheetFunction.Max(0, end_80CCC_1b_Type, end_80CCC_1b_Name, end_80CCC_1b_Amount)

If Not ValidateType_1b_80CCC Then Validate_80CCC_1b = False
'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not ValidateIFSC_80E Then Validate_80E = False
'If Not ValidatePAN_80E Then Validate_80E = False

If Not ValidateName_1b_80CCC Then Validate_80CCC_1b = False
If Not ValidateAmount_1b_80CCC Then Validate_80CCC_1b = False




End Function



Sub setTblinfo_80CCC_1b_Type()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_1b_Type").count
    Set rangecells = Range("Sec80CCC_1b_Type").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_80CCC_1b_Type = ccount
 End Sub
 
 Sub setTblinfo_80CCC_1b_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_1b_Name").count
    Set rangecells = Range("Sec80CCC_1b_Name").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCC_1b_Name = ccount
End Sub

Sub setTblinfo_80CCC_1b_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec80CCC_1b_Amount").count
    Set rangecells = Range("Sec80CCC_1b_Amount").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCC_1b_Amount = ccount
End Sub


Function ValidateType_1b_80CCC() As Boolean
    ValidateType_1b_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_1b_Type").Cells
    
    
    
    ReDim Type_1b_80CCC(end_80CCC_1b)
    For i = 1 To end_80CCC_1b
        Type_1b_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Type_1b_80CCC(i)) Then
             msgError = msgError & "* Type of Identifier is mandatory in schedule 80CCD(1B) at sr. no. " & i & "" & Chr(13)
            ValidateType_1b_80CCC = False
            Exit Function
        End If
         If Len(Type_1b_80CCC(i)) > 125 Then
          msgError = msgError & "* Type of Identifier of Section 80CCD(1B) at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateType_1b_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Type_1b_80CCC(i)) Then
             msgError = msgError & "* Type of Identifier of Section 80CCD(1B) at Sl.no. " & i & " should not contain special characters." & Chr(13)
            ValidateType_1b_80CCC = False
            Exit Function
        End If
         
Next
End Function


Function ValidateName_1b_80CCC() As Boolean
    ValidateName_1b_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_1b_Name").Cells
    ReDim Name_1b_80CCC(end_80CCC_1b)
    For i = 1 To end_80CCC_1b
        Name_1b_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Name_1b_80CCC(i)) Then
             msgError = msgError & "* Name of Identifier is mandatory in schedule 80CCD(1B) at sr. no. " & i & "" & Chr(13)
            ValidateName_1b_80CCC = False
            Exit Function
        End If
         If Len(Name_1b_80CCC(i)) > 125 Then
          msgError = msgError & "* Name of Identifier of Section 80CCD(1B) length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateName_1b_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Name_1b_80CCC(i)) Then
             msgError = msgError & "* Name of Identifier of Section 80CCD(1B) at Sl.no. " & i & " should not contain special characters." & Chr(13)
            ValidateName_1b_80CCC = False
            Exit Function
        End If
         
Next
End Function

Function ValidateAmount_1b_80CCC() As Boolean
    ValidateAmount_1b_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Sec80CCC_1b_Amount").Cells
    ReDim Amount_1b_80CCC(end_80CCC_1b)
    For i = 1 To end_80CCC_1b
        Amount_1b_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Amount_1b_80CCC(i)) Then
             msgError = msgError & "* Amount is mandatory in schedule 80CCD(1B) at sr. no. " & i & "" & Chr(13)
            ValidateAmount_1b_80CCC = False
            Exit Function
        End If
         
        
        If Not IsNumeric(Amount_1b_80CCC(i)) Then
            msgError = msgError & "* Amount of Section 80CCD(1B) at Sr. No  " & i & "  in sheet Income Details should be Numeric value" & Chr(13)
            ValidateAmount_1b_80CCC = False
            Exit Function
        End If
        
        If Amount_1b_80CCC(i) > 99999999999999# Then
            msgError = msgError & "* Amount of Section 80CCD(1B) at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
            ValidateAmount_1b_80CCC = False
            Exit Function
        End If
         
Next
End Function


'Ayush_23/02/2026
'Sub FitComments()
'Dim xComment As Comment
'For Each xComment In Application.ActiveSheet.Comments
'    xComment.Shape.TextFrame.AutoSize = True
'Next
'End Sub
'Ayush_03/04/2026
Sub FitComments_DynamicHeight()
    Dim cmt As Comment
    Dim txt As String
    Dim charsPerLine As Double
    Dim lineCount As Long
    Dim lineHeight As Double
    Dim minLines As Long
    
    ' Tuning parameters
    charsPerLine = 35     ' approx characters per line (depends on width/font)
    lineHeight = 15       ' height of one text line
    minLines = 2          ' minimum number of lines
    
    For Each cmt In ActiveSheet.Comments
        With cmt.Shape
            .TextFrame.AutoSize = False
            .Width = 220        ' fixed width to enforce wrapping
            
            txt = cmt.text
            lineCount = Application.WorksheetFunction.RoundUp(Len(txt) / charsPerLine, 0)
            
            If lineCount < minLines Then lineCount = minLines
            
            .Height = lineCount * lineHeight
        End With
    Next cmt
End Sub
'----------------
Sub AutoPopulateSecondaryAddress()
    Application.EnableEvents = False
 
    If Sheet1.Range("Secondary_Address").value = "Yes" Then
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceNo2").value = Sheet1.Range("sheet1.ResidenceNo").value
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.LocalityOrArea2").value = Sheet1.Range("sheet1.LocalityOrArea").value
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceName2").value = Sheet1.Range("sheet1.ResidenceName").value
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.RoadOrStreet2").value = Sheet1.Range("sheet1.RoadOrStreet").value
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").value = Sheet1.Range("sheet1.CityOrTownOrDistrict").value
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.PinCode").Locked = True Then
            Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Else
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.PinCode2").value = Sheet1.Range("sheet1.PinCode").value
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        End If
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.StateCode2").value = Sheet1.Range("sheet1.StateCode1").value
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CountryCode2").value = Sheet1.Range("sheet1.CountryCode").value
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("HASZIP").Locked = True Then
            Sheet1.Range("HASZIP2").MergeArea.ClearContents
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"
        Else
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("HASZIP2").value = Sheet1.Range("HASZIP").value
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
        End If
'         If Sheet1.Range("HASZIP").value = "Yes" Then
'            Sheet1.Range("sheet1.ZipCode2").value = "XXXXXX"
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
'        End If
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.ZipCode").Locked = True Then
            If Sheet1.Range("HASZIP").value = "Yes" Then
                Sheet1.Range("sheet1.ZipCode2").value = "XXXXXX"
                Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
                Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
            Else
            Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
            End If
        Else
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.ZipCode2").value = Sheet1.Range("sheet1.ZipCode").value
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
        End If
Else

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = False

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.ClearContents

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = False

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = False

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.ClearContents

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = False

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.ClearContents

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = False

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents

        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.StateCode2").value = "(Select)"

        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
 
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.CountryCode2").value = "(Select)"

        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("HASZIP2").MergeArea.ClearContents

        Sheet1.Range("HASZIP2").MergeArea.Locked = False

        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"

    End If
 
End Sub
Function Validate_Pran()
Validate_Pran = True

'MsgBox_80CCC = ""

setTblinfo_pran



'end_80CCC = WorksheetFunction.Max(0, end_80CCCType, end_80CCCName, end_80CCCAmount)

If Not Validate_pran_new Then Validate_Pran = False
'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not ValidateIFSC_80E Then Validate_80E = False
'If Not ValidatePAN_80E Then Validate_80E = False






End Function
Sub setTblinfo_pran()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("pran_new").count
    Set rangecells = Range("pran_new").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_pran = ccount
 End Sub
Function Validate_pran_new() As Boolean
    Validate_pran_new = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("pran_new").Cells
    
    
    
    ReDim pran_80CCC(end_pran)
    If end_pran = 0 Then
            msgError = msgError & "* ""PRAN no. shall be mandatory if deduction u/s 80CCD(1) or 80CCD(1B) is claimed""" & Chr(13)
            Validate_pran_new = False
            Exit Function
    End If
    For i = 1 To end_pran
        pran_80CCC(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(pran_80CCC(i)) Then
'             msgError = msgError & "* pran mandatory at sr. no. " & i & " " & Chr(13)
'            Validate_pran_new = False
'            Exit Function
'        End If
         If Len(pran_80CCC(i)) > 12 Then
          msgError = msgError & "* Please enter a valid 12 digit PRAN at Sr. No " & i & "" & Chr(13)
            Validate_pran_new = False
            Exit Function
        End If
        
        If Not IsNumeric(pran_80CCC(i)) Then
             msgError = msgError & "* Please enter a valid 12 digit PRAN at Sl.no. " & i & "" & Chr(13)
            Validate_pran_new = False
            Exit Function
        End If
         
Next
End Function
