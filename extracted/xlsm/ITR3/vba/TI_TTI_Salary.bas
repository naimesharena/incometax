Attribute VB_Name = "TI_TTI_Salary"
Option Explicit
'Jyoti23/05/2025
Public MsgBox_EA10_13A As String

Public EA_Sch10of13A_2Plus4, endEA10_13A, endEA10_13A_1, endEA10_13A_2, endEA10_13A_3, endEA10_13A_4, endEA10_13A_5, endEA10_13A_6 As Variant
Public EA_Sch10of13A_PlaceofWrk, EA_Sch10of13A_ActlHRArecivedA, EA_Sch10of13A_ActlRentpaid, EA_Sch10of13A_DetlsofSalpersec17of1, EA_Sch10of13A_Actlrentpaid10persalaryB, EA_Sch10of13A_50Por40Pofsalary, EA_Sch10of13A_BasicSalary, EA_Sch10of13A_DearAllowance, EA_Sch10of13A_ElgiblExmptAllwnce10of13A As Variant
'---
Dim frmsize As Long
Dim frmRngname, rngname, cntrRng As String
Public noOfSalFilled As Variant
'PARTBTI added new
Public IncomeFromSal_8b As Variant
Public IncomeFromHP_8b As Variant
Public ProfGainNoSpecBus_8b As Variant
Public ProfGainSpecBus_8b As Variant
Public ProfGainSpecifiedBus_8b As Variant
Public ProfGain115BF_8b As Variant
Public ProfGain115BG_8b As Variant
Public TotProfBusGain_8b As Variant
Public ShortTermUs111A_8b As Variant
Public ShortTerm30_8b As Variant
Public ShortTermOther_8b, ShortTermSpslRate_8b As Variant
Public TotalShortTerm_8b As Variant
Public LongTerm_8b, LongTerm_8b_SR As Variant
Public LongTerm_8b_NP As Variant

Public DeductionsUnder10Aor10AA_8b As Variant
Public PartBchapterVIA_8b As Variant
Public PartCchapterVIA_8b As Variant

Public LongTerm_8b_P As Variant

Public TotalCapGains_8b As Variant
Public ShortTermLongTermTotal_8b As Variant
Public CapGains30Per115BBH_8b As Variant

Public OtherSrcThanOwnRaceHorse_8b As Variant
Public WinLotteriesRacesGambling_8b As Variant
Public FromOwnRaceHorse_8b As Variant
Public TotIncFromOS_8b As Variant
Public TotalTI_8b As Variant
Public CurrentYearLoss_8b As Variant
Public BalanceAfterSetoffLosses_8b As Variant
Public BroughtFwdLossesSetoff_8b As Variant
Public GrossTotalIncome_8b As Variant
Public IncChargeableTaxSplRates_8b As Variant
Public IncChargeTaxSplRate111A112_8b As Variant
Public DeductionsUnderScheduleVIA_8b As Variant
Public TotalIncome_8b As Variant
Public NetAgricultureIncomeOrOtherIncomeForRate_8b As Variant
Public AggregateIncome_8b As Variant
Public LossesOfCurrentYearCarriedFwd_8b As Variant
Public DeemedTotIncSec115JC_8b As Variant

Public Others_Amt  As Variant
Public Others_NOI As Variant
Public Others_NOI1 As Variant

Dim msgValidateSheetPartB As Variant
Dim msgValidateSheet8b As String
Dim msgValidateSheet8bBlanks As String
Dim msgValidateSheet8bspecialcharacters As String

Public NatureOfIncomeOthCount1 As Long
Public ExemptUSectionOthCount1 As Long
Public AnyOthDescriptionOthCount1 As Long

Public NatureOfIncomeOthCount2 As Long
Public ExemptUSectionOthCount2 As Long
Public AnyOthDescriptionOthCount2 As Long

Public NatureOfIncomeOthCount3 As Long
Public ExemptUSectionOthCount3 As Long
Public AnyOthDescriptionOthCount3 As Long

Public rngname_IncomeOth As Variant
'_______________________________________________________

'TRP added new
'----------------------------
Public TRPID As Variant
Public TRPName As Variant
Public TRPreimbursementAmnt As Variant
Dim errmsgTRP As Variant
'--------------------------------------



Public BankAccountNo, BankAccountNumber_9 As Variant
Public IFSCCode, MICRCode_9 As Variant
Public TypeOfAccount, BankAccountType_9, BankName_9 As Variant
Public IFSCCodeOth, NameOfBankOth, AccountNumOth, TypeOfAccountOth As Variant
Public VerificationName, AssesseeVerName_9 As Variant
Public FatherName, FatherName_9 As Variant
Public VarificationPlace, Place_9 As Variant
Public VerificationDate, Date_9 As Variant
Public VerificationPAN, verPAN, verCapacity As Variant
Public AssesseeVerName As Variant
Public AOIFlag_1, AOIFlag, ctFlag As Variant

'Schedule S modified
'-----------------------------------------------------
Dim SalCount As Long

Public SchSEmpPAN As Variant
Public SchSEmpName, SchSEmployer, SALEntertainment As Variant
Public SchSEmpAddr As Variant
Public SchSEmpTown_City As Variant
Public SchSEmpState As Variant
Public SchSEmpPinCode As Variant
Public SchSEmpZipCode As Variant
Public SchSSalary1 As Variant

Public NatureOfIncomeOthCount, ExemptUSectionOthCount, AnyOthDescriptionOthCount As String

Public SchSExemptUSectionTravel1 As Variant
Public SchSExemptUSectionNonMonetary1 As Variant
Public SchSExemptUSectionHouseRent1 As Variant
Public SchSExemptUSectionOth1 As Variant

Public SchSAllowancesNotExempt1 As Variant
Public SchSValueOfPerquisites1 As Variant
Public SchSProfitsinLieuOfSalary1 As Variant
Public SchSDeductionUnderSection1, SchSDeductionUnder6Section1 As Variant
Public SchSIncomeFromSalary1 As Variant

Public SchSTotIncUnderHeadSalaries1 As Variant
'-----------------------------------------------------------


Dim MsgPartBScheduleS() As Variant
Public MsgPartBSheet, MsgScheduleS, MsgScheduleS1 As Variant


Public Bucket10_final As Double
Public Bucket20_final As Double
Public Bucket30_final As Double
Dim msgValidateSheet9 As String

Public currentDate As String

Public errmsgVerification As Variant
'<---Names Assigned By yogesh

Public totalIncome As Long


Public IBANSELECT As Variant
Public IBAN As Variant
Public IBANNameOfBank As Variant
Public IBANCOUNTRY As Variant
Public IBANACCNO As Variant
Public tempcheck, sRangeValue1  As Variant

Public end_IBAN, end_IBANACCNO, end_IBANCOUNTRY, end_IBANNameOfBank As Variant
Public rngname_IBAN, rngname_IBANACCNO, rngname_IBANCOUNTRY, rngname_IBANNameOfBank As Variant
Public Others_NOI_S, Others_NOI1_s, Others_Amt_s, end_OthersNOI_s, end_OthersAmt_s As Variant


Sub Cmd_ValidateSchSalary_Click()
    Dim vbMessgaeCaption As String
    vbMessgaeCaption = "ITR 3: AY: 2026-27"   'Chandru
    ValidatePartB
    If MsgPartBSheet = "" Then
        fmsgboxStatus "Sheet Part B - TI TTI is ok" ', vbOKOnly, vbMessgaeCaption
    End If
    CloseMsg
End Sub
Sub Cmd_AddRowsBank_Click()
Dim vRows As Long
Sheet7.Activate

tempcheck = ""
sRangeValue1 = ""
EfilingCommon.DefinedgridNameRange = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount||Sheet9.CheckBox||tempxml"
ActiveCellRange = EfilingCommon.searchLastRow("Sheet9.OthMICRCode")
vRows = EfilingCommon.insertRowUnderSectionWithFormula(1)

Dim tempcheck1, tempcheck2 As Variant
'tempcheck = Replace(tempcheck, "G", "R")
tempcheck = Replace(tempcheck, "G", "S")
tempcheck1 = Mid(tempcheck, 2, Len(tempcheck))
tempcheck1 = tempcheck1 + 1
tempcheck = Mid(tempcheck, 1, 1) & tempcheck1
'sRangeValue1 = Replace(sRangeValue1, "G", "R")

sRangeValue1 = Replace(sRangeValue1, "G", "S")

tempcheck = tempcheck + ":" + sRangeValue1

LinkCheckBoxes
LinkCheckBoxes1


End Sub
Sub LinkCheckBoxes1()
On Error Resume Next
Dim chk As CheckBox
Dim lCol As Long
Dim temp As Variant
lCol = 8 'number of columns to the right for link
'temp = Replace(tempcheck, "F", "I")

Dim tempcheck2 As Variant




For Each tempcheck2 In Sheet7.Range(tempcheck)
'MsgBox tempcheck2
For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = _
         .TopLeftCell.Offset(0, lCol).Address
   End With
  'MsgBox tempcheck2.AddressLocal
   If (tempcheck2.AddressLocal) = chk.TopLeftCell.Offset(0, lCol).Address Then
   chk.value = xlOff
   
   End If
 
Next chk
tempcheck2.value = False
Next

End Sub


Sub Cmd_AddSalary_Click()
addSalariesBlock
Application.EnableEvents = False
AddTotalSalary1
DescriptionClear
Application.EnableEvents = True
End Sub

Private Sub Cmd_AddCo_Owners_Click()
AddPropertyCoOWners
End Sub
Sub addSalariesBlock()
    Dim te As Long
    Dim ccount As Long
    Sheet39.Activate
    frmsize = Sheet39.Range("SalarySectionItem").value
    cntrRng = "SalarySectionCount"
    frmRngname = "salrptfrm"
    
    'Change-20.02.2023.102.Salary.
    'rngname = "SAL.NameOfEmployer1||sheet.EmployerCategory1||SAL.PANofEmployer1||SAL.AddrDetail1||SAL.CityOrTownOrDistrict1||SAL.StateCode1||SAL.PinCode1||SAL.ZipCode1||SAL.Salary1||SAL.AllowancesNotExempt1||SAL.ValueOfPerquisites1||SAL.ProfitsinLieuOfSalary1||Salary.Section1||Salary.salaryDescription1||Salary.SalaryAmount1||Salary.ValueSection1||Salary.ValueDescription1||Salary.ValueAmount1||Salary.ProfitSection1||Salary.ProfitDescription1||Salary.ProfitAmount1||SAL.ValueOfPerquisitesA1||IncomeNotified89A1||IncomeNotified89A_AmountUS1||IncomeNotified89A_AmountUK1||IncomeNotified89A_AmountCan1||IncomeNotifiedOther89A1||IncomeNotified89A_Amount1"
    rngname = "SAL.NameOfEmployer1||sheet.EmployerCategory1||SAL.PANofEmployer1||SAL.AddrDetail1||SAL.CityOrTownOrDistrict1||SAL.StateCode1||SAL.PinCode1||SAL.ZipCode1||SAL.Salary1||SAL.AllowancesNotExempt1||SAL.ValueOfPerquisites1||SAL.ProfitsinLieuOfSalary1||Salary.Section1||Salary.salaryDescription1||Salary.SalaryAmount1||Salary.ValueSection1||Salary.ValueDescription1||Salary.ValueAmount1||Salary.ProfitSection1||Salary.ProfitDescription1||Salary.ProfitAmount1||SAL.ValueOfPerquisitesA1||IncomeNotified89A1||IncomeNotified89A_AmountUS1||IncomeNotified89A_AmountUK1||IncomeNotified89A_AmountCan1||IncomeNotifiedOther89A1||IncomeNotified89A_Amount1||IncomeTaxablePrevYear89A1"
    
    Call EfilingCommon.addblock(rngname, frmRngname, cntrRng, frmsize)
    Sheet39.Unprotect Password:=getmsgstate
    Dim rowcount
    rowcount = Sheet39.Range("SalarySectionCount").value
    Sheet39.Range("SAL.ValueOfPerquisitesA" & rowcount).formula = "=SUMIF(Salary.ValueSection" & rowcount & ",""Stock options allotted or transferred by employer being an eligible start-up referred to in section 80-IAC"",Salary.ValueAmount" & rowcount & ")"
    Sheet39.Protect Password:=getmsgstate
End Sub
Sub addSalariesBlockold1()
    Dim te As Long
    Dim ccount As Long
    Sheet39.Activate
    frmsize = Sheet39.Range("SalarySectionItem").value
    cntrRng = "SalarySectionCount"
    frmRngname = "salrptfrm"
    rngname = "SAL.NameOfEmployer1||sheet.EmployerCategory1||SAL.PANofEmployer1||SAL.AddrDetail1||SAL.CityOrTownOrDistrict1||SAL.StateCode1||SAL.PinCode1||SAL.ZipCode1||SAL.Salary1||SAL.AllowancesNotExempt1||SAL.ValueOfPerquisites1||SAL.ProfitsinLieuOfSalary1||Salary.Section1||Salary.salaryDescription1||Salary.SalaryAmount1||Salary.ValueSection1||Salary.ValueDescription1||Salary.ValueAmount1||Salary.ProfitSection1||Salary.ProfitDescription1||Salary.ProfitAmount1"
    Call EfilingCommon.addblock(rngname, frmRngname, cntrRng, frmsize)
End Sub
'new

Sub AddSalary17_1(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet39.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet39.Range("SalarySectionCount").value
    
        
        EfilingCommon.DefinedgridNameRange = "Salary.Section||Salary.salaryDescription||Salary.SalaryAmount"
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.Section" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    'For i = 1 To Sheet39.Range("SalarySectionCount").value
        newfrmsize = Sheet39.Range("SalarySectionItem").value
        newfrmsize = newfrmsize + numberofrows
        Sheet39.Range("SalarySectionItem").value = newfrmsize
    'Next

endline:
   
'----------------Lock Password-------------------START---
   Sheet39.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub

Sub AddSalary17_2(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet39.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet39.Range("SalarySectionCount").value
    
        
        EfilingCommon.DefinedgridNameRange = "Salary.ValueSection||Salary.ValueDescription||Salary.ValueAmount"
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.ValueSection" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    'For i = 1 To Sheet39.Range("SalarySectionCount").value
        newfrmsize = Sheet39.Range("SalarySectionItem").value
        newfrmsize = newfrmsize + numberofrows
        Sheet39.Range("SalarySectionItem").value = newfrmsize
    'Next

endline:
   
'----------------Lock Password-------------------START---
   Sheet39.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub


Sub AddSalary17_3(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet39.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet39.Range("SalarySectionCount").value
    
        
        EfilingCommon.DefinedgridNameRange = "Salary.ProfitSection||Salary.ProfitDescription||Salary.ProfitAmount"
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.ProfitSection" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    'For i = 1 To Sheet39.Range("SalarySectionCount").value
        newfrmsize = Sheet39.Range("SalarySectionItem").value
        newfrmsize = newfrmsize + numberofrows
        Sheet39.Range("SalarySectionItem").value = newfrmsize
    'Next

endline:
   
'----------------Lock Password-------------------START---
   Sheet39.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub



'end
Sub AddPropertyCoOWners(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet8.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value
    
        
        EfilingCommon.DefinedgridNameRange = "HP.Co.Srno||HP.Co.Name||HP.Co.Pan||HP.Co.Aadhaar||HP.Co.Share||Co_OwnnerBlock"
        ActiveCellRange = EfilingCommon.searchLastRow("HP.Co.Srno" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
        newfrmsize = Sheet8.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet8.Range("NumRowsPropertyBlock").value = newfrmsize
    Next

endline:
   
'----------------Lock Password-------------------START---
   Sheet8.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub


Sub ValidatePartB()

Dim vbMessgaeCaption As String
Dim count, scount As Long

vbMessgaeCaption = "ITR 3: AY: 2025-26"  'Chandru

    MsgPartBSheet = ""
  
    ' msgbox9 = ""
     
'added new
'------------------------------------
If Not Validatesheet8b Then
        Sheet7.Activate
        fmsgboxsmall msgValidateSheet8b & msgValidateSheetPartB ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
 '------------------------------------
    
 'added new
 '------------------------------------------
 
'--------------------------------------
    
  

    setTblinfo_IBAN_1
    ValidateIBANSELECT


    ValidateIBAN
    ValidateIBANNameOfBank
    ValidateIBANBankAccountNumber_9
    ValidateIBANCountry_9

    'If IBANSELECT = "N" And end_IBAN < 1 Then
    '    msgbox9 ("Please fill all the mandatory fileds of Foreign bank account in Sheet Part B - TI TTI")
    'End If
    

    setTblinfo_IBAN_2
    setTblinfo_IBAN_3
    setTblinfo_IBAN_4
    
    If end_IBAN <> end_IBANNameOfBank Or _
       end_IBAN <> end_IBANACCNO Or _
       end_IBAN <> end_IBANCOUNTRY Then
        msgbox9 ("* Please fill all the fields under Foreign Bank Account details in Sheet Part B - TI TTI")
    End If


     ValidateBA
     ValidateAOIFlag

     ValidateFTFlag
     
     validatesheetTRP
     
     If msgValidateSheet9 <> "" Then
        Sheet7.Activate
        fmsgboxsmall msgValidateSheet9 ', vbOKOnly, vbMessgaeCaption
        CloseMsg
    End If

    
End Sub
Function validatesheetTRP() As Boolean

validatesheetTRP = True

TRPID = Trim(Range("Sheet9.IdentificationNoOfTRP").value)
TRPName = Trim(Range("Sheet9.NameOfTRP").value)
TRPreimbursementAmnt = Trim(Range("Sheet9.ReImbFrmGov").value)


If (TRPID <> "" Or TRPreimbursementAmnt <> "") And TRPName = "" Then msgbox9 ("Please enter TRP name" & Chr(13))
If (TRPName <> "" Or TRPreimbursementAmnt <> "") And TRPID = "" Then msgbox9 ("Please enter TRP ID" & Chr(13))
'If (TRPName <> "" Or TRPID <> "") And TRPreimbursementAmnt = "" Then msgbox9 ( "Please enter TRP Amount" & Chr(13))
If TRPID <> "" Then
    If Not ChkTRPID(TRPID) Then msgbox9 ("Please Enter a valid TRP ID" & Chr(13))
End If

'If errmsgTRP <> "" Then validatesheetTRP = False

End Function

Function ChkTRPID(item As Variant) As Boolean
On Error Resume Next
Dim i As Long
ChkTRPID = True

If Len(item) = 10 Then
    
    If UCase(Mid(item, 1, 1)) <> "T" Then
        ChkTRPID = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(item, 2, 9)) Then
        ChkTRPID = False
        Exit Function
    End If
    
ElseIf Len(item) = 6 Then

    If Not IsNumeric(Mid(item, 1, 6)) Then
        ChkTRPID = False
        Exit Function
    End If
    
Else
    ChkTRPID = False
    Exit Function
End If
End Function



Sub setTblinfo_IBAN_1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet7.Range("Sheet9.IBAN").count
 Set rangecells = Sheet7.Range("Sheet9.IBAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_IBAN = ccount
 rngname_IBAN = "Sheet9.IBAN;Sheet9.IBANBANK;Sheet9.IBANCOUNTRY;Sheet9.IBANACCNO;" '
End Sub

Sub setTblinfo_IBAN_2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet7.Range("Sheet9.IBANBANK").count
 Set rangecells = Sheet7.Range("Sheet9.IBANBANK").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_IBANNameOfBank = ccount
 rngname_IBANNameOfBank = "Sheet9.IBAN;Sheet9.IBANBANKSheet9.IBANCOUNTRY;;Sheet9.IBANACCNO;" '
End Sub

Sub setTblinfo_IBAN_3()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet7.Range("Sheet9.IBANCOUNTRY").count
 Set rangecells = Sheet7.Range("Sheet9.IBANCOUNTRY").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
         ccount = ccount + 1
     End If
 Next
 end_IBANCOUNTRY = ccount
 rngname_IBANCOUNTRY = "Sheet9.IBAN;Sheet9.IBANBANK;Sheet9.IBANCOUNTRY;Sheet9.IBANACCNO;" '
End Sub

Sub setTblinfo_IBAN_4()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet7.Range("Sheet9.IBANACCNO").count
 Set rangecells = Sheet7.Range("Sheet9.IBANACCNO").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_IBANACCNO = ccount
 rngname_IBANACCNO = "Sheet9.IBAN;Sheet9.IBANBANK;Sheet9.IBANCOUNTRY;Sheet9.IBANACCNO;" '
End Sub
    
 
Sub ValidateScheduleS_Click()
 validateSchS
 If MsgScheduleS = "" Or MsgScheduleS1 = "" Then
    fmsgboxStatus "Sheet Schedule S is OK." ', vbOKOnly, "ITR 3 : 2023-24"
 End If
End Sub
Sub validateSchS()
Dim scount, count As Variant
MsgScheduleS = ""
MsgScheduleS1 = ""
subProcCaption = "Validating Schedule S"
   If Not validateScheduleS Then
        Sheet39.Activate
        scount = Sheet39.Range("SalarySectionCount").value
        
        For count = 1 To scount
                MsgScheduleS = MsgScheduleS & Chr(13) & MsgPartBScheduleS(count)
        Next
        
     
    fmsgboxsmall MsgScheduleS ', vbOKOnly, "ITR 3 : 2023-24"
    CloseMsg
    End If
    validate10aaAmount
    If Len((Range("Salary.Nature1").item(1).value) > 0) Then
        If Not ValidateNatureOfIncome Then
        End If
        If Not ValidateAmount Then
        End If
    End If
    
    
    
    
    'd
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then
    If (Sheet39.Range("SAL.ExemptUSectionOth1").value - Sheet19.Range("Temp_HRA_Sec13A").value) > (Sheet39.Range("SAL.TotalGrossSalary").value - Sheet19.Range("Temp_HRA").value - Sheet19.Range("Temp_Inc89A").value - Sheet19.Range("Temp_Inc89AOth").value) Then
     MsgScheduleS1 = MsgScheduleS1 + "* Total of allowances excluding HRA shall not exceed total of gross salary at sr. no. 2 as reduced by HRA in Schedule Salary." & Chr(13)
'     ValidateLessAllowanceTable = False
    End If
    End If
    
    If Sheet19.Range("Temp_HRA_Sec13A").value > 0 And Sheet19.Range("scvia.Section80GG_Calc").value > 0 Then
     MsgScheduleS1 = MsgScheduleS1 + "* Deduction u/s 10(13A) & 80GG cannot be claimed for the same period " & Chr(13)
'     ValidateLessAllowanceTable = False
    End If
    
    setTblinfo_OthersNOI
    setTblinfo_OthersAmt

    If ((end_OthersNOI_s <> end_OthersAmt_s)) Then
        MsgScheduleS1 = MsgScheduleS1 + "* Enter all mandatory fields in Table Less allowances to the extent exempt u/s 10 in Sheet Salary" & Chr(13)
    End If
    
     If Sheet39.Range("SAL.Entertainment_allowance_1").value > 5000 Then
        MsgScheduleS1 = MsgScheduleS1 + "* Entertainment allowance u/s 16(ii) should not be greater than 5000"
        'validateScheduleS = False
    End If

    If Sheet39.Range("SAL.DeductionUnderSection1").value > 5000 Then
        fmsgbox "*  Deduction of Professional tax u/s 16(iii) is Rs. 200/- p.m. in most of the state. Kindly reconsider the amount claimed"
        'validateSchS = False
    End If
    
'Change 101 10.12.2023 Ver.0.6 SIT-62872
Dim ConvAll, OthAll, Othrs, Ex1014i, Ex1014ii, block As Long
Dim Gratuity_Sal, Leave_Sal, Ex1014iii_Gratuity, Ex1014iiv_Leave As Long

ConvAll = 0
OthAll = 0
Othrs = 0
Ex1014i = 0
Ex1014ii = 0

Gratuity_Sal = 0
Ex1014iii_Gratuity = 0
Leave_Sal = 0
Ex1014iiv_Leave = 0

scount = Sheet39.Range("SalarySectionCount").value
For block = 1 To scount

    Dim j As Long
    For j = 1 To Sheet39.Range("Salary.Section" & block).count
        If Sheet39.Range("Salary.Section" & block).item(j, 1).value = "Conveyance Allowance" Then
            ConvAll = ConvAll + Sheet39.Range("Salary.SalaryAmount" & block).item(j, 1).value
        End If
        If Sheet39.Range("Salary.Section" & block).item(j, 1).value = "Other Allowance" Then
            OthAll = OthAll + Sheet39.Range("Salary.SalaryAmount" & block).item(j, 1).value
        End If
        If Sheet39.Range("Salary.Section" & block).item(j, 1).value = "Others" Then
            Othrs = Othrs + Sheet39.Range("Salary.SalaryAmount" & block).item(j, 1).value
        End If
        
        
        'Newly added by Bindu as its Legacy
        If Sheet39.Range("Salary.Section" & block).item(j, 1).value = "Gratuity" Then
            Gratuity_Sal = Gratuity_Sal + Sheet39.Range("Salary.SalaryAmount" & block).item(j, 1).value
        End If
        
        'Newly added by Bindu as its Legacy
        If Sheet39.Range("Salary.Section" & block).item(j, 1).value = "Leave Travel Allowance(LTA)" Then
            Leave_Sal = Leave_Sal + Sheet39.Range("Salary.SalaryAmount" & block).item(j, 1).value
        End If
        
        
        
    Next j

Next block

Dim k As Long
For k = 1 To Sheet39.Range("Salary.Nature1").count
    If Sheet39.Range("Salary.Nature1").item(k, 1).value = "Sec 10(14)(i)-Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment" Then
        Ex1014i = Ex1014i + Sheet39.Range("Salary.Amount1").item(k, 1).value
    End If
    If Sheet39.Range("Salary.Nature1").item(k, 1).value = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living" Then
        Ex1014ii = Ex1014ii + Sheet39.Range("Salary.Amount1").item(k, 1).value
    End If
    
    
    'Newly added by Bindu as its Legacy
    If Trim(Sheet39.Range("Salary.Nature1").item(k, 1).value) = Trim("Sec 10(10)-Death-cum-retirement gratuity received") Then
        Ex1014iii_Gratuity = Ex1014iii_Gratuity + Sheet39.Range("Salary.Amount1").item(k, 1).value
    End If
    
    'Newly added by Bindu as its Legacy
'    If Trim(Sheet39.Range("Salary.Nature1").item(k, 1).value) = Trim("Sec 10(5)-Leave Travel concession/assistance received") Or Trim(Sheet39.Range("Salary.Nature1").item(k, 1).value) = Trim("Sec 10(5)-Leave Travel concession/assistance") Then
    If Trim(Sheet39.Range("Salary.Nature1").item(k, 1).value) = Trim("Sec 10(5)- Travel concession/assistance received") Then 'Added by Aavula Naresh on 16/05/2025
        Ex1014iiv_Leave = Ex1014iiv_Leave + Sheet39.Range("Salary.Amount1").item(k, 1).value
    End If
    
    
Next k
        
If (Ex1014i > ConvAll + OthAll + Othrs) Then
    MsgScheduleS1 = MsgScheduleS1 + "* Exemption u/s 10(14)(i) cannot be more than the value of Salary under  'Conveyance Allowance', 'Other Allowance' & 'Others' in Section 17(1)"
End If

If (Ex1014ii > ConvAll + OthAll + Othrs) Then
    MsgScheduleS1 = MsgScheduleS1 + "* Exemption u/s 10(14)(ii) cannot be more than the value of Salary under  'Conveyance Allowance', 'Other Allowance' & 'Others' in Section 17(1)"
End If


'Newly added by Bindu as its Legacy
If (Ex1014iii_Gratuity > Gratuity_Sal) Then
    MsgScheduleS1 = MsgScheduleS1 + "* Exemption u/s 10(10) cannot be more than the value of Salary under  'Gratuity' in Section 17(1)"
End If

If (Ex1014iiv_Leave > Leave_Sal) Then
    MsgScheduleS1 = MsgScheduleS1 + "* Exemption u/s 10(15) cannot be more than the value of Salary under  'Leave Travel Allowance(LTA)' in Section 17(1)"
End If



     If MsgScheduleS1 <> "" Then
    fmsgbox MsgScheduleS1 ', vbOKOnly, "ITR 3 : 2022-23"
    CloseMsg
    End If
    
  '    '============================================
    
     endEA10_13A = 0
    endEA10_13A_1 = 0
    endEA10_13A_2 = 0
    endEA10_13A_3 = 0
    endEA10_13A_4 = 0
    endEA10_13A_5 = 0
    endEA10_13A_6 = 0


            EA_Sch10of13A_PlaceofWrk = Sheet39.Range("Sch10of13A_PlaceofWrk")
            EA_Sch10of13A_ActlHRArecivedA = Sheet39.Range("Sch10of13A_ActlHRArecivedA")
            EA_Sch10of13A_ActlRentpaid = Sheet39.Range("Sch10of13A_ActlRentpaid")
            EA_Sch10of13A_DetlsofSalpersec17of1 = Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1")
            EA_Sch10of13A_Actlrentpaid10persalaryB = Sheet39.Range("Sch10of13A_Actlrentpaid10persalaryB")
            EA_Sch10of13A_50Por40Pofsalary = Sheet39.Range("Sch10of13A_50Por40Pofsalary")
            EA_Sch10of13A_ElgiblExmptAllwnce10of13A = Sheet39.Range("Sch10of13A_ElgiblExmptAllwnce10of13A")
            
    
     If EA_Sch10of13A_PlaceofWrk <> "" Then
        If Not isdropdownblank(EA_Sch10of13A_PlaceofWrk) And Not UCase(EA_Sch10of13A_PlaceofWrk) = UCase("Select") Then
        endEA10_13A_1 = endEA10_13A_1 + 1
        End If
     End If
     
     If EA_Sch10of13A_ActlHRArecivedA <> "" And EA_Sch10of13A_ActlHRArecivedA <> 0 Then
        If EA_Sch10of13A_ActlHRArecivedA <> "" Then
        endEA10_13A_2 = endEA10_13A_2 + 1
        End If
     End If
      
      If EA_Sch10of13A_ActlRentpaid <> "" Then
        If EA_Sch10of13A_ActlRentpaid <> "" Then
        endEA10_13A_3 = endEA10_13A_3 + 1
        End If
     End If
     
     If EA_Sch10of13A_BasicSalary <> "" Then
        If EA_Sch10of13A_BasicSalary <> "" Then
        endEA10_13A_4 = endEA10_13A_4 + 1
        End If
     End If
     
     If EA_Sch10of13A_DearAllowance <> "" Then
        If EA_Sch10of13A_DearAllowance <> "" Then
        endEA10_13A_5 = endEA10_13A_5 + 1
        End If
     End If
     
          

     If EA_Sch10of13A_50Por40Pofsalary <> "" Then
        If EA_Sch10of13A_50Por40Pofsalary <> "" Then
        endEA10_13A_6 = endEA10_13A_6 + 1
        End If
     End If
     
endEA10_13A = WorksheetFunction.Max(0, endEA10_13A_1, endEA10_13A_2, endEA10_13A_3, endEA10_13A_4, endEA10_13A_5, endEA10_13A_6)
        

 If endEA10_13A > 0 Then

If Sheet39.Range("Sch10of13A_PlaceofWrk").value <> "(Select)" Or Sheet39.Range("Sch10of13A_ActlHRArecivedA").value > 0 Or Sheet39.Range("Sch10of13A_ActlRentpaid").value > 0 Or Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").value > 0 Or Sheet39.Range("Sch10of13A_Actlrentpaid10persalaryB").value > 0 Then
        
        'Place of Residence
        If isdropdownblank(EA_Sch10of13A_PlaceofWrk) Then
        
        'changed by Chetan C M for AY 2026-27 as per DE V0.2
        'start
            'MsgScheduleS1 = MsgScheduleS1 + "* Please select dropdown from ""Place of work"" in schedule S." & Chr(13)
            MsgScheduleS1 = MsgScheduleS1 + "* Please select dropdown from ""Place of residence"" in schedule S." & Chr(13)
        '--end
         
        End If
        
        'Actual HRA received (A)
        If Not chkCompulsory(EA_Sch10of13A_ActlHRArecivedA) Or EA_Sch10of13A_ActlHRArecivedA = 0 Then
        Else
            If Len(EA_Sch10of13A_ActlHRArecivedA) > 14 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Actual HRA received (A) cannot be greater than 14 digits in schedule S""" & Chr(13)

            End If

             If Not IsNumeric(EA_Sch10of13A_ActlHRArecivedA) Or EA_Sch10of13A_ActlHRArecivedA < 0 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Actual HRA received (A) in schedule S""" & Chr(13)
             End If
         End If
        
        'Actual Rent paid
        If Not chkCompulsory(EA_Sch10of13A_ActlRentpaid) Or EA_Sch10of13A_ActlRentpaid = 0 Then
            MsgScheduleS1 = MsgScheduleS1 + "* ""Please provide actual rent paid in schedule S""" & Chr(13)
        Else
            If Len(EA_Sch10of13A_ActlRentpaid) > 14 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Actual Rent paid cannot be greater than 14 digits in schedule S""" & Chr(13)
            End If
            If Not IsNumeric(EA_Sch10of13A_ActlRentpaid) Or EA_Sch10of13A_ActlRentpaid < 0 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Actual Rent paid in schedule S""" & Chr(13)
            End If
        End If
        
        'Details of Salary as per section 17(1)
        If Not chkCompulsory(EA_Sch10of13A_DetlsofSalpersec17of1) Or EA_Sch10of13A_DetlsofSalpersec17of1 = 0 Then
            MsgScheduleS1 = MsgScheduleS1 + "* ""Please provide the salary considered for claiming exemption u/s 10(13A) in schedule S""" & Chr(13)
        Else
            If Len(EA_Sch10of13A_DetlsofSalpersec17of1) > 14 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Details of Salary as per section 17(1) cannot be greater than 14 digits in schedule S""" & Chr(13)
            End If
            If Not IsNumeric(EA_Sch10of13A_DetlsofSalpersec17of1) Or EA_Sch10of13A_DetlsofSalpersec17of1 < 0 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Details of Salary as per section 17(1) in schedule S""" & Chr(13)
            End If
        End If
        
        
        ''Note:-Actual HRA received(2) + Details of salary (4) shall not be more than salary as per 17(1) in field i(a) as per part A gen
        If EA_Sch10of13A_ActlHRArecivedA > 0 Or EA_Sch10of13A_DetlsofSalpersec17of1 > 0 Then
        EA_Sch10of13A_2Plus4 = EA_Sch10of13A_ActlHRArecivedA + EA_Sch10of13A_DetlsofSalpersec17of1

        End If
        
            
        'Actual rent paid-10% of salary (B)
        If EA_Sch10of13A_Actlrentpaid10persalaryB <> "" Then
            If Len(EA_Sch10of13A_Actlrentpaid10persalaryB) > 14 Then
            MsgScheduleS1 = MsgScheduleS1 + "* ""Actual rent paid-10% of salary (B) cannot be greater than 14 digits in schedule S""" & Chr(13)
            End If
            If Not IsNumeric(EA_Sch10of13A_Actlrentpaid10persalaryB) Or EA_Sch10of13A_Actlrentpaid10persalaryB < 0 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Actual rent paid-10% of salary (B) in schedule S""" & Chr(13)

            End If
        End If
            
        '50% /40% of salary (C)
        If Not chkCompulsory(EA_Sch10of13A_50Por40Pofsalary) Or EA_Sch10of13A_50Por40Pofsalary = 0 Then
            MsgScheduleS1 = MsgScheduleS1 + "* ""Please provide 50% /40% of salary (C) in schedule S""" & Chr(13)
         Else
            
           Dim Per50Salaryaspersection17 As Variant
              Per50Salaryaspersection17 = EA_Sch10of13A_DetlsofSalpersec17of1 * 0.5
            If Per50Salaryaspersection17 <> "" Then
                If EA_Sch10of13A_50Por40Pofsalary > Per50Salaryaspersection17 Then
                  MsgScheduleS1 = MsgScheduleS1 + "* ""The value in this Field 6 ""50% /40% of salary (C)"" shall not be more than 50% of Field 4 ""Details of Salary as per Section 17(1)""""" & Chr(13)
                End If
            End If
            
            If Len(EA_Sch10of13A_50Por40Pofsalary) > 14 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""50% /40% of salary (C) cannot be greater than 14 digits in schedule S""" & Chr(13)
            End If
            
            If Not IsNumeric(EA_Sch10of13A_50Por40Pofsalary) Or EA_Sch10of13A_50Por40Pofsalary < 0 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at 50% /40% of salary (C) in schedule S""" & Chr(13)
            End If
        End If
            
        'Eligible Exempt Allowance u/s 10(13A)
        If EA_Sch10of13A_ElgiblExmptAllwnce10of13A <> "" Then
            If Len(EA_Sch10of13A_ElgiblExmptAllwnce10of13A) > 14 Then
            MsgScheduleS1 = MsgScheduleS1 + "* ""Eligible Exempt Allowance u/s 10(13A) cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
            End If
            If Not IsNumeric(EA_Sch10of13A_ElgiblExmptAllwnce10of13A) Or EA_Sch10of13A_ElgiblExmptAllwnce10of13A < 0 Then
                MsgScheduleS1 = MsgScheduleS1 + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Eligible Exempt Allowance u/s 10(13A) in schedule S""" & Chr(13)
            End If
        End If
            
            
  
  End If
    
        If MsgScheduleS1 <> "" Then
    fmsgbox MsgScheduleS1 ', vbOKOnly, "ITR 3 : 2022-23"
    CloseMsg
    End If
    
    '============================================
    
   End If
    
   
    '--------
'    End If
'    End If

End Sub


Function validate10aaAmount()
    Dim count
    Dim isGovt, istan, TANCode, rCell
    isGovt = False
    istan = False
     For count = 1 To Sheet39.Range("SalarySectionCount").value
     
'Change-02.11.2023.102.Ver.0.6
'        If Sheet39.Range("sheet.EmployerCategory" & count).value = "Central Government" Or Sheet39.Range("sheet.EmployerCategory" & count).value = "State Government" Then
        If Sheet39.Range("sheet.EmployerCategory" & count).value = "Central Government" Or Sheet39.Range("sheet.EmployerCategory" & count).value = "State Government" Or Sheet39.Range("sheet.EmployerCategory" & count).value = "Public Sector Undertaking" Or Sheet39.Range("sheet.EmployerCategory" & count).value = "PSU-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & count).value = "SG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & count).value = "CG-Pensioners" Then
' End Change
             If isGovt = False Then
                isGovt = True
             End If
        End If
        If (isGovt = False) Then
        If (istan = False) Then
        
        For Each TANCode In SheetALL.Range("TanCodeSalary")
            If CStr(TANCode.value) = CStr(Sheet39.Range("SAL.PANofEmployer" & count).value) Then
                If istan = False Then
                    istan = True
                End If
                Exit For
            End If
        Next TANCode
        End If
        End If
     Next
     If isGovt = False Then
        If istan = False Then
            For Each rCell In Range("Salary.Nature1").Cells
                If rCell.value = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
                    If Range("J" & rCell.row).value > 300000 Then
                        fmsgboxStatus "Deduction u/s 10(10AA) shall be restricted to Rs. 3 Lakh in case of  employer catergory is selected as other than Central or State Government"
 
                        CloseMsg
                    End If
                End If
            Next
        End If
    End If
End Function
Sub validateSchSold1()
Dim scount, count As Variant
MsgScheduleS = ""
MsgScheduleS1 = ""
subProcCaption = "Validating Schedule S"
   If Not validateScheduleS Then
        Sheet39.Activate
        scount = Sheet39.Range("SalarySectionCount").value
        
        For count = 1 To scount
                MsgScheduleS = MsgScheduleS & Chr(13) & MsgPartBScheduleS(count)
        Next
        
     
    fmsgboxsmall MsgScheduleS ', vbOKOnly, "ITR 3 : 2023-24"
    CloseMsg
    End If
    
    If Len((Range("Salary.Nature1").item(1).value) > 0) Then
        If Not ValidateNatureOfIncome Then
        End If
        If Not ValidateAmount Then
        End If
    End If
    
    'd
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then
    If (Sheet39.Range("SAL.ExemptUSectionOth1").value - Sheet19.Range("Temp_HRA_Sec13A").value) > (Sheet39.Range("SAL.TotalGrossSalary").value - Sheet19.Range("Temp_HRA").value) Then
     MsgScheduleS1 = MsgScheduleS1 + "* Total of allowances excluding HRA shall not exceed total of gross salary at sr. no. 2 as reduced by HRA in Schedule Salary." & Chr(13)
'     ValidateLessAllowanceTable = False
    End If
    End If
     
    setTblinfo_OthersNOI
    setTblinfo_OthersAmt

    If ((end_OthersNOI_s <> end_OthersAmt_s)) Then
        MsgScheduleS1 = MsgScheduleS1 + "* Enter all mandatory fields in Table Less allowances to the extent exempt u/s 10 in Sheet Salary" & Chr(13)
    End If
    
     If Sheet39.Range("SAL.Entertainment_allowance_1").value > 5000 Then
        MsgScheduleS1 = MsgScheduleS1 + "* Entertainment allowance u/s 16(ii) should not be greater than 5000"
        'validateScheduleS = False
    End If

    If Sheet39.Range("SAL.DeductionUnderSection1").value > 5000 Then
        fmsgbox "*  Deduction of Professional tax u/s 16(iii) is Rs. 200/- p.m. in most of the state. Kindly reconsider the amount claimed"
        'validateSchS = False
    End If
    
    
    If MsgScheduleS1 <> "" Then
    fmsgbox MsgScheduleS1 ', vbOKOnly, "ITR 3 : 2023-24"
    CloseMsg
    End If
    
    

End Sub

Function ValidateBankAccountNumber(BankAccountNumber As Variant) As Boolean
    ValidateBankAccountNumber = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB() As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
    BankAccountNumber = Sheet7.Range("Sheet9.BankAccountNumber")
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        errmsgVerification = errmsgVerification & "Bank Account Number in Sheet : Taxes Paid and Verification   is Mandatory and must not exceed 20 characters " & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If

    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI. Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI cannot be 0" & Chr(13)
            ValidateBankAccountNumber = False
            Exit Function
        End If
    
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI is Mandatory for all cases" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI spaces, // and -- are not allowed" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI cannot start with / or -" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI cannot start with Special character" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        errmsgVerification = errmsgVerification & "Last Character of the Account Number should be a number in Sheet Part B - TI TTI" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        errmsgVerification = errmsgVerification & "Last Character of the Account Number should be a number in Sheet Part B - TI TTI" & Chr(13)
        ValidateBankAccountNumber = False
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
            errmsgVerification = errmsgVerification & "In Account Number, Any numeric string between alphabets or special characters cannot be all zeros in Sheet Part B - TI TTI." & Chr(13)
            ValidateBankAccountNumber = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "Bank Account Number in Sheet Part B - TI TTI is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If
End Function

Function ValidateIBANSELECT() As Boolean
ValidateIBANSELECT = True
    IBANSELECT = Sheet7.Range("Sheet9.IBANSELECT").value
    IBANSELECT = Mid(IBANSELECT, 1, 1)
    If isdropdownblank(IBANSELECT) Then
        IBANSELECT = ""
    End If
    
    If IBANSELECT = "" Or IsEmpty(IBANSELECT) Then
        msgbox9 ("Do you have any Bank Account in India? in Sheet PART B TI TTI  is mandatory")
        ValidateIBANSELECT = False
        Exit Function
    End If
End Function

Function ValidateIBAN() As Boolean
ValidateIBAN = True
    
    Dim i As Long
    Dim rangecells As Range
    setTblinfo_IBAN_1
    
    Set rangecells = Range("Sheet9.IBAN").Cells
    ReDim IBAN(end_IBAN)
    For i = 1 To end_IBAN
        IBAN(i) = rangecells.item(i).value
    
        If (Len(IBAN(i)) < 1) Or (Len(IBAN(i)) > 30) Then
            msgbox9 ("IBAN/SWIFT Code  at Sr.NO " & i & "  in Sheet PARTB - TI - TTI   is mandatory and must not exceed 30 characters ")
            ValidateIBAN = False
            Exit Function
        End If

        If Len(IBAN(i)) > 0 Then
            If Not checkfieldspecialcharacter(IBAN(i)) Then
                msgbox9 ("IBAN/SWIFT Code  at Sr.NO " & i & " in Sheet PART B TI TTI characters < > & ' " & Chr(34) & " are not allowed")
                ValidateIBAN = False
                Exit Function
            End If
        End If
    Next
End Function

Function ValidateIBANNameOfBank() As Boolean
ValidateIBANNameOfBank = True
    
    Dim i As Long
    Dim rangecells As Range
    setTblinfo_IBAN_1
    
    Set rangecells = Range("Sheet9.IBANBANK").Cells
    ReDim IBANNameOfBank(end_IBAN)
    For i = 1 To end_IBAN
        IBANNameOfBank(i) = rangecells.item(i).value
    
        If (Len(IBANNameOfBank(i)) < 1) Or (Len(IBANNameOfBank(i)) > 125) Then
            msgbox9 ("Name of the Bank  at Sr.NO " & i & " in Sheet PARTB - TI - TTI   is mandatory and must not exceed 125 characters ")
            ValidateIBANNameOfBank = False
            Exit Function
        End If

        If Len(IBANNameOfBank(i)) > 0 Then
            If Not checkfieldspecialcharacter(IBANNameOfBank(i)) Then
                msgbox9 ("Name of the Bank  at Sr.NO " & i & " in Sheet PART B TI TTI characters < > & ' " & Chr(34) & " are not allowed")
                ValidateIBANNameOfBank = False
                Exit Function
            End If
        End If
    Next
End Function

Function ValidateIBANBankAccountNumber_9() As Boolean
ValidateIBANBankAccountNumber_9 = True
    
    Dim i As Long
    Dim rangecells As Range
    setTblinfo_IBAN_1
    
    Set rangecells = Range("Sheet9.IBANACCNO").Cells
    ReDim IBANACCNO(end_IBAN)
    For i = 1 To end_IBAN
        IBANACCNO(i) = rangecells.item(i).value
    
        If (Len(IBANACCNO(i)) < 1) Or (Len(IBANACCNO(i)) > 40) Then
            msgbox9 ("IBAN at Sr.NO " & i & " in Sheet PARTB - TI - TTI   is mandatory and must not exceed 40 characters ")
            ValidateIBANBankAccountNumber_9 = False
            Exit Function
        End If

        If Not ValidateBankAccountNumber_FA(IBANACCNO(i), i - 1) Then
            ValidateIBANBankAccountNumber_9 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIBANCountry_9() As Boolean
ValidateIBANCountry_9 = True
    Dim i As Long
    Dim rangecells As Range
    Dim endto As Variant

    setTblinfo_IBAN_1

    Set rangecells = Range("Sheet9.IBANCOUNTRY").Cells
    ReDim IBANCOUNTRY(end_IBAN)
    For i = 1 To end_IBAN
        IBANCOUNTRY(i) = rangecells.item(i).value

        If isdropdownblank(IBANCOUNTRY(i)) Then
            IBANCOUNTRY(i) = ""
        End If

        If IBANCOUNTRY(i) = "" Or IsEmpty(IBANCOUNTRY(i)) Then
            msgbox9 ("Country  at Sr.NO " & i & " in Sheet PART B TI TTI  is mandatory")
            ValidateIBANCountry_9 = False
            Exit Function
        End If

        endto = Split(IBANCOUNTRY(i), "-")
        IBANCOUNTRY(i) = endto(0)
        If IBANCOUNTRY(i) = "" Or IsEmpty(IBANCOUNTRY(i)) Then
            msgbox9 ("Country  at Sr.NO " & i & " in Sheet PART B TI TTI  is mandatory")
            ValidateIBANCountry_9 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIFSCCode(IFSCCode As Variant) As Boolean
    ValidateIFSCCode = True

    If Len(IFSCCode) = 11 Then
        If Mid(IFSCCode, 5, 1) <> 0 Then
            ValidateIFSCCode = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(IFSCCode, 1, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 2, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 3, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 4, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 6, 1)) And Not IsNumeric(Mid(IFSCCode, 6, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 7, 1)) And Not IsNumeric(Mid(IFSCCode, 7, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 8, 1)) And Not IsNumeric(Mid(IFSCCode, 8, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 9, 1)) And Not IsNumeric(Mid(IFSCCode, 9, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 10, 1)) And Not IsNumeric(Mid(IFSCCode, 10, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    
        If Not CheckAtoZ(Mid(IFSCCode, 11, 1)) And Not IsNumeric(Mid(IFSCCode, 11, 1)) Then
            ValidateIFSCCode = False
            Exit Function
        End If
    Else
       ValidateIFSCCode = False
    End If
End Function


Function ValidateBankAccountType_9() As Boolean
    ValidateBankAccountType_9 = True
    BankAccountType_9 = Sheet7.Range("Sheet9.BankAccountType1").value
    BankAccountType_9 = Mid(BankAccountType_9, 1, 3)
    
    If isdropdownblank(BankAccountType_9) Then
        BankAccountType_9 = ""
    End If
    
    If UCase(BankAccountType_9) = "CAS" Then
        BankAccountType_9 = "CCT"
    End If
    
    If Not chkCompulsory(BankAccountType_9) Then
        msgbox9 ("BankAccountType in Sheet PART B TI TTI is Mandatory" & Chr(13))
        ValidateBankAccountType_9 = False
        Exit Function
    End If
End Function

Function ValidateAssesseeVerName_9() As Boolean
    ValidateAssesseeVerName_9 = True
    AssesseeVerName_9 = Sheet7.Range("sheet9.AssesseeVerName")
    
    If Len(AssesseeVerName_9) > 125 Then
        msgbox9 ("AssesseeVerName in Sheet : PART B - TI - TTI    should be at most 125 characters " & Chr(13))
        ValidateAssesseeVerName_9 = False
        Exit Function
    End If
    
    If AssesseeVerName_9 = "" Or IsEmpty(AssesseeVerName_9) Then
        msgbox9 ("AssesseeVerName in Sheet : PART B - TI - TTI    is Mandatory" & Chr(13))
        ValidateAssesseeVerName_9 = False
        Exit Function
    End If
End Function
                                                                 
Function ValidateDate_9() As Boolean
    Dim i As Long
    
    ValidateDate_9 = True
    Date_9 = Sheet52.Range("sheet9.Date").value
    If Date_9 = "" Or IsEmpty(Date_9) Then
        msgbox9 ("Date in Sheet : PART B - TI - TTI    is Mandatory" & Chr(13))
        ValidateDate_9 = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(Date_9) Then
        ValidateDate_9 = False
        msgbox9 ("Date in Sheet : PART B - TI - TTI    must be a valid dd/mm/yyyy format" & Chr(13))
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
    
    If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2022, "Verification date cannot be less than 01/04/2022") Then
        ValidateDate_9 = False
        msgbox9 ("Date in Verification , Sheet PART B - TI - TTI must not be less than 01/04/2022" & Chr(13))
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
End Function
                                                                    

Function ValidateAOIFlag() As Boolean
    ValidateAOIFlag = True
    AOIFlag = Sheet7.Range("AOIFLAG_1").value
    'ctFlag = Sheet7.Range("CTCheck").value
    
    If isdropdownblank(AOIFlag) Then
        AOIFlag = ""
    End If
    If Not (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") Then
      If Sheet7.Range("AOIFLAG_1").Locked = False Then 'Added by Aavula Naresh
        If Len(Trim(AOIFlag)) = 0 Then
            msgbox9 ("Please select the option from Drop Down. in sheet PARTB-TI TTI" & Chr(13))
            ValidateAOIFlag = False
            Exit Function
        End If
      End If ''
    
    End If
    

End Function


Function ValidateFTFlag() As Boolean
Dim FTFlag As Variant

End Function


Function validateScheduleS() As Boolean
'On Error Resume Next
    validateScheduleS = True
    noOfSalFilled = 0
   Dim i As Variant
SalCount = Sheet39.Range("SalarySectionCount").value
  ReDim MsgPartBScheduleS(SalCount) As Variant
    noOfProcessSub = SalCount
ReDim SchSSalary1(SalCount), SchSEmployer(SalCount), SchSExemptUSectionTravel1(SalCount), SchSExemptUSectionNonMonetary1(SalCount), _
SchSExemptUSectionHouseRent1(SalCount), SchSExemptUSectionOth1(SalCount), SchSAllowancesNotExempt1(SalCount), _
SchSValueOfPerquisites1(SalCount), SchSProfitsinLieuOfSalary1(SalCount), SchSDeductionUnder6Section1(SalCount), SchSDeductionUnderSection1(SalCount), SchSIncomeFromSalary1(SalCount), SchSEmpName(SalCount), SchSEmpAddr(SalCount), _
SchSEmpTown_City(SalCount), SchSEmpState(SalCount), SchSEmpPinCode(SalCount), SchSEmpPAN(SalCount), SchSEmpZipCode(SalCount), SALEntertainment(SalCount)

'SchSTotIncUnderHeadSalaries1 = Sheet39.Range("SAL.TotIncUnderHeadSalaries1").value
 For i = 1 To SalCount

    
    SchSSalary1(i) = Sheet39.Range("SAL.Salary" & i).value

    SchSAllowancesNotExempt1(i) = Sheet39.Range("SAL.AllowancesNotExempt" & i).value
    SchSValueOfPerquisites1(i) = Sheet39.Range("SAL.ValueOfPerquisites" & i).value
    SchSProfitsinLieuOfSalary1(i) = Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value
    SchSDeductionUnder6Section1 = Sheet39.Range("SAL.DeductionUnder6Section1").value
    SchSDeductionUnderSection1 = Sheet39.Range("SAL.DeductionUnderSection1").value
    SALEntertainment = Sheet39.Range("SAL.Entertainment_allowance_1").value
   ' SchSIncomeFromSalary1(i) = Sheet39.Range("SAL.IncomeFromSalary" & i + 1).value
    SchSEmpPAN(i) = Sheet39.Range("SAL.PANofEmployer" & i).value
    SchSEmpName(i) = Sheet39.Range("SAL.NameOfEmployer" & i).value
    SchSEmployer(i) = Sheet39.Range("sheet.EmployerCategory" & i).value
    SchSEmpAddr(i) = Sheet39.Range("SAL.AddrDetail" & i).value
    SchSEmpTown_City(i) = Sheet39.Range("SAL.CityOrTownOrDistrict" & i).value
    SchSEmpState(i) = Sheet39.Range("SAL.StateCode" & i).value
    SchSEmpPinCode(i) = Sheet39.Range("SAL.PinCode" & i).value
    SchSEmpZipCode(i) = Sheet39.Range("SAL.ZipCode" & i).value

    setTableInfo_SalarySection1 (i)
    setTableInfo_SalaryDescription1 (i)
    setTableInfo_SalaryAmount1 (i)
    
    setTableInfo_SalarySection2 (i)
    setTableInfo_SalaryDescription2 (i)
    setTableInfo_SalaryAmount2 (i)
    
    setTableInfo_SalarySection3 (i)
    setTableInfo_SalaryDescription3 (i)
    setTableInfo_SalaryAmount3 (i)

    If SchSEmpName(i) <> "" Then
        noOfSalFilled = noOfSalFilled + 1
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If
    ElseIf SchSEmployer(i) <> "" And SchSEmployer(i) <> "(Select)" Then
      noOfSalFilled = noOfSalFilled + 1
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If
        
        
    ElseIf SchSEmpPAN(i) <> "" Then
        If Not CheckTAN(SchSEmpPAN(i)) Then
                MsgPartBScheduleS(i) = MsgPartBScheduleS(i) + "* Tan number of employer is not valid." & Chr(13)
                validateScheduleS = False
                
               ' Exit Function
        End If
        
        If Not ValidateTANCodes(UCase(SchSEmpPAN(i))) Then
                MsgPartBScheduleS(i) = MsgPartBScheduleS(i) + "* Invalid TAN at Sr.No " & i & "  Schedule Salary.Please enter valid TAN" & Chr(13)
                validateScheduleS = False
                Exit Function
        End If
        
        If Not checKMandatoryFields(i) Then
                validateScheduleS = False
                
         End If
         
         
    ElseIf SchSEmpAddr(i) <> "" Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If
    ElseIf SchSEmpTown_City(i) <> "" Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If
    ElseIf Not ((SchSEmpState(i) = "") Or (SchSEmpState(i) = "(Select)")) Then
        'If Mid(SchSEmpState, 1, 1) <> "(" Then
            If Not checKMandatoryFields(i) Then
                validateScheduleS = False
                'Exit Function
         '   End If
        End If
        

        
        
   
    ElseIf (SchSSalary1(i) > 0 And SchSSalary1(i) <> "") Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If
    ElseIf (SchSAllowancesNotExempt1(i) <> "" And SchSAllowancesNotExempt1(i) > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If
    ElseIf (SchSValueOfPerquisites1(i) <> "" And SchSValueOfPerquisites1(i) > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
           ' Exit Function
        End If

    ElseIf (SchSProfitsinLieuOfSalary1(i) <> "" And SchSProfitsinLieuOfSalary1(i) > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
            'Exit Function
        End If

    ElseIf (NatureOfIncomeOthCount1 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    ElseIf (ExemptUSectionOthCount1 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    ElseIf (AnyOthDescriptionOthCount1 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
        
    ElseIf (NatureOfIncomeOthCount2 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    ElseIf (ExemptUSectionOthCount2 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    ElseIf (AnyOthDescriptionOthCount2 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
        
        ElseIf (NatureOfIncomeOthCount3 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    ElseIf (ExemptUSectionOthCount3 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    ElseIf (AnyOthDescriptionOthCount3 > 0) Then
        If Not checKMandatoryFields(i) Then
            validateScheduleS = False
        End If
    
    End If
    
    
    'Changed by sai on 16/10/24==============================================
    
    
    If SchSEmpState(i) <> "" Then
    Sheets("Schedule S").Activate
         Dim PIN_targetadd, state_targetadd As String
         Dim ws_S As Worksheet
     
      state_targetadd = Replace(Sheet39.Range("SAL.StateCode" & i).Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "H", "K")
        Set ws_S = Worksheets("Schedule S")
         If Not state_Validation_Level(ws_S, PIN_targetadd, state_targetadd) Then
             Sheet39.Range(Replace(state_targetadd, "H", "K")).value = ""
             MsgPartBScheduleS(i) = MsgPartBScheduleS(i) + "* Invalid Pincode" & Chr(13)
           validateScheduleS = False
         End If
    End If
    
        '============================================================
    
    

'Dim te As Long
'    For te = 1 To Sheet39.Range("SalarySectionCount").value
'        If Sheet39.Range("SAL.IncomeFromSalary1").value > 0 Or Sheet39.Range("SAL.TotalGrossSalary").value > 0 Or Sheet39.Range("SAL.ExemptUSectionOth1").value > 0 Or Sheet39.Range("SAL.Netsalary1").value > 0 Or Sheet39.Range("SAL.Deduction1").value > 0 Then
'          If Sheet39.Range("SAL.NameOfEmployer" & 1).value = "" Then
'          MsgPartBScheduleS(i) = MsgScheduleS1 + "*  Please enter Name of employer" & 1 & " in Salary Schedule" & Chr(13) & Chr(13) + MsgPartBScheduleS(i)
'            validateScheduleS = False
''            Exit Function
'        End If
'        End If
'    Next
'
'
'
    If MsgPartBScheduleS(i) <> "" Then
    MsgPartBScheduleS(i) = "Schedule S(Salaries) Block" & i & " in Salary Schedule" & Chr(13) & Chr(13) + MsgPartBScheduleS(i)
    End If
    UpdateProgressBar
Next
End Function

Function checKMandatoryFields(count As Variant) As Boolean
Dim temp As Variant
temp = Sheet39.Range("SalarySectionCount").value
On Error Resume Next

    checKMandatoryFields = True
    'MsgPartBScheduleS(SalCount) = "Schedule S(Salaries) Block" & SalCount &  Chr(13)
If Not SchSEmpPAN(count) = "" Then
If Not CheckTAN(SchSEmpPAN(count)) Then
                MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* TAN format should be First 4 Alphabets, next 5 digits, then 1 Alphabet." & Chr(13)
                checKMandatoryFields = False
                Exit Function
               ' Exit Function
End If
End If

If SchSEmpName(count) = "" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter Name of Employer in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
End If

If SchSEmployer(count) = "" Or SchSEmployer(count) = "(Select)" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Nature of employer is mandatory in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
End If

If SchSEmpAddr(count) = "" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter address of the employer in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
End If

If SchSEmpTown_City(count) = "" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter town/city of the employer in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
End If

If SchSEmpState(count) = "" Or Mid(SchSEmpState(count), 1, 1) = "(" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter state of employer in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
End If

If SchSEmpState(count) <> "" Or Mid(SchSEmpState(count), 1, 1) <> "(" Then
If Mid(SchSEmpState(count), 1, 2) = "99" Then
    If SchSEmpZipCode(count) = "" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter Zipcode in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
    End If
    
     If Not checkfieldspecialcharacter(SchSEmpZipCode(count)) Then
             MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* characters < > & ' " & Chr(34) & " are not allowed in ZipCode in Salary Schedule" & Chr(13)
            checKMandatoryFields = False
            Exit Function
        End If
Else
    If SchSEmpPinCode(count) = "" Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter Pincode in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
    End If
End If
End If


If Not SchSSalary1(count) > 0 Then
    MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Please enter Salary in Salary Schedule" & Chr(13)
    checKMandatoryFields = False
End If

setTableInfo_SalarySection1 (count)
    setTableInfo_SalaryDescription1 (count)
    setTableInfo_SalaryAmount1 (count)
  
Dim i As Long
NatureOfIncomeOthCount1 = WorksheetFunction.Max(0, NatureOfIncomeOthCount1, ExemptUSectionOthCount1)
For i = 1 To NatureOfIncomeOthCount1
    'If Len(Sheet39.Range("Salary.Section" & Count + 1).Cells.item(i).value) > 0 Then

        If isdropdownblank(Sheet39.Range("Salary.Section" & count).Cells.item(i).value) Then
           MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Nature of salary in  section 17(1) at Sr. No " & i & "in Salary Schedule" & Chr(13)
        checKMandatoryFields = False
        Exit Function
        End If
        
        
        If (Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value) = "" Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Amount  in section 17(1) at Sr. No " & i & "in Salary Schedule" & Chr(13)
            checKMandatoryFields = False
            Exit Function
        End If

'Change-04.05.2023.101. -> V0.4 Implementation
        If Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value <> "" And Sheet39.Range("Salary.Section" & count).Cells.item(i).value = "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH" Then
            Dim Calc_17_80CCH
            Calc_17_80CCH = (Sheet39.Range("SAL.AllowancesNotExempt" & count).value - Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value)
            If Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value > (Calc_17_80CCH * 0.3) Then
                MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "*Amount against Sl. No 17, cannot be more than 30% of total salary u/s 17(1) other than sl. no. 17." & Chr(13)
                checKMandatoryFields = False
                Exit Function
            End If
        End If
'----
        If Sheet39.Range("Salary.salaryDescription" & count).Cells.item(i).Locked = False Then
            If Sheet39.Range("Salary.salaryDescription" & count).Cells.item(i).value = "" Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Description for the selected section 17(1)at Sr. No " & i & "in Salary Schedule" & Chr(13)
            checKMandatoryFields = False
            Exit Function
            End If
        End If
     
    'End If
    
    'End If
'Change 102 30.11.2023 Ver.0.6 SIT-62849
'    If Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value <> "" And Sheet39.Range("Salary.Section" & count).Cells.item(i).value = "Leave Travel Allowance(LTA)" Then
'        If Sheet39.Range("Salary.Nature1").Cells.item(i).value = "Sec 10(5)-Leave Travel concession/assistance" Then
'            If Sheet39.Range("Salary.Amount1").Cells.item(i).value > Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value Then
'            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* value of drop down ""Sec 10(5)-Leave Travel concession/assistance"" cannot be more than the value of field 1a ""Leave Travel Allowance"" under salary as per section 17(1)." & Chr(13)
'            checKMandatoryFields = False
'            Exit Function
'            End If
'        End If
'    End If
If Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value <> "" And Sheet39.Range("Salary.Section" & count).Cells.item(i).value = "Leave Travel Allowance(LTA)" Then
        If Sheet39.Range("Salary.Nature1").Cells.item(i).value = "Sec 10(5)- Travel concession/assistance received" Then
            If Sheet39.Range("Salary.Amount1").Cells.item(i).value > Sheet39.Range("Salary.SalaryAmount" & count).Cells.item(i).value Then
'            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* value of drop down ""Sec 10(5)-Leave Travel concession/assistance"" cannot be more than the value of field 1a ""Leave Travel Allowance"" under salary as per section 17(1)." & Chr(13)

        'For SIT-69339
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* ""Kindly restrict the exemption u/s 10(5) to ""Leave Travel Allowance"" under Salary as per section 17(1)." & Chr(13)
            
            checKMandatoryFields = False
            Exit Function
            End If
        End If
    End If
'End Change Ver.0.6 SIT-62849

Next

setTableInfo_SalarySection2 (count)
    setTableInfo_SalaryDescription2 (count)
    setTableInfo_SalaryAmount2 (count)
  
'Dim i As Long
NatureOfIncomeOthCount2 = WorksheetFunction.Max(0, NatureOfIncomeOthCount2, ExemptUSectionOthCount2)
For i = 1 To NatureOfIncomeOthCount2
    'If Len(Sheet39.Range("Salary.ValueSection" & Count + 1).Cells.item(i).value) > 0 Then

        If isdropdownblank(Sheet39.Range("Salary.ValueSection" & count).Cells.item(i).value) Then
         MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Nature of  perquisites in Value of perquisites as per section 17(2) at Sr. No " & i & "in Salary Schedule" & Chr(13)
        checKMandatoryFields = False
        Exit Function
        End If
        
        If (Sheet39.Range("Salary.ValueAmount" & count).Cells.item(i).value) = "" Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Amount  in Value of perquisites as per section 17(2) at Sr. No " & i & "in Salary Schedule" & Chr(13)
            checKMandatoryFields = False
            Exit Function
        End If

        If Sheet39.Range("Salary.ValueDescription" & count).Cells.item(i).Locked = False Then
            If Sheet39.Range("Salary.ValueDescription" & count).Cells.item(i).value = "" Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Description for the selected Section in Value of perquisites as per section 17(2) " & i & "in Salary Schedule" & Chr(13)
            checKMandatoryFields = False
            Exit Function
            End If
        End If

    'End If
    
    'End If
Next
'Change 102 30.11.2023 Ver.0.6 SIT-62870
For i = 1 To Sheet39.Range("Salary.Nature1").count
    If Sheet39.Range("Salary.Nature1").Cells.item(i).value <> "" And Sheet39.Range("Salary.Nature1").Cells.item(i).value = "Sec 10(10CC)-Tax paid by employer on non-monetary perquisite" Then
'        If Sheet39.Range("Salary.Nature1").Cells.item(i).value = "Sec 10(5)-Leave Travel concession/assistance" Then
            If Sheet39.Range("Salary.Amount1").Cells.item(i).value > Sheet39.Range("SAL.ValueOfPerquisites" & count).value Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Kindly restrict the exemption u/s 10(10CC) to Salary as per section 17(2) " & Chr(13)
            checKMandatoryFields = False
            Exit Function
            End If
'        End If
    End If
Next
'End Change Ver.0.6 SIT-62870
setTableInfo_SalarySection3 (count)
    setTableInfo_SalaryDescription3 (count)
    setTableInfo_SalaryAmount3 (count)
  
'Dim i As Long
NatureOfIncomeOthCount3 = WorksheetFunction.Max(0, NatureOfIncomeOthCount3, ExemptUSectionOthCount3)
For i = 1 To NatureOfIncomeOthCount3
    'If Len(Sheet39.Range("Salary.ProfitSection" & Count + 1).Cells.item(i).value) > 0 Then

        If isdropdownblank(Sheet39.Range("Salary.ProfitSection" & count).Cells.item(i).value) Then
        MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Nature of Profit in lieu of Salary in Profit in lieu of salary as per section 17(3) at Sr. No " & i & "in Salary Schedule" & Chr(13)
        checKMandatoryFields = False
        Exit Function
        End If

        If (Sheet39.Range("Salary.ProfitAmount" & count).Cells.item(i).value) = "" Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Amount  in Profit in lieu of salary as per section 17(3) at Sr. No " & i & Chr(13)
            checKMandatoryFields = False
            Exit Function
        End If

        If Sheet39.Range("Salary.ProfitDescription" & count).Cells.item(i).Locked = False Then
            If Sheet39.Range("Salary.ProfitDescription" & count).Cells.item(i).value = "" Then
            MsgPartBScheduleS(count) = MsgPartBScheduleS(count) + "* Enter Description for the selected Section in Profit in lieu of salary as per section 17(3) at Sr. No " & i & "in Salary Schedule" & Chr(13)
            checKMandatoryFields = False
            Exit Function
            End If
        End If

   ' End If
   
    'End If
Next
    
End Function
Function Validatesheet8b() As Boolean
    Validatesheet8b = True
    
    If Not ValidateIncomeFromSal_8b() Then Validatesheet8b = False
    If Not ValidateIncomeFromHP_8b() Then Validatesheet8b = False
    If Not ValidateProfGainNoSpecBus_8b() Then Validatesheet8b = False
    If Not ValidateProfGainSpecBus_8b() Then Validatesheet8b = False
    If Not ValidateProfGainSpecifiedBus_8b() Then Validatesheet8b = False
    'If Not ValidateProfGainSpecifiedBus115BBG_8b() Then Validatesheet8b = False
    If Not ValidateTotProfBusGain_8b() Then Validatesheet8b = False
    If Not ValidateShortTermUs111A_8b() Then Validatesheet8b = False
    If Not ValidateShortTermOther_8b() Then Validatesheet8b = False
    If Not ValidateTotalShortTerm_8b() Then Validatesheet8b = False
    If Not ValidateLongTerm_8b() Then Validatesheet8b = False
    'If Not ValidateTotalCapGains_8b() Then Validatesheet8b = False
    
    'Change-22.02.2023.103.CG
    If Not ValidateTotalCapGains_8b() Then Validatesheet8b = False
    If Not ValidateShortTermLongTermTotal_8b() Then Validatesheet8b = False
    If Not ValidateCapGains30Per115BBH_8b() Then Validatesheet8b = False
    
    If Not ValidateOtherSrcThanOwnRaceHorse_8b() Then Validatesheet8b = False
    If Not ValidateWinLotteriesRacesGambling_8b() Then Validatesheet8b = False
    If Not ValidateFromOwnRaceHorse_8b() Then Validatesheet8b = False
    If Not ValidateTotIncFromOS_8b() Then Validatesheet8b = False
    If Not ValidateTotalTI_8b() Then Validatesheet8b = False
    If Not ValidateCurrentYearLoss_8b() Then Validatesheet8b = False
    If Not ValidateBalanceAfterSetoffLosses_8b() Then Validatesheet8b = False
    If Not ValidateBroughtFwdLossesSetoff_8b() Then Validatesheet8b = False
    If Not ValidateGrossTotalIncome_8b() Then Validatesheet8b = False
    If Not ValidateDeductionsUnderScheduleVIA_8b() Then Validatesheet8b = False
    If Not ValidateTotalIncome_8b() Then Validatesheet8b = False
    If Not ValidateNetAgricultureIncomeOrOtherIncomeForRate_8b() Then Validatesheet8b = False
    If Not ValidateAggregateIncome_8b() Then Validatesheet8b = False
    If Not ValidateLossesOfCurrentYearCarriedFwd_8b() Then Validatesheet8b = False
End Function

Function ValidateIncomeFromSal_8b() As Boolean
    Dim i As Long
    
    ValidateIncomeFromSal_8b = True
    IncomeFromSal_8b = Sheet7.Range("Sheet8b.Salaries").value
    
    If Len(IncomeFromSal_8b) > 14 Then
        fmsgboxStatus ("IncomeFromHP should not greater than 14 digits in Sheet PART - TI - TTI")
        ValidateIncomeFromSal_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(IncomeFromSal_8b)
        If Not IsNumeric(IncomeFromSal_8b) Then
            fmsgboxStatus ("IncomeFromHP must contain only digits from 0 to 9 in Sheet PART - TI - TTI")
            ValidateIncomeFromSal_8b = False
            Exit Function
        End If
    Next
End Function


Function ValidateIncomeFromHP_8b() As Boolean
    Dim i As Long
    
    ValidateIncomeFromHP_8b = True
    IncomeFromHP_8b = Sheet7.Range("Sheet8b.IncomeFromHP").value
    
    If Len(IncomeFromHP_8b) > 14 Then
        fmsgboxStatus ("IncomeFromHP should not greater than 14 digits in Sheet PART - TI - TTI")
        ValidateIncomeFromHP_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(IncomeFromHP_8b)
        If Not IsNumeric(IncomeFromHP_8b) Then
            fmsgboxStatus ("IncomeFromHP must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateIncomeFromHP_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateProfGainNoSpecBus_8b() As Boolean
    Dim i As Long
    
    ValidateProfGainNoSpecBus_8b = True
    ProfGainNoSpecBus_8b = Sheet7.Range("Sheet8b.ProfGainNoSpecBus").value
    
    If Len(ProfGainNoSpecBus_8b) > 14 Then
        fmsgboxStatus ("ProfGainNoSpecBus should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateProfGainNoSpecBus_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(ProfGainNoSpecBus_8b) Then
        fmsgboxStatus ("ProfGainNoSpecBus must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateProfGainNoSpecBus_8b = False
        Exit Function
    End If
End Function

Function ValidateProfGainSpecBus_8b() As Boolean
    Dim i As Long
    
    ValidateProfGainSpecBus_8b = True
    ProfGainSpecBus_8b = Sheet7.Range("Sheet8b.ProfGainSpecBus").value
    
    If Len(ProfGainSpecBus_8b) > 14 Then
        msgbox8b ("ProfGainSpecBus should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateProfGainSpecBus_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(ProfGainSpecBus_8b)
        If Not IsNumeric(ProfGainSpecBus_8b) Then
            msgbox8b ("ProfGainSpecBus must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateProfGainSpecBus_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateProfGainSpecifiedBus_8b() As Boolean
    Dim i As Long
    
    ValidateProfGainSpecifiedBus_8b = True
    ProfGainSpecifiedBus_8b = Sheet7.Range("Sheet8b.ProfGainSpecifiedBus").value
    ProfGain115BF_8b = Sheet7.Range("Sheet8b.TaxOnIncome115BF").value
    If Len(ProfGainSpecifiedBus_8b) > 14 Then
        msgbox8b ("ProfGainSpecifiedBus in Sheet should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateProfGainSpecifiedBus_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(ProfGainSpecifiedBus_8b)
        If Not IsNumeric(ProfGainSpecifiedBus_8b) Then
            msgbox8b ("ProfGainSpecifiedBus in Sheet must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateProfGainSpecifiedBus_8b = False
            Exit Function
        End If
    Next
    
   If Len(ProfGain115BF_8b) > 14 Then
        msgbox8b ("Tax on income from patent u/s 115BBF should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateProfGainSpecifiedBus_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(ProfGain115BF_8b)
        If Not IsNumeric(ProfGain115BF_8b) Then
            msgbox8b ("Tax on income from patent u/s 115BBF must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateProfGainSpecifiedBus_8b = False
            Exit Function
        End If
    Next
End Function


Function ValidateTotProfBusGain_8b() As Boolean
    Dim i As Long
    
    ValidateTotProfBusGain_8b = True
    TotProfBusGain_8b = Sheet7.Range("Sheet8b.TotProfBusGain").value
    
    If Len(TotProfBusGain_8b) > 14 Then
        msgbox8b ("TotProfBusGain should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotProfBusGain_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(TotProfBusGain_8b)
        If Not IsNumeric(TotProfBusGain_8b) Then
            msgbox8b ("TotProfBusGain must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateTotProfBusGain_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateShortTermUs111A_8b() As Boolean
    Dim i As Long
    ValidateShortTermUs111A_8b = True
       
'    ShortTermUs111A_8b = Sheet7.Range("Sheet8b.ShortTermUs111A").value

    ShortTerm30_8b = Sheet7.Range("Sheet8b.ShortTerm30").value
    
'    If Len(ShortTermUs111A_8b) > 14 Then
'        msgbox8b ("ShortTermUs111A should not greater than 14 digits in Sheet PART - TI - TTI ")
'        ValidateShortTermUs111A_8b = False
'        Exit Function
'    End If
    
    For i = 1 To Len(ShortTermUs111A_8b)
        If Not IsNumeric(ShortTermUs111A_8b) Then
            msgbox8b ("ShortTermUs111A must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateShortTermUs111A_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateShortTermOther_8b() As Boolean
    Dim i As Long
    
    ValidateShortTermOther_8b = True
    ShortTermOther_8b = Sheet7.Range("Sheet8b.ShortTermOther").value
    ShortTermSpslRate_8b = Sheet7.Range("Sheet8b.STCGspecialrate").value
    
    If Len(ShortTermOther_8b) > 14 Then
        msgbox8b ("ShortTermOther should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateShortTermOther_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(ShortTermOther_8b) Then
        msgbox8b ("ShortTermOther in Sheet must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateShortTermOther_8b = False
        Exit Function
    End If
    If Len(ShortTermSpslRate_8b) > 14 Then
        msgbox8b ("ShortTerm should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateShortTermOther_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(ShortTermSpslRate_8b) Then
        msgbox8b ("ShortTerm must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateShortTermOther_8b = False
        Exit Function
    End If
End Function

Function ValidateTotalShortTerm_8b() As Boolean
Dim i As Long
    ValidateTotalShortTerm_8b = True
    TotalShortTerm_8b = Sheet7.Range("Sheet8b.TotalShortTerm").value
    
    If Len(TotalShortTerm_8b) > 14 Then
        msgbox8b ("TotalShortTerm should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotalShortTerm_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(TotalShortTerm_8b) Then
        msgbox8b ("TotalShortTerm must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateTotalShortTerm_8b = False
        Exit Function
    End If
End Function

Function ValidateLongTerm_8b() As Boolean
    Dim i As Long
    
    ValidateLongTerm_8b = True
    LongTerm_8b = Sheet7.Range("Sheet8b.LongTerm").value
'    LongTerm_8b_NP = Sheet7.Range("Sheet8b.LongTermNP").value
    LongTerm_8b_SR = Sheet7.Range("Sheet8b.LTCGspecialrate").value
'    LongTerm_8b_P = Sheet7.Range("Sheet8b.LongTermP").value
    
    DeductionsUnder10Aor10AA_8b = Sheet7.Range("Sheet8b.DeductionsUnder10Aor10AA").value
    PartBchapterVIA_8b = Sheet7.Range("Sheet8b.PartBchapterVIA").value
    PartCchapterVIA_8b = Sheet7.Range("Sheet8b.PartCchapterVIA").value
    
    If Len(LongTerm_8b) > 14 Then
        msgbox8b ("LongTerm should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateLongTerm_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(LongTerm_8b)
        If Not IsNumeric(LongTerm_8b) Then
            msgbox8b ("LongTerm must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateLongTerm_8b = False
            Exit Function
        End If
    Next
    
    If Len(LongTerm_8b_SR) > 14 Then
        msgbox8b ("LongTerm should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateLongTerm_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(LongTerm_8b_SR)
        If Not IsNumeric(LongTerm_8b) Then
            msgbox8b ("LongTerm must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateLongTerm_8b = False
            Exit Function
        End If
    Next
End Function

'Function ValidateTotalCapGains_8b() As Boolean
'    Dim i As Long
'
'    ValidateTotalCapGains_8b = True
'    TotalCapGains_8b = Sheet7.Range("Sheet8b.TotalCapGains").value
'
'    If Len(TotalCapGains_8b) > 14 Then
'        msgbox8b ("TotalCapGains should not greater than 14 digits in Sheet PART - TI - TTI ")
'        ValidateTotalCapGains_8b = False
'        Exit Function
'    End If
'
'    For i = 1 To Len(TotalCapGains_8b)
'        If Not IsNumeric(TotalCapGains_8b) Then
'            msgbox8b ("TotalCapGains must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
'            ValidateTotalCapGains_8b = False
'            Exit Function
'        End If
'    Next
'End Function

Function ValidateOtherSrcThanOwnRaceHorse_8b() As Boolean
    Dim i As Long
    
    ValidateOtherSrcThanOwnRaceHorse_8b = True
    OtherSrcThanOwnRaceHorse_8b = Sheet7.Range("Sheet8b.OtherSrcThanOwnRaceHorse").value
    
    If Len(OtherSrcThanOwnRaceHorse_8b) > 14 Then
        msgbox8b ("OtherSrcThanOwnRaceHorse should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateOtherSrcThanOwnRaceHorse_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(OtherSrcThanOwnRaceHorse_8b)
        If Not IsNumeric(OtherSrcThanOwnRaceHorse_8b) Then
            msgbox8b ("OtherSrcThanOwnRaceHorse must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateOtherSrcThanOwnRaceHorse_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateWinLotteriesRacesGambling_8b() As Boolean
    Dim i As Long
    
    ValidateWinLotteriesRacesGambling_8b = True
    WinLotteriesRacesGambling_8b = Sheet7.Range("sheet8b.WinLotteriesRacesGambling").value
    
    If Len(WinLotteriesRacesGambling_8b) > 14 Then
        msgbox8b ("WinLotteriesRacesGambling should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateWinLotteriesRacesGambling_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(WinLotteriesRacesGambling_8b)
        If Not IsNumeric(WinLotteriesRacesGambling_8b) Then
            msgbox8b ("WinLotteriesRacesGambling must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateWinLotteriesRacesGambling_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateFromOwnRaceHorse_8b() As Boolean
    Dim i As Long
    
    ValidateFromOwnRaceHorse_8b = True
    FromOwnRaceHorse_8b = Sheet7.Range("Sheet8b.FromOwnRaceHorse").value
    
    If Len(FromOwnRaceHorse_8b) > 14 Then
        msgbox8b ("FromOwnRaceHorse should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateFromOwnRaceHorse_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(FromOwnRaceHorse_8b)
        If Not IsNumeric(FromOwnRaceHorse_8b) Then
            msgbox8b ("FromOwnRaceHorse must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateFromOwnRaceHorse_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotIncFromOS_8b() As Boolean
    Dim i As Long
    
    ValidateTotIncFromOS_8b = True
    TotIncFromOS_8b = Sheet7.Range("Sheet8b.TotIncFromOS").value
    
    If Len(TotIncFromOS_8b) > 14 Then
        msgbox8b ("TotIncFromOS should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotIncFromOS_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(TotIncFromOS_8b)
        If Not IsNumeric(TotIncFromOS_8b) Then
            msgbox8b ("TotIncFromOS must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateTotIncFromOS_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotalTI_8b() As Boolean
    Dim i As Long
    
    ValidateTotalTI_8b = True
    TotalTI_8b = Sheet7.Range("Sheet8b.TotalTI").value
    
    If Len(TotalTI_8b) > 14 Then
        msgbox8b ("TotalTI should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotalTI_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(TotalTI_8b)
        If Not IsNumeric(TotalTI_8b) Then
            msgbox8b ("TotalTI must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateTotalTI_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateCurrentYearLoss_8b() As Boolean
    Dim i As Long
    
    ValidateCurrentYearLoss_8b = True
    CurrentYearLoss_8b = Sheet7.Range("Sheet8b.CurrentYearLoss").value
    
    If Len(CurrentYearLoss_8b) > 14 Then
        msgbox8b ("CurrentYearLoss should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateCurrentYearLoss_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(CurrentYearLoss_8b)
        If Not IsNumeric(CurrentYearLoss_8b) Then
            msgbox8b ("CurrentYearLoss must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateCurrentYearLoss_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateBalanceAfterSetoffLosses_8b() As Boolean
    Dim i As Long
    
    ValidateBalanceAfterSetoffLosses_8b = True
    BalanceAfterSetoffLosses_8b = Sheet7.Range("Sheet8b.BalanceAfterSetoffLosses").value
    
    If Len(BalanceAfterSetoffLosses_8b) > 14 Then
        msgbox8b ("BalanceAfterSetoffLosses should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateBalanceAfterSetoffLosses_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(BalanceAfterSetoffLosses_8b)
        If Not IsNumeric(BalanceAfterSetoffLosses_8b) Then
            msgbox8b ("BalanceAfterSetoffLosses must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateBalanceAfterSetoffLosses_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateBroughtFwdLossesSetoff_8b() As Boolean
    Dim i As Long
    
    ValidateBroughtFwdLossesSetoff_8b = True
    BroughtFwdLossesSetoff_8b = Sheet7.Range("Sheet8b.BroughtFwdLossesSetoff").value
    
    If Len(BroughtFwdLossesSetoff_8b) > 14 Then
        msgbox8b ("BroughtFwdLossesSetoff should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateBroughtFwdLossesSetoff_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(BroughtFwdLossesSetoff_8b)
        If Not IsNumeric(BroughtFwdLossesSetoff_8b) Then
            msgbox8b ("BroughtFwdLossesSetoff must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateBroughtFwdLossesSetoff_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateGrossTotalIncome_8b() As Boolean
    Dim i As Long
    
    ValidateGrossTotalIncome_8b = True
    GrossTotalIncome_8b = Sheet7.Range("Sheet8b.GrossTotalIncome").value
    IncChargeableTaxSplRates_8b = Sheet7.Range("Sheet8b.IncChargeableTaxSplRates").value
    IncChargeTaxSplRate111A112_8b = Sheet7.Range("Sheet8b.IncChargeTaxSplRate111A112").value
    
    If Len(GrossTotalIncome_8b) > 14 Then
        msgbox8b ("GrossTotalIncome should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateGrossTotalIncome_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(GrossTotalIncome_8b) Then
        msgbox8b ("GrossTotalIncome must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateGrossTotalIncome_8b = False
        Exit Function
    End If
End Function

Function ValidateDeductionsUnderScheduleVIA_8b() As Boolean
    Dim i As Long
    
    ValidateDeductionsUnderScheduleVIA_8b = True
    DeductionsUnderScheduleVIA_8b = Sheet7.Range("Sheet8b.DeductionsUnderScheduleVIA").value
    
    If Len(DeductionsUnderScheduleVIA_8b) > 14 Then
        msgbox8b ("DeductionsUnderScheduleVIA should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateDeductionsUnderScheduleVIA_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(DeductionsUnderScheduleVIA_8b)
        If Not IsNumeric(DeductionsUnderScheduleVIA_8b) Then
            msgbox8b ("DeductionsUnderScheduleVIA must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateDeductionsUnderScheduleVIA_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotalIncome_8b() As Boolean
    Dim i As Long
    
    ValidateTotalIncome_8b = True
    TotalIncome_8b = Sheet7.Range("Sheet8b.TotalIncome").value
    
    If Len(TotalIncome_8b) > 14 Then
        msgbox8b ("TotalIncome should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotalIncome_8b = False
        Exit Function
    End If
    
    If Len(Sheet7.Range("Sheet9.TotTaxRelief").value) > 14 Then
        msgbox8b ("Total Tax Relief should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotalIncome_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(TotalIncome_8b) Then
        msgbox8b ("TotalIncome in Sheet must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateTotalIncome_8b = False
        Exit Function
    End If
End Function

Function ValidateNetAgricultureIncomeOrOtherIncomeForRate_8b() As Boolean
    Dim i As Long
    
    ValidateNetAgricultureIncomeOrOtherIncomeForRate_8b = True
    NetAgricultureIncomeOrOtherIncomeForRate_8b = Sheet7.Range("Sheet8b.NetAgricultureIncomeOrOtherIncomeForRate").value
    
    If Len(NetAgricultureIncomeOrOtherIncomeForRate_8b) > 14 Then
        msgbox8b ("NetAgricultureIncomeOrOtherIncomeForRate should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateNetAgricultureIncomeOrOtherIncomeForRate_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(NetAgricultureIncomeOrOtherIncomeForRate_8b)
        If Not IsNumeric(NetAgricultureIncomeOrOtherIncomeForRate_8b) Then
            msgbox8b ("NetAgricultureIncomeOrOtherIncomeForRate must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateNetAgricultureIncomeOrOtherIncomeForRate_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateAggregateIncome_8b() As Boolean
    Dim i As Long
    
    ValidateAggregateIncome_8b = True
    AggregateIncome_8b = Sheet7.Range("Sheet8b.AggregateIncome").value
    
    If Len(AggregateIncome_8b) > 14 Then
        msgbox8b ("AggregateIncome in Sheet should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateAggregateIncome_8b = False
        Exit Function
    End If
    
    If Not IsNumeric(AggregateIncome_8b) Then
        msgbox8b ("AggregateIncome must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
        ValidateAggregateIncome_8b = False
        Exit Function
    End If
End Function


Function ValidateLossesOfCurrentYearCarriedFwd_8b() As Boolean
    Dim i As Long
    
    ValidateLossesOfCurrentYearCarriedFwd_8b = True
    LossesOfCurrentYearCarriedFwd_8b = Sheet7.Range("Sheet8b.LossesOfCurrentYearCarriedFwd").value
    DeemedTotIncSec115JC_8b = Sheet7.Range("Sheet8b.ProfGainSpecBus").value
    
    If Len(LossesOfCurrentYearCarriedFwd_8b) > 14 Then
        msgbox8b ("LossesOfCurrentYearCarriedFwd should not be greater than 14 digits  in Sheet PART - TI - TTI")
        ValidateLossesOfCurrentYearCarriedFwd_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(LossesOfCurrentYearCarriedFwd_8b)
        If Not IsNumeric(LossesOfCurrentYearCarriedFwd_8b) Then
            msgbox8b ("LossesOfCurrentYearCarriedFwd must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateLossesOfCurrentYearCarriedFwd_8b = False
            Exit Function
        End If
    Next
End Function
          
Function msgbox8b(strmsg As String) As String
    msgValidateSheetPartB = msgValidateSheetPartB & strmsg & Chr(13)
End Function

Function ValidateSheetPartBBlanks() As Boolean
    ValidateSheetPartBBlanks = True
End Function

Function msgbox9(strmsg As String) As String
    msgValidateSheet9 = msgValidateSheet9 & strmsg & Chr(13)
End Function

Sub LockUnlockPartBTTI()
On Error Resume Next
    If Mid(Trim(Sheet1.Range("Sheet1.115H").value), 1, 1) = "Y" Then
        UNLOCKPARTBTI
    Else
        LOCKPARTBTI
    End If
End Sub

Sub LOCKPARTBTI()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    sPassword = EfilingCommon.getmsgstate
    Sheet7.Unprotect Password:=sPassword
    
    Dim rngname_PARTBTI As Variant
    
    rngname_PARTBTI = "Sheet9.TaxAtNormalRatesOnAggrInc||sheet9.TaxAtSpecialRates||sheet9.RebateOnAgriInc||Sheet9.Surcharge_ii||Sheet9.EducationCess"
    LOCKRANGEPARTBTI (rngname_PARTBTI)
    
    
    Sheet7.Protect Password:=sPassword
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKPARTBTI()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    sPassword = EfilingCommon.getmsgstate
    Sheet7.Unprotect Password:=sPassword
    
    Dim rngname_PARTBTI As Variant
    
    rngname_PARTBTI = "Sheet9.TaxAtNormalRatesOnAggrInc||sheet9.TaxAtSpecialRates||sheet9.RebateOnAgriInc||Sheet9.Surcharge_ii||Sheet9.EducationCess"
    UNLOCKRANGEPARTBTI (rngname_PARTBTI)
    
    
    Sheet7.Protect Password:=sPassword
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKRANGEPARTBTI(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet7.Range(rangenamestring(i))
            chcell.Locked = False
            chcell.Interior.Color = 13434879
            chcell.value = 0
            Next
        Next
End Sub

Sub LOCKRANGEPARTBTI(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet7.Range(rangenamestring(i))
            chcell.Locked = True
            chcell.Interior.Color = 16777215
            chcell.value = 0
            Next
        Next
End Sub


Sub setTableInfo_SalarySection1(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.Section" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.Section" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    NatureOfIncomeOthCount1 = ccount
End Sub



Sub setTableInfo_SalaryDescription1(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.salaryDescription" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.salaryDescription" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    
    If ccount <= NatureOfIncomeOthCount1 Then
    AnyOthDescriptionOthCount1 = NatureOfIncomeOthCount1
    Else
    AnyOthDescriptionOthCount1 = ccount
    End If
End Sub

Sub setTableInfo_SalaryAmount1(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.SalaryAmount" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.SalaryAmount" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ExemptUSectionOthCount1 = ccount
End Sub


'2nd table

Sub setTableInfo_SalarySection2(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.ValueSection" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.ValueSection" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    NatureOfIncomeOthCount2 = ccount
End Sub



Sub setTableInfo_SalaryDescription2(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.ValueDescription" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.ValueDescription" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    
    If ccount <= NatureOfIncomeOthCount2 Then
    AnyOthDescriptionOthCount2 = NatureOfIncomeOthCount2
    Else
    AnyOthDescriptionOthCount2 = ccount
    End If
End Sub

Sub setTableInfo_SalaryAmount2(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.ValueAmount" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.ValueAmount" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ExemptUSectionOthCount2 = ccount
End Sub

'3rd table


Sub setTableInfo_SalarySection3(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.ProfitSection" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.ProfitSection" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    NatureOfIncomeOthCount3 = ccount
End Sub



Sub setTableInfo_SalaryDescription3(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.ProfitDescription" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.ProfitDescription" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    
    If ccount <= NatureOfIncomeOthCount3 Then
    AnyOthDescriptionOthCount3 = NatureOfIncomeOthCount3
    Else
    AnyOthDescriptionOthCount3 = ccount
    End If
End Sub

Sub setTableInfo_SalaryAmount3(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet39.Range("Salary.ProfitAmount" & cnt).Cells
    mIntCells = Sheet39.Range("Salary.ProfitAmount" & cnt).count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ExemptUSectionOthCount3 = ccount
End Sub

'end

Sub AddTotalSalary1()
  Dim Row7RangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, tot1, tot2, tot3, tot4, tot6, Totsections, otherTotal, Tempsal, TotalTempGrosssal As Double
    
    Dim Deduction80CCD2Total, AllowanceTotal As Double
    Dim Exemptincome As Double
    
    Tempsal = 0
    tot = 0
    tot1 = 0
    tot2 = 0
    tot3 = 0
    tot4 = 0
    tot6 = 0
    otherTotal = 0
    AllowanceTotal = 0
    Deduction80CCD2Total = 0
    Exemptincome = 0
    TotalTempGrosssal = 0
    
    'Row7RangeName = "SAL.IncomeFromSalary"
    Totsections = Sheet39.Range("SalarySectionCount").value
    
         sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
         
        For i = 1 To Totsections
        'newDefinedName = Row7RangeName & i
        
        
        tot2 = WorksheetFunction.Sum(Sheet39.Range("Salary.SalaryAmount" & i))
        Sheet39.Range("SAL.AllowancesNotExempt" & i).value = tot2
        
        '--Bindu 27/05/2025
    Dim BA_SAL, DA_SAL
    BA_SAL = 0
    DA_SAL = 0
    
       BA_SAL = BA_SAL + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & i), "Basic Salary", Sheet39.Range("Salary.SalaryAmount" & i))
       DA_SAL = DA_SAL + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & i), "Dearness Allowance(DA)", Sheet39.Range("Salary.SalaryAmount" & i))
''----Bindu
        
        
         tot3 = WorksheetFunction.Sum(Sheet39.Range("Salary.ValueAmount" & i))
        Sheet39.Range("SAL.ValueOfPerquisites" & i).value = tot3
        
         tot4 = WorksheetFunction.Sum(Sheet39.Range("Salary.ProfitAmount" & i))
        Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value = tot4
        
         tot6 = WorksheetFunction.Sum(Sheet39.Range("IncomeNotified89A_Amount" & i))
        Sheet39.Range("IncomeNotified89A" & i).value = tot6
        
        'Change-20.02.2023.102.Salary.
        'Sheet39.Range("SAL.Salary" & i).value = Sheet39.Range("SAL.AllowancesNotExempt" & i).value + Sheet39.Range("SAL.ValueOfPerquisites" & i).value + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value + Sheet39.Range("IncomeNotified89A" & i).value + Sheet39.Range("IncomeNotifiedOther89A" & i).value
        Sheet39.Range("SAL.Salary" & i).value = Sheet39.Range("SAL.AllowancesNotExempt" & i).value + Sheet39.Range("SAL.ValueOfPerquisites" & i).value + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value + Sheet39.Range("IncomeNotified89A" & i).value + Sheet39.Range("IncomeNotifiedOther89A" & i).value + Sheet39.Range("IncomeTaxablePrevYear89A" & i).value
        
        Tempsal = Tempsal + Sheet39.Range("SAL.Salary" & i).value
        
        tot = tot + Sheet39.Range("SAL.Salary" & i)
        tot1 = tot1 + Sheet39.Range("SAL.AllowancesNotExempt" & i) + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i)
        TotalTempGrosssal = TotalTempGrosssal + Sheet39.Range("SAL.Salary" & i).value
        
        Next
    
    Sheet39.Protect Password:=sPassword
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
  ' Sheet7.Unprotect Password:=sPassword
  ' Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Sheet39.Range("Tempgross").value = tot
    Sheet39.Range("Tempgross1").value = tot1
    
    '--Bindu 27/05/2025
    Sheet39.Range("SAL_Basic").value = BA_SAL
    Sheet39.Range("SAL_DA").value = DA_SAL
    
'Jyoti 28/5/2025
Dim HouseRentAllowance_count As Variant
    Sheet39.Range("HRA_SAL").value = HouseRentAllowance_count
 ''--Bindu
    
   ' Sheet7.Range("Sheet8b.Salaries").value = tot
    'Sheet19.Range("Salaryded").value = Deduction80CCD2Total
    Sheet39.Range("SAL.Salary1_Temp").value = Tempsal
    
    Sheet39.Range("SAL.TotalGrossSalary").value = TotalTempGrosssal
'----------------Lock Password-------------------START---
   'Sheet7.Protect Password:=sPassword
   Sheet39.Protect Password:=sPassword
   'Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub
Sub setTblinfo_OthersNOI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Salary.Nature1").count
    Set rangecells = Range("Salary.Nature1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OthersNOI_s = ccount
rngname_IncomeOth = "Salary.Nature1;Salary.Description1;Salary.Amount1;"
 
 End Sub
 
 Sub setTblinfo_OthersAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Salary.Amount1").count
    Set rangecells = Range("Salary.Amount1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_OthersAmt_s = ccount
End Sub
Function ValidateNatureOfIncome() As Boolean
ValidateNatureOfIncome = False

    setTblinfo_OthersNOI
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Range("Salary.Nature1").Cells
    Set rangecells1 = Range("Salary.Description1").Cells
    Dim i As Long
    ReDim Others_NOI_S(end_OthersNOI_s)
    ReDim Others_NOI1_s(end_OthersNOI_s)

    
    
     
    For i = 1 To end_OthersNOI_s
        Others_NOI_S(i) = rangecells.item(i).value
        Others_NOI1_s(i) = rangecells1.item(i).value
            
         If (Others_NOI_S(i) = "(Select)" Or Others_NOI_S(i) = "") Then
             MsgScheduleS1 = MsgScheduleS1 + "* Please enter Nature of Income at Sr. No  " & i & "in Salary Schedule" & Chr(13)
             ValidateNatureOfIncome = False
             Exit Function
         End If

         If (Others_NOI_S(i) = "Any Other") Then
         If Others_NOI1_s(i) = "" Then
'             MsgScheduleS1 = MsgScheduleS1 + "* Please enter Description is mandatory at Sr. No  " & i & " in Salary Schedule" & Chr(13)
             MsgScheduleS1 = MsgScheduleS1 + "* ""Please enter description in Nature of Exempt Allowance at Sr. No " & i & " in schedule salary""." & Chr(13)
             ValidateNatureOfIncome = False
             Exit Function
         End If
         End If
         
         
         If Len(Others_NOI1_s(i)) > 125 Then
             MsgScheduleS1 = MsgScheduleS1 + "* Description at Sr. No  " & i & "cannot exceed 125 characters in Salary Schedule" & Chr(13)
             ValidateNatureOfIncome = False
             Exit Function
         End If
    Next
End Function


Function ValidateAmount() As Boolean
ValidateAmount = False

    setTblinfo_OthersNOI
    Dim rangecells As Range
    Set rangecells = Range("Salary.Amount1").Cells
    Dim i As Long
    ReDim Others_Amt_s(end_OthersNOI_s)
    For i = 1 To end_OthersNOI_s
        Others_Amt_s(i) = rangecells.item(i).value
        
        If Not chkCompulsory(Others_Amt_s(i)) Then
            MsgScheduleS1 = MsgScheduleS1 + "* Please enter Amount at Sr. No  " & i & " in Salary Schedule" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
        
        
        If Not IsNumeric(Others_Amt_s(i)) Then
            MsgScheduleS1 = MsgScheduleS1 + "* Amount at Sr. No  " & i & "should be Numeric value in Salary Schedule" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
        
        If Others_Amt_s(i) > 99999999999999# Then
            MsgScheduleS1 = MsgScheduleS1 + "* Amount at Sr. No  " & i & "cannot exceed 14 digits in Salary Schedule" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
    Next
End Function

Sub DescriptionClear()

sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
Dim chcell As Variant
For Each chcell In Sheet39.Range("Salary.salaryDescription" & Sheet39.Range("SalarySectionCount").value)

                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""

Next
For Each chcell In Sheet39.Range("Salary.ValueDescription" & Sheet39.Range("SalarySectionCount").value)

                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""

            Next

            For Each chcell In Sheet39.Range("Salary.ProfitDescription" & Sheet39.Range("SalarySectionCount").value)

                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""

            Next


Sheet39.Protect Password:=sPassword
End Sub



Sub tttaa()
Application.EnableEvents = True
End Sub

'Malli_03/10/2025
Sub SalaryExemptAllowance1013A()
Application.EnableEvents = False
Sheet39.Unprotect Password:=getmsgstate

If Sheet1.Range("sheet1.Status").value = "H-HUF" Then
            Sheet39.Range("Sch10of13A_PlaceofWrk").ClearContents
            Sheet39.Range("Sch10of13A_PlaceofWrk").Locked = True
            Sheet39.Range("Sch10of13A_PlaceofWrk").Interior.Color = "&HD8D8D8"
            
            Sheet39.Range("Sch10of13A_ActlRentpaid").ClearContents
            Sheet39.Range("Sch10of13A_ActlRentpaid").Locked = True
            Sheet39.Range("Sch10of13A_ActlRentpaid").Interior.Color = "&HD8D8D8"
            
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").ClearContents
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").Locked = True
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").Interior.Color = "&HD8D8D8"
            
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").ClearContents
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").Locked = True
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").Interior.Color = "&HD8D8D8"

ElseIf SheetALL.Range("bacValue").value = 1 Then
            Sheet39.Range("Sch10of13A_PlaceofWrk").ClearContents
            Sheet39.Range("Sch10of13A_PlaceofWrk").Locked = True
            Sheet39.Range("Sch10of13A_PlaceofWrk").Interior.Color = "&HD8D8D8"
            
            Sheet39.Range("Sch10of13A_ActlRentpaid").ClearContents
            Sheet39.Range("Sch10of13A_ActlRentpaid").Locked = True
            Sheet39.Range("Sch10of13A_ActlRentpaid").Interior.Color = "&HD8D8D8"
            
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").ClearContents
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").Locked = True
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").Interior.Color = "&HD8D8D8"
            
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").ClearContents
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").Locked = True
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").Interior.Color = "&HD8D8D8"
            
ElseIf Sheet1.Range("sheet1.Status").value <> "H-HUF" And SheetALL.Range("bacValue").value = 2 Then

            Sheet39.Range("Sch10of13A_PlaceofWrk").ClearContents
            Sheet39.Range("Sch10of13A_PlaceofWrk").Locked = False
            Sheet39.Range("Sch10of13A_PlaceofWrk").Interior.Color = "&HCCFFCC"
            Sheet39.Range("Sch10of13A_PlaceofWrk").value = "(Select)"
            
            Sheet39.Range("Sch10of13A_ActlRentpaid").ClearContents
            Sheet39.Range("Sch10of13A_ActlRentpaid").Locked = False
            Sheet39.Range("Sch10of13A_ActlRentpaid").Interior.Color = "&HCCFFCC"
            
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").ClearContents
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").Locked = False
            Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").Interior.Color = "&HCCFFCC"
            
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").ClearContents
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").Locked = False
            Sheet39.Range("Sch10of13A_50Por40Pofsalary").Interior.Color = "&HCCFFCC"

End If
Application.EnableEvents = True
Sheet39.Protect Password:=getmsgstate
End Sub


