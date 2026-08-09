Attribute VB_Name = "mdTDS"
Option Explicit

Public TAN_TDS1 As Variant
Public EmployerOrDeductorOrCollecterName_TDS1 As Variant
Public AmtClaimedOnOwnHands_TDS1 As Variant
Public AmtCarriedFwd_TDS1 As Variant

Dim rngname_TDS11 As Variant
Dim end_TDS11 As Variant
Dim rngname_TDS1 As Variant
Public end_TDS1 As Variant

Dim rngname_TDS1_8 As Variant
Dim end_TDS1_8 As Variant

Public MsgSchIT, MsgTCS, MsgTDS1 As Variant

Dim i As Long

' IT


Public NameOfBank_IT As Variant
Public NameOfBranch_IT As Variant
Public BSRCode_IT As Variant
Public DateDep_IT As Variant
Public SrlNoOfChaln_IT As Variant
Public Amt_IT As Variant

Dim rngname_IT2, rngname_IT22 As Variant
Dim end_IT2 As Variant

Dim rngname_IT As Variant
Public end_IT As Variant
Dim incBy_IT As Variant
'----TCS----
Public TAN_TCS As Variant
Public EmployerOrDeductorOrCollecterName_TCS As Variant
Public AmtPaid_TCS As Variant
Public DeductedYear_TCS As Variant
Public SectionTCSDeducted_TCS As Variant ''TCS_C1 AY_2024_25 change


Public BroughtFwdTCSAmt_TCS As Variant
Public TotTCSOnAmtPaidCY_TCS As Variant
Public AmtClaimedOnOwnHands_TCS As Variant
Public AmtCarriedFwd_TCS As Variant
Public AmtClaimedBySpouse_TCS As Variant

Dim rngname_TCS1 As Variant
Dim end_TCS1 As Variant
Dim rngname_TCS As Variant
Public end_TCS As Variant
Dim incBy_TCS As Variant

Dim rngname_TCS_7 As Variant
Dim end_TCS_7 As Variant


Public end_Name_TDS1, end_Inc_TDS1, end_TotTAx_TDS1 As Long
Public end_Name_TDS2, end_Year_TDS2, end_AmtPaid_TDS2, end_TDi_TDS2, end_Amt_TDS2, end_TDii_TDS2, end_TDiii_TDS2, end_TD_PAN_TDS2, end_TD_PANp_TDS2, end_TDS_SIncome_TDS2, end_TDS_TDS_TDS2, end_TDS_PAN_TDS2, end_TDS_PANp_TDS2 As Long
Public end_Name_TDS2ii, end_Year_TDS2ii, end_AmtPaid_TDS2ii, end_TDi_TDS2ii, end_Amt_TDS2ii, end_TDii_TDS2ii, end_TDiii_TDS2ii, end_TD_PAN_TDS2ii, end_TDS_SIncome_TDS2ii, end_TDS_TDS_TDS2ii, end_TDS_PAN_TDS2ii As Long
Public end_Date_IT, end_SerialCh_IT, end_Taxpd_IT As Long
Public end_CollectorName_TCS, end_SectionTCSDeducted_TCS, end_AmtPaid_TCS, end_TaxCollected_TCS, end_AmtClaim_TCS, end_AmtPortuguese_TCS As Long


'----TDS2----
Public TAN_TDS2 As Variant
Public EmployerOrDeductorOrCollecterName_TDS2 As Variant
Public UTN_TDS2 As Variant
Public DeductedYear_TDS2 As Variant
Public AmtPaid_TDS2 As Variant
Public AmtClaimedOnOwnHands_TDS2 As Variant
Public TD_SIncome_TDS2 As Variant
Public TD_OwnHands_TDS2 As Variant
Public TD_TDS_TDS2 As Variant
Public TDS_SIncome_TDS2 As Variant
Public TDS_TDS_TDS2 As Variant
Public PAN7_TDS As Variant
Public PAN8_TDS As Variant
Public PAN9other_TDS As Variant
Public PAN10other_TDS As Variant
Dim TotalSum As Variant
Dim TotalSum1 As Variant

Dim rngname_TDS22 As Variant
Dim end_TDS22 As Variant
Dim rngname_TDS2 As Variant
Dim rngname_TDS2_8 As Variant
Public end_TDS2 As Variant

Public end_TDS2_1, end_TDS2_2, end_TDS2_3, end_TDS2_4, end_TDS2_5, end_TDS2_6, end_TDS2_7, end_TDS2_8, end_TDS2_9, end_TDS2_10 As Variant

Public TAN_TDS2_1, SectionTDSDeducted_TDS2_1, TdsClaimed_TDS2_1, Tds_TDS2_1, FinYear_TDS2_1, Tdsdeducted_TDS2_1, SectionTDS_TDS2_1, GrossAmount_TDS2_1, Headincome_TDS2_1, Tdscarried_TDS2_1 As Variant

Public end_TDS2ii_1, end_TDS2ii_2, end_TDS2ii_3, end_TDS2ii_4, end_TDS2ii_5, end_TDS2ii_6, end_TDS2ii_7, end_TDS2ii_8, end_TDS2ii_9 As Variant

Public PAN_TDS2ii_1, SectionTDSDeducted_TDS2ii_1, Aadhaar_TDS2ii_1, TdsClaimed_TDS2ii_1, Tds_TDS2ii_1, FinYear_TDS2ii_1, Tdsdeducted_TDS2ii_1, SectionTDS_TDS2_2ii, GrossAmount_TDS2ii_1, Headincome_TDS2ii_1, Tdscarried_TDS2ii_1 As Variant


Dim end_TDSCredit_TDS2, end_TDSCredit_TDS3, end_Yeari_TDS2, end_Yearii_TDS2ii, end_TDSCredit2_TDS2, end_TDSCredit3_TDS3 As Variant

Public MsgTDS2, MsgTDS2i As Variant



'''''' TDS2ii''''''
Public PAN_TDS2ii As Variant
Public TDS2Credit, TDS3Credit As Variant
Public EmployerOrDeductorOrCollecterName_TDS2ii As Variant
Public UTN_TDS2ii As Variant
Public DeductedYear_TDS2ii As Variant
Public AmtPaid_TDS2ii As Variant
Public AmtClaimedOnOwnHands_TDS2ii As Variant
Public TD_SIncome_TDS2ii As Variant
Public TD_OwnHands_TDS2ii As Variant
Public TD_TDS_TDS2ii As Variant
Public TDS_SIncome_TDS2ii As Variant
Public TDS_TDS_TDS2ii As Variant
Public PAN7_TDSii As Variant
Public PAN8_TDSii As Variant
Public PAN9other_TDSii As Variant
Public PAN10other_TDSii As Variant
Dim TotalSum2ii As Variant
Dim TotalSum2iii


Dim rngname_TDS22ii As Variant
Dim end_TDS22ii As Variant
Dim rngname_TDS2ii As Variant
Dim rngname_TDS2_8ii As Variant
Public end_TDS2ii As Variant


Dim end_TDS2_8i As Variant

Public MsgTDS2ii As Variant


Sub Cmd_Validate_Click()
    ValidateTDS_TCS_IT
    fmsgboxStatus "Sheet TDS is Ok" ', vbOKOnly, vbMessageCaption
End Sub
Sub Cmd_ValidateIT_Click()
    Validate_IT
    fmsgboxStatus "Sheet IT is Ok" ', vbOKOnly, vbMessageCaption
End Sub
Sub Cmd_ValidateTCS_Click()
    Validate_TCS
    fmsgboxStatus "Sheet TCS is Ok"  ', vbOKOnly, vbMessageCaption
End Sub
Sub Cmd_Prev_Click_TDS()
Sheet3.Activate
End Sub

Sub Cmd_Next_Click_TDS()
Sheet18.Activate
End Sub

Sub CmdTDS1()
    Sheets("TDS").Activate
    mdCommon.DefinedgridNameRange = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSal||TDSal.TotalTDSSal"
    ActiveCellRange = mdCommon.searchLastRow("TDSal.TAN")
    mdCommon.insertRowUnderSectionWithFormula
End Sub
Sub CmdTDS2()
    Sheets("TDS").Activate
    mdCommon.DefinedgridNameRange = "TDsOthr.TAN||TDsOthr.SectionTDS||TDsOthr.finyear||TDsOthr.tds||TDsOthr.tdsdeducted||TDsOthr.tdsclaimed||TDsOthr.grossamount||TDsOthr.headincome||TDsOthr.tdscarried"
    ActiveCellRange = mdCommon.searchLastRow("TDsOthr.TAN")
    mdCommon.insertRowUnderSectionWithFormulaTDS HOIflag:=1
    
'    setTableinfo_TDSCredit_TDS2
'    Sheet4.Range("TDS2.TdsCredit").Cells.item(end_TDSCredit_TDS2 + 1).value = "SELF"
'    setTableinfo_Year_TDS1i
'    Sheet4.Range("TDSoth.DeductedYear").Cells.item(end_Yeari_TDS2 + 1).value = "(Select)"
End Sub
Sub CmdTDS2ii()
    Sheets("TDS").Activate
    mdCommon.DefinedgridNameRange = "TDsOthr2.PAN||TDsOthr2.AADHAAR||TDsOthr2.SectionTDSDeducted||TDsOthr2.finyear||TDsOthr2.tds||TDsOthr2.tdsdeducted||TDsOthr2.tdsclaimed||TDsOthr2.grossamount||TDsOthr2.headamount||TDsOthr2.carried"
    ActiveCellRange = mdCommon.searchLastRow("TDsOthr2.PAN")
    mdCommon.insertRowUnderSectionWithFormulaTDS3 HOIflag:=1
'    setTableinfo_TDSCredit_TDS3
'    Sheet4.Range("TDS3.TdsCredit").Cells.item(end_TDSCredit_TDS3 + 1).value = "SELF"
'    setTableinfo_Yearii_TDS2ii
'    Sheet4.Range("TDSoth.DeductedYear").Cells.item(end_Yearii_TDS2ii + 1).value = "(Select)"
End Sub
Sub CmdTCS()
    Sheets("TCS").Activate
    mdCommon.DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.AmountPaid||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
    ActiveCellRange = mdCommon.searchLastRow("TCS.TAN")
    mdCommon.insertRowUnderSectionWithFormula
End Sub
'-------------------------------------------
'Sub CmdTDS2()
'    Sheets("TDS").Activate
'    ''TDS_E1
'    ''mdCommon.DefinedgridNameRange = "TDsOthr.TAN||TDsOthr.finyear||TDsOthr.tds||TDsOthr.tdsdeducted||TDsOthr.tdsclaimed||TDsOthr.grossamount||TDsOthr.headincome||TDsOthr.tdscarried"
'    ''TDS_C1 AY_2024_25 change
'    mdCommon.DefinedgridNameRange = "TDsOthr.TAN||TDsOthr.SectionTDSDeducted||TDsOthr.finyear||TDsOthr.tds||TDsOthr.tdsdeducted||TDsOthr.tdsclaimed||TDsOthr.grossamount||TDsOthr.headincome||TDsOthr.tdscarried"
'    ActiveCellRange = mdCommon.searchLastRow("TDsOthr.TAN")
'    mdCommon.insertRowUnderSectionWithFormulaTDS HOIflag:=1
'
''    setTableinfo_TDSCredit_TDS2
''    Sheet4.Range("TDS2.TdsCredit").Cells.item(end_TDSCredit_TDS2 + 1).value = "SELF"
''    setTableinfo_Year_TDS1i
''    Sheet4.Range("TDSoth.DeductedYear").Cells.item(end_Yeari_TDS2 + 1).value = "(Select)"
'End Sub
'Sub CmdTDS2ii()
'    Sheets("TDS").Activate
'    ''TDS_E4
'    ''mdCommon.DefinedgridNameRange = "TDsOthr2.PAN||TDsOthr2.AADHAAR||TDsOthr2.finyear||TDsOthr2.tds||TDsOthr2.tdsdeducted||TDsOthr2.tdsclaimed||TDsOthr2.grossamount||TDsOthr2.headamount||TDsOthr2.carried"
'    ''TDS_C4 AY_2024_25 change
'    mdCommon.DefinedgridNameRange = "TDsOthr2.PAN||TDsOthr2.SectionTDSDeducted||TDsOthr2.AADHAAR||TDsOthr2.finyear||TDsOthr2.tds||TDsOthr2.tdsdeducted||TDsOthr2.tdsclaimed||TDsOthr2.grossamount||TDsOthr2.headamount||TDsOthr2.carried"
'    ActiveCellRange = mdCommon.searchLastRow("TDsOthr2.PAN")
'    mdCommon.insertRowUnderSectionWithFormulaTDS
''    setTableinfo_TDSCredit_TDS3
''    Sheet4.Range("TDS3.TdsCredit").Cells.item(end_TDSCredit_TDS3 + 1).value = "SELF"
''    setTableinfo_Yearii_TDS2ii
''    Sheet4.Range("TDSoth.DeductedYear").Cells.item(end_Yearii_TDS2ii + 1).value = "(Select)"
'End Sub


'Sub CmdTCS()
'    Sheets("TCS").Activate
'    ''TCS_E1
'    '''mdCommon.DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.AmountPaid||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
'
'    ''TCS_C1 AY_2024_25 change
'    mdCommon.DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.SectionTCSDeducted||TCS.AmountPaid||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
'    ActiveCellRange = mdCommon.searchLastRow("TCS.TAN")
'    mdCommon.insertRowUnderSectionWithFormula
'End Sub

Sub CmdIT()
    Sheets("IT").Activate
    '----------02/12/2025
    'mdCommon.DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1"
    mdCommon.DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT"
    '-------------------------
    ActiveCellRange = mdCommon.searchLastRow("TaxP.BSRCode")
    mdCommon.insertRowUnderSectionWithFormula
End Sub


Sub ValidateTDS_TCS_IT()
On Error Resume Next
    MsgTDS1 = ""
    If Not ValidatesheetTDS1 Then
        Sheet4.Activate
        fmsgboxsmall (MsgTDS1)
        CloseMsg
    End If
    
    MsgTDS2 = ""
    If Not ValidateSheetTDS2 Then
        Sheet4.Activate
        fmsgboxsmall (MsgTDS2) ' , vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    MsgTDS2ii = ""
    If Not ValidateSheetTDS2ii Then
        Sheet4.Activate
        fmsgboxsmall (MsgTDS2ii) ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
        
End Sub
Function Validate_IT()
       MsgSchIT = ""
        If Not ValidatesheetIT Then
         Sheet18.Activate
         fmsgboxsmall (MsgSchIT) ', vbOKOnly, "Error(s)!"
         CloseMsg
    End If
End Function
Function Validate_TCS()

  MsgTCS = ""
    If Not ValidatesheetTCS Then
        Sheet17.Activate
        fmsgboxsmall (MsgTCS) ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
End Function



'Function ValidatesheetTCS() As Boolean
'On Error Resume Next
'Dim rangename, intcntr, name, cellcount As Variant
'Dim k  As Long
'
'    ValidatesheetTCS = True
'    MsgTCS = "TCS." & Chr(13)
'
'    setTblinfo_TCS
'    setTableInfo_CollectorName_TCS
'
'    setTableInfo_SectionTCSDeducted_TCS  '''''''''''''''Added by Malli ''TCS_C1 AY_2024_25 change
'
'    setTableInfo_AmtPaid_TCS
'    setTableInfo_TaxCollect_TCS
'    setTableInfo_AmtClain_TCS
'
'    end_TCS = WorksheetFunction.Max(0, end_TCS, end_CollectorName_TCS, end_SectionTCSDeducted_TCS, end_AmtPaid_TCS, end_TaxCollected_TCS, end_AmtClaim_TCS)
'
'For k = 1 To end_TCS
'    If Not (Sheet17.Range("TCS.TAN").Cells.item(k) <> "" And Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" And Sheet17.Range("TCS.SectionTCSDeducted").Cells.item(k) <> "" And Sheet17.Range("TCS.AmountPaid").Cells.item(k) <> "" And Sheet17.Range("TCS.TotalTCS").Cells.item(k) <> "" And Sheet17.Range("TCS.AmtTCSClaimedThisYear").Cells.item(k) <> "") Then
'    'If Not (Sheet17.Range("TCS.TAN").Cells.item(k) <> "" Or Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Sheet17.Range("TCS.AmountPaid").Cells.item(k) <> "" Or Sheet17.Range("TCS.TotalTCS").Cells.item(k) <> "" Or Sheet17.Range("TCS.AmtTCSClaimedThisYear").Cells.item(k) <> "") Then
'        MsgTCS = MsgTCS & "* At Sr. No " & k & "* Please fill all the Mandatory Fields " & Chr(13)
'        ValidatesheetTCS = False
'        Exit Function
'    End If
'Next
'
'    If end_TCS > 0 Then GoTo xyz4
'    cellcount = Sheet17.Range("TCS.TAN").count
'
'    rangename = Split(rangename, "|")
'
'
'xyz4:
'    If Not ValidateTAN_TCS() Then ValidatesheetTCS = False
'    If Not ValidateEmployerOrDeductorOrCollecterName_TCS() Then ValidatesheetTCS = False
'    If Not ValidateSectionTCSDeducted_TCS() Then ValidatesheetTCS = False  '''''''Added by malli ''TCS_C1 AY_2024_25 change
'    If Not ValidateAmtPaid_TCS() Then ValidatesheetTCS = False
'    If Not ValidateClaimOutOfTotTCSOnAmtPaid_TCS() Then ValidatesheetTCS = False
'
'        setTblinfo_TCS
'    setTableInfo_CollectorName_TCS
'    setTableInfo_SectionTCSDeducted_TCS  ''''''''added by Malli '''TCS_C1 AY_2024_25 change
'    setTableInfo_AmtPaid_TCS
'    setTableInfo_TaxCollect_TCS
'    setTableInfo_AmtClain_TCS
''        setTableInfo_AmtPortugese_TCS
'
'
'    If Len(Sheet17.Range("TCS.Sum").value) > 14 Then
'        MsgTCS = MsgTCS & "* Toatl of TCS Should not be exceed 14 digits " & Chr(13)
'        ValidatesheetTCS = False
'    End If
'
'
'End Function
Function ValidatesheetTCS() As Boolean
On Error Resume Next
Dim rangename, intcntr, name, cellcount As Variant
Dim k  As Long

    ValidatesheetTCS = True
    MsgTCS = "TCS." & Chr(13)
    
    setTblinfo_TCS
    setTableInfo_CollectorName_TCS
    setTableInfo_AmtPaid_TCS
    setTableInfo_TaxCollect_TCS
    setTableInfo_AmtClain_TCS
    
    end_TCS = WorksheetFunction.Max(0, end_TCS, end_CollectorName_TCS, end_AmtPaid_TCS, end_TaxCollected_TCS, end_AmtClaim_TCS)

'SIT-56341 Starts-------------

'For k = 1 To end_TCS
'    If Not (Sheet17.Range("TCS.TAN").Cells.item(k) <> "" And Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" And Sheet17.Range("TCS.AmountPaid").Cells.item(k) <> "" And Sheet17.Range("TCS.TotalTCS").Cells.item(k) <> "" And Sheet17.Range("TCS.AmtTCSClaimedThisYear").Cells.item(k) <> "") Then
'    'If Not (Sheet17.Range("TCS.TAN").Cells.item(k) <> "" Or Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Sheet17.Range("TCS.AmountPaid").Cells.item(k) <> "" Or Sheet17.Range("TCS.TotalTCS").Cells.item(k) <> "" Or Sheet17.Range("TCS.AmtTCSClaimedThisYear").Cells.item(k) <> "") Then
'        MsgTCS = MsgTCS & "* At Sr. No " & k & "* Please fill all the Mandatory Fields " & Chr(13)
'        ValidatesheetTCS = False
'        Exit Function
'    End If
'Next

'SIT-56341 Ends-------------
   
    If end_TCS > 0 Then GoTo xyz4
    cellcount = Sheet17.Range("TCS.TAN").count

    rangename = Split(rangename, "|")


xyz4:
    If Not ValidateTAN_TCS() Then ValidatesheetTCS = False
    If Not ValidateEmployerOrDeductorOrCollecterName_TCS() Then ValidatesheetTCS = False
    If Not ValidateAmtPaid_TCS() Then ValidatesheetTCS = False
    If Not ValidateClaimOutOfTotTCSOnAmtPaid_TCS() Then ValidatesheetTCS = False

        setTblinfo_TCS
    setTableInfo_CollectorName_TCS
    setTableInfo_AmtPaid_TCS
    setTableInfo_TaxCollect_TCS
    setTableInfo_AmtClain_TCS
'        setTableInfo_AmtPortugese_TCS
 

    If Len(Sheet17.Range("TCS.Sum").value) > 14 Then
        MsgTCS = MsgTCS & "* Toatl of TCS Should not be exceed 14 digits " & Chr(13)
        ValidatesheetTCS = False
    End If


End Function
Function ValidateTAN_TCS() As Boolean
    ValidateTAN_TCS = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("TCS.TAN").Cells
    ReDim TAN_TCS(end_TCS)
       For i = 1 To end_TCS
        TAN_TCS(i) = rangecells.item(i).value
            
            If Not chkCompulsory(TAN_TCS(i)) Then
                MsgTCS = MsgTCS + "* Please enter the TAN of Collector at Sr. No " & i & " in Sheet TCS" & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If

            
            If Not ValidateTantype_text(Mid(TAN_TCS(i), 1, 4)) Then
                MsgTCS = MsgTCS + "* TAN at Sr. No " & i & " in Sheet TCS is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            If Not IsNumeric(Mid(TAN_TCS(i), 5, 5)) Then
                MsgTCS = MsgTCS + "* TAN at Sr. No " & i & "  in Sheet TCS is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            If Not ValidateTantype_text(Right(TAN_TCS(i), 1)) Then
                MsgTCS = MsgTCS + "* TAN at Sr. No " & i & "  in Sheet TCS is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            If Not ValidateTANCodes(UCase(TAN_TCS(i))) Then
                MsgTCS = MsgTCS + "* Invalid TAN at Sr.No " & i & " in Sheet TCS(i).Please enter valid TAN" & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            

    Next
End Function

Function ValidatesheetTDS1() As Boolean
On Error Resume Next
    Dim cellcountTDS1, intcntr, count, rangename, name, errmsgtemp As Variant
    ValidatesheetTDS1 = True
    MsgTDS1 = "TDS 1." & Chr(13)
    subProcCaption = "Validating TDS1"
    setTblinfo_TDS1
    noOfProcessSub = end_TDS1
    Dim k  As Long
    For k = 1 To end_TDS1
    If Not (Sheet4.Range("TDSal.TAN").Cells.item(k) <> "" Or Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Sheet4.Range("TDSal.IncChrgSal").Cells.item(k) <> "" Or Sheet4.Range("TDSal.TotalTDSSal").Cells.item(k) <> "") Then
        MsgTDS1 = MsgTDS1 & "* At Sr. No " & k & " Please fill all the Mandatory Fields in Sheet TDS(1)" & Chr(13)
'         MsgTDS1 = MsgTDS1 & "* At Sr. No " & k & " Please enter the TAN of Employer" & Chr(13)
        ValidatesheetTDS1 = False
        Exit Function
    End If
    UpdateProgressBar
    Next
    
    If end_TDS1 > 0 Then GoTo xyz1
    
   cellcountTDS1 = Sheet4.Range("TDSal.TAN").count
    rangename = "TDSal.TAN|TDSal.EmployerOrDeductorOrCollecterName|TDSal.IncChrgSal|TDSal.TotalTDSSal"
   rangename = Split(rangename, "|")
    For intcntr = 1 To cellcountTDS1
    If Sheet4.Range("TDSal.TAN").Cells.item(intcntr) <> "" Or Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Cells.item(intcntr) <> "" Or Sheet4.Range("TDSal.IncChrgSal").Cells.item(intcntr) <> "" Or Sheet4.Range("TDSal.TotalTDSSal").Cells.item(intcntr) <> "" Then
    For Each name In rangename
        If Sheet4.Range(name).item(intcntr) = "" Then ValidatesheetTDS1 = False
     Next
     If ValidatesheetTDS1 = False Then
'        MsgTDS1 = MsgTDS1 & "* At Sr. No " & intcntr & " Please fill all the Mandatory Fields in Sheet TDS(1)" & Chr(13)
         MsgTDS1 = MsgTDS1 & "* At Sr. No " & intcntr & " Please enter the TAN of Employer in Sheet TDS(1)" & Chr(13)
        Exit Function
    End If
    End If
   Next
    
xyz1:
    

    
    
    If Not ValidateTAN_TDS1() Then ValidatesheetTDS1 = False
    If Not ValidateEmployerOrDeductorOrCollecterName_TDS1() Then ValidatesheetTDS1 = False
    If Not ValidateClaimOutOfTotTDSOnAmtPaid_TDS1() Then ValidatesheetTDS1 = False
    
    setTableinfo_NameTDS
    setTableinfo_IncTDS1
    setTableinfo_Tot_TaxTDS1
    
    
    
'    If ((end_TDS1 <> end_Name_TDS1) Or (end_TDS1 <> end_Inc_TDS1) Or (end_TDS1 <> end_TotTAx_TDS1)) Then
'        MsgTDS1 = MsgTDS1 & "* Please fill all the Mandatory Fields in Sheet TDS(1)" & Chr(13)
'        ValidatesheetTDS1 = False
'    End If
    
    If Len(Sheet4.Range("TDSal.Sum").value) > 14 Then
    MsgTDS1 = MsgTDS1 & "* Total of TDS1 Should not be exceed 14 digits in Sheet TDS(1)" & Chr(13)
    End If
    'Change.03.03.2023.102.IDS.97
    'Ayush_18/08
    'grossSalary_exempt_familyPension
'    If Sheet1.Range("IncD.IncomeFromSal_i").value < Sheet4.Range("TotIncSal_TDSHidden").value Then
'     If Sheet5.Range("grosssalandexemptincome_hidden").value < Sheet4.Range("TotIncSal_TDSHidden").value Then
    'Ayush_26-27_TDS_C8_commented
'    If Sheet5.Range("grossSalary_exempt_familyPension").value < Sheet4.Range("TotIncSal_TDSHidden").value Then
''        MsgTDS1 = MsgTDS1 & "* Amount of Gross Salary declared is less than the amount of Gross Salary declared in Schedule TDS 1" & Chr(13)
''        MsgTDS1 = MsgTDS1 & "*  ""Amount of ""Gross Salary"" and ""exempt incomes"" which are part of salary declared is less than the amount of Gross Salary declared in Schedule TDS 1""" & Chr(13)
''        ValidatesheetTDS1 = False
'
'        'Changed by Ankita on 29/10/2025_IPIP-74303
''    MsgBox "* ""Amount of ""Gross Salary"" and ""exempt incomes"" which are part of salary declared is less than the amount of Gross Salary declared in Schedule TDS 1""", , "Alert:"
'    MsgBox "* ""Sum of ""Gross Salary"" , ""Family Pension"" and ""exempt incomes which are part of salary"" declared is less than the amount of Gross Salary declared in Schedule TDS 1""", , "Alert:"
'
'    End If
    'End Change IDS.97
End Function

Function ValidateEmployerOrDeductorOrCollecterName_TCS() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TCS = True
'    setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
    ReDim EmployerOrDeductorOrCollecterName_TCS(end_TCS)
    For i = 1 To end_TCS
        EmployerOrDeductorOrCollecterName_TCS(i) = rangecells.item(i).value
        If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TCS(i)) Then
             MsgTCS = MsgTCS + "* Please enter the Name of Collector at Sr. No " & i & "  in Sheet TCS" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TCS = False
            Exit Function
        End If


         If Len(EmployerOrDeductorOrCollecterName_TCS(i)) > 125 Then
            MsgTCS = MsgTCS + "* Name Of Collector as Sr. No " & i & " in Sheet TCS cannot be more than 125 charachters." & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TCS = False
            Exit Function
         End If
 Next
End Function

Function ValidateAmtPaid_TCS() As Boolean
 
    ValidateAmtPaid_TCS = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("TCS.AmountPaid").Cells
    ReDim AmtPaid_TCS(end_TCS)
    For i = 1 To end_TCS
        AmtPaid_TCS(i) = rangecells.item(i).value
        If Not chkCompulsory(AmtPaid_TCS(i)) Then
             MsgTCS = MsgTCS + "* Please enter the amount at Sr. No " & i & "  in Sheet TCS" & Chr(13)
            ValidateAmtPaid_TCS = False
            Exit Function
        End If


         If Len(AmtPaid_TCS(i)) > 14 Then
            MsgTCS = MsgTCS + "* Amount at Sr. No " & i & " in Sheet TCS  cannot exceed 99999999999999" & Chr(13)
            ValidateAmtPaid_TCS = False
            Exit Function
         End If
 Next
End Function
Function ValidateClaimOutOfTotTCSOnAmtPaid_TCS() As Boolean
    ValidateClaimOutOfTotTCSOnAmtPaid_TCS = True
    'setTblinfo_TCS
    
    Dim rangecells1 As Range
'    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    
    Dim i As Long
    
    Set rangecells1 = Sheet17.Range("TCS.TotalTCS").Cells
    Set rangecells3 = Sheet17.Range("TCS.AmtTCSClaimedThisYear").Cells
    'Set rangecells2 = Sheet17.Range("TCS.AmtClaimedBySpouse").Cells
    Set rangecells4 = Sheet17.Range("TCS.AmountPaid").Cells
    
    
    ReDim BroughtFwdTCSAmt_TCS(end_TCS)
    ReDim AmtClaimedOnOwnHands_TCS(end_TCS)
'    ReDim AmtClaimedBySpouse_TCS(end_TCS)
    ReDim AmtPaid_TCS(end_TCS)
    
    For i = 1 To end_TCS
      BroughtFwdTCSAmt_TCS(i) = rangecells1.item(i).value
      AmtClaimedOnOwnHands_TCS(i) = rangecells3.item(i).value
      'AmtClaimedBySpouse_TCS(i) = rangecells2.item(i).value
      AmtPaid_TCS(i) = rangecells4.item(i).value
      
    If (BroughtFwdTCSAmt_TCS(i) = "") Then
        MsgTCS = MsgTCS + "* Please enter the Total Tax Collected at Sr. No " & i & " in Sheet TCS" & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If
    
    
    If IsNumeric(BroughtFwdTCSAmt_TCS(i)) Then
       If BroughtFwdTCSAmt_TCS(i) < 0 Then
            MsgTCS = MsgTCS + "* Amount in Sheet TCS of col.no 4 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        End If
        If Len("" & BroughtFwdTCSAmt_TCS(i)) > 14 Then
            MsgTCS = MsgTCS + "* Amount in Sheet TCS of col.no 4 cannot exceed 99999999999999 in Row " & i & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        End If
    Else
        MsgTCS = MsgTCS + "* Amount in Sheet TCS of col.no 4 is invalid in Row " & i & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If
    
    
    If (AmtClaimedOnOwnHands_TCS(i) = "") Then
    'Ayush
       ' MsgTCS = MsgTCS + "* Amount claimed for this year cannot be more than total tax collected at Sr. No " & i & " in Sheet TCS" & Chr(13)
       MsgTCS = MsgTCS + "* ""Please enter Amount claimed for this year"" at Sr. No " & i & " in Sheet TCS" & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If

    
    If IsNumeric(AmtClaimedOnOwnHands_TCS(i)) Then
      If AmtClaimedOnOwnHands_TCS(i) < 0 Then
            MsgTCS = MsgTCS + "* Amount in Sheet TCS of col.no 5 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
      End If
     If Len("" & AmtClaimedOnOwnHands_TCS(i)) > 14 Then
            MsgTCS = MsgTCS + "* Amount in Sheet TCS of col.no 5 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
            Exit Function
     End If
Else
        MsgTCS = MsgTCS + "* Amount in Sheet TCS of col.no 5 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        Exit Function
End If
    
        
   
    If AmtClaimedOnOwnHands_TCS(i) > (BroughtFwdTCSAmt_TCS(i)) Then
'        MsgTCS = MsgTCS + "* ""Amount claimed for this year cannot be more than total tax collected"" in Col.No 4 in Sheet TCS" & Chr(13)
        'Changed as per Bhoomija from BA Team
        MsgTCS = MsgTCS + "Amount claimed for this year cannot be more than total tax collected"
        
        'MsgTCS = MsgTCS + "* Amount claimed in Col 5 cannot exceed Total tax Collected in Row in Sheet TCS " & i & Chr(13)

        
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        Exit Function
    End If


    If (AmtClaimedOnOwnHands_TCS(i) = "") Then
        AmtClaimedOnOwnHands_TCS(i) = 0
    End If


 Next
End Function

Sub setTblinfo_TCS()
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
 rngname_TCS = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.AmountPaid||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
 End Sub

Sub setTblinfo_TCS_7()
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
 end_TCS_7 = ccount
 rngname_TCS_7 = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.AmountPaid||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
 End Sub


Sub setTblinfo_TDS1()
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
 end_TDS1 = ccount
 rngname_TDS1 = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSal||TDSal.TotalTDSSal"
 End Sub


Sub setTblinfo_TDS1_8()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSal.TotalTDSSal").count
 Set rangecells = Sheet4.Range("TDSal.TotalTDSSal").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS1_8 = ccount
 rngname_TDS1_8 = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSal||TDSal.TotalTDSSal"
 End Sub
 
Function ValidateTAN_TDS1() As Boolean
    ValidateTAN_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet4.Range("TDSal.TAN").Cells
    ReDim TAN_TDS1(end_TDS1)
       For i = 1 To end_TDS1
        TAN_TDS1(i) = UCase(rangecells.item(i).value)
            
            If (TAN_TDS1(i) = "") Then
                MsgTDS1 = MsgTDS1 + "* Please enter the TAN of Employer at Sr. No " & i & " in Schedule TDS(1)" & Chr(13)
                ValidateTAN_TDS1 = False
                Exit Function
            End If
            
            
            If (Len(TAN_TDS1(i)) > 10) Then
'                MsgTDS1 = MsgTDS1 + "* TAN at Sr. No " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
                'SIT-56342
                MsgTDS1 = MsgTDS1 + "* Invalid TAN at Sr. No " & i & " . TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
                ValidateTAN_TDS1 = False
                Exit Function
            End If
            
            If Not ValidateTantype_text(Mid(TAN_TDS1(i), 1, 4)) Then
'                MsgTDS1 = MsgTDS1 + "* TAN at Sr. No " & i & "is invalid. First 4 alphabets, next 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
            'SIT-56342
                 MsgTDS1 = MsgTDS1 + "* Invalid TAN at Sr. No " & i & " . TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
                ValidateTAN_TDS1 = False
                Exit Function
            End If
            If Not IsNumeric(Mid(TAN_TDS1(i), 5, 5)) Then
'               MsgTDS1 = MsgTDS1 + "* TAN at Sr. No " & i & "is invalid. First 4 alphabets, next 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
             'SIT-56342
                MsgTDS1 = MsgTDS1 + "* Invalid TAN at Sr. No " & i & " . TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
                ValidateTAN_TDS1 = False
                Exit Function
            End If
            If Not ValidateTantype_text(Right(TAN_TDS1(i), 1)) Then
'                MsgTDS1 = MsgTDS1 + "* TAN at Sr. No " & i & "is invalid. First 4 alphabets, next 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)
               'SIT-56342
                 MsgTDS1 = MsgTDS1 + "* Invalid TAN at Sr. No " & i & " . TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS(1)" & Chr(13)

                ValidateTAN_TDS1 = False
                Exit Function
            End If
             If Not ValidateTANCodes(UCase(TAN_TDS1(i))) Then
                MsgTDS1 = MsgTDS1 + "* Invalid TAN at Sr.No " & i & " in Schedule TDS(1).Please enter valid TAN" & Chr(13)
                ValidateTAN_TDS1 = False
                Exit Function
            End If
    Next
End Function




Function ValidateEmployerOrDeductorOrCollecterName_TDS1() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Cells
    ReDim EmployerOrDeductorOrCollecterName_TDS1(end_TDS1)
    For i = 1 To end_TDS1
    
        EmployerOrDeductorOrCollecterName_TDS1(i) = rangecells.item(i).value
        
        If EmployerOrDeductorOrCollecterName_TDS1(i) = "" Then
            MsgTDS1 = MsgTDS1 + "* Please enter the Name of Employer at Sr. No " & i & " in Schedule TDS(1)" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TDS1 = False
            Exit Function
        End If
        
        
        If Len(EmployerOrDeductorOrCollecterName_TDS1(i)) > 125 Then
            MsgTDS1 = MsgTDS1 + "* Name of Employer at Sr. No " & i & "Cannot be more than 125 Charachters in Schedule TDS(1)" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TDS1 = False
            Exit Function
        End If

    Next
End Function


Function ValidateClaimOutOfTotTDSOnAmtPaid_TDS1() As Boolean
    ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim i As Long
    
    Set rangecells1 = Sheet4.Range("TDSal.IncChrgSal").Cells
    Set rangecells2 = Sheet4.Range("TDSal.TotalTDSSal").Cells
        
    ReDim AmtClaimedOnOwnHands_TDS1(end_TDS1)
    ReDim AmtCarriedFwd_TDS1(end_TDS1)
    
    For i = 1 To end_TDS1

        
        AmtClaimedOnOwnHands_TDS1(i) = rangecells1.item(i).value
        AmtCarriedFwd_TDS1(i) = rangecells2.item(i).value

        
        If AmtClaimedOnOwnHands_TDS1(i) = "" Then
            MsgTDS1 = MsgTDS1 + "* Please enter the Income under Salaries at Sr.No " & i & " in Schedule TDS(1)" & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
        End If
        
        If AmtCarriedFwd_TDS1(i) = "" Then
            MsgTDS1 = MsgTDS1 + "* Please enter the Total Tax Deducted at Sr.No " & i & " in Schedule TDS(1)" & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
        End If
        
        
               
        If IsNumeric(AmtClaimedOnOwnHands_TDS1(i)) Then
            If AmtClaimedOnOwnHands_TDS1(i) < 0 Then
                MsgTDS1 = MsgTDS1 + "* Amount in col.no 3 should be Numeric, Non Negative, not exceeding 14 digits in Schedule TDS(1)" & i & Chr(13)
                ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
            End If
            If Len("" & AmtClaimedOnOwnHands_TDS1(i)) > 14 Then
                MsgTDS1 = MsgTDS1 + "* Amount of col.no 3 cannot exceed 99999999999999 in Schedule TDS(1)" & i & Chr(13)
                ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
            End If
        Else
            MsgTDS1 = MsgTDS1 + "* Amount in col.no 3 is invalid in Schedule TDS(1)" & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
        End If
        
    If IsNumeric(AmtCarriedFwd_TDS1(i)) Then
        If AmtCarriedFwd_TDS1(i) < 0 Then
            MsgTDS1 = MsgTDS1 + "* Amount in col.no 4 should be Numeric, Non Negative, not exceeding 14 digits in Schedule TDS(1)" & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
            Exit Function
        End If
        
        
        If Len("" & AmtCarriedFwd_TDS1(i)) > 14 Then
            MsgTDS1 = MsgTDS1 + "* Amount in col.no 4 cannot exceed 99999999999999 in Schedule TDS(1)" & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
            Exit Function
        End If
    Else
        MsgTDS1 = MsgTDS1 + "* Amount in col.no 4 is invalid in Schedule TDS(1)" & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
        Exit Function
    End If
                
    If AmtCarriedFwd_TDS1(i) > (AmtClaimedOnOwnHands_TDS1(i)) Then
        MsgTDS1 = MsgTDS1 + "* Amount in col.no 4 cannot exceed the Col 3 in Schedule TDS(1)" & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS1 = False
        Exit Function
    End If
        
Next
End Function


Function ValidatesheetIT() As Boolean
On Error Resume Next
Dim rangname, cellcount, name, intcntr As Variant
     ValidatesheetIT = True
     MsgSchIT = "IT." & Chr(13)
     subProcCaption = "Validating IT"
     setTblinfo_IT
     noOfProcessSub = end_IT
    Dim k  As Long
For k = 1 To end_IT
    If Not (Sheet18.Range("TaxP.BSRCode").Cells.item(k) <> "" Or Sheet18.Range("TaxP.DateDep").Cells.item(k) <> "" Or Sheet18.Range("TaxP.SrlNoOfChaln").Cells.item(k) <> "" Or Sheet18.Range("TaxP.Amt").Cells.item(k) <> "") Then
        MsgSchIT = MsgSchIT & "* At Sr. No " & k & " Please fill all the Mandatory Fields in Sheet IT " & Chr(13)
        ValidatesheetIT = False
        Exit Function
    End If
    UpdateProgressBar
Next
    
 If end_IT > 0 Then GoTo xyz2
 
     cellcount = Sheet18.Range("TaxP.BSRCode").count
     rangname = "TaxP.BSRCode|TaxP.DateDep|TaxP.SrlNoOfChaln|TaxP.Amt"
     rangname = Split(rangname, "|")
    
    For intcntr = 1 To cellcount
        If Sheet18.Range("TaxP.BSRCode").Cells.item(intcntr) <> "" Or Sheet18.Range("TaxP.DateDep").Cells.item(intcntr) <> "" Or Sheet18.Range("TaxP.SrlNoOfChaln").Cells.item(intcntr) <> "" Or Sheet18.Range("TaxP.Amt").Cells.item(intcntr) <> "" Then
            For Each name In rangname
                If Sheet18.Range(name).Cells.item(intcntr) = "" Then ValidatesheetIT = False
            Next
'            If ValidatesheetIT = False Then
'                MsgSchIT = MsgSchIT & "* At Sr. No " & intcntr & " Please fill all the Mandatory Fields in Sheet IT" & Chr(13)
'                GoTo xyz2
'            End If
        End If
        
    Next
    
xyz2:
    
    If Not ValidateBSRCode() Then ValidatesheetIT = False
    If Not ValidateDateDep() Then ValidatesheetIT = False
    If Not ValidateSrlNoOfChaln() Then ValidatesheetIT = False
    If Not ValidateAmt() Then ValidatesheetIT = False
 If Not ValidateMandatoryShIT() Then ValidatesheetIT = False

        setTableInfo_ITDate
        setTableInfo_ITSerialChalan
        setTableInfo_ITTaxPd
        
'
'         If ((end_IT <> end_Date_IT) Or (end_IT <> end_SerialCh_IT) Or (end_IT <> end_Taxpd_IT)) Then
'            MsgSchIT = MsgSchIT & "Please fill all the Mandatory Fields " & Chr(13)
'            ValidatesheetIT = False
'        End If
         
    If Len(Sheet18.Range("TaxP.Sum").value) > 14 Then
        MsgSchIT = MsgSchIT & "* Total of IT Should not be exceed 14 digits " & Chr(13)
        ValidatesheetIT = False
    End If
         
End Function

Function ValidateMandatoryShIT() As Boolean
ValidateMandatoryShIT = True
Dim i As Long
Dim flag As Boolean
flag = True
For i = 1 To Sheet18.Range("TaxP.BSRCode").Rows.count
    If Sheet18.Range("TaxP.BSRCode").item(i).value <> "" Or Sheet18.Range("TaxP.DateDep").item(i).value <> "" Or _
        Sheet18.Range("TaxP.SrlNoOfChaln").item(i).value <> "" Or Sheet18.Range("TaxP.Amt").item(i).value <> "" Then
        
        If Sheet18.Range("TaxP.BSRCode").item(i).value = "" Then
            flag = False
'            MsgSchIT = MsgSchIT + "* Please enter the BSR code in Schedule IT" & Chr(13)
             'SIT-56341
                MsgSchIT = MsgSchIT + "* Please enter the BSR code " & Chr(13)

        End If
        
        
        
       'Added in this function ValidateBSRCode
''sudha_205 -26
''
''        If Sheet18.Range("TaxP.BSRCode").item(i).value = "0000000" Then
''            flag = False
''            MsgSchIT = MsgSchIT + "* Please enter the BSR code in Schedule IT" & Chr(13)
''             SIT -56341
''                MsgSchIT = MsgSchIT + "* Please enter a valid 7 characters BSR Code" & Chr(13)
''
''        End If

'-------------------------------------------------


        If flag = False Then
            ValidateMandatoryShIT = False
            Exit Function
        End If
        
    End If
Next i
End Function

Function ValidateBSRCode() As Boolean
    ValidateBSRCode = True
    setTblinfo_IT
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet18.Range("TaxP.BSRCode").Cells
    ReDim BSRCode_IT(end_IT)
For i = 1 To end_IT
    BSRCode_IT(i) = rangecells.item(i).value
    
    
            If Not chkCompulsory(BSRCode_IT(i)) Then
            MsgSchIT = MsgSchIT + "* BSR Code at Sr.NO " & i & " in Schedule IT is mandatory" & Chr(13)
            ValidateBSRCode = False
            Exit Function
            End If

        If Not Len(BSRCode_IT(i)) = 0 Then
                    If Not (chkNumeric(Mid(BSRCode_IT(i), 1, 3)) And checkfieldspecialcharacter(Mid(BSRCode_IT(i), 4, 4))) Then
                        'Change.02.03.2023.102.IDS.84
                        'MsgSchIT = MsgSchIT + "* Invalid BSR code in schedule TDS. BSR format should be First 3 numeric and next 4 alphanumeric, at Sr. No  " & i & " in Schedule IT.." & Chr(13)
'                        MsgSchIT = MsgSchIT + "* Invalid BSR code. BSR format should be First 3 numeric and next 4 alphanumeric, at Sr. No  " & i & " in Schedule IT.." & Chr(13)
               
                'SIT-56341
                        MsgSchIT = MsgSchIT + "* Invalid BSR. First three digits from the left stand are for district code; next three digits stand for revenue centre code within the district and last single digit stands for population range code at Sr. No  " & i & " in Schedule IT.." & Chr(13)
                        'End Change.IDS.84
                        ValidateBSRCode = False
                        Exit Function
                        End If
                  End If
                  
                  'Newly added by Bindu as per DE V4
             If BSRCode_IT(i) = "0000000" Then
                MsgSchIT = MsgSchIT + "* Please enter a valid 7 characters BSR Code at Sr. No " & i & "  in Sheet IT" & Chr(13)
                ValidateBSRCode = False
                Exit Function
             End If
        
    
    
    
        
'        If BSRCode_IT(i) = "" Then
'            MsgSchIT = MsgSchIT + "* Please enter the BSR code at Sr. No " & i & "  in Sheet IT" & Chr(13)
'            ValidateBSRCode = False
'            Exit Function
'        End If
'
'
'        If Not Len(BSRCode_IT(i)) = 0 Then
'            If Not (chkNumeric(Mid(BSRCode_IT(i), 1, 3)) And checkfieldspecialcharacter(Mid(BSRCode_IT(i), 4, 4))) Then
'                MsgSchIT = MsgSchIT + "* Invalid BSR Code.BSR format should be First 3 numeric and next 4 alphanumeric, at Sr. No  " & i & " in Sheet IT" & Chr(13)
'                ValidateBSRCode = False
'                Exit Function
'                End If
'          End If
'
'        If Not chkNumeric(BSRCode_IT(i)) Then
'        MsgSchIT = MsgSchIT + "* BSRCode at Sr. No " & i & "  in Sheet IT should be Numeric,Non negative,No decimal" & Chr(13)
'            ValidateBSRCode = False
'            Exit Function
'        End If
        
            
'        If (Not (IsNumeric(BSRCode_IT(i))) Or BSRCode_IT(i) < 0 Or Len(BSRCode_IT(i)) <> 7) Then
'            MsgSchIT = MsgSchIT + "* BSRCode at Sr. No " & i & "  in Sheet IT is invalid" & Chr(13)
'            ValidateBSRCode = False
'            Exit Function
'        End If
 Next




End Function
Function chkNumeric(field As Variant) As Boolean
    Dim k As Long
    Dim chkchar As String

    chkNumeric = True
    For k = 1 To Len(field)
    chkchar = Mid(field, k, 1)
        If Not IsNumeric(chkchar) Then
            chkNumeric = False
            Exit Function
        End If
    Next
End Function
Function ValidateSrlNoOfChaln() As Boolean
    ValidateSrlNoOfChaln = True
    'setTblinfo_IT
    setTableInfo_ITSerialChalan
    Dim i As Long
    Dim rangecells As Range
    Set rangecells = Sheet18.Range("TaxP.SrlNoOfChaln").Cells
    'ReDim SrlNoOfChaln_IT(end_IT)
    ReDim SrlNoOfChaln_IT(end_IT)
For i = 1 To end_IT
    SrlNoOfChaln_IT(i) = rangecells.item(i).value
         If Not chkCompulsory(SrlNoOfChaln_IT(i)) Then
         'Error Message Changed By Ayush for the SIT-56341
               ' MsgSchIT = MsgSchIT + "* Please enter the Serial No of challan at Sr. No " & i & " in Sheet IT" & Chr(13)
               MsgSchIT = MsgSchIT + """Please enter the serial number of Challan""" & Chr(13)
                ValidateSrlNoOfChaln = False
                Exit Function
         End If
         
         If Not checkspecialcharacterForPinCode(SrlNoOfChaln_IT(i)) Then
                MsgSchIT = MsgSchIT + "* Serial No of challan at Sr. No " & i & " is invalid in Sheet IT" & Chr(13)
                ValidateSrlNoOfChaln = False
                Exit Function
         End If
         

        If (Not IsNumeric(SrlNoOfChaln_IT(i))) Or SrlNoOfChaln_IT(i) < 0 Or Len(SrlNoOfChaln_IT(i)) > 5 Then
                MsgSchIT = MsgSchIT & "* Serial No of Challan at Sr. No " & i & " is Invalid in Sheet IT" & Chr(13)
                ValidateSrlNoOfChaln = False
                Exit Function
        End If
      
 Next
End Function
Function ValidateAmt() As Boolean
    ValidateAmt = True
    'setTblinfo_IT
    setTableInfo_ITSerialChalan
    Dim i As Long
    Dim rangecells As Range
    Set rangecells = Sheet18.Range("TaxP.Amt").Cells
    ReDim Amt_IT(end_IT)
For i = 1 To end_IT
        Amt_IT(i) = rangecells.item(i).value
     If Not chkCompulsory(Amt_IT(i)) Then
        MsgSchIT = MsgSchIT + "* Please enter the Amount at Sr. No " & i & "  in Sheet IT" & Chr(13)
        ValidateAmt = False
        Exit Function
     End If
         
    If IsNumeric(Amt_IT(i)) Then
        If Amt_IT(i) < 0 Then
            MsgSchIT = MsgSchIT + "* Amount at Sr. No " & i & " should be Numeric, Non Negative, not exceeding 14 digits in Sheet IT" & Chr(13)
            ValidateAmt = False
            Exit Function
        End If
    End If
              
         
    If (Not IsNumeric(Amt_IT(i))) Or Amt_IT(i) < 0 Or Len("" & Amt_IT(i)) > 14 Then
            MsgSchIT = MsgSchIT + "* Amount at Sr. No " & i & "  in Sheet IT is invalid in Sheet IT" & Chr(13)
            ValidateAmt = False
            Exit Function
    End If
   
Next
End Function



Function ValidateDateDep() As Boolean
    ValidateDateDep = True

    setTblinfo_IT
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet18.Range("TaxP.DateDep").Cells
    ReDim DateDep_IT(end_IT)
For i = 1 To end_IT
        DateDep_IT(i) = rangecells.item(i).value
        If Not chkCompulsory(DateDep_IT(i)) Then
                MsgSchIT = MsgSchIT + "* Please enter the Date of Deposit at Sr. No " & i & " in Sheet IT" & Chr(13)
                ValidateDateDep = False
                Exit Function
        End If
         
        If Not CheckDateddmmyyyy(DateDep_IT(i)) Then
                 ValidateDateDep = False
                 'Ayush_10/02
                 MsgSchIT = MsgSchIT + "* ""Date of Deposit at Sr. No " & i & "  in Sheet IT - Please enter in dd/mm/yyyy format.""" & Chr(13)

               'SIT-56341
'                   MsgSchIT = MsgSchIT + "* Date of Deposit at Sr. No " & i & "  in Sheet IT is invalid. Please enter in dd/mm/yyyy format" & Chr(13)


                 Exit Function
        Else
          DateDep_IT(i) = Dformat(DateDep_IT(i), "")
             If Not ChkMinInclusiveDate(DateDep_IT(i), "2021-04-01") Then '' dpk1201
             
                 
                 MsgSchIT = MsgSchIT + "* Date of Deposit at Sr. No " & i & " in Sheet IT should not be less than 01/04/2021" & Chr(13)
                ' MsgSchIT = MsgSchIT + "* Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1." & Chr(13)
                 ValidateDateDep = False
                 Exit Function
             End If
        End If
Next
End Function


Sub setTblinfo_IT()
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
 end_IT = ccount
 '----------02/12/2025
 'rngname_IT = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.AmtIT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1"
 rngname_IT = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.AmtIT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT"
 '------------------
 End Sub


Sub setTblinfo_IT2()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet18.Range("TaxP.Amt").count
 Set rangecells = Sheet18.Range("TaxP.Amt").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_IT2 = ccount
 '--------02/12/2025
 'rngname_IT22 = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.AmtIT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1"
 rngname_IT22 = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.AmtIT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT"
 '-----------------
 End Sub

Function CheckDateMinDDMMYYYY(ByVal dt As String, ByVal minday As Long, ByVal minmonth As Long, ByVal minyear As Long, ByVal errormsg As String) As Boolean
    CheckDateMinDDMMYYYY = True
        
    If (val(Mid(dt, 1, 4)) < minyear) Then
        CheckDateMinDDMMYYYY = False
        fmsgbox "* Invalid Date, " & errormsg
        GoTo exit1
    End If
    
    If (val(Mid(dt, 1, 4)) = minyear) And (val(Mid(dt, 6, 2)) < minmonth) Then
        CheckDateMinDDMMYYYY = False
        fmsgbox "* Invalid Date, " & errormsg
        GoTo exit1
    End If
    If (val(Mid(dt, 1, 4)) = minyear) And (val(Mid(dt, 6, 2)) < minmonth) And (val(Mid(dt, 9, 2)) < minday) Then
        CheckDateMinDDMMYYYY = False
        fmsgbox "* Invalid Date, " & errormsg
        GoTo exit1
    End If
exit1:
End Function

Function ChkMinInclusiveDate(Mininclusive As Variant, Mininclusivedate As Variant) As Boolean
'' both date must be in format yyyy-mm-dd
    ChkMinInclusiveDate = True
    If Len(Mininclusive) > 0 Then
        If Mid(Mininclusive, 1, 4) < Mid(Mininclusivedate, 1, 4) Then
            ChkMinInclusiveDate = False
            Exit Function
        Else
            If Mid(Mininclusive, 1, 4) = Mid(Mininclusivedate, 1, 4) Then
                If (Mid(Mininclusive, 6, 2) < Mid(Mininclusivedate, 6, 2)) Then
                    ChkMinInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Mininclusive, 6, 2) = Mid(Mininclusivedate, 6, 2))) Then
                        If (Mid(Mininclusive, 9, 2) < Mid(Mininclusivedate, 9, 2)) Then
                            ChkMinInclusiveDate = False
                            Exit Function
                        End If
                End If
            End If
        End If
    End If
  
End Function

Function CheckDateddmmyyyy(dt As Variant) As Boolean
CheckDateddmmyyyy = True
Dim dt1 As String
If Len(dt) > 0 Then
    dt1 = dt
If Mid(dt, 3, 1) <> "/" Then
    If Mid(dt, 3, 1) <> "\" Then
        If Mid(dt, 3, 1) <> "-" Then
            If Mid(dt, 3, 1) <> "." Then
                CheckDateddmmyyyy = False
            Else
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
            End If
        Else
            dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
        End If
    Else
        dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
    End If
End If
If Mid(dt, 6, 1) <> "/" Then
    If Mid(dt, 6, 1) <> "-" Then
        If Mid(dt, 6, 1) <> "\" Then
            If Mid(dt, 6, 1) <> "." Then
                CheckDateddmmyyyy = False
            Else
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
            End If
        Else
            dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
        End If
    Else
        dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
    End If
End If

If Not IsDate(dt) Then CheckDateddmmyyyy = False
If val(Mid(dt, 1, 2)) < 0 Then CheckDateddmmyyyy = False
If val(Mid(dt, 1, 2)) > 31 Then CheckDateddmmyyyy = False
If val(Mid(dt, 4, 2)) < 0 Then CheckDateddmmyyyy = False
If val(Mid(dt, 4, 2)) > 12 Then CheckDateddmmyyyy = False
If val(Mid(dt, 7, 4)) < 1500 Then CheckDateddmmyyyy = False
If val(Mid(dt, 7, 4)) > 3000 Then CheckDateddmmyyyy = False

If Mid(dt1, 3, 1) = "-" Then CheckDateddmmyyyy = False
If Mid(dt1, 3, 1) = "\" Then CheckDateddmmyyyy = False
If Mid(dt1, 3, 1) = "." Then CheckDateddmmyyyy = False

If Mid(dt1, 6, 1) = "-" Then CheckDateddmmyyyy = False
If Mid(dt1, 6, 1) = "\" Then CheckDateddmmyyyy = False
If Mid(dt1, 6, 1) = "." Then CheckDateddmmyyyy = False

End If

End Function

Function Dformat(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(dt) > 0 Then

Year = Mid(dt, 7, 4)
month = Mid(dt, 4, 2)
day = Mid(dt, 1, 2)
formateddate = Year & "-" & month & "-" & day
Dformat = formateddate

Else
Dformat = ""
End If
End Function
Function Dformat1(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet8.Range("DOB_1").value) > 0 Then

Year = Mid(Sheet8.Range("DOB_1").value, 7, 4)
month = Mid(Sheet8.Range("DOB_1").value, 4, 2)
day = Mid(Sheet8.Range("DOB_1").value, 1, 2)
'formateddate = Year & "-" & month & "-" & day
formateddate = day & "/" & month & "/" & Year
Dformat1 = formateddate

Else
Dformat1 = ""
End If
End Function
'Ayush_26-27
'139(8A)----
Function Dformat2(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet8.Range("date_139_8A").value) > 0 Then

Year = Mid(Sheet8.Range("date_139_8A").value, 7, 4)
month = Mid(Sheet8.Range("date_139_8A").value, 4, 2)
day = Mid(Sheet8.Range("date_139_8A").value, 1, 2)
'formateddate = Year & "-" & month & "-" & day
formateddate = day & "." & month & "." & Year
Dformat2 = formateddate

Else
Dformat2 = ""
End If
End Function
'Ayush_26-27
Function Dformat3(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet8.Range("IT_Date_Deposit_PY").value) > 0 Then

Year = Mid(Sheet8.Range("IT_Date_Deposit_PY").value, 7, 4)
month = Mid(Sheet8.Range("IT_Date_Deposit_PY").value, 4, 2)
day = Mid(Sheet8.Range("IT_Date_Deposit_PY").value, 1, 2)
formateddate = day & "-" & month & "-" & Year
Dformat3 = formateddate

Else
Dformat3 = ""
End If
End Function
Function Dformat4(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String

If Len(Sheet8.Range("Date_80GGC_PY").value) > 0 Then

Year = Mid(Sheet8.Range("Date_80GGC_PY").value, 7, 4)
month = Mid(Sheet8.Range("Date_80GGC_PY").value, 4, 2)
day = Mid(Sheet8.Range("Date_80GGC_PY").value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat4 = formateddate

Else
Dformat4 = ""
End If
End Function
Function Dformat5(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet8.Range("AY_80GGC").value) > 0 Then

Year = Mid(Sheet8.Range("AY_80GGC").value, 7, 4)
month = Mid(Sheet8.Range("AY_80GGC").value, 4, 2)
day = Mid(Sheet8.Range("AY_80GGC").value, 1, 2)
'formateddate = Year & "-" & month & "-" & day
formateddate = day & "/" & month & "/" & Year
Dformat5 = formateddate

Else
Dformat5 = ""
End If
End Function

Function ValidatePantype_text(strName As Variant) As Boolean
    ValidatePantype_text = True
    Dim len1, j As Long
    Dim s1 As String
    len1 = Len(strName)
    For j = 1 To len1
        s1 = Mid(strName, j, 1)
        If (((asc(s1) >= 65) And (asc(s1) <= 90)) Or (asc(s1) = 45)) Then
        Else
            ValidatePantype_text = False
        End If
    Next
End Function

Sub setTableinfo_NameTDS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").count
 Set rangecells = Sheet4.Range("TDSal.EmployerOrDeductorOrCollecterName").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Name_TDS1 = ccount
End Sub


Sub setTableinfo_IncTDS1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSal.IncChrgSal").count
 Set rangecells = Sheet4.Range("TDSal.IncChrgSal").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Inc_TDS1 = ccount
End Sub



Sub setTableinfo_Tot_TaxTDS1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSal.TotalTDSSal").count
 Set rangecells = Sheet4.Range("TDSal.TotalTDSSal").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TotTAx_TDS1 = ccount
End Sub






Sub setTableInfo_ITDate()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet18.Range("TaxP.DateDep").count
 Set rangecells = Sheet18.Range("TaxP.DateDep").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Date_IT = ccount
End Sub


Sub setTableInfo_ITSerialChalan()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet18.Range("TaxP.SrlNoOfChaln").count
 Set rangecells = Sheet18.Range("TaxP.SrlNoOfChaln").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_SerialCh_IT = ccount
End Sub


Sub setTableInfo_ITTaxPd()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet18.Range("TaxP.Amt").count
 Set rangecells = Sheet18.Range("TaxP.Amt").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Taxpd_IT = ccount
End Sub


'-------------------TCS Starts-------

Sub setTableInfo_CollectorName_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").count
 Set rangecells = Sheet17.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_CollectorName_TCS = ccount
End Sub

Sub setTableInfo_AmtPaid_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet17.Range("TCS.AmountPaid").count
 Set rangecells = Sheet17.Range("TCS.AmountPaid").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtPaid_TCS = ccount
End Sub

Sub setTableInfo_TaxCollect_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet17.Range("TCS.TotalTCS").count
 Set rangecells = Sheet17.Range("TCS.TotalTCS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TaxCollected_TCS = ccount
End Sub

Sub setTableInfo_AmtClain_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet17.Range("TCS.AmtTCSClaimedThisYear").count
 Set rangecells = Sheet17.Range("TCS.AmtTCSClaimedThisYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtClaim_TCS = ccount
End Sub


'Sub setTableInfo_AmtPortugese_TCS()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet17.Range("TCS.AmtClaimedBySpouse").Count
' Set rangecells = Sheet17.Range("TCS.AmtClaimedBySpouse").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_AmtPortuguese_TCS = ccount
'End Sub




Sub UNLOCKTDS1()
    Application.ScreenUpdating = False
    Sheet4.Unprotect getmsgstate


    Dim rngname_5_AL As Variant
    rngname_5_AL = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSal||TDSal.TotalTDSSal"
    UNLOCKRANGETDS1 (rngname_5_AL)

    Sheet4.Protect getmsgstate
    Application.ScreenUpdating = True
End Sub


Sub LOCKTDS1()
    Application.ScreenUpdating = False

    Sheet4.Unprotect getmsgstate

    Dim rngname_5_AL As Variant
    rngname_5_AL = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSal||TDSal.TotalTDSSal"
    LOCKRANGETDS1 (rngname_5_AL)

   Sheet4.Protect getmsgstate

    Application.ScreenUpdating = True
End Sub


Sub LOCKRANGETDS1(rangenamestring As Variant)
Application.EnableEvents = False
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet4.Range(rangenamestring(i))
            chcell.Locked = True
            chcell.Interior.Color = (&HD8D8D8)
            chcell.ClearContents
            Next
        Next
Application.EnableEvents = True
End Sub

Sub UNLOCKRANGETDS1(rangenamestring As Variant)
Application.EnableEvents = False
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet4.Range(rangenamestring(i))
            chcell.Locked = False
            chcell.Interior.Color = (&HCCFFCC)
            chcell.ClearContents
            Next
        Next
Application.EnableEvents = True
End Sub


Sub setTableinfo_TDS2_TAN()
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
 end_TDS2_1 = ccount
End Sub

Sub setTableinfo_TDS2_FinYear()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.finyear").count
 Set rangecells = Sheet4.Range("TDsOthr.finyear").Cells
 For mIntCtr = 1 To mIntCells
     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
            ccount = ccount + 1
        End If
 Next
 end_TDS2_2 = ccount
End Sub

Sub setTableinfo_TDS2_Tds()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.tds").count
 Set rangecells = Sheet4.Range("TDsOthr.tds").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_3 = ccount
End Sub

Sub setTableinfo_TDS2_TdsDeducted()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.tdsdeducted").count
 Set rangecells = Sheet4.Range("TDsOthr.tdsdeducted").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_4 = ccount
End Sub

Sub setTableinfo_TDS2_TdsClaimed()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.tdsclaimed").count
 Set rangecells = Sheet4.Range("TDsOthr.tdsclaimed").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_5 = ccount
End Sub

Sub setTableinfo_TDS2_GrossAmount()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.grossamount").count
 Set rangecells = Sheet4.Range("TDsOthr.grossamount").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_6 = ccount
End Sub

Sub setTableinfo_TDS2_HeadIncome()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.headincome").count
 Set rangecells = Sheet4.Range("TDsOthr.headincome").Cells
 For mIntCtr = 1 To mIntCells
     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
            ccount = ccount + 1
        End If
 Next
 end_TDS2_7 = ccount
End Sub

Sub setTableinfo_TDS2_TdsCarried()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr.tdscarried").count
 Set rangecells = Sheet4.Range("TDsOthr.tdscarried").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_8 = ccount
End Sub

'Ayush_30/01
'Sub setTableinfo_TDS2_SectionTDS()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet4.Range("TDsOthr.SectionTDS").count
' Set rangecells = Sheet4.Range("TDsOthr.SectionTDS").Cells
' For mIntCtr = 1 To mIntCells
'     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
'            ccount = ccount + 1
'        End If
' Next
' end_TDS2_10 = ccount
'End Sub

'Function ValidateSheetTDS2() As Boolean
'On Error Resume Next
'
'ValidateSheetTDS2 = True
'    MsgTDS2 = ""
'    setTableinfo_TDS2_TAN
'    setTableinfo_TDS2_SectionTDSDeducted '''''''''Added by malli
'    setTableinfo_TDS2_FinYear
'    setTableinfo_TDS2_Tds
'    setTableinfo_TDS2_TdsDeducted
'    setTableinfo_TDS2_TdsClaimed
'    setTableinfo_TDS2_GrossAmount
'    setTableinfo_TDS2_HeadIncome
'    'setTableinfo_TDS2_TdsCarried
'
'    end_TDS2_1 = WorksheetFunction.Max(0, end_TDS2_1, end_TDS2_2, end_TDS2_3, end_TDS2_4, end_TDS2_5, end_TDS2_6, end_TDS2_7, end_TDS2_9)
'
'If Not ValidateTAN_TDS2() Then ValidateSheetTDS2 = False
'If Not ValidateSectionTDSDeducted_TDS2() Then ValidateSheetTDS2 = False '''''''Added by Malli
'    If Not ValidateTdsClaimed_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidateFinYear_Tds_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidateGrossAmount_TDS2() Then ValidateSheetTDS2 = False
'
'    If Len(Sheet4.Range("TDS2i.Sum").value) > 14 Then
'        MsgTDS2 = MsgTDS2 & "* Total of TDS2i Should not be exceed 14 digits in Sheet TDS2i" & Chr(13)
'        ValidateSheetTDS2 = False
'    End If
'
'End Function
Function ValidateSheetTDS2() As Boolean
On Error Resume Next

ValidateSheetTDS2 = True
    MsgTDS2 = ""
    setTableinfo_TDS2_TAN
    setTableinfo_TDS2_FinYear
    setTableinfo_TDS2_Tds
    setTableinfo_TDS2_TdsDeducted
    setTableinfo_TDS2_TdsClaimed
    setTableinfo_TDS2_GrossAmount
    setTableinfo_TDS2_HeadIncome
    'Ayush_30/01
   ' setTableinfo_TDS2_SectionTDS
    'setTableinfo_TDS2_TdsCarried
    
   end_TDS2_1 = WorksheetFunction.Max(0, end_TDS2_1, end_TDS2_2, end_TDS2_3, end_TDS2_4, end_TDS2_5, end_TDS2_6, end_TDS2_7)
   ' end_TDS2_1 = WorksheetFunction.Max(0, end_TDS2_1, end_TDS2_2, end_TDS2_3, end_TDS2_4, end_TDS2_5, end_TDS2_6, end_TDS2_7, end_TDS2_10)

If Not ValidateTAN_TDS2() Then ValidateSheetTDS2 = False
    If Not ValidateTdsClaimed_TDS2() Then ValidateSheetTDS2 = False
    If Not ValidateFinYear_Tds_TDS2() Then ValidateSheetTDS2 = False
    If Not ValidateGrossAmount_TDS2() Then ValidateSheetTDS2 = False
    'Ayush_30/01
    
        'sudha_2025-26
    
    If Not ValidateSectionTDS_Tds_TDS2() Then ValidateSheetTDS2 = False
   '==================
   
   
   ' If Not ValidateSectionTDS_Tds_TDS2() Then ValidateSheetTDS2 = False
    
    If Len(Sheet4.Range("TDS2i.Sum").value) > 14 Then
        MsgTDS2 = MsgTDS2 & "* Total of TDS2i Should not be exceed 14 digits in Sheet TDS2i" & Chr(13)
        ValidateSheetTDS2 = False
    End If
    
End Function

Sub setTableinfo_TDS2ii_PAN()
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
End Sub
'
'Sub setTableinfo_TDS2ii_Aadhaar()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet4.Range("TDsOthr2.AADHAAR").Count
' Set rangecells = Sheet4.Range("TDsOthr2.AADHAAR").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2ii_A = ccount
'End Sub

Sub setTableinfo_TDS2ii_FinYear()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.finyear").count
 Set rangecells = Sheet4.Range("TDsOthr2.finyear").Cells
 For mIntCtr = 1 To mIntCells
     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
            ccount = ccount + 1
        End If
 Next
 end_TDS2ii_2 = ccount
End Sub

Sub setTableinfo_TDS2ii_Tds()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.tds").count
 Set rangecells = Sheet4.Range("TDsOthr2.tds").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2ii_3 = ccount
End Sub

Sub setTableinfo_TDS2ii_TdsDeducted()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.tdsdeducted").count
 Set rangecells = Sheet4.Range("TDsOthr2.tdsdeducted").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2ii_4 = ccount
End Sub

Sub setTableinfo_TDS2ii_TdsClaimed()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.tdsclaimed").count
 Set rangecells = Sheet4.Range("TDsOthr2.tdsclaimed").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2ii_5 = ccount
End Sub

Sub setTableinfo_TDS2ii_GrossAmount()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.grossamount").count
 Set rangecells = Sheet4.Range("TDsOthr2.grossamount").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2ii_6 = ccount
End Sub

Sub setTableinfo_TDS2ii_HeadIncome()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.headamount").count
 Set rangecells = Sheet4.Range("TDsOthr2.headamount").Cells
 For mIntCtr = 1 To mIntCells
     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
            ccount = ccount + 1
        End If
 Next
 end_TDS2ii_7 = ccount
End Sub

Sub setTableinfo_TDS2ii_TdsCarried()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDsOthr2.tdscarried").count
 Set rangecells = Sheet4.Range("TDsOthr2.tdscarried").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2ii_8 = ccount
End Sub

'Ayush_30/01
'Sub setTableinfo_TDS2_SectionTDS_ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet4.Range("TDsOthr.SectionTDS_ii").count
' Set rangecells = Sheet4.Range("TDsOthr.SectionTDS_ii").Cells
' For mIntCtr = 1 To mIntCells
'     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
'            ccount = ccount + 1
'        End If
' Next
' end_TDS2ii_9 = ccount
'End Sub
Function ValidatePAN_TDS2ii() As Boolean
ValidatePAN_TDS2ii = True

noOfProcessSub = end_TDS2ii_1

Dim rangecells As Range
Dim rangecells1 As Range

Set rangecells = Range("TDsOthr2.PAN").Cells
Set rangecells1 = Range("TDsOthr2.AADHAAR").Cells
Dim i As Long
 ReDim PAN_TDS2ii_1(end_TDS2ii_1)
 ReDim Aadhaar_TDS2ii_1(end_TDS2ii_1)

For i = 1 To end_TDS2ii_1
    PAN_TDS2ii_1(i) = rangecells.item(i).value
    
    If Not Len(PAN_TDS2ii_1(i)) = 0 Then
        If Not mdIncomeDetails.CheckPAN_Tenant(UCase(Mid(PAN_TDS2ii_1(i), 1, 10))) Then
        MsgTDS2ii = MsgTDS2ii + "* Invalid PAN.PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr. No  " & i & " in Schedule TDS2(ii)" & Chr(13)
        ValidatePAN_TDS2ii = False
        End If
    End If
    
    If Not chkCompulsory(PAN_TDS2ii_1(i)) Then
                MsgTDS2ii = MsgTDS2ii + "* Please enter the PAN  at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidatePAN_TDS2ii = False
                Exit Function
    End If
    
    Aadhaar_TDS2ii_1(i) = rangecells1.item(i).value
    
    If Not Len(Aadhaar_TDS2ii_1(i)) = 0 Then
        If Len(Aadhaar_TDS2ii_1(i)) < 12 Then
        MsgTDS2ii = MsgTDS2ii + "* Please enter the valid Aadhaar number at Sr. No  " & i & " in Schedule TDS2(ii)" & Chr(13)
        ValidatePAN_TDS2ii = False
        End If
    End If
    
    UpdateProgressBar
Next
End Function

'Change-28.11.2022.103.06
'Function ValidateTdsClaimed_TDS2ii() As Boolean
'ValidateTdsClaimed_TDS2ii = True
'noOfProcessSub = end_TDS2ii_1
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'Dim rangecells2 As Range
'Dim i As Long
'
'Set rangecells = Range("TDsOthr2.tdsclaimed").Cells
'Set rangecells1 = Range("TDsOthr2.tds").Cells
'Set rangecells2 = Range("TDsOthr2.tdsdeducted").Cells
'
' ReDim TdsClaimed_TDS2ii_1(end_TDS2ii_1)
' ReDim Tds_TDS2ii_1(end_TDS2ii_1)
' ReDim Tdsdeducted_TDS2ii_1(end_TDS2ii_1)
'
' For i = 1 To end_TDS2ii_1
'    TdsClaimed_TDS2ii_1(i) = rangecells.item(i).value
'    Tds_TDS2ii_1(i) = rangecells1.item(i).value
'    Tdsdeducted_TDS2ii_1(i) = rangecells2.item(i).value
'
'
'    If Not chkCompulsory(TdsClaimed_TDS2ii_1(i)) Then
'      MsgTDS2ii = MsgTDS2ii + "* Please enter Tds claimed at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'       ValidateTdsClaimed_TDS2ii = False
'       Exit Function
'    End If
'
'    If Tds_TDS2ii_1(i) <> "" Then
'    If TdsClaimed_TDS2ii_1(i) > Tds_TDS2ii_1(i) Then
'      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'      MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted. at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'       ValidateTdsClaimed_TDS2ii = False
'       Exit Function
'    End If
'    End If
'    If Tdsdeducted_TDS2ii_1(i) <> "" Then
'    If TdsClaimed_TDS2ii_1(i) > Tdsdeducted_TDS2ii_1(i) Then
'      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'      MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted. at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'       ValidateTdsClaimed_TDS2ii = False
'       Exit Function
'    End If
'    End If
'
'
'
' UpdateProgressBar
'Next
'End Function


Function ValidateFinYear_Tds_TDS2ii() As Boolean
ValidateFinYear_Tds_TDS2ii = True
noOfProcessSub = end_TDS2ii_1

Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
'Dim rangecells4 As Range    'added by Chetan C M for AY 2025-26 on 04/02/2025


Dim i As Long

Set rangecells = Range("TDsOthr2.finyear").Cells
Set rangecells1 = Range("TDsOthr2.tds").Cells
Set rangecells2 = Range("TDsOthr2.tdsdeducted").Cells
Set rangecells3 = Range("TDsOthr2.carried").Cells
'Set rangecells4 = Range("TDsOthr.SectionTDS_ii").Cells  'added by Chetan C M for AY 2025-26 on 04/02/2025


 ReDim FinYear_TDS2ii_1(end_TDS2ii_1)
 ReDim Tds_TDS2ii_1(end_TDS2ii_1)
 ReDim Tdsdeducted_TDS2ii_1(end_TDS2ii_1)
 ReDim Tdscarried_TDS2ii_1(end_TDS2ii_1)
 'ReDim SectionTDS_TDS2_2ii(end_TDS2ii_1)    'added by Chetan C M for AY 2025-26 on 04/02/2025


 For i = 1 To end_TDS2ii_1
    FinYear_TDS2ii_1(i) = rangecells.item(i).value
    Tds_TDS2ii_1(i) = rangecells1.item(i).value
    Tdsdeducted_TDS2ii_1(i) = rangecells2.item(i).value
    Tdscarried_TDS2ii_1(i) = rangecells3.item(i).value
   ' SectionTDS_TDS2_2ii(i) = rangecells4.item(i).value  'added by Chetan C M for AY 2025-26 on 04/02/2025
    
    
    If Not isdropdownblank(FinYear_TDS2ii_1(i)) Then
    If Tds_TDS2ii_1(i) = "" Then
    MsgTDS2ii = MsgTDS2ii + "* Please enter the Amount of TDS brought forward at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidateFinYear_Tds_TDS2ii = False
                Exit Function
    End If
    End If
    
    If Tds_TDS2ii_1(i) <> "" Then
    If isdropdownblank(FinYear_TDS2ii_1(i)) Then
    MsgTDS2ii = MsgTDS2ii + "* Please select Financial Year in which TDS deducted from drop-down  at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidateFinYear_Tds_TDS2ii = False
                Exit Function
    End If
    End If
    
    'If FinYear_TDS2ii_1(i) > 2021 And UCase(FinYear_TDS2ii_1(i)) <> "(SELECT)" Then dpk1201
'    If FinYear_TDS2ii_1(i) > 2022 And UCase(FinYear_TDS2ii_1(i)) <> "(SELECT)" Then
    'If FinYear_TDS2ii_1(i) > "2023-24" And UCase(FinYear_TDS2ii_1(i)) <> "(SELECT)" Then
    'Ayush_2026-27
'        If FinYear_TDS2ii_1(i) > "2024-25" And UCase(FinYear_TDS2ii_1(i)) <> "(SELECT)" Then
    If FinYear_TDS2ii_1(i) > "2025 - 26" And UCase(FinYear_TDS2ii_1(i)) <> "(SELECT)" Then


'    If FinYear_TDS2ii_1(i) > 2022 And UCase(FinYear_TDS2ii_1(i)) <> "(SELECT)" Then
    'MsgTDS2ii = MsgTDS2ii + "* Financial Year in which deducted cannot be more than 2021 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13) dpk1201
    MsgTDS2ii = MsgTDS2ii + "* Financial Year in which deducted cannot be more than 2025-26 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidateFinYear_Tds_TDS2ii = False
                Exit Function
    End If
    
    If TdsClaimed_TDS2ii_1(i) <> "" Then
    If isdropdownblank(FinYear_TDS2ii_1(i)) And Tds_TDS2ii_1(i) = "" Then
    If Tdsdeducted_TDS2ii_1(i) = "" Then
    MsgTDS2ii = MsgTDS2ii + "* Please enter TDS deducted  at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
           ValidateFinYear_Tds_TDS2ii = False
          Exit Function
    End If
    End If
    End If
    
     If Tds_TDS2ii_1(i) <> "" Then
    If Tdsdeducted_TDS2ii_1(i) <> "" Then
            'MsgTDS2ii = MsgTDS2ii + "* Column 4 and Column 5 cannot filled together at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
            MsgTDS2ii = MsgTDS2ii + "* Column TDS b/f  and Column TDS Deducted  cannot filled together at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
            ValidateFinYear_Tds_TDS2ii = False
            Exit Function
    End If
    End If
    
    If Tdsdeducted_TDS2ii_1(i) <> "" Then
    If Tds_TDS2ii_1(i) <> "" Then
            'MsgTDS2ii = MsgTDS2ii + "* Column 4 and Column 5 cannot filled together at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
            MsgTDS2ii = MsgTDS2ii + "* Column TDS b/f  and Column TDS Deducted  cannot filled together at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
            ValidateFinYear_Tds_TDS2ii = False
            Exit Function
    End If
    End If
    
 UpdateProgressBar
Next
End Function



'Ayush_30/01
Function ValidateSectionTDS_Tds_TDS2_ii() As Boolean
ValidateSectionTDS_Tds_TDS2_ii = True
noOfProcessSub = end_TDS2ii_1

Dim rangecells As Range


Dim i As Long

Set rangecells = Range("TDsOthr2.SectionTDSDeducted").Cells


 ReDim SectionTDS_TDS2_2ii(end_TDS2ii_1)
  For i = 1 To end_TDS2ii_1
    SectionTDS_TDS2_2ii(i) = rangecells.item(i).value


    If isdropdownblank(SectionTDS_TDS2_2ii(i)) Then

    MsgTDS2ii = MsgTDS2ii + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidateSectionTDS_Tds_TDS2_ii = False
                Exit Function
    End If

'Newly added by sai on 24/04/2025
Dim TDS_section_error_1 As Range
  Dim TDS_section_error_count_1, iCount_error_1 As Long
      Set TDS_section_error_1 = Sheet8.Range("TDS_section_error").Cells
       TDS_section_error_count_1 = Sheet8.Range("TDS_section_error").count
       
        For iCount_error_1 = 1 To TDS_section_error_count_1
       
          If UCase(Trim(SectionTDS_TDS2_2ii(i))) = UCase(Trim(TDS_section_error_1.item(iCount_error_1).value)) Then
            MsgTDS2ii = MsgTDS2ii + "*TDS relating to salaries is required to be claimed in Schedule TDS1 at Sr. No  " & i & "" & Chr(13)
            ValidateSectionTDS_Tds_TDS2_ii = False
          End If
       
        Next



 UpdateProgressBar
Next
End Function



Function ValidateGrossAmount_TDS2ii() As Boolean
ValidateGrossAmount_TDS2ii = True
noOfProcessSub = end_TDS2ii_1

Dim rangecells As Range
Dim rangecells1 As Range
Dim i As Long

Set rangecells = Range("TDsOthr2.grossamount").Cells
Set rangecells1 = Range("TDsOthr2.headamount").Cells
 ReDim GrossAmount_TDS2ii_1(end_TDS2ii_1)
 ReDim Headincome_TDS2ii_1(end_TDS2ii_1)
 For i = 1 To end_TDS2ii_1
    GrossAmount_TDS2ii_1(i) = rangecells.item(i).value
    Headincome_TDS2ii_1(i) = rangecells1.item(i).value
    
    
    If TdsClaimed_TDS2ii_1(i) <> "" And TdsClaimed_TDS2ii_1(i) > 0 Then
    If GrossAmount_TDS2ii_1(i) = "" Then
    MsgTDS2ii = MsgTDS2ii + "* Please enter the amount in Gross Amount at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidateGrossAmount_TDS2ii = False
                Exit Function
    End If
    End If
    
    'Ayush_02/06
    
    If Not IsNumeric(GrossAmount_TDS2ii_1(i)) Then
            MsgTDS2ii = MsgTDS2ii & "* Gross Amount at Sr. No  " & i & "  in schedule TDS2(ii) should be Numeric value" & Chr(13)
            ValidateGrossAmount_TDS2ii = False
            Exit Function
        End If
        '-------------
    
    If TdsClaimed_TDS2ii_1(i) <> "" And TdsClaimed_TDS2ii_1(i) > 0 Then
    If isdropdownblank(Headincome_TDS2ii_1(i)) Then
    MsgTDS2ii = MsgTDS2ii + "* Please select the drop down of head of income for which corresponding income offered at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
                ValidateGrossAmount_TDS2ii = False
                Exit Function
    End If
    End If
    
 UpdateProgressBar
Next
End Function

'Function ValidateSheetTDS2ii() As Boolean
'On Error Resume Next
'
'ValidateSheetTDS2ii = True
'    MsgTDS2ii = ""
'    setTableinfo_TDS2ii_PAN
'    setTableinfo_TDS2ii_SectionTDSDeducted ''''''''''''''''Added by malli
'    setTableinfo_TDS2ii_FinYear
'    setTableinfo_TDS2ii_Tds
'    setTableinfo_TDS2ii_TdsDeducted
'    setTableinfo_TDS2ii_TdsClaimed
'    setTableinfo_TDS2_GrossAmount
'    setTableinfo_TDS2_HeadIncome
'    'setTableinfo_TDS2ii_TdsCarried
'
'    end_TDS2ii_1 = WorksheetFunction.Max(0, end_TDS2ii_1, end_TDS2ii_2, end_TDS2ii_3, end_TDS2ii_4, end_TDS2ii_5, end_TDS2ii_6, end_TDS2ii_7, end_TDS2ii_9)
'
'If Not ValidatePAN_TDS2ii() Then ValidateSheetTDS2ii = False
'If Not ValidateSectionTDSDeducted_TDS2ii() Then ValidateSheetTDS2ii = False 'Malli
'    If Not ValidateTdsClaimed_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidateFinYear_Tds_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidateGrossAmount_TDS2ii() Then ValidateSheetTDS2ii = False
'
'    If Len(Sheet4.Range("TDS2ii.Sum").value) > 14 Then
'        MsgTDS2ii = MsgTDS2ii & "* Total of TDS2ii Should not be exceed 14 digits in Sheet TDS2ii" & Chr(13)
'        ValidateSheetTDS2ii = False
'    End If
'
'End Function
Function ValidateSheetTDS2ii() As Boolean
On Error Resume Next

ValidateSheetTDS2ii = True
    MsgTDS2ii = ""
    setTableinfo_TDS2ii_PAN
    setTableinfo_TDS2ii_FinYear
    setTableinfo_TDS2ii_Tds
    setTableinfo_TDS2ii_TdsDeducted
    setTableinfo_TDS2ii_TdsClaimed
    setTableinfo_TDS2_GrossAmount
    setTableinfo_TDS2_HeadIncome
    'Ayush_30/01
   ' setTableinfo_TDS2_SectionTDS_ii
    'setTableinfo_TDS2ii_TdsCarried
    
    end_TDS2ii_1 = WorksheetFunction.Max(0, end_TDS2ii_1, end_TDS2ii_2, end_TDS2ii_3, end_TDS2ii_4, end_TDS2ii_5, end_TDS2ii_6, end_TDS2ii_7)

If Not ValidatePAN_TDS2ii() Then ValidateSheetTDS2ii = False
    If Not ValidateTdsClaimed_TDS2ii() Then ValidateSheetTDS2ii = False
    If Not ValidateFinYear_Tds_TDS2ii() Then ValidateSheetTDS2ii = False
    'Ayush_30/01  ValidateSectionTDS_Tds_TDS2_ii
   If Not ValidateSectionTDS_Tds_TDS2_ii() Then ValidateSheetTDS2ii = False
    If Not ValidateGrossAmount_TDS2ii() Then ValidateSheetTDS2ii = False
    
    If Len(Sheet4.Range("TDS2ii.Sum").value) > 14 Then
        MsgTDS2ii = MsgTDS2ii & "* Total of TDS2ii Should not be exceed 14 digits in Sheet TDS2ii" & Chr(13)
        ValidateSheetTDS2ii = False
    End If
    
End Function


Function ValidateTAN_TDS2() As Boolean
ValidateTAN_TDS2 = True

noOfProcessSub = end_TDS2_1

Dim rangecells As Range
Dim rangecells1 As Range

Set rangecells = Range("TDsOthr.TAN").Cells
'Set rangecells1 = Range("TDsOthr.TAN").Cells
Dim i As Long
 ReDim TAN_TDS2_1(end_TDS2_1)

For i = 1 To end_TDS2_1
    TAN_TDS2_1(i) = rangecells.item(i).value
    If Not Len(TAN_TDS2_1(i)) = 0 Then
        If Not mdIncomeDetails.CheckTAN(UCase(Mid(TAN_TDS2_1(i), 1, 10))) Then
'        MsgTDS2 = MsgTDS2 + "* TAN of the Deductor at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet in Schedule TDS2(i)" & Chr(13)


        'SIT-56341
         MsgTDS2 = MsgTDS2 + "* Invalid TAN at Sr. No  " & i & " . TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS2(i)" & Chr(13)
        ValidateTAN_TDS2 = False
        End If
    End If
    
    If Not chkCompulsory(TAN_TDS2_1(i)) Then
                MsgTDS2 = MsgTDS2 + "* Please enter TAN of the Deductor  at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateTAN_TDS2 = False
                Exit Function
    End If
    
    If Not ValidateTANCodes(UCase(TAN_TDS2_1(i))) Then
                MsgTDS2 = MsgTDS2 + "* Invalid TAN at Sr.No " & i & " in Schedule TDS2(i).Please enter valid TAN" & Chr(13)
                ValidateTAN_TDS2 = False
                Exit Function
            End If
    
    UpdateProgressBar
Next
End Function

'Change-28.11.2022.103.04
'Function ValidateTdsClaimed_TDS2() As Boolean
'ValidateTdsClaimed_TDS2 = True
'noOfProcessSub = end_TDS2_1
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'Dim rangecells2 As Range
'Dim i As Long
'
'Set rangecells = Range("TDsOthr.tdsclaimed").Cells
'Set rangecells1 = Range("TDsOthr.tds").Cells
'Set rangecells2 = Range("TDsOthr.tdsdeducted").Cells
'
' ReDim TdsClaimed_TDS2_1(end_TDS2_1)
' ReDim Tds_TDS2_1(end_TDS2_1)
' ReDim Tdsdeducted_TDS2_1(end_TDS2_1)
'
' For i = 1 To end_TDS2_1
'    TdsClaimed_TDS2_1(i) = rangecells.item(i).value
'    Tds_TDS2_1(i) = rangecells1.item(i).value
'    Tdsdeducted_TDS2_1(i) = rangecells2.item(i).value
'
'
'    If Not chkCompulsory(TdsClaimed_TDS2_1(i)) Then
'      MsgTDS2 = MsgTDS2 + "* Please enter Tds claimed at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'       ValidateTdsClaimed_TDS2 = False
'       Exit Function
'
'    End If
'
'    If Tds_TDS2_1(i) <> "" Then
'    If TdsClaimed_TDS2_1(i) > Tds_TDS2_1(i) Then
'      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'      MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'       ValidateTdsClaimed_TDS2 = False
'       Exit Function
'    End If
'    End If
'    If Tdsdeducted_TDS2_1(i) <> "" Then
'    If TdsClaimed_TDS2_1(i) > Tdsdeducted_TDS2_1(i) Then
'      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'      MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'       ValidateTdsClaimed_TDS2 = False
'       Exit Function
'    End If
'    End If
'    If Tdsdeducted_TDS2_1(i) <> "" Then
'    If TdsClaimed_TDS2_1(i) > Tdsdeducted_TDS2_1(i) Then
'      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'      MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'       ValidateTdsClaimed_TDS2 = False
'       Exit Function
'    End If
'    End If
'
' UpdateProgressBar
'Next
'End Function


Function ValidateFinYear_Tds_TDS2() As Boolean
ValidateFinYear_Tds_TDS2 = True
noOfProcessSub = end_TDS2_1

Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
'Dim rangecells4 As Range    'added by Chetan C M for AY 2025-26 on 04/02/2025

Dim i As Long

Set rangecells = Range("TDsOthr.finyear").Cells
Set rangecells1 = Range("TDsOthr.tds").Cells
Set rangecells2 = Range("TDsOthr.tdsdeducted").Cells
Set rangecells3 = Range("TDsOthr.tdscarried").Cells
'Set rangecells4 = Range("TDsOthr.SectionTDS").Cells 'added by Chetan C M for AY 2025-26 on 04/02/2025

 ReDim FinYear_TDS2_1(end_TDS2_1)
 ReDim Tds_TDS2_1(end_TDS2_1)
 ReDim Tdsdeducted_TDS2_1(end_TDS2_1)
 ReDim Tdscarried_TDS2_1(end_TDS2_1)
' ReDim SectionTDS_TDS2_1(end_TDS2_1)    'added by Chetan C M for AY 2025-26 on 04/02/2025
 For i = 1 To end_TDS2_1
    FinYear_TDS2_1(i) = rangecells.item(i).value
    Tds_TDS2_1(i) = rangecells1.item(i).value
    Tdsdeducted_TDS2_1(i) = rangecells2.item(i).value
    Tdscarried_TDS2_1(i) = rangecells3.item(i).value
  '  SectionTDS_TDS2_1(i) = rangecells4.item(i).value    'added by Chetan C M for AY 2025-26 on 04/02/2025
    
    If Not isdropdownblank(FinYear_TDS2_1(i)) Then
    If Tds_TDS2_1(i) = "" Then
    MsgTDS2 = MsgTDS2 + "* Please enter the Amount of TDS brought forward at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateFinYear_Tds_TDS2 = False
                Exit Function
    End If
    End If
    
    If Tds_TDS2_1(i) <> "" Then
    If isdropdownblank(FinYear_TDS2_1(i)) Then
    MsgTDS2 = MsgTDS2 + "* Please select Financial Year in which TDS deducted from drop-down  at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateFinYear_Tds_TDS2 = False
                Exit Function
    End If
    End If
'    MsgBox FinYear_TDS2_1(i)
    
    'If FinYear_TDS2_1(i) > 2021 And UCase(FinYear_TDS2_1(i)) <> "(SELECT)" Then dpk1201
    'Ayush_26-27
  '  If FinYear_TDS2_1(i) > "2024 - 25" And UCase(FinYear_TDS2_1(i)) <> "(SELECT)" Then
  If FinYear_TDS2_1(i) > "2025 - 26" And UCase(FinYear_TDS2_1(i)) <> "(SELECT)" Then
    'MsgTDS2 = MsgTDS2 + "* Financial Year in which deducted cannot be more than 2021 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13) dpk1201
    MsgTDS2 = MsgTDS2 + "* Financial Year in which deducted cannot be more than 2025-26 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateFinYear_Tds_TDS2 = False
                Exit Function
    End If
    
    
    If TdsClaimed_TDS2_1(i) <> "" Then
    If isdropdownblank(FinYear_TDS2_1(i)) And Tds_TDS2_1(i) = "" Then
    If Tdsdeducted_TDS2_1(i) = "" Then
    MsgTDS2 = MsgTDS2 + "* Please enter TDS deducted   at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
           ValidateFinYear_Tds_TDS2 = False
          Exit Function
    End If
    End If
    End If
    
    If Tds_TDS2_1(i) <> "" Then
    If Tdsdeducted_TDS2_1(i) <> "" Then
            MsgTDS2 = MsgTDS2 + "* Column 4 and Column 5 cannot filled together at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
            ValidateFinYear_Tds_TDS2 = False
            Exit Function
    End If
    End If
    
    If Tdsdeducted_TDS2_1(i) <> "" Then
    If Tds_TDS2_1(i) <> "" Then
            MsgTDS2 = MsgTDS2 + "* Column 4 and Column 5 cannot filled together at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
            ValidateFinYear_Tds_TDS2 = False
            Exit Function
    End If
    End If
    
    
    
 UpdateProgressBar
Next
End Function


'Ayush_30/01
Function ValidateSectionTDS_Tds_TDS2() As Boolean
ValidateSectionTDS_Tds_TDS2 = True
noOfProcessSub = end_TDS2_1

Dim rangecells As Range


Dim i As Long

Set rangecells = Range("TDsOthr.SectionTDS").Cells


 ReDim SectionTDS_TDS2_1(end_TDS2_1)
  For i = 1 To end_TDS2_1
    SectionTDS_TDS2_1(i) = rangecells.item(i).value


    If isdropdownblank(SectionTDS_TDS2_1(i)) Then

    MsgTDS2 = MsgTDS2 + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateSectionTDS_Tds_TDS2 = False
                Exit Function
    End If

'Newly added by sai on 24/04/2025

  Dim TDS_section_error As Range
  Dim TDS_section_error_count, iCount_error As Long
      Set TDS_section_error = Sheet8.Range("TDS_Section_Error").Cells
       TDS_section_error_count = Sheet8.Range("TDS_Section_Error").count
       
        For iCount_error = 1 To TDS_section_error_count
       
          If UCase(Trim(SectionTDS_TDS2_1(i))) = UCase(Trim(TDS_section_error.item(iCount_error).value)) Then
            MsgTDS2 = MsgTDS2 + "*TDS relating to salaries is required to be claimed in Schedule TDS1 at Sr. No  " & i & "" & Chr(13)
            ValidateSectionTDS_Tds_TDS2 = False
          End If
       
        Next



 UpdateProgressBar
Next
End Function



Function ValidateGrossAmount_TDS2() As Boolean
ValidateGrossAmount_TDS2 = True
noOfProcessSub = end_TDS2_1

Dim rangecells As Range
Dim rangecells1 As Range
Dim i As Long

Set rangecells = Range("TDsOthr.grossamount").Cells
Set rangecells1 = Range("TDsOthr.headincome").Cells
 ReDim GrossAmount_TDS2_1(end_TDS2_1)
 ReDim Headincome_TDS2_1(end_TDS2_1)
 For i = 1 To end_TDS2_1
    GrossAmount_TDS2_1(i) = rangecells.item(i).value
    Headincome_TDS2_1(i) = rangecells1.item(i).value
    
    
    If TdsClaimed_TDS2_1(i) <> "" Then
    If GrossAmount_TDS2_1(i) = "" And rangecells.item(i).Locked = False Then
    MsgTDS2 = MsgTDS2 + "* Please enter the amount in Gross Amount at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateGrossAmount_TDS2 = False
                Exit Function
    End If
    End If
    
    'Ayush_02/06
    If Not IsNumeric(GrossAmount_TDS2_1(i)) Then
            MsgTDS2 = MsgTDS2 & "* Gross Amount at Sr. No  " & i & "  in schedule TDS2(i) should be Numeric value" & Chr(13)
            ValidateGrossAmount_TDS2 = False
            Exit Function
        End If
    
    If TdsClaimed_TDS2_1(i) <> "" Then
    If isdropdownblank(Headincome_TDS2_1(i)) And rangecells.item(i).Locked = False Then
    MsgTDS2 = MsgTDS2 + "* Please select the drop down of head of income for which corresponding income offered at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
                ValidateGrossAmount_TDS2 = False
                Exit Function
    End If
    End If
    
 UpdateProgressBar
Next
End Function


Function checkforNonMandatoryCol_TDS2() As Boolean
checkforNonMandatoryCol_TDS2 = True
    Dim temp As Double
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.UTN").count
    Set rangecells = Range("TDSoth.UTN").Cells
    
    temp = end_TDS2
    
    For mIntCtr = (temp + 1) To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                MsgTDS2 = MsgTDS2 & "* At Sr. No " & mIntCtr & " Please fill all the Mandatory Fields " & Chr(13)
                checkforNonMandatoryCol_TDS2 = False
                Exit Function
           End If
    Next
End Function
Function checkforNonMandatoryCol_TDS2ii() As Boolean
checkforNonMandatoryCol_TDS2ii = True
    Dim temp As Double
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth2.UTN").count
    Set rangecells = Range("TDSoth2.UTN").Cells
    
    temp = end_TDS2ii
    
    For mIntCtr = (temp + 1) To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                MsgTDS2ii = MsgTDS2ii & "* At Sr. No " & mIntCtr & " Please fill all the Mandatory Fields " & Chr(13)
                checkforNonMandatoryCol_TDS2ii = False
                Exit Function
           End If
    Next
End Function

Sub setTblinfo_TDS2()
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
 end_TDS2 = ccount
 rngname_TDS2 = "TDsOthr.TAN||TDsOthr.finyear||TDsOthr.tds||TDsOthr.tdsdeducted||TDsOthr.grossamount||TDsOthr.headincome||TDsOthr.tdscarried"
 End Sub
Sub setTblinfo_TDS2ii()
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
 end_TDS2ii = ccount
 rngname_TDS2ii = "TDsOthr2.PAN||TDsOthr2.finyear||TDsOthr2.tds||TDsOthr2.tdsdeducted||TDsOthr2.tdsclaimed ||TDsOthr2.grossamount ||TDsOthr2.headamount||TDsOthr2.carried"
 End Sub
Function ValidateUTN_TDS2() As Boolean
 
    ValidateUTN_TDS2 = True
    setTblinfo_TDS2
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet4.Range("TDSoth.UTN").Cells
    ReDim UTN_TDS2(end_TDS2)
    For i = 1 To end_TDS2
        UTN_TDS2(i) = rangecells.item(i).value
    If Not checkfieldspecialcharacter(UTN_TDS2(i)) Then
         MsgTDS2 = MsgTDS2 + "* Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
          ValidateUTN_TDS2 = False
          Exit Function
     End If
 
If InStr(UTN_TDS2(i), ".") > 0 Then
         MsgTDS2 = MsgTDS2 + "* Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
          ValidateUTN_TDS2 = False
          Exit Function
    End If
    
    If InStr(UTN_TDS2(i), "-") > 0 Then
         MsgTDS2 = MsgTDS2 + "* Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
          ValidateUTN_TDS2 = False
          Exit Function
    End If
    
    If UTN_TDS2(i) <> "" Then
    If Len(UTN_TDS2(i)) < 6 Or Len(UTN_TDS2(i)) > 8 Then
          MsgTDS2 = MsgTDS2 + "* Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
          ValidateUTN_TDS2 = False
          Exit Function
    End If
    End If
    
 
 Next
End Function
Function ValidateUTN_TDS2ii() As Boolean
 
    ValidateUTN_TDS2ii = True
    setTblinfo_TDS2ii
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet4.Range("TDSoth2.UTN").Cells
    ReDim UTN_TDS2ii(end_TDS2ii)
    For i = 1 To end_TDS2ii
        UTN_TDS2ii(i) = rangecells.item(i).value
    If Not checkfieldspecialcharacter(UTN_TDS2ii(i)) Then
         MsgTDS2ii = MsgTDS2ii + "* Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid"
          ValidateUTN_TDS2ii = False
          Exit Function
     End If
 
If InStr(UTN_TDS2ii(i), ".") > 0 Then
         MsgTDS2ii = MsgTDS2ii + "* Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
          ValidateUTN_TDS2ii = False
          Exit Function
    End If
    
    If InStr(UTN_TDS2ii(i), "-") > 0 Then
         MsgTDS2ii = MsgTDS2ii + "* Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
          ValidateUTN_TDS2ii = False
          Exit Function
    End If
    
    If UTN_TDS2ii(i) <> "" Then
    If Len(UTN_TDS2ii(i)) < 6 Or Len(UTN_TDS2ii(i)) > 8 Then
          MsgTDS2ii = MsgTDS2ii + "* Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
          ValidateUTN_TDS2ii = False
          Exit Function
    End If
    End If
    
 
 Next
End Function



Function ValidateEmployerOrDeductorOrCollecterName_TDS2() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TDS2 = True
    setTblinfo_TDS2
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet4.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
    ReDim EmployerOrDeductorOrCollecterName_TDS2(end_TDS2)
    For i = 1 To end_TDS2
        EmployerOrDeductorOrCollecterName_TDS2(i) = rangecells.item(i).value

        If EmployerOrDeductorOrCollecterName_TDS2(i) = "" Then
            MsgTDS2 = MsgTDS2 + "* Name of deductor at Sr. No  " & i & "  is Mandatory" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TDS2 = False
            Exit Function
        End If
            
            
            If Len(EmployerOrDeductorOrCollecterName_TDS2(i)) > 125 Then
            MsgTDS2 = MsgTDS2 + "* Name of deductor at Sr. No  " & i & "  in Sheet TDS  Cannot be more than 125 Charachters" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TDS2 = False
            Exit Function
        End If

    Next
End Function
Function ValidateEmployerOrDeductorOrCollecterName_TDS2ii() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TDS2ii = True
    setTblinfo_TDS2ii
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet4.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells
    ReDim EmployerOrDeductorOrCollecterName_TDS2ii(end_TDS2ii)
    For i = 1 To end_TDS2ii
        EmployerOrDeductorOrCollecterName_TDS2ii(i) = rangecells.item(i).value

        If EmployerOrDeductorOrCollecterName_TDS2ii(i) = "" Then
            MsgTDS2ii = MsgTDS2ii + "* Name of the tenant at Sr. No  " & i & "  is Mandatory" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TDS2ii = False
            Exit Function
        End If
            
            
            If Len(EmployerOrDeductorOrCollecterName_TDS2ii(i)) > 125 Then
            MsgTDS2ii = MsgTDS2ii + "* Name of the tenant at Sr. No  " & i & "  in Sheet TDS  Cannot be more than 125 Charachters" & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TDS2ii = False
            Exit Function
        End If

    Next
End Function

Function ValidateClaimOutOfTotTDSOnAmtPaid_TDS2() As Boolean
    ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = True
    setTblinfo_TDS2
    setTableinfo_TDS_SIncome_TDS2
    setTableinfo_TDsinc_TaxTDS1
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim rangecells7 As Range
    Dim rangecells8 As Range
    Dim rangecells9 As Range
    
    
    Dim i As Long
    Dim portugesecode As Variant
    
    Set rangecells6 = Sheet4.Range("TDSoth.AmountPaid").Cells
    Set rangecells1 = Sheet4.Range("TDSoth.TD_OwnHands").Cells
    Set rangecells5 = Sheet4.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
    

    
    ReDim AmtPaid_TDS2(end_TDS2)
    ReDim TD_OwnHands_TDS2(end_TDS2)
    ReDim AmtClaimedOnOwnHands_TDS2(end_TDS2)
    
    ReDim TD_SIncome_TDS2(end_TDS2)
    ReDim TD_TDS_TDS2(end_TDS2)
    
    ReDim TDS_SIncome_TDS2(end_TDS2)
    ReDim TDS_TDS_TDS2(end_TDS2)
    
    ReDim TotalSum(end_TDS2)
    ReDim TotalSum1(end_TDS2)
    
For i = 1 To end_TDS2

    AmtPaid_TDS2(i) = rangecells6.item(i).value
    AmtClaimedOnOwnHands_TDS2(i) = rangecells5.item(i).value
    TD_OwnHands_TDS2(i) = rangecells1.item(i).value
    

    If AmtClaimedOnOwnHands_TDS2(i) = "" And Sheet4.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells.item(i).Locked = False Then
        MsgTDS2 = MsgTDS2 + "* Amount claimed this year at Sr.No " & i & " is Mandatory" & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
    End If

    If AmtPaid_TDS2(i) = "" And Sheet4.Range("TDSoth.AmountPaid").Cells.item(i).Locked = False Then
        MsgTDS2 = MsgTDS2 + "* Details of Receipts as mentioned in 26AS at Sr.No " & i & " is Mandatory" & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
    End If
    
     If TD_OwnHands_TDS2(i) = "" And Sheet4.Range("TDSoth.TD_OwnHands").Cells.item(i).Locked = False Then
        MsgTDS2 = MsgTDS2 + "* Please enter Tax deducted in own hands at Sr.No " & i & " is Mandatory" & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
    End If
    

    
    If IsNumeric(AmtPaid_TDS2(i)) Then
        If AmtPaid_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 5 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & AmtPaid_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 5 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 5 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If
            
    If IsNumeric(AmtClaimedOnOwnHands_TDS2(i)) Then
        If AmtClaimedOnOwnHands_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & AmtClaimedOnOwnHands_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If
    
    
    
    If IsNumeric(TD_OwnHands_TDS2(i)) Then
        If TD_OwnHands_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 6 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & TD_OwnHands_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 6 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 6 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If
     

        
   If (AmtClaimedOnOwnHands_TDS2(i) = "") Then
        AmtClaimedOnOwnHands_TDS2(i) = 0
    End If


Next

For i = 1 To end_TDS2

  Set rangecells2 = Sheet4.Range("TDSoth.TD_SIncome").Cells
    Set rangecells7 = Sheet4.Range("TDSoth.TD_TDS").Cells
    
    TD_SIncome_TDS2(i) = rangecells2.item(i).value
    TD_TDS_TDS2(i) = rangecells7.item(i).value


    If IsNumeric(TD_TDS_TDS2(i)) Then
        If TD_TDS_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & TD_TDS_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 7 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If

    If IsNumeric(TD_SIncome_TDS2(i)) Then
        If TD_SIncome_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & TD_SIncome_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 7 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If



Next

For i = 1 To end_TDS2

    Set rangecells8 = Sheet4.Range("TDSoth.TDS_SIncome").Cells
    Set rangecells9 = Sheet4.Range("TDSoth.TDS_TDS").Cells

    TDS_SIncome_TDS2(i) = rangecells8.item(i).value
    TDS_TDS_TDS2(i) = rangecells9.item(i).value

        

    If IsNumeric(TDS_TDS_TDS2(i)) Then
        If TDS_TDS_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & TDS_TDS_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If
    
        If IsNumeric(TDS_SIncome_TDS2(i)) Then
        If TDS_SIncome_TDS2(i) < 0 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        End If
        If Len("" & TDS_SIncome_TDS2(i)) > 14 Then
            MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
            Exit Function
        End If
    Else
        MsgTDS2 = MsgTDS2 + "* Amount in col.no 8 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
        Exit Function
    End If

Next

End Function

Function ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii() As Boolean
    ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = True
    setTblinfo_TDS2ii
    setTableinfo_TDsinc_TaxTDS2ii
    setTableinfo_TDS_SIncome_TDS2ii
    
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim rangecells7 As Range
    Dim rangecells8 As Range
    Dim rangecells9 As Range
    
    Dim i As Long
    Dim portugesecode As Variant
    
    Set rangecells6 = Sheet4.Range("TDSoth2.AmountSubjectDed").Cells
    Set rangecells1 = Sheet4.Range("TDSoth2.TD_OwnHands").Cells
    Set rangecells5 = Sheet4.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells
    
    
    ReDim AmtPaid_TDS2ii(end_TDS2ii)
    ReDim TD_OwnHands_TDS2ii(end_TDS2ii)
    ReDim AmtClaimedOnOwnHands_TDS2ii(end_TDS2ii)
    
    ReDim TD_SIncome_TDS2ii(end_TDS2ii)
    ReDim TD_TDS_TDS2ii(end_TDS2ii)
    
    ReDim TDS_SIncome_TDS2ii(end_TDS2ii)
    ReDim TDS_TDS_TDS2ii(end_TDS2ii)
    
    ReDim TotalSum2ii(end_TDS2ii)
    ReDim TotalSum2iii(end_TDS2ii)
    
For i = 1 To end_TDS2ii

    AmtPaid_TDS2ii(i) = rangecells6.item(i).value
    AmtClaimedOnOwnHands_TDS2ii(i) = rangecells5.item(i).value
    TD_OwnHands_TDS2ii(i) = rangecells1.item(i).value
    
   
    If AmtClaimedOnOwnHands_TDS2ii(i) = "" And Sheet4.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells.item(i).Locked = False Then
        MsgTDS2ii = MsgTDS2ii + "* Amount claimed this year at Sr.No " & i & " is Mandatory" & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
    End If

    If AmtPaid_TDS2ii(i) = "" And Sheet4.Range("TDSoth2.AmountSubjectDed").Cells.item(i).Locked = False Then
        MsgTDS2ii = MsgTDS2ii + "* Details of Receipts as mentioned in 26AS at Sr.No " & i & " is Mandatory" & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
    End If

     If TD_OwnHands_TDS2ii(i) = "" And Sheet4.Range("TDSoth2.TD_OwnHands").Cells.item(i).Locked = False Then
        MsgTDS2ii = MsgTDS2ii + "* Please enter Tax deducted in own hands at Sr.No " & i & " is Mandatory" & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
    End If


    
    If IsNumeric(AmtPaid_TDS2ii(i)) Then
        If AmtPaid_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 5 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & AmtPaid_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 5 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 5 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If
            
    If IsNumeric(AmtClaimedOnOwnHands_TDS2ii(i)) Then
        If AmtClaimedOnOwnHands_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & AmtClaimedOnOwnHands_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If

    
    If IsNumeric(TD_OwnHands_TDS2ii(i)) Then
        If TD_OwnHands_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 6 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & TD_OwnHands_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 6 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 6 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If
        
   If (AmtClaimedOnOwnHands_TDS2ii(i) = "") Then
        AmtClaimedOnOwnHands_TDS2ii(i) = 0
    End If

Next



For i = 1 To end_TDS2ii

    Set rangecells2 = Sheet4.Range("TDSoth2.TD_SIncome").Cells
    Set rangecells7 = Sheet4.Range("TDSoth2.TD_TDS").Cells

    TD_SIncome_TDS2ii(i) = rangecells2.item(i).value
    TD_TDS_TDS2ii(i) = rangecells7.item(i).value
         
    

   
    
        
    If IsNumeric(TD_SIncome_TDS2ii(i)) Then
        If TD_SIncome_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & TD_SIncome_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 7 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If
        
    If IsNumeric(TD_TDS_TDS2ii(i)) Then
        If TD_TDS_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & TD_TDS_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 7 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If
Next



For i = 1 To end_TDS2ii

    Set rangecells8 = Sheet4.Range("TDSoth2.TDS_SIncome").Cells
    Set rangecells9 = Sheet4.Range("TDSoth2.TDS_TDS").Cells

    TDS_SIncome_TDS2ii(i) = rangecells8.item(i).value
    TDS_TDS_TDS2ii(i) = rangecells9.item(i).value

 

   


If IsNumeric(TDS_SIncome_TDS2ii(i)) Then
        If TDS_SIncome_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & TDS_SIncome_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If

If IsNumeric(TDS_TDS_TDS2ii(i)) Then
        If TDS_TDS_TDS2ii(i) < 0 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        End If
        If Len("" & TDS_TDS_TDS2ii(i)) > 14 Then
            MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
            Exit Function
        End If
    Else
        MsgTDS2ii = MsgTDS2ii + "* Amount in col.no 8 is invalid in Row  " & i & Chr(13)
        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
        Exit Function
    End If

Next


End Function

Function ValidatePAN7_TDS() As Boolean
ValidatePAN7_TDS = True
setTableinfo_TDS_SIncome_TDS2
setTableinfo_TDsinc_TaxTDS1
setTblinfo_TDS2
noOfProcessSub = end_TDS2

Dim rangecells As Range
Dim rangecells1 As Range

Set rangecells = Range("TDSoth.TD_PAN").Cells
Set rangecells1 = Range("TDSoth.TDS_PAN").Cells


Dim i As Long
ReDim PAN7_TDS(end_TDS2)
ReDim PAN8_TDS(end_TDS2)
For i = 1 To end_TDS2
    PAN7_TDS(i) = rangecells.item(i).value
    
    If (Not Len(PAN7_TDS(i)) = 0) Then
        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN7_TDS(i), 1, 10))) Then
        MsgTDS2 = MsgTDS2 + "* PAN at Sr. No " & i & "  in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
        ValidatePAN7_TDS = False
        End If
        End If

    
    If PAN7_TDS(i) = Sheet1.Range("sheet1.PAN").value And Range("TDSoth.TD_PAN").Cells.item(i).Locked = False Then
                MsgTDS2 = MsgTDS2 + "* PAN of spouse / other person at Sr.No " & i & " can not the same as PAN of user  " & Chr(13)
                ValidatePAN7_TDS = False
                Exit Function
    End If
    
    
Next
    
For i = 1 To end_TDS2
PAN8_TDS(i) = rangecells1.item(i).value

    If (Not Len(PAN8_TDS(i)) = 0) Then
        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN8_TDS(i), 1, 10))) Then
        MsgTDS2 = MsgTDS2 + "* PAN at Sr. No  " & i & " in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
        ValidatePAN7_TDS = False
        End If
        End If
    
        If ((PAN8_TDS(i) = Sheet1.Range("sheet1.PAN").value) And Sheet1.Range("sheet1.PAN").value <> "") Then
                MsgTDS2 = MsgTDS2 + "* PAN of spouse/other person at Sr.No " & i & " can not the same as PAN of user  " & Chr(13)
                ValidatePAN7_TDS = False
                Exit Function
    End If
    
    
    UpdateProgressBar
    Next
End Function

Function ValidatePAN7_TDSii() As Boolean
ValidatePAN7_TDSii = True
setTblinfo_TDS2
setTableinfo_TDsinc_TaxTDS2ii
setTableinfo_TDS_SIncome_TDS2ii
noOfProcessSub = end_TDS2ii

Dim rangecells As Range
Dim rangecells1 As Range
Set rangecells = Range("TDSoth2.TD_PAN").Cells
Set rangecells1 = Range("TDSoth2.TDS_PAN").Cells


Dim i As Long
ReDim PAN7_TDSii(end_TDS2ii)
ReDim PAN8_TDSii(end_TDS2ii)

For i = 1 To end_TDS2ii
    PAN7_TDSii(i) = rangecells.item(i).value
    
    If (Not Len(PAN7_TDSii(i)) = 0) And Range("TDSoth2.TD_PAN").Cells.item(i).Locked = False Then
        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN7_TDSii(i), 1, 10))) Then
        MsgTDS2ii = MsgTDS2ii + "* PAN  at Sr. No " & i & " in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet" & Chr(13)
        ValidatePAN7_TDSii = False
        End If
        End If

    
    If ((PAN7_TDSii(i) = Sheet1.Range("sheet1.PAN").value) And Sheet1.Range("sheet1.PAN").value <> "") Then
                MsgTDS2ii = MsgTDS2ii + "* PAN of spouse/other person at Sr.No " & i & " can not the same as PAN of user  " & Chr(13)
                ValidatePAN7_TDSii = False
                Exit Function
    End If
    Next
    
For i = 1 To end_TDS2ii

    PAN8_TDSii(i) = rangecells1.item(i).value

    If (Not Len(PAN8_TDSii(i)) = 0) And Range("TDSoth2.TDS_PAN").Cells.item(i).Locked = False Then
        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN8_TDSii(i), 1, 10))) Then
        MsgTDS2ii = MsgTDS2ii + "* PAN  at Sr. No " & i & " in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
        ValidatePAN7_TDSii = False
        End If
        End If


    
        If (PAN8_TDSii(i) = Sheet1.Range("sheet1.PAN").value And Sheet1.Range("sheet1.PAN").value <> "") Then
                MsgTDS2ii = MsgTDS2ii + "* PAN of spouse/other person at Sr.No " & i & " can not the same as PAN of user  " & Chr(13)
                ValidatePAN7_TDSii = False
                Exit Function
      End If
    
    
    UpdateProgressBar
    Next
End Function


Function ValidateDeductedYear_TDS2() As Boolean
 
    ValidateDeductedYear_TDS2 = True
    setTblinfo_TDS2
    Dim i As Long
    Dim rangecells, rangecells1 As Range
    Set rangecells = Sheet4.Range("TDSoth.DeductedYear").Cells
    Set rangecells1 = Sheet4.Range("TDS2.TdsCredit").Cells
    ReDim DeductedYear_TDS2(end_TDS2)
    ReDim TDS2Credit(end_TDS2)
For i = 1 To end_TDS2
        DeductedYear_TDS2(i) = rangecells.item(i).value
        TDS2Credit(i) = rangecells1.item(i).value
        
        If isdropdownblank(DeductedYear_TDS2(i)) Then
            MsgTDS2 = MsgTDS2 + "* Please select Year of tax deduction at Sr. No " & i & " from drop-down" & Chr(13)
            ValidateDeductedYear_TDS2 = False
            Exit Function
        End If
        
       
    
    
     If Not checkfieldspecialcharacter(DeductedYear_TDS2(i)) Then
         MsgTDS2 = MsgTDS2 + "* Deducted Year at Sr. No " & i & " in Sheet TDS schedule TDS2 is invalid" & Chr(13)
          ValidateDeductedYear_TDS2 = False
          Exit Function
     End If
     
     
     If DeductedYear_TDS2(i) > 2022 Then
          MsgTDS2 = MsgTDS2 + "* Deducted Year at Sr. No " & i & " in Sheet TDS cannot be greater than 2022 " & Chr(13)
          ValidateDeductedYear_TDS2 = False
          Exit Function
    End If
Next
End Function

Function ValidateDeductedYear_TDS2ii() As Boolean
 
    ValidateDeductedYear_TDS2ii = True
    setTblinfo_TDS2ii
    Dim i As Long
    Dim rangecells, rangecells1 As Range
    Set rangecells = Sheet4.Range("TDSoth2.DeductedYear").Cells
    Set rangecells1 = Sheet4.Range("TDS3.TdsCredit").Cells
    ReDim DeductedYear_TDS2ii(end_TDS2ii)
    ReDim TDS3Credit(end_TDS2ii)
For i = 1 To end_TDS2ii
        DeductedYear_TDS2ii(i) = rangecells.item(i).value
        TDS3Credit(i) = rangecells1.item(i).value
        
        If isdropdownblank(DeductedYear_TDS2ii(i)) Then
            MsgTDS2ii = MsgTDS2ii + "* Please select Year of tax deduction at Sr. No " & i & " from drop-down" & Chr(13)
            ValidateDeductedYear_TDS2ii = False
            Exit Function
        End If

    
    
     If Not checkfieldspecialcharacter(DeductedYear_TDS2ii(i)) Then
         MsgTDS2ii = MsgTDS2ii + "* Deducted Year at Sr. No " & i & " in Sheet TDS schedule TDS2 is invalid" & Chr(13)
          ValidateDeductedYear_TDS2ii = False
          Exit Function
          End If

     
     If DeductedYear_TDS2ii(i) > 2022 And UCase(DeductedYear_TDS2ii(i)) <> "(SELECT)" Then
          MsgTDS2ii = MsgTDS2ii + "* Deducted Year at Sr. No " & i & " in Sheet TDS cannot be greater than 2022 " & Chr(13)
          ValidateDeductedYear_TDS2ii = False
          Exit Function
    End If
Next
End Function

Sub setTableinfo_Name_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.EmployerOrDeductorOrCollecterName").count
 Set rangecells = Sheet4.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Name_TDS2 = ccount
End Sub

Sub setTableinfo_TDSCredit_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDS2.TdsCredit").count
 Set rangecells = Sheet4.Range("TDS2.TdsCredit").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDSCredit_TDS2 = ccount
End Sub
Sub setTableinfo_TDSCredit2_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDS2.TdsCredit").count
 Set rangecells = Sheet4.Range("TDS2.TdsCredit").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_TDSCredit2_TDS2 = ccount
End Sub

Sub setTableinfo_TDSCredit_TDS3()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDS3.TdsCredit").count
 Set rangecells = Sheet4.Range("TDS3.TdsCredit").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDSCredit_TDS3 = ccount
End Sub
Sub setTableinfo_TDSCredit3_TDS3()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDS3.TdsCredit").count
 Set rangecells = Sheet4.Range("TDS3.TdsCredit").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_TDSCredit3_TDS3 = ccount
End Sub
Sub setTableinfo_Year_TDS1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.DeductedYear").count
 Set rangecells = Sheet4.Range("TDSoth.DeductedYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_Year_TDS2 = ccount
End Sub
Sub setTableinfo_Year_TDS1i()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.DeductedYear").count
 Set rangecells = Sheet4.Range("TDSoth.DeductedYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "") Then ' Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_Yeari_TDS2 = ccount
End Sub

Sub setTableinfo_AmtPaidTDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.AmountPaid").count
 Set rangecells = Sheet4.Range("TDSoth.AmountPaid").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtPaid_TDS2 = ccount
End Sub
'
Sub setTableinfo_TD_OwnHand()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TD_OwnHands").count
 Set rangecells = Sheet4.Range("TDSoth.TD_OwnHands").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDi_TDS2 = ccount
End Sub


Sub setTableinfo_Amt_TaxTDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").count
 Set rangecells = Sheet4.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Amt_TDS2 = ccount
End Sub


Sub setTableinfo_TDsinc_TaxTDS1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TD_SIncome").count
 Set rangecells = Sheet4.Range("TDSoth.TD_SIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDii_TDS2 = ccount
End Sub

Sub setTableinfo_TD_TDS_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TD_TDS").count
 Set rangecells = Sheet4.Range("TDSoth.TD_TDS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDiii_TDS2 = ccount
End Sub


Sub setTableinfo_TD_PAN_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TD_PAN").count
 Set rangecells = Sheet4.Range("TDSoth.TD_PAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_TD_PAN_TDS2 = ccount
End Sub

Sub setTableinfo_TDS_SIncome_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TDS_SIncome").count
 Set rangecells = Sheet4.Range("TDSoth.TDS_SIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS_SIncome_TDS2 = ccount
End Sub


Sub setTableinfo_TDS_TDS_TDS2()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TDS_TDS").count
 Set rangecells = Sheet4.Range("TDSoth.TDS_TDS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS_TDS_TDS2 = ccount
End Sub


Sub setTableinfo_TDS_PAN_TaxTDS1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth.TDS_PAN").count
 Set rangecells = Sheet4.Range("TDSoth.TDS_PAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS_PAN_TDS2 = ccount
End Sub


Sub setTableinfo_Name_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.EmployerOrDeductorOrCollecterName").count
 Set rangecells = Sheet4.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Name_TDS2ii = ccount
End Sub


Sub setTableinfo_Year_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.DeductedYear").count
 Set rangecells = Sheet4.Range("TDSoth2.DeductedYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_Year_TDS2ii = ccount
End Sub
Sub setTableinfo_Yearii_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.DeductedYear").count
 Set rangecells = Sheet4.Range("TDSoth2.DeductedYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_Yearii_TDS2ii = ccount
End Sub
Sub setTableinfo_AmtPaidTDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.AmountSubjectDed").count
 Set rangecells = Sheet4.Range("TDSoth2.AmountSubjectDed").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtPaid_TDS2ii = ccount
End Sub
'
Sub setTableinfo_TD_OwnHand2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TD_OwnHands").count
 Set rangecells = Sheet4.Range("TDSoth2.TD_OwnHands").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDi_TDS2ii = ccount
End Sub


Sub setTableinfo_Amt_TaxTDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").count
 Set rangecells = Sheet4.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Amt_TDS2ii = ccount
End Sub


Sub setTableinfo_TDsinc_TaxTDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TD_SIncome").count
 Set rangecells = Sheet4.Range("TDSoth2.TD_SIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDii_TDS2ii = ccount
End Sub

Sub setTableinfo_TD_TDS_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TD_TDS").count
 Set rangecells = Sheet4.Range("TDSoth2.TD_TDS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDiii_TDS2ii = ccount
End Sub


Sub setTableinfo_TD_PAN_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TD_PAN").count
 Set rangecells = Sheet4.Range("TDSoth2.TD_PAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_TD_PAN_TDS2ii = ccount
End Sub


Sub setTableinfo_TDS_SIncome_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TDS_SIncome").count
 Set rangecells = Sheet4.Range("TDSoth2.TDS_SIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS_SIncome_TDS2ii = ccount
End Sub


Sub setTableinfo_TDS_TDS_TDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TDS_TDS").count
 Set rangecells = Sheet4.Range("TDSoth2.TDS_TDS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS_TDS_TDS2ii = ccount
End Sub


Sub setTableinfo_TDS_PAN_TaxTDS2ii()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("TDSoth2.TDS_PAN").count
 Set rangecells = Sheet4.Range("TDSoth2.TDS_PAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS_PAN_TDS2ii = ccount
End Sub


Function LockingTDS2()
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim rangecells4 As Range
Dim rangecells5 As Range
Dim rangecells6 As Range
Dim rangecells7 As Range
Dim rangecells8 As Range

Set rangecells = Sheet4.Range("TDS2.TdsCredit").Cells
Set rangecells1 = Sheet4.Range("TDSoth.TD_SIncome").Cells
Set rangecells2 = Sheet4.Range("TDSoth.TD_TDS").Cells
Set rangecells3 = Sheet4.Range("TDSoth.TD_PAN").Cells
Set rangecells4 = Sheet4.Range("TDSoth.TD_OwnHands").Cells
Set rangecells5 = Sheet4.Range("TDSoth.TDS_SIncome").Cells
Set rangecells6 = Sheet4.Range("TDSoth.TDS_TDS").Cells
Set rangecells7 = Sheet4.Range("TDSoth.TDS_PAN").Cells
Set rangecells8 = Sheet4.Range("TDSoth.AmountPaid").Cells

Sheet4.Unprotect Password:=getmsgstate

For i = 1 To Sheet4.Range("TDSoth.TAN").count

If Mid(rangecells.item(i).value, 1, 1) = "S" Then
rangecells1.item(i).Locked = True
rangecells2.item(i).Locked = True
rangecells3.item(i).Locked = True

rangecells1.item(i).Interior.Color = (&HD8D8D8)
rangecells2.item(i).Interior.Color = (&HD8D8D8)
rangecells3.item(i).Interior.Color = (&HD8D8D8)

ElseIf Mid(rangecells.item(i).value, 1, 1) = "O" Then
rangecells4.item(i).Locked = True
rangecells5.item(i).Locked = True
rangecells6.item(i).Locked = True
rangecells7.item(i).Locked = True
rangecells8.item(i).Locked = True

rangecells4.item(i).Interior.Color = (&HD8D8D8)
rangecells5.item(i).Interior.Color = (&HD8D8D8)
rangecells6.item(i).Interior.Color = (&HD8D8D8)
rangecells7.item(i).Interior.Color = (&HD8D8D8)
rangecells8.item(i).Interior.Color = (&HD8D8D8)

Else
rangecells1.item(i).Locked = False
rangecells2.item(i).Locked = False
rangecells3.item(i).Locked = False
rangecells4.item(i).Locked = False
rangecells5.item(i).Locked = False
rangecells6.item(i).Locked = False
rangecells7.item(i).Locked = False
rangecells8.item(i).Locked = False

rangecells1.item(i).Interior.Color = (&HCCFFCC)
rangecells2.item(i).Interior.Color = (&HCCFFCC)
rangecells3.item(i).Interior.Color = (&HCCFFCC)
rangecells4.item(i).Interior.Color = (&HCCFFCC)
rangecells5.item(i).Interior.Color = (&HCCFFCC)
rangecells6.item(i).Interior.Color = (&HCCFFCC)
rangecells7.item(i).Interior.Color = (&HCCFFCC)
rangecells8.item(i).Interior.Color = (&HCCFFCC)

End If
    
Next

Sheet4.Protect Password:=getmsgstate

End Function


Function LockingTDS3()
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim rangecells4 As Range
Dim rangecells5 As Range
Dim rangecells6 As Range
Dim rangecells7 As Range
Dim rangecells8 As Range

Set rangecells = Sheet4.Range("TDS3.TdsCredit").Cells
Set rangecells1 = Sheet4.Range("TDSoth2.TD_SIncome").Cells
Set rangecells2 = Sheet4.Range("TDSoth2.TD_TDS").Cells
Set rangecells3 = Sheet4.Range("TDSoth2.TD_PAN").Cells
Set rangecells4 = Sheet4.Range("TDSoth2.AmountSubjectDed").Cells
Set rangecells5 = Sheet4.Range("TDSoth2.TD_OwnHands").Cells
Set rangecells6 = Sheet4.Range("TDSoth2.TDS_SIncome").Cells
Set rangecells7 = Sheet4.Range("TDSoth2.TDS_TDS").Cells
Set rangecells8 = Sheet4.Range("TDSoth2.TDS_PAN").Cells

Sheet4.Unprotect Password:=getmsgstate

For i = 1 To Sheet4.Range("TDSoth2.PAN").count

If Mid(rangecells.item(i).value, 1, 1) = "S" Then
rangecells1.item(i).Locked = True
rangecells2.item(i).Locked = True
rangecells3.item(i).Locked = True

rangecells1.item(i).Interior.Color = (&HD8D8D8)
rangecells2.item(i).Interior.Color = (&HD8D8D8)
rangecells3.item(i).Interior.Color = (&HD8D8D8)

ElseIf Mid(rangecells.item(i).value, 1, 1) = "O" Then
rangecells4.item(i).Locked = True
rangecells5.item(i).Locked = True
rangecells6.item(i).Locked = True
rangecells7.item(i).Locked = True
rangecells8.item(i).Locked = True

rangecells4.item(i).Interior.Color = (&HD8D8D8)
rangecells5.item(i).Interior.Color = (&HD8D8D8)
rangecells6.item(i).Interior.Color = (&HD8D8D8)
rangecells7.item(i).Interior.Color = (&HD8D8D8)
rangecells8.item(i).Interior.Color = (&HD8D8D8)

Else
rangecells1.item(i).Locked = False
rangecells2.item(i).Locked = False
rangecells3.item(i).Locked = False
rangecells4.item(i).Locked = False
rangecells5.item(i).Locked = False
rangecells6.item(i).Locked = False
rangecells7.item(i).Locked = False
rangecells8.item(i).Locked = False

rangecells1.item(i).Interior.Color = (&HCCFFCC)
rangecells2.item(i).Interior.Color = (&HCCFFCC)
rangecells3.item(i).Interior.Color = (&HCCFFCC)
rangecells4.item(i).Interior.Color = (&HCCFFCC)
rangecells5.item(i).Interior.Color = (&HCCFFCC)
rangecells6.item(i).Interior.Color = (&HCCFFCC)
rangecells7.item(i).Interior.Color = (&HCCFFCC)
rangecells8.item(i).Interior.Color = (&HCCFFCC)


End If
    
Next

Sheet4.Protect Password:=getmsgstate

End Function

Function ValidateTANCodes(TAN As String) As Boolean
    ValidateTANCodes = False
    Dim TANCode As Range

    
    Dim TAN1 As Range


''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet2.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
    
    Set TANCode = Sheet8.Range("TANCodes")
  
    
    For Each TAN1 In TANCode
        If Mid(TAN1.value, 1, 3) = Mid(TAN, 1, 3) Then
        ValidateTANCodes = True
        Exit For
        End If
    Next
    
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet2.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''Added by malli
'''TDS_C1 AY_2024_25 change
'
'Function ValidateSectionTDSDeducted_TDS2() As Boolean
'ValidateSectionTDSDeducted_TDS2 = True
'
'noOfProcessSub = end_TDS2_1
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'
'Set rangecells = Range("TDsOthr.SectionTDSDeducted").Cells
''Set rangecells1 = Range("TDsOthr.SectionTDSDeducted").Cells
'Dim i As Long
' ReDim SectionTDSDeducted_TDS2_1(end_TDS2_1)
'
'For i = 1 To end_TDS2_1
'    SectionTDSDeducted_TDS2_1(i) = rangecells.item(i).value
'
'
'    If Not chkCompulsory(SectionTDSDeducted_TDS2_1(i)) Then
'                MsgTDS2 = MsgTDS2 + "* Please select the Section under which TDS is deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'                ValidateSectionTDSDeducted_TDS2 = False
'                Exit Function
'    End If
'
'    If (SectionTDSDeducted_TDS2_1(i) = "" Or SectionTDSDeducted_TDS2_1(i) = "(Select)") Then
'                MsgTDS2 = MsgTDS2 + "* Please select the Section under which TDS is deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
'                ValidateSectionTDSDeducted_TDS2 = False
'                Exit Function
'    End If
'
'
'         UpdateProgressBar
'Next
'End Function

'''''''''''''''''''''''''''Added by malli
'
'''TDS_C1 AY_2024_25 change
'
'Sub setTableinfo_TDS2_SectionTDSDeducted()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet4.Range("TDsOthr.SectionTDSDeducted").count
' Set rangecells = Sheet4.Range("TDsOthr.SectionTDSDeducted").Cells
' For mIntCtr = 1 To mIntCells
'     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2_9 = ccount
'End Sub
'
' '''''''''''''''''''''''''''''''''Added by Malli
' '''''TDS_C4 AY_2024_25 change
'Sub setTableinfo_TDS2ii_SectionTDSDeducted()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet4.Range("TDsOthr2.SectionTDSDeducted").count
' Set rangecells = Sheet4.Range("TDsOthr2.SectionTDSDeducted").Cells
' For mIntCtr = 1 To mIntCells
'     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2ii_9 = ccount
'End Sub
'
''''''''''''''''''''''''''''''''''''''''Added by malli
''''''TDS_C4 AY_2024_25 change
'Function ValidateSectionTDSDeducted_TDS2ii() As Boolean
'ValidateSectionTDSDeducted_TDS2ii = True
'
'noOfProcessSub = end_TDS2ii_1
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'
'Set rangecells = Range("TDsOthr2.SectionTDSDeducted").Cells
'
'Dim i As Long
' ReDim SectionTDSDeducted_TDS2ii_1(end_TDS2ii_1)
'
'
'For i = 1 To end_TDS2ii_1
'    SectionTDSDeducted_TDS2ii_1(i) = rangecells.item(i).value
'
'
'
'    If Not chkCompulsory(SectionTDSDeducted_TDS2ii_1(i)) Then
'                MsgTDS2ii = MsgTDS2ii + "* Please select the Section under which TDS is deducted  at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'                ValidateSectionTDSDeducted_TDS2ii = False
'                Exit Function
'    End If
'
'    If (SectionTDSDeducted_TDS2ii_1(i) = "" Or SectionTDSDeducted_TDS2ii_1(i) = "(Select)") Then
'                MsgTDS2ii = MsgTDS2ii + "* Please select the Section under which TDS is deducted  at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
'                ValidateSectionTDSDeducted_TDS2ii = False
'                Exit Function
'    End If
'
'
'    UpdateProgressBar
'Next
'End Function
'
''''''''''''Added by Malli
''''''TCS_C1 AY_2024_25 change
'Sub setTableInfo_SectionTCSDeducted_TCS()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet17.Range("TCS.SectionTCSDeducted").count
' Set rangecells = Sheet17.Range("TCS.SectionTCSDeducted").Cells
' For mIntCtr = 1 To mIntCells
'     If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
'         ccount = ccount + 1
'     End If
' Next
' end_SectionTCSDeducted_TCS = ccount
'End Sub
'
'
'''''''''''Added by Malli
'
''''''TCS_C1 AY_2024_25 change
'Function ValidateSectionTCSDeducted_TCS() As Boolean
'
'    ValidateSectionTCSDeducted_TCS = True
'    'setTblinfo_TCS
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet17.Range("TCS.SectionTCSDeducted").Cells
'    ReDim SectionTCSDeducted_TCS(end_TCS)
'
'For i = 1 To end_TCS
'        SectionTCSDeducted_TCS(i) = rangecells.item(i).value
'        If Not chkCompulsory(SectionTCSDeducted_TCS(i)) Then
'             MsgTCS = MsgTCS + "* Please select the Section under which TCS is deducted at Sr. No " & i & "  in Sheet TCS" & Chr(13)
'            ValidateSectionTCSDeducted_TCS = False
'            Exit Function
'        End If
'
'      If (SectionTCSDeducted_TCS(i) = "" Or SectionTCSDeducted_TCS(i) = "(Select)") Then
'                MsgTCS = MsgTCS + "* Please select the Section under which TCS is deducted at Sr. No " & i & "  in Sheet TCS" & Chr(13)
'                ValidateSectionTCSDeducted_TCS = False
'                Exit Function
'    End If
'
'Next
'End Function




