Attribute VB_Name = "md80_"
Option Explicit

Public IA80_DeductProfUs80_IA_4_I As Variant
Public IA80_DeductProfUs80_IA_4_Ii As Variant
Public IA80_DeductProfUs80_IA_4_iii As Variant
Public IA80_DeductProfUs80_IA_4_iv As Variant
Public IA80_TotSchedule80_IA As Variant

Public IB80_DeductJKLocUs80_IB_4 As Variant
Public IB80_DeductBackStatesUs80_IB_4 As Variant
Public IB80_DeductBackDisttUs80_IB_5 As Variant
Public IB80_DeductMultiplexUs80_IB_7A As Variant
Public IB80_DeductConvCentUs80_IB_7B As Variant
Public IB80_DeductMinOilUs80_IB_9 As Variant
Public IB80_DeductHousUs80_IB_10 As Variant
Public IB80_DeductColdChainUs80_IB_11 As Variant
Public IB80_DeductFruitVegUs80_IB_11A As Variant
Public IB80_DeductFoodGrainUs80_IB_11A As Variant
Public IB80_DeductRurHospUs80_IB_11B As Variant
Public IB80_DeductHospAnyAreaUs80IB_11C As Variant
Public IB80_TotSchedule80_IB As Variant

Public IC80_DeductInSikkim As Variant
Public IC80_DeductInHimachalP As Variant
Public IC80_DeductInUttaranchal As Variant
Public IC80_Assam As Variant
Public IC80_ArunachalPradesh As Variant
Public IC80_Manipur As Variant
Public IC80_Mizoram As Variant
Public IC80_Meghalaya As Variant
Public IC80_Nagaland As Variant
Public IC80_Tripura As Variant
Public IC80_TotDeductInNorthEast As Variant
Public IC80_TotSchedule80_IC As Variant

Public msgValidateSheetIA80, msgValidateSheetIB80, msgValidateSheetIC80 As Variant
'schedule via

Public scviaSection80C As Variant
Public scviaSection80CCC As Variant
Public scviaSection80CCD_SE As Variant
Public scviaSection80CCD1B_SE As Variant
Public scviaSection80CCD As Variant
Public scviaSection80CCG As Variant
Public scviaSection80D, scviaSection80DB, scviaSection80DC As Variant
Public SELECT80D, SELECT80DB, SELECT80DC, SELECT80DD, SELECT80DDB, SELECT80U As Variant
Public SYSTEM80D As Variant
Public scviaSection80DD As Variant
Public scviaSection80DDB As Variant
Public scviaSection80E As Variant
Public scviaSection80EE As Variant
Public scviaSection80G As Variant
Public scviaSection80GG As Variant
Public scviaSection80GGC As Variant
Public scviaTotPartBchapterVIA As Variant
Public scviaSection80IA As Variant
Public scviaSection80IAB As Variant
Public scviaSection80IB As Variant
Public scviaSection80IBA As Variant
Public scviaSection80IC As Variant
'Public scviaSection80ID As Variant
Public scviaSection80JJA As Variant
Public scviaSection80JJAA As Variant
Public scviaSection80QQB As Variant
Public scviaSection80RRB As Variant
Public scviaTotPartCchapterVIA As Variant
Public scviaSection80TTA, scviaSection80TTB As Variant
Public scviaSection80U As Variant
Public scviaTotPartCAandDchapterVIA As Variant
Public scviaUsrTotalChapVIADeductions As Variant

Public scviaSection80C_Calc As Variant
Public scviaSection80CCC_Calc As Variant
Public scviaSection80CCD_SE_Calc As Variant
Public scviaSection80CCD1B_SE_Calc  As Variant
Public scviaSection80CCD_Calc As Variant
Public scviaSection80CCG_Calc As Variant
Public scviaSection80D_Calc, scviaSection80DB_Calc, scviaSection80DC_Calc As Variant
Public scviaSection80DD_Calc As Variant
Public scviaSection80DDB_Calc As Variant
Public scviaSection80E_Calc As Variant
Public scviaSection80EE_Calc As Variant
Public scviaSection80G_Calc As Variant
Public scviaSection80GG_Calc As Variant
Public scviaSection80GGC_Calc As Variant
Public scviaTotPartBchapterVIA_Calc As Variant
Public scviaSection80IA_Calc As Variant
Public scviaSection80IAB_Calc As Variant
Public scviaSection80IB_Calc As Variant
Public scviaSection80IBA_Calc As Variant
Public scviaSection80IC_Calc As Variant
'Public scviaSection80ID_Calc As Variant
Public scviaSection80JJA_Calc As Variant
Public scviaSection80JJAA_Calc As Variant
Public scviaSection80QQB_Calc As Variant
Public scviaSection80RRB_Calc As Variant
Public scviaTotPartCchapterVIA_Calc As Variant
Public scviaSection80TTA_Calc, scviaSection80TTB_Calc As Variant
Public scviaSection80U_Calc As Variant
Public scviaTotPartCAandDchapterVIA_Calc As Variant
Public scviaTotalChapVIADeductions_Calc As Variant

'***************************************
'by sadineni on 31/03/2026
Public end_80CCC As Variant
Public end_80CCCType As Variant
Public end_80CCCName As Variant
Public end_80CCCAmount As Variant
Public Type_80CCC As Variant
Public Name_80CCC As Variant
Public Amount_80CCC As Variant

Public end_80CCD_1 As Variant
Public end_80CCD_1_Type As Variant
Public end_80CCD_1_Name As Variant
Public end_80CCD_1_Amount As Variant
Public Type_1_80CCD As Variant
Public Name_1_80CCD As Variant
Public Amount_1_80CCD As Variant

Public end_80CCD_1b As Variant
Public end_80CCD_1b_Type As Variant
Public end_80CCD_1b_Name As Variant
Public end_80CCD_1b_Amount As Variant
Public Type_1b_80CCD As Variant
Public Name_1b_80CCD As Variant
Public Amount_1b_80CCD As Variant

Public msgValidate_80CCC As String
Public msgValidate_80CCD_1 As String
Public msgValidate_80CCD_1b As String
'*************************************************
Dim msgValidateSheet19 As String
Dim msgValidateVIAerror As Variant

'added by Chetan C M for AY 2026-27 as per DE V 0.3
Public end_pran As Variant
Public pran_80CCC As Variant
'---------------end

Sub ValidateSheet80Click()
ValidateSheet80
fmsgboxStatus "Sheet 80 is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheet80()
subProcCaption = "Validating 80IA"
    If Not ValidatesheetIA80 Then
        Sheet37.Activate
        fmsgboxsmall msgValidateSheetIA80 ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
subProcCaption = "Validating 80IB"
    If Not ValidatesheetIB80 Then
        Sheet37.Activate
        fmsgboxsmall msgValidateSheetIB80 ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
subProcCaption = "Validating 80IC"
    If Not ValidatesheetIC80 Then
        Sheet37.Activate
        fmsgboxsmall msgValidateSheetIC80 ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub
Sub ValidateSheet80_IAClick()
ValidateSheet80_IA
fmsgboxStatus "Sheet VI-A is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheet80_IA()

    
    If Not ValidatesheetVIA Then
        Sheet19.Activate
        fmsgboxsmall msgValidateVIAerror ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
      
    
End Sub

Function msgbox_VIA(strmsg As String) As String
     msgValidateVIAerror = msgValidateVIAerror & strmsg & Chr(13)
End Function


Function msgbox_IA80(strmsg As String) As String
     msgValidateSheetIA80 = msgValidateSheetIA80 & strmsg & Chr(13)
End Function


Function msgbox_IB80(strmsg As String) As String
     msgValidateSheetIB80 = msgValidateSheetIB80 & strmsg & Chr(13)
End Function


Function msgbox_IC80(strmsg As String) As String
     msgValidateSheetIC80 = msgValidateSheetIC80 & strmsg & Chr(13)
End Function

Sub addrowsIA80_DeductProfUs80_IA_4_i()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IA80.DeductProfUs80_IA_4_i")
ActiveCellRange = EfilingCommon.searchLastRow("IA80.DeductProfUs80_IA_4_i")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub addrowsIA80_DeductProfUs80_IA_4_ii()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IA80.DeductProfUs80_IA_4_ii")
ActiveCellRange = EfilingCommon.searchLastRow("IA80.DeductProfUs80_IA_4_ii")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub addrowsIA80_DeductProfUs80_IA_4_iii()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IA80.DeductProfUs80_IA_4_iii")
ActiveCellRange = EfilingCommon.searchLastRow("IA80.DeductProfUs80_IA_4_iii")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub addrowsIA80_DeductProfUs80_IA_4_iv()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IA80.DeductProfUs80_IA_4_iv")
ActiveCellRange = EfilingCommon.searchLastRow("IA80.DeductProfUs80_IA_4_iv")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub addrowsIB80_DeductJKLocUs80_IB_4()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductJKLocUs80_IB_4")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductJKLocUs80_IB_4")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductBackStatesUs80_IB_4()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductBackStatesUs80_IB_4")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductBackStatesUs80_IB_4")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductBackDisttUs80_IB_5()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductBackDisttUs80_IB_5")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductBackDisttUs80_IB_5")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductMultiplexUs80_IB_7A()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductMultiplexUs80_IB_7A")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductMultiplexUs80_IB_7A")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductConvCentUs80_IB_7B()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductConvCentUs80_IB_7B")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductConvCentUs80_IB_7B")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductMinOilUs80_IB_9()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductMinOilUs80_IB_9")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductMinOilUs80_IB_9")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductHousUs80_IB_10()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductHousUs80_IB_10")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductHousUs80_IB_10")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductColdChainUs80_IB_11()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductColdChainUs80_IB_11")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductColdChainUs80_IB_11")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductFruitVegUs80_IB_11A()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductFruitVegUs80_IB_11A")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductFruitVegUs80_IB_11A")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductFoodGrainUs80_IB_11A()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductFoodGrainUs80_IB_11A")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductFoodGrainUs80_IB_11A")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductRurHospUs80_IB_11B()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductRurHospUs80_IB_11B")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductRurHospUs80_IB_11B")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIB80_DeductHospAnyAreaUs80IB_11C()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IB80.DeductHospAnyAreaUs80IB_11C")
ActiveCellRange = EfilingCommon.searchLastRow("IB80.DeductHospAnyAreaUs80IB_11C")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_DeductInSikkim()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.DeductInSikkim")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.DeductInSikkim")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_DeductInHimachalP()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.DeductInHimachalP")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.DeductInHimachalP")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_DeductInUttaranchal()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.DeductInUttaranchal")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.DeductInUttaranchal")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_Assam()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.Assam")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.Assam")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_ArunachalPradesh()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.ArunachalPradesh")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.ArunachalPradesh")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_Manipur()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.Manipur")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.Manipur")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_Mizoram()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.Mizoram")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.Mizoram")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_Meghalaya()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.Meghalaya")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.Meghalaya")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_Nagaland()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.Nagaland")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.Nagaland")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub addrowsIC80_Tripura()
Dim vRows As Variant
EfilingCommon.DefinedgridNameRange = ("IC80.Tripura")
ActiveCellRange = EfilingCommon.searchLastRow("IC80.Tripura")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

'by sadineni on 31/03/2026
Sub AddRows80CCC()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    
    Set sourceSheet = ThisWorkbook.Sheets("VI-A")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "S.No_80CCC||Type_of_Identifier_80CCC||Name_of_Identifier_80CCC||Amount_80CCC"
    ActiveCellRange = EfilingCommon.searchLastRow("Type_of_Identifier_80CCC")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_VIA
    
End Sub

'****************commented by Chetan C M for AY 2026-27(V 0.3)********************
    '--start--
''by sadineni on 31/03/2026
'Sub AddRows80CCD1()
'    Dim vRows As Long
'    Dim sourceSheet As Worksheet
'
'    Set sourceSheet = ThisWorkbook.Sheets("VI-A")
'    sourceSheet.Activate
'    EfilingCommon.DefinedgridNameRange = "S.No_80CCD1||Type_of_Identifier_80CCD1||Name_of_Identifier_80CCD1||Amount_80CCD1"
'    ActiveCellRange = EfilingCommon.searchLastRow("Type_of_Identifier_80CCD1")
'    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_VIA
'
'End Sub

''by sadineni on 31/03/2026
'Sub AddRows80CCD1b()
'    Dim vRows As Long
'    Dim sourceSheet As Worksheet
'
'    Set sourceSheet = ThisWorkbook.Sheets("VI-A")
'    sourceSheet.Activate
'    EfilingCommon.DefinedgridNameRange = "S.No_CCD1b||Type_of_Identifier_80CCD1B||Name_of_Identifier_80CCD1B||Amount_80CCD1B"
'    ActiveCellRange = EfilingCommon.searchLastRow("Type_of_Identifier_80CCD1B")
'    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_VIA
'
'End Sub
     '************************************----end

'****************added by Chetan C M for AY 2026-27(V 0.3)********************
    '--start--
    
Sub AddRows_Sec80CCD_pran()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("VI-A")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "S.No_80CCD_PRAN||Pran_Number"
    ActiveCellRange = EfilingCommon.searchLastRow("S.No_80CCD_PRAN")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_VIA
End Sub
'--end*******************************

Function ValidatesheetIA80() As Boolean
    ValidatesheetIA80 = True
    If Not ValidateDeductProfUs80_IA_4_i_IA80() Then ValidatesheetIA80 = False
    If Not ValidateDeductProfUs80_IA_4_ii_IA80() Then ValidatesheetIA80 = False
    If Not ValidateDeductProfUs80_IA_4_iii_IA80() Then ValidatesheetIA80 = False
    If Not ValidateDeductProfUs80_IA_4_iv_IA80() Then ValidatesheetIA80 = False
    If Not ValidateTotSchedule80_IA_IA80() Then ValidatesheetIA80 = False
End Function

Function ValidatesheetIB80() As Boolean
    ValidatesheetIB80 = True
    If Not ValidateDeductJKLocUs80_IB_4_IB80() Then ValidatesheetIB80 = False
'    If Not ValidateDeductBackStatesUs80_IB_4_IB80() Then ValidatesheetIB80 = False
'    If Not ValidateDeductBackDisttUs80_IB_5_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductMultiplexUs80_IB_7A_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductConvCentUs80_IB_7B_IB80() Then ValidatesheetIB80 = False
    'If Not ValidateDeductScientificResUs80_IB_8A_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductMinOilUs80_IB_9_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductHousUs80_IB_10_IB80() Then ValidatesheetIB80 = False
'    If Not ValidateDeductColdChainUs80_IB_11_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductFruitVegUs80_IB_11A_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductFoodGrainUs80_IB_11A_IB80() Then ValidatesheetIB80 = False
    If Not ValidateDeductRurHospUs80_IB_11B_IB80() Then ValidatesheetIB80 = False
    If Not ValidateTotSchedule80_IB_IB80() Then ValidatesheetIB80 = False
End Function

Function ValidatesheetIC80() As Boolean
    ValidatesheetIC80 = True
    If Not ValidateDeductInSikkim_IC80() Then ValidatesheetIC80 = False
    If Not ValidateDeductInHimachalP_IC80() Then ValidatesheetIC80 = False
    If Not ValidateDeductInUttaranchal_IC80() Then ValidatesheetIC80 = False
    If Not ValidateAssam_IC80() Then ValidatesheetIC80 = False
    If Not ValidateArunachalPradesh_IC80() Then ValidatesheetIC80 = False
    If Not ValidateManipur_IC80() Then ValidatesheetIC80 = False
    If Not ValidateMizoram_IC80() Then ValidatesheetIC80 = False
    If Not ValidateMeghalaya_IC80() Then ValidatesheetIC80 = False
    If Not ValidateNagaland_IC80() Then ValidatesheetIC80 = False
    If Not ValidateTripura_IC80() Then ValidatesheetIC80 = False
    If Not ValidateTotDeductInNorthEast_IC80() Then ValidatesheetIC80 = False
    If Not ValidateTotSchedule80_IC_IC80() Then ValidatesheetIC80 = False
End Function
Function ValidatesheetVIA() As Boolean
ValidatesheetVIA = True
 
  
 scviaSection80C = IIf(Len(Sheet19.Range("scvia.Section80C").value) = 0, 0, Sheet19.Range("scvia.Section80C").value)
 scviaSection80CCC = IIf(Len(Sheet19.Range("scvia.Section80CCC").value) = 0, 0, Sheet19.Range("scvia.Section80CCC").value)
 scviaSection80CCD_SE = IIf(Len(Sheet19.Range("scvia.Section80CCD_SE").value) = 0, 0, Sheet19.Range("scvia.Section80CCD_SE").value)
 scviaSection80CCD1B_SE = IIf(Len(Sheet19.Range("scvia.Section80CCD1B_SE").value) = 0, 0, Sheet19.Range("scvia.Section80CCD1B_SE").value)
 scviaSection80CCD = IIf(Len(Sheet19.Range("scvia.Section80CCD").value) = 0, 0, Sheet19.Range("scvia.Section80CCD").value)
 scviaSection80CCG = IIf(Len(Sheet19.Range("scvia.Section80CCG").value) = 0, 0, Sheet19.Range("scvia.Section80CCG").value)
' scviaSection80D = IIf(Len(Sheet19.Range("scvia.Section80D").value) = 0, 0, Sheet19.Range("scvia.Section80D").value)
 
 
 
    Dim Usr80DVal, Usr80DDVal, Usr80DDBVal, Usr80UVal As Variant
    Dim status_6 As Variant
    Dim DOB_6 As Variant
    
'    SELECT80D = Sheet19.Range("SELECT80D").value
   '' SYSTEM80D = Sheet19.Range("scvia.Section80D_Calc").value
    status_6 = Sheet1.Range("sheet1.Status").value
    status_6 = Mid(status_6, 1, 1)
    
    
    DOB_6 = Sheet1.Range("sheet1.DOB").value
    

    Dim HUFStatus As Variant
    HUFStatus = Mid(Sheet1.Range("sheet1.Status").value, 1, 1)


 scviaSection80DD = IIf(Len(Sheet19.Range("scvia.Section80DD").value) = 0, 0, Sheet19.Range("scvia.Section80DD").value)
 
 'SELECT80DD = Sheet19.Range("SELECT80DD").value
 'SELECT80DD = Mid(SELECT80DD, 1, 1)
  '  If isdropdownblank(SELECT80DD) Then
   '     SELECT80DD = ""
    'End If
    
   ' Usr80DDVal = Sheet19.Range("scvia.Section80DD").value
    
    'If SELECT80DD <> "" Then
     '   If IsEmpty(Usr80DDVal) Or Usr80DDVal = 0 Then
      '      msgbox_VIA ("* Please enter the value for 80DD deduction in Sheet VIA.")
       '     ValidatesheetVIA = False
        '    Exit Function
        'End If
   ' End If
    
'Jyotiuc
    
    Usr80DDVal = IIf(Trim(Usr80DDVal) = "", 0, Usr80DDVal)

    If status_6 = "I" And Usr80DDVal > 0 Then
        If isdropdownblank(SELECT80DD) Then
            'msgbox_VIA ("* Please select an option at dropdown of Sec.80DD under Chapter VIA")
            msgbox_VIA ("* Please select an option from the drop down in 80DD in Schedule VIA")
            ValidatesheetVIA = False
            Exit Function
        End If
    End If
    'uc end
 scviaSection80DDB = IIf(Len(Sheet19.Range("scvia.Section80DDB").value) = 0, 0, Sheet19.Range("scvia.Section80DDB").value)
 
 SELECT80DDB = Sheet19.Range("SELECT80DDB").value
 SELECT80DDB = Mid(SELECT80DDB, 1, 1)
 If isdropdownblank(SELECT80DDB) Then
        SELECT80DDB = ""
    End If
    
    Usr80DDBVal = Sheet19.Range("scvia.Section80DDB").value
    
    If SELECT80DDB <> "" Then
        If IsEmpty(Usr80DDBVal) Or Usr80DDBVal = 0 Then
        msgbox_VIA ("Please provide mandatory details under 80DDB deduction.") & Chr(13)
        
            msgbox_VIA ("* Please enter the value for 80DDB deduction in Sheet VIA.")
            ValidatesheetVIA = False
            Exit Function
        End If
    End If
    
    Usr80DDBVal = IIf(Trim(Usr80DDBVal) = "", 0, Usr80DDBVal)
        
    If status_6 = "I" And Usr80DDBVal > 0 Then
        If isdropdownblank(SELECT80DDB) Then
        'Jyoti69
           ' msgbox_VIA ("* Please select an option at dropdown of Sec.80DDB under Chapter VIA")
           msgbox_VIA ("Please provide mandatory details under 80DDB deduction.") & Chr(13)
            msgbox_VIA ("* Please select an option from the drop down in 80DDB in Schedule VIA")
            ValidatesheetVIA = False
            Exit Function
        End If
    End If
 
     'sudha_2025-26
        If Sheet19.Range("Sheet19.Specified_Disease").value <> "(Select)" Or Sheet19.Range("scvia.Section80DDB").value > 0 Then
        'If Sheet19.Range("scvia.Section80DDB").value > 0 Then
                If Sheet19.Range("SELECT80DDB") = "(Select)" Then
                      
                      msgbox_VIA ("Please provide mandatory details under 80DDB deduction.") & Chr(13)
                      msgbox_VIA ("* ""Please select an option from the drop down in 80DDB in Schedule VIA""") & Chr(13)
                      ValidatesheetVIA = False
                      'Exit Function
               End If
        End If

'commented by sadineni on 31/03/2026 as per DE V0.2
        
'If Sheet19.Range("scvia.Section80CCD_SE").value > 0 Or Sheet19.Range("scvia.Section80CCD1B_SE").value > 0 Then
'    If Sheet19.Range("Pran_Number").value = "" Then
'     msgbox_VIA ("* ""PRAN of the Taxpayer is mandatory""") & Chr(13)
'    ValidatesheetVIA = False
'    'Exit Function
'    End If
'End If

'****************comment by Chetan C M for AY 2026-27 as per DE V 0.3********************
'--start--
If Sheet19.Range("scvia.Section80CCD_SE").value > 0 Or Sheet19.Range("scvia.Section80CCD1B_SE").value > 0 Then
    If Not Validate_Pran Then ValidatesheetVIA = False
End If
'----end************************************
If Len(Sheet19.Range("scvia.Section80C").value) > 14 Then
msgbox_VIA ("* Amount should be Numeric, Non negative, No decimal and up to 14 digits at field 80C") & Chr(13)
ValidatesheetVIA = False
End If
 
If Len(Sheet19.Range("scvia.Section80E").value) > 14 Then
msgbox_VIA ("* Amount should be Numeric, Non negative, No decimal and up to 14 digits at field 80E") & Chr(13)
ValidatesheetVIA = False
End If
 
If Len(Sheet19.Range("scvia.Section80EE").value) > 14 Then
msgbox_VIA ("* Amount should be Numeric, Non negative, No decimal and up to 14 digits at field 80EE") & Chr(13)
ValidatesheetVIA = False
End If
 
 
If Len(Sheet19.Range("scvia.Section80EEA").value) > 14 Then
msgbox_VIA ("* Amount should be Numeric, Non negative, No decimal and up to 14 digits at field 80EEA") & Chr(13)
ValidatesheetVIA = False
End If
 
If Len(Sheet19.Range("scvia.Section80EEB").value) > 14 Then
msgbox_VIA ("* Amount should be Numeric, Non negative, No decimal and up to 14 digits at field 80EEB") & Chr(13)
ValidatesheetVIA = False
End If


'If Sheet19.Range("scvia.Section80CCC").value > 0 Or Sheet19.Range("scvia.Section80CCD_SE").value > 0 Or Sheet19.Range("scvia.Section80CCD1B_SE").value > 0 Then
'    'If Not checkspecialcharacterForPinCode(Sheet19.Range("Pran_Number").value) Then
'        If Len(Sheet19.Range("Pran_Number").value) > 125 Then
'            msgbox_VIA ("*""Please enter valid Pran Number""") & Chr(13)
'            ValidatesheetVIA = False
'            'Exit Function
'        End If
'    'End If
'End If


    'Added by Riyaz on 24/06/2024
        If Sheet19.Range("Sheet19.Specified_Disease").Locked = False Then
        If Sheet19.Range("SELECT80DDB").value <> "(Select)" And Sheet19.Range("scvia.Section80DDB").value > 0 Then
                If Sheet19.Range("Sheet19.Specified_Disease") = "" Or Sheet19.Range("Sheet19.Specified_Disease") = "(Select)" Then
                      msgbox_VIA ("* ""Name of Specified disease is mandatory""") & Chr(13)
                      ValidatesheetVIA = False
                      
               End If
        End If
        End If
        
  
    
If Sheet19.Range("scvia.Section80GG").value > 0 Then
      If Sheet19.Range("Sheet19.AckNum").value = "" Then
    
          msgbox_VIA ("*""Please enter Acknowledgement number of Form 10BA filed for AY 2025-26 against Section 80GG in Schedule VI A""") & Chr(13)
            ValidatesheetVIA = False
        'Exit Function
    End If

End If

'Newly added by Bindu on 11th June 25
If Len(Sheet19.Range("Sheet19.AckNum").value) > 0 Then
If SheetALL.Range("bacValue").value = 2 Then
If Sheet1.Range("sheet1.Status") <> "H-HUF" Then
      If Sheet19.Range("scvia.Section80GG").value = "" Or Sheet19.Range("scvia.Section80GG").value = 0 Then
    
          msgbox_VIA ("*""80GG amount shall be more than 0 when acknowledgement no. of form 10BA is provided""") & Chr(13)
            ValidatesheetVIA = False
        'Exit Function
    End If
 End If
  End If
End If

    
    
If Sheet19.Range("scvia.Section80QQB").value > 0 Then
      If Sheet19.Range("Sheet19.AckNum10CCD").value = "" Then
    
        msgbox_VIA ("*""Please enter Acknowledgement number of Form 10CCD filed for AY 2025-26 against Section 80QQB in Schedule VI A""") & Chr(13)
        ValidatesheetVIA = False
        'Exit Function
    End If

End If

'Newly added by Bindu on 11th June 25
If Len(Sheet19.Range("Sheet19.AckNum10CCD").value) > 0 Then
If SheetALL.Range("bacValue").value = 2 Then
If Sheet1.Range("sheet1.Status") <> "H-HUF" Then

      'If Sheet19.Range("scvia.Section80QQB").value = "" Or Sheet19.Range("scvia.Section80GG").value = 0 Then
      If Sheet19.Range("scvia.Section80QQB").value = "" Or Sheet19.Range("scvia.Section80QQB").value = 0 Then
          msgbox_VIA ("*""80QQB amount shall be more than 0 when acknowledgement no. of form 10CCD is provided""") & Chr(13)
            ValidatesheetVIA = False
        'Exit Function
    End If
 End If
End If
 End If

If Sheet19.Range("scvia.Section80RRB").value > 0 Then
      If Sheet19.Range("Sheet19.AckNum.10CCE").value = "" Then
    
      msgbox_VIA ("*""Please enter Acknowledgement number of Form 10CCE filed for AY 2025-26 against Section 80RRB in Schedule VI A""") & Chr(13)
    ValidatesheetVIA = False
    'Exit Function
    End If

End If

'Newly added by Bindu on 11th June 25
If Len(Sheet19.Range("Sheet19.AckNum.10CCE").value) > 0 Then
If SheetALL.Range("bacValue").value = 2 Then
If Sheet1.Range("sheet1.Status") <> "H-HUF" Then
     ' If Sheet19.Range("scvia.Section80RRB").value = "" Or Sheet19.Range("scvia.Section80GG").value = 0 Then
     If Sheet19.Range("scvia.Section80RRB").value = "" Or Sheet19.Range("scvia.Section80RRB").value = 0 Then
          msgbox_VIA ("*""80RRB amount shall be more than 0 when acknowledgement no. of form 10CCE is provided""") & Chr(13)
            ValidatesheetVIA = False
        'Exit Function
    End If
 End If
  End If
End If
    
 '======================================================
 
 scviaSection80E = IIf(Len(Sheet19.Range("scvia.Section80E").value) = 0, 0, Sheet19.Range("scvia.Section80E").value)
 scviaSection80EE = IIf(Len(Sheet19.Range("scvia.Section80EE").value) = 0, 0, Sheet19.Range("scvia.Section80EE").value)
 scviaSection80G = IIf(Len(Sheet19.Range("scvia.Section80G").value) = 0, 0, Sheet19.Range("scvia.Section80G").value)
 scviaSection80GG = IIf(Len(Sheet19.Range("scvia.Section80GG").value) = 0, 0, Sheet19.Range("scvia.Section80GG").value)
 'scviaSection80GGC = IIf(Len(Sheet19.Range("scvia.Section80GGC").value) = 0, 0, Sheet19.Range("scvia.Section80GGC").value)
 scviaTotPartBchapterVIA = IIf(Len(Sheet19.Range("scvia.TotPartBchapterVIA").value) = 0, 0, Sheet19.Range("scvia.TotPartBchapterVIA").value)
 scviaSection80IA = IIf(Len(Sheet19.Range("scvia.Section80IA").value) = 0, 0, Sheet19.Range("scvia.Section80IA").value)
 scviaSection80IAB = IIf(Len(Sheet19.Range("scvia.Section80IAB").value) = 0, 0, Sheet19.Range("scvia.Section80IAB").value)
 scviaSection80IB = IIf(Len(Sheet19.Range("scvia.Section80IB").value) = 0, 0, Sheet19.Range("scvia.Section80IB").value)
 scviaSection80IBA = IIf(Len(Sheet19.Range("scvia.Section80IBA").value) = 0, 0, Sheet19.Range("scvia.Section80IBA").value)
 scviaSection80IC = IIf(Len(Sheet19.Range("scvia.Section80IC").value) = 0, 0, Sheet19.Range("scvia.Section80IC").value)
 'scviaSection80ID = IIf(Len(Sheet19.Range("scvia.Section80ID").value) = 0, 0, Sheet19.Range("scvia.Section80ID").value)
 scviaSection80JJA = IIf(Len(Sheet19.Range("scvia.Section80JJA").value) = 0, 0, Sheet19.Range("scvia.Section80JJA").value)
 scviaSection80JJAA = IIf(Len(Sheet19.Range("scvia.Section80JJAA").value) = 0, 0, Sheet19.Range("scvia.Section80JJAA").value)
 scviaSection80QQB = IIf(Len(Sheet19.Range("scvia.Section80QQB").value) = 0, 0, Sheet19.Range("scvia.Section80QQB").value)
 scviaSection80RRB = IIf(Len(Sheet19.Range("scvia.Section80RRB").value) = 0, 0, Sheet19.Range("scvia.Section80RRB").value)
 scviaTotPartCchapterVIA = IIf(Len(Sheet19.Range("scvia.TotPartCchapterVIA").value) = 0, 0, Sheet19.Range("scvia.TotPartCchapterVIA").value)
 scviaSection80TTB = IIf(Len(Sheet19.Range("scvia.Section80TTB").value) = 0, 0, Sheet19.Range("scvia.Section80TTB").value)
 scviaSection80TTA = IIf(Len(Sheet19.Range("scvia.Section80TTA").value) = 0, 0, Sheet19.Range("scvia.Section80TTA").value)
 scviaSection80U = IIf(Len(Sheet19.Range("scvia.Section80U").value) = 0, 0, Sheet19.Range("scvia.Section80U").value)
 
' SELECT80U = Sheet19.Range("SELECT80U").value
' SELECT80U = Mid(SELECT80U, 1, 1)
'  If isdropdownblank(SELECT80U) Then
'        SELECT80U = ""
'    End If
'
'    Usr80UVal = Sheet19.Range("scvia.Section80U").value
'
'    If SELECT80U <> "" Then
'        If IsEmpty(Usr80UVal) Or Usr80UVal = 0 Then
'            msgbox_VIA ("* Please enter the value for 80U deduction in Sheet VIA.")
'            ValidatesheetVIA = False
'            Exit Function
'        End If
'    End If
    
'Change-25.04.2023.101.VIA -> SIT-49347
'If (Sheet19.Range("Section80CCH").value = "80CCH (applicable for only Central govt emplyees)" And Sheet19.Range("Anyother80CCH_Calc").value = 0 And Sheet19.Range("Anyother80CCH").value = 0) Then
  'Chandru (AY2024-25 Change)
  'If (Sheet19.Range("Anyother80CCH_Calc").value = 0 And Sheet19.Range("Anyother80CCH").value = 0) Then

'uncommented by Chetan C M for AY 2026-27 DE V0.3 on 05/05/2026 (SIT-118781)
'start**************
'Newly commented by Bindu as per DE 0.2 80CCH is not Mandatory
If (Sheet19.Range("Anyother80CCH_Calc").value = 0 And Sheet19.Range("Anyother80CCH").value = 0) And Sheet19.Range("Anyother80CCH").Locked = False Then
    msgbox_VIA ("* Please enter amount of 80CCH.")
    ValidatesheetVIA = False
    Exit Function
End If
'---
'end****************


'    Usr80UVal = IIf(Trim(Usr80UVal) = "", 0, Usr80UVal)
'
'    If status_6 = "I" And Usr80UVal > 0 Then
'        If isdropdownblank(SELECT80U) Then
'            msgbox_VIA ("* Please select an option at dropdown of Sec.80U under Chapter VIA")
'            ValidatesheetVIA = False
'            Exit Function
'        End If
'    End If
'
'************************************************************
'by sadineni on 31/03/2026
If Not Validate_80CCC() Then ValidatesheetVIA = False

'****************commented by Chetan C M for AY 2026-27(V 0.3)********************
'--start--
'If Not Validate_80CCD_1() Then ValidatesheetVIA = False
'If Not Validate_80CCD_1b() Then ValidatesheetVIA = False
    '************************************----end
    
'***********************************************************


 scviaTotPartCAandDchapterVIA = IIf(Len(Sheet19.Range("scvia.TotPartCAandDchapterVIA").value) = 0, 0, Sheet19.Range("scvia.TotPartCAandDchapterVIA").value)
 scviaUsrTotalChapVIADeductions = IIf(Len(Sheet19.Range("scvia.UsrTotalChapVIADeductions").value) = 0, 0, Sheet19.Range("scvia.UsrTotalChapVIADeductions").value)

 scviaSection80C_Calc = IIf(Len(Sheet19.Range("scvia.Section80C_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80C_Calc").value)
 scviaSection80CCC_Calc = IIf(Len(Sheet19.Range("scvia.Section80CCC_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80CCC_Calc").value)
 scviaSection80CCD_SE_Calc = IIf(Len(Sheet19.Range("scvia.Section80CCD_SE_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80CCD_SE_Calc").value)
 scviaSection80CCD1B_SE_Calc = IIf(Len(Sheet19.Range("scvia.Section80CCD1B_SE_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80CCD1B_SE_Calc").value)
 scviaSection80CCD_Calc = IIf(Len(Sheet19.Range("scvia.Section80CCD_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80CCD_Calc").value)
 scviaSection80CCG_Calc = IIf(Len(Sheet19.Range("scvia.Section80CCG_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80CCG_Calc").value)
 scviaSection80DD_Calc = IIf(Len(Sheet19.Range("scvia.Section80DD_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80DD_Calc").value)
 scviaSection80DDB_Calc = IIf(Len(Sheet19.Range("scvia.Section80DDB_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80DDB_Calc").value)
 scviaSection80E_Calc = IIf(Len(Sheet19.Range("scvia.Section80E_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80E_Calc").value)
 scviaSection80EE_Calc = IIf(Len(Sheet19.Range("scvia.Section80EE_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80EE_Calc").value)
 scviaSection80G_Calc = IIf(Len(Sheet19.Range("scvia.Section80G_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80G_Calc").value)
 scviaSection80GG_Calc = IIf(Len(Sheet19.Range("scvia.Section80GG_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80GG_Calc").value)
 'scviaSection80GGC_Calc = IIf(Len(Sheet19.Range("scvia.Section80GGC_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80GGC_Calc").value)
 scviaTotPartBchapterVIA_Calc = IIf(Len(Sheet19.Range("scvia.TotPartBchapterVIA_Calc").value) = 0, 0, Sheet19.Range("scvia.TotPartBchapterVIA_Calc").value)
 scviaSection80IA_Calc = IIf(Len(Sheet19.Range("scvia.Section80IA_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80IA_Calc").value)
 scviaSection80IAB_Calc = IIf(Len(Sheet19.Range("scvia.Section80IAB_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80IAB_Calc").value)
 scviaSection80IB_Calc = IIf(Len(Sheet19.Range("scvia.Section80IB_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80IB_Calc").value)
 scviaSection80IBA_Calc = IIf(Len(Sheet19.Range("scvia.Section80IBA_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80IBA_Calc").value)
 scviaSection80IC_Calc = IIf(Len(Sheet19.Range("scvia.Section80IC_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80IC_Calc").value)
 'scviaSection80ID_Calc = IIf(Len(Sheet19.Range("scvia.Section80ID_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80ID_Calc").value)
 scviaSection80JJA_Calc = IIf(Len(Sheet19.Range("scvia.Section80JJA_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80JJA_Calc").value)
 scviaSection80JJAA_Calc = IIf(Len(Sheet19.Range("scvia.Section80JJAA_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80JJAA_Calc").value)
 scviaSection80QQB_Calc = IIf(Len(Sheet19.Range("scvia.Section80QQB_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80QQB_Calc").value)
 scviaSection80RRB_Calc = IIf(Len(Sheet19.Range("scvia.Section80RRB_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80RRB_Calc").value)
 scviaTotPartCchapterVIA_Calc = IIf(Len(Sheet19.Range("scvia.TotPartCchapterVIA_Calc").value) = 0, 0, Sheet19.Range("scvia.TotPartCchapterVIA_Calc").value)
 scviaSection80TTB_Calc = IIf(Len(Sheet19.Range("scvia.Section80TTB_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80TTB_Calc").value)
 scviaSection80TTA_Calc = IIf(Len(Sheet19.Range("scvia.Section80TTA_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80TTA_Calc").value)
 scviaSection80U_Calc = IIf(Len(Sheet19.Range("scvia.Section80U_Calc").value) = 0, 0, Sheet19.Range("scvia.Section80U_Calc").value)
 scviaTotPartCAandDchapterVIA_Calc = IIf(Len(Sheet19.Range("scvia.TotPartCAandDchapterVIA_Calc").value) = 0, 0, Sheet19.Range("scvia.TotPartCAandDchapterVIA_Calc").value)
 scviaTotalChapVIADeductions_Calc = IIf(Len(Sheet19.Range("scvia.TotalChapVIADeductions").value) = 0, 0, Sheet19.Range("scvia.TotalChapVIADeductions").value)


End Function

Function ValidateDeductProfUs80_IA_4_i_IA80() As Boolean
    ValidateDeductProfUs80_IA_4_i_IA80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IA80.DeductProfUs80_IA_4_i").Cells
    count = 0
    
    noOfProcessSub = rangecells.Cells.count
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    UpdateProgressBar
    Next
    
    ReDim IA80_DeductProfUs80_IA_4_I(count)
    For i = 1 To count
    IA80_DeductProfUs80_IA_4_I(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductProfUs80_IA_4_ii_IA80() As Boolean
    ValidateDeductProfUs80_IA_4_ii_IA80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IA80.DeductProfUs80_IA_4_ii").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IA80_DeductProfUs80_IA_4_Ii(count)
    For i = 1 To count
    IA80_DeductProfUs80_IA_4_Ii(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductProfUs80_IA_4_iii_IA80() As Boolean
    ValidateDeductProfUs80_IA_4_iii_IA80 = True
   
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IA80.DeductProfUs80_IA_4_iii").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IA80_DeductProfUs80_IA_4_iii(count)
    For i = 1 To count
    IA80_DeductProfUs80_IA_4_iii(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductProfUs80_IA_4_iv_IA80() As Boolean
    ValidateDeductProfUs80_IA_4_iv_IA80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IA80.DeductProfUs80_IA_4_iv").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IA80_DeductProfUs80_IA_4_iv(count)
    For i = 1 To count
    IA80_DeductProfUs80_IA_4_iv(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductProfUs80_IA_4_v_IA80() As Boolean
    ValidateDeductProfUs80_IA_4_v_IA80 = True
        Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IA80.DeductProfUs80_IA_4_iv").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IA80_DeductProfUs80_IA_4_iii(count)
    For i = 1 To count
    IA80_DeductProfUs80_IA_4_iii(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotSchedule80_IA_IA80() As Boolean
    ValidateTotSchedule80_IA_IA80 = True
    IA80_TotSchedule80_IA = Sheet37.Range("IA80.TotSchedule80_IA").value
    
    If Len(IA80_TotSchedule80_IA) > 14 Then
        msgbox_IA80 ("* TotSchedule80_IA in Sch 80 IA should not be greater than 14 digits ")
        ValidateTotSchedule80_IA_IA80 = False
        Exit Function
    End If
End Function

Function ValidateDeductJKLocUs80_IB_4_IB80() As Boolean
    ValidateDeductJKLocUs80_IB_4_IB80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductJKLocUs80_IB_4").Cells
    count = 0
    noOfProcessSub = rangecells.Cells.count
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    UpdateProgressBar
    Next
    
    ReDim IB80_DeductJKLocUs80_IB_4(count)
    For i = 1 To count
    IB80_DeductJKLocUs80_IB_4(i) = rangecells.item(i).value
    Next
End Function

'Function ValidateDeductBackStatesUs80_IB_4_IB80() As Boolean
'    ValidateDeductBackStatesUs80_IB_4_IB80 = True
'    Dim rangecells As Range
'    Dim Count, i As Variant
'    Set rangecells = Sheet37.Range("IB80.DeductBackStatesUs80_IB_4").Cells
'    Count = 0
'    For i = 1 To (rangecells.Cells.Count)
'    If rangecells.item(i).value <> "" Then
'    Count = Count + 1
'    End If
'    Next
'
'    ReDim IB80_DeductBackStatesUs80_IB_4(Count)
'    For i = 1 To Count
'    IB80_DeductBackStatesUs80_IB_4(i) = rangecells.item(i).value
'    Next
'End Function

'Function ValidateDeductBackDisttUs80_IB_5_IB80() As Boolean
'    ValidateDeductBackDisttUs80_IB_5_IB80 = True
'     Dim rangecells As Range
'    Dim Count, i As Variant
'    Set rangecells = Sheet37.Range("IB80.DeductBackDisttUs80_IB_5").Cells
'    Count = 0
'    For i = 1 To (rangecells.Cells.Count)
'    If rangecells.item(i).value <> "" Then
'    Count = Count + 1
'    End If
'    Next
'
'    ReDim IB80_DeductBackDisttUs80_IB_5(Count)
'    For i = 1 To Count
'    IB80_DeductBackDisttUs80_IB_5(i) = rangecells.item(i).value
'    Next
'End Function

Function ValidateDeductMultiplexUs80_IB_7A_IB80() As Boolean
    ValidateDeductMultiplexUs80_IB_7A_IB80 = True
     Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductMultiplexUs80_IB_7A").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductMultiplexUs80_IB_7A(count)
    For i = 1 To count
    IB80_DeductMultiplexUs80_IB_7A(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductConvCentUs80_IB_7B_IB80() As Boolean
    ValidateDeductConvCentUs80_IB_7B_IB80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductConvCentUs80_IB_7B").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductConvCentUs80_IB_7B(count)
    For i = 1 To count
    IB80_DeductConvCentUs80_IB_7B(i) = rangecells.item(i).value
    Next
End Function


Function ValidateDeductMinOilUs80_IB_9_IB80() As Boolean
    ValidateDeductMinOilUs80_IB_9_IB80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductMinOilUs80_IB_9").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductMinOilUs80_IB_9(count)
    For i = 1 To count
    IB80_DeductMinOilUs80_IB_9(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductHousUs80_IB_10_IB80() As Boolean
    ValidateDeductHousUs80_IB_10_IB80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductHousUs80_IB_10").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductHousUs80_IB_10(count)
    For i = 1 To count
    IB80_DeductHousUs80_IB_10(i) = rangecells.item(i).value
    Next
End Function

'Function ValidateDeductColdChainUs80_IB_11_IB80() As Boolean
'    ValidateDeductColdChainUs80_IB_11_IB80 = True
'    Dim rangecells As Range
'    Dim Count, i As Variant
'    Set rangecells = Sheet37.Range("IB80.DeductColdChainUs80_IB_11").Cells
'    Count = 0
'    For i = 1 To (rangecells.Cells.Count)
'    If rangecells.item(i).value <> "" Then
'    Count = Count + 1
'    End If
'    Next
'
'    ReDim IB80_DeductColdChainUs80_IB_11(Count)
'    For i = 1 To Count
'    IB80_DeductColdChainUs80_IB_11(i) = rangecells.item(i).value
'    Next
'End Function

Function ValidateDeductFruitVegUs80_IB_11A_IB80() As Boolean
    ValidateDeductFruitVegUs80_IB_11A_IB80 = True

    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductFruitVegUs80_IB_11A").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next

    ReDim IB80_DeductFruitVegUs80_IB_11A(count)
    For i = 1 To count
    IB80_DeductFruitVegUs80_IB_11A(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductFoodGrainUs80_IB_11A_IB80() As Boolean
    ValidateDeductFoodGrainUs80_IB_11A_IB80 = True
    
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductFoodGrainUs80_IB_11A").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductFoodGrainUs80_IB_11A(count)
    For i = 1 To count
    IB80_DeductFoodGrainUs80_IB_11A(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductRurHospUs80_IB_11B_IB80() As Boolean
    ValidateDeductRurHospUs80_IB_11B_IB80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IB80.DeductRurHospUs80_IB_11B").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductRurHospUs80_IB_11B(count)
    For i = 1 To count
    IB80_DeductRurHospUs80_IB_11B(i) = rangecells.item(i).value
    Next
    
   
    Set rangecells = Sheet37.Range("IB80.DeductHospAnyAreaUs80IB_11C").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IB80_DeductHospAnyAreaUs80IB_11C(count)
    For i = 1 To count
    IB80_DeductHospAnyAreaUs80IB_11C(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotSchedule80_IB_IB80() As Boolean
    ValidateTotSchedule80_IB_IB80 = True
    IB80_TotSchedule80_IB = Sheet37.Range("IB80.TotSchedule80_IB").value
    
    If Len(IB80_TotSchedule80_IB) > 14 Then
        msgbox_IB80 ("* TotSchedule80_IB in Sch 80 IB should not be greater than 14 digits ")
        ValidateTotSchedule80_IB_IB80 = False
        Exit Function
    End If
End Function

Function ValidateDeductInSikkim_IC80() As Boolean
    ValidateDeductInSikkim_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.DeductInSikkim").Cells
    count = 0

    noOfProcessSub = rangecells.Cells.count
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    UpdateProgressBar
    Next

    ReDim IC80_DeductInSikkim(count)
    For i = 1 To count
    IC80_DeductInSikkim(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductInHimachalP_IC80() As Boolean
    ValidateDeductInHimachalP_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.DeductInHimachalP").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next

    ReDim IC80_DeductInHimachalP(count)
    For i = 1 To count
    IC80_DeductInHimachalP(i) = rangecells.item(i).value
    Next
End Function

Function ValidateDeductInUttaranchal_IC80() As Boolean
    ValidateDeductInUttaranchal_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.DeductInUttaranchal").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next

    ReDim IC80_DeductInUttaranchal(count)
    For i = 1 To count
    IC80_DeductInUttaranchal(i) = rangecells.item(i).value
    Next
End Function

Function ValidateAssam_IC80() As Boolean
    ValidateAssam_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.Assam").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_Assam(count)
    For i = 1 To count
    IC80_Assam(i) = rangecells.item(i).value
    Next
End Function

Function ValidateArunachalPradesh_IC80() As Boolean
    ValidateArunachalPradesh_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.ArunachalPradesh").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_ArunachalPradesh(count)
    For i = 1 To count
    IC80_ArunachalPradesh(i) = rangecells.item(i).value
    Next
End Function

Function ValidateManipur_IC80() As Boolean
    ValidateManipur_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.Manipur").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_Manipur(count)
    For i = 1 To count
    IC80_Manipur(i) = rangecells.item(i).value
    Next
End Function

Function ValidateMizoram_IC80() As Boolean
    ValidateMizoram_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.Mizoram").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_Mizoram(count)
    For i = 1 To count
    IC80_Mizoram(i) = rangecells.item(i).value
    Next
End Function

Function ValidateMeghalaya_IC80() As Boolean
    ValidateMeghalaya_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.Meghalaya").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_Meghalaya(count)
    For i = 1 To count
    IC80_Meghalaya(i) = rangecells.item(i).value
    Next
End Function

Function ValidateNagaland_IC80() As Boolean
    ValidateNagaland_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.Nagaland").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_Nagaland(count)
    For i = 1 To count
    IC80_Nagaland(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTripura_IC80() As Boolean
    ValidateTripura_IC80 = True
    Dim rangecells As Range
    Dim count, i As Variant
    Set rangecells = Sheet37.Range("IC80.Tripura").Cells
    count = 0
    For i = 1 To (rangecells.Cells.count)
    If rangecells.item(i).value <> "" Then
    count = count + 1
    End If
    Next
    
    ReDim IC80_Tripura(count)
    For i = 1 To count
    IC80_Tripura(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotDeductInNorthEast_IC80() As Boolean
    ValidateTotDeductInNorthEast_IC80 = True
    IC80_TotDeductInNorthEast = Sheet37.Range("IC80.TotDeductInNorthEast").value
    
    If Len(IC80_TotDeductInNorthEast) > 14 Then
        msgbox_IC80 ("* TotDeductInNorthEast in Sch 80-IC should not be greater than 14 digits ")
        ValidateTotDeductInNorthEast_IC80 = False
        Exit Function
    End If
End Function

Function ValidateTotSchedule80_IC_IC80() As Boolean
    ValidateTotSchedule80_IC_IC80 = True
    IC80_TotSchedule80_IC = Sheet37.Range("IC80.TotSchedule80_IC").value
    
    If Len(IC80_TotSchedule80_IC) > 14 Then
        msgbox_IC80 ("* TotSchedule80_IC in Sch 80-IC should not be greater than 14 digits ")
        ValidateTotSchedule80_IC_IC80 = False
        Exit Function
    End If
End Function

Function msgbox19(strmsg As String) As String
    msgValidateSheet19 = msgValidateSheet19 & strmsg & Chr(13)
End Function

'by sadineni on 31/03/2026
Sub setTblinfo_80CCC_Type()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Type_of_Identifier_80CCC").count
    Set rangecells = Range("Type_of_Identifier_80CCC").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_80CCCType = ccount
 End Sub

'by sadineni on 31/03/2026
 Sub setTblinfo_80CCC_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Name_of_Identifier_80CCC").count
    Set rangecells = Range("Name_of_Identifier_80CCC").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCCName = ccount
End Sub

'by sadineni on 31/03/2026
Sub setTblinfo_80CCC_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Amount_80CCC").count
    Set rangecells = Range("Amount_80CCC").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCCAmount = ccount
End Sub

 '****************commented by Chetan C M for AY 2026-27(V 0.3)********************
    '--start--
''by sadineni on 31/03/2026
'Sub setTblinfo_80CCD_1b_Type()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Type_of_Identifier_80CCD1B").count
'    Set rangecells = Range("Type_of_Identifier_80CCD1B").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not (rangecells.item(mIntCtr).value = "") Then
'                ccount = ccount + 1
'            End If
'    Next
' end_80CCD_1b_Type = ccount
' End Sub
'
''by sadineni on 31/03/2026
'
' Sub setTblinfo_80CCD_1b_Name()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Name_of_Identifier_80CCD1B").count
'    Set rangecells = Range("Name_of_Identifier_80CCD1B").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80CCD_1b_Name = ccount
'End Sub
'
''by sadineni on 31/03/2026
'Sub setTblinfo_80CCD_1b_Amount()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Amount_80CCD1B").count
'    Set rangecells = Range("Amount_80CCD1B").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80CCD_1b_Amount = ccount
'End Sub
'
''by sadineni on 31/03/2026
'Sub setTblinfo_80CCD_1_Type()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Type_of_Identifier_80CCD1").count
'    Set rangecells = Range("Type_of_Identifier_80CCD1").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not (rangecells.item(mIntCtr).value = "") Then
'                ccount = ccount + 1
'            End If
'    Next
' end_80CCD_1_Type = ccount
' End Sub
'
''by sadineni on 31/03/2026
' Sub setTblinfo_80CCD_1_Name()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Name_of_Identifier_80CCD1").count
'    Set rangecells = Range("Name_of_Identifier_80CCD1").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80CCD_1_Name = ccount
'End Sub
'
''by sadineni on 31/03/2026
'Sub setTblinfo_80CCD_1_Amount()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Amount_80CCD1").count
'    Set rangecells = Range("Amount_80CCD1").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80CCD_1_Amount = ccount
'End Sub
 '---end****************************************
 
'by sadineni on 31/03/2026
Function Validate_80CCC()
Validate_80CCC = True
setTblinfo_80CCC_Type
setTblinfo_80CCC_Name
setTblinfo_80CCC_Amount

end_80CCC = WorksheetFunction.Max(0, end_80CCCType, end_80CCCName, end_80CCCAmount)

If Not ValidateType_80CCC Then Validate_80CCC = False
If Not ValidateName_80CCC Then Validate_80CCC = False
If Not ValidateAmount_80CCC Then Validate_80CCC = False
End Function

'****************commented by Chetan C M for AY 2026-27(V 0.3)********************
'--start--
'by sadineni on 31/03/2026
'Function Validate_80CCD_1()
'Validate_80CCD_1 = True
'setTblinfo_80CCD_1_Type
'setTblinfo_80CCD_1_Name
'setTblinfo_80CCD_1_Amount
'end_80CCD_1 = WorksheetFunction.Max(0, end_80CCD_1_Type, end_80CCD_1_Name, end_80CCD_1_Amount)
'If Not ValidateType_1_80CCD Then Validate_80CCD_1 = False
'If Not ValidateName_1_80CCD Then Validate_80CCD_1 = False
'If Not ValidateAmount_1_80CCD Then Validate_80CCD_1 = False
'End Function
'
''by sadineni on 31/03/2026
'Function Validate_80CCD_1b()
'Validate_80CCD_1b = True
'setTblinfo_80CCD_1b_Type
'setTblinfo_80CCD_1b_Name
'setTblinfo_80CCD_1b_Amount
'end_80CCD_1b = WorksheetFunction.Max(0, end_80CCD_1b_Type, end_80CCD_1b_Name, end_80CCD_1b_Amount)
'If Not ValidateType_1b_80CCD Then Validate_80CCD_1b = False
'If Not ValidateName_1b_80CCD Then Validate_80CCD_1b = False
'If Not ValidateAmount_1b_80CCD Then Validate_80CCD_1b = False
'End Function
'************************************----end

'by sadineni on 31/03/2026
Function ValidateType_80CCC() As Boolean
    ValidateType_80CCC = True
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet19.Range("Type_of_Identifier_80CCC").Cells

    ReDim Type_80CCC(end_80CCC)
    For i = 1 To end_80CCC
    
        Type_80CCC(i) = rangecells.item(i, 1).value
        
        If Not chkCompulsory(Type_80CCC(i)) Then
            msgValidateVIAerror = msgValidateVIAerror + "* ""Type of Identifier is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If

         If Len(Type_80CCC(i)) > 125 Then
          msgValidateVIAerror = msgValidateVIAerror + "* Type of Identifier in schedule 80CCC at Sr. No " & i & """ should be less than or equal to 125 characters." & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Type_80CCC(i)) Then
             msgValidateVIAerror = msgValidateVIAerror + "* Type of Identifier in schedule 80CCC at Sl.no. " & i & """ should not Contain special characters." & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
    Next
End Function

'by sadineni on 31/03/2026
Function ValidateName_80CCC() As Boolean
    ValidateName_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet19.Range("Name_of_Identifier_80CCC").Cells
    
    ReDim Name_80CCC(end_80CCC)
    For i = 1 To end_80CCC
    
        Name_80CCC(i) = rangecells.item(i, 1).value
        
        If Not chkCompulsory(Name_80CCC(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Name of Identifier is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
             msgValidateVIAerror = msgValidateVIAerror + "* ""Identifier No. is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
        
         If Len(Name_80CCC(i)) > 125 Then
         
        'changed by Chetan C M for AY 2026-27 as per DE V 0.3--start--
          'msgValidateVIAerror = msgValidateVIAerror + "* Name of Identifier in schedule 80CCC at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
         msgValidateVIAerror = msgValidateVIAerror + "* Identifier No. in schedule 80CCC at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
        '--end
            
            ValidateName_80CCC = False
            Exit Function
        End If
        
        If Not checkallfieldspecialcharacter(Name_80CCC(i)) Then
        
        'changed by Chetan C M for AY 2026-27 as per DE V 0.3--start--
             'msgValidateVIAerror = msgValidateVIAerror + "* Name of Identifier in schedule 80CCC at Sl.no. " & i & " should not Contain special characters." & Chr(13)
             msgValidateVIAerror = msgValidateVIAerror + "*Identifier No. in schedule 80CCC at Sl.no. " & i & " should not Contain special characters." & Chr(13)
        '--end
            
            ValidateName_80CCC = False
            Sheet19.Range("Name_of_Identifier_80CCC").ClearContents
            Exit Function
        End If
         
    Next
End Function

'by sadineni on 31/03/2026
Function ValidateAmount_80CCC() As Boolean
    ValidateAmount_80CCC = True
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet19.Range("Amount_80CCC").Cells
    
    ReDim Amount_80CCC(end_80CCC)
    For i = 1 To end_80CCC
    
        Amount_80CCC(i) = rangecells.item(i, 1).value
        If Not chkCompulsory(Amount_80CCC(i)) Then
            msgValidateVIAerror = msgValidateVIAerror + "* ""Amount is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
        
        If Not IsNumeric(Amount_80CCC(i)) Then
          msgValidateVIAerror = msgValidateVIAerror + "* Amount in schedule 80CCC at Sr. No " & i & " should be Numeric value." & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
        If Amount_80CCC(i) > 99999999999999# Then
             msgValidateVIAerror = msgValidateVIAerror + "* Amount in schedule 80CCC at Sl.no. " & i & " should not Contain special characters." & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
         
    Next
End Function

'****************commented by Chetan C M for AY 2026-27(V 0.3)********************
'--start--
''by sadineni on 31/03/2026
'Function ValidateType_1_80CCD() As Boolean
'    ValidateType_1_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Type_of_Identifier_80CCD1").Cells
'    ReDim Type_1_80CCD(end_80CCD_1)
'    For i = 1 To end_80CCD_1
'        Type_1_80CCD(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(Type_1_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Type of Identifier is mandatory in schedule 80CCD(1) at Sr. No " & i & """" & Chr(13)
'            ValidateType_1_80CCD = False
'            Exit Function
'        End If
'         If Len(Type_1_80CCD(i)) > 125 Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Type of Identifier in schedule 80CCD(1) at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateType_1_80CCD = False
'            Exit Function
'        End If
'        If Not checkfieldspecialcharacter(Type_1_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Type of Identifier in schedule 80CCD(1) at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateType_1_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
''by sadineni on 31/03/2026
'Function ValidateName_1_80CCD() As Boolean
'    ValidateName_1_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Name_of_Identifier_80CCD1").Cells
'    ReDim Name_1_80CCD(end_80CCD_1)
'    For i = 1 To end_80CCD_1
'        Name_1_80CCD(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(Name_1_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Name of Identifier is mandatory in schedule 80CCD(1) at Sr. No " & i & """" & Chr(13)
'            ValidateName_1_80CCD = False
'            Exit Function
'        End If
'         If Len(Name_1_80CCD(i)) > 125 Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Name of Identifier in schedule 80CCD(1) length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateName_1_80CCD = False
'            Exit Function
'        End If
'
'        If Not checkfieldspecialcharacter(Name_1_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Name of Identifier in schedule 80CCD(1) character at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateName_1_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
''by sadineni on 31/03/2026
'Function ValidateAmount_1_80CCD() As Boolean
'    ValidateAmount_1_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Amount_80CCD1").Cells
'    ReDim Amount_1_80CCD(end_80CCD_1)
'    For i = 1 To end_80CCD_1
'        Amount_1_80CCD(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(Amount_1_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Amount is mandatory in schedule 80CCD(1) at Sr. No " & i & """" & Chr(13)
'            ValidateAmount_1_80CCD = False
'            Exit Function
'        End If
'
'
'        If Not IsNumeric(Amount_1_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Amount in schedule 80CCD(1) at Sr. No  " & i & " should be Numeric value" & Chr(13)
'            ValidateAmount_1_80CCD = False
'            Exit Function
'        End If
'
'        If Amount_1_80CCD(i) > 99999999999999# Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Amount in schedule 80CCD(1) at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
'            ValidateAmount_1_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
''by sadineni on 31/03/2026
'Function ValidateType_1b_80CCD() As Boolean
'    ValidateType_1b_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Type_of_Identifier_80CCD1B").Cells
'
'    ReDim Type_1b_80CCD(end_80CCD_1b)
'    For i = 1 To end_80CCD_1b
'        Type_1b_80CCD(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(Type_1b_80CCD(i)) Then
'
'    'SIT-116970 by sadineni on 06/04/2026
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Type of Identifier is mandatory in schedule 80CCD(1B) at Sr. No " & i & """" & Chr(13)
'            ValidateType_1b_80CCD = False
'            Exit Function
'        End If
'         If Len(Type_1b_80CCD(i)) > 125 Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Type of Identifier in schedule 80CCD(1B) at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateType_1b_80CCD = False
'            Exit Function
'        End If
'
'        If Not checkfieldspecialcharacter(Type_1b_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Type of Identifier in schedule 80CCD(1B) at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateType_1b_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
''by sadineni on 31/03/2026
'Function ValidateName_1b_80CCD() As Boolean
'    ValidateName_1b_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Name_of_Identifier_80CCD1B").Cells
'    ReDim Name_1b_80CCD(end_80CCD_1b)
'    For i = 1 To end_80CCD_1b
'        Name_1b_80CCD(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(Name_1b_80CCD(i)) Then
''SIT-116970 by sadineni on 06/04/2026
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Name of Identifier is mandatory in schedule 80CCD(1B) at Sr. No " & i & """" & Chr(13)
'            ValidateName_1b_80CCD = False
'            Exit Function
'        End If
'         If Len(Name_1b_80CCD(i)) > 125 Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Name of Identifier in schedule 80CCD(1B) length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateName_1b_80CCD = False
'            Exit Function
'        End If
'        If Not checkfieldspecialcharacter(Name_1b_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Name of Identifier in schedule 80CCD(1B) character at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateName_1b_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
''by sadineni on 31/03/2026
'Function ValidateAmount_1b_80CCD() As Boolean
'    ValidateAmount_1b_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Amount_80CCD1B").Cells
'    ReDim Amount_1b_80CCD(end_80CCD_1b)
'    For i = 1 To end_80CCD_1b
'        Amount_1b_80CCD(i) = rangecells.item(i, 1).value
'        If Not chkCompulsory(Amount_1b_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* ""Amount is mandatory in schedule 80CCD(1B) at Sr. No " & i & """" & Chr(13)
'            ValidateAmount_1b_80CCD = False
'            Exit Function
'        End If
'        If Not IsNumeric(Amount_1b_80CCD(i)) Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Amount in schedule 80CCD(1B) at Sr. No  " & i & " should be Numeric value" & Chr(13)
'            ValidateAmount_1b_80CCD = False
'            Exit Function
'        End If
'        If Amount_1b_80CCD(i) > 99999999999999# Then
'            msgValidateVIAerror = msgValidateVIAerror + "* Amount in schedule 80CCD(1B) at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
'            ValidateAmount_1b_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
'
''--end

'************************************----end


'****************added by Chetan C M for AY 2026-27(V 0.3)********************
'--start--
'Chetan C M on 23/04/2026
Function Validate_Pran()
Validate_Pran = True
setTblinfo_pran
If Not Validate_Pran_Number Then Validate_Pran = False
End Function

'Chetan C M on 23/04/2026
Sub setTblinfo_pran()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Pran_Number").count
    Set rangecells = Range("Pran_Number").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_pran = ccount
 End Sub


'Chetan C M on 23/04/2026
Function Validate_Pran_Number() As Boolean
    Validate_Pran_Number = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet19.Range("Pran_Number").Cells
    ReDim pran_80CCC(end_pran)
    If end_pran = 0 Then
            msgValidateVIAerror = msgValidateVIAerror + "* ""PRAN no. shall be mandatory if deduction u/s 80CCD(1) or 80CCD(1B) is claimed""" & Chr(13)
            Validate_Pran_Number = False
            Exit Function
    End If
    For i = 1 To end_pran
        pran_80CCC(i) = rangecells.item(i, 1).value
         If Len(pran_80CCC(i)) > 12 Then
          msgValidateVIAerror = msgValidateVIAerror + "* Please enter a valid 12 digit PRAN at Sr. No " & i & "" & Chr(13)
            Validate_Pran_Number = False
            Exit Function
        End If
        
        'If Not checkfieldspecialcharacter(pran_80CCC(i)) Then
        If Not checkallfieldspecialcharacter_80PRAN(pran_80CCC(i)) Then
            msgValidateVIAerror = msgValidateVIAerror + "* Please enter a valid 12 digit PRAN at Sl.no. " & i & "" & Chr(13)
            Validate_Pran_Number = False

            Exit Function
        End If
     If Not IsNumeric(pran_80CCC(i)) Then
          msgValidateVIAerror = msgValidateVIAerror + "* Please enter a valid 12 digit PRAN at Sl.no. " & i & "" & Chr(13)
            Validate_Pran_Number = False
            
            Exit Function
        End If
         
         
Next
End Function


'************************************----end
