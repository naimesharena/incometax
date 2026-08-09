Attribute VB_Name = "Module2"
'added by Chetan C M on 09/05/2025 for AY 2025-26
'start--
'Sub Addrows80C()
'Dim vRows As Long
'Dim sourceSheet As Worksheet
'Set sourceSheet = ThisWorkbook.Sheets("80C")
'    EfilingCommon.DefinedgridNameRange = "Identification_Number.80C||Amount.80C"
'    ActiveCellRange = EfilingCommon.searchLastRow("Identification_Number.80C")
'    vRows = insertRowUnderSectionWithFormula_80C
'End Sub

Sub Addrows80C()
Sheets("80C").Activate
    EfilingCommon.DefinedgridNameRange = "Identification_Number.80C||Amount.80C"
    ActiveCellRange = EfilingCommon.searchLastRow("Identification_Number.80C")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub Addrows80DA1()
Sheets("80D").Activate
'   EfilingCommon.DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||ReceiptNumA1.80D||AmtA1.80D"
EfilingCommon.DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||AmtA1.80D"
    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerA1.80D")
'    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
    EfilingCommon.insertRowUnderSectionWithFormula

End Sub
Sub Addrows80DB1()
Sheets("80D").Activate
'    EfilingCommon.DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||ReceiptNumB1.80D||AmtB1.80D"
 EfilingCommon.DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||AmtB1.80D"
    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerB1.80D")
'    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
    EfilingCommon.insertRowUnderSectionWithFormula

End Sub
Sub Addrows80DA2()
Sheets("80D").Activate
'    EfilingCommon.DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||ReceiptNumA2.80D||AmtA2.80D"
EfilingCommon.DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||AmtA2.80D"
    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerA2.80D")
'    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
    EfilingCommon.insertRowUnderSectionWithFormula

End Sub
Sub Addrows80DB2()
Sheets("80D").Activate
'    EfilingCommon.DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||ReceiptNumB2.80D||AmtB2.80D"
EfilingCommon.DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||AmtB2.80D"
    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerB2.80D")
'    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
    EfilingCommon.insertRowUnderSectionWithFormula

End Sub
'--end


