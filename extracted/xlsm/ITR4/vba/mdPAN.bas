Attribute VB_Name = "mdPAN"

Sub AddRowsPAN()
Dim vRows  As Long
'Sheets("PAN").Activate
mdCommon.DefinedgridNameRange = "TAN_1||TDSoth.TDS_SIncome||TDSoth.TDS_TDS||TDSoth.TDS_PAN"
ActiveCellRange = mdCommon.searchLastRow("TAN_1")
vRows = mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRowsPAN2()
Dim vRows  As Long
'Sheets("PAN").Activate
mdCommon.DefinedgridNameRange = "PAN_2||TDSoth2.TDS_SIncome||TDSoth2.TDS_TDS||TDSoth2.TDS_PAN"
ActiveCellRange = mdCommon.searchLastRow("PAN_2")
vRows = mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub GotoTDS()
Sheets("TDS").Activate
End Sub

Sub AddRowsSchTDS2PAN()
Dim vRows  As Long
Sheet16.Visible = True
Sheets("PAN").Activate
End Sub
