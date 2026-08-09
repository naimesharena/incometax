Attribute VB_Name = "Module6"


Sub Addrows24b()
Sheets("Schedule 24(b)").Activate
'Commented by Ayush_05/05/2025 for 0.7 changes
'    mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||IFSC.24b||bankName.24b||PAN.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
     mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
    ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.24b")
    mdCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80E()
Sheets("80E_80EE_80EEA_80EEB").Activate
'Commented by Ayush_05/05/2025 for 0.7 changes
'    mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||IFSC.80E||bankName.80E||PAN.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
     mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||bankName.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
    ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.80E")
    mdCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80EE()
Sheets("80E_80EE_80EEA_80EEB").Activate
'Commented by Ayush_05/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||IFSC.80EE||bankName.80EE||PAN.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
    mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||bankName.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
    ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.80EE")
    mdCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80EEA()
Sheets("80E_80EE_80EEA_80EEB").Activate
'Commented by Ayush_05/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||IFSC.80EEA||bankName.80EEA||PAN.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
    mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||bankName.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
    ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.80EEA")
    mdCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80EEB()
Sheets("80E_80EE_80EEA_80EEB").Activate
'Commented by Ayush_05/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||IFSC.80EEB||bankName.80EEB||PAN.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||Vehicle_value.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    mdCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||bankName.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    ActiveCellRange = mdCommon.searchLastRow("LoanfrmBankOrInstitute.80EEB")
    mdCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80c()
Sheets("80C").Activate
'Commented by Ayush_06/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "NaturePayment.80C||Amount.80C||Identification_Number.80C"
    mdCommon.DefinedgridNameRange = "Amount.80C||Identification_Number.80C"
    ActiveCellRange = mdCommon.searchLastRow("Amount.80C")
    mdCommon.insertRowUnderSectionWithFormula80C HOIflag:=1
End Sub
'Commented by Ayush_06/05/2025 for 0.7 changes
'Sub Addrows()
'Sheets("80C_").Activate
'   mdCommon.DefinedgridNameRange = "Name_of_insurer.||Policy_Document_Number.||Amount."
'    ActiveCellRange = mdCommon.searchLastRow("Name_of_insurer.")
'    mdCommon.insertRowUnderSectionWithFormula80C HOIflag:=1
'End Sub
Sub Addrows80DA1()
Sheets("80D").Activate
'Commented by Ayush_06/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||ReceiptNumA1.80D||AmtA1.80D"
    mdCommon.DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||AmtA1.80D"
    ActiveCellRange = mdCommon.searchLastRow("NameInsurerA1.80D")
    mdCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
Sub Addrows80DB1()
Sheets("80D").Activate
'Commented by Ayush_06/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||ReceiptNumB1.80D||AmtB1.80D"
     mdCommon.DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||AmtB1.80D"
    ActiveCellRange = mdCommon.searchLastRow("NameInsurerB1.80D")
    mdCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
Sub Addrows80DA2()
Sheets("80D").Activate
'Commented by Ayush_06/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||ReceiptNumA2.80D||AmtA2.80D"
    mdCommon.DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||AmtA2.80D"
    ActiveCellRange = mdCommon.searchLastRow("NameInsurerA2.80D")
    mdCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
Sub Addrows80DB2()
Sheets("80D").Activate
'Commented by Ayush_06/05/2025 for 0.7 changes
'   mdCommon.DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||ReceiptNumB2.80D||AmtB2.80D"
    mdCommon.DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||AmtB2.80D"
    ActiveCellRange = mdCommon.searchLastRow("NameInsurerB2.80D")
    mdCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
