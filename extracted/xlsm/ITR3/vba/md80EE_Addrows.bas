Attribute VB_Name = "md80EE_Addrows"
Sub AddRows80E()
    Sheets("80E_80EE_80EEA_80EEB").Activate
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||bankName.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E||"
 ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80E")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows80EE()
    Sheets("80E_80EE_80EEA_80EEB").Activate
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||bankName.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||"
 ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EE")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows80EEA()
    Sheets("80E_80EE_80EEA_80EEB").Activate
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||bankName.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||"
 ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EEA")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows80EEB()
    Sheets("80E_80EE_80EEA_80EEB").Activate
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||bankName.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||VehicleRegNum.80EEB||Intrst.80EEB||"
 ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EEB")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub
