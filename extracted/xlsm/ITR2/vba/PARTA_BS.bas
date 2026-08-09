Attribute VB_Name = "PartA_BS"
Option Explicit

Public Sheet2_PropCap As Variant
Public Sheet2_RevResr As Variant
Public Sheet2_CapResr As Variant
Public Sheet2_StatResr As Variant
Public Sheet2_OthResr As Variant
Public Sheet2_TotResrNSurp As Variant
Public Sheet2_TotPropFund As Variant
Public Sheet2_ForeignCurrLoan As Variant
Public Sheet2_FrmBank As Variant
Public Sheet2_FrmOthrs As Variant
Public Sheet2_TotRupeeLoan As Variant
Public Sheet2_TotSecrLoan As Variant
Public Sheet2_FrmBank1 As Variant
Public Sheet2_FrmOthrs1 As Variant
Public Sheet2_TotUnSecrLoan As Variant
Public Sheet2_TotLoanFund As Variant
Public Sheet2_DeferredTax As Variant
Public Sheet2_TotFundSrc As Variant
Public Sheet2_GrossBlock As Variant
Public Sheet2_Depreciation As Variant
Public Sheet2_NetBlock As Variant
Public Sheet2_CapWrkProg As Variant
Public Sheet2_TotFixedAsset As Variant
Public Sheet2_GovtOthSecQuoted As Variant
Public Sheet2_GovOthSecUnQoted As Variant
Public Sheet2_TotLongTermInv As Variant
Public Sheet2_EquityShares As Variant
Public Sheet2_PreferShares As Variant
Public Sheet2_Debenture As Variant
Public Sheet2_TotTradeInv As Variant
Public Sheet2_TotInvestments As Variant
Public Sheet3_StoresConsumables As Variant
Public Sheet3_RawMatl As Variant
Public Sheet3_StkInProcess As Variant
Public Sheet3_FinOrTradGood As Variant
Public Sheet3_TotInventries As Variant
Public Sheet3_SndryDebtors As Variant
Public Sheet3_CashinHand As Variant
Public Sheet3_BankBal As Variant
Public Sheet3_TotCashOrBankBal As Variant
Public Sheet3_OthCurrAsset As Variant
Public Sheet3_TotCurrAsset As Variant
Public Sheet3_AdvRecoverable As Variant
Public Sheet3_Deposits As Variant
Public Sheet3_BalWithRevAuth As Variant
Public Sheet3_TotLoanAdv As Variant
Public Sheet3_TotCurrAssetLoanAdv As Variant
Public Sheet3_SundryCred As Variant
Public Sheet3_LiabForLeasedAsset As Variant
Public Sheet3_AccrIntonLeasedAsset As Variant
Public Sheet3_AccrIntNotDue As Variant
Public Sheet3_TotCurrLiabilities As Variant
Public Sheet3_ITProvision As Variant
Public Sheet3_WTProvision As Variant
Public Sheet3_ELSuperAnnGratProvision As Variant
Public Sheet3_OthProvision As Variant
Public Sheet3_TotProvisions As Variant
Public Sheet3_TotCurrLiabilitiesProvision As Variant
Public Sheet3_NetCurrAsset As Variant
Public Sheet3_MiscExpndr As Variant
Public Sheet3_DefTaxAsset As Variant
Public Sheet3_AccumaltedLosses As Variant
Public Sheet3_TotMiscAdjust As Variant
Public Sheet3_TotFundApply As Variant
Public Sheet3_TotSundryDbtAmt As Variant
Public Sheet3_TotSundryCrdAmt As Variant
Public Sheet3_TotStkInTradAmt As Variant
Public Sheet3_CashBalAmt As Variant



Dim MsgBx_Funds, MsgBx_NOacct As String

Dim Count_NoAccount As Long

Sub CmdValidate_BS_Click()
ValidateBS_All
fmsgboxStatus "Sheet Balance Sheet is ok"
 End Sub

Sub ValidateBS_All()

InitializeValues
    If Not ValidateFunds_BS Then
        Sheet3.Activate
        fmsgboxsmall (MsgBx_Funds)
        CloseMsg
    End If

'To check whether both Balance sheet and No account sheet is enetred
   If Not ValidateNoAccounts_BS Then
        Sheet3.Activate
        fmsgboxStatus (MsgBx_NOacct)
        'CloseMsg
    End If
End Sub


Function ValidateFunds_BS() As Boolean
ValidateFunds_BS = True
    If (Sheet3.Range("sheet2.TotFundSrc").value <> Sheet3.Range("sheet3.TotFundApply").value) Then
        ValidateFunds_BS = False
        Sheet3.Activate
        Sheet3.Range("sheet2.TotFundSrc").Select
        MsgBx_Funds = ""
        MsgBx_Funds = "* Source of funds(4) must be equal to Application of funds(5)"
    End If
End Function

Function ValidateNoAccounts_BS() As Boolean
ValidateNoAccounts_BS = True
    Dim rangecells As Range
    Dim i, mIntCtr, countr, count1 As Long
    countr = 0
    count1 = 0
    Set rangecells = Sheet3.Range("NoAccountsCase").Cells
    mIntCtr = Sheet3.Range("NoAccountsCase").count
    For i = 1 To mIntCtr
        If Not rangecells.item(i).value = "" Then
            countr = countr + 1
        End If
    Next
    
    If Not ((Sheet3.Range("sheet2.TotFundSrc").value = 0) Or (Sheet3.Range("sheet3.TotFundApply").value = 0)) Then
        count1 = count1 + 1
    End If
    
    If ((countr > 0) And (count1 > 0)) Then
        ValidateNoAccounts_BS = False
        MsgBx_NOacct = ""
        MsgBx_NOacct = "Warning :: You have filled Both Balance Sheet and No Accounts" & Chr(13)
    End If
End Function

Function InitializeValues()
On Error Resume Next
 Sheet2_PropCap = IIf(Len(Sheet3.Range("sheet2.PropCap").value) = 0, 0, Sheet3.Range("sheet2.PropCap").value)
 Sheet2_RevResr = IIf(Len(Sheet3.Range("sheet2.RevResr").value) = 0, 0, Sheet3.Range("sheet2.RevResr").value)
 Sheet2_CapResr = IIf(Len(Sheet3.Range("sheet2.CapResr").value) = 0, 0, Sheet3.Range("sheet2.CapResr").value)
 Sheet2_StatResr = IIf(Len(Sheet3.Range("sheet2.StatResr").value) = 0, 0, Sheet3.Range("sheet2.StatResr").value)
 Sheet2_OthResr = IIf(Len(Sheet3.Range("sheet2.OthResr").value) = 0, 0, Sheet3.Range("sheet2.OthResr").value)
 Sheet2_TotResrNSurp = IIf(Len(Sheet3.Range("sheet2.TotResrNSurp").value) = 0, 0, Sheet3.Range("sheet2.TotResrNSurp").value)
 Sheet2_TotPropFund = IIf(Len(Sheet3.Range("sheet2.TotPropFund").value) = 0, 0, Sheet3.Range("sheet2.TotPropFund").value)
 Sheet2_ForeignCurrLoan = IIf(Len(Sheet3.Range("sheet2.ForeignCurrLoan").value) = 0, 0, Sheet3.Range("sheet2.ForeignCurrLoan").value)
 Sheet2_FrmBank = IIf(Len(Sheet3.Range("sheet2.FrmBank").value) = 0, 0, Sheet3.Range("sheet2.FrmBank").value)
 Sheet2_FrmOthrs = IIf(Len(Sheet3.Range("sheet2.FrmOthrs").value) = 0, 0, Sheet3.Range("sheet2.FrmOthrs").value)
 Sheet2_TotRupeeLoan = IIf(Len(Sheet3.Range("sheet2.TotRupeeLoan").value) = 0, 0, Sheet3.Range("sheet2.TotRupeeLoan").value)
 Sheet2_TotSecrLoan = IIf(Len(Sheet3.Range("sheet2.TotSecrLoan").value) = 0, 0, Sheet3.Range("sheet2.TotSecrLoan").value)
 Sheet2_FrmBank1 = IIf(Len(Sheet3.Range("sheet2.FrmBank1").value) = 0, 0, Sheet3.Range("sheet2.FrmBank1").value)
 Sheet2_FrmOthrs1 = IIf(Len(Sheet3.Range("sheet2.FrmOthrs1").value) = 0, 0, Sheet3.Range("sheet2.FrmOthrs1").value)
 Sheet2_TotUnSecrLoan = IIf(Len(Sheet3.Range("sheet2.TotUnSecrLoan").value) = 0, 0, Sheet3.Range("sheet2.TotUnSecrLoan").value)
 Sheet2_TotLoanFund = IIf(Len(Sheet3.Range("sheet2.TotLoanFund").value) = 0, 0, Sheet3.Range("sheet2.TotLoanFund").value)
 Sheet2_DeferredTax = IIf(Len(Sheet3.Range("sheet2.DeferredTax").value) = 0, 0, Sheet3.Range("sheet2.DeferredTax").value)
 Sheet2_TotFundSrc = IIf(Len(Sheet3.Range("sheet2.TotFundSrc").value) = 0, 0, Sheet3.Range("sheet2.TotFundSrc").value)
 Sheet2_GrossBlock = IIf(Len(Sheet3.Range("sheet2.GrossBlock").value) = 0, 0, Sheet3.Range("sheet2.GrossBlock").value)
 Sheet2_Depreciation = IIf(Len(Sheet3.Range("sheet2.Depreciation").value) = 0, 0, Sheet3.Range("sheet2.Depreciation").value)
 Sheet2_NetBlock = IIf(Len(Sheet3.Range("sheet2.NetBlock").value) = 0, 0, Sheet3.Range("sheet2.NetBlock").value)
 Sheet2_CapWrkProg = IIf(Len(Sheet3.Range("sheet2.CapWrkProg").value) = 0, 0, Sheet3.Range("sheet2.CapWrkProg").value)
 Sheet2_TotFixedAsset = IIf(Len(Sheet3.Range("sheet2.TotFixedAsset").value) = 0, 0, Sheet3.Range("sheet2.TotFixedAsset").value)
 Sheet2_GovtOthSecQuoted = IIf(Len(Sheet3.Range("sheet2.GovtOthSecQuoted").value) = 0, 0, Sheet3.Range("sheet2.GovtOthSecQuoted").value)
 Sheet2_GovOthSecUnQoted = IIf(Len(Sheet3.Range("sheet2.GovOthSecUnQoted").value) = 0, 0, Sheet3.Range("sheet2.GovOthSecUnQoted").value)
 Sheet2_TotLongTermInv = IIf(Len(Sheet3.Range("sheet2.TotLongTermInv").value) = 0, 0, Sheet3.Range("sheet2.TotLongTermInv").value)
 Sheet2_EquityShares = IIf(Len(Sheet3.Range("sheet2.EquityShares").value) = 0, 0, Sheet3.Range("sheet2.EquityShares").value)
 Sheet2_PreferShares = IIf(Len(Sheet3.Range("sheet2.PreferShares").value) = 0, 0, Sheet3.Range("sheet2.PreferShares").value)
 Sheet2_Debenture = IIf(Len(Sheet3.Range("sheet2.Debenture").value) = 0, 0, Sheet3.Range("sheet2.Debenture").value)
 Sheet2_TotTradeInv = IIf(Len(Sheet3.Range("sheet2.TotTradeInv").value) = 0, 0, Sheet3.Range("sheet2.TotTradeInv").value)
 Sheet2_TotInvestments = IIf(Len(Sheet3.Range("sheet2.TotInvestments").value) = 0, 0, Sheet3.Range("sheet2.TotInvestments").value)
 Sheet3_StoresConsumables = IIf(Len(Sheet3.Range("sheet3.StoresConsumables").value) = 0, 0, Sheet3.Range("sheet3.StoresConsumables").value)
 Sheet3_RawMatl = IIf(Len(Sheet3.Range("sheet3.RawMatl").value) = 0, 0, Sheet3.Range("sheet3.RawMatl").value)
 Sheet3_StkInProcess = IIf(Len(Sheet3.Range("sheet3.StkInProcess").value) = 0, 0, Sheet3.Range("sheet3.StkInProcess").value)
 Sheet3_FinOrTradGood = IIf(Len(Sheet3.Range("sheet3.FinOrTradGood").value) = 0, 0, Sheet3.Range("sheet3.FinOrTradGood").value)
 Sheet3_TotInventries = IIf(Len(Sheet3.Range("sheet3.TotInventries").value) = 0, 0, Sheet3.Range("sheet3.TotInventries").value)
 Sheet3_SndryDebtors = IIf(Len(Sheet3.Range("sheet3.SndryDebtors").value) = 0, 0, Sheet3.Range("sheet3.SndryDebtors").value)
 Sheet3_CashinHand = IIf(Len(Sheet3.Range("sheet3.CashinHand").value) = 0, 0, Sheet3.Range("sheet3.CashinHand").value)
 Sheet3_BankBal = IIf(Len(Sheet3.Range("sheet3.BankBal").value) = 0, 0, Sheet3.Range("sheet3.BankBal").value)
 Sheet3_TotCashOrBankBal = IIf(Len(Sheet3.Range("sheet3.TotCashOrBankBal").value) = 0, 0, Sheet3.Range("sheet3.TotCashOrBankBal").value)
 Sheet3_OthCurrAsset = IIf(Len(Sheet3.Range("sheet3.OthCurrAsset").value) = 0, 0, Sheet3.Range("sheet3.OthCurrAsset").value)
 Sheet3_TotCurrAsset = IIf(Len(Sheet3.Range("sheet3.TotCurrAsset").value) = 0, 0, Sheet3.Range("sheet3.TotCurrAsset").value)
 Sheet3_AdvRecoverable = IIf(Len(Sheet3.Range("sheet3.AdvRecoverable").value) = 0, 0, Sheet3.Range("sheet3.AdvRecoverable").value)
 Sheet3_Deposits = IIf(Len(Sheet3.Range("sheet3.Deposits").value) = 0, 0, Sheet3.Range("sheet3.Deposits").value)
 Sheet3_BalWithRevAuth = IIf(Len(Sheet3.Range("Sheet3.BalWithRevAuth").value) = 0, 0, Sheet3.Range("Sheet3.BalWithRevAuth").value)
 Sheet3_TotLoanAdv = IIf(Len(Sheet3.Range("Sheet3.TotLoanAdv").value) = 0, 0, Sheet3.Range("Sheet3.TotLoanAdv").value)
 Sheet3_TotCurrAssetLoanAdv = IIf(Len(Sheet3.Range("sheet3.TotCurrAssetLoanAdv").value) = 0, 0, Sheet3.Range("sheet3.TotCurrAssetLoanAdv").value)
 Sheet3_SundryCred = IIf(Len(Sheet3.Range("sheet3.SundryCred").value) = 0, 0, Sheet3.Range("sheet3.SundryCred").value)
 Sheet3_LiabForLeasedAsset = IIf(Len(Sheet3.Range("sheet3.LiabForLeasedAsset").value) = 0, 0, Sheet3.Range("sheet3.LiabForLeasedAsset").value)
 Sheet3_AccrIntonLeasedAsset = IIf(Len(Sheet3.Range("sheet3.AccrIntonLeasedAsset").value) = 0, 0, Sheet3.Range("sheet3.AccrIntonLeasedAsset").value)
 Sheet3_AccrIntNotDue = IIf(Len(Sheet3.Range("sheet3.AccrIntNotDue").value) = 0, 0, Sheet3.Range("sheet3.AccrIntNotDue").value)
 Sheet3_TotCurrLiabilities = IIf(Len(Sheet3.Range("sheet3.TotCurrLiabilities").value) = 0, 0, Sheet3.Range("sheet3.TotCurrLiabilities").value)
 Sheet3_ITProvision = IIf(Len(Sheet3.Range("sheet3.ITProvision").value) = 0, 0, Sheet3.Range("sheet3.ITProvision").value)
 'Sheet3_WTProvision = IIf(Len(Sheet3.Range("sheet3.WTProvision").value) = 0, 0, Sheet3.Range("sheet3.WTProvision").value)
 Sheet3_ELSuperAnnGratProvision = IIf(Len(Sheet3.Range("sheet3.ELSuperAnnGratProvision").value) = 0, 0, Sheet3.Range("sheet3.ELSuperAnnGratProvision").value)
 Sheet3_OthProvision = IIf(Len(Sheet3.Range("sheet3.OthProvision").value) = 0, 0, Sheet3.Range("sheet3.OthProvision").value)
 Sheet3_TotProvisions = IIf(Len(Sheet3.Range("sheet3.TotProvisions").value) = 0, 0, Sheet3.Range("sheet3.TotProvisions").value)
 Sheet3_TotCurrLiabilitiesProvision = IIf(Len(Sheet3.Range("sheet3.TotCurrLiabilitiesProvision").value) = 0, 0, Sheet3.Range("sheet3.TotCurrLiabilitiesProvision").value)
 Sheet3_NetCurrAsset = IIf(Len(Sheet3.Range("sheet3.NetCurrAsset").value) = 0, 0, Sheet3.Range("sheet3.NetCurrAsset").value)
 Sheet3_MiscExpndr = IIf(Len(Sheet3.Range("sheet3.MiscExpndr").value) = 0, 0, Sheet3.Range("sheet3.MiscExpndr").value)
 Sheet3_DefTaxAsset = IIf(Len(Sheet3.Range("sheet3.DefTaxAsset").value) = 0, 0, Sheet3.Range("sheet3.DefTaxAsset").value)
 Sheet3_AccumaltedLosses = IIf(Len(Sheet3.Range("sheet3.AccumaltedLosses").value) = 0, 0, Sheet3.Range("sheet3.AccumaltedLosses").value)
 Sheet3_TotMiscAdjust = IIf(Len(Sheet3.Range("sheet3.TotMiscAdjust").value) = 0, 0, Sheet3.Range("sheet3.TotMiscAdjust").value)
 Sheet3_TotFundApply = IIf(Len(Sheet3.Range("sheet3.TotFundApply").value) = 0, 0, Sheet3.Range("sheet3.TotFundApply").value)
 Sheet3_TotSundryDbtAmt = IIf(Len(Sheet3.Range("sheet3.TotSundryDbtAmt").value) = 0, 0, Sheet3.Range("sheet3.TotSundryDbtAmt").value)
 Sheet3_TotSundryCrdAmt = IIf(Len(Sheet3.Range("sheet3.TotSundryCrdAmt").value) = 0, 0, Sheet3.Range("sheet3.TotSundryCrdAmt").value)
 Sheet3_TotStkInTradAmt = IIf(Len(Sheet3.Range("sheet3.TotStkInTradAmt").value) = 0, 0, Sheet3.Range("sheet3.TotStkInTradAmt").value)
 Sheet3_CashBalAmt = IIf(Len(Sheet3.Range("sheet3.CashBalAmt").value) = 0, 0, Sheet3.Range("sheet3.CashBalAmt").value)
End Function

