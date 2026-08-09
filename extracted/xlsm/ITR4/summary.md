# ITR4 — Excel utility extraction

Source: `ITR4_AY_26-27_V1.1.xlsm` (5,883,165 bytes)

- Worksheets: **24** (7 visible, 17 hidden)
- Sheet contents (hidden + visible): **24 dumps** in `sheets/` (186,665 stored cell values, formulas extracted alongside)
- VBA modules: **79**, 96,921 lines of code
- Data-validation (dropdown) rules: **522** (resolved to actual list values in `dropdowns.json`)
- Defined names: **1493**

## Worksheets

| # | Sheet | State | Dimensions | Formulas | Validations |
|---|-------|-------|-----------|----------|-------------|
| 1 | Income Details | visible | A1:XEN36406 | 169 | 190 |
| 2 | HP | visible | A1:XFC243 | 72 | 52 |
| 3 | Sheet1 | hidden | J3:K3 | 0 | 0 |
| 4 | 44AE | hidden | B1:IU43 | 12 | 4 |
| 5 | Part A Gen_139(8A) | hidden | A1:XFC38 | 8 | 15 |
| 6 | Schedule EA 10(13A) | hidden | A1:O40 | 3 | 2 |
| 7 | Schedule 24(b) | hidden | A1:XFC24 | 10 | 7 |
| 8 | BP | visible | A1:XFC57158 | 69 | 47 |
| 9 | TDS | visible | A1:XFC65531 | 35 | 35 |
| 10 | TCS | visible | A1:M65460 | 5 | 8 |
| 11 | IT | visible | A1:AA99 | 93 | 7 |
| 12 | Part B ATI | hidden | A1:XFC52 | 23 | 12 |
| 13 | Taxes Paid and Verification | visible | A1:XFC65552 | 22 | 32 |
| 14 | 80D | hidden | A1:XFC98 | 21 | 14 |
| 15 | 80G | hidden | A1:AA394 | 196 | 31 |
| 16 | 80DD_80U | hidden | A1:T21 | 8 | 10 |
| 17 | 80GGC | hidden | A1:XFC22 | 29 | 10 |
| 18 | 80E_80EE_80EEA_80EEB | hidden | B1:X58 | 25 | 13 |
| 19 | 80C | hidden | A1:N32 | 8 | 6 |
| 20 | AL | hidden | A1:IU65536 | 6 | 27 |
| 21 | SUMMARY | hidden | A1:D73 | 11 | 0 |
| 22 | Help | hidden | A1:XFC85 | 0 | 0 |
| 23 | DB | hidden | A1:GD65539 | 10 | 0 |
| 24 | TaxCalc | hidden | A1:CL124 | 710 | 0 |

## VBA modules

| Module | Lines | File |
|--------|-------|------|
| AY23_24Changes.bas | 2989 | `vba/AY23_24Changes.bas` |
| AY25_26Changes.bas | 1473 | `vba/AY25_26Changes.bas` |
| ePayPrefill.bas | 12 | `vba/ePayPrefill.bas` |
| GenerateJSON.bas | 8257 | `vba/GenerateJSON.bas` |
| HS256.cls | 533 | `vba/HS256.cls` |
| ImportJSON.bas | 9639 | `vba/ImportJSON.bas` |
| ImportXML.bas | 4373 | `vba/ImportXML.bas` |
| md10_13A.bas | 321 | `vba/md10_13A.bas` |
| md44AE.bas | 121 | `vba/md44AE.bas` |
| md80C.bas | 215 | `vba/md80C.bas` |
| md80CCC.bas | 193 | `vba/md80CCC.bas` |
| md80E.bas | 822 | `vba/md80E.bas` |
| md80EE.bas | 1167 | `vba/md80EE.bas` |
| md80EEA.bas | 1312 | `vba/md80EEA.bas` |
| md80EEB.bas | 931 | `vba/md80EEB.bas` |
| md80G.bas | 998 | `vba/md80G.bas` |
| md80GGC.bas | 867 | `vba/md80GGC.bas` |
| mdAL.bas | 1428 | `vba/mdAL.bas` |
| mdATI.bas | 736 | `vba/mdATI.bas` |
| mdCommon.bas | 2915 | `vba/mdCommon.bas` |
| mdGen139_8A.bas | 367 | `vba/mdGen139_8A.bas` |
| mdHashing.bas | 172 | `vba/mdHashing.bas` |
| mdHouseProperty.bas | 3206 | `vba/mdHouseProperty.bas` |
| mdImportXL.bas | 1222 | `vba/mdImportXL.bas` |
| mdIncomeDetails.bas | 6745 | `vba/mdIncomeDetails.bas` |
| mdInitializer.bas | 91 | `vba/mdInitializer.bas` |
| mdInt24b.bas | 911 | `vba/mdInt24b.bas` |
| mdIT.bas | 9 | `vba/mdIT.bas` |
| mdNOBBP.bas | 1791 | `vba/mdNOBBP.bas` |
| mdPAN.bas | 28 | `vba/mdPAN.bas` |
| mdProgressBar.bas | 110 | `vba/mdProgressBar.bas` |
| mdTaxCalc.bas | 1893 | `vba/mdTaxCalc.bas` |
| mdTaxPaidVerification.bas | 823 | `vba/mdTaxPaidVerification.bas` |
| mdTCS.bas | 20 | `vba/mdTCS.bas` |
| mdTDS.bas | 4378 | `vba/mdTDS.bas` |
| mdTDS2.bas | 2841 | `vba/mdTDS2.bas` |
| mdTDS3.bas | 2781 | `vba/mdTDS3.bas` |
| MessageBox.frm | 13 | `vba/MessageBox.frm` |
| mfmessage.bas | 38 | `vba/mfmessage.bas` |
| Module1.bas | 12 | `vba/Module1.bas` |
| Module2.bas | 15 | `vba/Module2.bas` |
| Module3.bas | 8 | `vba/Module3.bas` |
| Module4.bas | 8 | `vba/Module4.bas` |
| Module5.bas | 8 | `vba/Module5.bas` |
| Module6.bas | 91 | `vba/Module6.bas` |
| PreFillJson.bas | 6234 | `vba/PreFillJson.bas` |
| pwd.bas | 43 | `vba/pwd.bas` |
| Sch80D.bas | 1093 | `vba/Sch80D.bas` |
| Sch80DD_80U.bas | 1136 | `vba/Sch80DD_80U.bas` |
| SchBA.bas | 1319 | `vba/SchBA.bas` |
| SchDI.bas | 269 | `vba/SchDI.bas` |
| Sheet1.cls | 10929 | `vba/Sheet1.cls` |
| Sheet10.cls | 47 | `vba/Sheet10.cls` |
| Sheet11.cls | 431 | `vba/Sheet11.cls` |
| Sheet12.cls | 19 | `vba/Sheet12.cls` |
| Sheet13.cls | 261 | `vba/Sheet13.cls` |
| Sheet14.cls | 357 | `vba/Sheet14.cls` |
| Sheet15.cls | 89 | `vba/Sheet15.cls` |
| Sheet16.cls | 637 | `vba/Sheet16.cls` |
| Sheet17.cls | 71 | `vba/Sheet17.cls` |
| Sheet18.cls | 215 | `vba/Sheet18.cls` |
| Sheet19.cls | 33 | `vba/Sheet19.cls` |
| Sheet2.cls | 10 | `vba/Sheet2.cls` |
| Sheet20.cls | 1879 | `vba/Sheet20.cls` |
| Sheet201.cls | 672 | `vba/Sheet201.cls` |
| Sheet202.cls | 384 | `vba/Sheet202.cls` |
| Sheet21.cls | 438 | `vba/Sheet21.cls` |
| Sheet22.cls | 597 | `vba/Sheet22.cls` |
| Sheet3.cls | 198 | `vba/Sheet3.cls` |
| Sheet4.cls | 593 | `vba/Sheet4.cls` |
| Sheet5.cls | 888 | `vba/Sheet5.cls` |
| Sheet6.cls | 987 | `vba/Sheet6.cls` |
| Sheet7.cls | 9 | `vba/Sheet7.cls` |
| Sheet8.cls | 16 | `vba/Sheet8.cls` |
| Sheet9.cls | 19 | `vba/Sheet9.cls` |
| ThisWorkbook.cls | 114 | `vba/ThisWorkbook.cls` |
| UserForm1.frm | 17 | `vba/UserForm1.frm` |
| UserForm2.frm | 17 | `vba/UserForm2.frm` |
| UserForm3.frm | 17 | `vba/UserForm3.frm` |

## Defined names

- `_44AD_Income` → TaxCalc!$T$17
- `_44AE_income` → TaxCalc!$T$16
- `_Final_44AD` → TaxCalc!$T$19
- `_NOB1` → DB!$CA$1:$CA$345
- `_xleta.CONCATENATE` → #NAME?
- `_xleta.IF` → #NAME?
- `_xleta.LEFT` → #NAME?
- `_xleta.MAX` → #NAME?
- `_xleta.MID` → #NAME?
- `_xleta.OR` → #NAME?
- `_xleta.SEC` → #NAME?
- `_xleta.SUM` → #NAME?
- `A1123456` → BP!$A$1
- `AadhaarDependent_80DD` → '80DD_80U'!$H$6
- `AadharHUF` → DB!#REF!
- `AadharIND` → DB!#REF!
- `AadharSelect` → DB!#REF!
- `AccntStatus_BA` → DB!$BQ$1:$BQ$3
- `AcknowledgeNum_80DD` → '80DD_80U'!$I$6
- `AcknowledgeNum11A2_80DD` → '80DD_80U'!#REF!
- `AcknowledgeNum11A2_80U` → '80DD_80U'!#REF!
- `AckNumber10IA_80U` → '80DD_80U'!$F$16
- `AE44.HEE3` → '44AE'!$G$9:$G$18
- `AE44.HEE4` → #REF!
- `AE44.TOTAL` → #REF!
- `AfterDateReturn` → DB!$CQ$3:$CQ$5
- `Age` → TaxCalc!$G$8
- `Agri_dropdown` → DB!$FW$1:$FW$4
- `Agri_income` → 'Taxes Paid and Verification'!$Q$13
- `AIR.TaxExmpIntInc` → 'Taxes Paid and Verification'!#REF!
- `AL.ArchCollDrawPaintSulpArt` → #REF!
- `AL.CashInHand` → #REF!
- `AL.DepositsInBank` → #REF!
- `AL.ImmovableAssetBuilding` → #REF!
- `AL.ImmovableAssetLand` → #REF!
- `AL.InsurancePolicies` → #REF!
- `AL.JewelleryBullionEtc` → #REF!
- `AL.LiabilityInRelatAssets` → #REF!
- `AL.LoansAndAdvancesGiven` → #REF!
- `AL.SharesAndSecurities` → #REF!
- `al.TotalImmovablMovablAssets` → #REF!
- `AL.VehiclYachtsBoatsAircrafts` → #REF!
- `All_Pincode_list` → DB!$EP$2:$EP$19303
- `All_Pincode_V` → DB!$EP$2:$EQ$19303
- `Amount.80C` → '80C'!$F$5:$F$8
- `Amount.80CCC` → '80C'!$G$23:$G$26
- `AmountDeduction_80U` → '80DD_80U'!$E$16
- `AmtA1.80D` → '80D'!$L$9:$L$12
- `AmtA2.80D` → '80D'!$L$33:$L$36
- `AmtB1.80D` → '80D'!$L$20:$L$23
- `AmtB2.80D` → '80D'!$L$44:$L$47
- `AmtDeduction_80DD` → '80DD_80U'!$E$6
- `AssessedTax` → TaxCalc!$I$5
- `AssessedTax_1` → TaxCalc!$J$44
- `AssessedTax_2` → TaxCalc!$L$44
- `Avalue` → 'Income Details'!$AZ$236
- `AY_80GGC` → DB!$C$28
- `AY23Return` → DB!$CS$2:$CS$3
- `AYinOpting` → 'Income Details'!$AO$76
- `AYinOptOut` → 'Income Details'!$AO$80
- `BAC115.A23b` → 'Income Details'!$AO$62
- `BAC115.A23ii_No` → 'Income Details'!$AO$56
- `BAC115.AcknowledgementNo` → 'Income Details'!$AN$50
- `BAC115.AcknowledgementNo_New_NA` → 'Income Details'!$AN$70
- `BAC115.AcknowledgementNo_New_No` → 'Income Details'!$AN$60
- `BAC115.AcknowledgementNo_New_Yes` → 'Income Details'!$AN$55
- `BAC115.AssessmentYear_2425` → 'Income Details'!$AO$49
- `BAC115.AssessmentYear_Aii` → 'Income Details'!$AO$54
- `BAC115.DateOfFiling` → 'Income Details'!$AO$48
- `BAC115.DateOfFiling_New_NA` → 'Income Details'!$AO$69
- `BAC115.DateOfFiling_New_No` → 'Income Details'!$AO$59
- `BAC115.DateOfFiling_New_Yes` → 'Income Details'!$AO$53
- `BAC115.NA` → DB!$DL$8
- `BAC115.NA_New` → DB!$DQ$4:$DQ$6
- `BAC115.No_New` → DB!$DP$4:$DP$6
- `BAC115.NY` → DB!$DL$4:$DL$6
- `BAC115.Yes_New` → DB!$DO$4:$DO$6
- `bacValue` → DB!$BN$33
- `bacValue_New` → 'Income Details'!$BB$50
- `Balance_Interest` → TaxCalc!$CD$6:$CD$124
- `bankName.24b` → 'Schedule 24(b)'!$G$4:$G$8
- `bankName.24b1` → HP!$F$36:$F$40
- `bankName.24b2` → HP!$F$77:$F$81
- `bankName.80E` → '80E_80EE_80EEA_80EEB'!$G$4:$G$7
- `bankName.80EE` → '80E_80EE_80EEA_80EEB'!$G$17:$G$20
- `bankName.80EEA` → '80E_80EE_80EEA_80EEB'!$G$30:$G$33
- `bankName.80EEB` → '80E_80EE_80EEA_80EEB'!$G$42:$G$45
- `bankpayment` → IT!$D$4
- `BeforeDateReturn` → DB!$CO$3:$CO$4
- `Block_1B` → HP!$P$48
- `BP.GSTRAmount` → BP!$F$75:$F$79
- `BP.GSTRno` → BP!$E$75:$E$79
- `BP.NatureAmount` → BP!$G$86:$G$89
- `BP.NatureDescription` → BP!$F$86:$F$89
- `BP.NatureItem` → BP!$E$86:$E$89
- `BP_E1` → BP!$H$19
- `BP_E1_GR` → BP!$I$19
- `BP_E10` → BP!$I$95
- `BP_E11` → BP!$I$96
- `BP_E12` → BP!$I$97
- `BP_E13` → BP!$I$98
- `BP_E14` → BP!$I$99
- `BP_E15_total` → BP!$I$100
- `BP_E16` → BP!$I$101
- `BP_E17` → BP!$I$103
- `BP_E18` → BP!$I$104
- `BP_E18a` → BP!$I$102
- `BP_E19` → BP!$I$105
- `BP_E1a` → BP!$I$20
- `BP_E1a_E1b` → BP!$I$21
- `BP_E1b` → BP!$I$22
- `BP_E2` → BP!$I$23
- `BP_E2_PI` → BP!$I$26
- `BP_E20` → BP!$I$106
- `BP_E21` → BP!$I$107
- `BP_E22` → BP!$I$108
- `BP_E23_total` → BP!$I$109
- `BP_E2a` → BP!$I$24
- `BP_E2b` → BP!$I$25
- `BP_E3` → BP!$I$65
- `BP_E3_2` → BP!$I$69
- `BP_E3_a` → BP!$I$38
- `BP_E3_b` → BP!$I$39
- `BP_E3_c` → BP!$I$40
- `BP_E3_GR` → BP!$I$37
- `BP_E4` → BP!#REF!
- `BP_E4_2` → BP!$H$68
- `BP_E4_PI44ADA` → BP!$I$41
- `BP_E5` → BP!#REF!
- `BP_E6` → BP!$H$82
- `BP_E7_IC44AE` → BP!$I$71
- `BP_E8_IC_BP` → BP!$I$72
- `BP_E9` → BP!$I$94
- `BS_A_E13` → BP!#REF!
- `BS_A_E14` → BP!#REF!
- `BS_A_E15` → BP!#REF!
- `BS_A_E16` → BP!#REF!
- `BS_A_E17` → BP!#REF!
- `BS_A_E18` → BP!#REF!
- `BS_A_E19` → BP!#REF!
- `BS_A_E20` → BP!#REF!
- `BS_B_E22` → BP!#REF!
- `BS_B_E23` → BP!#REF!
- `BS_B_E24` → BP!#REF!
- `BS_B_E25` → BP!#REF!
- `BS_B_E26` → BP!#REF!
- `BS_B_E27` → BP!#REF!
- `BS_B_E28` → BP!#REF!
- `BS_B_E29` → BP!#REF!
- `Button_epaytax` → 'Taxes Paid and Verification'!$H$9
- `Bvalue` → 'Income Details'!$AZ$237
- `C_Eligible` → '80G'!$T$3
- `CalculatedValue234A` → TaxCalc!$E$2
- `CalculatedValue234B` → TaxCalc!$E$3
- `CalculatedValue234C` → TaxCalc!$E$4
- `Capacity` → DB!$AC$1:$AC$5
- `CD_EligibleAmount` → '80G'!$W$3
- `CDE_EligibleAmount` → '80G'!$X$3
- `clauseiv7provisio139iAmount_1` → 'Income Details'!$AO$94
- `clauseiv7provisio139iAmount_2` → 'Income Details'!$AO$95
- `clauseiv7provisio139iAmount_3` → 'Income Details'!$AO$96
- `clauseiv7provisio139iAmount_4` → 'Income Details'!$AO$97
- `clauseiv7provisio139iFlg` → 'Income Details'!$AO$93
- `clauseiv7provisio139iFlg_1` → 'Income Details'!$AH$94
- `clauseiv7provisio139iFlg_2` → 'Income Details'!$AH$95
- `clauseiv7provisio139iFlg_3` → 'Income Details'!$AH$96
- `clauseiv7provisio139iFlg_4` → 'Income Details'!$AH$97
- `Co_OwnnerBlock1` → HP!$G$9:$J$18
- `Co_OwnnerBlock2` → HP!$G$50:$J$59
- `Column6_value` → #REF!
- `Column6_value3` → #REF!
- `comb_80G_A` → '80G'!$U$7:$U$12
- `comb_80G_B` → '80G'!$U$21:$U$26
- `comb_80G_C` → '80G'!$U$35:$U$40
- `comb_80G_D` → '80G'!$U$49:$U$54
- `comb_90` → '80G'!$W$7:$W$12
- `comb_91` → '80G'!$W$21
- `Comb_donation_80G_A` → '80G'!$V$7:$V$12
- `Comb_donation_80G_B` → '80G'!$V$21:$V$26
- `Comb_donation_80G_C` → '80G'!$V$35:$V$40
- `Comb_donation_80G_D` → '80G'!$V$49:$V$54
- `Comb_PAN_80G_A` → '80G'!$X$7:$X$12
- `Comb_PAN_80G_B` → '80G'!$X$21:$X$26
- `Comb_PAN_80G_C` → '80G'!$X$35:$X$40
- `Comb_PAN_80G_D` → '80G'!$X$49:$X$54
- `Combination_24B1` → HP!$R$36:$R$40
- `Combination_24B2` → HP!$R$77:$R$81
- `Combination_80EE` → '80E_80EE_80EEA_80EEB'!$X$17:$X$20
- `Combination_80EEA` → '80E_80EE_80EEA_80EEB'!$W$30:$W$33
- `Comp_dropdown` → DB!$FX$1:$FX$4
- `CONSUMED_TOTAL_INCOME` → 'Income Details'!$AZ$173
- `Count1` → 'Income Details'!$AY$112
- `Count2` → 'Income Details'!$AY$113
- `Count3` → 'Income Details'!$AY$114
- `Count4` → 'Income Details'!$AY$115
- `CountOS1` → 'Income Details'!$AZ$112
- `CountOS2` → 'Income Details'!$AZ$113
- `CountOS3` → 'Income Details'!$AZ$114
- `CountOS4` → 'Income Details'!$AZ$115
- `CountOS5` → 'Income Details'!$AZ$120

## Sheet dumps (`sheets/`) — hidden and visible

| Sheet | State | Rows with values | Cells | Formulas | File |
|-------|-------|------------------|-------|----------|------|
| Income Details | visible | 228 | 952 | 169 | `sheets/01_Income_Details.tsv` |
| HP | visible | 86 | 248 | 46 | `sheets/02_HP.tsv` |
| Sheet1 | hidden | 1 | 2 | 0 | `sheets/03_Sheet1.tsv` |
| 44AE | hidden | 19 | 36 | 3 | `sheets/04_44AE.tsv` |
| Part A Gen_139(8A) | hidden | 24 | 42 | 8 | `sheets/05_Part_A_Gen_139_8A.tsv` |
| Schedule EA 10(13A) | hidden | 10 | 24 | 3 | `sheets/06_Schedule_EA_10_13A.tsv` |
| Schedule 24(b) | hidden | 8 | 29 | 10 | `sheets/07_Schedule_24_b.tsv` |
| BP | visible | 98 | 241 | 46 | `sheets/08_BP.tsv` |
| TDS | visible | 52 | 181 | 35 | `sheets/09_TDS.tsv` |
| TCS | visible | 10 | 23 | 5 | `sheets/10_TCS.tsv` |
| IT | visible | 25 | 119 | 93 | `sheets/11_IT.tsv` |
| Part B ATI | hidden | 40 | 134 | 23 | `sheets/12_Part_B_ATI.tsv` |
| Taxes Paid and Verification | visible | 45 | 105 | 22 | `sheets/13_Taxes_Paid_and_Verification.tsv` |
| 80D | hidden | 46 | 90 | 21 | `sheets/14_80D.tsv` |
| 80G | hidden | 81 | 337 | 75 | `sheets/15_80G.tsv` |
| 80DD_80U | hidden | 9 | 33 | 8 | `sheets/16_80DD_80U.tsv` |
| 80GGC | hidden | 10 | 54 | 13 | `sheets/17_80GGC.tsv` |
| 80E_80EE_80EEA_80EEB | hidden | 30 | 84 | 25 | `sheets/18_80E_80EE_80EEA_80EEB.tsv` |
| 80C | hidden | 14 | 24 | 8 | `sheets/19_80C.tsv` |
| AL | hidden | 30 | 64 | 6 | `sheets/20_AL.tsv` |
| SUMMARY | hidden | 16 | 32 | 11 | `sheets/21_SUMMARY.tsv` |
| Help | hidden | 68 | 96 | 0 | `sheets/22_Help.tsv` |
| DB | hidden | 40439 | 182994 | 4 | `sheets/23_DB.tsv.gz` |
| TaxCalc | hidden | 124 | 721 | 366 | `sheets/24_TaxCalc.tsv` |
