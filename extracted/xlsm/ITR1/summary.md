# ITR1 — Excel utility extraction

Source: `ITR1_AY_26-27_V1.2.xlsm` (4,189,334 bytes)

- Worksheets: **21** (5 visible, 16 hidden)
- Sheet contents (hidden + visible): **21 dumps** in `sheets/` (184,072 stored cell values, formulas extracted alongside)
- VBA modules: **73**, 73,599 lines of code
- Data-validation (dropdown) rules: **387** (resolved to actual list values in `dropdowns.json`)
- Defined names: **1355**

## Worksheets

| # | Sheet | State | Dimensions | Formulas | Validations |
|---|-------|-------|-----------|----------|-------------|
| 1 | Income Details | visible | A1:XFC221 | 156 | 151 |
| 2 | HP | visible | A1:XFC82 | 60 | 47 |
| 3 | Schedule EA 10(13A) | hidden | A1:I13 | 3 | 2 |
| 4 | Schedule 24(b) | hidden | A1:W13 | 8 | 8 |
| 5 | Part A Gen_139(8A) | hidden | A1:XFC36 | 8 | 16 |
| 6 | TDS | visible | A1:XFD65555 | 104 | 21 |
| 7 | TCS | visible | A1:XFC1048576 | 4 | 6 |
| 8 | Taxes Paid and Verification | visible | A1:XFC65576 | 13 | 40 |
| 9 | Part B ATI | hidden | A1:XFC49 | 22 | 11 |
| 10 | 80D | hidden | A1:O56 | 17 | 19 |
| 11 | 80G | hidden | A1:XFC65544 | 157 | 18 |
| 12 | 80GGA | hidden | A1:XFC65524 | 19 | 11 |
| 13 | 80GGC | hidden | A1:XFC20 | 48 | 7 |
| 14 | 80U-80DD | hidden | A1:XFC29 | 7 | 9 |
| 15 | 80C | hidden | A1:XFC25 | 8 | 8 |
| 16 | 80E_80EE_80EEA_80EEB | hidden | A1:XFC46 | 13 | 13 |
| 17 | BankCode | hidden | A1:J318 | 0 | 0 |
| 18 | IFSC | hidden | E3:I45138 | 0 | 0 |
| 19 | DataBase | hidden | A1:GM19303 | 361 | 0 |
| 20 | SUMMARY | hidden | B1:IU47 | 10 | 0 |
| 21 | Help | hidden | A1:IT147 | 0 | 0 |

## VBA modules

| Module | Lines | File |
|--------|-------|------|
| AY23_24Changes.bas | 1736 | `vba/AY23_24Changes.bas` |
| checkBoxModule.bas | 29 | `vba/checkBoxModule.bas` |
| EfilingCommon.bas | 4069 | `vba/EfilingCommon.bas` |
| ePayPrefill.bas | 7 | `vba/ePayPrefill.bas` |
| FilingSectRadioButton.bas | 59 | `vba/FilingSectRadioButton.bas` |
| GenerateJson.bas | 7121 | `vba/GenerateJson.bas` |
| HS256.cls | 534 | `vba/HS256.cls` |
| ImportExcel.bas | 1302 | `vba/ImportExcel.bas` |
| ImportJson.bas | 6764 | `vba/ImportJson.bas` |
| md80C.bas | 213 | `vba/md80C.bas` |
| md80CCC.bas | 196 | `vba/md80CCC.bas` |
| md80E.bas | 835 | `vba/md80E.bas` |
| md80EE.bas | 1061 | `vba/md80EE.bas` |
| md80EEA.bas | 1153 | `vba/md80EEA.bas` |
| md80EEB.bas | 889 | `vba/md80EEB.bas` |
| mdATI.bas | 783 | `vba/mdATI.bas` |
| mdCalInterst234B.bas | 708 | `vba/mdCalInterst234B.bas` |
| mdGen139_8A.bas | 412 | `vba/mdGen139_8A.bas` |
| mdHashing.bas | 169 | `vba/mdHashing.bas` |
| mdImportXML.bas | 2299 | `vba/mdImportXML.bas` |
| mdInitializer.bas | 94 | `vba/mdInitializer.bas` |
| mdInt24b.bas | 957 | `vba/mdInt24b.bas` |
| mdProgressBar.bas | 108 | `vba/mdProgressBar.bas` |
| MessageBox.frm | 13 | `vba/MessageBox.frm` |
| mfMessage.bas | 49 | `vba/mfMessage.bas` |
| mIncmDtls.bas | 7348 | `vba/mIncmDtls.bas` |
| Module1.bas | 1917 | `vba/Module1.bas` |
| Module2.bas | 1913 | `vba/Module2.bas` |
| Module3.bas | 2 | `vba/Module3.bas` |
| Module4.bas | 2 | `vba/Module4.bas` |
| Module5.bas | 104 | `vba/Module5.bas` |
| Module8.bas | 36 | `vba/Module8.bas` |
| ModuleEA10_13A.bas | 327 | `vba/ModuleEA10_13A.bas` |
| Pincodechanges.bas | 62 | `vba/Pincodechanges.bas` |
| PreFillJson.bas | 4431 | `vba/PreFillJson.bas` |
| PWD.bas | 42 | `vba/PWD.bas` |
| Sch80D.bas | 1152 | `vba/Sch80D.bas` |
| Sch80G.bas | 2412 | `vba/Sch80G.bas` |
| Sch80GGA.bas | 606 | `vba/Sch80GGA.bas` |
| Sch80GGC.bas | 922 | `vba/Sch80GGC.bas` |
| Sch80U_DD.bas | 1085 | `vba/Sch80U_DD.bas` |
| SchAL.bas | 167 | `vba/SchAL.bas` |
| SchBA.bas | 931 | `vba/SchBA.bas` |
| SchDI.bas | 12 | `vba/SchDI.bas` |
| SchHP.bas | 3089 | `vba/SchHP.bas` |
| SchTaxVerify.bas | 676 | `vba/SchTaxVerify.bas` |
| SchTCS.bas | 568 | `vba/SchTCS.bas` |
| SchTDS.bas | 3221 | `vba/SchTDS.bas` |
| Sheet1.cls | 2606 | `vba/Sheet1.cls` |
| Sheet10.cls | 9 | `vba/Sheet10.cls` |
| Sheet11.cls | 78 | `vba/Sheet11.cls` |
| Sheet12.cls | 723 | `vba/Sheet12.cls` |
| Sheet13.cls | 544 | `vba/Sheet13.cls` |
| Sheet14.cls | 394 | `vba/Sheet14.cls` |
| Sheet15.cls | 15 | `vba/Sheet15.cls` |
| Sheet16.cls | 368 | `vba/Sheet16.cls` |
| Sheet17.cls | 1797 | `vba/Sheet17.cls` |
| Sheet18.cls | 94 | `vba/Sheet18.cls` |
| Sheet19.cls | 709 | `vba/Sheet19.cls` |
| Sheet2.cls | 315 | `vba/Sheet2.cls` |
| Sheet201.cls | 698 | `vba/Sheet201.cls` |
| Sheet202.cls | 222 | `vba/Sheet202.cls` |
| Sheet3.cls | 377 | `vba/Sheet3.cls` |
| Sheet4.cls | 1095 | `vba/Sheet4.cls` |
| Sheet5.cls | 9 | `vba/Sheet5.cls` |
| Sheet6.cls | 56 | `vba/Sheet6.cls` |
| Sheet7.cls | 65 | `vba/Sheet7.cls` |
| Sheet8.cls | 9 | `vba/Sheet8.cls` |
| Sheet9.cls | 699 | `vba/Sheet9.cls` |
| ThisWorkbook.cls | 85 | `vba/ThisWorkbook.cls` |
| UserForm1.frm | 15 | `vba/UserForm1.frm` |
| UserForm2.frm | 16 | `vba/UserForm2.frm` |
| UserForm3.frm | 16 | `vba/UserForm3.frm` |

## Defined names

- `_xleta.CONCATENATE` → #NAME?
- `_xleta.DAY` → #NAME?
- `_xleta.MID` → #NAME?
- `_xleta.SEC` → #NAME?
- `_xleta.TODAY` → #NAME?
- `_xleta.TRIM` → #NAME?
- `_xleta.YEAR` → #NAME?
- `Aadhaardependent_80DD` → '80U-80DD'!$H$20
- `AccountofDonor_80GGC` → '80GGC'!#REF!
- `AckNoFm10IAfiled_80DD` → '80U-80DD'!$I$20
- `AckNoFm10IAfiled_80U` → '80U-80DD'!$F$6:$F$6
- `AcknowledgeNum11A2_80DD` → '80U-80DD'!#REF!
- `AcknowledgeNum11A2_80U` → '80U-80DD'!#REF!
- `Address_80GGA` → '80GGA'!$G$8:$G$11
- `AfterDateReturn` → DataBase!$AZ$3:$AZ$5
- `age` → #REF!
- `age` → 'Taxes Paid and Verification'!$Y$46
- `Agri_dropdown` → DataBase!$BE$39:$BE$42
- `AL.ArchCollDrawPaintSulpArt` → #REF!
- `AL.ArchCollDrawPaintSulpArt` → #REF!
- `AL.CashInHand` → #REF!
- `AL.CashInHand` → #REF!
- `AL.DepositsInBank` → #REF!
- `AL.DepositsInBank` → #REF!
- `AL.ImmovableAssetBuilding` → #REF!
- `AL.ImmovableAssetBuilding` → #REF!
- `AL.ImmovableAssetLand` → #REF!
- `AL.ImmovableAssetLand` → #REF!
- `AL.InsurancePolicies` → #REF!
- `AL.InsurancePolicies` → #REF!
- `AL.JewelleryBullionEtc` → #REF!
- `AL.JewelleryBullionEtc` → #REF!
- `AL.LiabilityInRelatAssets` → #REF!
- `AL.LiabilityInRelatAssets` → #REF!
- `AL.LoansAndAdvancesGiven` → #REF!
- `AL.LoansAndAdvancesGiven` → #REF!
- `AL.SharesAndSecurities` → #REF!
- `AL.SharesAndSecurities` → #REF!
- `al.TotalImmovablMovablAssets` → #REF!
- `al.TotalImmovablMovablAssets` → #REF!
- `AL.VehiclYachtsBoatsAircrafts` → #REF!
- `AL.VehiclYachtsBoatsAircrafts` → #REF!
- `All_Pincode_List` → DataBase!$CP$2:$CP$19303
- `All_Pincode_V` → DataBase!$CP$2:$CQ$19303
- `Amount.80C` → '80C'!$E$5:$E$8
- `Amount.80CCC` → '80C'!$F$19:$F$22
- `Amount_80CCC` → 'Income Details'!$AN$118:$AW$119
- `Amount_80CCD1` → 'Income Details'!$AN$123:$AW$124
- `Amount_80CCD1b` → 'Income Details'!$AN$128:$AW$129
- `AmtA1.80D` → '80D'!$L$10:$L$12
- `AmtA2.80D` → '80D'!$L$34:$L$36
- `AmtB1.80D` → '80D'!$L$21:$L$23
- `AmtB2.80D` → '80D'!$L$45:$L$47
- `Amtdeduction_80DD` → '80U-80DD'!$E$20
- `Amtdeduction_80U` → '80U-80DD'!$E$6:$E$6
- `AY23Return` → #REF!
- `AY23Return` → DataBase!$BB$2:$BB$3
- `BA.BankCode` → #REF!
- `BA.BankCode` → #REF!
- `BA.Button1` → #REF!
- `BA.Button1` → #REF!
- `BacValue` → #REF!
- `BacValue` → DataBase!$AP$37
- `Balance_Interest` → DataBase!$GE$6:$GE$124
- `BankCode` → BankCode!$A$1:$A$180
- `bankName.24b` → 'Schedule 24(b)'!$G$5:$G$8
- `bankName.24b1` → HP!$F$32:$F$35
- `bankName.24b2` → HP!$F$68:$F$71
- `bankName.80E` → '80E_80EE_80EEA_80EEB'!$G$5:$G$6
- `bankName.80EE` → '80E_80EE_80EEA_80EEB'!$G$17:$G$18
- `bankName.80EEA` → '80E_80EE_80EEA_80EEB'!$G$29:$G$30
- `bankName.80EEB` → '80E_80EE_80EEA_80EEB'!$G$40:$G$41
- `BankNameList` → #REF!
- `BankNameList` → BankCode!#REF!
- `BankNameWithCode` → #REF!
- `BankNameWithCode` → #REF!
- `BeforeDateReturn` → DataBase!$AX$3:$AX$4
- `Block_1B` → HP!$Q$43
- `Button_80GGA` → '80GGA'!$C$14
- `Button_epaytax` → 'Taxes Paid and Verification'!$H$9
- `Button_nature` → 'Income Details'!$G$68
- `Button_nature_1` → 'Income Details'!$G$91:$AW$91
- `Button_nature1` → 'Income Details'!$F$170
- `C_Eligible` → #REF!
- `C_Eligible` → '80G'!$AE$3
- `Capacity` → #REF!
- `Capacity` → DataBase!$F$3:$F$5
- `Capacitym` → #REF!
- `Capacitym` → DataBase!$F$7:$F$8
- `CD_EligibleAmount` → #REF!
- `CD_EligibleAmount` → '80G'!$AH$3
- `CDE_EligibleAmount` → #REF!
- `CDE_EligibleAmount` → '80G'!$AI$3
- `Chequeno_80GGC` → '80GGC'!$L$8:$L$16
- `City_Town_District_80GGA` → '80GGA'!$H$8:$H$11
- `clauseiv7provisio139iAmount_1` → 'Income Details'!$AN$38
- `clauseiv7provisio139iAmount_2` → 'Income Details'!$AN$39
- `clauseiv7provisio139iAmount_3` → 'Income Details'!$AN$40
- `clauseiv7provisio139iAmount_4` → 'Income Details'!$AN$41
- `clauseiv7provisio139iFlg` → 'Income Details'!$AN$37
- `clauseiv7provisio139iFlg_1` → 'Income Details'!$AD$38
- `clauseiv7provisio139iFlg_2` → 'Income Details'!$AD$39
- `clauseiv7provisio139iFlg_3` → 'Income Details'!$AD$40
- `clauseiv7provisio139iFlg_4` → 'Income Details'!$AD$41
- `Co_OwnnerBlock1` → HP!$G$9:$J$15
- `Co_OwnnerBlock2` → HP!$G$45:$J$51
- `comb_80G_A` → '80G'!$Y$8:$Y$11
- `comb_80G_B` → '80G'!$Y$22:$Y$25
- `comb_80G_C` → '80G'!$Y$35:$Y$38
- `comb_80G_D` → '80G'!$Y$48:$Y$51
- `Comb_donation_80G_A` → '80G'!$Z$8:$Z$11
- `Comb_donation_80G_B` → '80G'!$Z$22:$Z$25
- `Comb_donation_80G_C` → '80G'!$Z$35:$Z$38
- `Comb_donation_80G_D` → '80G'!$Z$48:$Z$51
- `Comb_PAN_80G_A` → '80G'!$AA$8:$AA$11
- `Comb_PAN_80G_B` → '80G'!$AA$22:$AA$25
- `Comb_PAN_80G_C` → '80G'!$AA$35:$AA$38
- `Comb_PAN_80G_D` → '80G'!$AA$48:$AA$51
- `Combination_24B1` → HP!$O$32:$O$35
- `Combination_24B2` → HP!$O$68:$O$71
- `Combination_80EE` → '80E_80EE_80EEA_80EEB'!$X$17:$X$18
- `Combination_80EEA` → '80E_80EE_80EEA_80EEB'!$W$29:$W$30
- `Comp_dropdown` → DataBase!$BE$44:$BE$47
- `Count1` → 'Income Details'!$AZ$55
- `Count2` → 'Income Details'!$AZ$56
- `Count3` → 'Income Details'!$AZ$57
- `Count4` → 'Income Details'!$AZ$58
- `CountList` → #REF!
- `CountList` → DataBase!$Z$3:$Z$253
- `CountOS1` → 'Income Details'!$BA$55
- `CountOS2` → 'Income Details'!$BA$56
- `CountOS3` → 'Income Details'!$BA$57
- `CountOS4` → 'Income Details'!$BA$58
- `CountOS5` → 'Income Details'!$BA$63
- `Country` → [1]DropDownValues!$BU$3:$BU$251
- `Country_1B` → HP!$T$41
- `D_Eligible` → '80G'!$AF$3
- `Data_80G_A` → '80G'!$X$8:$X$11
- `Data_80G_B` → '80G'!$X$22:$X$25
- `Data_80G_C` → '80G'!$X$35:$X$38
- `Data_80G_D` → '80G'!$X$48:$X$51
- `Date_1` → DataBase!$B$21
- `Date_8A` → DataBase!$B$22
- `Date_Donation` → '80GGA'!$L$8:$L$11
- `Date_Notice` → DataBase!$D$19
- `DatefilingFm10IA_80DD` → '80U-80DD'!#REF!
- `DatefilingFm10IA_80U` → '80U-80DD'!#REF!
- `DateOfCont1` → DataBase!$B$27
- `DateOfCont2` → DataBase!$B$28
- `DateofDonation_80GGC` → '80GGC'!$E$8:$E$16
- `DateOfFiling` → DataBase!$D$4
- `DateOfFiling234A` → DataBase!$D$5
- `DateOfProcessing` → DataBase!$B$1
- `Deductions_16` → #REF!
- `Deductions_16` → 'Income Details'!$AO$72
- `DI_80C` → #REF!
- `DI_80C` → #REF!
- `DI_80CCC` → #REF!
- `DI_80CCC` → #REF!
- `DI_80CCD_1` → #REF!
- `DI_80CCD_1` → #REF!
- `DI_80CCD_1B` → #REF!
- `DI_80CCD_1B` → #REF!
- `DI_80CCD_2` → #REF!
- `DI_80CCD_2` → #REF!
- `DI_80CCG` → #REF!
- `DI_80CCG` → #REF!
- `DI_80D` → #REF!
- `DI_80D` → #REF!
- `DI_80DD` → #REF!
- `DI_80DD` → #REF!
- `DI_80DDB` → #REF!
- `DI_80DDB` → #REF!
- `DI_80E` → #REF!
- `DI_80E` → #REF!
- `DI_80EE` → #REF!
- `DI_80EE` → #REF!
- `DI_80EEA` → #REF!
- `DI_80EEA` → #REF!
- `DI_80EEB` → #REF!
- `DI_80EEB` → #REF!
- `DI_80G` → #REF!
- `DI_80G` → #REF!
- `DI_80GG` → #REF!
- `DI_80GG` → #REF!
- `DI_80GGA` → #REF!
- `DI_80GGA` → #REF!
- `DI_80GGC` → #REF!
- `DI_80GGC` → #REF!
- `DI_Deduction_80C` → #REF!
- `DI_Deduction_80C` → #REF!
- `DI_Deduction_80CCC` → #REF!
- `DI_Deduction_80CCC` → #REF!
- `DI_Deduction_80CCD_1` → #REF!
- `DI_Deduction_80CCD_1` → #REF!
- `DI_Deduction_80CCD_1B` → #REF!
- `DI_Deduction_80CCD_1B` → #REF!
- `DI_Deduction_80CCD_2` → #REF!
- `DI_Deduction_80CCD_2` → #REF!
- `DI_Deduction_80CCG` → #REF!

## Sheet dumps (`sheets/`) — hidden and visible

| Sheet | State | Rows with values | Cells | Formulas | File |
|-------|-------|------------------|-------|----------|------|
| Income Details | visible | 178 | 635 | 156 | `sheets/01_Income_Details.tsv` |
| HP | visible | 75 | 219 | 49 | `sheets/02_HP.tsv` |
| Schedule EA 10(13A) | hidden | 10 | 24 | 3 | `sheets/03_Schedule_EA_10_13A.tsv` |
| Schedule 24(b) | hidden | 8 | 35 | 6 | `sheets/04_Schedule_24_b.tsv` |
| Part A Gen_139(8A) | hidden | 24 | 45 | 8 | `sheets/05_Part_A_Gen_139_8A.tsv` |
| TDS | visible | 41 | 182 | 66 | `sheets/06_TDS.tsv` |
| TCS | visible | 8 | 21 | 4 | `sheets/07_TCS.tsv` |
| Taxes Paid and Verification | visible | 47 | 108 | 13 | `sheets/08_Taxes_Paid_and_Verification.tsv` |
| Part B ATI | hidden | 40 | 134 | 22 | `sheets/09_Part_B_ATI.tsv` |
| 80D | hidden | 46 | 106 | 17 | `sheets/10_80D.tsv` |
| 80G | hidden | 43 | 260 | 141 | `sheets/11_80G.tsv` |
| 80GGA | hidden | 11 | 50 | 17 | `sheets/12_80GGA.tsv` |
| 80GGC | hidden | 15 | 74 | 36 | `sheets/13_80GGC.tsv` |
| 80U-80DD | hidden | 10 | 42 | 7 | `sheets/14_80U-80DD.tsv` |
| 80C | hidden | 16 | 30 | 6 | `sheets/15_80C.tsv` |
| 80E_80EE_80EEA_80EEB | hidden | 26 | 117 | 13 | `sheets/16_80E_80EE_80EEA_80EEB.tsv` |
| BankCode | hidden | 318 | 636 | 0 | `sheets/17_BankCode.tsv` |
| IFSC | hidden | 40437 | 140437 | 0 | `sheets/18_IFSC.tsv.gz` |
| DataBase | hidden | 19303 | 40792 | 17 | `sheets/19_DataBase.tsv.gz` |
| SUMMARY | hidden | 15 | 30 | 10 | `sheets/20_SUMMARY.tsv` |
| Help | hidden | 67 | 95 | 0 | `sheets/21_Help.tsv` |
