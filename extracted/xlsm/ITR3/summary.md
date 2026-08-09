# ITR3 — Excel utility extraction

Source: `ITR3_AY_26-27_V1.2/ITR3_AY_26-27_V1.2.xlsm` (12,253,218 bytes)

- Worksheets: **66** (40 visible, 26 hidden)
- Sheet contents (hidden + visible): **66 dumps** in `sheets/` (495,612 stored cell values, formulas extracted alongside)
- VBA modules: **158**, 240,617 lines of code
- Data-validation (dropdown) rules: **1257** (resolved to actual list values in `dropdowns.json`)
- Defined names: **5821**

## Worksheets

| # | Sheet | State | Dimensions | Formulas | Validations |
|---|-------|-------|-----------|----------|-------------|
| 1 | ISIN List | hidden | A1:E124306 | 0 | 0 |
| 2 | Home | visible | A1:IU57 | 50 | 1 |
| 3 | PART A - General | visible | A1:BG201 | 16 | 151 |
| 4 | Part A Gen_139(8A) | hidden | A1:BB36 | 8 | 16 |
| 5 | Nature Of Business | visible | A1:Q65557 | 16 | 3 |
| 6 | Part A - BS | visible | A1:P97 | 24 | 4 |
| 7 | Manufacturing Account | visible | B1:V43 | 6 | 3 |
| 8 | Trading Account | visible | A1:R1005 | 12 | 7 |
| 9 | Profit and Loss | visible | A1:IU1019 | 83 | 170 |
| 10 | Part A - OI | visible | A1:O1003 | 11 | 14 |
| 11 | Quantitative Details | visible | A1:IU65543 | 0 | 14 |
| 12 | Sheet1 | hidden | A1:I26 | 16 | 1 |
| 13 | ITold | veryHidden | A1:V1000 | 64 | 6 |
| 14 | Schedule S | visible | A1:U148 | 27 | 35 |
| 15 | House Property | visible | A1:T65605 | 46 | 41 |
| 16 | BP | visible | A1:XFC1010 | 81 | 26 |
| 17 | DPM - DOA | visible | A1:XFC1002 | 125 | 47 |
| 18 | DEP_DCG | visible | A1:L46 | 34 | 2 |
| 19 | ESR | visible | A1:I1000 | 12 | 3 |
| 20 | CG | visible | A1:BN1121 | 757 | 110 |
| 21 | Schedule 112A | visible | A1:AG17 | 43 | 19 |
| 22 | HelpCSV | hidden | A1:IU129 | 0 | 0 |
| 23 | Schedule 115AD(1)(iii) proviso | visible | A1:AC17 | 39 | 19 |
| 24 | OS | visible | A1:XFC1015 | 331 | 34 |
| 25 | CYLA - BFLA | visible | A1:XFC65545 | 223 | 3 |
| 26 | CFL | visible | A1:XFC65542 | 62 | 5 |
| 27 | Unabsorbed Depreciation | visible | A1:XFC33 | 40 | 10 |
| 28 | 80GGC | hidden | A1:N16 | 18 | 7 |
| 29 | VDA | visible | A1:M21 | 7 | 10 |
| 30 | 80U-80DD | hidden | A1:AB129 | 6 | 15 |
| 31 | ICDS | visible | A1:L19 | 15 | 5 |
| 32 | 10AA | hidden | A1:IV1000 | 8 | 4 |
| 33 | 80G | hidden | A1:IX1003 | 181 | 17 |
| 34 | 80C | hidden | B1:G11 | 4 | 4 |
| 35 | 80D | hidden | B1:M54 | 21 | 20 |
| 36 | RA | hidden | A1:XFC15 | 11 | 9 |
| 37 | 80GGA | hidden | A1:XFC15 | 21 | 9 |
| 38 | 80 | hidden | A1:J1005 | 88 | 27 |
| 39 | 80E_80EE_80EEA_80EEB | hidden | A2:W44 | 19 | 8 |
| 40 | VI-A | visible | A1:BC74 | 123 | 23 |
| 41 | SPI - SI - IF | visible | A1:XFC181 | 927 | 29 |
| 42 | AMT | hidden | A1:AA15 | 37 | 8 |
| 43 | AMTC | visible | A1:IV27 | 25 | 14 |
| 44 | EI | visible | A1:IV1000 | 15 | 22 |
| 45 | FSI1 | hidden | A1:AL286 | 45 | 10 |
| 46 | PTI | visible | A1:IU54 | 62 | 7 |
| 47 | TPSA | visible | B1:Q36 | 8 | 11 |
| 48 | FSI | visible | A1:XFC32 | 39 | 9 |
| 49 | TR_FA | visible | A1:XFC168 | 49 | 72 |
| 50 | Sch 5A | visible | A1:IV1000 | 4 | 9 |
| 51 | AL | visible | A1:R65536 | 6 | 25 |
| 52 | GST | visible | A1:IU65522 | 3 | 2 |
| 53 | Tax Calculated | hidden | A1:CM377 | 2019 | 8 |
| 54 | Part B - TI TTI | visible | A1:XFC138 | 152 | 26 |
| 55 | IT | visible | A1:XFC32 | 161 | 5 |
| 56 | ESOP | visible | A1:W81 | 20 | 17 |
| 57 | Part B ATI | hidden | A1:Z56 | 24 | 11 |
| 58 | TDS | visible | A1:IX90 | 44 | 56 |
| 59 | Verification | visible | B1:O24 | 1 | 5 |
| 60 | OLDAL | veryHidden | A1:K19 | 1 | 2 |
| 61 | Temporary Values | hidden | A1:F62 | 5 | 0 |
| 62 | DropDownValues | hidden | A1:HD19302 | 101 | 0 |
| 63 | CG Pop up_prefill | hidden | A1:D6 | 0 | 0 |
| 64 | SUMMARY | hidden | A1:IU77 | 45 | 0 |
| 65 | BA | hidden | A1:IU12 | 2 | 7 |
| 66 | Instructions | hidden | A1:IU128 | 6 | 0 |

## VBA modules

| Module | Lines | File |
|--------|-------|------|
| AY23_24_Changes.bas | 3617 | `vba/AY23_24_Changes.bas` |
| AY25_26_Changes.bas | 1327 | `vba/AY25_26_Changes.bas` |
| AY26_27_Changes.bas | 918 | `vba/AY26_27_Changes.bas` |
| BFLA_Calculations.bas | 405 | `vba/BFLA_Calculations.bas` |
| Calculations.bas | 238 | `vba/Calculations.bas` |
| CG_Calc.bas | 1819 | `vba/CG_Calc.bas` |
| CGDeductions.bas | 3300 | `vba/CGDeductions.bas` |
| checkBoxModule.bas | 35 | `vba/checkBoxModule.bas` |
| CYLACalculations.bas | 11132 | `vba/CYLACalculations.bas` |
| DEP_DCG.bas | 302 | `vba/DEP_DCG.bas` |
| EfilingCommon.bas | 5060 | `vba/EfilingCommon.bas` |
| ePayPrefill.bas | 8 | `vba/ePayPrefill.bas` |
| FilingSectRadioButton.bas | 100 | `vba/FilingSectRadioButton.bas` |
| Generate_XML.bas | 18565 | `vba/Generate_XML.bas` |
| GenerateJson.bas | 34343 | `vba/GenerateJson.bas` |
| GST.bas | 199 | `vba/GST.bas` |
| HS256.cls | 535 | `vba/HS256.cls` |
| ImportJson.bas | 28968 | `vba/ImportJson.bas` |
| ImportPrefill.bas | 16709 | `vba/ImportPrefill.bas` |
| ImportSchedule112A.bas | 351 | `vba/ImportSchedule112A.bas` |
| ImportSchedule115AD.bas | 358 | `vba/ImportSchedule115AD.bas` |
| Initializer.bas | 77 | `vba/Initializer.bas` |
| Manufacturing_Account.bas | 6 | `vba/Manufacturing_Account.bas` |
| md112A.bas | 998 | `vba/md112A.bas` |
| md115AD.bas | 961 | `vba/md115AD.bas` |
| md80_.bas | 2118 | `vba/md80_.bas` |
| md80C.bas | 174 | `vba/md80C.bas` |
| md80E.bas | 562 | `vba/md80E.bas` |
| md80EE.bas | 703 | `vba/md80EE.bas` |
| md80EE_Addrows.bas | 29 | `vba/md80EE_Addrows.bas` |
| md80EEA.bas | 727 | `vba/md80EEA.bas` |
| md80EEB.bas | 634 | `vba/md80EEB.bas` |
| mdAMT.bas | 185 | `vba/mdAMT.bas` |
| mdAMTC.bas | 1646 | `vba/mdAMTC.bas` |
| mdATI.bas | 877 | `vba/mdATI.bas` |
| mdBA.bas | 1199 | `vba/mdBA.bas` |
| mdCFL.bas | 997 | `vba/mdCFL.bas` |
| mdEI.bas | 1210 | `vba/mdEI.bas` |
| mdESR.bas | 274 | `vba/mdESR.bas` |
| mdGen139_8A.bas | 385 | `vba/mdGen139_8A.bas` |
| mdHashing.bas | 98 | `vba/mdHashing.bas` |
| mdICDS.bas | 127 | `vba/mdICDS.bas` |
| mdImportXL.bas | 3683 | `vba/mdImportXL.bas` |
| mdImportXML.bas | 2590 | `vba/mdImportXML.bas` |
| mdInterestCalc.bas | 924 | `vba/mdInterestCalc.bas` |
| mdOI.bas | 2246 | `vba/mdOI.bas` |
| mdProgressbar.bas | 90 | `vba/mdProgressbar.bas` |
| mdPTI.bas | 1659 | `vba/mdPTI.bas` |
| mdVDA.bas | 100 | `vba/mdVDA.bas` |
| MessageBox.frm | 14 | `vba/MessageBox.frm` |
| mfMessage.bas | 37 | `vba/mfMessage.bas` |
| Module1.bas | 8 | `vba/Module1.bas` |
| Module2.bas | 45 | `vba/Module2.bas` |
| Module3.bas | 119 | `vba/Module3.bas` |
| Module4.bas | 13 | `vba/Module4.bas` |
| msSI.bas | 532 | `vba/msSI.bas` |
| Nature_of_business.bas | 135 | `vba/Nature_of_business.bas` |
| ParseJson.bas | 347 | `vba/ParseJson.bas` |
| Part_A_General.bas | 8440 | `vba/Part_A_General.bas` |
| PARTA_BS.bas | 210 | `vba/PARTA_BS.bas` |
| Profit_Loss.bas | 2478 | `vba/Profit_Loss.bas` |
| pwd.bas | 86 | `vba/pwd.bas` |
| Quantitative_Details.bas | 922 | `vba/Quantitative_Details.bas` |
| Sch10A.bas | 207 | `vba/Sch10A.bas` |
| Sch5A.bas | 451 | `vba/Sch5A.bas` |
| Sch80D.bas | 1122 | `vba/Sch80D.bas` |
| Sch80G.bas | 2876 | `vba/Sch80G.bas` |
| Sch80GGC.bas | 836 | `vba/Sch80GGC.bas` |
| Sch80U_80DD.bas | 1127 | `vba/Sch80U_80DD.bas` |
| SchAL.bas | 1648 | `vba/SchAL.bas` |
| SchBP.bas | 371 | `vba/SchBP.bas` |
| SchCG.bas | 7842 | `vba/SchCG.bas` |
| SchDPM_DOA.bas | 2128 | `vba/SchDPM_DOA.bas` |
| SchFSI.bas | 906 | `vba/SchFSI.bas` |
| SchHP.bas | 3160 | `vba/SchHP.bas` |
| SchIT.bas | 564 | `vba/SchIT.bas` |
| SchOI.bas | 81 | `vba/SchOI.bas` |
| SchOS.bas | 2736 | `vba/SchOS.bas` |
| SchRA.bas | 1109 | `vba/SchRA.bas` |
| SchTDS.bas | 3715 | `vba/SchTDS.bas` |
| SchTPSA.bas | 305 | `vba/SchTPSA.bas` |
| SchTR_FA.bas | 6238 | `vba/SchTR_FA.bas` |
| SchUD.bas | 574 | `vba/SchUD.bas` |
| Sheet1.cls | 8730 | `vba/Sheet1.cls` |
| Sheet10.cls | 246 | `vba/Sheet10.cls` |
| Sheet11.cls | 46 | `vba/Sheet11.cls` |
| Sheet12.cls | 62 | `vba/Sheet12.cls` |
| Sheet13.cls | 1883 | `vba/Sheet13.cls` |
| Sheet14.cls | 1115 | `vba/Sheet14.cls` |
| Sheet15.cls | 73 | `vba/Sheet15.cls` |
| Sheet16.cls | 116 | `vba/Sheet16.cls` |
| Sheet17.cls | 96 | `vba/Sheet17.cls` |
| Sheet18.cls | 62 | `vba/Sheet18.cls` |
| Sheet19.cls | 419 | `vba/Sheet19.cls` |
| Sheet2.cls | 578 | `vba/Sheet2.cls` |
| Sheet20.cls | 1218 | `vba/Sheet20.cls` |
| Sheet201.cls | 940 | `vba/Sheet201.cls` |
| Sheet202.cls | 315 | `vba/Sheet202.cls` |
| Sheet21.cls | 467 | `vba/Sheet21.cls` |
| Sheet22.cls | 54 | `vba/Sheet22.cls` |
| Sheet23.cls | 120 | `vba/Sheet23.cls` |
| Sheet24.cls | 401 | `vba/Sheet24.cls` |
| Sheet25.cls | 334 | `vba/Sheet25.cls` |
| Sheet26.cls | 255 | `vba/Sheet26.cls` |
| Sheet27.cls | 402 | `vba/Sheet27.cls` |
| Sheet28.cls | 1338 | `vba/Sheet28.cls` |
| Sheet29.cls | 49 | `vba/Sheet29.cls` |
| Sheet3.cls | 142 | `vba/Sheet3.cls` |
| Sheet30.cls | 40 | `vba/Sheet30.cls` |
| Sheet31.cls | 11 | `vba/Sheet31.cls` |
| Sheet32.cls | 11 | `vba/Sheet32.cls` |
| Sheet33.cls | 18 | `vba/Sheet33.cls` |
| Sheet34.cls | 18 | `vba/Sheet34.cls` |
| Sheet35.cls | 18 | `vba/Sheet35.cls` |
| Sheet36.cls | 45 | `vba/Sheet36.cls` |
| Sheet37.cls | 138 | `vba/Sheet37.cls` |
| Sheet38.cls | 179 | `vba/Sheet38.cls` |
| Sheet39.cls | 2451 | `vba/Sheet39.cls` |
| Sheet4.cls | 257 | `vba/Sheet4.cls` |
| Sheet40.cls | 540 | `vba/Sheet40.cls` |
| Sheet41.cls | 114 | `vba/Sheet41.cls` |
| Sheet42.cls | 39 | `vba/Sheet42.cls` |
| Sheet43.cls | 162 | `vba/Sheet43.cls` |
| Sheet44.cls | 15 | `vba/Sheet44.cls` |
| Sheet45.cls | 18 | `vba/Sheet45.cls` |
| Sheet46.cls | 701 | `vba/Sheet46.cls` |
| Sheet47.cls | 112 | `vba/Sheet47.cls` |
| Sheet48.cls | 51 | `vba/Sheet48.cls` |
| Sheet49.cls | 110 | `vba/Sheet49.cls` |
| Sheet5.cls | 192 | `vba/Sheet5.cls` |
| Sheet50.cls | 194 | `vba/Sheet50.cls` |
| Sheet51.cls | 70 | `vba/Sheet51.cls` |
| Sheet52.cls | 227 | `vba/Sheet52.cls` |
| Sheet53.cls | 260 | `vba/Sheet53.cls` |
| Sheet54.cls | 276 | `vba/Sheet54.cls` |
| Sheet55.cls | 677 | `vba/Sheet55.cls` |
| Sheet56.cls | 114 | `vba/Sheet56.cls` |
| Sheet57.cls | 103 | `vba/Sheet57.cls` |
| Sheet58.cls | 339 | `vba/Sheet58.cls` |
| Sheet59.cls | 273 | `vba/Sheet59.cls` |
| Sheet6.cls | 56 | `vba/Sheet6.cls` |
| Sheet60.cls | 652 | `vba/Sheet60.cls` |
| Sheet61.cls | 19 | `vba/Sheet61.cls` |
| Sheet62.cls | 9 | `vba/Sheet62.cls` |
| Sheet7.cls | 413 | `vba/Sheet7.cls` |
| Sheet8.cls | 761 | `vba/Sheet8.cls` |
| Sheet9.cls | 393 | `vba/Sheet9.cls` |
| Sheet98.cls | 103 | `vba/Sheet98.cls` |
| SheetALL.cls | 14 | `vba/SheetALL.cls` |
| SPI_SI.bas | 880 | `vba/SPI_SI.bas` |
| Tax_Calc.bas | 1339 | `vba/Tax_Calc.bas` |
| ThisWorkbook.cls | 24 | `vba/ThisWorkbook.cls` |
| TI_TTI_Salary.bas | 3109 | `vba/TI_TTI_Salary.bas` |
| Trading_Account.bas | 318 | `vba/Trading_Account.bas` |
| UserForm1.frm | 18 | `vba/UserForm1.frm` |
| UserForm2.frm | 16 | `vba/UserForm2.frm` |
| UserForm3.frm | 103 | `vba/UserForm3.frm` |
| Verification.bas | 206 | `vba/Verification.bas` |

## Defined names

- `_xlnm._FilterDatabase` → 'ISIN List'!$A$1:$B$1
- `_Ins1` → HelpCSV!$K$14
- `_Ins1` → Instructions!$K$14
- `_Ins10` → HelpCSV!$K$60
- `_Ins10` → Instructions!$K$59
- `_Ins11` → HelpCSV!$K$70
- `_Ins11` → Instructions!$K$69
- `_Ins2` → HelpCSV!$K$23
- `_Ins2` → Instructions!$K$22
- `_Ins3` → HelpCSV!$K$27
- `_Ins3` → Instructions!$K$26
- `_Ins4` → HelpCSV!$K$31
- `_Ins4` → Instructions!$K$30
- `_Ins5` → HelpCSV!$K$41
- `_Ins5` → Instructions!$K$40
- `_Ins6` → HelpCSV!$K$44
- `_Ins6` → Instructions!$K$43
- `_Ins7` → HelpCSV!$K$47
- `_Ins7` → Instructions!$K$46
- `_Ins8` → HelpCSV!$K$50
- `_Ins8` → Instructions!$K$49
- `_Ins9` → HelpCSV!$K$57
- `_Ins9` → Instructions!$K$56
- `_xleta.MAX` → #NAME?
- `_xleta.MIN` → #NAME?
- `_xleta.N` → #NAME?
- `_xleta.OR` → #NAME?
- `_xleta.SUMIFS` → #NAME?
- `_xleta.T` → #NAME?
- `AA10.AssmtYrUnit` → '10AA'!$H$11:$H$13
- `AA10.DedFromUndertaking` → '10AA'!$I$11:$I$13
- `AA10.TotalDedUs10Sub` → '10AA'!$I$14
- `AA10_undertaking` → '10AA'!$F$11:$F$13
- `Aadhaardependent_80DD` → '80U-80DD'!$H$16
- `Aadhar_TDS3` → TDS!$G$38:$G$41
- `Aadhar1` → 'PART A - General'!$BB$108:$BB$111
- `Aadhar2` → 'PART A - General'!#REF!
- `AccLTCG.Up16Of12To15Of3` → CG!$Q$543
- `AccLTCG.Up16Of3To31Of3` → CG!$S$543
- `AccLTCG.Up16Of9To15Of12` → CG!$O$543
- `AccLTCG.Upto15Of6` → CG!$L$543
- `AccLTCG.Upto15Of9` → CG!$M$543
- `AccLTCG12.5.Up16Of12To15Of3` → CG!$Q$544
- `AccLTCG12.5.Up16Of3To31Of3` → CG!$S$544
- `AccLTCG12.5.Up16Of9To15Of12` → CG!$O$544
- `AccLTCG12.5.Upto15Of6` → CG!$L$544
- `AccLTCG12.5.Upto15Of9` → CG!$M$544
- `AccLTCGDTAA.Up16Of12To15Of3` → CG!$Q$546
- `AccLTCGDTAA.Up16Of3To31Of3` → CG!$S$546
- `AccLTCGDTAA.Up16Of9To15Of12` → CG!$O$546
- `AccLTCGDTAA.Upto15Of6` → CG!$L$546
- `AccLTCGDTAA.Upto15Of9` → CG!$M$546
- `AccLTCGNP.Up16Of12To15Of3` → CG!$Q$545
- `AccLTCGNP.Up16Of3To31Of3` → CG!$S$545
- `AccLTCGNP.Up16Of9To15Of12` → CG!$O$545
- `AccLTCGNP.Upto15Of6` → CG!$L$545
- `AccLTCGNP.Upto15Of9` → CG!$M$545
- `AccSTCG.Up16Of12To15Of3` → CG!$Q$538
- `AccSTCG.Up16Of3To31Of3` → CG!$S$538
- `AccSTCG.Up16Of9To15Of12` → CG!$O$538
- `AccSTCG.Upto15Of6` → CG!$L$538
- `AccSTCG.Upto15Of9` → CG!$M$538
- `AccSTCG20.Up16Of12To15Of3` → CG!$Q$539
- `AccSTCG20.Up16Of3To31Of3` → CG!$S$539
- `AccSTCG20.Up16Of9To15Of12` → CG!$O$539
- `AccSTCG20.Upto15Of6` → CG!$L$539
- `AccSTCG20.Upto15Of9` → CG!$M$539
- `ACCSTCG30.Up16Of12To15Of3` → CG!$Q$540
- `ACCSTCG30.Up16Of3To31Of3` → CG!$S$540
- `ACCSTCG30.Up16Of9To15Of12` → CG!$O$540
- `AccSTCG30.Upto15Of6` → CG!$L$540
- `AccSTCG30.Upto15Of9` → CG!$M$540
- `AccSTCGDTAA.Up16Of12To15Of3` → CG!$Q$542
- `AccSTCGDTAA.Up16Of3To31Of3` → CG!$S$542
- `AccSTCGDTAA.Up16Of9To15Of12` → CG!$O$542
- `AccSTCGDTAA.Upto15Of6` → CG!$L$542
- `AccSTCGDTAA.Upto15Of9` → CG!$M$542
- `AccSTCGOTH.Up16Of12To15Of3` → CG!$Q$541
- `AccSTCGOTH.Up16Of3To31Of3` → CG!$S$541
- `AccSTCGOTH.Up16Of9To15Of12` → CG!$O$541
- `AccSTCGOTH.Upto15Of6` → CG!$L$541
- `AccSTCGOTH.Upto15Of9` → CG!$M$541
- `AccVDA.Up16Of12To15Of3` → CG!$Q$547
- `AccVDA.Up16Of3To31Of3` → CG!$S$547
- `AccVDA.Up16Of9To15Of12` → CG!$O$547
- `AccVDA.Upto15Of6` → CG!$L$547
- `AccVDA.Upto15Of9` → CG!$M$547
- `AckNoFm10IAfiled_80DD` → '80U-80DD'!$J$16
- `AckNoFm10IAfiled_80U` → '80U-80DD'!$G$6
- `acknowledgmentnumber` → 'PART A - General'!$AK$57
- `Address_80GGA` → RA!$F$7:$F$10
- `Address_80GGANew` → '80GGA'!$G$7:$G$10
- `adjstGTI` → '80G'!$AB$23
- `adjtotloss.BusLossOthThanSpecLossCF9` → CFL!$L$23
- `adjtotloss.HPLossCF9` → CFL!$G$23
- `adjtotloss.LossFrmSpecBusCF9` → CFL!$M$23
- `adjtotloss.LossFrmSpecifiedBusCF9` → CFL!$N$23
- `adjtotloss.LTCGLossCF9` → CFL!$R$23
- `adjtotloss.OthSrcLossRaceHorseCF9` → CFL!$U$23
- `adjtotloss.STCGLossCF9` → CFL!$O$23
- `AfterDateReturn` → DropDownValues!$DM$2:$DM$4
- `age` → 'Temporary Values'!$B$54
- `Age_TX` → Sheet1!$G$8
- `Agri_dropdown` → DropDownValues!$GP$27:$GP$30
- `AL.ArchCollDrawPaintSulpArt` → OLDAL!$J$14
- `AL.CashInHand` → OLDAL!$J$12
- `AL.DepositsInBank` → OLDAL!$J$8
- `AL.ImmovableAssetBuilding` → OLDAL!$J$5
- `AL.ImmovableAssetLand` → OLDAL!$J$4
- `AL.InsurancePolicies` → OLDAL!$J$10
- `AL.JewelleryBullionEtc` → OLDAL!$J$13
- `AL.LiabilityInRelatAssets` → OLDAL!$J$17
- `AL.LoansAndAdvancesGiven` → OLDAL!$J$11
- `AL.SharesAndSecurities` → OLDAL!$J$9
- `al.TotalImmovablMovablAssets` → OLDAL!$J$16
- `AL.VehiclYachtsBoatsAircrafts` → OLDAL!$J$15
- `All_Pincode_List` → DropDownValues!$EL$2:$EL$19302
- `All_Pincode_V` → DropDownValues!$EL$2:$EM$19302
- `AllBS.names` → 'Part A - BS'!$J$8,'Part A - BS'!$L$5,'Part A - BS'!$J$7,'Part A - BS'!$J$8,'Part A - BS'!$J$9,'Part A - BS'!$J$10,'Part A - BS'!$J$15,'Part A - BS'!$J$17,'Part A - BS'!$J$18,'Part A - BS'!$J$22,'Part A - BS'!$J$23,'Part A - BS'!$L$26,'Part A - BS'!$J$33,'Part A - BS'!$J$34,'Part A - BS'!$J$36,'Part A - BS'!$J$40,'Part A - BS'!$J$41,'Part A - BS'!$J$44,'Part A - BS'!$J$45,'Part A - BS'!$J$46,'Part A - BS'!$J$52,'Part A - BS'!$J$53,'Part A - BS'!$J$54,'Part A - BS'!$J$55,'Part A - BS'!$L$57,'Part A - BS'!$J$59,'Part A - BS'!$J$60,'Part A - BS'!$L$62,'Part A - BS'!$J$65,'Part A - BS'!$J$66,'Part A - BS'!$J$67,'Part A - BS'!$J$72,'Part A - BS'!$J$73,'Part A - BS'!$J$74,'Part A - BS'!$J$75,'Part A - BS'!$J$78,'Part A - BS'!$J$80,'Part A - BS'!$J$81,'Part A - BS'!$J$85,'Part A - BS'!$J$86,'Part A - BS'!$J$87,'Part A - BS'!$L$91
- `ALLMA_NAMES` → 'Manufacturing Account'!$R$6:$T$7,'Manufacturing Account'!$R$9:$T$10,'Manufacturing Account'!$R$12:$T$14,'Manufacturing Account'!$R$16:$T$21,'Manufacturing Account'!$R$25:$T$26
- `Allow3ai` → OS!$U$94
- `AllowanceBACNo` → DropDownValues!$DY$3:$DY$16
- `AllowanceBACNoAllPGov` → DropDownValues!$EF$3:$EF$14
- `AllowanceBACNoCGSG` → DropDownValues!$EC$3:$EC$17
- `AllowanceBACNoCGSGPCGPSG` → DropDownValues!$EA$3:$EA$15
- `AllowanceBACNoPCGPSG` → DropDownValues!$EE$3:$EE$14
- `AllowanceBACYes` → DropDownValues!$DY$24:$DY$35
- `AllowanceBACYesAllPGov` → DropDownValues!$EF$24:$EF$33
- `AllowanceBACYesCGSG` → DropDownValues!$EC$26:$EC$35
- `AllowanceBACYesCGSGPCGPSG` → DropDownValues!$EA$24:$EA$33
- `AllowanceBACYesPCGPSG` → DropDownValues!$EE$24:$EE$33
- `Allpl.name` → 'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!$J$6:$J$15,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!$L$29:$L$34,'Profit and Loss'!$J$36:$J$45,'Profit and Loss'!$J$47:$J$48,'Profit and Loss'!$J$50:$J$53,'Profit and Loss'!$L$55:$L$60,'Profit and Loss'!$J$62:$J$63,'Profit and Loss'!$J$66:$J$67,'Profit and Loss'!$J$70:$J$71,'Profit and Loss'!$L$73:$L$83,'Profit and Loss'!$J$85:$J$93,'Profit and Loss'!$L$95,'Profit and Loss'!$J$115:$J$123
- `ALLPL_NAMES1` → 'Profit and Loss'!$J$6,'Profit and Loss'!$J$6:$J$15,'Profit and Loss'!$H$20:$H$24,'Profit and Loss'!$J$20:$J$24,'Profit and Loss'!$L$29:$L$34,'Profit and Loss'!$J$36:$J$45,'Profit and Loss'!$J$47:$J$48,'Profit and Loss'!$J$50:$J$53,'Profit and Loss'!$L$55:$L$60,'Profit and Loss'!$J$62:$J$63,'Profit and Loss'!$J$66:$J$67,'Profit and Loss'!$J$70:$J$71,'Profit and Loss'!$L$73:$L$83,'Profit and Loss'!$J$85:$J$93,'Profit and Loss'!$L$95,'Profit and Loss'!$H$98:$H$101,'Profit and Loss'!$J$98:$J$101,'Profit and Loss'!$H$106:$H$112,'Profit and Loss'!$J$106:$J$112
- `ALLPL_NAMES2` → 'Profit and Loss'!$L$125,'Profit and Loss'!$L$125:$L$126,'Profit and Loss'!$J$129:$J$130,'Profit and Loss'!$L$132,'Profit and Loss'!$L$134:$L$135,'Profit and Loss'!$L$137,'Profit and Loss'!$L$139,'Profit and Loss'!$E$143:$L$145,'Profit and Loss'!$K$147:$L$149,'Profit and Loss'!$K$151:$L$152,'Profit and Loss'!$E$156:$L$158,'Profit and Loss'!$K$159:$L$163,'Profit and Loss'!$E$167:$L$169,'Profit and Loss'!$E$172:$L$175,'Profit and Loss'!$L$189:$L$192,'Profit and Loss'!$L$196:$L$199,'Profit and Loss'!$L$202:$L$204
- `ALLPL_NAMES3` → 'Profit and Loss'!$F$117:$Q$118,'Profit and Loss'!$J$123
- `AllTA_Names` → 'Trading Account'!$K$6,'Trading Account'!$K$6,'Trading Account'!$K$7,'Trading Account'!$G$10:$H$11,'Trading Account'!$K$16,'Trading Account'!$K$18:$L$25,'Trading Account'!$K$28,'Trading Account'!$K$30:$L$31,'Trading Account'!$K$33:$L$34,'Trading Account'!$G$37:$H$38,'Trading Account'!$K$43:$L$53
- `Amount.80C` → '80C'!$E$5:$E$8
- `Amount_80CCC` → 'VI-A'!$G$9:$G$11
- `Amount_80CCD1` → 'VI-A'!$G$15:$G$16
- `Amount_80CCD1B` → 'VI-A'!$G$20:$G$21
- `AMT.AdjustedUnderSec115JC` → AMT!$J$10
- `AMT.AdjustedUnderSec115JC3a` → AMT!$J$11
- `AMT.AdjustedUnderSec115JC3a2_1cr` → AMT!$N$13
- `AMT.AdjustedUnderSec115JC3a2_2cr` → AMT!$N$11
- `AMT.AdjustedUnderSec115JC3a2_50L` → AMT!$N$7
- `AMT.AdjustedUnderSec115JC3a2_5cr` → AMT!$N$9
- `AMT.AdjustedUnderSec115JC3a2New_1cr` → AMT!$N$14
- `AMT.AdjustedUnderSec115JC3a2New_2cr` → AMT!$N$12
- `AMT.AdjustedUnderSec115JC3a2New_50L` → AMT!$N$8
- `AMT.AdjustedUnderSec115JC3a2New_5cr` → AMT!$N$10
- `AMT.AdjustedUnderSec115JC3aNew` → AMT!$M$9
- `AMT.AdjustedUnderSec115JC3b` → AMT!$J$12
- `AMT.AdjustedUnderSec115JC3b2_1cr` → AMT!$O$13
- `AMT.AdjustedUnderSec115JC3b2_2cr` → AMT!$O$11
- `AMT.AdjustedUnderSec115JC3b2_50L` → AMT!$O$7
- `AMT.AdjustedUnderSec115JC3b2_5cr` → AMT!$O$9
- `AMT.AdjustedUnderSec115JC3b2New_1cr` → AMT!$O$14
- `AMT.AdjustedUnderSec115JC3b2New_2cr` → AMT!$O$12
- `AMT.AdjustedUnderSec115JC3b2New_50L` → AMT!$O$8
- `AMT.AdjustedUnderSec115JC3b2New_5cr` → AMT!$O$10
- `AMT.AdjustedUnderSec115JC3bNew` → AMT!$M$10
- `AMT.BTax_1cr` → AMT!$P$13
- `AMT.BTax_2cr` → AMT!$P$11
- `AMT.BTax_50L` → AMT!$P$7
- `AMT.BTax_5cr` → AMT!$P$9
- `AMT.DeductClaimSec10AA` → AMT!$H$7
- `AMT.DeductClaimSec35AD` → AMT!$H$8
- `AMT.DeductClaimSec6A` → AMT!$H$6
- `AMT.TaxPayableUnderSec115JC` → AMT!$J$13
- `AMT.Total` → AMT!$H$9
- `AMT.TotalIncItem11` → AMT!$J$4
- `Amt_1` → BP!$I$42:$I$44
- `Amt_Condn` → AMT!$M$3
- `AmtA1.80D` → '80D'!$K$10:$L$12
- `AmtA2.80D` → '80D'!$K$34:$L$36
- `AmtB1.80D` → '80D'!$K$21:$L$23
- `AmtB2.80D` → '80D'!$K$45:$L$47
- `AMTC.AmtCreditBalBroughtFwd` → AMTC!$I$12
- `AMTC.AmtCreditBalBroughtFwd_Total` → AMTC!$I$23
- `AMTC.AmtCreditBalBroughtFwd0` → AMTC!$I$22
- `AMTC.AmtCreditBalBroughtFwd1` → AMTC!$I$9
- `AMTC.AmtCreditBalBroughtFwd10` → AMTC!$I$19
- `AMTC.AmtCreditBalBroughtFwd11` → AMTC!$I$20
- `AMTC.AmtCreditBalBroughtFwd12` → AMTC!$I$21
- `AMTC.AmtCreditBalBroughtFwd2` → AMTC!$I$10
- `AMTC.AmtCreditBalBroughtFwd3` → AMTC!$I$11
- `AMTC.AmtCreditBalBroughtFwd4` → AMTC!$I$14
- `AMTC.AmtCreditBalBroughtFwd5` → AMTC!$I$13
- `AMTC.AmtCreditBalBroughtFwd6` → AMTC!$I$15
- `AMTC.AmtCreditBalBroughtFwd7` → AMTC!$I$16
- `AMTC.AmtCreditBalBroughtFwd8` → AMTC!$I$17
- `AMTC.AmtCreditBalBroughtFwd9` → AMTC!$I$18
- `AMTC.AmtCreditFwd` → AMTC!$G$12
- `AMTC.AmtCreditFwd_Total` → AMTC!$G$23
- `AMTC.AmtCreditFwd0` → AMTC!$G$22
- `AMTC.AmtCreditFwd1` → AMTC!$G$9
- `AMTC.AmtCreditFwd10` → AMTC!$G$19
- `AMTC.AmtCreditFwd11` → AMTC!$G$20
- `AMTC.AmtCreditFwd12` → AMTC!$G$21
- `AMTC.AmtCreditFwd2` → AMTC!$G$10

## Companion CSV templates

- `ITR3_AY_26-27_V1.2/112A & 115AD/CSV_112A and 115AD.csv` — header: `Share/Unit acquired(1a),ISIN Code(2),Name of the Share/Unit(3),No. of Shares/Units(4),Sale-price per Share/Unit(5),Full Value of Consideration(Total Sale Value)`
- `ITR3_AY_26-27_V1.2/TCS & TDS & IT/CSV_IT.csv` — header: `"BSR Code`
- `ITR3_AY_26-27_V1.2/TCS & TDS & IT/CSV_TCS.csv` — header: `TCS credit relating to Self/Other Person [Spouse as per section 5A/Other person as per Rule 37i(1)] (Col 2(i)),"TAN of the`
- `ITR3_AY_26-27_V1.2/TCS & TDS & IT/CSV_TDS1.csv` — header: `"Tax Deduction Account`
- `ITR3_AY_26-27_V1.2/TCS & TDS & IT/CSV_TDS2.csv` — header: `"TDS Credit related to self/ other person as per rule 37BA(2) `
- `ITR3_AY_26-27_V1.2/TCS & TDS & IT/CSV_TDS3.csv` — header: `TDS credit in the name of  (Col 2),"PAN Of Other Person (If TDS Credit related to other person)`

## Sheet dumps (`sheets/`) — hidden and visible

| Sheet | State | Rows with values | Cells | Formulas | File |
|-------|-------|------------------|-------|----------|------|
| ISIN List | hidden | 124306 | 425748 | 0 | `sheets/01_ISIN_List.tsv.gz` |
| Home | visible | 52 | 307 | 7 | `sheets/02_Home.tsv` |
| PART A - General | visible | 180 | 369 | 11 | `sheets/03_PART_A_-_General.tsv` |
| Part A Gen_139(8A) | hidden | 24 | 45 | 8 | `sheets/04_Part_A_Gen_139_8A.tsv` |
| Nature Of Business | visible | 5 | 30 | 16 | `sheets/05_Nature_Of_Business.tsv` |
| Part A - BS | visible | 92 | 284 | 24 | `sheets/06_Part_A_-_BS.tsv` |
| Manufacturing Account | visible | 26 | 79 | 6 | `sheets/07_Manufacturing_Account.tsv` |
| Trading Account | visible | 54 | 132 | 12 | `sheets/08_Trading_Account.tsv` |
| Profit and Loss | visible | 211 | 588 | 52 | `sheets/09_Profit_and_Loss.tsv` |
| Part A - OI | visible | 108 | 333 | 11 | `sheets/10_Part_A_-_OI.tsv` |
| Quantitative Details | visible | 22 | 67 | 0 | `sheets/11_Quantitative_Details.tsv` |
| Sheet1 | hidden | 21 | 66 | 16 | `sheets/12_Sheet1.tsv` |
| ITold | veryHidden | 22 | 86 | 64 | `sheets/13_ITold.tsv` |
| Schedule S | visible | 76 | 214 | 27 | `sheets/14_Schedule_S.tsv` |
| House Property | visible | 73 | 213 | 30 | `sheets/15_House_Property.tsv` |
| BP | visible | 164 | 544 | 81 | `sheets/16_BP.tsv` |
| DPM - DOA | visible | 52 | 321 | 105 | `sheets/17_DPM_-_DOA.tsv` |
| DEP_DCG | visible | 40 | 149 | 34 | `sheets/18_DEP_DCG.tsv` |
| ESR | visible | 13 | 41 | 5 | `sheets/19_ESR.tsv` |
| CG | visible | 534 | 2665 | 669 | `sheets/20_CG.tsv` |
| Schedule 112A | visible | 12 | 95 | 40 | `sheets/21_Schedule_112A.tsv` |
| HelpCSV | hidden | 26 | 48 | 0 | `sheets/22_HelpCSV.tsv` |
| Schedule 115AD(1)(iii) proviso | visible | 11 | 90 | 37 | `sheets/23_Schedule_115AD_1_iii_proviso.tsv` |
| OS | visible | 116 | 761 | 276 | `sheets/24_OS.tsv` |
| CYLA - BFLA | visible | 52 | 531 | 190 | `sheets/25_CYLA_-_BFLA.tsv` |
| CFL | visible | 23 | 150 | 40 | `sheets/26_CFL.tsv` |
| Unabsorbed Depreciation | visible | 17 | 57 | 20 | `sheets/27_Unabsorbed_Depreciation.tsv` |
| 80GGC | hidden | 10 | 34 | 18 | `sheets/28_80GGC.tsv` |
| VDA | visible | 7 | 24 | 7 | `sheets/29_VDA.tsv` |
| 80U-80DD | hidden | 6 | 34 | 6 | `sheets/30_80U-80DD.tsv` |
| ICDS | visible | 15 | 51 | 15 | `sheets/31_ICDS.tsv` |
| 10AA | hidden | 13 | 30 | 8 | `sheets/32_10AA.tsv` |
| 80G | hidden | 43 | 288 | 135 | `sheets/33_80G.tsv` |
| 80C | hidden | 8 | 15 | 3 | `sheets/34_80C.tsv` |
| 80D | hidden | 46 | 102 | 21 | `sheets/35_80D.tsv` |
| RA | hidden | 8 | 31 | 11 | `sheets/36_RA.tsv` |
| 80GGA | hidden | 10 | 46 | 21 | `sheets/37_80GGA.tsv` |
| 80 | hidden | 91 | 185 | 88 | `sheets/38_80.tsv` |
| 80E_80EE_80EEA_80EEB | hidden | 25 | 72 | 17 | `sheets/39_80E_80EE_80EEA_80EEB.tsv` |
| VI-A | visible | 64 | 330 | 123 | `sheets/40_VI-A.tsv` |
| SPI - SI - IF | visible | 148 | 1659 | 525 | `sheets/41_SPI_-_SI_-_IF.tsv` |
| AMT | hidden | 12 | 76 | 37 | `sheets/42_AMT.tsv` |
| AMTC | visible | 23 | 82 | 25 | `sheets/43_AMTC.tsv` |
| EI | visible | 33 | 96 | 15 | `sheets/44_EI.tsv` |
| FSI1 | hidden | 259 | 353 | 27 | `sheets/45_FSI1.tsv` |
| PTI | visible | 48 | 176 | 44 | `sheets/46_PTI.tsv` |
| TPSA | visible | 16 | 38 | 8 | `sheets/47_TPSA.tsv` |
| FSI | visible | 27 | 99 | 39 | `sheets/48_FSI.tsv` |
| TR_FA | visible | 110 | 279 | 49 | `sheets/49_TR_FA.tsv` |
| Sch 5A | visible | 12 | 33 | 4 | `sheets/50_Sch_5A.tsv` |
| AL | visible | 29 | 65 | 6 | `sheets/51_AL.tsv` |
| GST | visible | 7 | 10 | 3 | `sheets/52_GST.tsv` |
| Tax Calculated | hidden | 332 | 2887 | 1617 | `sheets/53_Tax_Calculated.tsv` |
| Part B - TI TTI | visible | 127 | 612 | 152 | `sheets/54_Part_B_-_TI_TTI.tsv` |
| IT | visible | 28 | 194 | 64 | `sheets/55_IT.tsv` |
| ESOP | visible | 18 | 74 | 14 | `sheets/56_ESOP.tsv` |
| Part B ATI | hidden | 41 | 135 | 24 | `sheets/57_Part_B_ATI.tsv` |
| TDS | visible | 53 | 225 | 44 | `sheets/58_TDS.tsv` |
| Verification | visible | 8 | 17 | 1 | `sheets/59_Verification.tsv` |
| OLDAL | veryHidden | 17 | 49 | 1 | `sheets/60_OLDAL.tsv` |
| Temporary Values | hidden | 49 | 192 | 5 | `sheets/61_Temporary_Values.tsv` |
| DropDownValues | hidden | 19302 | 52584 | 10 | `sheets/62_DropDownValues.tsv.gz` |
| CG Pop up_prefill | hidden | 6 | 9 | 0 | `sheets/63_CG_Pop_up_prefill.tsv` |
| SUMMARY | hidden | 40 | 103 | 45 | `sheets/64_SUMMARY.tsv` |
| BA | hidden | 8 | 23 | 2 | `sheets/65_BA.tsv` |
| Instructions | hidden | 121 | 287 | 1 | `sheets/66_Instructions.tsv` |
