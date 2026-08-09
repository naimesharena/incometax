# ITR2 — Excel utility extraction

Source: `ITR2_AY_26-27_V1.3/ITR2_AY_26-27_V1.3.xlsm` (10,242,993 bytes)

- Worksheets: **66** (25 visible, 41 hidden)
- Sheet contents (hidden + visible): **66 dumps** in `sheets/` (284,950 stored cell values, formulas extracted alongside)
- VBA modules: **155**, 236,664 lines of code
- Data-validation (dropdown) rules: **1122** (resolved to actual list values in `dropdowns.json`)
- Defined names: **6037**

## Worksheets

| # | Sheet | State | Dimensions | Formulas | Validations |
|---|-------|-------|-----------|----------|-------------|
| 1 | ISIN List | hidden | A1:C88568 | 0 | 0 |
| 2 | Home | visible | A1:J57 | 50 | 1 |
| 3 | PART A - General | visible | A1:XFC319 | 7 | 87 |
| 4 | Nature Of Business | hidden | A1:XFC65536 | 14 | 2 |
| 5 | Part A - BS | hidden | A1:P93 | 23 | 4 |
| 6 | Manufacturing Account | hidden | B1:V43 | 6 | 3 |
| 7 | Trading Account | hidden | B1:Q999 | 12 | 5 |
| 8 | Profit and Loss | hidden | A1:IU1000 | 81 | 105 |
| 9 | Part A - OI | hidden | A1:O1000 | 11 | 13 |
| 10 | Quantitative Details | hidden | A1:IU65538 | 0 | 13 |
| 11 | Sheet1 | hidden | A1:I26 | 15 | 1 |
| 12 | ITold | hidden | A1:V1000 | 64 | 6 |
| 13 | Part A Gen_139(8A) | hidden | A1:BB39 | 8 | 16 |
| 14 | Schedule S | visible | A1:U175 | 35 | 41 |
| 15 | House Property | visible | A1:Y484 | 54 | 50 |
| 16 | BP | hidden | A1:XFC1001 | 69 | 17 |
| 17 | DPM - DOA | hidden | A1:XFC1000 | 120 | 47 |
| 18 | DEP_DCG | hidden | A1:L46 | 34 | 2 |
| 19 | ESR | hidden | A1:I1000 | 12 | 3 |
| 20 | CG | visible | A1:CG1112 | 689 | 120 |
| 21 | Schedule 112A | visible | A1:IW17 | 39 | 17 |
| 22 | HelpCSV | hidden | A1:XFC128 | 0 | 0 |
| 23 | Schedule 115AD(1)(iii) proviso | visible | A1:IW17 | 39 | 16 |
| 24 | VDA | visible | A1:P27 | 8 | 6 |
| 25 | OS | visible | A1:BC1019 | 385 | 37 |
| 26 | CYLA - BFLA | visible | A1:BY65544 | 225 | 3 |
| 27 | CFL | visible | A1:AD65542 | 70 | 4 |
| 28 | Unabsorbed Depreciation | hidden | A1:XFC25 | 34 | 10 |
| 29 | ICDS | hidden | A1:J19 | 3 | 1 |
| 30 | 10AA | hidden | A1:IV1000 | 8 | 4 |
| 31 | 80C | hidden | A1:K11 | 4 | 5 |
| 32 | 80G | hidden | A1:IX72 | 144 | 17 |
| 33 | 80D | hidden | B1:M76 | 18 | 20 |
| 34 | RA | hidden | A1:XFC15 | 11 | 9 |
| 35 | 80GGA | hidden | A1:XFC19 | 37 | 9 |
| 36 | 80 | hidden | A1:IU1000 | 85 | 24 |
| 37 | 80E_80EE_80EEA_80EEB | hidden | B1:AC44 | 13 | 16 |
| 38 | VI-A | visible | A1:XFC1048576 | 116 | 19 |
| 39 | 80U-80DD | hidden | C1:AD19 | 6 | 15 |
| 40 | 80GGC | hidden | A1:S35 | 15 | 7 |
| 41 | AMT | hidden | A1:XFC15 | 36 | 7 |
| 42 | SPI - SI | visible | A1:XFC173 | 784 | 22 |
| 43 | AMTC | visible | A1:IV27 | 38 | 12 |
| 44 | EI | visible | A1:IV1013 | 16 | 22 |
| 45 | FSI1 | hidden | A1:AL294 | 45 | 10 |
| 46 | PTI | visible | A1:IU54 | 62 | 9 |
| 47 | TPSA | hidden | B1:Q36 | 8 | 11 |
| 48 | FSI | visible | A1:S43 | 35 | 9 |
| 49 | Sch 5A | visible | A1:IV1000 | 4 | 9 |
| 50 | TR_FA | visible | A1:XFC170 | 49 | 71 |
| 51 | AL | visible | A1:R65536 | 6 | 29 |
| 52 | GST | hidden | A1:IU65522 | 3 | 2 |
| 53 | Tax Calculated | hidden | A1:DM379 | 1981 | 8 |
| 54 | Part B - TI TTI | visible | A1:AS138 | 157 | 30 |
| 55 | IT | visible | A1:AG33 | 161 | 5 |
| 56 | ESOP | visible | A1:XFC83 | 18 | 14 |
| 57 | TDS | visible | A1:IW161 | 40 | 52 |
| 58 | Part B ATI | hidden | A1:T59 | 26 | 12 |
| 59 | Verification | visible | B1:O37 | 1 | 4 |
| 60 | CG Pop up_prefill | hidden | A1:D6 | 0 | 0 |
| 61 | OLDAL | veryHidden | A1:K19 | 1 | 2 |
| 62 | Temporary Values | hidden | A1:F50 | 5 | 0 |
| 63 | DropDownValues | hidden | A1:GK19302 | 103 | 0 |
| 64 | SUMMARY | hidden | A1:IU77 | 45 | 0 |
| 65 | BA | hidden | A1:IU12 | 2 | 7 |
| 66 | Instructions | hidden | A1:IU128 | 6 | 0 |

## VBA modules

| Module | Lines | File |
|--------|-------|------|
| AY23_24Changes.bas | 1806 | `vba/AY23_24Changes.bas` |
| BFLA_Calculations.bas | 411 | `vba/BFLA_Calculations.bas` |
| Calculations.bas | 231 | `vba/Calculations.bas` |
| CG_Calc.bas | 985 | `vba/CG_Calc.bas` |
| CGDeductions.bas | 2984 | `vba/CGDeductions.bas` |
| checkBoxModule.bas | 33 | `vba/checkBoxModule.bas` |
| CYLACalculations.bas | 8837 | `vba/CYLACalculations.bas` |
| DEP_DCG.bas | 302 | `vba/DEP_DCG.bas` |
| EfilingCommon.bas | 5674 | `vba/EfilingCommon.bas` |
| ePayPrefill.bas | 8 | `vba/ePayPrefill.bas` |
| FilingSectRadioButton.bas | 75 | `vba/FilingSectRadioButton.bas` |
| Generate_XML.bas | 18334 | `vba/Generate_XML.bas` |
| GenerateJson.bas | 36879 | `vba/GenerateJson.bas` |
| GST.bas | 199 | `vba/GST.bas` |
| HS256.cls | 535 | `vba/HS256.cls` |
| ImportJson.bas | 30756 | `vba/ImportJson.bas` |
| ImportPrefill.bas | 15135 | `vba/ImportPrefill.bas` |
| ImportSchedule112A.bas | 348 | `vba/ImportSchedule112A.bas` |
| ImportSchedule115AD.bas | 354 | `vba/ImportSchedule115AD.bas` |
| Initializer.bas | 74 | `vba/Initializer.bas` |
| Manufacturing_Account.bas | 6 | `vba/Manufacturing_Account.bas` |
| md112A.bas | 1275 | `vba/md112A.bas` |
| md115AD.bas | 1176 | `vba/md115AD.bas` |
| md80_.bas | 1943 | `vba/md80_.bas` |
| md80C.bas | 202 | `vba/md80C.bas` |
| md80E.bas | 693 | `vba/md80E.bas` |
| md80EE.bas | 877 | `vba/md80EE.bas` |
| md80EEA.bas | 858 | `vba/md80EEA.bas` |
| md80EEB.bas | 675 | `vba/md80EEB.bas` |
| mdAMT.bas | 175 | `vba/mdAMT.bas` |
| mdAMTC.bas | 1440 | `vba/mdAMTC.bas` |
| mdATI.bas | 772 | `vba/mdATI.bas` |
| mdBA.bas | 1214 | `vba/mdBA.bas` |
| mdCFL.bas | 1052 | `vba/mdCFL.bas` |
| mdEI.bas | 1182 | `vba/mdEI.bas` |
| mdESR.bas | 274 | `vba/mdESR.bas` |
| mdGen139_8A.bas | 405 | `vba/mdGen139_8A.bas` |
| mdHashing.bas | 98 | `vba/mdHashing.bas` |
| mdICDS.bas | 47 | `vba/mdICDS.bas` |
| mdImportXL.bas | 3637 | `vba/mdImportXL.bas` |
| mdImportXML.bas | 2580 | `vba/mdImportXML.bas` |
| mdInterestCalc.bas | 711 | `vba/mdInterestCalc.bas` |
| mdOI.bas | 2185 | `vba/mdOI.bas` |
| mdProgressbar.bas | 91 | `vba/mdProgressbar.bas` |
| mdPTI.bas | 1679 | `vba/mdPTI.bas` |
| mdVDA.bas | 206 | `vba/mdVDA.bas` |
| MessageBox.frm | 14 | `vba/MessageBox.frm` |
| mfMessage.bas | 37 | `vba/mfMessage.bas` |
| Module1.bas | 5 | `vba/Module1.bas` |
| Module2.bas | 60 | `vba/Module2.bas` |
| Module3.bas | 56 | `vba/Module3.bas` |
| Module4.bas | 170 | `vba/Module4.bas` |
| msSI.bas | 402 | `vba/msSI.bas` |
| Nature_of_business.bas | 121 | `vba/Nature_of_business.bas` |
| ParseJson.bas | 347 | `vba/ParseJson.bas` |
| Part_A_General.bas | 8081 | `vba/Part_A_General.bas` |
| PARTA_BS.bas | 210 | `vba/PARTA_BS.bas` |
| Profit_Loss.bas | 1921 | `vba/Profit_Loss.bas` |
| pwd.bas | 86 | `vba/pwd.bas` |
| Quantitative_Details.bas | 875 | `vba/Quantitative_Details.bas` |
| Sch10A.bas | 207 | `vba/Sch10A.bas` |
| Sch5A.bas | 496 | `vba/Sch5A.bas` |
| Sch80D.bas | 1166 | `vba/Sch80D.bas` |
| Sch80G.bas | 2773 | `vba/Sch80G.bas` |
| Sch80GGC.bas | 833 | `vba/Sch80GGC.bas` |
| Sch80U_80DD.bas | 1168 | `vba/Sch80U_80DD.bas` |
| SchAL.bas | 1735 | `vba/SchAL.bas` |
| SchBP.bas | 330 | `vba/SchBP.bas` |
| SchCG.bas | 9176 | `vba/SchCG.bas` |
| SchDPM_DOA.bas | 2055 | `vba/SchDPM_DOA.bas` |
| SchFSI.bas | 917 | `vba/SchFSI.bas` |
| SchHP.bas | 3359 | `vba/SchHP.bas` |
| SchIT.bas | 1023 | `vba/SchIT.bas` |
| SchOI.bas | 81 | `vba/SchOI.bas` |
| SchOS.bas | 2741 | `vba/SchOS.bas` |
| SchRA.bas | 1012 | `vba/SchRA.bas` |
| SchTDS.bas | 4016 | `vba/SchTDS.bas` |
| SchTPSA.bas | 305 | `vba/SchTPSA.bas` |
| SchTR_FA.bas | 5679 | `vba/SchTR_FA.bas` |
| SchUD.bas | 545 | `vba/SchUD.bas` |
| Sheet1.cls | 7762 | `vba/Sheet1.cls` |
| Sheet10.cls | 231 | `vba/Sheet10.cls` |
| Sheet11.cls | 41 | `vba/Sheet11.cls` |
| Sheet12.cls | 43 | `vba/Sheet12.cls` |
| Sheet13.cls | 2512 | `vba/Sheet13.cls` |
| Sheet14.cls | 1265 | `vba/Sheet14.cls` |
| Sheet15.cls | 80 | `vba/Sheet15.cls` |
| Sheet16.cls | 97 | `vba/Sheet16.cls` |
| Sheet17.cls | 102 | `vba/Sheet17.cls` |
| Sheet18.cls | 62 | `vba/Sheet18.cls` |
| Sheet19.cls | 442 | `vba/Sheet19.cls` |
| Sheet2.cls | 571 | `vba/Sheet2.cls` |
| Sheet20.cls | 986 | `vba/Sheet20.cls` |
| Sheet201.cls | 888 | `vba/Sheet201.cls` |
| Sheet202.cls | 242 | `vba/Sheet202.cls` |
| Sheet21.cls | 124 | `vba/Sheet21.cls` |
| Sheet22.cls | 62 | `vba/Sheet22.cls` |
| Sheet23.cls | 113 | `vba/Sheet23.cls` |
| Sheet24.cls | 406 | `vba/Sheet24.cls` |
| Sheet25.cls | 271 | `vba/Sheet25.cls` |
| Sheet26.cls | 256 | `vba/Sheet26.cls` |
| Sheet27.cls | 402 | `vba/Sheet27.cls` |
| Sheet28.cls | 1142 | `vba/Sheet28.cls` |
| Sheet29.cls | 66 | `vba/Sheet29.cls` |
| Sheet3.cls | 67 | `vba/Sheet3.cls` |
| Sheet30.cls | 40 | `vba/Sheet30.cls` |
| Sheet31.cls | 11 | `vba/Sheet31.cls` |
| Sheet32.cls | 11 | `vba/Sheet32.cls` |
| Sheet33.cls | 18 | `vba/Sheet33.cls` |
| Sheet34.cls | 18 | `vba/Sheet34.cls` |
| Sheet35.cls | 20 | `vba/Sheet35.cls` |
| Sheet36.cls | 45 | `vba/Sheet36.cls` |
| Sheet37.cls | 134 | `vba/Sheet37.cls` |
| Sheet38.cls | 174 | `vba/Sheet38.cls` |
| Sheet39.cls | 2402 | `vba/Sheet39.cls` |
| Sheet4.cls | 157 | `vba/Sheet4.cls` |
| Sheet40.cls | 532 | `vba/Sheet40.cls` |
| Sheet41.cls | 113 | `vba/Sheet41.cls` |
| Sheet42.cls | 39 | `vba/Sheet42.cls` |
| Sheet43.cls | 263 | `vba/Sheet43.cls` |
| Sheet44.cls | 15 | `vba/Sheet44.cls` |
| Sheet45.cls | 18 | `vba/Sheet45.cls` |
| Sheet46.cls | 371 | `vba/Sheet46.cls` |
| Sheet47.cls | 142 | `vba/Sheet47.cls` |
| Sheet48.cls | 42 | `vba/Sheet48.cls` |
| Sheet49.cls | 62 | `vba/Sheet49.cls` |
| Sheet5.cls | 192 | `vba/Sheet5.cls` |
| Sheet50.cls | 103 | `vba/Sheet50.cls` |
| Sheet51.cls | 70 | `vba/Sheet51.cls` |
| Sheet52.cls | 191 | `vba/Sheet52.cls` |
| Sheet53.cls | 664 | `vba/Sheet53.cls` |
| Sheet54.cls | 673 | `vba/Sheet54.cls` |
| Sheet55.cls | 689 | `vba/Sheet55.cls` |
| Sheet56.cls | 112 | `vba/Sheet56.cls` |
| Sheet57.cls | 89 | `vba/Sheet57.cls` |
| Sheet58.cls | 423 | `vba/Sheet58.cls` |
| Sheet59.cls | 336 | `vba/Sheet59.cls` |
| Sheet6.cls | 42 | `vba/Sheet6.cls` |
| Sheet60.cls | 715 | `vba/Sheet60.cls` |
| Sheet61.cls | 18 | `vba/Sheet61.cls` |
| Sheet62.cls | 9 | `vba/Sheet62.cls` |
| Sheet7.cls | 445 | `vba/Sheet7.cls` |
| Sheet8.cls | 997 | `vba/Sheet8.cls` |
| Sheet9.cls | 354 | `vba/Sheet9.cls` |
| Sheet98.cls | 97 | `vba/Sheet98.cls` |
| SheetALL.cls | 14 | `vba/SheetALL.cls` |
| SPI_SI.bas | 709 | `vba/SPI_SI.bas` |
| Tax_Calc.bas | 1499 | `vba/Tax_Calc.bas` |
| ThisWorkbook.cls | 46 | `vba/ThisWorkbook.cls` |
| TI_TTI_Salary.bas | 3313 | `vba/TI_TTI_Salary.bas` |
| Trading_Account.bas | 276 | `vba/Trading_Account.bas` |
| UserForm1.frm | 18 | `vba/UserForm1.frm` |
| UserForm2.frm | 16 | `vba/UserForm2.frm` |
| UserForm3.frm | 82 | `vba/UserForm3.frm` |
| Verification.bas | 197 | `vba/Verification.bas` |

## Defined names

- `_xlnm._FilterDatabase` → 'ISIN List'!$A$1:$B$1
- `_Ins1` → HelpCSV!$K$14
- `_Ins1` → Instructions!$K$14
- `_Ins10` → HelpCSV!$K$59
- `_Ins10` → Instructions!$K$59
- `_Ins11` → HelpCSV!$K$69
- `_Ins11` → Instructions!$K$69
- `_Ins2` → HelpCSV!$K$22
- `_Ins2` → Instructions!$K$22
- `_Ins3` → HelpCSV!$K$26
- `_Ins3` → Instructions!$K$26
- `_Ins4` → HelpCSV!$K$30
- `_Ins4` → Instructions!$K$30
- `_Ins5` → HelpCSV!$K$40
- `_Ins5` → Instructions!$K$40
- `_Ins6` → HelpCSV!$K$43
- `_Ins6` → Instructions!$K$43
- `_Ins7` → HelpCSV!$K$46
- `_Ins7` → Instructions!$K$46
- `_Ins8` → HelpCSV!$K$49
- `_Ins8` → Instructions!$K$49
- `_Ins9` → HelpCSV!$K$56
- `_Ins9` → Instructions!$K$56
- `_xleta.AND` → #NAME?
- `_xleta.IF` → #NAME?
- `_xleta.MID` → #NAME?
- `_xleta.MIN` → #NAME?
- `_xleta.RATE` → #NAME?
- `_xleta.TRUNC` → #NAME?
- `_xleta.VALUE` → #NAME?
- `AA10.AssmtYrUnit` → '10AA'!$H$11:$H$13
- `AA10.DedFromUndertaking` → '10AA'!$I$11:$I$13
- `AA10.TotalDedUs10Sub` → '10AA'!$I$14
- `AA10_undertaking` → '10AA'!$F$11:$F$13
- `Aadhaardependent_80DD` → '80U-80DD'!$I$16
- `Aadhar_TDS3` → TDS!$G$39:$G$42
- `Aadhar1` → 'PART A - General'!$BB$63:$BB$66
- `Aadhar2` → 'PART A - General'!#REF!
- `AccLTCG.Up16Of12To15Of3` → CG!$Q$534
- `AccLTCG.Up16Of3To31Of3` → CG!$S$534
- `AccLTCG.Up16Of9To15Of12` → CG!$O$534
- `AccLTCG.Upto15Of6` → CG!$L$534
- `AccLTCG.Upto15Of9` → CG!$M$534
- `AccLTCG12.5.Up16Of12To15Of3` → CG!$Q$535
- `AccLTCG12.5.Up16Of3To31Of3` → CG!$S$535
- `AccLTCG12.5.Up16Of9To15Of12` → CG!$O$535
- `AccLTCG12.5.Upto15Of6` → CG!$L$535
- `AccLTCG12.5.Upto15Of9` → CG!$M$535
- `AccLTCGDTAA.Up16Of12To15Of3` → CG!$Q$537
- `AccLTCGDTAA.Up16Of3To31Of3` → CG!$S$537
- `AccLTCGDTAA.Up16Of9To15Of12` → CG!$O$537
- `AccLTCGDTAA.Upto15Of6` → CG!$L$537
- `AccLTCGDTAA.Upto15Of9` → CG!$M$537
- `AccLTCGNP.Up16Of12To15Of3` → CG!$Q$536
- `AccLTCGNP.Up16Of3To31Of3` → CG!$S$536
- `AccLTCGNP.Up16Of9To15Of12` → CG!$O$536
- `AccLTCGNP.Upto15Of6` → CG!$L$536
- `AccLTCGNP.Upto15Of9` → CG!$M$536
- `AccSTCG.Up16Of12To15Of3` → CG!$Q$529
- `AccSTCG.Up16Of3To31Of3` → CG!$S$529
- `AccSTCG.Up16Of9To15Of12` → CG!$O$529
- `AccSTCG.Upto15Of6` → CG!$L$529
- `AccSTCG.Upto15Of9` → CG!$M$529
- `ACCSTCG20.Up16Of12To15Of3` → CG!$Q$530
- `ACCSTCG20.Up16Of3To31Of3` → CG!$S$530
- `ACCSTCG20.Up16Of9To15Of12` → CG!$O$530
- `AccSTCG20.Upto15Of6` → CG!$L$530
- `AccSTCG20.Upto15Of9` → CG!$M$530
- `ACCSTCG30.Up16Of12To15Of3` → CG!$Q$531
- `ACCSTCG30.Up16Of3To31Of3` → CG!$S$531
- `ACCSTCG30.Up16Of9To15Of12` → CG!$O$531
- `AccSTCG30.Upto15Of6` → CG!$L$531
- `AccSTCG30.Upto15Of9` → CG!$M$531
- `AccSTCGDTAA.Up16Of12To15Of3` → CG!$Q$533
- `AccSTCGDTAA.Up16Of3To31Of3` → CG!$S$533
- `AccSTCGDTAA.Up16Of9To15Of12` → CG!$O$533
- `AccSTCGDTAA.Upto15Of6` → CG!$L$533
- `AccSTCGDTAA.Upto15Of9` → CG!$M$533
- `AccSTCGOTH.Up16Of12To15Of3` → CG!$Q$532
- `AccSTCGOTH.Up16Of3To31Of3` → CG!$S$532
- `AccSTCGOTH.Up16Of9To15Of12` → CG!$O$532
- `AccSTCGOTH.Upto15Of6` → CG!$L$532
- `AccSTCGOTH.Upto15Of9` → CG!$M$532
- `AccVDA.Up16Of12To15Of3` → CG!$Q$538
- `AccVDA.Up16Of3To31Of3` → CG!$S$538
- `AccVDA.Up16Of9To15Of12` → CG!$O$538
- `AccVDA.Upto15Of6` → CG!$L$538
- `AccVDA.Upto15Of9` → CG!$M$538
- `AckNoFm10IAfiled_80DD` → '80U-80DD'!$K$16
- `AckNoFm10IAfiled_80U` → '80U-80DD'!$H$6
- `Acknowledge_Num11A2_80U` → '80U-80DD'!$L$5
- `Acknowledge_Num11A2_80U1` → '80U-80DD'!$L$5
- `AcknowledgeNum11A2_80DD` → '80U-80DD'!$L$16
- `AcknowledgeNum11A2_80U` → '80U-80DD'!$I$6
- `acknowledgmentnumber` → 'PART A - General'!$AK$53
- `Address_80GGA` → RA!$F$7:$F$10
- `Address_80GGANew` → '80GGA'!$G$7:$G$14
- `adjstGTI` → '80G'!$W$21
- `adjtotloss.BusLossOthThanSpecLossCF9` → CFL!$L$23
- `adjtotloss.HPLossCF9` → CFL!$G$23
- `adjtotloss.LossFrmSpecBusCF9` → CFL!$M$23
- `adjtotloss.LossFrmSpecifiedBusCF9` → CFL!$N$23
- `adjtotloss.LTCGLossCF9` → CFL!$R$23
- `adjtotloss.OthSrcLossRaceHorseCF9` → CFL!$U$23
- `adjtotloss.STCGLossCF9` → CFL!$O$23
- `AfterDateReturn` → DropDownValues!$DL$2:$DL$5
- `age` → 'Temporary Values'!$B$42
- `Age_TX` → Sheet1!$G$8
- `Agri_dropdown` → DropDownValues!$FV$95:$FV$98
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
- `All` → DropDownValues!$BR$13
- `All_Pincode_list` → DropDownValues!$FG$3:$FG$19302
- `All_pincode_list_Old` → DropDownValues!$DU$2:$DU$19301
- `All_Pincode_V` → DropDownValues!$FG$3:$FH$19302
- `All_pincode_V_O` → DropDownValues!$DU$2:$DV$19301
- `AllBS.names` → 'Part A - BS'!$J$8,'Part A - BS'!$L$5,'Part A - BS'!$J$7,'Part A - BS'!$J$8,'Part A - BS'!$J$9,'Part A - BS'!$J$10,'Part A - BS'!$J$15,'Part A - BS'!$J$17,'Part A - BS'!$J$18,'Part A - BS'!$J$22,'Part A - BS'!$J$23,'Part A - BS'!$L$26,'Part A - BS'!$J$29,'Part A - BS'!$J$30,'Part A - BS'!$J$32,'Part A - BS'!$J$36,'Part A - BS'!$J$37,'Part A - BS'!$J$40,'Part A - BS'!$J$41,'Part A - BS'!$J$42,'Part A - BS'!$J$48,'Part A - BS'!$J$49,'Part A - BS'!$J$50,'Part A - BS'!$J$51,'Part A - BS'!$L$53,'Part A - BS'!$J$55,'Part A - BS'!$J$56,'Part A - BS'!$L$58,'Part A - BS'!$J$61,'Part A - BS'!$J$62,'Part A - BS'!$J$63,'Part A - BS'!$J$68,'Part A - BS'!$J$69,'Part A - BS'!$J$70,'Part A - BS'!$J$71,'Part A - BS'!$J$74,'Part A - BS'!$J$76,'Part A - BS'!$J$77,'Part A - BS'!$J$81,'Part A - BS'!$J$82,'Part A - BS'!$J$83,'Part A - BS'!$L$87
- `ALLMA_NAMES` → 'Manufacturing Account'!$R$6:$T$7,'Manufacturing Account'!$R$9:$T$10,'Manufacturing Account'!$R$12:$T$14,'Manufacturing Account'!$R$16:$T$21,'Manufacturing Account'!$R$25:$T$26
- `Allow3ai` → OS!$U$96
- `AllowanceBACNo` → DropDownValues!$DQ$3:$DQ$17
- `AllowanceBACNoAllPGov` → DropDownValues!$DQ$75:$DQ$86
- `AllowanceBACNoCGSG` → DropDownValues!$DQ$41:$DQ$55
- `AllowanceBACNoCGSGPCGPSG` → DropDownValues!$DS$3:$DS$15
- `AllowanceBACNoPCGPSG` → DropDownValues!$DQ$59:$DQ$70
- `AllowanceBACYes` → DropDownValues!$DQ$24:$DQ$35
- `AllowanceBACYesAllPGov` → DropDownValues!$DS$72:$DS$81
- `AllowanceBACYesCGSG` → DropDownValues!$DS$41:$DS$50
- `AllowanceBACYesCGSGPCGPSG` → DropDownValues!$DS$24:$DS$33
- `AllowanceBACYesPCGPSG` → DropDownValues!$DS$57:$DS$66
- `Allpl.name` → 'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!$J$6:$J$15,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!#REF!,'Profit and Loss'!$L$26:$L$31,'Profit and Loss'!$J$33:$J$42,'Profit and Loss'!$J$44:$J$45,'Profit and Loss'!$J$47:$J$50,'Profit and Loss'!$L$52:$L$57,'Profit and Loss'!$J$59:$J$60,'Profit and Loss'!$J$63:$J$64,'Profit and Loss'!$J$67:$J$68,'Profit and Loss'!$L$70:$L$80,'Profit and Loss'!$J$82:$J$90,'Profit and Loss'!$L$92,'Profit and Loss'!$J$112:$J$120
- `ALLPL_NAMES1` → 'Profit and Loss'!$J$6,'Profit and Loss'!$J$6:$J$15,'Profit and Loss'!$H$18:$H$21,'Profit and Loss'!$J$18:$J$21,'Profit and Loss'!$L$26:$L$31,'Profit and Loss'!$J$33:$J$42,'Profit and Loss'!$J$44:$J$45,'Profit and Loss'!$J$47:$J$50,'Profit and Loss'!$L$52:$L$57,'Profit and Loss'!$J$59:$J$60,'Profit and Loss'!$J$63:$J$64,'Profit and Loss'!$J$67:$J$68,'Profit and Loss'!$L$70:$L$80,'Profit and Loss'!$J$82:$J$90,'Profit and Loss'!$L$92,'Profit and Loss'!$H$95:$H$98,'Profit and Loss'!$J$95:$J$98,'Profit and Loss'!$H$103:$H$109,'Profit and Loss'!$J$103:$J$109
- `ALLPL_NAMES2` → 'Profit and Loss'!$L$122,'Profit and Loss'!$L$122:$L$123,'Profit and Loss'!$J$126:$J$127,'Profit and Loss'!$L$129,'Profit and Loss'!$L$131:$L$132,'Profit and Loss'!$L$134,'Profit and Loss'!$L$136,'Profit and Loss'!$E$140:$L$142,'Profit and Loss'!$K$144:$L$145,'Profit and Loss'!$K$147:$L$148,'Profit and Loss'!$E$152:$L$154,'Profit and Loss'!$K$155:$L$156,'Profit and Loss'!$E$160:$L$162,'Profit and Loss'!$E$165:$L$168,'Profit and Loss'!$L$182:$L$185,'Profit and Loss'!$L$189:$L$192,'Profit and Loss'!$L$195:$L$197
- `ALLPL_NAMES3` → 'Profit and Loss'!$F$114:$Q$115,'Profit and Loss'!$J$120
- `AllTA_Names` → 'Trading Account'!$K$6,'Trading Account'!$K$6,'Trading Account'!$K$7,'Trading Account'!$G$10:$H$11,'Trading Account'!$K$16,'Trading Account'!$K$18:$L$25,'Trading Account'!$K$28,'Trading Account'!$K$30:$L$31,'Trading Account'!$K$33:$L$34,'Trading Account'!$G$37:$H$38,'Trading Account'!$K$43:$L$53
- `Alternate_Address` → 'PART A - General'!#REF!
- `Amount.80C` → '80C'!$D$5:$D$8
- `Amount_80CCC` → 'VI-A'!$G$9:$G$10
- `Amount_80CCD1` → 'VI-A'!$G$14:$G$15
- `Amount_80CCD1B` → 'VI-A'!$G$19:$G$20
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
- `Amt_1` → BP!$I$38:$I$40
- `Amt_Condn` → AMT!$M$3
- `AmtA1.80D` → '80D'!$J$10:$L$12
- `AmtA2.80D` → '80D'!$J$34:$L$36
- `AmtB1.80D` → '80D'!$J$21:$L$23
- `AmtB2.80D` → '80D'!$J$45:$L$47
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

## Companion CSV templates

- `ITR2_AY_26-27_V1.3/112A & 115AD/CSV_112A.csv` — header: `"Share/Unit acquired (On or before / after 31st Jan 2018)`
- `ITR2_AY_26-27_V1.3/112A & 115AD/CSV_115AD.csv` — header: `"Share/Unit acquired (On or before/ after 31st Jan 2018)`
- `ITR2_AY_26-27_V1.3/TCS & TDS & IT/CSV_IT.csv` — header: `"BSR Code`
- `ITR2_AY_26-27_V1.3/TCS & TDS & IT/CSV_TCS.csv` — header: `TCS credit relating to Self/Other Person [Spouse as per section 5A/Other person as per Rule 37i(1)] (Col 2(i)),"TAN of the`
- `ITR2_AY_26-27_V1.3/TCS & TDS & IT/CSV_TDS1.csv` — header: `"Tax Deduction Account`
- `ITR2_AY_26-27_V1.3/TCS & TDS & IT/CSV_TDS2.csv` — header: `"TDS Credit related to self/ other person as per rule 37BA(2) `
- `ITR2_AY_26-27_V1.3/TCS & TDS & IT/CSV_TDS3.csv` — header: `TDS credit in the name of  (Col 2),"PAN Of Other Person (If TDS Credit related to other person)`

## Sheet dumps (`sheets/`) — hidden and visible

| Sheet | State | Rows with values | Cells | Formulas | File |
|-------|-------|------------------|-------|----------|------|
| ISIN List | hidden | 88568 | 177136 | 0 | `sheets/01_ISIN_List.tsv.gz` |
| Home | visible | 52 | 307 | 12 | `sheets/02_Home.tsv` |
| PART A - General | visible | 87 | 192 | 7 | `sheets/03_PART_A_-_General.tsv` |
| Nature Of Business | hidden | 5 | 27 | 14 | `sheets/04_Nature_Of_Business.tsv` |
| Part A - BS | hidden | 88 | 272 | 23 | `sheets/05_Part_A_-_BS.tsv` |
| Manufacturing Account | hidden | 26 | 79 | 6 | `sheets/06_Manufacturing_Account.tsv` |
| Trading Account | hidden | 50 | 120 | 12 | `sheets/07_Trading_Account.tsv` |
| Profit and Loss | hidden | 191 | 531 | 50 | `sheets/08_Profit_and_Loss.tsv` |
| Part A - OI | hidden | 105 | 324 | 11 | `sheets/09_Part_A_-_OI.tsv` |
| Quantitative Details | hidden | 69 | 114 | 0 | `sheets/10_Quantitative_Details.tsv` |
| Sheet1 | hidden | 21 | 65 | 15 | `sheets/11_Sheet1.tsv` |
| ITold | hidden | 22 | 86 | 64 | `sheets/12_ITold.tsv` |
| Part A Gen_139(8A) | hidden | 24 | 45 | 8 | `sheets/13_Part_A_Gen_139_8A.tsv` |
| Schedule S | visible | 100 | 272 | 35 | `sheets/14_Schedule_S.tsv` |
| House Property | visible | 73 | 220 | 34 | `sheets/15_House_Property.tsv` |
| BP | hidden | 155 | 505 | 69 | `sheets/16_BP.tsv` |
| DPM - DOA | hidden | 50 | 304 | 100 | `sheets/17_DPM_-_DOA.tsv` |
| DEP_DCG | hidden | 40 | 149 | 34 | `sheets/18_DEP_DCG.tsv` |
| ESR | hidden | 13 | 41 | 5 | `sheets/19_ESR.tsv` |
| CG | visible | 525 | 2501 | 598 | `sheets/20_CG.tsv` |
| Schedule 112A | visible | 11 | 92 | 29 | `sheets/21_Schedule_112A.tsv` |
| HelpCSV | hidden | 26 | 48 | 0 | `sheets/22_HelpCSV.tsv` |
| Schedule 115AD(1)(iii) proviso | visible | 11 | 92 | 29 | `sheets/23_Schedule_115AD_1_iii_proviso.tsv` |
| VDA | visible | 7 | 23 | 8 | `sheets/24_VDA.tsv` |
| OS | visible | 121 | 823 | 289 | `sheets/25_OS.tsv` |
| CYLA - BFLA | visible | 50 | 513 | 186 | `sheets/26_CYLA_-_BFLA.tsv` |
| CFL | visible | 23 | 156 | 51 | `sheets/27_CFL.tsv` |
| Unabsorbed Depreciation | hidden | 15 | 51 | 12 | `sheets/28_Unabsorbed_Depreciation.tsv` |
| ICDS | hidden | 15 | 35 | 3 | `sheets/29_ICDS.tsv` |
| 10AA | hidden | 13 | 30 | 8 | `sheets/30_10AA.tsv` |
| 80C | hidden | 8 | 13 | 4 | `sheets/31_80C.tsv` |
| 80G | hidden | 37 | 244 | 136 | `sheets/32_80G.tsv` |
| 80D | hidden | 46 | 99 | 18 | `sheets/33_80D.tsv` |
| RA | hidden | 8 | 31 | 11 | `sheets/34_RA.tsv` |
| 80GGA | hidden | 14 | 70 | 21 | `sheets/35_80GGA.tsv` |
| 80 | hidden | 88 | 179 | 85 | `sheets/36_80.tsv` |
| 80E_80EE_80EEA_80EEB | hidden | 26 | 117 | 13 | `sheets/37_80E_80EE_80EEA_80EEB.tsv` |
| VI-A | visible | 62 | 332 | 116 | `sheets/38_VI-A.tsv` |
| 80U-80DD | hidden | 6 | 33 | 6 | `sheets/39_80U-80DD.tsv` |
| 80GGC | hidden | 9 | 31 | 12 | `sheets/40_80GGC.tsv` |
| AMT | hidden | 12 | 76 | 36 | `sheets/41_AMT.tsv` |
| SPI - SI | visible | 123 | 1378 | 287 | `sheets/42_SPI_-_SI.tsv` |
| AMTC | visible | 23 | 95 | 38 | `sheets/43_AMTC.tsv` |
| EI | visible | 33 | 97 | 16 | `sheets/44_EI.tsv` |
| FSI1 | hidden | 259 | 353 | 27 | `sheets/45_FSI1.tsv` |
| PTI | visible | 48 | 176 | 44 | `sheets/46_PTI.tsv` |
| TPSA | hidden | 16 | 38 | 8 | `sheets/47_TPSA.tsv` |
| FSI | visible | 23 | 87 | 35 | `sheets/48_FSI.tsv` |
| Sch 5A | visible | 12 | 33 | 4 | `sheets/49_Sch_5A.tsv` |
| TR_FA | visible | 110 | 395 | 49 | `sheets/50_TR_FA.tsv` |
| AL | visible | 29 | 65 | 6 | `sheets/51_AL.tsv` |
| GST | hidden | 7 | 10 | 3 | `sheets/52_GST.tsv` |
| Tax Calculated | hidden | 336 | 3118 | 1599 | `sheets/53_Tax_Calculated.tsv` |
| Part B - TI TTI | visible | 127 | 622 | 157 | `sheets/54_Part_B_-_TI_TTI.tsv` |
| IT | visible | 29 | 192 | 64 | `sheets/55_IT.tsv` |
| ESOP | visible | 18 | 60 | 16 | `sheets/56_ESOP.tsv` |
| TDS | visible | 59 | 211 | 38 | `sheets/57_TDS.tsv` |
| Part B ATI | hidden | 45 | 136 | 24 | `sheets/58_Part_B_ATI.tsv` |
| Verification | visible | 8 | 15 | 1 | `sheets/59_Verification.tsv` |
| CG Pop up_prefill | hidden | 6 | 9 | 0 | `sheets/60_CG_Pop_up_prefill.tsv` |
| OLDAL | veryHidden | 17 | 49 | 1 | `sheets/61_OLDAL.tsv` |
| Temporary Values | hidden | 44 | 167 | 5 | `sheets/62_Temporary_Values.tsv` |
| DropDownValues | hidden | 19302 | 90853 | 10 | `sheets/63_DropDownValues.tsv.gz` |
| SUMMARY | hidden | 40 | 103 | 45 | `sheets/64_SUMMARY.tsv` |
| BA | hidden | 8 | 23 | 2 | `sheets/65_BA.tsv` |
| Instructions | hidden | 121 | 287 | 1 | `sheets/66_Instructions.tsv` |
