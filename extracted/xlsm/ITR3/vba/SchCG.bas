Attribute VB_Name = "SchCG"
Option Explicit

Public RepSec92EFlag_1, RepSec92EOtherFlag_1, end_a_gen1_sc, rngname_a_gen1_sc As Variant
'new
Public STCG_A1_Count1, STCG_A1_Count2, STCG_A1_CountA, STCG_A1_Count3, STCG_A1_Count4, STCG_A1_Count5, STCG_A1_Count6 As Variant
Public MsgPartBScheduleCG_STCG As Variant
Public STCG_NameOfBuyer, STCG_PANOfBuyer, STCG_AadhaarOfBuyer, STCG_PercentageShare, STCG_Amount, STCG_AddressOfProperty, STCG_PinCode As Variant
Public end_PrevYear1, end_PrevYear2, end_PrevYear3, end_PrevYear4, end_PrevYear5 As Variant
Public end_PrevYearLTCG1, end_PrevYearLTCG2, end_PrevYearLTCG3, end_PrevYearLTCG4, end_PrevYearLTCG5 As Variant
'by sai on 06/03/2025 for AY 2025-26
Public end_PrevYearLTCG6 As Variant

'end
'new
Public LTCG_B1_Count1, LTCG_B1_Count2, LTCG_B1_Count3, LTCG_B1_Count4, LTCG_B1_Count5, LTCG_B1_Count6, LTCG_B1_CountA As Variant
Public MsgPartBScheduleCG_LTCG As Variant
Public LTCG_NameOfBuyer, LTCG_PANOfBuyer, LTCG_AadhaarOfBuyer, LTCG_PercentageShare, LTCG_Amount, LTCG_AddressOfProperty, LTCG_PinCode As Variant
'end
'block STCG
Public frmsize_STCGrptfrm As Variant
Public end_STCGrptfrm As Variant
Public cntrRng_STCGrptfrm As Variant
Public frmRngname_STCGrptfrm As Variant
Public rngname_STCGrptfrm As Variant
'Public end_STCGrptfrm As Variant
'block LTCG
Public frmsize_LTCGrptfrm As Variant
Public end_LTCGrptfrm As Variant
Public cntrRng_LTCGrptfrm As Variant
Public frmRngname_LTCGrptfrm As Variant
Public rngname_LTCGrptfrm As Variant
'Public end_LTCGrptfrm As Variant
'Validation messageBox on Validate Button
Dim MsgBx_1a As String
Dim MsgBx_1aiii As String
Dim msgbx_5d5e As String
Dim Msgbox_LTCGDTAA As String
Dim Msgbox_STCGDTAA As String
'Temporary VAriable
Public FullValueConsideration_1ai As Variant
Public PropertyStampValue_1aii As Variant
'Public FullValAsPer50C_1aiii As Variant
'CG Setoff Variables
Public Gain_STCG_15p As Double
Public Gain_STCG_30p As Double
Public Gain_STCG_AppRate As Double
Public Gain_LTCG_10p As Double
Public Gain_LTCG_20p As Double
'---------------------------------------------------
Public InStcgAppRate As Double
Public InStcg15Per   As Double
Public InLtcg20Per As Double
Public InStcg20Per As Double
Public InLtcg125Per As Double
Public InLtcg10Per As Double
Public InLtcgDTAA As Double
'Public amtSetOff As Double
Public amtSetOff112 As Double
Public amtSetOff115ea As Double
Public amtSetOffPTI20 As Double
Public amtSetOff112_Proviso As Double
'Public amtSetOff112 As Long
Public InStcg30Per As Double
Public cgosIncome As Double
Public rngname_exemption As Variant
Public rngname_exemption2 As Variant
Public end_exemption As Variant
Public end_exemption2 As Variant
Public STCG_A1e_CapgainonAssets As Variant
Public STCG_A20ia_FullConsideration As Variant
Public STCG_A20ib_NetWorthOfDivision As Variant
Public STCG_A20ic_CapgainonAssets As Variant

Public STCG_A2ia_FullConsideration As Variant

'For 3ia_Q Column, newly added by sai on 24/02/2025
'*****************************************************************
Public STCG_A2ia_2eii_FullConsideration As Variant
Public STCG_A2ibi_2eii_AquisitCost As Variant
Public STCG_A2ibii_2eii_ImproveCost As Variant
Public STCG_A2ibiii_2eii_ExpOnTrans As Variant
Public STCG_A2ibiv_TotalDedn_ii As Variant
Public STCG_A2ic_2eii_BalanceCG As Variant
Public STCG_A2id_2eii_LossSec94of7Or94of8 As Variant
Public STCG_A2ie_2ei_CapgainonAssets As Variant
'3ia_Q End******************************************************************
Public STCG_A2ie_2eii_CapgainonAssets As Variant


Public STCG_A2ibi_AquisitCost As Variant
Public STCG_A2ibii_ImproveCost As Variant
Public STCG_A2ibiii_ExpOnTrans As Variant

Public STCG_A2ibiv_TotalDedn As Variant


Public STCG_A2ic_BalanceCG As Variant
Public STCG_A2id_LossSec94of7Or94of8 As Variant
Public STCG_A2ie_CapgainonAssets As Variant

Public STCG_A2iia_FullConsideration As Variant


'For 3iia_Q Column, newly added by sai on 24/02/2025
'*************************************************************************
Public STCG_A2iia_FullConsideration_ii As Variant
Public STCG_A2iibi_AquisitCost_ii As Variant
Public STCG_A2iibii_ImproveCost_ii As Variant
Public STCG_A2iibiii_ExpOnTrans_ii As Variant
Public STCG_A2iibiv_TotalDedn_ii As Variant
Public STCG_A2iic_BalanceCG_ii As Variant
Public STCG_A2iid_LossSec94of7Or94of8_ii As Variant
Public STCG_A2iie_CapgainonAssets_ii As Variant
Public STCG_A2iie_CapgainonAssets_i As Variant
'3iia_Q  End****************************************************************

Public STCG_A2iie_CapgainonAssets As Variant '(P+Q Sum)

Public STCG_A2iibi_AquisitCost As Variant
Public STCG_A2iibii_ImproveCost As Variant
Public STCG_A2iibiii_ExpOnTrans As Variant

Public STCG_A2iibiv_TotalDedn As Variant

 

Public STCG_A2iic_BalanceCG As Variant
Public STCG_A2iid_LossSec94of7Or94of8 As Variant

Public STCG_A3a_NRItaxSTTPaid As Variant

'For AY 2025-26********************************
Public STCG_A3ai_NRItaxSTTPaid As Variant
Public STCG_A3aii_NRItaxSTTPaid As Variant
'**********************************************


Public STCG_A3b_NRItaxSTTNotPaid As Variant
Public STCG_A4a_FullConsideration As Variant
Public STCG_A4a_FullConsideration1 As Variant
Public STCG_A4a_FullConsideration2 As Variant
Public STCG_A4a_FullConsideration3 As Variant
Public STCG_A4a_FullConsideration4 As Variant

Public STCG_A4bi_AquisitCost As Variant
Public STCG_A4bii_ImproveCost As Variant
Public STCG_A4biii_ExpOnTrans As Variant
Public STCG_A4biv_TotalDedn As Variant
Public STCG_A4c_BalanceCG As Variant
Public STCG_A4d_LossSec94of7Or94of8 As Variant
Public STCG_A4e_CapgainonAssets As Variant
Public STCG_A5a_FullConsideration As Variant
Public STCG_A5a_FullConsideration1 As Variant
Public STCG_A5a_FullConsideration2 As Variant
Public STCG_A5a_FullConsideration3 As Variant
Public STCG_A5a_FullConsideration4 As Variant
Public STCG_A5bi_AquisitCost As Variant
Public STCG_A5bii_ImproveCost As Variant
Public STCG_A5biii_ExpOnTrans As Variant
Public STCG_A5biv_TotalDedn As Variant
Public STCG_A5c_BalanceCG As Variant
Public STCG_A5d_LossSec94of7Or94of8 As Variant
Public STCG_A5ei_ExemptionAmount As Variant
Public STCG_A5eii_ExemptionAmount As Variant
Public STCG_A5eiii_ExemptionAmount As Variant
Public STCG_A5e_ExemptionGrandTotal As Variant

Public STCG_A5e_CapgainonAssets As Variant

Public STCG_A6_StatusUnutilizedCG As Variant
Public STCG_A7_PreviousYer_1 As Variant
Public STCG_A7_DeductionSection_1 As Variant
Public STCG_A7_AsstAcquiredYer_1 As Variant
Public STCG_A7_AmtUtilised_1  As Variant
Public STCG_A7_AmtNotUsed_1 As Variant
Public STCG_A7_PreviousYer_2 As Variant
Public STCG_A7_DeductionSection_2 As Variant
Public STCG_A7_AsstAcquiredYer_2 As Variant
Public STCG_A7_AmtUtilised_2  As Variant
Public STCG_A7_AmtNotUsed_2 As Variant
Public STCG_A7_AmtDeemed As Variant
Public STCG_A6_AmountDeemedOth As Variant

Public STCG_A8_Country As Variant
Public STCG_A8_CountryCode As Variant
Public STCG_A8_CountryName As Variant
Public STCG_A8_DTAA_Article As Variant
Public STCG_A8_CertiStatus As Variant
Public STCG_A8_ItemNo As Variant
Public STCG_A8_StcgAmt As Variant
Public STCG_A8_StcgDTAA As Variant
Public RateTreatySTCG_DTAA As Variant
Public STCG_A8b_StcgDTAA As Variant
Public STCG_A8_Total As Variant

Public STCG_A8_DeemedStcgOnAssets As Variant
Public STCG_A7_TotalSTCG As Variant

Public LTCG_B1e_CapgainonAssets As Variant

'For Slump sale 2aiii_P Column
Public LTCG_B02a_FullConsideration As Variant
Public LTCG_B02b_NetWorthOfDivision As Variant
Public LTCG_B02c_SlumpBalance As Variant
Public LTCG_B02di_ExemptionAmount As Variant
Public LTCG_B02dii_ExemptionAmount As Variant
Public LTCG_B02d_ExemptionGrandTotal As Variant

'For Slump sale
'2aiii_Q Column for AY 2025-26 by sai on 02/03/2025***********************
Public LTCG_B02a_FullConsideration_ii As Variant
Public LTCG_B02b_NetWorthOfDivision_ii As Variant
Public LTCG_B02c_SlumpBalance_ii As Variant
Public LTCG_B02di_ExemptionAmount_ii As Variant
Public LTCG_B02dii_ExemptionAmount_ii As Variant
Public LTCG_B02d_ExemptionGrandTotal_ii As Variant
'end of 2aiii_Q *********************************************************

Public LTCG_B02e_CapgainonAssets As Variant
Public LTCG_B2a_FullConsideration As Variant
Public LTCG_B2bi_AquisitCost As Variant
Public LTCG_B2bii_ImproveCost As Variant
Public LTCG_B2biii_ExpOnTrans As Variant
Public LTCG_B2biv_TotalDedn As Variant
Public LTCG_B2c_BalanceCG As Variant
Public LTCG_B2di_ExemptionAmount As Variant
Public LTCG_B2dii_ExemptionAmount As Variant
Public LTCG_B2d_ExemptionGrandTotal As Variant

Public LTCG_B9_dateof_limitorwithdrawal_1 As Variant ' by konda


Public LTCG_B2e_CapgainonAssets As Variant
Public LTCG_B3a_FullConsideration As Variant
Public LTCG_B3bi_AquisitCost As Variant
Public LTCG_B3bii_ImproveCost As Variant
Public LTCG_B3biii_ExpOnTrans As Variant
Public LTCG_B3biv_TotalDedn As Variant
Public LTCG_B3c_BalanceCG As Variant
Public LTCG_B3di_ExemptionAmount As Variant
Public LTCG_B3dii_ExemptionAmount As Variant
Public LTCG_B3d_ExemptionGrandTotal As Variant
'______________
Public LTCG_B3e_CapgainonAssets As Variant

Public LTCG_Bii3a_FullConsideration As Variant
Public LTCG_Bii3bi_AquisitCost As Variant
Public LTCG_Bii3bii_ImproveCost As Variant
Public LTCG_Bii3biii_ExpOnTrans As Variant
Public LTCG_Bii3biv_TotalDedn As Variant
Public LTCG_Bii3c_BalanceCG As Variant
Public LTCG_Bii3di_ExemptionAmount As Variant
Public LTCG_Bii3dii_ExemptionAmount As Variant
Public LTCG_Bii3d_ExemptionGrandTotal As Variant


'by sai on 27/03/2025 as per DE V0.4**************************
Public LTCG_B3a_ii_FullConsideration As Variant
Public LTCG_B3bia_AcqIndexation_ii As Variant
Public LTCG_B3bia_AcqIndexation As Variant
Public LTCG_B3biia_ImproveCost_WithoutIndexation As Variant
Public LTCG_B3biia_ImproveCost_WithoutIndexation_ii As Variant
Public LTCG_B3biii_ExpOnTrans_ii As Variant
Public LTCG_B3biv_TotalDedn_ii As Variant
Public LTCG_B3c_BalanceCG_ii As Variant
Public LTCG_B04d_ExemptionGrandTotal_ii As Variant
Public LTCG_B3e_CapgainonAssets_ii As Variant
Public LTCG_B3f_Tax20 As Variant
Public LTCG_B3g_Tax10 As Variant
Public LTCG_B3h_Excess As Variant
Public LTCG_B3ea_CapgainonAssets As Variant
'**************************************************************

'Sale of GDR 115ACA 4ii_Q Column by sai for AY 2025-26
'********************************************************
Public LTCG_Bii3a_FullConsideration_ii As Variant
Public LTCG_Bii3bi_AquisitCost_ii As Variant
Public LTCG_Bii3bii_ImproveCost_ii As Variant
Public LTCG_Bii3biii_ExpOnTrans_ii As Variant
Public LTCG_Bii3biv_TotalDedn_ii As Variant
Public LTCG_Bii3c_BalanceCG_ii As Variant
Public LTCG_Bii3di_ExemptionAmount_ii As Variant
Public LTCG_Bii3dii_ExemptionAmount_ii As Variant
'********************************************************


Public LTCG_Bii3e_CapgainonAssets As Variant
Public LTCG_B4a_LTCGWithoutBenefit As Variant
Public LTCG_B4bi_ExemptionAmount As Variant
Public LTCG_B4bii_ExemptionAmount As Variant
Public LTCG_B4b_ExemptionGrandTotal As Variant
Public LTCG_B4c_BalanceCG As Variant

Public LTCG_B4d_LTCGOnListedSecurity As Variant
Public LTCG_B4e_LTCGOnUnlistedSecurity As Variant
Public LTCG_Bi5a_FullConsideration1 As Variant
Public LTCG_Bi5a_FullConsideration2 As Variant
Public LTCG_Bi5a_FullConsideration3 As Variant
Public LTCG_Bi5a_FullConsideration4 As Variant
Public LTCG_Bi5a_FullConsideration As Variant
Public LTCG_Bi5bi_AquisitCost As Variant
Public LTCG_Bi5bii_ImproveCost As Variant
Public LTCG_Bi5biii_ExpOnTrans As Variant
Public LTCG_Bi5biv_TotalDedn As Variant
Public LTCG_Bi5c_BalanceCG As Variant
Public LTCG_Bi5di_ExemptionAmount As Variant
Public LTCG_Bi5dii_ExemptionAmount As Variant
Public LTCG_Bi5d_ExemptionGrandTotal As Variant

Public LTCG_Bi5e_CapgainonAssets As Variant
Public LTCG_Bii5a_FullConsideration1 As Variant
Public LTCG_Bii5a_FullConsideration2 As Variant
Public LTCG_Bii5a_FullConsideration3 As Variant
Public LTCG_Bii5a_FullConsideration4 As Variant
Public LTCG_Bii5a_FullConsideration As Variant
Public LTCG_Bii5bi_AquisitCost As Variant
Public LTCG_Bii5bii_ImproveCost As Variant
Public LTCG_Bii5biii_ExpOnTrans As Variant
Public LTCG_Bii5biv_TotalDedn As Variant
Public LTCG_Bii5c_BalanceCG As Variant
Public LTCG_Bii5di_ExemptionAmount As Variant
Public LTCG_Bii5dii_ExemptionAmount As Variant
Public LTCG_Bii5d_ExemptionGrandTotal As Variant

Public LTCG_Bii5e_CapgainonAssets As Variant
Public LTCG_Biii5a_FullConsideration As Variant
Public LTCG_Biii5a_FullConsideration1 As Variant
Public LTCG_Biii5a_FullConsideration2 As Variant
Public LTCG_Biii5a_FullConsideration3 As Variant
Public LTCG_Biii5a_FullConsideration4 As Variant
Public LTCG_Biii5bi_AquisitCost As Variant
Public LTCG_Biii5bii_ImproveCost As Variant
Public LTCG_Biii5biii_ExpOnTrans As Variant
Public LTCG_Biii5biv_TotalDedn As Variant
Public LTCG_Biii5c_BalanceCG As Variant
Public LTCG_Biii5di_ExemptionAmount As Variant
Public LTCG_Biii5dii_ExemptionAmount As Variant
Public LTCG_Biii5d_ExemptionGrandTotal As Variant

Public LTCG_Biii5e_CapgainonAssets As Variant
Public LTCG_B6a_SaleonSpecAsset As Variant
Public LTCG_B6b_DednSpecAssetus115 As Variant
Public LTCG_B6c_BalonSpeciAsset As Variant
Public LTCG_B6d_SaleOtherSpecAsset As Variant

'Commented by sai for AY 2025-26
'Public LTCG_B6e_DednOtherSpecAssetus115 As Variant
'Public LTCG_B6f_BalOtherthanSpecAsset As Variant

Public LTCG_B7a_FullConsideration As Variant
Public LTCG_B7a_FullConsideration1 As Variant
Public LTCG_B7a_FullConsideration2 As Variant
Public LTCG_B7a_FullConsideration3 As Variant
Public LTCG_B7a_FullConsideration4 As Variant
Public LTCG_B7bi_AquisitCost As Variant
Public LTCG_B7bii_ImproveCost As Variant
Public LTCG_B7biii_ExpOnTrans As Variant
Public LTCG_B7biv_TotalDedn As Variant
Public LTCG_B7c_BalanceCG As Variant
Public LTCG_B7di_ExemptionAmount0 As Variant
Public LTCG_B7di_ExemptionAmount As Variant
Public LTCG_B7dii_ExemptionAmount As Variant
Public LTCG_B7diii_ExemptionAmount As Variant
Public LTCG_B7div_ExemptionAmount As Variant
Public LTCG_B7d_ExemptionGrandTotal As Variant

Public LTCG_B7e_CapgainonAssets As Variant

Public LTCG_B9_StatusUnutilizedCG As Variant
Public LTCG_B9_PreviousYer_1 As Variant
Public LTCG_B9_DeductionSection_1 As Variant
Public LTCG_B9_AsstAcquiredYer_1 As Variant
Public LTCG_B9_AmtUtilised_1  As Variant
Public LTCG_B9_AmtNotUsed_1 As Variant
Public LTCG_B9_PreviousYer_2 As Variant
Public LTCG_B9_DeductionSection_2 As Variant
Public LTCG_B9_AsstAcquiredYer_2 As Variant
Public LTCG_B9_AmtUtilised_2  As Variant
Public LTCG_B9_AmtNotUsed_2 As Variant
Public LTCG_B9_AmtDeemed As Variant
Public LTCG_B9_AmountDeemedOth As Variant

Public LTCG_B10_Country As Variant
Public LTCG_B10_CountryCode As Variant
Public LTCG_B10_CountryName As Variant
Public LTCG_B10_DTAA_Article As Variant
Public LTCG_B10_CertiStatus As Variant
Public LTCG_B10_ItemNo As Variant
Public LTCG_B10_StcgAmt As Variant
Public LTCG_B10_ltcgDTAA As Variant
Public LTCG_B10b_ltcgDTAA As Variant
Public LTCG_Total As Variant

'**************************************************************************
Public STCG_AII_CapitalLoss As Variant
Public LTCG_BII_CapitalLoss As Variant 'For AY 2025-26 by sai on 06/03/2025
'**************************************************************************

Public LTCG_B9_TotalLTCG As Variant

Public CG_C_TotScheduleCGFor23 As Variant
Public CG_C_SumOfCGIncm As Variant
Public CG_C_IncmFromVDATrnsf As Variant

Public IHLA_Ei2_StclSetoff15Per As Variant
Public IHLA_Ei3_StclSetoff30Per As Variant
Public IHLA_Ei4_StclSetoffAppRate As Variant
Public IHLA_Ei5_LtclSetOff10Per As Variant
Public IHLA_Ei6_LtclSetOff20Per As Variant
Public IHLA_Eii1_CurrYearIncome As Variant
Public IHLA_Eii3_StclSetoff30Per As Variant
Public IHLA_Eii4_StclSetoffAppRate As Variant
Public IHLA_Eii7_CurrYrCapGain As Variant
Public IHLA_Eiii1_CurrYearIncome As Variant
Public IHLA_Eiii2_StclSetoff15Per As Variant
Public IHLA_Eiii4_StclSetoffAppRate As Variant
Public IHLA_Eiii7_CurrYrCapGain As Variant
Public IHLA_Eiv1_CurrYearIncome As Variant
Public IHLA_Eiv2_StclSetoff15Per As Variant
Public IHLA_Eiv3_StclSetoff30Per As Variant
Public IHLA_Eiv7_CurrYrCapGain As Variant
Public IHLA_Ev1_CurrYearIncome As Variant
Public IHLA_Ev2_StclSetoff15Per As Variant
Public IHLA_Ev3_StclSetoff30Per As Variant
Public IHLA_Ev4_StclSetoffAppRate As Variant
Public IHLA_Ev6_LtclSetOff20Per As Variant
Public IHLA_Ev7_CurrYrCapGain As Variant
Public IHLA_Evi1_CurrYearIncome As Variant
Public IHLA_Evi2_StclSetoff15Per As Variant
Public IHLA_Evi3_StclSetoff30Per As Variant
Public IHLA_Evi4_StclSetoffAppRate As Variant
Public IHLA_Evi5_LtclSetOff10Per As Variant
Public IHLA_Evi7_CurrYrCapGain As Variant
Public IHLA_Evii2_StclSetoff15Per As Variant
Public IHLA_Evii3_StclSetoff30Per As Variant
Public IHLA_Evii4_StclSetoffAppRate As Variant
Public IHLA_Evii5_LtclSetOff10Per As Variant
Public IHLA_Evii6_LtclSetOff20Per As Variant
Public IHLA_Eviii2_StclSetoff15Per As Variant
Public IHLA_Eviii3_StclSetoff30Per As Variant
Public IHLA_Eviii4_StclSetoffAppRate As Variant
Public IHLA_Eviii5_LtclSetOff10Per As Variant
Public IHLA_Eviii6_LtclSetOff20Per As Variant

Public AccSTCG_Upto15Of6 As Variant
Public AccSTCG_Upto15Of9 As Variant
Public ACCSTCG_Up16Of9To15Of12 As Variant
Public ACCSTCG_Up16Of12To15Of3 As Variant
Public ACCSTCG_Up16Of3To31Of3 As Variant

'For AY 2025-26 Newly added by sai on 27/02/2025
'*************************************************************************
Public ACCLTCGDTAA_Upto15Of6_20 As Variant
Public ACCLTCGDTAA_Upto15Of9_20 As Variant
Public ACCLTCGDTAA_Up16Of9To15Of12_20 As Variant
Public ACCLTCGDTAA_Up16Of12To15Of3_20 As Variant
Public ACCLTCGDTAA_Up16Of3To31Of3_20 As Variant
Public ACCLTCGDTAA_Upto15Of6_125 As Variant
Public ACCLTCGDTAA_Upto15Of9_125 As Variant
Public ACCLTCGDTAA_Up16Of9To15Of12_125 As Variant
Public ACCLTCGDTAA_Up16Of12To15Of3_125 As Variant
Public ACCLTCGDTAA_Up16Of3To31Of3_125 As Variant

'**************************************************************************

Public AccSTCG30_Upto15Of6 As Variant
Public AccSTCG30_Upto15Of9 As Variant
Public AccSTCG30_Up16Of9To15Of12 As Variant
Public AccSTCG30_Up16Of12To15Of3 As Variant
Public AccSTCG30_Up16Of3To31Of3 As Variant
Public ACCSTCGOTH_Upto15Of6 As Variant
Public ACCSTCGOTH_Upto15Of9 As Variant
Public ACCSTCGOTH_Up16Of9To15Of12 As Variant
Public ACCSTCGOTH_Up16Of12To15Of3 As Variant
Public ACCSTCGOTH_Up16Of3To31Of3 As Variant
Public ACCLTCG_Upto15Of6 As Variant
Public ACCLTCG_Upto15Of9 As Variant
Public ACCLTCG_Up16Of9To15Of12 As Variant
Public ACCLTCG_Up16Of12To15Of3 As Variant
Public ACCLTCG_Up16Of3To31Of3 As Variant
Public ACCLTCGNP_Upto15Of6 As Variant
Public ACCLTCGNP_Upto15Of9 As Variant
Public ACCLTCGNP_Up16Of9To15Of12 As Variant
Public ACCLTCGNP_Up16Of12To15Of3 As Variant
Public ACCLTCGNP_Up16Of3To31Of3 As Variant

Public ACCLTCGDTAA_Upto15Of6 As Variant
Public ACCLTCGDTAA_Upto15Of9 As Variant
Public ACCLTCGDTAA_Up16Of9To15Of12 As Variant
Public ACCLTCGDTAA_Up16Of12To15Of3 As Variant
Public ACCLTCGDTAA_Up16Of3To31Of3 As Variant

Public ACCSTCGDTAA_Upto15Of6 As Variant
Public ACCSTCGDTAA_Upto15Of9 As Variant
Public ACCSTCGDTAA_Up16Of9To15Of12 As Variant
Public ACCSTCGDTAA_Up16Of12To15Of3 As Variant
Public ACCSTCGDTAA_Up16Of3To31Of3 As Variant

'Change 102 21.12.2023 V0.6 SIT-65102 A
Public AccVDA_Upto15Of6 As Variant
Public AccVDA_Upto15Of9 As Variant
Public AccVDA_Up16Of9To15Of12 As Variant
Public AccVDA_Up16Of12To15Of3 As Variant
Public AccVDA_Up16Of3To31Of3 As Variant
'End Change SIT-65102 A

Public DEDN_DeductedSecCode As Variant
Public DEDN_AmtDeducted As Variant
Public DEDN_CostofAsset As Variant
Public DEDN_DateofAcquist As Variant
Public DEDN_AmtDeposited As Variant

Public DEDN_1c_TotDeductClaim As Variant
Public DEDN_D2_DeductPAN As Variant

Public rngname_LTCGDTAA As Variant
Public end_LTCGDTAA, end_LTCGDTAA2, end_LTCGDTAA3, end_LTCGDTAA4, end_LTCGDTAA5, end_LTCGDTAA6, end_LTCGDTAA7, end_LTCGDTAA8  As Long
Public CountryNameLTCG_DTAA As Variant
Public CountryCodeLTCG_DTAA, RateTreatyLTCG_DTAA As Variant
Public ArticleLTCG_DTAA, RateITActLTCG_DTAA, Applrate_LTCG  As Variant
Public TRCLTCG_DTAA, SectionITActLTCG_DTAA As Variant
Public SectionLTCG_DTAA As Variant
Public SectionCodeLTCG_DTAA As Variant
Public AmountLTCG_DTAA As Variant
Public TotAmountLTCG_DTAA As Variant

Public osdividend_22f As Variant 'by Sai on 03/03/2025

'STCG DTAA
'Dim Msgbox_STCG_DTAA, Msgbox_STCGDTAA As String

'*********************************************************************
'Newly added by sai on 25/02/2025
Dim Msgbox_STCG_AII As String
Public end_STCGAII_Captlos20r30rapplicblrate   As Variant
Public end_STCGAII_Captlos20r30rapplicblrate_amt   As Variant

Dim Msgbox_LTCG_BII As String
Public end_LTCGBII_Captlos20r30rapplicblrate   As Variant
Public end_LTCGBII_Captlos20r30rapplicblrate_amt   As Variant

'*********************************************************************


Public rngname_STCGDTAA As Variant
Public end_STCGDTAA, end_STCGDTAA2, end_STCGDTAA3, end_STCGDTAA4, end_STCGDTAA5, end_STCGDTAA6, end_STCGDTAA7, end_STCGDTAA8  As Long
Public CountryNameSTCG_DTAA, RateITActSTCG_DTAA, ApplRate_STCG As Variant
Public CountryCodeSTCG_DTAA As Variant
Public ArticleSTCG_DTAA As Variant
Public TRCSTCG_DTAA As Variant
Public SectionSTCG_DTAA, SectionITActSTCG_DTAA As Variant
Public SectionCodeSTCG_DTAA As Variant
Public AmountSTCG_DTAA As Variant
Public TotAmountSTCG_DTAA As Variant

Sub Cmd_ValidateSheet_CG_Click()
ValidateSheetCG_All
fmsgboxStatus ("Sheet CG is ok")
End Sub

Sub Cmd_CG_Setoff_Click()
On Error GoTo endline
Application.EnableEvents = False
Application.ScreenUpdating = False
CG_Calc.doSetoff
CYLACalculations.CYLABFLASetOff
Sheet13.Activate
endline:
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
'---------------------------------
'Validation of CG Sheet
Sub ValidateSheetCG_All()
    subProcCaption = "Validating CG"
   
   STCGBlock_Validation
   LTCGBlock_Validation
   
   
    If Not ValidatesheetSTCGDTAASection Then
        Sheet13.Activate
        fmsgboxsmall Msgbox_STCGDTAA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    If Not ValidatesheetLTCGDTAASection Then
        Sheet13.Activate
        fmsgboxsmall Msgbox_LTCGDTAA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    '***************************************
    'For AY 2025-26 by sai on 27/02/2025
    If Not ValidatesheetSTCG_AII Then
        Sheet13.Activate
        fmsgboxsmall Msgbox_STCG_AII
        CloseMsg
    End If
    
    If Not ValidatesheetLTCG_BII Then
        Sheet13.Activate
        fmsgboxsmall Msgbox_LTCG_BII
        CloseMsg
    End If
    '**************************************
    
    
    If Not Validate_ALL_CG() Then
        Sheet13.Activate
        fmsgboxsmall (msgboxSHEETCG) ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    CG_BiiiGreater
    
    If Not Validate_SUMMARY_CG Then
        Sheet13.Activate
        fmsgboxsmall (msgboxSHEETCG) ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
   
    If Not validateDeduction Then
        Sheet13.Activate
        fmsgboxsmall (msgboxSHEETCG) ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
End Sub
Function validateDeduction() As Boolean
validateDeduction = True
If Not Validate_ALL_CG() Then validateDeduction = False
If Not Validate_SUMMARY_CG() Then validateDeduction = False
'dpk
If (Len(Sheet13.Range("Deduction.54.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54_CG Then validateDeduction = False
End If

If (Len(Sheet13.Range("Deduction.54B.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54B_CG Then validateDeduction = False
End If

If (Len(Sheet13.Range("Deduction.54D.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54D_CG Then validateDeduction = False
End If


If (Len(Sheet13.Range("Deduction.54EC.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54EC_CG Then validateDeduction = False
End If

If (Len(Sheet13.Range("Deduction.54F.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54F_CG Then validateDeduction = False
End If

If (Len(Sheet13.Range("Deduction.54G.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54G_CG Then validateDeduction = False
End If

If (Len(Sheet13.Range("Deduction.54GA.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable54GA_CG Then validateDeduction = False
End If

If (Len(Sheet13.Range("Deduction.115F.DateOfTransfer").item(1).value) > 0) Then
     If Not ValidateDednTable115F_CG Then validateDeduction = False
End If

'If (Len(Sheet13.Range("DEDN.D2_DeductPAN").value) > 0) Then
'If Not CheckPAN((Sheet13.Range("DEDN.D2_DeductPAN").value)) Then
'                'MsgPartBScheduleS(i) = MsgPartBScheduleS(i) + "* Pan number of employer is not valid." & Chr(13)
'                 msgboxSHEETCG = msgboxSHEETCG + "* Enter Valid PAN." & Chr(13)
'                validateDeduction = False
'        End If
'        End If

End Function

Function chk5d5e_LTCG() As Boolean
chk5d5e_LTCG = True
msgbx_5d5e = ""

End Function

Function populateCGTab()

Dim neg15 As Double
Dim neg10 As Double
Dim neg20 As Double
Dim SetOff As Double
Dim amtSetOff As Double

' stcg = cgosIncome.cgInc.stcg
' prctg15 = cgosIncome.cgInc.stcg.prctg15
 prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
 neg15 = 0
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg15_sec111a) < 0) Then
        neg15 = neg15 + cgosIncome_cgInc_stcg_prctg15_sec111a
       cgosIncome_cgInc_stcg_prctg15_sec111a = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) < 0) Then
        neg15 = neg15 + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
       cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = 0
    End If
    
    If (neg15 < 0) Then
        SetOff = setOffAgainst111a(neg15, cgosIncome_cgInc_stcg_prctg15)
        neg15 = neg15 + SetOff
        SetOff = setOffAgainst115ad_1_b_ii(neg15, cgosIncome_cgInc_stcg_prctg15)
        neg15 = neg15 + SetOff
        cgosIncome_cgInc_stcg_prctg15_sec111a = neg15 + cgosIncome_cgInc_stcg_prctg15_sec111a
    End If
 
' prctg10 = cgosIncome.cgInc.ltcg.prctg10
prctg10Sum = cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3

                     
 neg10 = 0
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_secProviso) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_secProviso
       cgosIncome_cgInc_ltcg_prctg10_secProviso = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115AC_1) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
       cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115AD_3) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
       cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = 0
    End If
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115E_b) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b
        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = 0
    End If

    If (neg10 < 0) Then
      SetOff = setOffAgainst112_1_c_2(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainst115AC_1(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainst115ACA_1(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainst115AD_3(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainst115E_b(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainstsecProviso(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        
        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
    End If
        
    
'prctg20 = cgosIncome.cgInc.ltcg.prctg20
prctg20Sum = cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA
    
   neg20 = 0
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_sec112) < 0) Then
        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec112
        cgosIncome_cgInc_ltcg_prctg20_sec112 = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_sec11EA) < 0) Then
        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec11EA
        cgosIncome_cgInc_ltcg_prctg20_sec11EA = 0
    End If
    
    If (neg20 < 0) Then
         SetOff = setOffAgainst112(neg20, cgosIncome_cgInc_ltcg_prctg20)
        neg20 = neg20 + SetOff
        SetOff = setOffAgainst115Ea(neg20, cgosIncome_cgInc_ltcg_prctg20)
        neg20 = neg20 + SetOff
        cgosIncome_cgInc_ltcg_prctg20_sec112 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec112
    End If
    
    
End Function



Function setOffPctg30Loss(cgosIncome_cgInc_stcg_prctg30Loss As Double)
Dim amtSetOff As Double

InStcgAppRate = Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").value
'Added by Riyaz on 07/04/2025
InStcg20Per = Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").value
InLtcg20Per = Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value
InStcg15Per = Sheet13.Range("IHLA.Eii3_StclSetoff30Per").value
'Added by Riyaz on 07/04/2025
InLtcg125Per = Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").value
InLtcg10Per = Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value

       
   cgosIncome_cgInc_stcg_prctg30Loss = -cgosIncome_cgInc_stcg_prctg30Loss
    amtSetOff = 0
    
'STCG AR
    amtSetOff = setOffAgainstAr(cgosIncome_cgInc_stcg_prctg30Loss, cgosIncome_cgInc_stcg_prctgAppRate)
    cgosIncome_cgInc_stcg_prctg30Loss = amtSetOff + cgosIncome_cgInc_stcg_prctg30Loss
    InStcgAppRate = Abs(amtSetOff)
    
'Added by Riyaz on 07/04/2025
'STCG 20%
    amtSetOff = setOffAgainststcg20(cgosIncome_cgInc_stcg_prctg30Loss, cgosIncome_cgInc_stcg_prctg20)
    cgosIncome_cgInc_stcg_prctg30Loss = amtSetOff + cgosIncome_cgInc_stcg_prctg30Loss
    InStcg20Per = Abs(amtSetOff)

'LTCG 20%
   amtSetOff = setOffAgainst20(cgosIncome_cgInc_stcg_prctg30Loss, cgosIncome_cgInc_ltcg_prctg20)
   cgosIncome_cgInc_stcg_prctg30Loss = amtSetOff + cgosIncome_cgInc_stcg_prctg30Loss
   InLtcg20Per = Abs(amtSetOff)

'STCG 15%
   amtSetOff = setOffAgainst15(cgosIncome_cgInc_stcg_prctg30Loss, cgosIncome_cgInc_stcg_prctg15)
   cgosIncome_cgInc_stcg_prctg30Loss = amtSetOff + cgosIncome_cgInc_stcg_prctg30Loss
   InStcg15Per = Abs(amtSetOff)

'LTCG 12.5%
   amtSetOff = setOffAgainst125(cgosIncome_cgInc_stcg_prctg30Loss, cgosIncome_cgInc_ltcg_125prctg125)
   cgosIncome_cgInc_stcg_prctg30Loss = amtSetOff + cgosIncome_cgInc_stcg_prctg30Loss
   InLtcg125Per = Abs(amtSetOff)
   
'LTCG 10%
   amtSetOff = setOffAgainst10(cgosIncome_cgInc_stcg_prctg30Loss, cgosIncome_cgInc_ltcg_prctg10)
   cgosIncome_cgInc_stcg_prctg30Loss = amtSetOff + cgosIncome_cgInc_stcg_prctg30Loss
   InLtcg10Per = Abs(amtSetOff)
   

End Function

'new STCG DTAA loss setoff
Function setOffstcgDTAA(cgosIncome_cgInc_stcg_DTAALoss As Double)
Dim amtSetOff As Double

InStcg30Per = Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").value
InStcgAppRate = Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").value
'Added by Riyaz on 07/04/2025
InStcg20Per = Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").value
InLtcg20Per = Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").value
InStcg15Per = Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").value
'Added by Riyaz on 07/04/2025
InLtcg125Per = Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").value
InLtcg10Per = Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").value
InLtcgDTAA = Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value

       
   cgosIncome_cgInc_stcg_DTAALoss = -cgosIncome_cgInc_stcg_DTAALoss
    amtSetOff = 0
'STCG 30%
    amtSetOff = setOffAgainst30(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_stcg_prctg30)
    cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
    InStcg30Per = Abs(amtSetOff)

'STCG AR
    amtSetOff = setOffAgainstAr(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_stcg_prctgAppRate)
    cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
    InStcgAppRate = Abs(amtSetOff)

'Added by Riyaz on 07/04/2025
'STCG 20%
    amtSetOff = setOffAgainststcg20(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_stcg_prctg20)
    cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
    InStcg20Per = Abs(amtSetOff)
 
'LTCG 20%
   amtSetOff = setOffAgainst20(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_ltcg_prctg20)
   cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
   InLtcg20Per = Abs(amtSetOff)
   
'STCG 15%
   amtSetOff = setOffAgainst15(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_stcg_prctg15)
   cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
   InStcg15Per = Abs(amtSetOff)

'Added by Riyaz on 07/04/2025
'LTCG 12.5%
   amtSetOff = setOffAgainst125(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_ltcg_125prctg125)
   cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
   InLtcg125Per = Abs(amtSetOff)

'LTCG 10%
   amtSetOff = setOffAgainst10(cgosIncome_cgInc_stcg_DTAALoss, cgosIncome_cgInc_ltcg_prctg10)
   cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
   InLtcg10Per = Abs(amtSetOff)

'LTCG DTAA
   amtSetOff = Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value
   cgosIncome_cgInc_stcg_DTAALoss = amtSetOff + cgosIncome_cgInc_stcg_DTAALoss
   InLtcgDTAA = Abs(amtSetOff)
   
   
End Function

Function setOffltcgDTAA(cgosIncome_cgInc_ltcg_DTAALoss As Double)
Dim amtSetOff As Double
InLtcg20Per = Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").value
InLtcg10Per = Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").value

'LTCG 12.5% by sai on 07/04/2025
InLtcg125Per = Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").value

       
   cgosIncome_cgInc_ltcg_DTAALoss = -cgosIncome_cgInc_ltcg_DTAALoss
    amtSetOff = 0
 'LTCG 20%
   amtSetOff = setOffAgainst20(cgosIncome_cgInc_ltcg_DTAALoss, cgosIncome_cgInc_ltcg_prctg20)
   cgosIncome_cgInc_ltcg_DTAALoss = amtSetOff + cgosIncome_cgInc_ltcg_DTAALoss
   InLtcg20Per = Abs(amtSetOff)
   'Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value = InLtcg20Per
    
'LTCG 12.5% by Sai on 07/04/2025
'************************************************************************
   amtSetOff = setOffAgainst125(cgosIncome_cgInc_ltcg_DTAALoss, cgosIncome_cgInc_ltcg_125prctg125)
   cgosIncome_cgInc_ltcg_DTAALoss = amtSetOff + cgosIncome_cgInc_ltcg_DTAALoss
   InLtcg125Per = Abs(amtSetOff)
'******************************************************************************************************
  
 'LTCG 10%
    amtSetOff = setOffAgainst10(cgosIncome_cgInc_ltcg_DTAALoss, cgosIncome_cgInc_ltcg_prctg10)
   cgosIncome_cgInc_ltcg_DTAALoss = amtSetOff + cgosIncome_cgInc_ltcg_DTAALoss
   InLtcg10Per = Abs(amtSetOff)
   
   ' Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value = InLtcg10Per

  
   
End Function
'end

Function setOffPctgArLoss(cgosIncome_cgInc_stcg_prctgArLoss As Double)
Dim amtSetOff As Double
InStcg30Per = Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").value
'Added by Riyaz on 07/04/2025
InStcg20Per = Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").value
InLtcg20Per = Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").value
InStcg15Per = Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").value
'Added by Riyaz on 07/04/2025
InLtcg125Per = Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").value
InLtcg10Per = Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").value
    
amtSetOff = 0
cgosIncome_cgInc_stcg_prctgArLoss = -cgosIncome_cgInc_stcg_prctgArLoss

'STCG 30%
amtSetOff = setOffAgainst30(cgosIncome_cgInc_stcg_prctgArLoss, cgosIncome_cgInc_stcg_prctg30)
cgosIncome_cgInc_stcg_prctgArLoss = amtSetOff + cgosIncome_cgInc_stcg_prctgArLoss
InStcg30Per = Abs(amtSetOff)

'STCG 20%
    amtSetOff = setOffAgainststcg20(cgosIncome_cgInc_stcg_prctgArLoss, cgosIncome_cgInc_stcg_prctg20)
    cgosIncome_cgInc_stcg_prctgArLoss = amtSetOff + cgosIncome_cgInc_stcg_prctgArLoss
    InStcg20Per = Abs(amtSetOff)
    
'LTCG 20%
amtSetOff = setOffAgainst20(cgosIncome_cgInc_stcg_prctgArLoss, cgosIncome_cgInc_ltcg_prctg20)
cgosIncome_cgInc_stcg_prctgArLoss = amtSetOff + cgosIncome_cgInc_stcg_prctgArLoss
InLtcg20Per = Abs(amtSetOff)
    
'STCG 15%
amtSetOff = setOffAgainst15(cgosIncome_cgInc_stcg_prctgArLoss, cgosIncome_cgInc_stcg_prctg15)
cgosIncome_cgInc_stcg_prctgArLoss = amtSetOff + cgosIncome_cgInc_stcg_prctgArLoss
InStcg15Per = Abs(amtSetOff)
    
'LTCG 12.5%
   amtSetOff = setOffAgainst125(cgosIncome_cgInc_stcg_prctgArLoss, cgosIncome_cgInc_ltcg_125prctg125)
   cgosIncome_cgInc_stcg_prctgArLoss = amtSetOff + cgosIncome_cgInc_stcg_prctgArLoss ' Riyaz code modified by sai
   InLtcg125Per = Abs(amtSetOff)
    
'LTCG 10%
amtSetOff = setOffAgainst10(cgosIncome_cgInc_stcg_prctgArLoss, cgosIncome_cgInc_ltcg_prctg10)
cgosIncome_cgInc_stcg_prctgArLoss = amtSetOff + cgosIncome_cgInc_stcg_prctgArLoss
InLtcg10Per = Abs(amtSetOff)

End Function
'Addedby Riyaz on 07/04/2025
Function setOffPcstcg20Loss(cgosIncome_cgInc_stcg_Pcstcg20Loss As Double)
Dim amtSetOff As Double

InStcg30Per = Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").value
InStcgAppRate = Sheet13.Range("IHLA.Ev2_StclSetoff20Per").value
InLtcg20Per = Sheet13.Range("IHLA.Eix2_StclSetoff20Per").value
InStcg15Per = Sheet13.Range("IHLA.Eii2_StclSetoff20Per").value
'Added by Riyaz on 07/04/2025
InLtcg125Per = Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").value
InLtcg10Per = Sheet13.Range("IHLA.Evii2_StclSetoff20Per").value
    
amtSetOff = 0
cgosIncome_cgInc_stcg_Pcstcg20Loss = -cgosIncome_cgInc_stcg_Pcstcg20Loss

'STCG 30%
amtSetOff = setOffAgainst30(cgosIncome_cgInc_stcg_Pcstcg20Loss, cgosIncome_cgInc_stcg_prctg30)
cgosIncome_cgInc_stcg_Pcstcg20Loss = amtSetOff + cgosIncome_cgInc_stcg_Pcstcg20Loss
InStcg30Per = Abs(amtSetOff)

'STCG AR%
    amtSetOff = setOffAgainstAr(cgosIncome_cgInc_stcg_Pcstcg20Loss, cgosIncome_cgInc_stcg_prctgAppRate)
    cgosIncome_cgInc_stcg_Pcstcg20Loss = amtSetOff + cgosIncome_cgInc_stcg_Pcstcg20Loss
    InStcgAppRate = Abs(amtSetOff)
    
'LTCG 20%
amtSetOff = setOffAgainst20(cgosIncome_cgInc_stcg_Pcstcg20Loss, cgosIncome_cgInc_ltcg_prctg20)
cgosIncome_cgInc_stcg_Pcstcg20Loss = amtSetOff + cgosIncome_cgInc_stcg_Pcstcg20Loss
InLtcg20Per = Abs(amtSetOff)
    
'STCG 15%
amtSetOff = setOffAgainst15(cgosIncome_cgInc_stcg_Pcstcg20Loss, cgosIncome_cgInc_stcg_prctg15)
cgosIncome_cgInc_stcg_Pcstcg20Loss = amtSetOff + cgosIncome_cgInc_stcg_Pcstcg20Loss
InStcg15Per = Abs(amtSetOff)
    
'LTCG 12.5%
   amtSetOff = setOffAgainst125(cgosIncome_cgInc_stcg_Pcstcg20Loss, cgosIncome_cgInc_ltcg_125prctg125)
   cgosIncome_cgInc_stcg_Pcstcg20Loss = amtSetOff + cgosIncome_cgInc_stcg_Pcstcg20Loss
   InLtcg125Per = Abs(amtSetOff)
    
'LTCG 10%
amtSetOff = setOffAgainst10(cgosIncome_cgInc_stcg_Pcstcg20Loss, cgosIncome_cgInc_ltcg_prctg10)
cgosIncome_cgInc_stcg_Pcstcg20Loss = amtSetOff + cgosIncome_cgInc_stcg_Pcstcg20Loss
InLtcg10Per = Abs(amtSetOff)

End Function

Function setOffPctg20Loss(cgosIncome_cgInc_stcg_prct20Loss As Double)
Dim amtSetOff As Double
InLtcg10Per = Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value


'by sai on 07/04/2025***********************************************
InLtcg125Per = Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").value
'**************************************************************************
amtSetOff = 0

'by sai on 07/04/2025***********************************************************************

'cgosIncome_cgInc_ltcg_prctg20_secProviso ' by sai as per DE V4
 prctg20Sum = cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso
    
    amtSetOff = setOffAgainst125(prctg20Sum, cgosIncome_cgInc_stcg_prct20Loss) '12.5

    prctg20Sum = amtSetOff + prctg20Sum 'by sai on 08/04/2025
    
    amtSetOff = setOffAgainst10(prctg20Sum, cgosIncome_cgInc_ltcg_prctg10) '10
    

    
    InLtcg10Per = Abs(amtSetOff)
    'Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value = InLtcg10Per
    
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg20_sec11EA) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg20_sec11EA = cgosIncome_cgInc_ltcg_prctg20_sec11EA + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And (cgosIncome_cgInc_ltcg_prctg20_sec11EA < 0)) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg20_sec11EA
        cgosIncome_cgInc_ltcg_prctg20_sec11EA = 0
    End If


    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI20) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI20 = cgosIncome_cgInc_ltcg_PTI20 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And (cgosIncome_cgInc_ltcg_PTI20 < 0)) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI20
        cgosIncome_cgInc_ltcg_PTI20 = 0
    End If
    
'by sai on 07/04/2025***************************************************************************************************************
     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg20_secProviso) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg20_secProviso = cgosIncome_cgInc_ltcg_prctg20_secProviso + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And (cgosIncome_cgInc_ltcg_prctg20_secProviso < 0)) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg20_secProviso
        cgosIncome_cgInc_ltcg_prctg20_secProviso = 0
    End If
'****************************************************************************************************************



    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg20_sec112) > (amtSetOff)) Then
        cgosIncome_cgInc_ltcg_prctg20_sec112 = cgosIncome_cgInc_ltcg_prctg20_sec112 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg20_sec112 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg20_sec112
        cgosIncome_cgInc_ltcg_prctg20_sec112 = 0
    End If
    
   

End Function

Function setOffPctg15Loss(cgosIncome_cgInc_stcg_prctg15Loss As Double)
Dim amtSetOff As Double
Dim totalSetOff As Double
 
 InStcg30Per = Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").value
 InStcgAppRate = Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").value
'Added by Riyaz on 07/04/2025
 InStcg20Per = Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").value
 InLtcg20Per = Sheet13.Range("IHLA.Evi2_StclSetoff15Per").value
'Added by Riyaz on 07/04/2025
 InLtcg125Per = Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").value
 InLtcg10Per = Sheet13.Range("IHLA.Ev2_StclSetoff15Per").value
    
 amtSetOff = 0
 totalSetOff = 0

 prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI15
'STCG 30%
    amtSetOff = setOffAgainst30(prctg15Sum, cgosIncome_cgInc_stcg_prctg30)
    prctg15Sum = amtSetOff + prctg15Sum
    totalSetOff = amtSetOff
    InStcg30Per = Abs(amtSetOff)

'STCG AR
    amtSetOff = setOffAgainstAr(prctg15Sum, cgosIncome_cgInc_stcg_prctgAppRate)
    prctg15Sum = amtSetOff + prctg15Sum
    totalSetOff = totalSetOff + amtSetOff
    InStcgAppRate = Abs(amtSetOff)

'Added by Riyaz on 07/04/2025
'STCG 20%
    amtSetOff = setOffAgainststcg20(prctg15Sum, cgosIncome_cgInc_stcg_prctg20)
    prctg15Sum = amtSetOff + prctg15Sum
    totalSetOff = totalSetOff + amtSetOff
    InStcg20Per = Abs(amtSetOff)

'LTCG 20%
    amtSetOff = setOffAgainst20(prctg15Sum, cgosIncome_cgInc_ltcg_prctg20)
    prctg15Sum = amtSetOff + prctg15Sum
    totalSetOff = totalSetOff + amtSetOff
    InLtcg20Per = Abs(amtSetOff)

'Added by Riyaz on 07/04/2025
'LTCG 12.5%
   amtSetOff = setOffAgainst125(prctg15Sum, cgosIncome_cgInc_ltcg_125prctg125)
   prctg15Sum = amtSetOff + prctg15Sum
   totalSetOff = totalSetOff + amtSetOff
   InLtcg125Per = Abs(amtSetOff)
    
'LTCG 10%
    amtSetOff = setOffAgainst10(prctg15Sum, cgosIncome_cgInc_ltcg_prctg10)
    prctg15Sum = amtSetOff + prctg15Sum
    totalSetOff = totalSetOff + amtSetOff
    InLtcg10Per = Abs(amtSetOff)
    
    
    If (totalSetOff >= 0 And Abs(cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) > totalSetOff) Then
       cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + totalSetOff
        amtSetOff = 0
 ElseIf (totalSetOff >= 0 And cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii < 0) Then
        amtSetOff = totalSetOff + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
       cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = 0
    End If
    
    If (totalSetOff >= 0 And Abs(cgosIncome_cgInc_stcg_PTI15) > totalSetOff) Then
       cgosIncome_cgInc_stcg_PTI15 = cgosIncome_cgInc_stcg_PTI15 + totalSetOff
        amtSetOff = 0
 ElseIf (totalSetOff >= 0 And cgosIncome_cgInc_stcg_PTI15 < 0) Then
        amtSetOff = totalSetOff + cgosIncome_cgInc_stcg_PTI15
       cgosIncome_cgInc_stcg_PTI15 = 0
    End If
    
    
    If (totalSetOff >= 0 And Abs(cgosIncome_cgInc_stcg_prctg15_sec111a) > totalSetOff) Then
        cgosIncome_cgInc_stcg_prctg15_sec111a = cgosIncome_cgInc_stcg_prctg15_sec111a + totalSetOff
        amtSetOff = 0
    ElseIf (totalSetOff >= 0 And cgosIncome_cgInc_stcg_prctg15_sec111a < 0) Then
        amtSetOff = totalSetOff + cgosIncome_cgInc_stcg_prctg15_sec111a
       cgosIncome_cgInc_stcg_prctg15_sec111a = 0
   End If
 End Function
   
  
   
Function setOffPctg10Loss(cgosIncome_cgInc_stcg_prctgAr10Loss As Double)
Dim amtSetOff As Double
InLtcg20Per = Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value

'LTCG-12.5% by sai on 07/04/2025***************************************************
InLtcg125Per = Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").value
'*************************************************************************
 amtSetOff = 0
 Dim prctg10Sum As Double
'prctg10 = cgosIncome.cgInc.ltcg.prctg10
 
 'prctg10Sum = cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112
 ' cgosIncome_cgInc_ltcg_prctg10_secProviso Not using  commented  by Sai as this sec changed from 10 to 20
 
 prctg10Sum = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112

    amtSetOff = setOffAgainst20(prctg10Sum, cgosIncome_cgInc_ltcg_prctg20) '20
    prctg10Sum = amtSetOff + prctg10Sum 'by sai on 08/04/2025
'    amtSetOff = setOffAgainst125(prctg10Sum, cgosIncome_cgInc_stcg_prctgAr10Loss) '12.5
     amtSetOff = setOffAgainst125(prctg10Sum, cgosIncome_cgInc_ltcg_prctg125) '12.5 'Modified by sai on 08/04/2025
    
    InLtcg20Per = Abs(amtSetOff)
    'Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value = InLtcg20Per
    

    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2) > amtSetOff) Then
       cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = 0
    End If

    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec115AC_1) > amtSetOff) Then
         cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
         cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = 0
    End If

    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
       cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = 0
    End If
    
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec115AD_3) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
       cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = 0
    End If
    
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
       cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = 0
    End If
    
    

    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec115E_b) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = cgosIncome_cgInc_ltcg_prctg10_sec115E_b + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec115E_b < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec115E_b
        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = 0
    End If
    
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI10) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI10 = cgosIncome_cgInc_ltcg_PTI10 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI10 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI10
        cgosIncome_cgInc_ltcg_PTI10 = 0
    End If
    
    ' cgosIncome_cgInc_ltcg_prctg10_secProviso Not using  commented  by Sai as this sec changed from 10 to 20
'    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_secProviso) > amtSetOff) Then
'        cgosIncome_cgInc_ltcg_prctg10_secProviso = cgosIncome_cgInc_ltcg_prctg10_secProviso + amtSetOff
'        amtSetOff = 0
'    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_secProviso < 0) Then
'        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_secProviso
'        cgosIncome_cgInc_ltcg_prctg10_secProviso = 0
'    End If
    
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI10_112) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI10_112 = cgosIncome_cgInc_ltcg_PTI10_112 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI10_112 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI10_112
        cgosIncome_cgInc_ltcg_PTI10_112 = 0
    End If
    
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg10_sec112A) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg10_sec112A = cgosIncome_cgInc_ltcg_prctg10_sec112A + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg10_sec112A < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg10_sec112A
        cgosIncome_cgInc_ltcg_prctg10_sec112A = 0
    End If

End Function

    
Function setOffAgainstAr(value As Double, cgosIncome_cgInc_stcg_prctgAr As Double)
     Dim amtSetOff As Double
     setOffAgainstAr = 0
    If ((value) < 0 And (cgosIncome_cgInc_stcg_Applicable) > 0) Then
        If ((cgosIncome_cgInc_stcg_Applicable) >= Abs(value)) Then
            cgosIncome_cgInc_stcg_Applicable = (cgosIncome_cgInc_stcg_Applicable + value)
            setOffAgainstAr = Abs(value)
            Else
            setOffAgainstAr = Abs(cgosIncome_cgInc_stcg_Applicable)
            cgosIncome_cgInc_stcg_Applicable = 0
        End If
    End If
    
    If ((value + setOffAgainstAr) < 0 And (cgosIncome_cgInc_stcg_ApplicablePTI) > 0) Then
        If ((cgosIncome_cgInc_stcg_ApplicablePTI) >= Abs(value + setOffAgainstAr)) Then
            cgosIncome_cgInc_stcg_ApplicablePTI = (cgosIncome_cgInc_stcg_ApplicablePTI) + (value + setOffAgainstAr) 'Updated by Bindu
            setOffAgainstAr = setOffAgainstAr + Abs(value)
            Else
            setOffAgainstAr = setOffAgainstAr + Abs(cgosIncome_cgInc_stcg_ApplicablePTI)
            cgosIncome_cgInc_stcg_ApplicablePTI = 0
        End If
    End If
End Function



Function setOffAgainst30(value As Double, cgosIncome_cgInc_stcg_prctg30 As Double)
Dim amtSetOff, amtSetOff1 As Double
 amtSetOff = 0
 amtSetOff1 = 0
setOffAgainst30 = 0
    If (value < 0) And (cgosIncome_cgInc_stcg_prctg30_sec115AD > 0) Then
        If cgosIncome_cgInc_stcg_prctg30_sec115AD >= Abs(value) Then
           cgosIncome_cgInc_stcg_prctg30_sec115AD = (cgosIncome_cgInc_stcg_prctg30_sec115AD) + (value)
            amtSetOff = Abs(value)
        Else
        
            amtSetOff = Abs(cgosIncome_cgInc_stcg_prctg30_sec115AD)
            cgosIncome_cgInc_stcg_prctg30_sec115AD = 0
        End If
    End If
    
    If ((value + amtSetOff) < 0) And (cgosIncome_cgInc_stcg_PTI30 > 0) Then
        If cgosIncome_cgInc_stcg_PTI30 >= Abs(value + amtSetOff) Then
           cgosIncome_cgInc_stcg_PTI30 = (cgosIncome_cgInc_stcg_PTI30) + (value + amtSetOff) 'Updated by Bindu
            amtSetOff1 = Abs(value)
        Else
        
            amtSetOff1 = Abs(cgosIncome_cgInc_stcg_PTI30)
            cgosIncome_cgInc_stcg_PTI30 = 0
        End If
    End If
       
    setOffAgainst30 = amtSetOff + amtSetOff1
End Function
    
    
Function setOffAgainst20(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double)
Dim amtSetOff As Double
 amtSetOff112 = 0
 amtSetOff115ea = 0
 amtSetOffPTI20 = 0
 amtSetOff112_Proviso = 0
 
    setOffAgainst20 = 0
    If (value < 0) Then
        amtSetOff115ea = setOffAgainst115Ea(value, cgosIncome_cgInc_stcg_prctg20)
        amtSetOffPTI20 = setOffAgainstPTI20((value + amtSetOff115ea), cgosIncome_cgInc_stcg_prctg20)
'Newly Added by Riyaz on 30/04/2025
        amtSetOff112_Proviso = setOffAgainstsecProviso_20(value + amtSetOff115ea + amtSetOffPTI20, cgosIncome_cgInc_stcg_prctg20)

        amtSetOff112 = setOffAgainst112(value + amtSetOff115ea + amtSetOffPTI20 + amtSetOff112_Proviso, cgosIncome_cgInc_stcg_prctg20)

    End If
            
    setOffAgainst20 = amtSetOff112 + amtSetOff112_Proviso + amtSetOffPTI20 + amtSetOff115ea
End Function
'Added by Riyaz on 07/04/2025 for STCG20
Function setOffAgainststcg20(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double) 'Updated by BIndu 17th March 25
     
     amtSetOff111_20 = 0
     amtSetOff115ad_1_b_ii_20 = 0
     amtSetOffPTI20_STCG = 0
     
     setOffAgainststcg20 = 0
    If (value < 0) Then
    
        amtSetOff115ad_1_b_ii_20 = setOffAgainst115ad_1_b_ii_20(value, cgosIncome_cgInc_stcg_prctg20)
        amtSetOffPTI20_STCG = setOffAgainst15PTI_20((value + amtSetOff115ad_1_b_ii_20), cgosIncome_cgInc_stcg_prctg20)
        amtSetOff111_20 = setOffAgainst111a_20(value + amtSetOff115ad_1_b_ii_20 + amtSetOffPTI20_STCG, cgosIncome_cgInc_stcg_prctg20)
        
    End If
    setOffAgainststcg20 = (amtSetOff111_20) + (amtSetOff115ad_1_b_ii_20) + (amtSetOffPTI20_STCG)
End Function
Function setOffAgainst112(value As Double, cgosIncome_cgInc_ltcg_prctg20 As Double)
Dim amtSetOff As Double
    'Var prctg20 = cgosIncome_cgInc_ltcg_prctg20
    prctg20Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso)
    
   amtSetOff = 0
    If (value) < 0 And (cgosIncome_cgInc_ltcg_prctg20_sec112) > 0 Then
        If (cgosIncome_cgInc_ltcg_prctg20_sec112) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_prctg20_sec112 = (cgosIncome_cgInc_ltcg_prctg20_sec112) + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg20_sec112)
            cgosIncome_cgInc_ltcg_prctg20_sec112 = 0
        End If
        If (amtSetOff) > (prctg20Sum) Then
            cgosIncome_cgInc_ltcg_prctg20_sec112 = (cgosIncome_cgInc_ltcg_prctg20_sec112) + (amtSetOff) - (prctg20Sum)
            amtSetOff = prctg20Sum
        End If
    End If
    setOffAgainst112 = amtSetOff
   End Function
   
   
   Function setOffAgainstPTI20(value As Double, cgosIncome_cgInc_ltcg_prctg20 As Double)
Dim amtSetOff As Double
    'Var prctg20 = cgosIncome_cgInc_ltcg_prctg20
    prctg20Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso)
    
   amtSetOff = 0
    If (value) < 0 And (cgosIncome_cgInc_ltcg_PTI20) > 0 Then
        If (cgosIncome_cgInc_ltcg_PTI20) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_PTI20 = (cgosIncome_cgInc_ltcg_PTI20) + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI20)
            cgosIncome_cgInc_ltcg_PTI20 = 0
        End If
        If (amtSetOff) > (prctg20Sum) Then
            cgosIncome_cgInc_ltcg_PTI20 = (cgosIncome_cgInc_ltcg_PTI20) + (amtSetOff) - (prctg20Sum)
            amtSetOff = prctg20Sum
        End If
    End If
    setOffAgainstPTI20 = amtSetOff
   End Function
    
Function setOffAgainst115Ea(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double)
Dim amtSetOff As Double
prctg20Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso)
    
   amtSetOff = 0
    If (value < 0 And cgosIncome_cgInc_ltcg_prctg20_sec11EA > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg20_sec11EA >= Abs(value)) Then
           cgosIncome_cgInc_ltcg_prctg20_sec11EA = cgosIncome_cgInc_ltcg_prctg20_sec11EA + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg20_sec11EA)
            cgosIncome_cgInc_ltcg_prctg20_sec11EA = 0
        End If
        If (amtSetOff > prctg20Sum) Then
            cgosIncome_cgInc_ltcg_prctg20_sec11EA = cgosIncome_cgInc_ltcg_prctg20_sec11EA + amtSetOff - prctg20Sum
            amtSetOff = prctg20Sum
        End If
    End If
    setOffAgainst115Ea = amtSetOff

  End Function

Function setOffAgainst15(value As Double, cgosIncome_cgInc_stcg_prctg15 As Double)
     
     amtSetOff111 = 0
     amtSetOff115ad_1_b_ii = 0
     amtSetOffPTI15 = 0
     
     setOffAgainst15 = 0
    If (value < 0) Then
    
        amtSetOff115ad_1_b_ii = setOffAgainst115ad_1_b_ii(value, cgosIncome_cgInc_stcg_prctg15)
        amtSetOffPTI15 = setOffAgainst15PTI((value + amtSetOff115ad_1_b_ii), cgosIncome_cgInc_stcg_prctg15)
        amtSetOff111a = setOffAgainst111a(value + amtSetOff115ad_1_b_ii + amtSetOffPTI15, cgosIncome_cgInc_stcg_prctg15)
        
    End If
    setOffAgainst15 = (amtSetOff111a) + (amtSetOff115ad_1_b_ii) + (amtSetOffPTI15)
End Function

Function setOffAgainst111a(value As Double, cgosIncome_cgInc_stcg_prctg15 As Double)
  Dim amtSetOff As Double
   prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI15
     amtSetOff = 0
     
    If ((value) < 0 And (cgosIncome_cgInc_stcg_prctg15_sec111a) > 0) Then
        If (cgosIncome_cgInc_stcg_prctg15_sec111a) >= Abs(value) Then
            cgosIncome_cgInc_stcg_prctg15_sec111a = (cgosIncome_cgInc_stcg_prctg15_sec111a) + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_prctg15_sec111a)
            cgosIncome_cgInc_stcg_prctg15_sec111a = 0
        End If
        If (amtSetOff) > (prctg15Sum) Then
           cgosIncome_cgInc_stcg_prctg15_sec111a = (cgosIncome_cgInc_stcg_prctg15_sec111a) + (amtSetOff) - (prctg15Sum)
            amtSetOff = (prctg15Sum)
        End If
End If
    setOffAgainst111a = amtSetOff
End Function
'Added by Riyaz on 03/04/2025
Function setOffAgainst111a_20(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double)
  Dim amtSetOff As Double
   prsctg20Sum = cgosIncome_cgInc_stcg_prctg20_sec111a + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI20
     amtSetOff = 0
     
    If ((value) < 0 And (cgosIncome_cgInc_stcg_prctg20_sec111a) > 0) Then
        If (cgosIncome_cgInc_stcg_prctg20_sec111a) >= Abs(value) Then
            cgosIncome_cgInc_stcg_prctg20_sec111a = (cgosIncome_cgInc_stcg_prctg20_sec111a) + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_prctg20_sec111a)
            cgosIncome_cgInc_stcg_prctg20_sec111a = 0
        End If
        If (amtSetOff) > (prsctg20Sum) Then
           cgosIncome_cgInc_stcg_prctg20_sec111a = (cgosIncome_cgInc_stcg_prctg20_sec111a) + (amtSetOff) - (prsctg20Sum)
            amtSetOff = (prsctg20Sum)
        End If
End If
    setOffAgainst111a_20 = amtSetOff
End Function
    


Function setOffAgainst15PTI(value As Double, cgosIncome_cgInc_stcg_prctg15 As Double)
Dim amtSetOff As Double
    setOffAgainst15PTI = 0
    prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI15
    amtSetOff = 0
    If (value) < 0 And (cgosIncome_cgInc_stcg_PTI15) > 0 Then
        If (cgosIncome_cgInc_stcg_PTI15) >= Abs(value) Then
            cgosIncome_cgInc_stcg_PTI15 = cgosIncome_cgInc_stcg_PTI15 + (value)
                                                  
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_PTI15)
            cgosIncome_cgInc_stcg_PTI15 = 0
        End If
        If ((amtSetOff) > (prctg15Sum)) Then
            cgosIncome_cgInc_stcg_PTI15 = (cgosIncome_cgInc_stcg_PTI15) + (amtSetOff) - (prctg15Sum)
            amtSetOff = (prctg15Sum)
        End If
    End If

    setOffAgainst15PTI = amtSetOff

End Function
'Added by Riyaz on 03/04/2025
Function setOffAgainst15PTI_20(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double)
Dim amtSetOff As Double
    setOffAgainst15PTI_20 = 0
    prsctg20Sum = cgosIncome_cgInc_stcg_prctg20_sec111a + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI20
    amtSetOff = 0
    If (value) < 0 And (cgosIncome_cgInc_stcg_PTI20) > 0 Then
        If (cgosIncome_cgInc_stcg_PTI20) >= Abs(value) Then
            cgosIncome_cgInc_stcg_PTI20 = cgosIncome_cgInc_stcg_PTI20 + (value)
                                                  
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_PTI20)
            cgosIncome_cgInc_stcg_PTI20 = 0
        End If
'08/04/2024 Changed by Riyaz
        If ((amtSetOff) > (prsctg20Sum)) Then
            cgosIncome_cgInc_stcg_PTI20 = (cgosIncome_cgInc_stcg_PTI20) + (amtSetOff) - (prsctg20Sum)
            amtSetOff = (prsctg20Sum)
        End If
    End If

    setOffAgainst15PTI_20 = amtSetOff

End Function

Function setOffAgainst115ad_1_b_ii(value As Double, cgosIncome_cgInc_stcg_prctg15 As Double)
Dim amtSetOff As Double
    setOffAgainst115ad_1_b_ii = 0
    prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI15
    amtSetOff = 0
    If (value) < 0 And (cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) > 0 Then
        If (cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) >= Abs(value) Then
            cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + (value)
                                                  
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii)
            cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = 0
        End If
        If ((amtSetOff) > (prctg15Sum)) Then
            cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = (cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) + (amtSetOff) - (prctg15Sum)
            amtSetOff = (prctg15Sum)
        End If
    End If

    setOffAgainst115ad_1_b_ii = amtSetOff

End Function
'Added by Riyaz on 03/04/2025
Function setOffAgainst115ad_1_b_ii_20(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double)
Dim amtSetOff As Double
    setOffAgainst115ad_1_b_ii_20 = 0
    prsctg20Sum = cgosIncome_cgInc_stcg_prctg20_sec111a + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI20
    amtSetOff = 0
    If (value) < 0 And (cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii) > 0 Then
        If (cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii) >= Abs(value) Then
            cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii + (value)
                                                  
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii)
            cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = 0
        End If
        If ((amtSetOff) > (prsctg20Sum)) Then
            cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = (cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii) + (amtSetOff) - (prsctg20Sum)
            amtSetOff = (prsctg20Sum)
        End If
    End If

    setOffAgainst115ad_1_b_ii_20 = amtSetOff

End Function

Function setOffAgainst10(value As Double, cgosIncome_cgInc_stcg_prctg10 As Double)
Dim amtSetOff As Double
   amtSetOff = 0
      If (value < 0) Then
        amtSetOff = amtSetOff + setOffAgainst112_1_c_2(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainst115AC_1(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainst115ACA_1(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainst115AD_3(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainst115AD_3_Proviso(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainst115E_b(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainstPTI10(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainstsecProviso(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
'Added by Riyaz on 26/06/2025
        amtSetOff = amtSetOff + setOffAgainstPTI10_112(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
        amtSetOff = amtSetOff + setOffAgainst112A(value + amtSetOff, cgosIncome_cgInc_stcg_prctg10)
    End If
    setOffAgainst10 = amtSetOff
End Function

'checked

Function setOffAgainst112_1_c_2(value As Double, cgosIncome_cgInc_stcg_prctg10 As Double) As Double
  Dim amtSetOff As Double
     prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)

     amtSetOff = 0
     
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + value
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2)
            cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = 0
        End If
        If (amtSetOff) > (prctg10Sum) Then
            cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
    End If
    setOffAgainst112_1_c_2 = amtSetOff
End Function
'checked

'This function not using from 10% to 20% changed
Function setOffAgainstsecProviso(value As Double, cgosIncome_cgInc_stcg_prctg10 As Double)
   Dim amtSetOff As Double
     prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)

     amtSetOff = 0
     
     setOffAgainstsecProviso = amtSetOff
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_secProviso) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg10_secProviso) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg10_secProviso = cgosIncome_cgInc_ltcg_prctg10_secProviso + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_secProviso)
           cgosIncome_cgInc_ltcg_prctg10_secProviso = 0
        End If
        If (amtSetOff) > (prctg10Sum) Then
         cgosIncome_cgInc_ltcg_prctg10_secProviso = cgosIncome_cgInc_ltcg_prctg10_secProviso + amtSetOff - prctg10Sum
            amtSetOff = (prctg10Sum)
        End If
    End If
    
   setOffAgainstsecProviso = amtSetOff
End Function

Function setOffAgainst115AC_1(value As Double, cgosIncome_cgInc_stcg_prctg10 As Double)
   Dim amtSetOff As Double
    prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)
                

     amtSetOff = 0

    If (value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec115AC_1) > 0 Then
        If ((cgosIncome_cgInc_ltcg_prctg10_sec115AC_1) >= Abs(value)) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec115AC_1)
             cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = 0
        End If
        If (amtSetOff) > (prctg10Sum) Then
            cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = (cgosIncome_cgInc_ltcg_prctg10_sec115AC_1) + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
    End If
    setOffAgainst115AC_1 = amtSetOff
End Function
'checked
Function setOffAgainst115AD_3(value, cgosIncome_cgInc_stcg_prctg10)
Dim amtSetOff As Double
   prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)
                

    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec115AD_3) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg10_sec115AD_3) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec115AD_3)
           cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = 0
        End If
        If ((amtSetOff) > (prctg10Sum)) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
        End If
    
    setOffAgainst115AD_3 = amtSetOff
End Function

Function setOffAgainst115AD_3_Proviso(value, cgosIncome_cgInc_stcg_prctg10)
Dim amtSetOff As Double
   prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)
                

    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso)
           cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = 0
        End If
        If ((amtSetOff) > (prctg10Sum)) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
        End If
    
    setOffAgainst115AD_3_Proviso = amtSetOff
End Function

 'checked
Function setOffAgainst115E_b(value, cgosIncome_cgInc_stcg_prctg10)
  Dim amtSetOff As Double
prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)
    
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec115E_b) > 0) Then
        If ((cgosIncome_cgInc_ltcg_prctg10_sec115E_b) >= Abs(value)) Then
            cgosIncome_cgInc_ltcg_prctg10_sec115E_b = cgosIncome_cgInc_ltcg_prctg10_sec115E_b + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec115E_b)
            cgosIncome_cgInc_ltcg_prctg10_sec115E_b = 0
        End If
        If (amtSetOff) > (prctg10Sum) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115E_b = (cgosIncome_cgInc_ltcg_prctg10_sec115E_b) + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
    End If
    
    setOffAgainst115E_b = amtSetOff

  End Function

Function setOffAgainstPTI10(value, cgosIncome_cgInc_stcg_prctg10)
Dim amtSetOff As Double
   prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)

    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI10) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI10) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_PTI10 = cgosIncome_cgInc_ltcg_PTI10 + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI10)
           cgosIncome_cgInc_ltcg_PTI10 = 0
        End If
        If ((amtSetOff) > (prctg10Sum)) Then
           cgosIncome_cgInc_ltcg_PTI10 = cgosIncome_cgInc_ltcg_PTI10 + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
        End If
    
    setOffAgainstPTI10 = amtSetOff
End Function
Function setOffAgainstPTI10_112(value, cgosIncome_cgInc_stcg_prctg10)
Dim amtSetOff As Double
   prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)

    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI10_112) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI10_112) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_PTI10_112 = cgosIncome_cgInc_ltcg_PTI10_112 + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI10_112)
           cgosIncome_cgInc_ltcg_PTI10_112 = 0
        End If
        If ((amtSetOff) > (prctg10Sum)) Then
           cgosIncome_cgInc_ltcg_PTI10_112 = cgosIncome_cgInc_ltcg_PTI10_112 + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
        End If
    
    setOffAgainstPTI10_112 = amtSetOff
End Function
Function setOffAgainst112A(value, cgosIncome_cgInc_stcg_prctg10)
Dim amtSetOff As Double
   prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)

    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec112A) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg10_sec112A) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg10_sec112A = cgosIncome_cgInc_ltcg_prctg10_sec112A + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec112A)
           cgosIncome_cgInc_ltcg_prctg10_sec112A = 0
        End If
        If ((amtSetOff) > (prctg10Sum)) Then
           cgosIncome_cgInc_ltcg_prctg10_sec112A = cgosIncome_cgInc_ltcg_prctg10_sec112A + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
        End If
    
    setOffAgainst112A = amtSetOff
End Function
  
Function setOffAgainst115ACA_1(value, cgosIncome_cgInc_stcg_prctg10)
  Dim amtSetOff As Double
prctg10Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112)
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1) > 0) Then
        If ((cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1) >= Abs(value)) Then
            cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1)
            cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = 0
        End If
        If (amtSetOff) > (prctg10Sum) Then
           cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = (cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1) + (amtSetOff) - (prctg10Sum)
            amtSetOff = (prctg10Sum)
        End If
    End If
    
    setOffAgainst115ACA_1 = amtSetOff

  End Function
  

  Function setOffAgainst_stcg115AD(value, cgosIncome_cgInc_stcg_prctg10)
  Dim amtSetOff As Double
prctg30Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_stcg_prctg30_sec115AD + cgosIncome_cgInc_stcg_PTI30)
    
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_stcg_prctg30_sec115AD) > 0) Then
        If ((cgosIncome_cgInc_stcg_prctg30_sec115AD) >= Abs(value)) Then
            cgosIncome_cgInc_stcg_prctg30_sec115AD = cgosIncome_cgInc_stcg_prctg30_sec115AD + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_prctg30_sec115AD)
            cgosIncome_cgInc_stcg_prctg30_sec115AD = 0
        End If
        If (amtSetOff) > (prctg30Sum) Then
           cgosIncome_cgInc_stcg_prctg30_sec115AD = (cgosIncome_cgInc_stcg_prctg30_sec115AD) + (amtSetOff) - (prctg30Sum)
            amtSetOff = (prctg30Sum)
        End If
    End If
    
    setOffAgainst_stcg115AD = amtSetOff

  End Function
  
  Function setOffAgainst_stcgPTI30(value, cgosIncome_cgInc_stcg_prctg10)
  Dim amtSetOff As Double
prctg30Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_stcg_prctg30_sec115AD + cgosIncome_cgInc_stcg_PTI30)
    
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_stcg_PTI30) > 0) Then
        If ((cgosIncome_cgInc_stcg_PTI30) >= Abs(value)) Then
            cgosIncome_cgInc_stcg_PTI30 = cgosIncome_cgInc_stcg_PTI30 + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_PTI30)
            cgosIncome_cgInc_stcg_PTI30 = 0
        End If
        If (amtSetOff) > (prctg30Sum) Then
           cgosIncome_cgInc_stcg_PTI30 = (cgosIncome_cgInc_stcg_PTI30) + (amtSetOff) - (prctg30Sum)
            amtSetOff = (prctg30Sum)
        End If
    End If
    
    setOffAgainst_stcgPTI30 = amtSetOff

  End Function
    
  Function setOffAgainst_stcgAR(value, cgosIncome_cgInc_stcg_prctg10)
  Dim amtSetOff As Double
prctgARSum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_stcg_Applicable + cgosIncome_cgInc_stcg_ApplicablePTI)
    
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_stcg_Applicable) > 0) Then
        If ((cgosIncome_cgInc_stcg_Applicable) >= Abs(value)) Then
            cgosIncome_cgInc_stcg_Applicable = cgosIncome_cgInc_stcg_Applicable + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_Applicable)
            cgosIncome_cgInc_stcg_Applicable = 0
        End If
        If (amtSetOff) > (prctgARSum) Then
           cgosIncome_cgInc_stcg_Applicable = (cgosIncome_cgInc_stcg_Applicable) + (amtSetOff) - (prctgARSum)
            amtSetOff = (prctgARSum)
        End If
    End If
    
    setOffAgainst_stcgAR = amtSetOff

  End Function
  
  Function setOffAgainst_stcgARPTI(value, cgosIncome_cgInc_stcg_prctg10)
  Dim amtSetOff As Double
prctgARSum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_stcg_Applicable + cgosIncome_cgInc_stcg_ApplicablePTI)
    
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_stcg_ApplicablePTI) > 0) Then
        If ((cgosIncome_cgInc_stcg_ApplicablePTI) >= Abs(value)) Then
            cgosIncome_cgInc_stcg_ApplicablePTI = cgosIncome_cgInc_stcg_ApplicablePTI + value
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_stcg_ApplicablePTI)
            cgosIncome_cgInc_stcg_ApplicablePTI = 0
        End If
        If (amtSetOff) > (prctgARSum) Then
           cgosIncome_cgInc_stcg_ApplicablePTI = (cgosIncome_cgInc_stcg_ApplicablePTI) + (amtSetOff) - (prctgARSum)
            amtSetOff = (prctgARSum)
        End If
    End If
    
    setOffAgainst_stcgARPTI = amtSetOff

  End Function

Function Validate_SUMMARY_CG() As Boolean
Validate_SUMMARY_CG = True

'Newly added by Bindu for Table E Editable format on 20th May 25
    If Sheet13.Range("TableE_Negativecount").value > 0 Then
        msgboxSHEETCG = ("Amount of set off claimed cannot exceed amount available for set off")
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    Dim STCG_Setoff As Variant
    Dim LTCG_Setoff As Variant
    
    Dim STCG15_total As Variant
    Dim STCG20_total As Variant
    Dim STCG30_total As Variant
    Dim STCGAPP_total As Variant
    Dim STCGDTAA_total As Variant
    
    Dim LTCG10_total As Variant
    Dim LTCG12_5_total As Variant
    Dim LTCG20_total As Variant
    Dim LTCGDTAA_total As Variant
    
    Dim STCG_Total As Variant
    Dim LTCG_Total As Variant
    
    LTCG_Total = WorksheetFunction.Sum(Sheet13.Range("IHLA.Ev7_CurrYrCapGain").value, Sheet13.Range("IHLA.Eviii7_CurrYrCapGain12").value, Sheet13.Range("IHLA.Evi7_CurrYrCapGain").value, Sheet13.Range("IHLA.E_LTCGv7_DTAArates").value)

    STCG_Total = WorksheetFunction.Sum(Sheet13.Range("IHLA.Eii7_CurrYrCapGain").value, Sheet13.Range("IHLA.Eiii7_CurrYrCapGain20").value, Sheet13.Range("IHLA.Eiii7_CurrYrCapGain").value, Sheet13.Range("IHLA.Eiv7_CurrYrCapGain").value, Sheet13.Range("IHLA.Ev7_DTAArates").value) + LTCG_Total
    
    
    
    STCG15_total = Sheet13.Range("IHLA.Eviii2_StclSetoff15Per").value
    STCG20_total = Sheet13.Range("IHLA.Exii2_StclSetoff20Per").value
    STCG30_total = Sheet13.Range("IHLA.Eviii3_StclSetoff30Per").value
    STCGAPP_total = Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate").value
    STCGDTAA_total = Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate").value
    
    LTCG10_total = Sheet13.Range("IHLA.Eviii5_LtclSetOff10Per").value
    LTCG12_5_total = Sheet13.Range("IHLA.Exii8_LtclSetOff12.5Per").value
    LTCG20_total = Sheet13.Range("IHLA.Eviii6_LtclSetOff20Per").value
    LTCGDTAA_total = Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate").value
    
    STCG_Setoff = STCG15_total + STCG20_total + STCG30_total + STCGAPP_total + STCGDTAA_total
    LTCG_Setoff = LTCG10_total + LTCG12_5_total + LTCG20_total + LTCGDTAA_total
    
    
    
    If STCG_Setoff <> 0 And STCG_Total <> 0 Then
      If STCG_Total <> 0 And STCG_Setoff <= STCG_Total Then
        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
        Validate_SUMMARY_CG = False
        Exit Function
      ElseIf STCG_Total <> 0 And STCG_Setoff >= STCG_Total Then
        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
        Validate_SUMMARY_CG = False
        Exit Function
      End If
   End If
   
   If LTCG_Setoff <> 0 And LTCG_Total <> 0 Then
      If LTCG_Total <> 0 And LTCG_Setoff <= LTCG_Total Then
        msgboxSHEETCG = ("Entire Long term capital loss is not set off with income available for set off")
        Validate_SUMMARY_CG = False
        Exit Function
      ElseIf LTCG_Total <> 0 And LTCG_Setoff >= LTCG_Total Then
        msgboxSHEETCG = ("Entire Long term capital loss is not set off with income available for set off")
        Validate_SUMMARY_CG = False
        Exit Function
      End If
   End If
    
    
    
    
    '15% STCG
    
    
'    If STCG15_total <> 0 And STCG15_total < STCG_Total Then
'        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If STCG20_total <> 0 And STCG20_total < STCG_Total Then
'        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If STCG30_total <> 0 And STCG30_total < STCG_Total Then
'        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If STCGAPP_total <> 0 And STCGAPP_total < STCG_Total Then
'        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If STCGDTAA_total <> 0 And STCGDTAA_total < STCG_Total Then
'        msgboxSHEETCG = ("Entire short term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'
'
'    If LTCG10_total <> 0 And LTCG10_total < LTCG_Total Then
'        msgboxSHEETCG = ("Entire Long term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If LTCG12_5_total <> 0 And LTCG12_5_total < LTCG_Total Then
'        msgboxSHEETCG = ("Entire Long term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If LTCG20_total <> 0 And LTCG20_total < LTCG_Total Then
'        msgboxSHEETCG = ("Entire Long term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
'    If LTCGDTAA_total <> 0 And LTCGDTAA_total < LTCG_Total Then
'        msgboxSHEETCG = ("Entire Long term capital loss is not set off with income available for set off")
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
''----



    
    Dim a3_16 As Variant
    Dim a3a_16 As Variant
    Dim a3b_16 As Variant
    Dim a4_16 As Variant
    Dim a4a_16 As Variant
    Dim a5_16 As Variant
    Dim a6_16 As Variant
    
'by sai on 03/04/2025*************
    Dim a3c_16 As Variant
    Dim a4b_16 As Variant
'*********************************
    

    Dim F_54EC, F_54, F_54F, F_54GB, F_115F As Variant
    Dim F_54B As Variant
    Dim F_54D As Variant
    Dim F_54G As Variant
    Dim F_54GA As Variant
    Dim F_54EE As Variant
    
    Dim D_54EC, D_54EE, D_54B, D_54, D_54F, D_54GB, D_115F As Variant
    Dim D_54G As Variant
    Dim D_54D As Variant
    Dim D_54GA As Variant

    a3_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3").value
    
'by sai on 03/04/2025
    a3c_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3c").value    '20%
    
    
    a3a_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3a").value
    a3b_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b").value
    
    a4_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4").value
    
' by sai on 03/04/2025
    a4b_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4b").value '12.5%
    
    a4a_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4a").value
    
    a5_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b_DTAARates").value
    
    a6_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses_DTAARates").value
    
    'Change 102 21.12.2023 V0.6 SIT-65102 C
    Dim a8_16
    a8_16 = Sheet21.Range("SPI_incheadCG_VDAF8").value
'    End Change SIT-65102 C
     
 ' Messages modified by sai on 29/03/2025 as 2 rows were added newly for AY 2025-26
     
     'Commented by Bindu on 23rd Feb 26
'    If (a3_16 <> AccSTCG_Upto15Of6 + AccSTCG_Upto15Of9 + ACCSTCG_Up16Of9To15Of12 + ACCSTCG_Up16Of12To15Of3 + ACCSTCG_Up16Of3To31Of3) Then
'        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(1) in Sheet CG for STCG 15% is equal to item 3via of schedule BFLA - Rs." & a3_16)
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
    
    
'For AY 2025-26 by sai on 15/04/2025
'***************************************************************************************************
    If (a3c_16 <> ACCLTCGDTAA_Upto15Of6_20 + ACCLTCGDTAA_Upto15Of9_20 + ACCLTCGDTAA_Up16Of9To15Of12_20 + ACCLTCGDTAA_Up16Of12To15Of3_20 + ACCLTCGDTAA_Up16Of3To31Of3_20) Then
        'Updated by Bindu from F(2) to F(1) on 23rd Feb 26
        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(1) in Sheet CG for STCG 20% is equal to item 3vib of schedule BFLA - Rs." & a3c_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
'****************************************************************************************************



    If (a3a_16 <> AccSTCG30_Upto15Of6 + AccSTCG30_Upto15Of9 + AccSTCG30_Up16Of9To15Of12 + AccSTCG30_Up16Of12To15Of3 + AccSTCG30_Up16Of3To31Of3) Then
        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(2) in Sheet CG for STCG 30% is equal to item 3vii of schedule BFLA - Rs" & a3a_16)
'By Bindu on 23rd Feb 26 from F(3) to F(2)
'         msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(3) in Sheet CG for STCG 30% is equal to item 3vii of schedule BFLA - Rs" & a3a_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    If (a3b_16 <> ACCSTCGOTH_Upto15Of6 + ACCSTCGOTH_Upto15Of9 + ACCSTCGOTH_Up16Of9To15Of12 + ACCSTCGOTH_Up16Of12To15Of3 + ACCSTCGOTH_Up16Of3To31Of3) Then
        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(3) in Sheet CG for STCG applicable rate is equal to item 3viii of schedule BFLA - Rs." & a3b_16)
'By Bindu on 23rd Feb 26 from F(4) to F(3)
'         msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(4) in Sheet CG for STCG applicable rate is equal to item 3viii of schedule BFLA - Rs." & a3b_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    
  'By Bindu on 23rd Feb 26
'    If (a4_16 <> ACCLTCG_Upto15Of6 + ACCLTCG_Upto15Of9 + ACCLTCG_Up16Of9To15Of12 + ACCLTCG_Up16Of12To15Of3 + ACCLTCG_Up16Of3To31Of3) Then
' '        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(5) in Sheet CG for LTCG 10% is equal to item 3(x) of schedule BFLA - Rs." & a4_16)
'          msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(6) in Sheet CG for LTCG 10% is equal to item 3(xa) of schedule BFLA - Rs." & a4_16)
'
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
    
'For AY 2025-26 by sai on 15/04/2025
'*******************************************************************************
    If (a4b_16 <> ACCLTCGDTAA_Upto15Of6_125 + ACCLTCGDTAA_Upto15Of9_125 + ACCLTCGDTAA_Up16Of9To15Of12_125 + ACCLTCGDTAA_Up16Of12To15Of3_125 + ACCLTCGDTAA_Up16Of3To31Of3_125) Then
        'BY Bindu on 23rd Feb 26 from F(7) to F(6)
        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(5) in Sheet CG for LTCG 12.5% is equal to item 3(xb) of schedule BFLA - Rs." & a4b_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
'********************************************************************************

'By Bindu on 23rd Feb 26
'    If (a4a_16 <> ACCLTCGNP_Upto15Of6 + ACCLTCGNP_Upto15Of9 + ACCLTCGNP_Up16Of9To15Of12 + ACCLTCGNP_Up16Of12To15Of3 + ACCLTCGNP_Up16Of3To31Of3) Then
''        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(6) in Sheet CG for LTCG 20% is equal to item 3(xi) of schedule BFLA - Rs." & a4a_16)
'         msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(8) in Sheet CG for LTCG 20% is equal to item 3(xi) of schedule BFLA - Rs." & a4a_16)
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If
    
    If (a5_16 <> ACCSTCGDTAA_Upto15Of6 + ACCSTCGDTAA_Upto15Of9 + ACCSTCGDTAA_Up16Of9To15Of12 + ACCSTCGDTAA_Up16Of12To15Of3 + ACCSTCGDTAA_Up16Of3To31Of3) Then
       msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(4) in Sheet CG for STCG DTAA is equal to item 3(ix) of schedule BFLA - Rs." & a5_16)
 'By Bindu from F(5) to F(4) on 23rd Feb 26
'         msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(5) in Sheet CG for STCG DTAA is equal to item 3(ix) of schedule BFLA - Rs." & a5_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    If (a6_16 <> ACCLTCGDTAA_Upto15Of6 + ACCLTCGDTAA_Upto15Of9 + ACCLTCGDTAA_Up16Of9To15Of12 + ACCLTCGDTAA_Up16Of12To15Of3 + ACCLTCGDTAA_Up16Of3To31Of3) Then
'        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(7) in Sheet CG for LTCG DTAA is equal to item 3(xii) of schedule BFLA - Rs." & a6_16)
       'By Bindu on From F(9) to F(6) on 23rd Feb 26
         msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(6) in Sheet CG for LTCG DTAA is equal to item 3(xii) of schedule BFLA - Rs." & a6_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
'    If (a6_16 <> ACCLTCGDTAA_Upto15Of6 + ACCLTCGDTAA_Upto15Of9 + ACCLTCGDTAA_Up16Of9To15Of12 + ACCLTCGDTAA_Up16Of12To15Of3 + ACCLTCGDTAA_Up16Of3To31Of3) Then
'        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(7) in Sheet CG for LTCG DTAA is equal to item 3(xii) of schedule BFLA - Rs." & a6_16)
'        Validate_SUMMARY_CG = False
'        Exit Function
'    End If


    'Change 102 21.12.2023 V0.6 SIT-65102 D
    If (a8_16 <> AccVDA_Upto15Of6 + AccVDA_Upto15Of9 + AccVDA_Up16Of9To15Of12 + AccVDA_Up16Of12To15Of3 + AccVDA_Up16Of3To31Of3) Then
'        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(8) in Sheet CG for for Capital gains on transfer of Virtual digital asset taxable at the rate of 30% is not equal to value mentioned in schedule SI - Rs." & a8_16)
       'By Bindu from F(10) to F(7) on 23rd Feb 26
        msgboxSHEETCG = ("Please ensure that the Total of the values in Table F(7) in Sheet CG for for Capital gains on transfer of Virtual digital asset taxable at the rate of 30% is not equal to value mentioned in schedule SI - Rs." & a8_16)
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    'End Change SIT-65102 D

    
'    F_54EE = IIf(Len(Sheet13.Range("F_54EE")) = 0, 0, Sheet13.Range("F_54EE").value)
    F_54EC = IIf(Len(Sheet13.Range("F_54EC")) = 0, 0, Sheet13.Range("F_54EC").value)
    F_54D = IIf(Len(Sheet13.Range("F_54D")) = 0, 0, Sheet13.Range("F_54D").value)
    F_54G = IIf(Len(Sheet13.Range("F_54G")) = 0, 0, Sheet13.Range("F_54G").value)
    F_54GA = IIf(Len(Sheet13.Range("F_54GA")) = 0, 0, Sheet13.Range("F_54GA").value)
    
    F_54B = IIf(Len(Sheet13.Range("F_54B")) = 0, 0, Sheet13.Range("F_54B").value)
    F_54 = IIf(Len(Sheet13.Range("F_54")) = 0, 0, Sheet13.Range("F_54").value)
    F_54F = IIf(Len(Sheet13.Range("F_54F")) = 0, 0, Sheet13.Range("F_54F").value)
    F_54GB = IIf(Len(Sheet13.Range("F_54GB")) = 0, 0, Sheet13.Range("F_54GB").value)
    F_115F = IIf(Len(Sheet13.Range("F_115F")) = 0, 0, Sheet13.Range("F_115F").value)
    
'    D_54EE = IIf(Len(Sheet13.Range("D_54EE")) = 0, 0, Sheet13.Range("D_54EE").value)
    D_54EC = IIf(Len(Sheet13.Range("D_54EC")) = 0, 0, Sheet13.Range("D_54EC").value)
    D_54G = IIf(Len(Sheet13.Range("D_54G")) = 0, 0, Sheet13.Range("D_54G").value)
    D_54D = IIf(Len(Sheet13.Range("D_54D")) = 0, 0, Sheet13.Range("D_54D").value)
    D_54GA = IIf(Len(Sheet13.Range("D_54GA")) = 0, 0, Sheet13.Range("D_54GA").value)
    
    D_54B = IIf(Len(Sheet13.Range("D_54B")) = 0, 0, Sheet13.Range("D_54B").value)
    D_54 = IIf(Len(Sheet13.Range("D_54")) = 0, 0, Sheet13.Range("D_54").value)
    D_54F = IIf(Len(Sheet13.Range("D_54F")) = 0, 0, Sheet13.Range("D_54F").value)
    D_54GB = IIf(Len(Sheet13.Range("D_54GB")) = 0, 0, Sheet13.Range("D_54GB").value)
    

    D_115F = IIf(Len(Sheet13.Range("D_115F")) = 0, 0, Sheet13.Range("D_115F").value)
    
    If (F_54B <> D_54B) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54B match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If

    
    If (F_54D <> D_54D) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54D match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If

    If (F_54G <> D_54G) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54G match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    If (F_54GA <> D_54GA) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54GA match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    If (F_54EC > 5000000) Then
        msgboxSHEETCG = ("Total of the Deductions u/s 54EC does not exceed 5000000 in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If

    If (F_54EC <> D_54EC) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54EC match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    If (F_54 <> D_54) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54 match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If

    If (F_54G <> D_54G) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54G match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    If (F_54F <> D_54F) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 54F match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If

    If (F_115F <> D_115F) Then
        msgboxSHEETCG = ("Please ensure that the Total of the Deductions u/s 115F match the same in Table D in Sch CG")
        Validate_SUMMARY_CG = False
        Exit Function
    End If
    
    
'     If Sheet13.Range("LTCG_Temp54GB").value > 0 Then
'    DEDN_D2_DeductPAN = Sheet13.Range("DEDN.D2_DeductPAN").value
'
'        If DEDN_D2_DeductPAN = "" Then
'            msgboxSHEETCG = "PAN required since Deduction claimed u/s 54GB in Sch CG Table D"
'            Validate_SUMMARY_CG = False
'            Exit Function
'        Else
'            If Not ValidatePAN(CStr(DEDN_D2_DeductPAN)) Then
'                msgboxSHEETCG = "Please enter valid PAN for Deduction under S 54GB in Table D of Sch CG. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."
'                Validate_SUMMARY_CG = False
'                Exit Function
'            End If
'        End If
'    End If
    
End Function

Function Validate_ALL_CG() As Boolean
    Validate_ALL_CG = True
    
    STCG_A1e_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A1e_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A1e_CapgainonAssets").value)
    
    STCG_A20ia_FullConsideration = IIf(Len(Sheet13.Range("STCG.A20ia_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A20ia_FullConsideration").value)
    STCG_A20ib_NetWorthOfDivision = IIf(Len(Sheet13.Range("STCG.A20ib_NetWorthOfDivision")) = 0, 0, Sheet13.Range("STCG.A20ib_NetWorthOfDivision").value)
    STCG_A20ic_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A20ic_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A20ic_CapgainonAssets").value)
    
    'For 3ia P Coulumn old
     STCG_A2ia_FullConsideration = IIf(Len(Sheet13.Range("STCG.A2ia_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A2ia_FullConsideration").value)
     
'3ia Q Coulumns new fields For AY 2025-26 added by sai
     '*********************************************************************************************
        STCG_A2ia_2eii_FullConsideration = IIf(Len(Sheet13.Range("STCG.A2ia_2eii_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A2ia_2eii_FullConsideration").value)
        STCG_A2ibi_2eii_AquisitCost = IIf(Len(Sheet13.Range("STCG.A2ibi_2eii_AquisitCost")) = 0, 0, Sheet13.Range("STCG.A2ibi_2eii_AquisitCost").value)
        STCG_A2ibii_2eii_ImproveCost = IIf(Len(Sheet13.Range("STCG.A2ibii_2eii_ImproveCost")) = 0, 0, Sheet13.Range("STCG.A2ibii_2eii_ImproveCost").value)
        STCG_A2ibiii_2eii_ExpOnTrans = IIf(Len(Sheet13.Range("STCG.A2ibiii_2eii_ExpOnTrans")) = 0, 0, Sheet13.Range("STCG.A2ibiii_2eii_ExpOnTrans").value)
        STCG_A2ibiv_TotalDedn_ii = IIf(Len(Sheet13.Range("STCG.A2ibiv_2eii_TotalDedn")) = 0, 0, Sheet13.Range("STCG.A2ibiv_2eii_TotalDedn").value)
        STCG_A2ic_2eii_BalanceCG = IIf(Len(Sheet13.Range("STCG.A2ic_2eii_BalanceCG")) = 0, 0, Sheet13.Range("STCG.A2ic_2eii_BalanceCG").value)
        STCG_A2id_2eii_LossSec94of7Or94of8 = IIf(Len(Sheet13.Range("STCG.A2id_2eii_LossSec94of7Or94of8")) = 0, 0, Sheet13.Range("STCG.A2id_2eii_LossSec94of7Or94of8").value)
        
        
        STCG_A2ie_2eii_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A2ie_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A2ie_CapgainonAssets").value)
        STCG_A2ie_2ei_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A2ie_2ei_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A2ie_2ei_CapgainonAssets").value)

    
        'P+Q P+Q Sum Variant Sachin
'       STCG_A2ie_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A2ie_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A2ie_CapgainonAssets").value)
        STCG_A2ie_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A2ie_CapgainonAssets_i_ii")) = 0, 0, Sheet13.Range("STCG.A2ie_CapgainonAssets_i_ii").value)
'3ia end***************************************************************************************
         

         
    STCG_A2ibi_AquisitCost = IIf(Len(Sheet13.Range("STCG.A2ibi_AquisitCost")) = 0, 0, Sheet13.Range("STCG.A2ibi_AquisitCost").value)
    STCG_A2ibii_ImproveCost = IIf(Len(Sheet13.Range("STCG.A2ibii_ImproveCost")) = 0, 0, Sheet13.Range("STCG.A2ibii_ImproveCost").value)
    STCG_A2ibiii_ExpOnTrans = IIf(Len(Sheet13.Range("STCG.A2ibiii_ExpOnTrans")) = 0, 0, Sheet13.Range("STCG.A2ibiii_ExpOnTrans").value)
    
    'For 3ibiv P column old
    STCG_A2ibiv_TotalDedn = IIf(Len(Sheet13.Range("STCG.A2ibiv_TotalDedn")) = 0, 0, Sheet13.Range("STCG.A2ibiv_TotalDedn").value)

    STCG_A2ic_BalanceCG = IIf(Len(Sheet13.Range("STCG.A2ic_BalanceCG")) = 0, 0, Sheet13.Range("STCG.A2ic_BalanceCG").value)
    STCG_A2id_LossSec94of7Or94of8 = IIf(Len(Sheet13.Range("STCG.A2id_LossSec94of7Or94of8")) = 0, 0, Sheet13.Range("STCG.A2id_LossSec94of7Or94of8").value)
    

    'For 3iia P Column old
    STCG_A2iia_FullConsideration = IIf(Len(Sheet13.Range("STCG.A2iia_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A2iia_FullConsideration").value)
    
'3iia_Q Coulumns new fields For AY 2025-26 added by sai*******************************
        STCG_A2iia_FullConsideration_ii = IIf(Len(Sheet13.Range("STCG.A2iia_FullConsideration_ii")) = 0, 0, Sheet13.Range("STCG.A2iia_FullConsideration_ii").value)
        STCG_A2iibi_AquisitCost_ii = IIf(Len(Sheet13.Range("STCG.A2iibi_AquisitCost_ii")) = 0, 0, Sheet13.Range("STCG.A2iibi_AquisitCost_ii").value)
        STCG_A2iibii_ImproveCost_ii = IIf(Len(Sheet13.Range("STCG.A2iibii_ImproveCost_ii")) = 0, 0, Sheet13.Range("STCG.A2iibii_ImproveCost_ii").value)
        STCG_A2iibiii_ExpOnTrans_ii = IIf(Len(Sheet13.Range("STCG.A2iibiii_ExpOnTrans_ii")) = 0, 0, Sheet13.Range("STCG.A2iibiii_ExpOnTrans_ii").value)
        STCG_A2iibiv_TotalDedn_ii = IIf(Len(Sheet13.Range("STCG.A2iibiv_TotalDedn_ii")) = 0, 0, Sheet13.Range("STCG.A2iibiv_TotalDedn_ii").value)
        STCG_A2iic_BalanceCG_ii = IIf(Len(Sheet13.Range("STCG.A2iic_BalanceCG_ii")) = 0, 0, Sheet13.Range("STCG.A2iic_BalanceCG_ii").value)
        STCG_A2iid_LossSec94of7Or94of8_ii = IIf(Len(Sheet13.Range("STCG.A2iid_LossSec94of7Or94of8_ii")) = 0, 0, Sheet13.Range("STCG.A2iid_LossSec94of7Or94of8_ii").value)
        
        STCG_A2iie_CapgainonAssets_ii = IIf(Len(Sheet13.Range("STCG.A2iie_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A2iie_CapgainonAssets").value)
        STCG_A2iie_CapgainonAssets_i = IIf(Len(Sheet13.Range("STCG.A2iie_CapgainonAssets_i")) = 0, 0, Sheet13.Range("STCG.A2iie_CapgainonAssets_i").value)
      'P+Q Sum Variant
'       STCG_A2iie_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A2iie_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A2iie_CapgainonAssets").value)
        STCG_A2iie_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A2iie_CapgainonAssets_i_ii")) = 0, 0, Sheet13.Range("STCG.A2iie_CapgainonAssets_i_ii").value)
'3iia_Q end**************************************************************************
    

    
    
    
    STCG_A2iibi_AquisitCost = IIf(Len(Sheet13.Range("STCG.A2iibi_AquisitCost")) = 0, 0, Sheet13.Range("STCG.A2iibi_AquisitCost").value)
    STCG_A2iibii_ImproveCost = IIf(Len(Sheet13.Range("STCG.A2iibii_ImproveCost")) = 0, 0, Sheet13.Range("STCG.A2iibii_ImproveCost").value)
    STCG_A2iibiii_ExpOnTrans = IIf(Len(Sheet13.Range("STCG.A2iibiii_ExpOnTrans")) = 0, 0, Sheet13.Range("STCG.A2iibiii_ExpOnTrans").value)
    
    'For 3iia P Column old
    STCG_A2iibiv_TotalDedn = IIf(Len(Sheet13.Range("STCG.A2iibiv_TotalDedn")) = 0, 0, Sheet13.Range("STCG.A2iibiv_TotalDedn").value)


    STCG_A2iic_BalanceCG = IIf(Len(Sheet13.Range("STCG.A2iic_BalanceCG")) = 0, 0, Sheet13.Range("STCG.A2iic_BalanceCG").value)
    STCG_A2iid_LossSec94of7Or94of8 = IIf(Len(Sheet13.Range("STCG.A2iid_LossSec94of7Or94of8")) = 0, 0, Sheet13.Range("STCG.A2iid_LossSec94of7Or94of8").value)
    
'    STCG_A3a_NRItaxSTTPaid = IIf(Len(Sheet13.Range("STCG.A3a_NRItaxSTTPaid")) = 0, 0, Sheet13.Range("STCG.A3a_NRItaxSTTPaid").value)
    'For AY 2025-26 by Sai
        '================================================================================================================
          STCG_A3a_NRItaxSTTPaid = IIf(Len(Sheet13.Range("STCG.A3a_NRItaxSTTPaid_i_ii")) = 0, 0, Sheet13.Range("STCG.A3a_NRItaxSTTPaid_i_ii").value)
          STCG_A3ai_NRItaxSTTPaid = IIf(Len(Sheet13.Range("STCG.A3a_NRItaxSTTPaid")) = 0, 0, Sheet13.Range("STCG.A3a_NRItaxSTTPaid").value)
          STCG_A3aii_NRItaxSTTPaid = IIf(Len(Sheet13.Range("STCG.A3aii_NRItaxSTTPaid")) = 0, 0, Sheet13.Range("STCG.A3aii_NRItaxSTTPaid").value)
        '=================================================================================================================
    
    
    
    
    STCG_A3b_NRItaxSTTNotPaid = IIf(Len(Sheet13.Range("STCG.A3b_NRItaxSTTNotPaid")) = 0, 0, Sheet13.Range("STCG.A3b_NRItaxSTTNotPaid").value)
    
    STCG_A4a_FullConsideration1 = IIf(Len(Sheet13.Range("STCG.A4aia_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A4aia_FullConsideration").value)
    STCG_A4a_FullConsideration2 = IIf(Len(Sheet13.Range("STCG.A4aib_FullMarketvalue")) = 0, 0, Sheet13.Range("STCG.A4aib_FullMarketvalue").value)
    STCG_A4a_FullConsideration3 = IIf(Len(Sheet13.Range("STCG.A4aic_FullConsideration50CA")) = 0, 0, Sheet13.Range("STCG.A4aic_FullConsideration50CA").value)
    STCG_A4a_FullConsideration4 = IIf(Len(Sheet13.Range("STCG.A4aii_FullConsiderationS")) = 0, 0, Sheet13.Range("STCG.A4aii_FullConsiderationS").value)
    STCG_A4a_FullConsideration = IIf(Len(Sheet13.Range("STCG.A4a_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A4a_FullConsideration").value)
    STCG_A4bi_AquisitCost = IIf(Len(Sheet13.Range("STCG.A4bi_AquisitCost")) = 0, 0, Sheet13.Range("STCG.A4bi_AquisitCost").value)
    STCG_A4bii_ImproveCost = IIf(Len(Sheet13.Range("STCG.A4bii_ImproveCost")) = 0, 0, Sheet13.Range("STCG.A4bii_ImproveCost").value)
    STCG_A4biii_ExpOnTrans = IIf(Len(Sheet13.Range("STCG.A4biii_ExpOnTrans")) = 0, 0, Sheet13.Range("STCG.A4biii_ExpOnTrans").value)
    STCG_A4biv_TotalDedn = IIf(Len(Sheet13.Range("STCG.A4biv_TotalDedn")) = 0, 0, Sheet13.Range("STCG.A4biv_TotalDedn").value)
    STCG_A4c_BalanceCG = IIf(Len(Sheet13.Range("STCG.A4c_BalanceCG")) = 0, 0, Sheet13.Range("STCG.A4c_BalanceCG").value)
    STCG_A4d_LossSec94of7Or94of8 = IIf(Len(Sheet13.Range("STCG.A4d_LossSec94of7Or94of8")) = 0, 0, Sheet13.Range("STCG.A4d_LossSec94of7Or94of8").value)
    STCG_A4e_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A4e_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A4e_CapgainonAssets").value)
    
    STCG_A5a_FullConsideration = IIf(Len(Sheet13.Range("STCG.A5a_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A5a_FullConsideration").value)
    STCG_A5a_FullConsideration1 = IIf(Len(Sheet13.Range("STCG.A5aia_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A5aia_FullConsideration").value)
    STCG_A5a_FullConsideration2 = IIf(Len(Sheet13.Range("STCG.A5aib_Fullmarketvalue")) = 0, 0, Sheet13.Range("STCG.A5aib_Fullmarketvalue").value)
    STCG_A5a_FullConsideration3 = IIf(Len(Sheet13.Range("STCG.A5aic_FullConsideration50CA")) = 0, 0, Sheet13.Range("STCG.A5aic_FullConsideration50CA").value)
    STCG_A5a_FullConsideration4 = IIf(Len(Sheet13.Range("STCG.A5aii_FullConsideration")) = 0, 0, Sheet13.Range("STCG.A5aii_FullConsideration").value)

    STCG_A5bi_AquisitCost = IIf(Len(Sheet13.Range("STCG.A5bi_AquisitCost")) = 0, 0, Sheet13.Range("STCG.A5bi_AquisitCost").value)
    STCG_A5bii_ImproveCost = IIf(Len(Sheet13.Range("STCG.A5bii_ImproveCost")) = 0, 0, Sheet13.Range("STCG.A5bii_ImproveCost").value)
    STCG_A5biii_ExpOnTrans = IIf(Len(Sheet13.Range("STCG.A5biii_ExpOnTrans")) = 0, 0, Sheet13.Range("STCG.A5biii_ExpOnTrans").value)
    STCG_A5biv_TotalDedn = IIf(Len(Sheet13.Range("STCG.A5biv_TotalDedn")) = 0, 0, Sheet13.Range("STCG.A5biv_TotalDedn").value)
    STCG_A5c_BalanceCG = IIf(Len(Sheet13.Range("STCG.A5c_BalanceCG")) = 0, 0, Sheet13.Range("STCG.A5c_BalanceCG").value)
    STCG_A5d_LossSec94of7Or94of8 = IIf(Len(Sheet13.Range("STCG.A5d_LossSec94of7Or94of8")) = 0, 0, Sheet13.Range("STCG.A5d_LossSec94of7Or94of8").value)
    
    'STCG_A5ei_ExemptionAmount = IIf(Len(Sheet13.Range("STCG.A5ei_ExemptionAmount")) = 0, 0, Sheet13.Range("STCG.A5ei_ExemptionAmount").value)
    STCG_A5eii_ExemptionAmount = IIf(Len(Sheet13.Range("STCG.A5eii_ExemptionAmount")) = 0, 0, Sheet13.Range("STCG.A5eii_ExemptionAmount").value)
    STCG_A5eiii_ExemptionAmount = IIf(Len(Sheet13.Range("STCG.A5eiii_ExemptionAmount")) = 0, 0, Sheet13.Range("STCG.A5eiii_ExemptionAmount").value)
    STCG_A5e_ExemptionGrandTotal = IIf(Len(Sheet13.Range("STCG.A5e_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("STCG.A5e_ExemptionGrandTotal").value)
    
    STCG_A5e_CapgainonAssets = IIf(Len(Sheet13.Range("STCG.A5e_CapgainonAssets")) = 0, 0, Sheet13.Range("STCG.A5e_CapgainonAssets").value)
      
    STCG_A7_AmtDeemed = Sheet13.Range("STCG.A7_AmtDeemed").value
    
    STCG_A6_StatusUnutilizedCG = IIf((Sheet13.Range("STCG.A6_StatusUnutilizedCG") = "Not Applicable"), "X", Mid(Sheet13.Range("STCG.A6_StatusUnutilizedCG"), 1, 1))
    If STCG_A6_StatusUnutilizedCG = "Y" Then
       If Not validateSTCGValidate() Then Validate_ALL_CG = False
    End If
    
    LTCG_B9_StatusUnutilizedCG = IIf((Sheet13.Range("LTCG.A9_StatusUnutilizedCG") = "Not Applicable"), "X", Mid(Sheet13.Range("LTCG.A9_StatusUnutilizedCG"), 1, 1))
    If LTCG_B9_StatusUnutilizedCG = "Y" Then
       If Not validateLTCGvalidate() Then Validate_ALL_CG = False
    End If
    
    
   If Sheet13.Range("CG_FilledFlag").value <> 0 Then
   If isdropdownblank(Sheet13.Range("STCG.A6_StatusUnutilizedCG").value) Then
    msgboxSHEETCG = msgboxSHEETCG & "Please Select Whether any amount of unutilized capital gain on asset transferred during the previous years at pt 7 in STCG" & Chr(13)
    Validate_ALL_CG = False
    Else
    STCG_A6_StatusUnutilizedCG = IIf((Sheet13.Range("STCG.A6_StatusUnutilizedCG") = "Not Applicable") Or (Sheet13.Range("CG_FilledFlag") = 0), "X", Mid(Sheet13.Range("STCG.A6_StatusUnutilizedCG"), 1, 1))
    End If

   End If
   STCG_A8_StcgDTAA = Sheet13.Range("STCG.A8_StcgDTAA").value
   STCG_A8b_StcgDTAA = Sheet13.Range("STCG.A8b_StcgDTAA").value
   STCG_A8_Total = Sheet13.Range("STCGtOT").value
'   If (STCG_A8_StcgDTAA + STCG_A8b_StcgDTAA) <> STCG_A8_Total Then
'    msgboxSHEETCG = msgboxSHEETCG & "Total of [(a) Total amount of STCG not chargeable to tax in India as per DTAA] and [(b) Total amount of STCG chargeable at special rates in India as per DTAA] should be equal to sum of ""Amount of income (2)"" " & Chr(13)
'     Validate_ALL_CG = False
'    End If
   
    If Not ValidateSTCGDTAA() Then Validate_ALL_CG = False
     STCG_A6_AmountDeemedOth = IIf(Len(Sheet13.Range("STCG.A6_AmountDeemedOth")) = 0, 0, Sheet13.Range("STCG.A6_AmountDeemedOth").value)
    STCG_A8_DeemedStcgOnAssets = IIf(Len(Sheet13.Range("STCG.A8_DeemedStcgOnAssets")) = 0, 0, Sheet13.Range("STCG.A8_DeemedStcgOnAssets").value)
    STCG_A7_TotalSTCG = IIf(Len(Sheet13.Range("STCG.A7_TotalSTCG")) = 0, 0, Sheet13.Range("STCG.A7_TotalSTCG").value)
    
    'From Slump sale 2aiii_P Column
     LTCG_B02a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.B02a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B02a_FullConsideration").value)
     LTCG_B02b_NetWorthOfDivision = IIf(Len(Sheet13.Range("LTCG.B02b_NetWorthOfDivision")) = 0, 0, Sheet13.Range("LTCG.B02b_NetWorthOfDivision").value)
     LTCG_B02c_SlumpBalance = IIf(Len(Sheet13.Range("LTCG.B02c_SlumpBalance")) = 0, 0, Sheet13.Range("LTCG.B02c_SlumpBalance").value)
     LTCG_B02di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B02di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B02di_ExemptionAmount").value)
     LTCG_B02dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B02dii_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B02dii_ExemptionAmount").value)
     LTCG_B02d_ExemptionGrandTotal = IIf(Len(LTCG_B02di_ExemptionAmount) = 0 And Len(LTCG_B02dii_ExemptionAmount) = 0, 0, Sheet13.Range("LTCG.B02d_ExemptionGrandTotal"))
     
    'From Slump sale 2aiii_Q Column by sai for AY 2025-26
     LTCG_B02a_FullConsideration_ii = IIf(Len(Sheet13.Range("LTCG.B02a_FullConsideration_ii")) = 0, 0, Sheet13.Range("LTCG.B02a_FullConsideration_ii").value)
     LTCG_B02b_NetWorthOfDivision_ii = IIf(Len(Sheet13.Range("LTCG.B02b_NetWorthOfDivision_ii")) = 0, 0, Sheet13.Range("LTCG.B02b_NetWorthOfDivision_ii").value)
     LTCG_B02c_SlumpBalance_ii = IIf(Len(Sheet13.Range("LTCG.B02c_SlumpBalance_ii")) = 0, 0, Sheet13.Range("LTCG.B02c_SlumpBalance_ii").value)
     LTCG_B02di_ExemptionAmount_ii = IIf(Len(Sheet13.Range("LTCG.B02di_ExemptionAmount_ii")) = 0, 0, Sheet13.Range("LTCG.B02di_ExemptionAmount_ii").value)
     LTCG_B02dii_ExemptionAmount_ii = IIf(Len(Sheet13.Range("LTCG.B02dii_ExemptionAmount_ii")) = 0, 0, Sheet13.Range("LTCG.B02dii_ExemptionAmount_ii").value)
     LTCG_B02d_ExemptionGrandTotal_ii = IIf(Len(LTCG_B02di_ExemptionAmount_ii) = 0 And Len(LTCG_B02dii_ExemptionAmount_ii) = 0, 0, Sheet13.Range("LTCG.B02d_ExemptionGrandTotal_ii"))
    
'     LTCG_B02e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.B02e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.B02e_CapgainonAssets").value)
      LTCG_B02e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.B02e_CapgainonAssets_sum")) = 0, 0, Sheet13.Range("LTCG.B02e_CapgainonAssets_sum").value)
      
    
    
    
    'change-1 09.05.2022
    LTCG_B2a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.B2a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B2a_FullConsideration").value)
    LTCG_B2bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.B2bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.B2bi_AquisitCost").value)
    LTCG_B2bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.B2bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.B2bii_ImproveCost").value)
    LTCG_B2biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.B2biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.B2biii_ExpOnTrans").value)
    LTCG_B2biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.B2biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.B2biv_TotalDedn").value)
    LTCG_B2c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.B2c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.B2c_BalanceCG").value)
    LTCG_B2di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B2di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B2di_ExemptionAmount").value)
   
    LTCG_B2e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.B2e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.B2e_CapgainonAssets").value)
    LTCG_B3a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.B3a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B3a_FullConsideration").value)
    
'**************************************************************************************************************************
'by sai on 27/03/2025 as per DE V0.4
    LTCG_B3a_ii_FullConsideration = IIf(Len(Sheet13.Range("LTCG.B3a_FullConsideration_ii")) = 0, 0, Sheet13.Range("LTCG.B3a_FullConsideration_ii").value)
    LTCG_B3bia_AcqIndexation = IIf(Len(Sheet13.Range("LTCG.B3bia_AcqIndexation")) = 0, 0, Sheet13.Range("LTCG.B3bia_AcqIndexation").value)
    LTCG_B3bia_AcqIndexation_ii = IIf(Len(Sheet13.Range("LTCG.B3bia_AcqIndexation_ii")) = 0, 0, Sheet13.Range("LTCG.B3bia_AcqIndexation_ii").value)
    LTCG_B3biia_ImproveCost_WithoutIndexation = IIf(Len(Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation")) = 0, 0, Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation").value)
    LTCG_B3biia_ImproveCost_WithoutIndexation_ii = IIf(Len(Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation_ii")) = 0, 0, Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation_ii").value)
    LTCG_B3biii_ExpOnTrans_ii = IIf(Len(Sheet13.Range("LTCG.B3biii_ExpOnTrans_ii")) = 0, 0, Sheet13.Range("LTCG.B3biii_ExpOnTrans_ii").value)
    LTCG_B3biv_TotalDedn_ii = IIf(Len(Sheet13.Range("LTCG.B3biv_TotalDedn_ii")) = 0, 0, Sheet13.Range("LTCG.B3biv_TotalDedn_ii").value)
    LTCG_B3c_BalanceCG_ii = IIf(Len(Sheet13.Range("LTCG.B3c_BalanceCG_ii")) = 0, 0, Sheet13.Range("LTCG.B3c_BalanceCG_ii").value)
    LTCG_B04d_ExemptionGrandTotal_ii = IIf(Len(Sheet13.Range("LTCG.B04d_ExemptionGrandTotal_ii")) = 0, 0, Sheet13.Range("LTCG.B04d_ExemptionGrandTotal_ii").value)
    LTCG_B3e_CapgainonAssets_ii = IIf(Len(Sheet13.Range("LTCG.B3e_CapgainonAssets_ii")) = 0, 0, Sheet13.Range("LTCG.B3e_CapgainonAssets_ii").value)
    LTCG_B3ea_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.B3ea_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.B3ea_CapgainonAssets").value)
    LTCG_B3f_Tax20 = IIf(Len(Sheet13.Range("LTCG.B3f_Tax20")) = 0, 0, Sheet13.Range("LTCG.B3f_Tax20").value)
    LTCG_B3g_Tax10 = IIf(Len(Sheet13.Range("LTCG.B3g_Tax10")) = 0, 0, Sheet13.Range("LTCG.B3g_Tax10").value)
    LTCG_B3h_Excess = IIf(Len(Sheet13.Range("LTCG.B3h_Excess")) = 0, 0, Sheet13.Range("LTCG.B3h_Excess").value)
 '********************************************************************************************************************

    LTCG_B3bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.B3bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.B3bi_AquisitCost").value)
    LTCG_B3bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.B3bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.B3bii_ImproveCost").value)
    LTCG_B3biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.B3biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.B3biii_ExpOnTrans").value)
    LTCG_B3biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.B3biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.B3biv_TotalDedn").value)
    LTCG_B3c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.B3c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.B3c_BalanceCG").value)
    LTCG_B3di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B3di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B3di_ExemptionAmount").value)
    LTCG_B3dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B04d_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").value)
    'LTCG_B3d_ExemptionGrandTotal = IIf(Len(LTCG_B3di_ExemptionAmount) = 0 And Len(LTCG_B3dii_ExemptionAmount) = 0, 0, Sheet13.Range("LTCG.B3d_ExemptionGrandTotal"))
    '__________________________________________
    LTCG_B3e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.B3e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.B3e_CapgainonAssets").value)
    
    
    'Sale of GDR 115ACA 4iia_P Column
    LTCG_Bii3a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.Bii3a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.Bii3a_FullConsideration").value)
    LTCG_Bii3bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.Bii3bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.Bii3bi_AquisitCost").value)
    LTCG_Bii3bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.Bii3bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.Bii3bii_ImproveCost").value)
    LTCG_Bii3biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.Bii3biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.Bii3biii_ExpOnTrans").value)
    LTCG_Bii3biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.Bii3biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.Bii3biv_TotalDedn").value)
    LTCG_Bii3c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.Bii3c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.Bii3c_BalanceCG").value)
    LTCG_Bii3di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.Bii3di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.Bii3di_ExemptionAmount").value)
    LTCG_Bii3dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B04di_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value)
    'LTCG_Bii3d_ExemptionGrandTotal = IIf(Len(LTCG_Bii3di_ExemptionAmount) = 0 And Len(LTCG_Bii3dii_ExemptionAmount) = 0, 0, Sheet13.Range("LTCG.Bii3d_ExemptionGrandTotal"))
    
    'Sale of GDR 115ACA 4iia_Q Column by Sai for AY 2025-26*****************************************************************
    LTCG_Bii3a_FullConsideration_ii = IIf(Len(Sheet13.Range("LTCG.Bii3a_FullConsideration_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3a_FullConsideration_ii").value)
    LTCG_Bii3bi_AquisitCost_ii = IIf(Len(Sheet13.Range("LTCG.Bii3bi_AquisitCost_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3bi_AquisitCost_ii").value)
    LTCG_Bii3bii_ImproveCost_ii = IIf(Len(Sheet13.Range("LTCG.Bii3bii_ImproveCost_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3bii_ImproveCost_ii").value)
    LTCG_Bii3biii_ExpOnTrans_ii = IIf(Len(Sheet13.Range("LTCG.Bii3biii_ExpOnTrans_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3biii_ExpOnTrans_ii").value)
    LTCG_Bii3biv_TotalDedn_ii = IIf(Len(Sheet13.Range("LTCG.Bii3biv_TotalDedn_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3biv_TotalDedn_ii").value)
    LTCG_Bii3c_BalanceCG_ii = IIf(Len(Sheet13.Range("LTCG.Bii3c_BalanceCG_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3c_BalanceCG_ii").value)
    
'    LTCG_Bii3di_ExemptionAmount_ii = IIf(Len(Sheet13.Range("LTCG.Bii3di_ExemptionAmount_ii")) = 0, 0, Sheet13.Range("LTCG.Bii3di_ExemptionAmount_ii").value)
'    LTCG_Bii3dii_ExemptionAmount_ii = IIf(Len(Sheet13.Range("LTCG.B04di_ExemptionGrandTotal_ii")) = 0, 0, Sheet13.Range("LTCG.B04di_ExemptionGrandTotal_ii").value)
    
'    LTCG_Bii3e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.Bii3e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.Bii3e_CapgainonAssets").value)
     LTCG_Bii3e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.Bii3e_CapgainonAssets_sum")) = 0, 0, Sheet13.Range("LTCG.Bii3e_CapgainonAssets_sum").value)
    
  'end 4iia_Q Column********************************************************************************************************
    
    LTCG_B4a_LTCGWithoutBenefit = IIf(Len(Sheet13.Range("LTCG.B4a_LTCGWithoutBenefit")) = 0, 0, Sheet13.Range("LTCG.B4a_LTCGWithoutBenefit").value)
   ' LTCG_B4bi_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B4bi_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B4bi_ExemptionAmount").value)
    LTCG_B4bii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B4bii_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B4bii_ExemptionAmount").value)
    'LTCG_B4b_ExemptionGrandTotal = IIf(Len(LTCG_B4bi_ExemptionAmount) = 0 And Len(LTCG_B4bii_ExemptionAmount) = 0, 0, Sheet13.Range("LTCG.B4b_ExemptionGrandTotal").value)
    
    LTCG_B4c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.B4e_LTCGOnUnlistedSecurity")) = 0, 0, Sheet13.Range("LTCG.B4e_LTCGOnUnlistedSecurity").value)
    
    LTCG_Bi5a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.Bi5a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.Bi5a_FullConsideration").value)
    LTCG_Bi5a_FullConsideration1 = IIf(Len(Sheet13.Range("LTCG.B6aia_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B6aia_FullConsideration").value)
    LTCG_Bi5a_FullConsideration2 = IIf(Len(Sheet13.Range("LTCG.B6aib_FullMarketValue")) = 0, 0, Sheet13.Range("LTCG.B6aib_FullMarketValue").value)
    LTCG_Bi5a_FullConsideration3 = IIf(Len(Sheet13.Range("LTCG.B6aic_FullConsideration50CA")) = 0, 0, Sheet13.Range("LTCG.B6aic_FullConsideration50CA").value)
    LTCG_Bi5a_FullConsideration4 = IIf(Len(Sheet13.Range("LTCG.B6aii_FullConsiderationS")) = 0, 0, Sheet13.Range("LTCG.B6aii_FullConsiderationS").value)
    LTCG_Bi5bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.Bi5bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.Bi5bi_AquisitCost").value)
    LTCG_Bi5bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.Bi5bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.Bi5bii_ImproveCost").value)
    LTCG_Bi5biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.Bi5biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.Bi5biii_ExpOnTrans").value)
    LTCG_Bi5biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.Bi5biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.Bi5biv_TotalDedn").value)
    LTCG_Bi5c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.Bi5c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.Bi5c_BalanceCG").value)
    'LTCG_Bi5di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.Bi5di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.Bi5di_ExemptionAmount").value)
    LTCG_Bi5dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").value)
    'LTCG_Bi5d_ExemptionGrandTotal = Sheet13.Range("LTCG.Bi5d_ExemptionGrandTotal").value
    
    LTCG_Bi5e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.Bi5e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.Bi5e_CapgainonAssets").value)
    LTCG_Bii5a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.Bii5a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.Bii5a_FullConsideration").value)
    LTCG_Bii5a_FullConsideration1 = IIf(Len(Sheet13.Range("LTCG.B6ibia_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B6ibia_FullConsideration").value)
    LTCG_Bii5a_FullConsideration2 = IIf(Len(Sheet13.Range("LTCG.B6ibib_FullMarketValue")) = 0, 0, Sheet13.Range("LTCG.B6ibib_FullMarketValue").value)
    LTCG_Bii5a_FullConsideration3 = IIf(Len(Sheet13.Range("LTCG.B6ibic_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B6ibic_FullConsideration").value)
    LTCG_Bii5a_FullConsideration4 = IIf(Len(Sheet13.Range("LTCG.B6ibii_FullConsideration50CA")) = 0, 0, Sheet13.Range("LTCG.B6ibii_FullConsideration50CA").value)
    LTCG_Bii5bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.Bii5bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.Bii5bi_AquisitCost").value)
    LTCG_Bii5bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.Bii5bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.Bii5bii_ImproveCost").value)
    LTCG_Bii5biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.Bii5biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.Bii5biii_ExpOnTrans").value)
    LTCG_Bii5biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.Bii5biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.Bii5biv_TotalDedn").value)
    LTCG_Bii5c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.Bii5c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.Bii5c_BalanceCG").value)
    'LTCG_Bii5di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.Bii5di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.Bii5di_ExemptionAmount").value)
    LTCG_Bii5dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B07div_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value)
    'LTCG_Bii5d_ExemptionGrandTotal = Sheet13.Range("LTCG.Bii5d_ExemptionGrandTotal").value
    
    LTCG_Bii5e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.Bii5e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.Bii5e_CapgainonAssets").value)
    LTCG_Biii5a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.Biii5a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.Biii5a_FullConsideration").value)
    LTCG_Biii5a_FullConsideration1 = IIf(Len(Sheet13.Range("LTCG.B6iiiaia_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B6iiiaia_FullConsideration").value)
    LTCG_Biii5a_FullConsideration2 = IIf(Len(Sheet13.Range("LTCG.B6iiiaib_FullMarketValue")) = 0, 0, Sheet13.Range("LTCG.B6iiiaib_FullMarketValue").value)
    LTCG_Biii5a_FullConsideration3 = IIf(Len(Sheet13.Range("LTCG.B6iiiaic_FullConsideration50CA")) = 0, 0, Sheet13.Range("LTCG.B6iiiaic_FullConsideration50CA").value)
    LTCG_Biii5a_FullConsideration4 = IIf(Len(Sheet13.Range("LTCG.B6iiiaii_FullConsiderationS")) = 0, 0, Sheet13.Range("LTCG.B6iiiaii_FullConsiderationS").value)
    LTCG_Biii5bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.Biii5bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.Biii5bi_AquisitCost").value)
    LTCG_Biii5bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.Biii5bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.Biii5bii_ImproveCost").value)
    LTCG_Biii5biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.Biii5biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.Biii5biii_ExpOnTrans").value)
    LTCG_Biii5biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.Biii5biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.Biii5biv_TotalDedn").value)
    LTCG_Biii5c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.Biii5c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.Biii5c_BalanceCG").value)
   ' LTCG_Biii5di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.Biii5di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.Biii5di_ExemptionAmount").value)
    LTCG_Biii5dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").value)
    'LTCG_Biii5d_ExemptionGrandTotal = Sheet13.Range("LTCG.Biii5d_ExemptionGrandTotal").value
    
    LTCG_Biii5e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.Biii5e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.Biii5e_CapgainonAssets").value)
    LTCG_B6a_SaleonSpecAsset = IIf(Len(Sheet13.Range("LTCG.B6a_SaleonSpecAsset")) = 0, 0, Sheet13.Range("LTCG.B6a_SaleonSpecAsset").value)
    LTCG_B6b_DednSpecAssetus115 = IIf(Len(Sheet13.Range("LTCG.B6b_DednSpecAssetus115")) = 0, 0, Sheet13.Range("LTCG.B6b_DednSpecAssetus115").value)
    LTCG_B6c_BalonSpeciAsset = IIf(Len(Sheet13.Range("LTCG.B6c_BalonSpeciAsset")) = 0, 0, Sheet13.Range("LTCG.B6c_BalonSpeciAsset").value)
    LTCG_B6d_SaleOtherSpecAsset = IIf(Len(Sheet13.Range("LTCG.B6d_SaleOtherSpecAsset")) = 0, 0, Sheet13.Range("LTCG.B6d_SaleOtherSpecAsset").value)
    
    'Commented by Sai on 03/03/2025 for AY 2025-26 as it was removed
'    LTCG_B6e_DednOtherSpecAssetus115 = IIf(Len(Sheet13.Range("LTCG.B6e_DednOtherSpecAssetus115")) = 0, 0, Sheet13.Range("LTCG.B6e_DednOtherSpecAssetus115").value)
'    LTCG_B6f_BalOtherthanSpecAsset = IIf(Len(Sheet13.Range("LTCG.B6f_BalOtherthanSpecAsset")) = 0, 0, Sheet13.Range("LTCG.B6f_BalOtherthanSpecAsset").value)
     
    
    LTCG_B7a_FullConsideration = IIf(Len(Sheet13.Range("LTCG.B7a_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B7a_FullConsideration").value)
    LTCG_B7a_FullConsideration1 = IIf(Len(Sheet13.Range("LTCG.B8aia_FullConsideration")) = 0, 0, Sheet13.Range("LTCG.B8aia_FullConsideration").value)
    LTCG_B7a_FullConsideration2 = IIf(Len(Sheet13.Range("LTCG.B8aib_FullMarketValue")) = 0, 0, Sheet13.Range("LTCG.B8aib_FullMarketValue").value)
    LTCG_B7a_FullConsideration3 = IIf(Len(Sheet13.Range("LTCG.B8aic_FullConsideration50CA")) = 0, 0, Sheet13.Range("LTCG.B8aic_FullConsideration50CA").value)
    LTCG_B7a_FullConsideration4 = IIf(Len(Sheet13.Range("LTCG.B8aii_FullConsiderationO")) = 0, 0, Sheet13.Range("LTCG.B8aii_FullConsiderationO").value)
    LTCG_B7bi_AquisitCost = IIf(Len(Sheet13.Range("LTCG.B7bi_AquisitCost")) = 0, 0, Sheet13.Range("LTCG.B7bi_AquisitCost").value)
    LTCG_B7bii_ImproveCost = IIf(Len(Sheet13.Range("LTCG.B7bii_ImproveCost")) = 0, 0, Sheet13.Range("LTCG.B7bii_ImproveCost").value)
    LTCG_B7biii_ExpOnTrans = IIf(Len(Sheet13.Range("LTCG.B7biii_ExpOnTrans")) = 0, 0, Sheet13.Range("LTCG.B7biii_ExpOnTrans").value)
    LTCG_B7biv_TotalDedn = IIf(Len(Sheet13.Range("LTCG.B7biv_TotalDedn")) = 0, 0, Sheet13.Range("LTCG.B7biv_TotalDedn").value)
    LTCG_B7c_BalanceCG = IIf(Len(Sheet13.Range("LTCG.B7c_BalanceCG")) = 0, 0, Sheet13.Range("LTCG.B7c_BalanceCG").value)
    
    LTCG_B7di_ExemptionAmount0 = IIf(Len(Sheet13.Range("LTCG.B7di_ExemptionAmount0")) = 0, 0, Sheet13.Range("LTCG.B7di_ExemptionAmount0").value)
    LTCG_B7di_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B7di_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B7di_ExemptionAmount").value)
    LTCG_B7dii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B7dii_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B7dii_ExemptionAmount").value)
    LTCG_B7diii_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B7diii_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B7diii_ExemptionAmount").value)
    LTCG_B7div_ExemptionAmount = IIf(Len(Sheet13.Range("LTCG.B7div_ExemptionAmount")) = 0, 0, Sheet13.Range("LTCG.B7div_ExemptionAmount").value)
       
    LTCG_B7d_ExemptionGrandTotal = IIf(Len(Sheet13.Range("LTCG.B7d_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B7d_ExemptionGrandTotal").value)
    
    LTCG_B7e_CapgainonAssets = IIf(Len(Sheet13.Range("LTCG.B7e_CapgainonAssets")) = 0, 0, Sheet13.Range("LTCG.B7e_CapgainonAssets").value)

    LTCG_B9_AmtDeemed = Sheet13.Range("LTCG.B9_AmtDeemed").value
    
    If Sheet13.Range("CG_FilledFlag").value <> 0 Then
    If isdropdownblank(Sheet13.Range("LTCG.A9_StatusUnutilizedCG").value) Then
    msgboxSHEETCG = msgboxSHEETCG & "Please Select Whether any amount of unutilized capital gain on asset transferred during the previous years at pt 11 in LTCG" & Chr(13)
    Validate_ALL_CG = False
    Else
    LTCG_B9_StatusUnutilizedCG = IIf((Sheet13.Range("LTCG.A9_StatusUnutilizedCG") = "Not Applicable") Or (Sheet13.Range("CG_FilledFlag") = 0), "X", Mid(Sheet13.Range("LTCG.A9_StatusUnutilizedCG"), 1, 1))
    End If

    End If
    
    LTCG_B10_ltcgDTAA = Sheet13.Range("LTCG.B10_ltcgDTAA").value
    LTCG_B10b_ltcgDTAA = Sheet13.Range("LTCG.B10b_ltcgDTAA").value
    LTCG_Total = Sheet13.Range("LTCGTot").value
      
    
'For Auto populated sum
'STCG
  'For 111A after 23rd
  
    If Len(Sheet13.Range("STCG.A2ibiv_2eii_TotalDedn").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 3 (ibiv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
 
    If Len(Sheet13.Range("STCG.A2ic_2eii_BalanceCG").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 3 (ic) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 

    If Len(Sheet13.Range("STCG.A2ie_CapgainonAssets_i_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in STCG Table 2 (A2ie) " & Chr(13)
        Validate_ALL_CG = False
    End If
    
'for 115AD after 23rd
    If Len(Sheet13.Range("STCG.A2iibiv_TotalDedn_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 3 (iibiv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
 
    If Len(Sheet13.Range("STCG.A2iic_BalanceCG_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table        3 (iic) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

    If Len(Sheet13.Range("STCG.A2iie_CapgainonAssets_i_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in STCG Table 3 (A2iie) " & Chr(13)
        Validate_ALL_CG = False
    End If
 
    
'    If Len(Sheet13.Range("STCG.A3a_NRItaxSTTPaid_i_ii").value) > 14 Then  'Commenting as per V0.01by Bindu on12th Feb 26
'        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in STCG Table 4 (A4a) " & Chr(13)
'        Validate_ALL_CG = False
'    End If
    
    
    
    
    
'For SIT-94390,SIT-94389 by sai on 13/06/2025
'Below fields will always be zero by Bindu 16th Feb 26
  If Sheet13.Range("LTCG.B3bii_ImproveCost").value < Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation").value Then
         msgboxSHEETCG = msgboxSHEETCG & "* ""4bii cant be less than 4biia"" Where transfer was before 23rd July 2024" & Chr(13)
         Validate_ALL_CG = False
  End If
  
  'Below fields LTCG.B3bi_AquisitCost/LTCG.B3bia_AcqIndexation will alyaws be zero by Bindu on 16th Feb 26
  If Sheet13.Range("LTCG.B3bi_AquisitCost").value < Sheet13.Range("LTCG.B3bia_AcqIndexation").value Then
          msgboxSHEETCG = msgboxSHEETCG & "* ""4bi cant be less than 4bia"" Where transfer was before 23rd July 2024" & Chr(13)
          Validate_ALL_CG = False
  End If
  '===================================================================
    
    
'For LTCG
'==============================================================================================================================
'If Sheet13.Range("LTCG.B02b_NetWorthOfDivision").value < 0 Then
'        If Len(Sheet13.Range("LTCG.B02b_NetWorthOfDivision").value) > 15 Then
'            msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG For Slump Sale Where transfer was before 23rd July 2024""" & Chr(13)
'            Validate_ALL_CG = False
'        End If
'    Else
'        If Len(Sheet13.Range("LTCG.B02b_NetWorthOfDivision").value) > 14 Then
'            msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG For Slump Sale Where transfer was before 23rd July 2024""" & Chr(13)
'            Validate_ALL_CG = False
'        End If
'End If
'
'
'If Sheet13.Range("LTCG.B02b_NetWorthOfDivision_ii").value < 0 Then
'        If Len(Sheet13.Range("LTCG.B02b_NetWorthOfDivision_ii").value) > 15 Then
'            msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG For Slump Sale Where transfer was on or after 23rd July 2024""" & Chr(13)
'            Validate_ALL_CG = False
'        End If
'    Else
'        If Len(Sheet13.Range("LTCG.B02b_NetWorthOfDivision_ii").value) > 14 Then
'            msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG For Slump Sale Where transfer was on or after 23rd July 2024""" & Chr(13)
'            Validate_ALL_CG = False
'        End If
'End If
'
'=====================================================================================================================
    
    
     If Len(Sheet13.Range("LTCG.Bbia_ImproveCost_Sum1").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 1 biib(a) " & Chr(13)
        Validate_ALL_CG = False
    End If
    If Len(Sheet13.Range("LTCG.Bbii_ImproveCost_Sum1").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 1 biib(c)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
'Table 2 Slump sale
    If Len(Sheet13.Range("LTCG.B02d_ExemptionGrandTotal").value) > 14 Then 'Not updating anything as it's always be zero By Bindu on 12th Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 2 (2d) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

    If Len(Sheet13.Range("LTCG.B02d_ExemptionGrandTotal_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 2 (2d) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
    
    If Len(Sheet13.Range("LTCG.B02e_CapgainonAssets_sum").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 2 (B2e) " & Chr(13)
        Validate_ALL_CG = False
    End If
    
'For Table 4i 112(1)
 
 'Before 23rd
    If Len(Sheet13.Range("LTCG.B3biv_TotalDedn").value) > 14 Then 'Not updating by Bindu as its alwys be zero on 16th Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 4i (ibiv) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
  
    If Len(Sheet13.Range("LTCG.B3biva_Total").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG       Table 4i (ibiva) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
 
    If Len(Sheet13.Range("LTCG.B3c_BalanceCG").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG       Table 4i (ic) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
 
    If Len(Sheet13.Range("LTCG.B3ca_BalanceCG").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG       Table 4i (ica) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
 'After 23rd
   
    If Len(Sheet13.Range("LTCG.B3biv_TotalDedn_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 4i (ibiv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
 
 
    If Len(Sheet13.Range("LTCG.B3c_BalanceCG_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG       Table 4i (ic) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    

    
 'For 4ii table 115ACA
 
   'Before 23rd
   
    If Len(Sheet13.Range("LTCG.Bii3biv_TotalDedn").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 4ii (iibiv) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
   
    If Len(Sheet13.Range("LTCG.Bii3c_BalanceCG").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 4ii (iic) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    'After 23rd
    
    If Len(Sheet13.Range("LTCG.Bii3biv_TotalDedn_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 4ii (iibiv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.Bii3c_BalanceCG_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 4ii (iic) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    
    If Len(Sheet13.Range("LTCG.Bii3e_CapgainonAssets_sum").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 4ii (B4iie) " & Chr(13)
        Validate_ALL_CG = False
    End If
    
 'For Table 5
 
    If Len(Sheet13.Range("LTCG.5.ci_LTCG_OnSaleCapitalAssets").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 5 (B5c) " & Chr(13)
        Validate_ALL_CG = False
    End If
    
'For Table 6

    If Len(Sheet13.Range("LTCG.B4a_LTCGWithoutBenefit").value) > 14 Then 'Not using this field "LTCG.B4a_LTCGWithoutBenefit" by Bindu on 22nd Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 6 (6a) " & Chr(13)
        Validate_ALL_CG = False
    End If


'    If Len(Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value) > 14 Then
'        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 6 (6b) " & Chr(13)
'        Validate_ALL_CG = False
'    End If

'by sai on 26/06/2025
    If Len(Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value) > 9 Then 'Not using this field "LTCG.B06d_ExemptionGrandTotal" by Bindu on 22nd Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,9999in LTCG Table 6 (6b) " & Chr(13)
        Validate_ALL_CG = False
    End If



    If Len(Sheet13.Range("LTCG.B4ei_LTCGOnUnlistedSecurity").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 6 (B6c) " & Chr(13)
        Validate_ALL_CG = False
    End If
    
'For Table 7 112(1)

  'before 23rd
    
    If Len(Sheet13.Range("LTCG.Bi5biv_TotalDedn").value) > 14 Then 'Not updating as this field will always be Zero by Bindu on 18th Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 7 (7biv) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If


    If Len(Sheet13.Range("LTCG.Bi5c_BalanceCG").value) > 14 Then 'Not updating as this field will always be Zero by Bindu on 18th Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (7ic) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

 'After 23rd
   
    If Len(Sheet13.Range("LTCG.Bi5biv_TotalDedn_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 7 (7biv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

    If Len(Sheet13.Range("LTCG.Bi5c_BalanceCG_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (7ic) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.Bi5e_CapgainonAssets_sum").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (B7ie)" & Chr(13)
        Validate_ALL_CG = False
    End If

 '115AC
     
       'before 23rd
    
    If Len(Sheet13.Range("LTCG.Bii5biv_TotalDedn").value) > 14 Then 'Not updating as this field "LTCG.Bii5biv_TotalDedn" will always be zero by Bindu on 18th feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (7biv) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If


    If Len(Sheet13.Range("LTCG.Bii5c_BalanceCG").value) > 14 Then 'Not updating as this field "LTCG.Bii5c_BalanceCG" will always be zero by Bindu on 18th feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (7iic) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

 'After 23rd
   
    If Len(Sheet13.Range("LTCG.Bii5biv_TotalDedn_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (7biv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

    If Len(Sheet13.Range("LTCG.Bii5c_BalanceCG_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (7iic) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.Bii5e_CapgainonAssets_sum").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 7 (B7iie)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
 'For table 8
 
    If Len(Sheet13.Range("LTCG.5.ci_LTCG_AssetsB8Above").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 8 (B8c)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
  'For table 9
  
    If Len(Sheet13.Range("LTCG.B6a_SaleonSpecAsset").value) > 14 Then 'Not updated By bindu as this field refers to this"LTCG.B6aii_SaleonSpecAsset" for AY 26-27 by Bindu on18th Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 9 (9a)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.B6b_DednSpecAssetus115").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 9 (9b)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
 'For SIT-91119********************************************************************
    If Len(Sheet13.Range("LTCG.B6ci_BalonSpeciAsset").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 9 (9c)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.B6c_BalonSpeciAsset").value) > 14 Then 'NOt updating anything as this field will be alwayyz zero by Bindu on 22nd Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 9 (9ci)" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.B6cii_BalonSpeciAsset").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 9 (9cii)" & Chr(13)
        Validate_ALL_CG = False
    End If
  '*************************************************************************************
    
'    If Len(Sheet13.Range("LTCG.B6d_SaleOtherSpecAsset").value) > 14 Then
'        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 9 (9d)" & Chr(13)
'        Validate_ALL_CG = False
'    End If
    
'For Table 10

    'before 23rd
    
    If Len(Sheet13.Range("LTCG.B7biv_TotalDedn").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 10 b (biv) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If


    If Len(Sheet13.Range("LTCG.B7c_BalanceCG").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 10 b (10c) ""Where transfer was before 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

 'After 23rd
   
    If Len(Sheet13.Range("LTCG.B7biv_TotalDedn_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 10 b (biv) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If

    If Len(Sheet13.Range("LTCG.B7c_BalanceCG_ii").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 10 b (10c) ""Where transfer was on or after 23rd July 2024" & Chr(13)
        Validate_ALL_CG = False
    End If
    
    If Len(Sheet13.Range("LTCG.B7e_CapgainonAssets_sum").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 10 (B710e)" & Chr(13)
        Validate_ALL_CG = False
    End If

'For table 11

    If Len(Sheet13.Range("LTCG.B9_AmountDeemedOth").value) > 14 Then 'NOt updating anything as this field will be alwayyz zero by Bindu on 22nd Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 11 (b)" & Chr(13)
        Validate_ALL_CG = False
    End If
  
    
    If Len(Sheet13.Range("LTCG.B9i_AmountDeemedOth").value) > 14 Then 'NOt updating anything as this field will be alwayyz zero by Bindu on 22nd Feb 26
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 11 (bi)" & Chr(13)
        Validate_ALL_CG = False
    End If
  
    If Len(Sheet13.Range("LTCG.B9ii_AmountDeemedOth").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 11 (bii)" & Chr(13)
        Validate_ALL_CG = False
    End If
  
  
    If Len(Sheet13.Range("LTCG.B9i_AmtDeemed").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 11 (b) B11" & Chr(13)
        Validate_ALL_CG = False
    End If
  
    If Len(Sheet13.Range("LTCG.B12_PassThroughIncome").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,non-negative, no decimal, upto 99,999,999,999,999 in LTCG Table 11 (B12)" & Chr(13)
        Validate_ALL_CG = False
    End If

'For Table 12
  
    If Len(Sheet13.Range("LTCG.B12ii_PassThroughIncome10_112A").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 12 a1(ii)" & Chr(13)
        Validate_ALL_CG = False
    End If


    If Len(Sheet13.Range("LTCG.B12ii_PassThroughIncome10").value) > 14 Then
        msgboxSHEETCG = msgboxSHEETCG & """Cell value can be Numeric,negative, no decimal, upto -99,999,999,999,999 in LTCG Table 12 a2(ii)" & Chr(13)
        Validate_ALL_CG = False
    End If


'End by Sai**********************************************************************************************************************


'    If (LTCG_B10_ltcgDTAA + LTCG_B10b_ltcgDTAA) <> LTCG_Total Then
'    msgboxSHEETCG = msgboxSHEETCG & "Total of [(a) Total amount of LTCG not chargeable to tax in India as per DTAA] and [(b) Total amount of LTCG chargeable at special rates in India as per DTAA] should be equal to sum of ""Amount of income (2)"" " & Chr(13)
'    Validate_ALL_CG = False
'    End If
    
    If Not validateLTCGDTAA() Then Validate_ALL_CG = False
    
    LTCG_B9_AmountDeemedOth = IIf(Len(Sheet13.Range("LTCG.B9_AmountDeemedOth")) = 0, 0, Sheet13.Range("LTCG.B9_AmountDeemedOth").value)
    'LTCG_B8_ExemptionGrandTotal = IIf(Len(Sheet13.Range("LTCG.B8_ExemptionGrandTotal")) = 0, 0, Sheet13.Range("LTCG.B8_ExemptionGrandTotal").value)
' Konda uncommented on 16-06-2025 SIT-94532
    LTCG_B9_TotalLTCG = IIf(Len(Sheet13.Range("LTCG.B9_TotalLTCG")) = 0, 0, Sheet13.Range("LTCG.B9_TotalLTCG").value)
'-----------------on 16-06-2025 SIT-94532
    'CG_C_TotScheduleCGFor23 = IIf(Len(Sheet13.Range("CG.C_TotScheduleCGFor23")) = 0, 0, Sheet13.Range("CG.C_TotScheduleCGFor23").value)
    
    CG_C_SumOfCGIncm = IIf(Len(Sheet13.Range("CG.C_TotScheduleCGFor23")) = 0, 0, Sheet13.Range("CG.C_TotScheduleCGFor23").value)
    CG_C_IncmFromVDATrnsf = IIf(Len(Sheet13.Range("CG.IncomeVDA")) = 0, 0, Sheet13.Range("CG.IncomeVDA").value)
    CG_C_TotScheduleCGFor23 = IIf(Len(Sheet13.Range("CG.CapitalGain")) = 0, 0, Sheet13.Range("CG.CapitalGain").value)
   
    
    IHLA_Ei2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Ei2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Ei2_StclSetoff15Per").value)
    IHLA_Ei3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Ei3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Ei3_StclSetoff30Per").value)
    IHLA_Ei4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Ei4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Ei4_StclSetoffAppRate").value)
    IHLA_Ei5_LtclSetOff10Per = IIf(Len(Sheet13.Range("IHLA.Ei5_LtclSetOff10Per")) = 0, 0, Sheet13.Range("IHLA.Ei5_LtclSetOff10Per").value)
    IHLA_Ei6_LtclSetOff20Per = IIf(Len(Sheet13.Range("IHLA.Ei6_LtclSetOff20Per")) = 0, 0, Sheet13.Range("IHLA.Ei6_LtclSetOff20Per").value)
    IHLA_Eii1_CurrYearIncome = IIf(Len(Sheet13.Range("IHLA.Eii1_CurrYearIncome")) = 0, 0, Sheet13.Range("IHLA.Eii1_CurrYearIncome").value)
    IHLA_Eii3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Eii3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Eii3_StclSetoff30Per").value)
    IHLA_Eii4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Eii4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").value)
    IHLA_Eii7_CurrYrCapGain = IIf(Len(Sheet13.Range("IHLA.Eii7_CurrYrCapGain")) = 0, 0, Sheet13.Range("IHLA.Eii7_CurrYrCapGain").value)
    IHLA_Eiii1_CurrYearIncome = IIf(Len(Sheet13.Range("IHLA.Eiii1_CurrYearIncome")) = 0, 0, Sheet13.Range("IHLA.Eiii1_CurrYearIncome").value)
    IHLA_Eiii2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Eiii2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").value)
    IHLA_Eiii4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").value)
    IHLA_Eiii7_CurrYrCapGain = IIf(Len(Sheet13.Range("IHLA.Eiii7_CurrYrCapGain")) = 0, 0, Sheet13.Range("IHLA.Eiii7_CurrYrCapGain").value)
    IHLA_Eiv1_CurrYearIncome = IIf(Len(Sheet13.Range("IHLA.Eiv1_CurrYearIncome")) = 0, 0, Sheet13.Range("IHLA.Eiv1_CurrYearIncome").value)
    IHLA_Eiv2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Eiv2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").value)
    IHLA_Eiv3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Eiv3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").value)
    IHLA_Eiv7_CurrYrCapGain = IIf(Len(Sheet13.Range("IHLA.Eiv7_CurrYrCapGain")) = 0, 0, Sheet13.Range("IHLA.Eiv7_CurrYrCapGain").value)
    IHLA_Ev1_CurrYearIncome = IIf(Len(Sheet13.Range("IHLA.Ev1_CurrYearIncome")) = 0, 0, Sheet13.Range("IHLA.Ev1_CurrYearIncome").value)
    IHLA_Ev2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Ev2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Ev2_StclSetoff15Per").value)
    IHLA_Ev3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Ev3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value)
    IHLA_Ev4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Ev4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").value)
    IHLA_Ev6_LtclSetOff20Per = IIf(Len(Sheet13.Range("IHLA.Ev6_LtclSetOff20Per")) = 0, 0, Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value)
    IHLA_Ev7_CurrYrCapGain = IIf(Len(Sheet13.Range("IHLA.Ev7_CurrYrCapGain")) = 0, 0, Sheet13.Range("IHLA.Ev7_CurrYrCapGain").value)
    IHLA_Evi1_CurrYearIncome = IIf(Len(Sheet13.Range("IHLA.Evi1_CurrYearIncome")) = 0, 0, Sheet13.Range("IHLA.Evi1_CurrYearIncome").value)
    IHLA_Evi2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Evi2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Evi2_StclSetoff15Per").value)
    IHLA_Evi3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Evi3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value)
    IHLA_Evi4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Evi4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").value)
    IHLA_Evi5_LtclSetOff10Per = IIf(Len(Sheet13.Range("IHLA.Evi5_LtclSetOff10Per")) = 0, 0, Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value)
    IHLA_Evi7_CurrYrCapGain = IIf(Len(Sheet13.Range("IHLA.Evi7_CurrYrCapGain")) = 0, 0, Sheet13.Range("IHLA.Evi7_CurrYrCapGain").value)
    IHLA_Evii2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Evii2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Evii2_StclSetoff15Per").value)
    IHLA_Evii3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Evii3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Evii3_StclSetoff30Per").value)
    IHLA_Evii4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Evii4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Evii4_StclSetoffAppRate").value)
    IHLA_Evii5_LtclSetOff10Per = IIf(Len(Sheet13.Range("IHLA.Evii5_LtclSetOff10Per")) = 0, 0, Sheet13.Range("IHLA.Evii5_LtclSetOff10Per").value)
    IHLA_Evii6_LtclSetOff20Per = IIf(Len(Sheet13.Range("IHLA.Evii6_LtclSetOff20Per")) = 0, 0, Sheet13.Range("IHLA.Evii6_LtclSetOff20Per").value)
    IHLA_Eviii2_StclSetoff15Per = IIf(Len(Sheet13.Range("IHLA.Eviii2_StclSetoff15Per")) = 0, 0, Sheet13.Range("IHLA.Eviii2_StclSetoff15Per").value)
    IHLA_Eviii3_StclSetoff30Per = IIf(Len(Sheet13.Range("IHLA.Eviii3_StclSetoff30Per")) = 0, 0, Sheet13.Range("IHLA.Eviii3_StclSetoff30Per").value)
    IHLA_Eviii4_StclSetoffAppRate = IIf(Len(Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate")) = 0, 0, Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate").value)
    IHLA_Eviii5_LtclSetOff10Per = IIf(Len(Sheet13.Range("IHLA.Eviii5_LtclSetOff10Per")) = 0, 0, Sheet13.Range("IHLA.Eviii5_LtclSetOff10Per").value)
    IHLA_Eviii6_LtclSetOff20Per = IIf(Len(Sheet13.Range("IHLA.Eviii6_LtclSetOff20Per")) = 0, 0, Sheet13.Range("IHLA.Eviii6_LtclSetOff20Per").value)
    
    AccSTCG_Upto15Of6 = IIf(Len(Sheet13.Range("AccSTCG.Upto15Of6")) = 0, 0, Sheet13.Range("AccSTCG.Upto15Of6").value)
    AccSTCG_Upto15Of9 = IIf(Len(Sheet13.Range("AccSTCG.Upto15Of9")) = 0, 0, Sheet13.Range("AccSTCG.Upto15Of9").value)
    ACCSTCG_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("ACCSTCG.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("ACCSTCG.Up16Of9To15Of12").value)
    ACCSTCG_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("ACCSTCG.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("ACCSTCG.Up16Of12To15Of3").value)
    ACCSTCG_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("ACCSTCG.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("ACCSTCG.Up16Of3To31Of3").value)
    
    '********************************************************** newly added by sai on 27/02/2025
    ACCLTCGDTAA_Upto15Of6_20 = IIf(Len(Sheet13.Range("AccSTCG20.Upto15Of6")) = 0, 0, Sheet13.Range("AccSTCG20.Upto15Of6").value)
    ACCLTCGDTAA_Upto15Of9_20 = IIf(Len(Sheet13.Range("AccSTCG20.Upto15Of9")) = 0, 0, Sheet13.Range("AccSTCG20.Upto15Of9").value)
    ACCLTCGDTAA_Up16Of9To15Of12_20 = IIf(Len(Sheet13.Range("ACCSTCG20.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("ACCSTCG20.Up16Of9To15Of12").value)
    ACCLTCGDTAA_Up16Of12To15Of3_20 = IIf(Len(Sheet13.Range("ACCSTCG20.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("ACCSTCG20.Up16Of12To15Of3").value)
    ACCLTCGDTAA_Up16Of3To31Of3_20 = IIf(Len(Sheet13.Range("ACCSTCG20.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("ACCSTCG20.Up16Of3To31Of3").value)
    
    ACCLTCGDTAA_Upto15Of6_125 = IIf(Len(Sheet13.Range("AccLTCG12.5.Upto15Of6")) = 0, 0, Sheet13.Range("AccLTCG12.5.Upto15Of6").value)
    ACCLTCGDTAA_Upto15Of9_125 = IIf(Len(Sheet13.Range("AccLTCG12.5.Upto15Of9")) = 0, 0, Sheet13.Range("AccLTCG12.5.Upto15Of9").value)
    ACCLTCGDTAA_Up16Of9To15Of12_125 = IIf(Len(Sheet13.Range("AccLTCG12.5.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("AccLTCG12.5.Up16Of9To15Of12").value)
    ACCLTCGDTAA_Up16Of12To15Of3_125 = IIf(Len(Sheet13.Range("AccLTCG12.5.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("AccLTCG12.5.Up16Of12To15Of3").value)
    ACCLTCGDTAA_Up16Of3To31Of3_125 = IIf(Len(Sheet13.Range("AccLTCG12.5.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("AccLTCG12.5.Up16Of3To31Of3").value)
    '***********************************************************************************************************************
    
    
    
    
    
    
    
    AccSTCG30_Upto15Of6 = IIf(Len(Sheet13.Range("AccSTCG30.Upto15Of6")) = 0, 0, Sheet13.Range("AccSTCG30.Upto15Of6").value)
    AccSTCG30_Upto15Of9 = IIf(Len(Sheet13.Range("AccSTCG30.Upto15Of9")) = 0, 0, Sheet13.Range("AccSTCG30.Upto15Of9").value)
    AccSTCG30_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("AccSTCG30.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("AccSTCG30.Up16Of9To15Of12").value)
    AccSTCG30_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("AccSTCG30.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("AccSTCG30.Up16Of12To15Of3").value)
    AccSTCG30_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("AccSTCG30.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("AccSTCG30.Up16Of3To31Of3").value)
    ACCSTCGOTH_Upto15Of6 = IIf(Len(Sheet13.Range("ACCSTCGOTH.Upto15Of6")) = 0, 0, Sheet13.Range("ACCSTCGOTH.Upto15Of6").value)
    ACCSTCGOTH_Upto15Of9 = IIf(Len(Sheet13.Range("ACCSTCGOTH.Upto15Of9")) = 0, 0, Sheet13.Range("ACCSTCGOTH.Upto15Of9").value)
    ACCSTCGOTH_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("ACCSTCGOTH.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("ACCSTCGOTH.Up16Of9To15Of12").value)
    ACCSTCGOTH_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("ACCSTCGOTH.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("ACCSTCGOTH.Up16Of12To15Of3").value)
    ACCSTCGOTH_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("ACCSTCGOTH.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("ACCSTCGOTH.Up16Of3To31Of3").value)
    ACCLTCG_Upto15Of6 = IIf(Len(Sheet13.Range("ACCLTCG.Upto15Of6")) = 0, 0, Sheet13.Range("ACCLTCG.Upto15Of6").value)
    ACCLTCG_Upto15Of9 = IIf(Len(Sheet13.Range("ACCLTCG.Upto15Of9")) = 0, 0, Sheet13.Range("ACCLTCG.Upto15Of9").value)
    ACCLTCG_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("ACCLTCG.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("ACCLTCG.Up16Of9To15Of12").value)
    ACCLTCG_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("ACCLTCG.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("ACCLTCG.Up16Of12To15Of3").value)
    ACCLTCG_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("ACCLTCG.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("ACCLTCG.Up16Of3To31Of3").value)
    ACCLTCGNP_Upto15Of6 = IIf(Len(Sheet13.Range("ACCLTCGNP.Upto15Of6")) = 0, 0, Sheet13.Range("ACCLTCGNP.Upto15Of6").value)
    ACCLTCGNP_Upto15Of9 = IIf(Len(Sheet13.Range("ACCLTCGNP.Upto15Of9")) = 0, 0, Sheet13.Range("ACCLTCGNP.Upto15Of9").value)
    ACCLTCGNP_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("ACCLTCGNP.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("ACCLTCGNP.Up16Of9To15Of12").value)
    ACCLTCGNP_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("ACCLTCGNP.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("ACCLTCGNP.Up16Of12To15Of3").value)
    ACCLTCGNP_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("ACCLTCGNP.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("ACCLTCGNP.Up16Of3To31Of3").value)

    ACCLTCGDTAA_Upto15Of6 = IIf(Len(Sheet13.Range("AccLTCGDTAA.Upto15Of6")) = 0, 0, Sheet13.Range("AccLTCGDTAA.Upto15Of6").value)
    ACCLTCGDTAA_Upto15Of9 = IIf(Len(Sheet13.Range("AccLTCGDTAA.Upto15Of9")) = 0, 0, Sheet13.Range("AccLTCGDTAA.Upto15Of9").value)
    ACCLTCGDTAA_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("AccLTCGDTAA.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("AccLTCGDTAA.Up16Of9To15Of12").value)
    ACCLTCGDTAA_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("AccLTCGDTAA.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("AccLTCGDTAA.Up16Of12To15Of3").value)
    ACCLTCGDTAA_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("AccLTCGDTAA.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("AccLTCGDTAA.Up16Of3To31Of3").value)
    
    ACCSTCGDTAA_Upto15Of6 = IIf(Len(Sheet13.Range("AccSTCGDTAA.Upto15Of6")) = 0, 0, Sheet13.Range("AccSTCGDTAA.Upto15Of6").value)
    ACCSTCGDTAA_Upto15Of9 = IIf(Len(Sheet13.Range("AccSTCGDTAA.Upto15Of9")) = 0, 0, Sheet13.Range("AccSTCGDTAA.Upto15Of9").value)
    ACCSTCGDTAA_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("AccSTCGDTAA.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("AccSTCGDTAA.Up16Of9To15Of12").value)
    ACCSTCGDTAA_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("AccSTCGDTAA.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("AccSTCGDTAA.Up16Of12To15Of3").value)
    ACCSTCGDTAA_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("AccSTCGDTAA.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("AccSTCGDTAA.Up16Of3To31Of3").value)
    
'Change 102 21.12.2023 V0.6 SIT-65102 B
    AccVDA_Upto15Of6 = IIf(Len(Sheet13.Range("AccVDA.Upto15Of6")) = 0, 0, Sheet13.Range("AccVDA.Upto15Of6").value)
    AccVDA_Upto15Of9 = IIf(Len(Sheet13.Range("AccVDA.Upto15Of9")) = 0, 0, Sheet13.Range("AccVDA.Upto15Of9").value)
    AccVDA_Up16Of9To15Of12 = IIf(Len(Sheet13.Range("AccVDA.Up16Of9To15Of12")) = 0, 0, Sheet13.Range("AccVDA.Up16Of9To15Of12").value)
    AccVDA_Up16Of12To15Of3 = IIf(Len(Sheet13.Range("AccVDA.Up16Of12To15Of3")) = 0, 0, Sheet13.Range("AccVDA.Up16Of12To15Of3").value)
    AccVDA_Up16Of3To31Of3 = IIf(Len(Sheet13.Range("AccVDA.Up16Of3To31Of3")) = 0, 0, Sheet13.Range("AccVDA.Up16Of3To31Of3").value)
'End Change SIT-65102 B
    DEDN_1c_TotDeductClaim = IIf(Len(Sheet13.Range("DEDN.1c_TotDeductClaim")) = 0, 0, Sheet13.Range("DEDN.1c_TotDeductClaim").value)

End Function

Sub setTblinfo_STCG_PrevYear1()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A7_PreviousYer_1").count
 Set rangecells = Sheet13.Range("STCG.A7_PreviousYer_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYear1 = ccount
End Sub

Sub setTblinfo_STCG_PrevYear2()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A7_DeductionSection_1").count
 Set rangecells = Sheet13.Range("STCG.A7_DeductionSection_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYear2 = ccount
End Sub

Sub setTblinfo_STCG_PrevYear3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A7_AsstAcquiredYer_1").count
 Set rangecells = Sheet13.Range("STCG.A7_AsstAcquiredYer_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYear3 = ccount
End Sub

Sub setTblinfo_STCG_PrevYear4()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A7_AmtUtilised_1").count
 Set rangecells = Sheet13.Range("STCG.A7_AmtUtilised_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYear4 = ccount
End Sub

Sub setTblinfo_STCG_PrevYear5()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A7_AmtNotUsed_1").count
 Set rangecells = Sheet13.Range("STCG.A7_AmtNotUsed_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYear5 = ccount
End Sub

Function validateSTCGValidate() As Boolean
validateSTCGValidate = True
        Dim rangecells1, rangecells2, rangecells3, rangecells4, rangecells5 As Range
        Dim i As Variant
        Dim count, temp As Variant
        
        setTblinfo_STCG_PrevYear1
        setTblinfo_STCG_PrevYear2
        setTblinfo_STCG_PrevYear3
        setTblinfo_STCG_PrevYear4
        setTblinfo_STCG_PrevYear5
        
        If (Mid(Sheet13.Range("STCG.A6_StatusUnutilizedCG").value, 1, 1) = "Y") Then
        If (end_PrevYear2 = 0 And end_PrevYear1 = 0 And end_PrevYear5 = 0) Then
            msgboxSHEETCG = ("Please enter amount of unutilized capital gain on asset transferred during the previous years.")
            validateSTCGValidate = False
            Exit Function
        End If

        If (end_PrevYear2 <> end_PrevYear1) Or (end_PrevYear2 <> end_PrevYear5) Then
            msgboxSHEETCG = ("Enter All Mandatory Fields in Table at Point A7. of STCG in schedule CG.")
            validateSTCGValidate = False
            Exit Function
        End If
        End If
        
        
        
        end_PrevYear1 = WorksheetFunction.Max(0, end_PrevYear1, end_PrevYear2, end_PrevYear3, end_PrevYear4, end_PrevYear5)

        Set rangecells1 = Sheet13.Range("STCG.A7_PreviousYer_1").Cells
        Set rangecells2 = Sheet13.Range("STCG.A7_DeductionSection_1").Cells
        Set rangecells3 = Sheet13.Range("STCG.A7_AsstAcquiredYer_1").Cells
        Set rangecells4 = Sheet13.Range("STCG.A7_AmtUtilised_1").Cells
        Set rangecells5 = Sheet13.Range("STCG.A7_AmtNotUsed_1").Cells
        
        
        ReDim STCG_A7_PreviousYer_1(end_PrevYear1)
        ReDim STCG_A7_DeductionSection_1(end_PrevYear1)
        ReDim STCG_A7_AsstAcquiredYer_1(end_PrevYear1)
        ReDim STCG_A7_AmtUtilised_1(end_PrevYear1)
        ReDim STCG_A7_AmtNotUsed_1(end_PrevYear1)
        
        
        For i = 1 To end_PrevYear1
            STCG_A7_PreviousYer_1(i) = rangecells1.item(i).value
            STCG_A7_DeductionSection_1(i) = rangecells2.item(i).value
            STCG_A7_AsstAcquiredYer_1(i) = Mid(rangecells3.item(i).value, 1, 4)
            STCG_A7_AmtUtilised_1(i) = rangecells4.item(i).value
            STCG_A7_AmtNotUsed_1(i) = rangecells5.item(i).value
        
        If isdropdownblank(STCG_A7_PreviousYer_1(i)) Then
            validateSTCGValidate = False
            msgboxSHEETCG = msgboxSHEETCG + "* Please select Previous year in which asset transferred at Pt 7.  (STCG) at Sr No." & i & "   in Schedule CG" & Chr(13)
            Exit Function
        End If
           
        If (isdropdownblank(STCG_A7_DeductionSection_1(i))) Then
            validateSTCGValidate = False
            msgboxSHEETCG = msgboxSHEETCG + "* Please select section under which deduction is claimed at Pt 7.  (STCG) at Sr No." & i & "  in Schedule CG" & Chr(13)
            Exit Function
        End If
        
         If isdropdownblank(STCG_A7_AsstAcquiredYer_1(i)) Then
            validateSTCGValidate = False
            msgboxSHEETCG = msgboxSHEETCG + "* Please select year in which asset aquired in which asset transferred at Pt 7. (STCG) at Sr No." & i & "  in Schedule CG" & Chr(13)
            Exit Function
        End If
        
          If STCG_A7_AmtUtilised_1(i) = "" Then
                    validateSTCGValidate = False
                    msgboxSHEETCG = msgboxSHEETCG + "* Amount Utilised at Pt No. 7 (STCG) at Sr No." & i & " is mandatory in Schedule CG" & Chr(13)
                    Exit Function
         End If
        
        If STCG_A7_AmtNotUsed_1(i) = "" Then
                    validateSTCGValidate = False
                    msgboxSHEETCG = msgboxSHEETCG + "* Amount not Used at Pt No. 7 (STCG) at Sr No." & i & " is mandatory in Schedule CG" & Chr(13)
                    Exit Function
        End If
        
       
        Next
        
            
        
End Function
Sub setTblinfo_LTCG_PrevYear1()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B9_PreviousYer_1").count
 Set rangecells = Sheet13.Range("LTCG.B9_PreviousYer_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYearLTCG1 = ccount
End Sub

Sub setTblinfo_LTCG_PrevYear2()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B9_DeductionSection_1").count
 Set rangecells = Sheet13.Range("LTCG.B9_DeductionSection_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYearLTCG2 = ccount
End Sub

Sub setTblinfo_LTCG_PrevYear3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B9_AsstAcquiredYer_1").count
 Set rangecells = Sheet13.Range("LTCG.B9_AsstAcquiredYer_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYearLTCG3 = ccount
End Sub

Sub setTblinfo_LTCG_PrevYear4()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B9_AmtUtilised_1").count
 Set rangecells = Sheet13.Range("LTCG.B9_AmtUtilised_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYearLTCG4 = ccount
End Sub

Sub setTblinfo_LTCG_PrevYear5()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B9_AmtNotUsed_1").count
 Set rangecells = Sheet13.Range("LTCG.B9_AmtNotUsed_1").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PrevYearLTCG5 = ccount
End Sub

'Newly added by sai on 06/03/2025 for AY 2025-26

'commented by sadineni for AY 2026-27 on 26/02/2026
'=======================================================================================
'Sub setTblinfo_LTCG_PrevYear6()
' Dim rangecells As Range
' Dim mIntCells  As Long
' Dim mIntCtr  As Long
' Dim ccount  As Long
' ccount = 0
' mIntCells = Sheet13.Range("LTCG.B9_dateof_limitorwithdrawal_1").count
' Set rangecells = Sheet13.Range("LTCG.B9_dateof_limitorwithdrawal_1").Cells
'  Dim countrycd As Variant
' For mIntCtr = 1 To mIntCells
'     countrycd = rangecells.item(mIntCtr).value
'        If isdropdownblank(countrycd) Then
'            countrycd = ""
'        End If
'
'     If Not countrycd = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_PrevYearLTCG6 = ccount
'End Sub
'=========================================================================================

Function validateLTCGvalidate() As Boolean
validateLTCGvalidate = True
        Dim rangecells1, rangecells2, rangecells3, rangecells4, rangecells5 As Range
        Dim rangecells6 As Range
        Dim i As Variant
        Dim count, temp As Variant
        
        setTblinfo_LTCG_PrevYear1
        setTblinfo_LTCG_PrevYear2
        setTblinfo_LTCG_PrevYear3
        setTblinfo_LTCG_PrevYear4
        setTblinfo_LTCG_PrevYear5
'        setTblinfo_LTCG_PrevYear6 'commented by sadineni on 26/02/2026 for AY 2026-27
        
'        end_PrevYearLTCG1 = WorksheetFunction.Max(0, end_PrevYearLTCG1, end_PrevYearLTCG2, end_PrevYearLTCG3, end_PrevYearLTCG4, end_PrevYearLTCG5, end_PrevYearLTCG6)
         end_PrevYearLTCG1 = WorksheetFunction.Max(0, end_PrevYearLTCG1, end_PrevYearLTCG2, end_PrevYearLTCG3, end_PrevYearLTCG4, end_PrevYearLTCG5)

        Set rangecells1 = Sheet13.Range("LTCG.B9_PreviousYer_1").Cells
        Set rangecells2 = Sheet13.Range("LTCG.B9_DeductionSection_1").Cells
        Set rangecells3 = Sheet13.Range("LTCG.B9_AsstAcquiredYer_1").Cells
        Set rangecells4 = Sheet13.Range("LTCG.B9_AmtUtilised_1").Cells
        Set rangecells5 = Sheet13.Range("LTCG.B9_AmtNotUsed_1").Cells
        'Set rangecells6 = Sheet13.Range("LTCG.B9_dateof_limitorwithdrawal_1").Cells
        
        
        ReDim LTCG_B9_PreviousYer_1(end_PrevYearLTCG1)
        ReDim LTCG_B9_DeductionSection_1(end_PrevYearLTCG1)
        ReDim LTCG_B9_AsstAcquiredYer_1(end_PrevYearLTCG1)
        ReDim LTCG_B9_AmtUtilised_1(end_PrevYearLTCG1)
        ReDim LTCG_B9_AmtNotUsed_1(end_PrevYearLTCG1)
        
    'by sai on 06/03/2025 for AY 2025-26************************************
    'commented by sadineni on 26/02/2026 for AY 2026-27
'        ReDim LTCG_B9_dateof_limitorwithdrawal_1(end_PrevYearLTCG1)
     '*********************************************************************
        For i = 1 To end_PrevYearLTCG1
            LTCG_B9_PreviousYer_1(i) = rangecells1.item(i).value
            LTCG_B9_DeductionSection_1(i) = rangecells2.item(i).value
            LTCG_B9_AsstAcquiredYer_1(i) = Mid(rangecells3.item(i).value, 1, 4)
            LTCG_B9_AmtUtilised_1(i) = rangecells4.item(i).value
            LTCG_B9_AmtNotUsed_1(i) = rangecells5.item(i).value
        'by sai on 06/03/2025 for AY 2025-26
        'commented by sadineni on 26/02/2026 for AY 2026-27
'            LTCG_B9_dateof_limitorwithdrawal_1(i) = rangecells6.item(i).value
        
        If isdropdownblank(LTCG_B9_PreviousYer_1(i)) Then
            validateLTCGvalidate = False
            msgboxSHEETCG = msgboxSHEETCG + "* Please select Previous year in which asset transferred at Pt 11. (LTCG) in Schedule CG" & Chr(13)
            Exit Function
        End If
           
        If (isdropdownblank(LTCG_B9_DeductionSection_1(i))) Then
            validateLTCGvalidate = False
            msgboxSHEETCG = msgboxSHEETCG + "* Please select section under which deduction is claimed at Pt 11. (LTCG) in Schedule CG" & Chr(13)
            Exit Function
        End If
        
'commented by sadineni on 17/06/2026 as the field is non mandatory as per ITR-2 IPIP-103807
'         If isdropdownblank(LTCG_B9_AsstAcquiredYer_1(i)) Then
'            validateLTCGvalidate = False
'            msgboxSHEETCG = msgboxSHEETCG + "* Please select year in which asset aquired in which asset transferred at Pt 11. (LTCG) at Sr No." & i & "  in Schedule CG" & Chr(13)
'            Exit Function
'        End If
        
          If LTCG_B9_AmtUtilised_1(i) = "" Then
                    validateLTCGvalidate = False
                    msgboxSHEETCG = msgboxSHEETCG + "* Amount Utilised at Pt No. 7 (STCG) at Sr No." & i & " is mandatory in Schedule CG" & Chr(13)
                    Exit Function
         End If
        
        If LTCG_B9_AmtNotUsed_1(i) = "" Then
                    validateLTCGvalidate = False
                    msgboxSHEETCG = msgboxSHEETCG + "* Amount used is Mandatory at Pt 11 SrNo." & Chr(13)
                    Exit Function
        End If
    'by sai on 06/03/2025 for AY 2025-26
    'commented by sadineni on 26/02/2026 for AY 2026-27
'        If LTCG_B9_dateof_limitorwithdrawal_1(i) = "" Then
'                    validateLTCGvalidate = False
'                    msgboxSHEETCG = msgboxSHEETCG + "* Please select Whether date of limitation / withdrawal was before 23rd July 2024  at Pt 11 (LTCG) at SrNo." & i & " in Schedule CG" & Chr(13)
'                    Exit Function
'        End If
       
        Next
        
            
        
End Function

Function validateLTCGDTAA() As Boolean
    validateLTCGDTAA = True
    setTblinfo_LTCGDTAA5
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "B10: LTCG DTAA :" & Chr(13) & Chr(13)
    If Not ValidateAmountLTCG_DTAA() Then validateLTCGDTAA = False
    If (Len(Sheet13.Range("LTCG.B10_LtcgAmt").item(0).value) > 0) Then
        If Not ValidateArticleLTCG_DTAA() Then validateLTCGDTAA = False
        If Not ValidateTRCLTCG_DTAA() Then validateLTCGDTAA = False
        If Not ValidateSectionLTCG_DTAA() Then validateLTCGDTAA = False
        If Not ValidateCountry_LTCGDTAA() Then validateLTCGDTAA = False
        If Not ValidateRateTreatyLTCG_DTAA() Then validateLTCGDTAA = False
        If Not ValidateSectionITActLTCG_DTAA() Then validateLTCGDTAA = False
        If Not ValidateRateITActLTCG_DTAA() Then validateLTCGDTAA = False

    End If
    
    setTblinfo_LTCGDTAA2
    setTblinfo_LTCGDTAA3
    setTblinfo_LTCGDTAA4
    setTblinfo_LTCGDTAA
    setTblinfo_LTCGDTAA6
    setTblinfo_LTCGDTAA7
    setTblinfo_LTCGDTAA8
    
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
    
     If (end_LTCGDTAA5 <> end_LTCGDTAA2) Or (end_LTCGDTAA5 <> end_LTCGDTAA3) Or (end_LTCGDTAA5 <> end_LTCGDTAA4) Or _
        (end_LTCGDTAA5 <> end_LTCGDTAA) Or (end_LTCGDTAA5 <> end_LTCGDTAA6) Or (end_LTCGDTAA5 <> end_LTCGDTAA7) Or (end_LTCGDTAA5 <> end_LTCGDTAA8) Then
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields." & Chr(13)
             validateLTCGDTAA = False
             Exit Function
    
    End If
    Else
    If (end_LTCGDTAA5 <> end_LTCGDTAA2) Or (end_LTCGDTAA5 <> end_LTCGDTAA4) Or _
        (end_LTCGDTAA5 <> end_LTCGDTAA) Or (end_LTCGDTAA5 <> end_LTCGDTAA6) Or (end_LTCGDTAA5 <> end_LTCGDTAA7) Or (end_LTCGDTAA5 <> end_LTCGDTAA8) Then
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields." & Chr(13)
             validateLTCGDTAA = False
             Exit Function
    
    End If
    End If
    
End Function

 Function ValidateCountry_LTCGDTAA() As Boolean
 
    Dim CountryName As String
    ValidateCountry_LTCGDTAA = True
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_Country").Cells
    ReDim CountryCodeLTCG_DTAA(end_LTCGDTAA)
    ReDim CountryNameLTCG_DTAA(end_LTCGDTAA)
    For i = 1 To end_LTCGDTAA
        CountryName = rangecells.item(i).value
        
        If isdropdownblank(CountryName) Then
         msgboxSHEETCG = msgboxSHEETCG + "* CountryName at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateCountry_LTCGDTAA = False
         Exit Function
     End If
        CountryNameLTCG_DTAA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
        CountryCodeLTCG_DTAA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        
     If Not chkCompulsory(CountryCodeLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* CountryName at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateCountry_LTCGDTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateArticleLTCG_DTAA() As Boolean
 
    ValidateArticleLTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_DTAA_Article").Cells
    ReDim ArticleLTCG_DTAA(end_LTCGDTAA)
    For i = 1 To end_LTCGDTAA
        ArticleLTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(ArticleLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Article of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateArticleLTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateTRCLTCG_DTAA() As Boolean
 
    ValidateTRCLTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_CertiStatus").Cells
    ReDim TRCLTCG_DTAA(end_LTCGDTAA)
    
    
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
    
    For i = 1 To end_LTCGDTAA
        TRCLTCG_DTAA(i) = rangecells.item(i).value
        
     If isdropdownblank(TRCLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Whether TRC obtained at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateTRCLTCG_DTAA = False
         Exit Function
     End If
 Next
 End If
End Function

Function ValidateSectionLTCG_DTAA() As Boolean
    ValidateSectionLTCG_DTAA = True
    Dim rangecells, rangecells2 As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_ItemNo").Cells
    'Set rangecells2 = Sheet13.Range("CG_LtcgDtaa_Xml_code").Cells
    
    ReDim SectionLTCG_DTAA(end_LTCGDTAA)
    ReDim SectionCodeLTCG_DTAA(end_LTCGDTAA)
    
    For i = 1 To end_LTCGDTAA
        SectionLTCG_DTAA(i) = rangecells.item(i).value
        'SectionCodeLTCG_DTAA(i) = rangecells2.item(i).value
        
     If Not isdropdownblank(SectionLTCG_DTAA(i)) Then
            SectionLTCG_DTAA(i) = SectionLTCG_DTAA(i)


'by sai on 16/03/2025 For AY 2025-26
'***********************************************************************************************************
'            If SectionLTCG_DTAA(i) = "B1e_20%" Then SectionLTCG_DTAA(i) = "B1e_20%"
'            If SectionLTCG_DTAA(i) = "B1e_12.5%" Then SectionLTCG_DTAA(i) = "B1e_12.5%"
'            If SectionLTCG_DTAA(i) = "B2e_20%" Then SectionLTCG_DTAA(i) = "B2e_20%"
'            If SectionLTCG_DTAA(i) = "B2e_12.5%" Then SectionLTCG_DTAA(i) = "B2e_12.5%"
'            If SectionLTCG_DTAA(i) = "B3e" Then SectionLTCG_DTAA(i) = "B3e"
'            If SectionLTCG_DTAA(i) = "B4ie_20%" Then SectionLTCG_DTAA(i) = "B4ie_20%"
'            If SectionLTCG_DTAA(i) = "B4ie_12.5%" Then SectionLTCG_DTAA(i) = "B4ie_12.5%"
'            If SectionLTCG_DTAA(i) = "B4iie_10%" Then SectionLTCG_DTAA(i) = "B4iie_10%"
'            If SectionLTCG_DTAA(i) = "B4iie_12.5%" Then SectionLTCG_DTAA(i) = "B4iie_12.5%"
'            If SectionLTCG_DTAA(i) = "B5c_10%" Then SectionLTCG_DTAA(i) = "B5c_10%"
'            If SectionLTCG_DTAA(i) = "B5c_12.5%" Then SectionLTCG_DTAA(i) = "B5c_12.5%"
'            If SectionLTCG_DTAA(i) = "B6c_20%" Then SectionLTCG_DTAA(i) = "B6c_20%"
'            If SectionLTCG_DTAA(i) = "B6c_10%" Then SectionLTCG_DTAA(i) = "B6c_10%"
'            If SectionLTCG_DTAA(i) = "B6c_12.5%" Then SectionLTCG_DTAA(i) = "B6c_12.5%"
'            If SectionLTCG_DTAA(i) = "B7ie_10%" Then SectionLTCG_DTAA(i) = "B7ie_10%"
'            If SectionLTCG_DTAA(i) = "B7ie_12.5%" Then SectionLTCG_DTAA(i) = "B7ie_12.5%"
'            If SectionLTCG_DTAA(i) = "B7iie_10%" Then SectionLTCG_DTAA(i) = "B7iie_10%"
'            If SectionLTCG_DTAA(i) = "B7iie_12.5%" Then SectionLTCG_DTAA(i) = "B7iie_12.5%"
'            If SectionLTCG_DTAA(i) = "B7iiie" Then SectionLTCG_DTAA(i) = "B7iiie"
'            If SectionLTCG_DTAA(i) = "B8c_10%" Then SectionLTCG_DTAA(i) = "B8c_10%"
'            If SectionLTCG_DTAA(i) = "B8c_12.5%" Then SectionLTCG_DTAA(i) = "B8c_12.5%"
'            If SectionLTCG_DTAA(i) = "B9c_10%" Then SectionLTCG_DTAA(i) = "B9c_10%"
'            If SectionLTCG_DTAA(i) = "B9c_12.5%" Then SectionLTCG_DTAA(i) = "B9c_12.5%"
''            If SectionLTCG_DTAA(i) = "B9d_20%" Then SectionLTCG_DTAA(i) = "B9d_20%"
''            If SectionLTCG_DTAA(i) = "B9d_12.5%" Then SectionLTCG_DTAA(i) = "B9d_12.5%"
'            If SectionLTCG_DTAA(i) = "B10e_20%" Then SectionLTCG_DTAA(i) = "B10e_20%"
'            If SectionLTCG_DTAA(i) = "B10e_12.5%" Then SectionLTCG_DTAA(i) = "B10e_12.5%"
'            If SectionLTCG_DTAA(i) = "B11_20%" Then SectionLTCG_DTAA(i) = "B11_20%"
'            If SectionLTCG_DTAA(i) = "B11_12.5%" Then SectionLTCG_DTAA(i) = "B11_12.5%"
'            If SectionLTCG_DTAA(i) = "B12a1_10%" Then SectionLTCG_DTAA(i) = "B12a1_10%"
'            If SectionLTCG_DTAA(i) = "B12a1_12.5%" Then SectionLTCG_DTAA(i) = "B12a1_12.5%"
'            If SectionLTCG_DTAA(i) = "B12a2_10%" Then SectionLTCG_DTAA(i) = "B12a2_10%"
'            If SectionLTCG_DTAA(i) = "B12a2_12.5%" Then SectionLTCG_DTAA(i) = "B12a2_12.5%"
'
'            'If SectionLTCG_DTAA(i) = "B12(b)" Then SectionLTCG_DTAA(i) = "B12b"
'
'             If SectionLTCG_DTAA(i) = "B12(b)" Then SectionLTCG_DTAA(i) = "B12(b)" 'by Malli on 23/06/2025

'Updated by sadineni on 13/05/2026
'*************************************************************************************************************************************************************
If SectionLTCG_DTAA(i) = "B1g-Total LTCG on all Immovable properties" Then SectionLTCG_DTAA(i) = "B1g"
If SectionLTCG_DTAA(i) = "B2e-LTCG from Slump sale" Then SectionLTCG_DTAA(i) = "B2e"
If SectionLTCG_DTAA(i) = "B3ie-112(1) LTCG u/s 112(1) from listed securities (other than a unit) or zero coupon bonds" Then SectionLTCG_DTAA(i) = "B3ie"
If SectionLTCG_DTAA(i) = "B3iie-115ACA - LTCG u/s 115ACA from GDR of an Indian company" Then SectionLTCG_DTAA(i) = "B3iie_115ACA"
If SectionLTCG_DTAA(i) = "B4c-112A- LTCG from sale of equity share in a company or unit of equity oriented fund or unit of a business trust on which STT is paid" Then SectionLTCG_DTAA(i) = "B4c_112A"
If SectionLTCG_DTAA(i) = "B5c-LTCG on unlisted shares or listed debentures" Then SectionLTCG_DTAA(i) = "B5c"
If SectionLTCG_DTAA(i) = "B6ie-112(1)(c) - LTCG u/s 112(1)(c) from unlisted securities" Then SectionLTCG_DTAA(i) = "B6ie_112(1)(c)"
If SectionLTCG_DTAA(i) = "B6iie-115AC - LTCG u/s 115AC from bonds or GDR" Then SectionLTCG_DTAA(i) = "B6iie_115AC"
If SectionLTCG_DTAA(i) = "B6iiie-115AD - LTCG u/s 115AD from securities by FII" Then SectionLTCG_DTAA(i) = "B6iiie_115AD"
If SectionLTCG_DTAA(i) = "B7c-115AD(1)(iii) proviso - LTCG under section 112A read with section 115AD" Then SectionLTCG_DTAA(i) = "B7c_115AD(1)(iii)_proviso"
If SectionLTCG_DTAA(i) = "B8c-LTCG from sale of foreign exchange asset by NRI" Then SectionLTCG_DTAA(i) = "B8c"
If SectionLTCG_DTAA(i) = "B9e-LTCG from sale of assets where B1 to B8 are not applicable" Then SectionLTCG_DTAA(i) = "B9e"
If SectionLTCG_DTAA(i) = "B10-Amount deemed to be long-term capital gains" Then SectionLTCG_DTAA(i) = "B10"
If SectionLTCG_DTAA(i) = "B11a1-Pass Through Income/ Loss in the nature of LTCG, chargeable @ 12.5% u/s 112A" Then SectionLTCG_DTAA(i) = "B11a1"
If SectionLTCG_DTAA(i) = "B11a2-Pass Through Income/ Loss in the nature of LTCG, chargeable @ 12.5% other than u/s 112A" Then SectionLTCG_DTAA(i) = "B11a2"
'**************************************************************************************************************************************************************
       End If
        
     If isdropdownblank(SectionLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Item B1 to B8 above in which included at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateSectionLTCG_DTAA = False
         Exit Function
     End If
   
 Next
End Function

Function ValidateAmountLTCG_DTAA() As Boolean
     ValidateAmountLTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_LtcgAmt").Cells
    ReDim AmountLTCG_DTAA(end_LTCGDTAA)
    For i = 1 To end_LTCGDTAA
        AmountLTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(AmountLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount of LTCG at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateAmountLTCG_DTAA = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountLTCG_DTAA(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount of LTCG at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateAmountLTCG_DTAA = False
        Exit Function
    End If
    
    If (Len(AmountLTCG_DTAA(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount of LTCG at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateAmountLTCG_DTAA = False
        Exit Function
    End If
 Next
End Function

Function ValidateRateTreatyLTCG_DTAA() As Boolean
 
    ValidateRateTreatyLTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_RatepTreaty").Cells
    ReDim RateTreatyLTCG_DTAA(end_LTCGDTAA)
    For i = 1 To end_LTCGDTAA6
        RateTreatyLTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(RateTreatyLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Rate as per Treaty of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateRateTreatyLTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateSectionITActLTCG_DTAA() As Boolean
 
    ValidateSectionITActLTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_SecIT").Cells
    ReDim SectionITActLTCG_DTAA(end_LTCGDTAA)
    For i = 1 To end_LTCGDTAA7
        SectionITActLTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(SectionITActLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Section of IT Act of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateSectionITActLTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateRateITActLTCG_DTAA() As Boolean
 
    ValidateRateITActLTCG_DTAA = True
    Dim rangecells, rangecells1 As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("LTCG.B10_RatePITAct").Cells
    Set rangecells1 = Sheet13.Range("LTCG.B10_ApplRate").Cells
    ReDim RateITActLTCG_DTAA(end_LTCGDTAA)
    ReDim Applrate_LTCG(end_LTCGDTAA)
    For i = 1 To end_LTCGDTAA8
        RateITActLTCG_DTAA(i) = rangecells.item(i).value
        Applrate_LTCG(i) = rangecells1.item(i).value
        
     If Not chkCompulsory(RateITActLTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Rate as per IT Act of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateRateITActLTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateTotAmountLTCG_DTAA() As Boolean
    ValidateTotAmountLTCG_DTAA = True
        TotAmountLTCG_DTAA = Sheet13.Range("LTCG.B10_ltcgDTAA").value

    If (Len(TotAmountLTCG_DTAA) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Total amount of LTCG chargeable to tax under DTAA cannot exceed 14 digits." & Chr(13)
        ValidateTotAmountLTCG_DTAA = False
        Exit Function
    End If
End Function

Sub setTblinfo_LTCGDTAA()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_Country").count
 Set rangecells = Sheet13.Range("LTCG.B10_Country").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA = ccount
rngname_LTCGDTAA = "LTCG.B10_LtcgAmt;LTCG.B10_ItemNo;LTCG.B10_Country;LTCG.B10_DTAA_Article;LTCG.B10_RatepTreaty;LTCG.B10_CertiStatus;LTCG.B10_SecIT;LTCG.B10_RatePITAct;LTCG.B10_ApplRate;"
End Sub

 Sub setTblinfo_LTCGDTAA2()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_DTAA_Article").count
 Set rangecells = Sheet13.Range("LTCG.B10_DTAA_Article").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA2 = ccount
 End Sub
 
 Sub setTblinfo_LTCGDTAA3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim trc As Variant
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_CertiStatus").count
 Set rangecells = Sheet13.Range("LTCG.B10_CertiStatus").Cells
 For mIntCtr = 1 To mIntCells
 
  trc = rangecells.item(mIntCtr).value
        If isdropdownblank(trc) Then
            trc = ""
        End If
    
     If Not trc = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA3 = ccount
 End Sub
 
 Sub setTblinfo_LTCGDTAA4()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim section As Variant
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_ItemNo").count
 Set rangecells = Sheet13.Range("LTCG.B10_ItemNo").Cells
 For mIntCtr = 1 To mIntCells
 
  section = rangecells.item(mIntCtr).value
        If isdropdownblank(section) Then
            section = ""
        End If
    
     If Not section = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA4 = ccount
 End Sub
 

 Sub setTblinfo_LTCGDTAA5()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_LtcgAmt").count
 Set rangecells = Sheet13.Range("LTCG.B10_LtcgAmt").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA5 = ccount
 rngname_LTCGDTAA = "LTCG.B10_LtcgAmt;LTCG.B10_ItemNo;LTCG.B10_Country;LTCG.B10_DTAA_Article;LTCG.B10_RatepTreaty;LTCG.B10_CertiStatus;LTCG.B10_SecIT;LTCG.B10_RatePITAct;LTCG.B10_ApplRate;"

 End Sub
 
 Sub setTblinfo_LTCGDTAA6()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_RatepTreaty").count
 Set rangecells = Sheet13.Range("LTCG.B10_RatepTreaty").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA6 = ccount
 End Sub
 
 
 Sub setTblinfo_LTCGDTAA7()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_SecIT").count
 Set rangecells = Sheet13.Range("LTCG.B10_SecIT").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA7 = ccount
 End Sub
 
 
 Sub setTblinfo_LTCGDTAA8()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("LTCG.B10_RatePITAct").count
 Set rangecells = Sheet13.Range("LTCG.B10_RatePITAct").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGDTAA8 = ccount
 End Sub
 
 Function ValidateSTCGDTAA() As Boolean
    ValidateSTCGDTAA = True
    setTblinfo_STCGDTAA5
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "A8: STCG DTAA :" & Chr(13) & Chr(13)
    If Not ValidateAmountSTCG_DTAA() Then ValidateSTCGDTAA = False
    If (Len(Sheet13.Range("STCG.A8_AmtI").item(0).value) > 0) Then
        If Not ValidateArticleSTCG_DTAA() Then ValidateSTCGDTAA = False
        If Not ValidateTRCSTCG_DTAA() Then ValidateSTCGDTAA = False
        If Not ValidateSectionSTCG_DTAA() Then ValidateSTCGDTAA = False
        'If Not ValidateAmountSTCG_DTAA() Then ValidateSTCGDTAA = False
        If Not ValidateCountry_STCGDTAA() Then ValidateSTCGDTAA = False
        If Not ValidateRateTreatySTCG_DTAA() Then ValidateSTCGDTAA = False
        If Not ValidateRateITActSTCG_DTAA() Then ValidateSTCGDTAA = False
        If Not ValidateSectionITActSTCG_DTAA() Then ValidateSTCGDTAA = False

        
    End If
    
    setTblinfo_STCGDTAA2
    setTblinfo_STCGDTAA3
    setTblinfo_STCGDTAA4
    setTblinfo_STCGDTAA
    setTblinfo_STCGDTAA6
    setTblinfo_STCGDTAA7
    setTblinfo_STCGDTAA8
    
    If end_STCGDTAA Or end_STCGDTAA2 Or end_STCGDTAA3 Or end_STCGDTAA4 Or end_STCGDTAA6 Or end_STCGDTAA7 Or end_STCGDTAA8 >= 1 Then
    If end_STCGDTAA5 = "" Or end_STCGDTAA5 = 0 Then
    
    
 ValidateAmountSTCG_DTAA_New
    ValidateSTCGDTAA = False
    End If
   
    
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
    
      If (end_STCGDTAA5 <> end_STCGDTAA2) Or (end_STCGDTAA5 <> end_STCGDTAA3) Or (end_STCGDTAA5 <> end_STCGDTAA4) Or _
        (end_STCGDTAA5 <> end_STCGDTAA) Or (end_STCGDTAA5 <> end_STCGDTAA6) Or (end_STCGDTAA5 <> end_STCGDTAA7) Or (end_STCGDTAA5 <> end_STCGDTAA8) Then
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields." & Chr(13)
             ValidateSTCGDTAA = False
             Exit Function
    
    End If
    Else
     If (end_STCGDTAA5 <> end_STCGDTAA2) Or (end_STCGDTAA5 <> end_STCGDTAA4) Or _
        (end_STCGDTAA5 <> end_STCGDTAA) Or (end_STCGDTAA5 <> end_STCGDTAA6) Or (end_STCGDTAA5 <> end_STCGDTAA7) Or (end_STCGDTAA5 <> end_STCGDTAA8) Then
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields." & Chr(13)
             ValidateSTCGDTAA = False
             Exit Function
    
    End If
    
    End If
    End If
   
End Function
 Function ValidateCountry_STCGDTAA() As Boolean
 
    Dim CountryName As String
    ValidateCountry_STCGDTAA = True
    Dim rangecells As Range
    Dim i As Long
    
    
    Set rangecells = Sheet13.Range("STCG.A8_Country").Cells
    ReDim CountryCodeSTCG_DTAA(end_STCGDTAA5)
    ReDim CountryNameSTCG_DTAA(end_STCGDTAA5)
    
    For i = 1 To end_STCGDTAA5
        CountryName = rangecells.item(i).value
        
        If isdropdownblank(CountryName) Then
         msgboxSHEETCG = msgboxSHEETCG + "* CountryName at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateCountry_STCGDTAA = False
         Exit Function
     End If
        CountryNameSTCG_DTAA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
        CountryCodeSTCG_DTAA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        
     If Not chkCompulsory(CountryCodeSTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* CountryName at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateCountry_STCGDTAA = False
         Exit Function
     End If

     
 Next
End Function
Function ValidateCountryStateCode_STCGDTAA() As Boolean
    Dim i As Long
    Dim Country1, state1 As Variant
    ValidateCountryStateCode_STCGDTAA = True
    'setTblinfo_hprptfrm
    
    For i = 1 To end_STCGDTAA5C
       ' If Len(AddrDetail_HP(i)) = 0 Then
          '  Exit For
       ' End If

        Country1 = Sheet13.Range("CG_A_CountryOfBuyer" & i).value
        state1 = Sheet13.Range("CG_A_StateCodeOfBuyer" & i).value
        
        If isdropdownblank(Country1) Then
            Country1 = ""
        End If
        
        If isdropdownblank(state1) Then
            state1 = ""
        End If
        
        If chkCompulsory(Country1) And chkCompulsory(state1) Then
            Country1 = Mid(Country1, 1, WorksheetFunction.Search("-", Country1) - 1)
            state1 = Mid(state1, 1, WorksheetFunction.Search("-", state1) - 1)
                        
            If (Country1 <> "91" And state1 <> "99") Then
                'msgbox_hprptfrm ("* Country cannot be other than India as you have selected an Indian state " & i & "  in Sheet HOUSE_PROPERTY")
                msgboxSHEETCG = msgboxSHEETCG + "* Country cannot be other than India as you have selected an Indian state at Sr. No  " & i & "." & Chr(13)
                 ValidateCountryStateCode_STCGDTAA = False
                Exit Function
            End If

           ' If (Country1 = "91" And State1 = "99") Then
               '  msgboxSHEETCG = msgboxSHEETCG + ("* Country cannot be India as you have selected a Foreign state" & i & "")
                 'ValidateCountryStateCode_STCGDTAA = False
                'Exit Function
            'End If

        End If
    Next
End Function



Function ValidateArticleSTCG_DTAA() As Boolean
 
    ValidateArticleSTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_DTAA_Article").Cells
    ReDim ArticleSTCG_DTAA(end_STCGDTAA5)
    For i = 1 To end_STCGDTAA5
        ArticleSTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(ArticleSTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Article of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateArticleSTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateTRCSTCG_DTAA() As Boolean
 
    ValidateTRCSTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_CertiStatus").Cells
    ReDim TRCSTCG_DTAA(end_STCGDTAA5)
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
    For i = 1 To end_STCGDTAA5
        TRCSTCG_DTAA(i) = rangecells.item(i).value
        
     If isdropdownblank(TRCSTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Whether TRC obtained at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateTRCSTCG_DTAA = False
         Exit Function
     End If
 Next
 End If
End Function

Function ValidateSectionSTCG_DTAA() As Boolean
 
 
    ValidateSectionSTCG_DTAA = True
    Dim rangecells, rangecells2 As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_IA1tA6").Cells
    'Set rangecells2 = Sheet13.Range("CG_stcgDtaa_Xml_code").Cells
    
    ReDim SectionSTCG_DTAA(end_STCGDTAA5)
    ReDim SectionCodeSTCG_DTAA(end_STCGDTAA5)
    
    For i = 1 To end_STCGDTAA5
        SectionSTCG_DTAA(i) = rangecells.item(i).value
        'SectionCodeSTCG_DTAA(i) = rangecells2.item(i).value
        
     If isdropdownblank(SectionSTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Item no. A1 to A6 above in which included at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateSectionSTCG_DTAA = False
         Exit Function
     End If
     
      If Not isdropdownblank(SectionSTCG_DTAA(i)) Then
            SectionSTCG_DTAA(i) = SectionSTCG_DTAA(i)
'            If SectionSTCG_DTAA(i) = "A1e" Then SectionSTCG_DTAA(i) = "A1e"
'            If SectionSTCG_DTAA(i) = "A2c" Then SectionSTCG_DTAA(i) = "A2c"
'            If SectionSTCG_DTAA(i) = "A3ie" Then SectionSTCG_DTAA(i) = "A3ie"
'            If SectionSTCG_DTAA(i) = "A3ie" Then SectionSTCG_DTAA(i) = "A3ie"
'            If SectionSTCG_DTAA(i) = "A3iie" Then SectionSTCG_DTAA(i) = "A3iie"
'
'            'If SectionSTCG_DTAA(i) = "A3ie" Then SectionSTCG_DTAA(i) = "A3e_111A"
'            'If SectionSTCG_DTAA(i) = "A3iie" Then SectionSTCG_DTAA(i) = "A3e_115AD"
'            If SectionSTCG_DTAA(i) = "A4a" Then SectionSTCG_DTAA(i) = "A4a"
'            If SectionSTCG_DTAA(i) = "A4b" Then SectionSTCG_DTAA(i) = "A4b"
'            If SectionSTCG_DTAA(i) = "A5e" Then SectionSTCG_DTAA(i) = "A5e"
'            If SectionSTCG_DTAA(i) = "A6g" Then SectionSTCG_DTAA(i) = "A6g"
'            If SectionSTCG_DTAA(i) = "A7" Then SectionSTCG_DTAA(i) = "A7"
'            If SectionSTCG_DTAA(i) = "A8a" Then SectionSTCG_DTAA(i) = "A8a"
'            If SectionSTCG_DTAA(i) = "A8b" Then SectionSTCG_DTAA(i) = "A8b"
'            If SectionSTCG_DTAA(i) = "A8c" Then SectionSTCG_DTAA(i) = "A8c"
'
'************************************************************************************************
'by sai on 16/03/2025 for AY 2025-26
'            If SectionSTCG_DTAA(i) = "A1e" Then SectionSTCG_DTAA(i) = "A1e"
'            If SectionSTCG_DTAA(i) = "A2c" Then SectionSTCG_DTAA(i) = "A2c"
'            If SectionSTCG_DTAA(i) = "A3ie(i)" Then SectionSTCG_DTAA(i) = "A3ie(i)"
'            If SectionSTCG_DTAA(i) = "A3ie(ii)" Then SectionSTCG_DTAA(i) = "A3ie(ii)"
'            If SectionSTCG_DTAA(i) = "A3iie(i)" Then SectionSTCG_DTAA(i) = "A3iie(i)"
'            If SectionSTCG_DTAA(i) = "A3iie(ii)" Then SectionSTCG_DTAA(i) = "A3iie(ii)"
'            If SectionSTCG_DTAA(i) = "A4a(i)" Then SectionSTCG_DTAA(i) = "A4a(i)"
'            If SectionSTCG_DTAA(i) = "A4a(ii)" Then SectionSTCG_DTAA(i) = "A4a(ii)"
'            If SectionSTCG_DTAA(i) = "A4b" Then SectionSTCG_DTAA(i) = "A4b"
'            If SectionSTCG_DTAA(i) = "A5e" Then SectionSTCG_DTAA(i) = "A5e"
'            If SectionSTCG_DTAA(i) = "A6g" Then SectionSTCG_DTAA(i) = "A6g"
'            If SectionSTCG_DTAA(i) = "A7" Then SectionSTCG_DTAA(i) = "A7"
'            If SectionSTCG_DTAA(i) = "A8(a)(i)" Then SectionSTCG_DTAA(i) = "A8(a)(i)"
'            If SectionSTCG_DTAA(i) = "A8a(ii)" Then SectionSTCG_DTAA(i) = "A8a(ii)"
'            If SectionSTCG_DTAA(i) = "A8(b)" Then SectionSTCG_DTAA(i) = "A8(b)"
'            If SectionSTCG_DTAA(i) = "A8(c)" Then SectionSTCG_DTAA(i) = "A8(c)"

'*********************************************************************************************************
         If SectionSTCG_DTAA(i) = "A1e- Short-term Capital Gains on Immovable property" Then SectionSTCG_DTAA(i) = "A1e"
         If SectionSTCG_DTAA(i) = "A2c- Short-term Capital Gains from Slump sale" Then SectionSTCG_DTAA(i) = "A2c"
         If SectionSTCG_DTAA(i) = "A3ie- 111A- Short-term capital gain on equity share or equity oriented MF (STT paid)" Then SectionSTCG_DTAA(i) = "A3ie"
         If SectionSTCG_DTAA(i) = "A3iie- 115AD- Short-term capital gain on equity share or equity oriented MF (STT paid)" Then SectionSTCG_DTAA(i) = "A3iie"
         If SectionSTCG_DTAA(i) = "A4a- STCG on transactions covered u/s 111A for Non Residents" Then SectionSTCG_DTAA(i) = "A4a"
         If SectionSTCG_DTAA(i) = "A4b- STCG from sale of shares not covered in sl. no 4a or sale of debentures for Non Residents" Then SectionSTCG_DTAA(i) = "A4b"
         If SectionSTCG_DTAA(i) = "A5e- Short-term capital gain on sale of securities by an FII as per section 115AD" Then SectionSTCG_DTAA(i) = "A5e"
         If SectionSTCG_DTAA(i) = "A6g- STCG on assets other than at A1 or A2 or A3 or A4 or A5" Then SectionSTCG_DTAA(i) = "A6g"
         If SectionSTCG_DTAA(i) = "A7- Amount deemed to be short term capital gains" Then SectionSTCG_DTAA(i) = "A7"
         If SectionSTCG_DTAA(i) = "A8a_20%- Pass Through Income/ Loss in the nature of Short-Term Capital Gain, chargeable @ 20%" Then SectionSTCG_DTAA(i) = "A8a"
         If SectionSTCG_DTAA(i) = "A8b- Pass Through Income/ Loss in the nature of Short-Term Capital Gain, chargeable @ 30%" Then SectionSTCG_DTAA(i) = "A8b"
         If SectionSTCG_DTAA(i) = "A8c- Pass Through Income/ Loss in the nature of Short-Term Capital Gain, chargeable at applicable rates" Then SectionSTCG_DTAA(i) = "A8c"
'***********************************************************************************************************
       End If
     
 Next
End Function

Function ValidateAmountSTCG_DTAA() As Boolean
 
    ValidateAmountSTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_AmtI").Cells
    ReDim AmountSTCG_DTAA(end_STCGDTAA5)
    For i = 1 To end_STCGDTAA5
        AmountSTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(AmountSTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount of income at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateAmountSTCG_DTAA = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountSTCG_DTAA(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount of income at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateAmountSTCG_DTAA = False
        Exit Function
    End If
    
    If (Len(AmountSTCG_DTAA(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount of income at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateAmountSTCG_DTAA = False
        Exit Function
    End If
 Next
End Function
Function ValidateAmountSTCG_DTAA_New() As Boolean
    ValidateAmountSTCG_DTAA_New = True
    Dim rangecells As Range
       Dim i As Long
       
    Dim k As Long
    k = Sheet13.Range("STCG.A8_AmtI").Cells.count
    
    Set rangecells = Sheet13.Range("STCG.A8_AmtI").Cells
'    ReDim AmountSTCG_DTAA(end_STCGDTAA5)
    For i = 1 To k
        'AmountSTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(rangecells.item(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount of income at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateAmountSTCG_DTAA_New = False
         Exit Function
     End If
     
     If Not IsNumeric(rangecells.item(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount of income at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateAmountSTCG_DTAA_New = False
        Exit Function
    End If
    'End Function
    
    If (Len(rangecells.item(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount of income at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateAmountSTCG_DTAA_New = False
        Exit Function
    End If
 Next
End Function
Function ValidateTotAmountSTCG_DTAA() As Boolean
    ValidateTotAmountSTCG_DTAA = True
        TotAmountSTCG_DTAA = Sheet13.Range("STCG.A7_TotalSTCG").value

    If (Len(TotAmountSTCG_DTAA) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Total amount of STCG chargeable to tax under DTAA cannot exceed 14 digits." & Chr(13)
        ValidateTotAmountSTCG_DTAA = False
        Exit Function
    End If
End Function

Function ValidateRateTreatySTCG_DTAA() As Boolean
 
    ValidateRateTreatySTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_RatePTreaty").Cells
    ReDim RateTreatySTCG_DTAA(end_STCGDTAA5)
    For i = 1 To end_STCGDTAA5
        RateTreatySTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(RateTreatySTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Rate as per Treaty of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateRateTreatySTCG_DTAA = False
         Exit Function
     End If
 Next
End Function
Function ValidateRateITActSTCG_DTAA() As Boolean
 
    ValidateRateITActSTCG_DTAA = True
    Dim rangecells, rangecells1 As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_RateIT").Cells
    Set rangecells1 = Sheet13.Range("STCG.A8_ApplRate").Cells
    ReDim RateITActSTCG_DTAA(end_STCGDTAA5)
    ReDim ApplRate_STCG(end_STCGDTAA5)
    For i = 1 To end_STCGDTAA5
        RateITActSTCG_DTAA(i) = rangecells.item(i).value
        ApplRate_STCG(i) = rangecells1.item(i).value
        
     If Not chkCompulsory(RateITActSTCG_DTAA(i)) Then
    'changed by Chetan C M on 09/09/2025
     'start--
         'msgboxSHEETCG = msgboxSHEETCG + "* Rate as per I.T Act of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         msgboxSHEETCG = msgboxSHEETCG + "* Rate should be within 100 in A9 of CG schedule." & Chr(13)
     '--end
         ValidateRateITActSTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Function ValidateSectionITActSTCG_DTAA() As Boolean
 
    ValidateSectionITActSTCG_DTAA = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("STCG.A8_SecIT").Cells
    ReDim SectionITActSTCG_DTAA(end_STCGDTAA5)
    For i = 1 To end_STCGDTAA5
        SectionITActSTCG_DTAA(i) = rangecells.item(i).value
        
     If Not chkCompulsory(SectionITActSTCG_DTAA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Section of I.T Act of DTAA at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateSectionITActSTCG_DTAA = False
         Exit Function
     End If
 Next
End Function

Sub setTblinfo_STCGDTAA()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_Country").count
 Set rangecells = Sheet13.Range("STCG.A8_Country").Cells
  Dim countrycd As Variant
 For mIntCtr = 1 To mIntCells
     countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
    
     If Not countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA = ccount
 'rngname_STCGDTAA = "STCG.A7_Country;STCG.A7_DTAA_Article;STCG.A7_CertiStatus;STCG.A7_ItemNo;STCG.A7_StcgAmt;STCG.A7_Rate_Treaty;STCG.A7_Section;STCG.A7_Rate_ITAct;STCG.A7_Rate_Applicable"
End Sub

 Sub setTblinfo_STCGDTAA2()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_DTAA_Article").count
 Set rangecells = Sheet13.Range("STCG.A8_DTAA_Article").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA2 = ccount
 End Sub
 
 Sub setTblinfo_STCGDTAA3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim trc As Variant
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_CertiStatus").count
 Set rangecells = Sheet13.Range("STCG.A8_CertiStatus").Cells
 For mIntCtr = 1 To mIntCells
 
  trc = rangecells.item(mIntCtr).value
        If isdropdownblank(trc) Then
            trc = ""
        End If
    
     If Not trc = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA3 = ccount
 End Sub
 
 Sub setTblinfo_STCGDTAA4()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim section As Variant
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_IA1tA6").count
 Set rangecells = Sheet13.Range("STCG.A8_IA1tA6").Cells
 For mIntCtr = 1 To mIntCells
 
  section = rangecells.item(mIntCtr).value
        If isdropdownblank(section) Then
            section = ""
        End If
    
     If Not section = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA4 = ccount
 End Sub
 

 Sub setTblinfo_STCGDTAA5()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_AmtI").count
 Set rangecells = Sheet13.Range("STCG.A8_AmtI").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA5 = ccount
 rngname_STCGDTAA = "STCG.A8_AmtI;STCG.A8_IA1tA6;STCG.A8_Country;STCG.A8_DTAA_Article;STCG.A8_RatePTreaty;STCG.A8_CertiStatus;STCG.A8_SecIT;STCG.A8_RateIT;STCG.A8_ApplRate"
 End Sub
 
 Sub setTblinfo_STCGDTAA6()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_RatePTreaty").count
 Set rangecells = Sheet13.Range("STCG.A8_RatePTreaty").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA6 = ccount
 End Sub
 
 Sub setTblinfo_STCGDTAA7()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_SecIT").count
 Set rangecells = Sheet13.Range("STCG.A8_SecIT").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA7 = ccount
 End Sub
 
 Sub setTblinfo_STCGDTAA8()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet13.Range("STCG.A8_RateIT").count
 Set rangecells = Sheet13.Range("STCG.A8_RateIT").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGDTAA8 = ccount
 End Sub
 

Function ValidateDate(ByVal Date_1 As Variant, ByVal Minyyyymmdd As String, ByVal Maxyyyymmdd As String, ByVal compulsory As Boolean, ByVal caption As String) As String
    Dim mydate_1 As Variant
    
    ValidateDate = ""
    If (Len(Date_1) = 0 And compulsory) Then
        ValidateDate = " must not be blank"
        GoTo exitfunction
    End If
    
    If Not CheckDateddmmyyyy(Date_1) Then
        ValidateDate = " must be a valid dd/mm/yyyy format"
        GoTo exitfunction
    End If
    
    mydate_1 = Dformat(Date_1, "yyyy-mm-dd")
    
    If (Len(mydate_1) > 0 And Len(Minyyyymmdd) > 0) Then
        If Not ChkMinInclusiveDate(mydate_1, Minyyyymmdd) Then
            ValidateDate = "cannot be less than " & Minyyyymmdd
            GoTo exitfunction
        End If
    End If
     
    mydate_1 = Dformat(Date_1, "yyyy-mm-dd")
    
    If (Len(mydate_1) > 0 And Len(Maxyyyymmdd) > 0) Then
        If Not ChkMaxInclusiveDate(mydate_1, Maxyyyymmdd) Then
            ValidateDate = "cannot be greater than " & "2022/03/31"
            GoTo exitfunction
        End If
    End If

exitfunction:
    If Len(ValidateDate) > 0 Then
        ValidateDate = caption & ValidateDate
    End If
End Function

Sub setTblinfo_exemption()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant

    ccount = 0
    mIntCells = Sheet13.Range("DEDN.DeductedSecCode").count
    Set rangecells = Sheet13.Range("DEDN.DeductedSecCode").Cells

    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).value
        If isdropdownblank(scode) Then
            scode = ""
        End If

        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_exemption = ccount
    rngname_exemption = "DEDN.DeductedSecCode;DEDN.AmtDeducted;DEDN.CostofAsset;DEDN.DateofAcquist;DEDN.AmtDeposited;"
End Sub


Function ValidateAuditedSection_1() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAuditedSection_1 = True
    RepSec92EFlag_1 = False
    RepSec92EOtherFlag_1 = False
    
    settblinfo_a_gen1_sc
    Set rangecells = Sheet1.Range("sheet1.AuditedSection").Cells
    ReDim AuditedSection_1(end_a_gen1_sc)
    For i = 1 To end_a_gen1_sc
        AuditedSection_1(i) = rangecells.item(((i - 1) * 7) + 1).value
        If AuditedSection_1(i) = "92E" Then
            RepSec92EFlag_1 = True
        Else
            RepSec92EOtherFlag_1 = True
        End If
    Next
End Function

Function get92eO() As Variant
    get92eO = RepSec92EOtherFlag_1
End Function
Function get92E() As Variant
    get92E = RepSec92EFlag_1
End Function
Function ValidatePAN(panentry As String) As Boolean
    ValidatePAN = True

    If Len(panentry) > 0 Then
        If Not IsNumeric(Mid(panentry, 6, 4)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 1, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 2, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 3, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 4, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 5, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 10, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
    End If
End Function
Function ValidateAadhaar(Aadhaar As String) As Boolean
    ValidateAadhaar = True

   If Aadhaar <> "" Then
        If Not IsNumeric(Aadhaar) Then
            'errmsgAdhaar = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
    
        If Aadhaar = "000000000000" Then
            'errmsgAdhaar = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
        
        If Aadhaar = "111111111111" Then
            'errmsgAdhaar = "is invalid"
            ValidateAadhaar = False
            Exit Function
        End If
            
        If Len(Aadhaar) <> 12 Then
            'errmsgAdhaar = "For Aadhaar Number- 12 digit."
            ValidateAadhaar = False
            Exit Function
        End If
        Dim i, j As Long
        Dim arr As Variant
        'arr = Array("@", "*", "!", "-", "_", "|", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
        arr = Array(".", ",")
        For i = 1 To Len(Aadhaar)
        For j = 0 To UBound(arr)
        If Mid(Aadhaar, i, 1) = arr(j) Then
        ValidateAadhaar = False
            Exit Function
        End If
        Next
        Next
    End If
End Function
Sub settblinfo_a_gen1_sc()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim mintnew As Variant
 
    ccount = 0
    mIntCells = Sheet1.Range("sheet1.AuditedSection").count

 
    Set rangecells = Sheet1.Range("sheet1.AuditedSection").Cells
 
    mintnew = ""
    For mIntCtr = 1 To mIntCells / 7
        mintnew = rangecells.item(((mIntCtr - 1) * 7) + 1).value
        If isdropdownblank(mintnew) Then
            mintnew = ""
        End If
        If Not mintnew = "" Then
            ccount = ccount + 1
        End If
    Next
    end_a_gen1_sc = ccount
    rngname_a_gen1_sc = "sheet1.AuditedSection;sheet1.AuditYN;sheet1.DateOfAudit;"
End Sub




Function ChkMinInclusiveDate(Mininclusive As Variant, Mininclusivedate As Variant) As Boolean
'Note:
'' both date must be in format yyyy-mm-dd
    ChkMinInclusiveDate = True
    If Len(Mininclusive) > 0 Then
        If Mid(Mininclusive, 1, 4) < Mid(Mininclusivedate, 1, 4) Then
            ChkMinInclusiveDate = False
            Exit Function
        Else
            If Mid(Mininclusive, 1, 4) = Mid(Mininclusivedate, 1, 4) Then
                If (Mid(Mininclusive, 6, 2) < Mid(Mininclusivedate, 6, 2)) Then
                    ChkMinInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Mininclusive, 6, 2) = Mid(Mininclusivedate, 6, 2))) Then
                    If (Mid(Mininclusive, 9, 2) < Mid(Mininclusivedate, 9, 2)) Then
                        ChkMinInclusiveDate = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
  
End Function
Function ChkMaxInclusiveDate(Maxinclusive As Variant, Maxinclusivedate As Variant) As Boolean
'Note:
'' both date must be in format yyyy-mm-dd
    
    ChkMaxInclusiveDate = True
    If Len(Maxinclusive) > 0 Then
        If Mid(Maxinclusive, 1, 4) > Mid(Maxinclusivedate, 1, 4) Then
            ChkMaxInclusiveDate = False
            Exit Function
        Else
            If Mid(Maxinclusive, 1, 4) = Mid(Maxinclusivedate, 1, 4) Then
                If (Mid(Maxinclusive, 6, 2) > Mid(Maxinclusivedate, 6, 2)) Then
                    ChkMaxInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Maxinclusive, 6, 2) = Mid(Maxinclusivedate, 6, 2))) Then
                    If (Mid(Maxinclusive, 9, 2) > Mid(Maxinclusivedate, 9, 2)) Then
                        ChkMaxInclusiveDate = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
End Function

Function ValidatesheetSTCGDTAASection() As Boolean
ValidatesheetSTCGDTAASection = True
Msgbox_STCGDTAA = ""
Msgbox_STCGDTAA = "Schedule CG : " & Chr(13) & Chr(13)

If Range("Stcg_DTAA_A1e").value > 0 Then
    If Not Range("STCG.Aaiii_FullConsiderationFiftyC_Temp").value > 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A1aiii should not be zero as there is an entry in DTAA table in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

If Range("STCG_DTAA_A2c").value > 0 Then
    If Range("STCG.A20ia_FullConsideration").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A2a should not be zero as there is an entry in DTAA table. in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If


If Range("Stcg_DTAA_A3ie").value > 0 Then
    If Range("STCG.A2ia_FullConsideration").value = 0 Then 'Not updating anything as Stcg_DTAA_A3ie will always be Zero by Bindu on 11th feb 26
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A3ia should not be zero as there is an entry in DTAA table. in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("STCG_DTAA_A3ie_20").value > 0 Then
    If Range("STCG.A2ia_2eii_FullConsideration").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A3ia where transfer was on or after 23rd july 2024 should not be zero as there is an entry in DTAA table. in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If



If Range("Stcg_DTAA_A3iie").value > 0 Then
    If Range("STCG.A2iia_FullConsideration").value = 0 Then 'Not updating anything as this field Stcg_DTAA_A3iie will alwayz be zero by Bindu on 11th feb 26
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A3iia  should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("STCG_DTAA_A3iie_20").value > 0 Then
    If Range("STCG.A2iia_FullConsideration_ii").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A3iia where transfer was on or after 23rd july 2024 should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If


If Range("Stcg_DTAA_A4a").value > 0 Then
    If Range("STCG.A3a_NRItaxSTTPaid").value = 0 Then 'Not updating anything as thif field Stcg_DTAA_A4a will always be zero by Bindu on12th feb 26
         Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A4ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("STCG_DTAA_A4a_20").value > 0 Then
    If Range("STCG.A3aii_NRItaxSTTPaid").value = 0 Then 'Not updating
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A4aii should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If


If Range("Stcg_DTAA_A4b").value > 0 Then
    If Range("STCG.A3b_NRItaxSTTNotPaid").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A4b should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

If Range("Stcg_DTAA_A5e").value > 0 Then
    If Range("STCG.A4a_FullConsideration").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A5a should not be zero as there is an entry in DTAA table. in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

If Range("Stcg_DTAA_A6f").value > 0 Then
    If Range("STCG.A5a_FullConsideration").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A6f should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

If Range("Stcg_DTAA_A7").value > 0 Then
    If Range("STCG.A7_AmtDeemed").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A7 should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

If Range("Stcg_DTAA_A8a").value > 0 Then
    If Range("STCG.A8_PassThroughIncome15").value = 0 Then 'Not updating as this field  Stcg_DTAA_A8a will always be zero by Bindu on 11th feb 26
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A8a should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If



'For AY 2025-26 by sai on 15/04/2025
 If Range("STCG_DTAA_A8a_20").value > 0 Then
    If Range("STCG.A8_PassThroughIncome20").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A8aii should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
 End If



If Range("Stcg_DTAA_A8b").value > 0 Then
    If Range("STCG.A8_PassThroughIncome30").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A8b should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

If Range("Stcg_DTAA_A8c").value > 0 Then
    If Range("STCG.A8_PassThroughIncome_ApplicableRate").value = 0 Then
        Msgbox_STCGDTAA = Msgbox_STCGDTAA + "* The value in A8c should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetSTCGDTAASection = False
    End If
End If

End Function

Function ValidatesheetLTCGDTAASection() As Boolean
ValidatesheetLTCGDTAASection = True
Msgbox_LTCGDTAA = ""
Msgbox_LTCGDTAA = "Schedule CG : " & Chr(13) & Chr(13)

If Range("Ltcg_DTAA_B1e").value > 0 Then
    If Not Range("LTCG.Baiii_FullConsiderationFiftyC_Temp").value > 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B1aiii should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

If Range("Ltcg_DTAA_B2e").value > 0 Then
    If Range("LTCG.B02a_FullConsideration").value = 0 Then 'Not updating anuthing as this field Ltcg_DTAA_B2e always be zero by Bindu on 12th Fen 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B2a should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

If Range("Ltcg_DTAA_B3e").value > 0 Then
    If Range("LTCG.B2a_FullConsideration").value = 0 Then 'No updating as this field Ltcg_DTAA_B3e will be zero by Bindu on 16th Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B3a should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

If Range("Ltcg_DTAA_B4ie").value > 0 Then 'Will alwayz be zero by Bindu on 16th feb 26
    If Range("LTCG.B3a_FullConsideration").value = 0 Then 'Not updating anything as this field  Ltcg_DTAA_B4ie will be zero by Bindu on 16th Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B4ia should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2024
If Range("LTCG_DTAA_B4ieii_12.5").value > 0 Then
    If Range("LTCG.B3a_FullConsideration_ii").value = 0 Then
        'Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B4ia (where transfer was on or after 23rd july 2024) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
'Updated by Bindu on 23rd Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B3ia should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

If Range("Ltcg_DTAA_B4iie").value > 0 Then 'Not updated as this field will always be zero by Bindu on 18th Feb 26
    If Range("LTCG.Bii3a_FullConsideration").value = 0 Then 'Not updated as this field Ltcg_DTAA_B4iie will always be zero by Bindu on 18th Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B4iia should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2024
If Range("LTCG_DTAA_B4iie_12.5").value > 0 Then
    If Range("LTCG.Bii3a_FullConsideration_ii").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B3iia should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        'Updated by Bindu on 23rd Feb 26
        'Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B4iia (where transfer was on or after 23rd july 2024) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


'new

If Range("Ltcg_DTAA_B5c").value > 0 Then 'Not updated as this field will always be zero by Bindu on 18th Feb 26
    'If Range("LTCG.5.LTCG_AfterThreshold").value = 0 Then
     'Not updated as this field Ltcg_DTAA_B5c will always be zero by Bindu on 18th Feb 26
     If Range("LTCG.5ai_LTCG_AfterThreshold").value = 0 Then 'For AY 2025-26 by sai on 15/04/2025
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B5ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2024
If Range("LTCG_DTAA_B5c_12.5").value > 0 Then
'Updated by Bindu on 23rd Feb 26
    If Range("LTCG.5.LTCG_AfterThreshold").value = 0 Then
    'If Range("LTCG.5aii_LTCG_AfterThreshold").value = 0 Then
    'Updated from B5aii to B4a by Bindu on 23rd Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B4a should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


If Range("Ltcg_DTAA_B6c").value > 0 Then ''Not updated as this field will always be zero by Bindu on 18th Feb 26
    'If Range("LTCG.B4a_LTCGWithoutBenefit").value = 0 Then
     If Range("LTCG.B4ai_LTCGWithoutBenefit").value = 0 Then 'For AY 2025-26 by sai on 15/04/2025'
     'Not updated as this field"Ltcg_DTAA_B6c" will always be zero by Bindu on 18th Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B6ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
'*********************************************************************************************
If Range("LTCG_DTAA_B6cii_10").value > 0 Then ''Not updated as this field will always be zero by Bindu on 18th Feb 26
     If Range("LTCG.B4aii_LTCGWithoutBenefit").value = 0 Then ''Not updated as this field"LTCG_DTAA_B6cii_10" will always be zero by Bindu on 18th Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B6aii should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

If Range("LTCG_DTAA_B6ciii_12.5").value > 0 Then
     If Range("LTCG.B4aiii_LTCGWithoutBenefit").value = 0 Then
     'Updated from B6aiii to B5a by Bindu on 23rd Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B5a should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If
'*********************************************************************************************


If Range("Ltcg_DTAA_B7ie").value > 0 Then 'Not updating as this field will always be zero by Bindu on 18th Feb 26
    If Range("LTCG.B6aia_FullConsideration").value = 0 Then 'Not updating as this field "Ltcg_DTAA_B7ie" will always be zero by Bindu on 18th feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B7ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B7ie_12.5").value > 0 Then
    If Range("LTCG.B6aia_FullConsideration_ii").value = 0 Then
        'Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B7ai (where transfer was on or after 23rd july 2024) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        'Updated by Binduon 23rd Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B6ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)

        ValidatesheetLTCGDTAASection = False
    End If
End If



If Range("Ltcg_DTAA_B7iie").value > 0 Then ''Not updating as this field will always be zero by Bindu on 18th Feb 26
    If Range("LTCG.B6ibia_FullConsideration").value = 0 Then ''Not updating as this field "Ltcg_DTAA_B7iie" will always be zero by Bindu on 18th Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B7iiai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B7iie_12.5").value > 0 Then
    If Range("LTCG.B6ibia_FullConsideration_ii").value = 0 Then
        'Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B7iiai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        'Updated by Bindu on 23rdFeb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B6iiai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)

        ValidatesheetLTCGDTAASection = False
    End If
End If

If Range("Ltcg_DTAA_B7iiie").value > 0 Then
    If Range("LTCG.B6iiiaia_FullConsideration").value = 0 Then
        'Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B7iiiai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        'Updated by Binduon 23rd Feb 26
         Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B6iiiai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


''Not updating anything as this field "Ltcg_DTAA_B8c"  willbe zero by Bindu on 23rd Feb 26
'If Range("Ltcg_DTAA_B8c").value > 0 Then
''    If Range("LTCG.8.LTCG_AfterThreshold").value = 0 Then
'     If Range("LTCG.8.ai_LTCG_AfterThreshold").value = 0 Then 'For AY 2025-26 by sai on 15/04/2025
'        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B8ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
'        ValidatesheetLTCGDTAASection = False
'    End If
'End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B8c_12.5").value > 0 Then
'    If Range("LTCG.8.aii_LTCG_AfterThreshold").value = 0 Then 'Updated byBindu on 23rd Feb 26
    If Range("LTCG.8.LTCG_AfterThreshold").value = 0 Then
    'Updated by Bindu from B8aii to B7a on 23rd Feb26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B7a should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


If Range("Ltcg_DTAA_B9c").value > 0 Then 'Not updating anything as this field "Ltcg_DTAA_B9c" will always be zero by Bindu on 18th feb 26
'    If Range("LTCG.B6a_SaleonSpecAsset").value = 0 Then
'Not updating anything as this field "Ltcg_DTAA_B9c" will always be zero by Bindu on 18th feb 26
     If Range("LTCG.B6ai_SaleonSpecAsset").value = 0 Then 'For AY 2025-26 by sai on 15/04/2025
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B9ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B9c_12.5").value > 0 Then
     If Range("LTCG.B6aii_SaleonSpecAsset").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B9aii should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If




'If Range("Ltcg_DTAA_B9f").value > 0 Then
'    If Range("LTCG.B6d_SaleOtherSpecAsset").value = 0 Then
'        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B9d should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
'        ValidatesheetLTCGDTAASection = False
'    End If
'End If

If Range("Ltcg_DTAA_B10e").value > 0 Then
    If Range("LTCG.B8aia_FullConsideration").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B10ai should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B10e_12.5").value > 0 Then
    If Range("LTCG.B8aia_FullConsideration_ii").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B10ai (where transfer was on or after 23rd july 2024) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If



If Range("Ltcg_DTAA_B11").value > 0 Then  'Notupdating anything as this field  will always be zero by Bindu on 22ndFeb 26
    If Range("LTCG.B9_AmtDeemed").value = 0 Then ''Notupdating anything as this field "Ltcg_DTAA_B11" will always be zero by Bindu on 22ndFeb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B11i should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B11_12.5").value > 0 Then
    If Range("LTCG.B9ii_AmtDeemed").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B11ii should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


If Range("Ltcg_DTAA_B12a1").value > 0 Then 'Not updating as this field will alwyaz be Zero by Bindu on 22nd Feb 26
    If Range("LTCG.B12_PassThroughIncome10_112A").value = 0 Then 'Not updating as this field "Ltcg_DTAA_B12a1" will be alwayz zero by Bindu on 22nd Feb 26
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B12a1(i) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B12a1_12.5").value > 0 Then
    If Range("LTCG.B12ii_PassThroughIncome10_112A").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B12a1(ii) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


If Range("Ltcg_DTAA_B12a2").value > 0 Then
    If Range("LTCG.B12_PassThroughIncome10").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B12a2(i) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If

'For AY 2025-26 by sai on 15/04/2025
If Range("LTCG_DTAA_B12a2_12.5").value > 0 Then
    If Range("LTCG.B12ii_PassThroughIncome10").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B12a2(ii) should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


If Range("Ltcg_DTAA_B12b").value > 0 Then
    If Range("LTCG.B12_PassThroughIncome20").value = 0 Then
        Msgbox_LTCGDTAA = Msgbox_LTCGDTAA + "* The value in B12b should not be zero as there is an entry in DTAA table.in schedule CG" & Chr(13)
        ValidatesheetLTCGDTAASection = False
    End If
End If


End Function

Sub AddBlockCall_STCGrptfrm()
    setTblinfo_STCGrptfrm
    Call EfilingCommon.addblock(rngname_STCGrptfrm, frmRngname_STCGrptfrm, cntrRng_STCGrptfrm, frmsize_STCGrptfrm)
    
    Dim blockcount As Variant
    Dim address1 As Variant
    Dim address2 As Variant
    Dim cellrange As Range
    Dim formula As Variant
    
    blockcount = Sheet13.Range("PropertySectionCount_STCG").value
    address1 = Sheet13.Range("STCG.Aai_FullConsideration" & blockcount).Address
    address2 = Sheet13.Range("STCG.Aaii_PropertyValuation" & blockcount).Address
    Set cellrange = Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC" & blockcount)
    formula = "=" & address1 & ":" & address2
    
     sPassword = EfilingCommon.getmsgstate
    Sheet13.Unprotect Password:=sPassword
    
    With cellrange.Validation
        .Delete
        .add Type:=xlValidateList, _
         AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, _
         Formula1:=formula
    End With
    Sheet13.Range("STCG_SLNo" & blockcount).value = blockcount
    Sheet13.Protect Password:=sPassword
    
     'new change
        Dim rangecells As Range
        Dim mIntCells  As Long
        Dim mIntCtr  As Long
        Dim ccount  As Long
        ccount = 0
        Set rangecells = Sheet13.Range("CG_A_AddressOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_A_AddressOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_A_AddressOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_A_AddressOfBuyer" & blockcount).Cells.item(mIntCtr).ClearContents
            End If
        Next
        Set rangecells = Sheet13.Range("CG_A_PinCodeOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_A_PinCodeOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_A_PinCodeOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_A_PinCodeOfBuyer" & blockcount).Cells.item(mIntCtr).ClearContents
            End If
        Next
        Set rangecells = Sheet13.Range("CG_A_ZipcodeOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_A_ZipcodeOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_A_ZipcodeOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_A_ZipcodeOfBuyer" & blockcount).Cells.item(mIntCtr).ClearContents
            End If
        Next
        Set rangecells = Sheet13.Range("CG_A_StateCodeOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_A_StateCodeOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_A_StateCodeOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_A_StateCodeOfBuyer" & blockcount).Cells.item(mIntCtr).value = "(Select)"
            End If
        Next
        Set rangecells = Sheet13.Range("CG_A_CountryOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_A_CountryOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_A_CountryOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_A_CountryOfBuyer" & blockcount).Cells.item(mIntCtr).value = "(Select)"
            End If
        Next
    'new change
    
End Sub

Sub setTblinfo_STCGrptfrm()
    Dim te As Long
    Dim ccount As Long
    
    frmsize_STCGrptfrm = Sheet13.Range("NumRowsPropertyBlock_STCG").value

    end_STCGrptfrm = Sheet13.Range("PropertySectionCount_STCG").value
    cntrRng_STCGrptfrm = "PropertySectionCount_STCG"
    frmRngname_STCGrptfrm = "CG_STCG_rptfrm"
    'rngname_STCGrptfrm = "STCG_SLNo1||STCG.A_AcqDate1||STCG.AB_SaleTrfDate1||STCG.Aai_FullConsideration1||STCG.Aaii_PropertyValuation1||STCG.Aaiii_FullConsiderationFiftyC1||STCG.Abi_AquisitCost1||STCG.Abii_ImproveCost1||STCG.Abiii_ExpOnTrans1||STCG.Abiv_TotalDedn1||STCG.Ac_BalanceCG1||STCG.Adi_ExemptionAmount1||STCG.Adii_ExemptionAmount1||STCG.Adiii_ExemptionAmount1||STCG.Adiv_ExemptionAmount1||STCG.Ad_ExemptionGrandTotal1||STCG.Ae_CapgainonAssets1||CG_A_NameOfBuyer1||CG_A_PanOfBuyer1||CG_A_AadhaarOfBuyer1||CG_A_PercentageShare1||CG_A_Amount1||CG_A_AddressOfBuyer1||CG_A_PinCodeOfBuyer1||CG_A_StateCodeOfBuyer1||CG_A_CountryOfBuyer1||CG_A_ZipcodeOfBuyer1"
    rngname_STCGrptfrm = "STCG_SLNo1||STCG.A_AcqDate1||STCG.AB_SaleTrfDate1||STCG.Aai_FullConsideration1||STCG.Aaii_PropertyValuation1||STCG.Aaiii_FullConsiderationFiftyC1||STCG.Abi_AquisitCost1||STCG.Abii_ImproveCost1||STCG.Abiii_ExpOnTrans1||STCG.Abiv_TotalDedn1||STCG.Ac_BalanceCG1||STCG.Adi_ExemptionAmount1||STCG.Adiii_ExemptionAmount1||STCG.Adiv_ExemptionAmount1||STCG.Ad_ExemptionGrandTotal1||STCG.Ae_CapgainonAssets1||CG_A_NameOfBuyer1||CG_A_PanOfBuyer1||CG_A_AadhaarOfBuyer1||CG_A_PercentageShare1||CG_A_Amount1||CG_A_AddressOfBuyer1||CG_A_PinCodeOfBuyer1||CG_A_StateCodeOfBuyer1||CG_A_CountryOfBuyer1||CG_A_ZipcodeOfBuyer1"
 
 End Sub
Sub AddRowSTCGA1_click()
AddSTCGA1
End Sub

Sub AddSTCGA1(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet13.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet13.Range("PropertySectionCount_STCG").value
    
        
        EfilingCommon.DefinedgridNameRange = "CG_A_NameOfBuyer||CG_A_PanOfBuyer||CG_A_AadhaarOfBuyer||CG_A_PercentageShare||CG_A_Amount||CG_A_AddressOfBuyer||CG_A_PinCodeOfBuyer||CG_A_StateCodeOfBuyer||CG_A_CountryOfBuyer||CG_A_ZipcodeOfBuyer"
        ActiveCellRange = EfilingCommon.searchLastRow("CG_A_NameOfBuyer" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False

    newfrmsize = Sheet13.Range("NumRowsPropertyBlock_STCG").value
        newfrmsize = newfrmsize + numberofrows
        Sheet13.Range("NumRowsPropertyBlock_STCG").value = newfrmsize

endline:
   
'----------------Lock Password-------------------START---
   Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub

'LTCG Block

Sub AddBlockCall_LTCGrptfrm()
    setTblinfo_LTCGrptfrm
    Call EfilingCommon.addblock(rngname_LTCGrptfrm, frmRngname_LTCGrptfrm, cntrRng_LTCGrptfrm, frmsize_LTCGrptfrm)
   Dim blockcount As Variant
    Dim address1 As Variant
    Dim address2 As Variant
    Dim cellrange As Range
    Dim formula As Variant
    
    blockcount = Sheet13.Range("PropertySectionCount_LTCG").value
    
'    'Newly added by sai on 19/02/2025 for AY 2025-26========================================================================
'     sPassword = EfilingCommon.getmsgstate
'     Sheet13.Unprotect Password:=sPassword
'     Sheet13.Range("LTCG.Bbi_AquisitCostTot" & blockcount).formula = "=SUM(LTCG.Bbii_ImproveCost" & blockcount & ")"
'    '=======================================================================================================================
'
'    'by Sai on 11/03/2025 as per DE V3
'     Sheet13.Unprotect Password:=sPassword
'     Sheet13.Range("LTCG.Bbia_ImproveCost_Sum" & blockcount).formula = "=SUM(LTCG.Bbia_ImproveCost" & blockcount & ")"
'     Sheet13.Range("LTCG.Bbii_ImproveCost_Sum" & blockcount).formula = "=SUM(LTCG.Bbii_ImproveCost" & blockcount & ")"
    
        'Malli added
     sPassword = EfilingCommon.getmsgstate
     Sheet13.Unprotect Password:=sPassword
'Commented by Riyaz on 17/03/2025 for 2025-26
'     Sheet13.Range("LTCG.Bbi_AquisitCostTot" & blockcount).formula = "=SUM(LTCG.Bbii_ImproveCost" & blockcount & ")"
      Sheet13.Range("LTCG.Bbia_ImproveCost_Sum" & blockcount).formula = "=SUM(LTCG.Bbia_ImproveCost" & blockcount & ")"
'--------------
    
    'Newly added by Bindu on 11th March 25 as per DE V3
     Sheet13.Unprotect Password:=sPassword
     Sheet13.Range("LTCG.Bbii_ImproveCost_Sum" & blockcount).formula = "=SUM(LTCG.Bbii_ImproveCost" & blockcount & ")"
    
    address1 = Sheet13.Range("LTCG.Bai_FullConsideration" & blockcount).Address
    address2 = Sheet13.Range("LTCG.Baii_PropertyValuation" & blockcount).Address
    Set cellrange = Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & blockcount)
    formula = "=" & address1 & ":" & address2
    
     sPassword = EfilingCommon.getmsgstate
    Sheet13.Unprotect Password:=sPassword
    
'    With cellrange.Validation
'        .Delete
'        .add Type:=xlValidateList, _
'         AlertStyle:=xlValidAlertStop, _
'         Operator:=xlBetween, _
'         Formula1:=formula
'    End With
    
'Change-17.04.2023.101.CG -> Formula Issue
    'To set LTCG.Bbi_AquisitCostTot2 formula to =SUM(LTCG.Bbii_ImproveCost2)
'    Set cellrange = Sheet13.Range("LTCG.Bbi_AquisitCostTot" & blockcount)
'    formula = "=SUM(LTCG.Bbii_ImproveCost" & blockcount & ")"
'    cellrange.formula = formula
'    Sheet13.Unprotect Password:=sPassword
'-------

    Sheet13.Range("LTCG_SLNo" & blockcount).value = blockcount
    Sheet13.Protect Password:=sPassword
    
 'new change
        Dim rangecells As Range
        Dim mIntCells  As Long
        Dim mIntCtr  As Long
        Dim ccount  As Long
        ccount = 0
         Set rangecells = Sheet13.Range("CG_B_AddressOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_B_AddressOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_B_AddressOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_B_AddressOfBuyer" & blockcount).Cells.item(mIntCtr).ClearContents
            End If
        Next
        Set rangecells = Sheet13.Range("CG_B_PinCodeOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_B_PinCodeOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_B_PinCodeOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_B_PinCodeOfBuyer" & blockcount).Cells.item(mIntCtr).ClearContents
            End If
        Next
        Set rangecells = Sheet13.Range("CG_B_ZipcodeOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_B_ZipcodeOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_B_ZipcodeOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_B_ZipcodeOfBuyer" & blockcount).Cells.item(mIntCtr).ClearContents
            End If
        Next
        Set rangecells = Sheet13.Range("CG_B_StateCodeOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_B_StateCodeOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_B_StateCodeOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_B_StateCodeOfBuyer" & blockcount).Cells.item(mIntCtr).value = "(Select)"
            End If
        Next
        Set rangecells = Sheet13.Range("CG_B_CountryOfBuyer" & blockcount).Cells
        mIntCells = Sheet13.Range("CG_B_CountryOfBuyer" & blockcount).count
        For mIntCtr = 1 To mIntCells
            If Sheet13.Range("CG_B_CountryOfBuyer" & blockcount).Cells.item(mIntCtr).Locked = False Then
                Sheet13.Range("CG_B_CountryOfBuyer" & blockcount).Cells.item(mIntCtr).value = "(Select)"
            End If
        Next
        
        
        
        
    'new change
    
    'Newly added by sai on 20/03/2025
    Worksheets("CG").Cost_of_Improvement
    Application.EnableEvents = True
    Sheet13.Protect Password:=EfilingCommon.getmsgstate
    
End Sub
'Newly added by sai on 11/03/2025
Sub AddRowsCOST()
Dim vRows As Long
Sheets("CG").Activate
' LTCG.Bbib_ImproveCostYear_Vlookup1
EfilingCommon.DefinedgridNameRange = "LTCG.Bbia_ImproveCost1||LTCG.Bbib_ImproveCostYear1||LTCG.Bbii_ImproveCost1||LTCG.Bbib_ImproveCostYear_Vlookup1"
ActiveCellRange = EfilingCommon.searchLastRow("LTCG.Bbii_ImproveCost1")
vRows = EfilingCommon.insertRowUnderSectionWithFormula_CG
 
End Sub


Sub setTblinfo_LTCGrptfrm()
    Dim te As Long
    Dim ccount As Long
    
     frmsize_LTCGrptfrm = Sheet13.Range("NumRowsPropertyBlock_LTCG").value

    end_LTCGrptfrm = Sheet13.Range("PropertySectionCount_LTCG").value
    cntrRng_LTCGrptfrm = "PropertySectionCount_LTCG"
   frmRngname_LTCGrptfrm = "CG_LTCG_rptfrm"
    'rngname_LTCGrptfrm = "LTCG_SLNo1||LTCG.B_AcqDate1||LTCG.B_SaleTrfDate1||LTCG.Bai_FullConsideration1||LTCG.Baii_PropertyValuation1||LTCG.Baiii_FullConsiderationFiftyC1||LTCG.Bbi_AquisitCost1||LTCG.Bbii_ImproveCost1||LTCG.Bbiii_ExpOnTrans1||LTCG.Bbiv_TotalDedn1||LTCG.Bc_BalanceCG1||LTCG.Bdi_ExemptionAmount1||LTCG.Bdii_ExemptionAmount1||LTCG.Bdiii_ExemptionAmount1||LTCG.Bdiv_ExemptionAmount1||LTCG.Bdv_ExemptionAmount1||LTCG.Bdvi_ExemptionAmount1||LTCG.Bdvii_ExemptionAmount1||LTCG.Bdviii_ExemptionAmount1||LTCG.Bd_ExemptionGrandTotal1||LTCG.Be_CapgainonAssets1||CG_B_NameOfBuyer1||CG_B_PanOfBuyer1||CG_B_AadhaaeOfBuyer1||CG_B_PercentageShare1||CG_B_Amount1||CG_B_AddressOfBuyer1||CG_B_PinCodeOfBuyer1||CG_B_StateCodeOfBuyer1||CG_B_CountryOfBuyer1||CG_B_ZipcodeOfBuyer1"
    'rngname_LTCGrptfrm = "LTCG_SLNo1||LTCG.B_AcqDate1||LTCG.B_SaleTrfDate1||LTCG.Bai_FullConsideration1||LTCG.Baii_PropertyValuation1||LTCG.Baiii_FullConsiderationFiftyC1||LTCG.Bbi_AquisitCostNew1||LTCG.Bbi_AquisitCost1||LTCG.Bbii_ImproveCost1||LTCG.Bbi_AquisitCostTot1||LTCG.Bbia_ImproveCost1||LTCG.Bbib_ImproveCostYear1||LTCG.Bbii_ImproveCost1||LTCG.Bbiii_ExpOnTrans1||LTCG.Bbiv_TotalDedn1||LTCG.Bc_BalanceCG1||LTCG.Bdi_ExemptionAmount1||LTCG.Bdii_ExemptionAmount1||LTCG.Bdiii_ExemptionAmount1||LTCG.Bdiv_ExemptionAmount1||LTCG.Bdv_ExemptionAmount1||LTCG.Bdvi_ExemptionAmount1||LTCG.Bdvii_ExemptionAmount1||LTCG.Bdviii_ExemptionAmount1||LTCG.Bd_ExemptionGrandTotal1||LTCG.Be_CapgainonAssets1||CG_B_NameOfBuyer1||CG_B_PanOfBuyer1||CG_B_AadhaaeOfBuyer1||CG_B_PercentageShare1||CG_B_Amount1||CG_B_AddressOfBuyer1||CG_B_PinCodeOfBuyer1||CG_B_StateCodeOfBuyer1||CG_B_CountryOfBuyer1||CG_B_ZipcodeOfBuyer1"
    
'     rngname_LTCGrptfrm = "LTCG_SLNo1||LTCG.B_AcqDate1||LTCG.B_SaleTrfDate1||LTCG.Bai_FullConsideration1||LTCG.Baii_PropertyValuation1||LTCG.Baiii_FullConsiderationFiftyC1||LTCG.Bbi_AquisitCostNew1||LTCG.Bbi_AquisitCost1||LTCG.Bbii_ImproveCost1||LTCG.Bbi_AquisitCostTot1||LTCG.Bbiii_ExpOnTrans1||LTCG.Bbiv_TotalDedn1||LTCG.Bc_BalanceCG1||LTCG.Bdi_ExemptionAmount1||LTCG.Bdii_ExemptionAmount1||LTCG.Bdiii_ExemptionAmount1||LTCG.Bdiv_ExemptionAmount1||LTCG.Bdv_ExemptionAmount1||LTCG.Bdvi_ExemptionAmount1||LTCG.Bdvii_ExemptionAmount1||LTCG.Bd_ExemptionGrandTotal1||LTCG.Be_CapgainonAssets1||LTCG.Bbia_ImproveCost1||LTCG.Bbib_ImproveCostYear1||LTCG.Bbii_ImproveCost1||CG_B_NameOfBuyer1||CG_B_PanOfBuyer1||CG_B_AadhaaeOfBuyer1||CG_B_PercentageShare1||CG_B_Amount1||CG_B_AddressOfBuyer1||CG_B_PinCodeOfBuyer1||CG_B_StateCodeOfBuyer1||CG_B_CountryOfBuyer1||CG_B_ZipcodeOfBuyer1"

'    'newly modified by sai on 11/03/2025
'      rngname_LTCGrptfrm = "LTCG_SLNo1||LTCG.B_AcqDate1||LTCG.B_SaleTrfDate1||LTCG.Bai_FullConsideration1||LTCG.Baii_PropertyValuation1||LTCG.Baiii_FullConsiderationFiftyC1||LTCG.Bbi_AquisitCostNew1||LTCG.Bbi_AquisitCost1||LTCG.Bbii_ImproveCost1||LTCG.Bbi_AquisitCostTot1||LTCG.Bbiii_ExpOnTrans1||LTCG.Bbiv_TotalDedn1||LTCG.Bbiva_TotalDedn1||LTCG.Bca_BalanceCG1||LTCG.Bc_BalanceCG1||LTCG.Bdi_ExemptionAmount1||LTCG.Bdii_ExemptionAmount1||LTCG.Bdiii_ExemptionAmount1||LTCG.Bdiv_ExemptionAmount1||LTCG.Bdv_ExemptionAmount1||LTCG.Bdvi_ExemptionAmount1||LTCG.Bdvii_ExemptionAmount1||LTCG.Bd_ExemptionGrandTotal1||LTCG.Be_CapgainonAssets1||LTCG.Bea_CapgainonAssets1||LTCG.BeiA_TaxSec1||LTCG.BeiB_TaxSec1||LTCG.Beii_Excessamt1||LTCG.Bbib_ImproveCostYear_Vlookup1" & _
'"LTCG.Bbia_ImproveCost_Sum1||LTCG.Bbii_ImproveCost_Sum1||LTCG.Bbia_ImproveCost1||LTCG.Bbib_ImproveCostYear1||LTCG.Bbii_ImproveCost1||CG_B_NameOfBuyer1||CG_B_PanOfBuyer1||CG_B_AadhaaeOfBuyer1||CG_B_PercentageShare1||CG_B_Amount1||CG_B_AddressOfBuyer1||CG_B_PinCodeOfBuyer1||CG_B_StateCodeOfBuyer1||CG_B_CountryOfBuyer1||CG_B_ZipcodeOfBuyer1||"
     
   'Malli_________________26/03/2025
rngname_LTCGrptfrm = "LTCG_SLNo1||LTCG.B_AcqDate1||LTCG.B_SaleTrfDate1||LTCG.Bai_FullConsideration1||LTCG.Baii_PropertyValuation1||LTCG.Baiii_FullConsiderationFiftyC1||LTCG.Bbi_AquisitCostNew1||LTCG.Bbi_AquisitCost1||LTCG.Bbii_ImproveCost1||LTCG.Bbi_AquisitCostTot1||LTCG.Bbiii_ExpOnTrans1||LTCG.Bbiv_TotalDedn1||LTCG.Bbiva_TotalDedn1||LTCG.Bca_BalanceCG1||LTCG.Bc_BalanceCG1||LTCG.Bdi_ExemptionAmount1||LTCG.Bdii_ExemptionAmount1||LTCG.Bdiii_ExemptionAmount1||LTCG.Bdiv_ExemptionAmount1||LTCG.Bdv_ExemptionAmount1||LTCG.Bdvi_ExemptionAmount1||LTCG.Bdvii_ExemptionAmount1||LTCG.Bd_ExemptionGrandTotal1||LTCG.Be_CapgainonAssets1||LTCG.Bea_CapgainonAssets1||LTCG.BeiA_TaxSec1||LTCG.BeiB_TaxSec1||LTCG.Beii_Excessamt1||LTCG.Bbib_ImproveCostYear_Vlookup1||" & _
"LTCG.Bbia_ImproveCost_Sum1||LTCG.Bbii_ImproveCost_Sum1||LTCG.Bbia_ImproveCost1||LTCG.Bbib_ImproveCostYear1||LTCG.Bbii_ImproveCost1||CG_B_NameOfBuyer1||CG_B_PanOfBuyer1||CG_B_AadhaaeOfBuyer1||CG_B_PercentageShare1||CG_B_Amount1||CG_B_AddressOfBuyer1||CG_B_PinCodeOfBuyer1||CG_B_StateCodeOfBuyer1||CG_B_CountryOfBuyer1||CG_B_ZipcodeOfBuyer1||"

     
 End Sub

Sub AddRowLTCGB1_click()
AddLTCGB1
End Sub
Sub AddRowLTCGB12_click()
AddLTCGB12
End Sub
Sub AddLTCGB12(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet13.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
'    For te = 1 To Sheet13.Range("PropertySectionCount1_LTCG").value
     For te = 1 To Sheet13.Range("PropertySectionCount_LTCG").value  'by sai on 01/04/2025
    
        
'        EfilingCommon.DefinedgridNameRange = "LTCG.Bbia_ImproveCost||LTCG.Bbib_ImproveCostYear||LTCG.Bbii_ImproveCost"
'by sai on 11/03/2025 For AY 2024-25
        EfilingCommon.DefinedgridNameRange = "LTCG.Bbia_ImproveCost||LTCG.Bbib_ImproveCostYear||LTCG.Bbii_ImproveCost||LTCG.Bbib_ImproveCostYear_Vlookup"
        ActiveCellRange = EfilingCommon.searchLastRow("LTCG.Bbia_ImproveCost" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    'For i = 1 To Sheet13.Range("PropertySectionCount1_LTCG").value
        newfrmsize = Sheet13.Range("NumRowsPropertyBlock_LTCG").value
        newfrmsize = newfrmsize + numberofrows
        Sheet13.Range("NumRowsPropertyBlock_LTCG").value = newfrmsize
   ' Next
    
  
endline:
   
'----------------Lock Password-------------------START---
   Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
Sub AddLTCGB1(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheet13.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet13.Range("PropertySectionCount_LTCG").value
    
        
        EfilingCommon.DefinedgridNameRange = "CG_B_NameOfBuyer||CG_B_PanOfBuyer||CG_B_AadhaaeOfBuyer||CG_B_PercentageShare||CG_B_Amount||CG_B_AddressOfBuyer||CG_B_PinCodeOfBuyer||CG_B_StateCodeOfBuyer||CG_B_CountryOfBuyer||CG_B_ZipcodeOfBuyer"
        ActiveCellRange = EfilingCommon.searchLastRow("CG_B_NameOfBuyer" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    'For i = 1 To Sheet13.Range("PropertySectionCount_LTCG").value
        newfrmsize = Sheet13.Range("NumRowsPropertyBlock_LTCG").value
        newfrmsize = newfrmsize + numberofrows
        Sheet13.Range("NumRowsPropertyBlock_LTCG").value = newfrmsize
   ' Next
    
  
endline:
   
'----------------Lock Password-------------------START---
   Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'new STCG code
Sub setTableInfo_STCG_A1_1(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_NameOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_NameOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_Count1 = ccount
End Sub

Sub setTableInfo_STCG_A1_2(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_PanOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_PanOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_Count2 = ccount
End Sub
Sub setTableInfo_STCG_A1_A(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_AadhaarOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_AadhaarOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_CountA = ccount
End Sub

Sub setTableInfo_STCG_A1_3(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_PercentageShare" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_PercentageShare" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_Count3 = ccount
End Sub

Sub setTableInfo_STCG_A1_4(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_Amount" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_Amount" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_Count4 = ccount
End Sub

Sub setTableInfo_STCG_A1_5(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_AddressOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_AddressOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_Count5 = ccount
End Sub

Sub setTableInfo_STCG_A1_6(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_A_PinCodeOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_A_PinCodeOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    STCG_A1_Count6 = ccount
End Sub
 Function STCGBlock_Validation()
Dim STCG_A1_Count As Variant
STCG_A1_Count = Sheet13.Range("PropertySectionCount_STCG").value
'Dim i As Long
Dim count As Variant
Dim FullValueConsideration_1ai As Variant
Dim PropertyStampValue_1aii As Variant
Dim FullValAsPer50C_1aiii As Variant

ReDim Stcg_TotalPercentage(STCG_A1_Count)
For count = 1 To STCG_A1_Count

FullValueConsideration_1ai = Sheet13.Range("STCG.Aai_FullConsideration" & count).value
PropertyStampValue_1aii = Sheet13.Range("STCG.Aaii_PropertyValuation" & count).value
FullValAsPer50C_1aiii = Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC" & count).value

If ((FullValueConsideration_1ai <> "" Or FullValueConsideration_1ai <> 0) And (PropertyStampValue_1aii = 0 Or PropertyStampValue_1aii = "")) Then

    MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please Enter Value of property as per Stamp Valuation in Short term capital gain in schedule CG " & Chr(13)
    
End If


'Added by sai on 16/03/2025
'*****************************************************
If (Len(Sheet13.Range("STCG.AII_CapitalLoss").value) > 15) Then
    MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Capital Loss on buy back of shares on or after 01st October 2024 [ Short Term 20% or 30% or applicable rate] Amount cannot exceed 14 digits. in Short term capital gain in schedule CG " & Chr(13)
End If
'******************************************************************************


If (PropertyStampValue_1aii <> "" Or PropertyStampValue_1aii <> 0) And (FullValueConsideration_1ai = 0 Or FullValueConsideration_1ai = "") Then

   MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please Enter  Full Value of Consideration Received/Receivable in Short term capital gain in schedule CG" & Chr(13)
   
End If


If ((FullValAsPer50C_1aiii <> FullValueConsideration_1ai) And (FullValAsPer50C_1aiii <> PropertyStampValue_1aii)) Then
 MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please Ensure STCG  (1)aiii is either ai or aii (Selct from Drop Down) in schedule CG" & Chr(13)
End If


If Sheet13.Range("STCG.Abiv_TotalDedn" & count).value > 0 Then
    If Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC" & count).value = 0 Then
        'Added by Aavula on 02/08/2024
        MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero" & Chr(13)
    End If
End If

    setTableInfo_STCG_A1_1 (count)
    setTableInfo_STCG_A1_2 (count)
    setTableInfo_STCG_A1_3 (count)
    setTableInfo_STCG_A1_4 (count)
    setTableInfo_STCG_A1_5 (count)
    setTableInfo_STCG_A1_6 (count)
    setTableInfo_STCG_A1_A (count)
    
STCG_A1_Count1 = WorksheetFunction.Max(0, STCG_A1_Count1, STCG_A1_Count2, STCG_A1_CountA, STCG_A1_Count3, STCG_A1_Count4, STCG_A1_Count5, STCG_A1_Count6)
ReDim STCG_NameOfBuyer(STCG_A1_Count1), STCG_PANOfBuyer(STCG_A1_Count1), STCG_AadhaarOfBuyer(STCG_A1_Count1), STCG_PercentageShare(STCG_A1_Count1), STCG_Amount(STCG_A1_Count1), STCG_AddressOfProperty(STCG_A1_Count1), STCG_PinCode(STCG_A1_Count1), STCG_StateCode(STCG_A1_Count1)
Dim i As Long
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim cellrange As Variant
Dim cellrange1 As Variant
Dim cellrange2 As Variant

If FullValAsPer50C_1aiii > 0 Then
If STCG_A1_Count1 = 0 Then

 MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* In case of transfer of immovable property please furnish the following details Table in Short term capital gain is mandatory in schedule CG " & Chr(13)

End If
End If

'Change-26.04.2023.101.CG -> SIT-49129
'If STCG_A1_Count1 > 0 Then
'If (FullValueConsideration_1ai = 0 Or FullValueConsideration_1ai = "") Then
'MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please Enter Full value of consideration received/receivable in Short term capital gain in schedule CG" & Chr(13)
'End If
'End If

If (FullValueConsideration_1ai > 0 Or PropertyStampValue_1aii > 0 Or FullValAsPer50C_1aiii > 0) Then
    If (Sheet13.Range("STCG.AB_SaleTrfDate" & count).value) = "" Then
        MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please select the date of Sale  in order to calculate the type of  capital Gain" & Chr(13)
    End If
    If (Sheet13.Range("STCG.A_AcqDate" & count).value) = "" Then
        MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please select the date of purchase in order to calculate the type of  capital Gain" & Chr(13)
    End If
End If
'-----

For i = 1 To STCG_A1_Count1
Set rangecells = Sheet13.Range("CG_A_NameOfBuyer" & count).Cells
Set rangecells1 = Sheet13.Range("CG_A_PanOfBuyer" & count).Cells
Set rangecells2 = Sheet13.Range("CG_A_AadhaarOfBuyer" & count).Cells
        cellrange = GetMergedAddressCell(rangecells, i)
        STCG_NameOfBuyer(i) = Sheet13.Range(cellrange).value
        cellrange1 = GetMergedAddressCell(rangecells1, i)
        STCG_PANOfBuyer(i) = Sheet13.Range(cellrange1).value
        cellrange2 = GetMergedAddressCell(rangecells2, i)
        STCG_AadhaarOfBuyer(i) = Sheet13.Range(cellrange2).value
        
'STCG_NameOfBuyer(i) = Sheet13.Range("CG_A_NameOfBuyer" & Count).Cells.item(i).value
'STCG_PANOfBuyer(i) = Sheet13.Range("CG_A_PanOfBuyer" & Count).Cells.item(i).value
STCG_PercentageShare(i) = Sheet13.Range("CG_A_PercentageShare" & count).Cells.item(i).value
STCG_Amount(i) = Sheet13.Range("CG_A_Amount" & count).Cells.item(i).value
STCG_AddressOfProperty(i) = Sheet13.Range("CG_A_AddressOfBuyer" & count).Cells.item(i).value
STCG_PinCode(i) = Sheet13.Range("CG_A_PinCodeOfBuyer" & count).Cells.item(i).value
'new change

STCG_StateCode(i) = Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value

'new change
Stcg_TotalPercentage(count) = Stcg_TotalPercentage(count) + STCG_PercentageShare(i)
    If FullValAsPer50C_1aiii > 0 Then
        

        If Not chkCompulsory(STCG_NameOfBuyer(i)) Then
         'changed by Chetan C M on 09/09/2025
        'start--
            'MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Name of Buyer in Short term capital gain is mandatory at Sr. No " & i & " in schedule CG" & Chr(13)
            MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please enter the name of buyer in A1f in Schedule CG" & Chr(13)
        '--end
        End If
        
        If Len(STCG_NameOfBuyer(i)) > 125 Then
            MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Name of Buyer in Short term capital gain cannot be more than 125 characters at Sr. No " & i & " in schedule CG" & Chr(13)
        End If
        
        If Len(STCG_PANOfBuyer(i)) = 10 Then
        If Not CheckPAN(STCG_PANOfBuyer(i)) Then
            MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please enter a valid 10 digit PAN in A1f at Sr. No " & i & " in Schedule CG" & Chr(13)
        End If
        End If
        
        If STCG_AadhaarOfBuyer(i) <> "" Then
        If Not IsNumeric(STCG_AadhaarOfBuyer(i)) Then
             MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid Aadhaar Number at Sr. No " & i & " in schedule CG" & Chr(13)
            End If
       
    
        If STCG_AadhaarOfBuyer(i) = "000000000000" Then
           MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid Aadhaar Number at Sr. No " & i & " in schedule CG" & Chr(13)

        End If
        
        If STCG_AadhaarOfBuyer(i) = "111111111111" Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid Aadhaar Number at Sr. No " & i & " in schedule CG" & Chr(13)

        End If
           
        If Len(STCG_AadhaarOfBuyer(i)) <> 12 Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid 12 digit Aadhaar Number at Sr. No " & i & " in schedule CG" & Chr(13)
        End If
        
        End If
        
        
        If Not (STCG_PercentageShare(i)) > 0 Then
           'changed by Chetan C M on 09/09/2025
        'start--
            'MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Percentage Share in Short term capital gain is mandatory at Sr. No " & i & "in schedule CG" & Chr(13)
            MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please enter percentage share in A1f in Schedule CG" & Chr(13)
        '--end
        End If
        
        If Not (STCG_Amount(i)) > 0 Then
           ' MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Amount in Short term capital gain is mandatory at Sr. No " & i & "in schedule CG" & Chr(13)
            'Ankita
             MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please enter Amount in A1f at Sr. No " & i & " in Schedule CG" & Chr(13)
 
        End If
        
        
'    'Changed by sai on 16/10/2024=================================================================
        If Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value <> "" Then

        Sheets("CG").Activate

        Dim PIN_targetadd, state_targetadd As String

        Dim ws_1 As Worksheet
        Set ws_1 = Worksheets("CG")


      'state_targetadd = Target.address
        state_targetadd = Replace(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "P", "Q")

            If Not state_Validation_Level(ws_1, PIN_targetadd, state_targetadd) Then
            
                     Sheet13.Range(Replace(state_targetadd, "P", "Q")).value = ""
            
                     MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Invalid Pincode" & Chr(13)

                 End If
            End If
             

   '========================================================================================

        If Not chkCompulsory(STCG_AddressOfProperty(i)) Then
        'changed by Chetan C M on 09/09/2025
        'start--
            'MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Address of Property in Short term capital gain is mandatory at Sr. No " & i & "in schedule CG" & Chr(13)
            MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please enter address of the property in A1f in Schedule CG" & Chr(13)
        '--end
        End If
        
        If Len(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value) < 0 Or Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value = "(Select)" Then
            MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Please enter State code in A1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
        End If
        
        If Len(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value) > 0 And Mid(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 2) <> "99" And Mid(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 1) <> "(" Then
             If Sheet13.Range("CG_A_PinCodeOfBuyer" & count).Cells.item(i).value = "" Then
                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please enter Pincode in A1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
            End If
    
'            If Len(Sheet13.Range("CG_A_PinCodeOfBuyer" & count).Cells.item(i).value) <> 6 Then
'                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Pincode in A1f" & " in Block No" & count & " should be 6 digits in Schedule CG" & Chr(13)
'            End If
    
            If Sheet13.Range("CG_A_PinCodeOfBuyer" & count).Cells.item(i).value = "000000" Then
                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Pincode in A1f" & " in Block No" & count & " is Invalid in Schedule CG" & Chr(13)
            End If
    
            If Mid(Sheet13.Range("CG_A_PinCodeOfBuyer" & count).Cells.item(i).value, 1, 1) = 0 Then
                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Pincode in A1f" & " in Block No" & count & " cannot begin with 0 in Schedule CG" & Chr(13)
            End If
    
            If Not checkfieldspecialcharacter(Sheet13.Range("CG_A_PinCodeOfBuyer" & count).Cells.item(i).value) Then
                 MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Pincode in A1f" & " in Block No" & count & " cannot contain special characters in Schedule CG" & Chr(13)
            End If
        End If
        
        
        
       'Added by Aavula on 02/08/2024  SIT-70650

           'If Len(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value) > 0 And Mid(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 2) = "99" Then
           If Len(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value) > 0 Then
   
       
            If Sheet13.Range("CG_A_CountryOfBuyer" & count).Cells.item(i).value = "" Then
               MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Please select country" & Chr(13)
            End If
    
           'Added by Aavula on 08/08/2024  SIT-70650
            If Mid(Sheet13.Range("CG_A_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 2) <> "99" And Mid(Sheet13.Range("CG_A_CountryOfBuyer" & count).Cells.item(i).value, 1, 2) <> "91" Then
             MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Country cannot be other than India as you have selected an Indian state" & Chr(13)
            End If
    
    '''''''''''''
            
            If Sheet13.Range("CG_A_ZipcodeOfBuyer" & count).Cells.item(i).value <> "" Then
'                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Please enter Zipcode in A1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
            End If
    
            If Len(Sheet13.Range("CG_A_ZipcodeOfBuyer" & count).Cells.item(i).value) > 8 Then
                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Zipcode in A1f" & " in Block No" & count & " should be 8 digits in Schedule CG" & Chr(13)
            End If
    
            If Sheet13.Range("CG_A_ZipcodeOfBuyer" & count).Cells.item(i).value = "000000" Then
                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Zipcode in A1f" & " in Block No" & count & " is Invalid in Schedule CG" & Chr(13)
            End If
    
            If Mid(Sheet13.Range("CG_A_ZipcodeOfBuyer" & count).Cells.item(i).value, 1, 1) = 0 Then
                MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Zipcode in A1f" & " in Block No" & count & " cannot begin with 0 in Schedule CG" & Chr(13)
            End If
    
            If Not checkfieldspecialcharacter(Sheet13.Range("CG_A_ZipcodeOfBuyer" & count).Cells.item(i).value) Then
                 MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Zipcode in A1f" & " in Block No" & count & " cannot contain special characters in Schedule CG" & Chr(13)
            End If
        End If
    
   End If
Next

'added by Chetan C M on 16/09/2025 SIT-95268
'start--
If Sheet13.Range("STCG.A20ia_FullConsideration11UAEi").value = "" Then
    If Sheet13.Range("STCG.A20iia_FullConsideration11UAEii").value <> "" Then
        MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Please enter Fair market value as per Rule 11UAE(2)" & Chr(13)
    End If
End If

If Sheet13.Range("STCG.A20iia_FullConsideration11UAEii").value = "" Then
    If Sheet13.Range("STCG.A20ia_FullConsideration11UAEi").value <> "" Then
        MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "Please enter Fair market value as per Rule 11UAE(3)" & Chr(13)
    End If
End If
'--end

If Stcg_TotalPercentage(count) <> "" Then
If Stcg_TotalPercentage(count) <> 100 Then
 MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Percentage Share in Short term capital gain is should be 100 in schedule CG " & Chr(13)
End If

If Stcg_TotalPercentage(count) > 100 Then
 MsgPartBScheduleCG_STCG = MsgPartBScheduleCG_STCG + "* Percentage Share in Short term capital gain is cannot be more than 100 in schedule CG" & Chr(13)
End If
End If

'If MsgPartBScheduleCG_STCG <> "" Then
'    fmsgboxStatus MsgPartBScheduleCG_STCG ', , "STCG Block in schedule CG" & Count
'    End
'End If

If MsgPartBScheduleCG_STCG <> "" Then
   MsgBox MsgPartBScheduleCG_STCG ', , "STCG Block in schedule CG" & Count
    End
End If
    UpdateProgressBar

Next
    
End Function
 
'end

' new LTCG code
Sub setTableInfo_LTCG_B1_1(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_NameOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_NameOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_Count1 = ccount
End Sub

Sub setTableInfo_LTCG_B1_2(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_PanOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_PanOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_Count2 = ccount
End Sub

Sub setTableInfo_LTCG_B1_A(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_AadhaaeOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_AadhaaeOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_CountA = ccount
End Sub

Sub setTableInfo_LTCG_B1_3(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_PercentageShare" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_PercentageShare" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_Count3 = ccount
End Sub

Sub setTableInfo_LTCG_B1_4(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_Amount" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_Amount" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_Count4 = ccount
End Sub

Sub setTableInfo_LTCG_B1_5(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_AddressOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_AddressOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_Count5 = ccount
End Sub

Sub setTableInfo_LTCG_B1_6(cnt As Variant)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("CG_B_PinCodeOfBuyer" & cnt).Cells
    mIntCells = Sheet13.Range("CG_B_PinCodeOfBuyer" & cnt).count
    For mIntCtr = 1 To mIntCells
        If chkCompulsory(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    LTCG_B1_Count6 = ccount
End Sub
Function CG_BiiiGreater()
    
    Dim msgboxGreatercheck As Variant
    msgboxGreatercheck = ""
If Sheet13.Range("STCG.A2ibiv_TotalDedn").value > 0 Then 'Not updating anything as this field will always Zero by Bindu on 11th Feb 26
    If Sheet13.Range("STCG.A2ia_FullConsideration").value = 0 Then 'Not updating anything as STCG.A2ibiv_TotalDedn will always Zero by Bindu on 11th Feb 26
        'msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        
        'Added by Aavula on 02/08/2024
         msgboxGreatercheck = msgboxGreatercheck + " 'Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero'" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If
 
'Newly added by sai for 3iibiv Q Column on 24/02/2025=================================
If Sheet13.Range("STCG.A2ibiv_2eii_TotalDedn").value > 0 Then
    If Sheet13.Range("STCG.A2ia_2eii_FullConsideration").value = 0 Then
        'msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'Added by Aavula on 02/08/2024
         msgboxGreatercheck = msgboxGreatercheck + " 'Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero'" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If
'================================================================================

If Sheet13.Range("STCG.A2iibiv_TotalDedn").value > 0 Then
    If Sheet13.Range("STCG.A2iia_FullConsideration").value = 0 Then ' 'Not updating anything as STCG.A2iibiv_TotalDedn will always Zero by Bindu on 11th Feb 26
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if  Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = Falsefv
        GoTo endline
    End If
End If

'Newly added by sai for 3iiaibiv Q Column on 24/02/2025=========================
If Sheet13.Range("STCG.A2iibiv_TotalDedn_ii").value > 0 Then
    If Sheet13.Range("STCG.A2iia_FullConsideration_ii").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if  Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = Falsefv
        GoTo endline
    End If
End If
'============================================================


If Sheet13.Range("STCG.A4biv_TotalDedn").value > 0 Then
    If Sheet13.Range("STCG.A4a_FullConsideration").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if  Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If

If Sheet13.Range("STCG.A5biv_TotalDedn").value > 0 Then
    If Sheet13.Range("STCG.A5a_FullConsideration").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if  Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If


'For Slump sale Table B2***********************************************************************
'Before
If Sheet13.Range("LTCG.B02d_ExemptionGrandTotal").value > 0 Then 'This field always be zero so not update required by Bindu on 12th feb 26
    'nOT UPDATING any thing as this field "LTCG.B02d_ExemptionGrandTotal" will always be zero by Bindu on 12th Feb 26
    If Sheet13.Range("LTCG.B02a_FullConsiderationUAEi").value = 0 Then 'For VF issue updated by sai on 23/06/2025
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if  Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If

'SIT-91233 After
If Sheet13.Range("LTCG.B02d_ExemptionGrandTotal_ii").value > 0 Then
    If Sheet13.Range("LTCG.B02a_FullConsiderationUAEi_ii").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if  Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If
'****************************************************************************************************


If Sheet13.Range("LTCG.B2biv_TotalDedn").value > 0 Then 'Will always be zero by Bindu on 16th Feb 26
    If Sheet13.Range("LTCG.B2a_FullConsideration").value = 0 Then 'Not updating as this field will alway be zero by Bindu on 16th Feb 26
    'SIT-69959 Changed by sai
       'msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
     
    'Added by Aavula on 02/08/2024 'SIT-69959
     
       msgboxGreatercheck = msgboxGreatercheck + "Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If

'Not updating anything as it will be zero by Bindu on 16th Feb 26
If Sheet13.Range("LTCG.B3biv_TotalDedn").value > 0 Then
    If Sheet13.Range("LTCG.B3a_FullConsideration").value = 0 Then
        'msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        
        'Added by Aavula on 02/08/2024 SIT-69953
        msgboxGreatercheck = msgboxGreatercheck + "'Deduction u/s 48 cannot be claimed if Full Value of Consideration at aiii is zero'" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If

'by sai on 25/03/2025 as per DE v0.4 for After****************************************************************
If Sheet13.Range("LTCG.B3biv_TotalDedn_ii").value > 0 Then
    If Sheet13.Range("LTCG.B3a_FullConsideration_ii").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + "'Deduction u/s 48 cannot be claimed if Full Value of Consideration at a is zero'" & Chr(13)
        GoTo endline
    End If
End If

'For Before
If Sheet13.Range("LTCG.B3biva_Total").value > 0 Then
    If Sheet13.Range("LTCG.B3a_FullConsideration").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + "'Deduction u/s 48 cannot be claimed if Full Value of Consideration at a is zero'" & Chr(13)
        GoTo endline
    End If
End If
'*****************************************************************************************************************


'SIT-69953 change by sai
If Sheet13.Range("LTCG.Bii3biv_TotalDedn").value > 0 Then 'Not updating anything as this field will always be zero by Bindu on 18th Feb 26
    If Sheet13.Range("LTCG.Bii3a_FullConsideration").value = 0 Then
    'Not updating anything as this field will always be zero by Bindu on 18th Feb 26
        
''Deduction u/s 48 cannot be claimed if Full Value of Consideration at aiii is zero'
        'msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration at aiii is zero in Schedule CG""" & Chr(13)
        'Added by Aavula
        msgboxGreatercheck = msgboxGreatercheck + "'Deduction u/s 48 cannot be claimed if Full Value of Consideration at aiii is zero'" & Chr(13)
        
        'ValidateA1Block = False
        GoTo endline
    End If
End If

'For AY 2025-26 by sai on 03/03/2025*************************************************
If Sheet13.Range("LTCG.Bii3biv_TotalDedn_ii").value > 0 Then
    If Sheet13.Range("LTCG.Bii3a_FullConsideration_ii").value = 0 Then
        
        msgboxGreatercheck = msgboxGreatercheck + "'Deduction u/s 48 cannot be claimed if  Full Value of Consideration at aiii is zero'" & Chr(13)
        
        GoTo endline
    End If
End If
'*************************************************************************************



If Sheet13.Range("LTCG.Bi5biv_TotalDedn").value > 0 Then 'Not updating as this field "LTCG.Bi5biv_TotalDedn" will always be zeroby Bindu on 18thFeb 26
    If Sheet13.Range("LTCG.Bi5a_FullConsideration").value = 0 Then 'Not updating as this field "LTCG.Bi5biv_TotalDedn" will always be zeroby Bindu on 18thFeb 26
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If

If Sheet13.Range("LTCG.Bii5biv_TotalDedn").value > 0 Then ''Not updating as this field "LTCG.Bii5biv_TotalDedn" will always be zeroby Bindu on 18thFeb 26
    If Sheet13.Range("LTCG.Bii5a_FullConsideration").value = 0 Then ''Not updating as this field "LTCG.Bii5biv_TotalDedn" will always be zeroby Bindu on 18thFeb 26
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If


If Sheet13.Range("LTCG.Biii5biv_TotalDedn").value > 0 Then
    If Sheet13.Range("LTCG.Biii5a_FullConsideration").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If

If Sheet13.Range("LTCG.B7biv_TotalDedn").value > 0 Then
    If Sheet13.Range("LTCG.B7a_FullConsideration").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If


'For AY 2025-26 by sai on 03/03/2025***********************************
If Sheet13.Range("LTCG.B7biv_TotalDedn_ii").value > 0 Then
    If Sheet13.Range("LTCG.B7a_FullConsideration_ii").value = 0 Then
        msgboxGreatercheck = msgboxGreatercheck + """Deduction u/s 48 cannot be claimed if Full Value of Consideration is zero in Schedule CG""" & Chr(13)
        'ValidateA1Block = False
        GoTo endline
    End If
End If
'***********************************************************************

''
endline:
    If msgboxGreatercheck <> "" Then
        Sheet13.Activate
        fmsgboxsmall msgboxGreatercheck ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Function
 Function LTCGBlock_Validation()
Dim LTCG_B1_Count As Variant
LTCG_B1_Count = Sheet13.Range("PropertySectionCount_LTCG").value
'Dim i As Long
Dim count As Variant
Dim FullValueConsideration_1ai As Variant
Dim PropertyStampValue_1aii As Variant
Dim FullValAsPer50C_1aiii As Variant

ReDim Ltcg_TotalPercentage(LTCG_B1_Count)
For count = 1 To LTCG_B1_Count

FullValueConsideration_1ai = Sheet13.Range("LTCG.Bai_FullConsideration" & count).value
PropertyStampValue_1aii = Sheet13.Range("LTCG.Baii_PropertyValuation" & count).value
FullValAsPer50C_1aiii = Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & count).value

If ((FullValueConsideration_1ai <> "" Or FullValueConsideration_1ai <> 0) And (PropertyStampValue_1aii = 0 Or PropertyStampValue_1aii = "")) Then

    MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please Enter  Value of property as per stamp valuation authority in Long term capital gain in schedule CG" & Chr(13)
End If

If (PropertyStampValue_1aii <> "" Or PropertyStampValue_1aii <> 0) And (FullValueConsideration_1ai = 0 Or FullValueConsideration_1ai = "") Then

   MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please Enter Full value of consideration received/receivable in Long term capital gain in schedule CG" & Chr(13)
End If


'Added by sai on 16/03/2025
'***************************************************************
If (Len(Sheet13.Range("LTCG.BII_CapitalLoss").value) > 15) Then
    MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Capital Loss on buy back of shares on or after 01st October 2024 [ Long Term 20% or 30% or applicable rate] Amount cannot exceed 14 digits. in Long term capital gain in schedule CG " & Chr(13)
End If

'by sadineni on 24/04/2026 as per DE V0.2

If Sheet13.Range("LTCG.BII_CapitalLoss").value = "" Or Sheet13.Range("LTCG.BII_CapitalLoss").value > 0 Then
    MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please enter value Sl.no.B(A)." & Chr(13)
End If

'***************************************************************

If ((FullValAsPer50C_1aiii <> FullValueConsideration_1ai) And (FullValAsPer50C_1aiii <> PropertyStampValue_1aii)) Then
 MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please Ensure LTCG  (1)aiii is either ai or aii (Selct from Drop Down) in schedule CG" & Chr(13)
End If

If Sheet13.Range("LTCG.Bbiv_TotalDedn" & count).value > 0 Then
    If Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & count).value = 0 Then
        MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Deduction u/s 48 cannot be claimed if Full Value of Consideration at aiii is zero" & Chr(13)
'        ValidateB1Block = False
    End If
End If

'**********************************************************************************************************
'by sadineni on 04/05/2026 as per DE V0.2
If Sheet13.Range("LTCG.Bbiva_TotalDedn" & count).value > 0 Then
    If Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & count).value = 0 Then
        MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Deduction at Sl. No. biva cannot be claimed if a(iii) is zero" & Chr(13)
    End If
End If
'***********************************************************************************************************


    setTableInfo_LTCG_B1_1 (count)
    setTableInfo_LTCG_B1_2 (count)
    setTableInfo_LTCG_B1_3 (count)
    setTableInfo_LTCG_B1_4 (count)
    setTableInfo_LTCG_B1_5 (count)
    setTableInfo_LTCG_B1_6 (count)
    setTableInfo_LTCG_B1_A (count)
    
LTCG_B1_Count1 = WorksheetFunction.Max(0, LTCG_B1_Count1, LTCG_B1_Count2, LTCG_B1_Count3, LTCG_B1_Count4, LTCG_B1_Count5, LTCG_B1_Count6, LTCG_B1_CountA)
ReDim LTCG_NameOfBuyer(LTCG_B1_Count1), LTCG_PANOfBuyer(LTCG_B1_Count1), LTCG_AadhaarOfBuyer(LTCG_B1_Count1), LTCG_PercentageShare(LTCG_B1_Count1), LTCG_Amount(LTCG_B1_Count1), LTCG_AddressOfProperty(LTCG_B1_Count1), LTCG_PinCode(LTCG_B1_Count1), LTCG_StateCode(LTCG_B1_Count1)
Dim i As Long
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim cellrange As Variant
Dim cellrange1 As Variant
Dim cellrange2 As Variant

If FullValAsPer50C_1aiii > 0 Then
If LTCG_B1_Count1 = 0 Then
 MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* In case of transfer of immovable property please furnish following details Table in Long term capital gain is mandatory in schedule CG" & Chr(13)
End If
End If

'Change-26.04.2023.101.CG -> SIT-49129
'If LTCG_B1_Count1 > 0 Then
'If (FullValueConsideration_1ai = 0 Or FullValueConsideration_1ai = "") Then
'MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please Enter Full value of consideration received/receivable in Long term capital gain in schedule CG" & Chr(13)
'End If
'End If


'Added by Sai on 28/02/2025******************************************
'For LTCG 5b
If Sheet13.Range("B5_DeductionTotal").value > 100000000 Then
    MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Amount against section 54F can not be more than 10 crores." & Chr(13)
End If

'For LTCG 6b
If Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value > 100000000 Then
    MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Amount against section 54F can not be more than 10 crores." & Chr(13)
End If


'For LTCG 8b
If Sheet13.Range("LTCG.8.DeductionUnder54F").value > 100000000 Then 'Not updating anything as this field refers to this field "LTCG.8_ii.DeductionUnder54F" by bindu on 18th feb 26
    MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Amount against section 54F can not be more than 10 crores." & Chr(13)
End If



'*********************************************************************



If (FullValueConsideration_1ai > 0 Or PropertyStampValue_1aii > 0 Or FullValAsPer50C_1aiii > 0) Then
    If (Sheet13.Range("LTCG.B_SaleTrfDate" & count).value) = "" Then
'        MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please select the date of Sale  in order to calculate the type of  capital Gain" & Chr(13)
        'by sai as per DE V0.4 on 26/03/2025
        MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* ""Date of sale/ transfer in sl. no. B1 of LTCG is mandatory""." & Chr(13)
        
    End If
    If (Sheet13.Range("LTCG.B_AcqDate" & count).value) = "" Then
'        MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please select the date of purchase in order to calculate the type of  capital Gain" & Chr(13)
        
        'by sai as per DE V0.4 on 26/03/2025
        MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* ""Date of purchase/ acquisition in sl. no. B1 of LTCG is mandatory""." & Chr(13)
        
    End If
End If
'-----
For i = 1 To LTCG_B1_Count1
Set rangecells = Sheet13.Range("CG_B_NameOfBuyer" & count).Cells
Set rangecells1 = Sheet13.Range("CG_B_PanOfBuyer" & count).Cells
Set rangecells2 = Sheet13.Range("CG_B_AadhaaeOfBuyer" & count).Cells

        cellrange = GetMergedAddressCell(rangecells, i)
        LTCG_NameOfBuyer(i) = Sheet13.Range(cellrange).value
        cellrange1 = GetMergedAddressCell(rangecells1, i)
        LTCG_PANOfBuyer(i) = Sheet13.Range(cellrange1).value
        cellrange2 = GetMergedAddressCell(rangecells2, i)
        LTCG_AadhaarOfBuyer(i) = Sheet13.Range(cellrange2).value

LTCG_PercentageShare(i) = Sheet13.Range("CG_B_PercentageShare" & count).Cells.item(i).value
LTCG_Amount(i) = Sheet13.Range("CG_B_Amount" & count).Cells.item(i).value
LTCG_AddressOfProperty(i) = Sheet13.Range("CG_B_AddressOfBuyer" & count).Cells.item(i).value
LTCG_PinCode(i) = Sheet13.Range("CG_B_PinCodeOfBuyer" & count).Cells.item(i).value
LTCG_StateCode(i) = Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value

Ltcg_TotalPercentage(count) = Ltcg_TotalPercentage(count) + LTCG_PercentageShare(i)
    If FullValAsPer50C_1aiii > 0 Then
        

        If Not chkCompulsory(LTCG_NameOfBuyer(i)) Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Name of Buyer in Long term capital gain is mandatory at Sr. No " & i & "in schedule CG" & Chr(13)
        End If
        
        If Len(LTCG_NameOfBuyer(i)) > 125 Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Name of Buyer in Long term capital gain cannot be more than 125 characters at Sr. No " & i & "in schedule CG" & Chr(13)
        End If
        
        If Len(LTCG_PANOfBuyer(i)) = 10 Then
        If Not CheckPAN(LTCG_PANOfBuyer(i)) Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* PAN of Buyer in Long term capital gain is Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr. No " & i & "in schedule CG" & Chr(13)
        End If
        End If
        
        If LTCG_AadhaarOfBuyer(i) <> "" Then
        If Not IsNumeric(LTCG_AadhaarOfBuyer(i)) Then
             MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid Aadhaar Number in schedule CG" & Chr(13)
            End If
       
    
        If LTCG_AadhaarOfBuyer(i) = "000000000000" Then
           MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid Aadhaar Number in schedule CG" & Chr(13)

        End If
        
        If LTCG_AadhaarOfBuyer(i) = "111111111111" Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid Aadhaar Number in schedule CG" & Chr(13)

        End If
           
        If Len(LTCG_AadhaarOfBuyer(i)) <> 12 Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + Chr(13) & "Please enter valid 12 digit Aadhaar Number in schedule CG" & Chr(13)
        End If
        
        End If
        
        If Not (LTCG_PercentageShare(i)) > 0 Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Percentage Share in Long term capital gain is mandatory at Sr. No " & i & "in schedule CG" & Chr(13)
        End If
        
        If Not (LTCG_Amount(i)) > 0 Then
           ' MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Amount in Long term capital gain is mandatory ccin schedule CG" & Chr(13)
           'Added by Aavula on 02/08/2024 SIT-70570
           MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Please enter the amount in schedule CG in item B2 at Sr. No " & i & "" & Chr(13)
        
        End If
        
        If Not chkCompulsory(LTCG_AddressOfProperty(i)) Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Address of Property in Long term capital gain is mandatory at Sr. No " & i & "in schedule CG" & Chr(13)
        End If
        
         If Len(Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value) < 0 Or Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value = "(Select)" Then
            MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Please enter State code in B1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
        End If
        
        If Len(Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value) > 0 And Mid(Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 2) <> "99" And Mid(Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 1) <> "(" Then
             If Sheet13.Range("CG_B_PinCodeOfBuyer" & count).Cells.item(i).value = "" Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Please enter Pincode in B1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
            End If
    
            If Len(Sheet13.Range("CG_B_PinCodeOfBuyer" & count).Cells.item(i).value) <> 6 Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Pincode in B1f" & " in Block No" & count & " should be 6 digits in Schedule CG" & Chr(13)
            End If
    
            If Sheet13.Range("CG_B_PinCodeOfBuyer" & count).Cells.item(i).value = "000000" Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Pincode in B1f" & " in Block No" & count & " is Invalid in Schedule CG" & Chr(13)
            End If
    
            If Mid(Sheet13.Range("CG_B_PinCodeOfBuyer" & count).Cells.item(i).value, 1, 1) = 0 Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Pincode in B1f" & " in Block No" & count & " cannot begin with 0 in Schedule CG" & Chr(13)
            End If
    
            If Not checkfieldspecialcharacter(Sheet13.Range("CG_B_PinCodeOfBuyer" & count).Cells.item(i).value) Then
                 MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Pincode in B1f" & " in Block No" & count & " cannot contain special characters in Schedule CG" & Chr(13)
            End If
        End If
        
        If Len(Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value) > 0 And Mid(Sheet13.Range("CG_B_StateCodeOfBuyer" & count).Cells.item(i).value, 1, 2) = "99" Then
            If Len(Sheet13.Range("CG_B_CountryOfBuyer" & count).Cells.item(i).value) < 0 Or Mid(Sheet13.Range("CG_B_CountryOfBuyer" & count).Cells.item(i).value, 1, 1) = "(" Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Please enter Country code in B1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
            End If
            If Sheet13.Range("CG_B_ZipcodeOfBuyer" & count).Cells.item(i).value <> "" Then
                'MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Please enter Zipcode in B1f" & " in Block No" & count & " in Schedule CG" & Chr(13)
            End If
    
            If Len(Sheet13.Range("CG_B_ZipcodeOfBuyer" & count).Cells.item(i).value) > 8 Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Zipcode in B1f" & " in Block No" & count & " should be 8 digits in Schedule CG" & Chr(13)
            End If
    
            If Sheet13.Range("CG_B_ZipcodeOfBuyer" & count).Cells.item(i).value = "000000" Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Zipcode in B1f" & " in Block No" & count & " is Invalid in Schedule CG" & Chr(13)
            End If
    
            If Mid(Sheet13.Range("CG_B_ZipcodeOfBuyer" & count).Cells.item(i).value, 1, 1) = 0 Then
                MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Zipcode in B1f" & " in Block No" & count & " cannot begin with 0 in Schedule CG" & Chr(13)
            End If
    
            If Not checkfieldspecialcharacter(Sheet13.Range("CG_B_ZipcodeOfBuyer" & count).Cells.item(i).value) Then
                 MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "Zipcode in B1f" & " in Block No" & count & " cannot contain special characters in Schedule CG" & Chr(13)
            End If
        End If
               
    End If
Next


If Ltcg_TotalPercentage(count) <> "" Then
If Ltcg_TotalPercentage(count) <> 100 Then
 MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Percentage Share in Long term capital gain is should be 100 in schedule CG" & Chr(13)
End If

If Ltcg_TotalPercentage(count) > 100 Then
 MsgPartBScheduleCG_LTCG = MsgPartBScheduleCG_LTCG + "* Percentage Share in Long term capital gain is cannot be more than 100 in schedule CG" & Chr(13)
End If
End If

If MsgPartBScheduleCG_LTCG <> "" Then
    MsgBox MsgPartBScheduleCG_LTCG ', , "LTCG Block in schedule CG" & Count
    End
End If

    UpdateProgressBar

Next
    
End Function

'Newly added by sai for AY 2025-26 on 18/02/2025
Sub AddRowsAII()
Dim vRows As Long
Sheets("CG").Activate
'Newly added STCG_AII.sino
EfilingCommon.DefinedgridNameRange = "STCG_AII.sino||STCG.AII_CapitalLoss20r30rapplicablerate||STCG.AII_CapitalLoss20r30rapplicablerate_amt"
ActiveCellRange = EfilingCommon.searchLastRow("STCG.AII_CapitalLoss20r30rapplicablerate")
vRows = EfilingCommon.insertRowUnderSectionWithFormula_CG

 
Dim rangecells, rangecells1, rangecells2 As Range
    Dim mIntCells, q As Long
Set rangecells = Range("STCG.AII_CapitalLoss20r30rapplicablerate").Cells
 
Set rangecells2 = Range("STCG_AII.sino").Cells

mIntCells = Range("STCG_AII.sino").count

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Application.EnableEvents = False

For q = 1 To mIntCells

    If rangecells.Cells(q, 1).value = "" Then rangecells.Cells(q, 1).value = "(Select)"
    
    If rangecells2.item(q).value = "" Then rangecells2.item(q).value = rangecells2.item(q).Offset(-1, 0).value + 1
Next

Application.EnableEvents = True
'----------------Lock Password-------------------START---
  Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
 
End Sub
'Newly added by sai for AY 2025-26 on 18/02/2025
Sub AddRowsBII() 'Will not using anywhere So not commenting by Bindu on 22nd Feb 26
Dim vRows As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "LTCG.BII_siNo||LTCG.BII_Rate||LTCG.BII_Amount"
ActiveCellRange = EfilingCommon.searchLastRow("LTCG.BII_Rate")
vRows = EfilingCommon.insertRowUnderSectionWithFormula_CG

Dim rangecells, rangecells1, rangecells2 As Range
    Dim mIntCells, q As Long
 Set rangecells2 = Range("LTCG.BII_siNo").Cells

mIntCells = Range("LTCG.BII_siNo").count
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Application.EnableEvents = False
For q = 1 To mIntCells
     
    If rangecells2.item(q).value = "" Then rangecells2.item(q).value = rangecells2.item(q).Offset(-1, 0).value + 1
Next

Application.EnableEvents = True
'----------------Lock Password-------------------START---
  Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub

'Newly added by sai on 27/02/2025
Sub setTblinfo_STCGAII()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 
 ccount = 0
 mIntCells = Sheet13.Range("STCG_AII.sino").count
 Set rangecells = Sheet13.Range("STCG.AII_CapitalLoss20r30rapplicablerate").Cells
  For mIntCtr = 1 To mIntCells
  'MsgBox rangecells.item(mIntCtr, 1).value
     If Not rangecells.item(mIntCtr, 1).value = "(Select)" And rangecells.item(mIntCtr, 1).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 end_STCGAII_Captlos20r30rapplicblrate = ccount
 
 End Sub
'Newly added by sai on 27/02/2025
 Sub setTblinfo_STCGAII_amt()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 
 ccount = 0
 
 mIntCells = Sheet13.Range("STCG_AII.sino").count
 Set rangecells = Sheet13.Range("STCG.AII_CapitalLoss20r30rapplicablerate_amt").Cells
  For mIntCtr = 1 To mIntCells
  'MsgBox rangecells.item(mIntCtr).value
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
    end_STCGAII_Captlos20r30rapplicblrate_amt = ccount
End Sub
'Newly mdified by sai on 24/03/2025 for AY 2025-26
  Function ValidatesheetSTCG_AII() As Boolean
 
    ValidatesheetSTCG_AII = True
    Dim i As Long
    
    Msgbox_STCG_AII = ""
    Msgbox_STCG_AII = "Schedule CG : " & Chr(13) & Chr(13)
    
    setTblinfo_STCGAII
    setTblinfo_STCGAII_amt
    
    'end_STCGAII_Captlos20r30rapplicblrate = Application.WorksheetFunction.Max(0, end_STCGAII_Captlos20r30rapplicblrate, end_STCGAII_Captlos20r30rapplicblrate_amt)
    
    Dim rangecells, rangecells1 As Range
    
    Set rangecells = Sheet13.Range("STCG.AII_CapitalLoss20r30rapplicablerate").Cells
    Set rangecells1 = Sheet13.Range("STCG.AII_CapitalLoss20r30rapplicablerate_amt").Cells
    
    
   ReDim Captlos20r30rapplicblrate_STCGAII(end_STCGAII_Captlos20r30rapplicblrate)
   ReDim Captlos20r30rapplicblrate_amt_STCGAII(end_STCGAII_Captlos20r30rapplicblrate)
   
    For i = 1 To end_STCGAII_Captlos20r30rapplicblrate
    
        Captlos20r30rapplicblrate_STCGAII(i) = rangecells.item(i, 1).value
        Captlos20r30rapplicblrate_amt_STCGAII(i) = rangecells1.item(i).value
        
         If Captlos20r30rapplicblrate_amt_STCGAII(i) = "" Then
            'Error msg updated as per DE V4
'            Msgbox_STCG_AII = Msgbox_STCG_AII + "*  ""Please enter value at Sl.no.A(A) against the dropdown selected  in schedule CG""" & Chr(13)
             Msgbox_STCG_AII = Msgbox_STCG_AII + "*  ""Please enter value at Sl.no.A(A) against the dropdown selected """ & Chr(13)
             ValidatesheetSTCG_AII = False
            Exit Function
         End If
   Next
   
   
   ReDim Captlos20r30rapplicblrate_STCGAII(end_STCGAII_Captlos20r30rapplicblrate_amt)
   ReDim Captlos20r30rapplicblrate_amt_STCGAII(end_STCGAII_Captlos20r30rapplicblrate_amt)
   For i = 1 To end_STCGAII_Captlos20r30rapplicblrate_amt
    
        Captlos20r30rapplicblrate_STCGAII(i) = rangecells.item(i, 1).value
        Captlos20r30rapplicblrate_amt_STCGAII(i) = rangecells1.item(i).value
        
         If Captlos20r30rapplicblrate_STCGAII(i) = "" Or UCase(Captlos20r30rapplicblrate_STCGAII(i)) = UCase("(Select)") Then
         
'Added by sudha_2025-26

            'Msgbox_STCG_AII = Msgbox_STCG_AII + "*  ""Please select dropdown in Sl.no.A(A) in schedule CG""" & Chr(13)
'             Msgbox_STCG_AII = Msgbox_STCG_AII + "*  ""Please select dropdown at Sl.no.A(A) in schedule CG""" & Chr(13)
              Msgbox_STCG_AII = Msgbox_STCG_AII + "*  ""Please select dropdown at Sl.no.A(A)""" & Chr(13)
             
            ValidatesheetSTCG_AII = False
            Exit Function
         End If
         

  
   Next
   
   
End Function

'Newly added by sai on 27/02/2025

 Sub setTblinfo_LTCGBII()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 
 ccount = 0
 
 mIntCells = Sheet13.Range("LTCG.BII_siNo").count
 Set rangecells = Sheet13.Range("LTCG.BII_Rate").Cells
  For mIntCtr = 1 To mIntCells
  'MsgBox rangecells.item(mIntCtr, 1).value
     If Not rangecells.item(mIntCtr, 1).value = "(Select)" And rangecells.item(mIntCtr, 1).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 end_LTCGBII_Captlos20r30rapplicblrate = ccount
 
 End Sub
'Newly added by sai on 27/02/2025
Sub setTblinfo_LTCGBII_amt()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 
 ccount = 0
 
 mIntCells = Sheet13.Range("LTCG.BII_siNo").count
 Set rangecells = Sheet13.Range("LTCG.BII_Amount").Cells
  For mIntCtr = 1 To mIntCells
  'MsgBox rangecells.item(mIntCtr).value
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
    end_LTCGBII_Captlos20r30rapplicblrate_amt = ccount
End Sub
'Newly modified by sai on 24/03/2025 for AY 2025-26
  Function ValidatesheetLTCG_BII() As Boolean
 
    ValidatesheetLTCG_BII = True
    Dim i As Long
    
    Msgbox_LTCG_BII = ""
    Msgbox_LTCG_BII = "Schedule CG : " & Chr(13) & Chr(13)
    
    setTblinfo_LTCGBII
    setTblinfo_LTCGBII_amt
    
    'end_LTCGBII_Captlos20r30rapplicblrate = Application.WorksheetFunction.Max(0, end_LTCGBII_Captlos20r30rapplicblrate, end_LTCGBII_Captlos20r30rapplicblrate_amt)
    
    Dim rangecells, rangecells1 As Range
    
    Set rangecells = Sheet13.Range("LTCG.BII_Rate").Cells
    Set rangecells1 = Sheet13.Range("LTCG.BII_Amount").Cells
    
    
   ReDim Captlos20r30rapplicblrate_LTCGBII(end_LTCGBII_Captlos20r30rapplicblrate)
   ReDim Captlos20r30rapplicblrate_amt_LTCGBII(end_LTCGBII_Captlos20r30rapplicblrate)
   
    For i = 1 To end_LTCGBII_Captlos20r30rapplicblrate
    
        Captlos20r30rapplicblrate_LTCGBII(i) = rangecells.item(i, 1).value
        Captlos20r30rapplicblrate_amt_LTCGBII(i) = rangecells1.item(i).value
        
         If Captlos20r30rapplicblrate_amt_LTCGBII(i) = "" Then
            'Error msg updated as per DE V4
'            Msgbox_LTCG_BII = Msgbox_LTCG_BII + "*  ""Please enter value at Sl.no.B(A) against the dropdown selected in schedule CG""" & Chr(13)
             Msgbox_LTCG_BII = Msgbox_LTCG_BII + "*  ""Please enter value at Sl.no.B(A) against the dropdown selected""" & Chr(13)
            
            ValidatesheetLTCG_BII = False
         End If
   Next

   ReDim Captlos20r30rapplicblrate_LTCGBII(end_LTCGBII_Captlos20r30rapplicblrate_amt)
   ReDim Captlos20r30rapplicblrate_amt_LTCGBII(end_LTCGBII_Captlos20r30rapplicblrate_amt)
   For i = 1 To end_LTCGBII_Captlos20r30rapplicblrate_amt
    
        Captlos20r30rapplicblrate_LTCGBII(i) = rangecells.item(i, 1).value
        Captlos20r30rapplicblrate_amt_LTCGBII(i) = rangecells1.item(i).value
        
         If Captlos20r30rapplicblrate_LTCGBII(i) = "" Or UCase(Captlos20r30rapplicblrate_LTCGBII(i)) = UCase("(Select)") Then

'Added by sudha_2025-26

            'Msgbox_LTCG_BII = Msgbox_LTCG_BII + "*  ""Please select dropdown in Sl.no.B(A) in schedule CG""" & Chr(13)
'             Msgbox_LTCG_BII = Msgbox_LTCG_BII + "*  ""Please select dropdown at Sl.no.B(A) in schedule CG""" & Chr(13)
              Msgbox_LTCG_BII = Msgbox_LTCG_BII + "*  ""Please select dropdown at Sl.no.B(A)""" & Chr(13)
             
            ValidatesheetLTCG_BII = False
            Exit Function
         End If
         


   Next
   
   
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst115E_a125(value, cgosIncome_cgInc_ltcg_125prctg125)
Dim amtSetOff As Double
   'prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)



    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec115E_a) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg125_sec115E_a) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115E_a = (cgosIncome_cgInc_ltcg_prctg125_sec115E_a + value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec115E_a)
           cgosIncome_cgInc_ltcg_prctg125_sec115E_a = 0
        End If
        If ((amtSetOff) > (prctg125Sum)) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115E_a = cgosIncome_cgInc_ltcg_prctg125_sec115E_a + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
        End If
    
    setOffAgainst115E_a125 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst112ASEC48(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double)
   Dim amtSetOff As Double
    ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


           
     amtSetOff = 0
     
     setOffAgainst112ASEC48 = amtSetOff
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI112_sec48) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI112_sec48) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_PTI112_sec48 = (cgosIncome_cgInc_ltcg_PTI112_sec48 + value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI112_sec48)
           cgosIncome_cgInc_ltcg_PTI112_sec48 = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
         cgosIncome_cgInc_ltcg_PTI112_sec48 = cgosIncome_cgInc_ltcg_PTI112_sec48 + amtSetOff - prctg125Sum
            amtSetOff = (prctg125Sum)
        End If
    End If
    
   setOffAgainst112ASEC48 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst112AB9(value, cgosIncome_cgInc_ltcg_125prctg125)
Dim amtSetOff As Double
  ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
   prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


             

  ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + cgosIncome_cgInc_ltcg_PTI125_112_B9)
             
    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112_B9) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112_B9) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_PTI125_112_B9 = cgosIncome_cgInc_ltcg_PTI125_112_B9 + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112_B9)
           cgosIncome_cgInc_ltcg_PTI125_112_B9 = 0
        End If
        If ((amtSetOff) > (prctg125Sum)) Then
           cgosIncome_cgInc_ltcg_PTI125_112_B9 = cgosIncome_cgInc_ltcg_PTI125_112_B9 + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
        End If
    
    setOffAgainst112AB9 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst112_1_c_125(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double) As Double
  Dim amtSetOff As Double
    ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
               cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)



     amtSetOff = 0
     
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + (value)
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2)
            cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst112_1_c_125 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst115AC_125(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double)
   Dim amtSetOff As Double
   ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + cgosIncome_cgInc_ltcg_PTI125_112_B9)
    prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


 
     amtSetOff = 0

    If (value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec115AC_1) > 0 Then
        If ((cgosIncome_cgInc_ltcg_prctg125_sec115AC_1) >= Abs(value)) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec115AC_1)
             cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = (cgosIncome_cgInc_ltcg_prctg125_sec115AC_1) + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst115AC_125 = amtSetOff
End Function

'Newly added by Riyaz on 25/02/2026

Function setOffAgainst115AD_125(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double)
   Dim amtSetOff As Double
    'prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + cgosIncome_cgInc_ltcg_PTI125_112_B9 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3)
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)



     amtSetOff = 0

    If (value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) > 0 Then
        If ((cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) >= Abs(value)) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec115AD_3)
             cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = (cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst115AD_125 = amtSetOff
End Function


'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst115ACA_125(value, cgosIncome_cgInc_ltcg_125prctg125)
  Dim amtSetOff As Double
'prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)



amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1) > 0) Then
        If ((cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1) >= Abs(value)) Then
            cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1)
            cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = (cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1) + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    
    setOffAgainst115ACA_125 = amtSetOff

  End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst115AD_3_Proviso_125(value, cgosIncome_cgInc_ltcg_125prctg125)
Dim amtSetOff As Double
  ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)

prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)



    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso)
           cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = 0
        End If
        If ((amtSetOff) > (prctg125Sum)) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
        End If
    
    setOffAgainst115AD_3_Proviso_125 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst115E_b125(value, cgosIncome_cgInc_ltcg_125prctg125)
  Dim amtSetOff As Double
'prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
     
 prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


    
     amtSetOff = 0

    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec115E_b) > 0) Then
        If ((cgosIncome_cgInc_ltcg_prctg125_sec115E_b) >= Abs(value)) Then
            cgosIncome_cgInc_ltcg_prctg125_sec115E_b = cgosIncome_cgInc_ltcg_prctg125_sec115E_b + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec115E_b)
            cgosIncome_cgInc_ltcg_prctg125_sec115E_b = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
           cgosIncome_cgInc_ltcg_prctg125_sec115E_b = (cgosIncome_cgInc_ltcg_prctg125_sec115E_b) + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    
    setOffAgainst115E_b125 = amtSetOff

  End Function
'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainstPTI125(value, cgosIncome_cgInc_ltcg_125prctg125)
Dim amtSetOff As Double
   'prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
    prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


            
    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_PTI125 = cgosIncome_cgInc_ltcg_PTI125 + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125)
           cgosIncome_cgInc_ltcg_PTI125 = 0
        End If
        If ((amtSetOff) > (prctg125Sum)) Then
           cgosIncome_cgInc_ltcg_PTI125 = cgosIncome_cgInc_ltcg_PTI125 + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
        End If
    
    setOffAgainstPTI125 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainstsecProviso_125(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double)
   Dim amtSetOff As Double
   '    prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
                
    prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


            
     amtSetOff = 0
     
     setOffAgainstsecProviso_125 = amtSetOff
     If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_secProviso) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg125_secProviso) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg125_secProviso = cgosIncome_cgInc_ltcg_prctg125_secProviso + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_secProviso)
           cgosIncome_cgInc_ltcg_prctg125_secProviso = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
         cgosIncome_cgInc_ltcg_prctg125_secProviso = cgosIncome_cgInc_ltcg_prctg125_secProviso + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    
   setOffAgainstsecProviso_125 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainstPTI10_1125(value, cgosIncome_cgInc_ltcg_125prctg125)
Dim amtSetOff As Double
  ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
   prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
               cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


 
    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_PTI125_112 = cgosIncome_cgInc_ltcg_PTI125_112 + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112)
           cgosIncome_cgInc_ltcg_PTI125_112 = 0
        End If
        If ((amtSetOff) > (prctg125Sum)) Then
           cgosIncome_cgInc_ltcg_PTI125_112 = cgosIncome_cgInc_ltcg_PTI125_112 + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
        End If
    
    setOffAgainstPTI10_1125 = amtSetOff
End Function

'by Sai on 04/04/2025 for LTCG_12.5%
Function setOffAgainst112A125(value, cgosIncome_cgInc_ltcg_125prctg125)
Dim amtSetOff As Double
  ' prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9)
                
                
   prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)


             
    amtSetOff = 0
      
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_prctg125_sec112A) > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg125_sec112A) >= Abs(value) Then
           cgosIncome_cgInc_ltcg_prctg125_sec112A = cgosIncome_cgInc_ltcg_prctg125_sec112A + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg125_sec112A)
           cgosIncome_cgInc_ltcg_prctg125_sec112A = 0
        End If
        If ((amtSetOff) > (prctg125Sum)) Then
           cgosIncome_cgInc_ltcg_prctg125_sec112A = cgosIncome_cgInc_ltcg_prctg125_sec112A + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
        End If
    
    setOffAgainst112A125 = amtSetOff
End Function
'LTCG 12.5% by sai on 07/04/2025
 Function setOffPctg125Loss(cgosIncome_cgInc_stcg_prctgAr10Loss As Double)
Dim amtSetOff As Double
InLtcg20Per = Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").value
InLtcg10Per = Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").value

'cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
'prctg10 = cgosIncome.cgInc.ltcg.prctg10
' prctg125Sum = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_PTI125_112_B9 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3

prctg125Sum = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10


 amtSetOff = 0
 
    amtSetOff = setOffAgainst20(prctg125Sum, cgosIncome_cgInc_ltcg_prctg20) '20
    prctg125Sum = amtSetOff + prctg125Sum 'by sai on 08/04/2025
    amtSetOff = setOffAgainst10(prctg125Sum, cgosIncome_cgInc_ltcg_prctg10) '10

    
    InLtcg20Per = Abs(amtSetOff)
    'Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value = InLtcg20Per
    
    '-not using
'    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec115E_a) > amtSetOff) Then
'        cgosIncome_cgInc_ltcg_prctg125_sec115E_a = cgosIncome_cgInc_ltcg_prctg125_sec115E_a + amtSetOff
'        amtSetOff = 0
'    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec115E_a < 0) Then
'        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec115E_a
'       cgosIncome_cgInc_ltcg_prctg125_sec115E_a = 0
'    End If
    
    '--not using
'     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI112_sec48) > amtSetOff) Then
'        cgosIncome_cgInc_ltcg_PTI112_sec48 = cgosIncome_cgInc_ltcg_PTI112_sec48 + amtSetOff
'        amtSetOff = 0
'    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI112_sec48 < 0) Then
'        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI112_sec48
'        cgosIncome_cgInc_ltcg_PTI112_sec48 = 0
'    End If
'    '--10th 'Not using
'    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112_B9) > amtSetOff) Then
'        cgosIncome_cgInc_ltcg_PTI125_112_B9 = cgosIncome_cgInc_ltcg_PTI125_112_B9 + amtSetOff
'        amtSetOff = 0
'    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112_B9 < 0) Then
'        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112_B9
'        cgosIncome_cgInc_ltcg_PTI125_112_B9 = 0
'    End If
    
    
    '--15. B10
     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112_B10) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125_112_B10 = cgosIncome_cgInc_ltcg_PTI125_112_B10 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112_B10 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112_B10
        cgosIncome_cgInc_ltcg_PTI125_112_B10 = 0
    End If
    
    '--14. B9e
     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112_B9e) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125_112_B9e = cgosIncome_cgInc_ltcg_PTI125_112_B9e + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112_B9e < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112_B9e
        cgosIncome_cgInc_ltcg_PTI125_112_B9e = 0
    End If
    
    
    '--13. B5c
     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112_B5c) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125_112_B5c = cgosIncome_cgInc_ltcg_PTI125_112_B5c + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112_B5c < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112_B5c
        cgosIncome_cgInc_ltcg_PTI125_112_B5c = 0
    End If
    
    
    '---12. B2e
     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112_B2e) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125_112_B2e = cgosIncome_cgInc_ltcg_PTI125_112_B2e + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112_B2e < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112_B2e
        cgosIncome_cgInc_ltcg_PTI125_112_B2e = 0
    End If
    
     '--11. B6ie - 112(1)(c )
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
       cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = 0
    End If
    
    
    '--10. B6iie - 115AC(1)(c )
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec115AC_1) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
       cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = 0
    End If
    
       '--9. B3iie - 115ACA(1)(b)
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1) > amtSetOff) Then
       cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = 0
    End If
    
'Newly added by Riyaz on 25/02/2026
'8. B6iiie - 115AD(1)(iii)
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
       cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = 0
    End If
    
      '--7. B8c - 115E(b)
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec115E_b) > amtSetOff) Then
         cgosIncome_cgInc_ltcg_prctg125_sec115E_b = cgosIncome_cgInc_ltcg_prctg125_sec115E_b + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec115E_b < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec115E_b
         cgosIncome_cgInc_ltcg_prctg125_sec115E_b = 0
    End If
    
  
    '--6. B11a2 - PTI_other than 112A_12.5%
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125 = cgosIncome_cgInc_ltcg_PTI125 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125
        cgosIncome_cgInc_ltcg_PTI125 = 0
    End If
    
    '5. B1(g)
     If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112_B1g) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125_112_B1g = cgosIncome_cgInc_ltcg_PTI125_112_B1g + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112_B1g < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112_B1g
        cgosIncome_cgInc_ltcg_PTI125_112_B1g = 0
    End If
    
    
    '--4. B3ie - 112 Proviso_12.5%
     '-15th March
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_secProviso) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg125_secProviso = cgosIncome_cgInc_ltcg_prctg125_secProviso + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_secProviso < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_secProviso
        cgosIncome_cgInc_ltcg_prctg125_secProviso = 0
    End If
    
     '--3. B11a1 - PTI_112A_12.5%
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_PTI125_112) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_PTI125_112 = cgosIncome_cgInc_ltcg_PTI125_112 + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_PTI125_112 < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_PTI125_112
        cgosIncome_cgInc_ltcg_PTI125_112 = 0
    End If
    
 
   '--2. B7c - 115AD(1)(iii) proviso
   If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = 0
    End If
  
    '--1. B4c - 112A
    If (amtSetOff >= 0 And Abs(cgosIncome_cgInc_ltcg_prctg125_sec112A) > amtSetOff) Then
        cgosIncome_cgInc_ltcg_prctg125_sec112A = cgosIncome_cgInc_ltcg_prctg125_sec112A + amtSetOff
        amtSetOff = 0
    ElseIf (amtSetOff >= 0 And cgosIncome_cgInc_ltcg_prctg125_sec112A < 0) Then
        amtSetOff = amtSetOff + cgosIncome_cgInc_ltcg_prctg125_sec112A
        cgosIncome_cgInc_ltcg_prctg125_sec112A = 0
    End If
End Function

'LTCG-12.5% by sai on 10/04/2025
Function setOffAgainst125(value As Double, cgosIncome_cgInc_ltcg_prctg125 As Double)
Dim amtSetOff As Double

If (value < 0) Then
     
         'amtSetOff = amtSetOff + setOffAgainst112AB9(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '10 '_Not using
         
         
        
        amtSetOff = amtSetOff + setOffAgainst112_B10(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'B10  15
        amtSetOff = amtSetOff + setOffAgainst112_B9e(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'B9e  14
        amtSetOff = amtSetOff + setOffAgainst112_B5c(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'B5c  13
        amtSetOff = amtSetOff + setOffAgainst112_B2e(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'B2e  12
        amtSetOff = amtSetOff + setOffAgainst112_1_c_125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'B6ie - 112(1)(c )-11
        amtSetOff = amtSetOff + setOffAgainst115AC_125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'B6iie - 115AC(1)(c )-10
        amtSetOff = amtSetOff + setOffAgainst115ACA_125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '9. B3iie - 115ACA(1)(b)
        amtSetOff = amtSetOff + setOffAgainst115AD_125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '8. B6iiie - 115AD(1)(iii)
        amtSetOff = amtSetOff + setOffAgainst115E_b125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '7. B8c - 115E(b)
        amtSetOff = amtSetOff + setOffAgainstPTI125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '6. B11a2 - PTI_other than 112A_12.5%
        amtSetOff = amtSetOff + setOffAgainst112_B1g(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '5. B1(g)
        amtSetOff = amtSetOff + setOffAgainstsecProviso_125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '4. B3ie - 112 Proviso_12.5%
        amtSetOff = amtSetOff + setOffAgainstPTI10_1125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '3. B11a1 - PTI_112A_12.5%
        amtSetOff = amtSetOff + setOffAgainst115AD_3_Proviso_125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '2. B7c - 115AD(1)(iii) proviso
         amtSetOff = amtSetOff + setOffAgainst112A125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) '1. B4c - 112A


         'amtSetOff = amtSetOff + setOffAgainst115E_a125(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'not using
         'amtSetOff = amtSetOff + setOffAgainst112ASEC48(value + amtSetOff, cgosIncome_cgInc_ltcg_prctg125) 'not using
         
         
    
End If

    setOffAgainst125 = amtSetOff
End Function

'by sai on 10/04/2025
Function setOffAgainstsecProviso_20(value As Double, cgosIncome_cgInc_stcg_prctg20 As Double) 'Newly added by Sai on 10/04/2025
Dim amtSetOff As Double
    
    prctg20Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso)

   amtSetOff = 0
    If (value < 0 And cgosIncome_cgInc_ltcg_prctg20_secProviso > 0) Then
        If (cgosIncome_cgInc_ltcg_prctg20_secProviso >= Abs(value)) Then
           cgosIncome_cgInc_ltcg_prctg20_secProviso = cgosIncome_cgInc_ltcg_prctg20_secProviso + (value)
            amtSetOff = Abs(value)
        Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_prctg20_secProviso)
            cgosIncome_cgInc_ltcg_prctg20_secProviso = 0
        End If
        If (amtSetOff > prctg20Sum) Then
            cgosIncome_cgInc_ltcg_prctg20_secProviso = cgosIncome_cgInc_ltcg_prctg20_secProviso + amtSetOff - prctg20Sum
            amtSetOff = prctg20Sum
        End If
    End If
    setOffAgainstsecProviso_20 = amtSetOff

  End Function
'--4th June 26
Function setOffAgainst112_B1g(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double) As Double
  Dim amtSetOff As Double
     
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)

     amtSetOff = 0
     
    If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112_B1g) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112_B1g) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B1g = cgosIncome_cgInc_ltcg_PTI125_112_B1g + (value)
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112_B1g)
            cgosIncome_cgInc_ltcg_PTI125_112_B1g = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B1g = cgosIncome_cgInc_ltcg_PTI125_112_B1g + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst112_B1g = amtSetOff
End Function


Function setOffAgainst112_B2e(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double) As Double
  Dim amtSetOff As Double
     
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)

     amtSetOff = 0
If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112_B2e) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112_B2e) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B2e = cgosIncome_cgInc_ltcg_PTI125_112_B2e + (value)
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112_B2e)
            cgosIncome_cgInc_ltcg_PTI125_112_B2e = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B2e = cgosIncome_cgInc_ltcg_PTI125_112_B2e + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst112_B2e = amtSetOff
End Function

Function setOffAgainst112_B5c(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double) As Double
  Dim amtSetOff As Double
     
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
               cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)

     amtSetOff = 0
If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112_B5c) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112_B5c) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B5c = cgosIncome_cgInc_ltcg_PTI125_112_B5c + (value)
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112_B5c)
            cgosIncome_cgInc_ltcg_PTI125_112_B5c = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B5c = cgosIncome_cgInc_ltcg_PTI125_112_B5c + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst112_B5c = amtSetOff
End Function
Function setOffAgainst112_B9e(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double) As Double
  Dim amtSetOff As Double
     
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)

     amtSetOff = 0
If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112_B9e) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112_B9e) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B9e = cgosIncome_cgInc_ltcg_PTI125_112_B9e + (value)
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112_B9e)
            cgosIncome_cgInc_ltcg_PTI125_112_B9e = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B9e = cgosIncome_cgInc_ltcg_PTI125_112_B9e + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst112_B9e = amtSetOff
End Function
Function setOffAgainst112_B10(value As Double, cgosIncome_cgInc_ltcg_125prctg125 As Double) As Double
  Dim amtSetOff As Double
     
     prctg125Sum = Application.WorksheetFunction.Max(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10)

     amtSetOff = 0
If ((value) < 0 And (cgosIncome_cgInc_ltcg_PTI125_112_B10) > 0) Then
        If (cgosIncome_cgInc_ltcg_PTI125_112_B10) >= Abs(value) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B10 = cgosIncome_cgInc_ltcg_PTI125_112_B10 + (value)
            amtSetOff = Abs(value)
       Else
            amtSetOff = Abs(cgosIncome_cgInc_ltcg_PTI125_112_B10)
            cgosIncome_cgInc_ltcg_PTI125_112_B10 = 0
        End If
        If (amtSetOff) > (prctg125Sum) Then
            cgosIncome_cgInc_ltcg_PTI125_112_B10 = cgosIncome_cgInc_ltcg_PTI125_112_B10 + (amtSetOff) - (prctg125Sum)
            amtSetOff = (prctg125Sum)
        End If
    End If
    setOffAgainst112_B10 = amtSetOff
End Function



