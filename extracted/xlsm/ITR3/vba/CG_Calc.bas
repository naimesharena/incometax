Attribute VB_Name = "CG_Calc"
Option Explicit

'Public cgosIncome_cgInc_stcg_prctg30 As Long

Public cgosIncome_cgInc_ltcg_prctg20 As Double
Public cgosIncome_cgInc_ltcg_prctg10 As Double

Public cgosIncome_cgInc_ltcg_prctg125 As Double 'by sai on 08/04/2025

Public cgosIncome_cgInc_stcg_prctg15 As Double
'added by Riyaz on 03/04/2025
Public cgosIncome_cgInc_stcg_prctg20 As Double
'--
Public cgosIncome_cgInc_stcg_prctg30 As Double
Public cgosIncome_cgInc_stcg_prctgAppRate As Double

Public cgosIncome_cgInc_ltcg_prctg20_sec112 As Double
Public cgosIncome_cgInc_ltcg_prctg20_sec11EA As Double
Public cgosIncome_cgInc_ltcg_PTI20 As Double

'LTCG-20% by sai on 09/04/2025*****************************************
Public cgosIncome_cgInc_ltcg_prctg20_secProviso As Double
Public cgosIncome_cgInc_ltcg_sumof10 As Double
'**********************************************************************


'cgosIncome_cgInc_ltcg_prctg20_sec11EA

Public prctg20Sum As Double

 Public cgosIncome_cgInc_stcg_prctg15_sec111a As Double
 Public cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii As Double
 Public cgosIncome_cgInc_stcg_PTI15 As Double
 
 
 'Added by Sai on 21/03/2025 for AY 2025-26
 Public cgosIncome_cgInc_stcg_prctg20_sec111a As Double
 Public cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii As Double
 Public cgosIncome_cgInc_stcg_PTI20 As Double
 Public cgosIncome_cgInc_stcg_sumof20 As Double
 '**************************************************************
 
 
 
 Public cgosIncome_cgInc_ltcg_prctg10_secProviso As Double
 Public cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 As Double
 Public cgosIncome_cgInc_ltcg_prctg10_sec115E_b  As Double
 Public cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 As Double
 Public cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 As Double
 Public cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 As Double
 
'Newly added by Riyaz on 25/02/2026
 Public cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 As Double
 
 Public cgosIncome_cgInc_ltcg_prctg10_sec112A As Double
 Public cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso As Double
 
 Public cgosIncome_cgInc_ltcg_PTI10 As Double
 Public cgosIncome_cgInc_ltcg_PTI10_112 As Double
 
'LTCG_12.5% by sai on 04/04/2025********************************************************
Public cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec115E_b As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec115E_a As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec112A As Double
Public cgosIncome_cgInc_ltcg_PTI125 As Double
Public cgosIncome_cgInc_ltcg_PTI125_112 As Double
Public cgosIncome_cgInc_ltcg_PTI112_sec48 As Double
Public cgosIncome_cgInc_ltcg_PTI125_112_B9 As Double
'--Newly added by Bindu on 4th June 26 As per V2,3 DE Sheet

Public cgosIncome_cgInc_ltcg_PTI125_112_B1g As Double
Public cgosIncome_cgInc_ltcg_PTI125_112_B2e As Double
Public cgosIncome_cgInc_ltcg_PTI125_112_B5c As Double
Public cgosIncome_cgInc_ltcg_PTI125_112_B9e As Double
Public cgosIncome_cgInc_ltcg_PTI125_112_B10 As Double



Public cgosIncome_cgInc_ltcg_125prctg125 As Double
Public cgosIncome_cgInc_ltcg_prctg125_secProviso As Double

Public amtSetOff111_20 As Double
Public amtSetOff115ad_1_b_ii_20 As Double
Public amtSetOffPTI20_STCG As Double
'******************************************************************************

'LTCG-12.5% by sai on 04/04/2025*****************************
Public cgosIncome_cgInc_ltcg_sumof12_5 As Double
Public cgosIncome_cgInc_stcg_sumofAPP As Double
Public cgosIncome_cgInc_stcg_sumof30 As Double
'************************************************************



 'new stcg30
 Public cgosIncome_cgInc_stcg_prctg30_sec115AD As Double
 Public cgosIncome_cgInc_stcg_PTI30 As Double
 'end
 
 'new Applicable rate PTI
 Public cgosIncome_cgInc_stcg_Applicable As Double
 Public cgosIncome_cgInc_stcg_ApplicablePTI As Double
 'end

Public amtSetOff111a As Double
Public amtSetOff111  As Double
Public amtSetOff115ad_1_b_ii As Double
Public amtSetOffPTI15 As Double

Public prctg15 As Double
Public prsctg20Sum As Double

Public prctg15Sum As Double
Public prcstcg20Sum As Double
Public prctg10Sum As Double
Public prctg125Sum As Double
Public prctg30Sum As Double
Public prctgARSum As Double

'TABLE E EDIATBLE
Public STCG_30 As Double
Public STCG_20 As Double
Public STCG_15 As Double
Public LTCG_20 As Double
Public LTCG_125 As Double
Public LTCG_10 As Double
Public STCG_APP As Double
Public STCG_DTAA As Double
Public LTCG_DTAA As Double

Sub doSetoff()
Application.ScreenUpdating = False
cgosIncome_cgInc_stcg_prctg15_sec111a = Sheet13.Range("Inc_Frm_stcg_111A").value
cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = Sheet13.Range("Inc_Frm_stcg_115AD_1bii").value
cgosIncome_cgInc_stcg_PTI15 = Sheet13.Range("Inc_Frm_stcg_PTI15").value


'20%STCG added by sai on 21/02/2025 for AY 2025-26
'****************************************************************************************************
cgosIncome_cgInc_stcg_prctg20_sec111a = Sheet13.Range("Inc_Frm_stcg_111A_20").value
cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = Sheet13.Range("Inc_Frm_stcg_115AD_1bii_20").value
cgosIncome_cgInc_stcg_PTI20 = Sheet13.Range("Inc_Frm_stcg_PTI15_20").value
'****************************************************************************************************




'LTCG-10%

cgosIncome_cgInc_ltcg_prctg10_secProviso = Sheet13.Range("Inc_Frm_ltcg_10_112proviso").value
cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = Sheet13.Range("Inc_Frm_ltcg_10_115ACA").value
cgosIncome_cgInc_ltcg_prctg10_sec115E_b = Sheet13.Range("Inc_Frm_ltcg_10_115Eb").value
cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = Sheet13.Range("Inc_Frm_ltcg_10_112_1ciii").value
cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = Sheet13.Range("Inc_frm_ltcg_10_115AC").value
cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii").value
'new
cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii_Proviso").value
cgosIncome_cgInc_ltcg_prctg10_sec112A = Sheet13.Range("Inc_Frm_ltcg_10_112A").value
cgosIncome_cgInc_ltcg_PTI10 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10").value
cgosIncome_cgInc_ltcg_PTI10_112 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_112").value
'end

'LTCG-12.5% by sai on 04/04/2025
'**********************************************************************
cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = Sheet13.Range("Inc_Frm_ltcg_10_115ACA_125").value '115ACA(1) - B4iieii
cgosIncome_cgInc_ltcg_prctg125_sec115E_b = Sheet13.Range("Inc_Frm_ltcg_10_115Eb_125").value '115E(b)_B9cii
cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = Sheet13.Range("Inc_Frm_ltcg_10_112_1ciii_125").value '112(1)c(iii) - B7ieii
cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = Sheet13.Range("Inc_frm_ltcg_10_115AC_125").value '115AC(1) - B7iieii
cgosIncome_cgInc_ltcg_prctg125_sec115E_a = Sheet13.Range("Inc_Frm_ltcg_10_115Ea_125").value '115E(a)_B9dii -Not required
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii_Proviso_125").value '115AD(iii)_Proviso - B7cii
cgosIncome_cgInc_ltcg_prctg125_sec112A = Sheet13.Range("Inc_Frm_ltcg_10_112A_125").value '112A - B5cii
cgosIncome_cgInc_ltcg_PTI125 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_125").value 'PTI - 12a2ii
cgosIncome_cgInc_ltcg_PTI125_112 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_112_125").value 'PTI_112 - 12a1ii
cgosIncome_cgInc_ltcg_PTI112_sec48 = Sheet13.Range("Inc_Frm_ltcg_125_112_Sec48").value '112 sec48 - B6ciii - Not required Inc_Frm_ltcg_125_112
 
'cgosIncome_cgInc_ltcg_PTI125_112_B9 = Sheet13.Range("Inc_Frm_ltcg_125_112").value '112 - (B1gb,B10eii,B11ii,B6ciii)
'Commented by Bindu on 4th June 26 as per V3,2 DE Sheet

cgosIncome_cgInc_ltcg_PTI125_112_B1g = Sheet13.Range("Inc_Frm_ltcg_125_112_B1g").value 'B1g
cgosIncome_cgInc_ltcg_PTI125_112_B2e = Sheet13.Range("Inc_Frm_ltcg_125_112_B2e").value 'B2e
cgosIncome_cgInc_ltcg_PTI125_112_B5c = Sheet13.Range("Inc_Frm_ltcg_125_112_B5c").value 'B5c
cgosIncome_cgInc_ltcg_PTI125_112_B9e = Sheet13.Range("Inc_Frm_ltcg_125_112_B9e").value 'B9e
cgosIncome_cgInc_ltcg_PTI125_112_B10 = Sheet13.Range("Inc_Frm_ltcg_125_112_B10").value 'B10




cgosIncome_cgInc_ltcg_prctg125_secProviso = Sheet13.Range("Inc_Frm_ltcg_125_112proviso").value  '112 Proviso(B4iei)

'Newly added by Riyaz on 25/02/2026
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = Sheet13.Range("Inc_Frm_ltcg_125_115ADiii").value
'*******************************************************************************************



'new stcg 30%
cgosIncome_cgInc_stcg_PTI30 = Sheet13.Range("Inc_Frm_stcg_30_PTI30").value
cgosIncome_cgInc_stcg_prctg30_sec115AD = Sheet13.Range("Inc_Frm_stcg_30_115AD").value

'end

'new ltcg 20%
cgosIncome_cgInc_ltcg_prctg20_sec112 = Sheet13.Range("Inc_frm_ltcg_20_112").value

'Commented by sai on 03/03/2025 for AY 2025-26
'cgosIncome_cgInc_ltcg_prctg20_sec11EA = Sheet13.Range("Inc_Frm_ltcg_20_115Ea").value

'LTCG-PTI-20%
cgosIncome_cgInc_ltcg_PTI20 = Sheet13.Range("Inc_Frm_ltcg_20_PTI20").value
'end

cgosIncome_cgInc_ltcg_prctg20_secProviso = Sheet13.Range("Inc_Frm_ltcg_20_112proviso") ' by sai on 09/04/2025

'new Applicable PTI
cgosIncome_cgInc_stcg_Applicable = Sheet13.Range("Inc_Frm_stcg_Applicable").value
cgosIncome_cgInc_stcg_ApplicablePTI = Sheet13.Range("Inc_Frm_stcg_ApplicablePTI").value
'end

cgosIncome_cgInc_stcg_prctgAppRate = Sheet13.Range("IHLA.Eiv1_CurrYearIncome").value
cgosIncome_cgInc_ltcg_prctg20 = Sheet13.Range("IHLA.Evi1_CurrYearIncome").value
cgosIncome_cgInc_stcg_prctg15 = Sheet13.Range("IHLA.Eii1_CurrYearIncome").value
cgosIncome_cgInc_stcg_prctg20 = Sheet13.Range("IHLA.Eiii_CurrYearIncome20").value 'STCG-20%
cgosIncome_cgInc_ltcg_prctg10 = Sheet13.Range("IHLA.Ev1_CurrYearIncome").value

'LTCG-12.5% by sai on 04/04/2025***********************************************************
cgosIncome_cgInc_ltcg_125prctg125 = Sheet13.Range("IHLA.Eix_CurrYearIncome").value
'*******************************************************************************************

cgosIncome_cgInc_stcg_prctg30 = Sheet13.Range("IHLA.Eiii1_CurrYearIncome").value


'LTCG-12.5% by sai on 04/04/2025***********************************************************
cgosIncome_cgInc_stcg_sumof20 = Sheet13.Range("STCG_sumof20").value
cgosIncome_cgInc_ltcg_sumof12_5 = Sheet13.Range("Sumof_12.5").value
'******************************************************************************************
'Added by Riyaz on 08/04/2025
cgosIncome_cgInc_stcg_sumof30 = Sheet13.Range("STCG_sumof30").value
cgosIncome_cgInc_stcg_sumofAPP = Sheet13.Range("STCG_Appli_Rates").value

' by sai on 09/04/2025
cgosIncome_cgInc_ltcg_sumof10 = Sheet13.Range("Sumof_10").value


populateCGTab

STCG_30 = Sheet13.Range("STCG_30").value
STCG_20 = Sheet13.Range("STCG_20").value
STCG_15 = Sheet13.Range("STCG_15").value
LTCG_20 = Sheet13.Range("LTCG_20").value
LTCG_125 = Sheet13.Range("LTCG_12.5").value
LTCG_10 = Sheet13.Range("LTCG_10").value
STCG_APP = Sheet13.Range("STCG_APP").value
STCG_DTAA = Sheet13.Range("STCG_DTAA").value
LTCG_DTAA = Sheet13.Range("LTCG_DTAA").value

If UCase(Sheet13.Range("CG_TableE_Checkbox").value) = UCase("True") Then

    Call setOffAgainst30(STCG_30, cgosIncome_cgInc_stcg_prctg30) 'STCG 30%
    Call setOffAgainstAr(STCG_APP, cgosIncome_cgInc_stcg_prctgAppRate) 'STCG APP RATE
    Call setOffAgainststcg20(STCG_20, cgosIncome_cgInc_stcg_prctg20) 'STCG 20%
    Call setOffAgainst20(LTCG_20, cgosIncome_cgInc_ltcg_prctg125) 'LTCG 20%
    Call setOffAgainst15(STCG_15, cgosIncome_cgInc_stcg_prctg15) 'STCG 15%
    Call setOffAgainst125(LTCG_125, cgosIncome_cgInc_ltcg_125prctg125) 'LTCG 12.5%
    Call setOffAgainst10(LTCG_10, cgosIncome_cgInc_ltcg_prctg10) 'LTCG 10%
    'Call setOffstcgDTAA(Sheet13.Range("IHLA.Ei5_StclSetoffDTAARate").value)  'STCGD DTAA
    'Call setOffltcgDTAA(Sheet13.Range("IHLA.Ei8_LtclSetOffDTAARate").value)  'LTCG DTAA

Else
    
    setOffPctg30Loss (Sheet13.Range("IHLA.Ei3_StclSetoff30Per").value) 'STCG 30% Done -Changes by Riyaz
    setOffPctgArLoss (Sheet13.Range("IHLA.Ei4_StclSetoffAppRate").value) 'STCG AR Done -Changes by Riyaz
'Newly Added by Riyaz on 07/04/2025
    setOffPcstcg20Loss (Sheet13.Range("IHLA.Ei2_StclSetoff20Per").value) 'STCG 20% Done -Changes by Riyaz
    
    setOffPctg20Loss (Sheet13.Range("IHLA.Ei6_LtclSetOff20Per").value) 'LTCG 20% -Changes by Sai
    
    setOffPctg15Loss (Sheet13.Range("IHLA.Ei2_StclSetoff15Per").value)  'STCG 15% Done -Changes by Riyaz
    
'LTCG 12.5% by sai on 07/04/2025**********************************************
    setOffPctg125Loss (Sheet13.Range("IHLA.Ei8_LtclSetOff12.5Per").value) 'LTCG 12.5% -Changes by Sai
'*********************************************************************************
    setOffPctg10Loss (Sheet13.Range("IHLA.Ei5_LtclSetOff10Per").value) 'LTCG 10%-Changes by Sai
    
    
    setOffstcgDTAA (Sheet13.Range("IHLA.Ei5_StclSetoffDTAARate").value) 'STCG DTAA Done -Changes by Riyaz
    setOffltcgDTAA (Sheet13.Range("IHLA.Ei8_LtclSetOffDTAARate").value) 'LTCG DTAA-Changes by Sai
    
End If




'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet31.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Sheet31.Range("temp20Sec112").value = cgosIncome_cgInc_ltcg_prctg20_sec112
    Sheet31.Range("temp20Sec11EA").value = cgosIncome_cgInc_ltcg_prctg20_sec11EA
    
    
'LTCG-20% by sai on 09/04/2025
'***************************************************************************************
    Sheet31.Range("temp20SecPro").value = cgosIncome_cgInc_ltcg_prctg20_secProviso
'***************************************************************************************

    Sheet31.Range("temp15Sec111a").value = cgosIncome_cgInc_stcg_prctg15_sec111a
    Sheet31.Range("temp15Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii

    Sheet31.Range("temp10SecPro").value = cgosIncome_cgInc_ltcg_prctg10_secProviso
    Sheet31.Range("temp10Sec115ACA1").value = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
    Sheet31.Range("temp10Sec115Eb").value = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
    Sheet31.Range("temp10Sec1121c2").value = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
    Sheet31.Range("temp10Sec115AC1").value = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
    Sheet31.Range("temp10Sec115AD3").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
    
    Sheet31.Range("temp10Sec112A").value = cgosIncome_cgInc_ltcg_prctg10_sec112A
    Sheet31.Range("temp10Sec115AD3_Proviso").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
    
     Sheet31.Range("tempPTI10").value = cgosIncome_cgInc_ltcg_PTI10
     Sheet31.Range("tempPTI10_112").value = cgosIncome_cgInc_ltcg_PTI10_112
     Sheet31.Range("tempPTI15").value = cgosIncome_cgInc_stcg_PTI15
     Sheet31.Range("tempPTI20").value = cgosIncome_cgInc_ltcg_PTI20
     Sheet31.Range("tempPTI30").value = cgosIncome_cgInc_stcg_PTI30
     Sheet31.Range("temp115AD30").value = cgosIncome_cgInc_stcg_prctg30_sec115AD

'Added by Riyaz on 03/04/2025
     Sheet31.Range("temp20Sec111a").value = cgosIncome_cgInc_stcg_prctg20_sec111a
     Sheet31.Range("temp20Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
     Sheet31.Range("tempPTI20_STCG").value = cgosIncome_cgInc_stcg_PTI20
    
    
    
    
'LTCG 12.5% by sai on 04/04/2025****************************************************************************
     Sheet31.Range("temp125Sec_115ACA").value = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 '1
     Sheet31.Range("temp125Sec_115EB").value = cgosIncome_cgInc_ltcg_prctg125_sec115E_b '2
     Sheet31.Range("temp125Sec112_c").value = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 '3
     Sheet31.Range("temp125Sec_115AC").value = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 '4
     Sheet31.Range("temp125Sec_115Ea").value = cgosIncome_cgInc_ltcg_prctg125_sec115E_a '5
     Sheet31.Range("temp125Sec_115AD_iii").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso '6
     Sheet31.Range("temp125Sec_112A").value = cgosIncome_cgInc_ltcg_prctg125_sec112A '7
     Sheet31.Range("temp125Sec_PTI").value = cgosIncome_cgInc_ltcg_PTI125 '8
     Sheet31.Range("temp125Sec_112_PTI").value = cgosIncome_cgInc_ltcg_PTI125_112 '9
     Sheet31.Range("temp125Sec_112sec_48").value = cgosIncome_cgInc_ltcg_PTI112_sec48 '10
     'Sheet31.Range("temp125Sec_112").value = cgosIncome_cgInc_ltcg_PTI125_112_B9 '11
     '''--Newly added by Bindu on 4th July 26 as per V2,3 DE Sheet
     Sheet31.Range("temp125Sec_112_B1g").value = cgosIncome_cgInc_ltcg_PTI125_112_B1g '11
     Sheet31.Range("temp125Sec_112_B2e").value = cgosIncome_cgInc_ltcg_PTI125_112_B2e '11
     Sheet31.Range("temp125Sec_112_B5c").value = cgosIncome_cgInc_ltcg_PTI125_112_B5c '11
     Sheet31.Range("temp125Sec_112_B9e").value = cgosIncome_cgInc_ltcg_PTI125_112_B9e '11
     Sheet31.Range("temp125Sec_112_B10").value = cgosIncome_cgInc_ltcg_PTI125_112_B10 '11

     
     
     
     
     Sheet31.Range("temp125SecPro").value = cgosIncome_cgInc_ltcg_prctg125_secProviso '12
'Newly added by Riyaz on 25/02/2026
     Sheet31.Range("temp125Sec115AD3").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
     
'**************************************************************************************************************

'----------------Lock Password-------------------START---
   Sheet31.Protect Password:=sPassword

'----------------Lock Password-------------------END-----

Application.ScreenUpdating = True
End Sub




Function populateCGTab()

Dim neg15 As Double
Dim neg10 As Double

Dim neg125 As Double 'by sai on 04/04/2025

Dim negstcg20 As Double
Dim neg20 As Double
Dim neg30 As Double
Dim negAR As Double
Dim SetOff As Double
Dim amtSetOff As Double

' stcg = cgosIncome.cgInc.stcg
' prctg15 = cgosIncome.cgInc.stcg.prctg15
 prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI20
 neg15 = 0
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg15_sec111a) < 0) Then
        neg15 = neg15 + cgosIncome_cgInc_stcg_prctg15_sec111a
       cgosIncome_cgInc_stcg_prctg15_sec111a = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) < 0) Then
        neg15 = neg15 + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
       cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_PTI15) < 0) Then
        neg15 = neg15 + cgosIncome_cgInc_stcg_PTI15
       cgosIncome_cgInc_stcg_PTI15 = 0
    End If
    
    If (neg15 < 0) Then
    
        SetOff = setOffAgainst115ad_1_b_ii(neg15, cgosIncome_cgInc_stcg_prctg15)
        neg15 = neg15 + SetOff
        
        
        SetOff = setOffAgainst15PTI(neg15, cgosIncome_cgInc_stcg_prctg15)
        neg15 = neg15 + SetOff
        
        SetOff = setOffAgainst111a(neg15, cgosIncome_cgInc_stcg_prctg15)
        neg15 = neg15 + SetOff
        
        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = neg15 + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
    End If
    
    
 'Newly added by Riyaz on 03/04/205
  prsctg20Sum = cgosIncome_cgInc_stcg_prctg20_sec111a + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI20
  negstcg20 = 0
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg20_sec111a) < 0) Then
        negstcg20 = negstcg20 + cgosIncome_cgInc_stcg_prctg20_sec111a
        cgosIncome_cgInc_stcg_prctg20_sec111a = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii) < 0) Then
        negstcg20 = negstcg20 + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_PTI20) < 0) Then
        negstcg20 = negstcg20 + cgosIncome_cgInc_stcg_PTI20
       cgosIncome_cgInc_stcg_PTI20 = 0
    End If
    
'Added by Riyaz on 07/04/2025
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_sumof20) < 0) Then
        negstcg20 = negstcg20 + cgosIncome_cgInc_stcg_sumof20
       cgosIncome_cgInc_stcg_sumof20 = 0
    End If
    
    If (negstcg20 < 0) Then
    
        SetOff = setOffAgainst115ad_1_b_ii_20(negstcg20, cgosIncome_cgInc_stcg_prctg20)
        negstcg20 = negstcg20 + SetOff
        
        
        SetOff = setOffAgainst15PTI_20(negstcg20, cgosIncome_cgInc_stcg_prctg20)
        negstcg20 = negstcg20 + SetOff
        
        SetOff = setOffAgainst111a_20(negstcg20, cgosIncome_cgInc_stcg_prctg20)
        negstcg20 = negstcg20 + SetOff
        
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = negstcg20 + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
         
         End If
    
    


'--End--
' prctg10 = cgosIncome.cgInc.ltcg.prctg10
    prctg10Sum = cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112
                

                     
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
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
       cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec112A) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112A
       cgosIncome_cgInc_ltcg_prctg10_sec112A = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI10) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_PTI10
       cgosIncome_cgInc_ltcg_PTI10 = 0
    End If
    
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI10_112) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_PTI10_112
       cgosIncome_cgInc_ltcg_PTI10_112 = 0
    End If
    
'B(A)10% by sai on 09/04/2025
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_sumof10) < 0) Then
        neg10 = neg10 + cgosIncome_cgInc_ltcg_sumof10
       cgosIncome_cgInc_ltcg_sumof10 = 0
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
        
        
        SetOff = setOffAgainst115AD_3_Proviso(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        
        SetOff = setOffAgainst115E_b(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
         SetOff = setOffAgainstPTI10(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
                
'        SetOff = setOffAgainstsecProviso(neg10, cgosIncome_cgInc_ltcg_prctg10) 'Not using
'        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainstPTI10_112(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        SetOff = setOffAgainst112A(neg10, cgosIncome_cgInc_ltcg_prctg10)
        neg10 = neg10 + SetOff
        
        
        
        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
    End If
        
        
'Newly added by Sai on 04/04/2025 for LTCG 12.5%
'cgosIncome_cgInc_ltcg_sumof12_5
'prctg125Sum = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_PTI125_112_B9 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_sumof12_5 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
                
 prctg125Sum = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + _
                cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_prctg125_secProviso + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 + cgosIncome_cgInc_ltcg_PTI125_112_B1g + cgosIncome_cgInc_ltcg_PTI125_112_B2e + cgosIncome_cgInc_ltcg_PTI125_112_B5c + cgosIncome_cgInc_ltcg_PTI125_112_B9e + cgosIncome_cgInc_ltcg_PTI125_112_B10
               
                    
 neg125 = 0
    
      '--1st 1. B4c - 112A
If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec112A) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec112A
        cgosIncome_cgInc_ltcg_prctg125_sec112A = 0
    End If
    
    
    '--2nd 2. B7c - 115AD(1)(iii) proviso
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
       cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = 0
    End If
    '--3. B11a1 - PTI_112A_12.5%
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112
       cgosIncome_cgInc_ltcg_PTI125_112 = 0
    End If
     '4. B3ie - 112 Proviso_12.5%
    '15th march
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_secProviso) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_secProviso
       cgosIncome_cgInc_ltcg_prctg125_secProviso = 0
    End If
    
    '--5. B1(g)
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B1g) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B1g
       cgosIncome_cgInc_ltcg_PTI125_112_B1g = 0
    End If
    '--6. B11a2 - PTI_other than 112A_12.5%
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125
       cgosIncome_cgInc_ltcg_PTI125 = 0
    End If
     '--7. B8c - 115E(b)
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115E_b) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b
        cgosIncome_cgInc_ltcg_prctg125_sec115E_b = 0
    End If
    
    '--8. B6iiie - 115AD(1)(iii)
    'Newly Added by Riyaz on 25/02/2026
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
       cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = 0
    End If
    '--9. B3iie - 115ACA(1)(b)
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
       cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = 0
    End If
    
   '--10. B6iie - 115AC(1)(c )
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115AC_1) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
       cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = 0
    End If
    '-11. B6ie - 112(1)(c )
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = 0
    End If

     '-12 B2e
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B2e) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B2e
        cgosIncome_cgInc_ltcg_PTI125_112_B2e = 0
    End If
    
     '-'-13 B5c
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B5c) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B5c
        cgosIncome_cgInc_ltcg_PTI125_112_B5c = 0
    End If
    
     '-14 B9e
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B9e) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B9e
        cgosIncome_cgInc_ltcg_PTI125_112_B9e = 0
    End If
    
     '-15 B10
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B10) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B10
        cgosIncome_cgInc_ltcg_PTI125_112_B10 = 0
    End If
    

    
'    '--10th Not using
'     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B9) < 0) Then
'        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B9
'       cgosIncome_cgInc_ltcg_PTI125_112_B9 = 0
'    End If
    
    '-Below 2 not required
'     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115E_a) < 0) Then
'        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a
'       cgosIncome_cgInc_ltcg_prctg125_sec115E_a = 0
'     End If
'
'     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI112_sec48) < 0) Then
'        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI112_sec48
'       cgosIncome_cgInc_ltcg_PTI112_sec48 = 0
'     End If

    '18th March
    'B(A)
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_sumof12_5) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_sumof12_5
       cgosIncome_cgInc_ltcg_sumof12_5 = 0
    End If
  
    If (neg125 < 0) Then
       '--Below section not required
'        SetOff = setOffAgainst115E_a125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'Done
'        neg125 = neg125 + SetOff
'        SetOff = setOffAgainst112ASEC48(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
'        neg125 = neg125 + SetOff
    '--10th - Not using
        'SetOff = setOffAgainst112AB9(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        'neg125 = neg125 + SetOff
        
    ' 15. B10 -
       SetOff = setOffAgainst112_B10(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        ' 14. B9e -
       SetOff = setOffAgainst112_B9e(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
    ' 13. B5c -
       SetOff = setOffAgainst112_B5c(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
    '12. B2e -
       SetOff = setOffAgainst112_B2e(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        
    '--11. B6ie - 112(1)(c )
        SetOff = setOffAgainst112_1_c_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '10. B6iie - 115AC(1)(c )
        SetOff = setOffAgainst115AC_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
         '--9. B3iie - 115ACA(1)(b)
        SetOff = setOffAgainst115ACA_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
'Newly added by Riyaz on 25/02/2026
        '8. B6iiie - 115AD(1)(iii)
        SetOff = setOffAgainst115AD_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        '--7. B8c - 115E(b)
        SetOff = setOffAgainst115E_b125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        
        '--6. B11a2 - PTI_other than 112A_12.5%
        SetOff = setOffAgainstPTI125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        
        '--5. B1(g)
        SetOff = setOffAgainst112_B1g(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        '--4. B3ie - 112 Proviso_12.5%
         '15th March
        SetOff = setOffAgainstsecProviso_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        '--3. B11a1 - PTI_112A_12.5%
        SetOff = setOffAgainstPTI10_1125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        
        '--2. B7c - 115AD(1)(iii) proviso
        SetOff = setOffAgainst115AD_3_Proviso_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'dONE
        neg125 = neg125 + SetOff
        
        '--1. B4c - 112A
        SetOff = setOffAgainst112A125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff

        
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
    End If
        

 'cgosIncome_cgInc_ltcg_prctg20_secProviso newly added by Bindu 15th March 25 as per DE V4
prctg20Sum = cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso
    
    
    
   neg20 = 0
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_sec112) < 0) Then
        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec112
        cgosIncome_cgInc_ltcg_prctg20_sec112 = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_sec11EA) < 0) Then
        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec11EA
        cgosIncome_cgInc_ltcg_prctg20_sec11EA = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI20) < 0) Then
        neg20 = neg20 + cgosIncome_cgInc_ltcg_PTI20
        cgosIncome_cgInc_ltcg_PTI20 = 0
    End If
    
    
'B(A)cgosIncome_cgInc_ltcg_prctg20_secProviso newly added by Sai on 10/04/2025
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_secProviso) < 0) Then
        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_secProviso
        cgosIncome_cgInc_ltcg_prctg20_secProviso = 0
    End If



    If (neg20 < 0) Then
       SetOff = setOffAgainst115Ea(neg20, cgosIncome_cgInc_ltcg_prctg20)
          neg20 = neg20 + SetOff
     
         SetOff = setOffAgainstPTI20(neg20, cgosIncome_cgInc_ltcg_prctg20)
        neg20 = neg20 + SetOff
        
'B(A)cgosIncome_cgInc_ltcg_prctg20_secProviso newly added by sai 10/04/2025
        SetOff = setOffAgainstsecProviso_20(neg20, cgosIncome_cgInc_ltcg_prctg20)
        neg20 = neg20 + SetOff
        
        
        SetOff = setOffAgainst112(neg20, cgosIncome_cgInc_ltcg_prctg20)
        neg20 = neg20 + SetOff
        
        cgosIncome_cgInc_ltcg_prctg20_sec11EA = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec11EA
    End If
    
    
    
 'prctg30 = cgosIncome.cgInc.stcg.prctg30
prctg30Sum = cgosIncome_cgInc_stcg_prctg30_sec115AD + cgosIncome_cgInc_stcg_PTI30
    neg30 = 0
If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg30_sec115AD) < 0) Then
        neg30 = neg30 + cgosIncome_cgInc_stcg_prctg30_sec115AD
        cgosIncome_cgInc_stcg_prctg30_sec115AD = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_PTI30) < 0) Then
        neg30 = neg30 + cgosIncome_cgInc_stcg_PTI30
        cgosIncome_cgInc_stcg_PTI30 = 0
    End If
    
    'Added by Riyaz on 08/04/2025
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_sumof30) < 0) Then
        neg30 = neg30 + cgosIncome_cgInc_stcg_sumof30
       cgosIncome_cgInc_stcg_sumof30 = 0
    End If
    
 If (neg30 < 0) Then
       SetOff = setOffAgainst_stcg115AD(neg30, cgosIncome_cgInc_stcg_prctg30)
          neg30 = neg30 + SetOff
     
         SetOff = setOffAgainst_stcgPTI30(neg30, cgosIncome_cgInc_stcg_prctg30)
        neg30 = neg30 + SetOff
        
        cgosIncome_cgInc_stcg_prctg30_sec115AD = neg30 + cgosIncome_cgInc_stcg_prctg30_sec115AD
    End If


'AR = cgosIncome.cgInc.AR

prctgARSum = cgosIncome_cgInc_stcg_Applicable + cgosIncome_cgInc_stcg_ApplicablePTI
    negAR = 0
If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_Applicable) < 0) Then
        negAR = negAR + cgosIncome_cgInc_stcg_Applicable
        cgosIncome_cgInc_stcg_Applicable = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_ApplicablePTI) < 0) Then
        negAR = negAR + cgosIncome_cgInc_stcg_ApplicablePTI
        cgosIncome_cgInc_stcg_ApplicablePTI = 0
    End If
    
    'Added by Riyaz on 08/04/2025
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_sumofAPP) < 0) Then
        negAR = negAR + cgosIncome_cgInc_stcg_sumofAPP
        cgosIncome_cgInc_stcg_sumofAPP = 0
    End If
    
 If (negAR < 0) Then
       SetOff = setOffAgainst_stcgAR(negAR, cgosIncome_cgInc_stcg_prctgAppRate)
          negAR = negAR + SetOff
     
         SetOff = setOffAgainst_stcgARPTI(negAR, cgosIncome_cgInc_stcg_prctgAppRate)
        negAR = negAR + SetOff
        
        cgosIncome_cgInc_stcg_Applicable = negAR + cgosIncome_cgInc_stcg_Applicable
    End If
    
    
    
End Function

Function TableE_Editable()
Application.EnableEvents = False
If UCase(Sheet13.Range("CG_TableE_Checkbox").value) = UCase("TRUE") Then

'Commented by Riyaz on 19/02/2026
''---Loss STCG 15% Column wise
'
''20% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").Locked = False
'    Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
'
'    '30% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").Locked = False
'    Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").Interior.Color = "&HCCFFCC"
'
'    'APP STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").Locked = False
'    Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").Interior.Color = "&HCCFFCC"
'
'    'DTAA STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev2_DTAArates").Locked = False
'    Sheet13.Range("IHLA.Ev2_DTAArates").Interior.Color = "&HCCFFCC"
'
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev2_StclSetoff15Per").Locked = False
'    Sheet13.Range("IHLA.Ev2_StclSetoff15Per").Interior.Color = "&HCCFFCC"
'
'    '12.5% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").Locked = False
'    Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").Interior.Color = "&HCCFFCC"
'
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi2_StclSetoff15Per").Locked = False
'    Sheet13.Range("IHLA.Evi2_StclSetoff15Per").Interior.Color = "&HCCFFCC"
'
'    'DTAA LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv15_DTAArates").Locked = False
'    Sheet13.Range("IHLA.E_LTCGv15_DTAArates").Interior.Color = "&HCCFFCC"
    
'---Loss STCG 20% Column wise
'Commented by Riyaz on 19/02/2026
''15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii2_StclSetoff20Per").Locked = False
'    Sheet13.Range("IHLA.Eii2_StclSetoff20Per").Interior.Color = "&HCCFFCC"

    '30% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").Locked = False
    Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
    'App STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev2_StclSetoff20Per").Locked = False
    Sheet13.Range("IHLA.Ev2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
    'DTAA STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Evi2_StclSetoff20Per").Locked = False
    Sheet13.Range("IHLA.Evi2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evii2_StclSetoff20Per").Locked = False
'    Sheet13.Range("IHLA.Evii2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").Locked = False
    Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eix2_StclSetoff20Per").Locked = False
'    Sheet13.Range("IHLA.Eix2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ex2_StclSetoff20Per").Locked = False
    Sheet13.Range("IHLA.Ex2_StclSetoff20Per").Interior.Color = "&HCCFFCC"
    
'---Loss STCG 30% Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii3_StclSetoff30Per").Locked = False
'    Sheet13.Range("IHLA.Eii3_StclSetoff30Per").Interior.Color = "&HCCFFCC"
    
    '20% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").Locked = False
    Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").Interior.Color = "&HCCFFCC"
    
    'APP STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").Locked = False
    Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").Interior.Color = "&HCCFFCC"
    
    'DTAA STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev3_DTAArates").Locked = False
    Sheet13.Range("IHLA.Ev3_DTAArates").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev3_StclSetoff30Per").Locked = False
'    Sheet13.Range("IHLA.Ev3_StclSetoff30Per").Interior.Color = "&HCCFFCC"
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").Locked = False
    Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi3_StclSetoff30Per").Locked = False
'    Sheet13.Range("IHLA.Evi3_StclSetoff30Per").Interior.Color = "&HCCFFCC"
    
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGv30_DTAArates").Locked = False
    Sheet13.Range("IHLA.E_LTCGv30_DTAArates").Interior.Color = "&HCCFFCC"

'---Loss STCG APP Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").Locked = False
'    Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").Interior.Color = "&HCCFFCC"
    
    '20% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").Locked = False
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").Interior.Color = "&HCCFFCC"
    
    '30% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").Locked = False
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").Interior.Color = "&HCCFFCC"
    
    'DTAA STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev4_DTAArates").Locked = False
    Sheet13.Range("IHLA.Ev4_DTAArates").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").Locked = False
'    Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").Interior.Color = "&HCCFFCC"
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").Locked = False
    Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").Locked = False
'    Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").Interior.Color = "&HCCFFCC"
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").Locked = False
    Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").Interior.Color = "&HCCFFCC"
    
'---Loss STCG DTAA Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").Locked = False
'    Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").Interior.Color = "&HCCFFCC"
    
    '20% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").Locked = False
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").Interior.Color = "&HCCFFCC"
    
    '30% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").Locked = False
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").Interior.Color = "&HCCFFCC"
    
    'APP STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").Locked = False
    Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").Locked = False
'    Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").Interior.Color = &HCCFFCC
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").Locked = False
    Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").Locked = False
'    Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").Interior.Color = "&HCCFFCC"
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").Locked = False
    Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
''---Loss LTCG 10% Column wise
'    '12.5% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").Locked = False
'    Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").Interior.Color = "&HCCFFCC"
'
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").Locked = False
'    Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").Interior.Color = "&HCCFFCC"
'
'    'DTAA LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv5_DTAArates").Locked = False
'    Sheet13.Range("IHLA.E_LTCGv5_DTAArates").Interior.Color = "&HCCFFCC"

'---Loss LTCG 12.5% Column wise
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").Locked = False
'    Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").Locked = False
'    Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").Interior.Color = "&HCCFFCC"
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").Locked = False
    Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
''---Loss LTCG 20% Column wise
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").Locked = False
'    Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").Interior.Color = "&HCCFFCC"
'
'    '12.5% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").Locked = False
'    Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").Interior.Color = "&HCCFFCC"
    
'    'DTAA LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv6_DTAArates").Locked = False
'    Sheet13.Range("IHLA.E_LTCGv6_DTAArates").Interior.Color = "&HCCFFCC"
    
    
'---Loss LTCG DTAA Column wise
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").Locked = False
'    Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").Interior.Color = "&HCCFFCC"
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").Locked = False
    Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").Interior.Color = "&HCCFFCC"
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").Locked = False
'    Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").Interior.Color = "&HCCFFCC"


'Commented by Riyaz on 19/02/2026
''--Loss STCG 15% Column wise
'If Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").value = 0 Then 'STCG 20%
'       Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").value = 0
'End If
'If Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").value = 0 Then 'STCG 30%
'       Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").value = 0
'End If
'If Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").value = 0 Then 'STCG APP
'       Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").value = 0
'End If
'If Sheet13.Range("IHLA.Ev2_DTAArates").value = 0 Then 'STCG DTAA
'       Sheet13.Range("IHLA.Ev2_DTAArates").value = 0
'End If
'If Sheet13.Range("IHLA.Ev2_StclSetoff15Per").value = 0 Then ''10% LTCG
'       Sheet13.Range("IHLA.Ev2_StclSetoff15Per").value = 0
'End If
'If Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").value = 0 Then ''12.5% LTCG
'       Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").value = 0
'End If
'If Sheet13.Range("IHLA.Evi2_StclSetoff15Per").value = 0 Then ''20% LTCG
'       Sheet13.Range("IHLA.Evi2_StclSetoff15Per").value = 0
'End If
'If Sheet13.Range("IHLA.E_LTCGv15_DTAArates").value = 0 Then ''DTAA LTCG
'       Sheet13.Range("IHLA.E_LTCGv15_DTAArates").value = 0
'End If

'---Loss STCG 20% Column wise
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Eii2_StclSetoff20Per").value = 0 Then '15% STCG
'       Sheet13.Range("IHLA.Eii2_StclSetoff20Per").value = 0
'End If
If Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").value = 0 Then '30% STCG
       Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").value = 0
End If
If Sheet13.Range("IHLA.Ev2_StclSetoff20Per").value = 0 Then 'App STCG
       Sheet13.Range("IHLA.Ev2_StclSetoff20Per").value = 0
End If
If Sheet13.Range("IHLA.Evi2_StclSetoff20Per").value = 0 Then 'DTAA STCG
       Sheet13.Range("IHLA.Evi2_StclSetoff20Per").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Evii2_StclSetoff20Per").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Evii2_StclSetoff20Per").value = 0
'End If
If Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").value = 0 Then '12.5% LTCG
       Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Eix2_StclSetoff20Per").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Eix2_StclSetoff20Per").value = 0
'End If
If Sheet13.Range("IHLA.Ex2_StclSetoff20Per").value = 0 Then 'DTAA LTCG
       Sheet13.Range("IHLA.Ex2_StclSetoff20Per").value = 0
End If

'---Loss STCG 30% Column wise
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Eii3_StclSetoff30Per").value = 0 Then '15% STCG
'       Sheet13.Range("IHLA.Eii3_StclSetoff30Per").value = 0
'End If
If Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").value = 0 Then '20% STCG
       Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").value = 0
End If
If Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").value = 0 Then 'APP STCG
       Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").value = 0
End If
If Sheet13.Range("IHLA.Ev3_DTAArates").value = 0 Then 'DTAA STCG
       Sheet13.Range("IHLA.Ev3_DTAArates").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value = 0
'End If
If Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").value = 0 Then '12.5% LTCG
       Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value = 0
'End If
If Sheet13.Range("IHLA.E_LTCGv30_DTAArates").value = 0 Then 'DTAA LTCG
       Sheet13.Range("IHLA.E_LTCGv30_DTAArates").value = 0
End If

'---Loss STCG APP Column wise
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").value = 0 Then '15% STCG
'       Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").value = 0
'End If
If Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").value = 0 Then '20% STCG
       Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").value = 0
End If
If Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").value = 0 Then '30% STCG
       Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").value = 0
End If
If Sheet13.Range("IHLA.Ev4_DTAArates").value = 0 Then 'DTAA STCG
       Sheet13.Range("IHLA.Ev4_DTAArates").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").value = 0
'End If
If Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").value = 0 Then '12.5% LTCG
       Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").value = 0
'End If
If Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").value = 0 Then 'DTAA LTCG
       Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").value = 0
End If

'---Loss STCG DTAA Column wise
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").value = 0 Then '15% STCG
'       Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").value = 0
'End If
If Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").value = 0 Then '20% STCG
       Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").value = 0
End If
If Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").value = 0 Then '30% STCG
       Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").value = 0
End If
If Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").value = 0 Then 'APP STCG
       Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").value = 0
'End If
If Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").value = 0 Then '12.5% LTCG
       Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").value = 0
'End If
If Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value = 0 Then 'DTAA LTCG
       Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value = 0
End If
'Commented by Riyaz on 19/02/2026
''---Loss LTCG 10% Column wise
'If Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").value = 0 Then '12.5% LTCG
'       Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").value = 0
'End If
'If Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value = 0
'End If
'If Sheet13.Range("IHLA.E_LTCGv5_DTAArates").value = 0 Then 'DTAA LTCG
'       Sheet13.Range("IHLA.E_LTCGv5_DTAArates").value = 0
'End If

'---Loss LTCG 12.5% Column wise
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").value = 0
'End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").value = 0
'End If
If Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").value = 0 Then 'DTAA LTCG
       Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").value = 0
End If

'Commented by Riyaz on 19/02/2026
''---Loss LTCG 20% Column wise
'If Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value = 0
'End If
'If Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").value = 0 Then '12.5% LTCG
'       Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").value = 0
'End If
'If Sheet13.Range("IHLA.E_LTCGv6_DTAArates").value = 0 Then 'DTAA LTCG
'       Sheet13.Range("IHLA.E_LTCGv6_DTAArates").value = 0
'End If

'---Loss LTCG DTAA Column wise
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").value = 0 Then '10% LTCG
'       Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").value = 0
'End If
If Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").value = 0 Then '12.5% LTCG
       Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").value = 0
End If
'Commented by Riyaz on 19/02/2026
'If Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").value = 0 Then '20% LTCG
'       Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").value = 0
'End If


Else

'Commented by Riyaz on 19/02/2026
''---Loss STCG 15% Column wise
'    '20% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per-IHLA.Eiv2_StclSetoff15Per,IHLA.Eiii_CurrYearIncome20-IHLA.Eiii3_StclSetoff30Per-IHLA.Eiii4_StclSetoffAppRate20)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").Locked = True
'
'    '30% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").formula = "=MIN(IHLA.Ei2_StclSetoff15Per,IHLA.Eiii1_CurrYearIncome-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiv2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").Locked = True
'
'    'APP STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per,IHLA.Eiv1_CurrYearIncome-IHLA.Eiv3_StclSetoff30Per-IHLA.Ev2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").Locked = True
'
'    'DTAA STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev2_DTAArates").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per-IHLA.Eiv2_StclSetoff15Per-IHLA.Eiii2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per-IHLA.Eviii2_StclSetoff12.5Per-IHLA.Ev2_StclSetoff15Per,IHLA.Ev_DTAArates-IHLA.Ev3_DTAArates-IHLA.Ev4_DTAArates-IHLA.Evi2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev2_DTAArates").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev2_DTAArates").Locked = True
'
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev2_StclSetoff15Per").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per-IHLA.Eiv2_StclSetoff15Per-IHLA.Eiii2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per-IHLA.Eviii2_StclSetoff12.5Per,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per-IHLA.Ev4_StclSetoffAppRate-IHLA.Evii2_StclSetoff20Per-IHLA.Ev6_LtclSetOff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev2_StclSetoff15Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev2_StclSetoff15Per").Locked = True
'
'    '12.5% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per-IHLA.Eiv2_StclSetoff15Per-IHLA.Eiii2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per-IHLA.Eviii4_StclSetoffAppRate12-IHLA.Eviii2_StclSetoff20Per-IHLA.Eviii6_LtclSetOff2012Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").Locked = True
'
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi2_StclSetoff15Per").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per-IHLA.Eiv2_StclSetoff15Per-IHLA.Eiii2_StclSetoff20Per,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per-IHLA.Evi4_StclSetoffAppRate-IHLA.Eix2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi2_StclSetoff15Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evi2_StclSetoff15Per").Locked = True
'
'    'DTAA LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv15_DTAArates").formula = "=MIN(IHLA.Ei2_StclSetoff15Per-IHLA.Eiii2_StclSetoff15Per-IHLA.Eiv2_StclSetoff15Per-IHLA.Eiii2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per-IHLA.Eviii2_StclSetoff12.5Per-IHLA.Ev2_StclSetoff15Per-IHLA.Ev2_DTAArates,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates-IHLA.E_LTCGvapplicable_DTAArates-IHLA.Ex2_StclSetoff20Per-IHLA.E_LTCGv6_DTAArates)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv15_DTAArates").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.E_LTCGv15_DTAArates").Locked = True
    
'---Loss STCG 20% Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per-IHLA.Ev2_StclSetoff20Per-IHLA.Eix2_StclSetoff20Per,IHLA.Eii1_CurrYearIncome-IHLA.Eii3_StclSetoff30Per-IHLA.Eii4_StclSetoffAppRate)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii2_StclSetoff20Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eii2_StclSetoff20Per").Locked = True
    
    '30% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per,IHLA.Eiii1_CurrYearIncome-IHLA.Eiii4_StclSetoffAppRate)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").Locked = True
    
    'App STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per,IHLA.Eiv1_CurrYearIncome-IHLA.Eiv3_StclSetoff30Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev2_StclSetoff20Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Ev2_StclSetoff20Per").Locked = True
    
    'DTAA STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Evi2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per-IHLA.Ev2_StclSetoff20Per-IHLA.Eviii2_StclSetoff20Per,IHLA.Ev_DTAArates-IHLA.Ev3_DTAArates-IHLA.Ev4_DTAArates)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Evi2_StclSetoff20Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Evi2_StclSetoff20Per").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evii2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per-IHLA.Ev2_StclSetoff20Per-IHLA.Eix2_StclSetoff20Per-IHLA.Eii2_StclSetoff20Per-IHLA.Eviii2_StclSetoff20Per,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per-IHLA.Ev4_StclSetoffAppRate)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evii2_StclSetoff20Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evii2_StclSetoff20Per").Locked = True
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per-IHLA.Ev2_StclSetoff20Per,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per-IHLA.Eviii4_StclSetoffAppRate12)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eix2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per-IHLA.Ev2_StclSetoff20Per,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per-IHLA.Evi4_StclSetoffAppRate)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eix2_StclSetoff20Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eix2_StclSetoff20Per").Locked = True
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ex2_StclSetoff20Per").formula = "=MIN(IHLA.Ei2_StclSetoff20Per-IHLA.Eiv2_StclSetoff20Per-IHLA.Ev2_StclSetoff20Per-IHLA.Eviii2_StclSetoff20Per-IHLA.Evi2_StclSetoff20Per,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates-IHLA.E_LTCGvapplicable_DTAArates)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ex2_StclSetoff20Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Ex2_StclSetoff20Per").Locked = True


'---Loss STCG 30% Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii3_StclSetoff30Per").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per-IHLA.Eiii3_StclSetoff30Per-IHLA.Evi3_StclSetoff30Per,IHLA.Eii1_CurrYearIncome)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii3_StclSetoff30Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eii3_StclSetoff30Per").Locked = True
    
    '20% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per,IHLA.Eiii_CurrYearIncome20)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").Locked = True
    
    'APP STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").formula = "=MIN(IHLA.Ei3_StclSetoff30Per,IHLA.Eiv1_CurrYearIncome)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").Locked = True
    
    'DTAA STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev3_DTAArates").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per-IHLA.Eiii3_StclSetoff30Per-IHLA.Eviii3_StclSetoff3012Per,IHLA.Ev_DTAArates)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev3_DTAArates").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Ev3_DTAArates").Locked = True

'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev3_StclSetoff30Per").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per-IHLA.Eiii3_StclSetoff30Per-IHLA.Evi3_StclSetoff30Per-IHLA.Eii3_StclSetoff30Per-IHLA.Eviii3_StclSetoff3012Per,IHLA.Ev1_CurrYearIncome)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev3_StclSetoff30Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev3_StclSetoff30Per").Locked = True
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per-IHLA.Eiii3_StclSetoff30Per,IHLA.Eix_CurrYearIncome)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").Locked = True

'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi3_StclSetoff30Per").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per-IHLA.Eiii3_StclSetoff30Per,IHLA.Evi1_CurrYearIncome)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi3_StclSetoff30Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evi3_StclSetoff30Per").Locked = True
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGv30_DTAArates").formula = "=MIN(IHLA.Ei3_StclSetoff30Per-IHLA.Eiv3_StclSetoff30Per-IHLA.Eiii3_StclSetoff30Per-IHLA.Eviii3_StclSetoff3012Per-IHLA.Ev3_DTAArates,IHLA.E_LTCGv_DTAArates)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGv30_DTAArates").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.E_LTCGv30_DTAArates").Locked = True
    
'---Loss STCG APP Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate20-IHLA.Evi4_StclSetoffAppRate,IHLA.Eii1_CurrYearIncome-IHLA.Eii3_StclSetoff30Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").Locked = True
    
    
    '20% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate,IHLA.Eiii_CurrYearIncome20-IHLA.Eiii3_StclSetoff30Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").Locked = True
    
    '30% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate,IHLA.Eiii1_CurrYearIncome)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").Locked = True
    
    'DTAA STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev4_DTAArates").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate20-IHLA.Eviii4_StclSetoffAppRate12,IHLA.Ev_DTAArates-IHLA.Ev3_DTAArates)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ev4_DTAArates").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Ev4_DTAArates").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate20-IHLA.Evi4_StclSetoffAppRate-IHLA.Eii4_StclSetoffAppRate-IHLA.Eviii4_StclSetoffAppRate12,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").Locked = True
    
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate20,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate20,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").Locked = True
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").formula = "=MIN(IHLA.Ei4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiii4_StclSetoffAppRate20-IHLA.Eviii4_StclSetoffAppRate12-IHLA.Ev4_DTAArates,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").Locked = True
    
'---Loss STCG DTAA Column wise
'Commented by Riyaz on 19/02/2026
'    '15% STCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate-IHLA.Eiv5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate20-IHLA.Evi5_StclSetoffDTAARate,IHLA.Eii1_CurrYearIncome-IHLA.Eii3_StclSetoff30Per-IHLA.Eii4_StclSetoffAppRate-IHLA.Eii2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").Locked = True
    
    '20% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate-IHLA.Eiv5_StclSetoffDTAARate,IHLA.Eiii_CurrYearIncome20-IHLA.Eiii3_StclSetoff30Per-IHLA.Eiii4_StclSetoffAppRate20)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").Locked = True
    
    '30% STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate,IHLA.Eiii1_CurrYearIncome-IHLA.Eiii4_StclSetoffAppRate-IHLA.Eiv2_StclSetoff20Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").Locked = True
    
    
    'APP STCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate,IHLA.Eiv1_CurrYearIncome-IHLA.Eiv3_StclSetoff30Per-IHLA.Ev2_StclSetoff20Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate-IHLA.Eiv5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate20-IHLA.Evi5_StclSetoffDTAARate-IHLA.Eii5_StclSetoffDTAARate-IHLA.Eviii5_StclSetoffDTAARate12,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per-IHLA.Ev4_StclSetoffAppRate-IHLA.Evii2_StclSetoff20Per-IHLA.Ev6_LtclSetOff20Per-IHLA.Ev2_StclSetoff15Per-IHLA.Evii8_LtclSetOff12.5Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").Locked = True
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate-IHLA.Eiv5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate20,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per-IHLA.Eviii4_StclSetoffAppRate12-IHLA.Eviii2_StclSetoff20Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate-IHLA.Eiv5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate20,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per-IHLA.Evi4_StclSetoffAppRate-IHLA.Eix2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per-IHLA.Eix8_LtclSetOff12.5Per-IHLA.Evi5_LtclSetOff10Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").Locked = True
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").formula = "=MIN(IHLA.Ei5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate-IHLA.Eiv5_StclSetoffDTAARate-IHLA.Eiii5_StclSetoffDTAARate20-IHLA.Eviii5_StclSetoffDTAARate12,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates-IHLA.E_LTCGvapplicable_DTAArates-IHLA.Ex2_StclSetoff20Per-IHLA.Ex8_LtclSetOff12.5Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").Locked = True

'Commented by Riyaz on 19/02/2026
''---Loss LTCG 10% Column wise
'    '12.5% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").formula = "=MIN(IHLA.Ei5_LtclSetOff10Per-IHLA.Evi5_LtclSetOff10Per,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per-IHLA.Eviii4_StclSetoffAppRate12-IHLA.Eviii2_StclSetoff20Per-IHLA.Eviii6_LtclSetOff2012Per-IHLA.Eviii2_StclSetoff12.5Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").Locked = True
'
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").formula = "=MIN(IHLA.Ei5_LtclSetOff10Per,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per-IHLA.Evi4_StclSetoffAppRate-IHLA.Eix2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per-IHLA.Eix8_LtclSetOff12.5Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").Locked = True
'
'
'    'DTAA LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv5_DTAArates").formula = "=MIN(IHLA.Ei5_LtclSetOff10Per-IHLA.Evi5_LtclSetOff10Per-IHLA.Eviii5_LtclSetOff1012Per,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates-IHLA.E_LTCGvapplicable_DTAArates-IHLA.Ex2_StclSetoff20Per-IHLA.E_LTCGv6_DTAArates-IHLA.E_LTCGv15_DTAArates-IHLA.Ex8_LtclSetOff12.5Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv5_DTAArates").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.E_LTCGv5_DTAArates").Locked = True
    
    
'---Loss LTCG 12.5% Column wise
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").formula = "=MIN(IHLA.Ei8_LtclSetOff12.5Per-IHLA.Eix8_LtclSetOff12.5Per,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per-IHLA.Ev4_StclSetoffAppRate-IHLA.Evii2_StclSetoff20Per-IHLA.Ev6_LtclSetOff20Per-IHLA.Ev2_StclSetoff15Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").Locked = True
'
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").formula = "=MIN(IHLA.Ei8_LtclSetOff12.5Per,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per-IHLA.Evi4_StclSetoffAppRate-IHLA.Eix2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").Locked = True
    
    'DTAA LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").formula = "=MIN(IHLA.Ei8_LtclSetOff12.5Per,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates-IHLA.E_LTCGvapplicable_DTAArates-IHLA.Ex2_StclSetoff20Per)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").Locked = True
    
'Commented by Riyaz on 19/02/2026
''---Loss LTCG 20% Column wise
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").formula = "=MIN(IHLA.Ei6_LtclSetOff20Per-IHLA.Eviii6_LtclSetOff2012Per,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per-IHLA.Ev4_StclSetoffAppRate-IHLA.Evii2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").Locked = True
'
'    '12.5% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").formula = "=MIN(IHLA.Ei6_LtclSetOff20Per,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per-IHLA.Eviii4_StclSetoffAppRate12-IHLA.Eviii2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").Locked = True
'
'    'DTAA LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv6_DTAArates").formula = "=MIN(IHLA.Ei6_LtclSetOff20Per-IHLA.Eviii6_LtclSetOff2012Per-IHLA.Ev6_LtclSetOff20Per,IHLA.E_LTCGv_DTAArates-IHLA.E_LTCGv30_DTAArates-IHLA.E_LTCGvapplicable_DTAArates-IHLA.Ex2_StclSetoff20Per)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.E_LTCGv6_DTAArates").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.E_LTCGv6_DTAArates").Locked = True
    
    
'---Loss LTCG DTAA Column wise
'Commented by Riyaz on 19/02/2026
'    '10% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").formula = "=MIN(IHLA.Ei8_LtclSetOffDTAARate-IHLA.Evi8_LtclSetOffDTAARate-IHLA.Eviii8_LtclSetOffDTAARate12,IHLA.Ev1_CurrYearIncome-IHLA.Ev3_StclSetoff30Per-IHLA.Ev4_StclSetoffAppRate-IHLA.Evii2_StclSetoff20Per-IHLA.Ev6_LtclSetOff20Per-IHLA.Ev2_StclSetoff15Per-IHLA.Evii8_LtclSetOff12.5Per-IHLA.Ev5_StclSetoffDTAARate)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").Locked = True
    
    
    '12.5% LTCG
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").formula = "=MIN(IHLA.Ei8_LtclSetOffDTAARate,IHLA.Eix_CurrYearIncome-IHLA.Eviii3_StclSetoff3012Per-IHLA.Eviii4_StclSetoffAppRate12-IHLA.Eviii2_StclSetoff20Per-IHLA.Eviii5_StclSetoffDTAARate12)"
    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").Interior.ColorIndex = 2
    Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").Locked = True
    
'Commented by Riyaz on 19/02/2026
'    '20% LTCG
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").formula = "=MIN(IHLA.Ei8_LtclSetOffDTAARate,IHLA.Evi1_CurrYearIncome-IHLA.Evi3_StclSetoff30Per-IHLA.Evi4_StclSetoffAppRate-IHLA.Eix2_StclSetoff20Per-IHLA.Evi2_StclSetoff15Per-IHLA.Eix8_LtclSetOff12.5Per-IHLA.Evi5_LtclSetOff10Per-IHLA.Evi5_StclSetoffDTAARate)"
'    Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").Interior.ColorIndex = 2
'    Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").Locked = True
End If



End Function
Sub LinkCheckBoxes_CG() 'Newly added by Bindu on 19th May 25
Dim chk As CheckBox
Dim lCol As Long

lCol = 11 'number of columns to the right for link(AA)
'lCol = 10 'number of columns to the right for link(AA)'updated by Sadineni on 27/02/2026


Sheet13.Unprotect Password:=EfilingCommon.getmsgstate


For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = _
         .TopLeftCell.Offset(0, lCol).Address
   End With
Next chk

TableE_Editable

Sheet13.Protect Password:=EfilingCommon.getmsgstate


End Sub

