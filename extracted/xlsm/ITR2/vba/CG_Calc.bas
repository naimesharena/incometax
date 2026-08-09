Attribute VB_Name = "CG_Calc"

Option Explicit
'Public cgosIncome_cgInc_stcg_prctg30 As Long
 
Public cgosIncome_cgInc_ltcg_prctg20 As Double
Public cgosIncome_cgInc_ltcg_prctg10 As Double
Public cgosIncome_cgInc_stcg_prctg15 As Double
Public cgosIncome_cgInc_ltcg_prctg125 As Double
Public cgosIncome_cgInc_stcg_prctg20 As Double
Public cgosIncome_cgInc_stcg_prctg30 As Double
Public cgosIncome_cgInc_stcg_prctgAppRate As Double
 
Public cgosIncome_cgInc_ltcg_prctg20_sec112 As Double
Public cgosIncome_cgInc_ltcg_prctg20_sec11EA As Double
Public cgosIncome_cgInc_ltcg_PTI20 As Double
'Newly added by Bindu on 15th March 25 ad per DE V4
Public cgosIncome_cgInc_ltcg_prctg20_secProviso As Double
 
'cgosIncome_cgInc_ltcg_prctg20_sec11EA
 
Public prctg20Sum As Double
 
Public cgosIncome_cgInc_stcg_prctg15_sec111a As Double
Public cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii As Double
Public cgosIncome_cgInc_stcg_PTI15 As Double
'Added by Riyaz 18/02/2025
Public cgosIncome_cgInc_stcg_prctg20_sec111a As Double
Public cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii As Double
Public cgosIncome_cgInc_stcg_PTI20 As Double
'Added by Riyaz on 21/02/2025
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
Public cgosIncome_cgInc_ltcg_125prctg125 As Double
Public cgosIncome_cgInc_ltcg_prctg125_secProviso As Double
Public cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 As Double

Public amtSetOff111_20 As Double
Public amtSetOff115ad_1_b_ii_20 As Double
Public amtSetOffPTI20_STCG As Double
 

Public cgosIncome_cgInc_ltcg_prctg10_secProviso As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec115E_b  As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec112A As Double
Public cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso As Double
Public cgosIncome_cgInc_ltcg_PTI10 As Double
Public cgosIncome_cgInc_ltcg_PTI10_112 As Double
'new stcg30
Public cgosIncome_cgInc_stcg_prctg30_sec115AD As Double
Public cgosIncome_cgInc_stcg_PTI30 As Double
'end
'new Applicable rate PTI
Public cgosIncome_cgInc_stcg_Applicable As Double
Public cgosIncome_cgInc_stcg_ApplicablePTI As Double
'end
Public cgosIncome_cgInc_stcg_sumof20 As Double
Public cgosIncome_cgInc_stcg_sumof30 As Double
Public cgosIncome_cgInc_stcg_sumofAPP As Double
Public cgosIncome_cgInc_ltcg_sumof12_5 As Double
Public cgosIncome_cgInc_ltcg_sumof10 As Double

Public amtSetOff111a As Double
Public amtSetOff111  As Double
Public amtSetOff115ad_1_b_ii As Double
Public amtSetOffPTI15 As Double
Public prctg15 As Double
Public prctg15Sum As Double
Public prcstcg20Sum As Double
Public prctg10Sum As Double
Public prctg125Sum As Double
Public prctg30Sum As Double
Public prctgARSum As Double
 
'Newly added by Bindu
Public rngname_account_ifsc As Range
Public end_account_ifsc
 
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
 
'CYLA-BFLA EDITABLE
Public SI_LTCG125_115AD_125 As Double 'added by sai on 13/03/2026

Public BFLA_STCG_15, SI_STCG15_111A, SI_STCG15_PTI, SI_STCG15_115ad_1_b_ii As Double
Public BFLA_STCG_20, SI_STCG20_111A, SI_STCG20_PTI, SI_STCG20_115ad_1_b_ii As Double
Public BFLA_STCG_30, SI_STCG30_115AD, SI_STCG30_PTI As Double
Public BFLA_LTCG_10, SI_LTCG10_112A, SI_LTCG10_PTI112, SI_LTCG10_PTI, SI_LTCG10_115EB, SI_LTCG10_115AD3_Pro, SI_LTCG10_115AD3, SI_LTCG10_115ACA, SI_LTCG10_115AC, SI_LTCG10_1121C As Double
Public BFLA_LTCG_125, SI_LTCG125_112A, SI_LTCG125_PTI112, SI_LTCG125_112_Pro, SI_LTCG125_PTI, SI_LTCG125_115EB, SI_LTCG125_115AD3_Pro, SI_LTCG125_115ACA, SI_LTCG125_115AC, SI_LTCG125_1121C, SI_LTCG125_112 As Double
Public BFLA_LTCG_20, SI_LTCG20_112, SI_LTCG20_PTI, SI_LTCG20_112_PRO As Double



Sub doSetoff()
Application.ScreenUpdating = False
'STCG 15%
'Commented by Riyaz on 29/01/2026
'cgosIncome_cgInc_stcg_prctg15_sec111a = Sheet13.Range("Inc_Frm_stcg_111A").value
'cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = Sheet13.Range("Inc_Frm_stcg_115AD_1bii").value
'cgosIncome_cgInc_stcg_PTI15 = Sheet13.Range("Inc_Frm_stcg_PTI15").value
 
'--20%STCG added by Riyaz on 18/02/2025
cgosIncome_cgInc_stcg_prctg20_sec111a = Sheet13.Range("Inc_Frm_stcg_111A_20").value
cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = Sheet13.Range("Inc_Frm_stcg_115AD_1bii_20").value
cgosIncome_cgInc_stcg_PTI20 = Sheet13.Range("Inc_Frm_stcg_PTI15_20").value
 
 
'LTCG 10%
' cgosIncome_cgInc_ltcg_prctg10_secProviso NOt using  commented  by Bindu as this sec changed from 10 to 20
'cgosIncome_cgInc_ltcg_prctg10_secProviso = Sheet13.Range("Inc_Frm_ltcg_10_112proviso").value

cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = Sheet13.Range("Inc_Frm_ltcg_10_115ACA").value
cgosIncome_cgInc_ltcg_prctg10_sec115E_b = Sheet13.Range("Inc_Frm_ltcg_10_115Eb").value
cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = Sheet13.Range("Inc_Frm_ltcg_10_112_1ciii").value
cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = Sheet13.Range("Inc_frm_ltcg_10_115AC").value
cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii").value
cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii_Proviso").value
cgosIncome_cgInc_ltcg_prctg10_sec112A = Sheet13.Range("Inc_Frm_ltcg_10_112A").value
cgosIncome_cgInc_ltcg_PTI10 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10").value
cgosIncome_cgInc_ltcg_PTI10_112 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_112").value
 
 
'LTCG 12.5% added by Riyaz on 21/02/2025
cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = Sheet13.Range("Inc_Frm_ltcg_10_115ACA_125").value '115ACA(1) - B3iieii
cgosIncome_cgInc_ltcg_prctg125_sec115E_b = Sheet13.Range("Inc_Frm_ltcg_10_115Eb_125").value '115E(b)_B8cii
cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = Sheet13.Range("Inc_Frm_ltcg_10_112_1ciii_125").value '112(1)c(iii) - B6ieii
cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = Sheet13.Range("Inc_frm_ltcg_10_115AC_125").value '115AC(1) - B6iieii
cgosIncome_cgInc_ltcg_prctg125_sec115E_a = Sheet13.Range("Inc_Frm_ltcg_10_115Ea_125").value '115E(a)_B8dii
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii_Proviso_125").value '115AD(iii)_Proviso - B7cii
cgosIncome_cgInc_ltcg_prctg125_sec112A = Sheet13.Range("Inc_Frm_ltcg_10_112A_125").value '112A - B4cii
cgosIncome_cgInc_ltcg_PTI125 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_125").value 'PTI - 11 a2ii
cgosIncome_cgInc_ltcg_PTI125_112 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_112_125").value 'PTI_112 - 11 a1ii
cgosIncome_cgInc_ltcg_PTI112_sec48 = Sheet13.Range("Inc_Frm_ltcg_125_112_Sec48").value '112 sec48 - B5ciii - NOT REQUIRED Inc_Frm_ltcg_125_112
 
cgosIncome_cgInc_ltcg_PTI125_112_B9 = Sheet13.Range("Inc_Frm_ltcg_125_112").value '112 - (B1gb,B9eii,B10ii,B5ciii)
cgosIncome_cgInc_ltcg_prctg125_secProviso = Sheet13.Range("Inc_Frm_ltcg_125_112proviso").value  '112 Proviso(B3iei)
 
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = Sheet13.Range("Inc_Frm_ltcg_125_115ADiii").value
 
'new stcg 30%
cgosIncome_cgInc_stcg_PTI30 = Sheet13.Range("Inc_Frm_stcg_30_PTI30").value
cgosIncome_cgInc_stcg_prctg30_sec115AD = Sheet13.Range("Inc_Frm_stcg_30_115AD").value
 
'end
 

cgosIncome_cgInc_ltcg_prctg125 = Sheet13.Range("IHLA.Evi1_CurrYearIncome").value

'LTCG 20%
cgosIncome_cgInc_ltcg_prctg20_sec112 = Sheet13.Range("Inc_frm_ltcg_20_112").value
cgosIncome_cgInc_ltcg_prctg20_sec11EA = Sheet13.Range("Inc_Frm_ltcg_20_115Ea").value
cgosIncome_cgInc_ltcg_PTI20 = Sheet13.Range("Inc_Frm_ltcg_20_PTI20").value
''Newly added by Bindu as per DE V4 on 15th March 25
cgosIncome_cgInc_ltcg_prctg20_secProviso = Sheet13.Range("Inc_Frm_ltcg_20_112proviso")


'end
 
'new Applicable PTI
cgosIncome_cgInc_stcg_Applicable = Sheet13.Range("Inc_Frm_stcg_Applicable").value
cgosIncome_cgInc_stcg_ApplicablePTI = Sheet13.Range("Inc_Frm_stcg_ApplicablePTI").value
'end
 
cgosIncome_cgInc_stcg_prctgAppRate = Sheet13.Range("IHLA.Eiv1_CurrYearIncome").value
cgosIncome_cgInc_ltcg_prctg20 = Sheet13.Range("IHLA.Evi1_CurrYearIncome").value
cgosIncome_cgInc_stcg_prctg15 = Sheet13.Range("IHLA.Eii1_CurrYearIncome").value
 
'---STCG 15% NEWLY ADDED BY RIYAZ
cgosIncome_cgInc_stcg_prctg20 = Sheet13.Range("IHLA.Eiii_CurrYearIncome20").value
 
'----LTCG 12.5% NEWLY ADDED BY RIYAZ
cgosIncome_cgInc_ltcg_125prctg125 = Sheet13.Range("IHLA.Eix_CurrYearIncome").value
 
'--20%
cgosIncome_cgInc_ltcg_prctg10 = Sheet13.Range("IHLA.Ev1_CurrYearIncome").value
cgosIncome_cgInc_stcg_prctg30 = Sheet13.Range("IHLA.Eiii1_CurrYearIncome").value


cgosIncome_cgInc_stcg_sumof20 = Sheet13.Range("STCG_sumof20").value
cgosIncome_cgInc_stcg_sumof30 = Sheet13.Range("STCG_sumof30").value
cgosIncome_cgInc_stcg_sumofAPP = Sheet13.Range("STCG_Appli_Rates").value

'Updated by Riyaz as range name changed
cgosIncome_cgInc_ltcg_sumof12_5 = Sheet13.Range("LTCG.BII_CapitalLoss").value
cgosIncome_cgInc_ltcg_sumof10 = Sheet13.Range("Sumof_10").value
 
populateCGTab

'Commented by Riyaz on 29/01/2026
STCG_30 = Sheet13.Range("STCG_30").value
STCG_20 = Sheet13.Range("STCG_20").value
'STCG_15 = Sheet13.Range("STCG_15").value
'LTCG_20 = Sheet13.Range("LTCG_20").value
LTCG_125 = Sheet13.Range("LTCG_12.5").value
'LTCG_10 = Sheet13.Range("LTCG_10").value
STCG_APP = Sheet13.Range("STCG_APP").value
STCG_DTAA = Sheet13.Range("STCG_DTAA").value
LTCG_DTAA = Sheet13.Range("LTCG_DTAA").value

If UCase(Sheet13.Range("CG_TableE_Checkbox").value) = UCase("True") Then

    Call setOffAgainst30(STCG_30, cgosIncome_cgInc_stcg_prctg30) 'STCG 30%
    Call setOffAgainstAr(STCG_APP, cgosIncome_cgInc_stcg_prctgAppRate) 'STCG APP RATE
    Call setOffAgainststcg20(STCG_20, cgosIncome_cgInc_stcg_prctg20) 'STCG 20%
'Commented by Riyaz on 29/01/2026
'    Call setOffAgainst20(LTCG_20, cgosIncome_cgInc_ltcg_prctg125) 'LTCG 20%
'    Call setOffAgainst15(STCG_15, cgosIncome_cgInc_stcg_prctg15) 'STCG 15%
    Call setOffAgainst125(LTCG_125, cgosIncome_cgInc_ltcg_125prctg125) 'LTCG 12.5%
'Commented by Riyaz on 29/01/2026
'    Call setOffAgainst10(LTCG_10, cgosIncome_cgInc_ltcg_prctg10) 'LTCG 10%
    'Call setOffstcgDTAA(Sheet13.Range("IHLA.Ei5_StclSetoffDTAARate").value)  'STCGD DTAA
    'Call setOffltcgDTAA(Sheet13.Range("IHLA.Ei8_LtclSetOffDTAARate").value)  'LTCG DTAA

Else
    setOffPctg30Loss (Sheet13.Range("IHLA.Ei3_StclSetoff30Per").value) 'STCG 30% - DONE
    setOffPctgArLoss (Sheet13.Range("IHLA.Ei4_StclSetoffAppRate").value) 'STCG AR - DONE
    setOffPctg20Loss_STCG (Sheet13.Range("IHLA.Ei2_StclSetoff20Per").value) '-STCG 20% NEW ADDED BY ME
'Commented by Riyaz on 29/01/2026
'    setOffPctg20Loss (Sheet13.Range("IHLA.Ei6_LtclSetOff20Per").value) 'LTCG20% - NA(STCG 20%,LTCG 12.5)
'    setOffPctg15Loss (Sheet13.Range("IHLA.Ei2_StclSetoff15Per").value) 'STCG 15% - DONE
    setOffPctg125Loss (Sheet13.Range("IHLA.Ei4_StclSetoffAppRate").value) 'LTCG 12.5%
'Commented by Riyaz on 29/01/2026
'    setOffPctg10Loss (Sheet13.Range("IHLA.Ei5_LtclSetOff10Per").value) 'LTCG 10%  - NA (STCG 20% NA,12.5)
    setOffstcgDTAA (Sheet13.Range("IHLA.Ei5_StclSetoffDTAARate").value) 'STCGD DTAA - DONE
    setOffltcgDTAA (Sheet13.Range("IHLA.Ei8_LtclSetOffDTAARate").value) 'LTCG DTAA - NA ONLY FOR STCG 20% - DONE
    
End If

 
 
 
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet31.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
 
    Sheet31.Range("temp20Sec112").value = cgosIncome_cgInc_ltcg_prctg20_sec112
    Sheet31.Range("temp20Sec11EA").value = cgosIncome_cgInc_ltcg_prctg20_sec11EA
    'Newly added by Bindu 15th March 25 as per DE V4
    Sheet31.Range("temp20SecPro").value = cgosIncome_cgInc_ltcg_prctg20_secProviso
'Commented by Riyaz on 29/01/2026
'    Sheet31.Range("temp15Sec111a").value = cgosIncome_cgInc_stcg_prctg15_sec111a
'    Sheet31.Range("temp15Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
    '---NEWLY ADDED BY BINDU STCG 20% ON 19TH FEB 25
    Sheet31.Range("temp20Sec111a").value = cgosIncome_cgInc_stcg_prctg20_sec111a '111a
    Sheet31.Range("temp20Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii '115AD

 
    'Sheet31.Range("temp10SecPro").value = cgosIncome_cgInc_ltcg_prctg10_secProviso '' cgosIncome_cgInc_ltcg_prctg10_secProviso NOt using  commented  by Bindu as this sec changed from 10 to 20

    Sheet31.Range("temp10Sec115ACA1").value = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
    Sheet31.Range("temp10Sec115Eb").value = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
    Sheet31.Range("temp10Sec1121c2").value = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
    Sheet31.Range("temp10Sec115AC1").value = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
    Sheet31.Range("temp10Sec115AD3").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
    Sheet31.Range("temp10Sec112A").value = cgosIncome_cgInc_ltcg_prctg10_sec112A
    Sheet31.Range("temp10Sec115AD3_Proviso").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
     Sheet31.Range("tempPTI10").value = cgosIncome_cgInc_ltcg_PTI10
     Sheet31.Range("tempPTI10_112").value = cgosIncome_cgInc_ltcg_PTI10_112
'Commented by Riyaz on 29/01/2026
'     Sheet31.Range("tempPTI15").value = cgosIncome_cgInc_stcg_PTI15
     'STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25 PTI
     Sheet31.Range("tempPTI20_STCG").value = cgosIncome_cgInc_stcg_PTI20
     Sheet31.Range("tempPTI20").value = cgosIncome_cgInc_ltcg_PTI20
     Sheet31.Range("tempPTI30").value = cgosIncome_cgInc_stcg_PTI30
     Sheet31.Range("temp115AD30").value = cgosIncome_cgInc_stcg_prctg30_sec115AD

     '---LYCG 12.5% NEWLY ADDED BY BINDU 23RD FEB 25
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
     Sheet31.Range("temp125Sec_112").value = cgosIncome_cgInc_ltcg_PTI125_112_B9 '11
     Sheet31.Range("temp125SecPro").value = cgosIncome_cgInc_ltcg_prctg125_secProviso '12
     Sheet31.Range("temp115AD125").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3



 
'----------------Lock Password-------------------START---
   Sheet31.Protect Password:=sPassword
 
'----------------Lock Password-------------------END-----
 
Application.ScreenUpdating = True
End Sub



Function populateCGTab()
Dim neg15 As Double
Dim stcgneg20 As Double
Dim neg10 As Double
Dim neg20 As Double
Dim neg125 As Double
Dim neg30 As Double
Dim negAR As Double
Dim SetOff As Double
Dim amtSetOff As Double

'Commented by Riyaz on 29/01/2026
' stcg = cgosIncome.cgInc.stcg
' prctg15 = cgosIncome.cgInc.stcg.prctg15
' prctg15Sum = cgosIncome_cgInc_stcg_prctg15_sec111a + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI15
' neg15 = 0
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg15_sec111a) < 0) Then
'        neg15 = neg15 + cgosIncome_cgInc_stcg_prctg15_sec111a
'       cgosIncome_cgInc_stcg_prctg15_sec111a = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii) < 0) Then
'        neg15 = neg15 + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
'       cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_PTI15) < 0) Then
'        neg15 = neg15 + cgosIncome_cgInc_stcg_PTI15
'       cgosIncome_cgInc_stcg_PTI15 = 0
'    End If
'
'    If (neg15 < 0) Then
'
'        SetOff = setOffAgainst115ad_1_b_ii(neg15, cgosIncome_cgInc_stcg_prctg15)
'        neg15 = neg15 + SetOff
'
'
'        SetOff = setOffAgainst15PTI(neg15, cgosIncome_cgInc_stcg_prctg15)
'        neg15 = neg15 + SetOff
'
'        SetOff = setOffAgainst111a(neg15, cgosIncome_cgInc_stcg_prctg15)
'        neg15 = neg15 + SetOff
'
'        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = neg15 + cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
'    End If
 
 '--Newly Added by Riyaz on 19/02/2025 for 20% STCG
 'cgosIncome_cgInc_stcg_sumof20 17TH March 25
 prcstcg20Sum = cgosIncome_cgInc_stcg_prctg20_sec111a + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii + cgosIncome_cgInc_stcg_PTI20 + cgosIncome_cgInc_stcg_sumof20
 stcgneg20 = 0
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg20_sec111a) < 0) Then
        stcgneg20 = stcgneg20 + cgosIncome_cgInc_stcg_prctg20_sec111a
       cgosIncome_cgInc_stcg_prctg20_sec111a = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii) < 0) Then
        stcgneg20 = stcgneg20 + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
       cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_PTI20) < 0) Then
        stcgneg20 = stcgneg20 + cgosIncome_cgInc_stcg_PTI20
       cgosIncome_cgInc_stcg_PTI20 = 0
    End If
    
    '17th March 25
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_sumof20) < 0) Then
        stcgneg20 = stcgneg20 + cgosIncome_cgInc_stcg_sumof20
       cgosIncome_cgInc_stcg_sumof20 = 0
    End If
    
    
    If (stcgneg20 < 0) Then
    
        SetOff = setOffAgainst115ad_1_b_ii20(stcgneg20, cgosIncome_cgInc_stcg_prctg20)
        stcgneg20 = stcgneg20 + SetOff
        
        
        SetOff = setOffAgainst15PTI20(stcgneg20, cgosIncome_cgInc_stcg_prctg20)
        stcgneg20 = stcgneg20 + SetOff
        
        SetOff = setOffAgainst111a20(stcgneg20, cgosIncome_cgInc_stcg_prctg20)
        stcgneg20 = stcgneg20 + SetOff
        
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = stcgneg20 + cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
    End If
'----End STCG 20%--
 
'Commented by Riyaz on 29/01/2026
' prctg10 = cgosIncome.cgInc.ltcg.prctg10
    
    ' prctg10Sum = cgosIncome_cgInc_ltcg_prctg10_secProviso + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112
    ' cgosIncome_cgInc_ltcg_prctg10_secProviso NOt using  commented  by Bindu as this sec changed from 10 to 20
    'cgosIncome_cgInc_ltcg_sumof10
'    prctg10Sum = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 + _
'                cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 + cgosIncome_cgInc_ltcg_prctg10_sec112A + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI10 + cgosIncome_cgInc_ltcg_PTI10_112 + cgosIncome_cgInc_ltcg_sumof10
'
'
'
' neg10 = 0
'    ' cgosIncome_cgInc_ltcg_prctg10_secProviso NOt using  commented  by Bindu as this sec changed from 10 to 20
''    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_secProviso) < 0) Then
''        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_secProviso
''       cgosIncome_cgInc_ltcg_prctg10_secProviso = 0
''    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115AC_1) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
'        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
'       cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115AD_3) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
'       cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115E_b) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115E_b
'        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
'       cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg10_sec112A) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112A
'       cgosIncome_cgInc_ltcg_prctg10_sec112A = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI10) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_PTI10
'       cgosIncome_cgInc_ltcg_PTI10 = 0
'    End If
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI10_112) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_PTI10_112
'       cgosIncome_cgInc_ltcg_PTI10_112 = 0
'    End If
'
'    '--18th march
'    'B(A)
'     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_sumof10) < 0) Then
'        neg10 = neg10 + cgosIncome_cgInc_ltcg_sumof10
'       cgosIncome_cgInc_ltcg_sumof10 = 0
'    End If
'
'
'    If (neg10 < 0) Then
'      SetOff = setOffAgainst112_1_c_2(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'        SetOff = setOffAgainst115AC_1(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'        SetOff = setOffAgainst115ACA_1(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'        SetOff = setOffAgainst115AD_3(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'
'        SetOff = setOffAgainst115AD_3_Proviso(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'
'        SetOff = setOffAgainst115E_b(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'         SetOff = setOffAgainstPTI10(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'    ' setOffAgainstsecProviso NOt using  commented  by Bindu as this sec changed from 10 to 20
'
''        SetOff = setOffAgainstsecProviso(neg10, cgosIncome_cgInc_ltcg_prctg10)
''        neg10 = neg10 + SetOff
'
'        SetOff = setOffAgainstPTI10_112(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'        SetOff = setOffAgainst112A(neg10, cgosIncome_cgInc_ltcg_prctg10)
'        neg10 = neg10 + SetOff
'
'
'
'        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = neg10 + cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'    End If

'Newly added by Riyaz on 21/02/2025 for LTCG 12.5%

'cgosIncome_cgInc_ltcg_sumof12_5
prctg125Sum = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 + _
                cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a + cgosIncome_cgInc_ltcg_prctg125_sec112A + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso + cgosIncome_cgInc_ltcg_PTI125 + cgosIncome_cgInc_ltcg_PTI125_112 + cgosIncome_cgInc_ltcg_PTI112_sec48 + cgosIncome_cgInc_ltcg_PTI125_112_B9 + cgosIncome_cgInc_ltcg_prctg125_secProviso + cgosIncome_cgInc_ltcg_sumof12_5 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
                
                    
 neg125 = 0
    
    '--1st
If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec112A) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec112A
        cgosIncome_cgInc_ltcg_prctg125_sec112A = 0
    End If
    
    '--2nd
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112
       cgosIncome_cgInc_ltcg_PTI125_112 = 0
    End If
    '--3rd
    '15th march
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_secProviso) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_secProviso
       cgosIncome_cgInc_ltcg_prctg125_secProviso = 0
    End If
    
    '--4th
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125
       cgosIncome_cgInc_ltcg_PTI125 = 0
    End If
    '--5th
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115E_b) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_b
        cgosIncome_cgInc_ltcg_prctg125_sec115E_b = 0
    End If
    '--6th
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
       cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = 0
    End If

'Added by Riyaz on 03/03/2026
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
       cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = 0
    End If
    
    
    '--7th
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
       cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = 0
    End If
    
    '--8th
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115AC_1) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
       cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = 0
    End If
    
    '-9th
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = 0
    End If
    '--10th
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI125_112_B9) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI125_112_B9
       cgosIncome_cgInc_ltcg_PTI125_112_B9 = 0
    End If
    '-Below 2 not required
     If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg125_sec115E_a) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115E_a
       cgosIncome_cgInc_ltcg_prctg125_sec115E_a = 0
    End If
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI112_sec48) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_PTI112_sec48
       cgosIncome_cgInc_ltcg_PTI112_sec48 = 0
    End If

    '18th March
    'B(A)
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_sumof12_5) < 0) Then
        neg125 = neg125 + cgosIncome_cgInc_ltcg_sumof12_5
       cgosIncome_cgInc_ltcg_sumof12_5 = 0
    End If
  
    If (neg125 < 0) Then
       '--Below section not required
        SetOff = setOffAgainst115E_a125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'Done
        neg125 = neg125 + SetOff
        SetOff = setOffAgainst112ASEC48(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
    '--10th
        SetOff = setOffAgainst112AB9(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
    '--9th
        SetOff = setOffAgainst112_1_c_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '--8th
        SetOff = setOffAgainst115AC_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '--7th
        SetOff = setOffAgainst115ACA_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
'Added by Riyaz on 03/03/2026
        SetOff = setOffAgainst115AD_125(neg125, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3) 'DONE
        neg125 = neg125 + SetOff
        
        '--6th
        SetOff = setOffAgainst115AD_3_Proviso_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'dONE
        neg125 = neg125 + SetOff
        '--5th
        SetOff = setOffAgainst115E_b125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '--4th
        SetOff = setOffAgainstPTI125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '--3rd
         '15th March
        SetOff = setOffAgainstsecProviso_125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '--2nd
        SetOff = setOffAgainstPTI10_1125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff
        '--1st
        SetOff = setOffAgainst112A125(neg125, cgosIncome_cgInc_ltcg_125prctg125) 'DONE
        neg125 = neg125 + SetOff

        
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = neg125 + cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
    End If
    
'Commented by Riyaz on 29/01/2026
'prctg20 = cgosIncome.cgInc.ltcg.prctg20
'cgosIncome_cgInc_ltcg_prctg20_secProviso newly added by Bindu 15th March 25 as per DE V4
'prctg20Sum = cgosIncome_cgInc_ltcg_prctg20_sec112 + cgosIncome_cgInc_ltcg_prctg20_sec11EA + cgosIncome_cgInc_ltcg_PTI20 + cgosIncome_cgInc_ltcg_prctg20_secProviso
'
'   neg20 = 0
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_sec112) < 0) Then
'        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec112
'        cgosIncome_cgInc_ltcg_prctg20_sec112 = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_sec11EA) < 0) Then
'        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec11EA
'        cgosIncome_cgInc_ltcg_prctg20_sec11EA = 0
'    End If
'
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_PTI20) < 0) Then
'        neg20 = neg20 + cgosIncome_cgInc_ltcg_PTI20
'        cgosIncome_cgInc_ltcg_PTI20 = 0
'    End If
'
'    ''cgosIncome_cgInc_ltcg_prctg20_secProviso newly added by Bindu 15th March 25 as per DE V4
'    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_ltcg_prctg20_secProviso) < 0) Then
'        neg20 = neg20 + cgosIncome_cgInc_ltcg_prctg20_secProviso
'        cgosIncome_cgInc_ltcg_prctg20_secProviso = 0
'    End If
'
'
'
'
'    If (neg20 < 0) Then
'       'Not using the below section
'       SetOff = setOffAgainst115Ea(neg20, cgosIncome_cgInc_ltcg_prctg20)
'          neg20 = neg20 + SetOff
'
'       '--PTI
'        SetOff = setOffAgainstPTI20(neg20, cgosIncome_cgInc_ltcg_prctg20)
'        neg20 = neg20 + SetOff
'
'         'cgosIncome_cgInc_ltcg_prctg20_secProviso newly added by Bindu 15th March 25 as per DE V4
'        SetOff = setOffAgainstsecProviso_20(neg20, cgosIncome_cgInc_ltcg_prctg20)
'        neg20 = neg20 + SetOff
'
'        SetOff = setOffAgainst112(neg20, cgosIncome_cgInc_ltcg_prctg20)
'        neg20 = neg20 + SetOff
'
'
'        cgosIncome_cgInc_ltcg_prctg20_sec11EA = neg20 + cgosIncome_cgInc_ltcg_prctg20_sec11EA
'    End If
    
    
    
'prctg30 = cgosIncome.cgInc.stcg.prctg30
'cgosIncome_cgInc_stcg_sumof30 17th March 25
prctg30Sum = cgosIncome_cgInc_stcg_prctg30_sec115AD + cgosIncome_cgInc_stcg_PTI30 + cgosIncome_cgInc_stcg_sumof30
    neg30 = 0
If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_prctg30_sec115AD) < 0) Then
        neg30 = neg30 + cgosIncome_cgInc_stcg_prctg30_sec115AD
        cgosIncome_cgInc_stcg_prctg30_sec115AD = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_PTI30) < 0) Then
        neg30 = neg30 + cgosIncome_cgInc_stcg_PTI30
        cgosIncome_cgInc_stcg_PTI30 = 0
    End If
    
    'March17th 25
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
'cgosIncome_cgInc_stcg_sumofAPP 17th March 25
prctgARSum = cgosIncome_cgInc_stcg_Applicable + cgosIncome_cgInc_stcg_ApplicablePTI + cgosIncome_cgInc_stcg_sumofAPP
    negAR = 0
If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_Applicable) < 0) Then
        negAR = negAR + cgosIncome_cgInc_stcg_Applicable
        cgosIncome_cgInc_stcg_Applicable = 0
    End If
    
    If (Application.WorksheetFunction.Min(0, cgosIncome_cgInc_stcg_ApplicablePTI) < 0) Then
        negAR = negAR + cgosIncome_cgInc_stcg_ApplicablePTI
        cgosIncome_cgInc_stcg_ApplicablePTI = 0
    End If
    
    '17th March 25
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
Sub setTblinfo_CG_Account_IFSC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet15.Range("CG.AccNumber").count
    Set rangecells = Sheet15.Range("CG.AccNumber").Cells
    
    For mIntCtr = 1 To mIntCells
    
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    
    end_account_ifsc = ccount
    rngname_account_ifsc = "CG.AccNumber;CG.IFSC;"
End Sub
Sub AddRows_Account_IFSC()
    Dim numberofrows As Long
    setTblinfo_CG_Account_IFSC
    SelectLastRow ("CG.AccNumber")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_account_ifsc)
End Sub

Function CYLA_BFLA_Editable() 'Added by Bindu


'Commented by Riyaz on 30/01/2026
'---STCG15%
'BFLA_STCG_15 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3").value
'
'cgosIncome_cgInc_stcg_prctg15_sec111a = Sheet13.Range("Inc_Frm_stcg_111A").value
'cgosIncome_cgInc_stcg_PTI15 = Sheet13.Range("Inc_Frm_stcg_PTI15").value
'cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = Sheet13.Range("Inc_Frm_stcg_115AD_1bii").value
'
'
'
'SI_STCG15_111A = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_15 - 0, cgosIncome_cgInc_stcg_prctg15_sec111a))
'SI_STCG15_PTI = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_15 - SI_STCG15_111A, cgosIncome_cgInc_stcg_PTI15))
'SI_STCG15_115ad_1_b_ii = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_15 - SI_STCG15_111A - SI_STCG15_PTI, cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii))



'---STCG 20%
BFLA_STCG_20 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3c").value


cgosIncome_cgInc_stcg_prctg20_sec111a = Sheet13.Range("Inc_Frm_stcg_111A_20").value
cgosIncome_cgInc_stcg_PTI20 = Sheet13.Range("Inc_Frm_stcg_PTI15_20").value
cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = Sheet13.Range("Inc_Frm_stcg_115AD_1bii_20").value


SI_STCG20_111A = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_20 - 0, cgosIncome_cgInc_stcg_prctg20_sec111a))
SI_STCG20_PTI = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_20 - SI_STCG20_111A, cgosIncome_cgInc_stcg_PTI20))
SI_STCG20_115ad_1_b_ii = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_20 - SI_STCG20_111A - SI_STCG20_PTI, cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii))



'STCG 30%

BFLA_STCG_30 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3a").value

cgosIncome_cgInc_stcg_prctg30_sec115AD = Sheet13.Range("Inc_Frm_stcg_30_115AD").value
cgosIncome_cgInc_stcg_PTI30 = Sheet13.Range("Inc_Frm_stcg_30_PTI30").value


SI_STCG30_115AD = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_30 - 0, cgosIncome_cgInc_stcg_prctg30_sec115AD))
SI_STCG30_PTI = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_STCG_30 - SI_STCG30_115AD, cgosIncome_cgInc_stcg_PTI30))

'Commented by Riyaz on 30/01/2026
'---LTCG 10%

'BFLA_LTCG_10 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4").value
'
'
'
'cgosIncome_cgInc_ltcg_prctg10_sec112A = Sheet13.Range("Inc_Frm_ltcg_10_112A").value '1st
'cgosIncome_cgInc_ltcg_PTI10_112 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_112").value '2nd
'cgosIncome_cgInc_ltcg_PTI10 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10").value ' 3rd
'cgosIncome_cgInc_ltcg_prctg10_sec115E_b = Sheet13.Range("Inc_Frm_ltcg_10_115Eb").value '4th
'cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii_Proviso").value '5th
'cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii").value '6th
'cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = Sheet13.Range("Inc_Frm_ltcg_10_115ACA").value '7th
'cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = Sheet13.Range("Inc_frm_ltcg_10_115AC").value '8th
'cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = Sheet13.Range("Inc_Frm_ltcg_10_112_1ciii").value '9th
'
'
'SI_LTCG10_112A = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - 0, cgosIncome_cgInc_ltcg_prctg10_sec112A))
'SI_LTCG10_PTI112 = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A, cgosIncome_cgInc_ltcg_PTI10_112))
'SI_LTCG10_PTI = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112, cgosIncome_cgInc_ltcg_PTI10))
'SI_LTCG10_115EB = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112 - SI_LTCG10_PTI, cgosIncome_cgInc_ltcg_prctg10_sec115E_b))
'SI_LTCG10_115AD3_Pro = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112 - SI_LTCG10_PTI - SI_LTCG10_115EB, cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso))
'SI_LTCG10_115AD3 = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112 - SI_LTCG10_PTI - SI_LTCG10_115EB - SI_LTCG10_115AD3_Pro, cgosIncome_cgInc_ltcg_prctg10_sec115AD_3))
'SI_LTCG10_115ACA = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112 - SI_LTCG10_PTI - SI_LTCG10_115EB - SI_LTCG10_115AD3_Pro - SI_LTCG10_115AD3, cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1))
'SI_LTCG10_115AC = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112 - SI_LTCG10_PTI - SI_LTCG10_115EB - SI_LTCG10_115AD3_Pro - SI_LTCG10_115AD3 - SI_LTCG10_115ACA, cgosIncome_cgInc_ltcg_prctg10_sec115AC_1))
'SI_LTCG10_1121C = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_10 - SI_LTCG10_112A - SI_LTCG10_PTI112 - SI_LTCG10_PTI - SI_LTCG10_115EB - SI_LTCG10_115AD3_Pro - SI_LTCG10_115AD3 - SI_LTCG10_115ACA - SI_LTCG10_115AC, cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2))



'---LTCG 12.5%
BFLA_LTCG_125 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4b").value


cgosIncome_cgInc_ltcg_prctg125_sec112A = Sheet13.Range("Inc_Frm_ltcg_10_112A_125").value '112A - B4cii '1ST
cgosIncome_cgInc_ltcg_PTI125_112 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_112_125").value 'PTI_112 - 11 a1ii '2ND
cgosIncome_cgInc_ltcg_prctg125_secProviso = Sheet13.Range("Inc_Frm_ltcg_125_112proviso").value  '112 Proviso(B3iei) '3RD
cgosIncome_cgInc_ltcg_PTI125 = Sheet13.Range("Inc_Frm_ltcg_10_PTI10_125").value 'PTI - 11 a2ii '4TH
cgosIncome_cgInc_ltcg_prctg125_sec115E_b = Sheet13.Range("Inc_Frm_ltcg_10_115Eb_125").value '115E(b)_B8cii '5TH
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = Sheet13.Range("Inc_Frm_ltcg_10_115ADiii_Proviso_125").value '115AD(iii)_Proviso - B7cii '6TH
cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = Sheet13.Range("Inc_Frm_ltcg_10_115ACA_125").value '115ACA(1) - B3iieii '7TH
cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = Sheet13.Range("Inc_frm_ltcg_10_115AC_125").value '115AC(1) - B6iieii '8TH
cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = Sheet13.Range("Inc_Frm_ltcg_10_112_1ciii_125").value '112(1)c(iii) - B6ieii '9TH
cgosIncome_cgInc_ltcg_PTI125_112_B9 = Sheet13.Range("Inc_Frm_ltcg_125_112").value '112 - (B1gb,B9eii,B10ii,B5ciii) '10TH
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = Sheet13.Range("Inc_Frm_ltcg_125_115ADiii").value


SI_LTCG125_112A = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - 0, cgosIncome_cgInc_ltcg_prctg125_sec112A))
SI_LTCG125_PTI112 = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A, cgosIncome_cgInc_ltcg_PTI125_112))
SI_LTCG125_112_Pro = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112, cgosIncome_cgInc_ltcg_prctg125_secProviso))
SI_LTCG125_PTI = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro, cgosIncome_cgInc_ltcg_PTI125))
SI_LTCG125_115EB = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI, cgosIncome_cgInc_ltcg_prctg125_sec115E_b))
SI_LTCG125_115AD3_Pro = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI - SI_LTCG125_115EB, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso))
'Added by Riyaz on 03/03/2026
SI_LTCG125_115AD_125 = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI - SI_LTCG125_115EB - SI_LTCG125_115AD3_Pro, cgosIncome_cgInc_ltcg_prctg125_sec115AD_3))

SI_LTCG125_115ACA = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI - SI_LTCG125_115EB - SI_LTCG125_115AD3_Pro - SI_LTCG125_115AD_125, cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1))
SI_LTCG125_115AC = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI - SI_LTCG125_115EB - SI_LTCG125_115AD3_Pro - SI_LTCG125_115AD_125 - SI_LTCG125_115ACA, cgosIncome_cgInc_ltcg_prctg125_sec115AC_1))
SI_LTCG125_1121C = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI - SI_LTCG125_115EB - SI_LTCG125_115AD3_Pro - SI_LTCG125_115AD_125 - SI_LTCG125_115ACA - SI_LTCG125_115AC, cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2))
SI_LTCG125_112 = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_125 - SI_LTCG125_112A - SI_LTCG125_PTI112 - SI_LTCG125_112_Pro - SI_LTCG125_PTI - SI_LTCG125_115EB - SI_LTCG125_115AD3_Pro - SI_LTCG125_115AD_125 - SI_LTCG125_115ACA - SI_LTCG125_115AC - SI_LTCG125_1121C, cgosIncome_cgInc_ltcg_PTI125_112_B9))


'Commented by Riyaz on 30/01/2026
'---LTCG 20%

'BFLA_LTCG_20 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4a").value
'
'
'
'cgosIncome_cgInc_ltcg_prctg20_sec112 = Sheet13.Range("Inc_frm_ltcg_20_112").value
'cgosIncome_cgInc_ltcg_PTI20 = Sheet13.Range("Inc_Frm_ltcg_20_PTI20").value
'cgosIncome_cgInc_ltcg_prctg20_secProviso = Sheet13.Range("Inc_Frm_ltcg_20_112proviso")
'
'
'
'SI_LTCG20_112 = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_20 - 0, cgosIncome_cgInc_ltcg_prctg20_sec112))
'SI_LTCG20_PTI = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_20 - SI_LTCG20_112, cgosIncome_cgInc_ltcg_PTI20))
'SI_LTCG20_112_PRO = WorksheetFunction.Max(0, WorksheetFunction.Min(BFLA_LTCG_20 - SI_LTCG20_112 - SI_LTCG20_PTI, cgosIncome_cgInc_ltcg_prctg20_secProviso))




'--Temp Sheet
Sheet31.Unprotect Password:=EfilingCommon.getmsgstate
Application.EnableEvents = False

'Commented by Riyaz on 30/01/2026
'---STCG 15%
'Sheet31.Range("BFLAtemp15Sec111a").value = SI_STCG15_111A
'Sheet31.Range("BFLAtempPTI15").value = SI_STCG15_PTI
'Sheet31.Range("BFLAtemp15Sec115ad1bii").value = SI_STCG15_115ad_1_b_ii


'--STCG 20%
Sheet31.Range("BFLAtemp20Sec111a").value = SI_STCG20_111A
Sheet31.Range("BFLAtempPTI20_STCG").value = SI_STCG20_PTI
Sheet31.Range("BFLAtemp20Sec115ad1bii").value = SI_STCG20_115ad_1_b_ii

'--STCG 30%
Sheet31.Range("BFLAtemp115AD30").value = SI_STCG30_115AD
Sheet31.Range("BFLAtempPTI30").value = SI_STCG30_PTI


'Commented by Riyaz on 30/01/2026
'--LTCG 10%
'
'Sheet31.Range("BFLAtemp10Sec112A").value = SI_LTCG10_112A
'Sheet31.Range("BFLAtempPTI10_112").value = SI_LTCG10_PTI112
'Sheet31.Range("BFLAtempPTI10").value = SI_LTCG10_PTI
'Sheet31.Range("BFLAtemp10Sec115Eb").value = SI_LTCG10_115EB
'Sheet31.Range("BFLAtemp10Sec115AD3_Proviso").value = SI_LTCG10_115AD3_Pro
'Sheet31.Range("BFLAtemp10Sec115AD3").value = SI_LTCG10_115AD3
'Sheet31.Range("BFLAtemp10Sec115ACA1").value = SI_LTCG10_115ACA
'Sheet31.Range("BFLAtemp10Sec115AC1").value = SI_LTCG10_115AC
'Sheet31.Range("BFLAtemp10Sec1121c2").value = SI_LTCG10_1121C

'--LTCG 12.5%
Sheet31.Range("BFLAtemp125Sec_112A").value = SI_LTCG125_112A
Sheet31.Range("BFLAtemp125Sec_112_PTI").value = SI_LTCG125_PTI112
Sheet31.Range("BFLAtemp125SecPro").value = SI_LTCG125_112_Pro
Sheet31.Range("BFLAtemp125Sec_PTI").value = SI_LTCG125_PTI
Sheet31.Range("BFLAtemp125Sec_115EB").value = SI_LTCG125_115EB
Sheet31.Range("BFLAtemp125Sec_115AD_iii").value = SI_LTCG125_115AD3_Pro
Sheet31.Range("BFLAtemp125Sec_115ACA").value = SI_LTCG125_115ACA
Sheet31.Range("BFLAtemp125Sec_115AC").value = SI_LTCG125_115AC
Sheet31.Range("BFLAtemp125Sec112_c").value = SI_LTCG125_1121C
Sheet31.Range("BFLAtemp125Sec_112").value = SI_LTCG125_112

Sheet31.Range("BFLAtemp115AD125").value = SI_LTCG125_115AD_125
'Commented by Riyaz on 30/01/2026
'---LTCG 20%
'Sheet31.Range("BFLAtemp20Sec112").value = SI_LTCG20_112
'Sheet31.Range("BFLAtempPTI20").value = SI_LTCG20_PTI
'Sheet31.Range("BFLAtemp20SecPro").value = SI_LTCG20_112_PRO



Sheet31.Protect Password:=EfilingCommon.getmsgstate
Application.EnableEvents = True

End Function


