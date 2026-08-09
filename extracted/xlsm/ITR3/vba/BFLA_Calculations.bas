Attribute VB_Name = "BFLA_Calculations"
Option Explicit

Function calcBFLA()

Application.EnableEvents = False
Application.ScreenUpdating = False
   
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet16.Unprotect Password:=sPassword  'CYLA-BFLA
   'Sheet16.Unprotect Password:=sPassword   'CG
   'Sheet9.Unprotect Password:=sPassword    'CFL
'----------------Unlock Password-------------------END-----

Dim stcgBreakUp1 As Double
Dim stcgBreakUp2 As Double
Dim stcgBreakUp3 As Double
Dim stcgBreakUp4 As Double

Dim ltcgBreakUp1 As Double
Dim ltcgBreakUp2 As Double
Dim ltcgBreakUp3 As Double
Dim ltcgBreakUp4 As Double
Dim ltcgBreakUp5 As Double
Dim ltcgBreakUp6 As Double
Dim ltcgBreakUp7 As Double
Dim ltcgBreakUp8 As Double


Dim tempstcgBreakUp1 As Double
Dim tempstcgBreakUp2 As Double
Dim tempstcgBreakUp3 As Double
Dim tempstcgBreakUp4 As Double
Dim templtcgBreakUp1 As Double
Dim templtcgBreakUp2 As Double
Dim templtcgBreakUp3 As Double
Dim templtcgBreakUp4 As Double
Dim templtcgBreakUp5 As Double
Dim templtcgBreakUp6 As Double
Dim templtcgBreakUp7 As Double
Dim templtcgBreakUp8 As Double

Dim tempLtcgCFL As Double
Dim tempstcgCFL As Double

stcgBreakUp1 = Sheet8.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value  'cgosIncome.cgInc.stcg.prctg30
stcgBreakUp2 = Sheet8.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value 'cgosIncome.cgInc.stcg.prctgAr
stcgBreakUp3 = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
stcgBreakUp4 = cgosIncome_cgInc_stcg_prctg15_sec111a
        
ltcgBreakUp2 = cgosIncome_cgInc_ltcg_prctg20_sec11EA
ltcgBreakUp1 = cgosIncome_cgInc_ltcg_prctg20_sec112
ltcgBreakUp4 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
ltcgBreakUp5 = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
ltcgBreakUp6 = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
ltcgBreakUp7 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
ltcgBreakUp8 = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
ltcgBreakUp3 = cgosIncome_cgInc_ltcg_prctg10_secProviso
 
tempstcgBreakUp1 = stcgBreakUp1
tempstcgBreakUp2 = stcgBreakUp2
tempstcgBreakUp3 = stcgBreakUp3
tempstcgBreakUp4 = stcgBreakUp4
        
templtcgBreakUp1 = ltcgBreakUp1
templtcgBreakUp2 = ltcgBreakUp2
templtcgBreakUp3 = ltcgBreakUp3
templtcgBreakUp4 = ltcgBreakUp4
templtcgBreakUp5 = ltcgBreakUp5
templtcgBreakUp6 = ltcgBreakUp6
templtcgBreakUp7 = ltcgBreakUp7
templtcgBreakUp8 = ltcgBreakUp8
 
 tempLtcgCFL = Sheet9.Range("totofbfloss.LTCGLossCF8").value
 tempstcgCFL = Sheet9.Range("totofbfloss.STCGLossCF8").value

'stage-1 setting off LTCG_CFL
'setting with templtcgBreakUp1
        
If (tempLtcgCFL > templtcgBreakUp2) Then
      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp2
      templtcgBreakUp2 = 0
Else
    templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - tempLtcgCFL)
    tempLtcgCFL = 0
End If
    
If (tempLtcgCFL > templtcgBreakUp1) Then
       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp1
       templtcgBreakUp1 = 0
Else
       templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp4) Then
     tempLtcgCFL = tempLtcgCFL - templtcgBreakUp4
     templtcgBreakUp4 = 0
Else
     templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - tempLtcgCFL)
     tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp5) Then
      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp5
     templtcgBreakUp5 = 0
Else
      templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - tempLtcgCFL)
      tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp6) Then
       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp6
       templtcgBreakUp6 = 0
Else
     templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - tempLtcgCFL)
     tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp7) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7
            templtcgBreakUp7 = 0
Else
      templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp8) Then
         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp8
         templtcgBreakUp8 = 0
Else
         templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - tempLtcgCFL)
         tempLtcgCFL = 0
End If
        
 If (tempLtcgCFL > templtcgBreakUp3) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp3
        templtcgBreakUp3 = 0
Else
        templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - tempLtcgCFL)
        tempLtcgCFL = 0
End If


Sheet9.Range("adjtotloss.LTCGLossCF9").value = Application.WorksheetFunction.Max(0, ltcgBreakUp1 - templtcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp2 - templtcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp3 - templtcgBreakUp3) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp4 - templtcgBreakUp4) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp5 - templtcgBreakUp5) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp6 - templtcgBreakUp6) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7 - templtcgBreakUp7) + _
                                             Application.WorksheetFunction.Max(0, ltcgBreakUp8 - templtcgBreakUp8)


Dim storeSTCG1, storeSTCG2, storeSTCG3, storeSTCG4 As Double
Dim storeLTCG1, storeLTCG2, storeLTCG3, storeLTCG4, storeLTCG5, storeLTCG6, storeLTCG7, storeLTCG8 As Double

storeSTCG1 = tempstcgBreakUp1
storeSTCG2 = tempstcgBreakUp2
storeSTCG3 = tempstcgBreakUp3
storeSTCG4 = tempstcgBreakUp4
storeLTCG1 = templtcgBreakUp1
storeLTCG2 = templtcgBreakUp2
storeLTCG3 = templtcgBreakUp3
storeLTCG4 = templtcgBreakUp4
storeLTCG5 = templtcgBreakUp5
storeLTCG6 = templtcgBreakUp6
storeLTCG7 = templtcgBreakUp7
storeLTCG8 = templtcgBreakUp8
        


'SCTCG SetOff

If (tempstcgCFL > tempstcgBreakUp1) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp1
            tempstcgBreakUp1 = 0
Else
            tempstcgBreakUp1 = Application.WorksheetFunction.Max(0, tempstcgBreakUp1 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > tempstcgBreakUp2) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp2
            tempstcgBreakUp2 = 0
Else
            tempstcgBreakUp2 = Application.WorksheetFunction.Max(0, tempstcgBreakUp2 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > templtcgBreakUp2) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp2
            templtcgBreakUp2 = 0
Else
            templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > templtcgBreakUp1) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp1
            templtcgBreakUp1 = 0
Else
            templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > tempstcgBreakUp3) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp3
            tempstcgBreakUp3 = 0
Else
            tempstcgBreakUp3 = Application.WorksheetFunction.Max(0, tempstcgBreakUp3 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > tempstcgBreakUp4) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp4
            tempstcgBreakUp4 = 0
Else
            tempstcgBreakUp4 = Application.WorksheetFunction.Max(0, tempstcgBreakUp4 - tempstcgCFL)
            tempstcgCFL = 0
End If

        
 If (tempstcgCFL > templtcgBreakUp4) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp4
            templtcgBreakUp4 = 0
Else
            templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - tempstcgCFL)
            tempstcgCFL = 0
End If
If (tempstcgCFL > templtcgBreakUp5) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp5
            templtcgBreakUp5 = 0
Else
            templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > templtcgBreakUp6) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp6
            templtcgBreakUp6 = 0
Else
            templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > templtcgBreakUp7) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp7
            templtcgBreakUp7 = 0
Else
            templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > templtcgBreakUp8) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp8
            templtcgBreakUp8 = 0
Else
            templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
If (tempstcgCFL > templtcgBreakUp3) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp3
            templtcgBreakUp3 = 0
Else
            templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
 If (tempLtcgCFL > templtcgBreakUp2) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp2
            templtcgBreakUp2 = 0
Else
            templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
 If (tempLtcgCFL > templtcgBreakUp1) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp1
            templtcgBreakUp1 = 0
Else
            templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp4) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp4
            templtcgBreakUp4 = 0
Else
            templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp5) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp5
            templtcgBreakUp5 = 0
Else
            templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp6) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp6
            templtcgBreakUp6 = 0
Else
            templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - tempLtcgCFL)
            tempLtcgCFL = 0
End If

If (tempLtcgCFL > templtcgBreakUp7) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7
            templtcgBreakUp7 = 0
Else
            templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp8) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp8
            templtcgBreakUp8 = 0
Else
            templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - tempLtcgCFL)
            tempLtcgCFL = 0
End If
        
If (tempLtcgCFL > templtcgBreakUp3) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp3
            templtcgBreakUp3 = 0
Else
            templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - tempLtcgCFL)
            tempLtcgCFL = 0
End If

Sheet8.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value = Application.WorksheetFunction.Max(0, stcgBreakUp1 - tempstcgBreakUp1)

Sheet8.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value = Application.WorksheetFunction.Max(0, stcgBreakUp2 - tempstcgBreakUp2)

Sheet8.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value = Application.WorksheetFunction.Max(0, stcgBreakUp3 - tempstcgBreakUp3) + _
                                                            Application.WorksheetFunction.Max(0, stcgBreakUp4 - tempstcgBreakUp4)
                                                            
Sheet8.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value = Application.WorksheetFunction.Max(0, ltcgBreakUp1 - templtcgBreakUp1) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp2 - templtcgBreakUp2)
                                                            
Sheet8.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value = Application.WorksheetFunction.Max(0, ltcgBreakUp3 - templtcgBreakUp3) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp4 - templtcgBreakUp4) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp5 - templtcgBreakUp5) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp6 - templtcgBreakUp6) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp7 - templtcgBreakUp7) + _
                                                           Application.WorksheetFunction.Max(0, ltcgBreakUp8 - templtcgBreakUp8)
                                                            
Sheet9.Range("adjtotloss.STCGLossCF9").value = Application.WorksheetFunction.Max(0, storeSTCG1 - tempstcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG2 - tempstcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG3 - tempstcgBreakUp3) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG4 - tempstcgBreakUp4) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG1 - templtcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG2 - templtcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG3 - templtcgBreakUp3) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG4 - templtcgBreakUp4) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG5 - templtcgBreakUp5) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG6 - templtcgBreakUp6) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG7 - templtcgBreakUp7) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG8 - templtcgBreakUp8)
        
        cgosIncome_cgInc_stcg_prctg30 = tempstcgBreakUp1
        cgosIncome_cgInc_stcg_prctgAppRate = tempstcgBreakUp2
        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = tempstcgBreakUp3
        cgosIncome_cgInc_stcg_prctg15_sec111a = tempstcgBreakUp4
        cgosIncome_cgInc_ltcg_prctg20_sec112 = templtcgBreakUp1
        cgosIncome_cgInc_ltcg_prctg20_sec11EA = templtcgBreakUp2
        cgosIncome_cgInc_ltcg_prctg10_secProviso = templtcgBreakUp3
        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = templtcgBreakUp4
        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = templtcgBreakUp5
        cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = templtcgBreakUp6
        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = templtcgBreakUp7
        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = templtcgBreakUp8

 'storing diff section remaining value
    Sheet16.Range("BFLAtemp20Sec112").value = cgosIncome_cgInc_ltcg_prctg20_sec112
    Sheet16.Range("BFLAtemp20Sec11EA").value = cgosIncome_cgInc_ltcg_prctg20_sec11EA

    Sheet16.Range("BFLAtemp15Sec111a").value = cgosIncome_cgInc_stcg_prctg15_sec111a
    Sheet16.Range("BFLAtemp15Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii

    Sheet16.Range("BFLAtemp10SecPro").value = cgosIncome_cgInc_ltcg_prctg10_secProviso
    Sheet16.Range("BFLAtemp10Sec115ACA1").value = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
    Sheet16.Range("BFLAtemp10Sec115Eb").value = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
    Sheet16.Range("BFLAtemp10Sec1121c2").value = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
    Sheet16.Range("BFLAtemp10Sec115AC1").value = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
        Sheet16.Range("BFLAtemp10Sec115AD3").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3

  '----------------Lock Password-------------------START---
  Sheet8.Protect Password:=sPassword
  Sheet16.Protect Password:=sPassword
  Sheet9.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

   Application.ScreenUpdating = True
    Application.EnableEvents = True
End Function



