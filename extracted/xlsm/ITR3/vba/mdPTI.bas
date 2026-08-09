Attribute VB_Name = "mdPTI"
Option Explicit

Dim sPassword As String

  Dim end_PTI_1, end_PTI_2, end_PTI_3, end_PTI_4, end_PTI_5, end_PTI_6, end_PTI_7, end_PTI_8, end_PTI_9, end_PTI_10, end_PTI_11, end_PTI_12, end_PTI_13, end_PTI_14, end_PTI_15, end_PTI_16, end_PTI_17, end_PTI_18, end_PTI_19, end_PTI_20, end_PTI_21, end_PTI_22, end_PTI_23, end_PTI_24, end_PTI_25, end_PTI_26, end_PTI_27, end_PTI_28, end_PTI_29, end_PTI_30, end_PTI_31 As Variant
Dim end_PTI3 As Variant

'PTI
Public NameOfBusiness_PTI As Variant
Public PANOfBusiness_PTI As Variant
Public Investment115UAOfBusiness_PTI As Variant

Public PTI_HPIncomeAmt As Variant
Public PTI_HPTDSAmt As Variant

Public PTI_CG_STCGIncomeAmt As Variant
Public PTI_CG_STCGTDSAmt As Variant
Public PTI_CG_LTCGIncomeAmt As Variant
Public PTI_CG_LTCGTDSAmt As Variant

Public PTI_OSIncomeAmt As Variant
Public PTI_OSTDSAmt As Variant

Public PTI_Usa As Variant
Public PTI_UsIncomeAmta As Variant
Public PTI_UsTDSAmta As Variant

Public PTI_Us_iv_TdsAmta As Variant

Public PTI_Usb As Variant
Public PTI_UsIncomeAmtb As Variant
Public PTI_UsTDSAmtb As Variant

Public PTI_Usc As Variant
Public PTI_UsIncomeAmtc As Variant
Public PTI_UsTDSAmtc As Variant

Public rngname_PTI As Variant
Public end_PTI, end_PTI1 As Variant

Public rngname_PTI2 As Variant
Public end_PTI2 As Variant
Dim msgValidatesheetPTI As String

Public frmsize_Ptirptfrm As Variant
Public cntrRng_Ptirptfrm As Variant
Public frmRngname_Ptirptfrm As Variant

Sub ValidatePTI()
    ValidateSheetSchPTI
    fmsgboxStatus "Sheet PTI is ok" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheetSchPTI()
    If Not ValidatesheetPTI Then
        Sheet41.Activate
        fmsgboxsmall msgValidatesheetPTI ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub

Sub AddRows_PTI()
    setTblinfo_PTI
    Call addPtiblock(rngname_PTI, frmRngname_Ptirptfrm, cntrRng_Ptirptfrm, frmsize_Ptirptfrm)
End Sub



Function MyTotalPTI() As Long
    'dpk
    Dim te As Long
    MyTotalPTI = 0
    
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
       MyTotalPTI = MyTotalPTI + Sheet41.Range("PTI_Us_iv_NetAmta" & te).value
    Next
    End Function

Sub setTblinfo_PTI()
    Dim Investment_UA_UB As Variant
    Dim ccount As Long
    Dim te As Long
    
    ccount = 0
    frmsize_Ptirptfrm = Sheet41.Range("Ptirptfrm.size").value
    
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        Investment_UA_UB = Sheet41.Range("PTI_Investment_UA_UB" & te).value

        If Not isdropdownblank(Investment_UA_UB) Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    end_PTI = ccount
    cntrRng_Ptirptfrm = "cntr.Ptirptfrm"
    frmRngname_Ptirptfrm = "Ptirptfrm"
    
rngname_PTI = "PTI_Investment_UA_UB1;PTI_NameOfBusiness1;PTI_PAN1;PTI_HPIncomeAmt1;PTI_HP_ShareAmt1;PTI_HP_NetAmt1;PTI_HPTDSAmt1;PTI_CG_STCGIncomeAmt1;PTI_CG_STCG_ShareAmt1;PTI_CG_STCG_NetAmt1;PTI_CG_STCGTDSAmt1;PTI_CG_STCG_SectionA_IncomeAmt1;PTI_CG_STCG_SectionA_ShareAmt1;PTI_CG_STCG_SectionA_NetAmt1;PTI_CG_STCG_SectionA_TdsAmt1;PTI_CG_STCG_Others_IncomeAmt1;PTI_CG_STCG_Others_ShareAmt1;PTI_CG_STCG_Others_NetAmt1;PTI_CG_STCG_Others_TdsAmt1;PTI_CG_LTCGIncomeAmt1;PTI_CG_LTCG_ShareAmt1;PTI_CG_LTCG_NetAmt1;PTI_CG_LTCGTDSAmt1;PTI_CG_LTCG_SectionA_IncomeAmt1;PTI_CG_LTCG_SectionA_ShareAmt1;PTI_CG_LTCG_SectionA_NetAmt1;PTI_CG_LTCG_SectionA_TdsAmt1;PTI_CG_LTCG_Others_IncomeAmt1;PTI_CG_LTCG_Others_ShareAmt1;PTI_CG_LTCG_Others_NetAmt1;PTI_CG_LTCG_Others_TdsAmt1;PTI_OSIncomeAmt1;PTI_OS_NetAmt1;PTI_OSTDSAmt1;PTI_OS_Dividend_IncomeAmt1;PTI_OS_Dividend_NetAmt1;PTI_OS_Dividend_TdsAmt1;PTI_OS_Others_IncomeAmt1;PTI_OS_Others_NetAmt1;PTI_OS_Others_TdsAmt1;PTI_Us_iv_IncomeAmta1;PTI_Us_iv_NetAmta1;PTI_Us_iv_TdsAmta1;" & _
 "PTI_Usa1;PTI_UsIncomeAmta1;PTI_Us_NetAmta1;PTI_UsTDSAmta1;PTI_Usb1;PTI_UsIncomeAmtb1;PTI_Us_NetAmtb1;PTI_UsTDSAmtb1;PTI_Usc1;PTI_UsIncomeAmtc1;PTI_Us_NetAmtc1;PTI_UsTDSAmtc1;"
End Sub

Sub setTblinfo_PTI2()
    Dim Investment_UA_UB As Variant
    Dim ccount As Long
    Dim te As Long
    
    ccount = 0
    frmsize_Ptirptfrm = Sheet41.Range("Ptirptfrm.size").value
    
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
       Investment_UA_UB = Sheet41.Range("PTI_Investment_UA_UB" & te).value
        If isdropdownblank(Investment_UA_UB) Then
            Investment_UA_UB = ""
        End If
        
        If Not Investment_UA_UB = "" Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    
    end_PTI2 = ccount
    cntrRng_Ptirptfrm = "cntr.Ptirptfrm"
    frmRngname_Ptirptfrm = "Ptirptfrm"
    
rngname_PTI = "PTI_Investment_UA_UB1;PTI_NameOfBusiness1;PTI_PAN1;PTI_HPIncomeAmt1;PTI_HP_ShareAmt1;PTI_HP_NetAmt1;PTI_HPTDSAmt1;PTI_CG_STCGIncomeAmt1;PTI_CG_STCG_ShareAmt1;PTI_CG_STCG_NetAmt1;PTI_CG_STCGTDSAmt1;PTI_CG_STCG_SectionA_IncomeAmt1;PTI_CG_STCG_SectionA_ShareAmt1;PTI_CG_STCG_SectionA_NetAmt1;PTI_CG_STCG_SectionA_TdsAmt1;PTI_CG_STCG_Others_IncomeAmt1;PTI_CG_STCG_Others_ShareAmt1;PTI_CG_STCG_Others_NetAmt1;PTI_CG_STCG_Others_TdsAmt1;PTI_CG_LTCGIncomeAmt1;PTI_CG_LTCG_ShareAmt1;PTI_CG_LTCG_NetAmt1;PTI_CG_LTCGTDSAmt1;PTI_CG_LTCG_SectionA_IncomeAmt1;PTI_CG_LTCG_SectionA_ShareAmt1;PTI_CG_LTCG_SectionA_NetAmt1;PTI_CG_LTCG_SectionA_TdsAmt1;PTI_CG_LTCG_Others_IncomeAmt1;PTI_CG_LTCG_Others_ShareAmt1;PTI_CG_LTCG_Others_NetAmt1;PTI_CG_LTCG_Others_TdsAmt1;PTI_OSIncomeAmt1;PTI_OS_NetAmt1;PTI_OSTDSAmt1;PTI_OS_Dividend_IncomeAmt1;PTI_OS_Dividend_NetAmt1;PTI_OS_Dividend_TdsAmt1;PTI_OS_Others_IncomeAmt1;PTI_OS_Others_NetAmt1;PTI_OS_Others_TdsAmt1;PTI_Us_iv_IncomeAmta1;PTI_Us_iv_NetAmta1;PTI_Us_iv_TdsAmta1;" & _
 "PTI_Usa1;PTI_UsIncomeAmta1;PTI_Us_NetAmta1;PTI_UsTDSAmta1;PTI_Usb1;PTI_UsIncomeAmtb1;PTI_Us_NetAmtb1;PTI_UsTDSAmtb1;PTI_Usc1;PTI_UsIncomeAmtc1;PTI_Us_NetAmtc1;PTI_UsTDSAmtc1;"
End Sub

Sub setTblinfo_PTI1()
    Dim PAN_PTI As Variant
    Dim ccount As Long
    Dim te As Long
    
    ccount = 0
    frmsize_Ptirptfrm = Sheet41.Range("Ptirptfrm.size").value
    
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
       PAN_PTI = Sheet41.Range("PTI_PAN" & te).value
        If isdropdownblank(PAN_PTI) Then
            PAN_PTI = ""
        End If
        
        If Not PAN_PTI = "" Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    
    end_PTI1 = ccount
    cntrRng_Ptirptfrm = "cntr.Ptirptfrm"
    frmRngname_Ptirptfrm = "Ptirptfrm"
    
rngname_PTI = "PTI_Investment_UA_UB1;PTI_NameOfBusiness1;PTI_PAN1;PTI_HPIncomeAmt1;PTI_HP_ShareAmt1;PTI_HP_NetAmt1;PTI_HPTDSAmt1;PTI_CG_STCGIncomeAmt1;PTI_CG_STCG_ShareAmt1;PTI_CG_STCG_NetAmt1;PTI_CG_STCGTDSAmt1;PTI_CG_STCG_SectionA_IncomeAmt1;PTI_CG_STCG_SectionA_ShareAmt1;PTI_CG_STCG_SectionA_NetAmt1;PTI_CG_STCG_SectionA_TdsAmt1;PTI_CG_STCG_Others_IncomeAmt1;PTI_CG_STCG_Others_ShareAmt1;PTI_CG_STCG_Others_NetAmt1;PTI_CG_STCG_Others_TdsAmt1;PTI_CG_LTCGIncomeAmt1;PTI_CG_LTCG_ShareAmt1;PTI_CG_LTCG_NetAmt1;PTI_CG_LTCGTDSAmt1;PTI_CG_LTCG_SectionA_IncomeAmt1;PTI_CG_LTCG_SectionA_ShareAmt1;PTI_CG_LTCG_SectionA_NetAmt1;PTI_CG_LTCG_SectionA_TdsAmt1;PTI_CG_LTCG_Others_IncomeAmt1;PTI_CG_LTCG_Others_ShareAmt1;PTI_CG_LTCG_Others_NetAmt1;PTI_CG_LTCG_Others_TdsAmt1;PTI_OSIncomeAmt1;PTI_OS_NetAmt1;PTI_OSTDSAmt1;PTI_OS_Dividend_IncomeAmt1;PTI_OS_Dividend_NetAmt1;PTI_OS_Dividend_TdsAmt1;PTI_OS_Others_IncomeAmt1;PTI_OS_Others_NetAmt1;PTI_OS_Others_TdsAmt1;PTI_Us_iv_IncomeAmta1;PTI_Us_iv_NetAmta1;PTI_Us_iv_TdsAmta1;" & _
 "PTI_Usa1;PTI_UsIncomeAmta1;PTI_Us_NetAmta1;PTI_UsTDSAmta1;PTI_Usb1;PTI_UsIncomeAmtb1;PTI_Us_NetAmtb1;PTI_UsTDSAmtb1;PTI_Usc1;PTI_UsIncomeAmtc1;PTI_Us_NetAmtc1;PTI_UsTDSAmtc1;"
End Sub


Sub setTblinfo_PTI_Check()
    Dim countrycd As Variant
    Dim ccount1, ccount2, ccount3, ccount4, ccount5, ccount6, ccount7, ccount8, ccount9, ccount10, ccount11, ccount12, ccount13, ccount14, ccount15, ccount16, ccount17, ccount18, ccount19, ccount20, ccount21, ccount22, ccount23, ccount24, ccount25, ccount26, ccount27, ccount28, ccount29, ccount30, ccount31 As Long
    Dim te As Long
    
    ccount1 = 0
    ccount2 = 0
    ccount3 = 0
    ccount4 = 0
    ccount5 = 0
    ccount6 = 0
    ccount7 = 0
    ccount8 = 0
    ccount9 = 0
    ccount10 = 0
    ccount11 = 0
    ccount12 = 0
    ccount13 = 0
    ccount14 = 0
    ccount15 = 0
    ccount16 = 0
    ccount17 = 0
    ccount18 = 0
     ccount19 = 0
    ccount20 = 0
    ccount21 = 0
    ccount22 = 0
    ccount23 = 0
    ccount24 = 0
    ccount25 = 0
    ccount26 = 0
    ccount27 = 0
    ccount28 = 0
    ccount29 = 0
    ccount30 = 0
     ccount31 = 0
    
    frmsize_Ptirptfrm = Sheet41.Range("Ptirptfrm.size").value
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_NameOfBusiness" & te).value
        
        If Not countrycd = "" Then
            ccount1 = ccount1 + 1
        Else
            'Exit For
        End If
    Next
    
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_PAN" & te).value
        
        If Not countrycd = "" Then
            ccount2 = ccount2 + 1
        Else
            'Exit For
        End If
    Next
    
        countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_HPIncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount3 = ccount3 + 1
        Else
            'Exit For
        End If
    Next
    
          countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_HP_ShareAmt" & te).value
        
        If Not countrycd = "" Then
            ccount4 = ccount4 + 1
        Else
            'Exit For
        End If
    Next
    
        
          countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_HPTDSAmt" & te).value
        
        If Not countrycd = "" Then
            ccount5 = ccount5 + 1
        Else
            'Exit For
        End If
    Next
    
              countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCGTDSAmt" & te).value
        
        If Not countrycd = "" Then
            ccount6 = ccount6 + 1
        Else
            'Exit For
        End If
    Next
    
    
    
           countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount7 = ccount7 + 1
        Else
            'Exit For
        End If
    Next
    
    
        
           countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & te).value
        
        If Not countrycd = "" Then
            ccount8 = ccount8 + 1
        Else
            'Exit For
        End If
    Next
    
              countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & te).value
        
        If Not countrycd = "" Then
            ccount9 = ccount9 + 1
        Else
            'Exit For
        End If
    Next
    
        
              countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount10 = ccount10 + 1
        Else
            'Exit For
        End If
    Next
    
                  countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & te).value
        
        If Not countrycd = "" Then
            ccount11 = ccount11 + 1
        Else
            'Exit For
        End If
    Next
    
                      countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & te).value
        
        If Not countrycd = "" Then
            ccount12 = ccount12 + 1
        Else
            'Exit For
        End If
    Next
    
                      countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount13 = ccount13 + 1
        Else
            'Exit For
        End If
    Next
        
                      countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & te).value
        
        If Not countrycd = "" Then
            ccount14 = ccount14 + 1
        Else
            'Exit For
        End If
    Next
    
                          countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & te).value
        
        If Not countrycd = "" Then
            ccount15 = ccount15 + 1
        Else
            'Exit For
        End If
    Next
    
                              countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount16 = ccount16 + 1
        Else
            'Exit For
        End If
    Next
    
                                  countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & te).value
        
        If Not countrycd = "" Then
            ccount17 = ccount17 + 1
        Else
            'Exit For
        End If
    Next
    
                                      countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & te).value
        
        If Not countrycd = "" Then
            ccount18 = ccount18 + 1
        Else
            'Exit For
        End If
    Next
    
                                          countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount19 = ccount19 + 1
        Else
            'Exit For
        End If
    Next
    
                                              countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_OS_Dividend_TdsAmt" & te).value
        
        If Not countrycd = "" Then
            ccount20 = ccount20 + 1
        Else
            'Exit For
        End If
    Next
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_OS_Others_IncomeAmt" & te).value
        
        If Not countrycd = "" Then
            ccount21 = ccount21 + 1
        Else
            'Exit For
        End If
    Next
    
        countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_OS_Others_TdsAmt" & te).value
        
        If Not countrycd = "" Then
            ccount22 = ccount22 + 1
        Else
            'Exit For
        End If
    Next
    
            countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_Us_iv_TdsAmta" & te).value
        
        If Not countrycd = "" Then
            ccount23 = ccount23 + 1
        Else
            'Exit For
        End If
    Next

    
  '12
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_UsIncomeAmta" & te).value
        
        If Not countrycd = "" Then
            ccount24 = ccount24 + 1
        Else
            'Exit For
        End If
    Next
    
    
'--
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_UsTDSAmta" & te).value
        
        If Not countrycd = "" Then
            ccount25 = ccount25 + 1
        Else
            'Exit For
        End If
    Next
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_Usb" & te).value
        
        If Not countrycd = "" Then
            ccount26 = ccount26 + 1
        Else
            'Exit For
        End If
    Next
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_UsIncomeAmtb" & te).value
        
        If Not countrycd = "" Then
            ccount27 = ccount27 + 1
        Else
            'Exit For
        End If
    Next
    
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_UsTDSAmtb" & te).value
        
        If Not countrycd = "" Then
            ccount28 = ccount28 + 1
        Else
            'Exit For
        End If
    Next
        
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_Usc" & te).value
        
        If Not countrycd = "" Then
            ccount29 = ccount29 + 1
        Else
            'Exit For
        End If
    Next
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_UsIncomeAmtc" & te).value
        
        If Not countrycd = "" Then
            ccount30 = ccount30 + 1
        Else
            'Exit For
        End If
    Next
    
    countrycd = ""
    For te = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        countrycd = Sheet41.Range("PTI_UsTDSAmtc" & te).value
        
        If Not countrycd = "" Then
            ccount31 = ccount31 + 1
        Else
            'Exit For
        End If
    Next
    
    end_PTI_1 = ccount1
    end_PTI_2 = ccount2
    end_PTI_3 = ccount3
    end_PTI_4 = ccount4
    end_PTI_5 = ccount5
    end_PTI_6 = ccount6
    end_PTI_7 = ccount7
    end_PTI_8 = ccount8
    end_PTI_9 = ccount9
    end_PTI_10 = ccount10
    end_PTI_11 = ccount11
    end_PTI_12 = ccount12
    end_PTI_13 = ccount13
    end_PTI_14 = ccount14
    end_PTI_15 = ccount15
    end_PTI_16 = ccount16
    end_PTI_17 = ccount17
    end_PTI_18 = ccount18
    end_PTI_19 = ccount19
    end_PTI_20 = ccount20
    end_PTI_21 = ccount21
    end_PTI_22 = ccount22
    end_PTI_23 = ccount23
    end_PTI_24 = ccount24
    end_PTI_25 = ccount25
    end_PTI_26 = ccount26
    end_PTI_27 = ccount27
    end_PTI_28 = ccount28
    end_PTI_29 = ccount29
    end_PTI_30 = ccount30
    end_PTI_31 = ccount31
       
End Sub

Function ValidatesheetPTI() As Boolean
    ValidatesheetPTI = True
     If Not ValidateInvestmentEntityOfBusiness_PTI() Then ValidatesheetPTI = False
    If (Len(Sheet41.Range("PTI_Investment_UA_UB1").item(1).value) > 0 And Sheet41.Range("PTI_Investment_UA_UB1").item(1).value <> "(Select)") Then
        If Not ValidateNameOfBusiness_PTI() Then ValidatesheetPTI = False
        If Not ValidatePANOfBusiness_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncFromHP_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncSTCGCapGain_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncLTCGCapGain_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncOthSrc_PTI() Then ValidatesheetPTI = False
        If Not ValidateSectionClaimedExempt_A_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncClaimedExempt_A_PTI() Then ValidatesheetPTI = False
        
        If Not ValidateSectionClaimedExempt_B_PTI() Then ValidatesheetPTI = False
        If Not ValidateSectionClaimedExempt_C_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncClaimedExempt_B_PTI() Then ValidatesheetPTI = False
        If Not ValidateIncClaimedExempt_C_PTI() Then ValidatesheetPTI = False
    End If
    
    If ValidatesheetPTI Then
        setTblinfo_PTI2
        setTblinfo_PTI1
       If (end_PTI <> end_PTI2) Then
            msgbox_PTI ("* PTI Details from business trust/ investment fund as per section 115UA, 115UB in Schedule PTI are mandatory")
            ValidatesheetPTI = False
            Exit Function
        End If
    
    
        setTblinfo_PTI_Check
    
        If (end_PTI <> end_PTI_1) Or (end_PTI <> end_PTI_2) Or (end_PTI <> end_PTI_3) Or (end_PTI <> end_PTI_4) Or (end_PTI <> end_PTI_5) Or (end_PTI <> end_PTI_6) Or (end_PTI <> end_PTI_7) Or (end_PTI <> end_PTI_8) Or (end_PTI <> end_PTI_9) Or (end_PTI <> end_PTI_10) Or (end_PTI <> end_PTI_11) Or (end_PTI <> end_PTI_12) Or (end_PTI <> end_PTI_13) Or (end_PTI <> end_PTI_14) Or _
                 (end_PTI <> end_PTI_15) Or (end_PTI <> end_PTI_16) Or (end_PTI <> end_PTI_17) Or (end_PTI <> end_PTI_18) Or (end_PTI <> end_PTI_19) Or (end_PTI <> end_PTI_20) Or (end_PTI <> end_PTI_21) Or (end_PTI <> end_PTI_22) Or (end_PTI <> end_PTI_23) Or (end_PTI <> end_PTI_24) Or (end_PTI <> end_PTI_25) Then
            ChkMandPTI
            'msgbox_PTI ("* PTI Details from business trust/ investment fund as per section 115UA, 115UB in Sch PTI are mandatory")
            ValidatesheetPTI = False
            Exit Function
        End If

        If Not ChkMandPTI_1 Then
            ValidatesheetPTI = False
            Exit Function
        End If
        
        If ((end_PTI_26 <> end_PTI_27) Or _
            (end_PTI_26 <> end_PTI_28) Or _
            (end_PTI_29 <> end_PTI_30) Or _
            (end_PTI_29 <> end_PTI_31)) Then
            
            'msgbox_PTI ("* PTI Details from business trust/ investment fund as per section 115UA, 115UB in Sch PTI are mandatory")
            If Not ChkMandPTI_2 Then
                ValidatesheetPTI = False
                Exit Function
            End If
        End If
    End If
End Function

Function ChkMandPTI() As Boolean
On Error Resume Next
    ChkMandPTI = True
    Dim i, CntrPTI As Long
    Dim PANOfBusiness_PTI_1 As Variant
    Dim NameOfBusiness_PTI_1, Investment115UA_PTI_1 As Variant
        
    CntrPTI = Range("cntr.Ptirptfrm")
    ReDim PANOfBusiness_PTI_1(CntrPTI)
    ReDim NameOfBusiness_PTI_1(CntrPTI)
    ReDim Investment115UA_PTI_1(CntrPTI)
    For i = 1 To CntrPTI
        Investment115UA_PTI_1(i) = Sheet41.Range("PTI_Investment_UA_UB" & i).value
        PANOfBusiness_PTI_1(i) = Sheet41.Range("PTI_PAN" & i).value
        NameOfBusiness_PTI_1(i) = Sheet41.Range("PTI_NameOfBusiness" & i).value
        
        If isdropdownblank(Investment115UA_PTI_1(i)) Then
'            msgbox_PTI ("* Investment entity covered by section 115UA/115UB at Section   " & i & " in Schedule PTI is mandatory.")
            msgbox_PTI ("*""please Select relevant section""")
            
            
            ChkMandPTI = False
            Exit Function
        End If
        If NameOfBusiness_PTI_1(i) = "" Then
            msgbox_PTI ("* Name of business trust/ investment fund is mandatory in Schedule PTI in Sl.no.1")
            ChkMandPTI = False
            Exit Function
        End If
        
        If PANOfBusiness_PTI_1(i) = "" Then
            msgbox_PTI ("* PAN of business trust/ investment fund at Section   " & i & " in Schedule PTI is mandatory.")
            ChkMandPTI = False
            Exit Function
        End If
                
    Next
End Function


Function ChkMandPTI_1() As Boolean
On Error Resume Next
    ChkMandPTI_1 = True
    Dim i, CntrPTI As Long
    Dim PTI_Usb_1, PTI_Usc_1 As Variant
    Dim NameOfBusiness_PTI_1 As Variant
    
    CntrPTI = Range("cntr.Ptirptfrm")
    ReDim PTI_Usb_1(CntrPTI)
    ReDim PTI_Usc_1(CntrPTI)
    ReDim NameOfBusiness_PTI_1(CntrPTI)
    
    For i = 1 To CntrPTI
        PTI_Usb_1(i) = Sheet41.Range("PTI_Usb" & i).value
        PTI_Usc_1(i) = Sheet41.Range("PTI_Usc" & i).value
        NameOfBusiness_PTI_1(i) = Sheet41.Range("PTI_NameOfBusiness" & i).value

        If PTI_Usb_1(i) <> "" Then
            If NameOfBusiness_PTI_1(i) = "" Then
                msgbox_PTI ("* Name Of Business at Section " & i & " in Schedule PTI is mandatory.")
                ChkMandPTI_1 = False
                Exit Function
            End If
        End If
        
        If PTI_Usc_1(i) <> "" Then
            If NameOfBusiness_PTI_1(i) = "" Then
                msgbox_PTI ("* Name Of Business at Section " & i & " in Schedule PTI is mandatory.")
                ChkMandPTI_1 = False
                Exit Function
            End If
        End If
    Next
End Function

Function ChkMandPTI_2() As Boolean
On Error Resume Next
    ChkMandPTI_2 = True
    Dim i, CntrPTI As Long
    
    Dim PTI_Usb_1, PTI_Usc_1 As Variant
    Dim PTI_UsTDSAmtb_1, PTI_UsIncomeAmtb_1 As Variant
    Dim PTI_UsIncomeAmtc_1, PTI_UsTDSAmtc_1 As Variant
    
    CntrPTI = Range("cntr.Ptirptfrm")
    ReDim PTI_Usb_1(CntrPTI)
    ReDim PTI_Usc_1(CntrPTI)
    ReDim PTI_UsTDSAmtb_1(CntrPTI)
    ReDim PTI_UsIncomeAmtb_1(CntrPTI)
    ReDim PTI_UsIncomeAmtc_1(CntrPTI)
    ReDim PTI_UsTDSAmtc_1(CntrPTI)
    For i = 1 To CntrPTI
    
        PTI_Usb_1(i) = Sheet41.Range("PTI_Usb" & i).value
        PTI_Usc_1(i) = Sheet41.Range("PTI_Usc" & i).value
        PTI_UsIncomeAmtb_1(i) = Sheet41.Range("PTI_UsIncomeAmtb" & i).value
        PTI_UsTDSAmtb_1(i) = Sheet41.Range("PTI_UsTDSAmtb" & i).value

        PTI_UsIncomeAmtc_1(i) = Sheet41.Range("PTI_UsIncomeAmtc" & i).value
        PTI_UsTDSAmtc_1(i) = Sheet41.Range("PTI_UsTDSAmtc" & i).value
    
        If ((PTI_UsIncomeAmtb_1(i) <> "" Or PTI_UsTDSAmtb_1(i) <> "") And Trim(PTI_Usb_1(i)) = "") Then
            msgbox_PTI ("* U/s(b) under which Income claimed to be exempt at Section  " & i & " in Schedule PTI is mandatory.")
            ChkMandPTI_2 = False
            Exit Function
        End If
        
        If ((PTI_UsIncomeAmtc_1(i) <> "" Or PTI_UsTDSAmtc_1(i) <> "") And Trim(PTI_Usc_1(i)) = "") Then
            msgbox_PTI ("* U/s(c) under which Income claimed to be exempt at Section   " & i & " Schedule PTI is mandatory.")
            ChkMandPTI_2 = False
            Exit Function
        End If
    Next
End Function


Function ValidateNameOfBusiness_PTI() As Boolean
    Dim i As Long

    ValidateNameOfBusiness_PTI = True
    setTblinfo_PTI

    ReDim NameOfBusiness_PTI(end_PTI)
    For i = 1 To end_PTI
        NameOfBusiness_PTI(i) = Sheet41.Range("PTI_NameOfBusiness" & i).value
        
        If Len(NameOfBusiness_PTI(i)) = 0 Then
            'Exit For
        End If
        
        If Not chkCompulsory(NameOfBusiness_PTI(i)) Then
'            msgbox_PTI ("* Name Of Business at Section   " & i & "  in Sheet PTI  is mandatory")

        'SIT-69617 Changed by sai
             'msgbox_PTI ("*Name of business trust/ investment fund is mandatory in Schedule-PTI in Sl.no " & i & """")
             
             'Added by Aavula
             msgbox_PTI ("* Name of business trust/ investment fund is mandatory in Schedule PTI in Sl.no.1""")
            ValidateNameOfBusiness_PTI = False
            Exit Function
        End If
        
        If Not CheckfewSpecialCharacter(NameOfBusiness_PTI(i)) Then
            msgbox_PTI ("* Name Of Business trust/ investment fund at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateNameOfBusiness_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateInvestmentEntityOfBusiness_PTI() As Boolean
Dim i As Long

    ValidateInvestmentEntityOfBusiness_PTI = True
    setTblinfo_PTI

    ReDim Investment115UAOfBusiness_PTI(end_PTI)
    For i = 1 To end_PTI
        Investment115UAOfBusiness_PTI(i) = Sheet41.Range("PTI_Investment_UA_UB" & i).value
        
        If Len(Investment115UAOfBusiness_PTI(i)) = 0 Then
            'Exit For
        End If

        If Not CheckfewSpecialCharacter(Investment115UAOfBusiness_PTI(i)) Then
            msgbox_PTI ("* Investment entity covered by section 115UA/115UB at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateInvestmentEntityOfBusiness_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidatePANOfBusiness_PTI() As Boolean
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells2B As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim i As Long
  
    ValidatePANOfBusiness_PTI = True
    setTblinfo_PTI
    
    ReDim PANOfBusiness_PTI(end_PTI)
    For i = 1 To end_PTI
        PANOfBusiness_PTI(i) = Sheet41.Range("PTI_PAN" & i).value
        
        If Len(NameOfBusiness_PTI(i)) = 0 Then
         'Exit For
        End If
        
        If Not CheckDoneePAN(PANOfBusiness_PTI(i)) Then
        
        'SIT 91167 2
           ' msgbox_PTI ("* PAN Of Business Trust at Section   " & i & " in Schedule PTI is invalid")
             msgbox_PTI ("* ""Invalid PAN""")
            ValidatePANOfBusiness_PTI = False
            Exit Function
        End If
        
        If Not chkCompulsory(PANOfBusiness_PTI(i)) Then
         'SIT 91167 3
           ' msgbox_PTI ("* PAN Of Business Trust at Section   " & i & "  in Schedule PTI is mandatory")
            msgbox_PTI ("*""Enter 10 digit valid Pan in Schedule PTI in Sl.no." & i & """")
            ValidatePANOfBusiness_PTI = False
            Exit Function
       End If
        
        If Not checkfieldspecialcharacter(PANOfBusiness_PTI(i)) Then
         'SIT 91167 4
           'msgbox_PTI ("* PAN Of Business Trust at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            msgbox_PTI ("*""Invalid PAN in Schedule PTI in Sl.no." & i & " PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.""")
            ValidatePANOfBusiness_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncFromHP_PTI() As Boolean
    Dim i As Long
    
    ValidateIncFromHP_PTI = True
    setTblinfo_PTI

    ReDim PTI_HPIncomeAmt(end_PTI)
    ReDim PTI_HPTDSAmt(end_PTI)
    ReDim PTI_HPShareAmt(end_PTI)
    For i = 1 To end_PTI
        PTI_HPIncomeAmt(i) = Sheet41.Range("PTI_HPIncomeAmt" & i).value
        PTI_HPTDSAmt(i) = Sheet41.Range("PTI_HPTDSAmt" & i).value
        PTI_HPShareAmt(i) = Sheet41.Range("PTI_HP_ShareAmt" & i).value
        
        If Not chkCompulsory(PTI_HPIncomeAmt(i)) Then
            msgbox_PTI ("* Income From HP at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncFromHP_PTI = False
            Exit Function
        End If
        
          If Not chkCompulsory(PTI_HPShareAmt(i)) Then
            msgbox_PTI ("* Share of current year loss distributed by Investment fund From HP at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncFromHP_PTI = False
            Exit Function
        End If

        
        If Not chkCompulsory(PTI_HPTDSAmt(i)) Then
            msgbox_PTI ("* TDS From HP at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncFromHP_PTI = False
            Exit Function
        End If
        
        If (PTI_HPTDSAmt(i) < 0) Then
            msgbox_PTI ("* TDS From HP at Section   " & i & "  in Schedule PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncFromHP_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PTI_HPTDSAmt(i)) Then
            msgbox_PTI ("* TDS From HP at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncFromHP_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncSTCGCapGain_PTI() As Boolean
    Dim i As Long
    
    ValidateIncSTCGCapGain_PTI = True
    setTblinfo_PTI

    ReDim PTI_CG_STCGIncomeAmt(end_PTI)
    ReDim PTI_CG_STCGTDSAmt(end_PTI)
    For i = 1 To end_PTI
        'PTI_CG_STCGIncomeAmt(i) = Sheet41.Range("PTI_CG_STCGIncomeAmt" & i).value
        PTI_CG_STCGTDSAmt(i) = Sheet41.Range("PTI_CG_STCGTDSAmt" & i).value

        If Not chkCompulsory(Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & i).value) Then
            msgbox_PTI ("* Income From Capital Gain(STCG) Section 111A  at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
            If Not chkCompulsory(Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & i).value) Then
            msgbox_PTI ("*share of current year loss Income From Capital Gain(STCG) Section 111A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        
            If Not chkCompulsory(Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & i).value) Then
            msgbox_PTI ("* Income From Capital Gain(STCG) Others at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        
            If Not chkCompulsory(Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & i).value) Then
            msgbox_PTI ("* share of current year loss Income From Capital Gain(STCG) Others at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If

        If Not chkCompulsory(PTI_CG_STCGTDSAmt(i)) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        If (PTI_CG_STCGTDSAmt(i) < 0) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) at Section   " & i & "  in Sheet PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PTI_CG_STCGTDSAmt(i)) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        'new
        
             
        If Not chkCompulsory(Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) Section 111A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        If (Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & i).value < 0) Then
            msgbox_PTI ("* TDS From capital Gain(STCG)Section 111A at Section   " & i & "  in Sheet PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) Section 111A at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        '2
         If Not chkCompulsory(Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) Others at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        If (Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & i).value < 0) Then
            msgbox_PTI ("* TDS From capital Gain(STCG)Others at Section   " & i & "  in Sheet PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(STCG) Others at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncSTCGCapGain_PTI = False
            Exit Function
        End If

    Next
End Function

Function ValidateIncLTCGCapGain_PTI() As Boolean
    Dim i As Long
    
    ValidateIncLTCGCapGain_PTI = True
    setTblinfo_PTI

    ReDim PTI_CG_LTCGIncomeAmt(end_PTI)
    ReDim PTI_CG_LTCGTDSAmt(end_PTI)
    For i = 1 To end_PTI
        PTI_CG_LTCGIncomeAmt(i) = Sheet41.Range("PTI_CG_LTCGIncomeAmt" & i).value
        PTI_CG_LTCGTDSAmt(i) = Sheet41.Range("PTI_CG_LTCGTDSAmt" & i).value

        If Not chkCompulsory(Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & i).value) Then
            msgbox_PTI ("* Income From Capital Gain(LTCG) Section 112A  at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
            If Not chkCompulsory(Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & i).value) Then
            msgbox_PTI ("*share of current year loss Income From Capital Gain(LTCG) Section 112A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        
            If Not chkCompulsory(Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & i).value) Then
            msgbox_PTI ("* Income From Capital Gain(LTCG) Other than Section 112A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        
            If Not chkCompulsory(Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & i).value) Then
            msgbox_PTI ("* share of current year loss Income From Capital Gain(LTCG) Other than Section 112A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
     'TDS
        If Not chkCompulsory(PTI_CG_LTCGTDSAmt(i)) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        If (PTI_CG_LTCGTDSAmt(i) < 0) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) at Section   " & i & "  in Sheet PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PTI_CG_LTCGTDSAmt(i)) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
         'new
        
             
        If Not chkCompulsory(Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) Section 112A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        If (Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & i).value < 0) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG)Section 112A at Section   " & i & "  in Sheet PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) Section 112A at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        '2
         If Not chkCompulsory(Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) Other than Section 112A at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        If (Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & i).value < 0) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) Other than Section 112A at Section   " & i & "  in Sheet PTI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From capital Gain(LTCG) Other than Section 112A at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncLTCGCapGain_PTI = False
            Exit Function
        End If
        
    Next
End Function

Function ValidateIncOthSrc_PTI() As Boolean
    Dim i As Long
    
    ValidateIncOthSrc_PTI = True
    setTblinfo_PTI
    
    ReDim PTI_OSIncomeAmt(end_PTI)
    ReDim PTI_OSTDSAmt(end_PTI)
    For i = 1 To end_PTI
        PTI_OSIncomeAmt(i) = Sheet41.Range("PTI_OSIncomeAmt" & i).value
        PTI_OSTDSAmt(i) = Sheet41.Range("PTI_OSTDSAmt" & i).value

        If Not chkCompulsory(Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & i).value) Then
            msgbox_PTI ("* Income From Dividend (referred to in section 115O)s at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
        
           If Not chkCompulsory(Sheet41.Range("PTI_OS_Others_IncomeAmt" & i).value) Then
            msgbox_PTI ("* Income From Others(Other Sources) at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
    
        
        If Not chkCompulsory(PTI_OSTDSAmt(i)) Then
            msgbox_PTI ("* TDS From Other Sources at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PTI_OSTDSAmt(i)) Then
            msgbox_PTI ("* TDS From Other Sources at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
      
      
      
      'new
        If Not chkCompulsory(Sheet41.Range("PTI_OS_Dividend_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From Dividend (referred to in section 115O) at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Sheet41.Range("PTI_OS_Dividend_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From Dividend (referred to in section 115O) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
        '2
        If Not chkCompulsory(Sheet41.Range("PTI_OS_Others_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From  Others(Other Sources) at Section   " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Sheet41.Range("PTI_OS_Others_TdsAmt" & i).value) Then
            msgbox_PTI ("* TDS From Others(Other Sources) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncOthSrc_PTI = False
            Exit Function
        End If

    
    Next
End Function

Function ValidateSectionClaimedExempt_A_PTI()
ValidateSectionClaimedExempt_A_PTI = True
    Dim i As Long

    setTblinfo_PTI
    
    ReDim PTI_Usa(end_PTI)
    For i = 1 To end_PTI
        PTI_Usa(i) = Sheet41.Range("PTI_Usa" & i).value
        If Len(PTI_Usa(i)) = 0 Then
        'Exit For
        End If

    Next
End Function

Function ValidateSectionClaimedExempt_B_PTI()
ValidateSectionClaimedExempt_B_PTI = True
    Dim i As Long

    setTblinfo_PTI
    
    ReDim PTI_Usb(end_PTI)
    For i = 1 To end_PTI
        PTI_Usb(i) = Sheet41.Range("PTI_Usb" & i).value
        If Len(PTI_Usb(i)) = 0 Then
        'Exit For
        End If

        If Not CheckfewSpecialCharacter(PTI_Usb(i)) Then
            msgbox_PTI ("* u/s(b) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateSectionClaimedExempt_B_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateSectionClaimedExempt_C_PTI()
ValidateSectionClaimedExempt_C_PTI = True
    Dim i As Long

    setTblinfo_PTI
    
    ReDim PTI_Usc(end_PTI)
    For i = 1 To end_PTI
        PTI_Usc(i) = Sheet41.Range("PTI_Usc" & i).value
        If Len(PTI_Usc(i)) = 0 Then
        'Exit For
        End If
       
        If Not CheckfewSpecialCharacter(PTI_Usc(i)) Then
            msgbox_PTI ("* u/s(c) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateSectionClaimedExempt_C_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncClaimedExempt_A_PTI()
ValidateIncClaimedExempt_A_PTI = True

    Dim i As Long

    setTblinfo_PTI
    
    ReDim PTI_UsIncomeAmta(end_PTI)
    ReDim PTI_UsTDSAmta(end_PTI)
    ReDim PTI_Us_iv_TdsAmta(end_PTI)
    For i = 1 To end_PTI
        PTI_UsIncomeAmta(i) = Sheet41.Range("PTI_UsIncomeAmta" & i).value
        PTI_UsTDSAmta(i) = Sheet41.Range("PTI_UsTDSAmta" & i).value
        PTI_Us_iv_TdsAmta(i) = Sheet41.Range("PTI_Us_iv_TdsAmta" & i).value
        If Len(PTI_UsIncomeAmta(i)) = 0 Then
        'Exit For
        End If
        If Not chkCompulsory(PTI_Us_iv_TdsAmta(i)) Then
            msgbox_PTI ("* TDS Amount claimed to be exempt at Section " & i & " Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncClaimedExempt_A_PTI = False
            Exit Function
        End If
        
        If Not chkCompulsory(PTI_UsIncomeAmta(i)) Then
            msgbox_PTI ("* Income Amount claimed to be exempt(a) at Section " & i & "  in Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncClaimedExempt_A_PTI = False
            Exit Function
        End If
    
        If Not chkCompulsory(PTI_UsTDSAmta(i)) Then
            msgbox_PTI ("* TDS Amount claimed to be exempt(a) at Section " & i & " Schedule PTI is mandatory. Pl fill zero if no Income.")
            ValidateIncClaimedExempt_A_PTI = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PTI_UsTDSAmta(i)) Then
            msgbox_PTI ("* TDS Amount claimed to be exempt(a) at Section " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncClaimedExempt_A_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncClaimedExempt_B_PTI()
ValidateIncClaimedExempt_B_PTI = True

    Dim i As Long

    setTblinfo_PTI
    
    ReDim PTI_UsIncomeAmtb(end_PTI)
    ReDim PTI_UsTDSAmtb(end_PTI)
    
    For i = 1 To end_PTI
        PTI_UsIncomeAmtb(i) = Sheet41.Range("PTI_UsIncomeAmtb" & i).value
        PTI_UsTDSAmtb(i) = Sheet41.Range("PTI_UsTDSAmtb" & i).value

        
        If PTI_Usb(i) <> "" Then
            If Not chkCompulsory(PTI_UsIncomeAmtb(i)) Then
                msgbox_PTI ("* Income Amount claimed to be exempt(b) at Section   " & i & " in Schedule PTI is mandatory. Pl fill zero if no Income.")
                ValidateIncClaimedExempt_B_PTI = False
                Exit Function
            End If
        End If

        If PTI_Usb(i) <> "" Then
            If Not chkCompulsory(PTI_UsTDSAmtb(i)) Then
                msgbox_PTI ("* TDS Amount claimed to be exempt(b) at Section   " & i & " in Schedule PTI is mandatory. Pl fill zero if no Income.")
                ValidateIncClaimedExempt_B_PTI = False
                Exit Function
            End If
        End If
        
        If Not checkfieldspecialcharacter(PTI_UsTDSAmtb(i)) Then
            msgbox_PTI ("* TDS Amount claimed to be exempt(b) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncClaimedExempt_B_PTI = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncClaimedExempt_C_PTI()
ValidateIncClaimedExempt_C_PTI = True

    Dim i As Long

    setTblinfo_PTI
    
    ReDim PTI_UsIncomeAmtc(end_PTI)
    ReDim PTI_UsTDSAmtc(end_PTI)
    For i = 1 To end_PTI
        PTI_UsIncomeAmtc(i) = Sheet41.Range("PTI_UsIncomeAmtc" & i).value
        PTI_UsTDSAmtc(i) = Sheet41.Range("PTI_UsTDSAmtc" & i).value
        
        If PTI_Usc(i) <> "" Then
            If Not chkCompulsory(PTI_UsIncomeAmtc(i)) Then
                msgbox_PTI ("* Income Amount claimed to be exempt(c) at Section   " & i & " in Schedule PTI is mandatory. Pl fill zero if no Income.")
                ValidateIncClaimedExempt_C_PTI = False
                Exit Function
            End If
        End If
  
        If PTI_Usc(i) <> "" Then
            If Not chkCompulsory(PTI_UsTDSAmtc(i)) Then
                msgbox_PTI ("* TDS Amount claimed to be exempt(c) at Section   " & i & " in Schedule PTI is mandatory. Pl fill zero if no Income.")
                ValidateIncClaimedExempt_C_PTI = False
                Exit Function
            End If
        End If
        
        If Not checkfieldspecialcharacter(PTI_UsTDSAmtc(i)) Then
            msgbox_PTI ("* TDS Amount claimed to be exempt(c) at Section   " & i & " in Schedule PTI characters < > & ' " & Chr(34) & " are not allowed")
            ValidateIncClaimedExempt_C_PTI = False
            Exit Function
        End If
    
    Next
End Function

Function msgbox_PTI(strmsg As String) As String
     msgValidatesheetPTI = msgValidatesheetPTI & strmsg & Chr(13)
End Function

Sub LOCKPTI()
On Error GoTo pfsi
    Dim strpassword As String
    strpassword = getmsgstate
    Sheet41.Unprotect strpassword
    
    Dim prngname_5_PTI As Variant
    prngname_5_PTI = "PTI_NameOfBusiness;PTI_PAN;PTI_HPIncomeAmt;PTI_HPTDSAmt;PTI_CG_STCGIncomeAmt;PTI_CG_STCGTDSAmt;PTI_CG_LTCGIncomeAmt;PTI_CG_LTCGTDSAmt;PTI_OSIncomeAmt;PTI_OSTDSAmt;PTI_Usa;PTI_UsIncomeAmta;PTI_UsTDSAmta;PTI_Usb;PTI_UsIncomeAmtb;PTI_UsTDSAmtb;PTI_Usc;PTI_UsIncomeAmtc;PTI_UsTDSAmtc;"
    BlankPTI
    LOCKRANGEPTI (prngname_5_PTI)
pfsi:
    Sheet41.Protect Password:=strpassword
End Sub

Sub BlankPTI()
    Dim i, j As Long
    setTblinfo_PTI
    
    For j = 1 To Sheet41.Range("cntr.ptirptfrm").value
        Sheet41.Range("PTI_NameOfBusiness" & j).value = "(Select)"
        Sheet41.Range("PTI_PAN" & j).value = ""
        Sheet41.Range("PTI_HPIncomeAmt" & j).value = ""
        Sheet41.Range("PTI_CG_STCGIncomeAmt" & j).value = ""
        Sheet41.Range("PTI_CG_LTCGIncomeAmt" & j).value = ""
        Sheet41.Range("PTI_OSIncomeAmt" & j).value = ""
        Sheet41.Range("PTI_UsIncomeAmta" & j).value = ""
        Sheet41.Range("PTI_UsTDSAmta" & j).value = ""
        Sheet41.Range("PTI_Usb" & j).value = ""
        Sheet41.Range("PTI_UsIncomeAmtb" & j).value = ""
        Sheet41.Range("PTI_UsTDSAmtb" & j).value = ""
        Sheet41.Range("PTI_Usc" & j).value = ""
        Sheet41.Range("PTI_UsIncomeAmtc" & j).value = ""
        Sheet41.Range("PTI_UsTDSAmtc" & j).value = ""
    Next
End Sub

Sub UNLOCKPTI()
On Error GoTo pfsi
    Dim strpassword As String
    Dim prngname_5_PTI As Variant
    
    strpassword = getmsgstate
    Sheet41.Unprotect strpassword
    
    prngname_5_PTI = "PTI_NameOfBusiness;PTI_PAN;PTI_HPIncomeAmt;PTI_HPTDSAmt;PTI_CG_STCGIncomeAmt;PTI_CG_STCGTDSAmt;PTI_CG_LTCGIncomeAmt;PTI_CG_LTCGTDSAmt;PTI_OSIncomeAmt;PTI_OSTDSAmt;PTI_Usa;PTI_UsIncomeAmta;PTI_UsTDSAmta;PTI_Usb;PTI_UsIncomeAmtb;PTI_UsTDSAmtb;PTI_Usc;PTI_UsIncomeAmtc;PTI_UsTDSAmtc;"
    UNLOCKRANGEPTI (prngname_5_PTI)

pfsi:
    Sheet41.Protect Password:=strpassword
End Sub


Sub LOCKRANGEPTI(rangenamestring As Variant)
    Dim i, j As Long
    Dim chCells As Range
    rangenamestring = Split(rangenamestring, ";")
    For j = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        For i = 0 To UBound(rangenamestring) - 1
            For Each chCells In Sheet41.Range(rangenamestring(i) & j)
               chCells.value = ""
               chCells.Locked = True
               chCells.Interior.Color = (&HD8D8D8)
            Next
        Next
    Next
End Sub

Sub UNLOCKRANGEPTI(rangenamestring As Variant)
    Dim i, j As Long
    Dim chCells As Range
    rangenamestring = Split(rangenamestring, ";")
    For j = 1 To Sheet41.Range("cntr.Ptirptfrm").value
        For i = 0 To UBound(rangenamestring) - 1
            For Each chCells In Sheet41.Range(rangenamestring(i) & j)
                chCells.Locked = False
                chCells.Interior.Color = (&HCCFFCC)
            Next
        Next
    Next
End Sub

Function addPtiblock(prevrangename As Variant, blockname As Variant, frmcounter As Variant, totalblocksize As Variant)
    Dim i, j, k, l, commindex As Long
    Dim newnamerefersto As Variant
    Dim gapbtn, dcounter, counter, offset1, offset2 As Long
    Dim lenn, startrow, endrow, endrow1, endrow2 As Long
    Dim destinationrowindex, destinationcolumnindex, lastdestinationrow As String
    Dim rangearr, newranges, precrange, oldrangeaddress, newrangeaddress, newrangeaddress1, newrangeaddress2 As Variant
    Dim ndestinationrowindex, ndestinationcolumnindex As Long
    Dim noofrowsadded As Long
    Dim rangeCount As Long

    
    Application.EnableEvents = False
    
    sPassword = getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    gapbtn = 0
    dcounter = 0
    counter = Range(frmcounter).value
    
    Dim lastrow As Variant
    lastrow = onlyDigits(Sheet41.Range("PTI.LastCell").AddressLocal, "I")
    
    If counter > 0 Then
        noofrowsadded = (totalblocksize - 15) / counter
        offset1 = 15 + (1) * noofrowsadded
        offset2 = 15 + (counter) * noofrowsadded
        
        'offset1 = Range(blockname & 1).Rows.count
        ndestinationrowindex = Range(blockname & counter).row
        ndestinationrowindex = ndestinationrowindex + gapbtn + offset1
        ndestinationrowindex = lastrow
        ndestinationcolumnindex = Range(blockname & counter).Column
        newnamerefersto = Range(blockname & counter).Address
        lastdestinationrow = ndestinationrowindex + offset1 - 1
        Cells(CLng(ndestinationrowindex - 1), CLng(ndestinationcolumnindex)).Select
    
        Call InsertBlock(offset1)
    
        For i = 1 To Len(newnamerefersto)
            If (Mid(newnamerefersto, i, 1) = "$") Then
                dcounter = dcounter + 1
                If dcounter = 2 Then
                    lenn = (InStr(1, newnamerefersto, ":") - 1) - i
                    startrow = Mid(newnamerefersto, i + 1, lenn)
                    newnamerefersto = Replace(newnamerefersto, startrow, ndestinationrowindex)
                End If
                
                If dcounter = 4 Then
                    endrow = Mid(newnamerefersto, i + 1, Len(newnamerefersto) - i)
                    newnamerefersto = Replace(newnamerefersto, endrow, lastdestinationrow)
                End If
            End If
        Next
        
        Application.EnableEvents = False
    
        counter = counter + 1
        ThisWorkbook.Names.add name:=blockname & counter, _
        RefersTo:="=" & newnamerefersto, Visible:=True
    
        sPassword = getmsgstate
        ActiveSheet.Unprotect Password:=sPassword
        
        Range(frmcounter).value = counter
        
        Application.EnableEvents = False
    
        ActiveSheet.Unprotect Password:=sPassword
    
        'Range(blockname & counter - 1).Copy Destination:=Cells(ndestinationrowindex, ndestinationcolumnindex)
        Range(blockname & counter - 1).Copy Destination:=Cells(ndestinationrowindex, ndestinationcolumnindex)
    End If
    
    '' add all range name to block
    
    commindex = 2
    j = 0
    
    prevrangename = Split(prevrangename, ";")
    rangeCount = UBound(prevrangename)
    
    ReDim rangearr(rangeCount)
    commindex = 0
    j = 1
    
    For i = 0 To UBound(prevrangename)
        rangearr(i) = prevrangename(i)
    Next
    
    ReDim newranges(rangeCount)
    For i = 0 To UBound(rangearr)
        newranges(i) = Replace(rangearr(i), CStr(1), CStr(counter))
    Next
    
    If counter > 2 Then
        ReDim precrange(rangeCount)
        For i = 0 To UBound(rangearr)
            precrange(i) = Replace(rangearr(i), CStr(1), CStr(counter - 1))
        Next
    
        ReDim oldrangeaddress(rangeCount)
        For i = 0 To UBound(precrange)
            If Not (precrange(i) = "") Then
                oldrangeaddress(i) = Range(precrange(i)).Address
            End If
        Next
    Else
        ReDim oldrangeaddress(rangeCount)
        For i = 0 To UBound(rangearr)
            If Not (rangearr(i) = "") Then
                oldrangeaddress(i) = Range(rangearr(i)).Address
            End If
        Next
    End If
    
    ReDim newrangeaddress(rangeCount)
    For i = 0 To UBound(oldrangeaddress)
        dcounter = 0
        Dim myrangearr() As String
        myrangearr = Split(oldrangeaddress(i), ":")
        If UBound(myrangearr) <= 0 Then
            For k = 1 To Len(oldrangeaddress(i))
                If (Mid(oldrangeaddress(i), k, 1) = "$") Then
                    dcounter = dcounter + 1
                    If (dcounter = 2) Then
                        endrow = Mid(oldrangeaddress(i), k + 1, Len(oldrangeaddress(i)) - k)
                        newrangeaddress(i) = Replace(oldrangeaddress(i), endrow, (endrow + gapbtn + offset1))
                    End If
                End If
            Next
        Else
    
            Dim myoldrangeaddress() As String
            myoldrangeaddress = Split(oldrangeaddress(i), ":")
    
            For k = 1 To Len(myrangearr(0))
                If (Mid(myoldrangeaddress(0), k, 1) = "$") Then
                    dcounter = dcounter + 1
                    If (dcounter = 2) Then
                        endrow1 = Mid(myoldrangeaddress(0), k + 1, Len(myoldrangeaddress(0)) - k)
                        newrangeaddress1 = Replace(myoldrangeaddress(0), endrow1, (endrow1 + gapbtn + offset1))
                    End If
                End If
            Next
            
            dcounter = 0
            For k = 1 To Len(myrangearr(1))
                If (Mid(myoldrangeaddress(1), k, 1) = "$") Then
                    dcounter = dcounter + 1
                    If (dcounter = 2) Then
                        endrow2 = Mid(myoldrangeaddress(1), k + 1, Len(myoldrangeaddress(1)) - k)
                        newrangeaddress2 = Replace(myoldrangeaddress(1), endrow2, (endrow2 + gapbtn + offset1))
                    End If
                End If
            Next
            newrangeaddress(i) = newrangeaddress1 + ":" + newrangeaddress2
        End If
    Next
    
    For l = 0 To UBound(newrangeaddress)
        If Not newranges(l) = "" Then
            ThisWorkbook.Names.add name:=newranges(l), _
            RefersTo:="=" & newrangeaddress(l), Visible:=True
            If Range(newranges(l)).Interior.ColorIndex = 35 Then
                Range(newranges(l)).ClearContents
            End If
        End If
    Next
    ActiveSheet.Protect Password:=sPassword
    Application.EnableEvents = True
End Function

Function InsertPtiBlock(vRows1 As Variant)
    Dim x As Long
    Dim vRows As Long
    Dim sht As Worksheet, shts() As String, i As Long
    
    sPassword = getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    Application.EnableEvents = False
    ActiveCell.EntireRow.Select
    
    
    vRows = CLng(vRows1)
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
    Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        
        i = i + 1
        shts(i) = sht.name
        x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
        
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault
        
        On Error Resume Next
        
        Selection.Offset(1).Resize(vRows).EntireRow. _
        SpecialCells(xlConstants).ClearContents
    Next sht
    
    Worksheets(shts).Select
    ActiveSheet.Protect Password:=sPassword
    Application.EnableEvents = True
End Function

