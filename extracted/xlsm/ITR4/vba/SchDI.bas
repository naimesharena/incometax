Attribute VB_Name = "SchDI"
Public SectionDI, EligibleDI, MsgBox_DI As Variant
Sub ValidateSheetDI_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2021-22"
 
 ValidateDI

 
fmsgboxStatus "Sheet DI is OK"

End Sub
 Sub Prev80DI_Click()
Sheet5.Activate
End Sub
Sub Next80DI_Click()
Sheet16.Activate
End Sub


Function ValidateDI()
ValidateDI = True

If Not ChkDeductionDI Then ValidateDI = False
If Not ValidateDI Then
fmsgboxStatus MsgBox_DI
End
End If
End Function

Function ChkDeductionDI() As Boolean
    ChkDeductionDI = True
    
    
    
      If Sheet11.Range("DI_Deduction_80C").value > Sheet11.Range("DI_EligibleAmount_80C").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80C in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
      End If
 
 
 
 
 If Sheet11.Range("DI_Deduction_").value > Sheet11.Range("DI_EligibleAmount_").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction  in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 

 If Sheet11.Range("DI_Deduction_80CCD_1").value > Sheet11.Range("DI_EligibleAmount_80CCD_1").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80CCD(1) in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If
 
 
 If Sheet11.Range("DI_Deduction_80CCD_1B").value > Sheet11.Range("DI_EligibleAmount_80CCD_1B").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80CCD(1B) in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If
 
 If Sheet11.Range("DI_Deduction_80CCD_2").value > Sheet11.Range("DI_EligibleAmount_80CCD_2").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80CCD(2) in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If
 
' If Sheet11.Range("DI_Deduction_80CCG").Value > Sheet11.Range("DI_EligibleAmount_80CCG").Value Then
'            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80CCG in Sheet DI " & Chr(13)
'            ChkDeductionDI = False
'            Exit Function
' End If
 
 
 If Sheet11.Range("DI_Deduction_80D").value > Sheet11.Range("DI_EligibleAmount_80D").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80D in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 

 If Sheet11.Range("DI_Deduction_80DD").value > Sheet11.Range("DI_EligibleAmount_80DD").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80DD in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If
 

 If Sheet11.Range("DI_Deduction_80DDB").value > Sheet11.Range("DI_EligibleAmount_80DDB").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80DDB in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If
 

 If Sheet11.Range("DI_Deduction_80E").value > Sheet11.Range("DI_EligibleAmount_80E").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80E in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 If Sheet11.Range("DI_Deduction_80EE").value > Sheet11.Range("DI_EligibleAmount_80EE").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80EE in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 
 If Sheet11.Range("DI_Deduction_80EEA").value > Sheet11.Range("DI_EligibleAmount_80EEA").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80EEA in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 
 If Sheet11.Range("DI_Deduction_80EEB").value > Sheet11.Range("DI_EligibleAmount_80EEB").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80EEB in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 If Sheet11.Range("DI_Deduction_80G").value > Sheet11.Range("DI_EligibleAmount_80G").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80G in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 
 If Sheet11.Range("DI_Deduction_80GG").value > Sheet11.Range("DI_EligibleAmount_80GG").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80GG in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If

 

 
 
 
 If Sheet11.Range("DI_Deduction_80GGC").value > Sheet11.Range("DI_EligibleAmount_80GGC").value Then
            MsgBox_DI = MsgBox_DI + "* Col(2) cannot be greater than Col(1) for deduction 80GGC in Sheet DI " & Chr(13)
            ChkDeductionDI = False
            Exit Function
 End If
 


If Sheet1.Range("IncD.InvestmentDropDown").value = "Yes" Then
 If Sheet11.Range("DI_EligibleAmount_80C").value > 0 And Sheet11.Range("DI_Deduction_80C").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80C in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80C in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
  If Sheet11.Range("DI_EligibleAmount_").value > 0 And Sheet11.Range("DI_Deduction_").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
 If Sheet11.Range("DI_EligibleAmount_80CCD_1").value > 0 And Sheet11.Range("DI_Deduction_80CCD_1").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80CCD_1  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80CCD_1  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
End If
 
 If Sheet11.Range("DI_EligibleAmount_80CCD_1B").value > 0 And Sheet11.Range("DI_Deduction_80CCD_1B").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80CCD_1B in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80CCD_1B in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
End If
 
 
 If Sheet11.Range("DI_EligibleAmount_80CCD_2").value > 0 And Sheet11.Range("DI_Deduction_80CCD_2").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80CCD_2  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80CCD_2  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
 If Sheet11.Range("DI_EligibleAmount_80D").value > 0 And Sheet11.Range("DI_Deduction_80D").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80D  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80D  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
 If Sheet11.Range("DI_EligibleAmount_80DD").value > 0 And Sheet11.Range("DI_Deduction_80DD").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80DD  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80DD  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If

 If Sheet11.Range("DI_EligibleAmount_80DDB").value > 0 And Sheet11.Range("DI_Deduction_80DDB").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80DDB in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80DDB in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
 
 If Sheet11.Range("DI_EligibleAmount_80E").value > 0 And Sheet11.Range("DI_Deduction_80E").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80E  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80E  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
  If Sheet11.Range("DI_EligibleAmount_80EE").value > 0 And Sheet11.Range("DI_Deduction_80EE").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80EE  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80EE  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
 If Sheet11.Range("DI_EligibleAmount_80EEA").value > 0 And Sheet11.Range("DI_Deduction_80EEA").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80EEA in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80EEA in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
 If Sheet11.Range("DI_EligibleAmount_80EEB").value > 0 And Sheet11.Range("DI_Deduction_80EEB").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80EEB  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80EEB  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
  
 If Sheet11.Range("DI_EligibleAmount_80G").value > 0 And Sheet11.Range("DI_Deduction_80G").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80G  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80G  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If
 
  If Sheet11.Range("DI_EligibleAmount_80GG").value > 0 And Sheet11.Range("DI_Deduction_80GG").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80GG in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80GG in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If

 
  If Sheet11.Range("DI_EligibleAmount_80GGC").value > 0 And Sheet11.Range("DI_Deduction_80GGC").value = "" Then
 'MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2021 to 31.07.2021"" else , enter 0"" for deduction 80GGC  in Sheet DI " & Chr(13) dpk1201
 MsgBox_DI = MsgBox_DI + "* Please enter amount in field ""Deduction attributable to investment/expenditure made between 01.04.2022 to 31.07.2022"" else , enter 0"" for deduction 80GGC  in Sheet DI " & Chr(13)
 ChkDeductionDI = False
 Exit Function
 End If

End If
End Function





