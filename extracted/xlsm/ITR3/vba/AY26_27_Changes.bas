Attribute VB_Name = "AY26_27_Changes"
Sub BusiProfAssyr_26_27()
If UCase(Sheet1.Range("Sheet1.BusiProfAssyr").value) = UCase("Yes") Then

      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
      
        Sheet1.Range("Sheet1.Form10IEAoldtax").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Form10IEAoldtax").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.Form10IEAoldtax").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AssYr_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
           
         Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
         Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        

        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAB").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       ElseIf UCase(Sheet1.Range("Sheet1.BusiProfAssyr").value) = UCase("No") Then
       

        Sheet1.Range("Sheet1.Form10IEAoldtax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
        Sheet1.Range("Sheet1.AssYr_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
           
         Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
         Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.Form10IEAB").value = "(Select)"
        
 
  
       ElseIf UCase(Sheet1.Range("Sheet1.BusiProfAssyr").value) = UCase("(Select)") Then

        Sheet1.Range("Sheet1.Form10IEAoldtax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AssYr_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
           
         Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
         Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        Sheet1.Unprotect Password:=getmsgstate

       Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAB").value = ""
         Sheet1.Unprotect Password:=getmsgstate

        
        End If
        

End Sub


Sub Form10IEAoldtax_26_27()

    If UCase(Sheet1.Range("Sheet1.Form10IEAoldtax").value) = UCase("Yes") Then
    
      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
        
        Sheet1.Range("Sheet1.AssYr_IA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AssYr_IA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.AssYr_IA").value = ""
        
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AckNo_IA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.AckNo_IA").value = ""
        
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "(Select)"
        
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
    ElseIf UCase(Sheet1.Range("Sheet1.Form10IEAoldtax").value) = UCase("No") Then
        
        Sheet1.Range("Sheet1.AssYr_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
    
        Sheet1.Range("Sheet1.AckNo_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "(Select)"
        
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Form10IEAB").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        
    
    ElseIf UCase(Sheet1.Range("Sheet1.Form10IEAoldtax").value) = UCase("(Select)") Then
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AssYr_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
               
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        
                Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        
'        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Form10IEAB").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        
    End If
    
End Sub
Sub Form10IEAoldtax_IA_26_27()

    If UCase(Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value) = UCase("Yes") Then
    
      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
        
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        
        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
        NewTaxRegime_New

ElseIf UCase(Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value) = UCase("No") Then
        
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = "(Select)"
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        
        OldTaxRegime_New
        
    ElseIf UCase(Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value) = UCase("(Select)") Then
    
        Sheet1.Range("Sheet1.AssYr_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AssYr_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AssYr_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNo_IIA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IIA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IIA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Form10IEANewTax").value = ""
        Sheet1.Unprotect Password:=getmsgstate
''
''        Sheet1.Unprotect Password:=getmsgstate
''        Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = False
''        Sheet1.Range("Sheet1.Form10IEANewTax").Interior.Color = (&HCCFFCC)
''        Sheet1.Range("Sheet1.Form10IEANewTax").value = "(Select)"
''
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        
    End If
    
End Sub


Sub Form10IEANewtax_26_27()
   If UCase(Sheet1.Range("Sheet1.Form10IEANewTax").value) = UCase("Yes") Then
    
      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
        
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = False
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        
        NewTaxRegime_New

     ElseIf UCase(Sheet1.Range("Sheet1.Form10IEANewTax").value) = UCase("No") Then
        
        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        OldTaxRegime_New
        
     ElseIf UCase(Sheet1.Range("Sheet1.Form10IEANewTax").value) = UCase("(Select)") Then

        Sheet1.Range("Sheet1.DOFNewTax_bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOFNewTax_bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = ""
        Sheet1.Unprotect Password:=getmsgstate

   End If
End Sub
  
Sub Form10IEANewTaxIB_26_27()

   If UCase(Sheet1.Range("Sheet1.Form10IEANewTaxIB").value) = UCase("Yes") Then
    
      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = False
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = False
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.Form10IEAB").value = "(Select)"
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 2
        SheetALL.Protect Password:=getmsgstate
        
        If UCase(Sheet1.Range("Sheet1.Form10IEANewTaxIB").value) = UCase("Yes") Then
            OldTaxRegime_New
        ElseIf UCase(Sheet1.Range("Sheet1.Form10IEANewTaxIB").value) = UCase("(Select)") Then
           NewTaxRegime_New
        End If
        
        

     ElseIf UCase(Sheet1.Range("Sheet1.Form10IEANewTaxIB").value) = UCase("No") Then
        
        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Form10IEAB").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        
        
        NewTaxRegime_New
        

     ElseIf UCase(Sheet1.Range("Sheet1.Form10IEANewTaxIB").value) = UCase("(Select)") Then

        Sheet1.Range("Sheet1.DOF_Ibi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.DOF_Ibi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.DOF_Ibi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("Sheet1.AckNo_IBi").MergeArea.Locked = True
        Sheet1.Range("Sheet1.AckNo_IBi").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.AckNo_IBi").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.Form10IEAB").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Form10IEAB").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

   End If

End Sub
Sub DateOfFilingform_bi(ByVal Target As Range)

        If Not Application.Intersect(Sheet1.Range("Sheet1.DOFNewTax_bi"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          If Range("Sheet1.DOFNewTax_bi").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("Sheet1.DOFNewTax_bi").value)) Or Not CheckDateddmmyyyy212(Trim(Range("Sheet1.DOFNewTax_bi").value)) Then
            
            fmsgbox ("*""Please enter correct date in dd/mm/yyyy format for Form 10IEA for AY 2026-27.""")
                
                Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
                
            Else
            
        If Sheet1.Range("Sheet1.Form10IEANewTax").value = "Yes" Then
    
                    If Dformat(Trim(Range("Sheet1.DOFNewTax_bi").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2026"), "yyyy-mm-dd") Then
                             
                             fmsgbox ("*""Date of filing of Form 10IEA  for AY 2026-27 cannot be prior to  01/04/2026.""")
                            
                            Sheet1.Range("Sheet1.DOFNewTax_bi").value = ""
                            
                    ElseIf Dformat(Trim(Range("Sheet1.DOFNewTax_bi").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                        
                      fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")

                            Range("Sheet1.DOFNewTax_bi").value = ""
                            
                               Else
                           
'                   If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/07/2026" Then
                   If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/08/2026" Then
                   
                    If Dformat(Trim(Range("Sheet1.DOFNewTax_bi").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("NonAudit_DueDate").value, "yyyy-mm-dd") Then
                             fmsgbox ("* ""Since form 10IEA for AY 2026-27 is filed after due date please select appropriate option""")
                            Range("Sheet1.DOFNewTax_bi").value = ""
                    End If
               End If
               
                If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/10/2026" Then
                                
                    If Dformat(Trim(Range("Sheet1.DOFNewTax_bi").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_44AB").value, "yyyy-mm-dd") Then
                             fmsgbox ("* ""Since form 10IEA for AY 2026-27 is filed after due date please select appropriate option""")
                            Range("Sheet1.DOFNewTax_bi").value = ""
                    End If
               End If
               
                If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "30/11/2026" Then
                                
                    If Dformat(Trim(Range("Sheet1.DOFNewTax_bi").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_92E").value, "yyyy-mm-dd") Then
                             fmsgbox ("* ""Since form 10IEA for AY 2026-27 is filed after due date please select appropriate option""")
                            Range("Sheet1.DOFNewTax_bi").value = ""
                    End If
               End If
               
                            
                    End If
                End If
            End If
        End If
    End If
        Sheet1.Protect Password:=getmsgstate

End Sub

Sub DateOfFilingform_Ibi(ByVal Target As Range)
        If Not Application.Intersect(Sheet1.Range("Sheet1.DOF_Ibi"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          If Range("Sheet1.DOF_Ibi").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("Sheet1.DOF_Ibi").value)) Or Not CheckDateddmmyyyy212(Trim(Range("Sheet1.DOF_Ibi").value)) Then
            
            fmsgbox ("*""Please enter correct date in dd/mm/yyyy format for Form 10IEA for AY 2026-27.""")
                
                Sheet1.Range("Sheet1.DOF_Ibi").value = ""
                
            Else
            
        If Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "Yes" Then
    
                    If Dformat(Trim(Range("Sheet1.DOF_Ibi").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2026"), "yyyy-mm-dd") Then
                             
                             fmsgbox ("*""Date of filing of Form 10IEA  for AY 2026-27 cannot be prior to  01/04/2026.""")
                            
                            Sheet1.Range("Sheet1.DOF_Ibi").value = ""
                            
                    ElseIf Dformat(Trim(Range("Sheet1.DOF_Ibi").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                        
                      fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")

                            Range("Sheet1.DOF_Ibi").value = ""
                            
                        Else
                           
'                   If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/07/2026" Then
                   If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/08/2026" Then
                                
                    If Dformat(Trim(Range("Sheet1.DOF_Ibi").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("NonAudit_DueDate").value, "yyyy-mm-dd") Then
                             fmsgbox ("* ""Since form 10IEA for AY 2026-27 is filed after due date, old tax regime is not applicable to you. Please recheck the due date, date of filing of Form 10IEA and the option chosen to avoid any adjustment by the department""")
                            Range("Sheet1.DOF_Ibi").value = ""
                    End If
               End If
               
                If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/10/2026" Then
                                
                    If Dformat(Trim(Range("Sheet1.DOF_Ibi").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_44AB").value, "yyyy-mm-dd") Then
                             fmsgbox ("* ""Since form 10IEA for AY 2026-27 is filed after due date, old tax regime is not applicable to you. Please recheck the due date, date of filing of Form 10IEA and the option chosen to avoid any adjustment by the department""")
                            Range("Sheet1.DOF_Ibi").value = ""
                    End If
               End If
               
                If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "30/11/2026" Then
                                
                    If Dformat(Trim(Range("Sheet1.DOF_Ibi").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_92E").value, "yyyy-mm-dd") Then
                             fmsgbox ("* ""Since form 10IEA for AY 2026-27 is filed after due date, old tax regime is not applicable to you. Please recheck the due date, date of filing of Form 10IEA and the option chosen to avoid any adjustment by the department""")
                            Range("Sheet1.DOF_Ibi").value = ""
                    End If
               End If
                            
                            
                    End If
                End If
            End If
        End If
    End If
        Sheet1.Protect Password:=getmsgstate

End Sub
Sub Validate_BusiProfAssyr_2026_27()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.BusiProfAssyr").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.BusiProfAssyr").value = "" Or Sheet1.Range("Sheet1.BusiProfAssyr").value = "(Select)" Then
                 MsgBx_General = MsgBx_General + "*""Please select correct option for 'Do you have income from business or profession for current AY'.""" & Chr(13)
              End If
       End If
   End If
End Sub

Sub Validate_Form10IEAB_2026_27()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.Form10IEAB").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.Form10IEAB").value = "" Or Sheet1.Range("Sheet1.Form10IEAB").value = "(Select)" Then
                 MsgBx_General = MsgBx_General + "*""Please select option if you wish to opt for old tax regime for the current Assessment Year.""" & Chr(13)
              End If
       End If
   End If
End Sub

Sub Validate_Form10IEAoldtax_I()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.Form10IEAoldtax").MergeArea.Locked = False Then
         
            If Sheet1.Range("Sheet1.Form10IEAoldtax").value = "" Or Sheet1.Range("Sheet1.Form10IEAoldtax").value = "(Select)" Then
             MsgBx_General = MsgBx_General + "*""Please select option for have you filed form10IEA within due date for any earlier assessment year""" & Chr(13)
            End If
            
                 If Sheet1.Range("Sheet1.Form10IEAoldtax").value = "Yes" Then
                    If Sheet1.Range("Sheet1.AssYr_IA").value = "" Then
                       MsgBx_General = MsgBx_General + "*""Please enter Assessment year for which Form 10IEA for choosing old tax regime was filed.""" & Chr(13)
                    End If
            
                    If Sheet1.Range("Sheet1.AckNo_IA").value = "" Then
'                       MsgBx_General = MsgBx_General + "*""Please enter Form 10IEA acknowledgement number for which this form for choosing old tax regime was filed.""" & Chr(13)
                        MsgBx_General = MsgBx_General + "*""Please enter acknowledgement number of Form 10IEA filed for choosing old tax regime.""" & Chr(13)
                    End If
                    If Sheet1.Range("Sheet1.AckNo_IA").value <> "" Then
                       If Len(Sheet1.Range("Sheet1.AckNo_IA").value) > 15 Then
'                          MsgBx_General = MsgBx_General + "*""Please enter the correct Acknowledgement Number for which this form for choosing old tax regime was filed.""" & Chr(13)
                           MsgBx_General = MsgBx_General + "*""Please enter the correct Acknowledgement Number of Form 10IEA filed for choosing old tax regime.""" & Chr(13)
                       End If
                       
                       If Len(Sheet1.Range("Sheet1.AckNo_IA").value) < 15 Then
'                          MsgBx_General = MsgBx_General + "*""Please enter the correct Acknowledgement Number for which this form for choosing old tax regime was filed.""" & Chr(13)
                           MsgBx_General = MsgBx_General + "*""Please enter the correct Acknowledgement Number of Form 10IEA filed for choosing old tax regime.""" & Chr(13)
                       End If
                    End If
                 End If
         End If
     End If
End Sub
Sub Validate_Form10IEAoldtax_IA()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "" Or Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "(Select)" Then
              MsgBx_General = MsgBx_General + "*""Please select If you have re-entered new tax regime by filing form 10IEA for any assessment year subsequent to assessment year in which first form 10IEA was filed for choosing old tax regime?""" & Chr(13)
              End If
            
             If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "Yes" Then
            
                      If Sheet1.Range("Sheet1.AssYr_IIA").value = "" Then
                       MsgBx_General = MsgBx_General + "*""Please select Assessment Year in which you have re-entered into the New Tax Regime""" & Chr(13)

                      End If
              
                     If Sheet1.Range("Sheet1.AckNo_IIA").value = "" Then
                        MsgBx_General = MsgBx_General + "*""Please fill the Acknowledgement Number at Sl. No. A19(b)(I)(A)(ii)(a) of Part A General""" & Chr(13)
                      End If
                      
                     If Sheet1.Range("Sheet1.AckNo_IIA").value <> "" Then
                           If Len(Sheet1.Range("Sheet1.AckNo_IIA").value) > 15 Then
                              MsgBx_General = MsgBx_General + "*""Please enter the correct Acknowledgement Number at Sl. No. A19(b)(I)(A)(ii)(a) of Part A General""" & Chr(13)
                           End If
                           
                           If Len(Sheet1.Range("Sheet1.AckNo_IIA").value) < 15 Then
                              MsgBx_General = MsgBx_General + "*""Please enter the correct Acknowledgement Number at Sl. No. A19(b)(I)(A)(ii)(a) of Part A General""" & Chr(13)
                           End If
                     End If
            End If
       End If

   End If
End Sub

Sub Validate_Form10IEANewTax()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.Form10IEANewTax").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.Form10IEANewTax").value = "" Or Sheet1.Range("Sheet1.Form10IEANewTax").value = "(Select)" Then
              MsgBx_General = MsgBx_General + "*""Please select whether you have furnished Form 10IEA for re-entering in New Tax Regime in current AY""" & Chr(13)
              End If
            
             If Sheet1.Range("Sheet1.Form10IEANewTax").value = "Yes" Then
            
                      If Sheet1.Range("Sheet1.DOFNewTax_bi").value = "" Then
                       MsgBx_General = MsgBx_General + "*""Please enter the date of filing of Form 10IEA for AY 2026-27.""" & Chr(13)

                      End If
              
                     If Sheet1.Range("Sheet1.AckNoNewTax_Bi").value = "" Then
                        MsgBx_General = MsgBx_General + "*""Please enter Form 10IEA acknowledgement number for which this form for re-entering into New tax regime was filed.""" & Chr(13)
                      End If
            End If
       End If
   End If
End Sub
Sub Validate_Form10IEANewTaxIB()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.Form10IEANewTaxIB").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "" Or Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "(Select)" Then
              MsgBx_General = MsgBx_General + "*""Please select option if you have furnished Form 10IEA for choosing Old Tax Regime in current AY.""" & Chr(13)
              End If
            
             If Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "Yes" Then
            
                      If Sheet1.Range("Sheet1.DOF_Ibi").value = "" Then
                       MsgBx_General = MsgBx_General + "*""Please enter the date of filing of Form 10IEA for AY 2026-27.""" & Chr(13)

                      End If
              
                     If Sheet1.Range("Sheet1.AckNo_IBi").value = "" Then
                        MsgBx_General = MsgBx_General + "*""Please enter Form 10IEA acknowledgement number for AY 2026-27.""" & Chr(13)
                      End If
            End If
       End If

   End If
End Sub


Sub validate_FPI()

         If (Sheet1.Range("Foreign_institutinal_investment").value = "(Select)" Or Sheet1.Range("Foreign_institutinal_investment").value = "") Then
            If Sheet1.Range("Foreign_institutinal_investment").Locked = False Then
              MsgBx_General = MsgBx_General & """Please select whether you are an FPI""" & Chr(13)
            End If
     End If
End Sub
