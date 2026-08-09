Attribute VB_Name = "AY25_26_Changes"
Sub OptOutNewTaxregime_25_26()

    If UCase(Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").value) = UCase("by filing 10IEA (having income from business or profession)") Then

      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False

        
        
        Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = False
        Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.115BAC").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        '================================================================
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
  '================================================================
  'For Table B
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
       
       'For Table C
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate

       '================================================================
       
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = ""
        Sheet1.Unprotect Password:=getmsgstate
       
       
  ElseIf UCase(Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").value) = UCase("by exercising the option in the return of income only (form 10IEA is not applicable)") Then
        
        
        
        Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = True
        Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.115BAC").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
      
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '===========================================================
  
    
     'For Table B
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
  
     'For Table C
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
       
       
       ElseIf UCase(Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").value) = UCase("(Select)") Then
       
       
        Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = False
        Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.115BAC").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        
  '================================================================
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
  '================================================================
  
    
     'For Table B
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
  
  
     'For Table C
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
       
End If


End Sub

Sub Taxregime_25_26()

    If UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("Yes") Then
    
      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        
  '================================================================
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
  '================================================================
  
     'For Table B
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
       
    'For Table C
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        NewTaxRegime_New
        
       '================================================================
        
    ElseIf UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("No") Then
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
      
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '===========================================================
  
    
     'For Table B
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
  
     'For Table C
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        NewTaxRegime_New
    
    ElseIf UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("(Select)") Then
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        
  '================================================================
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
  '================================================================
  
    
     'For Table B
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '================================================================
  
  
     'For Table C
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        NewTaxRegime_New
        
    End If
    
End Sub

Sub Taxregime_25_26_A()

If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value) <> UCase("Yes") Then 'Only for NO & Select 'New Tax
        
        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False

        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 2
        SheetALL.Protect Password:=getmsgstate
        
        NewTaxRegime_New
        
        


ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value) = UCase("Yes") Then 'Only for YES 'Old Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False
        
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 2
        SheetALL.Protect Password:=getmsgstate
        
        OldTaxRegime_New

End If
Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub

Sub Taxregime_25_26_B()

If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) <> UCase("No") Then 'Only for Yes & Select 'Old Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False

        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 0
        SheetALL.Protect Password:=getmsgstate
        
        If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) = UCase("Yes") Then
        
            OldTaxRegime_New
            
        ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) = UCase("(Select)") Then
        
           NewTaxRegime_New
        
        End If


ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) = UCase("No") Then 'Only for No 'New Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 1
        SheetALL.Protect Password:=getmsgstate
        
        NewTaxRegime_New

End If

Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub


Sub Taxregime_25_26_C()

If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value) <> UCase("No") Then 'Only for Yes & Select 'Old Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False

        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 2
        SheetALL.Protect Password:=getmsgstate
        
        If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value) = UCase("Yes") Then
            OldTaxRegime_New
        ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value) = UCase("(Select)") Then
           NewTaxRegime_New
        End If


ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value) = UCase("No") Then 'Only for No 'New Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False
        
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
        SheetALL.Unprotect Password:=getmsgstate
        SheetALL.Range("bacValue").value = 1
        SheetALL.Protect Password:=getmsgstate
        
        NewTaxRegime_New

End If

Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub

Sub NewTaxRegime_New()

ThisWorkbook.Unprotect Password:=getmsgstate

      Sheet55.Unprotect Password:=getmsgstate
      Sheet55.Visible = xlSheetHidden
      Sheet55.Protect Password:=getmsgstate
      
      Sheet20.Unprotect Password:=getmsgstate
      Sheet20.Visible = xlSheetHidden
      Sheet20.Protect Password:=getmsgstate
  
      Sheet58.Unprotect Password:=getmsgstate
      Sheet58.Visible = xlSheetHidden
      Sheet58.Protect Password:=getmsgstate
    
      Sheet59.Unprotect Password:=getmsgstate
      Sheet59.Visible = xlSheetHidden
      Sheet59.Protect Password:=getmsgstate
         
      Sheet18.Unprotect Password:=getmsgstate
      Sheet18.Visible = xlSheetHidden
      Sheet18.Protect Password:=getmsgstate
      
      Sheet50.Unprotect Password:=getmsgstate
      Sheet50.Visible = xlSheetHidden
      Sheet50.Protect Password:=getmsgstate
      
      Sheet56.Unprotect Password:=getmsgstate
      Sheet56.Visible = xlSheetHidden
      Sheet56.Protect Password:=getmsgstate
      
      'Added by Aavula Naresh for AY 2025-26
      
      Sheet60.Unprotect Password:=getmsgstate
      Sheet60.Visible = xlSheetHidden
      Sheet60.Protect Password:=getmsgstate
      
      '''''''''''''''''''''''''''''''''''''
      'added by Chetan C M for AY 2025-26 on 27/05/2025
    'start--
    Sheet61.Unprotect Password:=getmsgstate
    Sheet61.Visible = xlSheetHidden
    Sheet61.Protect Password:=getmsgstate
    '--end

      
      Sheet37.Unprotect Password:=getmsgstate
      Sheet37.Visible = xlSheetHidden
      Sheet37.Protect Password:=getmsgstate
      
      Sheet22.Unprotect Password:=getmsgstate
      Sheet22.Visible = xlSheetHidden
      Sheet22.Protect Password:=getmsgstate
      
      SheetALL.Unprotect Password:=getmsgstate
      SheetALL.Range("bacValue").value = 1
      SheetALL.Protect Password:=getmsgstate
      
      Sheet1.Unprotect Password:=getmsgstate
      resetBacYes
      SheetALL.Unprotect Password:=getmsgstate
      
      SheetALL.Protect Password:=getmsgstate
      Sheet1.Protect Password:=getmsgstate
    
      Sheet1.Activate


     ThisWorkbook.Protect Password:=getmsgstate
     
     
     Sheet1.Unprotect Password:=getmsgstate
     SheetALL.Unprotect Password:=getmsgstate
     
     SheetALL.Protect Password:=getmsgstate
     Sheet1.Protect Password:=getmsgstate
    
     Sheet1.Activate
 
    SheetALL.Unprotect Password:=getmsgstate
'    SheetALL.Range("bacValue").Clear
    SheetALL.Protect Password:=getmsgstate
    
'    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
'    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
'     resetBacNo
'    sheetVisibleForNo

'Malli_03/10/2025
SalaryExemptAllowance1013A
'----------------


End Sub

Sub OldTaxRegime_New()

 ThisWorkbook.Unprotect Password:=getmsgstate
 
    sheetVisibleForNo
    
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Range("bacValue").value = 2
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Unprotect Password:=getmsgstate
    
   resetBacNo
   
    Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    SheetALL.Unprotect Password:=getmsgstate
    
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Protect Password:=getmsgstate

    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate

  
    ThisWorkbook.Protect Password:=getmsgstate
    Sheet1.Activate
    Sheet1.Unprotect Password:=getmsgstate
    SheetALL.Unprotect Password:=getmsgstate
    SheetALL.Protect Password:=getmsgstate
    Sheet1.Protect Password:=getmsgstate
    Sheet1.Activate
    
'Malli_03/10/2025
SalaryExemptAllowance1013A
'----------------


End Sub

'Newly Added by Aavula Naresh

Sub Validate_CurrentAY25_26_A23i()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)" Then
              MsgBx_General = MsgBx_General + "*""Do you wish to continue to opt out of New Tax Regime for current assessment year"" is mandatory""" & Chr(13)
              End If
            
             If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No" Then
            
                      If Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = "" Then
'                       MsgBx_General = MsgBx_General + "*""Please enter Date of filing""" & Chr(13)
                       MsgBx_General = MsgBx_General + "*""Please enter the date of filing of Form 10IEA.""" & Chr(13)

                      End If
              
                     If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = "" Then
                        MsgBx_General = MsgBx_General + "*""Acknowledgement Number of  Form 10IEA is mandatory.""" & Chr(13)
                      End If
            End If

       End If

   End If
End Sub

Sub Validate_CurrentAY25_26_A23Bi()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
          If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False Then
            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)" Then
                    MsgBx_General = MsgBx_General + "*""Do you wish to continue to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13)
            End If
            
             If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Then
            
               If Sheet1.Range("BAC115.DateOfFiling_New_No").value = "" Then
'                 MsgBx_General = MsgBx_General + "*""Date of filing of  Form 10-IEA for AY 2025-26 is mandatory""" & Chr(13)
                 MsgBx_General = MsgBx_General + "*""Please enter the date of filing of Form 10IEA.""" & Chr(13)

               End If

              If Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = "" Then
                 MsgBx_General = MsgBx_General + "*""Acknowledgement Number of  Form 10IEA is mandatory.""" & Chr(13)
               End If

          End If
      End If
    End If
End Sub

Sub Validate_CurrentAY25_26_A23Ci()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)" Then
                   MsgBx_General = MsgBx_General + "*""Do you wish to continue to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13)
              End If
            
            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Then
               
                  If Sheet1.Range("BAC115.DateOfFiling_New_NA").value = "" Then
                  
'                  MsgBx_General = MsgBx_General + "*""Date of filing of  Form 10-IEA for AY 2025-26 is mandatory""" & Chr(13)
                  MsgBx_General = MsgBx_General + "*""Please enter the date of filing of Form 10IEA.""" & Chr(13)

  
                  End If
                 
                  If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = "" Then
                  
                   MsgBx_General = MsgBx_General + "*""Acknowledgement Number of  Form 10IEA is mandatory.""" & Chr(13)
  
                  End If

             End If
          
       End If
    End If
End Sub

Sub ID_Opt_Date10IEA_New(ByVal Target As Range)
        If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          If Range("BAC115.DateOfFiling").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling").value)) Or Not CheckDateddmmyyyy212(Trim(Range("BAC115.DateOfFiling").value)) Then
            
            fmsgbox ("*""Please enter date in dd/mm/yyyy format.""")
                
                Sheet1.Range("BAC115.DateOfFiling").value = ""
                
            Else
            
        If Sheet1.Range("Sheet1.115BAC").value = "Yes" Then
    
                    If Dformat(Trim(Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2024"), "yyyy-mm-dd") Then
                             
                             fmsgbox ("*""Date of filing of Form 10IEA cannot be prior to 01/04/2024.""")
                            
                            Sheet1.Range("BAC115.DateOfFiling").value = ""
                            
                    ElseIf Dformat(Trim(Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                        
'                      fmsgbox ("*""Date of filing of Form 10IEA can Not be more than system Date in Income Details""")
                      fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")
'
                            Range("BAC115.DateOfFiling").value = ""
                            
                    End If
                End If
            End If
        End If
    End If
        Sheet1.Protect Password:=getmsgstate

End Sub

Sub ID_Opt_Date10IEA_New_Ai(ByVal Target As Range)
        If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling_New_Yes"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          If Range("BAC115.DateOfFiling_New_Yes").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling_New_Yes").value)) Or Not CheckDateddmmyyyy212(Trim(Range("BAC115.DateOfFiling_New_Yes").value)) Then
            
                 fmsgbox ("*""Please enter date in dd/mm/yyyy format.""")
                 
                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
                
            Else
            
        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No" Then
    
                    If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2025"), "yyyy-mm-dd") Then
                        
                          fmsgbox ("*""Date of filing of Form 10IEA cannot be prior to 01/04/2025.""")
                          
                       Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
                            
                   Else
                   
                   Dim DateOfFiling_chk As Variant
                       
                       If Sheet1.Range("Sheet1.DueDate").value = "31st July" Then
                          DateOfFiling_chk = "31/07/2025"
                       ElseIf Sheet1.Range("Sheet1.DueDate").value = "31st October" Then
                          DateOfFiling_chk = "31/10/2025"
                       ElseIf Sheet1.Range("Sheet1.DueDate").value = "30th November" Then
                          DateOfFiling_chk = "30/11/2025"
                       End If
                       
'                   If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim("31/12/2025"), "yyyy-mm-dd") Then
''                          fmsgbox ("*""Since form 10IEA is filed after due date please select ""Yes"" for field ""Do you wish to continue to opt out of New Tax Regime for current assessment year""")
'                          fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
'                                Range("BAC115.DateOfFiling_New_Yes").value = ""
'
'                    ElseIf Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'
''                          fmsgbox ("*""Date of filing of Form 10IEA can Not be more than system Date in Income Details""")
'                          fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")
'
''
'                            Range("BAC115.DateOfFiling_New_Yes").value = ""
'
'                    End If
                    
                    'Malli_12/09/2025
                    If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                        fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")
                        Range("BAC115.DateOfFiling_New_Yes").value = ""
                    Else:
                                If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim("31/12/2025"), "yyyy-mm-dd") Then
                                    fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                                    Range("BAC115.DateOfFiling_New_Yes").value = ""
                                End If
                    End If
                    '----------------
                    
                End If
            End If
        End If
    End If
End If
        Sheet1.Protect Password:=getmsgstate

End Sub

Sub ID_Opt_Date10IEA_New_Bi(ByVal Target As Range)
        If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling_New_No"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          'If Trim(Range("BAC115.DateOfFiling_New_No").value) <> "00/00/0000" And Range("BAC115.DateOfFiling_New_No").value <> "" Then
        If Range("BAC115.DateOfFiling_New_No").value <> "" Then
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling_New_No").value)) Or Not CheckDateddmmyyyy212(Trim(Range("BAC115.DateOfFiling_New_No").value)) Then
            
                fmsgbox ("*""Please enter date in dd/mm/yyyy format.""")
                
                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
                
            Else
            
        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Then
    
                    If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2025"), "yyyy-mm-dd") Then
                            
                             fmsgbox ("*""Date of filing of Form 10IEA cannot be prior to 01/04/2025.""")
 
                            Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
                            
                   Else
                   
                    Dim DateOfFiling_chk As Variant
                       
                       If Sheet1.Range("Sheet1.DueDate").value = "31st July" Then
                          DateOfFiling_chk = "31/07/2025"
                       ElseIf Sheet1.Range("Sheet1.DueDate").value = "31st October" Then
                          DateOfFiling_chk = "31/10/2025"
                       ElseIf Sheet1.Range("Sheet1.DueDate").value = "30th November" Then
                          DateOfFiling_chk = "30/11/2025"
                       End If
                       
                   'Malli_12/09/2025
        If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                           fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")
                            Range("BAC115.DateOfFiling_New_No").value = ""
                Else:
                
        If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/07/2025" Then  'sadineni 30/10/2025
                                
               If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("NonAudit_DueDate").value, "yyyy-mm-dd") Then
                            fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                            Range("BAC115.DateOfFiling_New_No").value = ""
               End If
         End If


'sadineni 30/10/2025
'=================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
       If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/10/2025" Then
            If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_44AB").value, "yyyy-mm-dd") Then
                            fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                            Range("BAC115.DateOfFiling_New_No").value = ""
               End If
        End If
        
        If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "30/11/2025" Then  'sadineni 30/10/2025
            If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_92E").value, "yyyy-mm-dd") Then
                            fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                            Range("BAC115.DateOfFiling_New_No").value = ""
               End If
        End If
'==========================================================>>>>>>>>>>>>>>>>>
        End If
                End If
            End If
        End If
    End If
End If
        Sheet1.Protect Password:=getmsgstate
End Sub


Sub ID_Opt_Date10IEA_New_Ci(ByVal Target As Range)
        If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling_New_NA"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          'If Trim(Range("BAC115.DateOfFiling_New_NA").value) <> "00/00/0000" And Range("BAC115.DateOfFiling_New_NA").value <> "" Then

        If Range("BAC115.DateOfFiling_New_NA").value <> "" Then
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling_New_NA").value)) Or Not CheckDateddmmyyyy212(Trim(Range("BAC115.DateOfFiling_New_NA").value)) Then

                   fmsgbox ("*""Please enter date in dd/mm/yyyy format.""")

                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""

            Else
        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Then

                    If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2025"), "yyyy-mm-dd") Then

                   fmsgbox ("*""Date of filing of Form 10IEA cannot be prior to 01/04/2025.""")
                            Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""

                   Else
                   
                   Dim DateOfFiling_chk As Variant
                       
                       If Sheet1.Range("Sheet1.DueDate").value = "31st July" Then
                          DateOfFiling_chk = "31/07/2025"
                       ElseIf Sheet1.Range("Sheet1.DueDate").value = "31st October" Then
                          DateOfFiling_chk = "31/10/2025"
                       ElseIf Sheet1.Range("Sheet1.DueDate").value = "30th November" Then
                          DateOfFiling_chk = "30/11/2025"
                       End If

                               
If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                            fmsgbox ("* Date of filing Form 10IEA shall not be beyond system date.")
                            Range("BAC115.DateOfFiling_New_NA").value = ""
                Else:
                
                
        If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/07/2025" Then  '24/10/2025
               If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("NonAudit_DueDate").value, "yyyy-mm-dd") Then
                            fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                            Range("BAC115.DateOfFiling_New_NA").value = ""
               End If
        End If
        
'sadineni 30/10/2025
'=====================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
        If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "31/10/2025" Then
            If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_44AB").value, "yyyy-mm-dd") Then
                            fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                            Range("BAC115.DateOfFiling_New_NA").value = ""
               End If
        End If
        
        If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) = "30/11/2025" Then
            If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Sheet33.Range("AuditcaseDate_92E").value, "yyyy-mm-dd") Then
                            fmsgbox ("* Since form 10IEA is filed after due date please select appropriate option")
                            Range("BAC115.DateOfFiling_New_NA").value = ""
               End If
        End If
'==========================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        End If

                End If
            End If
        End If
    End If
End If
        Sheet1.Protect Password:=getmsgstate
End Sub

Sub Validate_CurrentAY25_26()

    If (Sheet1.Range("Sheet1.115BAC").value = "Yes") Then
        If Sheet1.Range("BAC115.DateOfFiling").value = "" Then
         
'         MsgBx_General = MsgBx_General + "*""Date of filing of Form 10IEA is mandatory.""" & Chr(13)
'          MsgBx_General = MsgBx_General + "* Please enter Date of filing of Form 10-IEA for AY 2024-25 is mandatory" & Chr(13)
          MsgBx_General = MsgBx_General + "* Please enter Date of filing of Form 10-IEA for AY 2024-25" & Chr(13)
        End If
        
        If Sheet1.Range("BAC115.AcknowledgementNo").value = "" Then
            If Sheet1.Range("BAC115.AcknowledgementNo").Locked = False Then
                MsgBx_General = MsgBx_General + "*""Acknowledgement Number of Form 10IEA is mandatory.""" & Chr(13)
            End If
        End If

'        If Not Sheet1.ChkMaxDate_filing_2425(Trim(Sheet1.Range("BAC115.DateOfFiling").value), "31-07-2024") Then

'        If Not Sheet1.ChkMaxDate_filing_2425(Trim(Sheet1.Range("BAC115.DateOfFiling").value), "15-12-2024") Then
'
'
'            MsgBx_General = MsgBx_General + "*""Please select ‘No’, if Form 10IEA was filed after due date for AY 2024-25.""" & Chr(13)
'
'        End If
        
        If Sheet1.Range("BAC115.AcknowledgementNo").value <> "" Then
           If Len(Sheet1.Range("BAC115.AcknowledgementNo").value) > 15 Then
            MsgBx_General = MsgBx_General + "*Acknowledgement Number of Form 10IEA is  cannot exceed 15 characters" & Chr(13)
          End If
          
           If Len(Sheet1.Range("BAC115.AcknowledgementNo").value) < 15 Then
            MsgBx_General = MsgBx_General + "*Acknowledgement Number of Form 10IEA is should 15 characters" & Chr(13)
           End If
        End If

     End If
     
End Sub


Sub Validate_CurrentAY25_26_SetB()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "(Select)" Then
                 MsgBx_General = MsgBx_General + "*""Please Select the Option for Do you wish to opt out of New Tax Regime for current assessment year""" & Chr(13)
              End If
       End If
   End If
End Sub


'ChkMaxDate_filing_2425
Public Function ChkMaxDate_filing_2425(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

    ChkMaxDate_filing_2425 = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))
     
        If Year > 2024 Then
            ChkMaxDate_filing_2425 = False
            Exit Function
        Else
            If Year = 2024 Then
                If month > 8 Then
                    ChkMaxDate_filing_2425 = False
                    Exit Function
                Else
                    If month = 8 Then
                        If dat > 1 Then
                           ChkMaxDate_filing_2425 = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_filing_2425 = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function

Sub Validate_CurrentAY25_26_Method()

     If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
         If Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").MergeArea.Locked = False Then
         
              If Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").value = "" Or Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").value = "(Select)" Then
                 MsgBx_General = MsgBx_General + "*""Please select Method of opting-out of new tax regime for current AY .""" & Chr(13)
              End If
       End If
   End If
End Sub
Sub Validate_CurrentAY24_25_SetA()
  
        If (Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H-HUF") Then
        If Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = False Then
           If Sheet1.Range("Sheet1.115BAC").value = "" Or Sheet1.Range("Sheet1.115BAC").value = "(Select)" Then
              MsgBx_General = MsgBx_General + "*""Please select an option from ""Have you exercised the option u/s 115BAC(6) of Opting out of new tax regime in Form 10-IEA in AY 2024-25?""" & Chr(13)
           End If
        End If
        End If
        
End Sub
Function CheckDateddmmyyyy212(dt As Variant) As Boolean
    CheckDateddmmyyyy212 = True
    
    If InStr(1, dt, ".") > 0 Then
        CheckDateddmmyyyy212 = False
    End If
    
    If Len(dt) > 0 Then
        
        If Mid(dt, 3, 1) <> "/" Then
            If Mid(dt, 3, 1) <> "\" Then
                If Mid(dt, 3, 1) <> "-" Then
                    If Mid(dt, 3, 1) <> "." Then
                        CheckDateddmmyyyy212 = False
                    Else
                        dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                        CheckDateddmmyyyy212 = False
                    End If
                Else
                    dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                    CheckDateddmmyyyy212 = False
                End If
            Else
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                CheckDateddmmyyyy212 = False
            End If
        End If
        
        If Mid(dt, 6, 1) <> "/" Then
            If Mid(dt, 6, 1) <> "-" Then
                If Mid(dt, 6, 1) <> "\" Then
                    If Mid(dt, 6, 1) <> "." Then
                        CheckDateddmmyyyy212 = False
                    Else
                        dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                        CheckDateddmmyyyy212 = False
                    End If
                Else
                    dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                    CheckDateddmmyyyy212 = False
                End If
            Else
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                CheckDateddmmyyyy212 = False
            End If
        End If
        
        If Not IsDate(dt) Then CheckDateddmmyyyy212 = False
        If val(Mid(dt, 1, 2)) < 0 Then CheckDateddmmyyyy212 = False
        If val(Mid(dt, 1, 2)) > 31 Then CheckDateddmmyyyy212 = False
        If val(Mid(dt, 4, 2)) < 0 Then CheckDateddmmyyyy212 = False
        If val(Mid(dt, 4, 2)) > 12 Then CheckDateddmmyyyy212 = False
        If val(Mid(dt, 7, 4)) < 1 Then CheckDateddmmyyyy212 = False
        If val(Mid(dt, 7, 4)) > 3000 Then CheckDateddmmyyyy212 = False
    End If
End Function

Function ValidateEmailNoConsecutiveSpecialChars(email As String) As Boolean
    Dim i As Long
    
    ' Loop through each character in the email
    For i = 1 To Len(email) - 1
        ' Check if the current and next characters are special characters
        If Not (Mid(email, i, 1) Like "[A-Za-z0-9]") And _
           Not (Mid(email, i + 1, 1) Like "[A-Za-z0-9]") Then
            ValidateEmailNoConsecutiveSpecialChars = False
            Exit Function
        End If
    Next i
    
    ' Passed validation if no consecutive special characters found
    ValidateEmailNoConsecutiveSpecialChars = True
End Function

Function checkspecialcharacterForPinCode(field As Variant) As Boolean
    Dim i, j As Long
    checkspecialcharacterForPinCode = True
    Dim arr As Variant
    arr = Array(".", "@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkspecialcharacterForPinCode = False
            Exit Function
        End If
        Next
    Next
End Function

Sub Validate_BAC115DateOfFiling()

If Sheet1.Range("Sheet1.115BAC").Locked = False Then
    If (Sheet1.Range("Sheet1.115BAC").value = "Yes") Then
        If Sheet1.Range("BAC115.DateOfFiling").value <> "" Then
            If Dformat(Trim(Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
               MsgBx_General = MsgBx_General + "*""Date of filing Form 10IEA shall not be beyond system date." & Chr(13)
            End If
        End If
     End If
End If
End Sub


