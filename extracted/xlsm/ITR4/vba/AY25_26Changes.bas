Attribute VB_Name = "AY25_26Changes"
Sub Taxregime_25_26()



If UCase(Left(Sheet1.Range("sheet1.Status").value, 1)) <> UCase("F") Then 'Only for Select,IND,HUF

Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False

    If UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("Yes") Then
    
'Ayush_Regime
'        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Range("BAC115.AssessmentYear_2425").MergeArea.Locked = False
        Sheet1.Range("BAC115.AssessmentYear_2425").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AssessmentYear_2425").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
  '================================================================
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        
        
        Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = False
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
 '       Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        'Ayush_regime
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
'
  '================================================================
  
     'For Table B
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

         Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.A23ii_No").value = "(Select)"
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
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("BAC115.A23b").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23b").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23b").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
 
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        NewTaxRegime_New
        
       '================================================================
  

        
    ElseIf UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("No") Then
 'Ayush_regime
'        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling").value = ""
        Sheet1.Range("BAC115.AssessmentYear_2425").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_2425").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_2425").value = ""
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
        
        Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       
        
        
 'Ayush_regime
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '===========================================================
  
    
     'For Table B
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)"
'        Sheet1.Unprotect Password:=getmsgstate

         Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23ii_No").value = ""
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
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("BAC115.A23b").MergeArea.Locked = False
        Sheet1.Range("BAC115.A23b").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.A23b").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        

        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        NewTaxRegime_New
        
'Ayush_regime
       '================================================================
'    ElseIf UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("Not Applicable") Then
'
''        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
''        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
''        Sheet1.Range("BAC115.DateOfFiling").value = ""
''        Sheet1.Unprotect Password:=getmsgstate
''
''        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
''        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
''        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
''        Sheet1.Unprotect Password:=getmsgstate
'
'
'
'     'For Table A
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'       '===========================================================
'
'
'     'For Table B
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'       '================================================================
'
'     'For Table C
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)"
'        Sheet1.Unprotect Password:=getmsgstate
'
' 'Ayush_regime
''        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = False
''        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HCCFFCC)
''        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
''        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = False
'        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        NewTaxRegime_New
    
    ElseIf UCase(Sheet1.Range("Sheet1.115BAC").value) = UCase("(Select)") Then
        
'        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling").value = ""
'Ayush_Regime
        Sheet1.Range("BAC115.AssessmentYear_2425").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_2425").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_2425").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        
  '================================================================
     'For Table A
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = True
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = ""
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
  '================================================================
  
    
     'For Table B
     'Ayush_regime
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)"
'        Sheet1.Unprotect Password:=getmsgstate
         Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23ii_No").value = ""
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
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)"
'        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.A23b").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23b").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23b").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'Ayush_regime
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        NewTaxRegime_New
        
    End If
    

ElseIf UCase(Left(Sheet1.Range("sheet1.Status").value, 1)) = UCase("F") Then 'Only for FIRM

        
            
     '===========================================================
  Sheet1.Unprotect Password:=getmsgstate
  'Ayush_regime
'        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
    
        Sheet1.Range("BAC115.AssessmentYear_2425").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_2425").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_2425").value = ""
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
        
 'Ayush_regime
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
       '===========================================================
  
    
     'For Table B
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23ii_No").value = ""
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
       'For Table B
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = ""
'        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("BAC115.A23b").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23b").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23b").value = ""
        Sheet1.Unprotect Password:=getmsgstate
          
        Sheet1.Range("BAC115.DateOfFiling_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        
        
        'Malli_SIT-118203 '27/04/2026
        Sheet1.Range("Sec80CCC_Type").Locked = True
        Sheet1.Range("Sec80CCC_Type").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sec80CCC_Type").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sec80CCC_Name").Locked = True
        Sheet1.Range("Sec80CCC_Name").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sec80CCC_Name").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("Sec80CCC_Amount").Locked = True
        Sheet1.Range("Sec80CCC_Amount").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sec80CCC_Amount").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        '----------------
        
        'For FRIM FEW SHEETS BELOW SHEETS HIDE/UNHIDE
        ThisWorkbook.Unprotect Password:=getmsgstate
        Sheet11.Unprotect Password:=getmsgstate
        
        Sheet11.Visible = xlSheetVisible
        Sheet11.Protect Password:=getmsgstate
        
        Sheet6.Unprotect Password:=getmsgstate
        Sheet6.Visible = xlSheetVisible
        Sheet6.Protect Password:=getmsgstate
        
        
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Visible = xlSheetHidden
        Sheet16.Protect Password:=getmsgstate
        
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Visible = xlSheetHidden
        Sheet14.Protect Password:=getmsgstate
        
        'Ayush_15/05
        Sheet20.Unprotect Password:=getmsgstate
        Sheet20.Visible = xlSheetHidden
        Sheet20.Protect Password:=getmsgstate
        
        Sheet15.Unprotect Password:=getmsgstate
        Sheet15.Visible = xlSheetHidden
        Sheet15.Protect Password:=getmsgstate
        
        Sheet19.Unprotect Password:=getmsgstate
        Sheet19.Visible = xlSheetHidden
        Sheet19.Protect Password:=getmsgstate
        
        '----------------------------
        
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").Clear
        Sheet8.Protect Password:=getmsgstate
    
End If
    
Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub
Sub Taxregime_25_26_A()



'If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value) <> UCase("Yes") Then 'Only for NO & Select 'New Tax
If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value) = UCase("Yes") Then
        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False

'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
        Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = False
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.A23ii_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").value = 2
        Sheet8.Protect Password:=getmsgstate
        
        NewTaxRegime_New
        


ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value) = UCase("No") Then 'Only for YES 'Old Tax
'
        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False
'
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""

         Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = True
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
'
        
        Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.A23ii_No").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
'
'        Sheet8.Unprotect Password:=getmsgstate
'        Sheet8.Range("bacValue").value = 2
'        Sheet8.Protect Password:=getmsgstate
'
'        OldTaxRegime_New


ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value) = UCase("(Select)") Then 'Only for YES 'Old Tax
'
        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False
'
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""

         Sheet1.Range("BAC115.AssessmentYear_Aii").MergeArea.Locked = False
        Sheet1.Range("BAC115.AssessmentYear_Aii").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AssessmentYear_Aii").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
         Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = ""
        Sheet1.Unprotect Password:=getmsgstate
'
        
        Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.A23ii_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.A23ii_No").value = "(Select)"
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        
        NewTaxRegime_New

End If
Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub

'Ayush_Regime
Sub Taxregime_26_27_A23ii_b()

    If UCase(Sheet1.Range("BAC115.A23ii_No").value) <> UCase("No") Then
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
            
            Sheet8.Unprotect Password:=getmsgstate
            Sheet8.Range("bacValue").value = 1
            Sheet8.Protect Password:=getmsgstate
        
            NewTaxRegime_New
    Else
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
            
            Sheet8.Unprotect Password:=getmsgstate
            Sheet8.Range("bacValue").value = 2
            Sheet8.Protect Password:=getmsgstate
            
            OldTaxRegime_New
            
    End If


End Sub

'------------

'Modified by sai on 31/01/2025

Sub Taxregime_25_26_B()

If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) <> UCase("No") Then 'Only for Yes & Select 'Old Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False

'        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = False
'        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = False
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").value = 0
        Sheet8.Protect Password:=getmsgstate
        
        If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) = UCase("Yes") Then
        
            OldTaxRegime_New
            
        ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) = UCase("(Select)") Then
        
           NewTaxRegime_New
        
        End If


ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value) = UCase("No") Then 'Only for No 'New Tax

        Sheet1.Unprotect Password:=getmsgstate
        Application.EnableEvents = False
        
'        Sheet1.Range("BAC115.DateOfFiling_New_No").MergeArea.Locked = True
'        Sheet1.Range("BAC115.DateOfFiling_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").MergeArea.Locked = True
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = ""
'        Sheet1.Unprotect Password:=getmsgstate
        
        
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").value = 1
        Sheet8.Protect Password:=getmsgstate
        
        NewTaxRegime_New

End If

Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub
Sub Taxregime_25_26_C()



'If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value) <> UCase("No") Then 'Only for Yes & Select 'Old Tax
If UCase(Sheet1.Range("BAC115.A23b").value) <> UCase("No") Then
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
        
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").value = 2
        Sheet8.Protect Password:=getmsgstate
        
        If UCase(Sheet1.Range("BAC115.A23b").value) = UCase("Yes") Then
            OldTaxRegime_New
            
        ElseIf UCase(Sheet1.Range("BAC115.A23b").value) = UCase("(Select)") Then
           NewTaxRegime_New
        End If

'
'ElseIf UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value) = UCase("No") Then 'Only for No 'New Tax

ElseIf UCase(Sheet1.Range("BAC115.A23b").value) = UCase("No") Then
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
        
        
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").value = 1
        Sheet8.Protect Password:=getmsgstate
        
        NewTaxRegime_New

End If

Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub

Sub Taxregime_25_26_Status()

If (UCase(Left(Sheet1.Range("sheet1.Status").value, 1)) = UCase("F")) Then
         
            Sheet1.Unprotect Password:=getmsgstate
            Application.EnableEvents = False

            Sheet1.Range("Sheet1.115BAC").Validation.Delete
            Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NA"
            Application.EnableEvents = True
            Sheet1.Range("Sheet1.115BAC").value = "NA"
            Application.EnableEvents = False
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = True
      
ElseIf (UCase(Left(Sheet1.Range("sheet1.Status").value, 1)) <> UCase("F")) Then

            Sheet1.Unprotect Password:=getmsgstate
            Application.EnableEvents = False

            Sheet1.Range("Sheet1.115BAC").Validation.Delete
            Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NY"
            Sheet1.Range("Sheet1.115BAC").value = "(Select)"
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HCCFFCC)
            Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = False
            
            NewTaxRegime_New


            Sheet1.Protect Password:=getmsgstate
            Application.EnableEvents = True
End If

 Taxregime_25_26

End Sub
Sub ttt()
Application.EnableEvents = True
End Sub

Sub NewTaxRegime_New()

If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
      Sheet1.Range("Sheet1.115BAC").value = "(Select)"
      fmsgboxsmall "* In case of Firm, New tax regime u/s 115BAC is not applicable."
      Application.EnableEvents = True
      Exit Sub
 End If
 
 'Hiding & Unhiding
  ThisWorkbook.Unprotect Password:=getmsgstate
  Sheet16.Unprotect Password:=getmsgstate
  Sheet16.Visible = xlSheetHidden
  Sheet16.Protect Password:=getmsgstate
  
  'Ayush_13/04
  Sheet19.Unprotect Password:=getmsgstate
  Sheet19.Visible = xlSheetHidden
  Sheet19.Protect Password:=getmsgstate
  
'    Ayush_05/05/2025
'  Sheet15.Unprotect Password:=getmsgstate
'  Sheet15.Visible = xlSheetHidden
'  Sheet15.Protect Password:=getmsgstate
'----------------------------

'Ayush_08/05
  Sheet15.Unprotect Password:=getmsgstate
  Sheet15.Visible = xlSheetHidden
  Sheet15.Protect Password:=getmsgstate
  
   'Ayush_08/05
  Sheet20.Unprotect Password:=getmsgstate
  Sheet20.Visible = xlSheetHidden
  Sheet20.Protect Password:=getmsgstate
  '-------------------------------
  
  Sheet6.Unprotect Password:=getmsgstate
  Sheet6.Visible = xlSheetHidden
  Sheet6.Protect Password:=getmsgstate
  
  'AY 2024-25 Sheet(80GGC) implementation visible/unvisible Starts------------
  Sheet11.Unprotect Password:=getmsgstate
  Sheet11.Visible = xlSheetHidden
  Sheet11.Protect Password:=getmsgstate
  'AY 2024-25 Sheet(80GGC) implementation visible/unvisible End------------
  
  'New Schedule 80DD-80U Added by Sai
  Sheet14.Unprotect Password:=getmsgstate
  Sheet14.Visible = xlSheetHidden
  Sheet14.Protect Password:=getmsgstate
  
  
  
  Sheet8.Unprotect Password:=getmsgstate
  Sheet8.Range("bacValue").value = 1
  Sheet8.Protect Password:=getmsgstate
  Sheet1.Unprotect Password:=getmsgstate
  
  '----
  'Ayush_19
  Sheet1.LockUnlock80E
  Sheet22.LockUnlock_24bHP
  
  resetBacYes
  'Ayush_27/01
 ' Sheet1.LockUnlock24b 'Added by Bindu on 23rd April 25
  
  Sheet1.Activate
  Sheet1.Protect Password:=getmsgstate
  ThisWorkbook.Protect Password:=getmsgstate
 
End Sub
Sub OldTaxRegime_New()

'Lock/Unlock 80D/80D Sheets
Call Worksheets("Income Details").LockUnlock80DU
Call Worksheets("Income Details").LockUnlock80D
'Ayush_13/04
'Call Worksheets("Income Details").LockUnlock80C


'Hiding & Unhiding

        ThisWorkbook.Unprotect Password:=getmsgstate
        
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Visible = xlSheetVisible
        Sheet16.Protect Password:=getmsgstate
        
        'Ayush_13/04
         Sheet19.Unprotect Password:=getmsgstate
        Sheet19.Visible = xlSheetVisible
        Sheet19.Protect Password:=getmsgstate
        
        
        'AY 2024-25 Sheet(80GGC) implementation visible/unvisible Starts------------------
        Sheet11.Unprotect Password:=getmsgstate
        Sheet11.Visible = xlSheetVisible
        Sheet11.Protect Password:=getmsgstate
        'AY 2024-25 Sheet(80GGC) implementation visible/unvisible End------------------
        
        'New Schedule 80DD-80U Added by Sai
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Visible = xlSheetVisible
        Sheet14.Protect Password:=getmsgstate
        
        
        'Ayush_08/05
         If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Visible = xlSheetVisible
            Sheet19.Protect Password:=getmsgstate
        Else
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Visible = xlSheetHidden
            Sheet19.Protect Password:=getmsgstate
        End If
        
        
            'Ayush_08/05
 
        If Mid(Range("sheet1.NatureofEmployment").value, 1, 3) <> "Not" And Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
            Sheet15.Unprotect Password:=getmsgstate
            Sheet15.Visible = xlSheetVisible
            Sheet15.Protect Password:=getmsgstate
        Else
            Sheet15.Unprotect Password:=getmsgstate
            Sheet15.Visible = xlSheetHidden
            Sheet15.Protect Password:=getmsgstate
        End If
        
        
        If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Visible = xlSheetVisible
            Sheet20.Protect Password:=getmsgstate
        Else
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Visible = xlSheetHidden
            Sheet20.Protect Password:=getmsgstate
         End If
        
        
        
        
        Sheet6.Unprotect Password:=getmsgstate
        Sheet6.Visible = xlSheetVisible
        Sheet6.Protect Password:=getmsgstate
        Sheet8.Unprotect Password:=getmsgstate
        Sheet8.Range("bacValue").value = 2
        Sheet8.Protect Password:=getmsgstate
        Sheet1.Unprotect Password:=getmsgstate
        
        '   Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
        '   Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
        
        ' Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HD8D8D8)  ', "AU"
        'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = True
       'Ayush_19
        Sheet1.LockUnlock80E
        Sheet22.LockUnlock_24bHP
        
    
        
        resetBacNo
        'Ayush_27/01
      '  Sheet1.LockUnlock24b ' Updated by Bindu on 23rd April 25
        Sheet1.Protect Password:=getmsgstate
        
        Sheet1.Activate
        ThisWorkbook.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
        
End Sub

Function Validate_CurrentAY25_26_A23Bi() As Boolean 'added by sai on 27/01/2025

Validate_CurrentAY25_26_A23Bi = True


     If (Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H - HUF") Then
'          If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").MergeArea.Locked = False Then
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "(Select)" Then
'             'Ayush_10/02
''                 msgError = msgError & "*""Do you wish to continue to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13) 'by sai on 30/01/2025
'                  msgError = msgError & "*""Do you wish to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13) 'by sai on 30/01/2025
'                  Validate_CurrentAY25_26_A23Bi = False
'            End If

        
            
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "No" Then

'             If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Then 'by sai on 30/01/2025
             
             If Sheet1.Range("BAC115.A23ii_No").MergeArea.Locked = False Then
             
                If Sheet1.Range("BAC115.A23ii_No").value = "" Or Sheet1.Range("BAC115.A23ii_No").value = "(Select)" Then
                   msgError = msgError & "* If answer to A23(A)(ii) is No,  ""Have you furnished form 10IEA for re-entering in new tax regime in current assessment year?"" is mandatory" & Chr(13)
                   Validate_CurrentAY25_26_A23Bi = False
                End If
             
              If Sheet1.Range("BAC115.A23ii_No").value = "Yes" Then
              
            
               If Sheet1.Range("BAC115.DateOfFiling_New_No").value = "" Then

                 msgError = msgError & "*""Date of filing of  Form 10-IEA for AY 2026-27 is mandatory""" & Chr(13) 'by sai on 30/01/2025
                 Validate_CurrentAY25_26_A23Bi = False
               End If

              If Sheet1.Range("BAC115.AcknowledgementNo_New_No").value = "" Then

                 msgError = msgError & "*""Acknowledgement Number of  Form 10IEA is mandatory.""" & Chr(13)
                 Validate_CurrentAY25_26_A23Bi = False
               End If

          End If
      End If
    End If
    

 If Not ID_Opt_Date10IEA_New_Bi Then Validate_CurrentAY25_26_A23Bi = False
    
End Function

Function Validate_CurrentAY25_26_A23Ci() As Boolean 'added by sai on 27/01/2025

Validate_CurrentAY25_26_A23Ci = True



     If (Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H - HUF") Then
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").MergeArea.Locked = False Then
         If Sheet1.Range("BAC115.A23b").MergeArea.Locked = False Then
         
'              If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "(Select)" Then
'                 'Ayush_10/02
''                 msgError = msgError & "*""Do you wish to continue to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13) 'by sai on 30/01/2025
'                  msgError = msgError & "*""Do you wish to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13) 'by sai on 30/01/2025
'                  Validate_CurrentAY25_26_A23Ci = False
'              End If

                 If Sheet1.Range("BAC115.A23b").value = "" Or Sheet1.Range("BAC115.A23b").value = "(Select)" Then
                   msgError = msgError & "* Please select a dropdown from ""If answer to A23 is No, have you furnished  form 10IEA within due date for current assessment year for choosing old tax regime?"" " & Chr(13)
                  Validate_CurrentAY25_26_A23Ci = False
              End If
            
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Then
                If Sheet1.Range("BAC115.A23b").value = "Yes" Then
               
                  If Sheet1.Range("BAC115.DateOfFiling_New_NA").value = "" Then
    
                    msgError = msgError & "*""Date of filing of  Form 10-IEA for AY 2026-27 is mandatory""" & Chr(13) 'by sai on 30/01/2025
                    Validate_CurrentAY25_26_A23Ci = False
                  End If
                 
                 If Sheet1.Range("BAC115.AcknowledgementNo_New_NA").value = "" Then
    
                    msgError = msgError & "*""Acknowledgement Number of  Form 10IEA is mandatory.""" & Chr(13)
                    Validate_CurrentAY25_26_A23Ci = False
                  End If

          End If
          
       End If
    End If
    
   If Not ID_Opt_Date10IEA_New_Ci Then Validate_CurrentAY25_26_A23Ci = False
End Function

'added by sai on 27/01/2025 for AY 2025-26


Function Validate_CurrentAY25_26_A23i() As Boolean 'added by sai on 27/01/2025

Validate_CurrentAY25_26_A23i = True

     If (Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H - HUF") Then
         If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").MergeArea.Locked = False Then
         
'              If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)" Then
'                'Ayush_10/02
'                msgError = msgError & "*""Do you wish to continue to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13) ' by sai on 30/01/2025
'        '        msgError = msgError & "*""Do you wish to opt out of New Tax Regime for current assessment year"" is mandatory" & Chr(13) ' by sai on 30/01/2025
'                Validate_CurrentAY25_26_A23i = False
'              End If

            'Ayush_15/01/2026_V0.2
               If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "(Select)" Then
                
                msgError = msgError & "* Please select an option from the dropdown for ""Have you filed ITR 3/4 in past and have re-entered new tax regime by filing form 10IEA for any assessment year subsequent to assessment year in which first form 10IEA was filed for choosing old tax regime?""" & Chr(13)
                Validate_CurrentAY25_26_A23i = False
              End If
            
            'Ayush_15/01/2026_V0.2
             If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "Yes" Then
            
'                      If Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = "" Then
'
'                        msgError = msgError & "*""Date of filing of  Form 10-IEA for AY 2025-26 is mandatory""" & Chr(13)
'                        Validate_CurrentAY25_26_A23i = False
'                      End If

'                 If Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = "" Then
'
'                    msgError = msgError & "*""Date of filing of  Form 10-IEA for AY 2025-26 is mandatory""" & Chr(13)
'                    Validate_CurrentAY25_26_A23i = False
'                  End If

                 If Sheet1.Range("BAC115.AssessmentYear_Aii").value = "" Or Sheet1.Range("BAC115.AssessmentYear_Aii").value = "(Select)" Then
    
                    msgError = msgError & "* Please select a dropdown from ""Assessment year for which form 10IEA for choosing new tax regime was filed""" & Chr(13)
                    Validate_CurrentAY25_26_A23i = False
                  End If
          
                 If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value = "" Then
    
                    msgError = msgError & "*""Acknowledgement Number of  Form 10IEA is mandatory.""" & Chr(13)
                    Validate_CurrentAY25_26_A23i = False
                  End If
                  
                If Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value <> "" Then
                     If Len(Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value) > 15 Then
                        msgError = msgError & "*Acknowledgement Number of Form 10IEA is  cannot exceed 15 characters" & Chr(13)
                        Validate_CurrentAY25_26_A23i = False
                    End If
          
                    If Len(Sheet1.Range("BAC115.AcknowledgementNo_New_Yes").value) < 15 Then
                      msgError = msgError & "*Acknowledgement Number of Form 10IEA is should 15 characters" & Chr(13)
                      Validate_CurrentAY25_26_A23i = False
                    End If
                End If

     End If
            End If

       End If

   
'Ayush_15/01/2026_V0.2_Regime
'If Not ID_Opt_Date10IEA_New_Ai Then Validate_CurrentAY25_26_A23i = False
   
   
   
End Function


'added by sai on 27/01/2025 for AY 2025-26
Function ID_Opt_Date10IEA_New_Ai() As Boolean

ID_Opt_Date10IEA_New_Ai = True
            Sheet1.Unprotect Password:=getmsgstate
          If Trim(Range("BAC115.DateOfFiling_New_Yes").value) <> "00/00/0000" And Range("BAC115.DateOfFiling_New_Yes").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling_New_Yes").value)) Then
 'Ayush_10/02
'                fmsgboxStatus """Please enter date in dd/mm/yyyy format."""
                msgError = msgError & """Date of filing of Form 10-IEA for AY 2025-26 - Please enter valid date in valid format."""
                ID_Opt_Date10IEA_New_Ai = False
                
                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
                
            Else
        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_Yes").value = "No" Then
    
                    If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2025"), "yyyy-mm-dd") Then
                        
                            msgError = msgError & """Date of filing of Form 10IEA cannot be prior to  01/04/2025."""

                            Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
                            ID_Opt_Date10IEA_New_Ai = False
                            
                   Else
                   'Ayush_09/06
'                   If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim("31/12/2025"), "yyyy-mm-dd") Then
'                      ' If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim("31/07/2025"), "yyyy-mm-dd") Then
''                                msgError = msgError & "*""Since form 10IEA is filed after due date please select ""Yes"" for field ""Do you wish to continue to opt out of New Tax Regime for current assessment year"""
'                                     msgError = msgError & "*""Since form 10IEA is filed after due date please select appropriate option"""
'                                Range("BAC115.DateOfFiling_New_Yes").value = ""   ' added by sai on 27/01/2025 for AY 2025-26
'                                ID_Opt_Date10IEA_New_Ai = False
'      '  Ayush_09/06
'                    ElseIf Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
''                            msgError = msgError & " Date of filing of Form 10IEA can Not be more than system Date in Income Details"
'
'                            msgError = msgError & " Date of filing Form 10IEA shall not be beyond system date."
'                            Range("BAC115.DateOfFiling_New_Yes").value = ""
'                            ID_Opt_Date10IEA_New_Ai = False
'
'                    End If
                    
  '******************************************************
    'sadineni 12_09_2025 Changed
        If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                msgError = msgError & " Date of filing Form 10IEA shall not be beyond system date."
                ID_Opt_Date10IEA_New_Ai = False
                Range("BAC115.DateOfFiling_New_Yes").value = ""
                Else
                                
               If Dformat(Trim(Range("BAC115.DateOfFiling_New_Yes").value), "yyyy-mm-dd") > Dformat(Trim("31/12/2025"), "yyyy-mm-dd") Then
                    msgError = msgError & "*""Since form 10IEA is filed after due date please select appropriate option"""
                    ID_Opt_Date10IEA_New_Ai = False
                    Range("BAC115.DateOfFiling_New_Yes").value = ""
                End If
                    
        End If
   '************************************************************
                End If
            End If
        End If
    End If
    
    'Ayush_28/02
    If Trim(Range("BAC115.DateOfFiling_New_Yes").value) = "00/00/0000" Then
        
            
                msgError = msgError & """Date of filing of Form 10-IEA for AY 2025-26 - Please enter valid date in valid format."""
                ID_Opt_Date10IEA_New_Ai = False
                
                Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
    End If
    

        Sheet1.Protect Password:=getmsgstate
End Function

'added by sai on 27/01/2025 for AY 2025-26

Function ID_Opt_Date10IEA_New_Bi() As Boolean

ID_Opt_Date10IEA_New_Bi = True
        
            Sheet1.Unprotect Password:=getmsgstate
          If Trim(Range("BAC115.DateOfFiling_New_No").value) <> "00/00/0000" And Range("BAC115.DateOfFiling_New_No").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling_New_No").value)) Then
'Ayush_10/02
'                fmsgboxStatus """Please enter date in dd/mm/yyyy format."""
                msgError = msgError & """Date of filing of Form 10-IEA for AY 2026-27 - Please enter valid date in valid format."""
                ID_Opt_Date10IEA_New_Bi = False
                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
                
            Else
'        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "No" Then
'        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Then 'by sai on 30/01/2025
         If Sheet1.Range("BAC115.A23ii_No").value = "Yes" Then
    
                If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2026"), "yyyy-mm-dd") Then
                        
                            msgError = msgError & """Date of filing of Form 10IEA cannot be prior to  01/04/2026."""
                            ID_Opt_Date10IEA_New_Bi = False
                            Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
                            
                   Else
                   'Ayush_10/06
'                       If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Trim("31/07/2025"), "yyyy-mm-dd") Then
                         'If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Trim("15/09/2025"), "yyyy-mm-dd") Then
'                If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Sheet9.Range("DueDate_Extend").value, "yyyy-mm-dd") Then 'Bindu 8th Sep 25
'
'
'                              '  msgError = msgError & "*""Since form 10IEA is filed after due date please select ""No"" for field ""Do you wish to opt out of New Tax Regime for current assessment year"""
'                                 msgError = msgError & "*""Since form 10IEA is filed after due date please select appropriate option"""
'                                ID_Opt_Date10IEA_New_Bi = False
'                                Range("BAC115.DateOfFiling_New_No").value = ""
'
'
'                    ElseIf Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                          '  msgError = msgError & " Date of filing of Form 10IEA can Not be more than system Date in Income Details"
'                             msgError = msgError & " Date of filing Form 10IEA shall not be beyond system date."
'                            ID_Opt_Date10IEA_New_Bi = False
'
'                            Range("BAC115.DateOfFiling_New_No").value = ""
'
'                    End If
                    
'***********************************************************
'sadineni 12_09_2025 Changed
        If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                'msgError = msgError & " Date of filing Form 10IEA shall not be beyond system date."
                    msgError = msgError & """Date of filing Form 10IEA shall not be beyond system date."""
                    ID_Opt_Date10IEA_New_Bi = False
                    Range("BAC115.DateOfFiling_New_No").value = ""
                Else
                                
'               If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Sheet9.Range("DueDate_Extend").value, "yyyy-mm-dd") Then
'                    msgError = msgError & "*""Since form 10IEA is filed after due date please select appropriate option"""
'                    ID_Opt_Date10IEA_New_Bi = False
'                    Range("BAC115.DateOfFiling_New_No").value = ""
'                End If

        'Ayush_15/01/2026_V0.2
                If Dformat(Trim(Range("BAC115.DateOfFiling_New_No").value), "yyyy-mm-dd") > Dformat(Trim("31/12/2026"), "yyyy-mm-dd") Then
                                
                    msgError = msgError & """Date of filing of Form 10IEA filed for re-entering in new tax regime in AY 2026-27 cannot be after 31/12/2026"""
                    ID_Opt_Date10IEA_New_Bi = False
                    Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
                End If
                            
                End If
'*****************************************************************
                End If
            End If
        End If
    End If
    
    'Ayush_28/02
    If Trim(Range("BAC115.DateOfFiling_New_No").value) = "00/00/0000" Then
                msgError = msgError & """Date of filing of Form 10-IEA for AY 2026-27 - Please enter valid date in valid format."""
                ID_Opt_Date10IEA_New_Bi = False
                Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
    End If

        Sheet1.Protect Password:=getmsgstate


End Function

Function ID_Opt_Date10IEA_New_Ci() As Boolean

ID_Opt_Date10IEA_New_Ci = True

        
            Sheet1.Unprotect Password:=getmsgstate
          If Trim(Range("BAC115.DateOfFiling_New_NA").value) <> "00/00/0000" And Range("BAC115.DateOfFiling_New_NA").value <> "" Then

            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling_New_NA").value)) Then
'Ayush_10/02
         '       fmsgboxStatus """Please enter date in dd/mm/yyyy format."""
                msgError = msgError & """Date of filing of Form 10-IEA for AY 2026-27 - Please enter valid date in valid format."""
                ID_Opt_Date10IEA_New_Ci = False
                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""

            Else
'        If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Then
         If Sheet1.Range("BAC115.A23b").value = "Yes" Then

                    If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2026"), "yyyy-mm-dd") Then

                            msgError = msgError & """Date of filing of Form 10IEA for AY 2026-27 cannot be prior to 01/04/2026."""
                            ID_Opt_Date10IEA_New_Ci = False
                            Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""

                   Else
                   'Ayush_10/09
'                       If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Trim("31/07/2025"), "yyyy-mm-dd") Then
                         'If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Trim("15/09/2025"), "yyyy-mm-dd") Then
                         
                         
'            If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Sheet9.Range("DueDate_Extend").value, "yyyy-mm-dd") Then 'Bindu 8th Sep 25
'                              '  msgError = msgError & "*""Since form 10IEA is filed after due date please select ""No"" for field ""Do you wish to opt out of New Tax Regime for current assessment year"""
'                                msgError = msgError & "*""Since form 10IEA is filed after due date please select appropriate option"""
'                                ID_Opt_Date10IEA_New_Ci = False
'                                Range("BAC115.DateOfFiling_New_NA").value = ""
'
'
'                    ElseIf Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                         '   msgError = msgError & " Date of filing of Form 10IEA can Not be more than system Date in Income Details"
'                            msgError = msgError & " Date of filing Form 10IEA shall not be beyond system date."
'                            ID_Opt_Date10IEA_New_Ci = False
'
'                            Range("BAC115.DateOfFiling_New_NA").value = ""
'
'            End If
                    

'***********************************************************
'sadineni 12_09_2025 Changed
        If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                msgError = msgError & " Date of filing Form 10IEA shall not be beyond system date."
                ID_Opt_Date10IEA_New_Ci = False
                    Range("BAC115.DateOfFiling_New_NA").value = ""
                Else
                                
               If Dformat(Trim(Range("BAC115.DateOfFiling_New_NA").value), "yyyy-mm-dd") > Dformat(Sheet9.Range("DueDate_Extend").value, "yyyy-mm-dd") Then
                            'msgError = msgError & "* Since form 10IEA is filed after due date i.e. 31/07/2026 please select appropriate option "
                            'By Bindu on 24th Feb 26
                            msgError = msgError & "* Since form 10IEA is filed after due date i.e. " & Sheet9.Range("DueDate_Extend").value & " please select appropriate option "
                            ID_Opt_Date10IEA_New_Ci = False
                            Range("BAC115.DateOfFiling_New_NA").value = ""
                End If
                    
        End If
'*****************************************************************

                End If
            End If
        End If
    End If
    'Ayush_28/02
    
    If Trim(Range("BAC115.DateOfFiling_New_NA").value) = "00/00/0000" Then

            
                msgError = msgError & """Date of filing of Form 10-IEA for AY 2026-27 - Please enter valid date in valid format."""
                ID_Opt_Date10IEA_New_Ci = False
                Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
    End If

        Sheet1.Protect Password:=getmsgstate


End Function


