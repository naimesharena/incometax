Attribute VB_Name = "Sheet13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String
Dim count As Variant


Private Sub Worksheet_Activate()
On Error Resume Next


fmsgbox ("If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F")

'Newly added by sai on 28/03/2025
 AddTotalSTCGA1
 AddTotalLTCGB1
 Cost_of_Improvement


'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet13.Shapes("Button 13").Left = Range("Button_CG").Left + 2
Sheet13.Shapes("Button 13").Top = Range("Button_CG").Top + 2

Sheet13.Shapes("Button 87").Left = Range("Button_STCG").Left + 6
Sheet13.Shapes("Button 87").Top = Range("Button_STCG").Top + 6

Sheet13.Shapes("Button 89").Left = Range("Button_STCG_A8_DTAA").Left + 6
Sheet13.Shapes("Button 89").Top = Range("Button_STCG_A8_DTAA").Top + 6

Sheet13.Shapes("Button 159").Left = Range("Button_LTCG").Left + 6
Sheet13.Shapes("Button 159").Top = Range("Button_LTCG").Top + 6

Sheet13.Shapes("Button 160").Left = Range("BUTTON_160_LTCG10").Left + 6
Sheet13.Shapes("Button 160").Top = Range("BUTTON_160_LTCG10").Top + 6

    Sheet13.Shapes("Button 270").Left = Range("Button_Ded54D").Left + 3
    Sheet13.Shapes("Button 270").Top = Range("Button_Ded54D").Top + 3

    Sheet13.Shapes("Button 267").Left = Range("Button_Ded54").Left + 3
    Sheet13.Shapes("Button 267").Top = Range("Button_Ded54").Top + 3

    Sheet13.Shapes("Button 269").Left = Range("Button_Ded54B").Left + 3
    Sheet13.Shapes("Button 269").Top = Range("Button_Ded54B").Top + 3

    Sheet13.Shapes("Button 271").Left = Range("Button_Ded54EC").Left + 3
    Sheet13.Shapes("Button 271").Top = Range("Button_Ded54EC").Top + 3

    Sheet13.Shapes("Button 272").Left = Range("Button_Ded54F").Left + 3
    Sheet13.Shapes("Button 272").Top = Range("Button_Ded54F").Top + 3

    Sheet13.Shapes("Button 273").Left = Range("Button_Ded54G").Left + 3
    Sheet13.Shapes("Button 273").Top = Range("Button_Ded54G").Top + 3


    Sheet13.Shapes("Button 274").Left = Range("Button_Ded54GA").Left + 3
    Sheet13.Shapes("Button 274").Top = Range("Button_Ded54GA").Top + 3

    Sheet13.Shapes("Button 275").Left = Range("Button_Ded115F").Left + 3
    Sheet13.Shapes("Button 275").Top = Range("Button_Ded115F").Top + 3


    Sheet13.Shapes("Button 402").Left = Range("Button_CGA1").Left + 3
    Sheet13.Shapes("Button 402").Top = Range("Button_CGA1").Top + 3

    Sheet13.Shapes("Button 403").Left = Range("Button_CGA1").Left + 150
    Sheet13.Shapes("Button 403").Top = Range("Button_CGA1").Top + 3

    Sheet13.Shapes("Button 405").Left = Range("Button_CGB1").Left + 3
    Sheet13.Shapes("Button 405").Top = Range("Button_CGB1").Top + 3

    Sheet13.Shapes("Button 406").Left = Range("Button_CGB1").Left + 150
    Sheet13.Shapes("Button 406").Top = Range("Button_CGB1").Top + 3

endline:
Sheet13.Protect Password:=sPassword
Application.EnableEvents = True
End Sub
Sub checkfuturedate()

End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim Rangename, i As Variant
Application.EnableEvents = False


'--Newly added by Bindu for Table E Editable
If Not Application.Intersect(Sheet13.Range("CG_TABLE_E_Editable"), Sheet13.Range(Target.Address)) Is Nothing Then

   If UCase(Sheet13.Range("CG_TableE_Checkbox").value) = UCase("True") Then
    
        If Sheet13.Range("TableE_Negativecount").value > 0 Then
              MsgBox ("**Amount of set off claimed cannot exceed amount available for set off")
              Application.EnableEvents = False

              Target.value = 0
              Application.EnableEvents = True
        GoTo endline
         End If
    End If
    
End If
    
    
'LockUnlock112A9
'LockUnlock112A11
'ChangeCG-22.12.2022.102.03A

'For AY 2025-26 by sai on 24/03/2025**********************************************
Dim c As Long
For c = 1 To Sheet13.Range("PropertySectionCount_STCG").value
   If Not Sheet13.Range("STCG.AB_SaleTrfDate" & c) = "" Then
       If Not Application.Intersect(Sheet13.Range("STCG.AB_SaleTrfDate" & c), Sheet13.Range(Target.Address)) Is Nothing Then
         If Dformat(Sheet13.Range("STCG.AB_SaleTrfDate" & c).value, "22/07/2024") < Dformat(Sheet13.Range("STCG.A_AcqDate" & c).value, "22/07/2024") Then
'           MsgBox ("Date of sale can't be earlier to Date of purchase / acquisition in Short term capital gain in schedule CG at S.no " & C)
       'SIT-91152
            MsgBox ("""Date of sale can't be earlier to Date of purchase / acquisition in sl.no.A1 of STCG at S.no " & c)
           Sheet13.Range("STCG.AB_SaleTrfDate" & c).value = ""
         End If
       End If
   End If
Next
'**********************************************************************************


'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.54.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year Change by sai on 19/02/2025
        'Year changes by Bindu on 23rd Feb 26
        'AY updated by Bindu  on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1a In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next


'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54.DateofDeposit").count
     If Not Intersect(Range("Deduction.54.DateofDeposit"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format at D1a In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
    
    'SIT-91423 Sadineni_13_09_2025
      '****************************************
        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("""Date cannot be after system date""")
            Target.value = ""
        End If
       '*****************************************
    
      End If
Next

'Newly added by Bindu DE 0.2

For i = 1 To Sheet13.Range("Deduction.54.IFSC").count
   If Not Application.Intersect(Range("Deduction.54.IFSC"), Range(Target.Address)) Is Nothing Then
    
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then

                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
          
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
  
Next



'For SIT-94390,SIT-94389 by sai on 13/06/2025*************************************************************
'--Below fields will alyws be Zero by Bindu 16th Feb 26
If Not Application.Intersect(Sheet13.Range("LTCG.B3bii_ImproveCost"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation"), Range(Target.Address)) Is Nothing Then
  If Sheet13.Range("LTCG.B3bii_ImproveCost").value < Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation").value Then
        fmsgboxStatus "4bii cant be less than 4biia"
        Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation").value = ""
  End If
End If

'--Below fields will alyws be Zero by Bindu 16th Feb 26
If Not Application.Intersect(Sheet13.Range("LTCG.B3bi_AquisitCost"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet13.Range("LTCG.B3bia_AcqIndexation"), Range(Target.Address)) Is Nothing Then
  If Sheet13.Range("LTCG.B3bi_AquisitCost").value < Sheet13.Range("LTCG.B3bia_AcqIndexation").value Then
        fmsgboxStatus "4bi cant be less than 4bia"
        Sheet13.Range("LTCG.B3bia_AcqIndexation").value = ""
  End If
End If

'**********************************************************************************************************
'Newly modified by sai on 28/02/2025 for AY 2025-26
'Not updating anything as this field "LTCG.B04d_ExemptionGrandTotal" will always be zero by Bindu on 17th Feb 26
If Not Intersect(Range("LTCG.B04d_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.B04d_ExemptionGrandTotal_ii"), Range(Target.Address)) Is Nothing Then
   Dim B4id As Variant
        B4id = Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").value + Sheet13.Range("LTCG.B04d_ExemptionGrandTotal_ii").value
    If B4id > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If
'********************************************************************************************************

'Not updating any this this field "LTCG.B04di_ExemptionGrandTotal" will alywas be zero By Bindu on 18th Feb 26
If Not Intersect(Range("LTCG.B04di_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.B04di_ExemptionGrandTotal_ii"), Range(Target.Address)) Is Nothing Then
   Dim B4iid As Variant
   B4iid = Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value + Sheet13.Range("LTCG.B04di_ExemptionGrandTotal_ii").value
    If B4iid > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If


'For 5bi and 5bii

      'Not updating any this this field "B5_i_DeductionTotal" will alywas be zero By Bindu on 18th Feb 26
If Not Intersect(Range("B5_i_DeductionTotal"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("B5_ii_DeductionTotal"), Range(Target.Address)) Is Nothing Then
   Dim B5i As Variant
   B5i = Sheet13.Range("B5_i_DeductionTotal").value + Sheet13.Range("B5_ii_DeductionTotal").value
    If B5i > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If
      
'*******************************************************

'For 7d before and after 112(1)(C)

'Not updating anything as this field will always be zero "LTCG.B07dii_ExemptionGrandTotal"
If Not Intersect(Range("LTCG.B07dii_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.B07dii_ExemptionGrandTotal_ii"), Range(Target.Address)) Is Nothing Then
   Dim B7di As Variant
   B7di = Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").value + Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal_ii").value
    If B7di > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If


'For 7dii before and after 115 (C)
     'Not updating anything as this field "LTCG.B07div_ExemptionGrandTotal" will alywas be zero by Bindu on 18th Feb 26
If Not Intersect(Range("LTCG.B07div_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.B07div_ExemptionGrandTotal_ii"), Range(Target.Address)) Is Nothing Then
   Dim B7dii As Variant
   B7dii = Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value + Sheet13.Range("LTCG.B07div_ExemptionGrandTotal_ii").value
    If B7dii > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If

'For 7 115AD

      If Not Intersect(Range("LTCG.B07dv_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
         If Target.value > 100000000 Or Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").value > 100000000 Then
                fmsgbox ("Amount against section 54F can not be more than 10 crores. ")
                Target.value = ""
                GoTo endline
         End If
      End If

'For 8bi and 8bii

'Not updating anything as this field "LTCG.8_i.DeductionUnder54F" will always be zero by Bindu on 18th Fen 26
If Not Intersect(Range("LTCG.8_i.DeductionUnder54F"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.8_ii.DeductionUnder54F"), Range(Target.Address)) Is Nothing Then
   Dim B8bi As Variant
   B8bi = Sheet13.Range("LTCG.8_i.DeductionUnder54F").value + Sheet13.Range("LTCG.8_ii.DeductionUnder54F").value
    If B8bi > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If

'*************************************************************************************************************


''Newly added by Bindu as per DE 0.2
'For i = 1 To Sheet13.Range("Deduction.54.AmtDedClaimed").count
'     If Not Intersect(Range("Deduction.54.AmtDedClaimed"), Range(Target.Address)) Is Nothing Then
'
'        If Target.value > 100000000 Then
'                fmsgbox ("Please ensure Amount of deduction in Section 54 cannot exceed Rs.10 crore ")
'                Target.value = ""
'                GoTo endline
'        End If
'      End If
'Next
'
''Newly added by Bindu as per DE 0.2
'For i = 1 To Sheet13.Range("Deduction.54F.AmtDedClaimed").count
'     If Not Intersect(Range("Deduction.54F.AmtDedClaimed"), Range(Target.Address)) Is Nothing Then
'
'        If Target.value > 100000000 Then
'                fmsgbox ("Please ensure Amount of deduction in Section 54F cannot exceed Rs.10 crore ")
'                Target.value = ""
'                GoTo endline
'        End If
'      End If
'Next

'JYOTI SIT-78223
'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54.AmtDedClaimed").count
     If Not Intersect(Range("Deduction.54.AmtDedClaimed"), Range(Target.Address)) Is Nothing Then
        
        'If Target.value > 100000000 Then
        If Target.value > 100000000 Or Sheet13.Range("Deduction.54.Total").value > 100000000 Then
                fmsgbox ("Please ensure Amount of deduction in Section 54F cannot exceed Rs.10 crore ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54F.AmtDedClaimed").count
     If Not Intersect(Range("Deduction.54F.AmtDedClaimed"), Range(Target.Address)) Is Nothing Then
        'If Target.value > 100000000 then
        If Target.value > 100000000 Or Sheet13.Range("Deduction.54F.Total").value > 100000000 Then
                fmsgbox ("Please ensure Amount of deduction in Section 54F cannot exceed Rs.10 crore ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next



'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.54B.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54B.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        
        'Year change from 31/03/2025 to 31/03/2026 by Bindu on 23rd Feb 26
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year Change by sai on 19/02/2025
        ' 'Year change from 31/03/2025 to 31/03/2026  & also AY change by Bindu on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1b In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54B.DateofDeposit").count
     If Not Intersect(Range("Deduction.54B.DateofDeposit"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format at D1b In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
        
    'SIT-91423 Sadineni_13_09_2025
      '****************************************
        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("""Date cannot be after system date""")
            Target.value = ""
            GoTo endline
        End If
       '*****************************************
        
      End If
Next

'Newly added by Bindu DE 0.2

For i = 1 To Sheet13.Range("Deduction.54B.IFSC").count
   If Not Application.Intersect(Range("Deduction.54B.IFSC"), Range(Target.Address)) Is Nothing Then
    
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then

                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
          
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
  
Next

'Newly added by Bindu as per DE 0.2
'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54D.DateofDeposit").count
     If Not Intersect(Range("Deduction.54D.DateofDeposit"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
        
        
    'SIT-91423 Sadineni_13_09_2025
      '****************************************
        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("""Date cannot be after system date""")
            Target.value = ""
            GoTo endline
        End If
       '*****************************************
        
      End If
Next

'Newly added by Bindu DE 0.2

For i = 1 To Sheet13.Range("Deduction.54D.IFSC").count
   If Not Application.Intersect(Range("Deduction.54D.IFSC"), Range(Target.Address)) Is Nothing Then
    
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then

                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
          
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
  
Next


'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54F.DateofDeposit").count
     If Not Intersect(Range("Deduction.54F.DateofDeposit"), Range(Target.Address)) Is Nothing Then
     
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format at D1e In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
        
'SIT-91423 Sadineni_13_09_2025
      '****************************************
        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("""Date cannot be after system date""")
            Target.value = ""
            GoTo endline
        End If
       '*****************************************
        
        
        
      End If
Next
 
'Newly added by Bindu DE 0.2
 
For i = 1 To Sheet13.Range("Deduction.54F.IFSC").count
   If Not Application.Intersect(Range("Deduction.54F.IFSC"), Range(Target.Address)) Is Nothing Then
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then
 
                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
Next


'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54G.DateofDeposit").count
     If Not Intersect(Range("Deduction.54G.DateofDeposit"), Range(Target.Address)) Is Nothing Then
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format at D1f In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
        
'SIT-91423 Sadineni_13_09_2025
      '****************************************
        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("""Date cannot be after system date""")
            Target.value = ""
            GoTo endline
        End If
       '*****************************************

      End If
Next
 
'Newly added by Bindu DE 0.2
 
For i = 1 To Sheet13.Range("Deduction.54G.IFSC").count
   If Not Application.Intersect(Range("Deduction.54G.IFSC"), Range(Target.Address)) Is Nothing Then
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then
 
                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
Next


'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54GA.DateofDeposit").count
     If Not Intersect(Range("Deduction.54GA.DateofDeposit"), Range(Target.Address)) Is Nothing Then
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format at D1g In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
        
'SIT-91423 Sadineni_13_09_2025
      '****************************************
        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("""Date cannot be after system date""")
            Target.value = ""
            GoTo endline
        End If
       '*****************************************

      End If
Next
 
'Newly added by Bindu DE 0.2
 
For i = 1 To Sheet13.Range("Deduction.54GA.IFSC").count
   If Not Application.Intersect(Range("Deduction.54GA.IFSC"), Range(Target.Address)) Is Nothing Then
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then
 
                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
Next


'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.54EC.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54EC.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        ' 'Year change from 31/03/2025 to 31/03/2026   by Bindu on 23rd Feb 26
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year Change by sai on 19/02/2025
        '' 'Year change from 31/03/2025 to 31/03/2026  & also AY change by Bindu on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1d In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.54F.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54F.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        '' 'Year change from 31/03/2025 to 31/03/2026  by Bindu on 23rd Feb 26
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year Change by sai on 19/02/2025
' 'Year change from 31/03/2025 to 31/03/2026  & also AY change by Bindu on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1e In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Newly added by Bindu as per DE 0.2
For i = 1 To Sheet13.Range("Deduction.54F.DateofDeposit").count
     If Not Intersect(Range("Deduction.54F.DateofDeposit"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value, "") Then
                fmsgbox ("Date of Deposit should be dd/mm/yyyy format at D1e In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Newly added by Bindu DE 0.2

For i = 1 To Sheet13.Range("Deduction.54F.IFSC").count
   If Not Application.Intersect(Range("Deduction.54F.IFSC"), Range(Target.Address)) Is Nothing Then
    
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then

                    fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                    Sheet13.Range(Target.Address).value = ""
                    Sheet13.Range(Target.Address).Select
                    GoTo endline
          
             ElseIf Not ValidateIFSCCode(Target.value) Then
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If
  
Next


'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.54G.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54G.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year Change by sai on 19/02/2025
        ' 'Year change from 31/03/2025 to 31/03/2026  & also AY change by Bindu on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1f In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.54GA.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54GA.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year Change by sai on 19/02/2025
        ' 'Year change from 31/03/2025 to 31/03/2026  & also AY change by Bindu on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1g In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'Changed as per DE 0.2 By Bindu
For i = 1 To Sheet13.Range("Deduction.115F.DateOfTransfer").count
     If Not Intersect(Range("Deduction.115F.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        
        If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then
        '' 'Year change from 31/03/2025 to 31/03/2026  & also AY change by Bindu on 23rd Feb 26
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1h In schedule CG ")
                Target.value = ""
                GoTo endline
        End If
      End If
Next
Dim ACG_A_Count
ACG_A_Count = Sheet13.Range("PropertySectionCount_STCG").value

'Updated as per DE 0.2 by Bindu
For i = 1 To ACG_A_Count
    If Not Intersect(Range("STCG.A_AcqDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
        If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2001") Then
'            fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2024 at Srno A1 In schedule CG ")
             'Changes from 31/03/2025 to 31/03/2026 by Bindu on 11th Feb 26
             fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2026 at Srno A1 In schedule CG ")
            Target.value = ""
            GoTo endline
        End If
    If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year change by sai on 19/02/2025'Year change by Bindu on 11th fen 26
            fmsgbox ("*Date of Purchase should not be after 31/03/2026 for AY 2026-27 A1 In schedule CG") 'Year change by Bindu on 11th fen 26
            Target.value = ""
            GoTo endline
        End If
        End If
    End If
    
  If Not Intersect(Range("STCG.AB_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
            If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year change by sai on 19/02/2025'Year change by Bindu on 11th fen 26
'            fmsgbox ("*Date of sale should not be after 31/03/2025 for AY AY 2025-26 A1 In schedule CG")
'Change by sai on 13/05/2025 as per DE V0.5
             'fmsgbox ("""Please select the date within Financial year 2025-26""") 'Year change by Bindu on 11th fen 26 earlier it was 2024-25
            'by sadineni on 24/04/2026 as per DE V0.2
              fmsgbox ("""Date of Sale/Transfer cannot be beyond 31.03.2026""")
            Target.value = ""
            GoTo endline
           End If
        End If
    End If
    
 'by sai on 13/05/2025 as per DE V0.5
 '********************************************************************************************************
    If Not Intersect(Range("STCG.AB_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
'           If Not ChkMinDOBDate_2526(target.value, "31/03/2025") Then 'Year change by Bindu on 11th fen 26
           If Not ChkMinAuditDate2627(Target.value, "31/03/2025") Then ''SIT-115930 by sadineni on 26/03/2026
            fmsgbox ("""Please check the correctness of date of sale.""")
            Target.value = ""
            GoTo endline
            End If
        End If
    End If
  '********************************************************************************************************

Next

Dim BCG_A_Count
BCG_A_Count = Sheet13.Range("PropertySectionCount_LTCG").value
'Changed as per DE 0.2 by Bindu
For i = 1 To BCG_A_Count
    If Not Intersect(Range("LTCG.B_AcqDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
           If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2001") Then
'            fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2024 at Srno B1 In schedule CG ")
'             fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2025 at Srno B1 In schedule CG ")
             
              fmsgbox ("* ""Please enter date from  01/04/2001 to 31/03/2026.""") 'Year changes from 31/03/2025 to 31/03/2026 by Bindu on 12th Feb 26
             
            Target.value = ""
            GoTo endline
        End If
        
        'Year changes from 31/03/2025 to 31/03/2026 by Bindu on 12th Feb 26
    If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year change by sai on 19/02/2025
'            fmsgbox ("*Date of Purchase should not be after 31/03/2025 for AY 2025-26 B1 In schedule CG")
            'by sai on 29/03/2025 as per DE V0.4
             fmsgbox ("* ""Please enter date from  01/04/2001 to 31/03/2026.""") 'Year changes from 31/03/2025 to 31/03/2026 by Bindu on 12th Feb 26
            Target.value = ""
            GoTo endline
        End If
        End If
    End If
        

 'Change by sai on 28/05/2025
If Not Intersect(Range("LTCG.B_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
             ''Year changes from 31/03/2025 to 31/03/2026 by Bindu on 12th Feb 26
             If Not ChkMaxDOBDate23_24(Target.value, "31/03/2026") Then 'Year change by sai on 19/02/2025
'            fmsgbox ("*Date of sale should not be after 31/03/2025 for AY 2025-26 B1 In schedule CG")
          'by sai on 29/03/2025 as per DE V0.4
'            fmsgbox ("*""Please enter valid date in sl.no.B1""date of sale/transfer' field.""")
 
'Change by sai on 13/05/2025 as per DE V0.7
'                    fmsgbox ("""Please select the date within Financial year 2025-26""") 'Year change from 2024-25 to 2025-26 by Bindu on 12th feb 26
                'by sadineni on 24/04/2026 as per DE V0.2
                    fmsgbox ("""Date of Sale/Transfer cannot be beyond 31.03.2026""")
                    Target.value = ""
             GoTo endline
             End If
        End If
        
If Sheet13.Range("LTCG.B_SaleTrfDate" & i).value <> "" Then
'             If Not ChkMinDOBDate_2526(target.value, "31/03/2025") Then 'Year changes by Bindu on 12th Feb 26 from 31/03/2024 to 31/03/2025
             If Not ChkMinAuditDate2627(Target.value, "31/03/2025") Then ''SIT-115930 by sadineni on 26/03/2026
             fmsgbox ("""Please check the correctness of date of sale.""")
               Target.value = ""
               GoTo endline
             End If
        End If
    End If
Next
'****************************************************************************************************************



'For AY 2025-26 by sai on 24/03/2025*********************************

Dim s As Long
For s = 1 To Sheet13.Range("PropertySectionCount_LTCG").value
If Not Sheet13.Range("LTCG.B_SaleTrfDate" & s) = "" Then
     If Not Application.Intersect(Sheet13.Range("LTCG.B_SaleTrfDate" & s), Sheet13.Range(Target.Address)) Is Nothing Then
         If Dformat(Sheet13.Range("LTCG.B_SaleTrfDate" & s).value, "22/07/2024") < Dformat(Sheet13.Range("LTCG.B_AcqDate" & s).value, "22/07/2024") Then
'            MsgBox ("Date of sale can't be earlier to Date of purchase / acquisition in Long term capital gain in schedule CG at S.no " & s)
            
            'by sai as per DE V0.4 on 26/03/2025
            MsgBox ("Date of sale can't be earlier to Date of purchase / acquisition in sl.no.B1 of LTCG at S.no " & s)
            
            Sheet13.Range("LTCG.B_SaleTrfDate" & s).value = ""
         End If
     End If
End If
Next
'***********************************************************************


'commented by sai 12/09/2024===================================================================================

'end ChangeCG-22.12.2022.102.03A
'        If Target.Validation.Type = 3 Then
'
'            'new change
'    Dim CG_A_Count, CG_B_Count
'     CG_A_Count = Sheet13.Range("PropertySectionCount_STCG").value
'    CG_B_Count = Sheet13.Range("PropertySectionCount_LTCG").value
'
'    For i = 1 To CG_A_Count
'        If Not Intersect(Range("CG_A_StateCodeOfBuyer" & i), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=sPassword
'            If isdropdownblank(Range(Target.Address).value) Then
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = True
'
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
'
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = True
'            End If
'            Sheet13.Protect Password:=sPassword
'            GoTo endline
'        End If
'    Next
'
'    For i = 1 To CG_B_Count
'        If Not Intersect(Range("CG_B_StateCodeOfBuyer" & i), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=sPassword
'            If isdropdownblank(Range(Target.Address).value) Then
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = True
'
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
'
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
'                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = True
'            End If
'            Sheet13.Protect Password:=sPassword
'            GoTo endline
'        End If
'    Next
'
'
'
'
'    'new change
'
'End If

'============================End by sai 12/09/2024=================================

'by sai 12/09/2024===========================================================

 Dim CG_A__STCG_Count, CG_B__LTCG_Count
 CG_A__STCG_Count = Sheet13.Range("PropertySectionCount_STCG").value
 CG_B__LTCG_Count = Sheet13.Range("PropertySectionCount_LTCG").value

 For i = 1 To CG_A__STCG_Count
            If Not Application.Intersect(Sheet13.Range("CG_A_StateCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then


      Dim PIN_targetadd, state_targetadd As String
       
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "P", "Q")
      If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet13.Range(Replace(state_targetadd, "P", "Q")).value = ""
    
End If
 
If Not Application.Intersect(Sheet13.Range("CG_A_PinCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then

     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
      'pinrng_chk = Target.address
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "Q", "P")
      cntryrng_chk = Replace(pinrng_chk, "Q", "R")
      ZIPrng_chk = Replace(pinrng_chk, "Q", "S")
      
      Set ws = Worksheets("CG")
      
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
      Sheet13.Protect Password:=sPassword
            GoTo endline
End If

 
Next


For i = 1 To CG_B__LTCG_Count
 
  

    If Not Application.Intersect(Sheet13.Range("CG_B_StateCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then

      
      Dim PIN_targetadd_LTCG, state_targetadd_LTCG As String
       
      state_targetadd_LTCG = Replace(Target.Address, "$", "")
      PIN_targetadd_LTCG = Replace(state_targetadd_LTCG, "P", "Q")
      If state_Validation(PIN_targetadd_LTCG, state_targetadd_LTCG) = False Then Sheet13.Range(Replace(state_targetadd_LTCG, "P", "Q")).value = ""
    
      
      
    End If
      

 
If Not Application.Intersect(Sheet13.Range("CG_B_PinCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then
 



 Dim pinrng_chk_LTCG, staterng_chk_LTCG, cntryrng_chk_LTCG, ZIPrng_chk_LTCG As String
     Dim ws1 As Worksheet
      'pinrng_chk = Target.address
      pinrng_chk_LTCG = Replace(Target.Address, "$", "")
      staterng_chk_LTCG = Replace(pinrng_chk_LTCG, "Q", "P")
      cntryrng_chk_LTCG = Replace(pinrng_chk_LTCG, "Q", "R")
      ZIPrng_chk_LTCG = Replace(pinrng_chk_LTCG, "Q", "S")
      
      Set ws1 = Worksheets("CG")
      
      Call PinState_codes_validation(ws1, pinrng_chk_LTCG, staterng_chk_LTCG, cntryrng_chk_LTCG, ZIPrng_chk_LTCG)
      
           Sheet13.Protect Password:=sPassword
            GoTo endline
End If

Next
'==================================================================================

        If Target.Validation.Type = 3 Then
        
            'new change
    Dim CG_A_Count, CG_B_Count
     CG_A_Count = Sheet13.Range("PropertySectionCount_STCG").value
    CG_B_Count = Sheet13.Range("PropertySectionCount_LTCG").value

    For i = 1 To CG_A_Count
        If Not Intersect(Range("CG_A_StateCodeOfBuyer" & i), Range(Target.Address)) Is Nothing Then
            
            If isdropdownblank(Range(Target.Address).value) Then
            'Change-15.12.2022.104
                'Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
                
                Sheet13.Unprotect Password:=sPassword
                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False

                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
                
                Sheet13.Unprotect Password:=sPassword
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HD8D8D8)
                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = True

                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
                
            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
                
                Sheet13.Unprotect Password:=sPassword
                 If Sheet13.Range(Replace(Target.Address, "P", "Q")).value = "" Then 'by sai
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
                End If
                
                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HD8D8D8)
                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = True
            End If
            Sheet13.Protect Password:=sPassword
            GoTo endline
        End If
    Next
    
    'Change-19.01.2023.104.23
    For i = 1 To CG_A_Count
         If Not Intersect(Range("CG_A_CountryOfBuyer" & i), Range(Target.Address)) Is Nothing Then
            If Not Mid(Target.value, 1, 2) = "91" Then
            If Not (Mid(Sheet13.Range(Replace(Target.Address, "R", "P")).value, 1, 2) = "99" Or (UCase(Sheet13.Range(Replace(Target.Address, "R", "P")).value)) = "(SELECT)") Then
                fmsgbox ("Country cannot be other than India as you have selected an Indian State in Schedule CG")
'                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
            End If
            End If
         End If
    Next
    'End Change
    
    For i = 1 To CG_B_Count
        If Not Intersect(Range("CG_B_StateCodeOfBuyer" & i), Range(Target.Address)) Is Nothing Then
           
            If isdropdownblank(Range(Target.Address).value) Then
            'Change-15.12.2022.104
                'Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
                
                 Sheet13.Unprotect Password:=sPassword
                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
    
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
                
            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
                
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
                
                 Sheet13.Unprotect Password:=sPassword
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HD8D8D8)
                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = True

                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = False
                
            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
                
                Sheet13.Unprotect Password:=sPassword
                If Sheet13.Range(Replace(Target.Address, "P", "Q")).value = "" Then 'by sai
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
                Sheet13.Range(Replace(Target.Address, "P", "Q")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "Q")).Locked = False
                End If
                
                Sheet13.Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HD8D8D8)
                Sheet13.Range(Replace(Target.Address, "P", "S")).value = ""
                Sheet13.Range(Replace(Target.Address, "P", "S")).Locked = True
            End If
            Sheet13.Protect Password:=sPassword
            GoTo endline
        End If
    Next
    
    
    'Change-19.01.2023.104.23
    For i = 1 To CG_B_Count
         If Not Intersect(Range("CG_B_CountryOfBuyer" & i), Range(Target.Address)) Is Nothing Then
            If Not Mid(Target.value, 1, 2) = "91" Then
                If Not (Mid(Sheet13.Range(Replace(Target.Address, "R", "P")).value, 1, 2) = "99" Or (UCase(Sheet13.Range(Replace(Target.Address, "R", "P")).value)) = "(SELECT)") Then
                    fmsgbox ("Country cannot be other than India as you have selected an Indian State in Schedule CG")
                    Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
                End If
            End If
         End If
    Next
    
    'new change
End If





 'Added by Aavula
 
 'SIT-94746 commented
'
'   If target.Column = 10 Then
'        If Not (target.value = "(Select)" Or target.value = "") Then
'        If 1 < Application.CountIf(Range("STCG.A8_Country"), target.value) Then
'             fmsgbox "* The same county code should not be allowed to added twice"
'             Range(target.Address).value = "(Select)"
'        End If
'        End If
'     End If
'
'
'    If target.Column = 10 Then
'        If Not (target.value = "(Select)" Or target.value = "") Then
'        If 1 < Application.CountIf(Range("LTCG.B10_Country"), target.value) Then
'             fmsgbox "* The same county code should not be allowed to added twice"
'             Range(target.Address).value = "(Select)"
'        End If
'        End If
'     End If
  ''''''''''''
  
  
  'Newly added by sai on 11/03/2025
  '***********************************************************************************************************
    If Target.Column = 8 Then
        If Not (Target.value = "(Select)" Or Target.value = "") Then
        If 1 < Application.CountIf(Range("STCG.AII_CapitalLoss20r30rapplicablerate"), Target.value) Then
             fmsgbox "* Same drop downs cannot be selected more than once "
             Range(Target.Address).value = "(Select)"
        End If
        End If
  End If

  If Target.Column = 8 Then
        If Not (Target.value = "(Select)" Or Target.value = "") Then
        If 1 < Application.CountIf(Range("LTCG.BII_Rate"), Target.value) Then
             fmsgbox "* Same drop downs cannot be selected more than once "
             Range(Target.Address).value = "(Select)"
        End If
        End If
  End If
  '***************************************************************************************************************
  
  

LockUnlock115AD9
LockUnlock115AD11
Dim temp, temp1, Temp2, Temp3, targetadd As Variant
If Not Application.Intersect(Sheet13.Range("STCG.A8_CertiStatus"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "M", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(targetadd).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet13.Range("STCG.A8_RatePTreaty").Cells.item(temp).value, Sheet13.Range("STCG.A8_RateIT").Cells.item(temp).value)
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Locked = True
        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else

        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).ClearContents
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Locked = True
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
    Sheet13.Protect Password:=sPassword
   End If


If Not Application.Intersect(Sheet13.Range("STCG.A8_RatePTreaty"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "L", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(Replace(targetadd, "L", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet13.Range("STCG.A8_RatePTreaty").Cells.item(temp).value, Sheet13.Range("STCG.A8_RateIT").Cells.item(temp).value)
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Locked = True
        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else

       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).ClearContents
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Locked = True
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
    Sheet13.Protect Password:=sPassword
   End If

   If Not Application.Intersect(Sheet13.Range("STCG.A8_RateIT"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "O", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(Replace(targetadd, "O", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet13.Range("STCG.A8_RatePTreaty").Cells.item(temp).value, Sheet13.Range("STCG.A8_RateIT").Cells.item(temp).value)
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Locked = True
        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else

        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).ClearContents
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Locked = True
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
    Sheet13.Protect Password:=sPassword
   End If



'new
If Not Application.Intersect(Sheet13.Range("LTCG.B10_CertiStatus"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "M", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(targetadd).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet13.Range("LTCG.B10_RatepTreaty").Cells.item(temp).value, Sheet13.Range("LTCG.B10_RatePITAct").Cells.item(temp).value)
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Locked = True
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Interior.ColorIndex = 2

     Else

        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).ClearContents
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Locked = True
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
    Sheet13.Protect Password:=sPassword
   End If


If Not Application.Intersect(Sheet13.Range("LTCG.B10_RatepTreaty"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "L", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(Replace(targetadd, "L", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet13.Range("LTCG.B10_RatepTreaty").Cells.item(temp).value, Sheet13.Range("LTCG.B10_RatePITAct").Cells.item(temp).value)
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Locked = True
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else

      Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).ClearContents
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Locked = True
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If

    Sheet13.Protect Password:=sPassword
End If

   If Not Application.Intersect(Sheet13.Range("LTCG.B10_RatePITAct"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "O", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(Replace(targetadd, "O", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet13.Range("LTCG.B10_RatepTreaty").Cells.item(temp).value, Sheet13.Range("LTCG.B10_RatePITAct").Cells.item(temp).value)
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Locked = True
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else

      Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).ClearContents
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Locked = True
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
    Sheet13.Protect Password:=sPassword
   End If
'Commented by sai on 03/03/2025
'If target.Validation.Type = 1 Or target.Validation.Type = 3 Then
'    Application.ScreenUpdating = False
'        AddTotalSTCGA1
'        AddTotalLTCGB1
'    Application.ScreenUpdating = True
'End If

        AddTotalSTCGA1
        AddTotalLTCGB1
      'by sai on 11/03/2025
        Cost_of_Improvement
        
       Application.EnableEvents = False

'****************************************************************************************************************
'Newly Modified by sai on 26/06/2025 as per Shubham Verma(BA Team)

'For STCG
'Commented by Bindu as per V0.01 on 11th feb 26
'uncommented by sadineni on 24/04/2026 as per DE V0.2
If Not Application.Intersect(Range("STCG.A8_AmtI"), Range(Target.Address)) Is Nothing Then
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
         If Trim(Target.value <> "") Then
            fmsgbox """Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""" & Chr(10) & Chr(10) ', vbOKOnly, "Alert"
         End If
      End If
End If

'For LTCG
 If Not Application.Intersect(Range("LTCG.B10_LtcgAmt"), Range(Target.Address)) Is Nothing Then
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Target.value <> "") Then
            fmsgbox """Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""" & Chr(10) & Chr(10) ', vbOKOnly, "Alert"
        End If
      End If
End If
'**************************************************************************************************


If Not Intersect(Range("STCG.A7_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
    STCGSecUndDedClaim (Target.row)
End If

If Not Intersect(Range("LTCG.B9_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
    LTCGSecUndDedClaim (Target.row)
End If


If Target.name.name = "STCG.A6_StatusUnutilizedCG" Then
    Sheet13.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
    Rangename = "STCG.A7_PreviousYer_1||STCG.A7_DeductionSection_1||STCG.A7_AsstAcquiredYer_1||STCG.A7_AmtUtilised_1||STCG.A7_AmtNotUsed_1" '||STCG.A7_AsstAcquiredYer_2||STCG.A7_AmtUtilised_2||STCG.A7_AmtNotUsed_2"
    Rangename = Split(Rangename, "||")
    If Mid(Sheet13.Range(Target.name.name).value, 1, 1) <> "Y" Then

        For i = 0 To UBound(Rangename)
            Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
            Sheet13.Range(Rangename(i)).Locked = True
        Next
    Else
        For i = 0 To UBound(Rangename)
            Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
            Sheet13.Range(Rangename(i)).Locked = False
        Next
    End If
Sheet13.Protect Password:=getmsgstate
Application.EnableEvents = True
End If



If Target.name.name = "LTCG.A9_StatusUnutilizedCG" Then
Sheet13.Unprotect Password:=getmsgstate
Application.EnableEvents = False
'Rangename = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1" '||LTCG.B9_AsstAcquiredYer_2||LTCG.B9_AmtUtilised_2||LTCG.B9_AmtNotUsed_2"
'Modified by sai on 13/02/2025 for AY 2025-26
'Rangename = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1||LTCG.B9_dateof_limitorwithdrawal_1" '||LTCG.B9_AsstAcquiredYer_2||LTCG.B9_AmtUtilised_2||LTCG.B9_AmtNotUsed_2"

'by sadineni For AY 2026-27 on 26/02/2026
Rangename = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1" '||LTCG.B9_AsstAcquiredYer_2||LTCG.B9_AmtUtilised_2||LTCG.B9_AmtNotUsed_2"



Rangename = Split(Rangename, "||")
    If Mid(Sheet13.Range(Target.name.name).value, 1, 1) <> "Y" Then

        For i = 0 To UBound(Rangename)
            Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
            Sheet13.Range(Rangename(i)).Locked = True
        Next
    Else
        For i = 0 To UBound(Rangename)
            Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
            Sheet13.Range(Rangename(i)).Locked = False
        Next
    End If
Sheet13.Protect Password:=getmsgstate
Application.EnableEvents = True

End If

If Target.Validation.Type = 6 Then
    Target.value = UCase(Target.value)
End If


endline:
Application.EnableEvents = True

Sheet13.Protect Password:=sPassword

End Sub


Sub LockUnlock115AD9()
    On Error Resume Next
    Sheet13.Unprotect Password:=getmsgstate

    If Sheet13.Range("LTCG.8.TotalDeductions").value = "" Then
        Sheet13.Range("LTCG.8.TotalDeductions").formula = "=Round(MAX(0,LTCG.8.CostOfAcquisition_ia+LTCG.8.CostOfImprovement+LTCG.8.ExpenditureWholly),0)"
        Sheet13.Range("LTCG.8.TotalDeductions").Interior.Color = 13434879
        Sheet13.Range("LTCG.8.TotalDeductions").Font.Color = RGB(0, 0, 255)

    End If

Sheet13.Protect Password:=getmsgstate
End Sub

Sub LockUnlock115AD11()
    On Error Resume Next
    Sheet13.Unprotect Password:=getmsgstate

    If Sheet54.Range("LTCG.8.Balance").value = "" Then
        Sheet13.Range("LTCG.8.Balance").formula = "=Round((LTCG.8.FullValue-LTCG.8.TotalDeductions),0)"
        Sheet13.Range("LTCG.8.Balance").Interior.Color = 13434879
        Sheet13.Range("LTCG.8.Balance").Font.Color = RGB(0, 0, 255)

    End If

Sheet13.Protect Password:=getmsgstate
End Sub


Private Sub Worksheet_Deactivate()
Dim activesh As Variant
activesh = ActiveSheet.name

If UCase(Sheet13.Range("CG_TableE_Checkbox").value) = UCase("True") Then
    
    If Sheet13.Range("TableE_Negativecount").value > 0 Then
      MsgBox ("**Amount of set off claimed cannot exceed amount available for set off")
      Sheet13.Activate
      Exit Sub
    End If
End If



CG_Calc.doSetoff
CYLACalculations.CYLABFLASetOff
Sheets(activesh).Activate
End Sub

Sub AddTotalSTCGA1()
    Dim Row7RangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim Totsections, tot, tot1 As Variant
    Dim tot2_1, tot7_1, tot8_1, tot10_1 As Variant


    'Row7RangeName = "SAL.IncomeFromSalary"
    Totsections = Sheet13.Range("PropertySectionCount_STCG").value
    tot = 0
    tot1 = 0
    tot2_1 = 0
    tot7_1 = 0
    tot8_1 = 0
    tot10_1 = 0

         sPassword = EfilingCommon.getmsgstate
         Sheet13.Unprotect Password:=sPassword

        For i = 1 To Totsections
        'newDefinedName = Row7RangeName & i


        tot = tot + Sheet13.Range("STCG.Ae_CapgainonAssets" & i).value
         tot1 = tot1 + Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC" & i).value
        tot2_1 = tot2_1 + Sheet13.Range("STCG.Adi_ExemptionAmount" & i).value

        tot7_1 = tot7_1 + Sheet13.Range("STCG.Adiv_ExemptionAmount" & i).value
        tot8_1 = tot8_1 + Sheet13.Range("STCG.Adiii_ExemptionAmount" & i).value
        'tot10_1 = tot10_1 + Sheet13.Range("STCG.Adii_ExemptionAmount" & i).value

        Next

    Sheet13.Protect Password:=sPassword
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword

'----------------Unlock Password-------------------END-----

    Sheet13.Range("STCG.A1e_CapgainonAssets").value = tot
    Sheet13.Range("STCG.Aaiii_FullConsiderationFiftyC_Temp").value = tot1
    Sheet13.Range("D_54B_Temp1").value = tot2_1
    Sheet13.Range("D_54GA_Temp1").value = tot7_1
    Sheet13.Range("D_54G_Temp1").value = tot8_1
    'Sheet13.Range("D_54D_Temp1").value = tot10_1


'----------------Lock Password-------------------START---

   Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub

Sub AddTotalLTCGB1()
    Dim Row7RangeName As String
    Dim i As Variant
    Dim j As Variant 'by sai for AY 2025-26
    Dim newDefinedName As String
    Dim Totsections, tot, tot1, tot2, tot3, tot4, tot5, tot6, tot7, tot8, tot9, tot10, tot11 As Variant
    
    Dim tot12, tot13, tot14 As Variant 'By sai on 03/03/2025

    'Row7RangeName = "SAL.IncomeFromSalary"
    Totsections = Sheet13.Range("PropertySectionCount_LTCG").value
    tot = 0
    tot1 = 0
    tot2 = 0
    tot3 = 0
    tot4 = 0
    tot5 = 0
    tot6 = 0
    tot7 = 0
    tot8 = 0
    tot9 = 0
    tot10 = 0
    tot11 = 0
    tot12 = 0
    tot13 = 0
    tot14 = 0

         sPassword = EfilingCommon.getmsgstate
         Sheet13.Unprotect Password:=sPassword

        For i = 1 To Totsections
        'newDefinedName = Row7RangeName & i


        tot = tot + Sheet13.Range("LTCG.Be_CapgainonAssets" & i).value
        'tot1 = tot1 + Sheet13.Range("LTCG.Bdviii_ExemptionAmount" & i).value

           tot2 = tot2 + Sheet13.Range("LTCG.Bdii_ExemptionAmount" & i).value
           tot3 = tot3 + Sheet13.Range("LTCG.Bdi_ExemptionAmount" & i).value
           tot4 = tot4 + Sheet13.Range("LTCG.Bdiv_ExemptionAmount" & i).value
           tot5 = tot5 + Sheet13.Range("LTCG.Bdv_ExemptionAmount" & i).value
          ' tot6 = tot6 + Sheet13.Range("LTCG.Bdviii_ExemptionAmount" & i).value
           tot7 = tot7 + Sheet13.Range("LTCG.Bdvii_ExemptionAmount" & i).value
           tot8 = tot8 + Sheet13.Range("LTCG.Bdvi_ExemptionAmount" & i).value
           tot10 = tot10 + Sheet13.Range("LTCG.Bdiii_ExemptionAmount" & i).value
'           tot11 = tot11 + Sheet13.Range("LTCG.Bdix_ExemptionAmount54EE" & i).value

           tot9 = tot9 + Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & i).value

        Next

    Sheet13.Protect Password:=sPassword
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword

'----------------Unlock Password-------------------END-----

    Sheet13.Range("LTCG.B1e_CapgainonAssets").value = tot
    'Sheet13.Range("LTCG_Temp54GB").value = tot1
    Sheet13.Range("D_54B_Temp").value = tot2
    Sheet13.Range("D_54_Temp").value = tot3
    Sheet13.Range("D_54EC_Temp").value = tot4
'    Sheet13.Range("D_54EE_Temp").value = tot11
    Sheet13.Range("D_54F_Temp").value = tot5
    Sheet13.Range("D_54GB_Temp").value = tot6
    Sheet13.Range("D_54GA_Temp").value = tot7
    Sheet13.Range("D_54G_Temp").value = tot8
    Sheet13.Range("D_54D_Temp").value = tot10
    Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC_Temp").value = tot9
    
'Added by Sai on 26/02/2024
    For j = 1 To Totsections
          
       If Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & j).value = 0 Then
            Sheet13.Range("LTCG.Bbiv_TotalDedn" & j).value = 0
        Else
'Commented by Riyaz on 18/02/2026
'        If Dformat(Sheet13.Range("LTCG.B_SaleTrfDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then '' updating from this LTCG.Bbii_ImproveCost to this  LTCG.Bbia_ImproveCost 'Before 'Added by Sai on 26/02/2024
            Sheet13.Range("LTCG.Bbiv_TotalDedn" & j).value = Sheet13.Range("LTCG.Bbi_AquisitCostNew" & j).value + Sheet13.Range("LTCG.Bbia_ImproveCost_Sum" & j).value + Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & j).value
        End If
'        End If
'Commented by Riyaz on 18/02/2026
'        If Sheet13.Range("LTCG.B_SaleTrfDate" & j).value = "" Then
'            Sheet13.Range("LTCG.Bbiv_TotalDedn" & j).value = 0
'        Else
'        If Dformat(Sheet13.Range("LTCG.B_SaleTrfDate" & j).value, "23/07/2024") >= Dformat("23/07/2024", "23/07/2024") Then ' updating from this Bbi_AquisitCost to this  LTCG.Bbia_ImproveCost ' 'Added by Sai on 26/02/2024
'            Sheet13.Range("LTCG.Bbiv_TotalDedn" & j).value = Sheet13.Range("LTCG.Bbi_AquisitCostNew" & j).value + Sheet13.Range("LTCG.Bbia_ImproveCost_Sum" & j).value + Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & j).value
'        End If
'        End If
            
'Added by Sai on 26/02/2024
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
'Changed by Riyaz on 18/02/2026
            If Sheet13.Range("LTCG.B_AcqDate" & j).value = "" Then
                Sheet13.Range("LTCG.Bbiva_TotalDedn" & j).value = 0
            Else
                If Dformat(Sheet13.Range("LTCG.B_AcqDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then
                    Sheet13.Range("LTCG.Bbiva_TotalDedn" & j).value = Sheet13.Range("LTCG.Bbi_AquisitCost" & j).value + Sheet13.Range("LTCG.Bbii_ImproveCost_Sum" & j).value + Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & j).value
                Else
                    Sheet13.Range("LTCG.Bbiva_TotalDedn" & j).value = 0
                End If
            End If
        Else
            Sheet13.Range("LTCG.Bbiva_TotalDedn" & j).value = 0
        End If
        
'Added by Sai on 26/02/2024
'S.no ca
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
'Changed by Riyaz on 18/02/2026
            If Sheet13.Range("LTCG.B_AcqDate" & j).value = "" Then
                Sheet13.Range("LTCG.Bca_BalanceCG" & j).value = 0
            Else
                If Dformat(Sheet13.Range("LTCG.B_AcqDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then
                    Sheet13.Range("LTCG.Bca_BalanceCG" & j).formula = "=MAX(0,(LTCG.Baiii_FullConsiderationFiftyC" & j & "-" & "LTCG.Bbiva_TotalDedn" & j & ")" & ")"
                Else
                    Sheet13.Range("LTCG.Bca_BalanceCG" & j).value = 0
                End If
            End If
        Else
            Sheet13.Range("LTCG.Bca_BalanceCG" & j).value = 0
        End If
'end
'S.no B1e(a)
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
'Changed by Riyaz on 18/02/2026
            If Sheet13.Range("LTCG.B_AcqDate" & j).value = "" Then
                Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).value = 0
            Else
                If Dformat(Sheet13.Range("LTCG.B_AcqDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then
                    Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).formula = "=IF(LTCG.Bca_BalanceCG" & j & "<0" & "," & "LTCG.Bca_BalanceCG" & j & "," & "IF(LTCG.Bd_ExemptionGrandTotal" & j & ">" & "LTCG.Bca_BalanceCG" & j & "," & 0 & "," & "LTCG.Bca_BalanceCG" & j & "-" & "LTCG.Bd_ExemptionGrandTotal" & j & ")" & ")"
                Else
                    Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).value = 0
                End If
            End If
        Else
            Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).value = 0
        End If
        
'S.no B1ei(A)
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
'Changed by Riyaz on 18/02/2026
            If Sheet13.Range("LTCG.B_AcqDate" & j).value = "" Then
                Sheet13.Range("LTCG.BeiA_TaxSec" & j).value = 0
            Else
            If Dformat(Sheet13.Range("LTCG.B_AcqDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then
                Sheet13.Range("LTCG.BeiA_TaxSec" & j).formula = "=MAX(0,ROUND(LTCG.Be_CapgainonAssets" & j & "*12.5%,0))"
            Else
                Sheet13.Range("LTCG.BeiA_TaxSec" & j).value = 0
            End If
            End If
        Else
            Sheet13.Range("LTCG.BeiA_TaxSec" & j).value = 0
        End If
        
'S.no B1ei(B) ' Need to check with BA Bcz in DE Sheet not mentioned
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
'Changed by Riyaz on 18/02/2026
            If Sheet13.Range("LTCG.B_AcqDate" & j).value = "" Then
                Sheet13.Range("LTCG.BeiB_TaxSec" & j).value = 0
            Else
            If Dformat(Sheet13.Range("LTCG.B_AcqDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then
                Sheet13.Range("LTCG.BeiB_TaxSec" & j).formula = "=MAX(0,ROUND(LTCG.Bea_CapgainonAssets" & j & "*20%,0))"
            Else
                Sheet13.Range("LTCG.BeiB_TaxSec" & j).value = 0
            End If
            End If
            Else
                Sheet13.Range("LTCG.BeiB_TaxSec" & j).value = 0
        End If
'Added by Sai on 26/02/2024
'Before 23rd July

'Changed by Sai
        If Sheet13.Range("LTCG.B_AcqDate" & j).value <> "" Then
                tot12 = tot12 + Sheet13.Range("LTCG.Be_CapgainonAssets" & j).value
        End If

'After 23rd July
'Changed by Sai
        If Sheet13.Range("LTCG.B_AcqDate" & j).value <> "" Then
                tot13 = tot13 + Sheet13.Range("LTCG.Beii_Excessamt" & j).value
        End If
        
           
           Next j
'--End
    Sheet13.Range("LTCG.B1g_ImmpropertyB1e1").value = tot12
    Sheet13.Range("LTCG.B1g_ImmpropertyB1e1c").value = tot13

'----------------Lock Password-------------------START---

   Sheet13.Protect Password:=sPassword
   
   Application.EnableEvents = True
'----------------Lock Password-------------------END-----
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet1.Range(targetaddress).Value = ""
ElseIf ActiveCell.Validation.Type = 1 Then
 Application.OnKey "%{DOWN}", ""
ElseIf ActiveCell.Validation.Type = 3 Then

    Application.OnKey "%{DOWN}"
End If
Else
Application.OnKey "%{DOWN}", ""
End If

End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub

Sub Cost_of_Improvement()

Dim CG_B__LTCG_Count, i, j, cellcount  As Long
Dim rangecells1, rangecells2, rangecells3, rangecells4 As Range
Dim sale_Date_after As Boolean
Dim Lookup_year As Variant
Dim loopkup_1, loopkup_2

CG_B__LTCG_Count = Sheet13.Range("PropertySectionCount_LTCG").value

 
For i = 1 To CG_B__LTCG_Count
    cellcount = Sheet13.Range("LTCG.Bbii_ImproveCost" & i).count
    
    Set rangecells1 = Sheet13.Range("LTCG.Bbia_ImproveCost" & i)
    Set rangecells2 = Sheet13.Range("LTCG.Bbii_ImproveCost" & i)
    Set rangecells3 = Sheet13.Range("LTCG.Bbib_ImproveCostYear_Vlookup" & i)
    Set rangecells4 = Sheet13.Range("LTCG.Bbib_ImproveCostYear" & i)
    
    If Sheet13.Range("LTCG.B_AcqDate" & i).value <> "" Then
    
            sale_Date_after = SchCG.ChkMinInclusiveDate(Dformat((Sheet13.Range("LTCG.B_AcqDate" & i).value), "yyyy-mm-dd"), "2024-07-23") ' If true means on or after 23rd July
               
            For j = 1 To cellcount
                Debug.Print rangecells1.Cells(j, 1).Address
                If Sheet13.Range("LTCG.B_AcqDate" & i).value = "" Then
                    Lookup_year = 2001
                End If
'Updated by sadineni on 06/06/2026 ("LTCG.B_AcqDate" was updated to "LTCG.B_SaleTrfDate")
                If Application.WorksheetFunction.ValueToText(Mid(Sheet13.Range("LTCG.B_SaleTrfDate" & i).value, 4, 2)) > 3 Then
                  Lookup_year = (Mid(Sheet13.Range("LTCG.B_SaleTrfDate" & i).value, 7, 4) + 1) - 1
                Else
                   Lookup_year = (Mid(Sheet13.Range("LTCG.B_SaleTrfDate" & i).value, 7, 4) - 1)
                End If
                
               loopkup_1 = Application.VLookup(Lookup_year, SheetALL.Range("CostInflationtbl"), 2, False)
               loopkup_2 = rangecells3.Cells(j, 1).value
                sPassword = EfilingCommon.getmsgstate
               
               If rangecells1.Cells(j, 1).value <> "" And rangecells4.Cells(j, 1).value <> "" Then
                  
                    Sheet13.Unprotect Password:=sPassword
                    Application.EnableEvents = False
        
            
                    If sale_Date_after = True Or Mid(Sheet1.Range("sheet1.ResidentialStatus1"), 1, 3) = "NRI" Then
                    
                        rangecells2.Cells(j, 1).value = 0
                    Else
                        rangecells2.Cells(j, 1).value = Application.WorksheetFunction.IfError(Round(rangecells1.Cells(j, 1).value * loopkup_1 / loopkup_2, 0), 0)
                        Debug.Print rangecells2.Cells(j, 1).value
                    End If
                Else
                    Sheet13.Unprotect Password:=sPassword
                    Application.EnableEvents = False
                    rangecells2.Cells(j, 1).value = 0
                    
              End If
              
            Next
   Else
      
       For j = 1 To cellcount
           Sheet13.Unprotect Password:=sPassword
           Application.EnableEvents = False
           Application.EnableEvents = False
           rangecells2.Cells(j, 1).value = 0
       Next
            
   End If
   

Next
Application.EnableEvents = True

End Sub



Sub test()
Sheet13.Range("LTCG.B7e_CapgainonAssets").EntireRow.Hidden = False
Sheet13.Range("LTCG.B7e_CapgainonAssets_ii").EntireRow.Hidden = False
Sheet13.Range("LTCG.B4eiii_LTCGOnUnlistedSecurity").EntireRow.Hidden = False

End Sub

'SIT-115930 by sadineni on 26/03/2026
Function ChkMinAuditDate2627(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinAuditDate2627 = True
    If Len(dob) > 0 Then
 
        If val(Mid(dob, 7, 4)) < 2025 Then
           ChkMinAuditDate2627 = False
           Exit Function
        ElseIf (val(Mid(dob, 7, 4)) = 2025 And val(Mid(dob, 4, 2)) <= 3) Then
           ChkMinAuditDate2627 = False
           Exit Function
       End If
   End If
End Function

