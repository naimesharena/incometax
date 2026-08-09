Attribute VB_Name = "Sheet53"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error Resume Next
    ActiveSheet.Shapes("Button 1").Left = Range("Button_112A").Left + 2
    ActiveSheet.Shapes("Button 1").Top = Range("Button_112A").Top + 2
    
    
'If UCase(ActiveSheet.name) = UCase("Schedule 112A") Then ' Added by Bindu for Table E approach
    
For Each cell In Sheet53.Range("ShareAcq_112A")

 If cell.Locked = False Then

    'If LockFlag112A = False Then
    
      MsgBox "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F."
    Exit Sub
    'End If
 End If
    
Next
    
'End If

'    ThisWorkbook.Protect Password:=getmsgstate
End Sub



Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim targetadd As Variant
Application.EnableEvents = False
Application.ScreenUpdating = False
Sheet53.Unprotect Password:=getmsgstate
targetadd = Replace(Target.Address, "$", "")

If Not Application.Intersect(Range("ShareAcq_112A"), Range(Target.Address)) Is Nothing Then

    If Target.value = "After 31st January 2018" Then
    
        'Changed by Riyaz on 02/01/2025
    
'        Range(Replace(Target.Address, "E", "F")).ClearContents
'        Range(Replace(Target.Address, "E", "F")).Locked = True
'        Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
'        Range(Replace(Target.Address, "E", "F")).value = "INNOTREQUIRD"
'
'        Range(Replace(Target.Address, "E", "G")).ClearContents
'        Range(Replace(Target.Address, "E", "G")).Locked = True
'        Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
'        Range(Replace(Target.Address, "E", "G")).value = "CONSOLIDATED"
'
'        Range(Replace(Target.Address, "E", "H")).ClearContents
'        Range(Replace(Target.Address, "E", "H")).Locked = True
'        Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HD8D8D8"
'        Range(Replace(Target.Address, "E", "H")).value = 0
'
'        Range(Replace(Target.Address, "E", "I")).ClearContents
'        Range(Replace(Target.Address, "E", "I")).Locked = True
'        Range(Replace(Target.Address, "E", "I")).Interior.Color = "&HD8D8D8"
'        Range(Replace(Target.Address, "E", "I")).value = 0
'
'        Range(Replace(Target.Address, "E", "N")).ClearContents
'        Range(Replace(Target.Address, "E", "N")).Locked = True
'        Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HD8D8D8"
'        Range(Replace(Target.Address, "E", "N")).value = 0
'
'        Range(Replace(Target.Address, "E", "J")).ClearContents
'        Range(Replace(Target.Address, "E", "J")).Locked = False
'        Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HCCFFCC"
'        Range(Replace(Target.Address, "E", "J")).Font.Color = vbBlack


        Range(Replace(Target.Address, "E", "G")).ClearContents
        Range(Replace(Target.Address, "E", "G")).Locked = True
        Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "G")).value = "INNOTREQUIRD"
        
        Range(Replace(Target.Address, "E", "H")).ClearContents
        Range(Replace(Target.Address, "E", "H")).Locked = True
        Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "H")).value = "CONSOLIDATED"

        Range(Replace(Target.Address, "E", "I")).ClearContents
        Range(Replace(Target.Address, "E", "I")).Locked = True
        Range(Replace(Target.Address, "E", "I")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "I")).value = 0

        Range(Replace(Target.Address, "E", "J")).ClearContents
        Range(Replace(Target.Address, "E", "J")).Locked = True
        Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "J")).value = 0

        Range(Replace(Target.Address, "E", "O")).ClearContents
        Range(Replace(Target.Address, "E", "O")).Locked = True
        Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "O")).value = 0
        
        Range(Replace(Target.Address, "E", "K")).ClearContents
        Range(Replace(Target.Address, "E", "K")).Locked = False
       ' Range(Replace(Target.Address, "E", "K")).FormulaHidden = False          'added by Chetan C M on 29/10/2025 for SIT-89185 (issue 21)
        Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
        Range(Replace(Target.Address, "E", "K")).Font.Color = vbBlack
        
    Else
    
''        Range(Replace(Target.address, "E", "F")).ClearContents
'        Range(Replace(Target.Address, "E", "F")).Locked = False
'        Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "G")).ClearContents
'        Range(Replace(Target.Address, "E", "G")).Locked = False
'        Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "H")).ClearContents
'        Range(Replace(Target.Address, "E", "H")).Locked = False
'        Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "I")).ClearContents
'        Range(Replace(Target.Address, "E", "I")).Locked = False
'        Range(Replace(Target.Address, "E", "I")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "N")).ClearContents
'        Range(Replace(Target.Address, "E", "N")).Locked = False
'        Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"
'
'
''        Range(Replace(Target.address, "E", "J")).ClearContents
'        Range(Replace(Target.Address, "E", "J")).Locked = True
'        Range(Replace(Target.Address, "E", "J")).Interior.Color = vbWhite
'        Range(Replace(Target.Address, "E", "J")).Font.Color = RGB(51, 102, 255)
        


'        Range(Replace(Target.address, "E", "F")).ClearContents
        Range(Replace(Target.Address, "E", "G")).Locked = False
        Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"
        
'        Range(Replace(Target.address, "E", "G")).ClearContents
        Range(Replace(Target.Address, "E", "H")).Locked = False
        Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HCCFFCC"

'        Range(Replace(Target.address, "E", "H")).ClearContents
        Range(Replace(Target.Address, "E", "I")).Locked = False
        Range(Replace(Target.Address, "E", "I")).Interior.Color = "&HCCFFCC"

'        Range(Replace(Target.address, "E", "I")).ClearContents
        Range(Replace(Target.Address, "E", "J")).Locked = False
        Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HCCFFCC"

'        Range(Replace(Target.address, "E", "N")).ClearContents
        Range(Replace(Target.Address, "E", "O")).Locked = False
        Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HCCFFCC"
        
        
'        Range(Replace(Target.address, "E", "J")).ClearContents
        Range(Replace(Target.Address, "E", "K")).Locked = True
       ' Range(Replace(Target.Address, "E", "K")).FormulaHidden = True          'added by Chetan C M on 29/10/2025 for SIT-89185 (issue 21)
        Range(Replace(Target.Address, "E", "K")).Interior.Color = vbWhite
        Range(Replace(Target.Address, "E", "K")).Font.Color = RGB(51, 102, 255)


        'MsgBox ("=MAX(0," & "ROUND(" & Replace(TargetAdd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(TargetAdd, "E", "I") & ",4)" & ")")
'        Range(Replace(Target.Address, "E", "J")).formula = "=ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "),0)"
        'Range(Replace(Target.Address, "E", "k")).formula = "=(ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "J") & ",4)" & "),0)"

        Range(Replace(Target.Address, "E", "k")).formula = "=IFERROR(ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "J") & ",4)" & "),0),0)"   'Added by Shrutika(30/10/2025)
        
 '-------------------------End-------------------------------------------------
    End If
    If Target.value = "" Then
       Target.value = "(Select)"
    End If
End If

'added by Chetan C M on 28/10/2025 for SIT-89185(issue 19,20,21)
'start--

If Not Application.Intersect(Range("SalePriceperShare_112A"), Range(Target.Address)) Is Nothing Then

        Dim int_num, decimal_num, find_num
        
        find_num = InStr(1, Target.value, ".")
        
        If find_num = 1 Then
                MsgBox ("Sale-price per Share/Unit should starts with number") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
        End If
        
        If find_num > 0 Then
            int_num = Left(Target.value, find_num - 1)
            decimal_num = Mid(Target.value, find_num + 1)
        Else
            int_num = Target.value
        End If
        
            If Not IsNumeric(int_num) Then
                MsgBox ("Sale-price per Share/Unit should be a numeric value.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If Len(int_num) > 14 Then
                MsgBox ("Sale-price per Share/Unit cannot exceed 14 digits before decimal.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If (int_num > 99999999999999#) Then
                MsgBox ("Sale-price per Share/Unit cannot exceed 14 digits before decimal.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
             
             If decimal_num > 0 Then
                If Not IsNumeric(decimal_num) Then
                    MsgBox ("Sale-price per Share/Unit should be a numeric value after decimal.") & Chr(13)
                    Sheet53.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
             End If
            
            If Len(decimal_num) > 4 Then
                MsgBox ("Sale-price per Share/Unit cannot exceed 4 digits after decimal.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
        End If

Target.value = md112A.RoundToFourDigits(Target.value)
End If

If Not Application.Intersect(Range("NoofShare_112A"), Range(Target.Address)) Is Nothing Then
        
        find_num = InStr(1, Target.value, ".")
        
        If find_num = 1 Then
        
         MsgBox ("No. of Shares/Units should starts with number") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
        
        If find_num > 0 Then
        
            int_num = Left(Target.value, find_num - 1)
            decimal_num = Mid(Target.value, find_num + 1)
        
        Else
            int_num = Target.value
        End If
        
            If Not IsNumeric(int_num) Then
                MsgBox ("No. of Shares/Units should be a numeric value.")
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If Len(int_num) > 14 Then
                MsgBox ("No. of Shares/Units cannot exceed 14 digits before decimal.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If (int_num > 99999999999999#) Then
                MsgBox ("No. of Shares/Units cannot exceed 14 digits.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
             
             If decimal_num > 0 Then
                If Not IsNumeric(decimal_num) Then
                    MsgBox ("No. of Shares/Units should be a numeric value after decimal.") & Chr(13)
                    Sheet53.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
            End If
            
            If Len(decimal_num) > 4 Then
                MsgBox ("No. of Shares/Units cannot exceed 4 digits after decimal.") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
        End If
            
        
        Target.value = md112A.RoundToFourDigits(Target.value)
        
        End If
        
        '================================
        'FairMarketValue_112A
            If Not Application.Intersect(Range("FairMarketValue_112A"), Range(Target.Address)) Is Nothing Then
            
            
            find_num = InStr(1, Target.value, ".")
            
            If find_num = 1 Then
            
             MsgBox ("Fair Market Value per share/unit should starts with number") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If find_num > 0 Then
            
                int_num = Left(Target.value, find_num - 1)
                decimal_num = Mid(Target.value, find_num + 1)
            
            Else
                int_num = Target.value
            End If
            
                If Not IsNumeric(int_num) Then
                    MsgBox ("Fair Market Value should be a numeric value.") & Chr(13)
                    Sheet53.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
            End If
                
                If Len(int_num) > 14 Then
                    MsgBox ("Fair Market Value cannot exceed 14 digits before decimal.") & Chr(13)
                    Sheet53.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
            End If
                
                If (int_num > 99999999999999#) Then
                    MsgBox ("Fair Market Value cannot exceed 14 digits.") & Chr(13)
                    Sheet53.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
            End If
                 
                 If decimal_num > 0 Then
                    If Not IsNumeric(decimal_num) Then
                        MsgBox ("Fair Market Value should be a numeric value after decimal.") & Chr(13)
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
            End If
                
                If Len(decimal_num) > 4 Then
                     MsgBox ("Fair Market Value cannot exceed 4 digits after decimal.") & Chr(13)
                    Sheet53.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
                End If
                 
           End If
                
            '----
                    Target.value = md112A.RoundToFourDigits(Target.value)
        
        End If
'========================================================
             'CostOfAcquisition_112A
       If Not Application.Intersect(Range("COAwithIndex_112A"), Range(Target.Address)) Is Nothing Then
            
            
            find_num = InStr(1, Target.value, ".")
            
            If find_num = 1 Then
            
             MsgBox ("Cost Of Acquisition should starts with number") & Chr(13)
                Application.EnableEvents = False
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If find_num > 0 Then
            
                int_num = Left(Target.value, find_num - 1)
                decimal_num = Mid(Target.value, find_num + 1)
            
            Else
                int_num = Target.value
            End If
            
                If Not IsNumeric(int_num) Then
                        MsgBox ("Cost Of Acquisition should be a numeric value.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(int_num) > 14 Then
                        MsgBox ("Cost Of Acquisition cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                
                If (int_num > 99999999999999#) Then
                        MsgBox ("Cost Of Acquisition cannot exceed 14 digits.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                 
                 If decimal_num > 0 Then
                    If Not IsNumeric(decimal_num) Then
                        MsgBox ("Cost Of Acquisition should be a numeric value after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(decimal_num) > 4 Then
                        MsgBox ("Cost Of Acquisition cannot exceed 4 digits after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                 
                 End If
                
            '----
          Target.value = md112A.RoundToFourDigits(Target.value)
        End If

        '================================
        'Expenditure_112A
       If Not Application.Intersect(Range("ExpenditureWholly_112A"), Range(Target.Address)) Is Nothing Then
            
            
            find_num = InStr(1, Target.value, ".")
            
            If find_num = 1 Then
             MsgBox ("Expenditure wholly and exclusively should starts with number") & Chr(13)
                Sheet53.Unprotect Password:=getmsgstate
                Target.ClearContents
                GoTo endline
            End If
            
            If find_num > 0 Then
            
                int_num = Left(Target.value, find_num - 1)
                decimal_num = Mid(Target.value, find_num + 1)
            
            Else
                int_num = Target.value
            End If
            
                If Not IsNumeric(int_num) Then
                        MsgBox ("Expenditure wholly and exclusively should be a numeric value.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(int_num) > 14 Then
                        MsgBox ("Expenditure wholly and exclusively cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If (int_num > 99999999999999#) Then
                        MsgBox ("Expenditure wholly and exclusively cannot exceed 14 digits.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                 
                 If decimal_num > 0 Then
                    If Not IsNumeric(decimal_num) Then
                        MsgBox ("Expenditure wholly and exclusively should be a numeric value after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(decimal_num) > 4 Then
                        MsgBox ("Expenditure wholly and exclusively cannot exceed 4 digits after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet53.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                 
                 End If
       

Target.value = md112A.RoundToFourDigits(Target.value)
End If
'--end

If Not Application.Intersect(Range("ISINCode_112A"), Range(Target.Address)) Is Nothing Then
If Target.value <> "" Then

Target.value = UCase(Target.value)
 If UCase(Mid(Range(Target.Address), 1, 2)) <> "IN" Then
  'Ankita
 'fmsgboxStatus "Please enter the valid ISIN Code in sheet 112A"
  fmsgboxStatus "Please enter the valid ISIN Code in Column 2 of Schedule 112A" & Chr(13)
 Target.value = ""
 End If
 
 'addded by Chetan C M on 28/10/2025 for SIT-89185(issue 19,20,21)
'start--
If Not checkfieldspecialcharacter(Target.value) Then
            MsgBox ("Please don't use Special Characters < and  > " & Chr(13))
            Target.value = ""
            GoTo endline
End If
'--end

    Sheet53.Range(Target.Address).Select
    Sheet53.Range(Target.Address).Offset(0, 1) = UCase(Sheet53.Range(Target.Address).Offset(0, 26).value)
End If



If Target.value = "" Then
    Sheet53.Range(Target.Address).Offset(0, 1) = ""
End If

End If


'Commented by Shrutika

'sudha
'If Application.Intersect(Range("ShareAcq_112A"), Range(Target.Address)) Is Nothing Then
'
'    If Target.value = "" Then
'    Target.value = "(Select)"
'    End If
'End If
'sudha

endline:
Application.EnableEvents = True
Application.ScreenUpdating = True
Sheet53.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Deactivate()
Dim activesh As Variant
activesh = ActiveSheet.name

CG_Calc.doSetoff
'CYLACalculations.CYLABFLASetOff

'Updated by Bindu
If (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("No") And UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("No")) Then
    CYLACalculations.CYLABFLASetOff
ElseIf (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("Yes") Or UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("Yes")) Then
    CYLA_BFLA_Editable
End If
'--

Sheets(activesh).Activate

If ThisWorkbook.ProtectStructure = False Then
    ThisWorkbook.Protect Password:=getmsgstate
End If
'Adding the condition of 115AD lock if any one row is filled in 112A
    Dim end_112A As Long
    Dim range115AD As Variant
    
    setTableInfo112A
    setTableInfo112A1
    setTableInfo112A2
    setTableInfo112A3
    setTableInfo112A4
    setTableInfo112A5
'    setTableInfo112A6
    setTableInfo112A7
    setTableInfo112A8
    
    end_112A = WorksheetFunction.Max(end112A, end112A1, end112A2, end112A3, end112A4, end112A5, end112A7, end112A8)
    
    Application.EnableEvents = False
    Sheet54.Unprotect Password:=getmsgstate
    If end_112A > 0 Then
        If LockFlag115A = False Then
            MsgBox ("Please fill either Schedule 112A or Schedule 115AD(1)(b)(iii) proviso as applicable in your case")
        End If
        
'Changed by sudha on 25th Dec 24

'        range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
'        LOCKRANGE115AD (range115AD)
'
'    Else
'        If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" And LockFlag115A = True Then
'            range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
'            UNLOCKRANGE115AD (range115AD)
'        End If

    'changed by Chetan C M on 22/01/2026 for AY 2026-27
        'start--
'        range115AD = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
        range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
        '--end
        
        LOCKRANGE115AD (range115AD)
        
    Else
        If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" And LockFlag115A = True Then
        
        'changed by Chetan C M on 22/01/2026 for AY 2026-27
        'start--
'            range115AD = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
            range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
        '--end
        
            UNLOCKRANGE115AD (range115AD)
        End If
        
'--------------------------end------------------------------
        
        'call the unlock115Ad
    End If
    Application.EnableEvents = True
    Sheet54.Protect Password:=getmsgstate
    
endline:
    Application.EnableEvents = True
    Sheet54.Protect Password:=getmsgstate
End Sub



