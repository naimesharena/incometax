Attribute VB_Name = "Sheet54"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Activate()
On Error Resume Next
    ActiveSheet.Shapes("Button 1").Left = Range("Button_115AD").Left + 2
    ActiveSheet.Shapes("Button 1").Top = Range("Button_115AD").Top + 2
    
    
For Each cell In Sheet54.Range("ShareAcq_115AD_1")

 If cell.Locked = False Then
     
    'If LockFlag115A = False Then
    
     fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F."
     Exit Sub
    'End If
 End If
    
Next
    
    

'    ThisWorkbook.Protect Password:=getmsgstate
End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False
Application.ScreenUpdating = False
Sheet54.Unprotect Password:=getmsgstate
targetadd = Replace(Target.Address, "$", "")

If Not Application.Intersect(Range("ShareAcq_115AD_1"), Range(Target.Address)) Is Nothing Then

    If Target.value = "After 31st January 2018" Then
    
'Changed by sudha on 25th Dec 24
    
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
        
        Range(Replace(Target.Address, "E", "k")).formula = "=IFERROR(ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "J") & ",4)" & "),0),0)"   'Added by Shrutika(30/10/2025)
        
        
 '-------------------------End-------------------------------------------------
        
        '=MAX(0,H6*I6)
        'vbWhite
    End If
    If Target.value = "" Then
       Target.value = "(Select)"
    End If
End If

'added by Chetan C M on 28/10/2025 for SIT-89185(issue 19,20,21)
'start--
'SalePriceperShare_115AD
If Not Application.Intersect(Range("SalePriceperShare_115AD"), Range(Target.Address)) Is Nothing Then

        Dim int_num, decimal_num, find_num
        
        find_num = InStr(1, Target.value, ".")
        
        If find_num = 1 Then
         MsgBox ("Sale-price per Share/Unit should starts with number") & Chr(13)
            Application.EnableEvents = False
            Sheet54.Unprotect Password:=getmsgstate
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
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
            
            If Len(int_num) > 14 Then
                 MsgBox ("Sale-price per Share/Unit cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
            
            If (int_num > 99999999999999#) Then
                 MsgBox ("Sale-price per Share/Unit cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
             
             If decimal_num > 0 Then
                If Not IsNumeric(decimal_num) Then
                    MsgBox ("Sale-price per Share/Unit should be a numeric value after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
            
            If Len(decimal_num) > 4 Then
                 MsgBox ("Sale-price per Share/Unit cannot exceed 4 digits after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
             End If
'====================================

Target.value = md112A.RoundToFourDigits(Target.value)
End If


'NoofShare_115AD
If Not Application.Intersect(Range("NoofShare_115AD"), Range(Target.Address)) Is Nothing Then
           'Dim int_num, decimal_num, find_num As String
        
        find_num = InStr(1, Target.value, ".")
        
        If find_num = 1 Then
            MsgBox ("No. of Shares/Units should starts with number") & Chr(13)
            Application.EnableEvents = False
            Sheet54.Unprotect Password:=getmsgstate
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
                      Application.EnableEvents = False
                      Sheet54.Unprotect Password:=getmsgstate
                      Target.ClearContents
                    GoTo endline
            End If
            
            If Len(int_num) > 14 Then
                 MsgBox ("No. of Shares/Units cannot exceed 14 digits before decimal.") & Chr(13)
                 Application.EnableEvents = False
                 Sheet54.Unprotect Password:=getmsgstate
                 Target.ClearContents
                 GoTo endline
             End If
            
            If (int_num > 99999999999999#) Then
                 MsgBox ("No. of Shares/Units cannot exceed 14 digits.") & Chr(13)
                 Application.EnableEvents = False
                 Sheet54.Unprotect Password:=getmsgstate
                 Target.ClearContents

                 GoTo endline
             End If
             
             If decimal_num > 0 Then
                If Not IsNumeric(decimal_num) Then
                    MsgBox ("No. of Shares/Units should be a numeric value after decimal.") & Chr(13)
                    Application.EnableEvents = False
                    Sheet54.Unprotect Password:=getmsgstate
                    Target.ClearContents
                    GoTo endline
            End If
            
            If Len(decimal_num) > 4 Then
                 MsgBox ("No. of Shares/Units cannot exceed 4 digits after decimal.") & Chr(13)
                 Application.EnableEvents = False
                 Sheet54.Unprotect Password:=getmsgstate
                 Target.ClearContents
                 GoTo endline
             End If
             
             End If
            
        
        Target.value = md112A.RoundToFourDigits(Target.value)
        
        End If
        
        '================================
        
'FairMarketValue_115AD
If Not Application.Intersect(Range("FairMarketValue_115AD"), Range(Target.Address)) Is Nothing Then
            
            
            find_num = InStr(1, Target.value, ".")
            
            If find_num = 1 Then
            
             MsgBox ("Fair Market Value per share/unit should starts with number") & Chr(13)
                 Application.EnableEvents = False
                 Sheet54.Unprotect Password:=getmsgstate
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
                         Application.EnableEvents = False
                         Sheet54.Unprotect Password:=getmsgstate
                         Target.ClearContents
                         GoTo endline
                End If
                
                If Len(int_num) > 14 Then
                     MsgBox ("Sale-price per Share/Unit cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                
                If (int_num > 99999999999999#) Then
                     MsgBox ("Sale-price per Share/Unit cannot exceed 14 digits.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                 
                 If decimal_num > 0 Then
                    If Not IsNumeric(decimal_num) Then
                        MsgBox ("Sale-price per Share/Unit should be a numeric value after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(decimal_num) > 4 Then
                     MsgBox ("Sale-price per Share/Unit cannot exceed 4 digits after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                     GoTo endline
                 End If
                 End If
                
            '----
            Target.value = md112A.RoundToFourDigits(Target.value)
        
        End If
            
                  '================================
                  
        'CostOfAcquisition_115AD
       If Not Application.Intersect(Range("COAwithIndex_115AD"), Range(Target.Address)) Is Nothing Then
            
            
            find_num = InStr(1, Target.value, ".")
            
            If find_num = 1 Then
            
             MsgBox ("Cost Of Acquisition should starts with number") & Chr(13)
                Application.EnableEvents = False
                Sheet54.Unprotect Password:=getmsgstate
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
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(int_num) > 14 Then
                        MsgBox ("Cost Of Acquisition cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                
                If (int_num > 99999999999999#) Then
                        MsgBox ("Cost Of Acquisition cannot exceed 14 digits.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                 
                 If decimal_num > 0 Then
                    If Not IsNumeric(decimal_num) Then
                        MsgBox ("Cost Of Acquisition should be a numeric value after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(decimal_num) > 4 Then
                        MsgBox ("Cost Of Acquisition cannot exceed 4 digits after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                 End If
                 
                 End If
                
            '----
          Target.value = md112A.RoundToFourDigits(Target.value)
        End If

        '================================
        
        'Expenditure_115AD
       If Not Application.Intersect(Range("ExpenditureWholly_115AD"), Range(Target.Address)) Is Nothing Then
            
            
            find_num = InStr(1, Target.value, ".")
            
            If find_num = 1 Then
            
             MsgBox ("Expenditure wholly and exclusively should starts with number") & Chr(13)
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
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(int_num) > 14 Then
                        MsgBox ("Expenditure wholly and exclusively cannot exceed 14 digits before decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If (int_num > 99999999999999#) Then
                        MsgBox ("Expenditure wholly and exclusively cannot exceed 14 digits.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                 
                 If decimal_num > 0 Then
                    If Not IsNumeric(decimal_num) Then
                        MsgBox ("Expenditure wholly and exclusively should be a numeric value after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                
                If Len(decimal_num) > 4 Then
                        MsgBox ("Expenditure wholly and exclusively cannot exceed 4 digits after decimal.") & Chr(13)
                        Application.EnableEvents = False
                        Sheet54.Unprotect Password:=getmsgstate
                        Target.ClearContents
                        GoTo endline
                End If
                 
                 End If
                

Target.value = md112A.RoundToFourDigits(Target.value)
End If
          

'===========end


If Not Application.Intersect(Range("ISINCode_115AD"), Range(Target.Address)) Is Nothing Then
If Target.value <> "" Then

 Target.value = UCase(Target.value)
 
 If UCase(Mid(Range(Target.Address), 1, 2)) <> "IN" Then
 'changed by Chetan C M on 27/10/2025
 'start--
' fmsgboxStatus "Please enter the valid ISIN Code in sheet 115AD(1)(iii) proviso"
fmsgboxStatus "Please enter the valid ISIN Code in Column 2 of Schedule 115AD(1)(iii)proviso"
'--end
 Target.value = ""
 End If
 
 'added by Chetan C M on 27/10/2025
 'start--
 If Not checkfieldspecialcharacter(Target.value) Then
            MsgBox ("Please don't use Special Characters < and  > " & Chr(13))
            Target.value = ""
            GoTo endline
 End If
 '--end
 
    Sheet54.Range(Target.Address).Select
    Sheet54.Range(Target.Address).Offset(0, 1) = UCase(Sheet54.Range(Target.Address).Offset(0, 22).value)
End If

If Target.value = "" Then
    Sheet54.Range(Target.Address).Offset(0, 1) = ""
End If
End If


endline:
Application.EnableEvents = True
Application.ScreenUpdating = True
Sheet54.Protect Password:=getmsgstate
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
'Change-1, ANK, 01.10.2022
'Adding the condition of 112AD lock if any one row is filled in 115AD
    Dim end_115AD As Long
    Dim range112AD As Variant
    
    setTableInfo115AD
    setTableInfo115AD1
    setTableInfo115AD2
    setTableInfo115AD3
    setTableInfo115AD4
    setTableInfo115AD5
'    setTableInfo115AD6
    setTableInfo115AD7
    setTableInfo115AD8
    
    end_115AD = WorksheetFunction.Max(end115AD, end115AD1, end115AD2, end115AD3, end115AD4, end115AD5, end115AD7, end115AD8)
    
    Application.EnableEvents = False
    Sheet53.Unprotect Password:=getmsgstate
    
     If end_115AD > 0 Then
        'call the lock112AD
        If LockFlag112A = False Then
            MsgBox ("Please fill either Schedule 112A or Schedule 115AD(1)(b)(iii) proviso as applicable in your case")
        End If
        'changed by Chetan  C M on 20/01/2026 for AY2026-27
        'start--
''Changed by Riyaz on 02/01/2025
       range112AD = "ShareAcq_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||COAwithIndex_112A||FairMarketValue_112A||ExpenditureWholly_112A"
'        range112AD = "ShareAcq_112A||ShareTra_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||COAwithIndex_112A||FairMarketValue_112A||ExpenditureWholly_112A"
        '--end
        
        LOCKRANGE112AD (range112AD)
        
    Else
        'call the unlock112AD
        If LockFlag112A = True Then
        'changed by Chetan  C M on 20/01/2026 for AY2026-27
        'start--
''Changed by Riyaz on 02/01/2025
            range112AD = "ShareAcq_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||COAwithIndex_112A||FairMarketValue_112A||ExpenditureWholly_112A"
'            range112AD = "ShareAcq_112A||ShareTra_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||COAwithIndex_112A||FairMarketValue_112A||ExpenditureWholly_112A"
        '--end
        
            UNLOCKRANGE112AD (range112AD)
        End If
        
    End If
    Application.EnableEvents = True
    Sheet53.Protect Password:=getmsgstate
'end change---



endline:
    Application.EnableEvents = True
    Sheet53.Protect Password:=getmsgstate
End Sub




