Attribute VB_Name = "Sheet24"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
Sheet24.Shapes("Button 10").Left = Range("ScheduleEI_Button1").Left + 3
Sheet24.Shapes("Button 10").Top = Range("ScheduleEI_Button1").Top + 3

Sheet24.Shapes("Button 12").Left = Range("ScheduleEI_Button2").Left + 3
Sheet24.Shapes("Button 12").Top = Range("ScheduleEI_Button2").Top + 3

Sheet24.Shapes("Button 18").Left = Range("ScheduleEI_Button3").Left + 3
Sheet24.Shapes("Button 18").Top = Range("ScheduleEI_Button3").Top + 3

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False
CalculateTaxCheckFlag = False
'Added by Shrutika(20/05/2025)
Dim i As Long
'-------------------
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword
         
          'Commented by Shrutika(20/04/2026)V0.3
'If Not Intersect(Range("EI.NatureOfIncome"), Range(Target.Address)) Is Nothing Then
'    Dim NewComment As Variant
'    NewComment = "Exemption u/s 10(12C) can only be claimed by nominees of Agniveer or Disabled Agniveer under the Agnipath Scheme"
'    If Mid(Target.value, 1, 11) = "Sec 10(12C)" Then
'         Sheet24.Range(Replace(Target.Address, "F", "H")).AddComment (NewComment)
'    Else
'        Sheet24.Range(Replace(Target.Address, "F", "H")).ClearComments
'    End If
    



If Not Application.Intersect(Sheet24.Range("Nature_Others"), Sheet24.Range(Target.Address)) Is Nothing Then
    If Target.value = "10(32)-Minor child’s income—small exemption" Then
        DPMAddComment45 Target
    ElseIf Target.value = "10(26)-Any income as referred to in section 10(26)" Then
        DPMAddComment46 Target
    ElseIf Target.value = "10(26AAA)-Any income as referred to in section 10(26AAA)" Then
        DPMAddComment47 Target
    ElseIf Target.value = "10(12A)-Any payment from the National Pension System Trust to an assessee" Then
        DPMAddComment48 Target
    ElseIf Target.value = "10(12AA)-any payment from the National Pension System Trust " Then
        DPMAddComment49 Target
    ElseIf Target.value = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee" Then
        DPMAddComment50 Target
    ElseIf Target.value = "10(12BA)-partial withdrawal made from the National Pension System" Then
        DPMAddComment51 Target
    ElseIf Target.value = "10(12C)-Agniveer Corpus Fund income" Then
        DPMAddComment52 Target
    Else
        DPMADeleteComment45 Target
    End If
End If
    
    'Commented by Shrutika(20/04/2026)V0.3
'    Dim ChangeAddress As String
'         ChangeAddress = Replace(Target.Address, "F", "G")
'    If Mid(Sheet24.Range(Target.Address).value, 1, 3) = "Any" Then
'         Sheet24.Range(ChangeAddress).Locked = False
'         Sheet24.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
'         Sheet24.Range(ChangeAddress).ClearContents
'    Else
'         Sheet24.Range(ChangeAddress).Locked = True
'         Sheet24.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
'         Sheet24.Range(ChangeAddress).ClearContents
'    End If
'End If
'------------------------------------
'Added by Shrutika(20/05/2025)
Dim s As Variant
If Not Application.Intersect(Range("SCHEI_DTAA_AmountOfIncome"), Range(Target.Address)) Is Nothing Then
    For s = 1 To Sheet24.Range("SCHEI_DTAA_AmountOfIncome").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Range("SCHEI_DTAA_AmountOfIncome").Cells(s) <> "") Then
             If Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Cells(i) <> "" Then
                   fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
              
               'Target.value = ""
                GoTo endline
             End If
        End If
      End If
    Next
End If


If Not Application.Intersect(Range("SCHEI_DTAA_NatureOfIncome"), Range(Target.Address)) Is Nothing Then
    For s = 1 To Sheet24.Range("SCHEI_DTAA_NatureOfIncome").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Range("SCHEI_DTAA_NatureOfIncome").Cells(s) <> "") Then
             If Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Cells(i) <> "" Then
                   fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = ""
                GoTo endline
             End If
        End If
      End If
    Next
End If

If Not Application.Intersect(Range("SCHEI_DTAA_CountryCode"), Range(Target.Address)) Is Nothing Then
    For s = 1 To Sheet24.Range("SCHEI_DTAA_CountryCode").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Range("SCHEI_DTAA_CountryCode").Cells(s) <> "") Then
            If Not (Target.value = "(Select)" Or Target.value = "") Then
                fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = "(Select)"
                GoTo endline
             End If
        End If
      End If
    Next
End If

If Not Application.Intersect(Range("SCHEI_DTAA_Article"), Range(Target.Address)) Is Nothing Then
    For s = 1 To Sheet24.Range("SCHEI_DTAA_Article").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Range("SCHEI_DTAA_Article").Cells(s) <> "") Then
             If Not (Target.value = "") Then
                fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = ""
                GoTo endline
             End If
        End If
      End If
    Next
End If

If Not Application.Intersect(Range("SCHEI_DTAA_HeadOfIncome"), Range(Target.Address)) Is Nothing Then
    For s = 1 To Sheet24.Range("SCHEI_DTAA_HeadOfIncome").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Range("SCHEI_DTAA_HeadOfIncome").Cells(s) <> "") Then
            If Not (Target.value = "(Select)" Or Target.value = "") Then
                fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = "(Select)"
                GoTo endline
             End If
        End If
      End If
    Next
End If

If Not Application.Intersect(Range("SCHEI_TRCObtained"), Range(Target.Address)) Is Nothing Then
    For s = 1 To Sheet24.Range("SCHEI_TRCObtained").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Range("SCHEI_TRCObtained").Cells(s) <> "") Then
            If Not (Target.value = "(Select)" Or Target.value = "") Then
                fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = "(Select)"
                GoTo endline
             End If
        End If
      End If
    Next
End If

    'Added by Shrutika(22/04/2026) v0.3
PartAGenChangeSub4 Target
'-----------------------------------------
'Commented by Shrutika(20/04/2026)V0.3
'--------------------------------------------------------------------------------------
'Newly added by Bindu as per DE & defect
' Dim nOfCount As Long
' If Not Intersect(Range("EI.NatureOfIncome"), Range(Target.Address)) Is Nothing Then
'
'        nOfCount = EfilingCommon.CountDropdown_EI(Trim(Target.value), Sheet24.Range("EI.NatureOfIncome").Cells, Sheet24.Range("EI.NatureOfIncome").count)
'        If nOfCount > 1 Then
'            fmsgboxStatus " " & Trim(Target.value) & "  cannot be selected twice " ', vbOKOnly, "ITR 3"
'            Target.value = "(Select)"
'        End If
'End If

If Not Application.Intersect(Sheet24.Range("Nature_Others"), Sheet24.Range(Target.Address)) Is Nothing Then

        If (Sheet24.Range(Target.Address).value = "Income exempt as per CBDT Circular") Or (Sheet24.Range(Target.Address).value = "Income exempt as per CBDT Notification") Or (Sheet24.Range(Target.Address).value = "Receipts not in the nature of income") Then

            Sheet24.Unprotect Password:=getmsgstate
           ' Sheet24.Range(Replace(Target.Address, "G", "H")).ClearContents
            Sheet24.Range(Replace(Target.Address, "G", "H")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet24.Range(Replace(Target.Address, "G", "H")).Locked = False
            Sheet24.Range(Replace(Target.Address, "G", "H")).value = ""
            Sheet24.Protect Password:=getmsgstate
        Else
            Sheet24.Unprotect Password:=getmsgstate
            
            'Sheet24.Range(Replace(Target.Address, "G", "H")).ClearContents
            Sheet24.Range(Replace(Target.Address, "G", "H")).Interior.Color = (&HD8D8D8)
            Sheet24.Range(Replace(Target.Address, "G", "H")).Locked = True
            Sheet24.Range(Replace(Target.Address, "G", "H")).value = ""
            Sheet24.Protect Password:=getmsgstate
   End If
            Sheet24.Unprotect Password:=getmsgstate
            Sheet24.Range(Replace(Target.Address, "G", "F")).value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.VLookup(Target.value, SheetALL.Range("Sub_Category_lookup"), 2, 0), "")
            Sheet24.Protect Password:=getmsgstate
      '  End If
End If

If Not Application.Intersect(Sheet24.Range("Others.NOI"), Sheet24.Range(Target.Address)) Is Nothing Then
If Target.value = "(Select)" Then
        Sheet24.Unprotect Password:=getmsgstate
'       Sheet24.Range(Replace(Target.Address, "F", "H")).ClearContents
        Sheet24.Range(Replace(Target.Address, "F", "H")).Interior.Color = "&HD8D8D8"
        Sheet24.Range(Replace(Target.Address, "F", "H")).Locked = True
        Sheet24.Range(Replace(Target.Address, "F", "H")).value = ""
        Sheet24.Protect Password:=getmsgstate
End If
End If
   Sheet24.Protect Password:=sPassword

endline:
Application.EnableEvents = True

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

'Added by Shrutika- 24/04/2026 V0.3 AY-26-27



Sub PartAGenChangeSub4(ByVal Target As Range)
On Error GoTo endline
Dim validateType As String
Application.EnableEvents = False
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

If Not Application.Intersect(Sheet24.Range("Others.NOI"), Sheet24.Range(Target.Address)) Is Nothing Then
    If SheetALL.Range("bacValue").value = 1 Then
               Sheet24.Unprotect Password:=sPassword
                Application.EnableEvents = False
                
                Dim Sub_Cat As Variant
                Sub_Cat = Target.Address
                Sub_Cat = Replace(Sub_Cat, "$", "")
                Sub_Cat = Replace(Sub_Cat, "F", "G")
                
            If Target.value = "Agricultural  & related incomes" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Agri_dropdown"
            ElseIf Target.value = "Compensation/other sums received by government or other approved entities" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Comp_dropdown"
            ElseIf Target.value = "Income from specified Investments" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Inco_dropdown"
            ElseIf Target.value = "Specified sums received by armed forces personnel" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Spec__dropdown"
            ElseIf Target.value = "Sums received by Senior Citizens/Minors" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Seminew_dropdown"
            ElseIf Target.value = "Sums received by specified Category of Taxpayers" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Taxp_dropdown"
            ElseIf Target.value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Yojana_dropdown"
'Konda updated on 24-06-2026----SIT-123441
'            ElseIf Target.value = "Other Incomes" Then
'                Sheet24.Range(Sub_Cat).Validation.Delete
'                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"

                ElseIf Target.value = "Other Incomes" And UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) <> "RES" Then
                    Sheet24.Range(Sub_Cat).Validation.Delete
                    Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
                
                ElseIf Target.value = "Other Incomes" And UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) = "RES" Then
                    Sheet24.Range(Sub_Cat).Validation.Delete
                    Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown_RES"
'=========================on 24-06-2026----SIT-123441
           ElseIf Target.value = "Other Exempt Income for Non Residents" Then
                Sheet24.Range(Sub_Cat).Validation.Delete
                Sheet24.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OthExempt_dropdown"
           ElseIf Target.value = "(Select)" Or Target.value = "" Then
                Sheet24.Range(Sub_Cat).Validation.Delete    'Uncommented by Konda 26-06-2026
                Sheet24.Range(Sub_Cat).value = ""
            
            End If
        Else
           Sheet24.Unprotect Password:=sPassword
            Application.EnableEvents = False
            
            Dim Sub_Cat1 As Variant
            Sub_Cat1 = Target.Address
            Sub_Cat1 = Replace(Sub_Cat1, "$", "")
            Sub_Cat1 = Replace(Sub_Cat1, "F", "G")
            
        If Target.value = "Agricultural  & related incomes" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Agri_dropdown"
        ElseIf Target.value = "Compensation/other sums received by government or other approved entities" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Comp_dropdown"
        ElseIf Target.value = "Income from specified Investments" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Inco_dropdown"
        ElseIf Target.value = "Specified sums received by armed forces personnel" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Spec__dropdown"
        ElseIf Target.value = "Sums received by Senior Citizens/Minors" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Semiold_dropdown"
        ElseIf Target.value = "Sums received by specified Category of Taxpayers" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Taxp_dropdown"
        ElseIf Target.value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Yojana_dropdown"
'Konda updated on 24-06-2026----SIT-123441
'        ElseIf Target.value = "Other Incomes" Then
'            Sheet24.Range(Sub_Cat1).Validation.Delete
'            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
            
             ElseIf Target.value = "Other Incomes" And UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) <> "RES" Then
                    Sheet24.Range(Sub_Cat1).Validation.Delete
                    Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
             
             ElseIf Target.value = "Other Incomes" And UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) = "RES" Then
                    Sheet24.Range(Sub_Cat1).Validation.Delete
                    Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown_RES"
'================================
        
        ElseIf Target.value = "Other Exempt Income for Non Residents" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete
            Sheet24.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OthExempt_dropdown"
        ElseIf Target.value = "(Select)" Or Target.value = "" Then
            Sheet24.Range(Sub_Cat1).Validation.Delete   'Uncommented by Konda 26-06-2026
            Sheet24.Range(Sub_Cat1).value = ""
        End If
               Sheet24.Protect Password:=sPassword
            Application.EnableEvents = True
    End If
End If

' If Not Application.Intersect(Sheet24.Range("Nature_Others"), Sheet24.Range(Target.Address)) Is Nothing Then
'        If 1 < Application.CountIf(Sheet24.Range("Nature_Others"), Target.value) Then
''Konda updated on 22-06-2026--SIT-123147
'            If (Range(Target.Address).value <> "Income exempt as per CBDT Circular") Then
'            If (Range(Target.Address).value <> "Income exempt as per CBDT Notification") Then
'            If (Range(Target.Address).value <> "Receipts not in the nature of income") Then
'             fmsgbox "* Same drop downs cannot be selected twice"
'             Range(Target.Address).value = "(Select)"
'             End If
'             End If
'        End If
'        End If
'End If

If Not Application.Intersect(Sheet24.Range("Nature_Others"), Sheet24.Range(Target.Address)) Is Nothing Then
    If WorksheetFunction.CountIf(Sheet24.Range("Nature_Others"), Target.value) > 1 Then
        Application.EnableEvents = False

            If Not (Target.value = "Income exempt as per CBDT Circular" Or Target.value = "Income exempt as per CBDT Notification" Or Target.value = "Receipts not in the nature of income") Then
           fmsgbox "* Same drop downs cannot be selected twice"
            Range(Target.Address).value = "(Select)"
            End If
           
        Application.EnableEvents = True
    End If

End If





   Sheet24.Protect Password:=sPassword
endline:
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

