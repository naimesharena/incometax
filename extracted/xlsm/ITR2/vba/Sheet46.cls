Attribute VB_Name = "Sheet46"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_Deactivate()
sPassword = getmsgstate
Sheet46.Protect Password:=sPassword
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False

If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If
'Chandru - Year change

If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing"), Sheet46.Range(Target.Address)) Is Nothing Then
Sheet46.Unprotect Password:=getmsgstate
'Malli------------

        If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not FormatNCheckDate(Trim(Target.value)) Then
            MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
            Target.ClearContents
        End If
        End If

        If Not CheckDateddmmyyyy(Target.value) Then
              MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
        Target.value = ""
        End If

'------------------
'Commeted by Riyaz on 24/12/2024
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2023") Then 'dpk1601
'        fmsgbox ("The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP")
        If Not Target.value = "" Then
        
        'changed by Chetan C M on 20/01/2026 for AY 2026-27
        'start--
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2024") Then 'dpk1601
'        'fmsgbox ("The date should not be before 01/04/2024 for A.Y. 2025-26 In schedule ESOP")
'        fmsgbox ("Date should not be before 01/04/2024")
        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2025") Then 'dpk1601
        fmsgbox ("Date should not be before 01/04/2025")
        '--end
        
        Sheet46.Range(Target.Address).value = ""
        GoTo endline
        End If
        End If
Sheet46.Protect Password:=getmsgstate

Sheet46.Unprotect Password:=getmsgstate
        Dim DateOfCeasing As Variant
        VerificationDate = Sheet52.Range("sheet9.Date").value
        DateOfCeasing = Sheet46.Range("ESOP.DateOfCeasing").value
        'If Dformat(DateOfCeasing, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
        If Not ChkMaxInclusiveDate(Dformat(DateOfCeasing, ""), Dformat(VerificationDate, "")) Then
            fmsgbox ("Date of Ceasing cannot be after System Date in Schedule ESOP")
            Sheet46.Range(Target.Address).value = ""
            GoTo endline
        End If
        
'        If Not ChkMaxDOBDate(Sheet46.Range(Target.Address).value, "31/03/2022") Then
'                fmsgbox ("Date of Ceasing should not be after 31/03/2022 for AY 2022-23 In schedule ESOP ")
'                Target.value = ""
'                GoTo endline
'        End If
        
Sheet46.Protect Password:=getmsgstate

End If

If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee1"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If
'Chandru - Year Change
If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing1"), Sheet46.Range(Target.Address)) Is Nothing Then
        Sheet46.Unprotect Password:=getmsgstate
'Commented by Riyaz on 24/12/2024
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2023") Then
'        fmsgbox ("The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP")
        If Not Target.value = "" Then
        
        'changed by Chetan C M on 20/01/2026 for AY 2026-27
        'start--
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2024") Then
'        'fmsgbox ("The date should not be before 01/04/2024 for A.Y. 2025-26 In schedule ESOP")
'        fmsgbox ("Date should not be before 01/04/2024")
        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2025") Then
        fmsgbox ("Date should not be before 01/04/2025")
        '-end
        
        Sheet46.Range(Target.Address).value = ""
        End If
        End If
        Sheet46.Protect Password:=getmsgstate

        Sheet46.Unprotect Password:=getmsgstate
        'Dim DateOfCeasing As Variant
        VerificationDate = Sheet52.Range("sheet9.Date").value
        DateOfCeasing = Sheet46.Range("ESOP.DateOfCeasing1").value
        If Not ChkMaxInclusiveDate(Dformat(DateOfCeasing, ""), Dformat(VerificationDate, "")) Then
            fmsgbox ("Date of Ceasing cannot be after System Date in Schedule ESOP")
            Sheet46.Range(Target.Address).value = ""
        End If
        
        
Sheet46.Protect Password:=getmsgstate

End If
'Chandru - New Row added
If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee2"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If

If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing2"), Sheet46.Range(Target.Address)) Is Nothing Then
        Sheet46.Unprotect Password:=getmsgstate
'Commented by Riyaz on 24/12/2024
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2023") Then
'        fmsgbox ("The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP")
        If Not Target.value = "" Then
        
        'changed by Chetan C M on 20/01/2026 for AY 2026-27
        'start--
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2024") Then
'        'fmsgbox ("The date should not be before 01/04/2024 for A.Y. 2025-26 In schedule ESOP")
'        fmsgbox ("Date should not be before 01/04/2024")
        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2025") Then
        fmsgbox ("Date should not be before 01/04/2025")
        '--end
        
        Sheet46.Range(Target.Address).value = ""
        End If
        End If
        Sheet46.Protect Password:=getmsgstate

        Sheet46.Unprotect Password:=getmsgstate
        'Dim DateOfCeasing As Variant
        VerificationDate = Sheet52.Range("sheet9.Date").value
        DateOfCeasing = Sheet46.Range("ESOP.DateOfCeasing2").value
        If Not ChkMaxInclusiveDate(Dformat(DateOfCeasing, ""), Dformat(VerificationDate, "")) Then
            fmsgbox ("Date of Ceasing cannot be after System Date in Schedule ESOP")
            Sheet46.Range(Target.Address).value = ""
        End If
        
        
Sheet46.Protect Password:=getmsgstate

End If

'Newly Added by Riyaz on 24/12/2024
If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee3"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If

    If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing3"), Sheet46.Range(Target.Address)) Is Nothing Then
        Sheet46.Unprotect Password:=getmsgstate
        If Not Target.value = "" Then
        
        'changed by Chetan C M on 20/01/2026 for AY 2026-27
        'start--
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2024") Then
'        'fmsgbox ("The date should not be before 01/04/2024 for A.Y. 2025-26 In schedule ESOP")
'        fmsgbox ("Date should not be before 01/04/2024")
        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2025") Then
        fmsgbox ("Date should not be before 01/04/2025")
        '--end
        
        Sheet46.Range(Target.Address).value = ""
        End If
        End If
        Sheet46.Protect Password:=getmsgstate

        Sheet46.Unprotect Password:=getmsgstate
        'Dim DateOfCeasing As Variant
        VerificationDate = Sheet52.Range("sheet9.Date").value
        DateOfCeasing = Sheet46.Range("ESOP.DateOfCeasing3").value
        If Not ChkMaxInclusiveDate(Dformat(DateOfCeasing, ""), Dformat(VerificationDate, "")) Then
            fmsgbox ("Date of Ceasing cannot be after System Date in Schedule ESOP")
            Sheet46.Range(Target.Address).value = ""
        End If
        
        
Sheet46.Protect Password:=getmsgstate

End If


'End change


'added by Chetan C M on 20/01/2026 for AY 2026-27
'start--
If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee4"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing4"), Sheet46.Range(Target.Address)) Is Nothing Then
        Sheet46.Unprotect Password:=getmsgstate
        If Not Target.value = "" Then
            If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2025") Then
                fmsgbox ("Date should not be before 01/04/2025")
                Sheet46.Range(Target.Address).value = ""
            End If
        End If
        Sheet46.Protect Password:=getmsgstate
        Sheet46.Unprotect Password:=getmsgstate
        VerificationDate = Sheet52.Range("sheet9.Date").value
        DateOfCeasing = Sheet46.Range("ESOP.DateOfCeasing4").value
        If Not ChkMaxInclusiveDate(Dformat(DateOfCeasing, ""), Dformat(VerificationDate, "")) Then
            fmsgbox ("Date of Ceasing cannot be after System Date in Schedule ESOP")
            Sheet46.Range(Target.Address).value = ""
        End If
        Sheet46.Protect Password:=getmsgstate
End If
'--end

'Chandru - Year Change done
Dim i As Long
    For i = 1 To Sheet46.Range("ESOP.Date").count
        If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then
'            If ChkMaxDOBDateDef(Target.value, "31/03/2023") Then 'dpk1601
'Changed by Riyaz on 24/12/2024
             If Not Target.value = "" Then
             
    'changed by Chetan C M on 20/01/2026 for AY 2026-27
    'start--
'             If ChkMaxDOBDateDef(Target.value, "31/03/2024") Then
'                'fmsgbox ("The date should not be before 01/04/2025 for A.Y. 2025-26 In schedule ESOP") 'dpk1601
'                fmsgbox ("Date should not be before 01/04/2024")
                If ChkMaxDOBDateDef(Target.value, "31/03/2025") Then
                fmsgbox ("Date should not be before 01/04/2025")
    '--end
    
                Target.value = ""
                GoTo endline
            End If
            End If
        VerificationDate = Sheet52.Range("sheet9.Date").value
        
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
        fmsgbox (" Date cannot be after System Date in Schedule ESOP")
         Target.value = ""
        GoTo endline
        End If
        
        
        End If
        
Next



'Newly added by Bindu
If Not Intersect(Sheet46.Range("ESOP.PAN"), Range(Target.Address)) Is Nothing Then

Dim TargetPAN As String
               TargetPAN = ""
               TargetPAN = Target.name.name
                Sheet46.Range("ESOP.PAN").value = UCase(Range(TargetPAN).value)
                If Not Part_A_General.CheckPAN(UCase(Trim(Range(TargetPAN).value))) Then
                     fmsgboxStatus "Invalid PAN in 'PAN of the employer being an eligible startup'. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."
                     Sheet46.Range(TargetPAN).Select
                     Sheet46.Range(TargetPAN).value = ""
                 
                End If
End If


If Not Intersect(Sheet46.Range("ESOP.DPIIT"), Range(Target.Address)) Is Nothing Then
Dim TargetDPIIT
        TargetDPIIT = ""
        TargetDPIIT = Target.name.name
        Sheet46.Range("ESOP.DPIIT").value = UCase(Range(TargetDPIIT).value)
        If Not Part_A_General.ChkDPIIT(UCase(Trim(Range(TargetDPIIT).value))) Then
             fmsgboxStatus "Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"
             Sheet46.Range(TargetDPIIT).Select
             Sheet46.Range(TargetDPIIT).value = ""
        End If
        If Not checkallfieldspecialcharacter_80U_80DD(UCase(Trim(Range(TargetDPIIT).value))) Then
             fmsgboxStatus "Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"
             Sheet46.Range(TargetDPIIT).Select
             Sheet46.Range(TargetDPIIT).value = ""

        End If
End If



endline:
Sheet46.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub








