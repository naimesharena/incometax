Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False


fmsgboxStatus "Please verify the TAN details and the amount of credit available in your form 16 and 26 AS."

 Sheet4.Unprotect Password:=getmsgstate

    ActiveSheet.Shapes("Button 1").Left = Range("SchTDS.Button1").Left + 2
    ActiveSheet.Shapes("Button 1").Top = Range("SchTDS.Button1").Top + 2

    ActiveSheet.Shapes("Button 725").Left = Range("SchTDS.Button2").Left + 2
    ActiveSheet.Shapes("Button 725").Top = Range("SchTDS.Button2").Top + 2


   ActiveSheet.Shapes("Button 727").Left = Range("SchTDS.Button5").Left + 2
    ActiveSheet.Shapes("Button 727").Top = Range("SchTDS.Button5").Top + 2


'    ActiveSheet.Shapes("Button 2").Left = Range("SchTDS.Button2").Left
'    ActiveSheet.Shapes("Button 2").Top = Range("SchTDS.Button2").Top
'
'    ActiveSheet.Shapes("Button 3").Left = Range("SchTDS.Button3").Left
'    ActiveSheet.Shapes("Button 3").Top = Range("SchTDS.Button3").Top
'
'    ActiveSheet.Shapes("Button 4").Left = Range("SchTDS.Button4").Left
'    ActiveSheet.Shapes("Button 4").Top = Range("SchTDS.Button4").Top

        ActiveSheet.Shapes("Button 663").Left = Range("SchTDS2.Button").Left
        ActiveSheet.Shapes("Button 663").Top = Range("SchTDS2.Button").Top

        ActiveSheet.Shapes("Button 665").Left = Range("SchTDS2ii.Button").Left
        ActiveSheet.Shapes("Button 665").Top = Range("SchTDS2ii.Button").Top

        LockingTDS2
        LockingTDS3

endline:
 Sheet4.Protect Password:=getmsgstate

Application.EnableEvents = True
End Sub

'Private Sub Worksheet_Change(ByVal Target As Range)
'On Error GoTo Endline
'Application.EnableEvents = False
'
'
'CalculateTaxCheckFlag = False
'
'
'
'
'If Not checkfieldSuperSpecialcharacter(Target.value) Then
'         fmsgbox ("* Please do not enter Special characters < , >" & Chr(13))
'        Target.value = ""
'End If
'
'
' 'If Target.Validation.Type = 3 Then
'
'  If Not Application.Intersect(Sheet4.Range("TDSal.TAN"), Sheet4.Range(Target.Address)) Is Nothing Then
'       Target.value = UCase(Target.value)
' End If
'
' If Not Application.Intersect(Sheet4.Range("TDsOthr.TAN"), Sheet4.Range(Target.Address)) Is Nothing Then
'       Target.value = UCase(Target.value)
' End If
'
' If Not Application.Intersect(Sheet4.Range("TDsOthr2.PAN"), Sheet4.Range(Target.Address)) Is Nothing Then
'       Target.value = UCase(Target.value)
' End If
'
'If Target.Validation.Type = 6 Then
'
'If Not Application.Intersect(Sheet4.Range("TDsOthr2.AADHAAR"), Sheet4.Range(Target.Address)) Is Nothing Then
'    If Len(Sheet4.Range(Target.Address).value) >= 12 Then
'                    If chkNumeric(Sheet4.Range(Target.Address).value) Then
'                       If Sheet4.Range(Target.Address).value <> "000000000000" Then
'                         Sheet4.Range(Target.Address).value = Sheet4.Range(Target.Address).value
'                       Else
'                         fmsgboxStatus "* Invalid Aadhaar Number"
'                         Sheet4.Range(Target.Address).value = ""
'                         Sheet4.Range("TDsOthr2.AADHAAR").Select
'                       End If
'
'
'                       If Sheet4.Range(Target.Address).value < 0 Then
'                            fmsgboxStatus "* Invalid Aadhaar Number"
'                         Sheet4.Range(Target.Address).value = ""
'                         Sheet4.Range("TDsOthr2.AADHAAR").Select
'                        End If
'
'
'                       If Sheet4.Range(Target.Address).value <> "111111111111" Then
'                         Sheet4.Range(Target.Address).value = Sheet4.Range(Target.Address).value
'                       Else
'                         fmsgboxStatus "* Invalid Aadhaar Number"
'                         Sheet4.Range(Target.Address).value = ""
'                         Sheet4.Range("TDsOthr2.AADHAAR").Select
'                       End If
'
'                    Else
'                      fmsgboxStatus "* Invalid Aadhaar Number"
'                      Sheet4.Range(Target.Address).value = ""
'                      Sheet4.Range("TDsOthr2.AADHAAR").Select
'                    End If
'                Else
'                    Sheet4.Range(Target.Address).value = ""
'                End If
'                End If
'End If
'
''----------------Unlock Password-------------------START---
'   Sheet4.Unprotect Password:=getmsgstate
''----------------Unlock Password-------------------END-----
'
'If Not Intersect(Range("TDsOthr.tdsdeducted"), Range(Target.Address)) Is Nothing Then
'Sheet4.Unprotect Password:=getmsgstate
'        If Sheet4.Range(Target.Address).value = 0 Then
'
'            Sheet4.Range(Replace(Target.Address, "I", "G")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "I", "G")).Interior.Color = "&HCCFFCC"
'            Sheet4.Range(Replace(Target.Address, "I", "G")).value = "(Select)"
'
'            Sheet4.Range(Replace(Target.Address, "I", "H")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "I", "H")).Interior.Color = "&HCCFFCC"
'            Sheet4.Protect Password:=getmsgstate
'        Else
'        Sheet4.Unprotect Password:=getmsgstate
'            Sheet4.Range(Replace(Target.Address, "I", "G")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "I", "G")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "I", "G")).Interior.Color = "&HD8D8D8"
'
'            Sheet4.Range(Replace(Target.Address, "I", "H")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "I", "H")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "I", "H")).Interior.Color = "&HD8D8D8"
'            Sheet4.Protect Password:=getmsgstate
'
'    End If
'End If
'
'If Not Intersect(Range("TDsOthr.tds"), Range(Target.Address)) Is Nothing Then
'Sheet4.Unprotect Password:=getmsgstate
'        If Sheet4.Range(Target.Address).value = 0 Then
'
'            Sheet4.Range(Replace(Target.Address, "H", "I")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HCCFFCC"
'            Sheet4.Protect Password:=getmsgstate
'        Else
'Sheet4.Unprotect Password:=getmsgstate
'            Sheet4.Range(Replace(Target.Address, "H", "I")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "H", "I")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HD8D8D8"
'            Sheet4.Protect Password:=getmsgstate
'
'    End If
'End If
'
'If Not Intersect(Range("TDsOthr2.tdsdeducted"), Range(Target.Address)) Is Nothing Then
'Sheet4.Unprotect Password:=getmsgstate
'        If Sheet4.Range(Target.Address).value = 0 Then
'
'            Sheet4.Range(Replace(Target.Address, "J", "H")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "J", "H")).Interior.Color = "&HCCFFCC"
'            Sheet4.Range(Replace(Target.Address, "J", "H")).value = "(Select)"
'
'            Sheet4.Range(Replace(Target.Address, "J", "I")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "J", "I")).Interior.Color = "&HCCFFCC"
'            Sheet4.Protect Password:=getmsgstate
'        Else
'        Sheet4.Unprotect Password:=getmsgstate
'
'            Sheet4.Range(Replace(Target.Address, "J", "H")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "J", "H")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "J", "H")).Interior.Color = "&HD8D8D8"
'
'            Sheet4.Range(Replace(Target.Address, "J", "I")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "J", "I")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "J", "I")).Interior.Color = "&HD8D8D8"
'Sheet4.Protect Password:=getmsgstate
'    End If
'End If
'
'If Not Intersect(Range("TDsOthr2.tds"), Range(Target.Address)) Is Nothing Then
'Sheet4.Unprotect Password:=getmsgstate
'        If Sheet4.Range(Target.Address).value = 0 Then
'
'            Sheet4.Range(Replace(Target.Address, "I", "J")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
'            Sheet4.Protect Password:=getmsgstate
'        Else
'Sheet4.Unprotect Password:=getmsgstate
'            Sheet4.Range(Replace(Target.Address, "I", "J")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "I", "J")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
'Sheet4.Protect Password:=getmsgstate
'    End If
'End If
'
'
'If Not Intersect(Range("TDsOthr.tdsclaimed"), Range(Target.Address)) Is Nothing Then
'Sheet4.Unprotect Password:=getmsgstate
'        If Sheet4.Range(Target.Address).value > 0 Then
'            Sheet4.Range(Replace(Target.Address, "J", "k")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "J", "k")).Interior.Color = "&HCCFFCC"
'
'            Sheet4.Range(Replace(Target.Address, "J", "L")).Locked = False
'            Sheet4.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HCCFFCC"
'            Sheet4.Range(Replace(Target.Address, "J", "L")).value = "(Select)"
'            Sheet4.Protect Password:=getmsgstate
'        Else
'        Sheet4.Unprotect Password:=getmsgstate
'            Sheet4.Range(Replace(Target.Address, "J", "k")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "J", "k")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "J", "k")).Interior.Color = "&HD8D8D8"
'
'            Sheet4.Range(Replace(Target.Address, "J", "L")).ClearContents
'            Sheet4.Range(Replace(Target.Address, "J", "L")).Locked = True
'            Sheet4.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HD8D8D8"
'Sheet4.Protect Password:=getmsgstate
'
'    End If
'End If
'
''If Not Intersect(Range("TDsOthr.headincome"), Range(Target.Address)) Is Nothing Then
''    If UCase(Mid(Sheet4.Range(Target.Address).value, 1, 1)) = "N" Then
''        Sheet4.Range(Replace(Target.Address, "K", "J")).value = 0
''        Sheet4.Range(Replace(Target.Address, "K", "J")).Locked = True
''        Sheet4.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HD8D8D8"
''     Else
''        Sheet4.Range(Replace(Target.Address, "K", "J")).ClearContents
''        Sheet4.Range(Replace(Target.Address, "K", "J")).Locked = False
''        Sheet4.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HCCFFCC"
''    End If
''End If
'
''''''''''''''''''''''''''''''''''''''Malli
'''''''''''''''''''TDS_C1 AY_2024_25 change
'Sheet4.Unprotect Password:=getmsgstate
'
'If Target.Column = 6 Then
'        If Not Target.value = "(Select)" Then
'        If 1 < Application.CountIf(Range("TDsOthr.SectionTDSDeducted"), Target.value) Then
'             MsgBox "* Sch TDS2(i)-Section under which TDS is deducted in Same drop downs cannot be selected more than once in Schedule TDS"
'             Range(Target.Address).value = "(Select)"
'        End If
'        End If
'End If
'
'
''''''TDS_C4 AY_2024_25 change
'If Target.Column = 7 Then
'        If Not Target.value = "(Select)" Then
'        If 1 < Application.CountIf(Range("TDsOthr2.SectionTDSDeducted"), Target.value) Then
'             MsgBox "* Sch TDS2(ii)-Section under which TDS is deducted in Same drop downs cannot be selected more than once in Schedule TDS"
'             Range(Target.Address).value = "(Select)"
'        End If
'        End If
'End If
'
''''''''''''''''''''''''''''''''''''''''Malli
'
'
'Endline:
'Sheet4.Protect Password:=getmsgstate
'Application.EnableEvents = True
'
'
'End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False


CalculateTaxCheckFlag = False

If Not checkfieldSuperSpecialcharacter(Target.value) Then
         fmsgbox ("* Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
End If

'Ayush_26-27_TDS_C8
 If Not Application.Intersect(Sheet4.Range("TDSal.IncChrgSal"), Sheet4.Range(Target.Address)) Is Nothing Then
    If Sheet5.Range("grossSalary_exempt_familyPension").value < Sheet4.Range("TotIncSal_TDSHidden").value Then

        fmsgbox ("""Sum of ""Gross Salary"" , ""Family Pension"" and ""exempt incomes which are part of salary"" declared is less than the amount of Gross Salary declared in Schedule TDS 1.""")
'    fmsgbox ("""Nature of Exempt Allowance(s) will get refreshed if the field ""Salary as per Section 17(1)"" is modified.""")
'    Target.ClearContents
    End If
 End If

 'If Target.Validation.Type = 3 Then

  If Not Application.Intersect(Sheet4.Range("TDSal.TAN"), Sheet4.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet4.Range("TDsOthr.TAN"), Sheet4.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet4.Range("TDsOthr2.PAN"), Sheet4.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

If Target.Validation.Type = 6 Then

If Not Application.Intersect(Sheet4.Range("TDsOthr2.AADHAAR"), Sheet4.Range(Target.Address)) Is Nothing Then
    If Len(Sheet4.Range(Target.Address).value) >= 12 Then
            If chkNumeric(Sheet4.Range(Target.Address).value) Then
               If Sheet4.Range(Target.Address).value <> "000000000000" Then
                 Sheet4.Range(Target.Address).value = Sheet4.Range(Target.Address).value
               Else
                 fmsgboxStatus "* Invalid Aadhaar Number"
                 Sheet4.Range(Target.Address).value = ""
                 Sheet4.Range("TDsOthr2.AADHAAR").Select
               End If


               If Sheet4.Range(Target.Address).value < 0 Then
                    fmsgboxStatus "* Invalid Aadhaar Number"
                 Sheet4.Range(Target.Address).value = ""
                 Sheet4.Range("TDsOthr2.AADHAAR").Select
                End If


               If Sheet4.Range(Target.Address).value <> "111111111111" Then
                 Sheet4.Range(Target.Address).value = Sheet4.Range(Target.Address).value
               Else
                 fmsgboxStatus "* Invalid Aadhaar Number"
                 Sheet4.Range(Target.Address).value = ""
                 Sheet4.Range("TDsOthr2.AADHAAR").Select
               End If

            Else
              fmsgboxStatus "* Invalid Aadhaar Number"
              Sheet4.Range(Target.Address).value = ""
              Sheet4.Range("TDsOthr2.AADHAAR").Select
            End If
        Else
            Sheet4.Range(Target.Address).value = ""
        End If
    End If
End If

'----------------Unlock Password-------------------START---
   Sheet4.Unprotect Password:=getmsgstate
'----------------Unlock Password-------------------END-----

If Not Intersect(Range("TDsOthr.tdsdeducted"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).value = 0 Then

            Sheet4.Range(Replace(Target.Address, "I", "G")).Locked = False
            Sheet4.Range(Replace(Target.Address, "I", "G")).Interior.Color = "&HCCFFCC"
            Sheet4.Range(Replace(Target.Address, "I", "G")).value = "(Select)"
            
            Sheet4.Range(Replace(Target.Address, "I", "H")).Locked = False
            Sheet4.Range(Replace(Target.Address, "I", "H")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "I", "G")).ClearContents
            Sheet4.Range(Replace(Target.Address, "I", "G")).Locked = True
            Sheet4.Range(Replace(Target.Address, "I", "G")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "I", "H")).ClearContents
            Sheet4.Range(Replace(Target.Address, "I", "H")).Locked = True
            Sheet4.Range(Replace(Target.Address, "I", "H")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If

If Not Intersect(Range("TDsOthr.tds"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).value = 0 Then
                    
            Sheet4.Range(Replace(Target.Address, "H", "I")).Locked = False
            Sheet4.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "H", "I")).ClearContents
            Sheet4.Range(Replace(Target.Address, "H", "I")).Locked = True
            Sheet4.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If

'================================================================================================================

Dim TDS_Section_warning1, TDS_Section_warning2, TDS_section_error As Range
Dim TDS_Section_warning_count1, TDS_Section_warning_count2, TDS_section_error_count, iCount_warning1, iCount_warning2, iCount_error As Long

'Newly added by sai on 24th April As per DE V6
If Not Intersect(Range("TDsOthr2.SectionTDSDeducted"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("TDsOthr.SectionTDS"), Range(Target.Address)) Is Nothing Then

     If Target.value <> "" Or UCase(Target.value) <> UCase("(Select)") Then
       Set TDS_Section_warning1 = Sheet8.Range("TDS_Section_Warning1").Cells
       TDS_Section_warning_count1 = Sheet8.Range("TDS_Section_warning1").count
       
       For iCount_warning1 = 1 To TDS_Section_warning_count1
         
         If UCase(Trim(Target.value)) = UCase(Trim(TDS_Section_warning1.item(iCount_warning1).value)) Then
            MsgBox "*Incomes chargeable to tax at special rates are not expected to be reported in ITR  4. Accordingly, TDS relating to the special rates income namely 194B,194BB,194BA,194IA,194IC,194LA,194S are also not expected to be claimed. Thus, please ensure that the return being filed in ITR 4 is not defective.", vbInformation
          End If
       Next
       
       Set TDS_Section_warning2 = Sheet8.Range("TDS_Section_warning2").Cells
       TDS_Section_warning_count2 = Sheet8.Range("TDS_Section_warning2").count
       
       For iCount_warning2 = 1 To TDS_Section_warning_count2
         
         If UCase(Trim(Target.value)) = UCase(Trim(TDS_Section_warning2.item(iCount_warning2).value)) Then
            MsgBox "*The Income on which the TDS under this section has been deducted is payable to a non-resident.Returns for income of non-residents are required to be filed in ITR 3/5. Thus, kindly ensure that the return being filed in ITR 4 is not defective.", vbInformation
          End If
       Next
       
    
       Set TDS_section_error = Sheet8.Range("TDS_Section_Error").Cells
       TDS_section_error_count = Sheet8.Range("TDS_Section_Error").count
       
        For iCount_error = 1 To TDS_section_error_count
       
          If UCase(Trim(Target.value)) = UCase(Trim(TDS_section_error.item(iCount_error).value)) Then
            fmsgbox "*TDS relating to salaries is required to be claimed in Schedule TDS 1."
            Target.value = "(Select)"
          End If
       
        Next
     
     
     End If

End If

'==========================================================================================================

If Not Intersect(Range("TDsOthr2.tdsdeducted"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).value = 0 Then

            Sheet4.Range(Replace(Target.Address, "j", "H")).Locked = False
            Sheet4.Range(Replace(Target.Address, "j", "H")).Interior.Color = "&HCCFFCC"
'            Sheet4.Range(Replace(Target.Address, "j", "H")).value = "(Select)"
            
            Sheet4.Range(Replace(Target.Address, "J", "I")).Locked = False
            Sheet4.Range(Replace(Target.Address, "J", "I")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
        
            Sheet4.Range(Replace(Target.Address, "J", "H")).ClearContents
            Sheet4.Range(Replace(Target.Address, "J", "H")).Locked = True
            Sheet4.Range(Replace(Target.Address, "J", "H")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "J", "I")).ClearContents
            Sheet4.Range(Replace(Target.Address, "J", "I")).Locked = True
            Sheet4.Range(Replace(Target.Address, "J", "I")).Interior.Color = "&HD8D8D8"
Sheet4.Protect Password:=getmsgstate
    End If
End If

If Not Intersect(Range("TDsOthr2.tds"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).value = 0 Then
                    
            Sheet4.Range(Replace(Target.Address, "I", "J")).Locked = False
            Sheet4.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "I", "J")).ClearContents
            Sheet4.Range(Replace(Target.Address, "I", "J")).Locked = True
            Sheet4.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
Sheet4.Protect Password:=getmsgstate
    End If
End If


'TDsOthr2.tdsclaimed


If Not Intersect(Range("TDsOthr.tdsclaimed"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).value > 0 Then
            Sheet4.Range(Replace(Target.Address, "J", "K")).Locked = False
            Sheet4.Range(Replace(Target.Address, "J", "K")).Interior.Color = "&HCCFFCC"

            Sheet4.Range(Replace(Target.Address, "J", "L")).Locked = False
            Sheet4.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HCCFFCC"
            Sheet4.Range(Replace(Target.Address, "J", "L")).value = "(Select)"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "J", "K")).ClearContents
            Sheet4.Range(Replace(Target.Address, "J", "K")).Locked = True
            Sheet4.Range(Replace(Target.Address, "J", "K")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "J", "L")).ClearContents
            Sheet4.Range(Replace(Target.Address, "J", "L")).Locked = True
            Sheet4.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HD8D8D8"
Sheet4.Protect Password:=getmsgstate

    End If
End If



'Newly added by Sai For SIT-67591 and Fixed

  If Not Intersect(Range("TDsOthr2.tdsclaimed"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).value > 0 Then
            Sheet4.Range(Replace(Target.Address, "K", "L")).Locked = False
            Sheet4.Range(Replace(Target.Address, "K", "L")).Interior.Color = "&HCCFFCC"

            Sheet4.Range(Replace(Target.Address, "K", "M")).Locked = False
            Sheet4.Range(Replace(Target.Address, "K", "M")).Interior.Color = "&HCCFFCC"
            Sheet4.Range(Replace(Target.Address, "K", "M")).value = "(Select)"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "K", "L")).ClearContents
            Sheet4.Range(Replace(Target.Address, "K", "L")).Locked = True
            Sheet4.Range(Replace(Target.Address, "K", "L")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "K", "M")).ClearContents
            Sheet4.Range(Replace(Target.Address, "K", "M")).Locked = True
            Sheet4.Range(Replace(Target.Address, "K", "M")).Interior.Color = "&HD8D8D8"
Sheet4.Protect Password:=getmsgstate

    End If
End If























'If Not Intersect(Range("TDsOthr.headincome"), Range(Target.Address)) Is Nothing Then
'    If UCase(Mid(Sheet4.Range(Target.Address).value, 1, 1)) = "N" Then
'        Sheet4.Range(Replace(Target.Address, "K", "J")).value = 0
'        Sheet4.Range(Replace(Target.Address, "K", "J")).Locked = True
'        Sheet4.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HD8D8D8"
'     Else
'        Sheet4.Range(Replace(Target.Address, "K", "J")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "K", "J")).Locked = False
'        Sheet4.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HCCFFCC"
'    End If
'End If

endline:
Sheet4.Protect Password:=getmsgstate
Application.EnableEvents = True


End Sub


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub
