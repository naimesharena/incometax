Attribute VB_Name = "Sheet21"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim ActiveCellRange As String

Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet21.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Sheet21.Shapes("Button 6").Left = Range("Button_IF").Left
Sheet21.Shapes("Button 6").Top = Range("Button_IF").Top + 6

Sheet21.Shapes("Button 5").Left = Range("Button_SPI").Left
Sheet21.Shapes("Button 5").Top = Range("Button_SPI").Top + 4

Sheet21.Protect Password:=sPassword

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim cell As Variant
Application.EnableEvents = False

'--Newly added by Bindu on 25th May 26
If Not Application.Intersect(Sheet21.Range("SI_Editable"), Sheet21.Range(Target.Address)) Is Nothing Then
SI_Editable
Sheet21.Protect Password:=EfilingCommon.getmsgstate
Sheet21.Range("A1").Select
End If
'------end by Bindu on 25th May 26----'


' cell = Sheet21.Range("IF.FirmPAN").Cells
For Each cell In Sheet21.Range("IF.FirmPAN").Cells
Sheet21.Range("" & cell.Address).value = UCase(Sheet21.Range("" & cell.Address).value)
Next

For Each cell In Sheet21.Range("SPI.PANofSpecPerson").Cells
Sheet21.Range("" & cell.Address).value = UCase(Sheet21.Range("" & cell.Address).value)
Next

If Not IsEmpty(Trim(Target.value)) Then
If Not Application.Intersect(Sheet21.Range("SPI.PANofSpecPerson"), Sheet21.Range(Target.Address)) Is Nothing Then

       If Not CheckPAN(Sheet21.Range(Target.Address).value) Then
       
'            fmsgboxStatus "Please enter a valid PAN" ', vbOKOnly, "ITR-3"

           'SIT-69222 changed by sai
             fmsgboxStatus """Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule SPI." ', vbOKOnly, "ITR-3"
             
            Range(Target.Address).value = ""
            Range(Target.Address).Select

       End If
       GoTo endline
End If
End If

If (Target.name.name = "DTAA_TAX_OS") Then
    LockUnlockSurEduCess
    GoTo endline
End If

If (Target.name.name = "DTAA_TAX_OS_1") Then
    LockUnlockSurEduCess
    GoTo endline
End If

If (Target.name.name = "DTAA_TAX_OS_2") Then
    LockUnlockSurEduCess
    GoTo endline
End If


'--Newly added by Bindu on 25th May 26
If Sheet21.Range("Editable_111A_20_STCG") = "Err" Then '111A 20%
    MsgBox ("**Amount reported at ""111A - STCG on shares units on which STT paid as applicable"" of Schedule SI can not be more than sum of amount at Sl. No. A3ie and A4a of Schedule CG")
    GoTo endline
End If

If Sheet21.Range("Editable_PTI_20_STCG") = "Err" Then 'PTI 20%
      MsgBox ("**Amount reported at ""Pass Through Income in the nature of Short Term Capital Gain chargeable @ 20% "" of Schedule SI can not be more than amount at Sl. No. A8a of Schedule CG")
      GoTo endline
End If
If Sheet21.Range("Editable_115AD_20_STCG") = "Err" Then '115AD 20%
        MsgBox ("**Amount reported at ""115AD(1)(b)(ii)-Proviso - Short term capital gains referred to in section 111A rws. 115AD by FII, as applicable u/s 115AD(1)(b)(ii)"" of Schedule SI can not be more than amount at Sl. No. A3iie of Schedule CG")
        GoTo endline
End If
'------'

'Konda updated on 26-05-2026
'STCG-30%==================================

If Sheet21.Range("Editable_115AD_30_STCG") = "Err" Then '115AD 50%
    MsgBox ("**Amount reported at ""115AD(ii) -Short term capital gains (other than on equity share or equity oriented mutual fund referred to in section 111A) by an FII"" of Schedule SI can not be more than amount at Sl. No. A5e of Schedule CG")
    GoTo endline
End If

If Sheet21.Range("Editable_PTI_30_STCG") = "Err" Then 'PTI 30%
      MsgBox ("**Amount reported at ""Pass Through Income in the nature of Short Term Capital Gain chargeable @ 30% "" of Schedule SI can not be more than amount at Sl. No. A8b of Schedule CG")
      GoTo endline
End If

'LTCG-12.5%==================================

If Sheet21.Range("Editable_21b_12_5_LTCG") = "Err" Then '21b--G23 Cell
    MsgBox ("**Amount reported at ""112 -LTCG on others"" of Schedule SI can not be more than sum of amount at Sl. No. B1g, B2e, B5c, B9e, B10 of Schedule CG")
    GoTo endline
End If

If Sheet21.Range("Editable_22_12_5_LTCG") = "Err" Then '12_5_22--G26 Cell
      MsgBox ("**Amount reported at ""112(1) - LTCG on listed securities/ units without indexation"" of Schedule SI can not be more than amount at Sl. No. B3ie of Schedule CG")
      GoTo endline
End If

If Sheet21.Range("Editable_112A_12_5_LTCG") = "Err" Then '112A--G28 Cell
    MsgBox ("**Amount reported at ""112A -LTCG on sale of shares /units of equity oriented fund/units of business trust on which STT is paid"" of Schedule SI can not be more than amount at Sl. No. B4c of Schedule CG")
    GoTo endline
End If

If Sheet21.Range("Editable_1121C_12_5_LTCG") = "Err" Then '1121c--G30 Cell
      MsgBox ("**Amount reported at ""112(1)(c)(iii) - Long term capital gains on transfer of unlisted securities in the case of non-residents"" of Schedule SI can not be more than amount at Sl. No. B6ie of Schedule CG")
      GoTo endline
End If


If Sheet21.Range("Editable_115AC1_12_5_LTCG") = "Err" Then '115AC1--G50
      MsgBox ("**Amount reported at ""115AC(1)(c) -Long term capital gains arising from their transfer of bonds or GDR purchased in foreign currency in case of a non-resident"" of Schedule SI can not be more than amount at Sl. No. B6iie of Schedule CG")
      GoTo endline
End If

If Sheet21.Range("Editable_115ACA1b_12_5_LTCG") = "Err" Then '115ACA1b--G53 Cell
    MsgBox ("**Amount reported at ""115ACA(1)(b) - Long term capital gains arising from the transfer of GDR purchased in foreign currency -resident"" of Schedule SI can not be more than amount at Sl. No. B3iie of Schedule CG")
    GoTo endline
End If

If Sheet21.Range("Editable_115AD1Biii_FII_12_5_LTCG") = "Err" Then '115AD1Biii_FII--G58 Cell
      MsgBox ("**Amount reported at ""115AD(1)(b)(iii) - Long term capital gains (other than on equity share or equity oriented mutual fund referred to in section 112A) by an FII"" of Schedule SI can not be more than amount at Sl. No. B6iiie of Schedule CG")
      GoTo endline
End If
   
If Sheet21.Range("Editable_115AD1Biii_12_5_LTCG") = "Err" Then '115AD1biii--G59 Cell
    MsgBox ("**Amount reported at ""115AD(1)(b)(iii) Proviso -For NON-RESIDENTS - From sale of equity share in a company or unit of equity oriented fund or unit of a business trust on which STT is paid under section 112A"" of Schedule SI can not be more than amount at Sl. No. B7c of Schedule CG")
    GoTo endline
End If
    
If Sheet21.Range("Editable_115EB_12_5_LTCG") = "Err" Then '115Eb--G72 Cell
    MsgBox ("**Amount reported at ""115E(b) - Long term capital gains of a non-resident Indian on any foreign exchange asset "" of Schedule SI can not be more than amount at Sl. No. B8c of Schedule CG")
    GoTo endline
End If
   
If Sheet21.Range("Editable_PTI12_12_5_LTCG") = "Err" Then 'PTI_LTCG12--G79 Cell
    MsgBox ("**Amount reported at ""Pass Through Income in the nature of Long Term Capital Gain chargeable @ 12.5% u/s 112A"" of Schedule SI can not be more than amount at Sl. No. B11a1 of Schedule CG")
    GoTo endline
End If
    
If Sheet21.Range("Editable_PTIOTH112A_12_5_LTCG") = "Err" Then 'PTIOTH112A_12--G81 Cell
    MsgBox ("**Amount reported at ""Pass Through Income in the nature of Long Term Capital Gain chargeable @ 12.5% other than u/s 112A"" of Schedule SI can not be more than amount at Sl. No. B11a2 of Schedule CG")
    GoTo endline
End If
    

'============================


endline:
Application.EnableEvents = True
End Sub

Sub LockUnlockSurEduCess()
On Error Resume Next
    Sheet7.Unprotect Password:=getmsgstate
    Application.EnableEvents = False

    If Range("DTAA_TAX_OS").value > 0 Or Sheet14.Range("os.TaxBenefitTotal").value > 0 Or Sheet21.Range("DTAA_TAX_OS_1").value > 0 Or Sheet21.Range("DTAA_TAX_OS_2").value > 0 Then
        Sheet7.Range("Sheet9.Surcharge_ii").Locked = False
        Sheet7.Range("Sheet9.Surcharge_ii").Interior.Color = 13434879

        Sheet7.Range("Sheet9.EducationCess").Locked = False
        Sheet7.Range("Sheet9.EducationCess").Interior.Color = 13434879

    Else
        Sheet7.Range("Sheet9.Surcharge_ii").Locked = True
        Sheet7.Range("Sheet9.Surcharge_ii").Interior.Color = 16777215
        'Sheet7.Range("Sheet9.SurchargeOnTaxPayable").Value = "0"

        Sheet7.Range("Sheet9.EducationCess").Locked = True
        Sheet7.Range("Sheet9.EducationCess").Interior.Color = 16777215
        'Sheet7.Range("Sheet9.EducationCess").Value = "0"
    End If
    Application.EnableEvents = True
    Sheet7.Protect Password:=getmsgstate
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

Sub SI_Editable()
'----Newly added by Bindu on 25th May 26-----'
If UCase(Sheet21.Range("SI_Editable")) = UCase(Trim("Yes")) Then
      
'--20% STCG
    '111A
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc111b").Locked = False
    Sheet21.Range("splRateInc111b").Copy
    Sheet21.Range("splRateInc111b").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc111b").Interior.Color = "&HCCFFCC"
    
    '115AD
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD1b").Locked = False
    Sheet21.Range("splRateInc115AD1b").Copy
    Sheet21.Range("splRateInc115AD1b").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115AD1b").Interior.Color = "&HCCFFCC"
    
    'PTI
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTI_STCG20").Locked = False
    Sheet21.Range("splRateIncPTI_STCG20").Copy
    Sheet21.Range("splRateIncPTI_STCG20").PasteSpecial xlPasteValues
    Sheet21.Range("splRateIncPTI_STCG20").Interior.Color = "&HCCFFCC"
'--20% END
'Konda updated on 26-05-2026
    '--30% STCG
    
    '115AD--G33 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD_STCG30").Locked = False
    Sheet21.Range("splRateInc115AD_STCG30").Copy
    Sheet21.Range("splRateInc115AD_STCG30").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115AD_STCG30").Interior.Color = "&HCCFFCC"
    
    'PTI--G77 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTI_STCG30").Locked = False
    Sheet21.Range("splRateIncPTI_STCG30").Copy
    Sheet21.Range("splRateIncPTI_STCG30").PasteSpecial xlPasteValues
    Sheet21.Range("splRateIncPTI_STCG30").Interior.Color = "&HCCFFCC"
'--30% STCG END
    
  '--12.5% LTCG
  
  '21b--G23 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc21b").Locked = False
    Sheet21.Range("splRateInc21b").Copy
    Sheet21.Range("splRateInc21b").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc21b").Interior.Color = "&HCCFFCC"
    
    '12_5_22--G26 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc12_5_22").Locked = False
    Sheet21.Range("splRateInc12_5_22").Copy
    Sheet21.Range("splRateInc12_5_22").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc12_5_22").Interior.Color = "&HCCFFCC"
    
    '112A--G28 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc112A").Locked = False
    Sheet21.Range("splRateInc112A").Copy
    Sheet21.Range("splRateInc112A").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc112A").Interior.Color = "&HCCFFCC"
    
    '1121c--G30 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc1121c").Locked = False
    Sheet21.Range("splRateInc1121c").Copy
    Sheet21.Range("splRateInc1121c").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc1121c").Interior.Color = "&HCCFFCC"
    
    '115AC1--G50
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AC1").Locked = False
    Sheet21.Range("splRateInc115AC1").Copy
    Sheet21.Range("splRateInc115AC1").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115AC1").Interior.Color = "&HCCFFCC"
    
    '115ACA1b--G53 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115ACA1b").Locked = False
    Sheet21.Range("splRateInc115ACA1b").Copy
    Sheet21.Range("splRateInc115ACA1b").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115ACA1b").Interior.Color = "&HCCFFCC"
    
    '115AD1Biii_FII--G58 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD1Biii_FII").Locked = False
    Sheet21.Range("splRateInc115AD1Biii_FII").Copy
    Sheet21.Range("splRateInc115AD1Biii_FII").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115AD1Biii_FII").Interior.Color = "&HCCFFCC"
    
    '115AD1biii--G59 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD1biii").Locked = False
    Sheet21.Range("splRateInc115AD1biii").Copy
    Sheet21.Range("splRateInc115AD1biii").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115AD1biii").Interior.Color = "&HCCFFCC"
    
    '115Eb--G72 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115Eb").Locked = False
    Sheet21.Range("splRateInc115Eb").Copy
    Sheet21.Range("splRateInc115Eb").PasteSpecial xlPasteValues
    Sheet21.Range("splRateInc115Eb").Interior.Color = "&HCCFFCC"
    
    'PTI_LTCG12--G79 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTI_LTCG12").Locked = False
    Sheet21.Range("splRateIncPTI_LTCG12").Copy
    Sheet21.Range("splRateIncPTI_LTCG12").PasteSpecial xlPasteValues
    Sheet21.Range("splRateIncPTI_LTCG12").Interior.Color = "&HCCFFCC"
    
    'PTIOTH112A_12--G81 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTIOTH112A_12").Locked = False
    Sheet21.Range("splRateIncPTIOTH112A_12").Copy
    Sheet21.Range("splRateIncPTIOTH112A_12").PasteSpecial xlPasteValues
    Sheet21.Range("splRateIncPTIOTH112A_12").Interior.Color = "&HCCFFCC"
    
  '--12.5% LTCG END
Else
    
'--20% STCG
    '111A
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc111b").formula = "=BFLAtemp20Sec111a"
    Sheet21.Range("splRateInc111b").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc111b").Locked = True
    
    ''115AD
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD1b").formula = "=BFLAtemp20Sec115ad1bii"
    Sheet21.Range("splRateInc115AD1b").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115AD1b").Locked = True

    ''PTI
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTI_STCG20").formula = "=BFLAtempPTI20_STCG"
    Sheet21.Range("splRateIncPTI_STCG20").Interior.ColorIndex = 2
    Sheet21.Range("splRateIncPTI_STCG20").Locked = True
''--20% END

'Konda updated on 26-05-2026
    
     ''115AD
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD_STCG30").formula = "=BFLAtemp115AD30"
    Sheet21.Range("splRateInc115AD_STCG30").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115AD_STCG30").Locked = True

    ''PTI
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTI_STCG30").formula = "=BFLAtempPTI30"
    Sheet21.Range("splRateIncPTI_STCG30").Interior.ColorIndex = 2
    Sheet21.Range("splRateIncPTI_STCG30").Locked = True
''--20% END

'--12.5% LTCG
  
'21b--G23 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc21b").formula = "=BFLAtemp125Sec_112"
    Sheet21.Range("splRateInc21b").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc21b").Locked = True
    
'12_5_22--G26 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc12_5_22").formula = "=BFLAtemp125SecPro"
    Sheet21.Range("splRateInc12_5_22").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc12_5_22").Locked = True
    
    '112A--G28 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc112A").formula = "=BFLAtemp125Sec_112A"
    Sheet21.Range("splRateInc112A").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc112A").Locked = True
    
    '1121c--G30 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc1121c").formula = "=BFLAtemp125Sec112_C"
    Sheet21.Range("splRateInc1121c").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc1121c").Locked = True
    
    '115AC1--G50
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AC1").formula = "=BFLAtemp125Sec_115AC"
    Sheet21.Range("splRateInc115AC1").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115AC1").Locked = True
    
    '115ACA1b--G53 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115ACA1b").formula = "=BFLAtemp125Sec_115ACA"
    Sheet21.Range("splRateInc115ACA1b").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115ACA1b").Locked = True
    
    '115AD1Biii_FII--G58 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD1Biii_FII").formula = "=BFLAtemp125Sec115AD3"
    Sheet21.Range("splRateInc115AD1Biii_FII").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115AD1Biii_FII").Locked = True
    
    '115AD1biii--G59 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115AD1biii").formula = "=BFLAtemp125Sec_115AD_iii"
    Sheet21.Range("splRateInc115AD1biii").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115AD1biii").Locked = True
    
    '115Eb--G72 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateInc115Eb").formula = "=BFLAtemp125Sec_115EB"
    Sheet21.Range("splRateInc115Eb").Interior.ColorIndex = 2
    Sheet21.Range("splRateInc115Eb").Locked = True
    
    'PTI_LTCG12--G79 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTI_LTCG12").formula = "=BFLAtemp125Sec_112_PTI"
    Sheet21.Range("splRateIncPTI_LTCG12").Interior.ColorIndex = 2
    Sheet21.Range("splRateIncPTI_LTCG12").Locked = True
    
    'PTIOTH112A_12--G81 Cell
    Sheet21.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet21.Range("splRateIncPTIOTH112A_12").formula = "=BFLAtemp125Sec_PTI"
    Sheet21.Range("splRateIncPTIOTH112A_12").Interior.ColorIndex = 2
    Sheet21.Range("splRateIncPTIOTH112A_12").Locked = True
    
  '--12.5% LTCG END

'---------------------


End If

End Sub
