Attribute VB_Name = "Sheet14"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String
Dim end_os As Long


Private Sub Worksheet_Activate()
On Error Resume Next


'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet14.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet14.Shapes("Button 7").Left = Range("Button_OS").Left + 2
Sheet14.Shapes("Button 7").Top = Range("Button_OS").Top + 2

Sheet14.Shapes("Button 68").Left = Range("Button_OS1").Left + 2
Sheet14.Shapes("Button 68").Top = Range("Button_OS1").Top + 2

Sheet14.Shapes("Button 68").Left = Range("Button_OS1").Left + 2
Sheet14.Shapes("Button 68").Top = Range("Button_OS1").Top + 2

Sheet14.Shapes("Button 97").Left = Range("Button_OS2").Left + 2
Sheet14.Shapes("Button 97").Top = Range("Button_OS2").Top + 2

Sheet14.Shapes("Button 40").Left = Range("Button_40").Left + 2
Sheet14.Shapes("Button 40").Top = Range("Button_40").Top + 6

Sheet14.Shapes("Button 163").Left = Range("Button_OS_e").Left + 2
Sheet14.Shapes("Button 163").Top = Range("Button_OS_e").Top + 2

Dim chcell As Variant
For Each chcell In Sheet14.Range("os.SourceDescription").Cells
If chcell.value = "1 - 111 - Tax on accumulated balance of recognised PF" Then
chcell.value = ""
End If
Next
'Added by Ayush
OS_Table_D_F
OS_Table_E
Sheet14.Protect Password:=sPassword

End Sub
Sub dpkshp()
'Dim shp As Shape
'
'For Each shp In ActiveSheet.Shapes
'MsgBox shp.name & ":" & shp.AlternativeText
'If Mid(shp.name, 1, 3) = "Com" Then
'shp.Delete
'End If
'Next shp


End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline


If Not Application.Intersect(Sheet14.Range("os.WinOnlineGame"), Range(Target.Address)) Is Nothing Then
If Target.value < 0 Or Target.value > 99999999999999# Then
fmsgboxStatus "Amount should be Numeric, Non negative, No decimal and up to 14 digits"
Target.value = ""
End If


If Target.value > 0 Then

fmsgboxStatus "If the benefit of quarterly computation of interest u/s 234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
End If
End If




CalculateTaxCheckFlag = False

Dim sourceDescription As Range
Dim SourceAmount As Range
Dim rangecell As Range
Dim i As Long

 Application.EnableEvents = False
 Application.ScreenUpdating = False
 
 
 Dim sPassword

    sPassword = getmsgstate
'    If Not Application.Intersect(Sheet14.Range("IntrstSec10XIFP"), Sheet14.Range(Target.Address)) Is Nothing Then
'    If (Range(Target.Address).value) < (Sheet14.Range("IntrstSec10XIFP").value + Sheet14.Range("IntrstSec10XISP").value + Sheet14.Range("IntrstSec10XIIFP").value + Sheet14.Range("IntrstSec10XIISP").value) Then
'    fmsgbox " *You can enter amount only in one out of the 4 Nature of Income 'First proviso to Section 10(11), Second proviso to Section 10(11),First proviso to Section 10(12) ,Second proviso to Section 10(12)'"")"
'    Range(Target.Address).value = ""
'    End If
'    End If
'
'    If Not Application.Intersect(Sheet14.Range("IntrstSec10XISP"), Sheet14.Range(Target.Address)) Is Nothing Then
'    If (Range(Target.Address).value) < (Sheet14.Range("IntrstSec10XIFP").value + Sheet14.Range("IntrstSec10XISP").value + Sheet14.Range("IntrstSec10XIIFP").value + Sheet14.Range("IntrstSec10XIISP").value) Then
'    fmsgbox " *You can enter amount only in one out of the 4 Nature of Income 'First proviso to Section 10(11), Second proviso to Section 10(11),First proviso to Section 10(12) ,Second proviso to Section 10(12)'"")"
'    Range(Target.Address).value = ""
'    End If
'    End If
'
'    If Not Application.Intersect(Sheet14.Range("IntrstSec10XIIFP"), Sheet14.Range(Target.Address)) Is Nothing Then
'    If (Range(Target.Address).value) < (Sheet14.Range("IntrstSec10XIFP").value + Sheet14.Range("IntrstSec10XISP").value + Sheet14.Range("IntrstSec10XIIFP").value + Sheet14.Range("IntrstSec10XIISP").value) Then
'    fmsgbox " *You can enter amount only in one out of the 4 Nature of Income 'First proviso to Section 10(11), Second proviso to Section 10(11),First proviso to Section 10(12) ,Second proviso to Section 10(12)'"")"
'    Range(Target.Address).value = ""
'    End If
'    End If
'
'    If Not Application.Intersect(Sheet14.Range("IntrstSec10XIISP"), Sheet14.Range(Target.Address)) Is Nothing Then
'    If (Range(Target.Address).value) < (Sheet14.Range("IntrstSec10XIFP").value + Sheet14.Range("IntrstSec10XISP").value + Sheet14.Range("IntrstSec10XIIFP").value + Sheet14.Range("IntrstSec10XIISP").value) Then
'    fmsgbox " *You can enter amount only in one out of the 4 Nature of Income 'First proviso to Section 10(11), Second proviso to Section 10(11),First proviso to Section 10(12) ,Second proviso to Section 10(12)'"")"
'    Range(Target.Address).value = ""
'    End If
'    End If
    
    If Not Application.Intersect(Sheet14.Range("OSIncreliefus89A"), Sheet14.Range(Target.Address)) Is Nothing Then
    
        Sheet14.Unprotect Password:=getmsgstate
        If Sheet14.Range("OSIncreliefus89A").value > Sheet14.Range("OSIncomeNotified89A").value Then
            fmsgbox "* Relief u/s 89A cannot be claimed more than income offered in Income from retirement benefit account maintained in a notified country u/s 89A."
            Sheet14.Range("OSIncreliefus89A").value = 0
        End If
        
        If Sheet14.Range("OSIncreliefus89A").value > 0 Then
            fmsgbox "*Please ensure to submit Form 10EE. "
        End If
        
    End If
    Sheet14.Protect Password:=getmsgstate
    
    If Not Intersect(Range("os.NatureOfIncome"), Range(Target.Address)) Is Nothing Then
            If UCase(Mid(Range(Target.Address).value, 1, 2)) = "2E" Then
                Sheet14.Unprotect sPassword
                Range("M" & Target.row).value = "(Select)"
                With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="=INDIRECT(IF(MID(Sheet1.115H,1,1)=""Y"",""OS_PTI_115H"",IF(MID(sheet1.ResidentialStatus1,1,3)<>""NRI"",""OS_PTI_RES"",""OS_PTI_NRI"")))"
            End With
'                 Range("J" & Target.Row).Validation.Formula1:=  "=IF(IsForeignResident,OS.PTIResFC,IF(IsForeignNRI,OS.PTINRI,IF(MID(sheet1.ResidentialStatus1,1,3)=""NRI"",OS.PTINRI,OS.PTIRes)))"
'
                Sheet14.Protect sPassword
            ElseIf UCase(Mid(Range(Target.Address).value, 1, 2)) = "2D" Then
                Sheet14.Unprotect sPassword
                Range("M" & Target.row).value = "(Select)"
                 With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="=INDIRECT(IF(MID(Sheet1.115H,1,1)=""Y"",""OS_ALL"",IF(MID(sheet1.ResidentialStatus1,1,3)<>""NRI"",""OS_RESandNOR"",""OS_NRI"")))"
                End With
                Sheet14.Protect sPassword
'            ElseIf UCase(Mid(Range(target.Address).value, 1, 3)) = "1AI" Then

            ElseIf UCase(Range(Target.Address).value) = "1AI" Then
                Sheet14.Unprotect sPassword
                 Range("M" & Target.row).value = "56(2)(i)-Dividends"
                  With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="56(2)(i)-Dividends"
                End With
                Sheet14.Protect sPassword
          '===============================================================================
           'Newly added by sai on 20/02/2025 For AY 2025-26

            ElseIf UCase(Mid(Range(Target.Address).value, 1, 5)) = "1AIII" Then
                Sheet14.Unprotect sPassword
                 Range("M" & Target.row).value = "56(2)(i)-Dividend u/s 2(22)(f)"
                  With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="56(2)(i)-Dividend u/s 2(22)(f)"
                End With
                Sheet14.Protect sPassword
           '===============================================================================
                
            ElseIf UCase(Mid(Range(Target.Address).value, 1, 2)) = "1B" Then
                Sheet14.Unprotect sPassword
                 Range("M" & Target.row).value = "56(2)-Interest"
                  With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="56(2)-Interest"
                 End With
                Sheet14.Protect sPassword
            ElseIf UCase(Mid(Range(Target.Address).value, 1, 2)) = "1C" Then
                Sheet14.Unprotect sPassword
                 Range("M" & Target.row).value = "56(2)(iii)-Rental income from machinery, plants, buildings etc.,"
                With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="56(2)(iii)-Rental income from machinery plants buildings etc.,"
                 End With
                Sheet14.Protect sPassword
            ElseIf UCase(Mid(Range(Target.Address).value, 1, 2)) = "1D" Then
                Sheet14.Unprotect sPassword
                 Range("M" & Target.row).value = "56(2)(x)-Income under section 56(2)(x)."
                 With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="56(2)(x)-Income under section 56(2)(x)."
                 End With
                Sheet14.Protect sPassword
            ElseIf UCase(Mid(Range(Target.Address).value, 1, 6)) = "2AI" Then
                Sheet14.Unprotect sPassword
                 'Range("M" & target.row).value = "5BB-Winnings from lotteries crossword puzzles etc.[S.115BB]"
                 
        'SIT-105227 Sadineni on 17-12-2025
                  Range("M" & Target.row).value = "5BB-Winnings from lotteries, crossword puzzles, races, card games etc."
                 With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
'                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
'                    Formula1:="5BB-Winnings from lotteries crossword puzzles etc.[S.115BB]"

            'SIT-105227 Sadineni on 17-12-2025
                  .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="=OS_5BB"
                
                 End With
                Sheet14.Protect sPassword
                
                'Chandru
                ElseIf UCase(Mid(Range(Target.Address).value, 1, 6)) = "2AII" Then
                Sheet14.Unprotect sPassword
                 Range("M" & Target.row).value = "5BBJ - Winnings from online games"
                 With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="5BBJ - Winnings from online games"
                 End With
                Sheet14.Protect sPassword
                
                
            Else
             Sheet14.Unprotect sPassword
                Range("M" & Target.row).value = "(Select)"
                 With Range("M" & Target.row).Validation
                .Delete 'delete previous validation
                .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:="=INDIRECT(IF(MID(Sheet1.115H,1,1)=""Y"",""OSDTAA_Section_115H"",IF(MID(sheet1.ResidentialStatus1,1,3)<>""NRI"",""OSDTAA_Section_RES"",""OSDTAA_Section_NRI"")))"
            End With
                Sheet14.Protect sPassword
            End If
    End If

 Dim targetadd, temp As Variant
    If Not Application.Intersect(Sheet14.Range("os.TRC"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "L", "F")).value
     sPassword = EfilingCommon.getmsgstate
    Sheet14.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Range(targetadd).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet14.Range("os.ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet14.Range("os.Rate_DTAA").Cells.item(temp).value, Sheet14.Range("os.Rate_ITAct").Cells.item(temp).value)
         Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
        Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else
       Sheet14.Range("os.ApplRate").Cells.item(temp).ClearContents
       Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
       Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
    Sheet14.Protect Password:=sPassword
   End If

       If Not Application.Intersect(Sheet14.Range("os.Rate_DTAA"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "K", "F")).value
     sPassword = EfilingCommon.getmsgstate
    Sheet14.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Sheet14.Range("os.TRC").Cells.item(temp).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet14.Range("os.ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet14.Range("os.Rate_DTAA").Cells.item(temp).value, Sheet14.Range("os.Rate_ITAct").Cells.item(temp).value)
         Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
         Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else
       Sheet14.Range("os.ApplRate").Cells.item(temp).ClearContents
       Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
       Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
  Sheet14.Protect Password:=sPassword
   End If

 If Not Application.Intersect(Sheet14.Range("os.Rate_ITAct"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "N", "F")).value
     sPassword = EfilingCommon.getmsgstate
    Sheet14.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
     If Sheet14.Range("os.TRC").Cells.item(temp).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet14.Range("os.ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet14.Range("os.Rate_DTAA").Cells.item(temp).value, Sheet14.Range("os.Rate_ITAct").Cells.item(temp).value)
         Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
         Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else
       Sheet14.Range("os.ApplRate").Cells.item(temp).ClearContents
       Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
       Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
  Sheet14.Protect Password:=sPassword
   End If

'Newly added by sai SIT-72691
If Not Intersect(Range("OSIncomeNotified89A_AmountUS"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("OSIncomeNotified89A_AmountUK"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("OSIncomeNotified89A_AmountCan"), Range(Target.Address)) Is Nothing Then

     If Range("OSIncomeNotified89A") > 0 Then

     MsgBox ("If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"), , "Alert!!"
     End If
     
End If




'If Not Intersect(Range("os.TaxBenefit"), Range(Target.address)) Is Nothing Then
'    sPassword = EfilingCommon.getmsgstate
'   Sheet9.Unprotect Password:=sPassword
'     If (Range(Target.address).value) > 0 Then
'        Sheet9.Range("os.Depreciation").Locked = False
'        Sheet9.Range("os.Depreciation").Interior.Color = &HCCFFCC
'    Else
'        Sheet9.Range("os.Depreciation").Locked = True
'        Sheet9.Range("os.Depreciation").value = ""
'        Sheet9.Range("os.Depreciation").Interior.Color = &HD8D8D8
' End If
' Sheet9.Protect Password:=sPassword
'End If



If Not IsEmpty(Trim(Target.value)) Then
        'Source Section Validation: On Data Entry
        If Target.Validation.Type = 3 Then
            Dim nOfCount As Long
            nOfCount = EfilingCommon.CountDropdown(Trim(Target.value), Sheet14.Range("os.SourceDescription").Cells, Sheet14.Range("os.SourceDescription").count)
            If nOfCount > 1 Then
                fmsgboxStatus " " & Trim(Target.value) & "  cannot be selected twice " ', vbOKOnly, "ITR 3"
                Target.value = "(Select)"
            End If
        End If

        If Target.Validation.Type = 3 Then
            Dim nOfCount1 As Long
            nOfCount1 = EfilingCommon.CountDropdown(Trim(Target.value), Sheet14.Range("os.PTIDescription").Cells, Sheet14.Range("os.PTIDescription").count)
            If nOfCount1 > 1 Then
                fmsgboxStatus " " & Trim(Target.value) & "  cannot be selected twice " ', vbOKOnly, "ITR 3"
                Target.value = "(Select)"
            End If
        End If


End If

 If Not Intersect(Range("os.SourceDescription"), Range(Target.Address)) Is Nothing Then
     If Mid(Range(Target.Address).value, 1, 5) = "5BBDA" Then
          fmsgboxStatus "Please enter the amount, if Dividend income from domestic company exceeds Rs.10 lakh"
       End If
 End If

    Set sourceDescription = Range("os.SourceDescription")
        If Not Application.Intersect(Range("os.SourceDescription"), Range(Target.Address)) Is Nothing Then
          setTblinfo_OS_local
        Set rangecell = Sheet14.Range("os.SourceDescription").Cells
        For i = 1 To end_os
            If rangecell.item(i).value = "" Then
                fmsgboxStatus "please fill all the mandatory fields of the row " & (i + 2) & " before filling the next row in sheet :OS " ', vbOKOnly, "ITR 3"
                Target.value = ""
                Exit Sub
            End If
        Next
        populateSI
       End If



If Not Intersect(Range("os.TaxBenefit"), Range(Target.Address)) Is Nothing Then

    LockUnlockSurEduCessOS
    GoTo endline
End If

'Change-14.05.2023.101.OS -> SIT-49871
' If Target.name.name = "os_FamilyPension" Then
'
' LockUnlockDeduction57iia
'
' End If
'---
'Change-14.05.2023.101.OS -> SIT-49871
 'Change, 17.04.2023.102.SIT 48765 B
' If Target.name.name = "os_Deduction57iia" Then
'    If (Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Then
'        Sheet14.Range("os_Deduction57iia").value = 0
'
'    End If
' End If
 'End Change
 '----
 
If Not Intersect(Range("os.RentFromMachPlantBldgs"), Range(Target.Address)) Is Nothing Then

'If Target.name.name = "os.RentFromMachPlantBldgs" Then

 sPassword = EfilingCommon.getmsgstate
   Sheet14.Unprotect Password:=sPassword
     If (Range(Target.Address).value) > 0 Then
        Sheet14.Range("os.Depreciation").ClearContents
        Sheet14.Range("os.Depreciation").Locked = False
        Sheet14.Range("os.Depreciation").Interior.Color = &HCCFFCC
    Else
        Sheet14.Range("os.Depreciation").Locked = True
        Sheet14.Range("os.Depreciation").ClearContents
        Sheet14.Range("os.Depreciation").Interior.Color = &HD8D8D8
 End If
 Sheet14.Protect Password:=sPassword
   End If

sPassword = EfilingCommon.getmsgstate
Sheet14.Unprotect Password:=sPassword
    If Sheet14.Range("Allow3ai").value > 0 Then
        Sheet14.Range("os.Expenses").Locked = False
        Sheet14.Range("os.Expenses").Interior.Color = &HCCFFCC
    Else
        Sheet14.Range("os.Expenses").Interior.Color = &HD8D8D8
        Sheet14.Range("os.Expenses").value = ""
        Sheet14.Range("os.Expenses").Locked = True
    End If
    
'modified by sai on 20/02/2025 For AY 2025-26

    'If Sheet14.Range("os.DividendGross").value > 0 Then===========================================================

    If Sheet14.Range("os.DividendGross_1ai_1aii").value > 0 Then
    
        Sheet14.Range("os.expenditureClaimed").Locked = False
        Sheet14.Range("os.expenditureClaimed").Interior.Color = &HCCFFCC
        'AY_2025-26 os.expenditureClaimed field taiking value always "zero" so comented
        'Sheet14.Range("os.expenditureClaimed").value = 0
       
    Else
        Sheet14.Range("os.expenditureClaimed").Interior.Color = &HD8D8D8
        Sheet14.Range("os.expenditureClaimed").value = 0
        Sheet14.Range("os.expenditureClaimed").Locked = True
    End If
 '===============================================================================================================
'uncommented by Chetan C M on 30/01/2026 for AY 2026-27
'start--

''Newly added by sai on 20/02/2025 for AY 2025-26
If Not Application.Intersect(Sheet14.Range("os.DividendGrossaiii"), Range(Target.Address)) Is Nothing Then

'Newly added by sadineni on 19/05/2026
'**********************************************************
age1 = calculateAge(Sheet1.Range("sheet1.DOB").value)
If (age1 <= 59) Then
'***********************************************************

    If Sheet14.Range("os.DividendGrossaiii").value > 0 Then
       fmsgboxStatus "If the benefit of quarterly computation of interest u/s 234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'" ', vbOKOnly, "ITR 3"
    End If
End If

End If
'--end
'==================================================================================

 'Newly added by Ayush
 
If Not Application.Intersect(Sheet14.Range("os.WinLottRacePuzz"), Range(Target.Address)) Is Nothing Then
    If Sheet14.Range("os.WinLottRacePuzz").value > 0 Then
    
      fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources' " ', vbOKOnly, "ITR 3"
    
    End If
 
End If

'uncommented by sai on 13/05/2025 as per DE V0.7

If Not Application.Intersect(Sheet14.Range("os.WinOnlineGame"), Range(Target.Address)) Is Nothing Then
    If Sheet14.Range("os.WinOnlineGame").value > 0 Then

      fmsgboxStatus "If the benefit of quarterly computation of interest u/s 234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
    End If

End If

'Newly Modified by Sai on 20/02/2025 for AY 2025-26
'If Not Application.Intersect(Sheet14.Range("os.WinOnlineGame2aii_i"), Range(target.Address)) Is Nothing Or _
'   Not Application.Intersect(Sheet14.Range("os.WinOnlineGame2aii_ii"), Range(target.Address)) Is Nothing Then
'
'    If target.value <> "" Then
'       If Sheet14.Range("os.WinOnlineGame").value > 0 And Sheet31.Range("age") < 59 Then
'          fmsgboxStatus " If the benefit of quarterly computation of interest u/s 234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
'       End If
'    End If
'End If
'--------------------------------------


    
Sheet14.Protect Password:=sPassword

 '=========================Malli_22_02_24
  If Not Application.Intersect(Sheet14.Range("os.PTIAmount"), Range(Target.Address)) Is Nothing Then
    
    If Target.value > 0 Then
        If Sheet14.Range(Replace(Target.Address, "H", "G")).value = "PTI_115A(1)(a)(i)-Dividends received by non-resident (not being company) or foreign company chargeable u/s 115A(1)(a)(i) other than first provisio to 115A(1)(a)(A)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "PTI_115A(1)(a)(A)-Dividend received by non-resident (not being company) or foreign company from a unit in an International Financial Services Centre, as referred to in sub-section (1A) of section 80LA chargeable under proviso to section 115A(1)(a)(A)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "PTI_115ACA(1)(a)-Income from GDR purchased in foreign currency -resident" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "PTI_5AD1iDiv-PTI-Income (being dividend) received by an FII in respect of securities (other than units referred to in section 115AB) - chargeable u/s 115AD(1)(i)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "PTI_5AC1b-PTI -Income by way of Dividend on GDRs purchased in foreign currency - non-resident" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "PTI_115AD(1)(i)-Income (other than dividend) received by an FII in respect of securities (other than units as per Sec 115AB)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        
        End If
    End If
 End If
 
 
 'Newly added
 
 If Not Application.Intersect(Sheet14.Range("os.SourceAmount"), Range(Target.Address)) Is Nothing Then
    
    If Target.value > 0 Then
        If Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5A1ai - 115A(1)(a)(i) - Dividends received by non-resident (not being company) or foreign company chargeable u/s 115A(1)(a)(i) other than first  proviso to section  115A(1)(a)(A)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5A1aA - 115A(1)(a)(A) - Dividend received by non-resident (not being company) or foreign company from a unit in an International Financial Services Centre, as referred to in sub-section (1A) of section 80LA chargeable under proviso to section 115A(1)(a)(A)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5ACA1a - 115ACA(1)(a) -  Income from GDR purchased in foreign currency  - resident" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5AD1iDiv-Income (being dividend) received by an FII in respect of securities (other than units referred to in section 115AB) - chargeable u/s 115AD(1)(i)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5AC1b - 115AC1b - Income by way of Dividend on GDRs purchased in foreign currency - non-resident" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
        ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5AD1i - 115AD(1)(i) - Income (other than dividend) received by an FII in respect of securities (other than units as per Sec 115AB)" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
         ElseIf Sheet14.Range(Replace(Target.Address, "H", "G")).value = "5ACA1a - 115ACA(1)(a) -  Income from GDR purchased in foreign currency  - resident" Then
        fmsgboxStatus "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table 10 'Information about accrual/receipt of income from Other Sources'"
                
        End If
    End If
 End If




'Added by Aavula

Dim cChell
Dim rangecells As Range
Dim rangecell1 As Range
Dim rangecell2 As Range
Dim mIntCells

Set rangecells = Sheet14.Range("os.Amount_income")
Set rangecell1 = Sheet14.Range("os.NatureOfIncome")
Set rangecell2 = Sheet14.Range("os.TRC")
mIntCells = Range("os.NatureOfIncome").count

'Modified by sai on 20/02/2025 For AY 2025-26
'If Not Application.Intersect(Sheet14.Range("os.WinOnlineGame2aii_i"), Range(target.Address)) Is Nothing Or _
'    Not Application.Intersect(Sheet14.Range("os.WinOnlineGame2aii_ii"), Range(target.Address)) Is Nothing Or _
'    Not Application.Intersect(Sheet14.Range("os.WinLottRacePuzz"), Range(target.Address)) Is Nothing Or _
'    Not Application.Intersect(Sheet14.Range("os.Amount_income"), Range(target.Address)) Is Nothing Or _
'    Not Application.Intersect(Sheet14.Range("os.NatureOfIncome"), Range(target.Address)) Is Nothing Or _
'    Not Application.Intersect(Sheet14.Range("os.TRC"), Range(target.Address)) Is Nothing Then

'by sai on 13/05/2025 as per DE V0.5
If Not Application.Intersect(Sheet14.Range("os.WinOnlineGame"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.WinLottRacePuzz"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.Amount_income"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.NatureOfIncome"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.TRC"), Range(Target.Address)) Is Nothing Then


    If Sheet14.Range("os.WinLottRacePuzz").value > 0 Or Sheet14.Range("os.WinOnlineGame") > 0 Then
    
      If Sheet14.Range("os.WinLottRacePuzz").value > 0 Then
         Sheet14.Range("IncomefrmLottery.Upto15Of6").value = Sheet14.Range("os.WinLottRacePuzz")
      Else
        Sheet14.Range("IncomefrmLottery.Upto15Of6").value = ""
        
      End If
       
       If Sheet14.Range("os.WinOnlineGame").value > 0 Then
        Sheet14.Range("IncomefrmOnlineGames.Upto15Of6") = Sheet14.Range("os.WinOnlineGame")
       Else
        Sheet14.Range("IncomefrmOnlineGames.Upto15Of6").value = ""
       End If
   

    
        For i = 1 To mIntCells
          
          If rangecell1.item(i).value = "2ai" And rangecells.item(i).value > 0 Then
    
             If rangecell2.item(i).Locked = True Or UCase(rangecell2.item(i).value) = UCase("Yes") Then
    
               If (Sheet14.Range("os.WinLottRacePuzz") - rangecells.item(i).value) > 0 Then
                   Sheet14.Range("IncomefrmLottery.Upto15Of6").value = Sheet14.Range("os.WinLottRacePuzz") - rangecells.item(i).value
                'Exit For
               End If
    
             ElseIf UCase(rangecell2.item(i).value) = UCase("No") Then
              If Sheet14.Range("os.WinLottRacePuzz") > 0 Then
                  Sheet14.Range("IncomefrmLottery.Upto15Of6").value = Sheet14.Range("os.WinLottRacePuzz")
               'Exit For
               End If
    
             End If
          End If
    
          If rangecell1.item(i).value = "2aii" And rangecells.item(i).value > 0 Then
    
             If rangecell2.item(i).Locked = True Or UCase(rangecell2.item(i).value) = UCase("Yes") Then
    
               If (Sheet14.Range("os.WinOnlineGame") - rangecells.item(i).value) > 0 Then
                  Sheet14.Range("IncomefrmOnlineGames.Upto15Of6").value = Sheet14.Range("os.WinOnlineGame") - rangecells.item(i).value
                'Exit For
               End If
    
             ElseIf UCase(rangecell2.item(i).value) = UCase("No") Then
             If Sheet14.Range("os.WinOnlineGame") > 0 Then
                 Sheet14.Range("IncomefrmOnlineGames.Upto15Of6").value = Sheet14.Range("os.WinOnlineGame")
              ' Exit For
              End If
    
             End If
          End If
    
        Next
    End If
End If

'uncommented by Chetan C M on 30/01/2026 for AY 2026-27
'start--

''Newly added by Sai on 25/02/2025
'*********************************************************************************************************************
If Not Application.Intersect(Sheet14.Range("os.DividendGrossaiii"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.Amount_income"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.NatureOfIncome"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.TRC"), Range(Target.Address)) Is Nothing Then

     If Sheet14.Range("Ifaai_1aiii").value > 0 Then
        Sheet14.Range("BBDA_1aiii.Upto15Of6").value = Sheet14.Range("Ifaai_1aiii")
        Else
        Sheet14.Range("BBDA_1aiii.Upto15Of6").value = ""
     End If

End If
'***********************************************************************************************************************

'--end

'If Not Application.Intersect(Sheet14.Range("OSIncreliefus89A"), Range(target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet14.Range("OSIncomeNotified89A"), Range(target.Address)) Is Nothing Then
'
'    If Sheet14.Range("OSIncreliefus89A").value > 0 Or Sheet14.Range("OSIncomeNotified89A") > 0 Then
'
'           If Sheet14.Range("OSIncreliefus89A") - Sheet14.Range("OSIncomeNotified89A") > 0 Then
'
'             Sheet14.Range("IncD_q1OS1").value = Sheet14.Range("OSIncreliefus89A") - Sheet14.Range("OSIncomeNotified89A")
'           End If
'
'    End If
'End If

'Newly added by sai on 21/03/2025

If Not Application.Intersect(Sheet14.Range("OSIncreliefus89A"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("OSIncomeNotified89A_AmountUS"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("OSIncomeNotified89A_AmountUK"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("OSIncomeNotified89A_AmountCan"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("OSIncomeNotified89A"), Range(Target.Address)) Is Nothing Then
    
    If Sheet14.Range("OSIncreliefus89A").value > 0 Or Sheet14.Range("OSIncomeNotified89A") > 0 Then
            
            
'           If Sheet14.Range("OSIncreliefus89A") - Sheet14.Range("OSIncomeNotified89A") > 0 Then
'              Sheet14.Range("IncD_q1OS1").value = Sheet14.Range("OSIncreliefus89A") - Sheet14.Range("OSIncomeNotified89A")
'           End If

           If Sheet14.Range("OSIncomeNotified89A") - Sheet14.Range("OSIncreliefus89A") > 0 Then
              Sheet14.Range("IncD_q1OS1").value = Sheet14.Range("OSIncomeNotified89A") - Sheet14.Range("OSIncreliefus89A")
           Else
           Sheet14.Range("IncD_q1OS1").value = ""
           End If
           '---------------------------------
    
     Else
           Sheet14.Range("IncD_q1OS1").value = ""
    End If
End If



'Added by Aavula


 If Not Application.Intersect(Sheet14.Range("os.SourceDescription"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.PTIDescription"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.PTIAmount"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.SourceAmount"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.Amount_income"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.NatureOfIncome"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.SectionAct"), Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet14.Range("os.TRC"), Range(Target.Address)) Is Nothing Then

    If Sheet14.Range("Ifaai_115A1ai").value > 0 Then
        Sheet14.Range("IncD_115A1aiQ1").value = Sheet14.Range("Ifaai_115A1ai")
        Else
        Sheet14.Range("IncD_115A1aiQ1").value = ""
     End If
      
     If Sheet14.Range("Ifaai_115A1aA").value > 0 Then
        Sheet14.Range("IncD_115A1aAQ1").value = Sheet14.Range("Ifaai_115A1aA")
        Else
        Sheet14.Range("IncD_115A1aAQ1").value = ""
     End If
     
       If Sheet14.Range("Ifaai_115AC").value > 0 Then
        Sheet14.Range("IncD_115ACQ1").value = Sheet14.Range("Ifaai_115AC")
        Else
        Sheet14.Range("IncD_115ACQ1").value = ""
     End If
     
      If Sheet14.Range("Ifaai_115ACA1a").value > 0 Then
        Sheet14.Range("IncD_115ACA1aQ1").value = Sheet14.Range("Ifaai_115ACA1a")
        Else
        Sheet14.Range("IncD_115ACA1aQ1").value = ""
     End If
     
     If Sheet14.Range("Ifaai_115AD1i").value > 0 Then
        Sheet14.Range("IncD_115AD1iQ1").value = Sheet14.Range("Ifaai_115AD1i")
        Else
        Sheet14.Range("IncD_115AD1iQ1").value = ""
     End If
     
     
'      If Sheet14.Range("Ifaai_DividendDTAA_rates").value > 0 Then
'        Sheet14.Range("DTAA.Upto15Of6").value = Sheet14.Range("Ifaai_DividendDTAA_rates")
'        Else
'        Sheet14.Range("DTAA.Upto15Of6").value = ""
'     End If
     
End If


Dim CheckDivDTAA, CheckDTAA5ACA1a, CheckDTAA5AC, CheckDTAA5A1ai, CheckDTAA5A1aA, CheckDTAA5AD1i, OSTotDTAA As Variant

Dim CheckDivDTAA_1aiii As Variant ' by sai on 28/03/2025


'Newly added by Bindu as per DE
'DTAA_5BB_OS = Sheet14.Range("DTAA_5BB_OS").value
CheckDivDTAA = Sheet14.Range("DivDTAA").value
CheckDTAA5ACA1a = Sheet14.Range("DTAA5ACA1a").value
CheckDTAA5AC = Sheet14.Range("DTAA5AC").value
CheckDTAA5A1ai = Sheet14.Range("DTAA5A1ai").value

CheckDivDTAA_1aiii = Sheet14.Range("DivDTAA_1aiii").value ' by sai on 28/03/2025

'Chandru
CheckDTAA5A1aA = Sheet14.Range("DTAA5A1aA").value
CheckDTAA5AD1i = Sheet14.Range("DTAA5AD1i").value


'OSTotDTAA = CheckDivDTAA + CheckDTAA5ACA1a + CheckDTAA5AC + CheckDTAA5A1ai + CheckDTAA5A1aA + CheckDTAA5AD1i
 OSTotDTAA = CheckDivDTAA + CheckDivDTAA_1aiii + CheckDTAA5ACA1a + CheckDTAA5AC + CheckDTAA5A1ai + CheckDTAA5A1aA + CheckDTAA5AD1i


If OSTotDTAA > 0 Then

    Sheet14.Range("DTAA.Upto15Of6").value = OSTotDTAA
Else
  Sheet14.Range("DTAA.Upto15Of6").value = ""


End If


'**********************************************************************************************************************
'commented by Chetan C M for AY 2026-27
'start--

''Modified by sai on 26/06/2025 as per Shubham Veram(BA) and updated the logic only for Amount field only
'

'uncommented by sadineni on 29/04/2026

If Not Application.Intersect(Sheet14.Range("os.Amount_income"), Range(Target.Address)) Is Nothing Then
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
         If Trim(Target.value <> "") Then
            fmsgbox """Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed for residents who may claim DTAA benefit under Schedule TR and FSI.""" & Chr(10) & Chr(10) ', vbOKOnly, "Alert"
        End If
      End If
End If

'--end

'********************************************* DE V0.6 Changes end by sai***************************************************

endline:

     Application.ScreenUpdating = True
     Application.EnableEvents = True
End Sub

Sub LockUnlockSurEduCessOS()
On Error Resume Next
    Sheet7.Unprotect Password:=getmsgstate
    Application.EnableEvents = False

    If Range("os.TaxBenefitTotal").value > 0 Or Sheet21.Range("DTAA_TAX_OS").value > 0 Or Sheet21.Range("DTAA_TAX_OS_1").value > 0 Or Sheet21.Range("DTAA_TAX_OS_2").value > 0 Then
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
Private Sub Worksheet_Deactivate()
On Error GoTo endline
Application.EnableEvents = False
Dim Activesht As Variant
Activesht = ActiveSheet.name
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet21.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim srcRangecells As Range
Dim srcRangecount As Long

Dim dstRangecells As Range
Dim dstRangecount As Long


Dim storeRange As Range
Dim lookupRange As Range

Dim srcSectionCode As String
Dim dstDectioncode As String


Set srcRangecells = Sheet14.Range("os.check").Cells
srcRangecount = Sheet14.Range("os.check").count

Set dstRangecells = Sheet21.Range("MAtchFUnction").Cells
dstRangecount = Sheet21.Range("MAtchFUnction").count

Set storeRange = Sheet21.Range("MatchedVAlue").Cells
Set lookupRange = Sheet14.Range("os.SourceAmount").Cells


If Sheet14.Range("os.TaxBenefitTotal").value > 0 Or Sheet21.Range("DTAA_TAX_OS").value > 0 Then
        Sheet7.Unprotect Password:=getmsgstate
        Sheet7.Range("Sheet9.Surcharge_ii").Locked = False
        Sheet7.Range("Sheet9.Surcharge_ii").Interior.Color = 13434879

        Sheet7.Range("Sheet9.EducationCess").Locked = False
        Sheet7.Range("Sheet9.EducationCess").Interior.Color = 13434879
        Sheet7.Protect Password:=getmsgstate

    Else
        Sheet7.Unprotect Password:=getmsgstate
        Sheet7.Range("Sheet9.Surcharge_ii").Locked = True
        Sheet7.Range("Sheet9.Surcharge_ii").Interior.Color = 16777215
        'Sheet7.Range("Sheet9.SurchargeOnTaxPayable").Value = "0"

        Sheet7.Range("Sheet9.EducationCess").Locked = True
        Sheet7.Range("Sheet9.EducationCess").Interior.Color = 16777215
        'Sheet7.Range("Sheet9.EducationCess").Value = "0"
        Sheet7.Protect Password:=getmsgstate
    End If


Dim i As Long
Dim j As Long

For j = 1 To dstRangecount
storeRange.item(j).value = ""
Next


For i = 1 To srcRangecount
    For j = 1 To dstRangecount
        srcSectionCode = srcRangecells.item(i).value
        dstDectioncode = dstRangecells.item(j).value


        If (Trim(srcSectionCode) = Trim(dstDectioncode)) Then

            storeRange.item(j).value = lookupRange.item(i).value
        End If
    Next
Next


endline:
'----------------Lock Password-------------------START---
   Sheet21.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Sheets(Activesht).Activate
Application.EnableEvents = True
End Sub
Sub setTblinfo_OS_local()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet14.Range("os.SourceDescription").count
 Set rangecells = Sheet14.Range("os.SourceDescription").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 end_os = ccount
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

'Change-14.05.2023.101 -> SIT-49871
'Sub LockUnlockDeduction57iia()
'    On Error Resume Next
'    Sheet14.Unprotect Password:=getmsgstate
'
'    If Range("os_FamilyPension").value > 0 And SheetALL.Range("bacValue").value = 2 Then
'        Sheet14.Range("os_Deduction57iia").Locked = False
'        Sheet14.Range("os_Deduction57iia").Interior.Color = &HCCFFCC
'    Else
'        Sheet14.Range("os_Deduction57iia").Locked = True
'        Sheet14.Range("os_Deduction57iia").value = ""
'        Sheet14.Range("os_Deduction57iia").Interior.Color = &HD8D8D8
'    End If
'    'Change, 17.04.2023.102.SIT 48765 A
'    If (Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Then
'        Sheet14.Range("os_Deduction57iia").value = 0
'    Else
'        Sheet14.Range("os_Deduction57iia").ClearContents
'    End If
'    'End Change
'Sheet14.Protect Password:=getmsgstate
'End Sub
'----
Sub OS_Table_D_F()

Application.EnableEvents = False
Sheet14.Unprotect Password:=getmsgstate


    If Mid(Sheet1.Range("Foreign_institutinal_investment").value, 1, 1) = "Y" Then
       If Mid(Sheet1.Range("Sheet1.115H").value, 1, 1) = "Y" Then
            Sheet14.Range("os.SourceDescription").Validation.Delete
            Sheet14.Range("os.SourceDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_ALL"
            
            'Sheet14.Range("os.SectionAct").Validation.Delete
            'Sheet14.Range("os.SectionAct").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_ALL"
            
            
        ElseIf Mid(Sheet1.Range("Sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
            
            Sheet14.Range("os.SourceDescription").Validation.Delete
            Sheet14.Range("os.SourceDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_RESandNOR"
            
            'Sheet14.Range("os.SectionAct").Validation.Delete
            'Sheet14.Range("os.SectionAct").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_RESandNOR"
            
            
        Else
            Sheet14.Range("os.SourceDescription").Validation.Delete
            Sheet14.Range("os.SourceDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_NRI"
        
        
            'Sheet14.Range("os.SectionAct").Validation.Delete
            'Sheet14.Range("os.SectionAct").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_NRI"
        
        End If
        
    ElseIf Mid(Sheet1.Range("Foreign_institutinal_investment").value, 1, 1) = "N" Then
        If Mid(Sheet1.Range("Sheet1.115H").value, 1, 1) = "Y" Then
            Sheet14.Range("os.SourceDescription").Validation.Delete
            Sheet14.Range("os.SourceDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_ALL1"
            
            'Sheet14.Range("os.SectionAct").Validation.Delete
            'Sheet14.Range("os.SectionAct").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_ALL1"
            
            
        ElseIf Mid(Sheet1.Range("Sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
            
            Sheet14.Range("os.SourceDescription").Validation.Delete
            Sheet14.Range("os.SourceDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_RESandNOR1"
            
            
            'Sheet14.Range("os.SectionAct").Validation.Delete
            'Sheet14.Range("os.SectionAct").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_RESandNOR1"
            
               
        Else
            Sheet14.Range("os.SourceDescription").Validation.Delete
            Sheet14.Range("os.SourceDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_NRI1"
        
            'Sheet14.Range("os.SectionAct").Validation.Delete
            'Sheet14.Range("os.SectionAct").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_NRI1"
        
        End If
    End If

Application.EnableEvents = True
Sheet14.Protect Password:=sPassword
End Sub

Sub OS_Table_E()

Application.EnableEvents = False
Sheet14.Unprotect Password:=getmsgstate


    If Mid(Sheet1.Range("Foreign_institutinal_investment").value, 1, 1) = "Y" Then
       If Mid(Sheet1.Range("Sheet1.115H").value, 1, 1) = "Y" Then
            Sheet14.Range("os.PTIDescription").Validation.Delete
            Sheet14.Range("os.PTIDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_PTI_115H"
            
        ElseIf Mid(Sheet1.Range("Sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
            
            Sheet14.Range("os.PTIDescription").Validation.Delete
            Sheet14.Range("os.PTIDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_PTI_RES"
            
        Else
            Sheet14.Range("os.PTIDescription").Validation.Delete
            Sheet14.Range("os.PTIDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_PTI_NRI"
        End If
        
    ElseIf Mid(Sheet1.Range("Foreign_institutinal_investment").value, 1, 1) = "N" Then
        If Mid(Sheet1.Range("Sheet1.115H").value, 1, 1) = "Y" Then
            Sheet14.Range("os.PTIDescription").Validation.Delete
            Sheet14.Range("os.PTIDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_PTI_115H1"
            
        ElseIf Mid(Sheet1.Range("Sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
            
            Sheet14.Range("os.PTIDescription").Validation.Delete
            Sheet14.Range("os.PTIDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_PTI_RES1"
            
        Else
            Sheet14.Range("os.PTIDescription").Validation.Delete
            Sheet14.Range("os.PTIDescription").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OS_PTI_NRI1"
        End If
    End If

Application.EnableEvents = True
Sheet14.Protect Password:=sPassword
End Sub

Sub test()
Application.EnableEvents = True
End Sub





