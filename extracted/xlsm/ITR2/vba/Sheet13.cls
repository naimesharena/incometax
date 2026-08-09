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
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
MsgBox ("If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F."), , "Message"

'Newly added by Bindu 27th March 25
 AddTotalSTCGA1
 AddTotalLTCGB1
 Cost_of_Improvement
 '--End

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


'added by Bindu for Table E Editable format
If Not Intersect(Range("IHLA.Eii2_StclSetoff20Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eii3_StclSetoff30Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eii4_StclSetoffAppRate"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eii5_StclSetoffDTAARate"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eiii2_StclSetoff20Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eiii3_StclSetoff30Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eiii4_StclSetoffAppRate20"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eiii5_StclSetoffDTAARate20"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eiii2_StclSetoff15Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eiv2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eiii4_StclSetoffAppRate"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eiii5_StclSetoffDTAARate"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eiv2_StclSetoff15Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Ev2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eiv3_StclSetoff30Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eiv5_StclSetoffDTAARate"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ev2_DTAArates"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Evi2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ev3_DTAArates"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Ev4_DTAArates"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ev2_StclSetoff15Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Evii2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ev3_StclSetoff30Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Ev4_StclSetoffAppRate"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ev5_StclSetoffDTAARate"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Evii8_LtclSetOff12.5Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ev6_LtclSetOff20Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Ev8_LtclSetOffDTAARate"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eviii2_StclSetoff12.5Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eviii2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eviii3_StclSetoff3012Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eviii4_StclSetoffAppRate12"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eviii5_StclSetoffDTAARate12"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eviii5_LtclSetOff1012Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eviii6_LtclSetOff2012Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eviii8_LtclSetOffDTAARate12"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Evi2_StclSetoff15Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Eix2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Evi3_StclSetoff30Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Evi4_StclSetoffAppRate"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Evi5_StclSetoffDTAARate"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Evi5_LtclSetOff10Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Eix8_LtclSetOff12.5Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Evi8_LtclSetOffDTAARate"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.E_LTCGv15_DTAArates"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.Ex2_StclSetoff20Per"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.E_LTCGv30_DTAArates"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.E_LTCGvapplicable_DTAArates"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.E_LTCGv_5_DTAArates"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.E_LTCGv5_DTAArates"), Range(Target.Address)) Is Nothing _
Or Not Intersect(Range("IHLA.Ex8_LtclSetOff12.5Per"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("IHLA.E_LTCGv6_DTAArates"), Range(Target.Address)) Is Nothing Then

Sheet13.Unprotect Password:=EfilingCommon.getmsgstate

'--Row wise 15%
If Sheet13.Range("IHLA.Eii2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Eii2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Eii3_StclSetoff30Per").value < 0 Then
       Sheet13.Range("IHLA.Eii3_StclSetoff30Per").value = 0
   End If
If Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").value < 0 Then
       Sheet13.Range("IHLA.Eii4_StclSetoffAppRate").value = 0
   End If
 If Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Eii5_StclSetoffDTAARate").value = 0
   End If
   
'--Row wise 20%
If Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Eiii2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").value < 0 Then
       Sheet13.Range("IHLA.Eiii3_StclSetoff30Per").value = 0
   End If
If Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").value < 0 Then
       Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate20").value = 0
   End If
If Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").value < 0 Then
       Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate20").value = 0
   End If



'--Row wise 30%
If Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").value < 0 Then
       Sheet13.Range("IHLA.Eiii2_StclSetoff15Per").value = 0
   End If
If Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Eiv2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").value < 0 Then
       Sheet13.Range("IHLA.Eiii4_StclSetoffAppRate").value = 0
   End If
If Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Eiii5_StclSetoffDTAARate").value = 0
   End If
   
   
'--Row Wise App Rates
If Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").value < 0 Then
       Sheet13.Range("IHLA.Eiv2_StclSetoff15Per").value = 0
   End If
If Sheet13.Range("IHLA.Ev2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Ev2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").value < 0 Then
       Sheet13.Range("IHLA.Eiv3_StclSetoff30Per").value = 0
   End If
If Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Eiv5_StclSetoffDTAARate").value = 0
   End If

'--Row Wise DTAA Rates
If Sheet13.Range("IHLA.Ev2_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.Ev2_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.Evi2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Evi2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Ev3_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.Ev3_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.Ev4_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.Ev4_DTAArates").value = 0
   End If
   
   
'--Row Wise 10% LTCG
If Sheet13.Range("IHLA.Ev2_StclSetoff15Per").value < 0 Then
       Sheet13.Range("IHLA.Ev2_StclSetoff15Per").value = 0
   End If
If Sheet13.Range("IHLA.Evii2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Evii2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value < 0 Then
       Sheet13.Range("IHLA.Ev3_StclSetoff30Per").value = 0
   End If
If Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").value < 0 Then
       Sheet13.Range("IHLA.Ev4_StclSetoffAppRate").value = 0
   End If
If Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Ev5_StclSetoffDTAARate").value = 0
   End If
 If Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").value < 0 Then
       Sheet13.Range("IHLA.Evii8_LtclSetOff12.5Per").value = 0
   End If
If Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value < 0 Then
       Sheet13.Range("IHLA.Ev6_LtclSetOff20Per").value = 0
   End If
 If Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Ev8_LtclSetOffDTAARate").value = 0
   End If


'--Row Wise 12.5% LTCG
If Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").value < 0 Then
       Sheet13.Range("IHLA.Eviii2_StclSetoff12.5Per").value = 0
   End If
If Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Eviii2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").value < 0 Then
       Sheet13.Range("IHLA.Eviii3_StclSetoff3012Per").value = 0
   End If
If Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").value < 0 Then
       Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate12").value = 0
   End If
If Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").value < 0 Then
       Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate12").value = 0
   End If
If Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").value < 0 Then
       Sheet13.Range("IHLA.Eviii5_LtclSetOff1012Per").value = 0
   End If
If Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").value < 0 Then
       Sheet13.Range("IHLA.Eviii6_LtclSetOff2012Per").value = 0
   End If
If Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").value < 0 Then
       Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate12").value = 0
   End If
   
'---Row Wise 20% LTCG
If Sheet13.Range("IHLA.Evi2_StclSetoff15Per").value < 0 Then
       Sheet13.Range("IHLA.Evi2_StclSetoff15Per").value = 0
   End If
If Sheet13.Range("IHLA.Eix2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Eix2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value < 0 Then
       Sheet13.Range("IHLA.Evi3_StclSetoff30Per").value = 0
   End If
If Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").value < 0 Then
       Sheet13.Range("IHLA.Evi4_StclSetoffAppRate").value = 0
   End If
If Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Evi5_StclSetoffDTAARate").value = 0
   End If
If Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value < 0 Then
       Sheet13.Range("IHLA.Evi5_LtclSetOff10Per").value = 0
   End If
If Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").value < 0 Then
       Sheet13.Range("IHLA.Eix8_LtclSetOff12.5Per").value = 0
   End If
If Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").value < 0 Then
       Sheet13.Range("IHLA.Evi8_LtclSetOffDTAARate").value = 0
   End If
   
   
'--Row Wise DTAA Rates LTCG
If Sheet13.Range("IHLA.E_LTCGv15_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.E_LTCGv15_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.Ex2_StclSetoff20Per").value < 0 Then
       Sheet13.Range("IHLA.Ex2_StclSetoff20Per").value = 0
   End If
If Sheet13.Range("IHLA.E_LTCGv30_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.E_LTCGv30_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.E_LTCGvapplicable_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.E_LTCGv_5_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.E_LTCGv5_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.E_LTCGv5_DTAArates").value = 0
   End If
If Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").value < 0 Then
       Sheet13.Range("IHLA.Ex8_LtclSetOff12.5Per").value = 0
   End If
 If Sheet13.Range("IHLA.E_LTCGv6_DTAArates").value < 0 Then
       Sheet13.Range("IHLA.E_LTCGv6_DTAArates").value = 0
   End If
   
   
   'Malli_18/09/2025
   If Not Application.Intersect(Range("CG_TABLE_E_Editable"), Range(Target.Address)) Is Nothing Then
            If UCase(Sheet13.Range("CG_TableE_Checkbox").value) = UCase("True") Then
             
                 If Sheet13.Range("TableE_Negativecount").value > 0 Then
                     MsgBox ("* Amount of set off claimed cannot exceed amount available for set off"), , "Alert"
                     Application.EnableEvents = False
                     Target.value = 0
                     GoTo endline
                 End If
             End If
    End If  'Malli_18/09/2025
   
   
End If



'LockUnlock112A9
'LockUnlock112A11
'=======================================================================================================
'Change Satya 18.07.2023.102 Updating the  Date Validation for the below fields,

'Date of purchase/construction of new residential house
'Date of purchase of new agricultural land
'Date of investment
'Date of purchase of plant and machinery

'Added by Riyaz on 14/03/2025
Dim s As Long
For s = 1 To Sheet13.Range("PropertySectionCount_LTCG").value
If Not Sheet13.Range("LTCG.B_SaleTrfDate" & s) = "" Then
If Not Application.Intersect(Sheet13.Range("LTCG.B_SaleTrfDate" & s), Sheet13.Range(Target.Address)) Is Nothing Then
    If Dformat(Sheet13.Range("LTCG.B_SaleTrfDate" & s).value, "22/07/2024") < Dformat(Sheet13.Range("LTCG.B_AcqDate" & s).value, "22/07/2024") Then
        fmsgbox ("Date of sale can't be earlier to Date of purchase / acquisition in Long term capital gain in schedule CG at S.no " & s)
        Sheet13.Range("LTCG.B_SaleTrfDate" & s).value = ""
    End If
End If
End If

'Added by Riyaz on 11/05/2026
If Not Application.Intersect(Sheet13.Range("LTCG.Bbi_AquisitCostNew" & s), Sheet13.Range(Target.Address)) Is Nothing Then
Sheet13.Range("LTCG.Bbi_AquisitCost" & s).formula = "=LTCG.Bbi_AquisitCost_editable" & s
End If

Next
'Malli AY_2025_26

Dim CG_A__STCG_Count_2025_26
 CG_A__STCG_Count_2025_26 = Sheet13.Range("PropertySectionCount_STCG").value
  
For i = 1 To CG_A__STCG_Count_2025_26
        If Not Intersect(Range("STCG.A_land_OR_building_AcqDate" & i), Range(Target.Address)) Is Nothing Then
                If (Target.value <> "") Then
                    
                    If Not FormatNCheckDate(Target.value) Then
                            MsgBox ("* ""Date should be in valid dd/mm/yyyy format, numeric."""), , "Error(s)"
                            Target.value = ""
                            Target.Select
                            GoTo endline
                     End If
'Changed by Riyaz 2025 to 2026 on 22/01/2026
                    If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2001") Then
                            MsgBox ("* ""Date of Purchase can be 01/04/2001 to 31/03/2026 at A1 In schedule CG"""), , "Error(s)"
                            Target.value = ""
                            GoTo endline
                    End If
'Changed by Riyaz 2025 to 2026 on 22/01/2026
                    If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then
                        MsgBox ("* ""Date of Purchase should not be after 31/03/2026 for AY 2026-27 A1 In schedule CG"""), , "Error(s)"
                        Target.value = ""
                        GoTo endline
                    End If
                    
                End If
        End If
        
        If Not Intersect(Range("STCG.A_land_OR_building_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
                If (Target.value <> "") Then
                
                If Not FormatNCheckDate(Target.value) Then
                            MsgBox ("* ""Date should be in valid dd/mm/yyyy format, numeric."""), , "Error(s)"
                            Target.value = ""
                            Target.Select
                            GoTo endline
                     End If
'Changed by Riyaz 2025 to 2026 on 22/01/2026
                If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then 'dpk1601
                       'MsgBox ("* ""Date of sale should not be after 31/03/2025 for AY 2025-26 A1 In schedule CG"""), , "Error(s)"
                       'by sai on 14/05/2025 as per DE V0.5 and V0.6
'                        MsgBox ("*""Please select the date within Financial year 2025-26"""), , "Error(s)"
                'by Konda on 15/04/2026 as per DE V0.3
                        MsgBox ("*""Date of Sale/Transfer cannot be beyond 31.03.2026"""), , "Error(s)"
                       Target.value = ""
                      GoTo endline
                End If
'Changed by Riyaz 2024 to 2025 on 22/01/2026
           'by sai on 14/05/2025 as per DE V0.5 and V0.6
                If Not ChkMinCGDate2627(Target.value, "31/03/2025") Then
                        MsgBox ("""Please check the correctness of date of sale."""), , "Error(s)"
                        Target.value = ""
                        GoTo endline
                    End If

                End If
        End If
         'Malli comented V0.2 change AY_2025-26
'       If Not Intersect(Range("STCG.A_land_OR_building_complcertifDate" & i), Range(Target.Address)) Is Nothing Then
'                If (Target.value <> "") Then
'                     If Not FormatNCheckDate(Target.value) Then
'                            MsgBox ("* ""Date should be in valid dd/mm/yyyy format, numeric."""), , "Error(s)"
'                            Target.value = ""
'                            Target.Select
'                            GoTo endline
'                     End If
'                     If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2024") Then
'                            MsgBox ("* ""Capital gain will be taxable in the year in which date of completion certificate is received. Please enter date correctly in field ""If Yes, date of completion certification"" at sl.no.A1."""), , "Error(s)"
'                            Target.value = ""
'                            GoTo endline
'                    End If
'
'                    If Not ChkMaxDOBDate(Target.value, "31/03/2025") Then
'                        MsgBox ("* ""Capital gain will be taxable in the year in which date of completion certificate is received. Please enter date correctly in field ""f Yes, date of completion certification"" at sl.no.A1."""), , "Error(s)"
'                        Target.value = ""
'                        GoTo endline
'                    End If
'                End If
'        End If
Next


'********************************added by sai on 14/05/2025 as per DE V0.5 and V0.6
'--Goto endline commented as its not updating Applicable lower rate filed By Bindu - SIT-94328

'AS per BA Remaining fields validations are removed on 20/06/2025

If Not Application.Intersect(Range("STCG.A8_AmtI"), Range(Target.Address)) Is Nothing Then
   ' For i = 1 To Sheet13.Range("STCG.A8_AmtI").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
'        If Trim(Range("STCG.A8_AmtI").Cells(i) <> "") Then
'             If Sheet13.Range("STCG.A8_AmtI").Cells(i) <> "" Then
            'Added by Shrutika(SIT-94521-13/06/25)
                If Trim(Target.value <> "") Then
                If Target.value <> "" Then
                    fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = ""
               ' GoTo endline
             End If
        End If
      End If
   ' Next
End If
 
 
'For LTCG

 If Not Application.Intersect(Range("LTCG.B10_LtcgAmt"), Range(Target.Address)) Is Nothing Then
    'For i = 1 To Sheet13.Range("LTCG.B10_LtcgAmt").count
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
'        If Trim(Range("LTCG.B10_LtcgAmt").Cells(i) <> "") Then
'             If Sheet13.Range("LTCG.B10_LtcgAmt").Cells(i) <> "" Then
              'Added by Shrutika(SIT-94521-13/06/25)
                If Trim(Target.value <> "") Then
                If Target.value <> "" Then
                      fmsgbox ("""Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed. Residents may claim DTAA benefit under Schedule TR and FSI.""")
                'Target.value = ""
                'GoTo endline
             End If
        End If
      End If
   ' Next
End If



'**************************************************************************************************

'-------------------------------
'Added by Riyaz on 11/03/2025
If Not Application.Intersect(Sheet13.Range("STCG.AII_CapitalLoss20r30rapplicablerate"), Sheet13.Range(Target.Address)) Is Nothing Then
    If WorksheetFunction.CountIf(Sheet13.Range("STCG.AII_CapitalLoss20r30rapplicablerate"), Target.value) > 1 Then
        Application.EnableEvents = False
            If Not (Target.value = "(Select)" Or Target.value = "") Then
                   fmsgbox "You cannot select same dropdown twice"
                   Target.value = "(Select)"
            End If
           '-----------------------
        Application.EnableEvents = True
    End If
End If
'Removed as per DE-sheet V0.3 0n 16-04-2026 by Konda
'If Not Application.Intersect(Sheet13.Range("LTCG.BII_Rate"), Sheet13.Range(Target.Address)) Is Nothing Then
'    If WorksheetFunction.CountIf(Sheet13.Range("LTCG.BII_Rate"), Target.value) > 1 Then
'        Application.EnableEvents = False
'            If Not (Target.value = "(Select)" Or Target.value = "") Then
'                   fmsgbox "You cannot select same dropdown twice"
'                   Target.value = "(Select)"
'            End If
'           '-----------------------
'        Application.EnableEvents = True
'    End If
'End If

If Not Intersect(Range("Deduction.54EC.AmountInvestd"), Range(Target.Address)) Is Nothing Then
   
   If Target.value > 5000000 Then
   
    fmsgbox ("Amount invested in specified/notified bonds section 54EC can not be more than 50 Lakhs.")
    Target.value = ""
    GoTo endline
   End If

End If

'CG_2024_25 Bindu

'If Not Intersect(Range("LTCG.B04d_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
'For SIT-92690 Changed by Shrutika on 28/05/2025
If Not Intersect(Range("LTCG.B04d_ExemptionGrandTotal_ii"), Range(Target.Address)) Is Nothing Then
   Dim B3i As Variant
        B3i = Sheet13.Range("LTCG.B04d_ExemptionGrandTotal_ii").value
    If B3i > 100000000 Then
        fmsgbox ("Amount against section 54F can not be more than 10 crores.")
        Target.value = ""
     GoTo endline
   End If
End If



If Trim(Mid(Sheet1.Range("sheet1.ResidentialStatus1"), 1, 2)) = "NR" Then
'Changed by Riyaz on 18/02/2025
'If Not Intersect(Range("LTCG.B07dv_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
If Not Intersect(Range("LTCG.Biii5dii_ExemptionAmount"), Range(Target.Address)) Is Nothing Then
   If Target.value > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If

End If
  
'Added by Riyaz on 23/02/2025
 If Not Intersect(Range("LTCG.B07div_ExemptionGrandTotal_ii"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.B07div_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
        Dim B6eii As Variant
        B6eii = Sheet13.Range("LTCG.B07div_ExemptionGrandTotal_ii").value + Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value
        If B6eii > 100000000 Then
            fmsgbox ("Amount against section 54F can not be more than 10 crores.")
            Target.value = ""
        End If
 End If
  
  
End If




If Not Intersect(Range("LTCG.B03d_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
   
   If Target.value > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If

End If

'Added by Riyaz on 23/02/2025
If Not Intersect(Range("LTCG.B04di_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
   Dim B3iie As Variant
        B3iie = Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value
   If B3iie > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If

End If
'--End



If Not Intersect(Range("B5_DeductionTotal"), Range(Target.Address)) Is Nothing Then
  If Target.value > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If
End If


If Not Intersect(Range("LTCG.B06d_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
   
   If Target.value > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If

End If


'Added by Riyaz on 23/02/2025
If Not Intersect(Range("LTCG.B07dii_ExemptionGrandTotal"), Range(Target.Address)) Is Nothing Then
   Dim B6ie As Variant
        B6ie = Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").value
   If B6ie > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If

End If
'--End

If Not Intersect(Range("LTCG.8.DeductionUnder54F"), Range(Target.Address)) Is Nothing Then
   If Target.value > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If
End If


'Added by Riyaz on 23/02/2025
If Not Intersect(Range("LTCG.B7dii_ExemptionAmount_i"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("LTCG.B7dii_ExemptionAmount"), Range(Target.Address)) Is Nothing Then
   Dim B9e As Variant
        B9e = Sheet13.Range("LTCG.B7dii_ExemptionAmount_i").value + Sheet13.Range("LTCG.B7dii_ExemptionAmount").value
   If B9e > 100000000 Then
    fmsgbox ("Amount against section 54F can not be more than 10 crores.")
    Target.value = ""
    GoTo endline
   End If
End If

For i = 1 To Sheet13.Range("Deduction.54.DateOfPursConst").count
     If Not Intersect(Range("Deduction.54.DateOfPursConst"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value) Then
            fmsgbox ("Date Should be in DD/MM/YYYY format only")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
        
        If Not checkFirstDateBefore(Target.value, Sheet52.Range("sheet9.Date").value) Then '
                fmsgbox ("Date can not be after System Date") '
                Target.value = ""
                Target.Select
                GoTo endline
        End If
      End If
Next


'Newly added by Bindu
'54 Table
For i = 1 To Sheet13.Range("Deduction.54.AmtDedClaimed").count
     If Not Intersect(Range("Deduction.54.AmtDedClaimed"), Range(Target.Address)) Is Nothing Then
        
        If Target.value > 100000000 Then
            fmsgbox ("Please ensure Amount of deduction in Section 54 cannot exceed Rs.10 crore")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
    End If

Next

'54F
For i = 1 To Sheet13.Range("Deduction.54F.AmtDedClaimed").count
     If Not Intersect(Range("Deduction.54F.AmtDedClaimed"), Range(Target.Address)) Is Nothing Then
        
        If Target.value > 100000000 Then
            fmsgbox ("Please ensure Amount of deduction in Section 54F cannot exceed Rs.10 crore")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
    End If

Next



For i = 1 To Sheet13.Range("Deduction.54B.DateOfPurLand").count
     If Not Intersect(Range("Deduction.54B.DateOfPurLand"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value) Then
            fmsgbox ("Date Should be in DD/MM/YYYY format only")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
        
        If Not checkFirstDateBefore(Target.value, Sheet52.Range("sheet9.Date").value) Then '
                fmsgbox ("Date can not be after System Date") '
                Target.value = ""
                Target.Select
                GoTo endline
        End If
      End If
Next

For i = 1 To Sheet13.Range("Deduction.54EC.DateOfInvestment").count
     If Not Intersect(Range("Deduction.54EC.DateOfInvestment"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value) Then
            fmsgbox ("Date Should be in DD/MM/YYYY format only")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
        
        If Not checkFirstDateBefore(Target.value, Sheet52.Range("sheet9.Date").value) Then 'dpk1601
                fmsgbox ("Date can not be after System Date") '
                Target.value = ""
                Target.Select
                GoTo endline
        End If
      End If
Next

For i = 1 To Sheet13.Range("Deduction.54F.DateOfPurc").count
     If Not Intersect(Range("Deduction.54F.DateOfPurc"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value) Then
            fmsgbox ("Date Should be in DD/MM/YYYY format only")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
        
        If Not checkFirstDateBefore(Target.value, Sheet52.Range("sheet9.Date").value) Then '
                fmsgbox ("Date can not be after System Date") '
                Target.value = ""
                Target.Select
                GoTo endline
        End If
      End If
Next

For i = 1 To Sheet13.Range("Deduction.54GA.DateOfPurc").count
     If Not Intersect(Range("Deduction.54GA.DateOfPurc"), Range(Target.Address)) Is Nothing Then
        
        If Not FormatNCheckDate(Target.value) Then
            fmsgbox ("Date Should be in DD/MM/YYYY format only")
            Target.value = ""
            Target.Select
            GoTo endline
        End If
        
        If Not checkFirstDateBefore(Target.value, Sheet52.Range("sheet9.Date").value) Then '
                fmsgbox ("Date can not be after System Date") '
                Target.value = ""
                Target.Select
                GoTo endline
        End If
      End If
Next

'end Change Satya 18.07.2023.102 Updating the  Date Validation
'===================================================================================================

For i = 1 To Sheet13.Range("Deduction.54.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        'CG_E11_AY_2023-24 OLD  Malli
        'If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
                'fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1a In schedule CG ") 'dpk1601
                'CG_C11_AY_2024-25 Change
                'If Not ChkMaxDOBDate(Target.value, "31/03/2024") Then 'dpk1601
                'fmsgbox ("Date of transfer of original asset should not be after 31/03/2024 for AY 2024-25 at D1a In schedule CG ") 'dpk1601
                'Malli_CG_C322_AY_2025_26
                If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then 'dpk1601
                fmsgbox ("""Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1a In schedule CG""") 'dpk1601
                '---------------------------------
                Target.value = ""
                GoTo endline
        End If
      End If
Next


'54 Account
For i = 1 To Sheet13.Range("Deduction.54.DateofDeposit").count
     If Not Intersect(Range("Deduction.54.DateofDeposit"), Range(Target.Address)) Is Nothing Then
     
       If Not FormatNCheckDate(Target.value) Then
            fmsgboxStatus ("Date of deposit in CG sheet must be a valid dd/mm/yyyy format")
            Target.value = ""
            GoTo endline
        End If
     End If
Next





'CG_NEW_C3 2024-25 Bindu
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
                'fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If

Next
'---Ends
For i = 1 To Sheet13.Range("Deduction.54B.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54B.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        'CG_E12_AY_2023-24 OLD  Malli
        'If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
                'fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1b In schedule CG ") 'dpk1601
                'CG_C12_AY_2024-25 Change
                'If Not ChkMaxDOBDate(Target.value, "31/03/2024") Then 'dpk1601
                'fmsgbox ("Date of transfer of original asset should not be after 31/03/2024 for AY 2024-25 at D1b In schedule CG ") 'dpk1601
                'Malli_CG_C332_AY_2025_26
                If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then 'dpk1601
                fmsgbox ("""Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1b In schedule CG""") 'dpk1601
                '-----------------------
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'CG_NEW_C5 2024-25 Bindu
For i = 1 To Sheet13.Range("Deduction.54B.DateofDeposit").count
     If Not Intersect(Range("Deduction.54B.DateofDeposit"), Range(Target.Address)) Is Nothing Then
     
       If Not FormatNCheckDate(Target.value) Then
            fmsgboxStatus ("Date of deposit in CG sheet must be a valid dd/mm/yyyy format")
            Target.value = ""
            GoTo endline
        End If
     End If
Next

'CG_NEW_C6 2024-25 Bindu
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
                'fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If

Next
'--Ends

For i = 1 To Sheet13.Range("Deduction.54EC.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54EC.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
        'CG_E13_AY_2023-24 OLD  Malli
        'If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
                'fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1d In schedule CG ") 'dpk1601
                'CG_C13_AY_2024-25 Change
                'If Not ChkMaxDOBDate(Target.value, "31/03/2024") Then 'dpk1601
                'fmsgbox ("Date of transfer of original asset should not be after 31/03/2024 for AY 2024-25 at D1c In schedule CG ") 'dpk1601
                'Malli_CG_C342_AY_2025_26
                If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then 'dpk1601
                fmsgbox ("""Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1c In schedule CG""") 'dpk1601
                '-------------------
                Target.value = ""
                GoTo endline
        End If
      End If
Next



For i = 1 To Sheet13.Range("Deduction.54F.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54F.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
         
      'Malli------AY_2024_25_OLD
     'If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
     'fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1d In schedule CG ") 'dpk1601
    'C348_AY_2025_26
    If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then 'dpk1601
     fmsgbox ("""Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1d In schedule CG""") 'dpk1601
    '---------------------------------------
                Target.value = ""
                GoTo endline
        End If
      End If
Next

'CG_NEW_C8 2024-25 Bindu
For i = 1 To Sheet13.Range("Deduction.54F.DateofDeposit").count
     If Not Intersect(Range("Deduction.54F.DateofDeposit"), Range(Target.Address)) Is Nothing Then

       If Not FormatNCheckDate(Target.value) Then
            fmsgboxStatus ("Date of deposit in CG sheet must be a valid dd/mm/yyyy format")
            Target.value = ""
            GoTo endline
        End If
     End If
Next

'CG_NEW_C9 2024-25 Bindu
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
                
                'fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                Range(Target.Address) = ""
                Range(Target.Address).Select
                GoTo endline
            End If
        End If
  End If

Next
'--Ends


For i = 1 To Sheet13.Range("Deduction.54G.DateOfTransfer").count
     If Not Intersect(Range("Deduction.54G.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
         
        If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1f In schedule CG ") 'dpk1601
                 
                Target.value = ""
                GoTo endline
        End If
      End If
Next


'Newly commented by Bindu - 2024-25 Version
'For i = 1 To Sheet13.Range("Deduction.54GA.DateOfTransfer").count
'     If Not Intersect(Range("Deduction.54GA.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
'
'        If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
'                fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1g In schedule CG ") 'dpk1601
'                Target.value = ""
'                GoTo Endline
'        End If
'      End If
'Next

For i = 1 To Sheet13.Range("Deduction.115F.DateOfTransfer").count
     If Not Intersect(Range("Deduction.115F.DateOfTransfer"), Range(Target.Address)) Is Nothing Then
         
'        If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
'                fmsgbox ("Date of transfer of original asset should not be after 31/03/2023 for AY 2023-24 at D1h In schedule CG ") 'dpk1601

'IPIP-71883
        If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then
                fmsgbox ("Date of transfer of original asset should not be after 31/03/2026 for AY 2026-27 at D1e In schedule CG ")
                 
                Target.value = ""
                GoTo endline
        End If
      End If
Next

Dim ACG_A_Count
ACG_A_Count = Sheet13.Range("PropertySectionCount_STCG").value
'For i = 1 To ACG_A_Count
'    If Not Intersect(Range("STCG.A_AcqDate" & i), Range(Target.Address)) Is Nothing Then
'        If (Target.value <> "") Then
'        If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2001") Then
'            fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2022 at Srno A1 In schedule CG ")
'            Target.value = ""
'            GoTo endline
'        End If
'        If Not ChkMaxDOBDate(Target.value, "31/03/2022") Then
'            fmsgbox ("*Date of Purchase should not be after 31/03/2022 for AY 2022-23 A1 In schedule CG")
'            Target.value = ""
'            GoTo endline
'        End If
'        End If
'    End If
'
'    If Not Intersect(Range("STCG.AB_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
'        If (Target.value <> "") Then
'        If Not ChkMaxDOBDate(Target.value, "31/03/2022") Then
'            fmsgbox ("*Date of sale should not be after 31/03/2022 for AY 2022-23 A1 In schedule CG")
'            Target.value = ""
'            GoTo endline
'        End If
'        End If
'    End If
'Next

Dim BCG_A_Count
BCG_A_Count = Sheet13.Range("PropertySectionCount_LTCG").value
For i = 1 To BCG_A_Count
    If Not Intersect(Range("LTCG.B_AcqDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
        
         
        'CG_E3_AY_2023-24 OLD   'Malli
        If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2001") Then
            'CG_E3_AY_2023-24 OLD
            'fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2023 at Srno B1 In schedule CG ") 'dpk1601
            'Malli--AY_2024_25 OLD-----
            'CG_C3_AY_2024-25 Change
            'fmsgbox ("* Date of Purchase can be 01/04/2001 to 31/03/2024 at Srno B1 In schedule CG ") 'dpk1601
            'Malli--CG_E123_AY_2025_26
            fmsgbox ("* ""Please enter date from 01/04/2001 to 31/03/2026"" at B1 In schedule CG")
            '---------------------------
            Target.value = ""
            GoTo endline
        End If
        
        'If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then
        'Malli----AY_2024_25 OLD
       'If Not ChkMaxDOBDate(Target.value, "31/03/2024") Then
        'CG_E123_AY_2025_26
        If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then
            'fmsgbox ("*Date of Purchase should not be after 31/03/2023 for AY 2023-24 B1 In schedule CG")
            'Malli--AY_2024_25 OLD
            'fmsgbox ("*Date of Purchase should not be after 31/03/2024 for AY 2024-25 B1 In schedule CG")
            'CG_E123_AY_2025_26
'            fmsgbox ("* ""Please enter date from 01/04/2001 to 31/03/2026"" for AY 2025-26 at B1 In schedule CG")
             fmsgbox ("* ""Please enter date from 01/04/2001 to 31/03/2026.""") 'Added by Aavula SIT-112478
                         
            '------------------
            Target.value = ""
            GoTo endline
        End If
        End If
    End If
    

    If Not Intersect(Range("LTCG.B_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
        'CG_E4_AY_2023-24 OLD 'Malli
        'If Not ChkMaxDOBDate(Target.value, "31/03/2023") Then 'dpk1601
        'fmsgbox ("*Date of sale should not be after 31/03/2023 for AY 2023-24 B1 In schedule CG") 'dpk1601
        
        'Malli------AY_2024_25 OLD
        'CG_C4_AY_2024-25 Change
        'If Not ChkMaxDOBDate(Target.value, "31/03/2024") Then 'dpk1601
        'fmsgbox ("*Date of sale should not be after 31/03/2024 for AY 2024-25 B1 In schedule CG") 'dpk1601
         'Malli_CG_E124_AY_2025_26

        'Malli--N124_AY_2025_26
'        If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2001") Then
'            fmsgbox ("* ""Please enter date from  01/04/2001 to 31/03/2025"" at B1 In schedule CG")
'            Target.value = ""
'            GoTo endline
'        End If
        
        
        
    If Not ChkMaxDOBDate(Target.value, "31/03/2026") Then 'dpk1601
         'fmsgbox ("* ""Please enter date from  01/04/2001 to 31/03/2025"" for AY 2024-25 B1 In schedule CG")
      'Shrutika(27-02-25)
'         fmsgbox ("* ""Please enter date from  01/04/2001 to 31/03/2025"" for AY 2025-26 B1 In schedule CG")
       'by sai on 14/05/2025 as per V0.5 and V0.6
'         fmsgbox ("""Please select the date within Financial year 2025-26""")
        'by Konda on 17/04/2026 as per V0.3
        fmsgbox ("""Date of Sale/Transfer cannot be beyond 31.03.2026.""")
            Target.value = ""
            GoTo endline
    End If
        
    'by sai on 14/05/2025 as per V0.5 and V0.6
     '*****************************************************************************************
        If Not ChkMinCGDate2627(Target.value, "31/03/2025") Then
            fmsgbox ("""Please check the correctness of date of sale.""")
            Target.value = ""
            GoTo endline
        End If
     '****************************************************************************************
        
        
    
        If Not Dformat(Target.value, "") <= Dformat(Sheet52.Range("sheet9.Date").value, "") Then
            fmsgbox ("* ""Date of sale should not be after system date for AY 2026-27 B1 In schedule CG""")
            Target.value = ""
            GoTo endline
        End If

        '-----------------------
        
        End If
    End If
    
   'Malli comented AY_2025_26 V0.2 change
'Change-15.12.2022.104
'    If Not Intersect(Range("LTCG.B_SaleTrfDate" & i), Range(Target.Address)) Is Nothing Then
'        LTCG_SaleTrfDate Target
''        If (Target.value <> "") Then
''        If Not ChkMinInclusiveDate(Dformat(Target.value, ""), "2021-04-01") Then
''            fmsgbox ("*Date of sale should not be before 01/04/2021 for AY 2022-23 B1 In schedule CG")
''            Target.value = ""
''            GoTo endline
''        End If
''        End If
'    End If
'------------------------------------
'malli --CG_N124 - -AY_2025_26

If Not Intersect(Range("LTCG.B_land_OR_building_WtherChrgble" & i), Range(Target.Address)) Is Nothing Then
   If (Target.value <> "") Then
         If Target.value <> "" And Target.value = "Yes" Then
           Sheet13.Unprotect Password:=sPassword
           Sheet13.Range("LTCG.B_land_OR_building_complcertifDate" & i).MergeArea.Locked = False
           Sheet13.Range("LTCG.B_land_OR_building_complcertifDate" & i).MergeArea.value = ""
           Sheet13.Range("LTCG.B_land_OR_building_complcertifDate" & i).Interior.Color = (&HCCFFCC)
         Else
           Sheet13.Unprotect Password:=sPassword
           Sheet13.Range("LTCG.B_land_OR_building_complcertifDate" & i).MergeArea.value = ""
           Sheet13.Range("LTCG.B_land_OR_building_complcertifDate" & i).MergeArea.Locked = True
           Sheet13.Range("LTCG.B_land_OR_building_complcertifDate" & i).MergeArea.Interior.Color = (&HD8D8D8)
         End If
   End If
End If
'-------------------------------------
'Malli--CG_N125--AY_2025_26
If Not Intersect(Range("LTCG.B_land_OR_building_complcertifDate" & i), Range(Target.Address)) Is Nothing Then
        If (Target.value <> "") Then
         
          If EfilingCommon.checkFirstDateBefore(Target.value, "31/03/2024") Then
            fmsgbox ("* ""Please enter the date between 01/04/2024 to 31/03/2025"" in sl. no. 1 of LTCG")
            Target.value = ""
            GoTo endline
          End If
        
         If Not ChkMaxDOBDate(Target.value, "31/03/2025") Then
            fmsgbox ("* ""Please enter the date between 01/04/2024 to 31/03/2025"" in sl. no. 1 of LTCG")
            Target.value = ""
            GoTo endline
          End If
          
        End If
End If
'--------------------------

Next
'------------------------------------
'AY_2024_25 CHANGE------------------------

'Malli---------------------
 Dim CG_A__STCG_Count, CG_B__LTCG_Count
 CG_A__STCG_Count = Sheet13.Range("PropertySectionCount_STCG").value
 CG_B__LTCG_Count = Sheet13.Range("PropertySectionCount_LTCG").value

 For i = 1 To CG_A__STCG_Count
            If Not Application.Intersect(Sheet13.Range("CG_A_StateCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then
    
'      Dim PIN_targetadd, state_targetadd As Variant
'      state_targetadd = Target.Address
'      state_targetadd = Replace(state_targetadd, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "P", "Q")
'
'    Call state_Validation(PIN_targetadd, state_targetadd)

      Dim PIN_targetadd, state_targetadd As String
       
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "P", "Q")
      
      Dim ws2 As Worksheet
                             'Set ws2 = Worksheets("House Property")  '23/09/24  'CG
  
      Set ws2 = Worksheets("CG")  ' Malli  17/10/2024
      If state_Validation(ws2, PIN_targetadd, state_targetadd) = False Then Sheet13.Range(Replace(state_targetadd, "P", "Q")).value = ""
    
End If
 
If Not Application.Intersect(Sheet13.Range("CG_A_PinCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then
 
'     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As Variant
'
'      pinrng_chk = Target.Address
'      pinrng_chk = Replace(pinrng_chk, "$", "")
'      staterng_chk = Replace(pinrng_chk, "Q", "P")
'      cntryrng_chk = Replace(pinrng_chk, "Q", "R")
'      ZIPrng_chk = Replace(pinrng_chk, "Q", "S")
'
'      Call PinState_codes_validation(pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
     
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

'          Dim PIN_targetadd_LTCG, state_targetadd_LTCG As Variant
'          state_targetadd_LTCG = Target.Address
'          state_targetadd_LTCG = Replace(state_targetadd_LTCG, "$", "")
'          PIN_targetadd_LTCG = Replace(state_targetadd_LTCG, "P", "Q")
'
'        Call state_Validation(PIN_targetadd_LTCG, state_targetadd_LTCG)

      
      Dim PIN_targetadd_LTCG, state_targetadd_LTCG As String
       
      state_targetadd_LTCG = Replace(Target.Address, "$", "")
      PIN_targetadd_LTCG = Replace(state_targetadd_LTCG, "P", "Q")
      
      Dim ws3 As Worksheet
      Set ws3 = Worksheets("CG")  '23/09/24
                            

      If state_Validation(ws3, PIN_targetadd_LTCG, state_targetadd_LTCG) = False Then Sheet13.Range(Replace(state_targetadd_LTCG, "P", "Q")).value = ""
    
      
      
    End If
      

 
If Not Application.Intersect(Sheet13.Range("CG_B_PinCodeOfBuyer" & i), Sheet13.Range(Target.Address)) Is Nothing Then
 
'     Dim pinrng_chk_LTCG, staterng_chk_LTCG, cntryrng_chk_LTCG, ZIPrng_chk_LTCG As Variant
'
'      pinrng_chk_LTCG = Target.Address
'      pinrng_chk_LTCG = Replace(pinrng_chk_LTCG, "$", "")
'      staterng_chk_LTCG = Replace(pinrng_chk_LTCG, "Q", "P")
'      cntryrng_chk_LTCG = Replace(pinrng_chk_LTCG, "Q", "R")
'      ZIPrng_chk_LTCG = Replace(pinrng_chk_LTCG, "Q", "S")
'
'      Call PinState_codes_validation(pinrng_chk_LTCG, staterng_chk_LTCG, cntryrng_chk_LTCG, ZIPrng_chk_LTCG)


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


'----------------------------------------------------
'END CHANGE---------Malli-----------------------

        'If Target.Validation.Type = 3 Then
        
            'new change
    Dim CG_A_Count, CG_B_Count
     CG_A_Count = Sheet13.Range("PropertySectionCount_STCG").value
    CG_B_Count = Sheet13.Range("PropertySectionCount_LTCG").value

    For i = 1 To CG_A_Count
        If Not Intersect(Range("CG_A_StateCodeOfBuyer" & i), Range(Target.Address)) Is Nothing Then
            
            If isdropdownblank(Range(Target.Address).value) Then
            'Change-15.12.2022.104
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
                
               ' Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"  'Malli 17/10/2024
                
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
                 If Sheet13.Range(Replace(Target.Address, "P", "Q")).value = "" Then 'Malli
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
                If (UCase(Sheet13.Range(Replace(Target.Address, "R", "P")).value)) <> "" Then 'Konda updated on 10-06-2026
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"
                End If
            End If
            '17/10/2024--------------Malli
            ElseIf Mid(Target.value, 1, 2) = "91" And Mid(Sheet13.Range(Replace(Target.Address, "R", "P")).value, 1, 2) = "99" Then
            Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
            '--------------------------
            End If
         End If
    Next
    'End Change
    
    For i = 1 To CG_B_Count
        If Not Intersect(Range("CG_B_StateCodeOfBuyer" & i), Range(Target.Address)) Is Nothing Then
           
            If isdropdownblank(Range(Target.Address).value) Then
            'Change-15.12.2022.104
                Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
                'Sheet13.Range(Replace(Target.Address, "P", "R")).value = "91-INDIA"   'Malli  17/10/2024
                
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
                If Sheet13.Range(Replace(Target.Address, "P", "Q")).value = "" Then 'Malli
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
            '17/10/2024--------------Malli
            ElseIf Mid(Target.value, 1, 2) = "91" And Mid(Sheet13.Range(Replace(Target.Address, "R", "P")).value, 1, 2) = "99" Then
            Sheet13.Range(Replace(Target.Address, "P", "R")).value = "(Select)"
            '--------------------------
            End If
            
         End If
    Next
    
    'new change
'End If


LockUnlock115AD9
LockUnlock115AD11

Dim temp, temp1, Temp2, Temp3, targetadd As Variant
If Not Application.Intersect(Sheet13.Range("STCG.A8_CertiStatus"), Range(Target.Address)) Is Nothing Then
    temp = Range(Replace(Target.Address, "M", "G")).value

    Sheet13.Unprotect Password:=sPassword
    targetadd = Target.Address
     'Replace(TargetAdd, "$", "")
'Konda updated on 16-04-2026--DE-Sheet ---V0.3
     If Range(targetadd).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'     If Range(targetadd).value = "Yes" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
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
'Konda updated on 16-04-2026--DE-Sheet ---V0.3
     If Range(Replace(targetadd, "L", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'     If Range(Replace(targetadd, "L", "M")).value = "Yes" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
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
'Konda updated on 16-04-2026--DE-Sheet ---V0.3
     If Range(Replace(targetadd, "O", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'     If Range(Replace(targetadd, "O", "M")).value = "Yes" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
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
'Konda updated on 16-04-2026--DE-Sheet ---V0.3
     If Range(targetadd).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'     If Range(targetadd).value = "Yes" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
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
'Konda updated on 16-04-2026--DE-Sheet ---V0.3
     If Range(Replace(targetadd, "L", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'     If Range(Replace(targetadd, "L", "M")).value = "Yes" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
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
'Konda updated on 16-04-2026--DE-Sheet ---V0.3
     If Range(Replace(targetadd, "O", "M")).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'     If Range(Replace(targetadd, "O", "M")).value = "Yes" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
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

'Commented by Riyaz on 24/02/2025
'If Target.Validation.Type = 1 Or Target.Validation.Type = 3 Then
'    Application.ScreenUpdating = False
'        AddTotalSTCGA1
'
'        AddTotalLTCGB1
'    Application.ScreenUpdating = True
'End If

AddTotalSTCGA1
AddTotalLTCGB1
'Newly added by Bindu on 9th March 25
Cost_of_Improvement

Application.EnableEvents = False
 
'Commented by Riyaz on 11/03/2025
' If Not Intersect(Target, Range("Q6:Q12")) Is Nothing Then
' Sheet13.Range("Q6:Q12").SpecialCells(xlCellTypeFormulas).ClearContents
' End If


'For SIT-90074 by sai on 19/03/2025*************************************************************************
'Commented by Riyaz on 23/01/2026
'If Not Application.Intersect(Sheet13.Range("LTCG.B3bii_ImproveCost"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation"), Range(Target.Address)) Is Nothing Then
'  If Sheet13.Range("LTCG.B3bii_ImproveCost").value < Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation").value Then
'        fmsgboxStatus "3bii cant be less than 3biia"
'        Sheet13.Range("LTCG.B3biia_ImproveCost_WithoutIndexation").value = ""
'  End If
'End If

'For SIT-90075 by sai on 19/03/2025
If Not Application.Intersect(Sheet13.Range("LTCG.B3bi_AquisitCost"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet13.Range("LTCG.B3bia_AcqIndexation"), Range(Target.Address)) Is Nothing Then
  If Sheet13.Range("LTCG.B3bi_AquisitCost").value < Sheet13.Range("LTCG.B3bia_AcqIndexation").value Then
        fmsgboxStatus "3bi cant be less than 3bia"
        Sheet13.Range("LTCG.B3bia_AcqIndexation").value = ""
  End If
End If

'***************************************************************************************




If Not Intersect(Range("STCG.A7_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
'If Not Intersect(Range("STCG.A7_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                If Sheet13.Range(Target.Address).value = "2020-21" Or Sheet13.Range(Target.Address).value = "2019-20" Then
                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "54B"
                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = True
                Else
                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "(Select)"
                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = False
                End If
                Sheet13.Protect Password:=sPassword
            End If

            If Not Intersect(Range("STCG.A7_DeductionSection_1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                If Sheet13.Range(Replace(Target.Address, "I", "H")).value = "2016-17" Then
               
                  If Sheet13.Range(Target.Address).value = "54B" Then
                    fmsgboxStatus "Please select correct Section for the corresponding year"
                    Sheet13.Range(Target.Address).value = "(Select)"
                  End If
                End If
                Sheet13.Protect Password:=sPassword
            End If

''Malli ________________________________________
''CG_E9_AY_2023-24 OLD
'If Not Intersect(Range("LTCG.B9_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet13.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'                If Sheet13.Range(Target.Address).value = "2020-21" Then
'
'                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "54B"
'                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = True
'                Else
'                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "(Select)"
'                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = False
'                End If
'                Sheet13.Protect Password:=sPassword
'            End If
''Malli ____________________________________________________


'For SIT-89202 by sai on 04/03/2025
If Not Intersect(Range("LTCG.B9_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
    LTCGSecUndDedClaim (Target.row)
End If


            
            
'            If Not Intersect(Range("LTCG.B9_DeductionSection_1"), Range(Target.Address)) Is Nothing Then
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet13.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'                If Sheet13.Range(Replace(Target.Address, "I", "H")).value = "2016-17" Then
'                  If Sheet13.Range(Target.Address).value = "54B" Then
'                    fmsgboxStatus "Please select correct Section for the corresponding year"
'                    Sheet13.Range(Target.Address).value = "(Select)"
'                  End If
'                End If
'                Sheet13.Protect Password:=sPassword
'            End If

If Not Intersect(Range("STCG.A6_StatusUnutilizedCG"), Range(Target.Address)) Is Nothing Then

'If Target.name.name = "STCG.A6_StatusUnutilizedCG" Then
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


If Not Intersect(Range("LTCG.A9_StatusUnutilizedCG"), Range(Target.Address)) Is Nothing Then

'If Target.name.name = "LTCG.A9_StatusUnutilizedCG" Then
Sheet13.Unprotect Password:=getmsgstate
Application.EnableEvents = False
'Malli----
'AY_2024_25_OLD
'Rangename = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1" '||LTCG.B9_AsstAcquiredYer_2||LTCG.B9_AmtUtilised_2||LTCG.B9_AmtNotUsed_2"
'AY_CG_C283_2025_26_NEW
Rangename = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1"
'----------------------
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

'------------------------------

'CG_C7_AY_2024-25 Change
'Added by Malli

If Not Intersect(Range("LTCG.Bdi_ExemptionAmount1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                 
                  If Sheet13.Range(Target.Address).value > 100000000 Then
                    fmsgboxStatus "Amount against section 54 and 54F shall restricted to 10 crores."

                    'fmsgboxStatus "The value in Amount of Deduction under section 54  can not be more than 10 crores in schedule CG"
                    Sheet13.Range(Target.Address).value = ""
                  
                End If
                Sheet13.Protect Password:=sPassword
            End If


If Not Intersect(Range("LTCG.Bdv_ExemptionAmount1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                 
                  If Sheet13.Range(Target.Address).value > 100000000 Then
                    'fmsgboxStatus "The value in Amount of Deduction under section 54  can not be more than 10 crores in schedule CG"
                    fmsgboxStatus "Amount against section 54 and 54F shall restricted to 10 crores."
                    Sheet13.Range(Target.Address).value = ""
                  
                End If
                Sheet13.Protect Password:=sPassword
            End If


'---------------------------------
    


endline:
Application.EnableEvents = True

Sheet13.Protect Password:=sPassword

End Sub
Private Sub Worksheet_Changeold1(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim Rangename, i As Variant
Application.EnableEvents = False

'LockUnlock112A9
'LockUnlock112A11

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

If Target.Validation.Type = 1 Or Target.Validation.Type = 3 Then
    Application.ScreenUpdating = False
        AddTotalSTCGA1
        AddTotalLTCGB1
    Application.ScreenUpdating = True
End If

 If Not Intersect(Target, Range("Q6:Q12")) Is Nothing Then
 Sheet13.Range("Q6:Q12").SpecialCells(xlCellTypeFormulas).ClearContents
 End If



If Not Intersect(Range("STCG.A7_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                If Sheet13.Range(Target.Address).value = "2020-21" Then
                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "54B"
                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = True
                Else
                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "(Select)"
                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = False
                End If
                Sheet13.Protect Password:=sPassword
            End If

            If Not Intersect(Range("STCG.A7_DeductionSection_1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                If Sheet13.Range(Replace(Target.Address, "I", "H")).value = "2016-17" Then
                  If Sheet13.Range(Target.Address).value = "54B" Then
                    fmsgboxStatus "Please select correct Section for the corresponding year"
                    Sheet13.Range(Target.Address).value = "(Select)"
                  End If
                End If
                Sheet13.Protect Password:=sPassword
            End If

 
If Not Intersect(Range("LTCG.B9_PreviousYer_1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

                If Sheet13.Range(Target.Address).value = "2020-21" Then
                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "54B"
                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = True
                Else
                    Sheet13.Range(Replace(Target.Address, "H", "I")).value = "(Select)"
                    Sheet13.Range(Replace(Target.Address, "H", "I")).Locked = False
                End If
                Sheet13.Protect Password:=sPassword
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not Intersect(Range("LTCG.B9_DeductionSection_1"), Range(Target.Address)) Is Nothing Then
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
                If Sheet13.Range(Replace(Target.Address, "I", "H")).value = "2016-17" Then
                  If Sheet13.Range(Target.Address).value = "54B" Then
                    fmsgboxStatus "Please select correct Section for the corresponding year"
                    Sheet13.Range(Target.Address).value = "(Select)"
                  End If
                End If
                Sheet13.Protect Password:=sPassword
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


'Updated by Bindu

CG_Calc.doSetoff

If (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("No") And UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("No")) Then
    CYLACalculations.CYLABFLASetOff
ElseIf (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("Yes") Or UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("Yes")) Then
    CYLA_BFLA_Editable
End If
'---

Sheets(activesh).Activate
End Sub

Sub AddTotalSTCGA1()
    Dim Row7RangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim Totsections, tot, tot1 As Variant
    Dim tot2_1, tot7_1, tot8_1, tot10_1 As Variant

Application.EnableEvents = False
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
        tot2_1 = tot2_1 + Sheet13.Range("STCG.Ad_ExemptionGrandTotal" & i).value

        'tot7_1 = tot7_1 + Sheet13.Range("STCG.Adiv_ExemptionAmount" & i).value
        'tot8_1 = tot8_1 + Sheet13.Range("STCG.Adiii_ExemptionAmount" & i).value
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
    Sheet13.Range("D_54D_Temp1").value = tot10_1


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
Sub AddTotalLTCGB1()
    Dim Row7RangeName As String
    Dim i, j As Variant
    Dim newDefinedName As String
    Dim Totsections, tot, tot1, tot2, tot3, tot4, tot5, tot6, tot7, tot8, tot9, tot10, tot11, tot12, tot13, tot14 As Variant
Application.EnableEvents = False

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
           tot2 = tot2 + Sheet13.Range("LTCG.Bdii_ExemptionAmount" & i).value
           tot3 = tot3 + Sheet13.Range("LTCG.Bdi_ExemptionAmount" & i).value
           tot4 = tot4 + Sheet13.Range("LTCG.Bdiv_ExemptionAmount" & i).value
           tot5 = tot5 + Sheet13.Range("LTCG.Bdv_ExemptionAmount" & i).value
           tot7 = tot7 + Sheet13.Range("LTCG.Bdvii_ExemptionAmount" & i).value
            
         tot9 = tot9 + Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & i).value

        Next

    Sheet13.Protect Password:=sPassword
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword

'----------------Unlock Password-------------------END-----

    Sheet13.Range("LTCG.B1e_CapgainonAssets").value = tot
       
    Sheet13.Range("D_54B_Temp").value = tot2
  
    Sheet13.Range("D_54_Temp").value = tot3
    Sheet13.Range("D_54EC_Temp").value = tot4
    Sheet13.Range("D_54F_Temp").value = tot5
    Sheet13.Range("D_54GB_Temp").value = tot6
    Sheet13.Range("D_54GA_Temp").value = tot7
    Sheet13.Range("D_54G_Temp").value = tot8
    Sheet13.Range("D_54D_Temp").value = tot10
    Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC_Temp").value = tot9

'Added by Riyaz on 24/02/2025
    For j = 1 To Totsections
          
'Commented by Riyaz on 22/01/2026 and updated on 18/02/2026
       If Sheet13.Range("LTCG.Baiii_FullConsiderationFiftyC" & j).value = "" Then
            Sheet13.Range("LTCG.Bbiv_TotalDedn" & j).value = 0
        Else
'        If Dformat(Sheet13.Range("LTCG.B_SaleTrfDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then '' updating from this LTCG.Bbii_ImproveCost to this  LTCG.Bbia_ImproveCost 'Before by Bindu
            Sheet13.Range("LTCG.Bbiv_TotalDedn" & j).value = Sheet13.Range("LTCG.Bbi_AquisitCostNew" & j).value + Sheet13.Range("LTCG.Bbia_ImproveCost_Sum" & j).value + Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & j).value
        End If
'        End If
            
'Changed by Riyaz on 23/01/2026
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
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
        
'Changed by Riyaz on 23/01/2026
'S.no ca
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
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
'Changed by Riyaz on 23/01/2026
'S.no B1e(a)
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
            If Sheet13.Range("LTCG.B_AcqDate" & j).value = "" Then
                Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).value = 0
            Else
                If Dformat(Sheet13.Range("LTCG.B_AcqDate" & j).value, "22/07/2024") <= Dformat("22/07/2024", "22/07/2024") Then
                    Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).formula = "=IF(LTCG.Bca_BalanceCG" & j & "<0" & "," & "LTCG.Bca_BalanceCG" & j & "," & "IF(LTCG.Bd_ExemptionGrandTotal" & j & ">" & "LTCG.Bca_BalanceCG" & j & "," & 0 & "," & "LTCG.Bca_BalanceCG" & j & "-" & "LTCG.Bd_ExemptionGrandTotal" & j & ")" & ")"

'                     Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).formula = "=IF(LTCG.Bca_BalanceCG" & j & "0" & "," & "LTCG.Bca_BalanceCG" & j & "," & "IF(LTCG.Bd_ExemptionGrandTotal" & j & ">" & "LTCG.Bca_BalanceCG" & j & "," & 0 & "," & "LTCG.Bca_BalanceCG" & j & "-" & "LTCG.Bd_ExemptionGrandTotal" & j & ")" & ")"
                Else
                    Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).value = 0
                End If
            End If
        Else
            Sheet13.Range("LTCG.Bea_CapgainonAssets" & j).value = 0
        End If
'Changed by Riyaz on 23/01/2026
'S.no B1ei(A)
        If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
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
'Changed by Riyaz on 23/01/2026
'S.no B1ei(B) ' Need to check with BA Bcz in DE Sheet not mentioned
    If (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "RES") Or (Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) = "NOR") Then
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
'Added by Riyaz on 26/02/2024
'Before 23rd July

'Changed by Bindu on 4th April 25for g
        If Sheet13.Range("LTCG.B_AcqDate" & j).value <> "" Then
            tot12 = tot12 + Sheet13.Range("LTCG.Be_CapgainonAssets" & j).value
        End If
'for h row
        If Sheet13.Range("LTCG.B_AcqDate" & j).value <> "" Then
            tot13 = tot13 + Sheet13.Range("LTCG.Beii_Excessamt" & j).value
        End If

           Next j
    Sheet13.Range("LTCG.B1g_ImmpropertyB1e1").value = tot12
    Sheet13.Range("LTCG.B1g_ImmpropertyB1e1c").value = tot13

'----------------Lock Password-------------------START---

   Sheet13.Protect Password:=sPassword
   
   Application.EnableEvents = True
'----------------Lock Password-------------------END-----
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub



Sub yyyyyyy()
Application.EnableEvents = True
End Sub

Sub Cost_of_Improvement()

Dim CG_B__LTCG_Count, i, j, cellcount  As Long
Dim rangecells1, rangecells2, rangecells3, rangecells4 As Range
Dim sale_Date_after As Boolean
Dim Lookup_year As Variant
Dim loopkup_1, loopkup_2

Application.EnableEvents = False
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
                If Sheet13.Range("LTCG.B_SaleTrfDate" & i).value = "" Then
                    Lookup_year = 2001
                End If
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

'by sadineni on 26/03/2026
Function ChkMinCGDate2627(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinCGDate2627 = True
    If Len(dob) > 0 Then
 
        If val(Mid(dob, 7, 4)) < 2025 Then
           ChkMinCGDate2627 = False
           Exit Function
        ElseIf (val(Mid(dob, 7, 4)) = 2025 And val(Mid(dob, 4, 2)) <= 3) Then
           ChkMinCGDate2627 = False
           Exit Function
       End If
   End If
End Function
Sub taa()
Application.EnableEvents = True
End Sub
