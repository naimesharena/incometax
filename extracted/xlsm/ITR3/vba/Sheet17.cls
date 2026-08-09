Attribute VB_Name = "Sheet17"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


Private Sub Worksheet_Activate()
mdCFL.addCFLformula
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo endline
'Sheet17.Unprotect Password:=getmsgstate
CalculateTaxCheckFlag = False

If (Target.name.name = "yr2007.HPLossCF10") Then
          If Not (chkMandatory(Target.value)) Then
                Sheet17.Range("yr2007.HPLossCF10").formula = "=sheet16.BalHPlossCurYrAftSetoff"
          End If
    End If

    If (Target.name.name = "yr2007.BusLossOthThanSpecLossCF10") Then
          If Not (chkMandatory(Target.value)) Then
                Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").formula = "=sheet16.BalBusLossAftSetoff"
          End If
    End If

    If (Target.name.name = "yr2007.LossFrmSpecBusCF10") Then
          If Not (chkMandatory(Target.value)) Then
                Sheet17.Range("yr2007.LossFrmSpecBusCF10").formula = "=ABS(MIN(sheet12.AdjustedPLFrmSpecuBus,0))"
          End If
    End If

    If (Target.name.name = "yr2007.LossFrmSpecifiedBusCF10") Then
          If Not (chkMandatory(Target.value)) Then
                Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").formula = "=ABS(MIN(sheet12.AdjustedPLFrmSpecifiedBus,0))"
          End If
    End If

    If (Target.name.name = "yr2007.STCGLossCF10") Then
          If Not (chkMandatory(Target.value)) Then
                Sheet17.Range("yr2007.STCGLossCF10").formula = "=IHLA.Eviii2_StclSetoff15Per+IHLA.Exii2_StclSetoff20Per+IHLA.Eviii3_StclSetoff30Per+IHLA.Eviii4_StclSetoffAppRate+IHLA.Eviii5_StclSetoffDTAARate"
          End If
    End If

    If (Target.name.name = "yr2007.LTCGLossCF10") Then
          If Not (chkMandatory(Target.value)) Then
                'Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"
                 Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Exii8_LtclSetOff12.5Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"
          End If
    End If

    If (Target.name.name = "yr2007.OthSrcLossRaceHorseCF10") Then
          If Not (chkMandatory(Target.value)) Then
                Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").formula = "=ABS(MIN(os.BalanceOwnRaceHorse,0))"
          End If
    End If
endline:
'Sheet17.Protect Password:=getmsgstate
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

