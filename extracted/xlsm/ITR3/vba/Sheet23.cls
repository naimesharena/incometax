Attribute VB_Name = "Sheet23"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
On Error Resume Next
   sPassword = EfilingCommon.getmsgstate
   Sheet23.Unprotect Password:=sPassword
Sheet23.Protect Password:=sPassword
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Application.EnableEvents = False
'On Error Resume Next

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

If Sheet23.Range("AMTC.AmtCreditSetOfEy4").value > Sheet23.Range("AMTC.AmtCreditFwd4").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy4").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy4").value = ""
    End If
If Sheet23.Range("AMTC.AmtCreditSetOfEy6").value > Sheet23.Range("AMTC.AmtCreditFwd6").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy6").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy6").value = ""
    End If

If Sheet23.Range("AMTC.AmtCreditSetOfEy7").value > Sheet23.Range("AMTC.AmtCreditFwd7").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy7").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy7").value = ""
    End If

'Malli comented

''Change-20.01.2023.101.AMTC.15A
'If Sheet23.Range("AMTC.AmtCreditSetOfEy9").value > Sheet23.Range("AMTC.AmtCreditFwd9").value Then
'        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
'        Sheet23.Range("AMTC.AmtCreditSetOfEy8").Select
'        Sheet23.Range("AMTC.AmtCreditSetOfEy8").value = ""
'    End If
'----------------------------------------
If Sheet23.Range("AMTC.AmtCreditSetOfEy8").value > Sheet23.Range("AMTC.AmtCreditFwd8").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy8").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy8").value = ""
    End If
'---

    If Sheet23.Range("AMTC.AmtCreditSetOfEy1").value > Sheet23.Range("AMTC.AmtCreditFwd1").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy1").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy1").value = ""
    End If

    If Sheet23.Range("AMTC.AmtCreditSetOfEy2").value > Sheet23.Range("AMTC.AmtCreditFwd2").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy2").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy2").value = ""
    End If

    If Sheet23.Range("AMTC.AmtCreditSetOfEy3").value > Sheet23.Range("AMTC.AmtCreditFwd3").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy3").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy3").value = ""
    End If

  If Sheet23.Range("AMTC.AmtCreditSetOfEy").value > Sheet23.Range("AMTC.AmtCreditFwd").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy").value = ""
    End If
    
    'AY_2024_25 Change 'Malli
    'AMTC.AmtCreditSetOfEy9  'AMTC.AmtCreditFwd9
    
    If Sheet23.Range("AMTC.AmtCreditSetOfEy9").value > Sheet23.Range("AMTC.AmtCreditFwd9").value Then
        fmsgboxStatus "Amount credited to B(2) cannot be greater than B(1)"
        Sheet23.Range("AMTC.AmtCreditSetOfEy9").Select
        Sheet23.Range("AMTC.AmtCreditSetOfEy9").value = ""
    End If
    '-----------------------------------------------------------
    
Application.EnableEvents = True
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)

CalculateTaxCheckFlag = False


End Sub


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub
Sub pppp()
Application.EnableEvents = True
End Sub
