Attribute VB_Name = "Sheet41"
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
Sheet24.Unprotect Password:=sPassword
Sheet24.Range("SCHEI_Pass_Through_Income").value = MyTotalPTI()
Sheet24.Protect Password:=sPassword
End Sub
Private Sub Worksheet_Activate()
On Error Resume Next


    ActiveSheet.Shapes("Button 8").Left = Range("PTI.Button2").Left + 2
    ActiveSheet.Shapes("Button 8").Top = Range("PTI.Button2").Top + 2

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False


Application.EnableEvents = False



Target.value = UCase(Target.value)
AddTotPTIHP
AddTotIncomeUnderHouseProperty

'Malli---------------
'SIT-72538
sPassword = getmsgstate
Sheet24.Unprotect Password:=sPassword
Sheet24.Range("SCHEI_Pass_Through_Income").value = MyTotalPTI()
Sheet24.Protect Password:=sPassword
'-----------------------------------

endline:
Application.EnableEvents = True
End Sub
Sub AddTotPTIHP()
    Dim TotIncomeRangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, Totsections As Long

    tot = 0
    TotIncomeRangeName = "PTI_HP_NetAmt"
    Totsections = Sheet41.Range("cntr.Ptirptfrm").value
        For i = 1 To Totsections
        newDefinedName = TotIncomeRangeName & i
        tot = tot + Sheet41.Range(newDefinedName)
        Next
        
        sPassword = getmsgstate
        Sheet8.Unprotect Password:=sPassword
        
        Sheet8.Range("HP.PassTroughIncome").value = tot '+ Sheet8.Range("HP.PassTroughIncome").value '+ Sheet8.Range("HP.RentArearsSec25BAfter30pcDeduct").value
        
        Sheet8.Protect Password:=sPassword
End Sub
Sub AddTotIncomeUnderHouseProperty()
    Dim TotIncomeRangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, Totsections As Long

    tot = 0
    TotIncomeRangeName = "HP.IncomeOfHP"
    Totsections = Sheet8.Range("PropertySectionCOunt").value
        For i = 1 To Totsections
        newDefinedName = TotIncomeRangeName & i
        tot = tot + Sheet8.Range(newDefinedName)
        Next
        
        sPassword = getmsgstate
        Sheet8.Unprotect Password:=sPassword
        
        
    Sheet8.Range("HP.TotalIncomeChargeableUnHP").value = tot + Sheet8.Range("HP.PassTroughIncome").value '+ Sheet8.Range("HP.RentArearsSec25BAfter30pcDeduct").value
    
    Sheet8.Protect Password:=sPassword
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

