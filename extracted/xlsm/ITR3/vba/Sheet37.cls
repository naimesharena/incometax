Attribute VB_Name = "Sheet37"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet37.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet37.Shapes("Button 1").Left = Range("Button_80IAa").Left + 275
Sheet37.Shapes("Button 1").Top = Range("Button_80IAa").Top + 16

Sheet37.Shapes("Button 47").Left = Range("Button_80IAb").Left + 275
Sheet37.Shapes("Button 47").Top = Range("Button_80IAb").Top + 16

Sheet37.Shapes("Button 49").Left = Range("Button_80IAc").Left + 275
Sheet37.Shapes("Button 49").Top = Range("Button_80IAc").Top + 15

Sheet37.Shapes("Button 50").Left = Range("Button_80IAd").Left + 275
Sheet37.Shapes("Button 50").Top = Range("Button_80IAd").Top + 40

Sheet37.Shapes("Button 51").Left = Range("Button_80IBa").Left + 275
Sheet37.Shapes("Button 51").Top = Range("Button_80IBa").Top + 16

Sheet37.Shapes("Button 30").Left = Range("Button_IBb").Left + 275
Sheet37.Shapes("Button 30").Top = Range("Button_IBb").Top + 17
Sheet37.Shapes("Button 30").Visible = False


Sheet37.Shapes("Button 31").Left = Range("Button_IBc").Left + 275
Sheet37.Shapes("Button 31").Top = Range("Button_IBc").Top + 16
Sheet37.Shapes("Button 31").Visible = False


Sheet37.Shapes("Button 32").Left = Range("Button_IBd").Left + 275
Sheet37.Shapes("Button 32").Top = Range("Button_IBd").Top + 10

Sheet37.Shapes("Button 33").Left = Range("Button_IBe").Left + 275
Sheet37.Shapes("Button 33").Top = Range("Button_IBe").Top + 10

Sheet37.Shapes("Button 34").Left = Range("Button_IBf").Left + 275
Sheet37.Shapes("Button 34").Top = Range("Button_IBf").Top + 16

Sheet37.Shapes("Button 35").Left = Range("Button_IBg").Left + 275
Sheet37.Shapes("Button 35").Top = Range("Button_IBg").Top + 15


Sheet37.Shapes("Button 36").Left = Range("Button_IBh").Left + 275
Sheet37.Shapes("Button 36").Top = Range("Button_IBh").Top + 14
Sheet37.Shapes("Button 36").Visible = False


Sheet37.Shapes("Button 37").Left = Range("BUtton_IBi").Left + 275
'Sheet37.Shapes("Button 37").Top = Range("BUtton_IBi").Top + 20
Sheet37.Shapes("Button 37").Top = Range("BUtton_IBi").Top + 30

Sheet37.Shapes("Button 38").Left = Range("Button_IBj").Left + 275
Sheet37.Shapes("Button 38").Top = Range("Button_IBj").Top + 20

Sheet37.Shapes("Button 39").Left = Range("Button_IBk").Left + 275
Sheet37.Shapes("Button 39").Top = Range("Button_IBk").Top + 13

Sheet37.Shapes("Button 40").Left = Range("Button_IBl").Left + 275
Sheet37.Shapes("Button 40").Top = Range("Button_IBl").Top + 20

Sheet37.Shapes("Button 42").Left = Range("Button_ICa").Left + 275
Sheet37.Shapes("Button 42").Top = Range("Button_ICa").Top + 8

Sheet37.Shapes("Button 43").Left = Range("Button_ICb").Left + 275
Sheet37.Shapes("Button 43").Top = Range("Button_ICb").Top + 4

Sheet37.Shapes("Button 44").Left = Range("Button_ICc").Left + 275
Sheet37.Shapes("Button 44").Top = Range("Button_ICc").Top + 16

Sheet37.Shapes("Button 21").Left = Range("Button_ICca").Left + 212
Sheet37.Shapes("Button 21").Top = Range("Button_ICca").Top

Sheet37.Shapes("Button 22").Left = Range("Button_ICcb").Left + 212
Sheet37.Shapes("Button 22").Top = Range("Button_ICcb").Top + 1

Sheet37.Shapes("Button 23").Left = Range("Button_ICcc").Left + 212
Sheet37.Shapes("Button 23").Top = Range("Button_ICcc").Top + 1

Sheet37.Shapes("Button 24").Left = Range("Button_ICcd").Left + 212
Sheet37.Shapes("Button 24").Top = Range("Button_ICcd").Top + 2

Sheet37.Shapes("Button 25").Left = Range("Button_ICce").Left + 212
Sheet37.Shapes("Button 25").Top = Range("Button_ICce").Top + 2

Sheet37.Shapes("Button 26").Left = Range("Button_ICcf").Left + 212
Sheet37.Shapes("Button 26").Top = Range("Button_ICcf").Top + 2

Sheet37.Shapes("Button 27").Left = Range("Button_ICcg").Left + 212
Sheet37.Shapes("Button 27").Top = Range("Button_ICcg").Top + 2

Sheet37.Shapes("Button 28").Left = Range("Button_ICch").Left + 212
Sheet37.Shapes("Button 28").Top = Range("Button_ICch").Top + 2

Sheet37.Protect Password:=sPassword

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

CalculateTaxCheckFlag = False


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

