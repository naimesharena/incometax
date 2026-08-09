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
       
            fmsgboxStatus "Please enter a valid PAN" ', vbOKOnly, "ITR-3"
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
