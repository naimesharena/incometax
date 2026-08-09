Attribute VB_Name = "Sheet15"
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
Sheet15.Shapes("Button 4").Top = Sheet15.Range("UD.Button4").Top + 5
Sheet15.Shapes("Button 4").Left = Sheet15.Range("UD.Button4").Left + 3
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
If (Not Intersect(Target, Range("UD.AssessmentYear")) Is Nothing) And (Target.Cells.count = 1) Then
             Dim var As Integer
             var = Application.CountIf(Range("UD.AssessmentYear"), Target.value)
             If 1 < var Then
                fmsgbox "* Same Assessment Year cannot be selected more than once."
                Range(Target.Address).value = ""
            End If
         End If
Application.EnableEvents = True
End Sub

'End Sub

Private Sub Worksheet_Deactivate()
On Error Resume Next

CalculateTaxCheckFlag = False

Application.EnableEvents = False

Dim activesh As Variant
activesh = ActiveSheet.name
'CYLACalculations.CYLABFLASetOff

'Updated by Bindu
If (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("No") And UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("No")) Then
    CYLACalculations.CYLABFLASetOff
ElseIf (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("Yes") Or UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("Yes")) Then
    CYLA_BFLA_Editable
End If
'--

Application.EnableEvents = True
Worksheets(activesh).Activate
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
