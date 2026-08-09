Attribute VB_Name = "Sheet51"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Activate()
On Error Resume Next
    ActiveSheet.Shapes("Button 1").Left = Range("Button_GST").Left + 2
    ActiveSheet.Shapes("Button 1").Top = Range("Button_GST").Top + 2

'    ThisWorkbook.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo endline

Application.EnableEvents = False
CalculateTaxCheckFlag = False
'Target.value = UCase(Target.value)
'If Not Intersect(Sheet51.Range("GST_GSTRno"), Range(Target.address)) Is Nothing Then
'                Dim TargetGSTRNo As String
'                TargetGSTRNo = ""
'                TargetGSTRNo = Target.address
'                Range(TargetGSTRNo).value = UCase(Range(TargetGSTRNo).value)
'                If Len((UCase(Trim(Range(TargetGSTRNo))))) = 15 Then
'                If Not GST.CheckGSTR(UCase(Trim(Range(TargetGSTRNo).value))) Then
'                    fmsgboxStatus "Invalid GSTIN.GSTIN format should be First 2 numeric, next 5 Alphabets, next 4 numeric , next 1 Alphabet, then next 3 alphanumeric.in schedule GST" ', vbOKCancel, mdCommon.vbMessageCaption
'                    Range("TargetGSTRNo").Select
'                    Range("TargetGSTRNo").value = ""
'                    'Sheet3.Range("GST.GSTRno").value = " " 'Sheet3.Range("GST.GSTRno").ClearContents
'                    'GoTo endline
'                End If
'
'                End If
'
'
'End If

endline:
Application.EnableEvents = True
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

