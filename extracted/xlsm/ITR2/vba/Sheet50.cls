Attribute VB_Name = "Sheet50"
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
    ActiveSheet.Shapes("Button 1").Left = Range("Button_80GGA").Left + 5
    ActiveSheet.Shapes("Button 1").Top = Range("Button_80GGA").Top + 5

'    ThisWorkbook.Protect Password:=getmsgstate
End Sub

'Private Sub Worksheet_Deactivate()
'ThisWorkbook.Unprotect Password:=getmsgstate
'End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False
CalculateTaxCheckFlag = False

If Not IsEmpty(Trim(Target.value)) Then

    If Target.Validation.Type = 6 Then
        Target.value = UCase(Target.value)
        If Not checkfieldSuperSpecialcharacter(Target.value) Then
            fmsgboxStatus ("Please don't use Special Characters < and  > " & Chr(13))
            Target.value = ""
            GoTo endline
        End If
    End If
End If

If Target.Column = 9 Then
        Dim PinCode As Range
        For Each PinCode In SheetALL.Range("PincodeList")
            Dim code As String
            code = Left(CStr(Target.value), 3)
            If CStr(PinCode.value) = CStr(code) Then
                Sheet50.Range("H" & Target.row).value = SheetALL.Range("CX" & PinCode.row).value
                Exit For
            End If
         Next PinCode
    End If
    
    If Target.Column = 10 Then
    If Not Range("K" & Target.row).value = "" Then
        If 1 < Application.CountIf(Range("J1:J" & Rows.count), Target.value) Then
             fmsgbox "* Duplicate Pan value is not allowed"
             Range(Target.Address).value = ""
             Range("K" & Target.row).value = ""
        End If
    End If
End If
    
    If Target.Column = 11 Then
    If Not Range("K" & Target.row).value = "" Then
    Dim var As Integer
    var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).value)
        If 1 < var Then
             fmsgbox "* Duplicate Pan value is not allowed"
             Range("J" & Target.row).value = ""
             Range(Target.Address).value = ""
        End If
    End If
End If

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


