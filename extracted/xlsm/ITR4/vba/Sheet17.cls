Attribute VB_Name = "Sheet17"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False


fmsgboxStatus " Please verify the TAN details and the amount of credit available in your form 26 AS." ', vbOKOnly, "Schedule TCS"

 Sheet17.Unprotect Password:=getmsgstate

    ActiveSheet.Shapes("Button 5").Left = Range("SchTDS.Button3").Left
    ActiveSheet.Shapes("Button 5").Top = Range("SchTDS.Button3").Top

'    ActiveSheet.Shapes("Button 30").Left = Range("SchTDS3.Button2").Left
'    ActiveSheet.Shapes("Button 30").Top = Range("SchTDS3.Button2").Top
'
'    ActiveSheet.Shapes("Button 26").Left = Range("SchTDS3.Button3").Left
'    ActiveSheet.Shapes("Button 26").Top = Range("SchTDS3.Button3").Top



endline:
 Sheet17.Protect Password:=getmsgstate

Application.EnableEvents = True
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False

CalculateTaxCheckFlag = False


'Target.value = UCase(Target.value)

If Not checkfieldSuperSpecialcharacter(Target.value) Then
         fmsgbox ("* Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
End If

If Not Application.Intersect(Sheet17.Range("TCS.TAN"), Sheet17.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If



''''''''''''TCS_C1 AY_2024_25 change
'
'If Target.Column = 7 Then
'        If Not Target.value = "(Select)" Then
'        If 1 < Application.CountIf(Range("TCS.SectionTCSDeducted"), Target.value) Then
'             MsgBox "* Sch TCS-Section under which TCS is deducted in Same drop downs cannot be selected more than once in Schedule TCS"
'             Range(Target.Address).value = "(Select)"
'        End If
'        End If
'End If



endline:
Application.EnableEvents = True

End Sub
