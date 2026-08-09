Attribute VB_Name = "Sheet3"
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

 Sheet3.Unprotect Password:=getmsgstate

    ActiveSheet.Shapes("Button 94").Left = Range("SCHBP_GST").Left + 6
    ActiveSheet.Shapes("Button 94").Top = Range("SCHBP_GST").Top + 6

'    ActiveSheet.Shapes("Button 123").Left = Range("SCHBP_NOB").Left + 6
'    ActiveSheet.Shapes("Button 123").Top = Range("SCHBP_NOB").Top + 5
'
'    ActiveSheet.Shapes("Button 127").Left = Range("SCHBP_NOB2").Left + 6
'    ActiveSheet.Shapes("Button 127").Top = Range("SCHBP_NOB2").Top + 5
'
'    ActiveSheet.Shapes("Button 129").Left = Range("SCHBP_NOB3").Left + 6
'    ActiveSheet.Shapes("Button 129").Top = Range("SCHBP_NOB3").Top + 5



endline:
 Sheet3.Protect Password:=getmsgstate

Application.EnableEvents = True
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False

'BP_NEW_E2_P1 2024-25 Bindu


If Not Application.Intersect(Range("BP_E2a"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Range("BP_E2b"), Range(Target.Address)) Is Nothing Then

   If Range("BP_E2_PI").value > 30000001 Then
      fmsgbox "*Presumptive income under section 44AD cannot exceed 30000000"
      Range(Target.Address).value = ""
   
   End If

End If

'BP_NEW_E1_P1 2024-25 Bindu
If Not Application.Intersect(Range("BP_E1a"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Range("BP_E1a_E1b"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Range("BP_E1b"), Range(Target.Address)) Is Nothing Then

   If Range("BP_E1_GR").value > 30000001 Then
      fmsgbox "*If Gross Receipts are more than Rs.3 Crore it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form."
      Range(Target.Address).value = ""
   End If

End If

'New change for datavalidation by Sai

   If Range("BP_E1a").value > 30000000 Or Range("BP_E1a_E1b").value > 30000000 Or Range("BP_E1b").value > 30000000 Then
      fmsgbox "*If Gross Receipts are more than Rs.3 Crore it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form."
      Range(Target.Address).value = ""
   End If

'New change for datavalidation by Sai

   If Range("BP_E3_a").value > 7500000 Or Range("BP_E3_b").value > 7500000 Or Range("BP_E3_c").value > 7500000 Then
      fmsgbox "*If Gross Receipts are more than Rs 7500000, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form.."
      Range(Target.Address).value = ""
   End If





If Not Intersect(Target, Range("Sec44AE.RegNo")) Is Nothing Then
             Dim var As Integer
             var = Application.CountIf(Range("Sec44AE.RegNo"), Target.value)
             If 1 < var Then
             'Ayush_Legacy24
'                fmsgbox "* Duplicate Registration No. of goods carriage is not allowed"
                fmsgbox "* ""Same registration no. of goods carriage cannot be entered more than one time."""
                Range(Target.Address).value = ""
            End If
         End If
         
If Target.Validation.Type = 3 Then


'If Not Application.Intersect(Sheet3.Range("BP.NatureItem"), Sheet3.Range(Target.Address)) Is Nothing Then
'
'
'        If Sheet3.Range(Target.Address).value = "Any Other" Then
'            sPassword = getmsgstate
'            Sheet3.Unprotect Password:=sPassword
'            Sheet3.Range(Replace(Target.Address, "E", "F")).value = ""
'            Sheet3.Range(Replace(Target.Address, "E", "F")).Interior.Color = (&HCCFFCC)  ', "AU"
'            Sheet3.Range(Replace(Target.Address, "E", "F")).Locked = False
'
'            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
'
'            Sheet3.Protect Password:=sPassword
'        Else
'            sPassword = getmsgstate
'            Sheet3.Unprotect Password:=sPassword
'            Sheet3.Range(Replace(Target.Address, "E", "F")).value = "NA"
'            Sheet3.Range(Replace(Target.Address, "E", "F")).Interior.Color = (&HD8D8D8)
'            Sheet3.Range(Replace(Target.Address, "E", "F")).Locked = True
'            Sheet3.Protect Password:=sPassword
'       End If
'       End If
       End If


CalculateTaxCheckFlag = False


If Not checkfieldSuperSpecialcharacter(Target.value) Then
         fmsgbox ("Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
End If

If Target.Validation.Type <> 3 Then
      '  Target.value = UCase(Target.value)
End If

If Not Intersect(Range("BP.GSTRno"), Range(Target.Address)) Is Nothing Then
                Dim TargetGSTRNo As String
                TargetGSTRNo = ""
                TargetGSTRNo = Target.Address
                Range(TargetGSTRNo).value = UCase(Range(TargetGSTRNo).value)
'Change-06.05.2023.101 -> Prod Issue
'                If Len((UCase((Range(TargetGSTRNo))))) = 15 Then
                If Not mdNOBBP.CheckGSTR(UCase((Range(TargetGSTRNo).value))) Then
                    fmsgbox "* Invalid GSTIN.GSTIN format should be First 2 numeric, next 5 Alphabets, next 4 numeric , next 1 Alphabet, then next 3 alphanumeric." ', vbOKCancel, mdCommon.vbMessageCaption
                    Target.value = ""
'                    Sheet3.Range("BP.GSTRno").Select
                    'Sheet3.Range("BP.GSTRno").ClearContents
                    'Sheet3.Range("BP.GSTRno").value = " " 'Sheet3.Range("BP.GSTRno").ClearContents
                    GoTo endline
                End If
'                Else
'                MsgBox "GSTR number entered in this field should be 15 character", vbOKCancel, mdCommon.vbMessageCaption
'                     Sheet3.Range(TargetGSTRNo).Select
'                    Sheet3.Range(TargetGSTRNo).value = ""
'                    GoTo endline
'                End If
'----

End If


If Not Intersect(Range("Sheet44AE.PresumptiveIncome"), Range(Target.Address)) Is Nothing Then
Dim TargetPI As String
TargetPI = ""
TargetPI = Target.Address

  If Target.Offset(0, -2).value > 12 Then
      If Range(TargetPI).value < Target.Offset(0, 5).value Then
            fmsgbox "* Income cannot be lower than Rs. 1000 per metric ton per month in case of Heavy Goods vehicle of 44AE section in Sheet BP " ', vbOKCancel, mdCommon.vbMessageCaption
            Range(TargetPI).value = ""
            Range(TargetPI).Select
      End If
   End If

   If Target.Offset(0, -2).value <= 12 Then
      If Range(TargetPI).value < Target.Offset(0, 6).value Then
            fmsgbox "* Income cannot be less than Rs. 7500 per month in case of other than Heavy Goods Vehicle of 44AE section in Sheet BP" ', vbOKCancel, mdCommon.vbMessageCaption
            Range(TargetPI).value = ""
            Range(TargetPI).Select
      End If
   End If

End If

If Target.Validation.Type = 2 Then
If Not Application.Intersect(Sheet3.Range("Sec44AE.TonnageCpct"), Sheet3.Range(Target.Address)) Is Nothing Then
If Sheet3.Range(Target.Address).value > 100 Then
fmsgboxsmall "*Please enter valid tonnage"
'Sheet3.Range(Target.Address).value = ""
End If
End If
End If

endline:
Application.EnableEvents = True

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub
