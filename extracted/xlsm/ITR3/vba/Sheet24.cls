Attribute VB_Name = "Sheet24"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
Sheet24.Shapes("Button 10").Left = Range("ScheduleEI_Button1").Left + 3
Sheet24.Shapes("Button 10").Top = Range("ScheduleEI_Button1").Top + 3

Sheet24.Shapes("Button 12").Left = Range("ScheduleEI_Button2").Left + 3
Sheet24.Shapes("Button 12").Top = Range("ScheduleEI_Button2").Top + 3

Sheet24.Shapes("Button 18").Left = Range("ScheduleEI_Button3").Left + 3
Sheet24.Shapes("Button 18").Top = Range("ScheduleEI_Button3").Top + 3

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False
CalculateTaxCheckFlag = False

sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword
         
         'Commented by Aavula Naresh for AY 2026-27 as per 0.2V=========
         
'If Not Intersect(Range("EI.NatureOfIncome"), Range(target.Address)) Is Nothing Then
'    Dim ChangeAddress As String
'         ChangeAddress = Replace(target.Address, "F", "G")
'    If Mid(Sheet24.Range(target.Address).value, 1, 3) = "Any" Then
'        ' If Mid(Sheet24.Range(ChangeAddress).value, 1, 3) <> "Div" Then  'Malli
'         Sheet24.Range(ChangeAddress).Locked = False
'         Sheet24.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
'         Sheet24.Range(ChangeAddress).ClearContents
'        ' End If
'
'    Else
'    'If Mid(Sheet24.Range(ChangeAddress).value, 1, 3) <> "Div" Then  'Malli
'         Sheet24.Range(ChangeAddress).Locked = True
'         Sheet24.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
'         Sheet24.Range(ChangeAddress).ClearContents
'         'End If
'    End If
'End If

'========================================================

If Not Application.Intersect(Sheet24.Range("Nature_Others"), Sheet24.Range(Target.Address)) Is Nothing Then
    If Target.value = "10(32)-Minor child’s income—small exemption" Then
        DPMAddComment45 Target
    ElseIf Target.value = "10(26)-Any income as referred to in section 10(26)" Then
        DPMAddComment46 Target
    ElseIf Target.value = "10(26AAA)-Any income as referred to in section 10(26AAA)" Then
        DPMAddComment47 Target
    ElseIf Target.value = "10(12A)-Any payment from the National Pension System Trust to an assessee" Then
        DPMAddComment48 Target
    ElseIf Target.value = "10(12AA)-any payment from the National Pension System Trust " Then
        DPMAddComment49 Target
    ElseIf Target.value = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee" Then
        DPMAddComment50 Target
    ElseIf Target.value = "10(12BA)-partial withdrawal made from the National Pension System" Then
        DPMAddComment51 Target
    ElseIf Target.value = "10(12C)-Agniveer Corpus Fund income" Then
        DPMAddComment52 Target
    Else
        DPMADeleteComment45 Target
    End If
End If






'                Dim ChangeAddress3 As String
'                ChangeAddress3 = Replace(Target.Address, "F", "H")
'                If Mid(Sheet24.Range(Target.Address).value, 1, 3) <> "Any" Then
'                        If Mid(Sheet24.Range(ChangeAddress).value, 1, 3) = "Div" Then
'                        Sheet24.Range(ChangeAddress3).Locked = True
'                        Sheet24.Range(ChangeAddress3).Interior.Color = "&HD8D8D8"
'                        Sheet24.Range(ChangeAddress3).ClearContents
'                        End If
'                        Else
'                        Sheet24.Range(ChangeAddress3).Locked = False
'                        Sheet24.Range(ChangeAddress3).Interior.Color = "&HCCFFCC"
'                        Sheet24.Range(ChangeAddress3).ClearContents
'                End If
'End If
'If Not Intersect(Range("EI.NatureOfIncome_1"), Range(Target.Address)) Is Nothing Then
'                Dim ChangeAddress3, ChangeAddress4 As String
'                ChangeAddress3 = Replace(Target.Address, "F", "H")
'                ChangeAddress4 = Replace(Target.Address, "F", "G")
'                If Mid(Sheet24.Range(Target.Address).value, 1, 3) <> "Any" Then
'                        If Mid(Sheet24.Range(ChangeAddress4).value, 1, 3) = "Div" Then
'                        Sheet24.Range(ChangeAddress3).Locked = True
'                        Sheet24.Range(ChangeAddress3).Interior.Color = "&HD8D8D8"
'                        Sheet24.Range(ChangeAddress3).ClearContents
'                        End If
'                        Else
'                        Sheet24.Range(ChangeAddress3).Locked = False
'                        Sheet24.Range(ChangeAddress3).Interior.Color = "&HCCFFCC"
'                        Sheet24.Range(ChangeAddress3).ClearContents
'                End If
'End If

'Updated by Malli 28/03/2024------------------------------
'If Not Intersect(Range("EI.Amount_1"), Range(Target.Address)) Is Nothing Then
'    Dim ChangeAddress1, ChangeAddress2 As String
'         ChangeAddress1 = Replace(Target.Address, "H", "G")
'         ChangeAddress2 = Replace(Target.Address, "H", "F")
'         If Mid(Sheet24.Range(ChangeAddress1).value, 1, 3) = "Div" And Mid(Sheet24.Range(ChangeAddress2).value, 1, 3) = "Any" Then   'Malli
'             Dim EI_DIV_TOTal As Variant
'             EI_DIV_TOTal = Target.value
'             If EI_DIV_TOTal <> "" Then
'                  If EI_DIV_TOTal < Sheet9.Range("DivIncome_1").value Then
'  'fmsgboxStatus "*""The amount of Dividend income mentioned in this field cannot be less that the amount of Dividend income mentioned in the Sl. No. 5c of Schedule BP."""
'     MsgBox """The amount of Dividend income mentioned in this field cannot be less that the amount of Dividend income mentioned in the Sl.No.5c(1) of Schedule BP.""", , "Error(s)"
'                  Target.value = "0"
'                  End If
'             End If
'         End If
'End If
 
'commented by Chetan C M on 09/02/2026 for AY 2026-27
'start--

''Added by Aavula Naresh AY 2025-26
''Modified by sai on 26/06/2025 as Per Shubham Verma(BA)


If Not Intersect(Sheet24.Range("SCHEI_DTAA_AmountOfIncome"), Sheet24.Range(Target.Address)) Is Nothing Then
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        If Trim(Target.value <> "") Then
             fmsgbox """Kindly re-check the DTAA benefits claimed by you in the ITR. Same may not be allowed for residents who may claim DTAA benefit under Schedule TR and FSI.""" & Chr(10) & Chr(10) ', vbOKOnly, "Alert"
        End If
    End If
End If

'--end

PartAGenChangeSub4 Target


'''''''=========================================================================
'Added by Aavula Naresh on 25-06-2026

If Not Application.Intersect(Sheet24.Range("Nature_Others"), Sheet24.Range(Target.Address)) Is Nothing Then

        If (Sheet24.Range(Target.Address).value = "Income exempt as per CBDT Circular") Or (Sheet24.Range(Target.Address).value = "Income exempt as per CBDT Notification") Or (Sheet24.Range(Target.Address).value = "Receipts not in the nature of income") Then

            Sheet24.Unprotect Password:=getmsgstate
            Sheet24.Range(Replace(Target.Address, "G", "H")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet24.Range(Replace(Target.Address, "G", "H")).Locked = False
            Sheet24.Range(Replace(Target.Address, "G", "H")).value = ""
            Sheet24.Protect Password:=getmsgstate
        Else
            Sheet24.Unprotect Password:=getmsgstate
            
            Sheet24.Range(Replace(Target.Address, "G", "H")).Interior.Color = (&HD8D8D8)
            Sheet24.Range(Replace(Target.Address, "G", "H")).Locked = True
            Sheet24.Range(Replace(Target.Address, "G", "H")).value = ""
            Sheet24.Protect Password:=getmsgstate
   End If
            Sheet24.Unprotect Password:=getmsgstate
            Sheet24.Range(Replace(Target.Address, "G", "F")).value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.VLookup(Target.value, SheetALL.Range("Sub_Category_lookup"), 2, 0), "")
            Sheet24.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet24.Range("Others.NOI"), Sheet24.Range(Target.Address)) Is Nothing Then
If Target.value = "(Select)" Then
        Sheet24.Unprotect Password:=getmsgstate
        Sheet24.Range(Replace(Target.Address, "F", "H")).Interior.Color = "&HD8D8D8"
        Sheet24.Range(Replace(Target.Address, "F", "H")).Locked = True
        Sheet24.Range(Replace(Target.Address, "F", "H")).value = ""
        Sheet24.Protect Password:=getmsgstate
End If
End If

'''''''=========================================================================

   Sheet24.Protect Password:=sPassword

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

Sub testt()
Application.EnableEvents = True

End Sub
Sub DPMAddComment45(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption is allowed to the extent such income does not exceed one thousand five hundred rupees in respect of each minor child whose income is so includible;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment46(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26) is available only to certain categories of tax payers in NER and Ladakh"""
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment47(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26AAA) is available only to certain categories of Sikkimese tax payers"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment48(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed sixty per cent of the total amount payable  at the time of such closure or opting out of the scheme"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment49(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent that it does not exceed sixty per cent. of the individual corpus"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment50(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment51(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub

Sub DPMAddComment52(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
With Sheet24.Range(targetadd)
.AddComment
With .Comment
.text text:="""Exemption u/s 10(12C) can only be claimed by nominees of Agniveer or Disabled Agniveer under the Agnipath Scheme"""
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet24.Protect Password:=sPassword
End Sub



Sub DPMADeleteComment45(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next
sPassword = EfilingCommon.getmsgstate
         Sheet24.Unprotect Password:=sPassword
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet24.Range(targetadd).Comment.Delete
Sheet24.Protect Password:=sPassword
End Sub
