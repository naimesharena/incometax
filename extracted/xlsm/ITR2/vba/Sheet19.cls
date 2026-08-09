Attribute VB_Name = "Sheet19"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
calSchS

Lock_Unlock_QQB_RR
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
CalculateTaxCheckFlag = False
Application.ScreenUpdating = False
Application.EnableEvents = False

''Change-13.04.2023.103.VIA.01
'AnyOtherDeductions80CCH Target

'Change.21.03.2023.102.IDS.40A
Sheet19.Unprotect Password:=getmsgstate

'****************changed by Chetan C M for AY 2026-27 as per DE V 0.3********************
'--start--
If Sheet19.Range("scvia.Section80CCC").value > 0 Then
         Sheet19.Shapes("Button 225").Visible = True
       Else
           Sheet19.Shapes("Button 225").Visible = False
End If

If (Sheet19.Range("scvia.Section80CCD_SE").value + Sheet19.Range("scvia.Section80CCD1B_SE").value) > 0 Then
         Sheet19.Shapes("Button 229").Visible = True
       Else
           Sheet19.Shapes("Button 229").Visible = False
End If
'--end

'added by Chetan C M for AY 2026-27 on 05/05/2026 as per DE V0.3 (SIT-118779)
'start--

If Not Application.Intersect(Sheet19.Range("Type_of_Identifier_80CCC"), Sheet19.Range(Target.Address)) Is Nothing Then
    If Target.value = "PRAN" Then
        DPMAddComment_PRAN Target
    Else
        DPMADeleteComment_PRAN Target
    
    End If
End If

'--end

'If Not Application.Intersect(Sheet19.Range("scvia.Section80GG"), Sheet19.Range(Target.Address)) Is Nothing Then

   If CheckSection13A = True Then
      If (Target.value) > 55000 Then
        Sheet19.Range(Target.Address).value = ""
        MsgBox ("Deduction u/s 10(13A) & 80GG cannot be claimed for the same period ")
      End If
   End If
   
If SheetALL.Range("bacValue").value = 1 Then 'Shrutika(16/06/25)
'If Target.value > 0 Then

If Sheet19.Range("scvia.Section80GG").MergeArea.Locked = False Then

Sheet19.Unprotect Password:=getmsgstate
Sheets("VI-A").Activate
Range("Sheet19.AckNum").MergeArea.Locked = False
'Range("Sheet1.AckNum").MergeArea.ClearContents
Sheet19.Range("Sheet19.AckNum").MergeArea.Interior.Color = (&HCCFFCC)
Sheet19.Range("Sheet19.AckNum").Font.Color = vbBlack
Else

Sheet19.Unprotect Password:=getmsgstate
Range("Sheet19.AckNum").MergeArea.Locked = True
'Range("Sheet19.AckNum").MergeArea.ClearContents
Range("Sheet19.AckNum").MergeArea.Interior.Color = (&HD8D8D8)
Sheet19.Range("Sheet19.AckNum").MergeArea.Font.Color = (&HD8D8D8)
End If


End If

Sheet19.Protect Password:=getmsgstate

'End Change.IDS.40A

'Sudha_2025-26

'If Target.name.name = "scvia.Section80EE" Then
'Sheet19.Unprotect Password:=getmsgstate
'If Sheet19.Range("scvia.Section80EE").value > 0 Then
'Sheet19.Range("scvia.Section80EEA").value = ""
'Sheet19.Range("scvia.Section80EEA").Locked = True
'Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HD8D8D8)
'Sheet19.Protect Password:=getmsgstate
'Else
'Sheet19.Unprotect Password:=getmsgstate
'Sheet19.Range("scvia.Section80EEA").value = ""
'Sheet19.Range("scvia.Section80EEA").Locked = False
'Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HCCFFCC)
'Sheet19.Protect Password:=getmsgstate
'End If
'
'End If
'
'If Target.name.name = "scvia.Section80EEA" Then
'If Sheet19.Range("scvia.Section80EEA").value > 0 Then
'Sheet19.Unprotect Password:=getmsgstate
'Sheet19.Range("scvia.Section80EE").value = ""
'Sheet19.Range("scvia.Section80EE").Locked = True
'Sheet19.Range("scvia.Section80EE").Interior.Color = (&HD8D8D8)
'Sheet19.Protect Password:=getmsgstate
'Else
'Sheet19.Unprotect Password:=getmsgstate
'Sheet19.Range("scvia.Section80EE").value = ""
'Sheet19.Range("scvia.Section80EE").Locked = False
'Sheet19.Range("scvia.Section80EE").Interior.Color = (&HCCFFCC)
'Sheet19.Protect Password:=getmsgstate
'End If
'End If

'===========================

'If Not Application.Intersect(Sheet19.Range("scvia.Section80U"), Sheet19.Range(Target.Address)) Is Nothing Then
'    If Sheet19.Range("scvia.Section80U").value <> "" Or Sheet19.Range("scvia.Section80U").value <> 0 Then
'    If isdropdownblank(Sheet19.Range("SELECT80U").value) Then
'        MsgBox "Please select the appropriate drop down for 80U"
'        Target.value = ""
'    ElseIf Mid(Sheet19.Range("SELECT80U").value, 1, 1) = "1" Then
'        If Sheet19.Range("scvia.Section80U").value <> 75000 And Sheet19.Range("scvia.Section80U").value <> 0 Then
'            MsgBox " You can only claim 75000 for selected drop down"
'            Target.value = 75000
'        End If
'
'    ElseIf Mid(Sheet19.Range("SELECT80U").value, 1, 1) = "2" Then
'        If Sheet19.Range("scvia.Section80U").value <> 125000 And Sheet19.Range("scvia.Section80U").value <> 0 Then
'            MsgBox " You can only claim 125000 for selected drop down"
'            Target.value = 125000
'        End If
'    End If
'    End If
'End If

'If Not Application.Intersect(Sheet19.Range("scvia.Section80DD"), Sheet19.Range(Target.Address)) Is Nothing Then
'    If Sheet19.Range("scvia.Section80DD").value <> "" Or Sheet19.Range("scvia.Section80DD").value <> 0 Then
'    If isdropdownblank(Sheet19.Range("SELECT80DD").value) Then
'        MsgBox "Please select the appropriate drop down for 80DD"
'        Sheet19.Range("scvia.Section80DD").value = ""
'    ElseIf Mid(Sheet19.Range("SELECT80DD").value, 1, 1) = "1" Then
'        If Sheet19.Range("scvia.Section80DD").value <> 75000 And Sheet19.Range("scvia.Section80DD").value <> 0 And Target.value <> "" Then
'            MsgBox " You can only claim 75000 for selected drop down"
'            Sheet19.Range("scvia.Section80DD").value = 75000
'        End If
'
'    ElseIf Mid(Sheet19.Range("SELECT80DD").value, 1, 1) = "2" Then
'        If Sheet19.Range("scvia.Section80DD").value <> 125000 And Target.value <> 0 And Sheet19.Range("scvia.Section80DD").value <> "" Then
'            MsgBox " You can only claim 125000 for selected drop down"
'            Sheet19.Range("scvia.Section80DD").value = 125000
'        End If
'    End If
'    End If
'End If


If Target.Validation.Type = 1 Then

    If Target.name.name = "scvia.Section80C" Or Target.name.name = "scvia.Section80CCC" Or _
        Target.name.name = "scvia.Section80CCD_SE" Or Target.name.name = "scvia.Section80CCD1B_SE" Or _
        Target.name.name = "scvia.Section80CCD" Or _
        Target.name.name = "scvia.Section80DDB" Or Target.name.name = "scvia.Section80E" Or _
        Target.name.name = "scvia.Section80EE" Or Target.name.name = "scvia.Section80EEA" Or _
        Target.name.name = "scvia.Section80EEB" Or Target.name.name = "scvia.Section80GG" Then
    
     'DI_Hidden
     'DI_Greyoff
     If Target.name.name = "scvia.Section80GG" Then
        Application.EnableEvents = False
        If Target.value = "" Then
            Sheet19.Range("scvia.Section80GG").value = 0
        End If
        Application.EnableEvents = True
     End If
    
    
'    If Target.name.name = "scvia.Section80DD" Then
'    If Sheet19.Range("SELECT80DD").value = "2-Dependent person with Severe Disability" And (Target.value > 125000 Or Target.value > Sheet19.Range("GTI_LIMIT").value) Then
'    MsgBox ("Dependent person with Severe disability is Having Maximum limit of Rs. 1,25,000 or Gross Total Income whichever is lower")
'
'    Target.value = ""
'    ElseIf Sheet19.Range("SELECT80DD").value = "1-Dependent person with Disability" And (Target.value > 75000 Or Target.value > Sheet19.Range("GTI_LIMIT").value) Then
'
'    MsgBox ("Dependent person with Disability is Having Maximum limit of Rs. 75,000. or Gross Total Income whichever is lower")
'    Target.value = ""
'        End If
'    End If
'
    End If
    
    
'    If Target.name.name = "scvia.Section80U" Or Target.name.name = "SELECT80U" Then
'    If Sheet19.Range("SELECT80U").value = "2-Self with severe disability" And (Target.value > 125000 Or Target.value > Sheet19.Range("GTI_LIMIT").value) Then
'    MsgBox ("Self with severe disability is Having Maximum limit of Rs. 1,25,000 or Gross Total Income whichever is lower")
'
'    Target.value = ""
'    ElseIf Sheet19.Range("SELECT80U").value = "1-Self with disability" And (Target.value > 75000 Or Target.value > Sheet19.Range("GTI_LIMIT").value) Then
'
'    MsgBox ("Self with disability is Having Maximum limit of Rs. 75,000 or Gross Total Income whichever is lower")
'    Target.value = ""
'    ElseIf Sheet19.Range("SELECT80U").value = "(Select)" Then
'    Target.value = ""
'        End If
'
'    End If

End If


If Target.Validation.Type = 3 Then

'    If Target.name.name = "IncD.InvestmentDropDown" Then
''        DI_Greyoff
''        DI_Hidden
'    End If

End If




'Chandru
''sudha_2025-26

'If Sheet1.Range("sheet1.Status").value = "H-HUF" Or (SheetALL.Range("bacValue").value = 1) Then
'
'If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or SheetALL.Range("bacValue").value = 1 Then
'

'
'        Sheet19.Unprotect Password:=getmsgstate
'       Sheet19.Range("scvia.Section80EEA").ClearContents
'        Sheet19.Range("scvia.Section80EEA").Locked = True
'        Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HD8D8D8)
'        Sheet19.Range("scvia.Section80EEA").value = ""
'        Sheet19.Protect Password:=getmsgstate
'
'    ElseIf Sheet1.Range("sheet1.Status").value <> "H-HUF" Or (SheetALL.Range("bacValue").value <> 1) Then
'
'      If Sheet19.Range("scvia.Section80EE").value > 0 Then
'            Sheet19.Range("scvia.Section80EEA").value = ""
'            Sheet19.Range("scvia.Section80EEA").Locked = True
'            Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HD8D8D8)
'            Sheet19.Protect Password:=getmsgstate
'      Else
'
'        Sheet19.Unprotect Password:=getmsgstate
'         'Sheet19.Range("scvia.Section80EEA").ClearContents
'        Sheet19.Range("scvia.Section80EEA").Locked = False
'        Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HCCFFCC)
'        'Sheet19.Range("scvia.Section80EEA").value = ""
'        Sheet19.Protect Password:=getmsgstate
'     End If
'
'End If

'sudha_2025-26

'commented by sadineni on 31/03/2026 as per DE V0.2

'****************changed by Chetan C M for AY 2026-27 as per DE V 0.3********************
'--start--
'If Range("scvia.Section80CCC").value > 0 Or Range("scvia.Section80CCD_SE").value > 0 Or Range("scvia.Section80CCD1B_SE").value > 0 Then
If Range("scvia.Section80CCD_SE").value > 0 Or Range("scvia.Section80CCD1B_SE").value > 0 Then
    Sheet19.Unprotect Password:=getmsgstate
    'Sheet1.Range("Pran_Number").MergeArea.Locked = False
    Sheet19.Range("Pran_Number").Locked = False
    Sheet19.Range("Pran_Number").Interior.Color = (&HCCFFCC)

    
Else
    Sheet19.Unprotect Password:=getmsgstate
    Sheet19.Range("Pran_Number").Locked = True
    Sheet19.Range("Pran_Number").ClearContents
    Sheet19.Range("Pran_Number").Interior.Color = (&HD8D8D8)

End If
'----end************************************


'If Not Application.Intersect(Sheet19.Range("scvia.Section80GG"), Sheet19.Range(Target.Address)) Is Nothing Then


'End If


'--------------------------
'If Not Application.Intersect(Sheet19.Range("scvia.Section80QQB"), Sheet19.Range(Target.Address)) Is Nothing Then
If SheetALL.Range("bacValue").value = 1 Then 'Shrutika(16/06/25)

'If Target.value > 0 Then
If Sheet19.Range("scvia.Section80QQB").MergeArea.Locked = False Then

Sheet19.Unprotect Password:=getmsgstate
Sheets("VI-A").Activate
Range("Sheet19.AckNum10CCD").MergeArea.Locked = False
'Range("Sheet1.AckNum").MergeArea.ClearContents
Sheet19.Range("Sheet19.AckNum10CCD").MergeArea.Interior.Color = (&HCCFFCC)
Sheet19.Range("Sheet19.AckNum.10CCD").Font.Color = vbBlack

Else

Sheet19.Unprotect Password:=getmsgstate
Range("Sheet19.AckNum10CCD").MergeArea.Locked = True
'Range("Sheet19.AckNum10CCD").MergeArea.ClearContents  'Shrutika(16/06/25)
Range("Sheet19.AckNum10CCD").MergeArea.Interior.Color = (&HD8D8D8)
Sheet19.Range("Sheet19.AckNum10CCD").MergeArea.Font.Color = (&HD8D8D8)

End If
End If

'End If
'--------------------------
'If Not Application.Intersect(Sheet19.Range("scvia.Section80RRB"), Sheet19.Range(Target.Address)) Is Nothing Then
If SheetALL.Range("bacValue").value = 1 Then 'Shrutika(16/06/25)

'If Target.value > 0 Then
If Sheet19.Range("scvia.Section80RRB").MergeArea.Locked = False Then
Sheet19.Unprotect Password:=getmsgstate
Sheets("VI-A").Activate
Range("Sheet19.AckNum.10CCE").MergeArea.Locked = False
'Range("Sheet1.AckNum").MergeArea.ClearContents
Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Interior.Color = (&HCCFFCC)
Sheet19.Range("Sheet19.AckNum.10CCE").Font.Color = vbBlack

Else

Sheet19.Unprotect Password:=getmsgstate
Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = True
'Range("Sheet19.AckNum.10CCE").MergeArea.ClearContents 'Shrutika(16/06/25)
Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Interior.Color = (&HD8D8D8)
Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Font.Color = (&HD8D8D8)
End If


End If


If Not Application.Intersect(Sheet19.Range("SELECT80DDB"), Sheet19.Range(Target.Address)) Is Nothing Then

Sheet19.Unprotect Password:=getmsgstate
If Sheet19.Range("SELECT80DDB").MergeArea.Locked = True Then
Sheet19.Range("Sheet19.Specified_Disease").MergeArea.ClearContents
Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Interior.Color = (&HD8D8D8)
Sheet19.Range("SELECT80DDB").MergeArea.Locked = True
Sheet19.Protect Password:=getmsgstate
End If
End If

'=======================================================================

endline:
Application.ScreenUpdating = True
Application.EnableEvents = True
Sheet19.Protect Password:=getmsgstate


End Sub


Private Sub Worksheet_Deactivate()

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet19.Range(targetaddress).Value = ""
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
Function Lock_Unlock_QQB_RR()

End Function


'added by Chetan C M for AY 2026-27 on 05/05/2026 as per DE sheet V0.3 (SIT-118779)
'start********************

Sub DPMAddComment_PRAN(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet19.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet19.Range(targetadd).Comment.Delete
With Sheet19.Range(targetadd)
.AddComment
With .Comment
.text text:="PRAN (Permanent Retirement Account Number) is a unique 12-digit number allotted to every subscriber of the National Pension System (NPS)."
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet19.Protect Password:=getmsgstate
End Sub


Sub DPMADeleteComment_PRAN(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next

Sheet19.Unprotect Password:=getmsgstate
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet19.Range(targetadd).Comment.Delete
Sheet19.Protect Password:=getmsgstate
End Sub
'end**********************
