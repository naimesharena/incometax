Attribute VB_Name = "Sheet9"
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
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet9.Shapes("Button 17").Left = Range("BP.OtherExempt").Left
Sheet9.Shapes("Button 17").Top = Range("BP.OtherExempt").Top + 3

Sheet9.Shapes("Button 11").Left = Range("dropdown35ad").Left
Sheet9.Shapes("Button 11").Top = Range("dropdown35ad").Top + 3

'        Sheet9.Range("EI.Drpdn").Locked = True
'        Sheet9.Range("EI.Drpdn").Interior.Color = "&HD8D8D8"

Sheet9.Protect Password:=sPassword

End Sub
Private Sub Worksheet_Change(ByVal Target As Range)

Application.EnableEvents = False

On Error GoTo endline




'Updated by Malli on 28/03/2024--------------------

If Not Application.Intersect(Range("sheet10.IncRecCredPLOthHeads4Div"), Range(Target.Address)) Is Nothing Or Not Application.Intersect(Range("DivIncome_1"), Range(Target.Address)) Is Nothing Then

Dim dividends_tot As Variant

dividends_tot = Sheet9.Range("sheet10.IncRecCredPLOthHeads4Div").value + Sheet9.Range("DivIncome_1").value

If dividends_tot > Sheet4.Range("PL.Dividends").value Then



If (Target.name.name) = "sheet10.IncRecCredPLOthHeads4Div" Then

'MsgBox """Amount mentioned in this field plus A.5c(i) - Dividend Income"""" cannot be more than the amount of Dividend income mentioned in the Sl. No. 14(iii) of Part A Profit & loss account""", , "Error(s)"

MsgBox """Amount mentioned in this field plus A.5c(1) - Dividend Income"""" cannot be more than the amount of Dividend income mentioned in the Sl. No. 14(iii) of Part A Profit & loss account""", , "Error(s)"

ElseIf (Target.name.name) = "DivIncome_1" Then

MsgBox """Amount mentioned in this field plus A.3d(i) - Dividend Income"""" cannot be more than the amount of Dividend income mentioned in the Sl. No. 14(iii) of Part A Profit & loss account""", , "Error(s)"

End If
 
End If

End If



If Target.Column = 7 Then
        If Not (Target.value = "(Select)" Or Target.value = "") Then
        If 1 < Application.CountIf(Range("EI.Drpdn"), Target.value) Then
             fmsgbox "* Same drop downs cannot be selected more than once in Schedule BP"
             Range(Target.Address).value = "(Select)"
        End If
        End If
End If

CalculateTaxCheckFlag = False

'If (Target.name.name) = "sheet10.NetPLFromSpecBus" Then
'
'If Sheet9.Range("sheet10.NetPLFromSpecBus").value = "" Then
'
''Sheet9.Range("sheet10.NetPLFromSpecBus").formula = "=PL_65iv"
'Sheet9.Range("sheet10.NetPLFromSpecBus").formula = "=IncomeIntraTrading_12b+PL_65iv"
'End If
'
'End If



'If (Target.name.name) = "sheet10.IncRecCredPLOthHeads4Div" Then

'    If Not Sheet4.Range("PL.Dividends").value = "" Then

'        If Target.value > Sheet4.Range("PL.Dividends").value Then

'            fmsgbox "* Divident Income can not be higher then Profit and loss divident income"

'            Range(Target.Address).value = "0"

'        End If

'    End If

'End If



If Target.name.name = "sheet12.DeductUs35AD" Then

    If Sheet9.Range("sheet12.DeductUs35AD").value > 0 Then

    

        sPassword = EfilingCommon.getmsgstate

        Sheet9.Unprotect Password:=sPassword

        Sheet9.Range("EI.Drpdn").value = "(Select)"

        Sheet9.Range("EI.Drpdn").Locked = False

        Sheet9.Range("EI.Drpdn").Interior.Color = "&HCCFFCC"

        Sheet9.Protect Password:=sPassword

    

    Else

        

        sPassword = EfilingCommon.getmsgstate

        Sheet9.Unprotect Password:=sPassword

        Sheet9.Range("EI.Drpdn").ClearContents

        Sheet9.Range("EI.Drpdn").Locked = True

        Sheet9.Range("EI.Drpdn").Interior.Color = "&HD8D8D8"

        Sheet9.Protect Password:=sPassword

        

    End If

End If



If Target.name.name = "sheet10.NetPLFromSpecifiedBus" Then
    If Sheet9.Range("sheet12.NetPLFrmSpecifiedBus").value > 0 And SheetALL.Range("bacValue").value = 2 Then
       sPassword = EfilingCommon.getmsgstate
       Sheet9.Unprotect Password:=sPassword
                    Sheet9.Range("sheet12.DeductUs35AD").Locked = False
                    Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = "&HCCFFCC"
                    Sheet9.Protect Password:=sPassword
     Else
       sPassword = EfilingCommon.getmsgstate
       Sheet9.Unprotect Password:=sPassword
            If SheetALL.Range("bacValue").value = 2 Or Sheet9.Range("sheet12.NetPLFrmSpecifiedBus").value = 0 Then
                    Sheet9.Range("sheet12.DeductUs35AD").ClearContents
                    Sheet9.Range("sheet12.DeductUs35AD").Locked = True
                    Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = "&HD8D8D8"
                    Sheet9.Protect Password:=sPassword
            End If
    End If
End If



endline:
Sheet9.Protect Password:=sPassword

Application.EnableEvents = True

End Sub
Private Sub Worksheet_Changeold2(ByVal Target As Range)
Application.EnableEvents = False
On Error Resume Next

CalculateTaxCheckFlag = False
If (Target.name.name) = "sheet10.NetPLFromSpecBus" Then
If Sheet9.Range("sheet10.NetPLFromSpecBus").value = "" Then
Sheet9.Range("sheet10.NetPLFromSpecBus").formula = "=PL_65iv"
End If
End If

'If (Target.name.name) = "sheet10.IncRecCredPLOthHeads4Div" Then
'    If Not Sheet4.Range("PL.Dividends").value = "" Then
'        If Target.value > Sheet4.Range("PL.Dividends").value Then
'            fmsgbox "* Divident Income can not be higher then Profit and loss divident income"
'            Range(Target.Address).value = "0"
'        End If
'    End If
'End If

If Target.name.name = "sheet12.DeductUs35AD" Then
    If Sheet9.Range("sheet12.DeductUs35AD").value > 0 Then
        sPassword = EfilingCommon.getmsgstate
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("EI.Drpdn").value = "(Select)"
        Sheet9.Range("EI.Drpdn").Locked = False
        Sheet9.Range("EI.Drpdn").Interior.Color = "&HCCFFCC"
        Sheet9.Protect Password:=sPassword
    Else
        sPassword = EfilingCommon.getmsgstate
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("EI.Drpdn").value = "(Select)"
        Sheet9.Range("EI.Drpdn").Locked = True
        Sheet9.Range("EI.Drpdn").Interior.Color = "&HD8D8D8"
        Sheet9.Protect Password:=sPassword
    End If
End If
If Target.name.name = "sheet10.NetPLFromSpecifiedBus" Then
If Sheet9.Range("sheet12.NetPLFrmSpecifiedBus").value > 0 And SheetALL.Range("bacValue").value = 2 Then

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
'change
'        Sheet9.Range("sheet12.DeductUs35AD").ClearContents
'change
        Sheet9.Range("sheet12.DeductUs35AD").Locked = False
        Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = "&HCCFFCC"
        Sheet9.Protect Password:=sPassword
 Else
 '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword

'----------------Unlock Password-------------------END-----

If SheetALL.Range("bacValue").value = 2 Then
        Sheet9.Range("sheet12.DeductUs35AD").ClearContents
        Sheet9.Range("sheet12.DeductUs35AD").Locked = True
        Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = "&HD8D8D8"
        Sheet9.Protect Password:=sPassword
End If
End If
End If

Sheet9.Protect Password:=sPassword
Application.EnableEvents = True
End Sub
Private Sub Worksheet_Changeold1(ByVal Target As Range)
Application.EnableEvents = False
On Error Resume Next

CalculateTaxCheckFlag = False
If (Target.name.name) = "sheet10.NetPLFromSpecBus" Then
If Sheet9.Range("sheet10.NetPLFromSpecBus").value = "" Then
Sheet9.Range("sheet10.NetPLFromSpecBus").formula = "=PL_65iv"
End If
End If

'If (Target.name.name) = "sheet10.IncRecCredPLOthHeads4Div" Then
'    If Not Sheet4.Range("PL.Dividends").value = "" Then
'        If Target.value > Sheet4.Range("PL.Dividends").value Then
'            fmsgbox "* Divident Income can not be higher then Profit and loss divident income"
'            Range(Target.Address).value = "0"
'        End If
'    End If
'End If

If Target.name.name = "sheet12.DeductUs35AD" Then
    If Sheet9.Range("sheet12.DeductUs35AD").value > 0 Then
    
        sPassword = EfilingCommon.getmsgstate
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("EI.Drpdn").value = "(Select)"
        Sheet9.Range("EI.Drpdn").Locked = False
        Sheet9.Range("EI.Drpdn").Interior.Color = "&HCCFFCC"
        Sheet9.Protect Password:=sPassword
    
    Else
        
        sPassword = EfilingCommon.getmsgstate
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("EI.Drpdn").value = "(Select)"
        Sheet9.Range("EI.Drpdn").Locked = True
        Sheet9.Range("EI.Drpdn").Interior.Color = "&HD8D8D8"
        Sheet9.Protect Password:=sPassword
        
    End If
End If

If Target.name.name = "sheet10.NetPLFromSpecifiedBus" Then

If Sheet9.Range("sheet12.NetPLFrmSpecifiedBus").value > 0 And SheetALL.Range("bacValue").value = 2 Then


'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
        Sheet9.Range("sheet12.DeductUs35AD").ClearContents
        Sheet9.Range("sheet12.DeductUs35AD").Locked = False
        Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = "&HCCFFCC"
        Sheet9.Protect Password:=sPassword
 Else
 '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
If SheetALL.Range("bacValue").value = 2 Then
        Sheet9.Range("sheet12.DeductUs35AD").ClearContents
        Sheet9.Range("sheet12.DeductUs35AD").Locked = True
        Sheet9.Range("sheet12.DeductUs35AD").Interior.Color = "&HD8D8D8"
        Sheet9.Protect Password:=sPassword
End If
End If
End If

Sheet9.Protect Password:=sPassword
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

Sub Worksheet_Deactivate()
If (Sheet3.Range("sheet2.TotFundSrc").value = 0) And (Sheet3.Range("sheet3.TotSundryDbtAmt").value = 0) And (Sheet3.Range("sheet3.TotSundryCrdAmt").value = 0) And (Sheet3.Range("sheet3.TotStkInTradAmt").value = 0) And (Sheet3.Range("sheet3.CashBalAmt").value = 0) And (Sheet9.Range("sheet12.IncChrgUnHdProftGain").value > 250000) Then
Dim QLine1, QLine2, Nchar As String
QLine1 = "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)"
QLine2 = "Note - If Nature of business code is only 00001-Share of Income from firm only, then above Pop up should not come"
Nchar = "\n"
'Change-23.05.2023.101 -> Issue number 4 in production build 1
'MsgBox QLine1 + Nchar + QLine2
MsgBox QLine1
'----
End If


Dim cnt00001
Dim cnttemp, cell
cnttemp = 0
For Each cell In Sheet2.Range("NOB.Code").Cells
    If Not isdropdownblank(cell.value) Then
        cnttemp = cnttemp + 1
    End If
Next cell
 cnt00001 = Application.CountIf(Sheet2.Range("NOB.Code"), "00001-Share of Income from firm only")
    If cnt00001 > 0 And cnt00001 = cnttemp And Sheet9.Range("sheet11.DebPLUs35ExcessAmt") = 0 And Sheet9.Range("sheet10.ProfBfrTaxPL") = 0 And SheetALL.Range("bacValue") = 2 Then
        ThisWorkbook.Unprotect Password:=getmsgstate
        Sheet56.Unprotect Password:=getmsgstate
        Sheet56.Visible = xlSheetVisible
        Sheet56.Protect Password:=getmsgstate
        ThisWorkbook.Protect Password:=getmsgstate
    
    Else
        ThisWorkbook.Unprotect Password:=getmsgstate
        Sheet56.Unprotect Password:=getmsgstate
        Sheet56.Visible = xlSheetHidden
        Sheet56.Protect Password:=getmsgstate
        ThisWorkbook.Protect Password:=getmsgstate
    
    End If




End Sub

Sub test()
Application.EnableEvents = True
End Sub
