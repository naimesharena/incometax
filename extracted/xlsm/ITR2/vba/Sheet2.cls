Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Sheet2.Unprotect Password:=getmsgstate
Application.EnableEvents = False

If Target.Column = 5 Then
    If (Not Target.value = "(Select)") And (Not Target.value = "") Then
      Dim cnt05001, cnt06008, cnt01003, cnt01002, cnt01001, cnt00001 As Variant
      cnt05001 = 0
      cnt06008 = 0
      cnt01003 = 0
      cnt01002 = 0
      cnt01001 = 0
      cnt00001 = 0
      
      cnt05001 = Application.CountIf(Range("NOB.Code"), "05001-Production, collection and distribution of electricity")
      cnt06008 = Application.CountIf(Range("NOB.Code"), "06008-Construction and maintenance of power transmission and telecommunication lines")
      cnt01003 = Application.CountIf(Range("NOB.Code"), "01003-Growing and manufacturing of rubber")
      cnt01002 = Application.CountIf(Range("NOB.Code"), "01002-Growing and manufacturing of coffee")
      cnt01001 = Application.CountIf(Range("NOB.Code"), "01001-Growing and manufacturing of tea")
      cnt00001 = Application.CountIf(Range("NOB.Code"), "00001-Share of Income from firm only")
      
        If cnt05001 > 0 Then
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.Locked = False
            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet9.Protect Password:=getmsgstate
        Else
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.Locked = True
            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").Interior.Color = (&HD8D8D8)
            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.ClearContents
            Sheet9.Protect Password:=getmsgstate
        End If
        If cnt01003 > 0 Then
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("BP_4b_ProfitRule7A").MergeArea.Locked = False
            Sheet9.Range("BP_4b_ProfitRule7A").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet9.Protect Password:=getmsgstate
        Else
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("BP_4b_ProfitRule7A").MergeArea.ClearContents
            Sheet9.Range("BP_4b_ProfitRule7A").MergeArea.Locked = True
            Sheet9.Range("BP_4b_ProfitRule7A").Interior.Color = (&HD8D8D8)
            Sheet9.Protect Password:=getmsgstate
        End If
        If cnt01002 > 0 Then
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("BP_4b_ProfitRule7B1").MergeArea.Locked = False
            Sheet9.Range("BP_4b_ProfitRule7B1").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet9.Range("BP_4b_ProfitRule7B2").MergeArea.Locked = False
            Sheet9.Range("BP_4b_ProfitRule7B2").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet9.Protect Password:=getmsgstate
        Else
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("BP_4b_ProfitRule7B1").MergeArea.ClearContents
            Sheet9.Range("BP_4b_ProfitRule7B1").MergeArea.Locked = True
            Sheet9.Range("BP_4b_ProfitRule7B1").Interior.Color = (&HD8D8D8)
            
            Sheet9.Range("BP_4b_ProfitRule7B2").MergeArea.ClearContents
            Sheet9.Range("BP_4b_ProfitRule7B2").MergeArea.Locked = True
            Sheet9.Range("BP_4b_ProfitRule7B2").Interior.Color = (&HD8D8D8)
            Sheet9.Protect Password:=getmsgstate
        End If
        If cnt01001 > 0 Then
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("BP_4b_ProfitRule8").MergeArea.Locked = False
            Sheet9.Range("BP_4b_ProfitRule8").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet9.Protect Password:=getmsgstate
        Else
            Sheet9.Unprotect Password:=getmsgstate
            Sheet9.Range("BP_4b_ProfitRule8").MergeArea.ClearContents
            Sheet9.Range("BP_4b_ProfitRule8").MergeArea.Locked = True
            Sheet9.Range("BP_4b_ProfitRule8").Interior.Color = (&HD8D8D8)
            Sheet9.Protect Password:=getmsgstate
        End If
        
        If cnt00001 > 0 Then
        SheetALL.Unprotect Password:=getmsgstate
            SheetALL.Range("NOBCODE00001").value = 1
            
            Else
             
            SheetALL.Range("NOBCODE00001").value = 0
         SheetALL.Protect Password:=getmsgstate
        End If
        
        
        If 1 < Application.CountIf(Range("NOB.Code"), Target.value) Then
             fmsgbox "* Duplicate Code-Sub Sector is not allowed"
             Range(Target.Address).value = "(Select)"
             Range("E" & Target.row).value = "(Select)"
        End If
End If
    
     If InStr(Target.value, "Other") <> 0 Then
            If Range("NOB.TradeNameOrDescription2").item(Target.row).value = "" Then
                 fmsgbox "* Description is compulsory: schedule Nature of business"
            End If
          End If
     
          
End If



If (Sheet2.Range("Sheet2.Code1").value <> Sheet2.Range("Sheet2.Code2").value) And (Sheet2.Range("Sheet2.Code1").value <> Sheet2.Range("Sheet.Code3").value) And (Sheet2.Range("Sheet2.Code2").value <> Sheet2.Range("Sheet.Code3").value) Then
    Sheet9.Unprotect Password:=getmsgstate

            
'            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.Locked = False
'            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.Interior.Color = (&HCCFFCC)
'            Sheet9.Protect Password:=getmsgstate
        Else
'            Sheet9.Unprotect Password:=getmsgstate
'
'            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.ClearContents
'            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").MergeArea.Locked = True
'            Sheet9.Range("sheet11.DepreciationAllowUs32_1_i").Interior.Color = (&HD8D8D8)
'
'            Sheet9.Protect Password:=getmsgstate
End If


endline:
Application.EnableEvents = True
Sheet2.Protect Password:=getmsgstate
End Sub

Sub LOCKPL1()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_PL1 As Variant
rngname_PL1 = "PL_61a||PL_61b||PL_61iia||PL_61iib||PL_62i||PL_62ii"
LOCKRANGEPL1 (rngname_PL1)




'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGEPL1(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))
           If chcell.MergeCells Then
           chcell.MergeArea.value = ""
           chcell.MergeArea.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Else
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           End If
           Next
       Next
End Sub

Sub UNLOCKBS()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet3.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_BS As Variant


rngname_BS = "AllBS.names||sheet3.TotSundryCrdAmt||sheet3.TotStkInTradAmt||sheet3.CashBalAmt"
UNLOCKRANGEBS (rngname_BS)

'----------------Lock Password-------------------START---
  Sheet3.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub


Sub LOCKBS()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet3.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_BS As Variant
rngname_BS = "AllBS.names||sheet3.TotSundryCrdAmt||sheet3.TotStkInTradAmt||sheet3.CashBalAmt"
LOCKRANGEBS (rngname_BS)




'----------------Lock Password-------------------START---
  Sheet3.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGEBS(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet3.Range(rangenamestring(i))
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Next
       Next
End Sub

Sub UNLOCKRANGEBS(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet3.Range(rangenamestring(i))
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Next
       Next
End Sub




' PL

Sub UNLOCKPL()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_PL As Variant


'rngname_PL = "Allpl.name||Nature_Name2||Nature_Amt2||Nature_Name3||Nature_Amt3||PLBD.PAN||PLBD.Amount||PL.ProvForBadDoubtDebt||PL.OthProvisionsExpdr||PLI.NonResOtherCompany||PLI.Others||PL.DepreciationAmort||PL.ProvForCurrTax||PL.ProvDefTax||PL.BalBFPrevYr||PL.TrfToReserves||Manufacture1"
rngname_PL = "ALLPL_NAMES1||ALLPL_NAMES2||ALLPL_NAMES3"
UNLOCKRANGEPL (rngname_PL)

'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub


Sub LOCKPL()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_PL As Variant
'rngname_PL = "Allpl.name||Nature_Name2||Nature_Amt2||Nature_Name3||Nature_Amt3||PLBD.PAN||PLBD.Amount||PL.ProvForBadDoubtDebt||PL.OthProvisionsExpdr||PLI.NonResOtherCompany||PLI.Others||PL.DepreciationAmort||PL.ProvForCurrTax||PL.ProvDefTax||PL.BalBFPrevYr||PL.TrfToReserves"
rngname_PL = "ALLPL_NAMES1||ALLPL_NAMES2||ALLPL_NAMES3"
LOCKRANGEPL (rngname_PL)




'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGEPL(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))
            
           If chcell.MergeCells Then
           chcell.MergeArea.value = ""
           chcell.MergeArea.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Else
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           End If
           Next
       Next
End Sub

Sub UNLOCKRANGEPL(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))
           If chcell.MergeCells Then
           chcell.MergeArea.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Else
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           End If
           Next
       Next
End Sub

' PL

Sub UNLOCKMA()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet48.Unprotect Password:=sPassword
'   Sheet48.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_MA As Variant


rngname_MA = "ALLMA_NAMES"
UNLOCKRANGEMA (rngname_MA)

'----------------Lock Password-------------------START---
  Sheet48.Protect Password:=sPassword
'   Sheet48.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub


Sub LOCKMA()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet48.Unprotect Password:=sPassword
'   Sheet48.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_MA As Variant
rngname_MA = "ALLMA_NAMES"
LOCKRANGEMA (rngname_MA)




'----------------Lock Password-------------------START---
  Sheet48.Protect Password:=sPassword
'   Sheet48.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGEMA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet48.Range(rangenamestring(i))
            
           If chcell.MergeCells Then
           chcell.MergeArea.value = ""
           chcell.MergeArea.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Else
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           End If
           Next
       Next
End Sub

Sub UNLOCKRANGEMA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet48.Range(rangenamestring(i))
           If chcell.MergeCells Then
           chcell.MergeArea.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Else
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           End If
           Next
       Next
End Sub

' PL

Sub UNLOCKTA()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet49.Unprotect Password:=sPassword
'   Sheet48.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_TA As Variant


rngname_TA = "AllTA_Names"
UNLOCKRANGETA (rngname_TA)

'----------------Lock Password-------------------START---
  Sheet49.Protect Password:=sPassword
'   Sheet48.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub


Sub LOCKTA()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet49.Unprotect Password:=sPassword
'   Sheet48.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_TA As Variant
rngname_TA = "AllTA_Names"
LOCKRANGETA (rngname_TA)




'----------------Lock Password-------------------START---
  Sheet49.Protect Password:=sPassword
'   Sheet48.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGETA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet49.Range(rangenamestring(i))
           
           If chcell.MergeCells Then
           chcell.MergeArea.value = ""
           chcell.MergeArea.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Else
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           End If
           Next
       Next
End Sub

Sub UNLOCKRANGETA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet49.Range(rangenamestring(i))
           
           If chcell.MergeCells Then
           chcell.MergeArea.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Else
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           End If
           Next
       Next
End Sub



Private Sub Worksheet_Deactivate()
'Dim cnt00001
' cnt00001 = Application.CountIf(Sheet2.Range("NOB.Code"), "00001-Share of Income from firm only")
'Dim cnttemp, cell
'cnttemp = 0
'For Each cell In Sheet2.Range("NOB.Code").Cells
'    If Not isdropdownblank(cell.value) Then
'        cnttemp = cnttemp + 1
'    End If
'Next cell
'    If cnt00001 > 0 And cnt00001 = cnttemp And Sheet9.Range("sheet11.DebPLUs35ExcessAmt") = 0 And Sheet9.Range("sheet10.ProfBfrTaxPL") = 0 And SheetALL.Range("bacValue") = 2 Then
'        ThisWorkbook.Unprotect Password:=getmsgstate
'        Sheet56.Unprotect Password:=getmsgstate
'        Sheet56.Visible = xlSheetVisible
'        Sheet56.Protect Password:=getmsgstate
'        ThisWorkbook.Protect Password:=getmsgstate
'
'    Else
'        ThisWorkbook.Unprotect Password:=getmsgstate
'        Sheet56.Unprotect Password:=getmsgstate
'        'Sheet56.Visible = xlSheetHidden
'        Sheet56.Protect Password:=getmsgstate
'        ThisWorkbook.Protect Password:=getmsgstate
'
'    End If
'

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
