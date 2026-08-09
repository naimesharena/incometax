Attribute VB_Name = "Sheet53"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

'
Private Sub Worksheet_Activate()
On Error Resume Next
    ActiveSheet.Shapes("Button 1").Left = Range("Button_112A").Left + 2
    ActiveSheet.Shapes("Button 1").Top = Range("Button_112A").Top + 2
    
'      MsgBox ("when user opens schedule CG--If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F. Give a button-""OK"". Only when user clicks on ""OK"", he will be able to proceed")
       fmsgboxStatus ("If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F.")

'    ThisWorkbook.Protect Password:=getmsgstate
End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim targetadd As Variant
Application.EnableEvents = False
Application.ScreenUpdating = False
Sheet53.Unprotect Password:=getmsgstate
targetadd = Replace(Target.Address, "$", "")

If Not Application.Intersect(Range("ShareAcq_112A"), Range(Target.Address)) Is Nothing Then

    If Target.value = "After 31st January 2018" Then
    
'        Range(Replace(target.Address, "E", "F")).ClearContents
'        Range(Replace(target.Address, "E", "F")).Locked = True
'        Range(Replace(target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
'        Range(Replace(target.Address, "E", "F")).value = "INNOTREQUIRD"
'
'        Range(Replace(target.Address, "E", "G")).ClearContents
'        Range(Replace(target.Address, "E", "G")).Locked = True
'        Range(Replace(target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
'        Range(Replace(target.Address, "E", "G")).value = "CONSOLIDATED"
'
'        Range(Replace(target.Address, "E", "H")).ClearContents
'        Range(Replace(target.Address, "E", "H")).Locked = True
'        Range(Replace(target.Address, "E", "H")).Interior.Color = "&HD8D8D8"
'        Range(Replace(target.Address, "E", "H")).value = 0
'
'        Range(Replace(target.Address, "E", "I")).ClearContents
'        Range(Replace(target.Address, "E", "I")).Locked = True
'        Range(Replace(target.Address, "E", "I")).Interior.Color = "&HD8D8D8"
'        Range(Replace(target.Address, "E", "I")).value = 0
'
'        Range(Replace(target.Address, "E", "N")).ClearContents
'        Range(Replace(target.Address, "E", "N")).Locked = True
'        Range(Replace(target.Address, "E", "N")).Interior.Color = "&HD8D8D8"
'        Range(Replace(target.Address, "E", "N")).value = 0
'
'        Range(Replace(target.Address, "E", "J")).ClearContents
'        Range(Replace(target.Address, "E", "J")).Locked = False
'        Range(Replace(target.Address, "E", "J")).Interior.Color = "&HCCFFCC"
'        Range(Replace(target.Address, "E", "J")).Font.Color = vbBlack
        Range(Replace(Target.Address, "E", "G")).ClearContents
        Range(Replace(Target.Address, "E", "G")).Locked = True
        Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "G")).value = "INNOTREQUIRD"
        
        Range(Replace(Target.Address, "E", "H")).ClearContents
        Range(Replace(Target.Address, "E", "H")).Locked = True
        Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "H")).value = "CONSOLIDATED"

        Range(Replace(Target.Address, "E", "I")).ClearContents
        Range(Replace(Target.Address, "E", "I")).Locked = True
        Range(Replace(Target.Address, "E", "I")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "I")).value = 0

        Range(Replace(Target.Address, "E", "J")).ClearContents
        Range(Replace(Target.Address, "E", "J")).Locked = True
        Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "J")).value = 0

        Range(Replace(Target.Address, "E", "O")).ClearContents
        Range(Replace(Target.Address, "E", "O")).Locked = True
        Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HD8D8D8"
        Range(Replace(Target.Address, "E", "O")).value = 0
        
        Range(Replace(Target.Address, "E", "K")).ClearContents
        Range(Replace(Target.Address, "E", "K")).Locked = False
        Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
        Range(Replace(Target.Address, "E", "K")).Font.Color = vbBlack

    Else
    
''        Range(Replace(Target.address, "E", "F")).ClearContents
'        Range(Replace(target.Address, "E", "F")).Locked = False
'        Range(Replace(target.Address, "E", "F")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "G")).ClearContents
'        Range(Replace(target.Address, "E", "G")).Locked = False
'        Range(Replace(target.Address, "E", "G")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "H")).ClearContents
'        Range(Replace(target.Address, "E", "H")).Locked = False
'        Range(Replace(target.Address, "E", "H")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "I")).ClearContents
'        Range(Replace(target.Address, "E", "I")).Locked = False
'        Range(Replace(target.Address, "E", "I")).Interior.Color = "&HCCFFCC"
'
''        Range(Replace(Target.address, "E", "N")).ClearContents
'        Range(Replace(target.Address, "E", "N")).Locked = False
'        Range(Replace(target.Address, "E", "N")).Interior.Color = "&HCCFFCC"
'
'
''        Range(Replace(Target.address, "E", "J")).ClearContents
'        Range(Replace(target.Address, "E", "J")).Locked = True
'        Range(Replace(target.Address, "E", "J")).Interior.Color = vbWhite
'        Range(Replace(target.Address, "E", "J")).Font.Color = RGB(51, 102, 255)
        Range(Replace(Target.Address, "E", "G")).Locked = False
        Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"
        
'        Range(Replace(Target.address, "E", "G")).ClearContents
        Range(Replace(Target.Address, "E", "H")).Locked = False
        Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HCCFFCC"

'        Range(Replace(Target.address, "E", "H")).ClearContents
        Range(Replace(Target.Address, "E", "I")).Locked = False
        Range(Replace(Target.Address, "E", "I")).Interior.Color = "&HCCFFCC"

'        Range(Replace(Target.address, "E", "I")).ClearContents
        Range(Replace(Target.Address, "E", "J")).Locked = False
        Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HCCFFCC"

'        Range(Replace(Target.address, "E", "N")).ClearContents
        Range(Replace(Target.Address, "E", "O")).Locked = False
        Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HCCFFCC"
        
        
'        Range(Replace(Target.address, "E", "J")).ClearContents
        Range(Replace(Target.Address, "E", "K")).Locked = True
        Range(Replace(Target.Address, "E", "K")).Interior.Color = vbWhite
        Range(Replace(Target.Address, "E", "K")).Font.Color = RGB(51, 102, 255)
        Range(Replace(Target.Address, "E", "K")).ClearContents

        'MsgBox ("=MAX(0," & "ROUND(" & Replace(TargetAdd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(TargetAdd, "E", "I") & ",4)" & ")")
'        Range(Replace(target.Address, "E", "J")).formula = "=ROUND(MAX(0," & "ROUND(" & Replace(TargetAdd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(TargetAdd, "E", "I") & ",4)" & "),0)"
         Range(Replace(Target.Address, "E", "k")).formula = "=ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "J") & ",4)" & "),0)"

        '=MAX(0,H6*I6)
        'vbWhite
    End If
    If Target.value = "" Then
       Target.value = "(Select)"
    End If
End If


If Not Application.Intersect(Range("ISINCode_112A"), Range(Target.Address)) Is Nothing Then
If Target.value <> "" Then

Target.value = UCase(Target.value)
 If UCase(Mid(Range(Target.Address), 1, 2)) <> "IN" Then
 fmsgboxStatus "Please enter the valid ISIN Code in sheet 112A"
 Target.value = ""
 End If
    Sheet53.Range(Target.Address).Select
    Sheet53.Range(Target.Address).Offset(0, 1) = UCase(Sheet53.Range(Target.Address).Offset(0, 26).value)
End If

If Target.value = "" Then
    Sheet53.Range(Target.Address).Offset(0, 1) = ""
End If

End If

endline:
Application.EnableEvents = True
Application.ScreenUpdating = True
Sheet53.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Deactivate()
On Error GoTo endline
Dim activesh As Variant
activesh = ActiveSheet.name

CG_Calc.doSetoff
CYLACalculations.CYLABFLASetOff
Sheets(activesh).Activate


If ThisWorkbook.ProtectStructure = False Then
    ThisWorkbook.Protect Password:=getmsgstate
End If
'Adding the condition of 115AD lock if any one row is filled in 112A
    Dim end_112A As Long
    Dim range115AD As Variant
    
    setTableInfo112A
    setTableInfo112A1
    setTableInfo112A2
    setTableInfo112A3
    setTableInfo112A4
    setTableInfo112A5
'    setTableInfo112A6
    setTableInfo112A7
    setTableInfo112A8
    
    end_112A = WorksheetFunction.Max(end112A, end112A1, end112A2, end112A3, end112A4, end112A5, end112A7, end112A8)
    
    Application.EnableEvents = False
    Sheet54.Unprotect Password:=getmsgstate
    If end_112A > 0 Then
        If LockFlag115A = False Then
            MsgBox ("Please fill either Schedule 112A or Schedule 115AD(1)(b)(iii) proviso as applicable in your case")
        End If
'        range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
  'Modified by sai on 07/03/2025
  
  'changed by Chetan C M on 30/01/2026 for AY 2026-27
    'start--
        'range115AD = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
        range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
    '--end
        
        LOCKRANGE115AD (range115AD)
        
    Else
        If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" And LockFlag115A = True Then
'            range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
'Modified by sai on 07/03/2025

'changed by Chetan C M on 30/01/2026 for AY 2026-27
    'start--
            'range115AD = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
            range115AD = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
    '--end
            
            UNLOCKRANGE115AD (range115AD)
        End If
        'call the unlock115Ad
    End If
    Application.EnableEvents = True
    Sheet54.Protect Password:=getmsgstate
    
endline:
    Application.EnableEvents = True
    Sheet54.Protect Password:=getmsgstate
End Sub




