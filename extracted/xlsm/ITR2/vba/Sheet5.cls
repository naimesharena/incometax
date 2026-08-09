Attribute VB_Name = "Sheet5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
'Application.ScreenUpdating = False
On Error GoTo endline

CalculateTaxCheckFlag = False

'If (Target.name.name) = "sheet5.ChangeInAcctMethFlg" Then Call LockUnlock("sheet5.ChangeInAcctMethFlg", "sheet5.ProfDeviatDueAcctMeth")
If (Target.name.name) = "sheet6.ChngStockValMetFlg" Then Call LockUnlock("sheet6.ChngStockValMetFlg", "sheet6.EffectOnPL", "sheet6.EffectOnPL4e")


If (Target.name.name) = "sheet7.Section92CE_Flag" Then

LockUnlockTPSA

If Sheet5.Range("sheet7.Section92CE_Flag").value = "Yes" And Sheet57.Range("TPSA_92CE_Amount").value = "" Then
MsgBox " please fill schedule TPSA"
End If

sPassword = getmsgstate
Sheet57.Unprotect Password:=sPassword

'If Sheet5.Range("sheet7.Section92CE_Flag").value = "Yes" Then
''Sheet57.Range("TPSA_92CE_Amount").value = ""
'Sheet57.Range("TPSA_92CE_Amount").Locked = True
'Sheet57.Range("TPSA_92CE_Amount").Interior.ColorIndex = 2
'
'Else
''Sheet57.Range("TPSA_92CE_Amount").value = ""
'Sheet57.Range("TPSA_92CE_Amount").Locked = True
'Sheet57.Range("TPSA_92CE_Amount").Interior.ColorIndex = 2
'
'End If
Sheet57.Protect Password:=sPassword

End If


endline:
Application.EnableEvents = True
'Application.ScreenUpdating = True
End Sub


Sub LockUnlock(Src As String, Dst As String, Dst1 As String)
On Error GoTo endline
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    If Range(Src) = "No" Or Range(Src) = "(Select)" Then
        Range(Dst).Select
        Selection.ClearContents
        Selection.Locked = True
        Selection.Interior.Color = (&HD8D8D8)
        Range(Dst1).Select
        Selection.ClearContents
        Selection.Locked = True
        Selection.Interior.Color = (&HD8D8D8)
    
    Else
        Range(Dst).Select
        Selection.Locked = False
        Selection.ClearContents
        Selection.Interior.Color = (&HCCFFCC)
         Range(Dst1).Select
        Selection.Locked = False
        Selection.ClearContents
        Selection.Interior.Color = (&HCCFFCC)
    End If
    
endline:
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
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
Sub LockUnlockTPSA()
On Error Resume Next

    If Sheet5.Range("sheet7.Section92CE_Flag").value = "Yes" Then
    'Or (Sheet1.Range("sheet1.ResidentialStatus1").value) = "(Select)" Then
        UnlockTPSA
    Else
        LockTPSA

    End If
End Sub

Sub LockTPSA()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet57.Unprotect Password:=sPassword

        rngname = "TPSA_92CE_Amount||TPSC_BSRCode||TPSC_BankName||TPSC_DateDep||TPSC_SrlNoOfChaln||TPSC_Amt||TPSA.Amount"
        LOCKRANGETPSA (rngname)

    Sheet57.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnlockTPSA()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet57.Unprotect Password:=sPassword

        rngname = "TPSA_92CE_Amount||TPSC_BSRCode||TPSC_BankName||TPSC_DateDep||TPSC_SrlNoOfChaln||TPSC_Amt||TPSA.Amount"
        UNLOCKRANGETPSA (rngname)

    Sheet57.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub LOCKRANGETPSA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet57.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub
Sub UNLOCKRANGETPSA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet57.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub
