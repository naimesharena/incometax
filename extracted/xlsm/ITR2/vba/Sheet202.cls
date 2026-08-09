Attribute VB_Name = "Sheet202"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Sub LockUnlockATI_1()
    Application.EnableEvents = False
    Sheet202.Unprotect Password:=getmsgstate
    
    If Sheet202.Range("U_LastAmtPayable").value > 0 Then
        Sheet202.Range("U_Refund").MergeArea.ClearContents
        Sheet202.Range("U_Refund").MergeArea.Locked = True
        Sheet202.Range("U_Refund").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_Refund").value = ""
        
        Sheet202.Range("U_TotRefund").MergeArea.ClearContents
        Sheet202.Range("U_TotRefund").MergeArea.Locked = True
        Sheet202.Range("U_TotRefund").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_TotRefund").value = ""
    Else
        Sheet202.Range("U_Refund").MergeArea.Locked = False
        Sheet202.Range("U_Refund").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_Refund").value = 0
        Sheet202.Range("U_TotRefund").MergeArea.Locked = False
        Sheet202.Range("U_TotRefund").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_TotRefund").value = 0
    End If
    
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

'Sheet1.Range("taxcheck").value = ""

Sheet202.Unprotect Password:=getmsgstate
Application.EnableEvents = False


'Added by Sadineni on 09/12/2025
'======================================================================================================================
If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then

Dim dateOfdeposit As Variant
Dim AssYear As Long
Dim yrdop As Long
Dim mthdop As Long

dateOfdeposit = Target.value
        
        AssYear = 2025
        dateOfdeposit = Dformat(dateOfdeposit, "yyyy-mm-dd")
        
        yrdop = CLng(Mid(dateOfdeposit, 1, 4))
        mthdop = CLng(Mid(dateOfdeposit, 6, 2))
        
            If (yrdop >= AssYear) Then
                dateOfdeposit = mthdop - 4 + (yrdop - AssYear) * 12
            End If
        
        dateOfdeposit = dateOfdeposit + 1
    Sheet202.Range(Replace(Replace(Target.Address, "$", ""), "G", "T")).value = dateOfdeposit
End If
'===========================================================================================================================

'Shrutika_UR
If Not Application.Intersect(Sheet202.Range("U_LastAmtPayable"), Sheet202.Range(Target.Address)) Is Nothing Then

     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_LastAmtPayable").value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then ""Amount payable as per prior ITR can’t be more than 0""", , "Error"
           Sheet202.Range("U_LastAmtPayable").ClearContents

        End If
     End If

End If

'Shrutika_UR
If Not Application.Intersect(Sheet202.Range("U_BSRCode2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_SrlNoOfChaln2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_Amt2"), Sheet202.Range(Target.Address)) Is Nothing Then
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" Then
         LOCKRANGEIT2
    End If
End If

If Not Application.Intersect(Sheet202.Range("U_Refund"), Sheet202.Range(Target.Address)) Is Nothing Then

     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_Refund").value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then ""Refund claimed can’t be more than 0""", , "Error"
           Sheet202.Range("U_Refund").ClearContents

        End If
     End If

End If


If Not Application.Intersect(Sheet202.Range("U_TotRefund"), Sheet202.Range(Target.Address)) Is Nothing Then

     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_TotRefund").value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"", then ""Refund issued can’t be more than 0""", , "Error"
           Sheet202.Range("U_TotRefund").ClearContents

        End If
     End If

End If


If Not Application.Intersect(Sheet202.Range("U_RegAssessementTAX"), Sheet202.Range(Target.Address)) Is Nothing Then

     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_RegAssessementTAX").value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"", then ""Regular Assessment Tax"" can't be claimed""", , "Error"
           Sheet202.Range("U_RegAssessementTAX").ClearContents

        End If
     End If

End If

'----------

'Change-25.11.2022.104.02A
If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then
                IT1_DateofDeposit Target
'                Dim temp
'                temp = Dformat(Sheet202.Range(Target.Address).value, "")
'                If Not ChkMinInclusiveDate(temp, "2023-01-01") Then
'
'                    fmsgboxsmall ("Date of Deposit in Sheet IT can not be prior to 01/01/2023")
'                    Sheet202.Range(Target.Address).value = ""
'
'                End If

       

End If



'Change-25.11.2022.104.02A
If Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Then
                IT2_DateofDeposit Target
'                    Dim temp1
'                temp1 = Dformat(Sheet202.Range(Target.Address).value, "")
'                If Not ChkMinInclusiveDate(temp1, "2021-04-01") Then
'
'                    fmsgboxsmall ("Date of Deposit in Sheet IT can not be prior to 01/04/2021")
'                    Sheet202.Range(Target.Address).value = ""
'
'                End If

       

End If

If Target.Column = 6 Then
    Target.value = UCase(Target.value)
End If
If Not Application.Intersect(Sheet202.Range("U_LastAmtPayable"), Sheet202.Range(Target.Address)) Is Nothing Then
LockUnlockATI_1
End If

If Not Application.Intersect(Sheet202.Range("U_Refund"), Sheet202.Range(Target.Address)) Is Nothing Then
    If Sheet202.Range("U_Refund").value > 0 Or Sheet202.Range("U_TotRefund").value > 0 Then
        Sheet202.Range("U_LastAmtPayable").MergeArea.ClearContents
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = True
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LastAmtPayable").value = ""
    Else
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = False
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_LastAmtPayable").value = 0

    End If
End If

If Not Application.Intersect(Sheet202.Range("U_TotRefund"), Sheet202.Range(Target.Address)) Is Nothing Then
    If Sheet202.Range("U_TotRefund").value > 0 Or Sheet202.Range("U_Refund").value > 0 Then
        Sheet202.Range("U_LastAmtPayable").MergeArea.ClearContents
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = True
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LastAmtPayable").value = ""

    Else
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = False
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_LastAmtPayable").value = 0
    End If
End If

'Added by Shrutika for SIT-103705
'start--
If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then
 
    If Not CheckDateddmmyyyy(Sheet202.Range(Target.Address).value) Then
     MsgBox ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2026 in schedule IT.""")
  
    Sheet202.Range(Target.Address).value = ""
    Target.Select
    End If
End If

If Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Then
 
    If Not CheckDateddmmyyyy(Sheet202.Range(Target.Address).value) Then
     MsgBox ("""Please enter valid date in dd/mm/yyyy format in schedule IT-2. Also date of credit to Central government cannot be prior to 01/04/2024 in schedule IT-2.""")

    Sheet202.Range(Target.Address).value = ""
    Target.Select
    End If
End If
'--end

endline:
Application.EnableEvents = True
Sheet202.Protect Password:=getmsgstate
End Sub







