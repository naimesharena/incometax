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
Sub temp()
'Sheet201.Unprotect Password:=getmsgstate
Application.EnableEvents = True



'Sheet202.Range("U_IncomeFromCG").Locked = True
'Sheet202.Range("U_IncomeFromCG").Interior.Color = "&HD8D8D8"
'Sheet202.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Activate()
 Sheet202.Protect Password:=getmsgstate

End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

'Sheet1.Range("taxcheck").value = ""
CalculateTaxCheckFlag = False
Sheet202.Unprotect Password:=getmsgstate
Application.EnableEvents = False

'Ayush_UR

'If Not Application.Intersect(Sheet202.Range("U_LatestTotInc"), Sheet202.Range(Target.Address)) Is Nothing Then
'
'     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("Yes") Then
'        If Sheet202.Range("U_LatestTotInc").value = "" Then
'           Sheet202.Unprotect Password:=getmsgstate
'           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then ""Amount payable as per prior ITR can’t be more than 0""", , "Error"
'           Sheet202.Range("U_LatestTotInc").ClearContents
'
'        End If
'     End If
'
'End If

'Added by Riyaz on 05/12/2025

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

'End

'Ayush_UR
If Not Application.Intersect(Sheet202.Range("U_LastAmtPayable"), Sheet202.Range(Target.Address)) Is Nothing Then
     
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_LastAmtPayable").value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then ""Amount payable as per prior ITR can’t be more than 0""", , "Error"
           Sheet202.Range("U_LastAmtPayable").ClearContents
        
        End If
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
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"", then ""Regular Assessment Tax"" can't be claimed", , "Error"
           Sheet202.Range("U_RegAssessementTAX").ClearContents
        
        End If
     End If

End If

'----------

'Change-24.11.2022.103.02A
If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then
                IT1_DateofDeposit Target
'                If Not FormatNCheckDate(Trim(Range("U_DateDep1").value)) Then
'                MsgBox "*Please enter date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
'               ' Sheet1.Range("U_DateDep1").Select
'                Sheet202.Range("U_DateDep1").value = ""

   'Ayush_14/11/2025
    If Target.value <> "00/00/0000" And Target.value <> "" Then
       If Not CheckDateddmmyyyy(Trim(Target.value)) Then
            fmsgbox "* ""Date of credit into Govt Account must be a valid dd/mm/yyyy format in schedule IT-1""" ', vbOKOnly, mdCommon.vbMessageCaption
            Target.ClearContents
        End If
    End If
                
                
         
'                Dim temp
'                temp = Dformat(Sheet202.Range(Target.Address).value, "")
'                If Not ChkMinInclusiveDate(temp, "2023-01-01") Then
'
'                    fmsgboxsmall ("Date of Deposit in Sheet IT should not be less than 01/01/2023")
'                    Sheet202.Range(Target.Address).value = ""
'
'                End If

End If


'Change-24.11.2022.103.02A
    If Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Then
                IT2_DateofDeposit Target
'                    Dim temp1
'                    temp1 = Dformat(Sheet202.Range(Target.Address).value, "")
'                    If Not ChkMinInclusiveDate(temp1, "2021-04-01") Then
'
'                        fmsgboxsmall ("Date of Deposit in sheet IT should not be less than 01/04/2021")
'                        Sheet202.Range(Target.Address).value = ""
'
'                    End If
'
    'Ayush_14/11/2025
    If Target.value <> "00/00/0000" And Target.value <> "" Then
       If Not CheckDateddmmyyyy(Trim(Target.value)) Then
            fmsgbox "* ""Date of credit into Govt Account must be a valid dd/mm/yyyy format in schedule IT-2""" ', vbOKOnly, mdCommon.vbMessageCaption
            Target.ClearContents
        End If
    End If
                   
    
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

'Malli----------------------11/10/2024
If Not Application.Intersect(Sheet202.Range("U_BSRCode2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_SrlNoOfChaln2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_Amt2"), Sheet202.Range(Target.Address)) Is Nothing Then

    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" Then
         LOCKRANGEIT2
    End If
End If



endline:
Application.EnableEvents = True
Sheet202.Protect Password:=getmsgstate


End Sub

Function ValidateDateDep_IT() As Boolean

    Dim rangecells As Range

    Dim i As Long

    Dim DateDep_IT

Dim mIntCells As Long

 Dim mIntCtr As Long

 mIntCells = Sheet202.Range("U_DateDep1").count

    

    ValidateDateDep_IT = True

    Dim end_IT

    end_IT = 0

    Set rangecells = Sheet202.Range("U_DateDep1").Cells

 For mIntCtr = 1 To mIntCells

     If Not rangecells.item(mIntCtr).value = "" Then

         end_IT = end_IT + 1

     End If

 Next

    ReDim DateDep_IT(end_IT)

    

    For i = 1 To end_IT

        DateDep_IT(i) = rangecells.item(i).value

        If Not chkCompulsory(DateDep_IT(i)) Then

            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT is mandatory")

            ValidateDateDep_IT = False

            Exit Function

        End If

        

        If Not CheckDateddmmyyyy(DateDep_IT(i)) Then

            ValidateDateDep_IT = False

'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
             fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
            rangecells.item(i).value = ""
            Exit Function

        Else

            DateDep_IT(i) = Dformat(DateDep_IT(i), "")

            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-01-01") Then

                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  should not be less than 01/01/2023")
                rangecells.item(i).value = ""
                ValidateDateDep_IT = False

                Exit Function

            End If

            
'
'            If DateDep_IT(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgboxsmall "* Date of Deposit at Sr. No  " & i & " in Sheet IT can not be more than system date in Income Details"
'
'                            rangecells.item(i).value = ""
'
'                            ValidateDateDep_IT = False
'
'                Exit Function
'
'                     End If

        End If

    Next

End Function













Sub ttt()
Application.EnableEvents = True
End Sub
