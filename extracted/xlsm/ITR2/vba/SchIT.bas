Attribute VB_Name = "SchIT"
Option Explicit

Public msgbox_TDS2, msgbox_ESOP As String

Public ColCount2 As Long
Public ColCount2_1 As Long
Public ColCount2_2 As Long
Public ColCount2_3 As Long

Public DateCredit_TDS As Variant
Public BSR_TDS As Variant
Public SerialNum_TDS As Variant
Public TaxPaid3_TDS As Variant
Public end_IT, rngname_IT As Variant
Public msgValidateSheetIT As String


Sub Cmd_AddRows_ESOP_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("ESOPCurrAY||ESOP.Date||ESOP.AttributedTax")
ActiveCellRange = EfilingCommon.searchLastRow("ESOPCurrAY")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'Change- esop
Sub AddDiffRowsEsop(DiffRows As Long)
Dim vRows
Sheet46.Activate
EfilingCommon.DefinedgridNameRange = ("ESOPCurrAY||ESOP.Date||ESOP.AttributedTax")
ActiveCellRange = EfilingCommon.searchLastRow("ESOPCurrAY")
vRows = EfilingCommon.insertRowUnderSectionWithFormula(DiffRows)
End Sub
Sub Cmd_Validate_ESOP_Click()
ValidateSheet_ESOP
'fmsgboxStatus "Sheet ESOP is ok"
End Sub
Sub ValidateSheet_ESOP()
On Error Resume Next
subProcCaption = "Validating ESOP"
If Not ValidateESOP Then
fmsgboxsmall (msgbox_ESOP)
Sheet46.Activate 'Konda updated 25-03-2026
CloseMsg
Else
fmsgboxsmall "Sheet ESOP is ok"
End If

End Sub
Sub ValidateSheet_ESOP_1()
On Error Resume Next
subProcCaption = "Validating ESOP"
If Not ValidateESOP Then
fmsgboxsmall (msgbox_ESOP)
Sheet46.Activate 'Konda updated 25-03-2026
CloseMsg

End If

End Sub
Function ValidateESOP() As Boolean
ValidateESOP = True
msgbox_ESOP = ""

Dim Pan_flag As Boolean
Dim dpiit_flag As Boolean

dpiit_flag = False
Pan_flag = False

'Changed by Riyaz by Riyaz on 13/05/2025
'If Sheet46.Range("ESOP.SecType").value = "Partly Sold" Then
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value < Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType").value = "Fully Sold" Then
    'Uncommented by Bindu as per Sonam on 2nd May 25
    If (Sheet46.Range("ESOP.AttributedTaxTot").value <> Sheet46.Range("ESOP.TaxPrevAY").value) Then
     msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' "
    'Error msg updated by Bindu as per Sonam
    'msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
    ValidateESOP = False
    End If
End If

If Sheet46.Range("ESOP.SecType").value = "Not sold" Then
    If (Sheet46.Range("ESOP.AttributedTaxTot").value <> 0) Then
    
    'Error msg updated as per Sonam on 2nd May 25 by Bindu
    msgbox_ESOP = " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
    'msgbox_ESOP = " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
    ValidateESOP = False
    End If
End If


'Newly added by Bindu
If Sheet46.Range("ESOP.PAN").value <> "" Then
    If Not Part_A_General.CheckPAN(UCase(Trim(Sheet46.Range("ESOP.PAN").value))) Then
       
       msgbox_ESOP = "Invalid PAN in 'PAN of the employer being an eligible startup'. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."""
       Sheet46.Range("ESOP.PAN").Select
       Sheet46.Range("ESOP.PAN").value = ""
       ValidateESOP = False
       
    End If
End If

If Sheet46.Range("ESOP.DPIIT").value <> "" Then
   
   If Not Part_A_General.ChkDPIIT(UCase(Trim(Sheet46.Range("ESOP.DPIIT").value))) Then
             msgbox_ESOP = "Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"
             Sheet46.Range("ESOP.DPIIT").Select
             Sheet46.Range("ESOP.DPIIT").value = ""
    End If
    
    If Not checkallfieldspecialcharacter_80U_80DD(UCase(Trim(Sheet46.Range("ESOP.DPIIT").value))) Then
             msgbox_ESOP = "Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"
             Sheet46.Range("ESOP.DPIIT").Select
             Sheet46.Range("ESOP.DPIIT").value = ""

    End If
End If
'---Ended

''Newly added by Shrutika IPIP-69017
If Sheet46.Range("ESOP.BalTaxCF23_24").value > 0 Then

    If Sheet46.Range("ESOP.PAN").value = "" Then
    If Pan_flag = False Then
      msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
      Pan_flag = True
     End If
     ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
      If dpiit_flag = False Then
         msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
        dpiit_flag = True
      End If
         ValidateESOP = False
    End If


End If


If Len(Sheet46.Range("ESOP.DPIIT").value) > 0 Then

    If Sheet46.Range("ESOP.PAN").value = "" Then
    If Pan_flag = False Then
      msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
      Pan_flag = True
     End If
     ValidateESOP = False
    End If
    
End If

If Len(Sheet46.Range("ESOP.PAN").value) > 0 Then

    If Sheet46.Range("ESOP.DPIIT").value = "" Then
      If dpiit_flag = False Then
         msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
        dpiit_flag = True
      End If
         ValidateESOP = False
    End If
    
End If

'--Ends

If (Sheet46.Range("ESOP.SecType").value <> "(Select)" Or Sheet46.Range("ESOP.SecType").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing").Locked = True) Then


    If Sheet46.Range("ESOP.PAN").value = "" Then
    If Pan_flag = False Then
      msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
      Pan_flag = True
     End If
     ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
      If dpiit_flag = False Then
         msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
        dpiit_flag = True
      End If
         ValidateESOP = False
    End If



ElseIf (Sheet46.Range("ESOP.SecType").value = "(Select)" Or Sheet46.Range("ESOP.SecType").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing")) = "" Then

Else

'Newly added By Bindu
   If Sheet46.Range("ESOP.PAN").value = "" Then
     If Pan_flag = False Then
        msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
        Pan_flag = True
      End If
        ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
            dpiit_flag = True
            msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
            ValidateESOP = False
    End If


'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'Added by Shrutika(25/03/2026)SIT-114136
msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields in Schedule ESOP." & Chr(13)
ValidateESOP = False
End If

'Change.10.01.2023.104
'Changed by Riyaz on 13/05/2024
If Sheet46.Range("ESOP.SecType1").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee1").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot1").value > Sheet46.Range("ESOP.TaxPrevAY1").value) Or (Sheet46.Range("ESOP.AttributedTaxTot1").value = Sheet46.Range("ESOP.TaxPrevAY1").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType1").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee1").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot1").value < Sheet46.Range("ESOP.TaxPrevAY1").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType1").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee1").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot1").value > Sheet46.Range("ESOP.TaxPrevAY1").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType1").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot1").value <> Sheet46.Range("ESOP.TaxPrevAY1").value) Then

'msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
'Newly updated error msg as per BA Confirmation
msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "

ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType1").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot1").value <> 0) Then
'Newly updated by Bindu on 2nd May 25 as per Sonam
msgbox_ESOP = " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
'msgbox_ESOP = " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If

If (Sheet46.Range("ESOP.SecType1").value <> "(Select)" Or Sheet46.Range("ESOP.SecType1").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY1").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee1").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee1").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing1")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing1").Locked = True) Then

   If Sheet46.Range("ESOP.PAN").value = "" Then
     If Pan_flag = False Then
        Pan_flag = True
        msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
     End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If


ElseIf (Sheet46.Range("ESOP.SecType1").value = "(Select)" Or Sheet46.Range("ESOP.SecType1").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY1").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee1").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee1").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing1")) = "" Then

Else

'Newly added by Bindu
If Sheet46.Range("ESOP.PAN").value = "" Then
    If Pan_flag = False Then
        Pan_flag = True
        msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
    End If
        ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
         If dpiit_flag = False Then
            dpiit_flag = True
            msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
        End If
         ValidateESOP = False
    End If

'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'Added by Shrutika(25/03/2026)SIT-114136
msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields in Schedule ESOP." & Chr(13)

ValidateESOP = False
End If

'Chandru
'Changed by Riyaz on 13/05/2025
If Sheet46.Range("ESOP.SecType2").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee2").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value > Sheet46.Range("ESOP.TaxPrevAY2").value) Or (Sheet46.Range("ESOP.AttributedTaxTot2").value = Sheet46.Range("ESOP.TaxPrevAY2").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType2").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee2").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value < Sheet46.Range("ESOP.TaxPrevAY2").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType2").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee2").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value > Sheet46.Range("ESOP.TaxPrevAY2").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType2").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value <> Sheet46.Range("ESOP.TaxPrevAY2").value) Then
'msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
'Newly updated error msg as per BA Confirmation
msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully  sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "


ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType2").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value <> 0) Then
'Newly updated by Bindu on 2nd May 25 as per Sonam
msgbox_ESOP = " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
'msgbox_ESOP = " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If

If (Sheet46.Range("ESOP.SecType2").value <> "(Select)" Or Sheet46.Range("ESOP.SecType2").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY2").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee2").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee2").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing2")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing2").Locked = True) Then

    If Sheet46.Range("ESOP.PAN").value = "" Then
      If Pan_flag = False Then
          Pan_flag = True
          msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
       End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If



ElseIf (Sheet46.Range("ESOP.SecType2").value = "(Select)" Or Sheet46.Range("ESOP.SecType2").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY2").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee2").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee2").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing2")) = "" Then


Else

'Newly added by Bindu
    If Sheet46.Range("ESOP.PAN").value = "" Then
      If Pan_flag = False Then
          Pan_flag = True
          msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
       End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If


'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'Added by Shrutika(25/03/2026)SIT-114136
msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields in Schedule ESOP." & Chr(13)

ValidateESOP = False
End If

'Newly added by Bindu 17th April 25
'Changed by Riyaz on 13/05/2025
If Sheet46.Range("ESOP.SecType3").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee3").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value > Sheet46.Range("ESOP.TaxPrevAY3").value) Or (Sheet46.Range("ESOP.AttributedTaxTot3").value = Sheet46.Range("ESOP.TaxPrevAY3").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType3").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee3").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value < Sheet46.Range("ESOP.TaxPrevAY3").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP = False
End If
End If

'Added by Riyaz on13/03/2025
If Sheet46.Range("ESOP.SecType3").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee3").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value > Sheet46.Range("ESOP.TaxPrevAY3").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType3").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value <> Sheet46.Range("ESOP.TaxPrevAY3").value) Then
'msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
'Newly updated error msg as per BA Confirmation
msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "


ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType3").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value <> 0) Then
'Newly updated by Bindu on 2nd May 25 as per Sonam
msgbox_ESOP = " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
'msgbox_ESOP = " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If

If (Sheet46.Range("ESOP.SecType3").value <> "(Select)" Or Sheet46.Range("ESOP.SecType3").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY3").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee3").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee3").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing3")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing3").Locked = True) Then

    If Sheet46.Range("ESOP.PAN").value = "" Then
      If Pan_flag = False Then
          Pan_flag = True
          msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
       End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If



ElseIf (Sheet46.Range("ESOP.SecType3").value = "(Select)" Or Sheet46.Range("ESOP.SecType3").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY3").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee3").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee3").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing3")) = "" Then


Else

'Newly added by Bindu
    If Sheet46.Range("ESOP.PAN").value = "" Then
      If Pan_flag = False Then
          Pan_flag = True
          msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
       End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If


'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'Added by Shrutika(25/03/2026)SIT-114136
msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields in Schedule ESOP." & Chr(13)

ValidateESOP = False
End If

'added by Chetan C M on 19/01/2026 for AY 2026-27
'start--
If Sheet46.Range("ESOP.SecType4").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee4").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value > Sheet46.Range("ESOP.TaxPrevAY4").value) Or (Sheet46.Range("ESOP.AttributedTaxTot4").value = Sheet46.Range("ESOP.TaxPrevAY4").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType4").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee4").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value < Sheet46.Range("ESOP.TaxPrevAY4").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType4").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee4").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value > Sheet46.Range("ESOP.TaxPrevAY4").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType4").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value <> Sheet46.Range("ESOP.TaxPrevAY4").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "


ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType4").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value <> 0) Then
msgbox_ESOP = " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If

If (Sheet46.Range("ESOP.SecType4").value <> "(Select)" Or Sheet46.Range("ESOP.SecType4").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY4").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee4").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee4").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing4")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing4").Locked = True) Then

    If Sheet46.Range("ESOP.PAN").value = "" Then
      If Pan_flag = False Then
          Pan_flag = True
          msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
       End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If



ElseIf (Sheet46.Range("ESOP.SecType4").value = "(Select)" Or Sheet46.Range("ESOP.SecType4").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY4").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee4").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee4").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing4")) = "" Then


Else

    If Sheet46.Range("ESOP.PAN").value = "" Then
      If Pan_flag = False Then
          Pan_flag = True
          msgbox_ESOP = msgbox_ESOP + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
       End If
      ValidateESOP = False
    End If
    
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
        If dpiit_flag = False Then
           dpiit_flag = True
           msgbox_ESOP = msgbox_ESOP + "Please enter the Start up recognition number of employer" & Chr(13)
         End If
         ValidateESOP = False
    End If

'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)

'Added by Shrutika(25/03/2026)SIT-114136
msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields in Schedule ESOP." & Chr(13)
ValidateESOP = False
End If
'--end

End Function


Sub Cmd_AddRows_IT_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("IT.BSRCode||IT.DateDep||IT.SrlNoOfChaln||IT.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT||FormulaOfExSAT1")
ActiveCellRange = EfilingCommon.searchLastRow("IT.BSRCode")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_Validate_IT_Click()
ValidateSheet_IT
fmsgboxStatus "Sheet IT is ok"
End Sub


Sub ValidateSheet_IT()
subProcCaption = "Validating IT"
If Not ValidateAdvanceTax Then
fmsgboxsmall (msgbox_TDS2)
CloseMsg
End If
End Sub


Function ValidateAdvanceTax() As Boolean
ValidateAdvanceTax = True
msgbox_TDS2 = "IT" & Chr(10)
    setTableInfo_Grid3
    setTableInfo1_Grid3
    setTableInfo2_Grid3
    setTableInfo3_Grid3
    If Len((Sheet25.Range("IT.BSRCode").item(1).value) > 0) Then
        If Not ValidateBSR_TDS Then ValidateAdvanceTax = False
        If Not ValidateDateCreditToGovt Then ValidateAdvanceTax = False
        If Not ValidateSerialNum Then ValidateAdvanceTax = False
        If Not ValidateTaxPaid Then ValidateAdvanceTax = False
    End If
    
    
'    If ((ColCount2 <> ColCount2_1) Or (ColCount2 <> ColCount2_2) Or (ColCount2 <> ColCount2_3)) Then
'    msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Details schedule IT"
'    ValidateAdvanceTax = False
'    End If
End Function

Function ValidateBSR_TDS() As Boolean
ValidateBSR_TDS = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.BSRCode").Cells
Dim i As Long

'Malli added
ColCount2 = Application.WorksheetFunction.Max(ColCount2, ColCount2_1, ColCount2_2, ColCount2_3)
'----------------------

ReDim BSR_TDS(ColCount2)
noOfProcessSub = ColCount2
For i = 1 To ColCount2
BSR_TDS(i) = rangecells.item(i).value
If Len(BSR_TDS(i)) = 0 Then
End If

If Not chkCompulsory(BSR_TDS(i)) Then
   ' msgbox_TDS2 = msgbox_TDS2 + "* Please enter the BSR Code at Sr.NO " & i & "in schedule IT" & Chr(13)
   'Added by Aavula
   msgbox_TDS2 = msgbox_TDS2 + "* Please enter the BSR code in Schedule IT at Sr.NO " & i & "" & Chr(13)
    ValidateBSR_TDS = False
    Exit Function
End If

'Added by Aavula
 If Not Validate_BSRCODE(BSR_TDS(i)) Then
         'msgbox_TDS2 = msgbox_TDS2 + "Please enter a valid 7 digit BSR Code in Schedule IT at Sr. No  " & i & " " & Chr(13)
         'Malli_V0.4V/18/03/2025
         msgbox_TDS2 = msgbox_TDS2 + "* ""Invalid BSR"" in Schedule IT at Sr. No  " & i & " " & Chr(13)
         ValidateBSR_TDS = False
        Exit Function
     End If
UpdateProgressBar
Next
End Function
Function Validate_BSRCODE(BsrCode As Variant) As Boolean
On Error Resume Next


    Validate_BSRCODE = True
    If Len(BsrCode) > 0 Then
        If Not IsNumeric(Mid(BsrCode, 1, 1)) Then
            Validate_BSRCODE = False
            Exit Function
        End If
        If Not IsNumeric(Mid(BsrCode, 2, 1)) Then
            Validate_BSRCODE = False
            Exit Function
        End If
        If Not IsNumeric(Mid(BsrCode, 3, 1)) Then
            Validate_BSRCODE = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter_BsRCode(Mid(BsrCode, 4, 4)) Then
            Validate_BSRCODE = False
            Exit Function
        End If

    End If
End Function

Function checkfieldspecialcharacter_BsRCode(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_BsRCode = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "_", "|", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_BsRCode = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateDateCreditToGovt() As Boolean
ValidateDateCreditToGovt = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.DateDep").Cells
Dim i As Long
ReDim DateCredit_TDS(ColCount2)
For i = 1 To ColCount2
DateCredit_TDS(i) = rangecells.item(i).value
If Len(DateCredit_TDS(i)) = 0 Then
End If

'rangecells.item(i) = Replace(DateCredit_TDS(i), "-", "/")

If Not chkCompulsory(DateCredit_TDS(i)) Then

'Added by Aavula
    'msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Date of Deposit into Govt Account at Sr.NO " & i & " in schedule IT" & Chr(13)
     msgbox_TDS2 = msgbox_TDS2 + "* Please enter date of deposit in Schedule IT at Sr.NO " & i & "" & Chr(13)
    
    ValidateDateCreditToGovt = False
    Exit Function
End If

'Malli----------30/10/2024
'AY_2024_25 OLD
'If DateCredit_TDS(i) <> "" Then
'AY_2025_26 SIT-88948
If DateCredit_TDS(i) <> "" And (Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21" Then
'If ChkMinInclusiveDate(Dformat(DateCredit_TDS(i), "yyyy-mm-dd"), "2025-01-01") Then
'If ChkMinInclusiveDate(Dformat(DateCredit_TDS(i), "yyyy-mm-dd"), "2026-01-01") Then  'Added by Shrutika 18/11/2025 SIT-103777

If ChkMinInclusiveDate(Dformat(DateCredit_TDS(i), "yyyy-mm-dd"), "2027-01-01") Then  'Added by Naresh 04/03/2026 SIT-112893
'If ChkMinInclusiveDate(Dformat(DateCredit_TDS(i), "yyyy-mm-dd"), "2025-01-01") Then

      'msgbox_TDS2 = msgbox_TDS2 + "* ""Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.""" & Chr(13)
     'Added by Shrutika(28/02/2025)
    ' msgbox_TDS2 = msgbox_TDS2 + "* ""Date of deposit can't be after 31-12-2025 in schedule IT. Please enter payments made after 31st December 2025 in schedule IT 1.""" & Chr(13)
      'Malli----AY_2025_26
'      msgbox_TDS2 = msgbox_TDS2 + "* ""If filing section is selected as 139(8A), then date of deposit cant be after 31.12.2025 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""" & Chr(13)
       msgbox_TDS2 = msgbox_TDS2 + "* ""If filing section is selected as 139(8A), then date of deposit cant be after 31.12.2026 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""" & Chr(13)
      
      ValidateDateCreditToGovt = False
      rangecells.item(i).value = ""
      Exit Function
End If
End If


Next
End Function


Function ValidateSerialNum() As Boolean
ValidateSerialNum = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.SrlNoOfChaln").Cells
Dim i As Long
ReDim SerialNum_TDS(ColCount2)
For i = 1 To ColCount2
SerialNum_TDS(i) = rangecells.item(i).value
If Len(SerialNum_TDS(i)) = 0 Then
End If

If Not chkCompulsory(SerialNum_TDS(i)) Then

    'msgbox_TDS2 = msgbox_TDS2 + "* Please enter 5 digit serial number of challan at Sr.NO " & i & " in schedule IT" & Chr(13)
    'Added by Aavula
    msgbox_TDS2 = msgbox_TDS2 + "* Please enter serial number of challan in Schedule IT at Sr.NO " & i & "" & Chr(13)
  
    ValidateSerialNum = False
    Exit Function
End If


If Not IsNumeric(SerialNum_TDS(i)) Then
    msgbox_TDS2 = msgbox_TDS2 + "* Serial Number of Challan at Sr.NO " & i & " can allow only Numeric Values in schedule IT" & Chr(13)
    ValidateSerialNum = False
    Exit Function
End If
Next
End Function

Function ValidateTaxPaid() As Boolean
ValidateTaxPaid = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.Amt").Cells
Dim i As Long
ReDim TaxPaid3_TDS(ColCount2)
For i = 1 To ColCount2
TaxPaid3_TDS(i) = rangecells.item(i).value
If Len(TaxPaid3_TDS(i)) = 0 Then
End If

If Not chkCompulsory(TaxPaid3_TDS(i)) Then
    'msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Tax Paid Amount at Sr.NO " & i & " in schedule IT" & Chr(13)
    
    'Added by Aavula
     msgbox_TDS2 = msgbox_TDS2 + "* Please enter amount in Schedule IT at Sr.NO " & i & "" & Chr(13)
 
    
    
    ValidateTaxPaid = False
    Exit Function
End If
Next
End Function

Sub setTableInfo_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.BSRCode").count
    Set rangecells = Sheet25.Range("IT.BSRCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2 = ccount
End Sub

Sub setTableInfo1_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.DateDep").count
    Set rangecells = Sheet25.Range("IT.DateDep").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_1 = ccount
End Sub


Sub setTableInfo2_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.SrlNoOfChaln").count
    Set rangecells = Sheet25.Range("IT.SrlNoOfChaln").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_2 = ccount
End Sub


Sub setTableInfo3_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.Amt").count
    Set rangecells = Sheet25.Range("IT.Amt").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_3 = ccount
End Sub

Function msgbox_IT(strmsg As String) As String
     msgValidateSheetIT = msgValidateSheetIT & strmsg & Chr(13)
End Function
'Added by Riyaz on 16/06/2025 for 2025-26 New Dev
Sub ImportScheduleITCSV()
Dim filePath, text As String

Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    msgbox_IT ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        .Filters.add "CSV File", "*.csv", 1
        If .Show = True Then
            filePath = .SelectedItems.item(1)
        Else
            Exit Sub
        End If
        On Error Resume Next
    End With
    
    Open filePath For Input As #1
    text = Input$(LOF(1), 1)
    Close #1
    
    importITCSVFunc (text)
End If
End Sub
'Added by Riyaz on 16/06/2025 for 2025-26 New Dev
    Sub importITCSVFunc(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    n = 0
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount As Variant
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(3, 1 To n)

            TwoDArray(0, n) = TmpAr(0)
            TwoDArray(1, n) = TmpAr(1)
            TwoDArray(2, n) = TmpAr(2)
            TwoDArray(3, n) = TmpAr(3)
            
        End If
    Next i
    
    Dim BSRCodeColNo, DateDepColNo, SrlNoOfChalnColNo, AmtColNo As Variant
    
    BSRCodeColNo = Sheet25.Range("IT.BSRCode").Column
    DateDepColNo = Sheet25.Range("IT.DateDep").Column
    SrlNoOfChalnColNo = Sheet25.Range("IT.SrlNoOfChaln").Column
    AmtColNo = Sheet25.Range("IT.Amt").Column
    
    
    TotalExRow = Sheet25.Range("IT.BSRCode").Rows.count
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet25.Range("IT.BSRCode").Locked = False Then
            Sheet25.Range("IT.BSRCode").ClearContents
        End If
        If Sheet25.Range("IT.DateDep").Locked = False Then
            Sheet25.Range("IT.DateDep").ClearContents
        End If
        If Sheet25.Range("IT.SrlNoOfChaln").Locked = False Then
            Sheet25.Range("IT.SrlNoOfChaln").ClearContents
        End If
        If Sheet25.Range("IT.Amt").Locked = False Then
            Sheet25.Range("IT.Amt").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_ScheduleIT (TotalDiffRow)
    End If
    rowcount = getRowNo(Sheet25.Range("IT.BSRCode").name)

    For i = 1 To n
        If Sheet25.Cells(rowcount, BSRCodeColNo).Locked = False Then
            Sheet25.Cells(rowcount, BSRCodeColNo).value = TwoDArray(0, i)
        End If
        If Sheet25.Cells(rowcount, DateDepColNo).Locked = False Then
            Sheet25.Cells(rowcount, DateDepColNo).value = TwoDArray(1, i)
        End If
        If Sheet25.Cells(rowcount, SrlNoOfChalnColNo).Locked = False Then
            Sheet25.Cells(rowcount, SrlNoOfChalnColNo).value = TwoDArray(2, i)
        End If
        If Sheet25.Cells(rowcount, AmtColNo).Locked = False Then
            Sheet25.Cells(rowcount, AmtColNo).value = TwoDArray(3, i)
        End If

        rowcount = rowcount + 1
        
        
    Next i
    
End Sub

