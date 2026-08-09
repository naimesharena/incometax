Attribute VB_Name = "Module3"
'added by Chetan C M for AY 2025-26
'start--
Sub Addrows24b(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheets("House Property").Activate

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value


        EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
        ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.24b" & te)

        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next



'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
        newfrmsize = Sheet8.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet8.Range("NumRowsPropertyBlock").value = newfrmsize
    Next


endline:

'----------------Lock Password-------------------START---
   Sheet8.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'--end

