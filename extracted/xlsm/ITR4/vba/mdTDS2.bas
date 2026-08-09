Attribute VB_Name = "mdTDS2"
''Option Explicit
''
''Public end_TDS2, end_TDS22, end_TDS2inner, end_TDS23, end_TDS24, end_TDS25, end_TDS26, end_TDS27, end_TDS28, end_TDS29, end_TDS210, end_TDS211, end_TDS212, end_TDS213, end_TDS214, end_TDS215, end_TDS216, end_TDS217 As Variant
''Public end_TDS2b, end_TDS2b2, end_TDS2b3, end_TDS2b4, end_TDS2b5, end_TDS2b6, end_TDS2b7, end_TDS2b8, end_TDS2b9, end_TDS2b10, end_TDS2b11, end_TDS2b12, end_TDS2b13 As Variant
''Public rngname_TDS2, rngname_TDS2b As Variant
''Public i As Long
''Public MsgTDS2, MsgTDS2b As Variant
'''Public TDSApportionedCheck As Variant
'''Public TDSClaimCheck As Variant
''Public TDSApportionedCheck As Variant
''Public TDSClaimCheck As Variant
''Public TAN_TDS2, TAN2_TDS2 As Variant
''Public EmployerOrDeductorOrCollecterName_TDS2 As Variant
''Public UTN_TDS2 As Variant
''Public DeductedYear_TDS2 As Variant
''Public DetailsOfReceipts_TDS2 As Variant
''Public TaxDeducted26AS_TDS2 As Variant
''
''Public AmtOfferedSelf_TDS2 As Variant
''Public TDSClaimedInOwnHands_TDS2 As Variant
''
''Public TDSApportOtherHands_TDS2 As Variant
''Public IncomeOfferedThisYear_TDS2 As Variant
''Public TDSClaimedThisYear_TDS2 As Variant
''Public IncomeCarriedFwd_TDS2 As Variant
''Public TDSCarriedFwd_TDS2 As Variant
''Public HeadOfIncome_TDS2 As Variant
''Public ReasonApportTDS_TDS2 As Variant
''Public PANOfOtherPerson_TDS2 As Variant
''Public PercentOfShare_TDS2 As Variant
''
''
''Public TAN_TDS2b As Variant
''Public EmployerOrDeductorOrCollecterName_TDS2b As Variant
''Public UTN_TDS2b As Variant
''Public DeductedYear_TDS2b As Variant
''Public IncomeApportioned_TDS2b As Variant
''Public TDSApportioned_TDS2b As Variant
''Public IncomeOfferedThisYear_TDS2b As Variant
''Public TDSClaimedThisYear_TDS2b As Variant
''Public IncomeCarriedFwd_TDS2b As Variant
''Public TDSCarriedFwd_TDS2b As Variant
''Public HeadOfIncome_TDS2b As Variant
''Public ReasonClaimOwnHands_TDS2b As Variant
''Public TDSDeductedPAN_TDS2b As Variant
''Sub CmdTDS2ii()
''    Sheets("TDS").Activate
''    mdCommon.DefinedgridNameRange = "TDS2ii.SlNo||TDSoth2.PAN||TDSoth2.EmployerOrDeductorOrCollecterName||TDSoth2.UTN||TDSoth2.DeductedYear||TDSoth2.AmountSubjectDed||TDSoth2.TD_OwnHands||TDSoth2.TD_SIncome||TDSoth2.TD_TDS||TDSoth2.ClaimOutOfTotTDSOnAmtPaid||TDSoth2.TD_PAN||TDS2ii.LOOKUP"
''    ActiveCellRange = mdCommon.searchLastRow("TDSoth2.PAN")
''    mdCommon.insertRowUnderSectionWithFormula
''End Sub
''
''
''
'
''Sub ValidateTDS2_Click()
''    ValidateSheetSchTDS2
''    MsgBox "Sheet is ok", vbOKOnly, "ITR-2"
''End Sub
''Sub ValidateSheetSchTDS2()
''   If Not ValidateTDSApportionedCheck Then
''        Sheet16.Activate
''        MsgBox (MsgTDS2)
''        CloseMsg
''    End If
''
''    If Not ValidateTDSClaimCheck Then
''        Sheet16.Activate
''        MsgBox (MsgTDS2)
''        CloseMsg
''    End If
''
''    If Not ValidateSheetTDS2 Then
''        Sheet16.Activate
''        MsgBox (MsgTDS2)
''        CloseMsg
''    End If
''
''    If Not ValidatesheetTDS2b Then
''        Sheet16.Activate
''        MsgBox MsgTDS2b, vbOKOnly, "ITR-2"
''        CloseMsg
''    End If
''End Sub
''
''Function ValidateTDSApportionedCheck() As Boolean
''ValidateTDSApportionedCheck = True
''TDSApportionedCheck = Sheet16.Range("TDS2.DrpA").value
''
''If isdropdownblank(TDSApportionedCheck) Then
''TDSApportionedCheck = ""
''End If
''
''If Mid((TDSApportionedCheck), 1, 1) = "(" Or TDSApportionedCheck = "" Then
''    MsgTDS2 = MsgTDS2 + "Do you have TDS reported in your 26 AS check has to be filled" & Chr(13)
''    ValidateTDSApportionedCheck = False
''Exit Function
''End If
''
''End Function
''
''Function ValidateTDSClaimCheck() As Boolean
''ValidateTDSClaimCheck = True
''TDSClaimCheck = Sheet16.Range("TDS2.DrpB").value
''
''If isdropdownblank(TDSClaimCheck) Then
''TDSClaimCheck = ""
''End If
''
''If Mid((TDSClaimCheck), 1, 1) = "(" Or TDSClaimCheck = "" Then
''    MsgTDS2 = MsgTDS2 + "Are you claiming TDS reported under other person's 26AS check has to be filled" & Chr(13)
''    ValidateTDSClaimCheck = False
''Exit Function
''End If
''
''End Function
''
''Sub UNLOCKTDS2()
''    Application.ScreenUpdating = False
''    Sheet16.Unprotect getmsgstate
''
''
''    Dim rngname_5_TDS2 As Variant
''    rngname_5_TDS2 = "TDS2.Reason||TDS2.DRPTAN||TDS2.OthrTds||TDSoth.TD_PAN||Column6_value"
''    UNLOCKRANGETDS2 (rngname_5_TDS2)
''
''    Sheet16.Range("TDS2.PANDetails").EntireRow.Hidden = False
''    Sheet16.Shapes("Button 58").Visible = msoTrue
''
''   ActiveSheet.Shapes("Button 58").Left = Range("SchTDS2.Button2").Left
''    ActiveSheet.Shapes("Button 58").Top = Range("SchTDS2.Button2").Top
''    'TDS2.PANDetails
''
''    Sheet16.Protect getmsgstate
''    Application.ScreenUpdating = True
''End Sub
''
''
''Sub LOCKTDS2()
''    Application.ScreenUpdating = False
''
''    Sheet16.Unprotect getmsgstate
''
''    Dim rngname_5_TDS2 As Variant
''    rngname_5_TDS2 = "TDS2.Reason||TDS2.DRPTAN||TDS2.OthrTds||TDSoth.TD_PAN||Column6_value"
''    LOCKRANGETDS2 (rngname_5_TDS2)
''
''Sheet16.Range("TDS2.PANDetails").EntireRow.Hidden = True
''Sheet16.Shapes("Button 58").Visible = msoFalse
''
''ActiveSheet.Shapes("Button 58").Left = Range("SchTDS2.Button2").Left
''ActiveSheet.Shapes("Button 58").Top = Range("SchTDS2.Button2").Top
''
''
''   Sheet16.Protect getmsgstate
''
''    Application.ScreenUpdating = True
''End Sub
''Sub LOCKRANGETDS2(rangenamestring As Variant)
''Dim i As Long
''Dim chCell As Range
''    rangenamestring = Split(rangenamestring, "||")
''        For i = 0 To UBound(rangenamestring)
''            For Each chCell In Sheet16.Range(rangenamestring(i))
''            chCell.Locked = True
''            chCell.Interior.Color = (&HD8D8D8)
''            chCell.ClearContents
''            Next
''        Next
''End Sub
''
''Sub UNLOCKRANGETDS2(rangenamestring As Variant)
''Dim i As Long
''Dim chCell As Range
''Dim addr1, addr2 As Variant
''
''
''    rangenamestring = Split(rangenamestring, "||")
''        For i = 0 To UBound(rangenamestring)
''            For Each chCell In Sheet16.Range(rangenamestring(i))
''             If rangenamestring(i) = "TDS2.Reason" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''            addr1 = Replace(chCell.AddressLocal, "L", "H")
''           addr2 = Replace(chCell.AddressLocal, "L", "K")
''            If Sheet16.Range(addr1).value > 0 Then
''            If Sheet16.Range(addr1).value = Sheet16.Range(addr2).value Then
''            Range(chCell.AddressLocal) = Replace(chCell.AddressLocal, "K", "L")
''            chCell.value = ""
''            chCell.Locked = True
''            chCell.Interior.Color = "&HD8D8D8"
''            End If
''            End If
''            ElseIf rangenamestring(i) = "TDSoth2b.DeductedYear" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''            ElseIf rangenamestring(i) = "TDSoth2b.HeadOfIncome" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''            ElseIf rangenamestring(i) = "TDSoth2b.Reason" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''            Else
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.ClearContents
''            End If
''            Next
''        Next
''End Sub
''
''Sub UNLOCKTDS2b()
''
''Application.ScreenUpdating = False
''  Application.EnableEvents = False
'''----------------Unlock Password-------------------START---
''
''   Sheet16.Unprotect getmsgstate
'''----------------Unlock Password-------------------END-----
''Dim rngname_TDS2b As Variant
''
''rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed"
''UNLOCKRANGETDS2 (rngname_TDS2b)
''
''
''
''Sheet16.Range("TDS2b.Hide").EntireRow.Hidden = False
''Sheet16.Shapes("TDS2b.button").Visible = msoTrue
''
''ActiveSheet.Shapes("TDS2b.button").Left = Range("SchTDS2.Button3").Left
''ActiveSheet.Shapes("TDS2b.button").Top = Range("SchTDS2.Button3").Top
''
''
'''----------------Lock Password-------------------START---
''   Sheet16.Protect getmsgstate
'''----------------Lock Password-------------------END-----
''
'' Application.EnableEvents = True
'' Application.ScreenUpdating = True
''End Sub
''Sub LOCKTDS2b()
''
''
''Application.ScreenUpdating = False
'' Application.EnableEvents = False
'''----------------Unlock Password-------------------START---
''
''   Sheet16.Unprotect getmsgstate
'''----------------Unlock Password-------------------END-----
''
''Dim rngname_TDS2b As Variant
''
''rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed"
''LOCKRANGETDS2 (rngname_TDS2b)
''
''
''
''Sheet16.Range("TDS2b.Hide").EntireRow.Hidden = True
''Sheet16.Shapes("TDS2b.button").Visible = msoFalse
''
''ActiveSheet.Shapes("TDS2b.button").Left = Range("SchTDS2.Button3").Left
''    ActiveSheet.Shapes("TDS2b.button").Top = Range("SchTDS2.Button3").Top
''
''
'''----------------Lock Password-------------------START---
''   Sheet16.Protect getmsgstate
'''----------------Lock Password-------------------END-----
''
'' Application.EnableEvents = True
'' Application.ScreenUpdating = True
''End Sub
''
''
''Sub setTblinfo_TDS2()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth.TAN").Count
'' Set rangecells = Sheet16.Range("TDSoth.TAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2 = ccount
'' rngname_TDS2 = "TDS2.SlNo||TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDSoth.UTN||TDSoth.DeductedYear||TDSoth.AmountPaid||TDSoth.TD_OwnHands||TDS2.SelfTdsClaimed||TDS2.HeadOfIncome||TDS2.TdsClaimedOwn||TDS2.Reason||TDS2.OthrTds||TDSoth.TD_PAN||TDS2.PercShare||TDS2.IncomeOffered||TDS2.IncomeClaimed||TDS2.IncBalCF||TDS2.IncBalCF"
'' End Sub
''
'' Sub setTblinfo_TDS22()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth.EmployerOrDeductorOrCollecterName").Count
'' Set rangecells = Sheet16.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS22 = ccount
'' End Sub
''
''  Sub setTblinfo_TDS23()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth.UTN").Count
'' Set rangecells = Sheet16.Range("TDSoth.UTN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS23 = ccount
'' End Sub
''  Sub setTblinfo_TDS24()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth.DeductedYear").Count
'' Set rangecells = Sheet16.Range("TDSoth.DeductedYear").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS24 = ccount
'' End Sub
''   Sub setTblinfo_TDS25()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth.AmountPaid").Count
'' Set rangecells = Sheet16.Range("TDSoth.AmountPaid").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS25 = ccount
'' End Sub
''
'' Sub setTblinfo_TDS26()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2.TD_OwnHands").Count
'' Set rangecells = Sheet16.Range("TDSoth2.TD_OwnHands").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS26 = ccount
'' End Sub
''
''  Sub setTblinfo_TDS27()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.SelfTdsClaimed").Count
'' Set rangecells = Sheet16.Range("TDS2.SelfTdsClaimed").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS27 = ccount
'' End Sub
''
''
''  Sub setTblinfo_TDS28()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.HeadOfIncome").Count
'' Set rangecells = Sheet16.Range("TDS2.HeadOfIncome").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS28 = ccount
'' End Sub
''   Sub setTblinfo_TDS29()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.TdsClaimedOwn").Count
'' Set rangecells = Sheet16.Range("TDS2.TdsClaimedOwn").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS29 = ccount
'' End Sub
''   Sub setTblinfo_TDS210()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.Reason").Count
'' Set rangecells = Sheet16.Range("TDS2.Reason").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS210 = ccount
'' End Sub
''  Sub setTblinfo_TDS2inner()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.DRPTAN").Count
'' Set rangecells = Sheet16.Range("TDS2.DRPTAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2inner = ccount
'' End Sub
''
''
''
''   Sub setTblinfo_TDS211()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.OthrTds").Count
'' Set rangecells = Sheet16.Range("TDS2.OthrTds").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS211 = ccount
'' End Sub
''   Sub setTblinfo_TDS212()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth.TD_PAN").Count
'' Set rangecells = Sheet16.Range("TDSoth.TD_PAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS212 = ccount
'' End Sub
''   Sub setTblinfo_TDS213()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.PercShare").Count
'' Set rangecells = Sheet16.Range("TDS2.PercShare").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS213 = ccount
'' End Sub
''   Sub setTblinfo_TDS214()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.IncomeOffered").Count
'' Set rangecells = Sheet16.Range("TDS2.IncomeOffered").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS214 = ccount
'' End Sub
''   Sub setTblinfo_TDS215()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.IncomeClaimed").Count
'' Set rangecells = Sheet16.Range("TDS2.IncomeClaimed").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS215 = ccount
'' End Sub
''
''   Sub setTblinfo_TDS216()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.IncBalCF").Count
'' Set rangecells = Sheet16.Range("TDS2.IncBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS216 = ccount
'' End Sub
''   Sub setTblinfo_TDS217()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDS2.TDSBalCF").Count
'' Set rangecells = Sheet16.Range("TDS2.TDSBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS217 = ccount
'' End Sub
'' Sub setTblinfo_TDS2b()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.TAN").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.TAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''  Sub setTblinfo_TDS2b2()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.EmployerOrDeductorOrCollecterName").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.EmployerOrDeductorOrCollecterName").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b2 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''   Sub setTblinfo_TDS2b3()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.UTN").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.UTN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b3 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''    Sub setTblinfo_TDS2b4()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.DeductedYear").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.DeductedYear").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b4 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''     Sub setTblinfo_TDS2b5()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.Income").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.Income").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b5 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
'' Sub setTblinfo_TDS2b6()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.TDS").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.TDS").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b6 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''  Sub setTblinfo_TDS2b7()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.PAN").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.PAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b7 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
'' Sub setTblinfo_TDS2b8()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.HeadOfIncome").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.HeadOfIncome").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b8 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''  Sub setTblinfo_TDS2b9()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.Reason").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.Reason").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b9 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
'' Sub setTblinfo_TDS2b10()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.IncomeOffered").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.IncomeOffered").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b10 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''
'' Sub setTblinfo_TDS2b11()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.TdsClaimed").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.TdsClaimed").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b11 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''  Sub setTblinfo_TDS2b12()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.IncBalCF").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.IncBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b12 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''   Sub setTblinfo_TDS2b13()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet16.Range("TDSoth2b.TdsBalCF").Count
'' Set rangecells = Sheet16.Range("TDSoth2b.TdsBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS2b13 = ccount
'' rngname_TDS2b = "TDSoth2b.TAN||TDSoth2b.EmployerOrDeductorOrCollecterName||TDSoth2b.UTN||TDSoth2b.DeductedYear||TDSoth2b.Income||TDSoth2b.TDS||TDSoth2b.PAN||TDSoth2b.HeadOfIncome||TDSoth2b.Reason||TDSoth2b.IncomeOffered||TDSoth2b.TdsClaimed||TDSoth2b.IncBalCF||TDSoth2b.TdsBalCF"
'' End Sub
''
''
'' Function ValidateSheetTDS2() As Boolean
''On Error Resume Next
''Dim cellcount, intcntr, rangname, name As Variant
''Dim SumTds, SumIft, tdssum, tdssum1, TDSvalue As Variant
''
''Dim k, j As Long
''
''ValidateSheetTDS2 = True
''    MsgTDS2 = "Schedule TDS2a." & Chr(13)
''    subProcCaption = "Validating TDS2"
''    setTblinfo_TDS2
''    noOfProcessSub = end_TDS2
''
''    If Len(Sheet16.Range("TDS2a.Total").value) > 14 Then
''       MsgTDS2 = MsgTDS2 & " Total of TDS2b should not excced 14 digits " & Chr(13)
''       ValidateSheetTDS2 = False
''       Exit Function
''    End If
''
''For k = 1 To end_TDS2
''
''        If Not (Sheet16.Range("TDSoth.TAN").Cells.item(k) <> "" Or Sheet16.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet16.Range("TDSoth.DeductedYear").Cells.item(k)) Or Sheet16.Range("TDSoth.AmountPaid").Cells.item(k) <> "" Or Sheet16.Range("TDSoth2.TD_OwnHands").Cells.item(k) <> "" Or Sheet16.Range("TDS2.IncomeClaimed").Cells.item(k) <> "") Then
''            MsgTDS2 = MsgTDS2 & " At Sr. No " & k & " Please fill all the Mandatory Fields " & Chr(13)
''            ValidateSheetTDS2 = False
''            Exit Function
''        End If
''
''    UpdateProgressBar
''Next
''
''
''If end_TDS2 > 0 Then GoTo xyz
''
''xyz:
''    If Not ValidateTAN_TDS2() Then ValidateSheetTDS2 = False
''    If Not ValidateEmployerOrDeductorOrCollecterName_TDS2() Then ValidateSheetTDS2 = False
''    If Not ValidateUTN_TDS2() Then ValidateSheetTDS2 = False
''    If Not ValidateDeductedYear_TDS2() Then ValidateSheetTDS2 = False
''    If Not ValidateDetailsOfReceipts_TDS2() Then ValidateSheetTDS2 = False
''        If Not ValidateTaxDeducted26AS_TDS2() Then ValidateSheetTDS2 = False
''        If Not ValidateAmtOfferedSelf_TDS2() Then ValidateSheetTDS2 = False
''        If Not ValidateTDSClaimedInOwnHands_TDS2() Then ValidateSheetTDS2 = False
''        If Not ValidateTDSApportOtherHands_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateIncomeOfferedThisYear_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateTDSClaimedThisYear_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateIncomeCarriedFwd_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateTDSCarriedFwd_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateHeadOfIncome_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateReasonApportTDS_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidatePANOfOtherPerson_TDS2 Then ValidateSheetTDS2 = False
''        If Not ValidateTANdetails Then ValidateSheetTDS2 = False
''
''
''
''
''     setTblinfo_TDS2
''     setTblinfo_TDS2inner
''    setTblinfo_TDS22
''    setTblinfo_TDS23
''    setTblinfo_TDS24
''    setTblinfo_TDS25
''    setTblinfo_TDS26
''    setTblinfo_TDS27
''    setTblinfo_TDS28
''    setTblinfo_TDS29
''    setTblinfo_TDS210
''    setTblinfo_TDS211
''    setTblinfo_TDS212
''    setTblinfo_TDS213
''    setTblinfo_TDS214
''    setTblinfo_TDS215
''    setTblinfo_TDS216
''    setTblinfo_TDS217
''
''
''       'setTblinfo_TDS2inner
''
''           If Mid(Range("TDS2.DrpA").Cells.value, 1, 1) = "Y" Then
''           If (end_TDS2 = 0 Or end_TDS22 = 0 Or end_TDS24 = 0 Or end_TDS25 = 0 Or end_TDS26 = 0 Or end_TDS27 = 0 Or end_TDS28 = 0 Or end_TDS29 = 0 Or end_TDS214 = 0 Or end_TDS215 = 0) Then
''           MsgTDS2 = MsgTDS2 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS2 = False
''            End If
''            End If
''
''
''
''            If ((end_TDS2 <> end_TDS22) Or (end_TDS2 <> end_TDS24) Or (end_TDS2 <> end_TDS25) Or (end_TDS2 <> end_TDS26) Or (end_TDS2 <> end_TDS27) Or (end_TDS2 <> end_TDS28) Or (end_TDS2 <> end_TDS29) Or (end_TDS2 <> end_TDS214) Or (end_TDS2 <> end_TDS215)) Then
''                MsgTDS2 = MsgTDS2 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS2 = False
''            End If
''
''
''            For k = 1 To end_TDS2
''            TDSvalue = 0
''        For j = 1 To end_TDS2inner
''
''        SumTds = Mid(Range("TDS2.DRPTAN").Cells.item(j).value, 3, 10)
''        SumIft = Range("TDSoth.TAN").Cells.item(k).value
''        If (SumTds = SumIft) And CStr(Range("TDS2.SlNo").Cells.item(k).value) = CStr(Mid(Range("TDS2.DRPTAN").Cells.item(j).value, 1, 1)) Then
''         TDSvalue = Range("TDSoth2.TD_OwnHands").Cells.item(k).value
''         End If
''         If CStr(Range("TDS2.SlNo").Cells.item(k).value) = Mid(Range("TDS2.DRPTAN").Cells.item(j).value, 1, 1) Then
''         Range("Column6_value").Cells.item(j).value = Round((Range("TDS2.OthrTds").Cells.item(j).value / TDSvalue) * 100, 2)
''         End If
''        Next
''
''        Next
''
''        If end_TDS2 > 0 Then
''        If Not ValidatePercentOfShare_TDS2 Then ValidateSheetTDS2 = False
''        End If
''        'If Not ValidatePercentOfShare_TDS2 Then ValidateSheetTDS2 = False
''
''            If Mid(Range("TDS2.DrpA").Cells.value, 1, 1) = "Y" Then
''             If end_TDS2 > 0 Then
''
''           ' If (end_TDS2inner <> end_TDS212) Or (end_TDS2inner <> end_TDS213) Or (end_TDS2inner <> end_TDS211) Then
''            If (end_TDS2inner <> end_TDS212) Or (end_TDS2inner <> end_TDS211) Then
''                    MsgTDS2 = MsgTDS2 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS2 = False
''            End If
''            End If
''            End If
''
''            If Mid(Range("TDS2.DrpA").Cells.value, 1, 1) = "Y" Then
''             If end_TDS210 > 0 Then
''           ' If (end_TDS2inner = 0) Or (end_TDS212 = 0) Or (end_TDS213 = 0) Or (end_TDS211 = 0) Then
''            If (end_TDS2inner = 0) Or (end_TDS212 = 0) Or (end_TDS211 = 0) Then
''                    MsgTDS2 = MsgTDS2 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS2 = False
''            End If
''            End If
''            End If
''
''     For k = 1 To end_TDS2
''     tdssum = 0
''     For j = 1 To end_TDS2inner
''     SumTds = Mid(Range("TDS2.DRPTAN").Cells.item(j).value, 3, 10)
''     SumIft = Range("TDSoth.TAN").Cells.item(k).value
''     If (SumTds = SumIft) And CStr(Range("TDS2.SlNo").Cells.item(k).value) = Mid(Range("TDS2.DRPTAN").Cells.item(j).value, 1, 1) Then
''
''     tdssum = WorksheetFunction.Sum(Range("TDS2.OthrTds").Cells.item(j).value, tdssum)
''     End If
''     Next
''
''     If (Range("TDS2.TdsClaimedOwn").item(k).value + tdssum <> (Range("TDSoth2.TD_OwnHands").item(k).value)) Then
''            MsgTDS2 = MsgTDS2 + "For TAN At Sr. No " & k & Chr(13) & "Amount at Tax  Deducted as per 26AS [Col (6)] should be equal to sum of amount at TDS claimed out of (6) in own hands [Col (9)] and amount at TDS out of (6) being apportioned in the hands of other person [Col (11)]" & Chr(13)
''            ValidateSheetTDS2 = False
''            Exit Function
''            End If
''
''     Next
''
''
''
'''If Not checkforNonMandatoryCol_TDS2 Then ValidateSheetTDS2 = False
''
''
''End Function
''
''Function ValidateTAN_TDS2() As Boolean
''    ValidateTAN_TDS2 = True
''    setTblinfo_TDS2
''    Dim rangecells, rangecells1 As Range
''    Dim i As Long
''    Set rangecells = Sheet16.Range("TDSoth.TAN").Cells
''    ReDim TAN_TDS2(end_TDS2)
''       For i = 1 To end_TDS2
''        TAN_TDS2(i) = UCase(rangecells.item(i).value)
''
''            If TAN_TDS2(i) = "" Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & " is Mandatory" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''
''            If Not ValidateTantype_text(Mid(UCase(TAN_TDS2(i)), 1, 4)) Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & " in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''            If Not IsNumeric(Mid(TAN_TDS2(i), 5, 5)) Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''            If Not ValidateTantype_text(Mid(UCase(TAN_TDS2(i)), 10, 1)) Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''
''
''
''    Next
''
''
''End Function
''Function ValidateEmployerOrDeductorOrCollecterName_TDS2() As Boolean
''
''    ValidateEmployerOrDeductorOrCollecterName_TDS2 = True
''    setTblinfo_TDS2
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet16.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
''    ReDim EmployerOrDeductorOrCollecterName_TDS2(end_TDS2)
''    For i = 1 To end_TDS2
''        EmployerOrDeductorOrCollecterName_TDS2(i) = rangecells.item(i).value
''
''        If EmployerOrDeductorOrCollecterName_TDS2(i) = "" Then
''            MsgTDS2 = MsgTDS2 + "Name of Deductor at Sr. No  " & i & "  is Mandatory" & Chr(13)
''            ValidateEmployerOrDeductorOrCollecterName_TDS2 = False
''            Exit Function
''        End If
''
''
''            If Len(EmployerOrDeductorOrCollecterName_TDS2(i)) > 125 Then
''            MsgTDS2 = MsgTDS2 + "Name of Deductor at Sr. No  " & i & "  in Sheet TDS  Cannot be more than 125 Charachters" & Chr(13)
''            ValidateEmployerOrDeductorOrCollecterName_TDS2 = False
''            Exit Function
''        End If
''
''    Next
''End Function
''
''Function ValidateUTN_TDS2() As Boolean
''
''    ValidateUTN_TDS2 = True
''    setTblinfo_TDS2
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet16.Range("TDSoth.UTN").Cells
''    ReDim UTN_TDS2(end_TDS2)
''    For i = 1 To end_TDS2
''        UTN_TDS2(i) = rangecells.item(i).value
''    If Not checkfieldspecialcharacter(UTN_TDS2(i)) Then
''         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
''          ValidateUTN_TDS2 = False
''          Exit Function
''     End If
''
''If InStr(UTN_TDS2(i), ".") > 0 Then
''         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
''          ValidateUTN_TDS2 = False
''          Exit Function
''    End If
''
'''    If Not CheckSpecialCharacter(UTN_TDS2(i)) Then
'''         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
'''          ValidateUTN_TDS2 = False
'''          Exit Function
'''    End If
''
''    If InStr(UTN_TDS2(i), "-") > 0 Then
''         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
''          ValidateUTN_TDS2 = False
''          Exit Function
''    End If
''
''    If UTN_TDS2(i) <> "" Then
''    If Len(UTN_TDS2(i)) < 6 Or Len(UTN_TDS2(i)) > 8 Or Len(UTN_TDS2(i)) <> 8 Then
''          MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
''          ValidateUTN_TDS2 = False
''          Exit Function
''    End If
''    End If
''
''
'' Next
''End Function
''Function ValidateDeductedYear_TDS2() As Boolean
''
''    ValidateDeductedYear_TDS2 = True
''    setTblinfo_TDS2
''    Dim i As Long
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth.DeductedYear").Cells
''    ReDim DeductedYear_TDS2(end_TDS2)
''For i = 1 To end_TDS2
''        DeductedYear_TDS2(i) = rangecells.item(i).value
''
''        If isdropdownblank(DeductedYear_TDS2(i)) Then
''            MsgTDS2 = MsgTDS2 + "Year of tax deduction at Sr. No  " & i & " is Mandatory" & Chr(13)
''            ValidateDeductedYear_TDS2 = False
''            Exit Function
''        End If
''
''
''     If Not checkfieldspecialcharacter(DeductedYear_TDS2(i)) Then
''         MsgTDS2 = MsgTDS2 + "Year of tax deduction at Sr. No  " & i & " in Sheet TDS schedule TDS2 is invalid" & Chr(13)
''          ValidateDeductedYear_TDS2 = False
''          Exit Function
''     End If
''     If DeductedYear_TDS2(i) > 2017 Then
''          MsgTDS2 = MsgTDS2 + "Year of tax deduction at Sr. No  " & i & " in Sheet TDS cannot be greater than 2017 " & Chr(13)
''          ValidateDeductedYear_TDS2 = False
''          Exit Function
''    End If
''Next
''End Function
''
''Function ValidateDetailsOfReceipts_TDS2() As Boolean
''    ValidateDetailsOfReceipts_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim DetailsOfReceipts_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDSoth.AmountPaid").Cells
''
''    For i = 1 To end_TDS2
''        DetailsOfReceipts_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(DetailsOfReceipts_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "Details Of Receipts at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateDetailsOfReceipts_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(DetailsOfReceipts_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "Details Of Receipts at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateDetailsOfReceipts_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTaxDeducted26AS_TDS2() As Boolean
''    ValidateTaxDeducted26AS_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TaxDeducted26AS_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDSoth2.TD_OwnHands").Cells
''
''    For i = 1 To end_TDS2
''        TaxDeducted26AS_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TaxDeducted26AS_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "Tax Deducted as per 26AS at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTaxDeducted26AS_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(TaxDeducted26AS_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "Tax Deducted as per 26AS at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTaxDeducted26AS_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateAmtOfferedSelf_TDS2() As Boolean
''    ValidateAmtOfferedSelf_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim AmtOfferedSelf_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDS2.SelfTdsClaimed").Cells
''
''    For i = 1 To end_TDS2
''        AmtOfferedSelf_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(AmtOfferedSelf_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "Amount offered for Tax at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateAmtOfferedSelf_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(AmtOfferedSelf_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "Amount offered for Tax at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateAmtOfferedSelf_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSClaimedInOwnHands_TDS2() As Boolean
''    ValidateTDSClaimedInOwnHands_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSClaimedInOwnHands_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDS2.TdsClaimedOwn").Cells
''
''    For i = 1 To end_TDS2
''        TDSClaimedInOwnHands_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSClaimedInOwnHands_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "TDS available in Own Hands at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSClaimedInOwnHands_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(TDSClaimedInOwnHands_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "TDS available in Own Hands at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSClaimedInOwnHands_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSApportOtherHands_TDS2() As Boolean
''    ValidateTDSApportOtherHands_TDS2 = True
''    setTblinfo_TDS2inner
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSApportOtherHands_TDS2(end_TDS2inner)
''    Set rangecells = Sheet16.Range("TDS2.OthrTds").Cells
''
''    For i = 1 To end_TDS2inner
''        TDSApportOtherHands_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSApportOtherHands_TDS2(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "TDS apportioned in the hands of other person at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSApportOtherHands_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(TDSApportOtherHands_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "TDS apportioned in the hands of other person at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSApportOtherHands_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeOfferedThisYear_TDS2() As Boolean
''    ValidateIncomeOfferedThisYear_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeOfferedThisYear_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDS2.IncomeOffered").Cells
''
''    For i = 1 To end_TDS2
''        IncomeOfferedThisYear_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeOfferedThisYear_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "Income Offered This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeOfferedThisYear_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(IncomeOfferedThisYear_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "Income Offered This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeOfferedThisYear_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSClaimedThisYear_TDS2() As Boolean
''    ValidateTDSClaimedThisYear_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSClaimedThisYear_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDS2.IncomeClaimed").Cells
''
''    For i = 1 To end_TDS2
''        TDSClaimedThisYear_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSClaimedThisYear_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "TDS Claimed This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSClaimedThisYear_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(TDSClaimedThisYear_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "TDS Claimed This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSClaimedThisYear_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeCarriedFwd_TDS2() As Boolean
''    ValidateIncomeCarriedFwd_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeCarriedFwd_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDS2.IncBalCF").Cells
''
''    For i = 1 To end_TDS2
''        IncomeCarriedFwd_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeCarriedFwd_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "Income Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeCarriedFwd_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(IncomeCarriedFwd_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "Income Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeCarriedFwd_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSCarriedFwd_TDS2() As Boolean
''    ValidateTDSCarriedFwd_TDS2 = True
''    setTblinfo_TDS2
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSCarriedFwd_TDS2(end_TDS2)
''    Set rangecells = Sheet16.Range("TDS2.TDSBalCF").Cells
''
''    For i = 1 To end_TDS2
''        TDSCarriedFwd_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSCarriedFwd_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "TDS Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSCarriedFwd_TDS2 = False
''         Exit Function
''    End If
''
''    If Len(TDSCarriedFwd_TDS2(i)) > 14 Then
''        MsgTDS2 = MsgTDS2 + "TDS Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSCarriedFwd_TDS2 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateHeadOfIncome_TDS2() As Boolean
''
''    ValidateHeadOfIncome_TDS2 = True
''    setTblinfo_TDS2
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDS2.HeadOfIncome").Cells
''    ReDim HeadOfIncome_TDS2(end_TDS2)
''    For i = 1 To end_TDS2
''        HeadOfIncome_TDS2(i) = rangecells.item(i).value
''        If isdropdownblank(HeadOfIncome_TDS2(i)) Then 'And rangecells.item(i).Locked = False Then
''            MsgTDS2 = MsgTDS2 + "Head Of Income at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateHeadOfIncome_TDS2 = False
''              Exit Function
''        End If
''
''        If Mid(HeadOfIncome_TDS2(i), 1, 1) = "H" Then HeadOfIncome_TDS2(i) = "HP"
''        If Mid(HeadOfIncome_TDS2(i), 1, 1) = "O" Then HeadOfIncome_TDS2(i) = "OS"
''
'''        If Not checkfieldspecialcharacter(HeadOfIncome_TDS2(i)) Then
'''              MsgTDS2 = MsgTDS2 + "Special characters are not allowed in Head Of Income at Sr. No  " & i & " ." & Chr(13)
'''              ValidateHeadOfIncome_TDS2 = False
'''              Exit Function
'''        End If
'' Next
''End Function
''Function ValidateTANdetails() As Boolean
''
''    ValidateTANdetails = True
''    setTblinfo_TDS211
''    setTblinfo_TDS2inner
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDS2.DRPTAN").Cells
''    ReDim TANdetails_TDS2(end_TDS2inner)
''    If end_TDS211 > 0 Then
''    For i = 1 To end_TDS2inner
''        TANdetails_TDS2(i) = rangecells.item(i).value
''
''        If Len(TANdetails_TDS2(i)) < 12 Or TANdetails_TDS2(i) = "" Then 'And rangecells.item(i).Locked = False Then
''            MsgTDS2 = MsgTDS2 + "Please select valid TAN From the dropdown at Sr. No  " & i & " ." & Chr(13)
''              ValidateTANdetails = False
''              Exit Function
''        End If
''    Next
''
''    End If
''End Function
''Function ValidateReasonApportTDS_TDS2() As Boolean
''
''    ValidateReasonApportTDS_TDS2 = True
''    setTblinfo_TDS2
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDS2.Reason").Cells
''    ReDim ReasonApportTDS_TDS2(end_TDS2)
''    For i = 1 To end_TDS2
''        ReasonApportTDS_TDS2(i) = rangecells.item(i).value
''        If isdropdownblank(ReasonApportTDS_TDS2(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS2 = MsgTDS2 + "Reason for Apportition at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateReasonApportTDS_TDS2 = False
''              Exit Function
''        End If
''
''        If Mid(ReasonApportTDS_TDS2(i), 1, 1) = "P" Then ReasonApportTDS_TDS2(i) = "PCC"
''        If Mid(ReasonApportTDS_TDS2(i), 1, 1) = "C" Then ReasonApportTDS_TDS2(i) = "COA"
'''        If Not checkfieldspecialcharacter(ReasonApportTDS_TDS2(i)) Then
'''              MsgTDS2 = MsgTDS2 + "Special characters are not allowed in Reason for Apportition at Sr. No  " & i & " ." & Chr(13)
'''              ValidateReasonApportTDS_TDS2 = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidatePANOfOtherPerson_TDS2() As Boolean
''ValidatePANOfOtherPerson_TDS2 = True
''setTblinfo_TDS2
''setTblinfo_TDS2inner
''Dim rangecells, rangecells1 As Range
''Set rangecells = Sheet16.Range("TDSoth.TD_PAN").Cells
''Set rangecells1 = Sheet16.Range("TDS2.DRPTAN").Cells
''Dim i As Long
''ReDim PANOfOtherPerson_TDS2(end_TDS2inner)
''ReDim TAN2_TDS2(end_TDS2inner)
''For i = 1 To end_TDS2inner
''PANOfOtherPerson_TDS2(i) = rangecells.item(i).value
''TAN2_TDS2(i) = rangecells1.item(i).value
''   If Not chkCompulsory(PANOfOtherPerson_TDS2(i)) And rangecells.item(i).Locked = False Then
''        MsgTDS2 = MsgTDS2 + "PAN at Sr.No " & i & " is Mandatory." & Chr(13)
''         ValidatePANOfOtherPerson_TDS2 = False
''         Exit Function
''  End If
''
''    If Not checkfieldspecialcharacter(PANOfOtherPerson_TDS2(i)) Then
''         MsgTDS2 = MsgTDS2 + "PAN at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
''         ValidatePANOfOtherPerson_TDS2 = False
''         Exit Function
''    End If
''
''    If Not CheckPAN(UCase(PANOfOtherPerson_TDS2(i))) Then
''         MsgTDS2 = MsgTDS2 + "PAN at Sr. No  " & i & " in Sheet TDS is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
''         ValidatePANOfOtherPerson_TDS2 = False
''         Exit Function
''    End If
''
''            If PANOfOtherPerson_TDS2(i) = Sheet1.Range("sheet1.PAN").value Then
''                MsgTDS2 = MsgTDS2 + "PAN (on which TDS deducted) at Sr.No " & i & " can not the same as PAN of user  " & Chr(13)
''                ValidatePANOfOtherPerson_TDS2 = False
''                Exit Function
''            End If
''
''Next
''End Function
''
''Function ValidatePercentOfShare_TDS2() As Boolean
''    ValidatePercentOfShare_TDS2 = True
''    setTblinfo_TDS2inner
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim PercentOfShare_TDS2(end_TDS2inner)
''    Set rangecells = Sheet16.Range("TDS2.PercShare").Cells
''
''    For i = 1 To end_TDS2inner
''        PercentOfShare_TDS2(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(PercentOfShare_TDS2(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2 = MsgTDS2 + "Percent Of Share at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidatePercentOfShare_TDS2 = False
''         Exit Function
''    End If
''
'''    If Len(PercentOfShare_TDS2(i)) > 5 Then
'''        MsgTDS2 = MsgTDS2 + "Percent Of Share at Sr. No  " & i & "  should not be greater than 5 digits." & Chr(13)
'''        ValidatePercentOfShare_TDS2 = False
'''    End If
''
''    Next
''End Function
''
''
'' Function ValidatesheetTDS2b() As Boolean
''    ValidatesheetTDS2b = True
''    MsgTDS2b = "Schedule TDS2b :" & Chr(10)
''
''    If Len(Sheet16.Range("TDS2b.Total").value) > 14 Then
''       MsgTDS2b = MsgTDS2b & " Total of TDS2b should not excced 14 digits " & Chr(13)
''       ValidatesheetTDS2b = False
''       Exit Function
''    End If
''
''    setTblinfo_TDS2b
''    setTblinfo_TDS2b2
''    setTblinfo_TDS2b3
''    setTblinfo_TDS2b4
''    setTblinfo_TDS2b5
''    setTblinfo_TDS2b6
''    setTblinfo_TDS2b7
''    setTblinfo_TDS2b8
''    setTblinfo_TDS2b9
''    setTblinfo_TDS2b10
''    setTblinfo_TDS2b11
''    setTblinfo_TDS2b12
''    setTblinfo_TDS2b13
''
''    For i = 1 To end_TDS2b
''        If Len(Sheet16.Range("TDSoth2b.TAN").item(i).value) = 0 Then
''            MsgTDS2b = MsgTDS2b + "Please fill all the mandatory fields of the row " & i & " before filling the next row." & Chr(13)
''            ValidatesheetTDS2b = False
''            Exit Function
''        End If
''    Next
''
''    If Not ValidateTAN_TDS2b() Then ValidatesheetTDS2b = False
''    If (Len(Sheet16.Range("TDSoth2b.TAN").item(1).value) > 0) Then
''        If Not ValidateEmployerOrDeductorOrCollecterName_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateUTN_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateDeductedYear_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateIncomeApportioned_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateTDSApportioned_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateIncomeOfferedThisYear_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateTDSClaimedThisYear_TDS2b() Then ValidatesheetTDS2b = False
''        If Not ValidateIncomeCarriedFwd_TDS2b Then ValidatesheetTDS2b = False
''        If Not ValidateTDSCarriedFwd_TDS2b Then ValidatesheetTDS2b = False
''        If Not ValidateHeadOfIncome_TDS2b Then ValidatesheetTDS2b = False
''        If Not ValidateReasonClaimOwnHands_TDS2b Then ValidatesheetTDS2b = False
''        If Not ValidateTDSDeductedPAN_TDS2b Then ValidatesheetTDS2b = False
''    End If
''
''    If Mid(Range("TDS2.DrpB").Cells.value, 1, 1) = "Y" Then
''           If (end_TDS2b = 0 Or end_TDS2b2 = 0 Or end_TDS2b4 = 0 Or end_TDS2b5 = 0 Or end_TDS2b6 = 0 Or end_TDS2b7 = 0 Or end_TDS2b8 = 0 Or end_TDS2b9 = 0 Or end_TDS2b10 = 0 Or end_TDS2b11 = 0 Or end_TDS2b12 = 0 Or end_TDS2b13 = 0) Then
''           MsgTDS2b = MsgTDS2b + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidatesheetTDS2b = False
''            End If
''            End If
''
''
''
''    For i = 1 To Sheet16.Range("TDSoth2b.TAN").Count
''        If Len(Sheet16.Range("TDSoth2b.UTN").item(i).value) > 0 Then
''
''            If Len(Sheet16.Range("TDSoth2b.TAN").item(i).value) = 0 Then
''                 MsgTDS2b = MsgTDS2b + "Enter All Mandatory Fields at Sr. No " & i & " ." & Chr(13)
''                ValidatesheetTDS2b = False
''                Exit Function
''            End If
''        End If
''    Next
''
''End Function
''
''Function ValidateTAN_TDS2b() As Boolean
''    ValidateTAN_TDS2b = True
''    setTblinfo_TDS2b
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth2b.TAN").Cells
''    ReDim TAN_TDS2b(end_TDS2b)
''
''    subProcCaption = "Validating TDS2b"
''    noOfProcessSub = end_TDS2b
''
''    For i = 1 To end_TDS2b
''        TAN_TDS2b(i) = UCase(rangecells.item(i).value)
'' If Not Len(TAN_TDS2b(i)) = 0 Then
''
''     If Not ValidateTantype_text(Mid(UCase(TAN_TDS2b(i)), 1, 4)) Then
''         MsgTDS2b = MsgTDS2b + "TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet." & Chr(13)
''         ValidateTAN_TDS2b = False
''         Exit Function
''     End If
''
''     If Not IsNumeric(Mid(TAN_TDS2b(i), 5, 5)) Then
''         MsgTDS2b = MsgTDS2b + "TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet." & Chr(13)
''         ValidateTAN_TDS2b = False
''         Exit Function
''     End If
''
''     If Not ValidateTantype_text(Right(UCase(TAN_TDS2b(i)), 1)) Then
''         MsgTDS2b = MsgTDS2b + "TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet." & Chr(13)
''         ValidateTAN_TDS2b = False
''         Exit Function
''     End If
''
'' ElseIf Not chkCompulsory(TAN_TDS2b(i)) Then
''         MsgTDS2b = MsgTDS2b + "TAN at Sr. No  " & i & "  is Mandatory." & Chr(13)
''     ValidateTAN_TDS2b = False
''     Exit Function
'' End If
'' UpdateProgressBar
'' Next
''End Function
''
''Function ValidateEmployerOrDeductorOrCollecterName_TDS2b() As Boolean
''
''    ValidateEmployerOrDeductorOrCollecterName_TDS2b = True
''    setTblinfo_TDS2b
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth2b.EmployerOrDeductorOrCollecterName").Cells
''    ReDim EmployerOrDeductorOrCollecterName_TDS2b(end_TDS2b)
''    For i = 1 To end_TDS2b
''        EmployerOrDeductorOrCollecterName_TDS2b(i) = rangecells.item(i).value
''     If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TDS2b(i)) Then
''         MsgTDS2b = MsgTDS2b + "EmployerOrDeductorOrCollecterNam at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateEmployerOrDeductorOrCollecterName_TDS2b = False
''         Exit Function
''     End If
'' Next
''End Function
''
''Function ValidateUTN_TDS2b() As Boolean
''
''    ValidateUTN_TDS2b = True
''    setTblinfo_TDS2b
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth2b.UTN").Cells
''    ReDim UTN_TDS2b(end_TDS2b)
''    For i = 1 To end_TDS2b
''        UTN_TDS2b(i) = rangecells.item(i).value
''
''    If Not checkfieldspecialcharacter(UTN_TDS2b(i)) Then
''         MsgTDS2b = MsgTDS2b + "Special characters are not allowed in Unique TDS Certificate Number at Sr. No  " & i & " ." & Chr(13)
''          ValidateUTN_TDS2b = False
''          Exit Function
''     End If
''
''If InStr(UTN_TDS2b(i), ".") > 0 Then
''         MsgTDS2b = MsgTDS2b + "Decimals are not allowed in Unique TDS Certificate Number at Sr. No  " & i & " ." & Chr(13)
''          ValidateUTN_TDS2b = False
''          Exit Function
''    End If
''
''    If (UTN_TDS2b(i)) <> "" Then
''    If Len(UTN_TDS2b(i)) <> 8 Then
''    MsgTDS2b = MsgTDS2b + "Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
''          ValidateUTN_TDS2b = False
''          Exit Function
''    End If
''    End If
''
''
''     If (UTN_TDS2b(i)) <> "" Then
''    If Len(UTN_TDS2b(i)) = 6 Then
''        UTN_TDS2b(i) = "00" & UTN_TDS2b(i)
''     End If
''    If Len(UTN_TDS2b(i)) = 7 Then
''        UTN_TDS2b(i) = "0" & UTN_TDS2b(i)
''     End If
''     End If
''
'' Next
''End Function
''
''Function ValidateDeductedYear_TDS2b() As Boolean
''
''    ValidateDeductedYear_TDS2b = True
''    setTblinfo_TDS2b
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth2b.DeductedYear").Cells
''    ReDim DeductedYear_TDS2b(end_TDS2b)
''    For i = 1 To end_TDS2b
''        DeductedYear_TDS2b(i) = rangecells.item(i).value
''        If isdropdownblank(DeductedYear_TDS2b(i)) Then
''            DeductedYear_TDS2b(i) = ""
''        End If
''        If Not checkfieldspecialcharacter(DeductedYear_TDS2b(i)) Then
''              MsgTDS2b = MsgTDS2b + "Special characters are not allowed in Deducted Year at Sr. No  " & i & " ." & Chr(13)
''              ValidateDeductedYear_TDS2b = False
''              Exit Function
''        End If
'' Next
''End Function
''
''Function ValidateIncomeApportioned_TDS2b() As Boolean
''    ValidateIncomeApportioned_TDS2b = True
''    setTblinfo_TDS2b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeApportioned_TDS2b(end_TDS2b)
''    Set rangecells = Sheet16.Range("TDSoth2b.Income").Cells
''
''    For i = 1 To end_TDS2b
''        IncomeApportioned_TDS2b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeApportioned_TDS2b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS2b + "Income Apportioned at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeApportioned_TDS2b = False
''         Exit Function
''    End If
''
''    If Len(IncomeApportioned_TDS2b(i)) > 14 Then
''        MsgTDS2b = MsgTDS2b + "Income Apportioned at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeApportioned_TDS2b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSApportioned_TDS2b() As Boolean
''    ValidateTDSApportioned_TDS2b = True
''    setTblinfo_TDS2b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSApportioned_TDS2b(end_TDS2b)
''    Set rangecells = Sheet16.Range("TDSoth2b.TDS").Cells
''
''    For i = 1 To end_TDS2b
''        TDSApportioned_TDS2b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSApportioned_TDS2b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS2b + "TDS Apportioned at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSApportioned_TDS2b = False
''         Exit Function
''    End If
''
''    If Len(TDSApportioned_TDS2b(i)) > 14 Then
''        MsgTDS2b = MsgTDS2b + "TDS Apportioned at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSApportioned_TDS2b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeOfferedThisYear_TDS2b() As Boolean
''    ValidateIncomeOfferedThisYear_TDS2b = True
''    setTblinfo_TDS2b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeOfferedThisYear_TDS2b(end_TDS2b)
''    Set rangecells = Sheet16.Range("TDSoth2b.IncomeOffered").Cells
''
''    For i = 1 To end_TDS2b
''        IncomeOfferedThisYear_TDS2b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeOfferedThisYear_TDS2b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS2b + "Income Offered This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeOfferedThisYear_TDS2b = False
''         Exit Function
''    End If
''
''    If Len(IncomeOfferedThisYear_TDS2b(i)) > 14 Then
''        MsgTDS2b = MsgTDS2b + "Income Offered This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeOfferedThisYear_TDS2b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSClaimedThisYear_TDS2b() As Boolean
''    ValidateTDSClaimedThisYear_TDS2b = True
''    setTblinfo_TDS2b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSClaimedThisYear_TDS2b(end_TDS2b)
''    Set rangecells = Sheet16.Range("TDSoth2b.TdsClaimed").Cells
''
''    For i = 1 To end_TDS2b
''        TDSClaimedThisYear_TDS2b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSClaimedThisYear_TDS2b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS2b + "TDS Claimed This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSClaimedThisYear_TDS2b = False
''         Exit Function
''    End If
''
''    If Len(TDSClaimedThisYear_TDS2b(i)) > 14 Then
''        MsgTDS2b = MsgTDS2b + "TDS Claimed This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSClaimedThisYear_TDS2b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeCarriedFwd_TDS2b() As Boolean
''    ValidateIncomeCarriedFwd_TDS2b = True
''    setTblinfo_TDS2b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeCarriedFwd_TDS2b(end_TDS2b)
''    Set rangecells = Sheet16.Range("TDSoth2b.IncBalCF").Cells
''
''    For i = 1 To end_TDS2b
''        IncomeCarriedFwd_TDS2b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeCarriedFwd_TDS2b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS2b + "Income Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeCarriedFwd_TDS2b = False
''         Exit Function
''    End If
''
''    If Len(IncomeCarriedFwd_TDS2b(i)) > 14 Then
''        MsgTDS2b = MsgTDS2b + "Income Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeCarriedFwd_TDS2b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSCarriedFwd_TDS2b() As Boolean
''    ValidateTDSCarriedFwd_TDS2b = True
''    setTblinfo_TDS2b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSCarriedFwd_TDS2b(end_TDS2b)
''    Set rangecells = Sheet16.Range("TDSoth2b.TdsBalCF").Cells
''
''    For i = 1 To end_TDS2b
''        TDSCarriedFwd_TDS2b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSCarriedFwd_TDS2b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS2b + "TDS Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSCarriedFwd_TDS2b = False
''         Exit Function
''    End If
''
''    If Len(TDSCarriedFwd_TDS2b(i)) > 14 Then
''        MsgTDS2b = MsgTDS2b + "TDS Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSCarriedFwd_TDS2b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateHeadOfIncome_TDS2b() As Boolean
''
''    ValidateHeadOfIncome_TDS2b = True
''    setTblinfo_TDS2b
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth2b.HeadOfIncome").Cells
''    ReDim HeadOfIncome_TDS2b(end_TDS2b)
''    For i = 1 To end_TDS2b
''        HeadOfIncome_TDS2b(i) = rangecells.item(i).value
''        If isdropdownblank(HeadOfIncome_TDS2b(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS2b = MsgTDS2b + "Head Of Income at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateHeadOfIncome_TDS2b = False
''              Exit Function
''        End If
''        If Mid(HeadOfIncome_TDS2b(i), 1, 1) = "H" Then HeadOfIncome_TDS2b(i) = "HP"
''        If Mid(HeadOfIncome_TDS2b(i), 1, 1) = "O" Then HeadOfIncome_TDS2b(i) = "OS"
''
'''        If Not checkfieldspecialcharacter(HeadOfIncome_TDS2b(i)) Then
'''              MsgTDS2b = MsgTDS2b + "Special characters are not allowed in Head Of Income at Sr. No  " & i & " ." & Chr(13)
'''              ValidateHeadOfIncome_TDS2b = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidateReasonClaimOwnHands_TDS2b() As Boolean
''
''    ValidateReasonClaimOwnHands_TDS2b = True
''    setTblinfo_TDS2b
''    Dim rangecells As Range
''    Set rangecells = Sheet16.Range("TDSoth2b.Reason").Cells
''    ReDim ReasonClaimOwnHands_TDS2b(end_TDS2b)
''    For i = 1 To end_TDS2b
''        ReasonClaimOwnHands_TDS2b(i) = rangecells.item(i).value
''        If isdropdownblank(ReasonClaimOwnHands_TDS2b(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS2b = MsgTDS2b + "Reason for Claiming in Own Hands at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateReasonClaimOwnHands_TDS2b = False
''              Exit Function
''        End If
''
''        If Mid(ReasonClaimOwnHands_TDS2b(i), 1, 1) = "P" Then ReasonClaimOwnHands_TDS2b(i) = "PCC"
''        If Mid(ReasonClaimOwnHands_TDS2b(i), 1, 1) = "C" Then ReasonClaimOwnHands_TDS2b(i) = "COA"
''        If Mid(ReasonClaimOwnHands_TDS2b(i), 1, 1) = "M" Then ReasonClaimOwnHands_TDS2b(i) = "SIC"
''        If Mid(ReasonClaimOwnHands_TDS2b(i), 1, 1) = "T" Then ReasonClaimOwnHands_TDS2b(i) = "TDP"
''        If Mid(ReasonClaimOwnHands_TDS2b(i), 1, 1) = "D" Then ReasonClaimOwnHands_TDS2b(i) = "DPI"
''
'''        If Not checkfieldspecialcharacter(ReasonClaimOwnHands_TDS2b(i)) Then
'''              MsgTDS2b = MsgTDS2b + "Special characters are not allowed in Reason for Claiming in Own Hands at Sr. No  " & i & " ." & Chr(13)
'''              ValidateReasonClaimOwnHands_TDS2b = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidateTDSDeductedPAN_TDS2b() As Boolean
''ValidateTDSDeductedPAN_TDS2b = True
''setTblinfo_TDS2b
''Dim rangecells As Range
''Set rangecells = Sheet16.Range("TDSoth2b.PAN").Cells
''Dim i As Long
''ReDim TDSDeductedPAN_TDS2b(end_TDS2b)
''For i = 1 To end_TDS2b
''TDSDeductedPAN_TDS2b(i) = rangecells.item(i).value
''
''   If Not chkCompulsory(TDSDeductedPAN_TDS2b(i)) And rangecells.item(i).Locked = False Then
''        MsgTDS2b = MsgTDS2b + "PAN at Sr.No " & i & " is Mandatory." & Chr(13)
''         ValidateTDSDeductedPAN_TDS2b = False
''         Exit Function
''  End If
''
''    If Not checkfieldspecialcharacter(TDSDeductedPAN_TDS2b(i)) Then
''         MsgTDS2b = MsgTDS2b + "PAN at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
''         ValidateTDSDeductedPAN_TDS2b = False
''         Exit Function
''    End If
''
''    If Not CheckPAN(UCase(TDSDeductedPAN_TDS2b(i))) Then
''         MsgTDS2b = MsgTDS2b + "PAN at Sr. No  " & i & " in Sheet TDS is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
''         ValidateTDSDeductedPAN_TDS2b = False
''         Exit Function
''    End If
''
''           If TDSDeductedPAN_TDS2b(i) = Sheet1.Range("sheet1.PAN").value Then
''                MsgTDS2 = MsgTDS2 + "PAN (on which TDS deducted) at Sr.No " & i & " can not the same as PAN of user  " & Chr(13)
''                ValidateTDSDeductedPAN_TDS2b = False
''                Exit Function
''            End If
''
''Next
''End Function
''
'Option Explicit
'
'Public TAN_TDS2 As Variant
'Public TAN2_TDS2 As Variant
'Public EmployerOrDeductorOrCollecterName_TDS2 As Variant
'Public UTN_TDS2 As Variant
'Public DeductedYear_TDS2 As Variant
'Public AmtPaid_TDS2 As Variant
'Public AmtClaimedOnOwnHands_TDS2 As Variant
'Public TD_SIncome_TDS2 As Variant
'Public TD_OwnHands_TDS2 As Variant
'Public TD_TDS_TDS2 As Variant
'Public TDS_SIncome_TDS2 As Variant
'Public TDS_TDS_TDS2 As Variant
'Public PAN7_TDS As Variant
'Public PAN8_TDS As Variant
'Public PAN9other_TDS As Variant
'Public PAN10other_TDS As Variant
'Dim TotalSum As Variant
'Dim TotalSum1 As Variant
'
'Dim rngname_TDS22 As Variant
'Dim end_TDS22 As Variant
'Dim rngname_TDS2 As Variant
'Dim rngname_TDS2_8 As Variant
'Public end_TDS2 As Variant
'
'
'Dim end_TDS2_8, end_TDS2inner As Variant
'
'Public MsgTDS1, MsgTDS2 As Variant
'
'Function ValidateSheetTDS2() As Boolean
'On Error Resume Next
'Dim cellcount, intcntr, rangname, name As Variant
'Dim SumTds, SumIft, tdssum As Variant
'
'Dim k, j As Long
'
'ValidateSheetTDS2 = True
'    MsgTDS2 = "TDS 2." & Chr(13)
'    subProcCaption = "Validating TDS2"
'    setTblinfo_TDS2
'    noOfProcessSub = end_TDS2
'For k = 1 To end_TDS2
'
'        If Not (Sheet22.Range("TDSoth.TAN").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet22.Range("TDSoth.DeductedYear").Cells.item(k)) Or Sheet22.Range("TDSoth.AmountPaid").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_OwnHands").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells.item(k) <> "") Then
'            MsgTDS2 = MsgTDS2 & " At Sr. No " & k & " Please fill all the Mandatory Fields " & Chr(13)
'            ValidateSheetTDS2 = False
'            Exit Function
'        End If
'
'    UpdateProgressBar
'Next
'
'If end_TDS2 > 0 Then GoTo xyz
'
'    cellcount = Sheet22.Range("TDSoth.TAN").Count
'    If Mid(Range("sheet1.PortugeseCC5A").value, 1, 3) <> "Yes" Then
'        rangname = "TDSoth.TAN|TDSoth.EmployerOrDeductorOrCollecterName|TDSoth.DeductedYear |TDSoth.AmountPaid|TDSoth.TD_OwnHands|TDSoth.TD_SIncome|TDSoth.TD_TDS|TDSoth.ClaimOutOfTotTDSOnAmtPaid|TDSoth.TDS_SIncome|TDSoth.TDS_TDS"
'    Else:
'        rangname = "TDSoth.TAN|TDSoth.EmployerOrDeductorOrCollecterName|TDSoth.DeductedYear|TDSoth.AmountPaid|TDSoth.TD_OwnHands|TDSoth.TD_SIncome|TDSoth.TD_TDS|TDSoth.ClaimOutOfTotTDSOnAmtPaid|TDSoth.TDS_SIncome|TDSoth.TDS_TDS|TDSoth.TD_PAN|TDSoth.TDS_PAN"
'    End If
'    rangname = Split(rangname, "|")
'
'If Mid(Range("sheet1.PortugeseCC5A").value, 1, 3) <> "Yes" Then
'    For intcntr = 1 To cellcount
'        If (Sheet22.Range("TDSoth.TAN").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet22.Range("TDSoth.DeductedYear").Cells.item(k)) Or Sheet22.Range("TDSoth.AmountPaid").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_OwnHands").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_SIncome").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_TDS").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TDS_SIncome").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TDS_TDS").Cells.item(k) <> "") Then
'            For Each name In rangname
'                If Sheet22.Range(name).Cells.item(intcntr) = "" Then ValidateSheetTDS2 = False
'            Next
'            If ValidateSheetTDS2 = False Then
'                MsgTDS2 = MsgTDS2 & "At Sr. No " & intcntr & " Please fill all the Mandatory Fields " & Chr(13)
'                GoTo xyz
'            End If
'
'        End If
'    Next
'
'Else:
'    For intcntr = 1 To cellcount
'        If (Sheet22.Range("TDSoth.TAN").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet22.Range("TDSoth.DeductedYear").Cells.item(k)) Or Sheet22.Range("TDSoth.AmountPaid").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_OwnHands").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_SIncome").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_TDS").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TDS_SIncome").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TDS_TDS").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TD_PAN").Cells.item(k) <> "" Or Sheet22.Range("TDSoth.TDS_PAN").Cells.item(k) <> "") Then
'            For Each name In rangname
'                If Sheet22.Range(name).Cells.item(intcntr) = "" Then ValidateSheetTDS2 = False
'            Next
'            If ValidateSheetTDS2 = False Then
'                MsgTDS2 = MsgTDS2 & "At Sr. No " & intcntr & " Please fill all the Mandatory Fields " & Chr(13)
'                GoTo xyz
'            End If
'        End If
'    Next
'End If
'
'xyz:
'    If Not ValidateTAN_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidateEmployerOrDeductorOrCollecterName_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidateUTN_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidateDeductedYear_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidateClaimOutOfTotTDSOnAmtPaid_TDS2() Then ValidateSheetTDS2 = False
'    If Not ValidatePAN7_TDS() Then ValidateSheetTDS2 = False
'
'    'If Not ValidatePAN7_TDS() Then ValidateSheetTDS2 = False
'
'
'        'setTableinfo_TDS_PANp_TaxTDS1
'        setTableinfo_TDS_PAN_TaxTDS1
'        setTableinfo_TDS_TDS_TDS2
'        setTableinfo_TDS_SIncome_TDS2
'       ' setTableinfo_TD_PANp_TDS2
'        setTableinfo_TD_PAN_TDS2
'        setTableinfo_TD_TDS_TDS2
'        setTableinfo_TDsinc_TaxTDS1
'        setTableinfo_Amt_TaxTDS2
'        setTableinfo_TD_OwnHand
'        setTableinfo_AmtPaidTDS2
'        setTableinfo_Year_TDS1
'        setTableinfo_Name_TDS2
'       setTblinfo_TDS2inner
'
'
'            If ((end_TDS2 <> end_Name_TDS2) Or (end_TDS2 <> end_Year_TDS2) Or (end_TDS2 <> end_AmtPaid_TDS2) Or (end_TDS2 <> end_TDi_TDS2) Or (end_TDS2 <> end_Amt_TDS2)) Then
'                MsgTDS2 = MsgTDS2 + "Please fill all the Mandatory Fields " & Chr(13)
'                ValidateSheetTDS2 = False
'            End If
'
'            If ((end_TDii_TDS2 <> end_TDiii_TDS2) Or (end_TDii_TDS2 <> end_TD_PAN_TDS2)) Or ((end_TDS_SIncome_TDS2 <> end_TDS_TDS_TDS2) Or (end_TDS_SIncome_TDS2 <> end_TDS_PAN_TDS2)) Then
'             MsgTDS2 = MsgTDS2 + "Please fill all the Mandatory Fields " & Chr(13)
'                ValidateSheetTDS2 = False
'            End If
'
''    For k = 1 To end_TDS2
''
''            Next
'
'     For k = 1 To end_TDS2
'     tdssum = 0
'     For j = 1 To end_TDS2inner
'     SumTds = Mid(Range("TAN_1").Cells.item(j).value, 3, 10) '('Application.WorksheetFunction.SumIf((Range("TDSoth.TDS_TDS")), (Mid(Range("TAN_1").Cells.item(j).value, 3, 10) = Range("TDSoth.TAN").Cells.item(k).value))
'     'SumIft = (Sheet22.Range("TAN_1").Cells.item(j).value)
'     'SumIft = Application.WorksheetFunction.Mid((Sheet22.Range("TAN_1").Cells.item(j).value), 1, 2)
'     SumIft = Range("TDSoth.TAN").Cells.item(k).value
'     If (SumTds = SumIft) Then
'     tdssum = WorksheetFunction.Sum(Range("TDSoth.TDS_TDS").Cells.item(j).value, tdssum)
'     End If
'     Next
'     If (Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").item(k).value + tdssum > (Range("TDSoth.TD_TDS").item(k).value + Range("TDSoth.TD_OwnHands").item(k).value)) Then
'            MsgTDS2 = MsgTDS2 + "For TAN At Sr. No " & k & Chr(13) & "Amount in field (Col (8)+ TDS at Col.(9)) cannot be more than sum of field (Col (6) and TDS at col.(7))." & Chr(13)
'            ValidateSheetTDS2 = False
'            Exit Function
'            End If
'
'     Next
'
'
'
'If ValidateSheetTDS2 = False Then GoTo Next1
'If Not checkforNonMandatoryCol_TDS2 Then ValidateSheetTDS2 = False
'
'Next1:
'End Function
'
'
'Function checkforNonMandatoryCol_TDS2() As Boolean
'checkforNonMandatoryCol_TDS2 = True
'    Dim temp As Double
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.UTN").Count
'    Set rangecells = Range("TDSoth.UTN").Cells
'
'    temp = end_TDS2
'
'    For mIntCtr = (temp + 1) To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                MsgTDS2 = MsgTDS2 & "At Sr. No " & mIntCtr & " Please fill all the Mandatory Fields " & Chr(13)
'                checkforNonMandatoryCol_TDS2 = False
'                Exit Function
'           End If
'    Next
'End Function
'
'Sub setTblinfo_TDS2()
' Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TAN").Count
' Set rangecells = Sheet22.Range("TDSoth.TAN").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2 = ccount
' rngname_TDS2 = "TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDSoth.UTN||TDSoth.DeductedYear||TDSoth.AmountPaid ||TDSoth.TD_OwnHands ||TDSoth.TD_SIncome||TDSoth.TD_TDS||TDSoth.TD_PAN||TDSoth.ClaimOutOfTotTDSOnAmtPaid||TDSoth.TDS_SIncome||TDSoth.TDS_TDS||TDSoth.TDS_PAN"
' End Sub
'
' Sub setTblinfo_TDS2inner()
' Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TAN_1").Count
' Set rangecells = Sheet22.Range("TAN_1").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2inner = ccount
' 'rngname_TDS2inner = "TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDSoth.UTN||TDSoth.DeductedYear||TDSoth.AmountPaid ||TDSoth.TD_OwnHands ||TDSoth.TD_SIncome||TDSoth.TD_TDS||TDSoth.TD_PAN||TDSoth.ClaimOutOfTotTDSOnAmtPaid||TDSoth.TDS_SIncome||TDSoth.TDS_TDS||TDSoth.TDS_PAN"
' End Sub
'
'
'Function ValidateUTN_TDS2() As Boolean
'
'    ValidateUTN_TDS2 = True
'    setTblinfo_TDS2
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet22.Range("TDSoth.UTN").Cells
'    ReDim UTN_TDS2(end_TDS2)
'    For i = 1 To end_TDS2
'        UTN_TDS2(i) = rangecells.item(i).value
'    If Not checkfieldspecialcharacter(UTN_TDS2(i)) Then
'         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
'          ValidateUTN_TDS2 = False
'          Exit Function
'     End If
'
'If InStr(UTN_TDS2(i), ".") > 0 Then
'         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
'          ValidateUTN_TDS2 = False
'          Exit Function
'    End If
'
'    If InStr(UTN_TDS2(i), "-") > 0 Then
'         MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
'          ValidateUTN_TDS2 = False
'          Exit Function
'    End If
'
'    If UTN_TDS2(i) <> "" Then
'    If Len(UTN_TDS2(i)) < 6 Or Len(UTN_TDS2(i)) > 8 Then
'          MsgTDS2 = MsgTDS2 + "Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
'          ValidateUTN_TDS2 = False
'          Exit Function
'    End If
'    End If
'
'
' Next
'End Function
'
'Function ValidateTAN_TDS2() As Boolean
'    ValidateTAN_TDS2 = True
'    setTblinfo_TDS2
'    setTblinfo_TDS2inner
'    Dim rangecells, rangecells1 As Range
'    Dim i As Long
'    Set rangecells = Sheet22.Range("TDSoth.TAN").Cells
'    Set rangecells1 = Sheet22.Range("TAN_1").Cells
'    ReDim TAN_TDS2(end_TDS2)
'    ReDim TAN2_TDS2(end_TDS2inner)
'       For i = 1 To end_TDS2
'        TAN_TDS2(i) = UCase(rangecells.item(i).value)
'
'            If TAN_TDS2(i) = "" Then
'                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & " is Mandatory" & Chr(13)
'                ValidateTAN_TDS2 = False
'                Exit Function
'            End If
'            If Not ValidateTantype_text(Mid(TAN_TDS2(i), 1, 4)) Then
'                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & " in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
'                ValidateTAN_TDS2 = False
'                Exit Function
'            End If
'            If Not IsNumeric(Mid(TAN_TDS2(i), 5, 5)) Then
'                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
'                ValidateTAN_TDS2 = False
'                Exit Function
'            End If
'            If Not ValidateTantype_text(Mid(TAN_TDS2(i), 10, 1)) Then
'                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
'                ValidateTAN_TDS2 = False
'                Exit Function
'            End If
'
'    Next
'
'    For i = 1 To end_TDS2inner
'        TAN2_TDS2(i) = UCase(rangecells1.item(i).value)
'
''            If TAN2_TDS2(i) = "" Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & " is Mandatory" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''            If Not ValidateTantype_text(Mid(TAN_TDS2(i), 1, 4)) Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & " in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''            If Not IsNumeric(Mid(TAN_TDS2(i), 5, 5)) Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
''            If Not ValidateTantype_text(Mid(TAN_TDS2(i), 10, 1)) Then
''                MsgTDS2 = MsgTDS2 + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
''                ValidateTAN_TDS2 = False
''                Exit Function
''            End If
'
'    Next
'
'
'End Function
'
'Function ValidateEmployerOrDeductorOrCollecterName_TDS2() As Boolean
'
'    ValidateEmployerOrDeductorOrCollecterName_TDS2 = True
'    setTblinfo_TDS2
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet22.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
'    ReDim EmployerOrDeductorOrCollecterName_TDS2(end_TDS2)
'    For i = 1 To end_TDS2
'        EmployerOrDeductorOrCollecterName_TDS2(i) = rangecells.item(i).value
'
'        If EmployerOrDeductorOrCollecterName_TDS2(i) = "" Then
'            MsgTDS2 = MsgTDS2 + "Name of Employer at Sr. No  " & i & "  is Mandatory" & Chr(13)
'            ValidateEmployerOrDeductorOrCollecterName_TDS2 = False
'            Exit Function
'        End If
'
'
'            If Len(EmployerOrDeductorOrCollecterName_TDS2(i)) > 125 Then
'            MsgTDS2 = MsgTDS2 + "Name of Employer at Sr. No  " & i & "  in Sheet TDS  Cannot be more than 125 Charachters" & Chr(13)
'            ValidateEmployerOrDeductorOrCollecterName_TDS2 = False
'            Exit Function
'        End If
'
'    Next
'End Function
'
'Function ValidateClaimOutOfTotTDSOnAmtPaid_TDS2() As Boolean
'    ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = True
'    setTblinfo_TDS2
'    setTblinfo_TDS2inner
'    setTableinfo_TDS_SIncome_TDS2
'    setTableinfo_TDsinc_TaxTDS1
'    Dim rangecells1 As Range
'    Dim rangecells2 As Range
'    Dim rangecells5 As Range
'    Dim rangecells6 As Range
'    Dim rangecells7 As Range
'    Dim rangecells8 As Range
'    Dim rangecells9 As Range
'
'
'    Dim i As Long
'    Dim portugesecode As Variant
'
'    Set rangecells6 = Sheet22.Range("TDSoth.AmountPaid").Cells
'    Set rangecells1 = Sheet22.Range("TDSoth.TD_OwnHands").Cells
'    Set rangecells5 = Sheet22.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
'
'    Set rangecells2 = Sheet22.Range("TDSoth.TD_SIncome").Cells
'    Set rangecells7 = Sheet22.Range("TDSoth.TD_TDS").Cells
'
'    Set rangecells8 = Sheet22.Range("TDSoth.TDS_SIncome").Cells
'    Set rangecells9 = Sheet22.Range("TDSoth.TDS_TDS").Cells
'
'    portugesecode = Trim(UCase(Sheet1.Range("sheet1.PortugeseCC5A").value))
'
'    ReDim AmtPaid_TDS2(end_TDS2)
'    ReDim TD_OwnHands_TDS2(end_TDS2)
'    ReDim AmtClaimedOnOwnHands_TDS2(end_TDS2)
'
'    ReDim TD_SIncome_TDS2(end_TDS2)
'    ReDim TD_TDS_TDS2(end_TDS2)
'
'    ReDim TDS_SIncome_TDS2(end_TDS2inner)
'    ReDim TDS_TDS_TDS2(end_TDS2inner)
'
'    ReDim TotalSum(end_TDS2)
'    ReDim TotalSum1(end_TDS2)
'
'For i = 1 To end_TDS2
'
'    AmtPaid_TDS2(i) = rangecells6.item(i).value
'    AmtClaimedOnOwnHands_TDS2(i) = rangecells5.item(i).value
'    TD_OwnHands_TDS2(i) = rangecells1.item(i).value
'
'
''    If BroughtFwdTDSAmt_TDS2(i) = "" Then
''        MsgTDS2 = MsgTDS2 + "Toatal Tax Deducted at Sr.No " & i & " is Mandatory" & Chr(13)
''        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''    End If
''
'    If AmtClaimedOnOwnHands_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Amount claimed this year at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
'    If AmtPaid_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Please enter the details of Receipt as mentioned in Form 26AS at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
'     If TD_OwnHands_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Please enter Tax deducted in own hands at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
''    If IsNumeric(BroughtFwdTDSAmt_TDS2(i)) Then
''        If BroughtFwdTDSAmt_TDS2(i) < 0 Then
''            MsgTDS2 = MsgTDS2 + "Amount in col.no 6 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
''            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''        End If
''        If Len("" & BroughtFwdTDSAmt_TDS2(i)) > 14 Then
''            MsgTDS2 = MsgTDS2 + "Amount in col.no 6 cannot exceed 14 digits in Row  " & i & Chr(13)
''            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''            Exit Function
''        End If
''    Else
''        MsgTDS2 = MsgTDS2 + "Amount in col.no 6 is invalid in Row  " & i & Chr(13)
''        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''        Exit Function
''    End If
'
'    If IsNumeric(AmtPaid_TDS2(i)) Then
'        If AmtPaid_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 5 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & AmtPaid_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 5 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 5 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'
'    If IsNumeric(AmtClaimedOnOwnHands_TDS2(i)) Then
'        If AmtClaimedOnOwnHands_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & AmtClaimedOnOwnHands_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 8 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'
'
'
'    If IsNumeric(TD_OwnHands_TDS2(i)) Then
'        If TD_OwnHands_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 6 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & TD_OwnHands_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 6 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 6 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'
'
''    If IsNumeric(AmtClaimedBySpouse_TDS2(i)) Then
''        If AmtClaimedBySpouse_TDS2(i) < 0 Then
''            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
''            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''        End If
''        If Len("" & AmtClaimedBySpouse_TDS2(i)) > 14 Then
''            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
''            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''            Exit Function
''        End If
''    Else
''        MsgTDS2 = MsgTDS2 + "Amount in col.no 8 is invalid in Row " & i & Chr(13)
''        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''        Exit Function
''    End If
'
'
''    If (BroughtFwdTDSAmt_TDS2(i) = "" Or BroughtFwdTDSAmt_TDS2(i) = 0) Then
''        BroughtFwdTDSAmt_TDS2(i) = 0
''    End If
'
'   If (AmtClaimedOnOwnHands_TDS2(i) = "") Then
'        AmtClaimedOnOwnHands_TDS2(i) = 0
'    End If
'
'' If portugesecode = "YES" Then
''    If AmtClaimedBySpouse_TDS2(i) = "" Or IsEmpty(AmtClaimedBySpouse_TDS2(i)) Then
''        MsgTDS2 = MsgTDS2 + "Amount claimed in the hands  of spouse is Mandatory as the assessee is governed by Portuguese Civil Code under Sec 5A. In case of nil amount, please enter zero, in Row  " & i & Chr(13)
''        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''        Exit Function
''    End If
''ElseIf portugesecode = "NO" Then
''  If AmtClaimedBySpouse_TDS2(i) <> "" Then
''        MsgTDS2 = MsgTDS2 + "Amount claimed in the hands of spouse is not Mandatory as the assessee is not governed by Portuguese Civil Code under Sec 5A. please dont enter any value, in Row  " & i & Chr(13)
''        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
''        Exit Function
''  End If
''End If
'
'
'
'    'If AmtPaid_TDS2(i) < (2 * (BroughtFwdTDSAmt_TDS2(i))) Then
'    '    MsgTDS2 = MsgTDS2 + "Details of Receipt as mentioned in Form 26AS should not be less than twice of the amount of tax deducted in Row " & i & Chr(13)
'    '    ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    '    Exit Function
'    'End If
'Next
'
'For i = 1 To end_TDS2
'
'    TD_SIncome_TDS2(i) = rangecells2.item(i).value
'    TD_TDS_TDS2(i) = rangecells7.item(i).value
'
'    If (TD_SIncome_TDS2(i) <> "" Or TD_TDS_TDS2(i) <> "") Then
'    If TD_SIncome_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Please enter Income which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No" & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
'    If TD_TDS_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Please enter Tax deducted in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
'    If IsNumeric(TD_TDS_TDS2(i)) Then
'        If TD_TDS_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & TD_TDS_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 7 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'
'    If IsNumeric(TD_SIncome_TDS2(i)) Then
'        If TD_SIncome_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & TD_SIncome_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 7 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'    End If
'
'
'
'Next
'
'For i = 1 To end_TDS2inner
'
'    TDS_SIncome_TDS2(i) = rangecells8.item(i).value
'    TDS_TDS_TDS2(i) = rangecells9.item(i).value
'
'
'    If TDS_SIncome_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Please enter Income which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
'    If TDS_TDS_TDS2(i) = "" Then
'        MsgTDS2 = MsgTDS2 + "Please enter Tax deducted in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'    End If
'
'
'    If IsNumeric(TDS_TDS_TDS2(i)) Then
'        If TDS_TDS_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & TDS_TDS_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 8 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'
'        If IsNumeric(TDS_SIncome_TDS2(i)) Then
'        If TDS_SIncome_TDS2(i) < 0 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        End If
'        If Len("" & TDS_SIncome_TDS2(i)) > 14 Then
'            MsgTDS2 = MsgTDS2 + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2 = MsgTDS2 + "Amount in col.no 8 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2 = False
'        Exit Function
'    End If
'
'Next
'
'End Function
'
'Function ValidatePAN7_TDS() As Boolean
'ValidatePAN7_TDS = True
'setTableinfo_TDS_SIncome_TDS2
'setTableinfo_TDsinc_TaxTDS1
'setTblinfo_TDS2inner
'setTblinfo_TDS2
'noOfProcessSub = end_TDS2
'
'ValidateClaimOutOfTotTDSOnAmtPaid_TDS2
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'
'Set rangecells = Range("TDSoth.TD_PAN").Cells
'Set rangecells1 = Range("TDSoth.TDS_PAN").Cells
'
'
'Dim i As Long
'ReDim PAN7_TDS(end_TDS2)
'ReDim PAN8_TDS(end_TDS2inner)
'For i = 1 To end_TDS2
'    PAN7_TDS(i) = rangecells.item(i).value
'
'    If (TD_SIncome_TDS2(i) <> "" Or TD_TDS_TDS2(i) <> "") Then
'
'    If Not Len(PAN7_TDS(i)) = 0 Then
'        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN7_TDS(i), 1, 10))) Then
'        MsgTDS2 = MsgTDS2 + "PAN of spouse / other person  at Sr. No  " & i & "  in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
'        ValidatePAN7_TDS = False
'        End If
'
'
'    ElseIf Not chkCompulsory(PAN7_TDS(i)) Then
'        MsgTDS2 = MsgTDS2 + " PAN of spouse / other person " & i & " in Sheet TDS, (20 TDS2(i)) is mandatory" & Chr(13)
'        ValidatePAN7_TDS = False
'    End If
'    End If
'Next
'
'For i = 1 To end_TDS2inner
'PAN8_TDS(i) = rangecells1.item(i).value
'
'    If Not Len(PAN8_TDS(i)) = 0 Then
'        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN8_TDS(i), 1, 10))) Then
'        MsgTDS2 = MsgTDS2 + "PAN of spouse / other person  at Sr. No  " & i & "  in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
'        ValidatePAN7_TDS = False
'        End If
'
'
'    ElseIf Not chkCompulsory(PAN8_TDS(i)) Then
'        MsgTDS2 = MsgTDS2 + " PAN of spouse / other person at Sr.No " & i & " in Sheet TDS, (20 TDS2(i)) is mandatory" & Chr(13)
'        ValidatePAN7_TDS = False
'    End If
'
'
'    UpdateProgressBar
'    Next
'End Function
'
'
'Function ValidateDeductedYear_TDS2() As Boolean
'
'    ValidateDeductedYear_TDS2 = True
'    setTblinfo_TDS2
'    Dim i As Long
'    Dim rangecells As Range
'    Set rangecells = Sheet22.Range("TDSoth.DeductedYear").Cells
'    ReDim DeductedYear_TDS2(end_TDS2)
'For i = 1 To end_TDS2
'        DeductedYear_TDS2(i) = rangecells.item(i).value
'
'        If isdropdownblank(DeductedYear_TDS2(i)) Then
'            MsgTDS2 = MsgTDS2 + "Deducted Year at Sr. No  " & i & " is Mandatory" & Chr(13)
'            ValidateDeductedYear_TDS2 = False
'            Exit Function
'        End If
'
'
'     If Not checkfieldspecialcharacter(DeductedYear_TDS2(i)) Then
'         MsgTDS2 = MsgTDS2 + "Deducted Year at Sr. No  " & i & " in Sheet TDS schedule TDS2 is invalid" & Chr(13)
'          ValidateDeductedYear_TDS2 = False
'          Exit Function
'     End If
'     If DeductedYear_TDS2(i) > 2017 Then
'          MsgTDS2 = MsgTDS2 + "Deducted Year at Sr. No  " & i & " in Sheet TDS cannot be greater than 2017 " & Chr(13)
'          ValidateDeductedYear_TDS2 = False
'          Exit Function
'    End If
'Next
'End Function
'
'
'
'
'
'' ------- FOR TDS2-----'
'
'Sub setTableinfo_Name_TDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.EmployerOrDeductorOrCollecterName").Count
' Set rangecells = Sheet22.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_Name_TDS2 = ccount
'End Sub
'
'
'Sub setTableinfo_Year_TDS1()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.DeductedYear").Count
' Set rangecells = Sheet22.Range("TDSoth.DeductedYear").Cells
' For mIntCtr = 1 To mIntCells
'     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
'         ccount = ccount + 1
'     End If
' Next
' end_Year_TDS2 = ccount
'End Sub
'
'Sub setTableinfo_AmtPaidTDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.AmountPaid").Count
' Set rangecells = Sheet22.Range("TDSoth.AmountPaid").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_AmtPaid_TDS2 = ccount
'End Sub
''
'Sub setTableinfo_TD_OwnHand()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TD_OwnHands").Count
' Set rangecells = Sheet22.Range("TDSoth.TD_OwnHands").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDi_TDS2 = ccount
'End Sub
'
'
'Sub setTableinfo_Amt_TaxTDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Count
' Set rangecells = Sheet22.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_Amt_TDS2 = ccount
'End Sub
'
'
'Sub setTableinfo_TDsinc_TaxTDS1()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TD_SIncome").Count
' Set rangecells = Sheet22.Range("TDSoth.TD_SIncome").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDii_TDS2 = ccount
'End Sub
'
'Sub setTableinfo_TD_TDS_TDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TD_TDS").Count
' Set rangecells = Sheet22.Range("TDSoth.TD_TDS").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDiii_TDS2 = ccount
'End Sub
'
'
'Sub setTableinfo_TD_PAN_TDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TD_PAN").Count
' Set rangecells = Sheet22.Range("TDSoth.TD_PAN").Cells
' For mIntCtr = 1 To mIntCells
'     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
'         ccount = ccount + 1
'     End If
' Next
' end_TD_PAN_TDS2 = ccount
'End Sub
'
'Sub setTableinfo_TDS_SIncome_TDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TDS_SIncome").Count
' Set rangecells = Sheet22.Range("TDSoth.TDS_SIncome").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS_SIncome_TDS2 = ccount
'End Sub
'
'
'Sub setTableinfo_TDS_TDS_TDS2()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TDS_TDS").Count
' Set rangecells = Sheet22.Range("TDSoth.TDS_TDS").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS_TDS_TDS2 = ccount
'End Sub
'
'
'Sub setTableinfo_TDS_PAN_TaxTDS1()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet22.Range("TDSoth.TDS_PAN").Count
' Set rangecells = Sheet22.Range("TDSoth.TDS_PAN").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS_PAN_TDS2 = ccount
'End Sub
'
'Sub ValidateTDS_TDS2()
'On Error Resume Next
'
'    MsgTDS2 = ""
'    If Not ValidateSheetTDS2 Then
'        Sheet22.Activate
'        MsgBox (MsgTDS2)
'        CloseMsg
'    End If
'
'    MsgBox "Sheet is Ok", vbOKOnly, vbMessageCaption
'
'End Sub
'Sub Cmd_Prev_Click_TDS2()
'Sheet4.Activate
'End Sub
'
'Sub Cmd_Next_Click_TDS2()
'Sheet21.Activate
'End Sub
'
'
