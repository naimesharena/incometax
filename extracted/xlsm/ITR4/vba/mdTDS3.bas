Attribute VB_Name = "mdTDS3"
''Option Explicit
''
''Public end_TDS3, end_TDS3inner, end_TDS32, end_TDS33, end_TDS34, end_TDS35, end_TDS36, end_TDS318, end_TDS37, end_TDS38, end_TDS39, end_TDS310, end_TDS311, end_TDS312, end_TDS313, end_TDS314, end_TDS315, end_TDS316, end_TDS317 As Variant
''Public end_TDS3b, end_TDS3b2, end_TDS3b3, end_TDS3b4, end_TDS3b5, end_TDS3b6, end_TDS3b7, end_TDS3b8, end_TDS3b9, end_TDS3b10, end_TDS3b11, end_TDS3b12, end_TDS3b13 As Variant
''Public rngname_TDS3, rngname_TDS3b As Variant
''Public i As Long
''Public MsgTDS3, MsgTDS3b As Variant
''Public TDS3ApportionedCheck As Variant
''Public TDS3ClaimCheck As Variant
''
''Public PAN_TDS3, PAN2_TDS3 As Variant
''Public NameOfTeanant_TDS3 As Variant
''Public UTN_TDS3 As Variant
''Public DeductedYear_TDS3 As Variant
''Public AmtDed_TDS3 As Variant
''Public DetailsOfReceipts_TDS3 As Variant
''Public TaxDeducted26AS_TDS3 As Variant
''Public AmtOfferedSelf_TDS3 As Variant
''Public TDSClaimedInOwnHands_TDS3 As Variant
''Public TDSApportOtherHands_TDS3 As Variant
''
''Public IncomeOfferedThisYear_TDS3 As Variant
''Public TDSClaimedThisYear_TDS3 As Variant
''Public IncomeCarriedFwd_TDS3 As Variant
''Public TDSCarriedFwd_TDS3 As Variant
''
''Public HeadOfIncome_TDS3 As Variant
''Public ReasonApportTDS_TDS3 As Variant
''Public PANOfOtherPerson_TDS3 As Variant
''Public PercentOfShare_TDS3 As Variant
''
''
''Public PAN_TDS3b As Variant
''Public Name_TDS3b As Variant
''Public UTN_TDS3b As Variant
''Public DeductedYear_TDS3b As Variant
''Public IncomeApportioned_TDS3b As Variant
''Public TDSApportioned_TDS3b As Variant
''Public IncomeOfferedThisYear_TDS3b As Variant
''Public TDSClaimedThisYear_TDS3b As Variant
''Public IncomeCarriedFwd_TDS3b As Variant
''Public TDSCarriedFwd_TDS3b As Variant
''Public HeadOfIncome_TDS3b As Variant
''Public ReasonClaimOwnHands_TDS3b As Variant
''Public TDSDeductedPAN_TDS3b As Variant
''
'
''
''Sub ValidateTDS3_Click()
''    ValidateSheetSchTDS3
''    MsgBox "Sheet is ok", vbOKOnly, "ITR-2"
''End Sub
''
''
''Sub ValidateSheetSchTDS3()
''   If Not ValidateTDS3ApportionedCheck Then
''        Sheet19.Activate
''        MsgBox (MsgTDS3)
''        CloseMsg
''    End If
''
''    If Not ValidateTDS3ClaimCheck Then
''        Sheet19.Activate
''        MsgBox (MsgTDS3)
''        CloseMsg
''    End If
''
''    If Not ValidateSheetTDS3 Then
''        Sheet19.Activate
''        MsgBox (MsgTDS3)
''        CloseMsg
''    End If
''
''    If Not ValidatesheetTDS3b Then
''        Sheet19.Activate
''        MsgBox MsgTDS3b, vbOKOnly, "ITR-2"
''        CloseMsg
''    End If
''End Sub
''
''Function ValidateTDS3ApportionedCheck() As Boolean
''ValidateTDS3ApportionedCheck = True
''TDS3ApportionedCheck = Sheet19.Range("TDS3b.DrpA").value
''
''If isdropdownblank(TDS3ApportionedCheck) Then
''TDS3ApportionedCheck = ""
''End If
''
''If Mid((TDS3ApportionedCheck), 1, 1) = "(" Or TDS3ApportionedCheck = "" Then
''    MsgTDS3 = MsgTDS3 + "Do you have TDS reported in your 26 AS check has to be filled" & Chr(13)
''    ValidateTDS3ApportionedCheck = False
''Exit Function
''End If
''
''End Function
''
''Function ValidateTDS3ClaimCheck() As Boolean
''ValidateTDS3ClaimCheck = True
''TDS3ClaimCheck = Sheet19.Range("TDS3b.DrpB").value
''
''If isdropdownblank(TDS3ClaimCheck) Then
''TDS3ClaimCheck = ""
''End If
''
''If Mid((TDS3ClaimCheck), 1, 1) = "(" Or TDS3ClaimCheck = "" Then
''    MsgTDS3 = MsgTDS3 + "Are you claiming TDS reported under other person's 26AS check has to be filled" & Chr(13)
''    ValidateTDS3ClaimCheck = False
''Exit Function
''End If
''
''End Function
''Sub UNLOCKTDS3()
''    Application.ScreenUpdating = False
''    Sheet19.Unprotect getmsgstate
''
''
''    Dim rngname_5_TDS3 As Variant
''    rngname_5_TDS3 = "TDS3.Reason||TDS3.DrpPAN ||TDS3.TDSothrPer||TDS3.PANOthrPer||Column6_value3"
''    UNLOCKRANGETDS3 (rngname_5_TDS3)
''
''    Sheet19.Range("TDS3.PANDetails").EntireRow.Hidden = False
''    Sheet19.Shapes("Button 30").Visible = msoTrue
''
''    ActiveSheet.Shapes("Button 30").Left = Range("SchTDS3.Button2").Left
''    ActiveSheet.Shapes("Button 30").Top = Range("SchTDS3.Button2").Top
''
''
''    Sheet19.Protect getmsgstate
''    Application.ScreenUpdating = True
''End Sub
''
''
''Sub LOCKTDS3()
''    Application.ScreenUpdating = False
''
''    Sheet19.Unprotect getmsgstate
''
''    Dim rngname_5_TDS3 As Variant
''    rngname_5_TDS3 = "TDS3.Reason||TDS3.DrpPAN ||TDS3.TDSothrPer||TDS3.PANOthrPer||Column6_value3"
''    LOCKRANGETDS3 (rngname_5_TDS3)
''
''    Sheet19.Range("TDS3.PANDetails").EntireRow.Hidden = True
''    Sheet19.Shapes("Button 30").Visible = msoFalse
''
''    ActiveSheet.Shapes("Button 30").Left = Range("SchTDS3.Button2").Left
''    ActiveSheet.Shapes("Button 30").Top = Range("SchTDS3.Button2").Top
''
''   Sheet19.Protect getmsgstate
''
''    Application.ScreenUpdating = True
''End Sub
''Sub LOCKRANGETDS3(rangenamestring As Variant)
''Dim i As Long
''Dim chCell As Range
''    rangenamestring = Split(rangenamestring, "||")
''        For i = 0 To UBound(rangenamestring)
''            For Each chCell In Sheet19.Range(rangenamestring(i))
''            chCell.Locked = True
''            chCell.Interior.Color = (&HD8D8D8)
''            chCell.ClearContents
''            Next
''        Next
''End Sub
''
''Sub UNLOCKRANGETDS3(rangenamestring As Variant)
''Dim i As Long
''Dim chCell As Range
''Dim addr1, addr2 As Variant
''    rangenamestring = Split(rangenamestring, "||")
''        For i = 0 To UBound(rangenamestring)
''            For Each chCell In Sheet19.Range(rangenamestring(i))
''            If rangenamestring(i) = "TDS3.Reason" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''
''
''            addr1 = Replace(chCell.AddressLocal, "L", "H")
''           addr2 = Replace(chCell.AddressLocal, "L", "K")
''            If Sheet19.Range(addr1).value > 0 Then
''            If Sheet19.Range(addr1).value = Sheet19.Range(addr2).value Then
''            Range(chCell.AddressLocal) = Replace(chCell.AddressLocal, "K", "L")
''            chCell.value = ""
''            chCell.Locked = True
''            chCell.Interior.Color = "&HD8D8D8"
''            End If
''
''            End If
''
''            ElseIf rangenamestring(i) = "TDS3b.Year" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''            ElseIf rangenamestring(i) = "TDS3b.HeadOfIncome" Then
''            chCell.Locked = False
''            chCell.Interior.Color = (&HCCFFCC)
''            chCell.value = "(Select)"
''            ElseIf rangenamestring(i) = "TDS3b.Reason" Then
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
''Sub UNLOCKTDS3b()
''
''Application.ScreenUpdating = False
''  Application.EnableEvents = False
'''----------------Unlock Password-------------------START---
''
''   Sheet19.Unprotect getmsgstate
'''----------------Unlock Password-------------------END-----
''Dim rngname_TDS3b As Variant
''
''rngname_TDS3b = "TDS3b.PAN||TDS3b.Name||TDS3b.UDN||TDS3b.Year||TDS3b.Income||TDS3b.TDS||TDS3b.PAN7||TDS3b.HeadOfIncome||TDS3b.Reason||TDS3b.IncomeOffered||TDS3b.TDSClaimedthisyear"
''UNLOCKRANGETDS3 (rngname_TDS3b)
''
''
''
''Sheet19.Range("TDS3b.hide").EntireRow.Hidden = False
''Sheet19.Shapes("Button 26").Visible = msoTrue
''
''ActiveSheet.Shapes("Button 26").Left = Range("SchTDS3.Button3").Left
''ActiveSheet.Shapes("Button 26").Top = Range("SchTDS3.Button3").Top
''
''
'''----------------Lock Password-------------------START---
''   Sheet19.Protect getmsgstate
'''----------------Lock Password-------------------END-----
''
'' Application.EnableEvents = True
'' Application.ScreenUpdating = True
''End Sub
''
'''Sub UNLOCKRANGEcg(rangenamestring As Variant)
'''Dim i As Integer
'''Dim chCell As Range
'''    rangenamestring = Split(rangenamestring, "||")
'''        For i = 0 To UBound(rangenamestring)
'''            For Each chCell In Sheet13.Range(rangenamestring(i))
'''            'chCell.Locked = False
'''            'chCell.Interior.Color = (&HCCFFCC)
'''            chCell.value = 0
'''            Next
'''        Next
'''End Sub
''
''Sub LOCKTDS3b()
''
''
''Application.ScreenUpdating = False
'' Application.EnableEvents = False
'''----------------Unlock Password-------------------START---
''
''   Sheet19.Unprotect getmsgstate
'''----------------Unlock Password-------------------END-----
''
''Dim rngname_TDS3b As Variant
''
''rngname_TDS3b = "TDS3b.PAN||TDS3b.Name||TDS3b.UDN||TDS3b.Year||TDS3b.Income||TDS3b.TDS||TDS3b.PAN7||TDS3b.HeadOfIncome||TDS3b.Reason||TDS3b.IncomeOffered||TDS3b.TDSClaimedthisyear"
''LOCKRANGETDS3 (rngname_TDS3b)
''
''
''
''Sheet19.Range("TDS3b.hide").EntireRow.Hidden = True
''Sheet19.Shapes("Button 26").Visible = msoFalse
''
''ActiveSheet.Shapes("Button 26").Left = Range("SchTDS3.Button3").Left
''ActiveSheet.Shapes("Button 26").Top = Range("SchTDS3.Button3").Top
''
''
'''----------------Lock Password-------------------START---
''   Sheet19.Protect getmsgstate
'''----------------Lock Password-------------------END-----
''
'' Application.EnableEvents = True
'' Application.ScreenUpdating = True
''End Sub
''
'''Sub LOCKRANGEcg(rangenamestring As Variant)
'''
'''Dim i As Integer
'''Dim chCell As Range
'''    rangenamestring = Split(rangenamestring, "||")
'''        For i = 0 To UBound(rangenamestring)
'''            For Each chCell In Sheet13.Range(rangenamestring(i))
'''            'chCell.Locked = True
'''            'chCell.Interior.Color = (&HD8D8D8)
'''            chCell.value = 0
'''            Next
'''        Next
'''End Sub
''Sub setTblinfo_TDS3()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDSoth2.PAN").Count
'' Set rangecells = Sheet19.Range("TDSoth2.PAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3 = ccount
'' rngname_TDS3 = "TDSoth2.PAN||TDSoth2.EmployerOrDeductorOrCollecterName||TDSoth2.UTN||TDSoth2.AmountSubjectDed||TDSoth2.DeductedYear||TDS3.DetailsOfRecAmt||TDS3.TaxDed||TDS3.AmountOffered||TDS3.HeadOfIncome||TDS3.TDSClaimed||TDS3.Reason||TDS3.TDSothrPer||TDS3.PANOthrPer||TDS3.PerShare||TDS3.IncomeOffered||TDS3.TDSClaimedthisyear||TDS3.IncBalCF||TDS3.TDSBalCF"
'' End Sub
''
'' Sub setTblinfo_TDS32()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Count
'' Set rangecells = Sheet19.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS32 = ccount
'' End Sub
''
''  Sub setTblinfo_TDS33()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDSoth2.UTN").Count
'' Set rangecells = Sheet19.Range("TDSoth2.UTN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS33 = ccount
'' End Sub
'''  Sub setTblinfo_TDS34()
''' Dim rangecells As Range
''' Dim mIntCells As Long
''' Dim mIntCtr As Long
''' Dim ccount As Long
''' ccount = 0
''' mIntCells = Sheet19.Range("TDSoth2.AmountSubjectDed").Count
''' Set rangecells = Sheet19.Range("TDSoth2.AmountSubjectDed").Cells
''' For mIntCtr = 1 To mIntCells
'''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
'''         ccount = ccount + 1
'''     End If
''' Next
''' end_TDS34 = ccount
''' End Sub
''   Sub setTblinfo_TDS35()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDSoth2.DeductedYear").Count
'' Set rangecells = Sheet19.Range("TDSoth2.DeductedYear").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS35 = ccount
'' End Sub
''
'' Sub setTblinfo_TDS36()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.DetailsOfRecAmt").Count
'' Set rangecells = Sheet19.Range("TDS3.DetailsOfRecAmt").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS36 = ccount
'' End Sub
''
''  Sub setTblinfo_TDS37()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.TaxDed").Count
'' Set rangecells = Sheet19.Range("TDS3.TaxDed").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS37 = ccount
'' End Sub
''
''
''  Sub setTblinfo_TDS38()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.AmountOffered").Count
'' Set rangecells = Sheet19.Range("TDS3.AmountOffered").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS38 = ccount
'' End Sub
''   Sub setTblinfo_TDS39()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.HeadOfIncome").Count
'' Set rangecells = Sheet19.Range("TDS3.HeadOfIncome").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS39 = ccount
'' End Sub
''   Sub setTblinfo_TDS310()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.TDSClaimed").Count
'' Set rangecells = Sheet19.Range("TDS3.TDSClaimed").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS310 = ccount
'' End Sub
''
''   Sub setTblinfo_TDS311()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.Reason").Count
'' Set rangecells = Sheet19.Range("TDS3.Reason").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS311 = ccount
'' End Sub
''    Sub setTblinfo_TDS3Inner()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.DrpPAN").Count
'' Set rangecells = Sheet19.Range("TDS3.DrpPAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3inner = ccount
'' End Sub
''   Sub setTblinfo_TDS312()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.TDSothrPer").Count
'' Set rangecells = Sheet19.Range("TDS3.TDSothrPer").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS312 = ccount
'' End Sub
''   Sub setTblinfo_TDS313()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.PANOthrPer").Count
'' Set rangecells = Sheet19.Range("TDS3.PANOthrPer").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS313 = ccount
'' End Sub
''   Sub setTblinfo_TDS314()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.PerShare").Count
'' Set rangecells = Sheet19.Range("TDS3.PerShare").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS314 = ccount
'' End Sub
''   Sub setTblinfo_TDS315()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.IncomeOffered").Count
'' Set rangecells = Sheet19.Range("TDS3.IncomeOffered").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS315 = ccount
'' End Sub
''
''   Sub setTblinfo_TDS316()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.TDSClaimedthisyear").Count
'' Set rangecells = Sheet19.Range("TDS3.TDSClaimedthisyear").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS316 = ccount
'' End Sub
''   Sub setTblinfo_TDS317()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.IncBalCF").Count
'' Set rangecells = Sheet19.Range("TDS3.IncBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS317 = ccount
'' End Sub
''   Sub setTblinfo_TDS318()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3.TDSBalCF").Count
'' Set rangecells = Sheet19.Range("TDS3.TDSBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS318 = ccount
'' End Sub
'' Sub setTblinfo_TDS3b()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.PAN").Count
'' Set rangecells = Sheet19.Range("TDS3b.PAN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b = ccount
'' rngname_TDS3b = "TDS3b.PAN||TDS3b.Name||TDS3b.UDN||TDS3b.Year||TDS3b.Income||TDS3b.TDS||TDS3b.PAN7||TDS3b.HeadOfIncome||TDS3b.Reason||TDS3b.IncomeOffered||TDS3b.TDSClaimedthisyear||TDS3b.IncBalCF||TDS3b.TDSBalCF"
'' End Sub
''  Sub setTblinfo_TDS3b2()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.Name").Count
'' Set rangecells = Sheet19.Range("TDS3b.Name").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b2 = ccount
'' End Sub
''   Sub setTblinfo_TDS3b3()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.UDN").Count
'' Set rangecells = Sheet19.Range("TDS3b.UDN").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b3 = ccount
''
'' End Sub
''    Sub setTblinfo_TDS3b4()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.Year").Count
'' Set rangecells = Sheet19.Range("TDS3b.Year").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b4 = ccount
''
'' End Sub
''     Sub setTblinfo_TDS3b5()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.Income").Count
'' Set rangecells = Sheet19.Range("TDS3b.Income").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b5 = ccount
'' End Sub
'' Sub setTblinfo_TDS3b6()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.TDS").Count
'' Set rangecells = Sheet19.Range("TDS3b.TDS").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b6 = ccount
''
'' End Sub
''  Sub setTblinfo_TDS3b7()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.PAN7").Count
'' Set rangecells = Sheet19.Range("TDS3b.PAN7").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b7 = ccount
''
'' End Sub
'' Sub setTblinfo_TDS3b8()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.HeadOfIncome").Count
'' Set rangecells = Sheet19.Range("TDS3b.HeadOfIncome").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b8 = ccount
''
'' End Sub
''  Sub setTblinfo_TDS3b9()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.Reason").Count
'' Set rangecells = Sheet19.Range("TDS3b.Reason").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b9 = ccount
''
'' End Sub
'' Sub setTblinfo_TDS3b10()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.IncomeOffered").Count
'' Set rangecells = Sheet19.Range("TDS3b.IncomeOffered").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b10 = ccount
''
'' End Sub
''
'' Sub setTblinfo_TDS3b11()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.TDSClaimedthisyear").Count
'' Set rangecells = Sheet19.Range("TDS3b.TDSClaimedthisyear").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b11 = ccount
''
'' End Sub
''  Sub setTblinfo_TDS3b12()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.IncBalCF").Count
'' Set rangecells = Sheet19.Range("TDS3b.IncBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b12 = ccount
''
'' End Sub
''   Sub setTblinfo_TDS3b13()
'' Dim rangecells As Range
'' Dim mIntCells As Long
'' Dim mIntCtr As Long
'' Dim ccount As Long
'' ccount = 0
'' mIntCells = Sheet19.Range("TDS3b.TDSBalCF").Count
'' Set rangecells = Sheet19.Range("TDS3b.TDSBalCF").Cells
'' For mIntCtr = 1 To mIntCells
''     If Not rangecells.item(mIntCtr).value = "" Then
''         ccount = ccount + 1
''     End If
'' Next
'' end_TDS3b13 = ccount
''
'' End Sub
''
''
''  Function ValidateSheetTDS3() As Boolean
''On Error Resume Next
''Dim cellcount, intcntr, rangname, name As Variant
''Dim SumTds, SumIft, tdssum, TDSvalue As Variant
''
''Dim k, j As Long
''
''ValidateSheetTDS3 = True
''    MsgTDS3 = "Schedule TDS3a." & Chr(13)
''    subProcCaption = "Validating TDS3"
''    setTblinfo_TDS3
''    noOfProcessSub = end_TDS3
''
''    If Len(Sheet19.Range("TDS3.Total").value) > 14 Then
''       MsgTDS3 = MsgTDS3 & " Total of TDS3 should not excced 14 digits " & Chr(13)
''       ValidateSheetTDS3 = False
''       Exit Function
''    End If
''
''
''
''
''For k = 1 To end_TDS3
''
''        If Not (Sheet19.Range("TDSoth2.PAN").Cells.item(k) <> "" Or Sheet19.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet19.Range("TDSoth2.DeductedYear").Cells.item(k)) Or Sheet19.Range("TDS3.DetailsOfRecAmt").Cells.item(k) <> "" Or Sheet19.Range("TDS3.TaxDed").Cells.item(k) <> "" Or (Sheet19.Range("TDS3.TDSClaimedthisyear").Cells.item(k) <> "")) Then
''            MsgTDS3 = MsgTDS3 & " At Sr. No " & k & " Please fill all the Mandatory Fields " & Chr(13)
''            ValidateSheetTDS3 = False
''            Exit Function
''        End If
''
''    UpdateProgressBar
''Next
''
''
''If end_TDS3 > 0 Or end_TDS3 = 0 Then
''
''
''    If Not ValidatePAN_TDS3() Then ValidateSheetTDS3 = False
''    If Not ValidateEmployerOrDeductorOrCollecterName_TDS3() Then ValidateSheetTDS3 = False
''    If Not ValidateUTN_TDS3() Then ValidateSheetTDS3 = False
''    If Not ValidateDeductedYear_TDS3() Then ValidateSheetTDS3 = False
''    If Not ValidateDetailsOfReceipts_TDS3() Then ValidateSheetTDS3 = False
''    'If Not ValidateAmtDed_TDS3() Then ValidateSheetTDS3 = False
''        If Not ValidateTaxDeducted26AS_TDS3() Then ValidateSheetTDS3 = False
''        If Not ValidateAmtOfferedSelf_TDS3() Then ValidateSheetTDS3 = False
''        If Not ValidateTDSClaimedInOwnHands_TDS3() Then ValidateSheetTDS3 = False
''        If Not ValidateTDSApportOtherHands_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateIncomeOfferedThisYear_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateTDSClaimedThisYear_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateIncomeCarriedFwd_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateTDSCarriedFwd_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateHeadOfIncome_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateReasonApportTDS_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidatePANOfOtherPerson_TDS3 Then ValidateSheetTDS3 = False
''        'If Not ValidatePercentOfShare_TDS3 Then ValidateSheetTDS3 = False
''        If Not ValidateTANdetails3 Then ValidateSheetTDS3 = False
''    End If
''
''
''     setTblinfo_TDS3
''     setTblinfo_TDS3Inner
''    setTblinfo_TDS32
''    setTblinfo_TDS33
''    'setTblinfo_TDS34
''    setTblinfo_TDS35
''    setTblinfo_TDS36
''    setTblinfo_TDS37
''    setTblinfo_TDS38
''    setTblinfo_TDS39
''    setTblinfo_TDS310
''    setTblinfo_TDS311
''    setTblinfo_TDS312
''    setTblinfo_TDS313
''    setTblinfo_TDS314
''    setTblinfo_TDS315
''    setTblinfo_TDS316
''    setTblinfo_TDS317
''    setTblinfo_TDS318
''
''       'setTblinfo_TDS2inner
''
''
''           If Mid(Range("TDS3b.DrpA").Cells.value, 1, 1) = "Y" Then
''           If end_TDS3 = 0 Or end_TDS32 = 0 Or end_TDS35 = 0 Or end_TDS36 = 0 Or end_TDS37 = 0 Or end_TDS38 = 0 Or end_TDS39 = 0 Or end_TDS310 = 0 Or end_TDS315 = 0 Or end_TDS316 = 0 Then
''              MsgTDS3 = MsgTDS3 + "Please fill all the Mandatory Fields " & Chr(13)
''                 ValidateSheetTDS3 = False
''            End If
''            End If
''
''
''            If ((end_TDS3 <> end_TDS32) Or (end_TDS3 <> end_TDS35) Or (end_TDS3 <> end_TDS36) Or (end_TDS3 <> end_TDS37) Or (end_TDS3 <> end_TDS38) Or (end_TDS3 <> end_TDS39) Or (end_TDS3 <> end_TDS310) Or (end_TDS3 <> end_TDS315) Or (end_TDS3 <> end_TDS316)) Then
''                MsgTDS3 = MsgTDS3 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS3 = False
''            End If
''
''
''
''            For k = 1 To end_TDS3
''            TDSvalue = 0
''        For j = 1 To end_TDS3inner
''        SumTds = Mid(Range("TDS3.DrpPAN").Cells.item(j).value, 3, 10)
''        SumIft = Range("TDSoth2.PAN").Cells.item(k).value
''        If (SumTds = SumIft) And CStr(Range("TDS2ii.SlNo").Cells.item(k).value) = CStr(Mid(Range("TDS3.DrpPAN").Cells.item(j).value, 1, 1)) Then
''         TDSvalue = Range("TDS3.TaxDed").Cells.item(k).value
''         End If
''         If CStr(Range("TDS2ii.SlNo").Cells.item(k).value) = CStr(Mid(Range("TDS3.DrpPAN").Cells.item(j).value, 1, 1)) Then
''         Range("Column6_value3").Cells.item(j).value = Round((Range("TDS3.TDSothrPer").Cells.item(j).value / TDSvalue) * 100, 2)
''         End If
''        Next
''        Next
''
''        If end_TDS3 > 0 Or end_TDS3 = 0 Then
''        If Not ValidatePercentOfShare_TDS3 Then ValidateSheetTDS3 = False
''       End If
''
''
''
''            If Mid(Range("TDS3b.DrpA").Cells.value, 1, 1) = "Y" Then
''             If end_TDS3 > 0 Then
''           ' If (end_TDS3inner <> end_TDS313) Or (end_TDS3inner <> end_TDS314) Or (end_TDS3inner <> end_TDS312) Then
''             If (end_TDS3inner <> end_TDS313) Or (end_TDS3inner <> end_TDS312) Then
''                    MsgTDS3 = MsgTDS3 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS3 = False
''            End If
''            End If
''            End If
''
''            If Mid(Range("TDS3b.DrpA").Cells.value, 1, 1) = "Y" Then
''             If end_TDS311 > 0 Then
''            'If (end_TDS3inner = 0) Or (end_TDS313 = 0) Or (end_TDS314 = 0) Or (end_TDS312 = 0) Then
''            If (end_TDS3inner = 0) Or (end_TDS313 = 0) Or (end_TDS312 = 0) Then
''                    MsgTDS3 = MsgTDS3 + "Please fill all the Mandatory Fields " & Chr(13)
''                ValidateSheetTDS3 = False
''            End If
''            End If
''            End If
''
''
''
''
''    For k = 1 To end_TDS3
''     tdssum = 0
''     For j = 1 To end_TDS3inner
''     SumTds = Mid(Range("TDS3.DrpPAN").Cells.item(j).value, 3, 10)
''     SumIft = Range("TDSoth2.PAN").Cells.item(k).value
''     If (SumTds = SumIft) And CStr(Range("TDS2ii.SlNo").Cells.item(k).value) = CStr(Mid(Range("TDS3.DrpPAN").Cells.item(j).value, 1, 1)) Then
''     tdssum = WorksheetFunction.Sum(Range("TDS3.TDSothrPer").Cells.item(j).value, tdssum)
''     End If
''     Next
''     If (Range("TDS3.TDSClaimed").item(k).value + tdssum <> (Range("TDS3.TaxDed").item(k).value)) Then
''            MsgTDS3 = MsgTDS3 + "For TAN At Sr. No " & k & Chr(13) & "Amount at Tax  Deducted as per 26AS [Col (6)] should be equal to sum of amount at TDS claimed out of (6) in own hands [Col (9)] and amount at TDS out of (6) being apportioned in the hands of other person [Col (11)]" & Chr(13)
''            ValidateSheetTDS3 = False
''            Exit Function
''            End If
''
''     Next
''
''
''
''
'''If Not checkforNonMandatoryCol_TDS2 Then ValidateSheetTDS2 = False
''
''
''End Function
''
''Function ValidatePAN_TDS3() As Boolean
''ValidatePAN_TDS3 = True
''setTblinfo_TDS3
''setTblinfo_TDS3Inner
''noOfProcessSub = end_TDS3
''
''Dim rangecells As Range
''Dim rangecells1 As Range
''Set rangecells = Range("TDSoth2.PAN").Cells
''Set rangecells1 = Range("TDS3.DrpPAN").Cells
''Dim i As Long
'' ReDim PAN_TDS3(end_TDS3)
'' ReDim PAN2_TDS3(end_TDS3inner)
''
''For i = 1 To end_TDS3
''    PAN_TDS3(i) = rangecells.item(i).value
''    If Not Len(PAN_TDS3(i)) = 0 Then
''        If Not mdIncomeDetails.CheckPAN_Tenant(UCase(Mid(PAN_TDS3(i), 1, 10))) Then
''        MsgTDS3 = MsgTDS3 + "PAN of the Tenant at Sr. No  " & i & "  in Sheet TDS3  is invalid. First 5 alphabets, next 4 digits, then alphabet (4th alphabet must be ""P"" or ""H"") " & Chr(13)
''        ValidatePAN_TDS3 = False
''        End If
''    End If
''
''           If PAN_TDS3(i) = Sheet1.Range("sheet1.PAN").value Then
''                MsgTDS3 = MsgTDS3 + "PAN of Tenant at Sr.No " & i & " can not be same as PAN of user  " & Chr(13)
''                ValidatePAN_TDS3 = False
''                Exit Function
''            End If
''
''    UpdateProgressBar
''Next
''
''For i = 1 To end_TDS3inner
''    PAN2_TDS3(i) = UCase(rangecells1.item(i).value)
''Next
''
''End Function
''Function ValidateEmployerOrDeductorOrCollecterName_TDS3() As Boolean
''
''    ValidateEmployerOrDeductorOrCollecterName_TDS3 = True
''    setTblinfo_TDS2
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet19.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells
''    ReDim NameOfTeanant_TDS3(end_TDS3)
''    For i = 1 To end_TDS3
''        NameOfTeanant_TDS3(i) = rangecells.item(i).value
''
''        If NameOfTeanant_TDS3(i) = "" Then
''            MsgTDS3 = MsgTDS3 + "Name of Employer at Sr. No  " & i & "  is Mandatory" & Chr(13)
''            ValidateEmployerOrDeductorOrCollecterName_TDS3 = False
''            Exit Function
''        End If
''
''
''            If Len(NameOfTeanant_TDS3(i)) > 125 Then
''            MsgTDS3 = MsgTDS3 + "Name of Employer at Sr. No  " & i & "  in Sheet TDS  Cannot be more than 125 Charachters" & Chr(13)
''            ValidateEmployerOrDeductorOrCollecterName_TDS3 = False
''            Exit Function
''        End If
''
''    Next
''End Function
''
''Function ValidateUTN_TDS3() As Boolean
''
''    ValidateUTN_TDS3 = True
''    setTblinfo_TDS3
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet19.Range("TDSoth2.UTN").Cells
''    ReDim UTN_TDS3(end_TDS3)
''    For i = 1 To end_TDS3
''        UTN_TDS3(i) = rangecells.item(i).value
''    If Not checkfieldspecialcharacter(UTN_TDS3(i)) Then
''         MsgTDS3 = MsgTDS3 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
''          ValidateUTN_TDS3 = False
''          Exit Function
''     End If
''
''If InStr(UTN_TDS3(i), ".") > 0 Then
''         MsgTDS3 = MsgTDS3 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
''          ValidateUTN_TDS3 = False
''          Exit Function
''    End If
''
''    If InStr(UTN_TDS3(i), "-") > 0 Then
''         MsgTDS3 = MsgTDS3 + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
''          ValidateUTN_TDS3 = False
''          Exit Function
''    End If
''
''    If UTN_TDS3(i) <> "" Then
''    If Len(UTN_TDS3(i)) < 6 Or Len(UTN_TDS3(i)) > 8 Or Len(UTN_TDS3(i)) <> 8 Then
''          MsgTDS3 = MsgTDS3 + "Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
''          ValidateUTN_TDS3 = False
''          Exit Function
''    End If
''    End If
''
''
'' Next
''End Function
''Function ValidateDeductedYear_TDS3() As Boolean
''
''    ValidateDeductedYear_TDS3 = True
''    setTblinfo_TDS3
''    Dim i As Long
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDSoth2.DeductedYear").Cells
''    ReDim DeductedYear_TDS3(end_TDS3)
''For i = 1 To end_TDS3
''        DeductedYear_TDS3(i) = rangecells.item(i).value
''
''        If isdropdownblank(DeductedYear_TDS3(i)) Then
''            MsgTDS3 = MsgTDS3 + " Year of tax Deduction at Sr. No  " & i & " is Mandatory" & Chr(13)
''            ValidateDeductedYear_TDS3 = False
''            Exit Function
''        End If
''
''
''     If Not checkfieldspecialcharacter(DeductedYear_TDS3(i)) Then
''         MsgTDS3 = MsgTDS3 + "Year of tax  Deduction at Sr. No  " & i & " in Sheet TDS schedule TDS2 is invalid" & Chr(13)
''          ValidateDeductedYear_TDS3 = False
''          Exit Function
''     End If
''     If DeductedYear_TDS3(i) > 2017 Then
''          MsgTDS3 = MsgTDS3 + "Year of tax  Deduction at Sr. No  " & i & " in Sheet TDS cannot be greater than 2017 " & Chr(13)
''          ValidateDeductedYear_TDS3 = False
''          Exit Function
''    End If
''Next
''End Function
'''Function ValidateAmtDed_TDS3() As Boolean
'''    ValidateAmtDed_TDS3 = True
'''    setTblinfo_TDS3
'''
'''    Dim i As Long
'''    Dim rangecells As Range
'''
'''    ReDim AmtDed_TDS3(end_TDS3)
'''    Set rangecells = Sheet19.Range("TDSoth2.AmountSubjectDed").Cells
'''
'''    For i = 1 To end_TDS3
'''        AmtDed_TDS3(i) = rangecells.item(i).value
'''
'''    If Not chkCompulsory(AmtDed_TDS3(i)) And rangecells.item(i).Locked = False Then
'''         MsgTDS3 = MsgTDS3 + "Details Of Receipts at Sr. No  " & i & " is Mandatory." & Chr(13)
'''         ValidateAmtDed_TDS3 = False
'''         Exit Function
'''    End If
'''
'''    If Len(AmtDed_TDS3(i)) > 14 Then
'''        MsgTDS3 = MsgTDS3 + "Details Of Receipts at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
'''        ValidateAmtDed_TDS3 = False
'''    End If
'''
'''    Next
'''End Function
''Function ValidateDetailsOfReceipts_TDS3() As Boolean
''    ValidateDetailsOfReceipts_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim DetailsOfReceipts_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.DetailsOfRecAmt").Cells
''
''    For i = 1 To end_TDS3
''        DetailsOfReceipts_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(DetailsOfReceipts_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "Details Of Receipts at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateDetailsOfReceipts_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(DetailsOfReceipts_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "Details Of Receipts at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateDetailsOfReceipts_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTaxDeducted26AS_TDS3() As Boolean
''    ValidateTaxDeducted26AS_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TaxDeducted26AS_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.TaxDed").Cells
''
''    For i = 1 To end_TDS3
''        TaxDeducted26AS_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TaxDeducted26AS_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "Tax Deducted as per 26AS at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTaxDeducted26AS_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(TaxDeducted26AS_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "Tax Deducted as per 26AS at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTaxDeducted26AS_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateAmtOfferedSelf_TDS3() As Boolean
''    ValidateAmtOfferedSelf_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim AmtOfferedSelf_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.AmountOffered").Cells
''
''    For i = 1 To end_TDS3
''        AmtOfferedSelf_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(AmtOfferedSelf_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "Amount offered for Tax at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateAmtOfferedSelf_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(AmtOfferedSelf_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "Amount offered for Tax at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateAmtOfferedSelf_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSClaimedInOwnHands_TDS3() As Boolean
''    ValidateTDSClaimedInOwnHands_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSClaimedInOwnHands_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.TDSClaimed").Cells
''
''    For i = 1 To end_TDS3
''        TDSClaimedInOwnHands_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSClaimedInOwnHands_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "TDS Claimed in Own Hands at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSClaimedInOwnHands_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(TDSClaimedInOwnHands_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "TDS Claimed in Own Hands at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSClaimedInOwnHands_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSApportOtherHands_TDS3() As Boolean
''    ValidateTDSApportOtherHands_TDS3 = True
''    setTblinfo_TDS3Inner
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSApportOtherHands_TDS3(end_TDS3inner)
''    Set rangecells = Sheet19.Range("TDS3.TDSothrPer").Cells
''
''    For i = 1 To end_TDS3inner
''        TDSApportOtherHands_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSApportOtherHands_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "TDS apportioned in the hands of other person at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSApportOtherHands_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(TDSApportOtherHands_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "TDS apportioned in the hands of other person at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSApportOtherHands_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeOfferedThisYear_TDS3() As Boolean
''    ValidateIncomeOfferedThisYear_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeOfferedThisYear_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.IncomeOffered").Cells
''
''    For i = 1 To end_TDS3
''        IncomeOfferedThisYear_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeOfferedThisYear_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "Income Offered This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeOfferedThisYear_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(IncomeOfferedThisYear_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "Income Offered This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeOfferedThisYear_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSClaimedThisYear_TDS3() As Boolean
''    ValidateTDSClaimedThisYear_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSClaimedThisYear_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.TDSClaimedthisyear").Cells
''
''    For i = 1 To end_TDS3
''        TDSClaimedThisYear_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSClaimedThisYear_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "TDS Claimed This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSClaimedThisYear_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(TDSClaimedThisYear_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "TDS Claimed This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSClaimedThisYear_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeCarriedFwd_TDS3() As Boolean
''    ValidateIncomeCarriedFwd_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeCarriedFwd_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.IncBalCF").Cells
''
''    For i = 1 To end_TDS3
''        IncomeCarriedFwd_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeCarriedFwd_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "Income Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeCarriedFwd_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(IncomeCarriedFwd_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "Income Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeCarriedFwd_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSCarriedFwd_TDS3() As Boolean
''    ValidateTDSCarriedFwd_TDS3 = True
''    setTblinfo_TDS3
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSCarriedFwd_TDS3(end_TDS3)
''    Set rangecells = Sheet19.Range("TDS3.TDSBalCF").Cells
''
''    For i = 1 To end_TDS3
''        TDSCarriedFwd_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSCarriedFwd_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "TDS Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSCarriedFwd_TDS3 = False
''         Exit Function
''    End If
''
''    If Len(TDSCarriedFwd_TDS3(i)) > 14 Then
''        MsgTDS3 = MsgTDS3 + "TDS Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSCarriedFwd_TDS3 = False
''    End If
''
''    Next
''End Function
''
''Function ValidateHeadOfIncome_TDS3() As Boolean
''
''    ValidateHeadOfIncome_TDS3 = True
''    setTblinfo_TDS3
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3.HeadOfIncome").Cells
''    ReDim HeadOfIncome_TDS3(end_TDS3)
''    For i = 1 To end_TDS3
''        HeadOfIncome_TDS3(i) = rangecells.item(i).value
''        If isdropdownblank(HeadOfIncome_TDS3(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS3 = MsgTDS3 + "Head Of Income at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateHeadOfIncome_TDS3 = False
''              Exit Function
''        End If
''
''        If Mid(HeadOfIncome_TDS3(i), 1, 1) = "H" Then HeadOfIncome_TDS3(i) = "HP"
''        If Mid(HeadOfIncome_TDS3(i), 1, 1) = "O" Then HeadOfIncome_TDS3(i) = "OS"
''
'''        If Not checkfieldspecialcharacter(HeadOfIncome_TDS3(i)) Then
'''              MsgTDS3 = MsgTDS3 + "Special characters are not allowed in Head Of Income at Sr. No  " & i & " ." & Chr(13)
'''              ValidateHeadOfIncome_TDS3 = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidateReasonApportTDS_TDS3() As Boolean
''
''    ValidateReasonApportTDS_TDS3 = True
''    setTblinfo_TDS3
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3.Reason").Cells
''    ReDim ReasonApportTDS_TDS3(end_TDS3)
''    For i = 1 To end_TDS3
''        ReasonApportTDS_TDS3(i) = rangecells.item(i).value
''        If isdropdownblank(ReasonApportTDS_TDS3(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS3 = MsgTDS3 + "Reason for Apportition at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateReasonApportTDS_TDS3 = False
''              Exit Function
''        End If
''
''        If Mid(ReasonApportTDS_TDS3(i), 1, 1) = "P" Then ReasonApportTDS_TDS3(i) = "PCC"
''        If Mid(ReasonApportTDS_TDS3(i), 1, 1) = "C" Then ReasonApportTDS_TDS3(i) = "COA"
'''        If Mid(ReasonApportTDS_TDS3(i), 1, 1) = "C" Then ReasonApportTDS_TDS3(i) = "SIC"
'''        If Mid(ReasonApportTDS_TDS3(i), 1, 1) = "D" Then ReasonApportTDS_TDS3(i) = "TDP"
'''        If Mid(ReasonApportTDS_TDS3(i), 1, 1) = "E" Then ReasonApportTDS_TDS3(i) = "DPI"
''
'''        If Not checkfieldspecialcharacter(ReasonApportTDS_TDS3(i)) Then
'''              MsgTDS3 = MsgTDS3 + "Special characters are not allowed in Reason for Apportition at Sr. No  " & i & " ." & Chr(13)
'''              ValidateReasonApportTDS_TDS3 = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidateTANdetails3() As Boolean
''
''    ValidateTANdetails3 = True
''    setTblinfo_TDS312
''    setTblinfo_TDS3Inner
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3.DrpPAN").Cells
''    ReDim TANdetails_TDS3(end_TDS3inner)
''    If end_TDS312 > 0 Then
''    For i = 1 To end_TDS3inner
''        TANdetails_TDS3(i) = rangecells.item(i).value
''
''        If Len(TANdetails_TDS3(i)) < 12 Or TANdetails_TDS3(i) = "" Then 'And rangecells.item(i).Locked = False Then
''            MsgTDS3 = MsgTDS3 + "Please select valid PAN From the dropdown at Sr. No  " & i & " ." & Chr(13)
''              ValidateTANdetails3 = False
''              Exit Function
''        End If
''    Next
''
''    End If
''End Function
''
''Function ValidatePANOfOtherPerson_TDS3() As Boolean
''ValidatePANOfOtherPerson_TDS3 = True
''setTblinfo_TDS3
''setTblinfo_TDS3Inner
''Dim rangecells As Range
''Set rangecells = Sheet19.Range("TDS3.PANOthrPer").Cells
''Dim i As Long
''ReDim PANOfOtherPerson_TDS3(end_TDS3inner)
''For i = 1 To end_TDS3inner
''PANOfOtherPerson_TDS3(i) = rangecells.item(i).value
''
''   If Not chkCompulsory(PANOfOtherPerson_TDS3(i)) And rangecells.item(i).Locked = False Then
''        MsgTDS3 = MsgTDS3 + "PAN at Sr.No " & i & " is Mandatory." & Chr(13)
''         ValidatePANOfOtherPerson_TDS3 = False
''         Exit Function
''  End If
''
''    If Not checkfieldspecialcharacter(PANOfOtherPerson_TDS3(i)) Then
''         MsgTDS3 = MsgTDS3 + "PAN at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
''         ValidatePANOfOtherPerson_TDS3 = False
''         Exit Function
''    End If
''
''    If Not CheckPAN(UCase(PANOfOtherPerson_TDS3(i))) Then
''         MsgTDS3 = MsgTDS3 + "PAN at Sr. No  " & i & " in Sheet TDS is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
''         ValidatePANOfOtherPerson_TDS3 = False
''         Exit Function
''    End If
''
''    If PANOfOtherPerson_TDS3(i) = Sheet1.Range("sheet1.PAN").value Then
''                MsgTDS3 = MsgTDS3 + "PAN (on which TDS deducted) Sr.No " & i & " can not be same as PAN of user  " & Chr(13)
''                ValidatePANOfOtherPerson_TDS3 = False
''                Exit Function
''    End If
''
''Next
''End Function
''
''Function ValidatePercentOfShare_TDS3() As Boolean
''    ValidatePercentOfShare_TDS3 = True
''    setTblinfo_TDS3Inner
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim PercentOfShare_TDS3(end_TDS3inner)
''    Set rangecells = Sheet19.Range("TDS3.PerShare").Cells
''
''    For i = 1 To end_TDS3inner
''        PercentOfShare_TDS3(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(PercentOfShare_TDS3(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3 = MsgTDS3 + "Percent Of Share at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidatePercentOfShare_TDS3 = False
''         Exit Function
''    End If
''
'''    If Len(PercentOfShare_TDS3(i)) > 5 Then
'''        MsgTDS2 = MsgTDS3 + "Percent Of Share at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
'''        ValidatePercentOfShare_TDS3 = False
'''    End If
''
''    Next
''End Function
''
''
'' Function ValidatesheetTDS3b() As Boolean
''    ValidatesheetTDS3b = True
''    MsgTDS2b = "Schedule TDS3b :" & Chr(10)
''
''    If Len(Sheet19.Range("TDS3b.Total").value) > 14 Then
''       MsgTDS3b = MsgTDS3b & " Total of TDS3b should not excced 14 digits " & Chr(13)
''       ValidatesheetTDS3b = False
''       Exit Function
''    End If
''
''    setTblinfo_TDS3b
''    setTblinfo_TDS3b2
''    setTblinfo_TDS3b3
''    setTblinfo_TDS3b4
''    setTblinfo_TDS3b5
''    setTblinfo_TDS3b6
''    setTblinfo_TDS3b7
''    setTblinfo_TDS3b8
''    setTblinfo_TDS3b9
''    setTblinfo_TDS3b10
''    setTblinfo_TDS3b11
''    setTblinfo_TDS3b12
''    setTblinfo_TDS3b13
''
''    For i = 1 To end_TDS3b
''        If Len(Sheet19.Range("TDS3b.PAN").item(i).value) = 0 Then
''            MsgTDS3b = MsgTDS3b + "Please fill all the mandatory fields of the row " & i & " before filling the next row." & Chr(13)
''            ValidatesheetTDS3b = False
''            Exit Function
''        End If
''    Next
''
''    If Not ValidatePAN_TDS3b() Then ValidatesheetTDS3b = False
''    If (Len(Sheet19.Range("TDS3b.PAN").item(1).value) > 0) Then
''        If Not ValidateEmployerOrDeductorOrCollecterName_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateUTN_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateDeductedYear_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateIncomeApportioned_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateTDSApportioned_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateIncomeOfferedThisYear_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateTDSClaimedThisYear_TDS3b() Then ValidatesheetTDS3b = False
''        If Not ValidateIncomeCarriedFwd_TDS3b Then ValidatesheetTDS3b = False
''        If Not ValidateTDSCarriedFwd_TDS3b Then ValidatesheetTDS3b = False
''        If Not ValidateHeadOfIncome_TDS3b Then ValidatesheetTDS3b = False
''        If Not ValidateReasonClaimOwnHands_TDS3b Then ValidatesheetTDS3b = False
''        If Not ValidateTDSDeductedPAN_TDS3b Then ValidatesheetTDS3b = False
''    End If
''
''
''    If Mid(Range("TDS3b.DrpB").Cells.value, 1, 1) = "Y" Then
''           If end_TDS3b = 0 Or end_TDS3b2 = 0 Or end_TDS3b4 = 0 Or end_TDS3b5 = 0 Or end_TDS3b6 = 0 Or end_TDS3b7 = 0 Or end_TDS3b8 = 0 Or end_TDS3b9 = 0 Or end_TDS3b10 = 0 Or end_TDS3b11 = 0 Or end_TDS3b12 = 0 Or end_TDS3b13 = 0 Then
''              MsgTDS3b = MsgTDS3b + "Please fill all the Mandatory Fields " & Chr(13)
''                 ValidatesheetTDS3b = False
''            End If
''            End If
''
''
''    For i = 1 To Sheet19.Range("TDS3b.PAN").Count
''        If Len(Sheet19.Range("TDS3b.UDN").item(i).value) > 0 Then
''
''            If Len(Sheet19.Range("TDS3b.UDN").item(i).value) = 0 Then
''                 MsgTDS3b = MsgTDS3b + "Enter All Mandatory Fields at Sr. No " & i & " ." & Chr(13)
''                ValidatesheetTDS3b = False
''                Exit Function
''            End If
''        End If
''    Next
''
''End Function
''
''Function ValidatePAN_TDS3b() As Boolean
''    ValidatePAN_TDS3b = True
''    setTblinfo_TDS3b
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3b.PAN").Cells
''    ReDim PAN_TDS3b(end_TDS3b)
''
''    subProcCaption = "Validating TDS3b"
''    noOfProcessSub = end_TDS3b
''
''    For i = 1 To end_TDS3b
''        PAN_TDS3b(i) = UCase(rangecells.item(i).value)
'' If Not Len(PAN_TDS3b(i)) = 0 Then
''
''     If Not ValidateTantype_text(Mid(PAN_TDS3b(i), 1, 5)) Then
''         MsgTDS3b = MsgTDS3b + "PAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet." & Chr(13)
''         ValidatePAN_TDS3b = False
''         Exit Function
''     End If
''
''     If Not IsNumeric(Mid(PAN_TDS3b(i), 6, 4)) Then
''         MsgTDS3b = MsgTDS3b + "PAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet." & Chr(13)
''         ValidatePAN_TDS3b = False
''         Exit Function
''     End If
''
''     If Not ValidateTantype_text(Right(PAN_TDS3b(i), 1)) Then
''         MsgTDS3b = MsgTDS3b + "PAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet." & Chr(13)
''         ValidatePAN_TDS3b = False
''         Exit Function
''     End If
''
'' ElseIf Not chkCompulsory(PAN_TDS3b(i)) Then
''         MsgTDS3b = MsgTDS3b + "PAN at Sr. No  " & i & "  is Mandatory." & Chr(13)
''     ValidatePAN_TDS3b = False
''     Exit Function
'' End If
''
''      If PAN_TDS3b(i) = Sheet1.Range("sheet1.PAN").value Then
''                MsgTDS3b = MsgTDS3b + "PAN of Tenant Sr.No " & i & " can not be same as PAN of user  " & Chr(13)
''                ValidatePAN_TDS3b = False
''                Exit Function
''      End If
''
''
''
'' UpdateProgressBar
'' Next
''End Function
''
''Function ValidateEmployerOrDeductorOrCollecterName_TDS3b() As Boolean
''
''    ValidateEmployerOrDeductorOrCollecterName_TDS3b = True
''    setTblinfo_TDS3b
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3b.Name").Cells
''    ReDim Name_TDS3b(end_TDS3b)
''    For i = 1 To end_TDS3b
''        Name_TDS3b(i) = rangecells.item(i).value
''     If Not chkCompulsory(Name_TDS3b(i)) Then
''         MsgTDS3b = MsgTDS3b + "EmployerOrTenant Name at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateEmployerOrDeductorOrCollecterName_TDS3b = False
''         Exit Function
''     End If
'' Next
''End Function
''
''Function ValidateUTN_TDS3b() As Boolean
''
''    ValidateUTN_TDS3b = True
''    setTblinfo_TDS3b
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3b.UDN").Cells
''    ReDim UTN_TDS3b(end_TDS3b)
''    For i = 1 To end_TDS3b
''        UTN_TDS3b(i) = rangecells.item(i).value
''
''    If Not checkfieldspecialcharacter(UTN_TDS3b(i)) Then
''         MsgTDS3b = MsgTDS3b + "Special characters are not allowed in Unique TDS Certificate Number at Sr. No  " & i & " ." & Chr(13)
''          ValidateUTN_TDS3b = False
''          Exit Function
''     End If
''
''If InStr(UTN_TDS3b(i), ".") > 0 Then
''         MsgTDS3b = MsgTDS3b + "Decimals are not allowed in Unique TDS Certificate Number at Sr. No  " & i & " ." & Chr(13)
''          ValidateUTN_TDS3b = False
''          Exit Function
''End If
''
'' If UTN_TDS3b(i) <> "" Then
''If Len(UTN_TDS3b(i)) <> 8 Then
''MsgTDS3b = MsgTDS3b + "Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
''ValidateUTN_TDS3b = False
''Exit Function
''End If
''End If
''
''
''    If UTN_TDS3b(i) <> "" Then
''    If Len(UTN_TDS3b(i)) = 6 Then
''        UTN_TDS3b(i) = "00" & UTN_TDS3b(i)
''     End If
''    If Len(UTN_TDS3b(i)) = 7 Then
''        UTN_TDS3b(i) = "0" & UTN_TDS3b(i)
''     End If
''     End If
''
'' Next
''End Function
''
''Function ValidateDeductedYear_TDS3b() As Boolean
''
''    ValidateDeductedYear_TDS3b = True
''    setTblinfo_TDS3b
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3b.Year").Cells
''    ReDim DeductedYear_TDS3b(end_TDS3b)
''    For i = 1 To end_TDS3b
''        DeductedYear_TDS3b(i) = rangecells.item(i).value
''        If isdropdownblank(DeductedYear_TDS3b(i)) Then
''            DeductedYear_TDS3b(i) = ""
''        End If
''        If Not checkfieldspecialcharacter(DeductedYear_TDS3b(i)) Then
''              MsgTDS3b = MsgTDS3b + "Special characters are not allowed in Deducted Year at Sr. No  " & i & " ." & Chr(13)
''              ValidateDeductedYear_TDS3b = False
''              Exit Function
''        End If
'' Next
''End Function
''
''Function ValidateIncomeApportioned_TDS3b() As Boolean
''    ValidateIncomeApportioned_TDS3b = True
''    setTblinfo_TDS3b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeApportioned_TDS3b(end_TDS3b)
''    Set rangecells = Sheet19.Range("TDS3b.Income").Cells
''
''    For i = 1 To end_TDS3b
''        IncomeApportioned_TDS3b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeApportioned_TDS3b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3b = MsgTDS3b + "Income Apportioned at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeApportioned_TDS3b = False
''         Exit Function
''    End If
''
''    If Len(IncomeApportioned_TDS3b(i)) > 14 Then
''        MsgTDS3b = MsgTDS3b + "Income Apportioned at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeApportioned_TDS3b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSApportioned_TDS3b() As Boolean
''    ValidateTDSApportioned_TDS3b = True
''    setTblinfo_TDS3b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSApportioned_TDS3b(end_TDS3b)
''    Set rangecells = Sheet19.Range("TDS3b.TDS").Cells
''
''    For i = 1 To end_TDS3b
''        TDSApportioned_TDS3b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSApportioned_TDS3b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3b = MsgTDS3b + "TDS Apportioned at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSApportioned_TDS3b = False
''         Exit Function
''    End If
''
''    If Len(TDSApportioned_TDS3b(i)) > 14 Then
''        MsgTDS3b = MsgTDS3b + "TDS Apportioned at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSApportioned_TDS3b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeOfferedThisYear_TDS3b() As Boolean
''    ValidateIncomeOfferedThisYear_TDS3b = True
''    setTblinfo_TDS3b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeOfferedThisYear_TDS3b(end_TDS3b)
''    Set rangecells = Sheet19.Range("TDS3b.IncomeOffered").Cells
''
''    For i = 1 To end_TDS3b
''        IncomeOfferedThisYear_TDS3b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeOfferedThisYear_TDS3b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3b = MsgTDS3b + "Income Offered This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeOfferedThisYear_TDS3b = False
''         Exit Function
''    End If
''
''    If Len(IncomeOfferedThisYear_TDS3b(i)) > 14 Then
''        MsgTDS3b = MsgTDS3b + "Income Offered This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeOfferedThisYear_TDS3b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSClaimedThisYear_TDS3b() As Boolean
''    ValidateTDSClaimedThisYear_TDS3b = True
''    setTblinfo_TDS3b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSClaimedThisYear_TDS3b(end_TDS3b)
''    Set rangecells = Sheet19.Range("TDS3b.TDSClaimedthisyear").Cells
''
''    For i = 1 To end_TDS3b
''        TDSClaimedThisYear_TDS3b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSClaimedThisYear_TDS3b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS2b = MsgTDS3b + "TDS Claimed This Year at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSClaimedThisYear_TDS3b = False
''         Exit Function
''    End If
''
''    If Len(TDSClaimedThisYear_TDS3b(i)) > 14 Then
''        MsgTDS3b = MsgTDS3b + "TDS Claimed This Year at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSClaimedThisYear_TDS3b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateIncomeCarriedFwd_TDS3b() As Boolean
''    ValidateIncomeCarriedFwd_TDS3b = True
''    setTblinfo_TDS3b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim IncomeCarriedFwd_TDS3b(end_TDS3b)
''    Set rangecells = Sheet19.Range("TDS3b.IncBalCF").Cells
''
''    For i = 1 To end_TDS3b
''        IncomeCarriedFwd_TDS3b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(IncomeCarriedFwd_TDS3b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3b = MsgTDS3b + "Income Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateIncomeCarriedFwd_TDS3b = False
''         Exit Function
''    End If
''
''    If Len(IncomeCarriedFwd_TDS3b(i)) > 14 Then
''        MsgTDS3b = MsgTDS3b + "Income Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateIncomeCarriedFwd_TDS3b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateTDSCarriedFwd_TDS3b() As Boolean
''    ValidateTDSCarriedFwd_TDS3b = True
''    setTblinfo_TDS3b
''
''    Dim i As Long
''    Dim rangecells As Range
''
''    ReDim TDSCarriedFwd_TDS3b(end_TDS3b)
''    Set rangecells = Sheet19.Range("TDS3b.TDSBalCF").Cells
''
''    For i = 1 To end_TDS3b
''        TDSCarriedFwd_TDS3b(i) = rangecells.item(i).value
''
''    If Not chkCompulsory(TDSCarriedFwd_TDS3b(i)) And rangecells.item(i).Locked = False Then
''         MsgTDS3b = MsgTDS3b + "TDS Carried Forward at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateTDSCarriedFwd_TDS3b = False
''         Exit Function
''    End If
''
''    If Len(TDSCarriedFwd_TDS3b(i)) > 14 Then
''        MsgTDS3b = MsgTDS3b + "TDS Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits." & Chr(13)
''        ValidateTDSCarriedFwd_TDS3b = False
''    End If
''
''    Next
''End Function
''
''Function ValidateHeadOfIncome_TDS3b() As Boolean
''
''    ValidateHeadOfIncome_TDS3b = True
''    setTblinfo_TDS3b
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3b.HeadOfIncome").Cells
''    ReDim HeadOfIncome_TDS3b(end_TDS3b)
''    For i = 1 To end_TDS3b
''        HeadOfIncome_TDS3b(i) = rangecells.item(i).value
''        If isdropdownblank(HeadOfIncome_TDS3b(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS3b = MsgTDS3b + "Head Of Income at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateHeadOfIncome_TDS3b = False
''              Exit Function
''        End If
''
''        If Mid(HeadOfIncome_TDS3b(i), 1, 1) = "H" Then HeadOfIncome_TDS3b(i) = "HP"
''        If Mid(HeadOfIncome_TDS3b(i), 1, 1) = "O" Then HeadOfIncome_TDS3b(i) = "OS"
''
'''        If Not checkfieldspecialcharacter(HeadOfIncome_TDS3b(i)) Then
'''              MsgTDS3b = MsgTDS3b + "Special characters are not allowed in Head Of Income at Sr. No  " & i & " ." & Chr(13)
'''              ValidateHeadOfIncome_TDS3b = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidateReasonClaimOwnHands_TDS3b() As Boolean
''
''    ValidateReasonClaimOwnHands_TDS3b = True
''    setTblinfo_TDS3b
''    Dim rangecells As Range
''    Set rangecells = Sheet19.Range("TDS3b.Reason").Cells
''    ReDim ReasonClaimOwnHands_TDS3b(end_TDS3b)
''    For i = 1 To end_TDS3b
''        ReasonClaimOwnHands_TDS3b(i) = rangecells.item(i).value
''        If isdropdownblank(ReasonClaimOwnHands_TDS3b(i)) And rangecells.item(i).Locked = False Then
''            MsgTDS3b = MsgTDS3b + "Reason for Claiming in Own Hands at Sr. No  " & i & " is Mandatory." & Chr(13)
''              ValidateReasonClaimOwnHands_TDS3b = False
''              Exit Function
''        End If
''
''        If Mid(ReasonClaimOwnHands_TDS3b(i), 1, 1) = "P" Then ReasonClaimOwnHands_TDS3b(i) = "PCC"
''        If Mid(ReasonClaimOwnHands_TDS3b(i), 1, 1) = "C" Then ReasonClaimOwnHands_TDS3b(i) = "COA"
''        If Mid(ReasonClaimOwnHands_TDS3b(i), 1, 1) = "M" Then ReasonClaimOwnHands_TDS3b(i) = "SIC"
''        If Mid(ReasonClaimOwnHands_TDS3b(i), 1, 1) = "T" Then ReasonClaimOwnHands_TDS3b(i) = "TDP"
''        If Mid(ReasonClaimOwnHands_TDS3b(i), 1, 1) = "D" Then ReasonClaimOwnHands_TDS3b(i) = "DPI"
''
'''        If Not checkfieldspecialcharacter(ReasonClaimOwnHands_TDS3b(i)) Then
'''              MsgTDS3b = MsgTDS3b + "Special characters are not allowed in Reason for Claiming in Own Hands at Sr. No  " & i & " ." & Chr(13)
'''              ValidateReasonClaimOwnHands_TDS3b = False
'''              Exit Function
'''        End If
'' Next
''End Function
''
''Function ValidateTDSDeductedPAN_TDS3b() As Boolean
''ValidateTDSDeductedPAN_TDS3b = True
''setTblinfo_TDS3b
''Dim rangecells As Range
''Set rangecells = Sheet19.Range("TDS3b.PAN7").Cells
''Dim i As Long
''ReDim TDSDeductedPAN_TDS3b(end_TDS3b)
''For i = 1 To end_TDS3b
''TDSDeductedPAN_TDS3b(i) = rangecells.item(i).value
''
''   If Not chkCompulsory(TDSDeductedPAN_TDS3b(i)) And rangecells.item(i).Locked = False Then
''        MsgTDS3b = MsgTDS3b + "PAN at Sr.No " & i & " is Mandatory." & Chr(13)
''         ValidateTDSDeductedPAN_TDS3b = False
''         Exit Function
''  End If
''
''    If Not checkfieldspecialcharacter(TDSDeductedPAN_TDS3b(i)) Then
''         MsgTDS3b = MsgTDS3b + "PAN at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
''         ValidateTDSDeductedPAN_TDS3b = False
''         Exit Function
''    End If
''
''    If Not CheckPAN(UCase(TDSDeductedPAN_TDS3b(i))) Then
''         MsgTDS3b = MsgTDS3b + "PAN at Sr. No  " & i & " in Sheet TDS is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
''         ValidateTDSDeductedPAN_TDS3b = False
''         Exit Function
''    End If
''
''    If TDSDeductedPAN_TDS3b(i) = Sheet1.Range("sheet1.PAN").value Then
''                MsgTDS3b = MsgTDS3b + "PAN (on which TDS deducted) Sr.No " & i & " can not be same as PAN of user  " & Chr(13)
''                ValidateTDSDeductedPAN_TDS3b = False
''                Exit Function
''    End If
''
''Next
''End Function
''
''
'Function ValidateSheetTDS2ii() As Boolean
'On Error Resume Next
'Dim cellcount, intcntr, rangname, name As Variant
'Dim k, j As Long
'Dim SumTds, SumIft, tdssum As Variant
'
'
'ValidateSheetTDS2ii = True
'    MsgTDS2ii = "TDS 3." & Chr(13)
'    subProcCaption = "Validating TDS2ii"
'    setTblinfo_TDS2ii
'    noOfProcessSub = end_TDS2ii
'For k = 1 To end_TDS2ii
'        If Not (Sheet21.Range("TDSoth2.PAN").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet21.Range("TDSoth2.DeductedYear").Cells.item(k)) Or Sheet21.Range("TDSoth2.AmountSubjectDed").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_OwnHands").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells.item(k) <> "") Then
'            MsgTDS2ii = MsgTDS2ii & " At Sr. No " & k & " Please fill all the Mandatory Fields " & Chr(13)
'            ValidateSheetTDS2ii = False
'            Exit Function
'        End If
'    UpdateProgressBar
'Next
'
'If end_TDS2ii > 0 Then GoTo xyz
'
'    cellcount = Sheet21.Range("TDSoth2.PAN").Count
'    If Mid(Range("sheet1.PortugeseCC5A").value, 1, 3) <> "Yes" Then
'        rangname = "TDSoth2.PAN|TDSoth2.EmployerOrDeductorOrCollecterName|TDSoth2.DeductedYear |TDSoth2.AmountSubjectDed|TDSoth2.TD_OwnHands|TDSoth2.TD_SIncome|TDSoth2.TD_TDS|TDSoth2.ClaimOutOfTotTDSOnAmtPaid|TDSoth2.TDS_SIncome|TDSoth2.TDS_TDS"
'    Else:
'        rangname = "TDSoth2.PAN|TDSoth2.EmployerOrDeductorOrCollecterName|TDSoth2.DeductedYear|TDSoth2.AmountSubjectDed|TDSoth2.TD_OwnHands|TDSoth2.TD_SIncome|TDSoth2.TD_TDS|TDSoth2.ClaimOutOfTotTDSOnAmtPaid|TDSoth2.TDS_SIncome|TDSoth2.TDS_TDS|TDSoth2.TD_PAN|TDSoth2.TDS_PAN"
'    End If
'    rangname = Split(rangname, "|")
'
'If Mid(Range("sheet1.PortugeseCC5A").value, 1, 3) <> "Yes" Then
'    For intcntr = 1 To cellcount
'        If (Sheet21.Range("TDSoth2.PAN").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet21.Range("TDSoth2.DeductedYear").Cells.item(k)) Or Sheet21.Range("TDSoth2.AmountSubjectDed").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_OwnHands").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_SIncome").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_TDS").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TDS_SIncome").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TDS_TDS").Cells.item(k) <> "") Then
'            For Each name In rangname
'                If Sheet21.Range(name).Cells.item(intcntr) = "" Then ValidateSheetTDS2ii = False
'            Next
'            If ValidateSheetTDS2ii = False Then
'                MsgTDS2ii = MsgTDS2ii & "At Sr. No " & intcntr & " Please fill all the Mandatory Fields " & Chr(13)
'                GoTo xyz
'            End If
'
'        End If
'    Next
'
'Else:
'    For intcntr = 1 To cellcount
'        If (Sheet21.Range("TDSoth2.PAN").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Not isdropdownblank(Sheet21.Range("TDSoth2.DeductedYear").Cells.item(k)) Or Sheet21.Range("TDSoth2.AmountSubjectDed").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_OwnHands").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_SIncome").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_TDS").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TDS_SIncome").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TDS_TDS").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TD_PAN").Cells.item(k) <> "" Or Sheet21.Range("TDSoth2.TDS_PAN").Cells.item(k) <> "") Then
'            For Each name In rangname
'                If Sheet21.Range(name).Cells.item(intcntr) = "" Then ValidateSheetTDS2ii = False
'            Next
'            If ValidateSheetTDS2 = False Then
'                MsgTDS2ii = MsgTDS2ii & "At Sr. No " & intcntr & " Please fill all the Mandatory Fields " & Chr(13)
'                GoTo xyz
'            End If
'        End If
'    Next
'End If
'
'xyz:
'     If Not ValidatePAN_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidateEmployerOrDeductorOrCollecterName_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidateUTN_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidateDeductedYear_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii() Then ValidateSheetTDS2ii = False
'    If Not ValidatePAN7_TDSii() Then ValidateSheetTDS2ii = False
'
'
'
'        setTableinfo_TDS_PAN_TaxTDS2ii
'        setTableinfo_TDS_TDS_TDS2ii
'        setTableinfo_TDS_SIncome_TDS2ii
'        setTableinfo_TD_PAN_TDS2ii
'        setTableinfo_TD_TDS_TDS2ii
'        setTableinfo_TDsinc_TaxTDS2ii
'        setTableinfo_Amt_TaxTDS2ii
'        setTableinfo_TD_OwnHand2ii
'        setTableinfo_AmtPaidTDS2ii
'        setTableinfo_Year_TDS2ii
'        setTableinfo_Name_TDS2ii
'
''        If (Sheet1.Range("sheet1.PortugeseCC5A").value = "Yes") Then
''         setTableinfo_TD2_PAN_TDS2ii
''          setTableinfo_TDS2_PAN_TaxTDS2ii
''    End If
'
'      If ((end_TDS2ii <> end_Name_TDS2ii) Or (end_TDS2ii <> end_Year_TDS2ii) Or (end_TDS2ii <> end_AmtPaid_TDS2ii) Or (end_TDS2ii <> end_TDi_TDS2ii) Or (end_TDS2ii <> end_Amt_TDS2ii)) Then
'                MsgTDS2ii = MsgTDS2ii + "Please fill all the Mandatory Fields " & Chr(13)
'                ValidateSheetTDS2ii = False
'            End If
'
'            If ((end_TDii_TDS2ii <> end_TDiii_TDS2ii) Or (end_TDii_TDS2ii <> end_TD_PAN_TDS2ii)) Or ((end_TDS_SIncome_TDS2ii <> end_TDS_TDS_TDS2ii) Or (end_TDS_SIncome_TDS2ii <> end_TDS_PAN_TDS2ii)) Then
'             MsgTDS2ii = MsgTDS2ii + "Please fill all the Mandatory Fields " & Chr(13)
'                ValidateSheetTDS2ii = False
'            End If
'
'
'     For k = 1 To end_TDS2ii
'     tdssum = 0
'     For j = 1 To end_TDS2iiinner
'     SumTds = Mid(Range("PAN_2").Cells.item(j).value, 3, 10) '('Application.WorksheetFunction.SumIf((Range("TDSoth.TDS_TDS")), (Mid(Range("TAN_1").Cells.item(j).value, 3, 10) = Range("TDSoth.TAN").Cells.item(k).value))
'     'SumIft = (Sheet21.Range("TAN_1").Cells.item(j).value)
'     'SumIft = Application.WorksheetFunction.Mid((Sheet21.Range("TAN_1").Cells.item(j).value), 1, 2)
'     SumIft = Range("TDSoth2.PAN").Cells.item(k).value
'     If (SumTds = SumIft) Then
'     tdssum = WorksheetFunction.Sum(Range("TDSoth2.TDS_TDS").Cells.item(j).value, tdssum)
'     End If
'     Next
'     If (Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").item(k).value + tdssum > (Range("TDSoth2.TD_TDS").item(k).value + Range("TDSoth2.TD_OwnHands").item(k).value)) Then
'            MsgTDS2ii = MsgTDS2ii + "For PAN At Sr. No " & k & Chr(13) & "Amount in field (Col (8)+ TDS at Col.(9)) cannot be more than sum of field (Col (6) and TDS at col.(7))." & Chr(13)
'            ValidateSheetTDS2ii = False
'            Exit Function
'            End If
'
'     Next
'
'
''     If Range("TDSoth2.Sum").value > 0 Then
''             MsgBox ("If TDS3(26QC) is filled ,then please ensure the ""Type of house property"" is Let out & ""Gross rent received /receivable /letable value"" is greater then zero"), vbExclamation, "Warning"
''            End If
''
'
'
'If ValidateSheetTDS2ii = False Then GoTo Next1
'If Not checkforNonMandatoryCol_TDS2ii Then ValidateSheetTDS2ii = False
'
'Next1:
'End Function
'
'
'Function checkforNonMandatoryCol_TDS2ii() As Boolean
'checkforNonMandatoryCol_TDS2ii = True
'    Dim temp As Double
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth2.UTN").Count
'    Set rangecells = Range("TDSoth2.UTN").Cells
'
'    temp = end_TDS2ii
'
'    For mIntCtr = (temp + 1) To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                MsgTDS2ii = MsgTDS2ii & "At Sr. No " & mIntCtr & " Please fill all the Mandatory Fields " & Chr(13)
'                checkforNonMandatoryCol_TDS2ii = False
'                Exit Function
'           End If
'    Next
'End Function
'
'Sub setTblinfo_TDS2ii()
' Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.PAN").Count
' Set rangecells = Sheet21.Range("TDSoth2.PAN").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2ii = ccount
' rngname_TDS2ii = "TDSoth2.PAN||TDSoth2.EmployerOrDeductorOrCollecterName||TDSoth2.UTN||TDSoth2.DeductedYear||TDSoth2.AmountSubjectDed ||TDSoth2.TD_OwnHands ||TDSoth2.TD_SIncome||TDSoth2.TD_TDS||TDSoth2.TD_PAN||TDSoth2.ClaimOutOfTotTDSOnAmtPaid||TDSoth2.TDS_SIncome||TDSoth2.TDS_TDS||TDSoth2.TDS_PAN"
' End Sub
' Sub setTblinfo_TDS2iiinner()
' Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("PAN_2").Count
' Set rangecells = Sheet21.Range("PAN_2").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS2iiinner = ccount
' 'rngname_TDS2ii = "TDSoth2.PAN||TDSoth2.EmployerOrDeductorOrCollecterName||TDSoth2.UTN||TDSoth2.DeductedYear||TDSoth2.AmountSubjectDed ||TDSoth2.TD_OwnHands ||TDSoth2.TD_SIncome||TDSoth2.TD_TDS||TDSoth2.TD_PAN||TDSoth2.ClaimOutOfTotTDSOnAmtPaid||TDSoth2.TDS_SIncome||TDSoth2.TDS_TDS||TDSoth2.TDS_PAN"
' End Sub
'
' Function ValidateUTN_TDS2ii() As Boolean
'
'    ValidateUTN_TDS2ii = True
'    setTblinfo_TDS2ii
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet21.Range("TDSoth2.UTN").Cells
'    ReDim UTN_TDS2ii(end_TDS2ii)
'    For i = 1 To end_TDS2ii
'        UTN_TDS2ii(i) = rangecells.item(i).value
'    If Not checkfieldspecialcharacter(UTN_TDS2ii(i)) Then
'         MsgTDS2ii = MsgTDS2ii + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid"
'          ValidateUTN_TDS2ii = False
'          Exit Function
'     End If
'
'If InStr(UTN_TDS2ii(i), ".") > 0 Then
'         MsgTDS2ii = MsgTDS2ii + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
'          ValidateUTN_TDS2ii = False
'          Exit Function
'    End If
'
'    If InStr(UTN_TDS2ii(i), "-") > 0 Then
'         MsgTDS2ii = MsgTDS2ii + "Unique TDS Certificate Number at Sr. No  " & i & " in Sheet TDS is invalid" & Chr(13)
'          ValidateUTN_TDS2ii = False
'          Exit Function
'    End If
'
'    If UTN_TDS2ii(i) <> "" Then
'    If Len(UTN_TDS2ii(i)) < 6 Or Len(UTN_TDS2ii(i)) > 8 Then
'          MsgTDS2ii = MsgTDS2ii + "Unique TDS Certificate Number at Sr. No  " & i & " is invalid. In case the Unique TDS Certificate Number is less than 8 digits, please prefix with zero's." & Chr(13)
'          ValidateUTN_TDS2ii = False
'          Exit Function
'    End If
'    End If
'
'
' Next
'End Function
'
'Function ValidateEmployerOrDeductorOrCollecterName_TDS2ii() As Boolean
'
'    ValidateEmployerOrDeductorOrCollecterName_TDS2ii = True
'    setTblinfo_TDS2ii
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet21.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells
'    ReDim EmployerOrDeductorOrCollecterName_TDS2ii(end_TDS2ii)
'    For i = 1 To end_TDS2ii
'        EmployerOrDeductorOrCollecterName_TDS2ii(i) = rangecells.item(i).value
'
'        If EmployerOrDeductorOrCollecterName_TDS2ii(i) = "" Then
'            MsgTDS2ii = MsgTDS2ii + "Name of the tenant at Sr. No  " & i & "  is Mandatory" & Chr(13)
'            ValidateEmployerOrDeductorOrCollecterName_TDS2ii = False
'            Exit Function
'        End If
'
'
'            If Len(EmployerOrDeductorOrCollecterName_TDS2ii(i)) > 125 Then
'            MsgTDS2ii = MsgTDS2ii + "Name of the tenant at Sr. No  " & i & "  in Sheet TDS  Cannot be more than 125 Charachters" & Chr(13)
'            ValidateEmployerOrDeductorOrCollecterName_TDS2ii = False
'            Exit Function
'        End If
'
'    Next
'End Function
'
'Function ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii() As Boolean
'    ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = True
'    setTblinfo_TDS2ii
'    setTblinfo_TDS2iiinner
'    setTableinfo_TDsinc_TaxTDS2ii
'    setTableinfo_TDS_SIncome_TDS2ii
'
'    Dim rangecells1 As Range
'    Dim rangecells2 As Range
'    Dim rangecells5 As Range
'    Dim rangecells6 As Range
'    Dim rangecells7 As Range
'    Dim rangecells8 As Range
'    Dim rangecells9 As Range
'
'    Dim i As Long
'    Dim portugesecode As Variant
'
'    Set rangecells6 = Sheet21.Range("TDSoth2.AmountSubjectDed").Cells
'    Set rangecells1 = Sheet21.Range("TDSoth2.TD_OwnHands").Cells
'    Set rangecells5 = Sheet21.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells
'
'    Set rangecells2 = Sheet21.Range("TDSoth2.TD_SIncome").Cells
'    Set rangecells7 = Sheet21.Range("TDSoth2.TD_TDS").Cells
'
'    Set rangecells8 = Sheet21.Range("TDSoth2.TDS_SIncome").Cells
'    Set rangecells9 = Sheet21.Range("TDSoth2.TDS_TDS").Cells
'
'    portugesecode = Trim(UCase(Sheet1.Range("sheet1.PortugeseCC5A").value))
'
'    ReDim AmtPaid_TDS2ii(end_TDS2ii)
'    ReDim TD_OwnHands_TDS2ii(end_TDS2ii)
'    ReDim AmtClaimedOnOwnHands_TDS2ii(end_TDS2ii)
'
'    ReDim TD_SIncome_TDS2ii(end_TDS2ii)
'    ReDim TD_TDS_TDS2ii(end_TDS2ii)
'
'    ReDim TDS_SIncome_TDS2ii(end_TDS2iiinner)
'    ReDim TDS_TDS_TDS2ii(end_TDS2iiinner)
'
'    ReDim TotalSum2ii(end_TDS2ii)
'    ReDim TotalSum2iii(end_TDS2ii)
'
'For i = 1 To end_TDS2ii
'
'    AmtPaid_TDS2ii(i) = rangecells6.item(i).value
'    AmtClaimedOnOwnHands_TDS2ii(i) = rangecells5.item(i).value
'    TD_OwnHands_TDS2ii(i) = rangecells1.item(i).value
'
'
'    If AmtClaimedOnOwnHands_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Amount claimed this year at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'    If AmtPaid_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Please enter the details of Receipt as mentioned in Form 26AS at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'     If TD_OwnHands_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Please enter Tax deducted in own hands at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'
'
'    If IsNumeric(AmtPaid_TDS2ii(i)) Then
'        If AmtPaid_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 5 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & AmtPaid_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 5 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 5 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'
'    If IsNumeric(AmtClaimedOnOwnHands_TDS2ii(i)) Then
'        If AmtClaimedOnOwnHands_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & AmtClaimedOnOwnHands_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'
'
'    If IsNumeric(TD_OwnHands_TDS2ii(i)) Then
'        If TD_OwnHands_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 6 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & TD_OwnHands_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 6 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 6 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'
'   If (AmtClaimedOnOwnHands_TDS2ii(i) = "") Then
'        AmtClaimedOnOwnHands_TDS2ii(i) = 0
'    End If
'
'Next
'
'
'
'For i = 1 To end_TDii_TDS2ii
'    TD_SIncome_TDS2ii(i) = rangecells2.item(i).value
'    TD_TDS_TDS2ii(i) = rangecells7.item(i).value
'
'     If (TD_SIncome_TDS2ii(i) <> "" Or TD_TDS_TDS2ii(i) <> "") Then
'
'    If TD_SIncome_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Please enter Income which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No" & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'
'    If TD_TDS_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Please enter Tax deducted in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'
'    If IsNumeric(TD_SIncome_TDS2ii(i)) Then
'        If TD_SIncome_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & TD_SIncome_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 7 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'
'    If IsNumeric(TD_TDS_TDS2ii(i)) Then
'        If TD_TDS_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & TD_TDS_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 7 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'    End If
'Next
'
'
'
'For i = 1 To end_TDS2iiinner
'
'    TDS_SIncome_TDS2ii(i) = rangecells8.item(i).value
'    TDS_TDS_TDS2ii(i) = rangecells9.item(i).value
'
' If TDS_SIncome_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Please enter Income which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'    If TDS_TDS_TDS2ii(i) = "" Then
'        MsgTDS2ii = MsgTDS2ii + "Please enter Tax deducted in the hands of spouse as per section 5A or any other person as per rule 37BA(2) at Sr.No " & i & " is Mandatory" & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'    End If
'
'
'
'If IsNumeric(TDS_SIncome_TDS2ii(i)) Then
'        If TDS_SIncome_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & TDS_SIncome_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'
'If IsNumeric(TDS_TDS_TDS2ii(i)) Then
'        If TDS_TDS_TDS2ii(i) < 0 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        End If
'        If Len("" & TDS_TDS_TDS2ii(i)) > 14 Then
'            MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'            Exit Function
'        End If
'    Else
'        MsgTDS2ii = MsgTDS2ii + "Amount in col.no 8 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii = False
'        Exit Function
'    End If
'
'Next
'
'
'End Function
'
'Function ValidatePAN7_TDSii() As Boolean
'ValidatePAN7_TDSii = True
'
'setTableinfo_TDsinc_TaxTDS2ii
'setTableinfo_TDS_SIncome_TDS2ii
'setTblinfo_TDS2iiinner
'setTblinfo_TDS2ii
'
'
'ValidateClaimOutOfTotTDSOnAmtPaid_TDS2ii
'
'noOfProcessSub = end_TDS2ii
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'Set rangecells = Range("TDSoth2.TD_PAN").Cells
'Set rangecells1 = Range("TDSoth2.TDS_PAN").Cells
'
'
'Dim i As Long
'ReDim PAN7_TDSii(end_TDS2ii)
'ReDim PAN8_TDSii(end_TDS2iiinner)
'
'For i = 1 To end_TDS2ii
'    PAN7_TDSii(i) = rangecells.item(i).value
'
'    'If (TD_SIncome_TDS2ii(i) <> "" Or TD_TDS_TDS2ii(i) <> "") Then
'
'    If Not Len(PAN7_TDSii(i)) = 0 Then
'        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN7_TDSii(i), 1, 10))) Then
'        MsgTDS2ii = MsgTDS2ii + "PAN of spouse / other person at Sr. No  " & i & "  in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet" & Chr(13)
'        ValidatePAN7_TDSii = False
'        End If
'
'
'    ElseIf Not chkCompulsory(PAN7_TDSii(i)) Then
'        MsgTDS2ii = MsgTDS2ii + " PAN of spouse / other person at Sr.No " & i & " in Sheet TDS, (20 TDS2(ii)) is mandatory" & Chr(13)
'        ValidatePAN7_TDSii = False
'    End If
'    'End If
'Next
'
'For i = 1 To end_TDS2iiinner
'
'    PAN8_TDSii(i) = rangecells1.item(i).value
'
'    If Not Len(PAN8_TDSii(i)) = 0 Then
'        If Not mdIncomeDetails.CheckPAN(UCase(Mid(PAN8_TDSii(i), 1, 10))) Then
'        MsgTDS2ii = MsgTDS2ii + "PAN of spouse / other person at Sr. No  " & i & "  in Sheet TDS is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
'        ValidatePAN7_TDSii = False
'        End If
'
'
'    ElseIf Not chkCompulsory(PAN8_TDSii(i)) Then
'        MsgTDS2ii = MsgTDS2ii + " PAN of spouse / other person at Sr.No " & i & " in Sheet TDS, (20 TDS2(ii)) is mandatory" & Chr(13)
'        ValidatePAN7_TDSii = False
'    End If
'
'
'    UpdateProgressBar
'Next
'End Function
'
'Function ValidatePAN_TDS2ii() As Boolean
'ValidatePAN_TDS2ii = True
'setTblinfo_TDS2ii
'setTblinfo_TDS2iiinner
'noOfProcessSub = end_TDS2ii
'
'Dim rangecells As Range
'Dim rangecells1 As Range
'Set rangecells = Range("TDSoth2.PAN").Cells
'Set rangecells1 = Range("PAN_2").Cells
'Dim i As Long
' ReDim PAN_TDS2ii(end_TDS2ii)
' ReDim PAN2_TDS2ii(end_TDS2iiinner)
'
'For i = 1 To end_TDS2ii
'    PAN_TDS2ii(i) = rangecells.item(i).value
'    If Not Len(PAN_TDS2ii(i)) = 0 Then
'        If Not mdIncomeDetails.CheckPAN_Tenant(UCase(Mid(PAN_TDS2ii(i), 1, 10))) Then
'        MsgTDS2ii = MsgTDS2ii + "PAN of the Tenant at Sr. No  " & i & "  in Sheet TDS3  is invalid. First 5 alphabets, next 4 digits, then alphabet (4th alphabet must be ""P"" or ""H"") " & Chr(13)
'        ValidatePAN_TDS2ii = False
'        End If
'    End If
'
'    UpdateProgressBar
'Next
'
'For i = 1 To end_TDS2iiinner
'    PAN2_TDS2ii(i) = rangecells1.item(i).value
'Next
'
'End Function
'
'Function ValidateDeductedYear_TDS2ii() As Boolean
'
'    ValidateDeductedYear_TDS2ii = True
'    setTblinfo_TDS2ii
'    Dim i As Long
'    Dim rangecells As Range
'    Set rangecells = Sheet21.Range("TDSoth2.DeductedYear").Cells
'    ReDim DeductedYear_TDS2ii(end_TDS2ii)
'For i = 1 To end_TDS2ii
'        DeductedYear_TDS2ii(i) = rangecells.item(i).value
'
'        If isdropdownblank(DeductedYear_TDS2ii(i)) Then
'            MsgTDS2ii = MsgTDS2ii + "Deducted Year at Sr. No  " & i & " is Mandatory" & Chr(13)
'            ValidateDeductedYear_TDS2ii = False
'            Exit Function
'        End If
'
'
'     If Not checkfieldspecialcharacter(DeductedYear_TDS2ii(i)) Then
'         MsgTDS2ii = MsgTDS2ii + "Deducted Year at Sr. No  " & i & " in Sheet TDS schedule TDS2 is invalid" & Chr(13)
'          ValidateDeductedYear_TDS2ii = False
'          Exit Function
'     End If
'     If DeductedYear_TDS2ii(i) > 2017 Then
'          MsgTDS2ii = MsgTDS2ii + "Deducted Year at Sr. No  " & i & " in Sheet TDS cannot be greater than 2017 " & Chr(13)
'          ValidateDeductedYear_TDS2ii = False
'          Exit Function
'    End If
'Next
'End Function
'
'Sub setTableinfo_Name_TDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Count
' Set rangecells = Sheet21.Range("TDSoth2.EmployerOrDeductorOrCollecterName").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_Name_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_Year_TDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.DeductedYear").Count
' Set rangecells = Sheet21.Range("TDSoth2.DeductedYear").Cells
' For mIntCtr = 1 To mIntCells
'     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
'         ccount = ccount + 1
'     End If
' Next
' end_Year_TDS2ii = ccount
'End Sub
'
'Sub setTableinfo_AmtPaidTDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.AmountSubjectDed").Count
' Set rangecells = Sheet21.Range("TDSoth2.AmountSubjectDed").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_AmtPaid_TDS2ii = ccount
'End Sub
''
'Sub setTableinfo_TD_OwnHand2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TD_OwnHands").Count
' Set rangecells = Sheet21.Range("TDSoth2.TD_OwnHands").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDi_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_Amt_TaxTDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Count
' Set rangecells = Sheet21.Range("TDSoth2.ClaimOutOfTotTDSOnAmtPaid").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_Amt_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_TDsinc_TaxTDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TD_SIncome").Count
' Set rangecells = Sheet21.Range("TDSoth2.TD_SIncome").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDii_TDS2ii = ccount
'End Sub
'
'Sub setTableinfo_TD_TDS_TDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TD_TDS").Count
' Set rangecells = Sheet21.Range("TDSoth2.TD_TDS").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDiii_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_TD_PAN_TDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TD_PAN").Count
' Set rangecells = Sheet21.Range("TDSoth2.TD_PAN").Cells
' For mIntCtr = 1 To mIntCells
'     If Not (rangecells.item(mIntCtr).value = "" Or UCase(rangecells.item(mIntCtr).value) = "(SELECT)") Then
'         ccount = ccount + 1
'     End If
' Next
' end_TD_PAN_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_TDS_SIncome_TDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TDS_SIncome").Count
' Set rangecells = Sheet21.Range("TDSoth2.TDS_SIncome").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS_SIncome_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_TDS_TDS_TDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TDS_TDS").Count
' Set rangecells = Sheet21.Range("TDSoth2.TDS_TDS").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS_TDS_TDS2ii = ccount
'End Sub
'
'
'Sub setTableinfo_TDS_PAN_TaxTDS2ii()
'Dim rangecells As Range
' Dim mIntCells As Long
' Dim mIntCtr As Long
' Dim ccount As Long
' ccount = 0
' mIntCells = Sheet21.Range("TDSoth2.TDS_PAN").Count
' Set rangecells = Sheet21.Range("TDSoth2.TDS_PAN").Cells
' For mIntCtr = 1 To mIntCells
'     If Not rangecells.item(mIntCtr).value = "" Then
'         ccount = ccount + 1
'     End If
' Next
' end_TDS_PAN_TDS2ii = ccount
'End Sub
'
'Sub ValidateTDS_TDS3()
'On Error Resume Next
'
'    MsgTDS2ii = ""
'    If Not ValidateSheetTDS2ii Then
'        Sheet21.Activate
'        MsgBox (MsgTDS2ii)
'        CloseMsg
'    End If
'MsgBox "Sheet is Ok", vbOKOnly, vbMessageCaption
'
'End Sub
'
'Sub Cmd_Prev_Click_TDS3()
'Sheet22.Activate
'End Sub
'
'Sub Cmd_Next_Click_TDS3()
'Sheet18.Activate
'End Sub
