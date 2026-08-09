Attribute VB_Name = "SchFSI"
Option Explicit

Public CountryCode_FSI As Variant
Public TaxIdentificationNo_FSI As Variant

Public TXNOSAL_FSI As Variant
Public TXNISAL_FSI As Variant
Public TXRSAL_FSI As Variant
Public DTAASAL_FSI As Variant

Public TXNOHP_FSI As Variant
Public TXNIHP_FSI As Variant
Public TXRHP_FSI As Variant
Public DTAAHP_FSI As Variant

Public TXNOBP_FSI As Variant
Public TXNIBP_FSI As Variant
Public TXRBP_FSI As Variant
Public DTAABP_FSI As Variant

Public TXNOCapGain_FSI As Variant
Public TXNICapGain_FSI As Variant
Public TXRCapGain_FSI As Variant
Public DTAACapGain_FSI As Variant

Public TXNOOthSrc_FSI As Variant
Public TXNIOthSrc_FSI As Variant
Public TXROthSrc_FSI As Variant
Public DTAAOthSrc_FSI As Variant


Public IncT_FSI As Variant
Public TXNOT_FSI As Variant
Public TXNIT_FSI As Variant
Public TXRT_FSI As Variant
    

Public IncFromSal_FSI As Variant
Public IncFromHP_FSI As Variant
Public IncFromBP_FSI As Variant
Public IncCapGain_FSI As Variant
Public IncOthSrc_FSI As Variant
Public TotalCountryWise_FSI As Variant
Public TotalIncomeOutIndia_FSI As Variant
Public TotalIncomeOutIndiaDTAA_FSI As Variant
Public TotalIncomeOutIndiaWoDTAA_FSI As Variant

'FSI
Public rngname_FSI As Variant
Public end_FSI, end_FSI11 As Variant

Dim rngname_FSI2 As Variant
Dim end_FSI2 As Variant

Dim incBy_FSI As Variant
Dim msgValidateSheetFSI As String
Dim Msgbox_FSI As String

Public Country_FSI As Variant
Public CountryName_FSI As Variant

Public end_FSI3, end_FSI4, end_FSI5, end_FSI6, end_FSI7, end_FSI8, end_FSI9 As Long
Public end_FSI10, end_FSI111, end_FSI12, end_FSI13, end_FSI14, end_FSI15, end_FSI16 As Long
Public end_FSI17, end_FSI18, end_FSI19, end_FSI20, end_FSI21, end_FSI22 As Long


Public frmsize_FSI, end_FSIfrm, cntrRng_FSIfrm, frmRngname_FSIfrm, rngname_FSIfrm As Variant

Public aa As Variant

Sub AddBlockCall_FSIfrm()
'Change-3, ANK, 01.06.2022
    Sheet40.Activate
'---
    settbl_FSI
    Call addblock_FSI(rngname_FSIfrm, frmRngname_FSIfrm, cntrRng_FSIfrm, frmsize_FSI)
    
    'commented by Chetan C M on 07/10/2025 (SIT-101803)
    
    'SIT-104166 uncommented by sai on 01/12/2025
    '***************************************************
    AddRows_TR_FSI
    Sheet40.Activate
    '***************************************************
    
End Sub

Sub settbl_FSI()

    Dim te As Long
    Dim ccount As Long
    
    frmsize_FSI = 6
    For te = 1 To Sheet40.Range("FSI_Count").value
        If Not Sheet40.Range("FSI_CountryCode" & te).value = "" Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    end_FSIfrm = ccount
    cntrRng_FSIfrm = "FSI_Count"
    frmRngname_FSIfrm = "FSIblock"
    rngname_FSIfrm = "FSI_CountryCode1;FSI_TaxIdentificationNo1;FSI_IncFromSal1;FSI_TXNOSAL1;FSI_TXNISAL1;FSI_TXRSAL1;FSI_DTAASAL1;FSI_IncFromHP1;FSI_TXNOHP1;FSI_TXNIHP1;FSI_TXRHP1;FSI_DTAAHP1;FSI_IncFromBP1;FSI_TXNOBP1;FSI_TXNIBP1;FSI_TXRBP1;FSI_DTAABP1;FSI_IncCapGain1;FSI_TXNOCapGain1;FSI_TXnICapGain1;FSI_TXRCapGain1;FSI_DTAACapGain1;FSI_IncOthSrc1;FSI_TXNOOthSrc1;FSI_TXNIOthSrc1;FSI_TXROthSrc1;FSI_DTAAOthSrc1;FSI_IncT1;FSI_TXNOT1;FSI_TXNIT1;FSI_TXRT1;"

End Sub

Function addblock_FSI(prevrangename As Variant, blockname As Variant, frmcounter As Variant, totalblocksize As Variant)
On Error GoTo endline
    Dim i, commindex, j, k, b, rangeCount, offset1, offset2 As Long
    Dim newnamerefersto As Variant
    Dim gapbtn, counter, dcounter, lenn, startrow, endrow, endrow1, endrow2 As Long
    Dim destinationrowindex, destinationcolumnindex, lastdestinationrow As String
    Dim rangearr, newranges, precrange, oldrangeaddress, newrangeaddress, newrangeaddress1, newrangeaddress2 As Variant
    Dim ndestinationrowindex, ndestinationcolumnindex As Long
    Dim noofrowsadded As Long
    
    Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    gapbtn = 0
    dcounter = 0
    counter = Range(frmcounter).value
    If counter > 0 Then
    
        noofrowsadded = (totalblocksize - totalblocksize) / counter
        offset1 = totalblocksize + (1) * noofrowsadded
        offset2 = totalblocksize + (counter) * noofrowsadded
        ndestinationrowindex = Range(blockname & counter).row
        ndestinationrowindex = ndestinationrowindex + gapbtn + offset1
        ndestinationcolumnindex = Range(blockname & counter).Column
        newnamerefersto = Range(blockname & counter).Address
        lastdestinationrow = ndestinationrowindex + offset1 - 1
        
        Cells(CLng(ndestinationrowindex - 1), CLng(ndestinationcolumnindex)).Select
        
        Call InsertBlockFSI(offset1)
        
        For i = 1 To Len(newnamerefersto)
            If (Mid(newnamerefersto, i, 1) = "$") Then
                 dcounter = dcounter + 1
    
                 If dcounter = 2 Then
                     lenn = (InStr(1, newnamerefersto, ":") - 1) - i
                     startrow = Mid(newnamerefersto, i + 1, lenn)
                     newnamerefersto = Replace(newnamerefersto, startrow, ndestinationrowindex)
                 End If
    
                 If dcounter = 4 Then
                     endrow = Mid(newnamerefersto, i + 1, Len(newnamerefersto) - i)
                     newnamerefersto = Replace(newnamerefersto, endrow, lastdestinationrow)
                 End If
    
            End If
        Next
            
        Application.EnableEvents = False
        
        counter = counter + 1
        ThisWorkbook.Names.add name:=blockname & counter, _
                RefersTo:="=" & newnamerefersto, Visible:=True
    
    
    
       
'---------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
'        Range(frmcounter).Locked = False
        
        Range(frmcounter).value = counter
    
 '       Range(frmcounter).Locked = True
        
    Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   'sPassword = EfilingCommon.getmsgstate
   'ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
        Range(blockname & counter - 1).Copy Destination:=Cells(ndestinationrowindex, ndestinationcolumnindex)
    End If


    commindex = 2
    j = 0
    
    prevrangename = Split(prevrangename, ";")
    rangeCount = UBound(prevrangename)
    
    ReDim rangearr(rangeCount)
    commindex = 0
    j = 1
    
    For i = 0 To UBound(prevrangename)
        rangearr(i) = prevrangename(i)
    Next
    
    ReDim newranges(rangeCount)
    For i = 0 To UBound(rangearr)
       newranges(i) = Replace(rangearr(i), CStr(1), CStr(counter))
    Next
    
    If counter > 2 Then
        ReDim precrange(rangeCount)
        For i = 0 To UBound(rangearr)
             precrange(i) = Replace(rangearr(i), CStr(1), CStr(counter - 1))
        Next
    
        ReDim oldrangeaddress(rangeCount)
        For i = 0 To UBound(precrange)
            If Not (precrange(i) = "") Then
               oldrangeaddress(i) = Range(precrange(i)).Address
            End If
        Next
    Else
        ReDim oldrangeaddress(rangeCount)
        For i = 0 To UBound(rangearr)
            If Not (rangearr(i) = "") Then
                oldrangeaddress(i) = Range(rangearr(i)).Address
            End If
        Next
    End If
    
    ReDim newrangeaddress(rangeCount)
    
    For i = 0 To UBound(oldrangeaddress)
        dcounter = 0
        Dim myrangearr() As String
        myrangearr = Split(oldrangeaddress(i), ":")
        If UBound(myrangearr) <= 0 Then
        For k = 1 To Len(oldrangeaddress(i))
            If (Mid(oldrangeaddress(i), k, 1) = "$") Then
                dcounter = dcounter + 1
                If (dcounter = 2) Then
                endrow = Mid(oldrangeaddress(i), k + 1, Len(oldrangeaddress(i)) - k)
                newrangeaddress(i) = Replace(oldrangeaddress(i), endrow, (endrow + gapbtn + offset1))
                End If
            End If
        Next
        
        Else
        
        Dim myoldrangeaddress() As String
        myoldrangeaddress = Split(oldrangeaddress(i), ":")
        
        For k = 1 To Len(myrangearr(0))
            If (Mid(myoldrangeaddress(0), k, 1) = "$") Then
                dcounter = dcounter + 1
                If (dcounter = 2) Then
                endrow1 = Mid(myoldrangeaddress(0), k + 1, Len(myoldrangeaddress(0)) - k)
                newrangeaddress1 = Replace(myoldrangeaddress(0), endrow1, (endrow1 + gapbtn + offset1))
                End If
            End If
        Next
        dcounter = 0
        For k = 1 To Len(myrangearr(1))
            If (Mid(myoldrangeaddress(1), k, 1) = "$") Then
                dcounter = dcounter + 1
                If (dcounter = 2) Then
                endrow2 = Mid(myoldrangeaddress(1), k + 1, Len(myoldrangeaddress(1)) - k)
                newrangeaddress2 = Replace(myoldrangeaddress(1), endrow2, (endrow2 + gapbtn + offset1))
                End If
            End If
    
        Next
            newrangeaddress(i) = newrangeaddress1 + ":" + newrangeaddress2
        End If
        
    
      Next
    
    
    'On Error Resume Next
    For b = 0 To UBound(newrangeaddress)
    
        If Not newranges(b) = "" Then
            ThisWorkbook.Names.add name:=newranges(b), _
                        RefersTo:="=" & newrangeaddress(b), Visible:=True
            If Range(newranges(b)).Interior.ColorIndex = 35 Then
                Range(newranges(b)).ClearContents
            End If
        End If
    Next
    
endline:

'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Function

Function InsertBlockFSI(vRows1 As Variant)
    Dim x As Long
    Dim vRows As Long
    Dim sht As Worksheet, shts() As String, i As Long
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    Application.EnableEvents = False
    ActiveCell.EntireRow.Select
    
    
    vRows = CLng(vRows1)
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
    Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        
        i = i + 1
        shts(i) = sht.name
        x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
        
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault
        
        On Error Resume Next
        
        Selection.Offset(1).Resize(vRows).EntireRow. _
        SpecialCells(xlConstants).ClearContents
    Next sht
    
    Worksheets(shts).Select
    ActiveSheet.Protect Password:=sPassword
    Application.EnableEvents = True
End Function

Sub Cmd_Validate_FSI_Click()
ValidateScheduleFSI
fmsgboxStatus "Sheet FSI  is ok" ', vbOKOnly, "ITR 3"
End Sub

Sub ValidateScheduleFSI()
    subProcCaption = "Validating FSI"
    If Not ValidatesheetFSI Then
        Sheet40.Activate
        fmsgboxsmall Msgbox_FSI ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub

Function ValidatesheetFSI() As Boolean
    Dim i As Variant
    ValidatesheetFSI = True
    
    If Not ValidateCountry_FSI() Then ValidatesheetFSI = False
        'If Len(Sheet40.Range("FSI_CountryCode" & i).value) > 0 Then
            If Not ValidateTaxIdentificationNo_FSI() Then ValidatesheetFSI = False
            If Not ValidateIncFromSal_FSI() Then ValidatesheetFSI = False
            If Not ValidateIncFromHP_FSI() Then ValidatesheetFSI = False
            If Not ValidateIncFromBusiness_FSI() Then ValidatesheetFSI = False
            If Not ValidateIncCapGain_FSI() Then ValidatesheetFSI = False
            If Not ValidateIncOthSrc_FSI() Then ValidatesheetFSI = False
            If Not validateTotalCountrywise() Then ValidatesheetFSI = False
        'End If
    
End Function


Function ValidateCountry_FSI() As Boolean
 Dim CountryName As String
    ValidateCountry_FSI = True
    
    setTblinfo_FSI
    
    Dim rangecells As Variant
    Dim i As Long
    ReDim Country_FSI(end_FSI)
    ReDim CountryName_FSI(end_FSI)
    noOfProcessSub = end_FSI
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_CountryCode" & i).value
        CountryName = rangecells
        If CountryName <> "(Select)" Then
            Country_FSI(i) = Mid(CountryName, WorksheetFunction.Search(":", CountryName) + 1, Len(CountryName))
            CountryName_FSI(i) = Mid(CountryName, 1, WorksheetFunction.Search(":", CountryName) - 1)
        End If

     If Not chkCompulsory(Country_FSI(i)) Then
        Msgbox_FSI = Msgbox_FSI + "* Country code at Sr. No  " & i & "  in Sheet  FSI  is Mandatory" & Chr(13)
        ValidateCountry_FSI = False
        Exit Function
     End If
    If Not checkfieldspecialcharacter(Country_FSI(i)) Then
        Msgbox_FSI = Msgbox_FSI + "* Country Code at Sr. No  " & i & " in Sheet  FSI  characters < > & ' " & Chr(34) & " are not allowed in Schedule FSI" & Chr(13)
        ValidateCountry_FSI = False
        Exit Function
     End If
     UpdateProgressBar
 Next
End Function


Function ValidateTaxIdentificationNo_FSI() As Boolean

    ValidateTaxIdentificationNo_FSI = True
    setTblinfo_FSI
    Dim rangecells As Variant
    Dim i As Long
    
    
    ReDim TaxIdentificationNo_FSI(end_FSI)
    
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_TaxIdentificationNo" & i).value

        TaxIdentificationNo_FSI(i) = rangecells
        
     If Not chkCompulsory(TaxIdentificationNo_FSI(i)) Then
         Msgbox_FSI = Msgbox_FSI + "* TaxIdentificationNo at Sr. No  " & i & "  in Sheet  FSI  is Mandatory" & Chr(13)
         ValidateTaxIdentificationNo_FSI = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter_FSI(TaxIdentificationNo_FSI(i)) Then
         Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed in TaxIdentificationNo at Sr. No  " & i & " in Sheet  FSI." & Chr(13)
         ValidateTaxIdentificationNo_FSI = False
         Exit Function
     End If
 Next

End Function

Function ValidateIncFromSal_FSI() As Boolean
    ValidateIncFromSal_FSI = True

    setTblinfo_FSI
    
    'Dim FSI_DTAASAL As Variant
    Dim rangecells As Variant
    Dim rangecells11 As Variant
    Dim rangecells12 As Variant
    Dim rangecells13 As Variant
    Dim rangecells14 As Variant
    
    Dim i As Long
    ReDim IncFromSal_FSI(end_FSI)
    ReDim TXNOSAL_FSI(end_FSI)
    ReDim TXNISAL_FSI(end_FSI)
    ReDim FSI_TXRSAL(end_FSI)
    ReDim DTAASAL_FSI(end_FSI)
    ReDim TXRSAL_FSI(end_FSI)
    
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_IncFromSal" & i).value
    rangecells11 = Sheet40.Range("FSI_TXNOSAL" & i).value
    rangecells12 = Sheet40.Range("FSI_TXNISAL" & i).value
    rangecells13 = Sheet40.Range("FSI_DTAASAL" & i).value
    rangecells14 = Sheet40.Range("FSI_TXRSAL" & i).value
        IncFromSal_FSI(i) = rangecells
        TXNOSAL_FSI(i) = rangecells11
        TXNISAL_FSI(i) = rangecells12
        DTAASAL_FSI(i) = rangecells13
        TXRSAL_FSI(i) = rangecells14
        If Not chkCompulsory(IncFromSal_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Income From Salary at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncFromSal_FSI = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(IncFromSal_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed in Income From Salary at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromSal_FSI = False
            Exit Function
        End If
        
    If (IncFromSal_FSI(i) > 0) Then
        If Not chkCompulsory(TXNOSAL_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India under Salary at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
            ValidateIncFromSal_FSI = False
        End If
        If Not checkfieldspecialcharacter(TXNOSAL_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "*  Special Characters are not allowed Tax paid outside India under Salary at Sr. No  " & i & " in Sheet  FSI." & Chr(13)
            ValidateIncFromSal_FSI = False
        End If
    
        If Not chkCompulsory(TXNISAL_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under Salary at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
            ValidateIncFromSal_FSI = False
        End If
        If Not checkfieldspecialcharacter(TXNISAL_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax payable on such income under Salary at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromSal_FSI = False
        End If
        
        If Not chkCompulsory(DTAASAL_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Relevant article of DTAA if relief claimed u/s 90 or 90A under Salary at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
            ValidateIncFromSal_FSI = False
        End If
        If Not checkfieldspecialcharacter_FSI(DTAASAL_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Relevant article of DTAA if relief claimed u/s 90 or 90A under Salary at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromSal_FSI = False
        End If
    End If
 Next


End Function

Function ValidateIncFromHP_FSI() As Boolean
    ValidateIncFromHP_FSI = True
    setTblinfo_FSI
    Dim rangecells As Variant
    Dim rangecells1 As Variant
    Dim rangecells2 As Variant
    Dim rangecells3 As Variant
    Dim rangecells4 As Variant
    'Dim FSI_TXNOHP, FSI_TXNIHP, FSI_DTAAHP As Variant
    Dim i As Long

    ReDim IncFromHP_FSI(end_FSI)
    ReDim TXNOHP_FSI(end_FSI)
    ReDim TXNIHP_FSI(end_FSI)
    ReDim TXRHP_FSI(end_FSI)
    ReDim DTAAHP_FSI(end_FSI)
    
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_IncFromHP" & i).value
    rangecells1 = Sheet40.Range("FSI_TXNOHP" & i).value
    rangecells2 = Sheet40.Range("FSI_TXNIHP" & i).value
    rangecells3 = Sheet40.Range("FSI_DTAAHP" & i).value
    rangecells4 = Sheet40.Range("FSI_TXRHP" & i).value
        
        IncFromHP_FSI(i) = rangecells
        TXNOHP_FSI(i) = rangecells1
        TXNIHP_FSI(i) = rangecells2
        TXRHP_FSI(i) = rangecells4
        DTAAHP_FSI(i) = rangecells3
        
        
        If Not chkCompulsory(IncFromHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Income From HP at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncFromHP_FSI = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(IncFromHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Income From HP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromHP_FSI = False
            Exit Function
        End If
        
    If (IncFromHP_FSI(i) > 0) Then
        If Not chkCompulsory(TXNOHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India under HP at Sr. No  " & i & "  in Sheet  FSI  is Mandatory. Please fill zero if no income" & Chr(13)
            ValidateIncFromHP_FSI = False
        End If
        If Not checkfieldspecialcharacter(TXNOHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax paid outside India under HP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromHP_FSI = False
        End If
    
        If Not chkCompulsory(TXRHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under HP at Sr. No  " & i & "  in Sheet  FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncFromHP_FSI = False
        End If
        If Not checkfieldspecialcharacter(TXRHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax payable on such income under HP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromHP_FSI = False
        End If
        
'Change-4, ANK-027, 25.05.2022
        If Not chkCompulsory(TXNIHP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under HP at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
                ValidateIncFromHP_FSI = False
            End If
            If Not checkfieldspecialcharacter(TXNIHP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax payable on such income under HP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
                ValidateIncFromHP_FSI = False
            End If
'---end change---
        
        If Not chkCompulsory(DTAAHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Relevant article of DTAA if relief claimed u/s 90 or 90A under HP at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncFromHP_FSI = False
        End If
        If Not checkfieldspecialcharacter_FSI(DTAAHP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Relevant article of DTAA if relief claimed u/s 90 or 90A under HP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
            ValidateIncFromHP_FSI = False
        End If
    End If
    Next
End Function

Function ValidateIncFromBusiness_FSI() As Boolean
    ValidateIncFromBusiness_FSI = True
    setTblinfo_FSI
    Dim rangecells As Variant
    Dim rangecells1 As Variant
    Dim rangecells2 As Variant
    Dim rangecells3 As Variant
    Dim rangecells4 As Variant
    'Dim FSI_TXNOBusiness, FSI_TXNIBusiness, FSI_DTAABusiness As Variant
    Dim i As Long
    
    
    ReDim IncFromBP_FSI(end_FSI)
    ReDim TXNOBP_FSI(end_FSI)
    ReDim TXNIBP_FSI(end_FSI)
    ReDim TXRBP_FSI(end_FSI)
    ReDim DTAABP_FSI(end_FSI)
    
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_IncFromBP" & i).value
    rangecells1 = Sheet40.Range("FSI_TXNOBP" & i).value
    rangecells2 = Sheet40.Range("FSI_TXNIBP" & i).value
    rangecells3 = Sheet40.Range("FSI_DTAABP" & i).value
    rangecells4 = Sheet40.Range("FSI_TXRBP" & i).value
        IncFromBP_FSI(i) = rangecells
        TXNOBP_FSI(i) = rangecells1
        TXNIBP_FSI(i) = rangecells2
        DTAABP_FSI(i) = rangecells3
        TXRBP_FSI(i) = rangecells4
        
        If Not chkCompulsory(IncFromBP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Income From Business at Sr.No  " & i & " in Sheet FSI is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncFromBusiness_FSI = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(IncFromBP_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Income From Business at Sr.No  " & i & " in Sheet FSI." & Chr(13)
             ValidateIncFromBusiness_FSI = False
             Exit Function
        End If
    
        If (IncFromBP_FSI(i) > 0) Then
            If Not chkCompulsory(TXNOBP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India under BP at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
                ValidateIncFromBusiness_FSI = False
              End If
            If Not checkfieldspecialcharacter(TXNOBP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax paid outside India under BP at Sr. No  " & i & " in Sheet FSI. Please fill zero if no income" & Chr(13)
                ValidateIncFromBusiness_FSI = False
            End If
        
            If Not chkCompulsory(TXNIBP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under BP at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
                ValidateIncFromBusiness_FSI = False
            End If
            If Not checkfieldspecialcharacter(TXNIBP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax payable on such income under BP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
                ValidateIncFromBusiness_FSI = False
            End If
            
            If Not chkCompulsory(DTAABP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Relevant article of DTAA if relief claimed u/s 90 or 90A under BP at Sr. No  " & i & "  in Sheet FSI  is Mandatory. Please fill zero if no income" & Chr(13)
                ValidateIncFromBusiness_FSI = False
            End If
            If Not checkfieldspecialcharacter_FSI(DTAABP_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Relevant article of DTAA if relief claimed u/s 90 or 90A under BP at Sr. No  " & i & " in Sheet FSI." & Chr(13)
                ValidateIncFromBusiness_FSI = False
            End If
        End If
    
    Next
End Function

Function ValidateIncCapGain_FSI() As Boolean
    ValidateIncCapGain_FSI = True
    setTblinfo_FSI
    Dim rangecells As Variant
    Dim rangecells1 As Variant
    Dim rangecells2 As Variant
    Dim rangecells3 As Variant
    Dim rangecells4 As Variant
    'Dim FSI_TXNOCapGain, FSI_TXNICapGain As Variant
    Dim i As Long
    
    
    ReDim IncCapGain_FSI(end_FSI)
    ReDim TXNOCapGain_FSI(end_FSI)
    ReDim TXNICapGain_FSI(end_FSI)
    ReDim TXRCapGain_FSI(end_FSI)
    ReDim DTAACapGain_FSI(end_FSI)

    
    
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_IncCapGain" & i).value
    rangecells1 = Sheet40.Range("FSI_TXNOCapGain" & i).value
    rangecells2 = Sheet40.Range("FSI_TXNICapGain" & i).value
    rangecells3 = Sheet40.Range("FSI_DTAACapGain" & i).value
    rangecells4 = Sheet40.Range("FSI_TXRCapGain" & i).value
    
        IncCapGain_FSI(i) = rangecells
        TXNOCapGain_FSI(i) = rangecells1
        TXNICapGain_FSI(i) = rangecells2
        DTAACapGain_FSI(i) = rangecells3
        TXRCapGain_FSI(i) = rangecells4
        
     If Not chkCompulsory(IncCapGain_FSI(i)) Then
         Msgbox_FSI = Msgbox_FSI + "* Income from Capital Gain at Sr.No" & i & " in Sheet FSI is Mandatory.Please fill zero if no income" & Chr(13)
         ValidateIncCapGain_FSI = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(IncCapGain_FSI(i)) Then
         Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Income from Capital Gain at Sr.No" & i & " in Sheet FSI." & Chr(13)
          ValidateIncCapGain_FSI = False
          Exit Function
     End If
 
        If (IncCapGain_FSI(i)) > 0 Then
            If Not chkCompulsory(TXNOCapGain_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India under CG at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
                ValidateIncCapGain_FSI = False
            End If
            If Not checkfieldspecialcharacter(TXNOCapGain_FSI(i)) Then
                Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax paid outside India under CG at Sr. No  " & i & " in Sheet FSI." & Chr(13)
                ValidateIncCapGain_FSI = False
            End If
    

           If Not chkCompulsory(TXNICapGain_FSI(i)) Then
               Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under CG at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
               ValidateIncCapGain_FSI = False
           End If
           If Not checkfieldspecialcharacter(TXNICapGain_FSI(i)) Then
               Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Tax payable on such income under CG at Sr. No  " & i & " in Sheet FSI." & Chr(13)
               ValidateIncCapGain_FSI = False
           End If
           
           If Not chkCompulsory(DTAACapGain_FSI(i)) Then
               Msgbox_FSI = Msgbox_FSI + "* Relevant article of DTAA if relief claimed u/s 90 or 90A under CG at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
               ValidateIncCapGain_FSI = False
           End If
           If Not checkfieldspecialcharacter_FSI(DTAACapGain_FSI(i)) Then
               Msgbox_FSI = Msgbox_FSI + "* Special Characters are not allowed Relevant article of DTAA if relief claimed u/s 90 or 90A under CG at Sr. No  " & i & " in Sheet FSI." & Chr(13)
               ValidateIncCapGain_FSI = False
           End If
        End If
 Next
End Function

Function ValidateIncOthSrc_FSI() As Boolean
ValidateIncOthSrc_FSI = True
    setTblinfo_FSI
    Dim rangecells As Variant
    Dim rangecells1 As Variant
    Dim rangecells2 As Variant
    Dim rangecells3 As Variant
    Dim rangecells4 As Variant
    'Dim FSI_TXNOOthSrc, FSI_TXNIOthSrc, FSI_DTAAOthSrc As Variant
    Dim i As Long
    
    ReDim IncOthSrc_FSI(end_FSI)
    ReDim TXNOOthSrc_FSI(end_FSI)
    ReDim TXNIOthSrc_FSI(end_FSI)
    ReDim TXROthSrc_FSI(end_FSI)
    ReDim DTAAOthSrc_FSI(end_FSI)
    
    For i = 1 To end_FSI
    rangecells = Sheet40.Range("FSI_IncOthSrc" & i).value
    rangecells1 = Sheet40.Range("FSI_TXNOOthSrc" & i).value
    rangecells2 = Sheet40.Range("FSI_TXNIOthSrc" & i).value
    rangecells3 = Sheet40.Range("FSI_DTAAOthSrc" & i).value
    rangecells4 = Sheet40.Range("FSI_TXROthSrc" & i).value
    
        IncOthSrc_FSI(i) = rangecells
        TXNOOthSrc_FSI(i) = rangecells1
        TXNIOthSrc_FSI(i) = rangecells2
        DTAAOthSrc_FSI(i) = rangecells3
        TXROthSrc_FSI(i) = rangecells4
        
     If Not chkCompulsory(IncOthSrc_FSI(i)) Then
         Msgbox_FSI = Msgbox_FSI + "* Income From OS at Sr. No" & i & "  in Sheet Schedule FSI  is Mandatory.Please fill zero if no income" & Chr(13)
         ValidateIncOthSrc_FSI = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(IncOthSrc_FSI(i)) Then
         Msgbox_FSI = Msgbox_FSI + "* Income From OS at Sr. No  " & i & " in Sheet FSI  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
          ValidateIncOthSrc_FSI = False
          Exit Function
     End If
 
        
     If (IncOthSrc_FSI(i)) > 0 Then
        
        If Not chkCompulsory(TXNOOthSrc_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India under OS at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncOthSrc_FSI = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(TXNOOthSrc_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India under OS at Sr. No  " & i & " in Sheet FSI  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
            ValidateIncOthSrc_FSI = False
            Exit Function
        End If
    
        If Not chkCompulsory(TXNIOthSrc_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under OS at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncOthSrc_FSI = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(TXNIOthSrc_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under OS at Sr. No  " & i & " in Sheet FSI  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
            ValidateIncOthSrc_FSI = False
            Exit Function
        End If
        
        If Not chkCompulsory(DTAAOthSrc_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Relevant article of DTAA if relief claimed u/s 90 or 90A under OS at Sr. No  " & i & "  in Sheet FSI  is Mandatory.Please fill zero if no income" & Chr(13)
            ValidateIncOthSrc_FSI = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter_FSI(DTAAOthSrc_FSI(i)) Then
            Msgbox_FSI = Msgbox_FSI + "* Relevant article of DTAA if relief claimed u/s 90 or 90A under OS at Sr. No  " & i & " in Sheet FSI  characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
            ValidateIncOthSrc_FSI = False
            Exit Function
        End If
    End If
 
 Next
End Function

Function validateTotalCountrywise() As Boolean
validateTotalCountrywise = True
Dim i As Long
Dim rangecells1 As Variant
Dim rangecells2 As Variant
Dim rangecells3 As Variant
Dim rangecells4 As Variant
setTblinfo_FSI



ReDim IncT_FSI(end_FSI)
ReDim TXNOT_FSI(end_FSI)
ReDim TXNIT_FSI(end_FSI)
ReDim TXRT_FSI(end_FSI)

For i = 1 To end_FSI
rangecells1 = Sheet40.Range("FSI_IncT" & i).value
rangecells2 = Sheet40.Range("FSI_TXNOT" & i).value
rangecells3 = Sheet40.Range("FSI_TXNIT" & i).value
rangecells4 = Sheet40.Range("FSI_TXRT" & i).value
IncT_FSI(i) = rangecells1
TXNOT_FSI(i) = rangecells2
TXNIT_FSI(i) = rangecells3
TXRT_FSI(i) = rangecells4

If Len(IncT_FSI(i)) > 14 Then
Msgbox_FSI = Msgbox_FSI + "* Income from outside India  at Sr. No  " & i & "  in Sheet FSI cannot be greater than 14 digits" & Chr(13)
validateTotalCountrywise = False
Exit Function
End If

If Len(IncT_FSI(i)) > 14 Then
Msgbox_FSI = Msgbox_FSI + "* Tax paid outside India at Sr. No  " & i & "  in Sheet FSI cannot be greater than 14 digits" & Chr(13)
validateTotalCountrywise = False
Exit Function
End If

If Len(IncT_FSI(i)) > 14 Then
Msgbox_FSI = Msgbox_FSI + "* Tax payable on such income under normal provisions in India  at Sr. No  " & i & "  in Sheet FSI cannot be greater than 14 digits" & Chr(13)
validateTotalCountrywise = False
Exit Function
End If

If Len(IncT_FSI(i)) > 14 Then
Msgbox_FSI = Msgbox_FSI + "* Tax relief available in India at Sr. No  " & i & "  in Sheet FSI cannot be greater than 14 digits" & Chr(13)
validateTotalCountrywise = False
Exit Function
End If

Next

End Function

Sub setTblinfo_FSI()
    Dim rangecells As Variant
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim i As Variant
    
    ccount = 0
    For i = 1 To Sheet40.Range("FSI_Count").value
   
    'mIntcells = Sheet40.Range("FSI_CountryCode" & i).count
    rangecells = Sheet40.Range("FSI_CountryCode" & i).value
    Dim countrycd As Variant
    
        countrycd = rangecells
        If isdropdownblank(countrycd) Then
            countrycd = 0
        End If
        If Not countrycd = 0 Then
            ccount = ccount + 1
            Else
            Exit For
        End If
    
    Next
 end_FSI = ccount
 end_FSI11 = ccount
 End Sub
