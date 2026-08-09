Attribute VB_Name = "mdImportXL"
Option Explicit
Public DestBook As Workbook
Public SrcBook As Workbook
Public retrievedSheetName As Variant
Public rngname_TDS3 As Variant
Public rngname_TDS1 As Variant
Public rngname_TCS As Variant
Sub IMPPrevVersion()
On Error Resume Next
    InitProgBar
    Dim rangenamearr() As Variant
    Dim Filename As Variant
    Dim dfilename, ndfilename, newfilename As Variant
    Dim flag As Boolean
    Dim add As Variant
    Dim rname, ws As Variant
    Dim A As Long
    Dim destadd As Variant
    
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim end_TDS2, rngname_TDS2 As Variant
    
    flag = True
    fmsgboxStatus "Please use this functionality only to 'Import from previous version' of current assessment year." ', vbOKOnly, "Alert"
     fmsgboxStatus "After importing, please check the utility, to ensure all rows are imported." ', vbOKOnly, "Alert"
    
    Filename = cmdFileDialog()
    noOfProcessMain = 2
    mainProcCaption = "Importing Excel"
    ProgressFrameCaption = "Import"

    If Filename <> "" Then UserForm1.Show vbModeless
    
    If Not Filename = "" Then
    
        Filename = Split(Filename, "\")
        newfilename = Filename(UBound(Filename))
        cnt = 0
        Application.ScreenUpdating = False
        Set SrcBook = Workbooks.Open(newfilename)
        Set DestBook = ThisWorkbook
        If Err = 1004 Then ProgressBarHide
        
        dfilename = Split((DestBook.FullName), "\")
        ndfilename = dfilename(UBound(dfilename))
        Application.ScreenUpdating = False
        subProcCaption = "Importing Tables"
        noOfProcessSub = 23
        If newfilename <> ndfilename Then
        
            JuridictionTable
            UpdateProgressBar
            
            DirectorCompany
            UpdateProgressBar
            
            Partnership_Firm
            UpdateProgressBar
            
            Trading_Account
            UpdateProgressBar
            
            Trading_Account1
            UpdateProgressBar
            
            Unlisted_Shares
            UpdateProgressBar
            
            ProfitandLoss1
            UpdateProgressBar
            
            ProfitandLoss2
            UpdateProgressBar
            
            ProfitandLoss3
            UpdateProgressBar
            
            ProfitandLoss4
            UpdateProgressBar
            
            ProfitandLoss5
            UpdateProgressBar
            
            Aufit_Furnishing_Report
            UpdateProgressBar
            
'            Audit_Act_Table
'            UpdateProgressBar
            
            CGTableExtend1
            UpdateProgressBar
            
            CGTableExtend2
            UpdateProgressBar
            
            CGTableExtend3
            UpdateProgressBar
            
            CGTableExtend4
            UpdateProgressBar
            
            CGTableExtend5
            UpdateProgressBar
            
            'CGTableExtend6
            UpdateProgressBar
            OSTableExtend1
            
            UpdateProgressBar
            OS1TableExtend1
            
            OSTableExtend11
            UpdateProgressBar
            
            OS1TableExtend111
            UpdateProgressBar
            
            OSTableExtend2
            UpdateProgressBar
            
            UDTableExtend
            UpdateProgressBar
            
            ExtendTable10A
            UpdateProgressBar
            
            ExtendTable80
            UpdateProgressBar
            
            RATableExtend
            UpdateProgressBar
            
            ExtendTable_80
            UpdateProgressBar
            
            EITableExtend
            UpdateProgressBar
            
            EITableExtend1
            UpdateProgressBar
            
            EITableExtend2
            UpdateProgressBar
            
             Schedule_TPSA
            UpdateProgressBar
            
            TRExtendTable
            UpdateProgressBar
            
            FAExtendTable
            UpdateProgressBar
            
            ExtendTableSPI
            UpdateProgressBar
            
            ExtendTableIF
            UpdateProgressBar
            
            ImportBP1
            UpdateProgressBar
            
            ImportGST
            UpdateProgressBar
            
            ImportBP2
            UpdateProgressBar
            
            ExtendITTable
            UpdateProgressBar
            
            ExtendTDSTable
            UpdateProgressBar
            
            ExtendTableTCS
            UpdateProgressBar
            
            ExtendTable_112A
            UpdateProgressBar
            

            
            ExtendTable_Salary
            UpdateProgressBar
             
            ExtendTable_115AD
            UpdateProgressBar
            
            ExtendHPTable
            UpdateProgressBar
            
            ExtendFSITable
            UpdateProgressBar
            
            ExtendSalTable
            UpdateProgressBar
            
            
            ExtendBankAccount
            
            UpdateProgressBar
            
            'ExtendTableBA2
            UpdateProgressBar
            
            ExtendPTITable
            UpdateProgressBar
            
            ExtendAL1Table
            UpdateProgressBar
            
            ExtendAL2Table
            UpdateProgressBar
            
            
           Application.EnableEvents = False
            subProcCaption = "Importing Static Contents"
            noOfProcessSub = Workbooks(newfilename).Names.count
            For Each rname In Workbooks(newfilename).Names
              retrievedSheetName = getSheetName(rname)
              Application.ScreenUpdating = False
            
            If retrievedSheetName <> "" Then
            If rname.name <> "os.InterestGross" And rname.name <> "SI_111_taxableInc" And rname.name <> "Os.taxthereon.111" And rname.name <> "PL.GrossReceipt" Then
            
              If DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True Then
                If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False Then
                   DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
               
                    
                    If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Locked = False
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Interior.Color = (&HCCFFCC)
                    Else
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HCCFFCC)
                    End If

                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                    
                    DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                Else
                    
                End If
                
              ElseIf DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False Then

                If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True Then
                   DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate


                    If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Locked = True
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Interior.Color = (&HD8D8D8)
                    Else
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HD8D8D8)
                    End If

                    'DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HD8D8D8)
                    
                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues

                    DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate


                End If
                
                
                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.value = SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.value
            End If
              
              If DestBook.Worksheets(retrievedSheetName).Range(rname.name).EntireRow.Hidden = True Then
                    If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).EntireRow.Hidden = False Then
                        DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
                            DestBook.Worksheets(retrievedSheetName).Range(rname.name).EntireRow.Hidden = False
                        DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                    End If
              End If
            End If
            End If
            UpdateProgressBar
            Next

            Dim rangecells As Range
                     Dim chCells As Variant
                     Set rangecells = DestBook.Sheets("Profit and Loss").Range("PL_Address_State").Cells

                    For Each chCells In rangecells
                    TemplockPL (chCells.AddressLocal)
                    Next

            Dim rangecells1 As Range
                     Dim chcells1 As Variant
                     Set rangecells1 = DestBook.Sheets("TDS").Range("TDS2.TDSCreditName").Cells

                    For Each chcells1 In rangecells1
                    Templock2 (chcells1.AddressLocal)
                    Next

            Dim rangecells2 As Range
                     Dim chcells2 As Variant
                     Set rangecells2 = SrcBook.Sheets("TDS").Range("TDS3.TDSCreditName").Cells

                    For Each chcells2 In rangecells2
                    Templock3 (chcells2.AddressLocal)
                    Next
                
             Dim rangecells3 As Range
                     Dim chcells3 As Variant
                     Set rangecells3 = SrcBook.Sheets("TDS").Range("TDS2.ClaimedInOwnHands").Cells

                    For Each chcells3 In rangecells3
                    Templock4 (chcells3.AddressLocal)
                    Next
            
            Dim rangecells4 As Range
                     Dim chcells4 As Variant
                     Set rangecells4 = SrcBook.Sheets("TDS").Range("TDS3.ClaimedInOwnHands").Cells

                    For Each chcells4 In rangecells4
                    Templock5 (chcells4.AddressLocal)
                    Next
                    

                 returnfilesecimport
                   salary_Lock
                     SrcBook.Sheets("Part B - TI TTI").Unprotect Password:=getmsgstate
                     DestBook.Sheets("Part B - TI TTI").Unprotect Password:=getmsgstate
                    If SrcBook.Worksheets("Part B - TI TTI").Range("Sheet9.intrstPayUs234F").MergeArea.Locked = False Then
                      DestBook.Worksheets("Part B - TI TTI").Range("Sheet9.intrstPayUs234F").MergeArea.Locked = False
                       DestBook.Worksheets("Part B - TI TTI").Range("Sheet9.intrstPayUs234F").MergeArea.Interior.ColorIndex = RGB(255, 255, 204)
                      
                    End If

                     SrcBook.Sheets("Part B - TI TTI").Protect Password:=getmsgstate
                     DestBook.Sheets("Part B - TI TTI").Protect Password:=getmsgstate

            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Sheet39.AddTotalSalary
            
            '112A
            Sheet53.Unprotect Password:=getmsgstate
            SrcBook.Sheets("Schedule 112A").Range("TotalSaleValue_112A").Copy
            DestBook.Sheets("Schedule 112A").Range("TotalSaleValue_112A").PasteSpecial xlValues
            Sheet53.Protect Password:=getmsgstate
            '115AD
            Sheet54.Unprotect Password:=getmsgstate
            SrcBook.Sheets("Schedule 115AD(1)(iii) proviso").Range("TotalSaleValue_115AD").Copy
            DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Range("TotalSaleValue_115AD").PasteSpecial xlValues
            Sheet54.Protect Password:=getmsgstate
            'TDS
            'CR40
            md115AD.validateImport115AD_1
            md112A.validateImport112A
            
            mdCFL.addCFLformula
            
            LockUnlockTables
            ProgressBarHide
            fmsgboxStatus "Import Completed"
            'DestBook.Save
            DestBook.Worksheets("PART A - GENERAL").Select
            Set SrcBook = Nothing
        Else
            fmsgboxStatus "Source file must not have same name As destination File"
            ProgressBarHide
        End If
    End If

End Sub


Sub setTblinfo_Per10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.DoneeName").count
    Set rangecells = Sheet20.Range("Per10080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G = ccount
    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
End Sub
Sub setTblinfo_PerNO5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.DoneeName").count
    Set rangecells = Sheet20.Range("PerNO5080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G = ccount
    rngname_PerNO5080G = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;"
End Sub
Sub setTblinfo_PerYES10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.DoneeName").count
    Set rangecells = Sheet20.Range("PerYES10080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G = ccount
    rngname_PerYES10080G = "PerYES10080G.DoneeName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;"
End Sub

Sub setTblinfo_Per5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.DoneeName").count
    Set rangecells = Sheet20.Range("Per5080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G = ccount
    rngname_Per5080G = "Per5080G.DoneeName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt"
End Sub
Sub setTblinfo_SPI()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet21.Range("SPI.SpecifiedPersonName").count
    Set rangecells = Sheet21.Range("SPI.SpecifiedPersonName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_SPI = ccount
    rngname_SPI = "SPI.SpecifiedPersonName;SPI.PANofSpecPerson;SPI.ReltnShip;SPI.NatureOfInc;SPI.AmtIncluded;"
End Sub
Sub setTblinfo_SI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet21.Range("SI.SecCode").count
    Set rangecells = Sheet21.Range("SI.SecCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_SI = ccount
    rngname_SI = "SI.SecCode;SI.SplRatePercent;SI.SplRateInc;SI.SplRateIncCalc;SI.SplRateIncTax;"
End Sub

Sub setTblinfo_IF()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet21.Range("IF.FirmName").count
    Set rangecells = Sheet21.Range("IF.FirmName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_IF = ccount
    rngname_IF = "IF.FirmName;IF.FirmPAN;IF.IsLiableToAudit;IF.Is92EApplicable;IF.ProfitSharePercent;IF.ProfitShareAmt;IF.FirmCapBalOn31Mar;isLiableForAudit;LiableFor92E;"
End Sub
Sub setTableInfo_IT()
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
    end_IT = ccount
    rngname_IT = "IT.BSRCode;IT.DateDep;IT.SrlNoOfChaln;IT.Amt;IT.FormulaOFS;FormulaOfQ;FormulaOfSAT;FormulaOfSAT1;FormulaOfExSAT;FormulaOfExSAT1;"
End Sub


Sub setTblinfo_FSIimportXML()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet26.Range("FSI_CountryCode").count
    Set rangecells = Sheet26.Range("FSI_CountryCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_FSI = ccount
    rngname_FSI = "FSI_CountryCode;FSI_TaxIdentificationNo;FSI_IncFromSal;FSI_TXNOSAL;FSI_TXNISAL;FSI_TXRSAL;FSI_DTAASAL;FSI_IncFromHP;FSI_TXNOHP;FSI_TXRHP;FSI_DTAAHP;FSI_IncFromBP;FSI_TXNOBP;FSI_TXNIBP;FSI_TXRBP;FSI_DTAABP;FSI_IncCapGain;FSI_TXNOCapGain;FSI_TXnICapGain;FSI_TXRCapGain;FSI_DTAACapGain;FSI_TXNOOthSrc;FSI_TXNIOthSrc;FSI_DTAAOthSrc;"
End Sub
'dpk4
Sub setTblinfo_OS_xl()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.othDesc").count
    Set rangecells = Sheet14.Range("os.othDesc").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os = ccount
    rngname_os = "os.othDesc;os.OtherSections;"
End Sub
Sub setTableinfo_TDS1import()
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet28.Range("TDS1.TAN").count
    Set rangecells = Sheet28.Range("TDS1.TAN").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS1 = ccount
    rngname_TDS1 = "TDS1.TAN;TDS1.EmployerOrDeductorOrCollecterName;TDS1.IncChrgSal;TDS1.TotalTDSSal;"
End Sub

Sub setTableinfo_TDS2()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.TDSCreditName").count
 Set rangecells = Sheet28.Range("TDS2.TDSCreditName").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS2 = ccount
 rngname_TDS2 = "TDS2.TDSCreditName;TDS2.PanOtherPerson;TDSAadharOfOther_TDS2;TDS2.GrossAmount;TDS2.HeadOfIncome;TDS2.TAN;TDS2.DeductedYear;TDS2.BroughtFwdTDSAmt;TDS2.DeductedInOwnHands;TDS2.DeductedIncome5A;TDS2.DeductedTDS5A;TDS2.ClaimedInOwnHands;TDS2.ClaimedIncome5A;TDS2.ClaimedTDS5A;TDS2.ClaimedPAN5A;ClaimedAadhar5A_TDS2;TDS2.AmtCarriedForward;"
 End Sub
Sub setTableinfo_TDS3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.TDSCreditName").count
 Set rangecells = Sheet28.Range("TDS3.TDSCreditName").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS3 = ccount
 rngname_TDS3 = "TDS3.TDSCreditName;TDS3.GrossAmount;TDS3.HeadOfIncome;TDS3.PanOtherPerson;Aadhar_TDS3;TDS3.PAN;TDS3.Aadhaar;TDS3.DeductedYear;TDS3.BroughtFwdTDSAmt;TDS3.DeductedInOwnHands;TDS3.DeductedIncome5A;TDS3.DeductedTDS5A;TDS3.ClaimedInOwnHands;TDS3.ClaimedIncome5A;TDS3.ClaimedTDS5A;TDS3.ClaimedPAN5A;ClaimedAadhar5A_TDS3;TDS3.AmtCarriedForward;"
 End Sub

Sub setTblinfo_TCS2XL()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").count
    Set rangecells = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TCS2 = ccount
'    rngname_TCS = "TCS.TAN;TCS.EmployerOrDeductorOrCollecterName;TCS.DeductedYear;TCS.BroughtFwdTDSAmt;TCS.TotalTCS;TCS.AmtTCSClaimedThisYear;TCS.AmtCarriedFwd;"
    rngname_TCS = "TCS.TAN;TCS.EmployerOrDeductorOrCollecterName;TCS.TotalTCS;TCS.AmtTCSClaimedThisYear;TCS.AmtCarriedFwd;"

End Sub


Function InsertRowsToImport(Optional vRows As Long = 0)
    Dim x As Long
    Dim strpassword As String
    Dim sht As Worksheet, shts() As String, i As Long
    
    strpassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=strpassword
    
    ActiveCell.EntireRow.Select
    
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
       Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        i = i + 1
        shts(i) = sht.name
        
        x = Sheets(sht.name).UsedRange.Rows.count
        
        ActiveSheet.Unprotect Password:=strpassword
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        ActiveSheet.Unprotect Password:=strpassword
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault

        On Error Resume Next
            ActiveSheet.Unprotect Password:=strpassword
            Selection.Offset(1).Resize(vRows).EntireRow. _
            SpecialCells(xlConstants).ClearContents
    Next sht
    ActiveSheet.Protect Password:=strpassword
End Function

Function cmdFileDialog() As String
    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
 
    cmdFileDialog = ""
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
      .AllowMultiSelect = False
      .Filters.Clear
      .Filters.add "Microsoft Office Excel Workbook", "*.xls,*.xlsm,*.xlsb"
        If .Show = True Then
            For Each varFile In .SelectedItems
               cmdFileDialog = varFile
            Next
        End If
    End With
End Function
Function JuridictionTable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "NRI_Juridiction"
                    sfirstbound = SrcBook.Sheets("PART A - General").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("PART A - General").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim rangename_JuridictionTable As Variant
                    rangename_JuridictionTable = "NRI_Juridiction;NRI_Taxpayer;"
                    cnt = SrcBook.Sheets("PART A - General").Range(name).count
                    dcnt = DestBook.Sheets("PART A - General").Range(name).count
                    DestBook.Sheets("PART A - General").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("PART A - General").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport ((cnt - dcnt) / 21)
                        Call ExendRangeNameToTable((cnt - dcnt) / 21, rangename_JuridictionTable)
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    End If
                
                
End Function


Function DirectorCompany()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Sheet1.CompanyName"
                    sfirstbound = SrcBook.Sheets("PART A - General").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("PART A - General").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim DirectorTable As Variant
                    DirectorTable = "Sheet1.CompanyName;Sheet1.CompanyType;Sheet1.CompPAN;Sheet1.Shares;Sheet1.DIN;"
                    cnt = SrcBook.Sheets("PART A - General").Range(name).count
                    dcnt = DestBook.Sheets("PART A - General").Range(name).count
                    DestBook.Sheets("PART A - General").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("PART A - General").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport ((cnt - dcnt) / 9)
                        Call ExendRangeNameToTable((cnt - dcnt) / 9, DirectorTable)
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    End If
               
End Function
Function Partnership_Firm()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Sheet1.NameOfFirm"
                    sfirstbound = SrcBook.Sheets("PART A - General").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("PART A - General").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim Partnership_Firm_Table As Variant
                    Partnership_Firm_Table = "Sheet1.NameOfFirm;Sheet1.FirmPAN;"
                    cnt = SrcBook.Sheets("PART A - General").Range(name).count
                    dcnt = DestBook.Sheets("PART A - General").Range(name).count
                    DestBook.Sheets("PART A - General").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("PART A - General").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport ((cnt - dcnt) / 17)
                        Call ExendRangeNameToTable((cnt - dcnt) / 17, Partnership_Firm_Table)
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    End If
               
End Function
Function Unlisted_Shares()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Share_NameofCompany"
                    sfirstbound = SrcBook.Sheets("PART A - General").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("PART A - General").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim Unlisted_Shares_Table As Variant
                    Unlisted_Shares_Table = "Share_NameofCompany;Share_TypeofCompany;Share_PAN;Share_OP_Count;Share_OP_Cost;Share_Acquired_Count;Share_Acquired_Date;Share_Acquired_FaceValue;Share_Acquired_Issue;Share_Acquired_Cost;Share_Transferred_Count;Share_Transferred_Sale;Share_CB_Count;Share_CB_Cost;"
                    cnt = SrcBook.Sheets("PART A - General").Range(name).count
                    dcnt = DestBook.Sheets("PART A - General").Range(name).count
                    DestBook.Sheets("PART A - General").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("PART A - General").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, Unlisted_Shares_Table)
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    End If
               
End Function
Function Trading_Account()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "TradingAcc_RevenueNature"
                    sfirstbound = SrcBook.Sheets("Trading Account").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Trading Account").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim Trading_Account_Table As Variant
                    Trading_Account_Table = "TradingAcc_RevenueNature;TradingAcc_RevenueAmount;"
                    cnt = SrcBook.Sheets("Trading Account").Range(name).count
                    dcnt = DestBook.Sheets("Trading Account").Range(name).count
                    DestBook.Sheets("Trading Account").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Trading Account").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, Trading_Account_Table)
                        SrcBook.Sheets("Trading Account").Range(name).Copy
                        DestBook.Sheets("Trading Account").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Trading Account").Range(name).Copy
                        DestBook.Sheets("Trading Account").Range(name).PasteSpecial xlValues
                    End If
               
End Function
Function Trading_Account1()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "TradingAcc_NatureOfDirectExpenses"
                    sfirstbound = SrcBook.Sheets("Trading Account").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Trading Account").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim Trading_Account_Table1 As Variant
                    Trading_Account_Table1 = "TradingAcc_NatureOfDirectExpenses;TradingAcc_OtherDirectAmount;"
                    cnt = SrcBook.Sheets("Trading Account").Range(name).count
                    dcnt = DestBook.Sheets("Trading Account").Range(name).count
                    DestBook.Sheets("Trading Account").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Trading Account").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, Trading_Account_Table1)
                        SrcBook.Sheets("Trading Account").Range(name).Copy
                        DestBook.Sheets("Trading Account").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Trading Account").Range(name).Copy
                        DestBook.Sheets("Trading Account").Range(name).PasteSpecial xlValues
                    End If
               
End Function

Function ProfitandLoss1()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Nature_Name2"
                    sfirstbound = SrcBook.Sheets("Profit and Loss").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Profit and Loss").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim ProfitandLoss_Table1 As Variant
                    ProfitandLoss_Table1 = "Nature_Name2;Nature_Amt2;"
                    cnt = SrcBook.Sheets("Profit and Loss").Range(name).count
                    dcnt = DestBook.Sheets("Profit and Loss").Range(name).count
                    DestBook.Sheets("Profit and Loss").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Profit and Loss").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, ProfitandLoss_Table1)
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    End If
               
End Function

Function ProfitandLoss2()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Nature_Name3"
                    sfirstbound = SrcBook.Sheets("Profit and Loss").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Profit and Loss").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim ProfitandLoss_Table2 As Variant
                    ProfitandLoss_Table2 = "Nature_Name3;Nature_Amt3;"
                    cnt = SrcBook.Sheets("Profit and Loss").Range(name).count
                    dcnt = DestBook.Sheets("Profit and Loss").Range(name).count
                    DestBook.Sheets("Profit and Loss").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Profit and Loss").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, ProfitandLoss_Table2)
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    End If
               
End Function


Function ProfitandLoss3()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "PLBD.PAN"
                    sfirstbound = SrcBook.Sheets("Profit and Loss").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Profit and Loss").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim ProfitandLoss_Table3 As Variant
                    ProfitandLoss_Table3 = "PLBD.PAN;PLBD.Aadhar;PLBD.Amount;"
                    cnt = SrcBook.Sheets("Profit and Loss").Range(name).count
                    dcnt = DestBook.Sheets("Profit and Loss").Range(name).count
                    DestBook.Sheets("Profit and Loss").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Profit and Loss").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, ProfitandLoss_Table3)
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    End If
               
End Function

Function ProfitandLoss4()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "PL_Name"
                    sfirstbound = SrcBook.Sheets("Profit and Loss").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Profit and Loss").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim ProfitandLoss_Table4 As Variant
                    ProfitandLoss_Table4 = "PL_Name;PL_Address_Flat;PL_Address_Village;PL_Address_Road;PL_Address_Area;PL_Address_City;PL_Address_State;PL_Address_Country;PL_Address_Pin;PL_Address_Zip;PL_Amount;"
                    cnt = SrcBook.Sheets("Profit and Loss").Range(name).count
                    dcnt = DestBook.Sheets("Profit and Loss").Range(name).count
                    DestBook.Sheets("Profit and Loss").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Profit and Loss").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport ((cnt - dcnt) / 2)
                        Call ExendRangeNameToTable((cnt - dcnt) / 2, ProfitandLoss_Table4)
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    End If
               
End Function

Function ProfitandLoss5()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Sec44AE.RegNo"
                    sfirstbound = SrcBook.Sheets("Profit and Loss").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Profit and Loss").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim ProfitandLoss_Table5 As Variant
                    ProfitandLoss_Table5 = "Sec44AE.RegNo;Sec44AE.OLH;Sec44AE.TonnageCpct;Sheet44AE.NoOfMonths;Sheet44AE.PresumptiveIncome;greater12;lesser12;"
                    cnt = SrcBook.Sheets("Profit and Loss").Range(name).count
                    dcnt = DestBook.Sheets("Profit and Loss").Range(name).count
                    DestBook.Sheets("Profit and Loss").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Profit and Loss").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport ((cnt - dcnt) / 3)
                        Call ExendRangeNameToTable((cnt - dcnt) / 3, ProfitandLoss_Table5)
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Profit and Loss").Range(name).Copy
                        DestBook.Sheets("Profit and Loss").Range(name).PasteSpecial xlValues
                    End If
               
End Function

'new
Function Schedule_TPSA()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "TPSC_BSRCode"
                    sfirstbound = SrcBook.Sheets("TPSA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TPSA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim TPSA_Table As Variant
                    TPSA_Table = "TPSC_BSRCode;TPSC_BankName;TPSC_DateDep;TPSC_SrlNoOfChaln;TPSC_Amt;"
                    cnt = SrcBook.Sheets("TPSA").Range(name).count
                    dcnt = DestBook.Sheets("TPSA").Range(name).count
                    DestBook.Sheets("TPSA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TPSA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, TPSA_Table)
                        SrcBook.Sheets("TPSA").Range(name).Copy
                        DestBook.Sheets("TPSA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TPSA").Range(name).Copy
                        DestBook.Sheets("TPSA").Range(name).PasteSpecial xlValues
                    End If
               
End Function


Function Aufit_Furnishing_Report()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "sheet1.AuditedSection"
                    sfirstbound = SrcBook.Sheets("PART A - General").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("PART A - General").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim Aufit_Furnishing_Report_Table As Variant
                    Aufit_Furnishing_Report_Table = "sheet1.AuditedSection;sheet1.AuditYN;sheet1.DateOfAudit;"
                    cnt = SrcBook.Sheets("PART A - General").Range(name).count
                    dcnt = DestBook.Sheets("PART A - General").Range(name).count
                    DestBook.Sheets("PART A - General").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("PART A - General").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, Aufit_Furnishing_Report_Table)
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    End If
               
End Function

Function Audit_Act_Table()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "Audit.Act"
                    sfirstbound = SrcBook.Sheets("PART A - General").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("PART A - General").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim Audit_Act As Variant
                    Audit_Act = "Audit.Act;Audit.Others;Audit.Sections;Audit.ACTYN;Audit.Date;"
                    cnt = SrcBook.Sheets("PART A - General").Range(name).count
                    dcnt = DestBook.Sheets("PART A - General").Range(name).count
                    DestBook.Sheets("PART A - General").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("PART A - General").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                      ' setTblinfo_OS
                        InsertRowsToImport ((cnt - dcnt) / 9)
                        Call ExendRangeNameToTable((cnt - dcnt) / 9, Audit_Act)
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("PART A - General").Range(name).Copy
                        DestBook.Sheets("PART A - General").Range(name).PasteSpecial xlValues
                    End If
               
End Function

Function CGTableExtend1()

    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "DEDN.DeductedSecCode"
                    sfirstbound = SrcBook.Sheets("CG").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("CG").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("CG").Range(name).count
                    dcnt = DestBook.Sheets("CG").Range(name).count
                    DestBook.Sheets("CG").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("CG").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_exemption
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_exemption)
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    End If
                
End Function

Function CGTableExtend2()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "STCG.A7_DeductionSection_1"
                    sfirstbound = SrcBook.Sheets("CG").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    Dim rngname_STCGvalidate As Variant
                    rngname_STCGvalidate = "STCG.A7_PreviousYer_1;STCG.A7_DeductionSection_1;STCG.A7_AsstAcquiredYer_1;STCG.A7_AmtUtilised_1;STCG.A7_AmtNotUsed_1;"
                    dfirstbound = DestBook.Sheets("CG").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                
                    cnt = SrcBook.Sheets("CG").Range(name).count
                    dcnt = DestBook.Sheets("CG").Range(name).count
                    DestBook.Sheets("CG").Activate
                    
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("CG").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_exemption
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_STCGvalidate)
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    End If
             
End Function

Function CGTableExtend3()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

name = "STCG.A8_Country"

     sfirstbound = SrcBook.Sheets("CG").Range(name).Address
     sTEMP = Split(sfirstbound, "$")
     supperbound = UBound(sTEMP)
     sTEMP = sTEMP(UBound(sTEMP))
     Dim rngname_STCGDTAA As Variant
     rngname_STCGDTAA = "STCG.A8_Country;STCG.A8_DTAA_Article;STCG.A8_CertiStatus;STCG.A8_ItemNo;STCG.A8_StcgAmt;Section_CodeStcgA8;Section_AmountStcgA8;Section_STCGbreakup;"
     dfirstbound = DestBook.Sheets("CG").Range(name).Address
     dTemp = Split(dfirstbound, "$")
     dupperbound = UBound(dTemp)
     ddTemp = dTemp(UBound(dTemp))

     cnt = SrcBook.Sheets("CG").Range(name).count
     dcnt = DestBook.Sheets("CG").Range(name).count
     DestBook.Sheets("CG").Activate
     If (cnt - dcnt) > 0 Then
         DestBook.Sheets("CG").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
         setTblinfo_exemption
         InsertRowsToImport (cnt - dcnt)
         Call ExendRangeNameToTable(cnt - dcnt, rngname_STCGDTAA)
         SrcBook.Sheets("CG").Range(name).Copy
         DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
     Else
         SrcBook.Sheets("CG").Range(name).Copy
         DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
    End If
    
End Function

Function CGTableExtend4()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "LTCG.B9_DeductionSection_1"
                    sfirstbound = SrcBook.Sheets("CG").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    Dim rngname_LTCGvalidate As Variant
                    rngname_LTCGvalidate = "LTCG.B9_PreviousYer_1;LTCG.B9_DeductionSection_1;LTCG.B9_AsstAcquiredYer_1;LTCG.B9_AmtUtilised_1;LTCG.B9_AmtNotUsed_1;"
                    dfirstbound = DestBook.Sheets("CG").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
           
                    cnt = SrcBook.Sheets("CG").Range(name).count
                    dcnt = DestBook.Sheets("CG").Range(name).count
                    DestBook.Sheets("CG").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("CG").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_exemption
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_LTCGvalidate)
                       SrcBook.Sheets("CG").Range(name).Copy
                       DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    End If
         
End Function
Function CGTableExtend5()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
             name = "LTCG.B10_Country"
                    sfirstbound = SrcBook.Sheets("CG").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    Dim rngname_LTCGDTAA As Variant
                    rngname_LTCGDTAA = "LTCG.B10_LtcgAmt;LTCG.B10_ItemNo;LTCG.B10_Country;LTCG.B10_DTAA_Article;LTCG.B10_RatepTreaty;LTCG.B10_CertiStatus;LTCG.B10_SecIT;LTCG.B10_RatePITAct;LTCG.B10_ApplRate;"
                    dfirstbound = DestBook.Sheets("CG").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))

                    cnt = SrcBook.Sheets("CG").Range(name).count
                    dcnt = DestBook.Sheets("CG").Range(name).count
                    DestBook.Sheets("CG").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("CG").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_exemption
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_LTCGDTAA)
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("CG").Range(name).Copy
                        DestBook.Sheets("CG").Range(name).PasteSpecial xlValues
                    End If
                    
End Function


Function OSTableExtend1()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "os.SourceDescription"
                    sfirstbound = SrcBook.Sheets("OS").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("OS").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim rangename_osImport As Variant
                    rangename_osImport = "os.SourceDescription;os.SourceAmount;OS_Section.Validate;os.check;SRCamt;"
                    cnt = SrcBook.Sheets("OS").Range(name).count
                    dcnt = DestBook.Sheets("OS").Range(name).count
                    DestBook.Sheets("OS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("OS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rangename_osImport)
                        SrcBook.Sheets("OS").Range(name).Copy
                        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("OS").Range(name).Copy
                        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
                    End If
                
                
End Function
Function OSTableExtend11()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
name = "os.PTIDescription"
                    sfirstbound = SrcBook.Sheets("OS").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("OS").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim rangename_osImport1 As Variant
                    rangename_osImport1 = "os.PTIDescription;os.PTIAmount;os_PTI_check;OS_PTI_Amount;OS_PTISectionCode;"
                    cnt = SrcBook.Sheets("OS").Range(name).count
                    dcnt = DestBook.Sheets("OS").Range(name).count
                    DestBook.Sheets("OS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("OS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rangename_osImport1)
                        SrcBook.Sheets("OS").Range(name).Copy
                        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("OS").Range(name).Copy
                        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
                    End If
                
                
End Function


Function OS1TableExtend1()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "os.othDesc"
    sfirstbound = SrcBook.Sheets("OS").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("OS").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_osImport As Variant
    rangename_osImport = "os.othDesc;os.OtherSections;"
    cnt = SrcBook.Sheets("OS").Range(name).count
    dcnt = DestBook.Sheets("OS").Range(name).count
    DestBook.Sheets("OS").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("OS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport ((cnt - dcnt) / 2)
        Call ExendRangeNameToTable((cnt - dcnt) / 2, rangename_osImport)
        SrcBook.Sheets("OS").Range(name).Copy
        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("OS").Range(name).Copy
        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
    End If
End Function
Function OS1TableExtend111()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "os.AssYear"
    sfirstbound = SrcBook.Sheets("OS").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("OS").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_osImport As Variant
    rangename_osImport = "os.AssYear;os.IncomeBenefit;os.TaxBenefit;"
    cnt = SrcBook.Sheets("OS").Range(name).count
    dcnt = DestBook.Sheets("OS").Range(name).count
    DestBook.Sheets("OS").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("OS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport (cnt - dcnt)
        Call ExendRangeNameToTable(cnt - dcnt, rangename_osImport)
        SrcBook.Sheets("OS").Range(name).Copy
        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("OS").Range(name).Copy
        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
    End If
End Function

Function OSTableExtend2()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
 name = "os.Amount_income"
                    sfirstbound = SrcBook.Sheets("OS").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("OS").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    Dim rangename_osImportDTAA As Variant
                    rangename_osImportDTAA = "os.Amount_income;os.NatureOfIncome;os.Country;os.Article_DTAA;os.Rate_DTAA;os.TRC;os.SectionAct;os.Rate_ITAct;os.ApplRate;os.DTAAcheck;OS_DTAASectionCode;os.DTAA_AmtChk;"
                    cnt = SrcBook.Sheets("OS").Range(name).count
                    dcnt = DestBook.Sheets("OS").Range(name).count
                    DestBook.Sheets("OS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("OS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_OS
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rangename_osImportDTAA)
                        SrcBook.Sheets("OS").Range(name).Copy
                        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("OS").Range(name).Copy
                        DestBook.Sheets("OS").Range(name).PasteSpecial xlValues
                    End If
               
End Function
Function UDTableExtend()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "UD.AssessmentYear"
    sfirstbound = SrcBook.Sheets("Unabsorbed Depreciation").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("Unabsorbed Depreciation").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_UDImport As Variant
    rangename_UDImport = "UD.AssessmentYear;UD.BF;UD.Setoff;UD2.BF;UD2.Setoff;"
    cnt = SrcBook.Sheets("Unabsorbed Depreciation").Range(name).count
    dcnt = DestBook.Sheets("Unabsorbed Depreciation").Range(name).count
    DestBook.Sheets("Unabsorbed Depreciation").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("Unabsorbed Depreciation").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport ((cnt - dcnt))
        Call ExendRangeNameToTable(cnt - dcnt, rangename_UDImport)
        SrcBook.Sheets("Unabsorbed Depreciation").Range(name).Copy
        DestBook.Sheets("Unabsorbed Depreciation").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("Unabsorbed Depreciation").Range(name).Copy
        DestBook.Sheets("Unabsorbed Depreciation").Range(name).PasteSpecial xlValues
    End If
End Function
Function ExtendTable10A()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

name = "SEZA10.DedFromUndertaking"
 
name = "AA10.DedFromUndertaking"
                    sfirstbound = SrcBook.Sheets("10AA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("10AA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("10AA").Range(name).count
                    dcnt = DestBook.Sheets("10AA").Range(name).count
                    DestBook.Sheets("10AA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("10AA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_AA10
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_AA10)
                        SrcBook.Sheets("10AA").Range(name).Copy
                        DestBook.Sheets("10AA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("10AA").Range(name).Copy
                        DestBook.Sheets("10AA").Range(name).PasteSpecial xlValues
                    End If
                

End Function

Function ExtendTable80()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

name = "Per10080G.DoneeName"
                    sfirstbound = SrcBook.Sheets("80G").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
        
                    dfirstbound = DestBook.Sheets("80G").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
        
                    cnt = SrcBook.Sheets("80G").Range(name).count
                    dcnt = DestBook.Sheets("80G").Range(name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        'setTblinfo_Per10080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationInCash;Per10080G.DonationAmt;Per10080G.TotalDonationAmt;Per10080G.EligibleAmt;")
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    End If
                
        
'Sch 80G- Section B
               name = "PerNO5080G.DoneeName"
                    sfirstbound = SrcBook.Sheets("80G").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                
                    cnt = SrcBook.Sheets("80G").Range(name).count
                    dcnt = DestBook.Sheets("80G").Range(name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        'setTblinfo_PerNO5080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationInCash;PerNO5080G.DonationAmt;PerNO5080G.TotalDonationAmt;PerNO5080G.EligibleAmt;")
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    End If

        
'Sch 80G- Section C

               name = "PerYES10080G.DoneeWithPanName"
                    sfirstbound = SrcBook.Sheets("80G").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80G").Range(name).count
                    dcnt = DestBook.Sheets("80G").Range(name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        'setTblinfo_PerYES10080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationInCash;PerYES10080G.DonationAmt;PerYES10080G.TotalDonationAmt;PerYES10080G.EligibleAmt;")
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    End If
                
        
'Sch 80G- Section D
                name = "Per5080G.DoneeWithPanName"
                    sfirstbound = SrcBook.Sheets("80G").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80G").Range(name).count
                    dcnt = DestBook.Sheets("80G").Range(name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        'setTblinfo_Per5080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationInCash;Per5080G.DonationAmt;Per5080G.EligibleAmt;Per5080G.TotalDonationAmt;Per5080G.DonationInCash1;")
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(name).Copy
                        DestBook.Sheets("80G").Range(name).PasteSpecial xlValues
                    End If
                
End Function

Function ExtendTable_80()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
Dim ranges80 As Variant
Dim i80cntr As Variant
            
ranges80 = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductJKLocUs80_IB_4||IB80.DeductBackStatesUs80_IB_4||IB80.DeductBackDisttUs80_IB_5||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductColdChainUs80_IB_11||IB80.DeductFruitVegUs80_IB_11A||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.DeductInSikkim||IC80.DeductInHimachalP||IC80.DeductInUttaranchal||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura"
ranges80 = Split(ranges80, "||")
For i80cntr = 0 To UBound(ranges80)
name = ranges80(i80cntr)
        sfirstbound = SrcBook.Sheets("80").Range(name).Address
        sTEMP = Split(sfirstbound, "$")
        supperbound = UBound(sTEMP)
        sTEMP = sTEMP(UBound(sTEMP))
        
        dfirstbound = DestBook.Sheets("80").Range(name).Address
        dTemp = Split(dfirstbound, "$")
        dupperbound = UBound(dTemp)
        ddTemp = dTemp(UBound(dTemp))
        
        cnt = SrcBook.Sheets("80").Range(name).count
        dcnt = DestBook.Sheets("80").Range(name).count
        DestBook.Sheets("80").Activate
        If (cnt - dcnt) > 0 Then
            DestBook.Sheets("80").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
            'setTblinfo_Per5080G
            InsertRowsToImport (cnt - dcnt)
            Call ExendRangeNameToTable(cnt - dcnt, ranges80(i80cntr) + ";")
            SrcBook.Sheets("80").Range(name).Copy
            DestBook.Sheets("80").Range(name).PasteSpecial xlValues
        Else
            SrcBook.Sheets("80").Range(name).Copy
            DestBook.Sheets("80").Range(name).PasteSpecial xlValues
        End If
    
Next
           
End Function
Function RATableExtend()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "Name_of_Donee_80GGA"
    sfirstbound = SrcBook.Sheets("RA").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("RA").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_RAImport As Variant
    rangename_RAImport = "Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_Eligible_80GGA;"
    cnt = SrcBook.Sheets("RA").Range(name).count
    dcnt = DestBook.Sheets("RA").Range(name).count
    DestBook.Sheets("RA").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("RA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport ((cnt - dcnt))
        Call ExendRangeNameToTable(cnt - dcnt, rangename_RAImport)
        SrcBook.Sheets("RA").Range(name).Copy
        DestBook.Sheets("RA").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("RA").Range(name).Copy
        DestBook.Sheets("RA").Range(name).PasteSpecial xlValues
    End If
End Function

Function TRExtendTable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
                name = "TR_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TR
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TR)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                
End Function

Function FAExtendTable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

 'Sch FA-A1
                name = "FA_A1_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_A_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "FA_A1_Country;FA_A1_BankName;FA_A1_BankAdd;FA_A1_ZipCode;FA_A1_ForeignAccountNumber;FA_A1_StatusBeneficiary;FA_A1_AccOpeningDate;FA_A1_PeakBal;FA_A1_ClosingBal;FA_A1_Grossinterest;")
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                    
 'Sch FA-A2
                name = "FA_A2_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_A_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "FA_A2_Country;FA_A2_BankName;FA_A2_BankAdd;FA_A2_ZipCode;FA_A2_ForeignAccountNumber;FA_A2_StatusBeneficiary;FA_A2_AccOpeningDate;FA_A2_PeakBal;FA_A2_ClosingBal;FA_A2_Grossinterest;FA_A2_Grossinterest_Nature;")
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                    
 'Sch FA-A3
                name = "FA_A3_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_A_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "FA_A3_Country;FA_A3_BankName;FA_A3_BankAdd;FA_A3_ZipCode;FA_A3_NatureOfEntity;FA_A3_AccOpeningDate;FA_A3_initialvalue;FA_A3_PeakBal;FA_A3_ClosingBal;FA_A3_Totalgrossamount;FA_A3_Totalgrosproceeds;")
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                    
                    
 'Sch FA-A4
                name = "FA_A4_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                       ' setTblinfo_A_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "FA_A4_Country;FA_A4_BankName;FA_A4_BankAdd;FA_A4_ZipCode;FA_A4_AccOpeningDate;FA_A4_the_cashvalue;FA_A4_Total_grossamount;")
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                
 'Sch FA-B
               name = "FA_B_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_B_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_B_FA)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                
                
'Sch FA-C
                name = "FA_C_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_C_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_C_FA)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                
                
'Sch FA-D
                name = "FA_D_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_D_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_D_FA)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
               
                
'Sch FA-E
                name = "FA_E_InstName"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_E_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_E_FA)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
               
                  
'Sch FA-F
                  
                name = "FA_F_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_F_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_F_FA)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                
                
                 name = "FA_G_Country"
                    sfirstbound = SrcBook.Sheets("TR_FA").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TR_FA").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TR_FA").Range(name).count
                    dcnt = DestBook.Sheets("TR_FA").Range(name).count
                    DestBook.Sheets("TR_FA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TR_FA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_G_FA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_G_FA)
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TR_FA").Range(name).Copy
                        DestBook.Sheets("TR_FA").Range(name).PasteSpecial xlValues
                    End If
                
                
End Function


Function ExtendTableSPI()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

  name = "SPI.SpecifiedPersonName"
    sfirstbound = SrcBook.Sheets("SPI - SI - IF").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("SPI - SI - IF").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    
    cnt = SrcBook.Sheets("SPI - SI - IF").Range(name).count
    dcnt = DestBook.Sheets("SPI - SI - IF").Range(name).count
    DestBook.Sheets("SPI - SI - IF").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("SPI - SI - IF").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
        setTblinfo_SPI
        InsertRowsToImport (cnt - dcnt)
        Call ExendRangeNameToTable(cnt - dcnt, rngname_F_FA)
        SrcBook.Sheets("SPI - SI - IF").Range(name).Copy
        DestBook.Sheets("SPI - SI - IF").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("SPI - SI - IF").Range(name).Copy
        DestBook.Sheets("SPI - SI - IF").Range(name).PasteSpecial xlValues
    End If

End Function


Function ExtendTableSI()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
                
name = "SI.SecCode"
sfirstbound = SrcBook.Sheets("SPI - SI - IF").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("SPI - SI - IF").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("SPI - SI - IF").Range(name).count
dcnt = DestBook.Sheets("SPI - SI - IF").Range(name).count
DestBook.Sheets("SPI - SI - IF").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("SPI - SI - IF").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    setTblinfo_SI
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, rngname_SI)
    SrcBook.Sheets("SPI - SI - IF").Range(name).Copy
    DestBook.Sheets("SPI - SI - IF").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("SPI - SI - IF").Range(name).Copy
    DestBook.Sheets("SPI - SI - IF").Range(name).PasteSpecial xlValues
    
    SrcBook.Sheets("SPI - SI - IF").Range("E7:E12").Copy
    DestBook.Sheets("SPI - SI - IF").Range("E7:E12").PasteSpecial xlValues
    SrcBook.Sheets("SPI - SI - IF").Range("F8:F8").Copy
    DestBook.Sheets("SPI - SI - IF").Range("F8:F8").PasteSpecial xlValues
    
    SrcBook.Sheets("SPI - SI - IF").Range("G8:G8").Copy
    DestBook.Sheets("SPI - SI - IF").Range("G8:G8").PasteSpecial xlValues
    
    SrcBook.Sheets("SPI - SI - IF").Range("C9:C13").Copy
    DestBook.Sheets("SPI - SI - IF").Range("C9:C13").PasteSpecial xlValues
    
    SrcBook.Sheets("SPI - SI - IF").Range("D7:D7").Copy
    DestBook.Sheets("SPI - SI - IF").Range("D7:D7").PasteSpecial xlValues
End If

End Function

 
Function ExtendTableIF()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

  
name = "IF.FirmName"
sfirstbound = SrcBook.Sheets("SPI - SI - IF").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("SPI - SI - IF").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("SPI - SI - IF").Range(name).count
dcnt = DestBook.Sheets("SPI - SI - IF").Range(name).count
DestBook.Sheets("SPI - SI - IF").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("SPI - SI - IF").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    setTblinfo_IF
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, rngname_F_FA)
    SrcBook.Sheets("SPI - SI - IF").Range(name).Copy
    DestBook.Sheets("SPI - SI - IF").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("SPI - SI - IF").Range(name).Copy
    DestBook.Sheets("SPI - SI - IF").Range(name).PasteSpecial xlValues
End If
               
End Function
Function EITableExtend()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "Name_of_district"
    sfirstbound = SrcBook.Sheets("EI").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("EI").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_EIImport As Variant
    rangename_EIImport = "Name_of_district;SchEI_Pincode;SchEI_AgriculturalLand;SchEI_OwedOrLease;SchEI_IrrigatedOrRainFed;"
    cnt = SrcBook.Sheets("EI").Range(name).count
    dcnt = DestBook.Sheets("EI").Range(name).count
    DestBook.Sheets("EI").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("EI").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport ((cnt - dcnt))
        Call ExendRangeNameToTable(cnt - dcnt, rangename_EIImport)
        SrcBook.Sheets("EI").Range(name).Copy
        DestBook.Sheets("EI").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("EI").Range(name).Copy
        DestBook.Sheets("EI").Range(name).PasteSpecial xlValues
    End If
End Function

Function EITableExtend1()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "EI.NatureOfIncome"
    sfirstbound = SrcBook.Sheets("EI").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("EI").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_EIImport1 As Variant
    rangename_EIImport1 = "EI.NatureOfIncome;EI.Description;EI.Amount;"
    cnt = SrcBook.Sheets("EI").Range(name).count
    dcnt = DestBook.Sheets("EI").Range(name).count
    DestBook.Sheets("EI").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("EI").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport ((cnt - dcnt))
        Call ExendRangeNameToTable(cnt - dcnt, rangename_EIImport1)
        SrcBook.Sheets("EI").Range(name).Copy
        DestBook.Sheets("EI").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("EI").Range(name).Copy
        DestBook.Sheets("EI").Range(name).PasteSpecial xlValues
    End If
End Function

Function EITableExtend2()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    name = "SCHEI_DTAA_AmountOfIncome"
    sfirstbound = SrcBook.Sheets("EI").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("EI").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    Dim rangename_EIImport2 As Variant
    rangename_EIImport2 = "SCHEI_DTAA_AmountOfIncome;SCHEI_DTAA_NatureOfIncome;SCHEI_DTAA_CountryCode;SCHEI_DTAA_Article;SCHEI_DTAA_HeadOfIncome;SCHEI_TRCObtained;"
    cnt = SrcBook.Sheets("EI").Range(name).count
    dcnt = DestBook.Sheets("EI").Range(name).count
    DestBook.Sheets("EI").Activate
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("EI").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
       ' setTblinfo_OS
        InsertRowsToImport ((cnt - dcnt))
        Call ExendRangeNameToTable(cnt - dcnt, rangename_EIImport2)
        SrcBook.Sheets("EI").Range(name).Copy
        DestBook.Sheets("EI").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("EI").Range(name).Copy
        DestBook.Sheets("EI").Range(name).PasteSpecial xlValues
    End If
End Function

Function ImportBP1()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

  
name = "Name_1"
sfirstbound = SrcBook.Sheets("BP").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("BP").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("BP").Range(name).count
dcnt = DestBook.Sheets("BP").Range(name).count
DestBook.Sheets("BP").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("BP").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    'setTblinfo_IF
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, "Name_1;Amt_1;")
    SrcBook.Sheets("BP").Range(name).Copy
    DestBook.Sheets("BP").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("BP").Range(name).Copy
    DestBook.Sheets("BP").Range(name).PasteSpecial xlValues
End If
               
End Function

Function ImportGST()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

  
name = "GST_GSTRno"
sfirstbound = SrcBook.Sheets("GST").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("GST").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("GST").Range(name).count
dcnt = DestBook.Sheets("GST").Range(name).count
DestBook.Sheets("GST").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("GST").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    'setTblinfo_IF
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, "GST_GSTRno;GST_GSTRAmount;")
    SrcBook.Sheets("GST").Range(name).Copy
    DestBook.Sheets("GST").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("GST").Range(name).Copy
    DestBook.Sheets("GST").Range(name).PasteSpecial xlValues
End If
               
End Function

Function ImportBP2()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

  
name = "EI.Drpdn"
sfirstbound = SrcBook.Sheets("BP").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("BP").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("BP").Range(name).count
dcnt = DestBook.Sheets("BP").Range(name).count
DestBook.Sheets("BP").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("BP").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    'setTblinfo_IF
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, "EI.Drpdn;")
    SrcBook.Sheets("BP").Range(name).Copy
    DestBook.Sheets("BP").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("BP").Range(name).Copy
    DestBook.Sheets("BP").Range(name).PasteSpecial xlValues
End If
               
End Function

Function ExtendITTable()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

                name = "IT.BSRCode"
                    sfirstbound = SrcBook.Sheets("IT").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("IT").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("IT").Range(name).count
                    dcnt = DestBook.Sheets("IT").Range(name).count
                    DestBook.Sheets("IT").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("IT").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTableInfo_IT
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_IT)
                        SrcBook.Sheets("IT").Range(name).Copy
                        DestBook.Sheets("IT").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("IT").Range(name).Copy
                        DestBook.Sheets("IT").Range(name).PasteSpecial xlValues
                    End If
               
 
End Function


Function ExtendTDSTable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim vRows As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

                name = "TDS1.TAN"
                    sfirstbound = SrcBook.Sheets("TDS").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(name).count
                    dcnt = DestBook.Sheets("TDS").Range(name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTableinfo_TDS1import
                        EfilingCommon.DefinedgridNameRange = rngname_TDS1
                        ActiveCellRange = EfilingCommon.searchLastRow("TDS1.TAN")
                        vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
                        SrcBook.Sheets("TDS").Range(name).Copy
                        DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(name).Copy
                        DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
                    End If
                
'Sch TDS2
        
               name = "TDS2.TDSCreditName"
                    sfirstbound = SrcBook.Sheets("TDS").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(name).count
                    dcnt = DestBook.Sheets("TDS").Range(name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTableinfo_TDS2
                        EfilingCommon.DefinedgridNameRange = "TDS2.TDSCreditName;TDS2.GrossAmount;TDS2.HeadOfIncome;TDS2.PanOtherPerson;TDSAadharOfOther_TDS2;TDS2.TAN;TDS2.DeductedYear;TDS2.BroughtFwdTDSAmt;TDS2.DeductedInOwnHands;TDS2.DeductedIncome5A;TDS2.DeductedTDS5A;TDS2.ClaimedInOwnHands;TDS2.ClaimedIncome5A;TDS2.ClaimedTDS5A;TDS2.ClaimedPAN5A;ClaimedAadhar5A_TDS2;TDS2.AmtCarriedForward;"
                        ActiveCellRange = EfilingCommon.searchLastRow("TDS2.TDSCreditName")
                        vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
                        SrcBook.Sheets("TDS").Range(name).Copy
                        DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(name).Copy
                        DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
                    End If
                

'Sch TDS3
                name = "TDS3.TDSCreditName"
                    sfirstbound = SrcBook.Sheets("TDS").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(name).count
                    dcnt = DestBook.Sheets("TDS").Range(name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTableinfo_TDS3
                        EfilingCommon.DefinedgridNameRange = "TDS3.TDSCreditName;TDS3.GrossAmount;TDS3.HeadOfIncome;TDS3.PanOtherPerson;Aadhar_TDS3;TDS3.PAN;TDS3.Aadhaar;TDS3.DeductedYear;TDS3.BroughtFwdTDSAmt;TDS3.DeductedInOwnHands;TDS3.DeductedIncome5A;TDS3.DeductedTDS5A;TDS3.ClaimedInOwnHands;TDS3.ClaimedIncome5A;TDS3.ClaimedTDS5A;TDS3.ClaimedPAN5A;ClaimedAadhar5A_TDS3;TDS3.AmtCarriedForward;"
                        ActiveCellRange = EfilingCommon.searchLastRow("TDS3.TDSCreditName")
                        vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
                        SrcBook.Sheets("TDS").Range(name).Copy
                        DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(name).Copy
                        DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
                    End If
               
  
End Function


Function ExtendTableFT()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
    
     name = "FT.Country"
                    sfirstbound = SrcBook.Sheets("FT").Range(name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    Dim rangename_FT As Variant
                    rangename_FT = "FT.Country;FT.NoOfTimes;FT.Expenses;"
                    dfirstbound = DestBook.Sheets("FT").Range(name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("FT").Range(name).count
                    dcnt = DestBook.Sheets("FT").Range(name).count
                    DestBook.Sheets("FT").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("FT").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rangename_FT)
                        SrcBook.Sheets("FT").Range(name).Copy
                        DestBook.Sheets("FT").Range(name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("FT").Range(name).Copy
                        DestBook.Sheets("FT").Range(name).PasteSpecial xlValues
                    End If
End Function

Function ExtendTableTCS()
Dim vRows As Variant
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "TCS.TAN"
sfirstbound = SrcBook.Sheets("TDS").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("TDS").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("TDS").Range(name).count
dcnt = DestBook.Sheets("TDS").Range(name).count
DestBook.Sheets("TDS").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    setTblinfo_TCS2XL
    EfilingCommon.DefinedgridNameRange = rngname_TCS
    ActiveCellRange = EfilingCommon.searchLastRow("TCS.TAN")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
    SrcBook.Sheets("TDS").Range(name).Copy
    DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("TDS").Range(name).Copy
    DestBook.Sheets("TDS").Range(name).PasteSpecial xlValues
End If

    
End Function


Function ExtendTable_112A()
Dim vRows As Variant
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "ISINCode_112A"
sfirstbound = SrcBook.Sheets("Schedule 112A").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("Schedule 112A").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("Schedule 112A").Range(name).count
dcnt = DestBook.Sheets("Schedule 112A").Range(name).count
DestBook.Sheets("Schedule 112A").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("Schedule 112A").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    EfilingCommon.DefinedgridNameRange = "ShareAcq_112A;ISINCode_112A;NameofShare_112A;NoofShare_112A;SalePriceperShare_112A;TotalSaleValue_112A;COAwithoutIndex_112A;COAwithIndex_112A;LTCGAssetAcquired_112A;FairMarketValue_112A;TotalFairMarketValue_112A;ExpenditureWholly_112A;TotalDeductions_112A;Balance_112A;"
    ActiveCellRange = EfilingCommon.searchLastRow("ISINCode_112A")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
    SrcBook.Sheets("Schedule 112A").Range(name).Copy
    DestBook.Sheets("Schedule 112A").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("Schedule 112A").Range(name).Copy
    DestBook.Sheets("Schedule 112A").Range(name).PasteSpecial xlValues
End If

    
End Function

Function ExtendTable_Salary()
Dim vRows As Variant
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "Salary.Nature1"
sfirstbound = SrcBook.Sheets("Schedule S").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("Schedule S").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("Schedule S").Range(name).count
dcnt = DestBook.Sheets("Schedule S").Range(name).count
DestBook.Sheets("Schedule S").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("Schedule S").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    'setTblinfo_TCS2XL
    EfilingCommon.DefinedgridNameRange = "Salary.Nature1;Salary.Description1;Salary.Amount1;"
    ActiveCellRange = EfilingCommon.searchLastRow("Salary.Nature1")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
    SrcBook.Sheets("Schedule S").Range(name).Copy
    DestBook.Sheets("Schedule S").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("Schedule S").Range(name).Copy
    DestBook.Sheets("Schedule S").Range(name).PasteSpecial xlValues
End If

    
End Function

Function ExtendTable_115AD()
Dim vRows As Variant
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "ISINCode_115AD"
sfirstbound = SrcBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).count
dcnt = DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).count
DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    'setTblinfo_TCS2FullValueConsideration_115AD
    EfilingCommon.DefinedgridNameRange = "ShareAcq_115AD_1;ISINCode_115AD;NameofShare_115AD;NoofShare_115AD;SalePriceperShare_115AD;TotalSaleValue_115AD;COAwithoutIndex_115AD;COAwithIndex_115AD;LTCGAssetAcquired_115AD;FairMarketValue_115AD;TotalFairMarketValue_115AD;ExpenditureWholly_115AD;TotalDeductions_115AD;Balance_115AD;"
    ActiveCellRange = EfilingCommon.searchLastRow("ISINCode_115AD")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula1(cnt - dcnt)
    SrcBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).Copy
    DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).Copy
    DestBook.Sheets("Schedule 115AD(1)(iii) proviso").Range(name).PasteSpecial xlValues
End If

    
End Function
Function ExtendHPTable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant

cnt = SrcBook.Sheets("House Property").Range("PropertySectionCOunt").value
dcnt = DestBook.Sheets("House Property").Range("PropertySectionCOunt").value
DestBook.Sheets("House Property").Activate
If cnt > dcnt Then
    Dim loopcount As Variant
    For loopcount = 1 To cnt - dcnt
        AddBlockCall_hprptfrm
    Next
End If
End Function

Function ExtendPTITable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
Dim loopcount As Variant

cnt = SrcBook.Sheets("PTI").Range("cntr.Ptirptfrm").value
dcnt = DestBook.Sheets("PTI").Range("cntr.Ptirptfrm").value
DestBook.Sheets("PTI").Activate

If cnt > dcnt Then
For loopcount = 1 To cnt - dcnt
    AddRows_PTI
Next
End If
                           
End Function

Function ExtendAL1Table()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "SchAL.A.Description"
sfirstbound = SrcBook.Sheets("AL").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("AL").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("AL").Range(name).count
dcnt = DestBook.Sheets("AL").Range(name).count
DestBook.Sheets("AL").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    setTableInfo_AL1
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, range_AL1)
    SrcBook.Sheets("AL").Range(name).Copy
    DestBook.Sheets("AL").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("AL").Range(name).Copy
    DestBook.Sheets("AL").Range(name).PasteSpecial xlValues
End If
End Function

Function ExtendAL2Table()

Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "SchAL.C.Name"
sfirstbound = SrcBook.Sheets("AL").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("AL").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("AL").Range(name).count
dcnt = DestBook.Sheets("AL").Range(name).count
DestBook.Sheets("AL").Activate
If (cnt - dcnt) > 0 Then
    DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    setTableInfo_AL2_1
    InsertRowsToImport (cnt - dcnt)
    Call ExendRangeNameToTable(cnt - dcnt, range_AL2_1)
    SrcBook.Sheets("AL").Range(name).Copy
    DestBook.Sheets("AL").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("AL").Range(name).Copy
    DestBook.Sheets("AL").Range(name).PasteSpecial xlValues
End If
End Function

Function ExtendFSITable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
Dim loopcount As Variant

cnt = SrcBook.Sheets("FSI").Range("FSI_Count").value
dcnt = DestBook.Sheets("FSI").Range("FSI_Count").value
DestBook.Sheets("FSI").Activate

If cnt > dcnt Then
For loopcount = 1 To cnt - dcnt
    AddBlockCall_FSIfrm
Next
End If
                           
End Function

Function ExtendSalTable()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant
Dim loopcount As Variant
cnt = SrcBook.Sheets("Schedule S").Range("SalarySectionCount").value
dcnt = DestBook.Sheets("Schedule S").Range("SalarySectionCount").value
DestBook.Sheets("Schedule S").Activate
sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword

If cnt > dcnt Then
For loopcount = 1 To cnt - dcnt
    addSalariesBlock
Next
Dim te As Long
Dim numberofrows As Variant
numberofrows = 0
For te = 1 To Sheet39.Range("SalarySectionCount").value
    cnt = SrcBook.Sheets("Schedule S").Range("Salary.Section" & te).count
dcnt = DestBook.Sheets("Schedule S").Range("Salary.Section" & te).count
        
        EfilingCommon.DefinedgridNameRange = "Salary.Section||Salary.salaryDescription||Salary.SalaryAmount"
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.Section" & te)
        
        If (cnt - dcnt) > 0 Then
           numberofrows = EfilingCommon.insertRowUnderSectionWithFormula((cnt - dcnt), True, te)
        End If
Next
Application.EnableEvents = False
Application.EnableEvents = False
'Sheet39.Range("SalarySectionItem").value = Sheet39.Range("SalarySectionItem").value + numberofrows
Dim numberofrows1 As Variant
numberofrows1 = 0
For te = 1 To Sheet39.Range("SalarySectionCount").value
    cnt = SrcBook.Sheets("Schedule S").Range("Salary.ValueSection" & te).count
dcnt = DestBook.Sheets("Schedule S").Range("Salary.ValueSection" & te).count
        
        EfilingCommon.DefinedgridNameRange = "Salary.ValueSection||Salary.ValueDescription||Salary.ValueAmount"
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.ValueSection" & te)
        
        If (cnt - dcnt) > 0 Then
           numberofrows1 = EfilingCommon.insertRowUnderSectionWithFormula((cnt - dcnt), True, te)
        End If
Next
'Sheet39.Range("SalarySectionItem").value = Sheet39.Range("SalarySectionItem").value + numberofrows
Dim numberofrows2 As Variant
numberofrows2 = 0
For te = 1 To Sheet39.Range("SalarySectionCount").value
    cnt = SrcBook.Sheets("Schedule S").Range("Salary.ProfitSection" & te).count
dcnt = DestBook.Sheets("Schedule S").Range("Salary.ProfitSection" & te).count
        
        EfilingCommon.DefinedgridNameRange = "Salary.ProfitSection||Salary.ProfitDescription||Salary.ProfitAmount"
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.ProfitSection" & te)
        
        If (cnt - dcnt) > 0 Then
           numberofrows2 = EfilingCommon.insertRowUnderSectionWithFormula((cnt - dcnt), True, te)
        End If
Next

Application.EnableEvents = False
Application.EnableEvents = False


sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
         
Sheet39.Range("SalarySectionItem").value = Sheet39.Range("SalarySectionItem").value + numberofrows + numberofrows1 + numberofrows2

sPassword = EfilingCommon.getmsgstate
         Sheet39.Protect Password:=sPassword

'----------------Unlock Password-------------------START---

End If

End Function

Function ExtendBankAccount()
Dim cnt, dcnt As Long
Dim newrname As String
Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
Dim name As Variant


name = "Sheet9.OthMICRCode"
sfirstbound = SrcBook.Sheets("Part B - TI TTI").Range(name).Address
sTEMP = Split(sfirstbound, "$")
supperbound = UBound(sTEMP)
sTEMP = sTEMP(UBound(sTEMP))

dfirstbound = DestBook.Sheets("Part B - TI TTI").Range(name).Address
dTemp = Split(dfirstbound, "$")
dupperbound = UBound(dTemp)
ddTemp = dTemp(UBound(dTemp))

cnt = SrcBook.Sheets("Part B - TI TTI").Range(name).count
dcnt = DestBook.Sheets("Part B - TI TTI").Range(name).count
DestBook.Sheets("Part B - TI TTI").Activate
Application.EnableEvents = False

If (cnt - dcnt) > 0 Then
    DestBook.Sheets("Part B - TI TTI").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
    'setTableInfo_BA_IFSC
    InsertRowsToImport (cnt - dcnt)
    'Call ExendRangeNameToTable(cnt - dcnt, "Sheet9.OthMICRCode;Sheet9.OthBankName;Sheet9.OthBankAccountNumber;Sheet9.CheckBox;tempxml;")
    Call ExendRangeNameToTable(cnt - dcnt, "Sheet9.OthMICRCode;Sheet9.OthBankName;Sheet9.OthBankAccountNumber;Sheet9.TypeOfAccount;")

    
    SrcBook.Sheets("Part B - TI TTI").Range(name).Copy
    DestBook.Sheets("Part B - TI TTI").Range(name).PasteSpecial xlValues
Else
    SrcBook.Sheets("Part B - TI TTI").Range(name).Copy
    DestBook.Sheets("Part B - TI TTI").Range(name).PasteSpecial xlValues
End If
Application.EnableEvents = True

End Function


Function ExtendTableBA2()
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    Dim name As Variant
    
    
    name = "Sheet9.IBAN"
    sfirstbound = SrcBook.Sheets("Part B - TI TTI").Range(name).Address
    sTEMP = Split(sfirstbound, "$")
    supperbound = UBound(sTEMP)
    sTEMP = sTEMP(UBound(sTEMP))
    
    dfirstbound = DestBook.Sheets("Part B - TI TTI").Range(name).Address
    dTemp = Split(dfirstbound, "$")
    dupperbound = UBound(dTemp)
    ddTemp = dTemp(UBound(dTemp))
    
    cnt = SrcBook.Sheets("Part B - TI TTI").Range(name).count
    dcnt = DestBook.Sheets("Part B - TI TTI").Range(name).count
    DestBook.Sheets("Part B - TI TTI").Activate
    Application.EnableEvents = False
    
    If (cnt - dcnt) > 0 Then
        DestBook.Sheets("Part B - TI TTI").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
        setTblinfo_IBAN_1
        InsertRowsToImport (cnt - dcnt)
        Call ExendRangeNameToTable(cnt - dcnt, rngname_IBAN)
        SrcBook.Sheets("Part B - TI TTI").Range(name).Copy
        DestBook.Sheets("Part B - TI TTI").Range(name).PasteSpecial xlValues
    Else
        SrcBook.Sheets("Part B - TI TTI").Range(name).Copy
        DestBook.Sheets("Part B - TI TTI").Range(name).PasteSpecial xlValues
    End If
    Application.EnableEvents = True
End Function

Function LockUnlockTables()
On Error GoTo endline

Dim Rangename As Variant
Dim i As Variant
DestBook.Activate

Sheet13.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Rangename = "STCG.A7_DeductionSection_1||STCG.A7_AsstAcquiredYer_1||STCG.A7_AmtUtilised_1||STCG.A7_AmtNotUsed_1||STCG.A7_AsstAcquiredYer_2||STCG.A7_AmtUtilised_2||STCG.A7_AmtNotUsed_2"
Rangename = Split(Rangename, "||")
    If Mid(Sheet13.Range("STCG.A6_StatusUnutilizedCG").value, 1, 1) <> "Y" Then
        
        For i = 0 To UBound(Rangename)
            Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
            Sheet13.Range(Rangename(i)).Locked = True
        Next
    Else
        For i = 0 To UBound(Rangename)
            'Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
            Sheet13.Range(Rangename(i)).Locked = False
        Next
    End If
Sheet13.Protect Password:=getmsgstate
Application.EnableEvents = True





Sheet13.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Rangename = "LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1||LTCG.B9_AsstAcquiredYer_2||LTCG.B9_AmtUtilised_2||LTCG.B9_AmtNotUsed_2"
Rangename = Split(Rangename, "||")
    If Mid(Sheet13.Range("LTCG.B9_DeductionSection_1").value, 1, 1) = "N" Then
       
        For i = 0 To UBound(Rangename)
            Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
            Sheet13.Range(Rangename(i)).Locked = True
        Next
    Else
        For i = 0 To UBound(Rangename)
            'Sheet13.Range(Rangename(i)).value = ""
            Sheet13.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
            Sheet13.Range(Rangename(i)).Locked = False
        Next
    End If
Sheet13.Protect Password:=getmsgstate
Application.EnableEvents = True

If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
    Sheet13.Unprotect Password:=getmsgstate
    Sheet14.Unprotect Password:=getmsgstate
    
    Sheet13.Range("STCG.A.NRI").EntireRow.Hidden = True
    Sheet13.Range("LTCG.B.NRI").EntireRow.Hidden = True
    Sheet13.Shapes("Button 89").Visible = msoFalse
    Sheet13.Shapes("Button 160").Visible = msoFalse
    Sheet13.Range("STCG.DTAA").EntireRow.Hidden = True
    Sheet13.Range("LTCG.DTAA").EntireRow.Hidden = True
    
    
    Sheet13.Range("LTCG.B4d_LTCGOnListedSecurity").EntireRow.Hidden = True
    
   ' Sheet13.Range("LTCG.B4bi_ExemptionAmount").EntireRow.Hidden = True
   ' Sheet13.Range("LTCG.B4bii_ExemptionAmount").EntireRow.Hidden = True
'    Sheet14.Range("OS_DTAA_hiding").EntireRow.Hidden = True
    Sheet14.Shapes("Button 40").Visible = msoFalse
    
    Sheet13.Range("LTCG.B4c_BalanceCG").EntireRow.Hidden = True
    
    Sheet13.Protect Password:=getmsgstate
    Sheet14.Protect Password:=getmsgstate

Else
    Sheet13.Unprotect Password:=getmsgstate
    Sheet14.Unprotect Password:=getmsgstate
    
    Sheet13.Range("STCG.A.NRI").EntireRow.Hidden = False
    Sheet13.Range("LTCG.B.NRI").EntireRow.Hidden = False
    Sheet13.Range("STCG.DTAA").EntireRow.Hidden = False
    Sheet13.Range("LTCG.DTAA").EntireRow.Hidden = False
    Sheet13.Shapes("Button 89").Visible = msoTrue
    Sheet13.Shapes("Button 160").Visible = msoTrue
    Sheet13.Range("LTCG.B4d_LTCGOnListedSecurity").EntireRow.Hidden = True
    
    
   ' Sheet13.Range("LTCG.B4bi_ExemptionAmount").EntireRow.Hidden = True
    Sheet13.Range("LTCG.B4bii_ExemptionAmount").EntireRow.Hidden = True
    Sheet13.Range("LTCG.B4c_BalanceCG").EntireRow.Hidden = True
    
'    Sheet14.Range("OS_DTAA_hiding").EntireRow.Hidden = False
    Sheet14.Shapes("Button 40").Visible = msoTrue
    
    Sheet13.Protect Password:=getmsgstate
    Sheet14.Protect Password:=getmsgstate
End If

endline:
    Application.EnableEvents = True
End Function

Function getSheetName(rn As Variant) As String
Dim temp As Variant
temp = Replace(Mid(rn, 2, InStr(1, rn, "!") - 2), "'", "")
getSheetName = temp
End Function

Sub TemplockPL(targetadd)


If Not Application.Intersect(Range("PL_Address_State"), Sheet4.Range(targetadd)) Is Nothing Then
            Sheet4.Unprotect Password:=EfilingCommon.getmsgstate
If Not isdropdownblank(Sheet4.Range(targetadd).value) Then
            If Mid(Sheet4.Range(targetadd).value, 1, 2) = "99" Then

                Sheet4.Range(Replace(targetadd, "M", "O")).Interior.Color = (&HD8D8D8)
                Sheet4.Range(Replace(targetadd, "M", "O")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "O")).Locked = True

                Sheet4.Range(Replace(targetadd, "M", "P")).Interior.Color = (&HCCFFCC)
                'Sheet4.Range(Replace(Targetadd, "M", "P")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "P")).Locked = False

             ElseIf Mid(Sheet4.Range(targetadd).value, 1, 2) <> "99" Then

                  Sheet4.Range(Replace(targetadd, "M", "O")).Interior.Color = (&HCCFFCC)
                'Sheet4.Range(Replace(Targetadd, "M", "O")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "O")).Locked = False

                Sheet4.Range(Replace(targetadd, "M", "P")).Interior.Color = (&HD8D8D8)
                Sheet4.Range(Replace(targetadd, "M", "P")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "P")).Locked = True

                Else
                
                  Sheet4.Range(Replace(targetadd, "M", "J")).Interior.Color = (&HCCFFCC)
                'Sheet4.Range(Replace(Targetadd, "M", "J")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "J")).Locked = False

                Sheet4.Range(Replace(targetadd, "M", "K")).Interior.Color = (&HCCFFCC)
                'Sheet4.Range(Replace(Targetadd, "M", "K")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "K")).Locked = False

                Sheet4.Range(Replace(targetadd, "M", "L")).Interior.Color = (&HCCFFCC)
                'Sheet4.Range(Replace(Targetadd, "M", "L")).value = ""
                Sheet4.Range(Replace(targetadd, "M", "L")).Locked = False
                
End If

          
 End If
 Sheet4.Protect Password:=EfilingCommon.getmsgstate
End If

End Sub
Sub Templock2(targetadd)

If Not Application.Intersect(Range("TDS2.TDSCreditName"), Sheet28.Range(targetadd)) Is Nothing Then
            Sheet28.Unprotect Password:=EfilingCommon.getmsgstate

       If UCase(Mid(Sheet28.Range(targetadd).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(targetadd, "E", "L")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "L")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "L")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "M")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "M")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "M")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "F")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "F")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(targetadd, "E", "G")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "G")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "O")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "O")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "O")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "P")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "P")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "Q")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "Q")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "Q")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(targetadd, "E", "R")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "R")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "R")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "K")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "K")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "K")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(targetadd).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(targetadd, "E", "L")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "L")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "L")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "M")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "M")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "F")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "F")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(targetadd, "E", "G")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "G")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "O")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "O")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "O")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "P")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "P")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "P")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "Q")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "Q")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "Q")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(targetadd, "E", "R")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "R")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "R")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "K")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "K")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "K")).Interior.Color = "&HD8D8D8"

    Else
        Sheet28.Range(Replace(targetadd, "E", "L")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "L")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "L")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "M")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "M")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "F")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "F")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(targetadd, "E", "G")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "G")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "O")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "O")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "O")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "P")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "P")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "Q")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "Q")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "Q")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(targetadd, "E", "R")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "R")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "R")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "K")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "K")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "K")).Interior.Color = "&HCCFFCC"

    End If
 Sheet28.Protect Password:=EfilingCommon.getmsgstate
End If
If Not Application.Intersect(Range("TDS2.HeadOfIncome"), Sheet28.Range(targetadd)) Is Nothing Then
 Sheet28.Unprotect Password:=EfilingCommon.getmsgstate
    If UCase(Mid(Sheet28.Range(targetadd).value, 1, 1)) = "N" Then
        Sheet28.Range(Replace(targetadd, "T", "S")).ClearContents
        Sheet28.Range(Replace(targetadd, "T", "S")).Locked = True
        Sheet28.Range(Replace(targetadd, "T", "S")).Interior.Color = "&HD8D8D8"
     Else
        Sheet28.Range(Replace(targetadd, "T", "S")).ClearContents
        Sheet28.Range(Replace(targetadd, "T", "S")).Locked = False
        Sheet28.Range(Replace(targetadd, "T", "S")).Interior.Color = "&HCCFFCC"
    End If
  Sheet28.Protect Password:=EfilingCommon.getmsgstate
End If

If Not Application.Intersect(Range("TDS2.ClaimedInOwnHands"), Sheet28.Range(targetadd)) Is Nothing Then

        If Sheet28.Range(targetadd).value > 0 Then
            Sheet28.Range(Replace(targetadd, "N", "S")).Locked = False
            Sheet28.Range(Replace(targetadd, "N", "S")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(targetadd, "N", "T")).Locked = False
            Sheet28.Range(Replace(targetadd, "N", "T")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(targetadd, "N", "T")).value = "(Select)"
        Else
            Sheet28.Range(Replace(targetadd, "N", "S")).ClearContents
            Sheet28.Range(Replace(targetadd, "N", "S")).Locked = True
            Sheet28.Range(Replace(targetadd, "N", "S")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(targetadd, "N", "T")).ClearContents
            Sheet28.Range(Replace(targetadd, "N", "T")).Locked = True
            Sheet28.Range(Replace(targetadd, "N", "T")).Interior.Color = "&HD8D8D8"


    End If
End If



End Sub

Sub Templock3(targetadd)


If Not Application.Intersect(Range("TDS3.TDSCreditName"), Sheet28.Range(targetadd)) Is Nothing Then
Sheet28.Unprotect Password:=EfilingCommon.getmsgstate
    If UCase(Mid(Sheet28.Range(targetadd).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(targetadd, "E", "M")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "M")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "M")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "N")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "N")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "N")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "F")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "F")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(targetadd, "E", "G")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "G")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "P")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "P")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "Q")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "Q")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "R")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "R")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "R")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(targetadd, "E", "S")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "S")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "S")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "L")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "L")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "L")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(targetadd).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(targetadd, "E", "M")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "M")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "N")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "N")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "F")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "F")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(targetadd, "E", "G")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "G")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "P")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "P")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "P")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "Q")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "Q")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "Q")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "R")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "R")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "R")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(targetadd, "E", "S")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "S")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "S")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "L")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "L")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "L")).Interior.Color = "&HD8D8D8"

    Else
        Sheet28.Range(Replace(targetadd, "E", "M")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "M")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "N")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "N")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "F")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "F")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(targetadd, "E", "G")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "G")).Locked = True
        Sheet28.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(targetadd, "E", "P")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "P")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "Q")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "Q")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "R")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "R")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "R")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(targetadd, "E", "S")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "S")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "S")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(targetadd, "E", "L")).ClearContents
        Sheet28.Range(Replace(targetadd, "E", "L")).Locked = False
        Sheet28.Range(Replace(targetadd, "E", "L")).Interior.Color = "&HCCFFCC"

    End If
Sheet28.Protect Password:=EfilingCommon.getmsgstate

End If




If Not Application.Intersect(Range("TDS3.ClaimedInOwnHands"), Sheet28.Range(targetadd)) Is Nothing Then
    Sheet28.Unprotect Password:=EfilingCommon.getmsgstate
        If Sheet28.Range(targetadd).value > 0 Then
            Sheet28.Range(Replace(targetadd, "O", "T")).Locked = False
            Sheet28.Range(Replace(targetadd, "O", "T")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(targetadd, "O", "U")).Locked = False
            Sheet28.Range(Replace(targetadd, "O", "U")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(targetadd, "O", "U")).value = "(Select)"
        Else
            Sheet28.Range(Replace(targetadd, "O", "T")).ClearContents
            Sheet28.Range(Replace(targetadd, "O", "T")).Locked = True
            Sheet28.Range(Replace(targetadd, "O", "T")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(targetadd, "O", "U")).ClearContents
            Sheet28.Range(Replace(targetadd, "O", "U")).Locked = True
            Sheet28.Range(Replace(targetadd, "O", "U")).Interior.Color = "&HD8D8D8"
    End If
Sheet28.Protect Password:=EfilingCommon.getmsgstate
End If

End Sub




Sub returnfilesecimport()
    DestBook.Sheets("PART A - General").Range("sheet1.ReturnFileSec").value = SrcBook.Sheets("PART A - General").Range("sheet1.ReturnFileSec").value


Sheet1.Unprotect Password:=EfilingCommon.getmsgstate

     Range("sheet1.ReturnFileSec1").formula = "=VLOOKUP(sheet1.ReturnFileSec,ReturnFileUnderSection,2,FALSE)"

Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub


Sub Templock4(targetadd)
If Not Intersect(Range("TDS2.ClaimedInOwnHands"), Range(targetadd)) Is Nothing Then
Sheet28.Unprotect Password:=EfilingCommon.getmsgstate
        If Sheet28.Range(targetadd).value > 0 Then
            Sheet28.Range(Replace(targetadd, "M", "Q")).Locked = False
            Sheet28.Range(Replace(targetadd, "M", "Q")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(targetadd, "M", "R")).Locked = False
            Sheet28.Range(Replace(targetadd, "M", "R")).Interior.Color = "&HCCFFCC"
'            Sheet28.Range(Replace(Targetadd, "M", "R")).value = "(Select)"
        Else
            Sheet28.Range(Replace(targetadd, "M", "Q")).ClearContents
            Sheet28.Range(Replace(targetadd, "M", "Q")).Locked = True
            Sheet28.Range(Replace(targetadd, "M", "Q")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(targetadd, "M", "R")).ClearContents
            Sheet28.Range(Replace(targetadd, "M", "R")).Locked = True
            Sheet28.Range(Replace(targetadd, "M", "R")).Interior.Color = "&HD8D8D8"


    End If
  Sheet28.Protect Password:=EfilingCommon.getmsgstate
End If
End Sub
Sub Templock5(targetadd)
If Not Intersect(Range("TDS3.ClaimedInOwnHands"), Range(targetadd)) Is Nothing Then
    Sheet28.Unprotect Password:=EfilingCommon.getmsgstate
        If Sheet28.Range(targetadd).value > 0 Then
            Sheet28.Range(Replace(targetadd, "M", "Q")).Locked = False
            Sheet28.Range(Replace(targetadd, "M", "Q")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(targetadd, "M", "R")).Locked = False
            Sheet28.Range(Replace(targetadd, "M", "R")).Interior.Color = "&HCCFFCC"
'            Sheet28.Range(Replace(Targetadd, "M", "R")).value = "(Select)"
        Else
            Sheet28.Range(Replace(targetadd, "M", "Q")).ClearContents
            Sheet28.Range(Replace(targetadd, "M", "Q")).Locked = True
            Sheet28.Range(Replace(targetadd, "M", "Q")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(targetadd, "M", "R")).ClearContents
            Sheet28.Range(Replace(targetadd, "M", "R")).Locked = True
            Sheet28.Range(Replace(targetadd, "M", "R")).Interior.Color = "&HD8D8D8"

    End If
    Sheet28.Protect Password:=EfilingCommon.getmsgstate
End If
End Sub


Sub salary_Lock()
Dim count, i, j As Long
count = Sheet39.Range("SalarySectionCount").value
Dim chcell As Range

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
For i = 1 To count

    For Each chcell In Sheet39.Range("Salary.Section" & i)
             If chcell.value = "Others" Then
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = False
                     DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HCCFFCC"
                     DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
                    
            Else
                     DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = True
                     DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HD8D8D8"
                     DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
            
             End If
    Next
    
        For Each chcell In Sheet39.Range("Salary.ValueSection" & i)
             If chcell.value = "Other benefits or amenities" Then
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = False
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HCCFFCC"
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
                    
            Else
              DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = True
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HD8D8D8"
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
            
             End If
    Next
    
        For Each chcell In Sheet39.Range("Salary.ProfitSection" & i)
             If chcell.value = "Any Other" Then
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = False
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HCCFFCC"
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
                    
            Else
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = True
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HD8D8D8"
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
            
             End If
       Next
       
          
Next
  
  
          For Each chcell In Sheet39.Range("Salary.Nature1")
             If chcell.value = "Any Other" Then
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = False
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HCCFFCC"
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
                    
            Else
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Locked = True
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).Interior.Color = "&HD8D8D8"
                    DestBook.Worksheets("Schedule S").Range(Replace(chcell.Address, "G", "H")).value = ""
            
             End If
         Next
        
         Sheet39.Protect Password:=sPassword
         
        SrcBook.Sheets("Schedule S").Unprotect Password:=getmsgstate
        DestBook.Sheets("Schedule S").Unprotect Password:=getmsgstate
    For j = 1 To count
        
        Dim rangenamestring As Variant
        rangenamestring = "Salary.salaryDescription||Salary.ValueDescription||Salary.ProfitDescription"
        rangenamestring = Split(rangenamestring, "||")
       
       For i = 0 To UBound(rangenamestring)
        SrcBook.Worksheets("Schedule S").Range(rangenamestring(i) & j).Copy
        DestBook.Worksheets("Schedule S").Range(rangenamestring(i) & j).PasteSpecial xlValues
        Next
        
    Next
    
        SrcBook.Worksheets("Schedule S").Range("Salary.Description1").Copy
        DestBook.Worksheets("Schedule S").Range("Salary.Description1").PasteSpecial xlValues
    
        SrcBook.Sheets("Schedule S").Protect Password:=getmsgstate
        DestBook.Sheets("Schedule S").Protect Password:=getmsgstate
         
End Sub




