Attribute VB_Name = "mdImportXL"
Option Explicit


Dim end_Per10080G, end_PerNO5080G, end_PerYES10080G, end_Per5080G, end_ITimport, end_TCSimport, end_TDS1import, end_TDS2import, end_TDS2iiimport, end_PerSchBA, end_CTAImport, end_CTBImport, end_CTCImport, end_CTDImport, end_CTEImport As Variant
Dim rngname_Per10080G, rngname_PerNO5080G, rngname_PerYES10080G, rngname_Per5080G, rngname_ITimport, rngname_TCSimport, rngname_TD1import, rngname_TDS2iiimport, rngname_TDS2import, rngname_PerSchBA, rngname_TDS1import, rngname_CTAImport, rngname_CTBImport, rngname_CTCImport, rngname_CTDImport, rngname_CTEImport As Variant
Dim end_TDS2subimport, end_TDS2bimport, end_TDS3aimport, end_TDS3subimport, end_TDS3bimport As Variant
Dim rngname_TDS2subimport, rngname_TDS2bimport, rngname_TDS3bimport, rngname_TDS3aimport, rngname_TDS3subimport As Variant
Public end_OtherEI, rngname_OtherEI As Variant
Public end_Allowances, rngname_Allowances As Variant
Public end_OS1, rngname_OS1 As Variant
Public retrievedSheetName As Variant
Public DestBook As Workbook, SrcBook As Workbook
Sub IMPPrevVersion()
On Error Resume Next
    
    Dim rangenamearr() As Variant
    Dim Filename As Variant
    Dim dfilename, ndfilename, newfilename As Variant
    Dim flag As Boolean
    Dim add As Variant
    Dim rname, ws As Variant
    Dim a As Long
    Dim destadd As Variant
    
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    
    InitProgBar
    ProgressFrameCaption = "Importing Excel"
    mainProcCaption = "Importing"
    noOfProcessMain = 2
    
    flag = True
    fmsgboxStatus "Please use this only to IMPORT from previous version of current assessment year only."
    fmsgboxStatus "Please check the full and final ITR4 after importing to ensure all rows are imported. "
    
    Filename = cmdFileDialog()
    If Filename <> "" Then UserForm1.Show vbModeless
    If Not Filename = "" Then
        Filename = Split(Filename, "\")
        newfilename = Filename(UBound(Filename))
        cnt = 0
        Application.ScreenUpdating = False
        Set SrcBook = Workbooks.Open(newfilename)
        Err.Clear
        If Err = "1004" Then ProgressBarHide
        Set DestBook = ThisWorkbook
        
        Dim targetSheet As Worksheet
Set targetSheet = DestBook.Worksheets(1)
Dim sourceSheet As Worksheet
Set sourceSheet = SrcBook.Worksheets(1)



     If (targetSheet.Range("AN1").value) <> (sourceSheet.Range("AN1").value) Then
        fmsgboxStatus "* Invalid ITR or Assessment Year"
        'SrcBook.Save
        SrcBook.Close
          Exit Sub
          
        Else
     
       End If
        
        
        dfilename = Split((DestBook.FullName), "\")
        ndfilename = dfilename(UBound(dfilename))
        noOfProcessSub = 10
        UpdateProgressBar
        UpdateProgressBar
        If newfilename <> ndfilename Then
            For Each rname In Workbooks(newfilename).Names
                newrname = rname.name
                newrname = Mid(newrname, InStr(1, newrname, "!") + 1)
                
                Application.EnableEvents = False
                Application.ScreenUpdating = False

'Sch 80G- Section A
                If newrname = "Per10080G.DoneeName" Then
                    subProcCaption = "Importing 80G(A)"
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
        
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
        
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        
                        setTblinfo_Per10080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_Per10080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
        
'Sch 80G- Section B
                If newrname = "PerNO5080G.DoneeName" Then
                    subProcCaption = "Importing 80G(B)"
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_PerNO5080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_PerNO5080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
        
        
'Sch 80G- Section C

                If newrname = "PerYES10080G.DoneeName" Then
                    subProcCaption = "Importing 80G(C)"
                
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_PerYES10080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_PerYES10080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
        
'Sch 80G- Section D
                If newrname = "Per5080G.DoneeName" Then
                    subProcCaption = "Importing 80G(D)"
                
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_Per5080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_Per5080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If

                
        
'Sch IT
                If newrname = "TaxP.BSRCode" Then
                    subProcCaption = "Importing IT"
                
                    sfirstbound = SrcBook.Sheets("IT").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("IT").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("IT").Range(rname.name).count
                    dcnt = DestBook.Sheets("IT").Range(rname.name).count
                    DestBook.Sheets("IT").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("IT").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_ITimport
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_ITimport)
                        SrcBook.Sheets("IT").Range(rname.name).Copy
                        DestBook.Sheets("IT").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("IT").Range(rname.name).Copy
                        DestBook.Sheets("IT").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
'Sch TDS1
        
                If newrname = "TDSal.TAN" Then
                    subProcCaption = "Importing TDS1"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TDS1import
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS1import)
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
'Sch TCS
                If newrname = "TCS.TAN" Then
                    subProcCaption = "Importing TCS"
                    sfirstbound = SrcBook.Sheets("TCS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TCS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TCS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TCS").Range(rname.name).count
                    DestBook.Sheets("TCS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TCS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TCSimport
'                       If SrcBook.Sheets("TDS").Range("TCS.AmtClaimedBySpouse").Locked = False Then
'                            DestBook.Sheets("TCS").Unprotect Password:=mdCommon.getmsgstate
'                            DestBook.Sheets("TCS").Range("TCS.AmtClaimedBySpouse").Locked = False
'                            DestBook.Sheets("TCS").Range("TCS.AmtClaimedBySpouse").Interior.Color = "&HCCFFCC"
''                           DestBook.Sheets("TDS").Range("TCS.6panspouse").Locked = False
''                           DestBook.Sheets("TDS").Range("TCS.6panspouse").Interior.Color = "&HCCFFCC"
'                            DestBook.Sheets("Income Details").Protect Password:=mdCommon.getmsgstate
'                            InsertRowsToImport (cnt - dcnt)
'                            Call ExendRangeNameToTable(cnt - dcnt, rngname_TCSimport)
'
'                            SrcBook.Sheets("TCS").Range(rname.name).Copy
'                            DestBook.Sheets("TCS").Range(rname.name).PasteSpecial xlValues
'                       Else
'                        InsertRowsToImport (cnt - dcnt)
'                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TCSimport)
'                        SrcBook.Sheets("TCS").Range(rname.name).Copy
''                        DestBook.Sheets("TCS").Range(rname.name).PasteSpecial xlValues
'                     End If
                    Else
                        SrcBook.Sheets("TCS").Range(rname.name).Copy
                        DestBook.Sheets("TCS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
'Sch TDS2
                If newrname = "TDSoth.TAN" Then
                    subProcCaption = "Importing TDS2"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TDS2import
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS2import)
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
                
                
                  
 'Sch TDS2ii
                If newrname = "TDSoth2.PAN" Then
                    subProcCaption = "Importing TDS2ii"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TDS2iiimport
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS2import)
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
                
'Scheduele BA
                If newrname = "SchBA.IFSC" Then
                    subProcCaption = "Importing BA"
                    sfirstbound = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
        
                    dfirstbound = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
        
                    cnt = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).count
                    dcnt = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).count
                    DestBook.Sheets("Taxes Paid and Verification").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Taxes Paid and Verification").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        
                        setTblinfo_ScheduleBA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_PerSchBA)
                        SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
                        DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
                        DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                End If
                
               ' GSTIN
                
                If newrname = "BP.GSTRno" Then
                subProcCaption = "Importing Exempt Income : Other"
                    sfirstbound = SrcBook.Sheets("BP").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("BP").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("BP").Range(rname.name).count
                    dcnt = DestBook.Sheets("BP").Range(rname.name).count
                    DestBook.Sheets("BP").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("BP").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        mdNOBBP.setTblinfo_GSTIN
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, "BP.GSTRno;BP.GSTRAmount;")
                        SrcBook.Sheets("BP").Range(rname.name).Copy
                        DestBook.Sheets("BP").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("BP").Range(rname.name).Copy
                        DestBook.Sheets("BP").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                
                
                
                          'Sch Other Income
                If newrname = "Sheet5.Nature" Then
                subProcCaption = "Importing Exempt Income : Other"
                    sfirstbound = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).count
                    dcnt = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).count
                    DestBook.Sheets("Taxes Paid and Verification").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Taxes Paid and Verification").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_OtherEI
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_OtherEI)
                        SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
                        DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
                        DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                
                
                If newrname = "Sheet1.Nature" Then
                subProcCaption = "Importing Allownaces"
                    sfirstbound = SrcBook.Sheets("Income Details").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Income Details").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Income Details").Range(rname.name).count
                    dcnt = DestBook.Sheets("Income Details").Range(rname.name).count
                    DestBook.Sheets("Income Details").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Income Details").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_Allowance
                        InsertRowsToImport ((cnt - dcnt) / 17)
                        Call ExendRangeNameToTable((cnt - dcnt) / 17, rngname_Allowances)
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                
                If newrname = "Sheet1.Nature1" Then
                subProcCaption = "Importing Income from OS"
                    sfirstbound = SrcBook.Sheets("Income Details").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Income Details").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Income Details").Range(rname.name).count
                    dcnt = DestBook.Sheets("Income Details").Range(rname.name).count
                    DestBook.Sheets("Income Details").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Income Details").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_OS1
                        InsertRowsToImport ((cnt - dcnt) / 17)
                        Call ExendRangeNameToTable((cnt - dcnt) / 17, rngname_OS1)
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                   
           Next
           
           
            
Application.EnableEvents = True
            
Application.EnableEvents = False
subProcCaption = "Importing Static Contents"
noOfProcessSub = Workbooks(newfilename).Names.count
For Each rname In Workbooks(newfilename).Names
retrievedSheetName = getSheetName(rname)
              
        If retrievedSheetName <> "" Then
            If Range(rname.name) <> "sheet1.ResidentialStatus1" Then
            If DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True Then
                    If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False Then
                        DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
                            If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Locked = False
                            Else
                                DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False
                            End If
                            
                            If (rname.name = "IncD.IntrstPayUs234A") Or (rname.name = "IncD.IntrstPayUs234B") Then
                                DestBook.Worksheets("Income Details").Range(rname.name).MergeArea.Interior.Color = 13434879
                            Else
                                If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Interior.Color = (&HCCFFCC)
                                Else
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HCCFFCC)
                                End If
                            End If
                                                       
                            If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Copy
                                DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.PasteSpecial xlValues
                            Else
                                SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                                DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                            End If
                        DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                    Else
                        DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
                            If rname.name = "sheet1.ZipCode" Then
                                If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Copy
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.PasteSpecial xlValues
                                Else
                                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                                End If
                            End If
                        DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                    End If
            Else
                    If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True Then
                        If (rname.name <> "sheet1.mobileCountryCode") And (rname.name <> "sheet1.mobileCountryCode2") Then
                            DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
                                If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Locked = True
                                Else
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True
                                End If
                                
                                If (rname.name = "IncD.IntrstPayUs234A") Or (rname.name = "IncD.IntrstPayUs234B") Then
                                    DestBook.Worksheets("Income Details").Range(rname.name).MergeArea.Interior.ColorIndex = 2
                                Else
                                    If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Interior.Color = (&HD8D8D8)
                                    Else
                                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HD8D8D8)
                                    End If
                                End If
                                
                                If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Copy
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.PasteSpecial xlValues
                                Else
                                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                                End If
                            DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                        Else
                                If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Copy
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.PasteSpecial xlValues
                                Else
                                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                                End If
                        End If
                    Else
                            If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                                SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Copy
                                DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.PasteSpecial xlValues
                            Else
                                SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                                DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                            End If
                    End If
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
            
            
            
            DestBook.Worksheets("Income Details").Range("sheet1.Status").value = SrcBook.Worksheets("Income Details").Range("sheet1.Status").value

            Dim rangecells As Variant
            Dim chcells As Variant
            Set rangecells = SrcBook.Sheets("Taxes Paid and Verification").Range("Sheet5.Nature").Cells
                            
            For Each chcells In rangecells
            Templock1 (chcells.AddressLocal)
            Next
            
            DestBook.Worksheets("Taxes Paid and Verification").Unprotect Password:=getmsgstate
                SrcBook.Worksheets("Taxes Paid and Verification").Range("Sheet5.Description").Copy
                DestBook.Worksheets("Taxes Paid and Verification").Range("Sheet5.Description").PasteSpecial xlValues
            DestBook.Worksheets("Taxes Paid and Verification").Protect Password:=getmsgstate
            
            
            Dim rangecells1 As Variant
            Dim chcells1 As Variant
            Set rangecells1 = SrcBook.Sheets("Income Details").Range("Sheet1.Nature").Cells
                            
            For Each chcells1 In rangecells1
            Templock2 (chcells1.AddressLocal)
            Next
            
            DestBook.Worksheets("Income Details").Unprotect Password:=getmsgstate
                SrcBook.Worksheets("Income Details").Range("Sheet1.Description").Copy
                DestBook.Worksheets("Income Details").Range("Sheet1.Description").PasteSpecial xlValues
            DestBook.Worksheets("Income Details").Protect Password:=getmsgstate
            
            Dim rangecells2 As Variant
            Dim chcells2 As Variant
            Set rangecells2 = SrcBook.Sheets("Income Details").Range("Sheet1.Nature1").Cells
                            
            For Each chcells2 In rangecells2
            Templock3 (chcells2.AddressLocal)
            Next
            
            DestBook.Worksheets("Income Details").Unprotect Password:=getmsgstate
                SrcBook.Worksheets("Income Details").Range("Sheet1.Description1").Copy
                DestBook.Worksheets("Income Details").Range("Sheet1.Description1").PasteSpecial xlValues
            DestBook.Worksheets("Income Details").Protect Password:=getmsgstate
            
            
            
    
            DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
            DestBook.Worksheets("Income Details").Range("sheet1.ReturnFileSec").value = SrcBook.Worksheets("Income Details").Range("sheet1.ReturnFileSec1").value
'            DestBook.Worksheets("Income Details").Range("sheet1.ReturnType1").value = SrcBook.Worksheets("Income Details").Range("sheet1.ReturnType1").value
            DestBook.Worksheets("Income Details").Range("sheet1.ReceiptNo").value = SrcBook.Worksheets("Income Details").Range("sheet1.ReceiptNo").value
            DestBook.Worksheets("Income Details").Range("sheet1.OrigRetFiledDate").value = SrcBook.Worksheets("Income Details").Range("sheet1.OrigRetFiledDate").value
            DestBook.Worksheets("Income Details").Range("sheet1.NoticeNo").value = SrcBook.Worksheets("Income Details").Range("sheet1.NoticeNo").value
            DestBook.Worksheets("Income Details").Range("sheet1.NoticeDate").value = SrcBook.Worksheets("Income Details").Range("sheet1.NoticeDate").value
            DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
            
            
            
            
            Dim return1 As Variant
            return1 = Mid(DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").value, 1, 2)

            If return1 = "11" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
            ElseIf return1 = "12" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "139(4)-After due date"
            ElseIf return1 = "13" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "142(1)"
            ElseIf return1 = "14" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "148"
            ElseIf return1 = "15" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "153A"
            ElseIf return1 = "16" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "153C"
            ElseIf return1 = "17" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "139(5)-Revised Return"
             ElseIf return1 = "18" Then
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "139(9)"
            ElseIf return1 = "20" Then
            
            'J200924
            'DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "119(2)(b)- after condonation of delay"
            DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = "139(9A) - After condonation of delay u/s 119(2)(b)"
            End If
            
            Sheet1.Range("sheet1.ReturnFileSec1").Formula = "=IF(ISERROR(VLOOKUP(sheet1.ReturnFileSec,ReturnTypenew,2,FALSE)),,VLOOKUP(sheet1.ReturnFileSec,ReturnTypenew,2,FALSE))"
            
            SrcBook.Sheets("Income Details").Unprotect Password:=getmsgstate
            DestBook.Sheets("Income Details").Unprotect Password:=getmsgstate
            If SrcBook.Worksheets("Income Details").Range("IncD.IntrstPayUs234F").MergeArea.Locked = False Then
            DestBook.Worksheets("Income Details").Range("IncD.IntrstPayUs234F").MergeArea.Locked = False
            DestBook.Worksheets("Income Details").Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = RGB(255, 255, 204)
            
            End If
            SrcBook.Sheets("Income Details").Protect Password:=getmsgstate
            DestBook.Sheets("Income Details").Protect Password:=getmsgstate
            

    ProgressBarHide
    fmsgboxStatus "Import Completed"
    'DestBook.Save
    DestBook.Worksheets("Income Details").Select
    Worksht_Active
    Set SrcBook = Nothing

Else
    fmsgboxStatus "* Source file must not have same name As destination File"
End If
End If
End Sub

'Sub returnfilesecimport()
''Dim DestBook As Workbook, SrcBook As Workbook
'
'    DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value = SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").value
'                 'MsgBox SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").value
'                    Dim return1 As Variant
'                    Dim Filingtype As Variant
'                    return1 = DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value
''                    Filingtype = DestBook.Sheets("Income Details").Range("sheet1.ReturnType1").value
'
'
'
''          Sheet1.Unprotect Password:=getmsgstate
''          If Mid(Filingtype, 1, 1) = "O" Then
''          Range("sheet1.ReturnType1").value = "Original"
''          ElseIf Mid(Filingtype, 1, 3) = "Rec" Then
''          Range("sheet1.ReturnType1").value = "Rectification"
''          ElseIf Mid(Filingtype, 1, 1) = "R" Then
''          Range("sheet1.ReturnType1").value = "Revised"
''          Else
''          Range("sheet1.ReturnType1").value = "In response to notice u/s 139(9)"
''          End If
'
'
'        If return1 = "(Select)" Then
'           Range("sheet1.ReturnFileSec1").value = "(Select)"
'        ElseIf Mid(return1, 1, 2) = "11" Then
'           Range("sheet1.ReturnFileSec").value = "Voluntarily on or  before the due date under section 139(1)"
'        ElseIf Mid(return1, 1, 2) = "12" Then
'           Range("sheet1.ReturnFileSec").value = "Voluntarily after the due date under section 139(4))"
'        ElseIf Mid(return1, 1, 2) = "13" Then
'           Range("sheet1.ReturnFileSec").value = "In response to notice under section 142(1)"
'        ElseIf Mid(return1, 1, 2) = "14" Then
'           Range("sheet1.ReturnFileSec").value = "In response to notice under section 148"
'        ElseIf Mid(return1, 1, 2) = "15" Then
'           Range("sheet1.ReturnFileSec").value = "In response to notice under section 153A"
'        ElseIf Mid(return1, 1, 2) = "16" Then
'           Range("sheet1.ReturnFileSec").value = "In response to notice under section 153C read with section 153A."
'        ElseIf Mid(return1, 1, 2) = "17" Then
'           Range("sheet1.ReturnFileSec").value = "Revised return under section 139(5)"
'        ElseIf Mid(return1, 1, 2) = "18" Then
'           Range("sheet1.ReturnFileSec").value = "In response to a notice u/s 139(9)where the original return filed was a Defective return."
'        ElseIf Mid(return1, 1, 2) = "20" Then
'           Range("sheet1.ReturnFileSec").value = "section 139(4) read with section 119(2)(b)"
'        End If
'
'        DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").value = DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").value
'
'
'        If Range("sheet1.ReturnFileSec").value = "(Select)" Then
'           Range("sheet1.ReturnFileSec1").value = "(Select)"
'        ElseIf Range("sheet1.ReturnFileSec").value = "Voluntarily on or  before the due date under section 139(1)" Then
'           Range("sheet1.ReturnFileSec1").value = "11 - Voluntarily on or  before the due date under section 139(1)"
'        ElseIf Range("sheet1.ReturnFileSec").value = "Voluntarily after the due date under section 139(4))" Then
'           Range("sheet1.ReturnFileSec1").value = "12 - Voluntarily after the due date under section 139(4))"
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice under section 142(1)" Then
'           Range("sheet1.ReturnFileSec1").value = "13 - In response to notice under section 142(1)"
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice under section 148" Then
'           Range("sheet1.ReturnFileSec1").value = "14 - In response to notice under section 148"
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice under section 153A" Then
'           Range("sheet1.ReturnFileSec1").value = "15 - In response to notice under section 153A"
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice under section 153C read with section 153A." Then
'           Range("sheet1.ReturnFileSec1").value = "16 - In response to notice under section 153C read with section 153A."
'        ElseIf Range("sheet1.ReturnFileSec").value = "Revised return under section 139(5)" Then
'           Range("sheet1.ReturnFileSec1").value = "17 - Revised return under section 139(5)"
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to a notice u/s 139(9)where the original return filed was a Defective return." Then
'           Range("sheet1.ReturnFileSec1").value = "18 - In response to a notice u/s 139(9)where the original return filed was a Defective return."
'        ElseIf Range("sheet1.ReturnFileSec").value = "section 139(4) read with section 92 CD" Then
'           Range("sheet1.ReturnFileSec1").value = "20 - section 139(4) read with section 119(2)(b)"
'        End If
'
'End Sub

Function getSheetName(rn As Variant) As String
Dim temp As Variant
temp = Replace(Mid(rn, 2, InStr(1, rn, "!") - 2), "'", "")
getSheetName = temp
End Function

Function InsertRowsToImport(Optional vRows As Long = 0)
    Dim x As Long
    Dim strpassword As String
    Dim sht As Worksheet, shts() As String, i As Long
    
    strpassword = mdCommon.getmsgstate
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
        
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault
    
    ActiveSheet.Protect Password:=strpassword
        On Error Resume Next
            Selection.Offset(1).Resize(vRows).EntireRow. _
            SpecialCells(xlCellTypeAllValidation).ClearContents
    Next sht
  '  ActiveSheet.Protect Password:=strpassword
End Function

Function cmdFileDialog() As String
    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
 
    cmdFileDialog = ""
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
      .AllowMultiSelect = False
      .Filters.Clear
      .Filters.add "Microsoft Office Excel Workbook", "*.xls,*.xlsm"
        If .Show = True Then
            For Each varFile In .SelectedItems
               cmdFileDialog = varFile
            Next
        End If
    End With
End Function

Sub setTblinfo_Per10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet6.Range("Per10080G.DoneeName").count
    Set rangecells = Sheet6.Range("Per10080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G = ccount
    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.Donationcash;Per10080G.DonationAmt;Per10080G.EligibleAmt;Data_80G_A;comb_80G_A;Comb_donation_80G_A;Comb_PAN_80G_A;"
End Sub
Sub setTblinfo_ScheduleBA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet5.Range("SchBA.IFSC").count
    Set rangecells = Sheet5.Range("SchBA.IFSC").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not Trim(rangecells.item(mIntCtr).value) = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerSchBA = ccount
    rngname_PerSchBA = "SchBA.IFSC;SchBA.BankName;SchBA.AcntNo;" 'SchBA.DepositedCash;"
End Sub
Sub setTblinfo_OtherEI()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Range("Sheet5.Nature").count
 Set rangecells = Range("Sheet5.Nature").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_OtherEI = ccount
 rngname_OtherEI = "Sheet5.Nature;Sheet5.Description;Sheet5.Amount;SEC38Check;"
End Sub
Sub setTblinfo_Allowance()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Range("Sheet1.Nature").count
 Set rangecells = Range("Sheet1.Nature").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_Allowances = ccount
 rngname_Allowances = "Sheet1.Nature;Sheet1.Description;Sheet1.Amount;"
End Sub
Sub setTblinfo_OS1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Range("Sheet1.Nature1").count
 Set rangecells = Range("Sheet1.Nature1").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_OS1 = ccount
 rngname_OS1 = "Sheet1.Nature1;Sheet1.Description1;Sheet1.Amount1;"
End Sub



Sub setTblinfo_PerNO5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet6.Range("PerNO5080G.DoneeName").count
    Set rangecells = Sheet6.Range("PerNO5080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G = ccount
    rngname_PerNO5080G = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationCash;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Data_80G_B;comb_80G_B;Comb_donation_80G_B;Comb_PAN_80G_B;"
End Sub
Sub setTblinfo_PerYES10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet6.Range("PerYES10080G.DoneeName").count
    Set rangecells = Sheet6.Range("PerYES10080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G = ccount
    rngname_PerYES10080G = "PerYES10080G.DoneeName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationCash;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Data_80G_C;comb_80G_C;Comb_donation_80G_C;Comb_PAN_80G_C;"
End Sub

Sub setTblinfo_Per5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet6.Range("Per5080G.DoneeName").count
    Set rangecells = Sheet6.Range("Per5080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G = ccount
    rngname_Per5080G = "Per5080G.DoneeName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationCash;Per5080G.DonationAmt;Per5080G.EligibleAmt;Data_80G_D;comb_80G_D;Comb_donation_80G_D;Comb_PAN_80G_D;"
End Sub

Sub Templock1(targetadd)


If Not Application.Intersect(Sheet5.Range("Sheet5.Nature"), Sheet5.Range(targetadd)) Is Nothing Then
 
        If Sheet5.Range(targetadd).value = "Any Other" Then
            Sheet5.Unprotect Password:=getmsgstate
'            Sheet5.Range(Replace(Targetadd, "G", "H")).value = ""
            Sheet5.Range(Replace(targetadd, "G", "H")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet5.Range(Replace(targetadd, "G", "H")).Locked = False
            
            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
            
            Sheet5.Protect Password:=getmsgstate
        Else
            Sheet5.Unprotect Password:=getmsgstate
            Sheet5.Range(Replace(targetadd, "G", "H")).value = "Not Applicable"
            Sheet5.Range(Replace(targetadd, "G", "H")).Interior.Color = (&HD8D8D8)
            Sheet5.Range(Replace(targetadd, "G", "H")).Locked = True
            Sheet5.Protect Password:=getmsgstate
    End If
End If

End Sub
Sub Templock2(targetadd)


If Not Application.Intersect(Sheet1.Range("Sheet1.Nature"), Sheet1.Range(targetadd)) Is Nothing Then

        If Sheet1.Range(targetadd).value = "Any Other" Then
            Sheet1.Unprotect Password:=getmsgstate
'            Sheet5.Range(Replace(Targetadd, "G", "H")).value = ""
            Sheet1.Range(Replace(targetadd, "I", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(targetadd, "I", "Z")).Locked = False
            
            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
            
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(targetadd, "I", "Z")).value = "Not Applicable"
            Sheet1.Range(Replace(targetadd, "I", "Z")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(targetadd, "I", "Z")).Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If

End Sub
Sub Templock3(targetadd)


If Not Application.Intersect(Sheet1.Range("Sheet1.Nature1"), Sheet1.Range(targetadd)) Is Nothing Then

        If Sheet1.Range(targetadd).value = "Any Other" Then
            Sheet1.Unprotect Password:=getmsgstate
'            Sheet5.Range(Replace(Targetadd, "G", "H")).value = ""
            Sheet1.Range(Replace(targetadd, "I", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(targetadd, "I", "Z")).Locked = False
            
            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
            
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(targetadd, "I", "Z")).value = "Not Applicable"
            Sheet1.Range(Replace(targetadd, "I", "Z")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(targetadd, "I", "Z")).Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If

End Sub

Sub ExendRangeNameToTable(numberofrows As Long, rangenamestring As Variant)
    Dim i As Long
    Dim x As Long
    Dim firstbound As String
    Dim temp As Variant
    Dim upperbound As String
    Dim lastbound As String
    Dim RangeAddress As Variant
    
    rangenamestring = Split(rangenamestring, ";")
    For i = 0 To UBound(rangenamestring) - 1
        firstbound = Range(rangenamestring(i)).Address
        temp = Split(firstbound, "$")
        upperbound = UBound(temp)
        temp = temp(UBound(temp))
        x = CLng(temp) + numberofrows
        lastbound = Replace(firstbound, temp, x)
        If upperbound < 3 Then
            RangeAddress = firstbound & ":" & lastbound
        Else
            RangeAddress = lastbound
        End If
        ThisWorkbook.Names.add name:=rangenamestring(i), _
                 RefersTo:="=" & RangeAddress, Visible:=True
    Next
End Sub

Sub setTblinfo_ITimport()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet18.Range("TaxP.BSRCode").count
 Set rangecells = Sheet18.Range("TaxP.BSRCode").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_ITimport = ccount
 '----------02/12/2025
 'rngname_ITimport = "TaxP.BSRCode;TaxP.DateDep;TaxP.SrlNoOfChaln;TaxP.Amt;IT.FormulaOFS;FormulaOfQ;FormulaOfSAT;FormulaOfSAT1;"
 rngname_ITimport = "TaxP.BSRCode;TaxP.DateDep;TaxP.SrlNoOfChaln;TaxP.Amt;IT.FormulaOFS;FormulaOfQ;FormulaOfSAT;FormulaOfSAT1;FormulaOfExSAT;"
 '-----------------
End Sub


Sub setTblinfo_TCSimport()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet17.Range("TCS.TAN").count
 Set rangecells = Sheet17.Range("TCS.TAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TCSimport = ccount
 rngname_TCSimport = "TCS.TAN;TCS.EmployerOrDeductorOrCollecterName;TCS.AmountPaid;TCS.TotalTCS;TCS.AmtTCSClaimedThisYear;"
 End Sub


Sub setTblinfo_TDS1import()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("TDSal.TAN").count
    Set rangecells = Sheet4.Range("TDSal.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS1import = ccount
    rngname_TDS1import = "TDSal.TAN;TDSal.EmployerOrDeductorOrCollecterName;TDSal.IncChrgSal;TDSal.TotalTDSSal;"
End Sub


Sub setTblinfo_TDS2import()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("TDsOthr.TAN").count
    Set rangecells = Sheet4.Range("TDsOthr.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS2import = ccount
    rngname_TDS2import = "TDsOthr.TAN;TDsOthr.finyear;TDsOthr.tds;TDsOthr.tdsdeducted;TDsOthr.tdsclaimed;TDsOthr.grossamount;TDsOthr.headincome;TDsOthr.tdscarried;"
End Sub

Sub setTblinfo_TDS2iiimport()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("TDsOthr2.PAN").count
    Set rangecells = Sheet4.Range("TDsOthr2.PAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS2iiimport = ccount
    rngname_TDS2iiimport = "TDsOthr2.PAN;TDsOthr2.AADHAAR;TDsOthr2.finyear;TDsOthr2.tds;TDsOthr2.tdsdeducted;TDsOthr2.tdsclaimed;TDsOthr2.grossamount;TDsOthr2.headamount;TDsOthr2.carried;"
End Sub


