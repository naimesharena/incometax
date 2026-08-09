Attribute VB_Name = "Sheet40"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String
Dim rngname As String

Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet40.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Sheet40.Shapes("Button 13").Left = Range("Button_FSI_1").Left
Sheet40.Shapes("Button 13").Top = Range("Button_FSI_1").Top + 5

Sheet40.Protect Password:=sPassword
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False


    Dim i As Variant
    Application.EnableEvents = False
'Change-27.01.2023.101.FSI.2A
'     If InStr(1, Target.name.name, "FSI_CountryCode") >= 1 Then
'        For i = 1 To Sheet40.Range("FSI_Count").value
'            If Not (Target.Address = Sheet40.Range("FSI_CountryCode" & i).Address) Then
'                If Range(Target.Address).value = Sheet40.Range("FSI_CountryCode" & i).value Then
'                    If Not Range(Target.Address).value = "(Select)" Or Range(Target.Address).value = "" Then
''                    Change-10.01.2023.102.FSI.02A
'                        'fmsgboxStatus "Do not enter same country twice" ', vbOKOnly, "ITR 3"
''                        fmsgboxStatus "Country code cannot be repeated in schedule FSI" ', vbOKOnly, "ITR 3"
''                        end---
'                        Range(Target.Address).value = "(Select)"
'                        Range(Target.Address).Select
'                    End If
'                End If
'            End If
'        Next
'    End If
    
'Change-2, RIS-153, 25.05.2022
    Dim s
    s = Trim(7 + 6 * (Sheet40.Range("FSI_Count").value - 1))
    If Not Intersect(Sheet40.Range(Target.Address), Sheet40.Range("E7:E" & s)) Is Nothing Then
        If Target.value = "" Or Target.value = "(Select)" Then
            Dim c
            c = Mid(Target.name.name, 16, Len(Target.name.name) - 15)
            Sheet27.Unprotect Password:=getmsgstate
            Sheet27.Range("TR_Country").Cells(CDbl(c)).value = ""
            Sheet27.Protect Password:=getmsgstate
        End If
    End If
'---end change---
    
    'Change-18.07.2023.103 -> TR Change
    If Target.Column = 5 Then
    
        Sheet27.Range("TR_ReliefClaimedUsSection").value = ""
    End If
    
    'Commented by Aavula Naresh New Dev
    
'    If InStr(1, Target.name.name, "FSI_TaxIdentificationNo") >= 1 Then
'        For i = 1 To Sheet40.Range("FSI_Count").value
'            If Not (Target.Address = Sheet40.Range("FSI_TaxIdentificationNo" & i).Address) Then
'                If Range(Target.Address).value = Sheet40.Range("FSI_TaxIdentificationNo" & i).value Then
'                    If Not Range(Target.Address).value = "" Then
'                        fmsgboxStatus "The Same TIN allowed only once" ', vbOKOnly, "ITR 3"
'                        Range(Target.Address).value = ""
'                        Range(Target.Address).Select
'                    End If
'                End If
'            End If
'        Next
'    End If
    
    
    If InStr(1, Target.name.name, "FSI_IncFromSal1") >= 0 Then
        For i = 1 To Sheet40.Range("FSI_Count").value
            If Not (Target.Address = Sheet40.Range("FSI_IncFromSal1" & i).Address) Then
                If Range(Target.Address).value <> "" Then
                      MsgBox ("MSG")
                    End If
                End If
            
        Next
    End If
    
    
    
    

endline:
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_Deactivate()
On Error Resume Next
Dim activesh As Variant

activesh = ActiveSheet.name
Application.EnableEvents = False

 PopulateScheduleTR


Application.EnableEvents = True
Worksheets(activesh).Activate
End Sub
Sub PopulateScheduleTR()
'On Error GoTo End1

Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Dim rangecellstr_c As Range
    Dim rangecellstr_tin As Range
    Dim rangecellstr_txno As Range
    Dim rangecellstr_txr As Range
    Dim rangecellstr_section As Range
    Dim i, j, k, b, end_TR, end_FSI As Long
    Dim Country_FSI, Country_TR, TaxIdentificationNo_FSI As Variant
    Dim TXNOSAL_FSI, TXRSAL_FSI, TXNOHP_FSI, TXRHP_FSI, TXNOBusiness_FSI, TXRBusiness_FSI, TXNOCapGain_FSI, TXRCapGain_FSI, TXNOOthSrc_FSI, TXROthSrc_FSI As Variant
    Dim TIN_TR, TaxPaidOutsideIndia_TR, TaxReliefOutsideIndia_TR, ReliefClaimedUsSection_TR As Variant

    Dim rangecellsfsi_c As Range
    Dim rangecellsfsi_tin As Range
    Dim rangecellsfsi_txno1 As Range
    Dim rangecellsfsi_txr1 As Range

    Dim rangecellsfsi_txno2 As Range
    Dim rangecellsfsi_txr2 As Range

    Dim rangecellsfsi_txno3 As Range
    Dim rangecellsfsi_txr3 As Range

    Dim rangecellsfsi_txno4 As Range
    Dim rangecellsfsi_txr4 As Range

    Dim rangecellsfsi_txno5 As Range
    Dim rangecellsfsi_txr5 As Range

    Dim IntCtr, mIntCells As Long
    Dim rangecells As Range
    Dim countrycd As Variant

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   sourceSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.ScreenUpdating = False

    'sourceSheet.Range("TR_Country").EntireRow.Hidden = True
    sourceSheet.Range("TR_TIN").value = ""
    sourceSheet.Range("TR_TaxPaidOutsideIndia").value = 0
    sourceSheet.Range("TR_TaxReliefOutsideIndia").value = 0

    'FSI Range
    Dim IntCells, mIntCtr As Long
    'IntCells = Sheet40.Range("FSI_CountryCode").count
    For i = 1 To Sheet40.Range("FSI_Count").value
    countrycd = Sheet40.Range("FSI_CountryCode" & i).value



        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
        If Not countrycd = "" Then
            end_FSI = end_FSI + 1
            Else
            Exit For
        End If
    Next
    rngname = "FSI_CountryCode||FSI_TaxIdentificationNo||FSI_IncFromSal||FSI_TXNOSAL||FSI_TXNISAL||FSI_TXRSAL||FSI_DTAASAL||FSI_IncFromHP||FSI_TXNOHP||FSI_TXNIHP||FSI_TXRHP||FSI_DTAAHP||FSI_IncFromBP||FSI_TXNOBP||FSI_TXNIBP||FSI_TXRBP||FSI_DTAABP||FSI_IncCapGain||FSI_TXNOCapGain||FSI_TXnICapGain||FSI_TXRCapGain||FSI_DTAACapGain||FSI_IncOthSrc||FSI_TXNOOthSrc||FSI_TXNIOthSrc||FSI_TXROthSrc||FSI_DTAAOthSrc||FSI_IncT||FSI_TXNOT||FSI_TXNIT||FSI_TXRT"
    'end_FSI = 2
    'TR range
    rngname = "TR_Country||TR_TIN||TR_TaxPaidOutsideIndia||TR_TaxReliefOutsideIndia||TR_ReliefClaimedUsSection"
'    end_TR = 250
    end_TR = end_FSI
   
    Set rangecellstr_c = sourceSheet.Range("TR_Country").Cells
    Set rangecellstr_tin = sourceSheet.Range("TR_TIN").Cells
    Set rangecellstr_txno = sourceSheet.Range("TR_TaxPaidOutsideIndia").Cells
    Set rangecellstr_txr = sourceSheet.Range("TR_TaxReliefOutsideIndia").Cells
    Set rangecellstr_section = sourceSheet.Range("TR_ReliefClaimedUsSection").Cells

    ReDim Country_TR(end_TR)
    ReDim TIN_TR(end_TR)
    ReDim TaxPaidOutsideIndia_TR(end_TR)
    ReDim TaxReliefOutsideIndia_TR(end_TR)
    ReDim ReliefClaimedUsSection_TR(end_TR)


    ReDim Country_FSI(end_FSI)
    ReDim TaxIdentificationNo_FSI(end_FSI)

    ReDim TXNOSAL_FSI(end_FSI)
    ReDim TXRSAL_FSI(end_FSI)

    ReDim TXNOHP_FSI(end_FSI)
    ReDim TXRHP_FSI(end_FSI)

    ReDim TXNOCapGain_FSI(end_FSI)
    ReDim TXRCapGain_FSI(end_FSI)


    ReDim TXNOOthSrc_FSI(end_FSI)
    ReDim TXROthSrc_FSI(end_FSI)

    ReDim TXNOBusiness_FSI(end_FSI)
    ReDim TXRBusiness_FSI(end_FSI)
     
    
    
     
    For i = 1 To end_FSI
    'Country & TIN
    
    Set rangecellsfsi_c = Sheet40.Range("FSI_CountryCode" & i).Cells
   
    Set rangecellsfsi_tin = Sheet40.Range("FSI_TaxIdentificationNo" & i).Cells
    
   
    'Salary
    Set rangecellsfsi_txno1 = Sheet40.Range("FSI_TXNOSAL" & i).Cells
    Set rangecellsfsi_txr1 = Sheet40.Range("FSI_TXRSAL" & i).Cells

    'House Property
    Set rangecellsfsi_txno2 = Sheet40.Range("FSI_TXNOHP" & i).Cells
    Set rangecellsfsi_txr2 = Sheet40.Range("FSI_TXRHP" & i).Cells

    'Capital Gain
    Set rangecellsfsi_txno3 = Sheet40.Range("FSI_TXNOCapGain" & i).Cells
    Set rangecellsfsi_txr3 = Sheet40.Range("FSI_TXRCapGain" & i).Cells

    'Other Source
    Set rangecellsfsi_txno4 = Sheet40.Range("FSI_TXNOOthSrc" & i).Cells
    Set rangecellsfsi_txr4 = Sheet40.Range("FSI_TXROthSrc" & i).Cells

    'Business
    Set rangecellsfsi_txno5 = Sheet40.Range("FSI_IncFromBP" & i).Cells
    Set rangecellsfsi_txr5 = Sheet40.Range("FSI_TXNOBP" & i).Cells
        Country_FSI(i) = rangecellsfsi_c.item(1).value
        TaxIdentificationNo_FSI(i) = rangecellsfsi_tin.item(1).value
        
        'TXNOSAL_FSI(i) = rangecellsfsi_txno1.Item(i).Value
        'TXRSAL_FSI(i) = rangecellsfsi_txr1.Item(i).Value

        TXNOHP_FSI(i) = rangecellsfsi_txno2.item(1).value
        TXRHP_FSI(i) = rangecellsfsi_txr2.item(1).value

        TXNOCapGain_FSI(i) = rangecellsfsi_txno3.item(1).value
        TXRCapGain_FSI(i) = rangecellsfsi_txr3.item(1).value

        TXNOOthSrc_FSI(i) = rangecellsfsi_txno4.item(1).value
        TXROthSrc_FSI(i) = rangecellsfsi_txr4.item(1).value

        TXNOBusiness_FSI(i) = rangecellsfsi_txno5.item(1).value
        TXRBusiness_FSI(i) = rangecellsfsi_txr5.item(1).value
        
        Country_TR(i) = Country_FSI(i)
        rangecellstr_c(i).value = Country_TR(i)
        
        TIN_TR(i) = TaxIdentificationNo_FSI(i)
        rangecellstr_tin(i).value = TIN_TR(i)
        
        TaxPaidOutsideIndia_TR(i) = Sheet40.Range("FSI_TXNOT" & i).item(1)
        rangecellstr_txno(i).value = TaxPaidOutsideIndia_TR(i)
         
        TaxReliefOutsideIndia_TR(i) = Sheet40.Range("FSI_TXRT" & i).item(1)
        rangecellstr_txr(i).value = TaxReliefOutsideIndia_TR(i)
         
        'FSI_TaxIdentificationNo
'        end_TR = 0
'        For j = 1 To end_TR
'
'            Country_TR(j) = rangecellstr_c.item(j).value
'
'            ' FSI_TXRSAL ETC'
'            If (Country_FSI(i) = Country_TR(j)) Then
'
'            TIN_TR(j) = TaxIdentificationNo_FSI(i)
'            'TaxPaidOutsideIndia_TR(j) = TXNOHP_FSI(i) + TXNOCapGain_FSI(i) + TXNOOthSrc_FSI(i) + TXNOBusiness_FSI(i)
'            'TaxReliefOutsideIndia_TR(j) = TXRHP_FSI(i) + TXRCapGain_FSI(i) + TXROthSrc_FSI(i) + TXRBusiness_FSI(i)
'            TaxPaidOutsideIndia_TR(j) = Sheet40.Range("FSI_TXNOT" & i).item(1)
'            TaxReliefOutsideIndia_TR(j) = Sheet40.Range("FSI_TXRT" & i).item(1)
'
'            rangecellstr_tin(j).value = TIN_TR(j)
'            rangecellstr_txno(j).value = TaxPaidOutsideIndia_TR(j)
'            rangecellstr_txr(j).value = TaxReliefOutsideIndia_TR(j)
''            rangecellstr_tin(j).EntireRow.Hidden = False
'            End If
'        Next

'Change-17.07.2023.103 -> TR Change
  For j = 1 To end_TR
        Country_TR(j) = rangecellstr_c.item(j).value
        Dim flagDTAA
                k = 0
                flagDTAA = False
                For k = 1 To SheetALL.Range("TR_DTAA_Countries").count
                    If Mid(Country_TR(j), InStr(1, Country_TR(j), ":") + 1) = Mid(SheetALL.Range("TR_DTAA_Countries").item(k, 1).value, InStr(1, SheetALL.Range("TR_DTAA_Countries").item(k, 1).value, ":") + 1) Then
                        flagDTAA = True
                    End If
                Next k
                If flagDTAA = True Then
                    If rangecellstr_section(j).value = "" Or rangecellstr_section(j).value = "(Select)" Then
                        rangecellstr_section(j).value = 90
                    End If
                Else
                    If rangecellstr_section(j).value = "" Or rangecellstr_section(j).value = "(Select)" Then
                        rangecellstr_section(j).value = 91
                    End If
                End If
        Next
    Next
End1:
'----------------Lock Password-------------------START---
   sourceSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.ScreenUpdating = True
End Sub

Sub PopulateScheduleTRold1()
'On Error GoTo End1
Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Dim rangecellstr_c As Range
    Dim rangecellstr_tin As Range
    Dim rangecellstr_txno As Range
    Dim rangecellstr_txr As Range
    Dim rangecellstr_section As Range
    Dim i, j, k, b, end_TR, end_FSI As Long
    Dim Country_FSI, Country_TR, TaxIdentificationNo_FSI As Variant
    Dim TXNOSAL_FSI, TXRSAL_FSI, TXNOHP_FSI, TXRHP_FSI, TXNOBusiness_FSI, TXRBusiness_FSI, TXNOCapGain_FSI, TXRCapGain_FSI, TXNOOthSrc_FSI, TXROthSrc_FSI As Variant
    Dim TIN_TR, TaxPaidOutsideIndia_TR, TaxReliefOutsideIndia_TR, ReliefClaimedUsSection_TR As Variant

    Dim rangecellsfsi_c As Range
    Dim rangecellsfsi_tin As Range
    Dim rangecellsfsi_txno1 As Range
    Dim rangecellsfsi_txr1 As Range

    Dim rangecellsfsi_txno2 As Range
    Dim rangecellsfsi_txr2 As Range

    Dim rangecellsfsi_txno3 As Range
    Dim rangecellsfsi_txr3 As Range

    Dim rangecellsfsi_txno4 As Range
    Dim rangecellsfsi_txr4 As Range

    Dim rangecellsfsi_txno5 As Range
    Dim rangecellsfsi_txr5 As Range

    Dim IntCtr, mIntCells As Long
    Dim rangecells As Range
    Dim countrycd As Variant

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   sourceSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.ScreenUpdating = False

    'sourceSheet.Range("TR_Country").EntireRow.Hidden = True
    sourceSheet.Range("TR_TIN").value = ""
    sourceSheet.Range("TR_TaxPaidOutsideIndia").value = 0
    sourceSheet.Range("TR_TaxReliefOutsideIndia").value = 0

    'FSI Range
    Dim IntCells, mIntCtr As Long
    'IntCells = Sheet40.Range("FSI_CountryCode").count
    For i = 1 To Sheet40.Range("FSI_Count").value
    countrycd = Sheet40.Range("FSI_CountryCode" & i).value



        If isdropdownblank(countrycd) Then
            countrycd = ""
        End If
        If Not countrycd = "" Then
            end_FSI = end_FSI + 1
            Else
            Exit For
        End If
    Next
    rngname = "FSI_CountryCode||FSI_TaxIdentificationNo||FSI_IncFromSal||FSI_TXNOSAL||FSI_TXNISAL||FSI_TXRSAL||FSI_DTAASAL||FSI_IncFromHP||FSI_TXNOHP||FSI_TXNIHP||FSI_TXRHP||FSI_DTAAHP||FSI_IncFromBP||FSI_TXNOBP||FSI_TXNIBP||FSI_TXRBP||FSI_DTAABP||FSI_IncCapGain||FSI_TXNOCapGain||FSI_TXnICapGain||FSI_TXRCapGain||FSI_DTAACapGain||FSI_IncOthSrc||FSI_TXNOOthSrc||FSI_TXNIOthSrc||FSI_TXROthSrc||FSI_DTAAOthSrc||FSI_IncT||FSI_TXNOT||FSI_TXNIT||FSI_TXRT"
    'end_FSI = 2
    'TR range
    rngname = "TR_Country||TR_TIN||TR_TaxPaidOutsideIndia||TR_TaxReliefOutsideIndia||TR_ReliefClaimedUsSection"
'    end_TR = 250
    end_TR = end_FSI
    
    Set rangecellstr_c = sourceSheet.Range("TR_Country").Cells
    Set rangecellstr_tin = sourceSheet.Range("TR_TIN").Cells
    Set rangecellstr_txno = sourceSheet.Range("TR_TaxPaidOutsideIndia").Cells
    Set rangecellstr_txr = sourceSheet.Range("TR_TaxReliefOutsideIndia").Cells
    Set rangecellstr_section = sourceSheet.Range("TR_ReliefClaimedUsSection").Cells

    ReDim Country_TR(end_TR)
    ReDim TIN_TR(end_TR)
    ReDim TaxPaidOutsideIndia_TR(end_TR)
    ReDim TaxReliefOutsideIndia_TR(end_TR)
    ReDim ReliefClaimedUsSection_TR(end_TR)


    ReDim Country_FSI(end_FSI)
    ReDim TaxIdentificationNo_FSI(end_FSI)

    ReDim TXNOSAL_FSI(end_FSI)
    ReDim TXRSAL_FSI(end_FSI)

    ReDim TXNOHP_FSI(end_FSI)
    ReDim TXRHP_FSI(end_FSI)

    ReDim TXNOCapGain_FSI(end_FSI)
    ReDim TXRCapGain_FSI(end_FSI)


    ReDim TXNOOthSrc_FSI(end_FSI)
    ReDim TXROthSrc_FSI(end_FSI)

    ReDim TXNOBusiness_FSI(end_FSI)
    ReDim TXRBusiness_FSI(end_FSI)
     
    
    
     
    For i = 1 To end_FSI
    'Country & TIN
    
    Set rangecellsfsi_c = Sheet40.Range("FSI_CountryCode" & i).Cells
    Set rangecellsfsi_tin = Sheet40.Range("FSI_TaxIdentificationNo" & i).Cells
    
   
    'Salary
    Set rangecellsfsi_txno1 = Sheet40.Range("FSI_TXNOSAL" & i).Cells
    Set rangecellsfsi_txr1 = Sheet40.Range("FSI_TXRSAL" & i).Cells

    'House Property
    Set rangecellsfsi_txno2 = Sheet40.Range("FSI_TXNOHP" & i).Cells
    Set rangecellsfsi_txr2 = Sheet40.Range("FSI_TXRHP" & i).Cells

    'Capital Gain
    Set rangecellsfsi_txno3 = Sheet40.Range("FSI_TXNOCapGain" & i).Cells
    Set rangecellsfsi_txr3 = Sheet40.Range("FSI_TXRCapGain" & i).Cells

    'Other Source
    Set rangecellsfsi_txno4 = Sheet40.Range("FSI_TXNOOthSrc" & i).Cells
    Set rangecellsfsi_txr4 = Sheet40.Range("FSI_TXROthSrc" & i).Cells

    'Business
    Set rangecellsfsi_txno5 = Sheet40.Range("FSI_IncFromBP" & i).Cells
    Set rangecellsfsi_txr5 = Sheet40.Range("FSI_TXNOBP" & i).Cells
        Country_FSI(i) = rangecellsfsi_c.item(1).value
        TaxIdentificationNo_FSI(i) = rangecellsfsi_tin.item(1).value

        'TXNOSAL_FSI(i) = rangecellsfsi_txno1.Item(i).Value
        'TXRSAL_FSI(i) = rangecellsfsi_txr1.Item(i).Value

        TXNOHP_FSI(i) = rangecellsfsi_txno2.item(1).value
        TXRHP_FSI(i) = rangecellsfsi_txr2.item(1).value

        TXNOCapGain_FSI(i) = rangecellsfsi_txno3.item(1).value
        TXRCapGain_FSI(i) = rangecellsfsi_txr3.item(1).value

        TXNOOthSrc_FSI(i) = rangecellsfsi_txno4.item(1).value
        TXROthSrc_FSI(i) = rangecellsfsi_txr4.item(1).value

        TXNOBusiness_FSI(i) = rangecellsfsi_txno5.item(1).value
        TXRBusiness_FSI(i) = rangecellsfsi_txr5.item(1).value

        'FSI_TaxIdentificationNo
        For j = 1 To end_TR
            Country_TR(j) = rangecellstr_c.item(j).value

            ' FSI_TXRSAL ETC'
            If (Country_FSI(i) = Country_TR(j)) Then

            TIN_TR(j) = TaxIdentificationNo_FSI(i)
            'TaxPaidOutsideIndia_TR(j) = TXNOHP_FSI(i) + TXNOCapGain_FSI(i) + TXNOOthSrc_FSI(i) + TXNOBusiness_FSI(i)
            'TaxReliefOutsideIndia_TR(j) = TXRHP_FSI(i) + TXRCapGain_FSI(i) + TXROthSrc_FSI(i) + TXRBusiness_FSI(i)
            TaxPaidOutsideIndia_TR(j) = Sheet40.Range("FSI_TXNOT" & i).item(1)
            TaxReliefOutsideIndia_TR(j) = Sheet40.Range("FSI_TXRT" & i).item(1)

            rangecellstr_tin(j).value = TIN_TR(j)
            rangecellstr_txno(j).value = TaxPaidOutsideIndia_TR(j)
            rangecellstr_txr(j).value = TaxReliefOutsideIndia_TR(j)
'            rangecellstr_tin(j).EntireRow.Hidden = False
            End If
        Next
    Next
End1:
'----------------Lock Password-------------------START---
   sourceSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.ScreenUpdating = True
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
    'Cancel = True
End Sub

