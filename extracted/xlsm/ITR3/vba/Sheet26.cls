Attribute VB_Name = "Sheet26"
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
   Sheet26.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Sheet26.Shapes("Button 48").Left = Range("Button_FSI").Left
Sheet26.Shapes("Button 48").Top = Range("Button_FSI").Top + 5

Sheet26.Protect Password:=sPassword
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

    Application.EnableEvents = False
    If Not IsEmpty(Trim(Target.value)) Then
            If Target.Validation.Type = 3 Then
                Dim nOfCount As Long
                nOfCount = EfilingCommon.CountDropdown(Trim(Target.value), Sheet26.Range("FSI_CountryCode").Cells, Sheet26.Range("FSI_CountryCode").count)
                If nOfCount > 1 Then
                    fmsgboxStatus "Do not enter same country twice in Schedule FSI" ', vbOKOnly, "ITR 3"
                    Target.value = "(Select)"
                End If
            End If
    End If
    
   
    
endline:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Deactivate()
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
      
    sourceSheet.Range("TR_Country").EntireRow.Hidden = True
    sourceSheet.Range("TR_TIN").value = ""
    sourceSheet.Range("TR_TaxPaidOutsideIndia").value = 0
    sourceSheet.Range("TR_TaxReliefOutsideIndia").value = 0

    'FSI Range
    Dim IntCells, mIntCtr As Long
    IntCells = Sheet26.Range("FSI_CountryCode").count
    Set rangecells = Sheet26.Range("FSI_CountryCode").Cells
    
    For IntCtr = 1 To IntCells
        countrycd = rangecells.item(IntCtr).value
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
    end_TR = 249
        
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
        
    'Country & TIN
    Set rangecellsfsi_c = Sheet26.Range("FSI_CountryCode").Cells
    Set rangecellsfsi_tin = Sheet26.Range("FSI_TaxIdentificationNo").Cells
    
    'Salary
    Set rangecellsfsi_txno1 = Sheet26.Range("FSI_TXNOSAL").Cells
    Set rangecellsfsi_txr1 = Sheet26.Range("FSI_TXRSAL").Cells
    
    'House Property
    Set rangecellsfsi_txno2 = Sheet26.Range("FSI_TXNOHP").Cells
    Set rangecellsfsi_txr2 = Sheet26.Range("FSI_TXRHP").Cells
    
    'Capital Gain
    Set rangecellsfsi_txno3 = Sheet26.Range("FSI_TXNOCapGain").Cells
    Set rangecellsfsi_txr3 = Sheet26.Range("FSI_TXRCapGain").Cells
    
    'Other Source
    Set rangecellsfsi_txno4 = Sheet26.Range("FSI_TXNOOthSrc").Cells
    Set rangecellsfsi_txr4 = Sheet26.Range("FSI_TXROthSrc").Cells
    
    'Business
    Set rangecellsfsi_txno5 = Sheet26.Range("FSI_IncFromBP").Cells
    Set rangecellsfsi_txr5 = Sheet26.Range("FSI_TXNOBP").Cells
        
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
        Country_FSI(i) = rangecellsfsi_c.item(i).value
        TaxIdentificationNo_FSI(i) = rangecellsfsi_tin.item(i).value
        
        'TXNOSAL_FSI(i) = rangecellsfsi_txno1.Item(i).Value
        'TXRSAL_FSI(i) = rangecellsfsi_txr1.Item(i).Value
        
        TXNOHP_FSI(i) = rangecellsfsi_txno2.item(i).value
        TXRHP_FSI(i) = rangecellsfsi_txr2.item(i).value
        
        TXNOCapGain_FSI(i) = rangecellsfsi_txno3.item(i).value
        TXRCapGain_FSI(i) = rangecellsfsi_txr3.item(i).value
        
        TXNOOthSrc_FSI(i) = rangecellsfsi_txno4.item(i).value
        TXROthSrc_FSI(i) = rangecellsfsi_txr4.item(i).value
        
        TXNOBusiness_FSI(i) = rangecellsfsi_txno5.item(i).value
        TXRBusiness_FSI(i) = rangecellsfsi_txr5.item(i).value
        
        'FSI_TaxIdentificationNo
        For j = 1 To end_TR
            Country_TR(j) = rangecellstr_c.item(j).value
            
            ' FSI_TXRSAL ETC'
            If (Country_FSI(i) = Country_TR(j)) Then
            
            TIN_TR(j) = TaxIdentificationNo_FSI(i)
            'TaxPaidOutsideIndia_TR(j) = TXNOHP_FSI(i) + TXNOCapGain_FSI(i) + TXNOOthSrc_FSI(i) + TXNOBusiness_FSI(i)
            'TaxReliefOutsideIndia_TR(j) = TXRHP_FSI(i) + TXRCapGain_FSI(i) + TXROthSrc_FSI(i) + TXRBusiness_FSI(i)
            TaxPaidOutsideIndia_TR(j) = Sheet26.Range("FSI_TXNOT").item(i)
            TaxReliefOutsideIndia_TR(j) = Sheet26.Range("FSI_TXRT").item(i)
            
            rangecellstr_tin(j).value = TIN_TR(j)
            rangecellstr_txno(j).value = TaxPaidOutsideIndia_TR(j)
            rangecellstr_txr(j).value = TaxReliefOutsideIndia_TR(j)
            rangecellstr_tin(j).EntireRow.Hidden = False
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
    Cancel = True
End Sub

