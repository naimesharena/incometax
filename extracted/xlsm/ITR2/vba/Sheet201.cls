Attribute VB_Name = "Sheet201"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Sub UNLOCKRANGE_General_1398A(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = rangenamestring
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet201.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                    chcell.value = ""
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    chcell.value = ""
                End If
            Next
        Next
End Sub

Sub temp()
SheetALL.Protect Password:=getmsgstate
ThisWorkbook.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub
Sub NameRanges1398A()

  Dim Rng As Range

  Set Rng = Sheet201.Range("E7")
  ThisWorkbook.Names.add name:="U_PAN", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("T7")
  ThisWorkbook.Names.add name:="U_Name", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("E9")
  ThisWorkbook.Names.add name:="U_AadhaarCardNo", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("Y9")
  ThisWorkbook.Names.add name:="U_AadhaarEnrolmentId", RefersTo:=Rng

  Set Rng = Sheet201.Range("AP9")
  ThisWorkbook.Names.add name:="U_AssessmentYear", RefersTo:=Rng

  Set Rng = Sheet201.Range("AS10")
  ThisWorkbook.Names.add name:="U_PreviouslyFiledForThisAY", RefersTo:=Rng

  Set Rng = Sheet201.Range("Z11")
  ThisWorkbook.Names.add name:="U_PreviouslyFiledForThisAY_139_8A", RefersTo:=Rng

  Set Rng = Sheet201.Range("Z13")
  ThisWorkbook.Names.add name:="U_ITRForm", RefersTo:=Rng

  Set Rng = Sheet201.Range("P14")
  ThisWorkbook.Names.add name:="U_AcknowledgementNo", RefersTo:=Rng

  Set Rng = Sheet201.Range("AN14")
  ThisWorkbook.Names.add name:="U_OrigRetFiledDate", RefersTo:=Rng

  Set Rng = Sheet201.Range("AS15")
  ThisWorkbook.Names.add name:="U_LaidOutIn_139_8A", RefersTo:=Rng

  Set Rng = Sheet201.Range("AS16")
  ThisWorkbook.Names.add name:="U_ITRFormUpdatingInc", RefersTo:=Rng

  Set Rng = Sheet201.Range("E18:AM19")
  ThisWorkbook.Names.add name:="U_ReasonsForUpdatingIncome", RefersTo:=Rng

  Set Rng = Sheet201.Range("AM21")
  ThisWorkbook.Names.add name:="U_UpdatedReturnDuringPeriod", RefersTo:=Rng

  Set Rng = Sheet201.Range("AS22")
  ThisWorkbook.Names.add name:="U_UnabsorbedDepreciation", RefersTo:=Rng

  Set Rng = Sheet201.Range("F25:K26")
  ThisWorkbook.Names.add name:="U_UnabsorbedDepreciationYear", RefersTo:=Rng

  Set Rng = Sheet201.Range("L25:Z26")
  ThisWorkbook.Names.add name:="U_RevisedReturnFile", RefersTo:=Rng

  Set Rng = Sheet201.Range("AA25:AM26")
  ThisWorkbook.Names.add name:="U_UpdatedReturnFile", RefersTo:=Rng

  Set Rng = Sheet201.Range("AN25:AR26")
  ThisWorkbook.Names.add name:="U_ReturnFiled", RefersTo:=Rng
  '---
  Set Rng = Sheet201.Range("AZ17")
  ThisWorkbook.Names.add name:="Reason1", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("BA17")
  ThisWorkbook.Names.add name:="Reason2", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("BB17")
  ThisWorkbook.Names.add name:="Reason3", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("AZ23")
  ThisWorkbook.Names.add name:="U_UnabsorbedDepreciationYear_1", RefersTo:=Rng
  
  Set Rng = Sheet201.Range("BA23")
  ThisWorkbook.Names.add name:="U_UnabsorbedDepreciationYear_2", RefersTo:=Rng

End Sub
Sub NameRangesBATI()

  Dim Rng As Range

  Set Rng = Sheet202.Range("P5")
  ThisWorkbook.Names.add name:="U_Salaries", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("P6")
  ThisWorkbook.Names.add name:="U_IncomeFromHP", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("P7")
  ThisWorkbook.Names.add name:="U_IncomeFromBP", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("P8")
  ThisWorkbook.Names.add name:="U_IncomeFromCG", RefersTo:=Rng

  Set Rng = Sheet202.Range("P9")
  ThisWorkbook.Names.add name:="U_IncomeFromOS", RefersTo:=Rng

  Set Rng = Sheet202.Range("P10")
  ThisWorkbook.Names.add name:="U_Total", RefersTo:=Rng

  Set Rng = Sheet202.Range("P11")
  ThisWorkbook.Names.add name:="U_LatestTotInc", RefersTo:=Rng

  Set Rng = Sheet202.Range("P12")
  ThisWorkbook.Names.add name:="U_UpdatedTotInc", RefersTo:=Rng

  Set Rng = Sheet202.Range("P13")
  ThisWorkbook.Names.add name:="U_AmtPayable", RefersTo:=Rng

  Set Rng = Sheet202.Range("P14")
  ThisWorkbook.Names.add name:="U_AmtRefundable", RefersTo:=Rng

  Set Rng = Sheet202.Range("P15")
  ThisWorkbook.Names.add name:="U_LastAmtPayable", RefersTo:=Rng

  Set Rng = Sheet202.Range("P16")
  ThisWorkbook.Names.add name:="U_Refund", RefersTo:=Rng

  Set Rng = Sheet202.Range("P17")
  ThisWorkbook.Names.add name:="U_TotRefund", RefersTo:=Rng

  Set Rng = Sheet202.Range("P18")
  ThisWorkbook.Names.add name:="U_FeeIncUS234F", RefersTo:=Rng

  Set Rng = Sheet202.Range("P19")
  ThisWorkbook.Names.add name:="U_RegAssessementTAX", RefersTo:=Rng

  Set Rng = Sheet202.Range("P21")
  ThisWorkbook.Names.add name:="U_AggrLiabilityRefund", RefersTo:=Rng

  Set Rng = Sheet202.Range("P22")
  ThisWorkbook.Names.add name:="U_AggrLiabilityNoRefund", RefersTo:=Rng

  Set Rng = Sheet202.Range("P23")
  ThisWorkbook.Names.add name:="U_AddtnlIncTax", RefersTo:=Rng

  Set Rng = Sheet202.Range("P24")
  ThisWorkbook.Names.add name:="U_NetPayable", RefersTo:=Rng
  '---
  Set Rng = Sheet202.Range("P25")
  ThisWorkbook.Names.add name:="U_TaxUS140B", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("P26")
  ThisWorkbook.Names.add name:="U_TaxDue10_11", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("F31:F32")
  ThisWorkbook.Names.add name:="U_BSRCode1", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("G31:G32")
  ThisWorkbook.Names.add name:="U_DateDep1", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("H31:J32")
  ThisWorkbook.Names.add name:="U_SrlNoOfChaln1", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("K31:K32")
  ThisWorkbook.Names.add name:="U_Amt1", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("K33")
  ThisWorkbook.Names.add name:="U_Total1", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("F39:F40")
  ThisWorkbook.Names.add name:="U_BSRCode2", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("G39:G40")
  ThisWorkbook.Names.add name:="U_DateDep2", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("H39:J40")
  ThisWorkbook.Names.add name:="U_SrlNoOfChaln2", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("K39:K40")
  ThisWorkbook.Names.add name:="U_Amt2", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("K41")
  ThisWorkbook.Names.add name:="U_Total2", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("P44")
  ThisWorkbook.Names.add name:="U_ReleifUS89", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("S17")
  ThisWorkbook.Names.add name:="Vmonth", RefersTo:=Rng
  
  Set Rng = Sheet202.Range("T17")
  ThisWorkbook.Names.add name:="Vyear", RefersTo:=Rng
  
  

End Sub
Sub LOCKRANGE_General_1398A(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet201.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub
Sub LockUnlockBATI()
    Application.EnableEvents = False
    Sheet202.Unprotect Password:=getmsgstate
    Sheet201.Unprotect Password:=getmsgstate
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" Then
        Sheet202.Range("U_LatestTotInc").MergeArea.ClearContents
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = True
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LatestTotInc").value = ""
        
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.ClearContents
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Locked = True
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = ""
        
        Sheet201.Range("U_ITRForm").MergeArea.ClearContents
        Sheet201.Range("U_ITRForm").MergeArea.Locked = True
        Sheet201.Range("U_ITRForm").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_ITRForm").value = ""
        
        Sheet201.Range("U_AcknowledgementNo").MergeArea.ClearContents
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Locked = True
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_AcknowledgementNo").value = ""
        
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.ClearContents
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Locked = True
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_OrigRetFiledDate").value = ""
        'Shrutika_Ur
        LOCKRANGEIT2
        
    Else
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = False
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Locked = False
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "(Select)"
        
        Sheet201.Range("U_ITRForm").MergeArea.Locked = False
        Sheet201.Range("U_ITRForm").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet201.Range("U_ITRForm").value = "(Select)"
        
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Locked = False
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Locked = False
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Interior.Color = "&HCCFFCC"
        
        'Shrutika_Ur
      UNLOCKRANGEIT2
    End If
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "(Select)" Then
        Sheet202.Range("U_LatestTotInc").MergeArea.ClearContents
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = True
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LatestTotInc").value = ""
    End If
    Application.EnableEvents = True
    Sheet202.Protect Password:=getmsgstate
    Sheet201.Protect Password:=getmsgstate
End Sub
Sub LockUnlockGenYear()
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
    
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value <> "1" Then
'
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_RevisedReturnFile").Value = "(Select)"
'
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_UpdatedReturnFile").Value = ""
'
'    ElseIf Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value = "1" Then
'
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_UpdatedReturnFile").Value = "(Select)"
'
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_RevisedReturnFile").Value = ""
'    Else
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_RevisedReturnFile").Value = ""
'
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_UpdatedReturnFile").Value = ""
'    End If
    
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value >= "1" Then
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_UpdatedReturnFile").Value = "(Select)"
'    End If
'
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_2").Value >= "1" Then
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_RevisedReturnFile").Value = "(Select)"
'    End If
    
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value = "0" Then
'        Sheet201.Range("U_UpdatedReturnFile").ClearContents
'        Sheet201.Range("U_UpdatedReturnFile").Locked = True
'        Sheet201.Range("U_UpdatedReturnFile").Interior.Color = "&HD8D8D8"
'    End If
'
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_2").Value = "0" Then
'        Sheet201.Range("U_RevisedReturnFile").ClearContents
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_RevisedReturnFile").Value = ""
'    End If
    
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub
Sub LockUnlockGenReason()
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
    
    If Sheet201.Range("Reason1").value = 1 Or Sheet201.Range("Reason2").value = 1 Or Sheet201.Range("Reason3").value = 1 Then
        
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Locked = False
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Interior.Color = "&HCCFFCC"
        'Sheet201.Range("U_UnabsorbedDepreciation").Value = "Yes"
        'Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Locked = True
    Else
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.ClearContents
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Locked = True
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Interior.Color = "&HD8D8D8"
                   
        Sheet201.Range("U_RevisedReturnFile").ClearContents
        Sheet201.Range("U_RevisedReturnFile").Locked = True
        Sheet201.Range("U_RevisedReturnFile").Interior.Color = "&HD8D8D8"

        
        Sheet201.Range("U_UpdatedReturnFile").ClearContents
        Sheet201.Range("U_UpdatedReturnFile").Locked = True
        Sheet201.Range("U_UpdatedReturnFile").Interior.Color = "&HD8D8D8"
        
        Sheet201.Range("U_UnabsorbedDepreciationYear").ClearContents
        Sheet201.Range("U_UnabsorbedDepreciationYear").Locked = True
        Sheet201.Range("U_UnabsorbedDepreciationYear").Interior.Color = "&HD8D8D8"

        
    End If
    
            
    
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub
Function chkNumeric_1398A(field As Variant) As Boolean
    Dim k As Long
    Dim chkchar As String

    chkNumeric_1398A = True
    For k = 1 To Len(field)
    chkchar = Mid(field, k, 1)
        If Not IsNumeric(chkchar) Then
            chkNumeric_1398A = False
            Exit Function
        End If
    Next
End Function

'Sub ReturnDropdownA12(RowIndex As Long)
'On Error GoTo endline
'Application.EnableEvents = False
'Sheet201.Unprotect Password:=getmsgstate
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2023-24" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'End If
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2024-25" Then
'    If ChkMinInclusiveDate(Dformat("31/12/2024", ""), Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'    End If
'End If
'endline:
'Sheet201.Protect Password:=getmsgstate
'    Application.EnableEvents = True
'End Sub

'Change-25.11.2022.104.08B
'Sub ReturnDropdownA12(RowIndex As Long)
'On Error GoTo endline
'Application.EnableEvents = False
'Sheet201.Unprotect Password:=getmsgstate
'
'
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2023-24" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2024", "") <= Dformat(Sheet52.Range("sheet9.Date").value, "")) And (Dformat("31/03/2024", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2023", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
'
'
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2024-25" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2025", "") <= Dformat(Sheet52.Range("sheet9.Date").value, "")) And (Dformat("31/03/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2024", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
'endline:
'Sheet201.Protect Password:=getmsgstate
'    Application.EnableEvents = True
'End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

'Sheet1.Range("taxcheck").value = ""

CalculateTaxCheckFlag = False
Application.EnableEvents = False



'If Not Application.Intersect(Sheet201.Range("U_AadhaarCardNo"), Sheet201.Range(Target.Address)) Is Nothing Then
'    If Not chkNumeric(Target.Value) Then
'        MsgBox "Please enter valid 12 digit aadhar number"
'        Target.Value = ""
'    End If
'End If

'Change-25.11.2022.104.08A
'ReturnDropdownA12 (Target.row)
ReturnDropdownA23_24 (Target.row)

'Shrutika_UR
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
 
Dim dd1 As String
Dim dd2 As String

'UR_Shrutika (11/07/2025)
'---------------------------------------
dd1 = "2026-03-31"
dd2 = "2027-03-31"
 
If Target.value = "2026-27" Then
If Dformat(Sheet52.Range("sheet9.Date"), "") <= dd1 Then
MsgBox "AY ""2026-27"" can be selected after 31/03/2026", vbOKOnly, "Alert"
Target.value = "(Select)"
End If
Else
If Target.value = "2027-28" Then
If Dformat(Sheet52.Range("sheet9.Date"), "") <= dd2 Then
MsgBox "AY ""2027-28"" can be selected after 31/03/2027", vbOKOnly, "Alert"
Target.value = "(Select)"
End If
End If
End If
End If
'-----------------------------------------------
 
'dd1 = "2025-03-31"
'dd2 = "2026-03-31"

 
'If Target.value = "2025-26" Then
'If Dformat(Sheet52.Range("sheet9.Date"), "") <= dd1 Then
'MsgBox "AY ""2025-26"" can be selected on or after 01/04/2025", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
'
'
'Else
'
'If Target.value = "2026-27" Then
'If Dformat(Sheet52.Range("sheet9.Date"), "") <= dd2 Then
'MsgBox "AY ""2026-27"" can be selected on or after 01/04/2026", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
'End If
'End If
'End If

'----------------------------------
If Not Application.Intersect(Sheet201.Range("U_PreviouslyFiledForThisAY"), Sheet201.Range(Target.Address)) Is Nothing Then
LockUnlockBATI
End If


'Chnage-25.11.2022.104.05
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
    PartA_139LockUnlockTable Target
'LockUnlockGenYear
'    Application.EnableEvents = False
'    Sheet201.Unprotect Password:=getmsgstate
'    If Target.value = "2023-24" Or Target.value = "2024-25" Then
'
'
''    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'
'    '    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = False
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'
''    ElseIf Target.Value = "2024" Then
''
'''    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
''    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = False
''    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
''
''    '    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
''    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'
'    Else
'
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = True
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    End If
'    Application.EnableEvents = True
'    Sheet201.Protect Password:=getmsgstate
'    Application.EnableEvents = False
End If

If Not Application.Intersect(Sheet201.Range("U_ReasonsForUpdatingIncome"), Sheet201.Range(Target.Address)) Is Nothing Then
    If WorksheetFunction.CountIf(Sheet201.Range("U_ReasonsForUpdatingIncome"), Target.value) > 1 Then
        Application.EnableEvents = False
            '08/11/2024  'Malli
            If Not (Target.value = "(Select)" Or Target.value = "") Then
                   MsgBox "You cannot select the same reason for updating your income twice"
            'Target.value = ""
                   Target.value = "(Select)"
            End If
           '-----------------------
        Application.EnableEvents = True
    End If
LockUnlockGenReason
End If

If Not Application.Intersect(Sheet201.Range("U_AcknowledgementNo"), Sheet201.Range(Target.Address)) Is Nothing Then
Application.EnableEvents = False
Sheet1.Range("sheet1.ReceiptNo").value = Sheet201.Range("U_AcknowledgementNo").value
Application.EnableEvents = True
End If

If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
Application.EnableEvents = False
Sheet1.Range("sheet1.OrigRetFiledDate").value = Sheet201.Range("U_OrigRetFiledDate").value
Application.EnableEvents = True
End If

If Not Application.Intersect(Sheet201.Range("U_LaidOutIn_139_8A"), Sheet201.Range(Target.Address)) Is Nothing Then
If Sheet201.Range(Target.Address).value = "No" Then
'MsgBox "Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second and third provisos to section 139(8A)"
MsgBox "Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second, third and fourth provisos to section 139(8A)" 'UR Shrutika(30/10/2025)

Sheet201.Range(Target.Address).value = "(Select)"
End If
End If

'Shrutika_UR
 
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
    If Not (Target.value = "(Select)" Or Target.value = "") Then
       If WorksheetFunction.CountIf(Sheet201.Range("U_UnabsorbedDepreciationYear"), Target.value) > 1 Then
        Application.EnableEvents = False
            'MsgBox "You cannot select the same drop down twice in Assessment Year"
            'Added by Shrutika
            MsgBox "You can not select the same dropdown for twice"
            Target.value = "(Select)"
        Application.EnableEvents = True
      End If
    End If
End If
'--------------------


''If (Target.name.name) = "U_UnabsorbedDepreciation" Then
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciation"), Sheet201.Range(Target.Address)) Is Nothing Then
    Application.EnableEvents = False
        Sheet201.Unprotect Password:=getmsgstate
    If Target.value = "Yes" Then
    
'    Sheet201.Range("U_UnabsorbedDepreciationYear").ClearContents
    
    Sheet201.Range("U_UnabsorbedDepreciationYear").Locked = False
    Sheet201.Range("U_UnabsorbedDepreciationYear").Interior.Color = "&HCCFFCC"
        
    Else
    
    Sheet201.Range("U_UnabsorbedDepreciationYear").ClearContents
    Sheet201.Range("U_UnabsorbedDepreciationYear").Locked = True
    Sheet201.Range("U_UnabsorbedDepreciationYear").Interior.Color = "&HD8D8D8"
    
    Sheet201.Range("U_RevisedReturnFile").ClearContents
    Sheet201.Range("U_RevisedReturnFile").Locked = True
    Sheet201.Range("U_RevisedReturnFile").Interior.Color = "&HD8D8D8"
    
    Sheet201.Range("U_UpdatedReturnFile").ClearContents
    Sheet201.Range("U_UpdatedReturnFile").Locked = True
    Sheet201.Range("U_UpdatedReturnFile").Interior.Color = "&HD8D8D8"
    End If
    Application.EnableEvents = True
        Sheet201.Protect Password:=getmsgstate
End If
'
'     Dim RangeArray As Variant
'     RangeArray = "U_UnabsorbedDepreciationYear"
'     Sheet201.Unprotect Password:=getmsgstate
'        If Sheet201.Range("Reason1").Value = 1 Or Sheet201.Range("Reason2").Value = 1 Or Sheet201.Range("Reason3").Value = 1 Then
'
'        UNLOCKRANGE_General_1398A (RangeArray)
'           Else
'
'          LOCKRANGE_General_1398A (RangeArray)
'
'         End If
'       Sheet201.Protect Password:=getmsgstate
'End If

'If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
'    If Sheet201.Range("U_UnabsorbedDepreciation").Value <> "Yes" Then
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value = 1 Or Sheet201.Range("U_UnabsorbedDepreciationYear_2").Value = 1 Then
'    fmsgbox "* Year you can not select if updated return is not to reduce carried forward loss or unabsorbed depreciation or tax credit" & Chr(13)
'    Sheet201.Range(Target.Address).Value = ""
'    End If
'    End If
'End If



'If Not Application.Intersect(Sheet201.Range("U_AadhaarEnrolmentId"), Sheet201.Range(Target.Address)) Is Nothing Then
''If (Target.name.name) = "U_AadhaarEnrolmentId" Then
'    If Len(Sheet201.Range("U_AadhaarEnrolmentId").Value) >= 12 Then
'        If IsNumeric(Sheet201.Range("U_AadhaarEnrolmentId").Value) Then
'
'            If Not checkfieldSuperSpecialcharacterDot_1398A(Sheet201.Range("U_AadhaarEnrolmentId").Value) Then
'                fmsgbox "* Please enter the valid Aadhaar Enrolment ID in Sheet Income Details"
'                Sheet1.Range("Sheet1.AadhaarEnrol").Value = ""
'            End If
'
'            If Sheet201.Range("U_AadhaarEnrolmentId").Value <> "0000000000000000000000000000" Then
'              Sheet201.Range("U_AadhaarEnrolmentId").Value = Sheet201.Range("U_AadhaarEnrolmentId").Value
'            Else
'              fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'              Sheet1.Range("U_AadhaarEnrolmentId").Value = ""
'            End If
'
'
'            If Sheet201.Range("U_AadhaarEnrolmentId").Value < 0 Then
'                fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'            End If
'
'            If Sheet201.Range("U_AadhaarEnrolmentId").Value <> "1111111111111111111111111111" Then
'                Sheet201.Range("U_AadhaarEnrolmentId").Value = Sheet201.Range("U_AadhaarEnrolmentId").Value
'            Else
'                fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'            End If
'
'        Else
'            fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'            Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'        End If
'    Else
'        Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'    End If
'End If
If (Target.name.name) = "U_AcknowledgementNo" Then
    Dim urecieptno As Variant
    Application.EnableEvents = False
    
    urecieptno = Sheet201.Range("U_AcknowledgementNo").value
    
    If Not chkNumeric_1398A(urecieptno) Then
        fmsgbox "* Invalid Acknowledgement Number. Please retry "
        Sheet201.Range("U_AcknowledgementNo").value = ""
        Sheet201.Range("U_AcknowledgementNo").Select
    End If
    
    If Sheet201.Range("U_AcknowledgementNo").value <> "000000000000000" Then
        Sheet201.Range("U_AcknowledgementNo").value = Sheet201.Range("U_AcknowledgementNo").value
    Else
        fmsgbox "* Invalid Acknowledgement Number. Please retry "
        Sheet201.Range("U_AcknowledgementNo").value = ""
        Sheet201.Range("U_AcknowledgementNo").Select
    
    End If
    Application.EnableEvents = True
End If

'Added by Shrutika on 25th Nov 24
 
If Not Application.Intersect(Sheet201.Range("U_PreviouslyFiledForThisAY"), Sheet201.Range(Target.Address)) Is Nothing Then
    If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
          Application.EnableEvents = False
        Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
        Sheet1.Range("sheet1.ReceiptNo").value = ""
         Application.EnableEvents = True
    End If
End If
 
'-------------- UR_Shrutika(04/11/2025)

'Shrutika_Ur
'If Not Application.Intersect(Sheet201.Range("U_UpdatedReturnDuringPeriod"), Sheet201.Range(Target.Address)) Is Nothing Then
'
'Dim dd3 As String
'Dim dd4 As String
'
'dd3 = "2026-04-01"
'dd4 = "2026-03-31"
'
''Commented by Shrutika
''If Target.value = "Up to 12 months from the end of Relevant Assessment Year" Then
''If Dformat(Sheet52.Range("sheet9.Date"), "") >= dd3 Then
''MsgBox "Please select appropriate option in 'Are you filing the updated return during the period.'", vbOKOnly, "Alert"
''Target.value = "(Select)"
''End If
'
'
''Else
'
'If Target.value = "Between 12 to 24 Months from the end of Relevant Assessment  Year" Then
'If Dformat(Sheet52.Range("sheet9.Date"), "") <= dd4 Then
'MsgBox "Please select appropriate option in 'Are you filing the updated return during the period.'", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
'End If
'End If
''End If
'-------------------------

'Malli------11/08/2024--------
 If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
         If Not CheckDateddmmyyyy(Sheet201.Range("U_OrigRetFiledDate").value) Then
             MsgBox "* ""Date of filing of Original return must be a valid dd/mm/yyyy format""", vbOKOnly, "Error(s)"
             Target.value = ""
         End If
 End If
 
'Added by Shrutika
 If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
 
            Application.EnableEvents = False
                If Trim(Range("U_OrigRetFiledDate").value) <> "00/00/0000" Then
                    If Not CheckDateBefore(Trim(Range("U_OrigRetFiledDate").value)) Then
                       Range("U_OrigRetFiledDate").Select
                       Range("U_OrigRetFiledDate").value = ""
                    End If
                    If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
                        MsgBox ("Date of filing of original return cannot be after sytem date."), "Error"
                        Target.value = ""
                        Target.Select
                    End If
                 End If
            Application.EnableEvents = True
           End If
 
 '---------------------------------

endline:
Application.EnableEvents = True

End Sub

'Added by Shruika
Function CheckDateBefore(dob As Variant) As Boolean
On Error Resume Next
 
 
    CheckDateBefore = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBefore = False
           MsgBox ("Date of Filing Original Return in Sheet : PART A-GEN_139(8A)  must be a valid dd/mm/yyyy format"), , "Error"
            Exit Function
        End If
        dob = Dformat(dob, "")
 
        'If Not ChkMinInclusiveDate(dob, "2024-04-01") Then 'UR Added by Shrutika(31/10/2025)
        If Not ChkMinInclusiveDate(dob, "2025-04-01") Then
            'MsgBox ("Date of filing of original return cannot be prior to 01/04/2024"), , "Error"
            MsgBox ("Date of filing of original return cannot be prior to 01/04/2025"), , "Error"  'UR Added by Shrutika(31/10/2025)
            CheckDateBefore = False
            Exit Function
        Else
       End If
    End If
End Function






