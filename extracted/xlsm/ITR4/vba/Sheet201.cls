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
        
        'Ayush
        Sheet1.Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
        Sheet1.Range("sheet1.ReceiptNo").MergeArea.ClearContents
        
        LOCKRANGEIT2  'Malli------------11/10/2024
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
        UNLOCKRANGEIT2             'Malli-------------------11/10/2024
    End If
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "Yes" Then
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = False
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HCCFFCC"
    Else
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


Private Sub Worksheet_Activate()
 Sheet201.Protect Password:=getmsgstate

End Sub

Private Sub Worksheet_Calculate()

End Sub

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
'    If ChkMinInclusiveDate(Dformat("31/12/2024", ""), Dformat(Sheet5.Range("Ver.Date").value, "")) Then
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


'Change-25.11.2022.103.08B
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
'        If (Dformat("01/01/2024", "") <= Dformat(Sheet5.Range("Ver.Date").value, "")) And (Dformat("31/03/2024", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2023", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
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
'        If (Dformat("01/01/2025", "") <= Dformat(Sheet5.Range("Ver.Date").value, "")) And (Dformat("31/03/2025", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2024", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
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


Application.EnableEvents = False

CalculateTaxCheckFlag = False

'Change-25.11.2022.103.08A
'ReturnDropdownA12 (Target.row)
'ReturnDropdownA23_24 (Target.row)

'Ayush_UR
'If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
'
'Dim dd1 As String
'Dim dd2 As String

'dd1 = "2025-03-31"
'dd2 = "2026-03-31"
'
'
'If Target.value = "2025-26" Then
'If Dformat(Sheet5.Range("Ver.Date"), "") <= dd1 Then
'MsgBox "AY ""2025-26"" can be selected on or after 01/04/2025", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
'
'
'Else
'
'If Target.value = "2026-27" Then
'If Dformat(Sheet5.Range("Ver.Date"), "") <= dd2 Then
'MsgBox "AY ""2026-27"" can be selected on or after 01/04/2026", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
'End If
'End If



'Added by Aavula Naresh for UR 2025-26

If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then

Dim DD1 As String
Dim dd2 As String

DD1 = "2026-03-31"
dd2 = "2027-03-31"

If Target.value = "2026-27" Then
If Dformat(Sheet5.Range("Ver.Date"), "") <= DD1 Then

MsgBox "AY ""2026-27"" can be selected after 31/03/2026", vbOKOnly, "Alert"

Target.value = "(Select)"
End If


Else

If Target.value = "2027-28" Then
If Dformat(Sheet5.Range("Ver.Date"), "") <= dd2 Then
MsgBox "AY ""2027-28"" can be selected after 31/03/2027", vbOKOnly, "Alert"

Target.value = "(Select)"
End If
End If
End If
End If
'-----------------------------------------


'Ayush_UR==============================

'If Not Application.Intersect(Sheet201.Range("U_UpdatedReturnDuringPeriod"), Sheet201.Range(Target.Address)) Is Nothing Then
'
'Dim dd3 As String
'Dim dd4 As String
'
''dd3 = "2026-04-01"
'dd4 = "2026-03-31"
 
 
'If Target.value = "Up to 12 months from the end of Relevant Assessment Year" Then
'If Dformat(Sheet5.Range("Ver.Date"), "") >= dd3 Then
'MsgBox "Please select appropriate option in 'Are you filing the updated return during the period.'", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
 
 

 
'If Target.value = "Between 12 to 24 Months from the end of Relevant Assessment  Year" Then
'If Dformat(Sheet5.Range("Ver.Date"), "") <= dd4 Then
'MsgBox "Please select appropriate option in 'Are you filing the updated return during the period.'", vbOKOnly, "Alert"
'Target.value = "(Select)"
'End If
'End If

'End If
'==============================

If Not Application.Intersect(Sheet201.Range("U_AadhaarCardNo"), Sheet201.Range(Target.Address)) Is Nothing Then
    If Not chkNumeric(Target.value) Then
        MsgBox "Please enter valid 12 digit Aadhaar number"
        Target.value = ""
    End If
End If

If Not Application.Intersect(Sheet201.Range("U_PreviouslyFiledForThisAY"), Sheet201.Range(Target.Address)) Is Nothing Then
LockUnlockBATI
End If

'Change-25.11.2022.103.05
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
    PartA_139_LockUnlockTable Target
''LockUnlockGenYear
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
'Ayush_SIT-84153
If Not (Target.value = "(Select)" Or Target.value = "") Then
    If WorksheetFunction.CountIf(Sheet201.Range("U_ReasonsForUpdatingIncome"), Target.value) > 1 Then
        Application.EnableEvents = False
            MsgBox "You cannot select the same reason for updating your income twice"
            Target.value = ""
            
        Application.EnableEvents = True
    End If
    End If
    
LockUnlockGenReason
End If

If Not Application.Intersect(Sheet201.Range("U_AcknowledgementNo"), Sheet201.Range(Target.Address)) Is Nothing Then
Application.EnableEvents = False
Sheet1.Range("sheet1.ReceiptNo").value = Sheet201.Range("U_AcknowledgementNo").value
Application.EnableEvents = True
End If

'Ayush_SIT-84153
  If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
      If Trim(Range("U_OrigRetFiledDate").value) <> "00/00/0000" Then
          If Not CheckDateddmmyyyy(Trim(Target.value)) Then
             MsgBox "Please enter valid date in dd/mm/yyyy format in date of filing of original return in Part A General 139(8A)", , "Error"
             Target.value = ""
             Else
              'Ayush_14/11/2025
                If ChkMaxDate(Trim(Target.value), "") Then
                    MsgBox "* ""Date of filing cannot be prior to 01/04/2025 in Part A General 139(8A)."""
                    Target.value = ""
                End If
      End If
    End If
  End If



If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
Sheet1.Range("sheet1.OrigRetFiledDate").value = Sheet201.Range("U_OrigRetFiledDate").value
End If

If Not Application.Intersect(Sheet201.Range("U_LaidOutIn_139_8A"), Sheet201.Range(Target.Address)) Is Nothing Then
If Sheet201.Range(Target.Address).value = "No" Then
'MsgBox "Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second and third provisos to section 139(8A)"
MsgBox "Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second, third and fourth provisos to section 139(8A)"

Sheet201.Range(Target.Address).value = "(Select)"
End If
End If


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
'Ayush_UR
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
      If UCase(Target.value) <> UCase("(Select)") Then
        If WorksheetFunction.CountIf(Sheet201.Range("U_UnabsorbedDepreciationYear"), Target.value) > 1 Then
            Application.EnableEvents = False
                MsgBox "You can not select the same dropdown for twice", , "Error"
                Target.value = ""
            Application.EnableEvents = True
        End If
      End If
End If


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


'Ayush_enrol===============================
'If Not Application.Intersect(Sheet201.Range("U_AadhaarEnrolmentId"), Sheet201.Range(Target.Address)) Is Nothing Then
''If (Target.name.name) = "U_AadhaarEnrolmentId" Then
'                Application.EnableEvents = False
'                If Len(Sheet201.Range("U_AadhaarEnrolmentId").value) >= 12 Then
'                    If IsNumeric(Sheet201.Range("U_AadhaarEnrolmentId").value) Then
'
'                    If Not checkfieldSuperSpecialcharacterDot_1398A(Sheet201.Range("U_AadhaarEnrolmentId").value) Then
'                    fmsgbox "* Please enter the valid Aadhaar Enrolment ID in Sheet Income Details"
'                    Sheet1.Range("Sheet1.AadhaarEnrol").value = ""
'                    End If
'                       If Sheet201.Range("U_AadhaarEnrolmentId").value <> "0000000000000000000000000000" Then
'                         Sheet201.Range("U_AadhaarEnrolmentId").value = Sheet201.Range("U_AadhaarEnrolmentId").value
'                       Else
'                         fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                         Sheet1.Range("U_AadhaarEnrolmentId").value = ""
'                       End If
'
'
'                       If Sheet201.Range("U_AadhaarEnrolmentId").value < 0 Then
'                       fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                         Sheet201.Range("U_AadhaarEnrolmentId").value = ""
'                         End If
'
'                       If Sheet201.Range("U_AadhaarEnrolmentId").value <> "1111111111111111111111111111" Then
'                         Sheet201.Range("U_AadhaarEnrolmentId").value = Sheet201.Range("U_AadhaarEnrolmentId").value
'                       Else
'                         fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                         Sheet201.Range("U_AadhaarEnrolmentId").value = ""
'                       End If
'
'                    Else
'                      fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                      Sheet201.Range("U_AadhaarEnrolmentId").value = ""
'                    End If
'                Else
'                    Sheet201.Range("U_AadhaarEnrolmentId").value = ""
'                End If
'                Application.EnableEvents = False
'            End If

'=======================
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

endline:
Application.EnableEvents = True

End Sub





