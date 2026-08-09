Attribute VB_Name = "AY23_24_Changes"
Option Explicit
Public msgValidateSheetTCS As String
Dim CreditDate As Variant
Public EnableEvents As Boolean
Sub StoreEventsStatus()
EnableEvents = Application.EnableEvents
End Sub
Sub RestoreEventsStatus()
    Application.EnableEvents = True
End Sub
Sub ab()
If IsNumeric(12) Then
MsgBox "HI"
End If
End Sub

Sub b()
Application.EnableEvents = False
End Sub
Function ChkMaxDOBDate23_24(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate23_24 = True
     If Len(dob) > 0 Then
'        If val(Mid(dob, 7, 4)) > 2023 Then
'         If val(Mid(dob, 7, 4)) > 2024 Then 'Year changed by Sai

'         If val(Mid(dob, 7, 4)) > 2025 Then 'Year changed by Aavula Naresh
         
         If val(Mid(dob, 7, 4)) > 2026 Then 'Year changed by Aavula Naresh for AY 2026-27


                    ChkMaxDOBDate23_24 = False
                    Exit Function
'         ElseIf val(Mid(dob, 7, 4)) = 2023 Then
'         ElseIf val(Mid(dob, 7, 4)) = 2024 Then 'Year changed by Sai
'          ElseIf val(Mid(dob, 7, 4)) = 2025 Then 'Year changed by Aavula Naresh
          
          ElseIf val(Mid(dob, 7, 4)) = 2026 Then 'Year changed by Aavula Naresh for AY 2026-27
         
            If val(Mid(dob, 4, 2)) >= 4 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate23_24 = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function

Sub ReturnDropdownA12_AY23_24(RowIndex As Long)
On Error GoTo endline
Application.EnableEvents = False
Sheet201.Unprotect Password:=getmsgstate



'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2024-25" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
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
'
'
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2026", "") <= Dformat(Sheet52.Range("sheet9.Date").value, "")) And (Dformat("31/03/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If

'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2026", "") < Dformat(Sheet52.Range("sheet9.Date").value, "")) Or (Dformat("31/03/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2026-27" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2027", "") < Dformat(Sheet52.Range("sheet9.Date").value, "")) Or (Dformat("31/03/2027", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If

'Konda updated on 12-11-2025
If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2026-27" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
    
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
        If (Dformat("31/12/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
        ElseIf (Dformat("01/01/2026", "") < Dformat(Sheet52.Range("sheet9.Date").value, "")) Or (Dformat("31/03/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
        End If
    'End If
End If

If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2026-27" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
    
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
        If (Dformat("31/12/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
        ElseIf (Dformat("01/01/2027", "") < Dformat(Sheet52.Range("sheet9.Date").value, "")) Or (Dformat("31/03/2027", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
        End If
   ' End If
End If


endline:
Sheet201.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub

Sub Pan_Other_Buyer(RowIndex As Long)
    If Not Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PanOtherPerson").Column).value = Empty Then
    If UCase(Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PanOtherPerson").Column).value) = UCase(Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PAN").Column).value) Then
    MsgBox "Pan of buyer and pan of the other person can't be same."
    Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PanOtherPerson").Column).value = ""
    Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PanOtherPerson").Column).Select
    End If
    End If


    
End Sub



Sub Pan_Other_Buyer2(RowIndex As Long)

    If Not Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PAN").Column).value = Empty Then
    If UCase(Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PAN").Column).value) = UCase(Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PanOtherPerson").Column).value) Then
    MsgBox "Pan of buyer and pan of the other person can't be same."
    Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PAN").Column).value = ""
    Sheet28.Cells(RowIndex, Sheet28.Range("TDS3.PAN").Column).Select
    End If
    End If


    
End Sub
'=============================================================================================================
'Sub reset10IEAy23_24()
'On Error Resume Next
'SheetALL.Unprotect Password:=getmsgstate
'Sheet1.Unprotect Password:=getmsgstate
'
'    If SheetALL.Range("OldBacValue").value = 2 Or SheetALL.Range("OldBacValue").value = "" Then
'
'        Sheet1.Range("AYinOpting").MergeArea.Locked = True
'        Sheet1.Range("AYinOpting").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("AYinOpting").value = ""
'
'        Sheet1.Range("Sheet1.date10IE_Opt").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.date10IE_Opt").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = ""
'
'        Sheet1.Range("Opt_Out115BAC").MergeArea.Locked = True
'        Sheet1.Range("Opt_Out115BAC").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Opt_Out115BAC").value = ""
'
'        Sheet1.Range("AYinOptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("AYinOptOut").MergeArea.Locked = True
'        Sheet1.Range("AYinOptOut").value = ""
'
'        Sheet1.Range("Sheet1.date10IE_OptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.date10IE_OptOut").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = ""
'
'    Else
'
'        Sheet1.Range("AYinOpting").value = "(Select)"
'        Sheet1.Range("AYinOpting").MergeArea.Locked = False
'        Sheet1.Range("AYinOpting").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'        Sheet1.Range("Sheet1.date10IE_Opt").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.date10IE_Opt").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Opt_Out115BAC").value = "(Select)"
'        Sheet1.Range("Opt_Out115BAC").MergeArea.Locked = False
'        Sheet1.Range("Opt_Out115BAC").Interior.Color = (&HCCFFCC)
'     End If
'
'    SheetALL.Protect Password:=getmsgstate
'    Sheet1.Protect Password:=getmsgstate
'End Sub

'Sub CurrentYear10IE_23_24(ByVal Target As Range)
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'    If Sheet1.Range("Opt_Out115BAC").value = "Yes" Then
'
'        Sheet1.Range("AYinOptOut").value = "(Select)"
'        Sheet1.Range("AYinOptOut").MergeArea.Locked = False
'        Sheet1.Range("AYinOptOut").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'        Sheet1.Range("Sheet1.date10IE_OptOut").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.date10IE_OptOut").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Interior.Color = (&HCCFFCC)
'
'    Else
'
'        Sheet1.Range("AYinOptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("AYinOptOut").MergeArea.Locked = True
'        Sheet1.Range("AYinOptOut").value = ""
'
'        Sheet1.Range("Sheet1.date10IE_OptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.date10IE_OptOut").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = ""
'
'    End If
'
'CurrentYearOptionAY23_24
'Sheet1.Protect Password:=getmsgstate
''Application.EnableEvents = True
'
'End Sub
'=====================================================================================================================
'Sub CurrentYearOptionAY23_24()
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
''Change-03.11.2023.102.Ver.0.6
'    If SheetALL.Range("OldBacValue").value = 2 And Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" And Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" Then
'
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1without_OptIn"
'        Application.EnableEvents = True
'        Sheet1.Range("OptingNewTaxRegime").value = "Not opting"
'        Application.EnableEvents = False
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
''        If SheetALL.Range("OldBacValue").value = 2 Then
'    ElseIf SheetALL.Range("OldBacValue").value = 2 Then
''End Change
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'
'    End If
'
'    If SheetALL.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "Yes" Then
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current3"
'        Application.EnableEvents = True
'        Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in"
'        Application.EnableEvents = False
''Change-ANK
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.capacityassessee").value = ""
'        Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'
'        Sheet1.Range("acknowledgmentnumber").value = ""
'        Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
'
'        'Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HD8D8D8)
'        'Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = True
''---
''Change-03.11.2023.102.Ver.0.6
'    ElseIf Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" And Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" Then
'        If Not Sheet1.Range("OptingNewTaxRegime").value = "Not opting" Then
'
'            Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'            Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1without_OptIn"
'            Application.EnableEvents = True
'            Sheet1.Range("OptingNewTaxRegime").value = "Not opting"
'            Application.EnableEvents = False
''            Sheet1.Unprotect Password:=getmsgstate
''            Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
''            Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
''
''            Sheet1.Range("sheet1.capacityassessee").value = ""
''            Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False
''            Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)
''
''            Sheet1.Range("acknowledgmentnumber").value = ""
''            Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
''            Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)
'        End If
''End Change
'    Else
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'
'        Sheet1.Range("sheet1.capacityassessee").value = ""
'        Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False
'        Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("acknowledgmentnumber").value = ""
'        Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
'        Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)
'    End If
'
''Change-03.11.2023.102.Ver.0.6
'    If SheetALL.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "No" And Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" And Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" Then
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current2without_OptOUt"
'        Application.EnableEvents = True
'        Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt"
'        Application.EnableEvents = False
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
''        If SheetALL.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "No" Then
'    ElseIf SheetALL.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "No" Then
''End Change
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current2"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'    End If
'
'Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True
'End Sub

'============================================================================================================================
'Sub CurrentYearOptionAY23_24()
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'    If SheetALL.Range("OldBacValue").value = 2 Then
'
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'
'    End If
'
'    If SheetALL.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "Yes" Then
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current3"
'        Application.EnableEvents = True
'        Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in"
'        Application.EnableEvents = False
''Change-ANK
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.capacityassessee").value = ""
'        Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'
'        Sheet1.Range("acknowledgmentnumber").value = ""
'        Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
'
'        'Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HD8D8D8)
'        'Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = True
''---
'    Else
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'
'        Sheet1.Range("sheet1.capacityassessee").value = ""
'        Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False
'        Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("acknowledgmentnumber").value = ""
'        Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
'        Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)
'    End If
'
'    If SheetALL.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "No" Then
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current2"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'    End If
'
'Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True
'End Sub
'==================================================================================================================

'Commented by Sai For AY 2023-24 Old
'Sub CurrentYear10IEDateAckAY_23_24(ByVal Target As Range)
'
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'If (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
''Change-ANK
'        Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = False
'        Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("sheet1.capacityassessee").value = ""
'
'        Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = False
'        Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("acknowledgmentnumber").value = ""
'
'Else
'        Sheet1.Range("sheet1.capacityassessee").value = ""
'        Sheet1.Range("sheet1.capacityassessee").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("sheet1.capacityassessee").MergeArea.Locked = True
'
'        Sheet1.Range("acknowledgmentnumber").value = ""
'        Sheet1.Range("acknowledgmentnumber").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("acknowledgmentnumber").MergeArea.Locked = True
''---end
'
'
'End If
'
'Sheet1.Protect Password:=getmsgstate
''Application.EnableEvents = True
'
'End Sub
'========================================End By Sai=============================================================



'===========================Commented by Aavula Naresh for AY 2025-26=======================================''''
'Newly added by Sai For AY 2024-25
'Sub CurrentYear10IEDateAckAY_24_25(ByVal target As Range)
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'If (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date") Or (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date") Then
'
'        Sheet1.Range("Sheet1.Form10IEADate").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.Form10IEADate").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.Form10IEADate").value = ""
'
'        Sheet1.Range("Sheet1.Form10IEAAckNo").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.Form10IEAAckNo").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.Form10IEAAckNo").value = ""
'
'Else
'        Sheet1.Range("Sheet1.Form10IEADate").value = ""
'        Sheet1.Range("Sheet1.Form10IEADate").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Form10IEADate").MergeArea.Locked = True
'
'        Sheet1.Range("Sheet1.Form10IEAAckNo").value = ""
'        Sheet1.Range("Sheet1.Form10IEAAckNo").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Form10IEAAckNo").MergeArea.Locked = True
''---end
'
'
'End If
'
'Sheet1.Protect Password:=getmsgstate
''Application.EnableEvents = True
'
'End Sub

'''''''''''''''================================='''''''''''''''''''''''''''''
'Sub ID_Opt_Date10IE(ByVal Target As Range)
'
'
'Sheet1.Unprotect Password:=getmsgstate
'        If Trim(Range("Sheet1.date10IE_Opt").value) <> "00/00/0000" And Range("Sheet1.date10IE_Opt").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE_Opt").value)) Then
'
''                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter date in dd/mm/yyyy format."
'                fmsgboxStatus "* Please enter date in dd/mm/yyyy format."
'                Sheet1.Range("Sheet1.date10IE_Opt").Select
'                Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'            Else
'
'                If (Sheet1.Range("AYinOpting").value = "2021-22") Then
'
'                        If Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2022"), "yyyy-mm-dd") Then
'
'                            fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2022 for AY 2021-22"
'                            Sheet1.Range("Sheet1.date10IE_Opt").Select
'                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        ElseIf Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2021"), "yyyy-mm-dd") Then
'
'                            fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2021 for AY 2021-22."
'                            Sheet1.Range("Sheet1.date10IE_Opt").Select
'                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        ElseIf Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                            Range("Sheet1.date10IE_Opt").Select
'                            Range("Sheet1.date10IE_Opt").value = ""
'
'                        End If
'
'                ElseIf (Sheet1.Range("AYinOpting").value = "2022-23") Then
'
'                        If Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2023"), "yyyy-mm-dd") Then
'
'                            fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2023 for AY 2022-23"
'                            Sheet1.Range("Sheet1.date10IE_Opt").Select
'                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        ElseIf Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2022"), "yyyy-mm-dd") Then
'
'                            fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2022 for AY 2022-23."
'                            Sheet1.Range("Sheet1.date10IE_Opt").Select
'                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        ElseIf Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                            Range("Sheet1.date10IE_Opt").Select
'                            Range("Sheet1.date10IE_Opt").value = ""
'
'                        End If
'
'                End If
'            End If
'        End If
'Sheet1.Protect Password:=getmsgstate
'
'End Sub


'Sub ID_OptOut_Date10IE(ByVal Target As Range)
'
'Sheet1.Unprotect Password:=getmsgstate
'
'        If Trim(Range("Sheet1.date10IE_OptOut").value) <> "00/00/0000" And Range("Sheet1.date10IE_OptOut").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE_OptOut").value)) Then
'
'                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter date in dd/mm/yyyy format."
'                Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'            Else
'
'                 If (Sheet1.Range("AYinOptOut").value = "2021-22") Then
'
'                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2022"), "yyyy-mm-dd") Then
'
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2022 in AY 2021-22"
'                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    ElseIf Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2021"), "yyyy-mm-dd") Then
'
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2021 in AY 2021-22."
'                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    ElseIf Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'
'                        fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("Sheet1.date10IE_OptOut").Select
'                        Range("Sheet1.date10IE_OptOut").value = ""
'
'                    End If
'
'                ElseIf (Sheet1.Range("AYinOptOut").value = "2022-23") Then
'
''Change-04.05.2023.101. -> V0.4 Implementation
''                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2023"), "yyyy-mm-dd") Then
''
''                        fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2023 in AY 2022-23"
''                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
''                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
''                    ElseIf Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2022"), "yyyy-mm-dd") Then
'                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2022"), "yyyy-mm-dd") Then
'
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2022 in AY 2022-23."
'                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    ElseIf Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'
'                        fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("Sheet1.date10IE_OptOut").Select
'                        Range("Sheet1.date10IE_OptOut").value = ""
'
'                     End If
'                End If
'
'            End If
'        End If
'Sheet1.Protect Password:=getmsgstate
'
'End Sub

'Sub ID_Date10IE(ByVal Target As Range)

'Sheet1.Unprotect Password:=getmsgstate
'
'        If Trim(Range("sheet1.capacityassessee").value) <> "00/00/0000" And Range("sheet1.capacityassessee").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("sheet1.capacityassessee").value)) Then
'
'                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter date in dd/mm/yyyy format."
'                Sheet1.Range("sheet1.capacityassessee").Select
'                Sheet1.Range("sheet1.capacityassessee").value = ""
'
'            Else
'                If Dformat(Trim(Range("sheet1.capacityassessee").value), "yyyy-mm-dd") < Dformat("01/04/2022", "yyyy-mm-dd") Then
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2022 in Income Details"
'                        Range("sheet1.capacityassessee").Select
'                        Range("sheet1.capacityassessee").value = ""
'
'                ElseIf Dformat(Trim(Range("sheet1.capacityassessee").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'                        fmsgboxStatus "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("sheet1.capacityassessee").Select
'                        Range("sheet1.capacityassessee").value = ""
'
'                ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt") Then
'
'                    If Dformat(Trim(Range("sheet1.capacityassessee").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2023"), "yyyy-mm-dd") Then
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2023 in Sheet Income Details"
'                        Sheet1.Range("sheet1.capacityassessee").Select
'                        Sheet1.Range("sheet1.capacityassessee").value = ""
'                    End If
'
'                ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
'
'                    If Dformat(Trim(Range("sheet1.capacityassessee").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2023"), "yyyy-mm-dd") Then
''                        fmsgboxStatus "* Date of filing cannot be prior to 1st April 2023"
''                        Sheet1.Range("sheet1.capacityassessee").Select
''                        Sheet1.Range("sheet1.capacityassessee").value = ""
'                    End If
'
'                End If
'
'            End If
'        End If
'
'Sheet1.Protect Password:=getmsgstate
'
'End Sub
'Change-25.12.2023.101.PAG.05A




'Sub Validate_CurrentAY23_24()


'Newly modified by Sai
  Sub Validate_CurrentAY24_25()


'Validation with respect to form 10IEA as per the Latest Requirement modified by Sai on 20/03/2024

    If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value <> "" Then

       If Sheet9.Range("Sheet12.NetPLAftAdjBusOthThanSpec").value <> 0 Or Sheet9.Range("Sheet12.AdjustedPLFrmSpecuBus").value <> 0 Or Sheet9.Range("Sheet12.AdjustedPLFrmSpecifiedBus").value <> 0 Then


         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then

                If Sheet1.Range("Sheet1.Form10IEADate").value = "" And Sheet1.Range("Sheet1.Form10IEADate").Locked = False Then
                   MsgBx_General = MsgBx_General & "*Please furnish date of filling of form 10IEA is mandatory" & Chr(13)
                End If
        
                 If Sheet1.Range("Sheet1.Form10IEAAckNo").value = "" And Sheet1.Range("Sheet1.Form10IEAAckNo").Locked = False Then
                    MsgBx_General = MsgBx_General & "*Acknowledgement Number of Form 10IEA is mandatory" & Chr(13)
                End If
        
        ElseIf SheetALL.Range("BacValue").value = 1 And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date" Then
        
                If Sheet1.Range("Sheet1.Form10IEADate").value = "" And Sheet1.Range("Sheet1.Form10IEADate").Locked = False Then
                  MsgBx_General = MsgBx_General & "*Please furnish date of filling of form 10IEA is mandatory" & Chr(13)
                End If
        
                If Sheet1.Range("Sheet1.Form10IEAAckNo").value = "" And Sheet1.Range("Sheet1.Form10IEAAckNo").Locked = False Then
                  MsgBx_General = MsgBx_General & "*Acknowledgement Number of Form 10IEA is mandatory" & Chr(13)
                End If
        
        
            End If
        End If
     End If

'Change-ANK
'     If (Sheet1.Range("OptingNewTaxRegime").value = "(Select)" Or Sheet1.Range("OptingNewTaxRegime").value = "") And (Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F") Then
'            If Sheet1.Range("OptingNewTaxRegime").Locked = False Then
'                MsgBx_General = MsgBx_General & "*Please select the applicable option 'Option for current assessment year'" & Chr(13)
'            End If
'     End If
'---end

'Change-ANK
'    If Sheet1.Range("Foreign_institutinal_investment").Locked = False And isdropdownblank(Sheet1.Range("Foreign_institutinal_investment").value) Then
''        MsgBx_General = MsgBx_General & "Please select Whether you are an FPI? in schedule Part A General" & Chr(13)
'        MsgBx_General = MsgBx_General & """Please select whether you are an FPI""" & Chr(13)
'    End If
    
'Change-ANK
'     If Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then
'    If Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then
'        If Trim(Range("sheet1.capacityassessee").value) = "" Or IsEmpty(Range("sheet1.capacityassessee").value) Then
'            MsgBx_General = MsgBx_General + "* Please furnish form 10-IE to opt for new tax regime" & Chr(13)
'        End If
'
'        If Trim(Range("acknowledgmentnumber").value) = "" Or IsEmpty(Range("acknowledgmentnumber").value) Then
'            MsgBx_General = MsgBx_General + "* Acknowledgement Number of Form 10IE is mandatory  for persons having income other than salary in Schedule Part A General " & Chr(13)
'        End If
'    End If
'If Sheet1.Range("AYinOpting").Locked = False Then
'
'End If


'---end
End Sub

Function ChkMaxDOBDate_2324(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate_2324 = True
     If Len(dob) > 0 Then
'        If val(Mid(dob, 7, 4)) > 2023 Then
'        If val(Mid(dob, 7, 4)) > 2024 Then 'Year Changed by Sai
'        If val(Mid(dob, 7, 4)) > 2025 Then 'Year Changed by Aavula Naresh
        If val(Mid(dob, 7, 4)) > 2026 Then 'Year Changed by Aavula Naresh
                    ChkMaxDOBDate_2324 = False
                    Exit Function
'        ElseIf val(Mid(dob, 7, 4)) = 2023 Then
'         ElseIf val(Mid(dob, 7, 4)) = 2024 Then
'         ElseIf val(Mid(dob, 7, 4)) = 2025 Then
         ElseIf val(Mid(dob, 7, 4)) = 2026 Then
         
            If val(Mid(dob, 4, 2)) >= 4 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate_2324 = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function

'Added by Aavula For AY 2026-27

Sub FurnishingAuditReport_2324(ByVal Target As Range)

    If Not Application.Intersect(Sheet1.Range("sheet1.AuditReportFurnishDate"), Sheet1.Range(Target.Address)) Is Nothing Then
    Sheet1.Unprotect Password:=getmsgstate

        If Dformat(Trim(Range("sheet1.AuditReportFurnishDate").value), "yyyy-mm-dd") < Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                                fmsgboxStatus "* Date of furnishing of the Audit Report in Part-A-General can not be more than System date"
                                Sheet1.Range("sheet1.AuditReportFurnishDate").Select
                                Sheet1.Range("sheet1.AuditReportFurnishDate").value = ""
        End If
    Sheet1.Protect Password:=getmsgstate
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''END



'Below code is commented for the Clarification issue of schedule S because of having so many blocks of employer category

'Sub Salary_Nature23_24(RowIndex As Long)
'    If Not Sheet39.Cells(RowIndex, Sheet39.Range("Salary.Nature1").Column).value = Empty Then
'
'        If Sheet39.Range("sheet.EmployerCategory2").value = "CG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory2").value = "SG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory2").value = "Central Government" Or Sheet39.Range("sheet.EmployerCategory2").value = "State Government" Then
'
'            If Sheet39.Cells(RowIndex, Sheet28.Range("Salary.Nature1").Column).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'
'
'
'                MsgBox "Exempt Allowance u/s 10(10B)-Second Proviso - Not allowed if Nature of employer is CG, SG, CG-Pensioners, SG-Pensioners"
'                Sheet39.Cells(RowIndex, Sheet39.Range("Salary.Nature1").Column).value = ""
'                Sheet39.Cells(RowIndex, Sheet39.Range("Salary.Nature1").Column).Select
'            End If
'        End If
'    End If
'
'
'
'End Sub

'If Not Application.Intersect(Sheet39.Range("Salary.Nature1"), Sheet39.Range(Target.Address)) Is Nothing Then
'
'        Salary_Nature23_24 (Target.row)
'
'End If
Sub AddRows_TCS1()
    Dim numberofrows As Long
    Dim rngname_TCS1 As Variant
    rngname_TCS1 = "TCS1.CreditName;TCS1.TAN;TCS1.PANOP;TCS1.CollectYear;TCS1.UnclaimedBFAmount;TCS1.CollectedOwnHands;TCS1.Collected37i;TCS1.ClaimedOwnHands;TCS1.ClaimedOPTCS;TCS1.ClaimedOPPAN;TCS1.CreditForward;"
    SelectLastRow ("TCS1.CreditName")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_TCS1)
End Sub

'Change-20.12.2022.101.TCS.7A
Function ValidateTCS1() As Boolean
ValidateTCS1 = True

If Not ValidateMandatoryTCS1() Then ValidateTCS1 = False
If Not ValidateTCS1ClaimedOwnHands Then ValidateTCS1 = False
If Not ValidateTCS1TAN Then ValidateTCS1 = False
If Not ValidateTCS1PANOP Then ValidateTCS1 = False
If Not ValidateTCS1ClaimedOPPAN Then ValidateTCS1 = False

End Function
Function msgbox_TCS(strmsg As String) As String
     msgValidateSheetTCS = msgValidateSheetTCS & strmsg & Chr(13)
End Function
Function ValidateMandatoryTCS1() As Boolean
ValidateMandatoryTCS1 = True
Dim i As Long


'First row must be filled first before filling second row - the first column must be filled

'Dim firstFilled
'Dim nextFilled
'For i = 1 To Sheet28.Range("TCS1.CreditName").count
'To be done later
'
'
'Next i
For i = 1 To Sheet28.Range("TCS1.CreditName").count

If (Sheet28.Range("TCS1.CreditName").Cells(i, 1).value <> "" And Sheet28.Range("TCS1.CreditName").Cells(i, 1).value <> "(Select)") Or _
    Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.PANOP").Cells(i, 1).value <> "" Or _
    (Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "(Select)") Or _
    Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.Collected37i").Cells(i, 1).value <> "" Or _
    Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value <> "" Then
    'Some thing is filled in the row
    
    If (Sheet28.Range("TCS1.CreditName").Cells(i, 1).value = "" Or Sheet28.Range("TCS1.CreditName").Cells(i, 1).value = "(Select)") Then
        ValidateMandatoryTCS1 = False
        msgbox_TCS ("Please select the drop down for the field TCS credit  relating to Self/Other Person in Schedule TCS at Sr. No. " & i)
    End If
    
    If Sheet28.Range("TCS1.TAN").Cells(i, 1).value = "" Then
        ValidateMandatoryTCS1 = False
        msgbox_TCS ("Please enter TAN of Collector at Sr. No. " & i)
    End If
    
    If Sheet28.Range("TCS1.PANOP").Cells(i, 1).value = "" And Sheet28.Range("TCS1.PANOP").Cells(i, 1).Locked = False Then
        ValidateMandatoryTCS1 = False
'Change-04.05.2023.101. -> V0.4 Implementation
'        msgbox_TCS ("Please enter the PAN of Other Person at Sr. No. " & i)
        msgbox_TCS ("Please enter the PAN of Other Person in schedule TCS at Sr. No. " & i)
    End If
'---
    If Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value = "" And Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).Locked = False Then
        ValidateMandatoryTCS1 = False
        msgbox_TCS ("Please enter Tax collected in own hands at Sr. No. " & i)
    End If
    
    If Sheet28.Range("TCS1.Collected37i").Cells(i, 1).value = "" And Sheet28.Range("TCS1.Collected37i").Cells(i, 1).Locked = False Then
        ValidateMandatoryTCS1 = False
        msgbox_TCS ("Please enter TCS collected in the hands of spouse or any other person at Sr. No. " & i)
    End If
    
    
    
    'These two field are mandatory for each other
    If (Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "(Select)") Or _
        Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value <> "" Then
        
      'Added by sudha as per SIT-94956
      
        'If Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).Locked = False Then
         If Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "" Or Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "(Select)" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).Locked = False Then
            
            ValidateMandatoryTCS1 = False
              msgbox_TCS ("Please select the Financial year in which TCS is collected at Sr. No. " & i)
             'msgbox_TCS ("Please select Year of tax deduction from drop-down in schedule TCS at Sr. No. " & i)
        End If
        
        If Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value = "" And Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).Locked = False Then
            ValidateMandatoryTCS1 = False
            msgbox_TCS ("Please enter the Amount of TCS brought forwarded at Sr. No. " & i)
        End If

    End If
    
    'These two fields are mandatory for each other
'    If (Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value <> "") Then
        
'Change-04.05.2023.101. -> V0.4 Implementation
'        If Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value = "" And Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).Locked = False Then
         If Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value = "" And Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).Locked = False And Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value <> "" Then
   
            ValidateMandatoryTCS1 = False
'            msgbox_TCS ("Please enter Tax deducted in own hands at Sr. No. " & i)
            msgbox_TCS ("Please enter Tax collected in Claimed in hands of spouse/other person at Sr. No. " & i)
        End If
'----
        If Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value = "" And Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).Locked = False And Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value <> "" Then
            ValidateMandatoryTCS1 = False
'Change-04.05.2023.101. -> V0.4 Implementation
'            msgbox_TCS ("Please enter the PAN/ Aadhaar of Other Person in schedule TCS at Sr. No. " & i)
            msgbox_TCS ("Please enter PAN in schedule TCS at Sr. No. " & i)
        End If
'---
'    End If
End If
Next i

End Function

Function ValidateTCS1ClaimedOwnHands() As Boolean
ValidateTCS1ClaimedOwnHands = True

Dim i As Long
'Change-04.05.2023.101. -> V0.4 Implementation
For i = 1 To Sheet28.Range("TCS1.CreditName").count
'    If Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value + Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value < Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value Then
     If Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value + Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value + Sheet28.Range("TCS1.Collected37i").Cells(i, 1).value < Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value Then
        ValidateTCS1ClaimedOwnHands = False
'        msgbox_TCS ("Amount in field Col 7(i) cannot be more than Col 5 or Col 6(i) as the case may be at Sr. No. " & i)
        msgbox_TCS ("Amount in field Col 7(i) cannot be more than Col 5 or Col 6(i)/Col 6(ii) as the case may be at Sr. No. " & i)
    End If
    
    If (Not isdropdownblank(Sheet28.Range("TCS1.CreditName").Cells(i, 1).value)) And Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value = "" Then
        ValidateTCS1ClaimedOwnHands = False
        msgbox_TCS ("Please enter Tax collected in own hands at Sr. No. " & i)
    End If
Next i
'----

End Function

Function ValidateTCS1TAN() As Boolean
ValidateTCS1TAN = True
Dim i As Long
For i = 1 To Sheet28.Range("TCS1.CreditName").count

If Not ValidateTantype_text(Mid(Sheet28.Range("TCS1.TAN").Cells(i, 1).value, 1, 4)) Then
         msgbox_TCS ("Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No. " & i & " in schedule TCS")
         ValidateTCS1TAN = False
         Exit Function
     End If
     
     If Not IsNumeric(Mid(Sheet28.Range("TCS1.TAN").Cells(i, 1).value, 5, 5)) And Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" Then
         msgbox_TCS ("Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No. " & i & " in schedule TCS")
         ValidateTCS1TAN = False
         Exit Function
     End If
     
     If Not ValidateTantype_text(Right(Sheet28.Range("TCS1.TAN").Cells(i, 1).value, 1)) Then
         msgbox_TCS ("Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No. " & i & " in schedule TCS")
         ValidateTCS1TAN = False
         Exit Function
     End If
     
     If Not ValidateTANFirstThreeLetters(Sheet28.Range("TCS1.TAN").Cells(i, 1).value) And Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" Then
         msgbox_TCS ("Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No. " & i & " in schedule TCS")
        ValidateTCS1TAN = False
        Exit Function
     End If
Next i

End Function

Function ValidateTCS1PANOP() As Boolean
ValidateTCS1PANOP = True
Dim i As Long
For i = 1 To Sheet28.Range("TCS1.CreditName").count

If Not checkfieldspecialcharacter(Sheet28.Range("TCS1.PANOP").Cells(i, 1).value) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1PANOP = False
         Exit Function
    End If
    
    If Not CheckPAN(UCase(Sheet28.Range("TCS1.PANOP").Cells(i, 1).value)) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1PANOP = False
         Exit Function
    End If
    
     If Sheet28.Range("TCS1.PANOP").Cells(i, 1).value = Sheet1.Range("sheet1.PAN").value And Sheet1.Range("sheet1.PAN").value <> "" Then
        'msgbox_TCS ("PAN of Other Person at Sr. No  " & i & " can not be same as Assessee PAN")
        
        msgbox_TCS ("Other person PAN cannot be same as assesse PAN or verification PAN")
        
         ValidateTCS1PANOP = False
         Exit Function
    End If
Next i

End Function

Function ValidateTCS1ClaimedOPPAN() As Boolean
ValidateTCS1ClaimedOPPAN = True
Dim i As Long
For i = 1 To Sheet28.Range("TCS1.CreditName").count

If Not checkfieldspecialcharacter(Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1ClaimedOPPAN = False
         Exit Function
    End If
    
    If Not CheckPAN(UCase(Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value)) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1ClaimedOPPAN = False
         Exit Function
    End If
    
     If Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value = Sheet1.Range("sheet1.PAN").value And Sheet1.Range("sheet1.PAN").value <> "" Then
        'msgbox_TCS ("PAN of Other Person at Sr. No  " & i & " can not be same as Assessee PAN")
         msgbox_TCS ("Other person PAN cannot be same as assesse PAN or verification PAN")
         ValidateTCS1ClaimedOPPAN = False
         Exit Function
    End If
Next i

End Function

Sub ChangeTCSTable(ByVal Target As Range)
Application.EnableEvents = False

Sheet28.Unprotect Password:=getmsgstate
Dim row As Long
row = Target.row

If Not Application.Intersect(Sheet28.Range(Target.Address), Sheet28.Range("TCS1.CreditName")) Is Nothing Then
    
    If Sheet28.Cells(row, Sheet28.Range("TCS1.CreditName").Column) = "Other Person" Then
    
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).Interior.Color = "&HCCFFCC"
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Interior.Color = "&HD8D8D8"
        
'Change-04.05.2023.101. -> V0.4 Implementation (From TCS sheet provided separately)
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Interior.Color = "&HD8D8D8"
        
        If (Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value = "" Or Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value = "(Select)") And _
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).value = "" Then

            Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Interior.Color = "&HCCFFCC"

        End If
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Interior.Color = "&HD8D8D8"
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Interior.Color = "&HD8D8D8"
        
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).Interior.Color = "&HD8D8D8"

        If (Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value = "" Or Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value = "(Select)") And _
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).value = "" Then
        
            Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Interior.Color = "&HCCFFCC"
            
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Interior.Color = "&HCCFFCC"
            
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Interior.Color = "&HCCFFCC"
        
        End If
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Interior.Color = "&HD8D8D8"
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Interior.Color = "&HCCFFCC"
        
    End If
    
    If Sheet28.Cells(row, Sheet28.Range("TCS1.CreditName").Column) = "(Select)" Or Sheet28.Cells(row, Sheet28.Range("TCS1.CreditName").Column) = "" Then
    
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TCS1.PANOP").Column).Interior.Color = "&HCCFFCC"
        
        If (Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value = "" Or Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value = "(Select)") And _
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).value = "" Then
        
            Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Interior.Color = "&HCCFFCC"
        
        End If
    End If
    
End If

If Not Application.Intersect(Sheet28.Range(Target.Address), Sheet28.Range("TCS1.CollectedOwnHands")) Is Nothing Or _
   Not Application.Intersect(Sheet28.Range(Target.Address), Sheet28.Range("TCS1.Collected37i")) Is Nothing Then

   If Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).value + Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).value > 0 Then

        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).Interior.Color = "&HD8D8D8"

        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).Interior.Color = "&HD8D8D8"

   Else

        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).Interior.Color = "&HCCFFCC"

        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).Interior.Color = "&HCCFFCC"

   End If

End If
     
'If Not Application.Intersect(Sheet28.Range(Target.Address), Sheet28.Range("TCS1.Collected37i")) Is Nothing Then
'
'    If Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).value <> "" Then
'
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Interior.Color = "&HD8D8D8"
'
'    Else
'
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Locked = False
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPPAN").Column).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If

If Not Application.Intersect(Sheet28.Range(Target.Address), Sheet28.Range("TCS1.CollectYear")) Is Nothing Or _
   Not Application.Intersect(Sheet28.Range(Target.Address), Sheet28.Range("TCS1.UnclaimedBFAmount")) Is Nothing Then
   
    If (Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value <> "" And Sheet28.Cells(row, Sheet28.Range("TCS1.CollectYear").Column).value <> "(Select)") Or _
        Sheet28.Cells(row, Sheet28.Range("TCS1.UnclaimedBFAmount").Column).value <> "" Then
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Interior.Color = "&HD8D8D8"
        
        Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Interior.Color = "&HD8D8D8"
        
'Change-04.05.2023.101. -> V0.4 Implementation
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Interior.Color = "&HD8D8D8"

'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Interior.Color = "&HD8D8D8"
'-----
    Else
    
        If (Sheet28.Cells(row, Sheet28.Range("TCS1.CreditName").Column).value = "Self") Then
        
            Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.CollectedOwnHands").Column).Interior.Color = "&HCCFFCC"
            
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOwnHands").Column).Interior.Color = "&HCCFFCC"
            
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.ClaimedOPTCS").Column).Interior.Color = "&HCCFFCC"
            
        End If
        
        If (Sheet28.Cells(row, Sheet28.Range("TCS1.CreditName").Column).value = "Other Person") Then
        
            Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Locked = False
            Sheet28.Cells(row, Sheet28.Range("TCS1.Collected37i").Column).Interior.Color = "&HCCFFCC"
            
        End If
        
    End If
   
End If

Application.EnableEvents = True
End Sub

'Change-08.05.2023.101 -> V0.4 Implementation
Sub ChangeTDS2Table(ByVal Target As Range)

Application.EnableEvents = False

Sheet28.Unprotect Password:=getmsgstate
Dim row As Long
row = Target.row


'col2 -> TDS2.TDSCreditName -> Other Person,Self
        'Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).ClearContents
        'Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).Locked = True
        'Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).Interior.Color = "&HD8D8D8"
        '
        'Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).Locked = False
        'Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).Interior.Color = "&HCCFFCC"
        
        
'col3i -> TDS2.PanOtherPerson
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Other Person") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.PanOtherPerson").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.PanOtherPerson").Column).Interior.Color = "&HCCFFCC"
    
    Else
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.PanOtherPerson").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.PanOtherPerson").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.PanOtherPerson").Column).Interior.Color = "&HD8D8D8"
    
    End If
        
        
        
'col3ii -> TDSAadharOfOther_TDS2
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Other Person") Then
    
        Sheet28.Cells(row, Sheet28.Range("TDSAadharOfOther_TDS2").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDSAadharOfOther_TDS2").Column).Interior.Color = "&HCCFFCC"
        
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDSAadharOfOther_TDS2").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDSAadharOfOther_TDS2").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDSAadharOfOther_TDS2").Column).Interior.Color = "&HD8D8D8"

        
    End If
    
'col4 -> TDS2.TAN
'        Sheet28.Cells(row, Sheet28.Range("TDS2.TAN").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TDS2.TAN").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TDS2.TAN").Column).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Cells(row, Sheet28.Range("TDS2.TAN").Column).Locked = False
'        Sheet28.Cells(row, Sheet28.Range("TDS2.TAN").Column).Interior.Color = "&HCCFFCC"
        
'col5 -> TDS2.DeductedYear
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).value <> "") Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).value <> "") Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).value <> "") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).Interior.Color = "&HD8D8D8"
    
    Else
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).Interior.Color = "&HCCFFCC"
    
    End If
        
'col6 -> TDS2.BroughtFwdTDSAmt
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).value <> "") Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).value <> "") Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).value <> "") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).Interior.Color = "&HD8D8D8"
        
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).Interior.Color = "&HCCFFCC"
        
    End If
    
'col7 -> TDS2.DeductedInOwnHands
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Other Person") Or _
        (Not isdropdownblank(Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).value)) Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).value <> "") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).Interior.Color = "&HD8D8D8"
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedInOwnHands").Column).Interior.Color = "&HCCFFCC"
    
    End If
        
'col8(Income) -> TDS2.DeductedIncome5A
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Self") Or _
        (Not isdropdownblank(Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).value)) Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).value <> "") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).Interior.Color = "&HD8D8D8"
    
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedIncome5A").Column).Interior.Color = "&HCCFFCC"
    
    End If
    
'col8(TDS) -> TDS2.DeductedTDS5A
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Self") Or _
        (Not isdropdownblank(Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedYear").Column).value)) Or _
        (Sheet28.Cells(row, Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column).value <> "") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).Interior.Color = "&HD8D8D8"
    
    Else
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.DeductedTDS5A").Column).Interior.Color = "&HCCFFCC"
    
    End If
        
'col9 -> TDS2.ClaimedInOwnHands
'        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).Locked = False
'        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).Interior.Color = "&HCCFFCC"
        
        
'col10(Income) -> TDS2.ClaimedIncome5A
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Other Person") Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedIncome5A").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedIncome5A").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedIncome5A").Column).Interior.Color = "&HD8D8D8"
    
    Else
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedIncome5A").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedIncome5A").Column).Interior.Color = "&HCCFFCC"
    
    End If
    
'col10(TDS) -> TDS2.ClaimedTDS5A
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Other Person") Then
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedTDS5A").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedTDS5A").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedTDS5A").Column).Interior.Color = "&HD8D8D8"
        
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedTDS5A").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedTDS5A").Column).Interior.Color = "&HCCFFCC"
        
    End If
    
'col10(PAN) -> TDS2.ClaimedPAN5A
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.TDSCreditName").Column).value = "Other Person") Then
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedPAN5A").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedPAN5A").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedPAN5A").Column).Interior.Color = "&HD8D8D8"
        
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedPAN5A").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedPAN5A").Column).Interior.Color = "&HCCFFCC"
        
    End If
    
'col10(Aadhaar) -> ClaimedAadhar5A_TDS2
'        Sheet28.Cells(row, Sheet28.Range("ClaimedAadhar5A_TDS2").Column).ClearContents
'        Sheet28.Cells(row, Sheet28.Range("ClaimedAadhar5A_TDS2").Column).Locked = True
'        Sheet28.Cells(row, Sheet28.Range("ClaimedAadhar5A_TDS2").Column).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Cells(row, Sheet28.Range("ClaimedAadhar5A_TDS2").Column).Locked = False
'        Sheet28.Cells(row, Sheet28.Range("ClaimedAadhar5A_TDS2").Column).Interior.Color = "&HCCFFCC"
        
        
'col11 -> TDS2.GrossAmount
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).value = 0 And (Not IsEmpty(Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column)))) Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.GrossAmount").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.GrossAmount").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.GrossAmount").Column).Interior.Color = "&HD8D8D8"
        
    Else
    
        Sheet28.Cells(row, Sheet28.Range("TDS2.GrossAmount").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.GrossAmount").Column).Interior.Color = "&HCCFFCC"
        
    End If
        
'col12 -> TDS2.HeadOfIncome
    If (Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column).value = 0 And (Not IsEmpty(Sheet28.Cells(row, Sheet28.Range("TDS2.ClaimedInOwnHands").Column)))) Then
        
        Sheet28.Cells(row, Sheet28.Range("TDS2.HeadOfIncome").Column).ClearContents
        Sheet28.Cells(row, Sheet28.Range("TDS2.HeadOfIncome").Column).Locked = True
        Sheet28.Cells(row, Sheet28.Range("TDS2.HeadOfIncome").Column).Interior.Color = "&HD8D8D8"
        
    Else

        Sheet28.Cells(row, Sheet28.Range("TDS2.HeadOfIncome").Column).Locked = False
        Sheet28.Cells(row, Sheet28.Range("TDS2.HeadOfIncome").Column).Interior.Color = "&HCCFFCC"

    End If



Sheet28.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
'--------

 Sub AddRows_TCS1_import(DiffRows As Long)
    Sheet28.Activate
    Dim rngname_TCS1 As Variant
    rngname_TCS1 = "TCS1.CreditName;TCS1.TAN;TCS1.PANOP;TCS1.CollectYear;TCS1.UnclaimedBFAmount;TCS1.CollectedOwnHands;TCS1.Collected37i;TCS1.ClaimedOwnHands;TCS1.ClaimedOPTCS;TCS1.ClaimedOPPAN;TCS1.CreditForward;"
    SelectLastRow ("TCS1.CreditName")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_TCS1)
End Sub
Function ValidateTANFirstThreeLetters(fieldValue As Variant) As Boolean
ValidateTANFirstThreeLetters = False
Dim i, j As Long
Dim arr As Variant

arr = Array("HYD", "VPN", "BBN", "BPL", "JBP", "CHE", "CMB", "MRI", "DEL", "CAL", "MRT", "AHM", "BRD", "RKT", "SRT", "BLR", "AGR", "KNP", "CHN", "TVD", _
        "ALD", "LKN", "MUM", "NGP", "AMR", "JLD", "PTL", "RTK", "KLP", "NSK", "PNE", "PTN", "RCH", "JDH", "JPR", "SHL")

For i = 1 To Len(fieldValue)
    For j = 0 To UBound(arr)
        If Mid(fieldValue, i, 3) = arr(j) Then
            ValidateTANFirstThreeLetters = True
            Exit Function
        End If
    Next
Next

End Function

Sub AddRows_Per5080G23_24()
    Dim numberofrows As Long
    Dim rngname_80G As Variant
'    rngname_80G = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.ARN_Reference;Per5080G.DonationInCash;Per5080G.DonationAmt;Per5080G.TotalDonationAmt;Per5080G.EligibleAmt;"
    rngname_80G = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.ARN_Reference;Per5080G.DonationInCash;Per5080G.DonationAmt;Per5080G.TotalDonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"

    SelectLastRow ("Per5080G.DoneeWithPanName")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_80G)
End Sub

'Change-20.01.2023.101.24C
Sub LockUnlock115AD_FII_115H(ByVal Target As Range)

        If UCase(Sheet1.Range("Sheet1.115H").value) = UCase("Yes") Or Sheet1.Range("Foreign_institutinal_investment").value = "Yes" Then
            Sheet1.UnLock115AD
        Else
            Sheet1.Lock115AD
        End If

End Sub

'Change-20.01.2023.101.24D
Sub LockUnlock115AD_FII_115H2(ByVal Target As Range)

        Application.EnableEvents = False
        Target.value = UCase(Target.value)
        LockUnlockPartBTTI
        If Mid(Trim(Sheet1.Range("Sheet1.115H").value), 1, 1) = "Y" Then
               UNLOCKcg
               ResetOSDropDown (Mid(Target.value, 1, 2))
        Else
            If Trim(Mid(Sheet1.Range("sheet1.ResidentialStatus1"), 1, 2)) <> "NR" Then
                LOCKCG
                ResetOSDropDown (Mid(Target.value, 1, 2))
            Else
                UNLOCKcg
                ResetOSDropDown (Mid(Target.value, 1, 2))
            End If
        End If

End Sub
'Change-20.01.2023.101.25A
Function isAlphanumeric(field As Variant) As Boolean
    Dim i As Long
    isAlphanumeric = True
    
    For i = 1 To Len(field)
        If ((asc(Mid(field, i, i)) > 47) And (asc(Mid(field, i, i)) < 58)) Or ((asc(Mid(field, i, i)) > 64) And (asc(Mid(field, i, i)) < 91)) Or ((asc(Mid(field, i, i)) > 96) And (asc(Mid(field, i, i)) < 123)) Then
        Else
            isAlphanumeric = False
            Exit Function
        End If
    Next

End Function
'Change-16.02.2023.103.25B
Sub LockUnLockSEBI(ByVal Target As Range)
Sheet1.Unprotect Password:=getmsgstate
    If UCase(Target.value) = "NO" Then
        Sheet1.Range("Sheet1.SEBI_Regn_No").value = ""
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sheet1.SEBI_Regn_No").MergeArea.Locked = True
        Sheet1.Range("Sheet1.SEBI_Regn_No").Interior.Color = (&HD8D8D8)
    Else
        Sheet1.Range("Sheet1.SEBI_Regn_No").MergeArea.Locked = False
        Sheet1.Range("Sheet1.SEBI_Regn_No").Interior.Color = (&HCCFFCC)
    End If
Sheet1.Protect Password:=getmsgstate
End Sub

Function MandatorySEBI() As Boolean
    MandatorySEBI = True
'Mandatory check
    If Sheet1.Range("Foreign_institutinal_investment").value = "Yes" And Sheet1.Range("Sheet1.SEBI_Regn_No").Locked = False Then
        If Sheet1.Range("Sheet1.SEBI_Regn_No").value = "" Then
            MsgBx_General = MsgBx_General & """ Please enter the SEBI Registration Number""" & Chr(13)
            MandatorySEBI = False
            Exit Function
        End If
    End If
'input alphanumeric check
End Function

Sub ValidateSEBI(Target As Range)
        If Not isAlphanumeric(Target.value) Then
            MsgBox "SEBI must be alphanumerics only with 12 characters"
            Target.value = ""
        End If
End Sub

'Newly added by sai for AY 2024-25
Function ValidateLEI() As Boolean

LEI_Number = Sheet1.Range("LEI_Number").value
LEI_Date = Sheet1.Range("LEI_Date").value

Refund_BTI = Sheet7.Range("Sheet9.RefundDue").value
ValidateLEI = True

If Refund_BTI > 500000000 Then
    
        If LEI_Number = "" Then
          MsgBx_General = "* Please, enter LEI number in Part A General Sheet" & Chr(13)
          ValidateLEI = False
          Exit Function
        End If
        
        If Trim(LEI_Date) = "" Or IsEmpty(LEI_Date) Or LEI_Date = "00/00/0000" Then
         MsgBx_General = "* Please, enter LEI Date in Part A General Sheet" & Chr(13)
         ValidateLEI = False
         Exit Function
        End If
End If

        
If Not CheckDateddmmyyyy(LEI_Date) Then
    ValidateLEI = False
    MsgBx_General = "*Please enter valid LEI date in dd/mm/yyyy format in schedule Part A General" & Chr(13)
    Exit Function
End If


If Len(LEI_Number) > 20 Then
    ValidateLEI = False
    MsgBx_General = "* LEI Number cannot exceed 20 Characters in Part A General" & Chr(13)
    Exit Function
End If

If Not CheckSpecialCharacter_New_1(LEI_Number) Then
    ValidateLEI = False
    MsgBx_General = "* LEI Number cannot  contain special characters in Part A General" & Chr(13)
    Exit Function

End If

'If Not IsNumeric(LEI_Number) Then
'    ValidateLEI = False
'    MsgBx_General = "* Please enter valid LEI Number in Part A General" & Chr(13)
'    Exit Function
'
'End If

End Function

'Change-20.01.2023.101.DPM.2A
Sub DPMAddComment45()
On Error Resume Next
Sheet10.Unprotect Password:=getmsgstate
Sheet10.Range("DPM45.RATE").Comment.Delete
With Sheet10.Range("DPM45.RATE")
.AddComment
With .Comment
.text text:="Maximum rate of depreciation shall be restricted to 40% as per Rule 5 in case of new regime is opted u/s 115BAC. Hence WDV on the first day of previous year should be added in 40% block."
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet10.Protect Password:=getmsgstate
End Sub
Sub DPMADeleteComment45()
On Error Resume Next
Sheet10.Unprotect Password:=getmsgstate
Sheet10.Range("DPM45.RATE").Comment.Delete
Sheet10.Protect Password:=getmsgstate
End Sub
Function CheckAuditDate_Before(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2017, for A.Y. 2017-18

   CheckAuditDate_Before = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckAuditDate_Before = False
          fmsgboxStatus ("""Please enter valid date in dd/mm/yyyy format in schedule Part A General""")
           Exit Function
       End If

'       If Not ChkMinAuditDate2324(dob, "31/03/2023") Then
'       If Not ChkMinAuditDate2324(dob, "31/03/2024") Then 'Year Changed by Sai
'       If Not ChkMinAuditDate2324(dob, "31/03/2025") Then 'Year Changed by Aavula Naresh
       If Not ChkMinAuditDate2324(dob, "31/03/2026") Then 'Year Changed by Aavula Naresh
'           fmsgboxStatus (errmsg & " in Part-A-General should not be less than 01/04/2023.")
'            fmsgboxStatus (errmsg & " in Part-A General should not be less than 01/04/2024.") 'Year Changed by Sai
            
'            fmsgboxStatus (errmsg & " in Part-A General should not be less than 01/04/2025.") 'Year Changed by Aavula Naresh
'           fmsgboxStatus (errmsg & " in Part A General should not be less than 01/04/2025.") 'Year Changed by Aavula Naresh
           fmsgboxStatus (errmsg & " in Part-A General should not be less than 01/04/2026""") 'Year Changed by Aavula Naresh

           CheckAuditDate_Before = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function

Function CheckAuditDate_Before_diii(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next

   CheckAuditDate_Before_diii = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckAuditDate_Before_diii = False
          fmsgboxStatus ("""Please enter valid date in dd/mm/yyyy format in schedule Part A General""")
           Exit Function
       End If

       If Not ChkMinAuditDate2324(dob, "31/03/2026") Then

           fmsgboxStatus (errmsg & " should not be less than 01/04/2026 in schedule Part A General.""")

           CheckAuditDate_Before_diii = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function
Function ChkMinAuditDate2324(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinAuditDate2324 = True
    If Len(dob) > 0 Then
'       If val(Mid(dob, 7, 4)) < 2023 Then
''       If val(Mid(dob, 7, 4)) < 2024 Then  'Year Changed by Sai
'        If val(Mid(dob, 7, 4)) < 2025 Then  'Year Changed by Aavula Naresh
        If val(Mid(dob, 7, 4)) < 2026 Then  'Year Changed by Aavula Naresh for 2026-27
       
           ChkMinAuditDate2324 = False
           Exit Function
'       ElseIf (val(Mid(dob, 7, 4)) = 2023 And val(Mid(dob, 4, 2)) <= 3) Then
'        ElseIf (val(Mid(dob, 7, 4)) = 2024 And val(Mid(dob, 4, 2)) <= 3) Then 'Year Changed by Sai
'        ElseIf (val(Mid(dob, 7, 4)) = 2025 And val(Mid(dob, 4, 2)) <= 3) Then 'Year Changed by Aavula Naresh
        ElseIf (val(Mid(dob, 7, 4)) = 2026 And val(Mid(dob, 4, 2)) <= 3) Then 'Year Changed by Aavula Naresh for 2026-27
           ChkMinAuditDate2324 = False
           Exit Function
       End If
   End If
End Function
Function ChkMinFurnishAuditDate2324(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinFurnishAuditDate2324 = True
    If Len(dob) > 0 Then
'Change-17.02.2023.103.PAG.29E
'       If val(Mid(dob, 7, 4)) < 2023 Then
'       If val(Mid(dob, 7, 4)) < 2024 Then   'Year Changed by Sai
'       If val(Mid(dob, 7, 4)) < 2025 Then   'Year Changed by Aavula Naresh
       If val(Mid(dob, 7, 4)) < 2026 Then   'Year Changed by Naresh
           ChkMinFurnishAuditDate2324 = False
           Exit Function
'       ElseIf (val(Mid(dob, 7, 4)) = 2023 And val(Mid(dob, 4, 2)) <= 3) Then
'       ElseIf (val(Mid(dob, 7, 4)) = 2024 And val(Mid(dob, 4, 2)) <= 3) Then   'Year Changed by Sai
'       ElseIf (val(Mid(dob, 7, 4)) = 2025 And val(Mid(dob, 4, 2)) <= 3) Then   'Year Changed by Aavula Naresh
       ElseIf (val(Mid(dob, 7, 4)) = 2026 And val(Mid(dob, 4, 2)) <= 3) Then   'Year Changed by Naresh
           ChkMinFurnishAuditDate2324 = False
           Exit Function
       End If
   End If
End Function

Function CheckFurnishAuditDate_Before(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2017, for A.Y. 2017-18

   CheckFurnishAuditDate_Before = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckFurnishAuditDate_Before = False
          fmsgboxStatus ("Please enter valid date in dd/mm/yyyy format in schedule Part A General")
           Exit Function
       End If
'Change-17.02.2023.103.PAG.29D
       'If Not ChkMinFurnishAuditDate2324(dob, "31/03/2022") Then
'       If Not ChkMinFurnishAuditDate2324(dob, "31/03/2023") Then
'       If Not ChkMinFurnishAuditDate2324(dob, "31/03/2024") Then  'Year Changed by Sai
'       If Not ChkMinFurnishAuditDate2324(dob, "31/03/2025") Then  'Year Changed by Aavula Naresh
       If Not ChkMinFurnishAuditDate2324(dob, "31/03/2026") Then  'Year Changed by Naresh for AY 2026-27
       
           'fmsgboxStatus (errmsg & " in Part-A-General should not be less than 01/04/2022.")
'           fmsgboxStatus (errmsg & " in Part-A-General should not be less than 01/04/2023.")
           
'            fmsgboxStatus (errmsg & " in Part-A General should not be less than 01/04/2024.") 'Year Changed by Sai
'            fmsgboxStatus (errmsg & " in Part-A General should not be less than 01/04/2025.") 'Year Changed by Aavula Naresh
            fmsgboxStatus (errmsg & " in Part-A General should not be less than 01/04/2026.") 'Year Changed by Naresh for AY 2026-27
           
           CheckFurnishAuditDate_Before = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function

Function CheckAuditDateLastTable_Before(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2017, for A.Y. 2017-18

   CheckAuditDateLastTable_Before = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckAuditDateLastTable_Before = False
          fmsgboxStatus ("Please enter valid date in dd/mm/yyyy format in schedule Part A General")
           Exit Function
       End If

'       If Not ChkMinAuditDate2324(dob, "31/03/2023") Then
'       If Not ChkMinAuditDate2324(dob, "31/03/2024") Then
'       If Not ChkMinAuditDate2324(dob, "31/03/2025") Then
       If Not ChkMinAuditDate2324(dob, "31/03/2026") Then  'Year Changed by Aavula Naresh for AY 2026-27
       
'           fmsgboxStatus ("Date should not be prior to 01/04/2023 in schedule Part A General")
'           fmsgboxStatus ("Date should not be prior to 01/04/2024 in schedule Part A General") 'Year Changed by Sai
'           fmsgboxStatus ("Date should not be prior to 01/04/2025 in schedule Part A General") 'Year Changed by Aavula Naresh
           fmsgboxStatus ("Date should not be prior to 01/04/2026 in schedule Part A General") 'Year Changed by Aavula Naresh for AY 2026-27
           CheckAuditDateLastTable_Before = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function

'Change-ANK
Function DformatExcel(inputDate As Variant) As String
    DformatExcel = ""
    If inputDate = "" Then
        DformatExcel = ""
    Else
        DformatExcel = Mid(inputDate, 9, 2) & "/" & Mid(inputDate, 6, 2) & "/" & Mid(inputDate, 1, 4)
    End If
End Function
'---

'Change-27.01.2023.101.IDS.2A
Function ValidateARNPer5080G() As Boolean
    ValidateARNPer5080G = True
    Dim i As Long
'    For i = 1 To Sheet20.Range("Per5080G.ARN_Reference").count
'        If Sheet20.Range("Per5080G.DoneeWithPanName").Cells(i, 1).value <> "" And Sheet20.Range("Per5080G.ARN_Reference").Cells(i, 1).value = "" Then
'            msgbox_Per5080G ("* Please enter the ARN number at Sr. No  " & i)
'            ValidateARNPer5080G = False
'        End If
'    Next i
End Function

'Change-28.01.2023.101.QD.2A
Function insertRowUnderSectionWithFormulaQD(tableRange As Range, maxAllowed As Long, Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional f_112A_115AD_1 As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Dim remainingRows As Long
        Dim currRows As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
        
        
        currRows = tableRange.count
        If currRows + nRows > 20 Then
            MsgBox "Maximum 20 rows allowed"
            remainingRows = maxAllowed - currRows
            nRows = remainingRows
        End If
        
        remainingRows = currRows - nRows
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If f_112A_115AD_1 = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("E" & rr.row).value = "(Select)"
            End If
            rrCount = rrCount + 1
        Next
    End If

    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormulaQD = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
End Function
'Change-28.01.2023.101.QD.2B
Sub AddRows_Trdng_QD()
    Dim vRows As Long
    Sheet6.Activate
    EfilingCommon.DefinedgridNameRange = "QDTradingConcern.ItemName||QDTradingConcern.UnitOfMeasure||QDTradingConcern.OpeningStock||QDTradingConcern.PurchaseQty||QDTradingConcern.SaleQty||QDTradingConcern.ClgStock||QDTradingConcern.AnyShortExces"
    ActiveCellRange = EfilingCommon.searchLastRow("QDTradingConcern.ItemName")
    vRows = insertRowUnderSectionWithFormulaQD(Range("QDTradingConcern.ItemName"), 20)
End Sub
'Change-28.01.2023.101.QD.2C
Sub AddRows_Rawmtrl_QD()
    Dim vRows As Long
    Sheet6.Activate
    EfilingCommon.DefinedgridNameRange = "QDRawMaterial.ItemName||QDRawMaterial.UnitOfMeasure||QDRawMaterial.OpeningStock||QDRawMaterial.PurchaseQty||QDRawMaterial.PrevYrConsum||QDRawMaterial.SaleQty||QDRawMaterial.ClgStock||QDRawMaterial.yldFinisProd||QDRawMaterial.PercentYld||QDRawMaterial.AnyShortExces"
    ActiveCellRange = EfilingCommon.searchLastRow("QDRawMaterial.ItemName")
    vRows = insertRowUnderSectionWithFormulaQD(Range("QDRawMaterial.ItemName"), 20)
End Sub
'Change-28.01.2023.101.QD.2D
Sub AddRows_Finished_QD()
    Dim vRows As Long
    Sheet6.Activate
    EfilingCommon.DefinedgridNameRange = "QDFinishrByProd.ItemName||QDFinishrByProd.UnitOfMeasure||QDFinishrByProd.OpeningStock||QDFinishrByProd.PurchaseQty||QDFinishrByProd.PrevYrConsum||QDFinishrByProd.SaleQty||QDFinishrByProd.ClgStock||QDFinishrByProd.AnyShortExces"
    ActiveCellRange = EfilingCommon.searchLastRow("QDFinishrByProd.ItemName")
    vRows = insertRowUnderSectionWithFormulaQD(Range("QDFinishrByProd.ItemName"), 20)
End Sub

'Change-28.01.2023.101.S.3A
Sub SalaryExemptAllowanceDropdown()
'Change-02.11.2023.102.Ver.0.5
'        Dim isPGov, isCGSG, j As Variant
        Dim isPGov, isCGSG, isAllPGov, j As Variant
'End Change
        isPGov = False
        isCGSG = False
'Change-02.11.2023.102.Ver.0.5
        isAllPGov = False
'End Change
        For j = 1 To Sheet39.Range("SalarySectionCount").value
            
            If Sheet39.Range("sheet.EmployerCategory" & j).value = "CG-Pensioners" Or _
            Sheet39.Range("sheet.EmployerCategory" & j).value = "SG-Pensioners" Then
                If isPGov = False Then
                   isPGov = True
                End If
            End If
            
            If Sheet39.Range("sheet.EmployerCategory" & j).value = "Central Government" Or _
            Sheet39.Range("sheet.EmployerCategory" & j).value = "State Government" Then
                If isCGSG = False Then
                   isCGSG = True
                End If
            End If
            
            If Sheet39.Range("sheet.EmployerCategory" & j).value = "PSU-Pensioners" Or _
            Sheet39.Range("sheet.EmployerCategory" & j).value = "Others-Pensioners" Then
                If isAllPGov = False Then
                   isAllPGov = True
                End If
            End If
            
        Next j
        
        Sheet39.Unprotect Password:=getmsgstate
        
'Change-16.05.2023.101 -> SIT-49865
        If SheetALL.Range("bacValue").value = 1 Then
            
            If isPGov = True And isCGSG = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesCGSGPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
                
            End If
            
            If isCGSG = True And isPGov = False Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesCGSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
'Change-02.11.2023.102.Ver.0.5
'            If isCGSG = False And isPGov = False Then
            If isCGSG = False And isPGov = False And isAllPGov = False Then
'End Change
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYes"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
'Change-02.11.2023.102.Ver.0.5
            If isCGSG = False And isPGov = False And isAllPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesAllPGov"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
'End Change
        ElseIf SheetALL.Range("bacValue").value = 2 Then
            If isCGSG = True And isPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoCGSGPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = True And isPGov = False Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoCGSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
                
'Change-02.11.2023.102.Ver.0.5
'            If isCGSG = False And isPGov = False Then
            If isCGSG = False And isPGov = False And isAllPGov = False Then
'End Change
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNo"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
'Change-02.11.2023.102.Ver.0.5
            If isCGSG = False And isPGov = False And isAllPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoAllPGov"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
'End Change
'----
        End If
        Sheet39.Protect Password:=getmsgstate
    
End Sub

Function YesNo1(inpStr As Variant) As String

    YesNo1 = ""
    If inpStr = "Y" Or inpStr = "Yes" Or inpStr = "YES" Then
        YesNo1 = "Yes"
    End If
    If inpStr = "N" Or inpStr = "No" Or inpStr = "NO" Then
        YesNo1 = "No"
    End If

End Function

Function YESNO2(inpStr As Variant) As String

    YESNO2 = ""
    If inpStr = "Y" Or inpStr = "Yes" Or inpStr = "YES" Then
        YESNO2 = "YES"
    End If
    If inpStr = "N" Or inpStr = "No" Or inpStr = "NO" Then
        YESNO2 = "NO"
    End If

End Function


Function ValidateESOP() As Boolean
ValidateESOP = True
msgbox_ESOP = "AY 2021-22" + Chr(13)

'Malli
Dim Pan_flag As Boolean
Dim dpiit_flag As Boolean

dpiit_flag = False
Pan_flag = False

'-------------
'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value < Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If

'Commented by Riyaz on 26/08/2025
'If Sheet46.Range("ESOP.SecType").value = "Partly Sold" Then
'
'If Sheet46.Range("ESOP.AttributedTaxTot").value = 0 Then
'msgbox_ESOP = msgbox_ESOP + "Total Amount of Table below  (4ii) is mandatory. In case of Col No.4 Selected as Partly" & Chr(13)
'ValidateESOP = False
'End If
'If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
'msgbox_ESOP = msgbox_ESOP + " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY " & Chr(13)
'ValidateESOP = False
'End If
'End If

If Sheet46.Range("ESOP.SecType").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value <> Sheet46.Range("ESOP.TaxPrevAY").value) Then
  ''Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
  msgbox_ESOP = msgbox_ESOP + " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' " & Chr(13)
'msgbox_ESOP = msgbox_ESOP + " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' " & Chr(13)
ValidateESOP = False
End If
End If

'Change-28.01.2023.101.ESOP -> 1
Dim cell, dateFlag
Dim row, totalTaxAtt21_22 As Long
row = 0
dateFlag = False
totalTaxAtt21_22 = 0
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" Or Sheet46.Range("ESOP.SecType").value = "Fully Sold" Then
    'The sum of the table below should be equal to the total amount 4ii
    For Each cell In Sheet46.Range("ESOP.AttributedTax").Cells
        row = row + 1
        If Sheet46.Range("ESOPCurrAY").item(row, 1).value = "2021-22" Then
            totalTaxAtt21_22 = totalTaxAtt21_22 + Sheet46.Range("ESOP.AttributedTax").item(row, 1).value
        
        
            If Sheet46.Range("ESOP.Date").item(row, 1).value = "" Then
                dateFlag = True
            End If
        
        End If
    Next cell
    
'Post confirmation with Shubham below validation commented by Riyaz - SIT-95647 on 29/09/2025
'    If totalTaxAtt21_22 <> Sheet46.Range("ESOP.AttributedTaxTot").value Then
'        msgbox_ESOP = msgbox_ESOP + "Amount of tax attributed out of sale (iii) should match Total Amount of Table below (4ii) In case of Col No.4 Selected as Partly or Fully" & Chr(13)
'    ValidateESOP = False
'    End If
    
    If dateFlag Then
        msgbox_ESOP = msgbox_ESOP + "Date (ii) Mandatory Yes if Col No 4 selected as Fully Sold or Partly Sold" & Chr(13)
    ValidateESOP = False
    End If
    
    
End If
'---------

If Sheet46.Range("ESOP.SecType").value = "Not Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value <> 0) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP = msgbox_ESOP + "  'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares " & Chr(13)
'msgbox_ESOP = msgbox_ESOP + " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares " & Chr(13)
ValidateESOP = False
End If
End If

'If (Not isdropdownblank(Sheet46.Range("ESOP.SecType").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY").value) <> "") And (Not isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee").value)) And ((Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" And Trim(Sheet46.Range("ESOP.DateOfCeasing")) <> "") Or Sheet46.Range("ESOP.CeasedEmployee").value = "No") Then
'
'ElseIf (isdropdownblank(Sheet46.Range("ESOP.SecType").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY").value) = "") And (isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee").value)) And Trim(Sheet46.Range("ESOP.DateOfCeasing")) = "" Then
'
'Else
''msgbox_ESOP = msgbox_ESOP + " Please Fill all the mandatory fields" & Chr(13)
'ValidateESOP = False
'End If
'
''Malli
''============================
'If Sheet46.Range("ESOP.PAN").value <> "" Then
'    If Not Part_A_General.CheckPAN(UCase(Trim(Sheet46.Range("ESOP.PAN").value))) Then
'
'       'msgbox_ESOP = "Invalid PAN in 'PAN of the employer being an eligible startup'. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."""
'      msgbox_ESOP = msgbox_ESOP + "Invalid PAN in 'PAN of the employer being an eligible startup'. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."""
'       Sheet46.Range("ESOP.PAN").Select
'       Sheet46.Range("ESOP.PAN").value = ""
'       ValidateESOP = False
'
'    End If
'End If
'
'If Sheet46.Range("ESOP.DPIIT").value <> "" Then
'
'   If Not Part_A_General.ChkDPIIT(UCase(Trim(Sheet46.Range("ESOP.DPIIT").value))) Then
'           msgbox_ESOP = msgbox_ESOP + "Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"
'             Sheet46.Range("ESOP.DPIIT").Select
'             Sheet46.Range("ESOP.DPIIT").value = ""
'    End If
'
'    If Not checkallfieldspecialcharacter_80U_80DD(UCase(Trim(Sheet46.Range("ESOP.DPIIT").value))) Then
'            msgbox_ESOP = msgbox_ESOP + "Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"
'             Sheet46.Range("ESOP.DPIIT").Select
'             Sheet46.Range("ESOP.DPIIT").value = ""
'
'    End If
'End If
''---Ended
'
'
'If (Sheet46.Range("ESOP.SecType").value <> "(Select)" Or Sheet46.Range("ESOP.SecType").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing").Locked = True) Then
'
'
'    If Sheet46.Range("ESOP.PAN").value = "" Then
'    If Pan_flag = False Then
'      'msgbox_ESOP = msgbox_ESOP + """Please enter PAN at Schedule ESOP Sheet""" & Chr(13)
'      'Added by Aavula
'      msgbox_ESOP = msgbox_ESOP + """Please enter PAN of donee  in schedule ESOP""" & Chr(13)
'      Pan_flag = True
'     End If
'     ValidateESOP = False
'    End If
'
'    If Sheet46.Range("ESOP.DPIIT").value = "" Then
'      If dpiit_flag = False Then
'         msgbox_ESOP = msgbox_ESOP + """Please enter the Start up recognition number of employer""" & Chr(13)
'        dpiit_flag = True
'      End If
'         ValidateESOP = False
'    End If
'
'
'
'ElseIf (Sheet46.Range("ESOP.SecType").value = "(Select)" Or Sheet46.Range("ESOP.SecType").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing")) = "" Then
'
'Else
'
''Newly added By Bindu
'   If Sheet46.Range("ESOP.PAN").value = "" Then
'     If Pan_flag = False Then
'        msgbox_ESOP = msgbox_ESOP + """Please enter PAN at Schedule ESOP Sheet""" & Chr(13)
'        Pan_flag = True
'      End If
'        ValidateESOP = False
'    End If
'
'    If Sheet46.Range("ESOP.DPIIT").value = "" Then
'        If dpiit_flag = False Then
'            dpiit_flag = True
'            msgbox_ESOP = msgbox_ESOP + """Please enter the Start up recognition number of employer""" & Chr(13)
'         End If
'            ValidateESOP = False
'    End If
'
'
'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'ValidateESOP = False
'End If

'Change.10.01.2023.104
'Commented by Riyaz on 26/08/2025
'If Sheet46.Range("ESOP.SecType").value = "Partly Sold" Then
'If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
'msgbox_ESOP = msgbox_ESOP + " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY "
'ValidateESOP = False
'End If
'End If

If Sheet46.Range("ESOP.SecType").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value <> Sheet46.Range("ESOP.TaxPrevAY").value) Then

''Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP = msgbox_ESOP + " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' "
'msgbox_ESOP = msgbox_ESOP + " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value <> 0) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP = msgbox_ESOP + "  'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
'msgbox_ESOP = msgbox_ESOP + " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If

'If (Sheet46.Range("ESOP.SecType2").value <> "(Select)" Or Sheet46.Range("ESOP.SecType2").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY2").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee2").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee2").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing2")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing2").Locked = True) Then
'
'   If Sheet46.Range("ESOP.PAN").value = "" Then
'     If Pan_flag = False Then
'        Pan_flag = True
'        msgbox_ESOP = msgbox_ESOP + """Please enter PAN at Schedule ESOP Sheet""" & Chr(13)
'     End If
'      ValidateESOP = False
'    End If
'
'    If Sheet46.Range("ESOP.DPIIT").value = "" Then
'        If dpiit_flag = False Then
'           dpiit_flag = True
'           msgbox_ESOP = msgbox_ESOP + """Please enter the Start up recognition number of employer""" & Chr(13)
'         End If
'         ValidateESOP = False
'    End If
'
'
'ElseIf (Sheet46.Range("ESOP.SecType2").value = "(Select)" Or Sheet46.Range("ESOP.SecType2").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY2").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee2").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee2").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing2")) = "" Then
'
'Else
'
''Newly added by Bindu
'If Sheet46.Range("ESOP.PAN").value = "" Then
'    If Pan_flag = False Then
'        Pan_flag = True
'        msgbox_ESOP = msgbox_ESOP + """Please enter PAN at Schedule ESOP Sheet""" & Chr(13)
'    End If
'        ValidateESOP = False
'    End If
'
'    If Sheet46.Range("ESOP.DPIIT").value = "" Then
'         If dpiit_flag = False Then
'            dpiit_flag = True
'            msgbox_ESOP = msgbox_ESOP + """Please enter the Start up recognition number of employer""" & Chr(13)
'        End If
'         ValidateESOP = False
'    End If
'
'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'ValidateESOP = False
'End If

'Chandru

If Sheet46.Range("ESOP.SecType2").value = "Partly Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value > Sheet46.Range("ESOP.TaxPrevAY2").value) Or (Sheet46.Range("ESOP.AttributedTaxTot2").value = Sheet46.Range("ESOP.TaxPrevAY2").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If
If Sheet46.Range("ESOP.SecType2").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value <> Sheet46.Range("ESOP.TaxPrevAY2").value) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' "
'msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType2").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value <> 0) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
 msgbox_ESOP = "'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares"
'msgbox_ESOP = " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If

'If (Sheet46.Range("ESOP.SecType3").value <> "(Select)" Or Sheet46.Range("ESOP.SecType3").value <> "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY3").value) <> "") And (Sheet46.Range("ESOP.CeasedEmployee3").value <> "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee3").value <> "") And (Trim(Sheet46.Range("ESOP.DateOfCeasing3")) <> "" Or Sheet46.Range("ESOP.DateOfCeasing3").Locked = True) Then
'
'    If Sheet46.Range("ESOP.PAN").value = "" Then
'      If Pan_flag = False Then
'          Pan_flag = True
'          msgbox_ESOP = msgbox_ESOP + """Please enter PAN at Schedule ESOP Sheet""" & Chr(13)
'       End If
'      ValidateESOP = False
'    End If
'
'    If Sheet46.Range("ESOP.DPIIT").value = "" Then
'        If dpiit_flag = False Then
'           dpiit_flag = True
'           msgbox_ESOP = msgbox_ESOP + """Please enter the Start up recognition number of employer""" & Chr(13)
'         End If
'         ValidateESOP = False
'    End If
'
'
'
'ElseIf (Sheet46.Range("ESOP.SecType3").value = "(Select)" Or Sheet46.Range("ESOP.SecType3").value = "") And (Trim(Sheet46.Range("ESOP.TaxPrevAY3").value) = "") And (Sheet46.Range("ESOP.CeasedEmployee3").value = "(Select)" Or Sheet46.Range("ESOP.CeasedEmployee3").value = "") And Trim(Sheet46.Range("ESOP.DateOfCeasing3")) = "" Then
'
'
'Else
'
''Newly added by Bindu
'    If Sheet46.Range("ESOP.PAN").value = "" Then
'      If Pan_flag = False Then
'          Pan_flag = True
'          msgbox_ESOP = msgbox_ESOP + """Please enter PAN at Schedule ESOP Sheet""" & Chr(13)
'       End If
'      ValidateESOP = False
'    End If
'
'    If Sheet46.Range("ESOP.DPIIT").value = "" Then
'        If dpiit_flag = False Then
'           dpiit_flag = True
'           msgbox_ESOP = msgbox_ESOP + """Please enter the Start up recognition number of employer""" & Chr(13)
'         End If
'         ValidateESOP = False
'    End If
'
'
'msgbox_ESOP = msgbox_ESOP + "Please Fill all the mandatory fields" & Chr(13)
'ValidateESOP = False
'End If

If Sheet46.Range("ESOP.SecType3").value = "Partly Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value > Sheet46.Range("ESOP.TaxPrevAY3").value) Or (Sheet46.Range("ESOP.AttributedTaxTot3").value = Sheet46.Range("ESOP.TaxPrevAY3").value) Then
msgbox_ESOP = " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP = False
End If
End If
If Sheet46.Range("ESOP.SecType3").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value <> Sheet46.Range("ESOP.TaxPrevAY3").value) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' "
'msgbox_ESOP = " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' "
ValidateESOP = False
End If
End If

If Sheet46.Range("ESOP.SecType3").value = "Not sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value <> 0) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP = " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares "
'msgbox_ESOP = " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares "
ValidateESOP = False
End If
End If
'----------
'============================
'----------
'============================

End Function

Function ValidateESOP2() As Boolean
ValidateESOP2 = True
msgbox_ESOP2 = "AY 2022-23" + Chr(13)

'Commented by Riyaz on 26/08/2025
'If Sheet46.Range("ESOP.SecType2").value = "Partly Sold" Then
'If Sheet46.Range("ESOP.AttributedTaxTot2").value = 0 Then
'msgbox_ESOP2 = msgbox_ESOP2 + "Total Amount of Table below  (4ii) is mandatory. In case of Col No.4 Selected as Partly" & Chr(13)
'ValidateESOP2 = False
'End If
'If (Sheet46.Range("ESOP.AttributedTaxTot2").value > Sheet46.Range("ESOP.TaxPrevAY2").value) Or (Sheet46.Range("ESOP.AttributedTaxTot2").value = Sheet46.Range("ESOP.TaxPrevAY2").value) Then
'msgbox_ESOP2 = msgbox_ESOP2 + " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY " & Chr(13)
'ValidateESOP2 = False
'End If
'End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP2 = msgbox_ESOP2 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP2 = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value < Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP2 = msgbox_ESOP2 + " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP2 = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP2 = msgbox_ESOP2 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP2 = False
End If
End If


If Sheet46.Range("ESOP.SecType2").value = "Fully Sold" Then
    If (Sheet46.Range("ESOP.AttributedTaxTot2").value <> Sheet46.Range("ESOP.TaxPrevAY2").value) Then
     
     'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
     msgbox_ESOP2 = msgbox_ESOP2 + " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY'  " & Chr(13)
    'msgbox_ESOP2 = msgbox_ESOP2 + " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' " & Chr(13)
    
    ValidateESOP2 = False
    End If
End If

'Change-28.01.2023.101.ESOP -> 2
Dim cell, dateFlag
Dim row, totalTaxAtt22_23 As Long
row = 0
totalTaxAtt22_23 = 0
dateFlag = False
If Sheet46.Range("ESOP.SecType2").value = "Partly Sold" Or Sheet46.Range("ESOP.SecType2").value = "Fully Sold" Then
    'The sum of the table below should be equal to the total amount 4ii
    For Each cell In Sheet46.Range("ESOP.AttributedTax").Cells
        row = row + 1
        If Sheet46.Range("ESOPCurrAY").item(row, 1).value = "2022-23" Then
            totalTaxAtt22_23 = totalTaxAtt22_23 + Sheet46.Range("ESOP.AttributedTax").item(row, 1).value
        
            If Sheet46.Range("ESOP.Date").item(row, 1).value = "" Then
                dateFlag = True
            End If
        
        End If
    Next cell
    
'Post confirmation with Shubham below validation commented by Riyaz - SIT-95647 on 29/09/2025
'    If totalTaxAtt22_23 <> Sheet46.Range("ESOP.AttributedTaxTot2").value Then
'        msgbox_ESOP2 = msgbox_ESOP2 + "Amount of tax attributed out of sale (iii) should match Total Amount of Table below (4ii) In case of Col No.4 Selected as Partly or Fully" & Chr(13)
'        ValidateESOP2 = False
'    End If
    
    If dateFlag Then
        msgbox_ESOP2 = msgbox_ESOP2 + "Date (ii) Mandatory Yes if Col No 4 selected as Fully Sold or Partly Sold" & Chr(13)
        ValidateESOP2 = False
    End If
    
End If
'------------

If Sheet46.Range("ESOP.SecType2").value = "Not Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot2").value <> 0) Then
 'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
 msgbox_ESOP2 = msgbox_ESOP2 + " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares " & Chr(13)
'msgbox_ESOP2 = msgbox_ESOP2 + " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares " & Chr(13)
ValidateESOP2 = False
End If
End If

If (Not isdropdownblank(Sheet46.Range("ESOP.SecType2").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY2").value) <> "") And (Not isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee2").value)) And ((Sheet46.Range("ESOP.CeasedEmployee2").value = "Yes" And Trim(Sheet46.Range("ESOP.DateOfCeasing2")) <> "") Or Sheet46.Range("ESOP.CeasedEmployee2").value = "No") Then

ElseIf (isdropdownblank(Sheet46.Range("ESOP.SecType2").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY2").value) = "") And (isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee2").value)) And Trim(Sheet46.Range("ESOP.DateOfCeasing2")) = "" Then

Else
msgbox_ESOP2 = msgbox_ESOP2 + " Please Fill all the mandatory fields" & Chr(13)
ValidateESOP2 = False
End If

End Function
Function ValidateESOP3() As Boolean
ValidateESOP3 = True
msgbox_ESOP3 = "AY 2023-24" + Chr(13)

'Commented by Riyaz on 26/08/2025
'If Sheet46.Range("ESOP.SecType3").value = "Partly Sold" Then
'If Sheet46.Range("ESOP.AttributedTaxTot3").value = 0 Then
'msgbox_ESOP3 = msgbox_ESOP3 + "Total Amount of Table below  (4ii) is mandatory. In case of Col No.4 Selected as Partly" & Chr(13)
'ValidateESOP3 = False
'End If
'If (Sheet46.Range("ESOP.AttributedTaxTot3").value > Sheet46.Range("ESOP.TaxPrevAY3").value) Or (Sheet46.Range("ESOP.AttributedTaxTot3").value = Sheet46.Range("ESOP.TaxPrevAY3").value) Then
'msgbox_ESOP3 = msgbox_ESOP3 + " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY " & Chr(13)
'ValidateESOP3 = False
'End If
'End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP3 = msgbox_ESOP3 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP3 = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value < Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP3 = msgbox_ESOP3 + " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP3 = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP3 = msgbox_ESOP3 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP3 = False
End If
End If

If Sheet46.Range("ESOP.SecType3").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value <> Sheet46.Range("ESOP.TaxPrevAY3").value) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP3 = msgbox_ESOP3 + " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' " & Chr(13)
'msgbox_ESOP3 = msgbox_ESOP3 + " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' " & Chr(13)
ValidateESOP3 = False
End If
End If


Dim cell, dateFlag
Dim row, totalTaxAtt23_24 As Long
row = 0
totalTaxAtt23_24 = 0
dateFlag = False
If Sheet46.Range("ESOP.SecType3").value = "Partly Sold" Or Sheet46.Range("ESOP.SecType3").value = "Fully Sold" Then
    'The sum of the table below should be equal to the total amount 4ii
    For Each cell In Sheet46.Range("ESOP.AttributedTax").Cells
        row = row + 1
        If Sheet46.Range("ESOPCurrAY").item(row, 1).value = "2023-24" Then
            totalTaxAtt23_24 = totalTaxAtt23_24 + Sheet46.Range("ESOP.AttributedTax").item(row, 1).value
        
            If Sheet46.Range("ESOP.Date").item(row, 1).value = "" Then
                dateFlag = True
            End If
        
        End If
    Next cell
    
'Post confirmation with Shubham below validation commented by Riyaz - SIT-95647 on 29/09/2025
'    If totalTaxAtt23_24 <> Sheet46.Range("ESOP.AttributedTaxTot3").value Then
'        msgbox_ESOP3 = msgbox_ESOP3 + "Amount of tax attributed out of sale (iii) should match Total Amount of Table below (4ii) In case of Col No.4 Selected as Partly or Fully" & Chr(13)
'        ValidateESOP3 = False
'    End If
    
    If dateFlag Then
        msgbox_ESOP3 = msgbox_ESOP3 + "Date (ii) Mandatory Yes if Col No 4 selected as Fully Sold or Partly Sold" & Chr(13)
        ValidateESOP3 = False
    End If
    
End If
'------------

If Sheet46.Range("ESOP.SecType3").value = "Not Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot3").value <> 0) Then
'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP3 = msgbox_ESOP3 + " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares " & Chr(13)
'msgbox_ESOP3 = msgbox_ESOP3 + " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares " & Chr(13)
ValidateESOP3 = False
End If
End If

If (Not isdropdownblank(Sheet46.Range("ESOP.SecType3").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY3").value) <> "") And (Not isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee3").value)) And ((Sheet46.Range("ESOP.CeasedEmployee3").value = "Yes" And Trim(Sheet46.Range("ESOP.DateOfCeasing3")) <> "") Or Sheet46.Range("ESOP.CeasedEmployee3").value = "No") Then

ElseIf (isdropdownblank(Sheet46.Range("ESOP.SecType3").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY3").value) = "") And (isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee3").value)) And Trim(Sheet46.Range("ESOP.DateOfCeasing3")) = "" Then

Else
msgbox_ESOP3 = msgbox_ESOP3 + " Please Fill all the mandatory fields" & Chr(13)
ValidateESOP3 = False
End If

End Function

Function ValidateESOP4() As Boolean
ValidateESOP4 = True
msgbox_ESOP4 = "AY 2024-25" + Chr(13)

'Commented by Riyazon 26/08/2025
'If Sheet46.Range("ESOP.SecType4").value = "Partly Sold" Then
'If Sheet46.Range("ESOP.AttributedTaxTot4").value = 0 Then
'msgbox_ESOP4 = msgbox_ESOP4 + "Total Amount of Table below  (4ii) is mandatory. In case of Col No.4 Selected as Partly" & Chr(13)
'ValidateESOP4 = False
'End If
'If (Sheet46.Range("ESOP.AttributedTaxTot4").value > Sheet46.Range("ESOP.TaxPrevAY4").value) Or (Sheet46.Range("ESOP.AttributedTaxTot4").value = Sheet46.Range("ESOP.TaxPrevAY4").value) Then
'msgbox_ESOP4 = msgbox_ESOP4 + " Total Amount of Table below  (4ii) can not exceed Amount of Tax deferred brought forward from earlier AY " & Chr(13)
'ValidateESOP4 = False
'End If
'End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP4 = msgbox_ESOP4 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP4 = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value < Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP4 = msgbox_ESOP4 + " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP4 = False
End If
End If

'Added by Riyaz on 26/08/2025
If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP4 = msgbox_ESOP4 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP4 = False
End If
End If


If Sheet46.Range("ESOP.SecType4").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value <> Sheet46.Range("ESOP.TaxPrevAY4").value) Then
 'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP4 = msgbox_ESOP4 + " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' " & Chr(13)
'msgbox_ESOP4 = msgbox_ESOP4 + " Total Amount of Table below  (4ii) for Fully paid sold should be equal to 'Amount of tax payable in the current Assessment Year ' " & Chr(13)
ValidateESOP4 = False
End If
End If


Dim cell, dateFlag
Dim row, totalTaxAtt24_25 As Long
row = 0
totalTaxAtt24_25 = 0
dateFlag = False
If Sheet46.Range("ESOP.SecType4").value = "Partly Sold" Or Sheet46.Range("ESOP.SecType4").value = "Fully Sold" Then
    'The sum of the table below should be equal to the total amount 4ii
    For Each cell In Sheet46.Range("ESOP.AttributedTax").Cells
        row = row + 1
        If Sheet46.Range("ESOPCurrAY").item(row, 1).value = "2024-25" Then
            totalTaxAtt24_25 = totalTaxAtt24_25 + Sheet46.Range("ESOP.AttributedTax").item(row, 1).value
        
            If Sheet46.Range("ESOP.Date").item(row, 1).value = "" Then
                dateFlag = True
            End If
        
        End If
    Next cell
    
'Post confirmation with Shubham below validation commented by Riyaz - SIT-95647 on 29/09/2025
'    If totalTaxAtt24_25 <> Sheet46.Range("ESOP.AttributedTaxTot4").value Then
'        msgbox_ESOP4 = msgbox_ESOP4 + "Amount of tax attributed out of sale (iii) should match Total Amount of Table below (4ii) In case of Col No.4 Selected as Partly or Fully" & Chr(13)
'        ValidateESOP4 = False
'    End If

    
    If dateFlag Then
        msgbox_ESOP4 = msgbox_ESOP4 + "Date (ii) Mandatory Yes if Col No 4 selected as Fully Sold or Partly Sold" & Chr(13)
        ValidateESOP4 = False
    End If
    
End If


If Sheet46.Range("ESOP.SecType4").value = "Not Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot4").value <> 0) Then
 'Msg updated by Bindu as per ITR2 Confirmed with Shubham(ITR3) & Sonam(ITR2) on 9th Sep 25
msgbox_ESOP4 = msgbox_ESOP4 + " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares " & Chr(13)
'msgbox_ESOP4 = msgbox_ESOP4 + " 'Amount of tax payable in the current Assessment Year ' should be 0 for Not Sold Shares " & Chr(13)
ValidateESOP4 = False
End If
End If

If (Not isdropdownblank(Sheet46.Range("ESOP.SecType4").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY4").value) <> "") And (Not isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee4").value)) And ((Sheet46.Range("ESOP.CeasedEmployee4").value = "Yes" And Trim(Sheet46.Range("ESOP.DateOfCeasing4")) <> "") Or Sheet46.Range("ESOP.CeasedEmployee4").value = "No") Then

ElseIf (isdropdownblank(Sheet46.Range("ESOP.SecType4").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY4").value) = "") And (isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee4").value)) And Trim(Sheet46.Range("ESOP.DateOfCeasing4")) = "" Then

Else
msgbox_ESOP4 = msgbox_ESOP4 + " Please Fill all the mandatory fields" & Chr(13)
ValidateESOP4 = False
End If

End Function
Function ValidateESOP5() As Boolean
ValidateESOP5 = True
msgbox_ESOP5 = "AY 2025-26" + Chr(13)


If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "No" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Or (Sheet46.Range("ESOP.AttributedTaxTot").value = Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP5 = msgbox_ESOP5 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP5 = False
End If
End If


If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value < Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP5 = msgbox_ESOP5 + " Total Amount of Table below  (4ii) for Partly Sold should be equal to 'Amount of Tax deferred brought forward from earlier AY ' "
ValidateESOP5 = False
End If
End If


If Sheet46.Range("ESOP.SecType").value = "Partly Sold" And Sheet46.Range("ESOP.CeasedEmployee").value = "Yes" Then
If (Sheet46.Range("ESOP.AttributedTaxTot").value > Sheet46.Range("ESOP.TaxPrevAY").value) Then
msgbox_ESOP5 = msgbox_ESOP5 + " Total Amount of Table below  (4ii) for Partly Sold can not exceed Amount of Tax deferred brought forward from earlier AY "
ValidateESOP5 = False
End If
End If


If Sheet46.Range("ESOP.SecType5").value = "Fully Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot5").value <> Sheet46.Range("ESOP.TaxPrevAY5").value) Then
msgbox_ESOP5 = msgbox_ESOP5 + " Total Amount of Table below  (4ii) for Fully sold should be equal to 'Amount of Tax deferred brought forward from earlier AY' " & Chr(13)
ValidateESOP5 = False
End If
End If


Dim cell, dateFlag
Dim row, totalTaxAtt25_26 As Long
row = 0
totalTaxAtt25_26 = 0
dateFlag = False
If Sheet46.Range("ESOP.SecType5").value = "Partly Sold" Or Sheet46.Range("ESOP.SecType5").value = "Fully Sold" Then
    'The sum of the table below should be equal to the total amount 4ii
    For Each cell In Sheet46.Range("ESOP.AttributedTax").Cells
        row = row + 1
        If Sheet46.Range("ESOPCurrAY").item(row, 1).value = "2025-26" Then
            totalTaxAtt25_26 = totalTaxAtt25_26 + Sheet46.Range("ESOP.AttributedTax").item(row, 1).value
        
            If Sheet46.Range("ESOP.Date").item(row, 1).value = "" Then
                dateFlag = True
            End If
        
        End If
    Next cell
    
    
    If dateFlag Then
        msgbox_ESOP5 = msgbox_ESOP5 + "Date (ii) Mandatory Yes if Col No 4 selected as Fully Sold or Partly Sold" & Chr(13)
        ValidateESOP5 = False
    End If
    
End If


If Sheet46.Range("ESOP.SecType5").value = "Not Sold" Then
If (Sheet46.Range("ESOP.AttributedTaxTot5").value <> 0) Then
msgbox_ESOP5 = msgbox_ESOP5 + " 'Total Amount of Table below  (4ii)' should be 0 for Not Sold Shares " & Chr(13)
ValidateESOP5 = False
End If
End If

If (Not isdropdownblank(Sheet46.Range("ESOP.SecType5").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY5").value) <> "") And (Not isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee5").value)) And ((Sheet46.Range("ESOP.CeasedEmployee5").value = "Yes" And Trim(Sheet46.Range("ESOP.DateOfCeasing5")) <> "") Or Sheet46.Range("ESOP.CeasedEmployee5").value = "No") Then

ElseIf (isdropdownblank(Sheet46.Range("ESOP.SecType5").value)) And (Trim(Sheet46.Range("ESOP.TaxPrevAY5").value) = "") And (isdropdownblank(Sheet46.Range("ESOP.CeasedEmployee5").value)) And Trim(Sheet46.Range("ESOP.DateOfCeasing5")) = "" Then

Else
msgbox_ESOP5 = msgbox_ESOP5 + " Please Fill all the mandatory fields" & Chr(13)
ValidateESOP5 = False
End If

End Function


''''''''''''''Added by Aavula Naresh For AY 2025-26-----------

Function ValidateESOP_PDA()
ValidateESOP_PDA = True
Dim Pan_flag As Boolean
Dim dpiit_flag As Boolean
msgbox_ESOPPD = ""
Pan_flag = False
dpiit_flag = False
 
If Sheet46.Range("ESOP.BalTaxCF25_26").value > 0 Then
 
    If Sheet46.Range("ESOP.PAN").value = "" Then
    If Pan_flag = False Then
 
       msgbox_ESOPPD = msgbox_ESOPPD + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
        ValidateESOP_PDA = True
      Pan_flag = True
     End If
     ValidateESOP_PDA = False
    End If
 
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
      If dpiit_flag = False Then
 
       msgbox_ESOPPD = msgbox_ESOPPD + "Please enter the Start up recognition number of employer" & Chr(13)
        ValidateESOP_PDA = True
        dpiit_flag = True
      End If
         ValidateESOP_PDA = False
    End If
 
 
End If
 
 
If Len(Sheet46.Range("ESOP.DPIIT").value) > 0 Then
 
    If Sheet46.Range("ESOP.PAN").value = "" Then
    If Pan_flag = False Then
 
       msgbox_ESOPPD = msgbox_ESOPPD + "Please enter PAN at Schedule ESOP Sheet" & Chr(13)
ValidateESOP_PDA = True
 
      Pan_flag = True
     End If
     ValidateESOP_PDA = False
    End If
 
End If
 
If Len(Sheet46.Range("ESOP.PAN").value) > 0 Then
 
    If Sheet46.Range("ESOP.DPIIT").value = "" Then
      If dpiit_flag = False Then
       msgbox_ESOPPD = msgbox_ESOPPD + "Please enter the Start up recognition number of employer" & Chr(13)
        ValidateESOP_PDA = True
        dpiit_flag = True
      End If
         ValidateESOP_PDA = False
    End If
 
End If
End Function


'Change-07.02.2023.101.80G.02A
Function ValidatePAN_ARN_80GD() As Boolean
ValidatePAN_ARN_80GD = True
'if pan and arn are same, then show error
Dim i, j As Long
For i = 1 To Range("Per5080G.ARN_Reference").count
    'If Range("Per5080G.DoneePAN").Cells(i, 1).value <> "" And Range("Per5080G.DoneeARN").Cells(i, 1).value <> "" Then
    If Range("Per5080G.DoneePAN").Cells(i, 1).value <> "" Then
    For j = i + 1 To Range("Per5080G.ARN_Reference").count
        If Range("Per5080G.DoneePAN").Cells(i, 1).value = Range("Per5080G.DoneePAN").Cells(j, 1).value And _
            Range("Per5080G.ARN_Reference").Cells(i, 1).value = Range("Per5080G.ARN_Reference").Cells(j, 1).value Then
            msgbox_Per5080G ("*Donation to same donee cannot be entered more than once, kindly enter gross amount of donation or enter different ARN, if applicable..")
            ValidatePAN_ARN_80GD = False
            Exit Function
        End If
    Next
    End If
Next
End Function
Function ValidateTotalCapGains_8b() As Boolean
    Dim i As Long

    ValidateTotalCapGains_8b = True
    TotalCapGains_8b = Sheet7.Range("Sheet8b.TotalCapitalGains").value

    If Len(TotalCapGains_8b) > 14 Then
        msgbox8b ("Total Capital Gains should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotalCapGains_8b = False
        Exit Function
    End If

    For i = 1 To Len(TotalCapGains_8b)
        If Not IsNumeric(TotalCapGains_8b) Then
            msgbox8b ("Total Capital Gains must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateTotalCapGains_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateShortTermLongTermTotal_8b() As Boolean
    Dim i As Long

    ValidateShortTermLongTermTotal_8b = True
    ShortTermLongTermTotal_8b = Sheet7.Range("Sheet8b.TotalCapGains").value

    If Len(ValidateShortTermLongTermTotal_8b) > 14 Then
        msgbox8b ("Sum of Short-term/Long-term Gains should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateShortTermLongTermTotal_8b = False
        Exit Function
    End If

    For i = 1 To Len(ValidateShortTermLongTermTotal_8b)
        If Not IsNumeric(ValidateShortTermLongTermTotal_8b) Then
            msgbox8b ("Sum of Short-term/Long-term  must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateShortTermLongTermTotal_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateCapGains30Per115BBH_8b() As Boolean
    Dim i As Long

    ValidateCapGains30Per115BBH_8b = True
    CapGains30Per115BBH_8b = Sheet7.Range("Sheet8b.TotalCapGains115BBH").value

    If Len(CapGains30Per115BBH_8b) > 14 Then
        msgbox8b ("Capital gain chargeable @ 30% u/s 115BBH should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateCapGains30Per115BBH_8b = False
        Exit Function
    End If

    For i = 1 To Len(CapGains30Per115BBH_8b)
        If Not IsNumeric(CapGains30Per115BBH_8b) Then
            msgbox8b ("Capital gain chargeable @ 30% u/s 115BBH must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateCapGains30Per115BBH_8b = False
            Exit Function
        End If
    Next
End Function

'Change-20.04.2023.101.PL -> SIT-48882
Sub LOCKRANGEPL2(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))
           If chcell.MergeCells Then
           chcell.MergeArea.value = ""
           chcell.MergeArea.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Else
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           End If
           Next
       Next
End Sub

Sub UNLOCKRANGEPL2(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))

           If chcell.MergeCells Then
           chcell.MergeArea.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Else
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           End If

           Next
       Next
End Sub
Sub UNLOCKPL2()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_PL2 As Variant


'rngname_PL2 = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription||PL_61a||PL_61b||PL_61iia||PL_61iib"
'New change by Sai
rngname_PL2 = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription||PL_61a||PL_61a_1b||PL_61b||PL_61iia||PL_61iib"
UNLOCKRANGEPL2 (rngname_PL2)

'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKPL2()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_PL2 As Variant
'rngname_PL2 = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription||PL_61a||PL_61b||PL_61iia||PL_61iib"
'New change by Sai
rngname_PL2 = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription||PL_61a||PL_61a_1b||PL_61b||PL_61iia||PL_61iib"
LOCKRANGEPL2 (rngname_PL2)


'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub UNLOCKPL3()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_PL3 As Variant


'rngname_PL3 = "Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1||PL_62i||PL_62ii"
'New change by Sai
rngname_PL3 = "Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1||PL_62i_A||PL_62i_B||PL_62i_C||PL_62ii"
UNLOCKRANGEPL2 (rngname_PL3)

'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKPL3()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_PL3 As Variant
'rngname_PL3 = "Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1||PL_62i||PL_62ii"
'New change by Sai
rngname_PL3 = "Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1||PL_62i_A||PL_62i_B||PL_62i_C||PL_62ii"
LOCKRANGEPL2 (rngname_PL3)


'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub
'Change-04.05.2023.101. -> V0.4 Implementation
Sub STCGSecUndDedClaim(RowIndex As Long) 'Changed as per DE 0.2 by Bindu
'Target will be Range("STCG.A7_PreviousYer_1")
sPassword = EfilingCommon.getmsgstate
Sheet13.Unprotect Password:=sPassword

Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).value = ""

'Years drop down list updated by sai on 03/03/2025 for AY 2025-26
'Year change by Bindu on 11th Feb 26
If Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_PreviousYer_1").Column).value = "2022-23" Then
    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54G,54GA"
End If

'Year change by Bindu on 11th Feb 26

If Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_PreviousYer_1").Column).value = "2023-24" Or Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_PreviousYer_1").Column).value = "2024-25" Then
    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54B,54G,54GA"
End If

'If Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_PreviousYer_1").Column).value = "2020-21" Then
'    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).value = "54B"
'    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).Validation.Delete
'    Sheet13.Cells(RowIndex, Sheet13.Range("STCG.A7_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="54B"
'End If

Sheet13.Protect Password:=sPassword
End Sub

'Change-04.05.2023.101. -> V0.4 Implementation
Sub LTCGSecUndDedClaim(RowIndex As Long) 'Changes as per DE 0.2 by Bindu
'Target will be Range("STCG.A7_PreviousYer_1")
sPassword = EfilingCommon.getmsgstate
Sheet13.Unprotect Password:=sPassword

Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).value = ""

'Years drop down list updated by sai on 22/02/2025 for AY 2025-26
'Year updated by Bindu on 2021-22 to 2022-23 on 22nd Feb 26
If Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_PreviousYer_1").Column).value = "2022-23" Then
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54D,54F,54G,54GA"
End If

'Year updated by Bindu on 2022-23 to 2023-24 on 22nd Feb 26
If Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_PreviousYer_1").Column).value = "2023-24" Then
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.Delete
'    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54B,54D,54F,54G,54GA"
     Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54B,54D,54F,54G,54GA,54GB"
End If

'Year updated by Bindu on 2023-24 to 2024-25 on 22nd Feb 26
If Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_PreviousYer_1").Column).value = "2024-25" Then
    'Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).value = "54B"
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54B,54D,54F,54G,54GA"
End If

Sheet13.Protect Password:=sPassword
End Sub
'Change-04.05.2023.101. -> V0.4 Implementation
Function check80CCH_Eligible() As Boolean
check80CCH_Eligible = True


Dim blockcount, CGemployee
blockcount = 0
CGemployee = False
For blockcount = 1 To Sheet39.Range("SalarySectionCount").value
    If Sheet39.Range("sheet.EmployerCategory" & blockcount).value = "Central Government" Then
        CGemployee = True
        Exit For
    End If
Next blockcount


If (Not CGemployee) Or Dformat(Sheet1.Range("sheet1.DOB"), "") < "2001-03-31" Or Dformat(Sheet1.Range("sheet1.DOB"), "") > "2006-03-31" Then
    check80CCH_Eligible = False
End If

End Function

Sub CalcVIA80CCH()

On Error GoTo endline
Dim blockcount, salaryAsPer17_1, nat80CCH, VIA80CCH As Long
blockcount = 0
salaryAsPer17_1 = 0
nat80CCH = 0
For blockcount = 1 To Sheet39.Range("SalarySectionCount").value
    salaryAsPer17_1 = salaryAsPer17_1 + Sheet39.Range("SAL.AllowancesNotExempt" & blockcount).value
    nat80CCH = nat80CCH + Application.WorksheetFunction.SumIfs(Sheet39.Range("salary.SalaryAmount" & blockcount), Sheet39.Range("salary.Section" & blockcount), "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH")
Next blockcount

VIA80CCH = 0.6 * (salaryAsPer17_1 - nat80CCH)
Sheet19.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Sheet19.Range("Anyother80CCH_PreCalc").value = VIA80CCH

endline:
Sheet19.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
'------

Sub ImportScheduleTCSCSV()
'validate_TDS2 = True
Dim filePath, text As String

'fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."
Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    'msgbox_TDS3 = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        'Filter to just the following types of files to narrow down selection options
        .Filters.add "CSV File", "*.csv", 1
        'Show the dialog box
        If .Show = True Then
            filePath = .SelectedItems.item(1)
        Else
            Exit Sub
        End If
        On Error Resume Next
    End With
    
    Open filePath For Input As #1
    text = Input$(LOF(1), 1)
    Close #1
    
    importTCSCSVFunc (text)
End If
End Sub



Sub importTCSCSVFunc(text As String)
Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount As Variant
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(10, 1 To n)
        
        
            TwoDArray(0, n) = TmpAr(0)      'slno
            TwoDArray(1, n) = TmpAr(1)      'self/other
            TwoDArray(2, n) = TmpAr(2)      'TAN
            TwoDArray(3, n) = TmpAr(3)      'PAN
            TwoDArray(4, n) = TmpAr(4)      'b/f financial year
            TwoDArray(5, n) = TmpAr(5)      'b/f amount
            TwoDArray(6, n) = TmpAr(6)      'FY collected in own hands
            TwoDArray(7, n) = TmpAr(7)      'FY year collected other person hands
            TwoDArray(8, n) = TmpAr(8)      'this year claimed own hands
            TwoDArray(9, n) = TmpAr(9)      'this year collected other person hands TCS
            TwoDArray(10, n) = TmpAr(10)      'this year collected other person hands PAN
            
        End If
    Next i
   
    Dim SelfOthColNo, TANColNo, PANColNo, DeductedYrColNo, BroughtFwdTCSAmtColNo, FYCollOwnHandsColNo, FYCollOthHandsColNo
    Dim CurrYrClaimOwnHandColNo, CurrYrClaimOthHandTCSColNo, CurrYrClaimOthHandPANColNo, AmtCarriedFwdColNo
    
    SelfOthColNo = Sheet28.Range("TCS1.CreditName").Column
    TANColNo = Sheet28.Range("TCS1.TAN").Column
    PANColNo = Sheet28.Range("TCS1.PANOP").Column
    DeductedYrColNo = Sheet28.Range("TCS1.CollectYear").Column
    BroughtFwdTCSAmtColNo = Sheet28.Range("TCS1.UnclaimedBFAmount").Column
    FYCollOwnHandsColNo = Sheet28.Range("TCS1.CollectedOwnHands").Column
    FYCollOthHandsColNo = Sheet28.Range("TCS1.Collected37i").Column
    CurrYrClaimOwnHandColNo = Sheet28.Range("TCS1.ClaimedOwnHands").Column
    CurrYrClaimOthHandTCSColNo = Sheet28.Range("TCS1.ClaimedOPTCS").Column
    CurrYrClaimOthHandPANColNo = Sheet28.Range("TCS1.ClaimedOPPAN").Column
    AmtCarriedFwdColNo = Sheet28.Range("TCS1.CreditForward").Column
    
    
    TotalExRow = Sheet28.Range("TCS1.CreditName").Rows.count
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet28.Range("TCS1.CreditName").Locked = False Then
            Sheet28.Range("TCS1.CreditName").ClearContents
        End If
        If Sheet28.Range("TCS1.TAN").Locked = False Then
            Sheet28.Range("TCS1.TAN").ClearContents
        End If
        If Sheet28.Range("TCS1.PANOP").Locked = False Then
            Sheet28.Range("TCS1.PANOP").ClearContents
        End If
        If Sheet28.Range("TCS1.CollectYear").Locked = False Then
            Sheet28.Range("TCS1.CollectYear").ClearContents
        End If
        If Sheet28.Range("TCS1.UnclaimedBFAmount").Locked = False Then
            Sheet28.Range("TCS1.UnclaimedBFAmount").ClearContents
        End If
        If Sheet28.Range("TCS1.CollectedOwnHands").Locked = False Then
            Sheet28.Range("TCS1.CollectedOwnHands").ClearContents
        End If
        If Sheet28.Range("TCS1.Collected37i").Locked = False Then
            Sheet28.Range("TCS1.Collected37i").ClearContents
        End If
        If Sheet28.Range("TCS1.ClaimedOwnHands").Locked = False Then
            Sheet28.Range("TCS1.ClaimedOwnHands").ClearContents
        End If
        If Sheet28.Range("TCS1.ClaimedOPTCS").Locked = False Then
            Sheet28.Range("TCS1.ClaimedOPTCS").ClearContents
        End If
        If Sheet28.Range("TCS1.ClaimedOPPAN").Locked = False Then
            Sheet28.Range("TCS1.ClaimedOPPAN").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
        AddRows_TCS1_import (TotalDiffRow)
    End If
    rowcount = getRowNo(Sheet28.Range("TCS1.CreditName").name)
    
    For i = 1 To n
        If Sheet28.Cells(rowcount, SelfOthColNo).Locked = False Then
            If UCase(Mid(TwoDArray(0, i), 1, 1)) = "S" Then
                Sheet28.Cells(rowcount, SelfOthColNo).value = "Self"
            ElseIf UCase(Mid(TwoDArray(0, i), 1, 1)) = "O" Then
                Sheet28.Cells(rowcount, SelfOthColNo).value = "Other Person"
            End If
        End If
        
        If Sheet28.Cells(rowcount, TANColNo).Locked = False Then
            Sheet28.Cells(rowcount, TANColNo).value = TwoDArray(1, i)
        End If
        If Sheet28.Cells(rowcount, PANColNo).Locked = False Then
            Sheet28.Cells(rowcount, PANColNo).value = TwoDArray(2, i)
        End If
        If Sheet28.Cells(rowcount, DeductedYrColNo).Locked = False Then
            Sheet28.Cells(rowcount, DeductedYrColNo).value = TwoDArray(3, i)
        End If
        If Sheet28.Cells(rowcount, BroughtFwdTCSAmtColNo).Locked = False Then
            Sheet28.Cells(rowcount, BroughtFwdTCSAmtColNo).value = TwoDArray(4, i)
        End If
        If Sheet28.Cells(rowcount, FYCollOwnHandsColNo).Locked = False Then
            Sheet28.Cells(rowcount, FYCollOwnHandsColNo).value = TwoDArray(5, i)
        End If
        If Sheet28.Cells(rowcount, FYCollOthHandsColNo).Locked = False Then
            Sheet28.Cells(rowcount, FYCollOthHandsColNo).value = TwoDArray(6, i)
        End If
        If Sheet28.Cells(rowcount, CurrYrClaimOwnHandColNo).Locked = False Then
            Sheet28.Cells(rowcount, CurrYrClaimOwnHandColNo).value = TwoDArray(7, i)
        End If
        If Sheet28.Cells(rowcount, CurrYrClaimOthHandTCSColNo).Locked = False Then
            Sheet28.Cells(rowcount, CurrYrClaimOthHandTCSColNo).value = TwoDArray(8, i)
        End If
        If Sheet28.Cells(rowcount, CurrYrClaimOthHandPANColNo).Locked = False Then
            Sheet28.Cells(rowcount, CurrYrClaimOthHandPANColNo).value = TwoDArray(9, i)
        End If
        
        rowcount = rowcount + 1
                
    Next i
    
End Sub
'----end change---


'Change-25.07.2023.101 -> SIT-49875
Function ValidateTAN_TDS1WithSalary() As Boolean
ValidateTAN_TDS1WithSalary = True

Dim flagThisTAN, currTDSTAN, currSalTAN, salaryBlocksCount, i, j
salaryBlocksCount = Sheet39.Range("SalarySectionCount").value

For i = 1 To Sheet28.Range("TDS1.TAN").count
    flagThisTAN = False
    currTDSTAN = ""
    currTDSTAN = Sheet28.Range("TDS1.TAN").item(i, 1).value
    If (currTDSTAN <> "") Then
        For j = 1 To salaryBlocksCount
            currSalTAN = ""
            currSalTAN = Sheet39.Range("SAL.PANofEmployer" & j).value
            
            If currTDSTAN = currSalTAN Then
                flagThisTAN = True
                Exit For
            End If
            
        Next j
        
        If flagThisTAN = False Then
            msgbox_TDS1 = msgbox_TDS1 + "Please enter TAN of employer as per TAN mentioned in details of Salary" & Chr(13)
            ValidateTAN_TDS1WithSalary = False
            Exit Function
        End If
    End If
    
Next i
End Function
'----
'Change Satya, 102 , 30/10/2023 ver.0.5
'Sub LockUnlock_FIIFPI()
'    If Sheet1.Range("sheet1.ResidentialStatus1").value = "RES - Resident" Or Sheet1.Range("sheet1.ResidentialStatus1").value = "NOR - Resident but not Ordinarily Resident" Then
'        Sheet1.Unprotect Password:=getmsgstate
''        Application.EnableEvents = False
'
'        Sheet1.Range("Foreign_institutinal_investment").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Foreign_institutinal_investment").MergeArea.Locked = True
''        Application.EnableEvents = True
'        Sheet1.Range("Foreign_institutinal_investment").value = "No"
'        Sheet1.Protect Password:=getmsgstate
'    Else 'If Sheet1.Range("sheet1.ResidentialStatus1").value = "NRI - Non Resident" Then
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("Foreign_institutinal_investment").MergeArea.Locked = False
'        Sheet1.Range("Foreign_institutinal_investment").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Foreign_institutinal_investment").value = "(Select)"
'        Sheet1.Protect Password:=getmsgstate
'    End If
'End Sub
'End Change
'Change Satya, 102 , 31/10/2023 ver.0.6
Function ChkMinSystemdate2324(Systemdate2324 As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinSystemdate2324 = True
    If Len(Systemdate2324) > 0 Then
       If val(Mid(Systemdate2324, 7, 4)) < 2026 Then
           ChkMinSystemdate2324 = False
           Exit Function
       ElseIf (val(Mid(Systemdate2324, 7, 4)) = 2026 And val(Mid(Systemdate2324, 4, 2)) <= 3) Then
           ChkMinSystemdate2324 = False
           Exit Function
       End If
   End If
End Function


'Malli AY2024-25 New changes

Sub CalcVIA80CCH_NewRegme()
On Error GoTo endline
Dim blockcountN, salaryAsPer17_1N, nat80CCHN, VIA80CCHN As Long
blockcountN = 0
salaryAsPer17_1N = 0
nat80CCHN = 0
For blockcountN = 1 To Sheet39.Range("SalarySectionCount").value
salaryAsPer17_1N = salaryAsPer17_1N + Sheet39.Range("SAL.AllowancesNotExempt" & blockcountN).value

nat80CCHN = nat80CCHN + Application.WorksheetFunction.SumIfs(Sheet39.Range("salary.SalaryAmount" & blockcountN), Sheet39.Range("salary.Section" & blockcountN), "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH")

Next blockcountN

VIA80CCHN = salaryAsPer17_1N - nat80CCHN
'Newly changed by Bindu as per Ayushi
VIA80CCHN = nat80CCHN


Sheet19.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Sheet19.Range("Anyother80CCH_PreCalc_Newreg").value = VIA80CCHN
Application.EnableEvents = True
Sheet19.Protect Password:=getmsgstate
endline:
End Sub

 
