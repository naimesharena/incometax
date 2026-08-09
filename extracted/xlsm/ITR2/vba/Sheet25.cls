Attribute VB_Name = "Sheet25"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String
Dim CreditDate As String

Private Sub Worksheet_Activate()
On Error Resume Next
Sheet25.Shapes("Button 16").Top = Sheet25.Range("Button_IT").Top + 5
Sheet25.Shapes("Button 16").Left = Sheet25.Range("Button_IT").Left + 3
Sheet25.Protect Password:=getmsgstate

If Mid(Sheet1.Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
'MsgBox " If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax / credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT-2 .Please be noted that credit will be allowed only once. Also, REGULAR ASSESSMENT TAX if any paid needs to be disclosed in sl.no.8 of Part B_ATI and schedule IT-2 at sl.no.14B"
'Shrutika_Ur
 MsgBox "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax / credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also in Schedule IT-2. Please be noted that credit will be allowed only once." & vbNewLine & "Also, ""REGULAR ASSESSMENT TAX"" if any paid needs to be disclosed in sl.no.8 of Part B_ATI and schedule IT-2 at sl.no.14B", , "ITR-2"
'---------------
MsgBox " If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments in Schedule IT-1", , "ITR-2"
'MsgBox "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax, credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also at Sl. No. 14B of Part B-ATI. Please be noted that credit will be allowed only once. Also, REGULAR ASSESSMENT TAX if any paid needs to be disclosed at Sl. No. 8 & Sl. No. 14B of Part B-ATI. If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments at Sl. No. 14A Part B-ATI "
End If


End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next

CalculateTaxCheckFlag = False

Application.EnableEvents = False
Dim DateRangecells  As Range
Dim dateCtr, x As Long


          
If Not Intersect(Range("IT.SrlNoOfChaln"), Range(Target.Address)) Is Nothing Then
    If Not chkNumeric(Target.value) Then
        MsgBox "Please enter 5 digit serial number of challan in schedule IT"
        Target.value = ""
    End If

End If


Set DateRangecells = Sheet25.Range("IT.DateDep").Cells
dateCtr = Sheet25.Range("IT.DateDep").count


If Not Intersect(Range("IT.BSRCode"), Range(Target.Address)) Is Nothing Then
Range(Target.Address).value = UCase(Range(Target.Address).value)
End If

If Not Application.Intersect(Sheet25.Range("IT.DateDep"), Sheet25.Range(Target.Address)) Is Nothing Then
       If Not (ValidateDateDep_IT()) Then
'            MsgBox "INVALID Date of Deposit"
       End If
       GoTo endline
   End If



'If Not Intersect(Range("IT.DateDep"), Range(Target.Address)) Is Nothing Then
'    For x = 1 To dateCtr
'        If Not DateRangecells.item(x).value = "" Then
'           If Not CheckDateBeforeinSheet_IT(DateRangecells.item(x).value) Then
'              DateRangecells.item(x).value = ""
'              DateRangecells.item(x).Select
'           Else
'              DateRangecells.item(x).value = CreditDate
'           End If
'        End If
'    Next
'End If
endline:
Application.EnableEvents = True
End Sub
Function ValidateDateDep_IT() As Boolean

    Dim rangecells As Range
    Dim i As Long
    Dim DateDep_IT
    Dim mIntCells As Long
    Dim mIntCtr As Long

    mIntCells = Sheet25.Range("IT.DateDep").count
       ValidateDateDep_IT = True
       Dim end_IT
       end_IT = 0
       Set rangecells = Sheet25.Range("IT.DateDep").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_IT = end_IT + 1
        End If

Next
   
   ReDim DateDep_IT(end_IT)
       For i = 1 To end_IT
           DateDep_IT(i) = rangecells.item(i).value
           If Not chkCompulsory(DateDep_IT(i)) Then
           'Added by Aavula
               fmsgboxsmall ("Please enter date of deposit in Schedule IT at Sr. No  " & i & "")
               ValidateDateDep_IT = False
               Exit Function
           End If
        
            If Not CheckDateddmmyyyy(DateDep_IT(i)) Then
               ValidateDateDep_IT = False
               'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
                
                'Added by Aavula
               fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format in Schedule IT at Sr. No  " & i & "")
              rangecells.item(i).value = ""
               Exit Function
           Else
           Dim IT_date_deposit As Date
            IT_date_deposit = CDate(SheetALL.Range("IT_Date").value)
               DateDep_IT(i) = Dformat(DateDep_IT(i), "")
               'Change-28.11.2022.104.06
                   'If Not ChkMinInclusiveDate(DateDep_IT(i), "2021-04-01") Then
                   'IT_E2_AY_2023-24 old
                   'If Not ChkMinInclusiveDate(DateDep_IT(i), "2022-04-01") Then
                   'IT_E2_AY_2023-24 Change 'Malli
'                   If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then 'dpk1601
                   'Added by Aavula New Dev 2025-26
'                   If Not ChkMinInclusiveDate(DateDep_IT(i), "2024-04-01") Then
                    If Not ChkMinInclusiveDate(DateDep_IT(i), Dformat4(SheetALL.Range("IT_Date").value, "yyyy-mm-dd")) Then
                   'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  should not be less than 01/04/2021")
                       'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  should not be less than 01/04/2023") 'dpk1601
                       
                       'Modify by Shrutiksa
                       'fmsgboxsmall """Date of deposit cannot be prior to  01/04/2022 in Schedule IT at Sr. No  " & i & """"
                       
'                       'Added by Aavula
'                       fmsgboxsmall """Date of deposit cannot be prior to  01/04/2023 in Schedule IT at Sr. No  " & i & """"
                       
                       'Added by Aavula New Dev 2025-26
                       fmsgboxsmall """Date of deposit cannot be prior to " & Dformat5(IT_date_deposit, "dd/mm/yyyy") & " in Schedule IT at Sr. No  " & i & """"

                       rangecells.item(i).value = ""
                       ValidateDateDep_IT = False
                       Exit Function
                   End If
            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
            
                'Malli-----------30/10/2024
                ' If ChkMinInclusiveDate(DateDep_IT(i), "2023-01-01") Then
                 'If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
'                 If ChkMinInclusiveDate(DateDep_IT(i), "2026-01-01") Then 'Commented by Shrutika confirm with bhoomija 19/11/2025
'                 '---------------------------
'                  fmsgboxsmall ("*Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
'                End If
            'End If
             
             'Malli-----30/10/2024
            ' If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
'             If ChkMinInclusiveDate(DateDep_IT(i), "2026-01-01") Then 'Added by Shrutika 18/11/2025 SIT-103777

              If ChkMinInclusiveDate(DateDep_IT(i), "2027-01-01") Then 'Added by Naresh 04/03/2026 SIT-112893
              
               '   fmsgboxsmall ("* ""Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.""")
              'Added by Shrutika(28/02/2025)
                 ' fmsgboxsmall ("* ""Date of deposit can't be after 31-12-2025 in schedule IT. Please enter payments made after 31st December 2025 in schedule IT 1.""")
                'Added by Shrutika(04/03/2025)
'                fmsgboxsmall "* ""If filing section is selected as 139(8A), then date of deposit cant be after 31.12.2025 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""" & Chr(13)
                
                fmsgboxsmall "* ""If filing section is selected as 139(8A), then date of deposit cant be after 31.12.2026 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""" & Chr(13)
                   rangecells.item(i).value = ""
                   ValidateDateDep_IT = False
                   Exit Function
             End If
             End If 'malli
             '--------------------

'            If DateDep_IT(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgboxsmall "* Date of Deposit at Sr. No  " & i & " in Sheet IT can not be more than system date in Income Details"
'
'                            rangecells.item(i).value = ""
'
'                            ValidateDateDep_IT = False
'
'                Exit Function
'
'                     End If
'
        End If
   Next



End Function
Function CheckDateBeforeinSheet_IT(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2016, for A.Y. 2018-19

    CheckDateBeforeinSheet_IT = True
    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBeforeinSheet_IT = False
           fmsgboxStatus ("Date of Credit to Govt. must be a valid dd/mm/yyyy format")
            Exit Function
        End If

        If Not ChkMinDOBDate_IT(dob, "01/04/2021") Then
            fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2021 in Schedule IT")
            CheckDateBeforeinSheet_IT = False
            Exit Function
        Else
            CreditDate = dob
        End If
    End If
 End Function


 Function ChkMinDOBDate_IT(Mininclusive As Variant, Mininclusivedate As Variant)
   '' both date must be in format yyyy-mm-dd
    ChkMinDOBDate_IT = True
    If Len(Mininclusive) > 0 Then
    If Mid(Mininclusive, 7, 4) < Mid(Mininclusivedate, 7, 4) Then
        ChkMinDOBDate_IT = False
        Exit Function
    Else
        If Mid(Mininclusive, 7, 4) = Mid(Mininclusivedate, 7, 4) Then
            If (Mid(Mininclusive, 4, 2) < Mid(Mininclusivedate, 4, 2)) Then
                ChkMinDOBDate_IT = False
                Exit Function
            ElseIf ((Mid(Mininclusive, 4, 2) = Mid(Mininclusivedate, 4, 2))) Then
                If (Mid(Mininclusive, 1, 2) < Mid(Mininclusivedate, 1, 2)) Then
                    ChkMinDOBDate_IT = False
                    Exit Function
                End If
                End If
            End If
        End If
    End If

 End Function

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

