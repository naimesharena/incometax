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


'Below msgs are Commented by Satyaprakash as Per Defect Lof named "ITR3_AY23_24_DefectLog_04_04_2023" on 5th April 2023


'MsgBox "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax /  Regular Assessment Tax, credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also in separately in Schedule IT-2 as the credit for the same is not to be allowed under section 140B (2). Please be noted that credit will be allowed only once. Also, ""REGULAR ASSESSMENT TAX"" if any paid needs to be disclosed in sl.no.8 of Part B_ATI and schedule IT-2 at sl.no.14B"
'MsgBox "If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments in Schedule IT-1"

'suhda_2025-26

If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then 'added this line by sudha
'End Change
'Change 102 28.11.2023 Ver 0.6 ScheduleIT
MsgBox "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax, credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also  in Schedule IT-2 Please be noted that credit will be allowed only once.       Also, ""REGULAR ASSESSMENT TAX"" if any paid needs to be disclosed in sl.no.8 of Part B_ATI and schedule IT-2 at sl.no.14B"

MsgBox "If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments in Schedule IT-1"
'End Change Ver 0.6 ScheduleIT
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
       End If
       GoTo endline
End If

'If Not Intersect(Range("IT.DateDep"), Range(target.Address)) Is Nothing Then
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



Function CheckDateBeforeinSheet_IT(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2016, for A.Y. 2018-19

    CheckDateBeforeinSheet_IT = True
    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBeforeinSheet_IT = False
           fmsgboxStatus ("Please enter a valid date in dd/mm/yyyy format in Schedule IT")
'            fmsgboxStatus ("Please enter a valid date in dd/mm/yyyy format")
            Exit Function
        End If
'Change-.3.11.2023.102.Ver.0.6
'AY_2024_25 Change Malli
        'If ChkMinDOBDate_IT(dob, "01/01/2024") Then
        
        
'        If ChkMinDOBDate_IT(dob, "01/01/2025") Then
'         If ChkMinDOBDate_IT(dob, "16/01/2025") Then ' Changed by sai as per BA on 06/01/2025




       If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then

'         If ChkMinDOBDate_IT(dob, "01/01/2026") Then ' Added by Aavula Naresh
         If ChkMinDOBDate_IT(dob, "31/12/2026") Then ' Added by Aavula Naresh

        'MsgBox ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
   'Added by Aavula
'         MsgBox ("Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1")
      ' Changed by sai as per BA on 06/01/2025
'          MsgBox ("""Date of deposit can't be after 15-01-2025 in schedule IT. Please enter payments made after 15th January 2025 in schedule IT 1.""")
            
            ' Added by Aavula Naresh
'            MsgBox ("""If filing section is selected as 139(8A), then date of deposit cant be after 31.12.2025 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""")
            MsgBox ("""If filing section is selected as 139(8A), then date of deposit cant be after 31.12.2026 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""")
            CheckDateBeforeinSheet_IT = False
            Exit Function
End If

        End If
''End Change


        'AY_2024_25 Change 'Malli
        'If Not ChkMinDOBDate_IT(dob, "01/04/2022") Then
'         If Not ChkMinDOBDate_IT(dob, "01/04/2023") Then
'         If Not ChkMinDOBDate_IT(dob, "01/04/2024") Then 'Added by Aavula Naresh New Dev
         If Not ChkMinDOBDate_IT(dob, "01/04/2025") Then 'Added by Aavula Naresh 2026-27
         
'            fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2021 in Schedule IT")
            'fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2022 in Schedule IT")
'             fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2023 in Schedule IT")
'             fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2024 in Schedule IT") 'Added by Aavula Naresh New Dev
             fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2025 in Schedule IT") 'Added by Aavula Naresh 2026-27
             
             
            CheckDateBeforeinSheet_IT = False
            Exit Function
            
'Change Satya, 102 , 30/10/2023 ver.0.5
'        ElseIf Dformat(dob, "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
'            fmsgboxStatus ("Date of Deposit can not be beyond system date in Schedule IT")
'            CheckDateBeforeinSheet_IT = False
'            Exit Function
 'End Change
        Else
            CreditDate = dob
        End If
        'AY_2024_25 Change 'Malli
        'If Not ChkMinDOBDate_IT(dob, "01/04/2022") Then
        If Not ChkMinDOBDate_IT(dob, "01/04/2023") Then
        MsgBox ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
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
 
 
Function Dformat4(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("IT_DateofDep_26_27").value) > 0 Then
 
Year = Mid(SheetALL.Range("IT_DateofDep_26_27").value, 7, 4)
month = Mid(SheetALL.Range("IT_DateofDep_26_27").value, 4, 2)
day = Mid(SheetALL.Range("IT_DateofDep_26_27").value, 1, 2)
formateddate = Year & "-" & month & "-" & day
Dformat4 = formateddate
 
Else
Dformat4 = ""
End If
End Function


Function Dformat5(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("IT_DateofDep_26_27").value) > 0 Then
 
Year = Mid(SheetALL.Range("IT_DateofDep_26_27").value, 7, 4)
month = Mid(SheetALL.Range("IT_DateofDep_26_27").value, 4, 2)
day = Mid(SheetALL.Range("IT_DateofDep_26_27").value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat5 = formateddate
 
Else
Dformat5 = ""
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
Sub test()

ActiveSheet.Columns("K:XFD").Hidden = True

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
               fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format in Schedule IT at Sr. No  " & i & "")
              rangecells.item(i).value = ""
               Exit Function
           Else
           
           
           Dim IT_date_deposit As Date
            IT_date_deposit = CDate(SheetALL.Range("IT_DateofDep_26_27").value)
               DateDep_IT(i) = Dformat(DateDep_IT(i), "")
                    If Not ChkMinInclusiveDate(DateDep_IT(i), Dformat4(SheetALL.Range("IT_DateofDep_26_27").value, "yyyy-mm-dd")) Then

               fmsgboxsmall (" ""Date of deposit cannot be prior to " & Dformat5(IT_date_deposit, "dd/mm/yyyy") & " in Schedule IT"" at Sr. No  " & i & "")

                       rangecells.item(i).value = ""
                       ValidateDateDep_IT = False
                       Exit Function
                   End If
                   
            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
            
             If ChkMinInclusiveDate(DateDep_IT(i), "2027-04-01") Then

                fmsgboxsmall "* ""If filing section is selected as 139(8A), then date of deposit cant be after 31.03.2027 in schedule IT. Please disclose payments u/s 140B in schedule IT-1""" & Chr(13)
  
                   rangecells.item(i).value = ""
                   ValidateDateDep_IT = False
                   Exit Function
             End If
             End If

        End If
   Next



End Function

