Attribute VB_Name = "Sheet57"
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
Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next

CalculateTaxCheckFlag = False

Application.EnableEvents = False
Dim DateRangecells  As Range
Dim dateCtr, x As Long

Set DateRangecells = Sheet57.Range("TPSC_DateDep").Cells
dateCtr = Sheet57.Range("TPSC_DateDep").count

If Not Intersect(Range("TPSC_DateDep"), Range(Target.Address)) Is Nothing Then
    For x = 1 To dateCtr
        If Not DateRangecells.item(x).value = "" Then
           If Not CheckDateBeforeinSheet_TPSA(DateRangecells.item(x).value) Then
              DateRangecells.item(x).value = ""
              DateRangecells.item(x).Select
           Else
              DateRangecells.item(x).value = CreditDate
           End If
        End If
    Next
End If

Application.EnableEvents = True
End Sub
Function CheckDateBeforeinSheet_TPSA(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2016, for A.Y. 2018-19

    CheckDateBeforeinSheet_TPSA = True
    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBeforeinSheet_TPSA = False
           'fmsgboxStatus ("Please enter a valid date in dd/mm/yyyy format in Schedule TPSA")
           fmsgboxStatus ("""Please enter a valid date in dd/mm/yyyy format in Schedule TPSA.""")
            Exit Function
        End If
'Change-10.01.2023.102.03A
        'AY_2024_25 Old 'Malli
        'If Not ChkMinDOBDate_TPSA(dob, "01/04/2022") Then
'        If Not ChkMinDOBDate_TPSA(dob, "01/04/2023") Then
        
'        If Not ChkMinDOBDate_TPSA(dob, "01/04/2024") Then 'Added by Aavula Naresh for A.Y 2025_26 Dev
        If Not ChkMinDOBDate_TPSA(dob, "01/04/2025") Then 'Added by Aavula Naresh for A.Y 2026-27
'            fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2021 in Schedule TPSA")
            'fmsgboxStatus ("Date of deposit cannot be prior to 01/04/2022 in Schedule TPSA")
'            fmsgboxStatus ("""Date of deposit cannot be prior to 01-04-2023 in Schedule TPSA.""")
            
            'Added by Aavula Naresh for A.Y 2025_26 Dev
'            fmsgboxStatus ("""Date of deposit cannot be prior to 01-04-2024 in Schedule TPSA.""")
           'Added by Aavula Naresh for A.Y 2026-27
            fmsgboxStatus ("""Date of deposit cannot be prior to 01-04-2025 in Schedule TPSA.""")
'            end---
            CheckDateBeforeinSheet_TPSA = False
            Exit Function
        Else
            CreditDate = dob
        End If
    End If
 End Function


 Function ChkMinDOBDate_TPSA(Mininclusive As Variant, Mininclusivedate As Variant)
   '' both date must be in format yyyy-mm-dd
    ChkMinDOBDate_TPSA = True
    If Len(Mininclusive) > 0 Then
    If Mid(Mininclusive, 7, 4) < Mid(Mininclusivedate, 7, 4) Then
        ChkMinDOBDate_TPSA = False
        Exit Function
    Else
        If Mid(Mininclusive, 7, 4) = Mid(Mininclusivedate, 7, 4) Then
            If (Mid(Mininclusive, 4, 2) < Mid(Mininclusivedate, 4, 2)) Then
                ChkMinDOBDate_TPSA = False
                Exit Function
            ElseIf ((Mid(Mininclusive, 4, 2) = Mid(Mininclusivedate, 4, 2))) Then
                If (Mid(Mininclusive, 1, 2) < Mid(Mininclusivedate, 1, 2)) Then
                    ChkMinDOBDate_TPSA = False
                    Exit Function
                End If
                End If
            End If
        End If
    End If

 End Function


