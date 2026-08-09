Attribute VB_Name = "Sheet98"
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
Sheet25.Shapes("Button 7").Top = Sheet25.Range("Button_IT").Top + 5
Sheet25.Shapes("Button 7").Left = Sheet25.Range("Button_IT").Left + 3
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
Application.EnableEvents = False
Dim DateRangecells  As Range
Dim dateCtr, x As Long

Set DateRangecells = Sheet25.Range("IT.DateDep").Cells
dateCtr = Sheet25.Range("IT.DateDep").count

If Not Intersect(Range("IT.DateDep"), Range(Target.Address)) Is Nothing Then
    For x = 1 To dateCtr
        If Not DateRangecells.item(x).value = "" Then
           If Not CheckDateBeforeinSheet_IT(DateRangecells.item(x).value) Then
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



Function CheckDateBeforeinSheet_IT(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2018, for A.Y. 2018-19

    CheckDateBeforeinSheet_IT = True
    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBeforeinSheet_IT = False
           fmsgboxStatus ("Date of Credit to Govt. must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
        If Not ChkMinDOBDate_IT(dob, "01/04/2021") Then
            fmsgboxStatus ("Date of Credit to Govt. should be on or after 01/04/2021")
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
 
 
 
