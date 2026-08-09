Attribute VB_Name = "Sheet47"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error Resume Next
    ActiveSheet.Shapes("Button 1").Left = Range("Button_VDA").Left + 2
    ActiveSheet.Shapes("Button 1").Top = Range("Button_VDA").Top + 2
    
For Each cell In Sheet53.Range("VDA_Date_of_Acquisition")

 If cell.Locked = False Then
'      MsgBox "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F."
       fmsgbox "If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F."
    Exit Sub
 End If
    
Next
    ThisWorkbook.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim i As Long

If Not Intersect(Sheet47.Range("VDA_Date_of_Acquisition"), Sheet47.Range(Target.Address)) Is Nothing Then
    For i = 1 To Sheet47.Range("VDA_Date_of_Acquisition").count
        If Not Target.value = "" Then
           If Not CheckDateAcqinSheet_VDA(Target.value) Then
              Target.value = ""
              Target.Select
            Else
            
           End If
        End If
    Next
End If

If Not Intersect(Sheet47.Range("VDA_Date_of_Transfer"), Sheet47.Range(Target.Address)) Is Nothing Then
    For i = 1 To Sheet47.Range("VDA_Date_of_Transfer").count
        If Not Target.value = "" Then
           If Not CheckDateTrinSheet_VDA(Target.value) Then
              Target.value = ""
              Target.Select
           End If
            
           If Dformat(Target.value, "") < Dformat(Sheet47.Range(Replace(Target.Address, "E", "D")).value, "") Then
'                MsgBox ("Date of Transfer cannot be before Date of Acquisition")
                 fmsgbox "Date of Transfer cannot be before Date of Acquisition."
                Target.value = ""
                Target.Select
           End If
        End If
    Next
End If

endline:
Application.EnableEvents = True
Sheet47.Protect Password:=getmsgstate
End Sub

Function CheckDateAcqinSheet_VDA(dob As Variant) As Boolean
On Error Resume Next
CheckDateAcqinSheet_VDA = True

    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDateVDA(dob) Then
            CheckDateAcqinSheet_VDA = False
'            MsgBox ("Date of Acquisition must be a valid dd/mm/yyyy format")
             fmsgbox "Date of Acquisition must be a valid dd/mm/yyyy format."
            Exit Function
        End If
        
'             If Dformat(dob, "") > Dformat("31/03/2025", "") Then 'Added by Aavula Naresh for A.Y 2025_26 Dev
              If Dformat(dob, "") > Dformat("31/03/2026", "") Then 'Added by Aavula Naresh for A.Y 2025_26 Dev
'Added by Aavula Naresh for 2026-27
'            MsgBox ("Date of Acquisition cannot be after 31/03/2026 for AY 26-27")
             fmsgbox "Date of Acquisition cannot be after 31/03/2026 for AY 26-27."
            CheckDateAcqinSheet_VDA = False
            Exit Function
        End If
    End If
 End Function

Function CheckDateTrinSheet_VDA(dob As Variant) As Boolean
On Error Resume Next
CheckDateTrinSheet_VDA = True

    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDateVDA(dob) Then
            CheckDateTrinSheet_VDA = False
'            MsgBox ("Date of Transfer must be a valid dd/mm/yyyy format")
             fmsgbox "Date of Transfer must be a valid dd/mm/yyyy format."
            Exit Function
        End If

'        If Dformat(dob, "") > Dformat("31/03/2025", "") Then 'Added by Aavula Naresh for A.Y 2025_26 Dev
         If Dformat(dob, "") > Dformat("31/03/2026", "") Then 'Added by Aavula Naresh for A.Y 2026-27
'Added by Aavula Naresh for 2026-27
            'MsgBox ("Date of transfer cannot be after 31/03/2026 for AY 26-27")
             fmsgbox "Date of transfer cannot be after 31/03/2026 for AY 26-27."
             CheckDateTrinSheet_VDA = False
            Exit Function
        End If
    End If
 End Function

