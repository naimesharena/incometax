Attribute VB_Name = "Sheet47"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()

MsgBox "*If the benefit of quarterly computation of interest u/s234C is sought, same may be claimed by making necessary modifications in Table F."

Sheet47.Protect Password:=getmsgstate


End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim i As Long

If Not Intersect(Sheet47.Range("VDA.DateOfAcq"), Sheet47.Range(Target.Address)) Is Nothing Then
    For i = 1 To Sheet47.Range("VDA.DateOfAcq").count
        If Not Target.value = "" Then
           If Not CheckDateAcqinSheet_VDA(Target.value) Then
              Target.value = ""
              Target.Select
           End If
        End If
    Next
End If

If Not Intersect(Sheet47.Range("VDA.DateOfTrans"), Sheet47.Range(Target.Address)) Is Nothing Then
    For i = 1 To Sheet47.Range("VDA.DateOfTrans").count
        If Not Target.value = "" Then
           If Not CheckDateTrinSheet_VDA(Target.value) Then
              Target.value = ""
              Target.Select
           End If
           'VDA_E1_AY_2023_24 old
           'If Dformat(Target.value, "") < Dformat(Sheet47.Range(Replace(Target.Address, "E", "D")).value, "") Then
           'VDA_E1_AY_2023_24 old 'Malli
           'If Dformat(Target.value, "") < Dformat(Sheet47.Range(Replace(Target.Address, "F", "E")).value, "") Then
              'Added by Shrutika(28-02-2025)
           If Dformat(Target.value, "") < Dformat(Sheet47.Range(Replace(Target.Address, "E", "D")).value, "") Then
                MsgBox ("Date of Transfer cannot be before Date of Acquisition")
                Target.value = ""
                Target.Select
           End If
        End If
    Next
End If

'added by Chetan C M on 17/10/2025 for SIT-89185
'start--
If Not Intersect(Sheet47.Range("VDA.DateOfAcq"), Sheet47.Range(Target.Address)) Is Nothing Then
Sheet47.Range(Replace(Target.Address, "D", "F")).value = "Capital Gain"
End If

If Not Intersect(Sheet47.Range("VDA.DateOfTrans"), Sheet47.Range(Target.Address)) Is Nothing Then
Sheet47.Range(Replace(Target.Address, "E", "F")).value = "Capital Gain"
End If

If Not Intersect(Sheet47.Range("VDA.CostOfAcq"), Sheet47.Range(Target.Address)) Is Nothing Then
Sheet47.Range(Replace(Target.Address, "G", "F")).value = "Capital Gain"
End If

If Not Intersect(Sheet47.Range("VDA.ConsiderationReceived"), Sheet47.Range(Target.Address)) Is Nothing Then
Sheet47.Range(Replace(Target.Address, "H", "F")).value = "Capital Gain"
End If
'--end


endline:
Application.EnableEvents = True
Sheet47.Protect Password:=getmsgstate
End Sub

Function CheckDateAcqinSheet_VDA(dob As Variant) As Boolean
On Error Resume Next
CheckDateAcqinSheet_VDA = True

    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDateAcqinSheet_VDA = False
            MsgBox ("Date of Acquisition must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        'VDA_E2_AY_2023_24 old
        'If Dformat(dob, "") > Dformat("31/03/2023", "") Then
        
        'VDA_C2_AY_2024_25  Change
        'If Dformat(dob, "") > Dformat("31/03/2024", "") Then
        
        'changed by Chetan C M on 21/01/2026 for AY 2026-27
        'start--
'     ' Added by Shrutika(AY-25)
'            If Dformat(dob, "") > Dformat("31/03/2025", "") Then
'            'MsgBox ("Date of Acquisition cannot be after 31/03/2024 for AY 24-25")
'            MsgBox ("Date of Acquisition cannot be after 31/03/2025 for AY 25-26")
            If Dformat(dob, "") > Dformat("31/03/2026", "") Then
            MsgBox ("Date of Acquisition cannot be after 31/03/2026 for AY 26-27")
        '--end
            CheckDateAcqinSheet_VDA = False
            Exit Function
        End If
    End If
 End Function

Function CheckDateTrinSheet_VDA(dob As Variant) As Boolean
On Error Resume Next
CheckDateTrinSheet_VDA = True

    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDateTrinSheet_VDA = False
            MsgBox ("Date of Transfer must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        'VDA_E3_AY_2023_24 old
        'If Dformat(dob, "") > Dformat("31/03/2023", "") Then
        'MsgBox ("Date of transfer cannot be after 31/03/2023 for AY 23-24")
        'VDA_C3_AY_2024_25  Change 'Malli
        'If Dformat(dob, "") > Dformat("31/03/2024", "") Then
        
        'changed by Chetan C M on 21/01/2026 for AY 2026-27
        'start--
'        'Added by Shrutika(AY-25)
'        If Dformat(dob, "") > Dformat("31/03/2025", "") Then
'            'MsgBox ("Date of transfer cannot be after 31/03/2024 for AY 24-25")
'            MsgBox ("Date of transfer cannot be after 31/03/2025 for AY 25-26")
        If Dformat(dob, "") > Dformat("31/03/2026", "") Then
            MsgBox ("Date of transfer cannot be after 31/03/2026 for AY 26-27")
        '-end
            CheckDateTrinSheet_VDA = False
            Exit Function
        End If
    End If
 End Function


