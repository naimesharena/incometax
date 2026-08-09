Attribute VB_Name = "Sheet16"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
 Sheet16.Protect Password:=getmsgstate

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim validateType As String
validateType = ""
Dim sourceSheet As Worksheet
Set sourceSheet = ThisWorkbook.Sheets("80D")

Application.ScreenUpdating = False
Application.EnableEvents = False


If (Target.name.name) = "DropDown_ValueOf_FamilyM_80D" Then
    If UCase(Target.value) = "YES" Then
        If (mdCommon.calculateAge(Sheet1.Range("sheet1.DOB").value) < 60) Then
        'Ayush_Legacy10
'            MsgBox ("You are not a senior citizen as per the date of birth. Please chose Yes option only if your spouse is a senior citizen")
            MsgBox ("You are not a senior citizen as per the date of birth. Please choose Yes option only if your spouse is a senior citizen")
        End If
    End If
End If
             

If (Target.name.name) = "Preventive_Health_80D" Then
    If (Range("Preventive_Health_80D").value + Range("Preventive_Health_SC_80D").value + Range("Preventive_Health2_80D").value + Range("Preventive_Health3_80D").value) > 5000 Then
        
        
       ' MsgBox "The amount of preventive health check up of all the fields combined together should not exceed 5000"
       
        MsgBox "Amount of preventive health check up claimed for self and family and/or parents cannot exceed Rs. 5000"
        
        Range("Preventive_Health_80D").value = ""
    End If
End If

If (Target.name.name) = "Preventive_Health_SC_80D" Then
    If (Range("Preventive_Health_80D").value + Range("Preventive_Health_SC_80D").value + Range("Preventive_Health2_80D").value + Range("Preventive_Health3_80D").value) > 5000 Then
        'Error Message Changed By Jyoti for the SIT-67661 on 25/06/2024
        
        'MsgBox "The amount of preventive health check up of all the fields combined together should not exceed 5000"
        MsgBox "Amount of preventive health check up claimed for self and family and/or parents cannot exceed Rs. 5000"
        
        
        Range("Preventive_Health_SC_80D").value = ""
    End If
End If

If (Target.name.name) = "Preventive_Health2_80D" Then
    If (Range("Preventive_Health_80D").value + Range("Preventive_Health_SC_80D").value + Range("Preventive_Health2_80D").value + Range("Preventive_Health3_80D").value) > 5000 Then
        'MsgBox "The amount of preventive health check up of all the fields combined together should not exceed 5000"
        MsgBox "Amount of preventive health check up claimed for self and family and/or parents cannot exceed Rs. 5000"
        Range("Preventive_Health2_80D").value = ""
    End If
End If

If (Target.name.name) = "Preventive_Health3_80D" Then
    If (Range("Preventive_Health_80D").value + Range("Preventive_Health_SC_80D").value + Range("Preventive_Health2_80D").value + Range("Preventive_Health3_80D").value) > 5000 Then
        'MsgBox "The amount of preventive health check up of all the fields combined together should not exceed 5000"
        MsgBox "Amount of preventive health check up claimed for self and family and/or parents cannot exceed Rs. 5000"
        Range("Preventive_Health3_80D").value = ""
    End If
End If



If (Target.name.name) = "DropDown_ValueOf_FamilyM_80D" Then
   Sheet16.Unprotect Password:=getmsgstate

If (Mid(Range("DropDown_ValueOf_FamilyM_80D").value, 1, 1) = "Y") Then

        Range("Senior_Citizen_80D").Locked = True
        Range("Senior_Citizen_80D").Interior.Color = 13434879
'        Range("Senior_Citizen_80D").Font.Color = RGB(0, 0, 0)
        Range("Senior_Citizen_80D").Formula = "=MIN(50000,SUM(Health_InsuranceSC_80D,Preventive_Health_SC_80D,Medical_Expenditure_SC_80D))"
'Ayush_commented
'        Range("Health_InsuranceSC_80D").Locked = False
'        Range("Health_InsuranceSC_80D").Interior.Color = (&HCCFFCC)
'Ayush_Added
        Range("NameInsurerB1.80D").Locked = False
        Range("NameInsurerB1.80D").Interior.Color = (&HCCFFCC)
        
        Range("PolicyNumB1.80D").Locked = False
        Range("PolicyNumB1.80D").Interior.Color = (&HCCFFCC)
        
 'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumB1.80D").Locked = False
'        Range("ReceiptNumB1.80D").Interior.Color = (&HCCFFCC)
        
        Range("AmtB1.80D").Locked = False
        Range("AmtB1.80D").Interior.Color = (&HCCFFCC)
        '----------------

        Range("Preventive_Health_SC_80D").Locked = False
        Range("Preventive_Health_SC_80D").Interior.Color = (&HCCFFCC)

        Range("Medical_Expenditure_SC_80D").Locked = False
        Range("Medical_Expenditure_SC_80D").Interior.Color = (&HCCFFCC)

        Range("Self_And_Family_80D").ClearContents
        Range("Self_And_Family_80D").Locked = True
        Range("Self_And_Family_80D").Interior.Color = (&HD8D8D8)
'Ayush_commented
'        Range("Health_Insurance_80D").ClearContents
'        Range("Health_Insurance_80D").Locked = True
'        Range("Health_Insurance_80D").Interior.Color = (&HD8D8D8)

'Ayush_Added
        Range("NameInsurerA1.80D").ClearContents
       Range("NameInsurerA1.80D").Locked = True
        Range("NameInsurerA1.80D").Interior.Color = (&HD8D8D8)
        
        Range("PolicyNumA1.80D").ClearContents
       Range("PolicyNumA1.80D").Locked = True
        Range("PolicyNumA1.80D").Interior.Color = (&HD8D8D8)
        
'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumA1.80D").ClearContents
'       Range("ReceiptNumA1.80D").Locked = True
'        Range("ReceiptNumA1.80D").Interior.Color = (&HD8D8D8)
        
        Range("AmtA1.80D").ClearContents
       Range("AmtA1.80D").Locked = True
        Range("AmtA1.80D").Interior.Color = (&HD8D8D8)
        '-------------------

        Range("Preventive_Health_80D").ClearContents
        Range("Preventive_Health_80D").Locked = True
        Range("Preventive_Health_80D").Interior.Color = (&HD8D8D8)

ElseIf (Mid(Range("DropDown_ValueOf_FamilyM_80D").value, 1, 1) = "N") Then

        Range("Self_And_Family_80D").Locked = True
        Range("Self_And_Family_80D").Interior.Color = 13434879
'        Range("Self_And_Family_80D").Font.Color = RGB(0, 0, 0)
        Range("Self_And_Family_80D").Formula = "=MIN(25000,SUM(Health_Insurance_80D,Preventive_Health_80D))"
'Ayush_commented
'        Range("Health_Insurance_80D").Locked = False
'        Range("Health_Insurance_80D").Interior.Color = (&HCCFFCC)
'Ayush_added
        Range("NameInsurerA1.80D").Locked = False
        Range("NameInsurerA1.80D").Interior.Color = (&HCCFFCC)
        
         Range("PolicyNumA1.80D").Locked = False
        Range("PolicyNumA1.80D").Interior.Color = (&HCCFFCC)
        
        
 'Commented by Ayush_06/05/2025 for 0.7 changes
'         Range("ReceiptNumA1.80D").Locked = False
'        Range("ReceiptNumA1.80D").Interior.Color = (&HCCFFCC)
        
         Range("AmtA1.80D").Locked = False
        Range("AmtA1.80D").Interior.Color = (&HCCFFCC)
        '----------------

        Range("Preventive_Health_80D").Locked = False
        Range("Preventive_Health_80D").Interior.Color = (&HCCFFCC)


        Range("Senior_Citizen_80D").ClearContents
        Range("Senior_Citizen_80D").Locked = True
        Range("Senior_Citizen_80D").Interior.Color = (&HD8D8D8)
'Ayush_commented
'        Range("Health_InsuranceSC_80D").ClearContents
'        Range("Health_InsuranceSC_80D").Locked = True
'        Range("Health_InsuranceSC_80D").Interior.Color = (&HD8D8D8)

'Ayush_added
        Range("NameInsurerB1.80D").ClearContents
        Range("NameInsurerB1.80D").Locked = True
        Range("NameInsurerB1.80D").Interior.Color = (&HD8D8D8)
        
         Range("PolicyNumB1.80D").ClearContents
        Range("PolicyNumB1.80D").Locked = True
        Range("PolicyNumB1.80D").Interior.Color = (&HD8D8D8)
        
 'Commented by Ayush_06/05/2025 for 0.7 changes
'         Range("ReceiptNumB1.80D").ClearContents
'        Range("ReceiptNumB1.80D").Locked = True
'        Range("ReceiptNumB1.80D").Interior.Color = (&HD8D8D8)
'
         Range("AmtB1.80D").ClearContents
        Range("AmtB1.80D").Locked = True
        Range("AmtB1.80D").Interior.Color = (&HD8D8D8)
        
    '----------------------

        Range("Preventive_Health_SC_80D").ClearContents
        Range("Preventive_Health_SC_80D").Locked = True
        Range("Preventive_Health_SC_80D").Interior.Color = (&HD8D8D8)

        Range("Medical_Expenditure_SC_80D").ClearContents
        Range("Medical_Expenditure_SC_80D").Locked = True
        Range("Medical_Expenditure_SC_80D").Interior.Color = (&HD8D8D8)

End If

If (Range("DropDown_ValueOf_FamilyM_80D").value) = "Not Claiming for Self/Family" Then

        Range("Senior_Citizen_80D").ClearContents
        Range("Senior_Citizen_80D").Locked = True
        Range("Senior_Citizen_80D").Interior.Color = (&HD8D8D8)
'Ayush_commented
'        Range("Health_InsuranceSC_80D").ClearContents
'        Range("Health_InsuranceSC_80D").Locked = True
'        Range("Health_InsuranceSC_80D").Interior.Color = (&HD8D8D8)

'Ayush_added
        Range("NameInsurerB1.80D").ClearContents
        Range("NameInsurerB1.80D").Locked = True
        Range("NameInsurerB1.80D").Interior.Color = (&HD8D8D8)
        
         Range("PolicyNumB1.80D").ClearContents
        Range("PolicyNumB1.80D").Locked = True
        Range("PolicyNumB1.80D").Interior.Color = (&HD8D8D8)
        
 'Commented by Ayush_06/05/2025 for 0.7 changes
'         Range("ReceiptNumB1.80D").ClearContents
'        Range("ReceiptNumB1.80D").Locked = True
'        Range("ReceiptNumB1.80D").Interior.Color = (&HD8D8D8)
'
         Range("AmtB1.80D").ClearContents
        Range("AmtB1.80D").Locked = True
        Range("AmtB1.80D").Interior.Color = (&HD8D8D8)
        
    '----------------------

        Range("Preventive_Health_SC_80D").ClearContents
        Range("Preventive_Health_SC_80D").Locked = True
        Range("Preventive_Health_SC_80D").Interior.Color = (&HD8D8D8)

        Range("Medical_Expenditure_SC_80D").ClearContents
        Range("Medical_Expenditure_SC_80D").Locked = True
        Range("Medical_Expenditure_SC_80D").Interior.Color = (&HD8D8D8)



        Range("Self_And_Family_80D").ClearContents
        Range("Self_And_Family_80D").Locked = True
        Range("Self_And_Family_80D").Interior.Color = (&HD8D8D8)
'Ayush_commented
'        Range("Health_Insurance_80D").ClearContents
'        Range("Health_Insurance_80D").Locked = True
'        Range("Health_Insurance_80D").Interior.Color = (&HD8D8D8)

'Ayush_Added
        Range("NameInsurerA1.80D").ClearContents
       Range("NameInsurerA1.80D").Locked = True
        Range("NameInsurerA1.80D").Interior.Color = (&HD8D8D8)
        
        Range("PolicyNumA1.80D").ClearContents
       Range("PolicyNumA1.80D").Locked = True
        Range("PolicyNumA1.80D").Interior.Color = (&HD8D8D8)
        
        
        'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumA1.80D").ClearContents
'       Range("ReceiptNumA1.80D").Locked = True
'        Range("ReceiptNumA1.80D").Interior.Color = (&HD8D8D8)
        
        Range("AmtA1.80D").ClearContents
       Range("AmtA1.80D").Locked = True
        Range("AmtA1.80D").Interior.Color = (&HD8D8D8)
        '-------------------

        Range("Preventive_Health_80D").ClearContents
        Range("Preventive_Health_80D").Locked = True
        Range("Preventive_Health_80D").Interior.Color = (&HD8D8D8)


 ElseIf (Range("DropDown_ValueOf_FamilyM_80D").value) = "(Select)" Then


        Range("Self_And_Family_80D").Locked = True
        Range("Self_And_Family_80D").Interior.Color = 13434879
'        Range("Self_And_Family_80D").Font.Color = RGB(0, 0, 0)
        Range("Self_And_Family_80D").Formula = "=MIN(25000,SUM(Health_Insurance2_80D,Preventive_Health2_80D))"
'Ayush_commented
'        Range("Health_Insurance_80D").Locked = False
'        Range("Health_Insurance_80D").Interior.Color = (&HCCFFCC)

'Ayush_added
        Range("NameInsurerA1.80D").Locked = False
        Range("NameInsurerA1.80D").Interior.Color = (&HCCFFCC)
        
         Range("PolicyNumA1.80D").Locked = False
        Range("PolicyNumA1.80D").Interior.Color = (&HCCFFCC)
        
'Commented by Ayush_06/05/2025 for 0.7 changes
'         Range("ReceiptNumA1.80D").Locked = False
'        Range("ReceiptNumA1.80D").Interior.Color = (&HCCFFCC)
        
         Range("AmtA1.80D").Locked = False
        Range("AmtA1.80D").Interior.Color = (&HCCFFCC)
        '----------------

        Range("Preventive_Health_80D").Locked = False
        Range("Preventive_Health_80D").Interior.Color = (&HCCFFCC)


        Range("Senior_Citizen_80D").Locked = True
        Range("Senior_Citizen_80D").Interior.Color = 13434879
'        Range("Senior_Citizen_80D").Font.Color = RGB(0, 0, 0)
        Range("Senior_Citizen_80D").Formula = "=MIN(50000,SUM(Health_Insurance3_80D,Preventive_Health3_80D,Medical_Expenditure2_80D))"

'        Range("Health_InsuranceSC_80D").Locked = False
'        Range("Health_InsuranceSC_80D").Interior.Color = (&HCCFFCC)

'Ayush_Added
        Range("NameInsurerB1.80D").Locked = False
        Range("NameInsurerB1.80D").Interior.Color = (&HCCFFCC)
        
        Range("PolicyNumB1.80D").Locked = False
        Range("PolicyNumB1.80D").Interior.Color = (&HCCFFCC)
        
'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumB1.80D").Locked = False
'        Range("ReceiptNumB1.80D").Interior.Color = (&HCCFFCC)
        
        Range("AmtB1.80D").Locked = False
        Range("AmtB1.80D").Interior.Color = (&HCCFFCC)
        '----------------

        Range("Preventive_Health_SC_80D").Locked = False
        Range("Preventive_Health_SC_80D").Interior.Color = (&HCCFFCC)

        Range("Medical_Expenditure_SC_80D").Locked = False
        Range("Medical_Expenditure_SC_80D").Interior.Color = (&HCCFFCC)

End If

Sheet16.Protect Password:=getmsgstate
End If


If (Target.name.name) = "DropDown_ValueOf_SC_80D" Then
   Sheet16.Unprotect Password:=getmsgstate

If (Mid(Range("DropDown_ValueOf_SC_80D").value, 1, 1) = "Y") Then

        Range("Parents_SC_80D").Locked = True
        Range("Parents_SC_80D").Interior.Color = 13434879
'        Range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Range("Parents_SC_80D").Formula = "=MIN(50000,SUM(Health_Insurance3_80D,Preventive_Health3_80D,Medical_Expenditure2_80D))"
'Ayush_commented
'        Range("Health_Insurance3_80D").Locked = False
'        Range("Health_Insurance3_80D").Interior.Color = (&HCCFFCC)

'Ayush_Added
        Range("NameInsurerB2.80D").Locked = False
        Range("NameInsurerB2.80D").Interior.Color = (&HCCFFCC)
        
        Range("PolicyNumB2.80D").Locked = False
        Range("PolicyNumB2.80D").Interior.Color = (&HCCFFCC)
        
        
         'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumB2.80D").Locked = False
'        Range("ReceiptNumB2.80D").Interior.Color = (&HCCFFCC)
        
        Range("AmtB2.80D").Locked = False
        Range("AmtB2.80D").Interior.Color = (&HCCFFCC)
'----------------

        Range("Preventive_Health3_80D").Locked = False
        Range("Preventive_Health3_80D").Interior.Color = (&HCCFFCC)

        Range("Medical_Expenditure2_80D").Locked = False
        Range("Medical_Expenditure2_80D").Interior.Color = (&HCCFFCC)


        Range("Parents_80D").ClearContents
        Range("Parents_80D").Locked = True
        Range("Parents_80D").Interior.Color = (&HD8D8D8)
'Ayush_commented
'        Range("Health_Insurance2_80D").ClearContents
'        Range("Health_Insurance2_80D").Locked = True
'        Range("Health_Insurance2_80D").Interior.Color = (&HD8D8D8)

'Ayush_Added
        Range("NameInsurerA2.80D").ClearContents
       Range("NameInsurerA2.80D").Locked = True
        Range("NameInsurerA2.80D").Interior.Color = (&HD8D8D8)
        
        Range("PolicyNumA2.80D").ClearContents
       Range("PolicyNumA2.80D").Locked = True
        Range("PolicyNumA2.80D").Interior.Color = (&HD8D8D8)
        
   'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumA2.80D").ClearContents
'       Range("ReceiptNumA2.80D").Locked = True
'        Range("ReceiptNumA2.80D").Interior.Color = (&HD8D8D8)
        
        Range("AmtA2.80D").ClearContents
       Range("AmtA2.80D").Locked = True
        Range("AmtA2.80D").Interior.Color = (&HD8D8D8)
'--------------------

        Range("Preventive_Health2_80D").ClearContents
        Range("Preventive_Health2_80D").Locked = True
        Range("Preventive_Health2_80D").Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("DropDown_ValueOf_SC_80D").value, 1, 1) = "N") Then

        Range("Parents_80D").Locked = True
        Range("Parents_80D").Interior.Color = 13434879
'        Range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Range("Parents_80D").Formula = "=MIN(25000,SUM(Health_Insurance2_80D,Preventive_Health2_80D))"
'Ayush_commented
'        Range("Health_Insurance2_80D").Locked = False
'        Range("Health_Insurance2_80D").Interior.Color = (&HCCFFCC)
'Ayush_Added
        Range("NameInsurerA2.80D").Locked = False
        Range("NameInsurerA2.80D").Interior.Color = (&HCCFFCC)
        
        Range("PolicyNumA2.80D").Locked = False
        Range("PolicyNumA2.80D").Interior.Color = (&HCCFFCC)
        
        
         'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumA2.80D").Locked = False
'        Range("ReceiptNumA2.80D").Interior.Color = (&HCCFFCC)
        
        Range("AmtA2.80D").Locked = False
        Range("AmtA2.80D").Interior.Color = (&HCCFFCC)
'----------------



        Range("Preventive_Health2_80D").Locked = False
        Range("Preventive_Health2_80D").Interior.Color = (&HCCFFCC)


        Range("Parents_SC_80D").ClearContents
        Range("Parents_SC_80D").Locked = True
        Range("Parents_SC_80D").Interior.Color = (&HD8D8D8)
'Ayush_commented
'        Range("Health_Insurance3_80D").ClearContents
'        Range("Health_Insurance3_80D").Locked = True
'        Range("Health_Insurance3_80D").Interior.Color = (&HD8D8D8)

'Ayush_Added
        Range("NameInsurerB2.80D").ClearContents
       Range("NameInsurerB2.80D").Locked = True
        Range("NameInsurerB2.80D").Interior.Color = (&HD8D8D8)
        
        Range("PolicyNumB2.80D").ClearContents
       Range("PolicyNumB2.80D").Locked = True
        Range("PolicyNumB2.80D").Interior.Color = (&HD8D8D8)
        
        
         'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumB2.80D").ClearContents
'       Range("ReceiptNumB2.80D").Locked = True
'        Range("ReceiptNumB2.80D").Interior.Color = (&HD8D8D8)
        
        Range("AmtB2.80D").ClearContents
       Range("AmtB2.80D").Locked = True
        Range("AmtB2.80D").Interior.Color = (&HD8D8D8)
'--------------------

        Range("Preventive_Health3_80D").ClearContents
        Range("Preventive_Health3_80D").Locked = True
        Range("Preventive_Health3_80D").Interior.Color = (&HD8D8D8)

        Range("Medical_Expenditure2_80D").ClearContents
        Range("Medical_Expenditure2_80D").Locked = True
        Range("Medical_Expenditure2_80D").Interior.Color = (&HD8D8D8)





End If

If (Range("DropDown_ValueOf_SC_80D").value) = "Not claiming for Parents" Then

        Range("Parents_80D").ClearContents
        Range("Parents_80D").Locked = True
        Range("Parents_80D").Interior.Color = (&HD8D8D8)
'AYUSH_COMMENTED
'        Range("Health_Insurance2_80D").ClearContents
'        Range("Health_Insurance2_80D").Locked = True
'        Range("Health_Insurance2_80D").Interior.Color = (&HD8D8D8)


'Ayush_Added
        Range("NameInsurerA2.80D").ClearContents
       Range("NameInsurerA2.80D").Locked = True
        Range("NameInsurerA2.80D").Interior.Color = (&HD8D8D8)
        
        Range("PolicyNumA2.80D").ClearContents
       Range("PolicyNumA2.80D").Locked = True
        Range("PolicyNumA2.80D").Interior.Color = (&HD8D8D8)
        
 'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumA2.80D").ClearContents
'       Range("ReceiptNumA2.80D").Locked = True
'        Range("ReceiptNumA2.80D").Interior.Color = (&HD8D8D8)
        
        Range("AmtA2.80D").ClearContents
       Range("AmtA2.80D").Locked = True
        Range("AmtA2.80D").Interior.Color = (&HD8D8D8)
'--------------------

        Range("Preventive_Health2_80D").ClearContents
        Range("Preventive_Health2_80D").Locked = True
        Range("Preventive_Health2_80D").Interior.Color = (&HD8D8D8)


        Range("Parents_SC_80D").ClearContents
        Range("Parents_SC_80D").Locked = True
        Range("Parents_SC_80D").Interior.Color = (&HD8D8D8)
 'AYUSH_COMMENTED

'        Range("Health_Insurance3_80D").ClearContents
'        Range("Health_Insurance3_80D").Locked = True
'        Range("Health_Insurance3_80D").Interior.Color = (&HD8D8D8)
        
        'Ayush_Added
        Range("NameInsurerB2.80D").ClearContents
       Range("NameInsurerB2.80D").Locked = True
        Range("NameInsurerB2.80D").Interior.Color = (&HD8D8D8)
        
        Range("PolicyNumB2.80D").ClearContents
       Range("PolicyNumB2.80D").Locked = True
        Range("PolicyNumB2.80D").Interior.Color = (&HD8D8D8)
        
        
         'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumB2.80D").ClearContents
'       Range("ReceiptNumB2.80D").Locked = True
'        Range("ReceiptNumB2.80D").Interior.Color = (&HD8D8D8)
'
        Range("AmtB2.80D").ClearContents
       Range("AmtB2.80D").Locked = True
        Range("AmtB2.80D").Interior.Color = (&HD8D8D8)
'--------------------

        Range("Preventive_Health3_80D").ClearContents
        Range("Preventive_Health3_80D").Locked = True
        Range("Preventive_Health3_80D").Interior.Color = (&HD8D8D8)

        Range("Medical_Expenditure2_80D").ClearContents
        Range("Medical_Expenditure2_80D").Locked = True
        Range("Medical_Expenditure2_80D").Interior.Color = (&HD8D8D8)

ElseIf (Range("DropDown_ValueOf_SC_80D").value) = "(Select)" Then

        Range("Parents_80D").Locked = True
        Range("Parents_80D").Interior.Color = 13434879
'        Range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Range("Parents_80D").Formula = "=MIN(25000,SUM(Health_Insurance2_80D,Preventive_Health2_80D))"
'AYUSH_COMMENTED
'        Range("Health_Insurance2_80D").Locked = False
'        Range("Health_Insurance2_80D").Interior.Color = (&HCCFFCC)


'Ayush_Added
        Range("NameInsurerA2.80D").Locked = False
        Range("NameInsurerA2.80D").Interior.Color = (&HCCFFCC)
        
        Range("PolicyNumA2.80D").Locked = False
        Range("PolicyNumA2.80D").Interior.Color = (&HCCFFCC)
        
        
         'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumA2.80D").Locked = False
'        Range("ReceiptNumA2.80D").Interior.Color = (&HCCFFCC)
        
        Range("AmtA2.80D").Locked = False
        Range("AmtA2.80D").Interior.Color = (&HCCFFCC)
'----------------

        Range("Preventive_Health2_80D").Locked = False
        Range("Preventive_Health2_80D").Interior.Color = (&HCCFFCC)


        Range("Parents_SC_80D").Locked = False
        Range("Parents_SC_80D").Interior.Color = 13434879
'        Range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Range("Parents_SC_80D").Formula = "=MIN(50000,SUM(Health_Insurance3_80D,Preventive_Health3_80D,Medical_Expenditure2_80D))"
        
        
'AYUSH_COMMENTED
'        Range("Health_Insurance3_80D").Locked = False
'        Range("Health_Insurance3_80D").Interior.Color = (&HCCFFCC)

'Ayush_Added
        Range("NameInsurerB2.80D").Locked = False
        Range("NameInsurerB2.80D").Interior.Color = (&HCCFFCC)
        
        Range("PolicyNumB2.80D").Locked = False
        Range("PolicyNumB2.80D").Interior.Color = (&HCCFFCC)
        
    'Commented by Ayush_06/05/2025 for 0.7 changes
'        Range("ReceiptNumB2.80D").Locked = False
'        Range("ReceiptNumB2.80D").Interior.Color = (&HCCFFCC)
        
        Range("AmtB2.80D").Locked = False
        Range("AmtB2.80D").Interior.Color = (&HCCFFCC)
'----------------

        Range("Preventive_Health3_80D").Locked = False
        Range("Preventive_Health3_80D").Interior.Color = (&HCCFFCC)

        Range("Medical_Expenditure2_80D").Locked = False
        Range("Medical_Expenditure2_80D").Interior.Color = (&HCCFFCC)

End If

Sheet16.Protect Password:=getmsgstate
End If
endline:

Application.EnableEvents = True

End Sub


