Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


Public currCnt22, currCnt33, currCnt44, currCnt66, currCnt67, currCnt68, currCnt69, currCnt70 As Variant
Public cutoff1 As Date
Sub a()
Application.EnableEvents = True
Dim sheet
For Each sheet In ThisWorkbook.Sheets
sheet.Protect Password:=getmsgstate
Next
End Sub
Sub LockUnlockClauseiv()
    
        If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
        
        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("clauseiv7provisio139iFlg").value = "(Select)"
        Application.EnableEvents = False
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("clauseiv7provisio139iFlg").MergeArea.Locked = True

        Sheet1.Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
        Application.EnableEvents = True
        Sheet1.Protect Password:=getmsgstate
        
        End If

End Sub
Sub LockUnlockQVIA()
'On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
Sheet3.Unprotect Password:=getmsgstate


If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then

        Sheet3.Range("Sec44AD.NOBName1").ClearContents
        Sheet3.Range("Sec44AD.NOBName1").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("Sec44AD.NOBName1").Locked = True
 
        Sheet3.Range("Sec44AD.NOBCode1").ClearContents
        Sheet3.Range("Sec44AD.NOBCode1").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("Sec44AD.NOBCode1").Locked = True
        
        Sheet3.Range("Sec44AD.NOBDescription1").ClearContents
        Sheet3.Range("Sec44AD.NOBDescription1").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("Sec44AD.NOBDescription1").Locked = True
        
'        Sheet3.Range("BP_E3_GR").ClearContents
'        Sheet3.Range("BP_E3_GR").Interior.ColorIndex = 15   ', "AU"
'        Sheet3.Range("BP_E3_GR").MergeArea.Locked = True

'AY 2024-25 change for Table E3 Starts-----------------------------

        Sheet3.Range("BP_E3_a").ClearContents
        Sheet3.Range("BP_E3_a").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("BP_E3_a").MergeArea.Locked = True

        Sheet3.Range("BP_E3_b").ClearContents
        Sheet3.Range("BP_E3_b").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("BP_E3_b").MergeArea.Locked = True

        Sheet3.Range("BP_E3_c").ClearContents
        Sheet3.Range("BP_E3_c").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("BP_E3_c").MergeArea.Locked = True
        
'AY 2024-25 change for Table E3 Ends-----------------------------


        Sheet3.Range("BP_E4_PI44ADA").ClearContents
        Sheet3.Range("BP_E4_PI44ADA").Interior.ColorIndex = 15   ', "AU"
        Sheet3.Range("BP_E4_PI44ADA").MergeArea.Locked = True
  
        Sheet1.Range("IncD.Section89").MergeArea.ClearContents
        Sheet1.Range("IncD.Section89").Interior.ColorIndex = 15   ', "AU"
        Sheet1.Range("IncD.Section89").MergeArea.Locked = True
        
Else
    
    Sheet3.Range("Sec44AD.NOBName1").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("Sec44AD.NOBName1").Locked = False
    
    Sheet3.Range("Sec44AD.NOBCode1").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("Sec44AD.NOBCode1").Locked = False
    
    Sheet3.Range("Sec44AD.NOBDescription1").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("Sec44AD.NOBDescription1").Locked = False
    
'    Sheet3.Range("BP_E3_GR").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet3.Range("BP_E3_GR").Locked = False

'AY 2024-25 change for Table E3 Starts-----------------------------

    Sheet3.Range("BP_E3_a").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("BP_E3_a").Locked = False
 
    Sheet3.Range("BP_E3_b").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("BP_E3_b").Locked = False

    Sheet3.Range("BP_E3_c").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("BP_E3_c").Locked = False
    
'AY 2024-25 change for Table E3 End-----------------------------



'
    Sheet3.Range("BP_E4_PI44ADA").Interior.Color = (&HCCFFCC)   ', "AU"
    Sheet3.Range("BP_E4_PI44ADA").Locked = False
    
    Sheet1.Range("IncD.Section89").MergeArea.Locked = False
    Sheet1.Range("IncD.Section89").MergeArea.Interior.Color = (&HCCFFCC)   ', "AU"
    
    
End If

Sheet3.Protect Password:=getmsgstate

    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or _
       Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then

'        Sheet1.Range("IncD.Section80DValue").MergeArea.ClearContents
'        Sheet1.Range("IncD.Section80DValue").MergeArea.Locked = True
'        Sheet1.Range("IncD.Section80DValue").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("IncD.Section80DValue").value = ""
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
        Sheet1.Range("IncD.Section80E").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80E").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80E").value = ""

        Sheet1.Range("IncD.Section80EE").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80EE").value = ""

        Sheet1.Range("IncD.Section80C").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80C").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80C").value = ""

        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80EEA").value = ""

        Sheet1.Range("IncD.Section80EEB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80EEB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80EEB").value = ""
    End If

        'Sheet1.Range("IncD.Section80U").MergeArea.ClearContents
        'Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80U").MergeArea.Interior.Color = "&HD8D8D8"
        'Sheet1.Range("IncD.Section80U").value = ""

'        Sheet1.Range("SELECT80U").MergeArea.ClearContents
'        Sheet1.Range("SELECT80U").MergeArea.Locked = True
'        Sheet1.Range("SELECT80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("SELECT80U").value = ""

'        Sheet1.Range("IncD.Section80DC").MergeArea.ClearContents
'        Sheet1.Range("IncD.Section80DC").MergeArea.Locked = True
'        Sheet1.Range("IncD.Section80DC").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("IncD.Section80DC").value = ""
'
'        Sheet1.Range("SELECT80DC").MergeArea.ClearContents
'        Sheet1.Range("SELECT80DC").MergeArea.Locked = True
'        Sheet1.Range("SELECT80DC").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("SELECT80DC").value = ""
    ElseIf Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" And Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
        'Sheet1.Range("IncD.Section80U").MergeArea.ClearContents
        'Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80U").MergeArea.Interior.Color = "&HD8D8D8"
        'Sheet1.Range("IncD.Section80U").value = ""

'        Sheet1.Range("SELECT80U").MergeArea.ClearContents
'        Sheet1.Range("SELECT80U").MergeArea.Locked = True
'        Sheet1.Range("SELECT80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("SELECT80U").value = ""

    ElseIf Sheet8.Range("bacValue").value = 2 Then
'        Sheet1.Range("IncD.Section80DValue").MergeArea.ClearContents
'        Sheet1.Range("IncD.Section80DValue").MergeArea.Locked = False
'        Sheet1.Range("IncD.Section80DValue").MergeArea.Interior.ColorIndex = 2
'        Sheet1.Range("IncD.Section80DValue").Formula = "=Eligible_Amount_80D"


        Sheet1.Range("IncD.Section80E").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80E").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80E").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Section80E").value = ""

        Sheet1.Range("IncD.Section80EE").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EE").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Section80EE").value = ""

        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Section80EEA").value = ""


        Sheet1.Range("IncD.Section80C").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80C").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80C").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Section80C").value = ""


        Sheet1.Range("IncD.Section80EEB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80EEB").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Section80EEB").value = ""

'        Sheet1.Range("SELECT80U").MergeArea.ClearContents
'        Sheet1.Range("SELECT80U").MergeArea.Locked = False
'        Sheet1.Range("SELECT80U").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("SELECT80U").value = "(Select)"

        'Sheet1.Range("IncD.Section80U").MergeArea.ClearContents
       ' Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80U").MergeArea.Interior.ColorIndex = 2
        'Sheet1.Range("IncD.Section80U").value = ""

'        Sheet1.Range("SELECT80DC").MergeArea.ClearContents
'        Sheet1.Range("SELECT80DC").MergeArea.Locked = False
'        Sheet1.Range("SELECT80DC").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("SELECT80DC").value = "(Select)"
'
'        Sheet1.Range("IncD.Section80DC").MergeArea.ClearContents
'        Sheet1.Range("IncD.Section80DC").MergeArea.Locked = False
'        Sheet1.Range("IncD.Section80DC").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("IncD.Section80DC").value = ""


    End If
'      If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
'
'        Sheet1.Range("DropDown_ValueOf_FamilyM_80D").MergeArea.ClearContents
'        Sheet1.Range("DropDown_ValueOf_FamilyM_80D").MergeArea.Locked = True
'        Sheet1.Range("DropDown_ValueOf_FamilyM_80D").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("DropDown_ValueOf_FamilyM_80D").value = "(Select)"
'
'        Sheet1.Range("Self_And_Family_80D").MergeArea.ClearContents
'        Sheet1.Range("Self_And_Family_80D").MergeArea.Locked = True
'        Sheet1.Range("Self_And_Family_80D").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("Self_And_Family_80D").value = "(Select)"
'
'        Sheet1.Range("Health_Insurance_80D").MergeArea.ClearContents
'        Sheet1.Range("Health_Insurance_80D").MergeArea.Locked = True
'        Sheet1.Range("Health_Insurance_80D").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("Health_Insurance_80D").value = "(Select)"
'
'        Sheet1.Range("Preventive_Health_80D").MergeArea.ClearContents
'        Sheet1.Range("Preventive_Health_80D").MergeArea.Locked = True
'        Sheet1.Range("Health_Insurance_80D").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("Health_Insurance_80D").value = "(Select)"
'
'
'
'        Else
'
'    End If



Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub




Sub ProtectAll()
ThisWorkbook.Protect Password:=getmsgstate
Sheet1.Protect Password:=getmsgstate
Sheet2.Protect Password:=getmsgstate
Sheet3.Protect Password:=getmsgstate
Sheet4.Protect Password:=getmsgstate
Sheet5.Protect Password:=getmsgstate
Sheet6.Protect Password:=getmsgstate
Sheet7.Protect Password:=getmsgstate
Sheet8.Protect Password:=getmsgstate
Sheet9.Protect Password:=getmsgstate
Sheet10.Protect Password:=getmsgstate
Sheet12.Protect Password:=getmsgstate
Sheet13.Protect Password:=getmsgstate
Sheet16.Protect Password:=getmsgstate
Sheet17.Protect Password:=getmsgstate
Sheet18.Protect Password:=getmsgstate
End Sub

Sub WorkbookProtectAll()
ThisWorkbook.Protect Password:=getmsgstate
Sheet9.Visible = False
End Sub

Sub LockUnlock80DDB()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate

If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then

'        Sheet1.Range("SELECT80DD").MergeArea.ClearContents
'        Sheet1.Range("SELECT80DD").MergeArea.Locked = True
'        Sheet1.Range("SELECT80DD").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("SELECT80DD").value = ""

        Sheet1.Range("SELECT80DDB").MergeArea.ClearContents
        Sheet1.Range("SELECT80DDB").MergeArea.Locked = True
        Sheet1.Range("SELECT80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("SELECT80DDB").value = ""
        
'For SIT-92091 Newly added by sai on 25/04/2025
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.ClearContents
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("Sheet1.Specified_Disease").value = ""
        


        'Sheet1.Range("IncD.Section80DD").MergeArea.ClearContents
        'Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80DD").MergeArea.Interior.Color = "&HD8D8D8"
        'Sheet1.Range("IncD.Section80DD").value = ""

        Sheet1.Range("IncD.Section80DDB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80DDB").value = ""



ElseIf Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" And Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then

'        Sheet1.Range("SELECT80DD").MergeArea.ClearContents
'        Sheet1.Range("SELECT80DD").MergeArea.Locked = True
'        Sheet1.Range("SELECT80DD").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("SELECT80DD").value = ""

        Sheet1.Range("SELECT80DDB").MergeArea.ClearContents
        Sheet1.Range("SELECT80DDB").MergeArea.Locked = True
        Sheet1.Range("SELECT80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("SELECT80DDB").value = ""


        'Sheet1.Range("IncD.Section80DD").MergeArea.ClearContents
        'Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80DD").MergeArea.Interior.Color = "&HD8D8D8"
        'Sheet1.Range("IncD.Section80DD").value = ""

        Sheet1.Range("IncD.Section80DDB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80DDB").value = ""

 ElseIf Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" And Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then

'        Sheet1.Range("SELECT80DD").MergeArea.ClearContents
'        Sheet1.Range("SELECT80DD").MergeArea.Locked = True
'        Sheet1.Range("SELECT80DD").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("SELECT80DD").value = ""

        Sheet1.Range("SELECT80DDB").MergeArea.ClearContents
        Sheet1.Range("SELECT80DDB").MergeArea.Locked = True
        Sheet1.Range("SELECT80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("SELECT80DDB").value = ""


        'Sheet1.Range("IncD.Section80DD").MergeArea.ClearContents
       ' Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80DD").MergeArea.Interior.Color = "&HD8D8D8"
        'Sheet1.Range("IncD.Section80DD").value = ""

        Sheet1.Range("IncD.Section80DDB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Section80DDB").value = ""





ElseIf (Sheet1.Range("sheet1.ResidentialStatus1").value) = "(Select)" And (Sheet1.Range("sheet1.Status").value) = "(Select)" Then


'        Sheet1.Range("SELECT80DD").MergeArea.Locked = False
'        Sheet1.Range("SELECT80DD").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("SELECT80DD").value = "(Select)"


        Sheet1.Range("SELECT80DDB").MergeArea.Locked = False
        Sheet1.Range("SELECT80DDB").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("SELECT80DDB").value = "(Select)"
        
        
'For SIT-92091 Newly added by sai on 25/04/2025
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("Sheet1.Specified_Disease").value = "(Select)"


        'Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
        'Sheet1.Range("IncD.Section80DD").MergeArea.Interior.ColorIndex = 2
        'Sheet1.Range("IncD.Section80DD").value = ""


        Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80DDB").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Section80DDB").value = ""

ElseIf Sheet8.Range("bacValue").value = 2 Then

'        Sheet1.Range("SELECT80DD").MergeArea.Locked = False
'        Sheet1.Range("SELECT80DD").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("SELECT80DD").value = "(Select)"


        Sheet1.Range("SELECT80DDB").MergeArea.Locked = False
        Sheet1.Range("SELECT80DDB").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("SELECT80DDB").value = "(Select)"
        
'For SIT-92091 Newly added by sai on 25/04/2025
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("Sheet1.Specified_Disease").value = "(Select)"


       ' Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
       ' Sheet1.Range("IncD.Section80DD").MergeArea.Interior.ColorIndex = 2
       
        Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80DDB").MergeArea.Interior.Color = "&HCCFFCC"


End If


Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub

Sub LockUnlock80TTB()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate

DOB_6 = Sheet1.Range("sheet1.DOB").value
If calculateAge(DOB_6) <= 59 Then

            Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
            Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
            Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HCCFFCC"

ElseIf calculateAge(DOB_6) > 59 And calculateAge(DOB_6) <> 2025 Then 'Changes from 2024 to 2025 by Bindu

            Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
            Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"

End If

If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Or (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") Then

        Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"

ElseIf (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") Then
        Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"

ElseIf Sheet8.Range("bacValue").value = 2 Then
        Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HCCFFCC"


End If


If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F") Then
        Sheet1.Range("IncD.Section80TTA").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HD8D8D8"

ElseIf Sheet8.Range("bacValue").value = 2 Then
        Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HCCFFCC"


End If


    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then


                 Sheet1.Range("IncD.Section80TTA").value = ""
                 Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
                 Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)

'                 Else
'                 Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                End If



    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then

                Sheet1.Range("IncD.Section80TTB").value = ""
                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                End If

If calculateAge(Trim(Range("sheet1.DOB").value)) <= 59 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then

                Sheet1.Range("IncD.Section80TTB").value = ""
                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)


                End If
      'Chandru
'        If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date") And Sheet1.Range("sheet1.DOB") <> "" Then
'         If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime") And Sheet1.Range("sheet1.DOB") <> "" Then
'Ayush_27/01
            If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I") And (Sheet8.Range("bacValue").value = 1) And Sheet1.Range("sheet1.DOB") <> "" Then
                Sheet1.Range("IncD.Section80TTB").value = ""
                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
                
                Sheet1.Range("IncD.Section80TTA").value = ""
                Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
                Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)
                
                
                End If


Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
Sub PANStatusChange()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
    If Mid(Range("sheet1.PAN").value, 4, 1) = "P" Then
        Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL"
    ElseIf Mid(Range("sheet1.PAN").value, 4, 1) = "H" Then
        Sheet1.Range("sheet1.Status").value = "H - HUF"
    ElseIf Mid(Range("sheet1.PAN").value, 4, 1) = "F" Then
        Sheet1.Range("sheet1.Status").value = "F - FIRM(Other than LLP)"
    Else
        Sheet1.Range("sheet1.Status").value = "(Select)"
    End If
Application.EnableEvents = True
Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockNatureofEmployment()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate

If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then

        Sheet1.Range("sheet1.NatureofEmployment").MergeArea.ClearContents
        Sheet1.Range("sheet1.NatureofEmployment").MergeArea.Locked = True
        Sheet1.Range("sheet1.NatureofEmployment").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.NatureofEmployment").value = "Not Applicable"
Else


        Sheet1.Range("sheet1.NatureofEmployment").MergeArea.ClearContents
        Sheet1.Range("sheet1.NatureofEmployment").MergeArea.Locked = False
        Sheet1.Range("sheet1.NatureofEmployment").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.NatureofEmployment").value = "(Select)"

End If
Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub

Sub LockUnlockAadhaar()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or _
       Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
        Sheet1.Range("sheet1.Aadhaar").MergeArea.ClearContents
        Sheet1.Range("sheet1.Aadhaar").MergeArea.Locked = True
        Sheet1.Range("sheet1.Aadhaar").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.Aadhaar").value = ""
'Ayush_enrol=======================
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = True
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("sheet1.AadhaarEnrol").value = ""
'========================================
    Else
        Sheet1.Range("sheet1.Aadhaar").MergeArea.ClearContents
        Sheet1.Range("sheet1.Aadhaar").MergeArea.Locked = False
        Sheet1.Range("sheet1.Aadhaar").MergeArea.Interior.Color = "&HCCFFCC"
        'Sheet1.Range("sheet1.Aadhaar").value = ""
'Ayush_enrol=================
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = "&HCCFFCC"

'================================
        'Sheet1.Range("sheet1.AadhaarEnrol").value = ""
    End If
Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub

Sub LockUnlockReturnfilesec()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12) Then
'        Range("sheet1.ReceiptNo1").MergeArea.Locked = False
'        Range("sheet1.ReceiptNo1").MergeArea.Interior.Color = (&HCCFFCC)
'
'
'        Range("sheet1.OrigRetFiledDate1").MergeArea.Locked = False
'        Range("sheet1.OrigRetFiledDate1").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.ReceiptNo").MergeArea.ClearContents
        Range("sheet1.ReceiptNo").MergeArea.Locked = True
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeNo").MergeArea.ClearContents
        Range("sheet1.NoticeNo").MergeArea.Locked = True
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeDate").MergeArea.ClearContents
        Range("sheet1.NoticeDate").MergeArea.Locked = True
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)


    ElseIf (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 13 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 14 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 15 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 16 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 20) Then

        Range("sheet1.ReceiptNo").MergeArea.ClearContents
        Range("sheet1.ReceiptNo").MergeArea.Locked = True
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)


'        Range("sheet1.NoticeNo").Value = ""
        Range("sheet1.NoticeNo").MergeArea.Locked = False
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HCCFFCC)

'        Range("sheet1.NoticeDate").Value = ""
        Range("sheet1.NoticeDate").MergeArea.Locked = False
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)

    ElseIf Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17 Or Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 21 Then

        Range("sheet1.ReceiptNo").MergeArea.Locked = False
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)


        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeNo").MergeArea.ClearContents
        Range("sheet1.NoticeNo").MergeArea.Locked = True
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeDate").MergeArea.ClearContents
        Range("sheet1.NoticeDate").MergeArea.Locked = True
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)


  ElseIf (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18) Then


        Range("sheet1.ReceiptNo").MergeArea.Locked = False
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeNo").MergeArea.Locked = False
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeDate").MergeArea.Locked = False
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)
    Else


'        Range("sheet1.ReceiptNo").MergeArea.ClearContents
        Range("sheet1.ReceiptNo").MergeArea.Locked = False
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)


'        Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

'        Range("sheet1.NoticeNo").MergeArea.ClearContents
        Range("sheet1.NoticeNo").MergeArea.Locked = False
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HCCFFCC)

'        Range("sheet1.NoticeDate").MergeArea.ClearContents
        Range("sheet1.NoticeDate").MergeArea.Locked = False
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)


End If


'            If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18) Then
'                Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = 13434879
'
'                Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = 13434879
''                Range("IncD.IntrstPayUs234B").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234C").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234C").MergeArea.Interior.Color = 13434879
''                Range("IncD.IntrstPayUs234C").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = 13434879
'
'
'             ElseIf ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12)) Then
'             'Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18)) Then
'
'                Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = 13434879
'
'                Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = 13434879
''                Range("IncD.IntrstPayUs234B").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234C").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234C").MergeArea.Interior.Color = 13434879
''                Range("IncD.IntrstPayUs234C").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = 13434879
'
'            ElseIf (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17) Then
'               Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = 13434879
'
'                Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = 13434879
''                Range("IncD.IntrstPayUs234B").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234C").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234C").MergeArea.Interior.Color = 13434879
''                Range("IncD.IntrstPayUs234C").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = 13434879
'
'
'            Else
'
'                Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = 13434879
'
'                Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = 13434879
'
'                Range("IncD.IntrstPayUs234C").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234C").MergeArea.Interior.Color = 13434879
'                Range("IncD.IntrstPayUs234C").MergeArea.value = 0
'
'                Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
'                Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = 13434879
'
'            End If






Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
Sub LockUnlockDed()
On Error Resume Next
Sheet1.Activate
Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False
If Range("sheet1.status").value = "I - INDIVIDUAL" And Sheet8.Range("bacValue").value = 2 Then


'If Range("Sheet1.NatureofEmployment").value = "State Government" Or Range("Sheet1.NatureofEmployment").value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958" Or Range("Sheet1.NatureofEmployment").value = "Central Government" Or Range("Sheet1.NatureofEmployment").value = "Public Sector Undertaking" Or Range("Sheet1.NatureofEmployment").value = "Others" Then
        
 'By Bindu as per DE V0.4 on 24th feb 26
 If Range("Sheet1.NatureofEmployment").value = "State Government" Or Range("Sheet1.NatureofEmployment").value = "Central Government" Or Range("Sheet1.NatureofEmployment").value = "Public Sector Undertaking" Or Range("Sheet1.NatureofEmployment").value = "Others" Then

        If Range("Sheet1.NatureofEmployment").value <> "Others" Then
            Range("IncD.IncomeFromSal_v").MergeArea.Locked = False

           Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)
            
        Else
         Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HD8D8D8)

        Range("IncD.IncomeFromSal_v").MergeArea.Locked = True

        Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
        
        End If
        Range("IncD.Section80CCD").MergeArea.Locked = False

        Range("IncD.Section80CCD").MergeArea.Interior.Color = (&HCCFFCC)



Else
    
     Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HD8D8D8)

        Range("IncD.IncomeFromSal_v").MergeArea.Locked = True

        Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
        
        Range("IncD.Section80CCD").MergeArea.Interior.Color = (&HD8D8D8)

        Range("IncD.Section80CCD").MergeArea.Locked = True

        Range("IncD.Section80CCD").MergeArea.ClearContents

        

End If


End If

End Sub
'Function LockUnlock115BAC()
'Application.EnableEvents = False
'Sheet1.Unprotect Password:=getmsgstate
'
'If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
'
'
'        Sheet8.Unprotect Password:=getmsgstate
'        Sheet8.Range("bacValue").Clear
'        Sheet8.Range("oldbacValue").Clear
'        Sheet8.Protect Password:=getmsgstate
'        Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
'        Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
'        Sheet1.Shapes("BacYes").OLEFormat.Object.Enabled = False
'        Sheet1.Shapes("BacNO").OLEFormat.Object.Enabled = False
'        Sheet1.Shapes("BacNO").Visible = False
'        Sheet1.Shapes("BacYes").Visible = False
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = True
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Interior.Color = (&HD8D8D8)
'        Sheet1.Range("AO31").Interior.Color = (&HD8D8D8)
'        ThisWorkbook.Unprotect Password:=getmsgstate
'        Sheet6.Visible = True
'        ThisWorkbook.Protect Password:=getmsgstate
'        reset10IE
'        'Change-16.01.2023.103
'        reset10IEAy23_24
'        'End Change
'
'
'Else
'        Sheet1.Shapes("BacNO").Visible = True
'        Sheet1.Shapes("BacYes").Visible = True
'        Sheet1.Shapes("BacYes").OLEFormat.Object.Enabled = True
'        Sheet1.Shapes("BacNO").OLEFormat.Object.Enabled = True
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Interior.Color = (&HCCFFCC)
'        Sheet1.Range("AO31").Interior.Color = vbWhite
'
'
'End If
'
'Application.EnableEvents = True
'Sheet1.Protect Password:=getmsgstate
'
'
'
'End Function

'Chandru
Private Sub CommandButton1_Click()
ActiveWorkbook.FollowHyperlink Address:="https://eportal.incometax.gov.in/iec/foservices/#/TaxCalc/calculator"

End Sub

Private Sub Worksheet_Activate()
    Sheet1.Unprotect Password:=getmsgstate
    
    'Ayush_23/02/2026
 ' Call FitComments
  'Ayush_03/04/2026
  Call FitComments_DynamicHeight
    'Commented By Ayush
  '  ActiveSheet.Range("Sheet1.Duedate").value = Sheet9.Range("DueDate").value
    
'Ayush_22/07

'If Dformat(Sheet5.Range("Ver.Date"), "") > due1 Then
'
'Sheet1.Range("sheet1.ReturnFileSec").value = "139(4)-After due date"
'
'Else
'
'Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'
'End If

 
''Ayush_29/07
'Dim due1, due2 As String
'
'due1 = "2025-09-15"
'
'due2 = "2025-09-16"
'
'Dim sheet1_ReturnFileSec_duplicate As Variant
'
'sheet1_ReturnFileSec_duplicate = Sheet1.Range("sheet1.ReturnFileSec").value
'
''Updated by Sadineni_29/07/2025
'
'If Dformat(Sheet5.Range("Ver.Date"), "") <= due1 Then
'
'    'Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'
'    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(4)-Belated" Or Sheet1.Range("sheet1.ReturnFileSec") = "(Select)" Then
'
'           Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'
''    Else
'
''        Sheet1.Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
'
'    End If
'
'     If Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
'
'       If Range("sheet1.ReceiptNo").MergeArea.Locked = False Or Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False Then
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Application.EnableEvents = False
'
'                Range("sheet1.ReceiptNo").value = ""
'
'                Range("sheet1.ReceiptNo").MergeArea.Locked = True
'
'                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)
'
'                Range("sheet1.OrigRetFiledDate").value = ""
'
'                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
'
'                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)
'
'        Application.EnableEvents = True
'
'        End If
'
'     End If
'
'End If
'
'
'If Dformat(Sheet5.Range("Ver.Date"), "") >= due2 And Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
'
'    Sheet1.Range("sheet1.ReturnFileSec").value = "139(4)-Belated"
'
'        ElseIf Dformat(Sheet5.Range("Ver.Date"), "") >= due2 Then
'
'        Sheet1.Range("sheet1.ReturnFileSec").value = sheet1_ReturnFileSec_duplicate
'
'End If
'

'*****************************************************

   If Sheet1.Range("sheet1.NatureofEmployment").value = Pensioners - PSU Then

Dim Sheet1_Nature_cellscnt, i As Variant

Sheet1_Nature_cellscnt = Sheet1.Range("Sheet1.Nature").Rows.count

        For i = 1 To Sheet5_Nature_cellscnt

                If Sheet5.Range("Sheet1.Nature").item(i, 1).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette " Then

                     Sheet5.Range("Sheet1.Nature").item(i, 1).value = "(Select)"

                     Sheet5.Range("Sheet1.Nature").item(i, 1).Offset(0, 1).value = ""

                     Sheet5.Range("Sheet1.Nature").item(i, 1).Offset(0, 2).value = ""

                 End If

        Next

End If
'
   
    Sheet1.Protect Password:=getmsgstate
End Sub


'Private Sub Worksheet_Activate()
'    If Sheet1.Range("sheet1.PAN").value = "" Then
'        Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No"
'    End If
'End Sub



Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub
Sub changexyz()
MsgBox Sheet1.Range("Sheet1.Amount").value
End Sub
Sub LockUnlock1398A_Gen()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.ClearContents
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = True
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.SeventhProvisoFlag").value = ""
'Change 01, Satya, 01.10.2022
        
        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents

        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True

        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.DepositAmount").MergeArea.ClearContents

        Range("Sheet1.DepositAmount").MergeArea.Locked = True

        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)





        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)





        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)



        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents

        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True

        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
'End Change

        
'        Sheet201.Visible = xlSheetVisible
'        Sheet202.Visible = xlSheetVisible
    Else
        If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = False
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = "&HCCFFCC"
            'Ayush_29/07
            If Sheet1.Range("sheet1.SeventhProvisoFlag").value = "" Then
                Sheet1.Range("sheet1.SeventhProvisoFlag").value = "(Select)"
            End If
            
        End If
'        Sheet201.Visible = xlSheetHidden
'        Sheet202.Visible = xlSheetHidden
    End If
    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
    
    'Ayush_UR
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Rows("37:37").Hidden = False
    
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet201.Visible = xlSheetVisible
    Sheet202.Visible = xlSheetVisible
    ThisWorkbook.Protect Password:=getmsgstate
    Else
    'Ayush_UR
    Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Rows("37:37").Hidden = True
     
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet201.Visible = xlSheetHidden
    Sheet202.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub


Sub PartAGenChangeSub1(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False


'Ayush_28/11/2024
 If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Then

 mIntCells = Range("Sheet1.Nature").count
 mIntCells = Range("Sheet1.Nature").Rows.count
    Set rangecells = Range("Sheet1.Nature").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
  '  MsgBox rangecells.item(mIntCtr, 1).value
            If (rangecells.item(mIntCtr, 1).value = "" Or rangecells.item(mIntCtr, 1).value = "(Select)") Then
                 
                 If Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "H")) < Sheet1.Range(Replace(Target.Address, "AA", "H")) Then
                 Sheet1.Unprotect Password:=getmsgstate
                 rangecells.item(mIntCtr, 1).value = Sheet1.Range(Replace(Target.Address, "AA", "I"))
                 Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "Z")).value = Sheet1.Range(Replace(Target.Address, "AA", "Z"))
                 Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "I", "AA")) = Target.value
                 Sheet1.Range(Replace(Target.Address, "AA", "I")) = "(Select)"
                 Sheet1.Range(Replace(Target.Address, "AA", "Z")) = ""
                 Target.value = ""
                 
            End If
            End If
    Next
    
    
     

End If
'--------------------
If Not Intersect(Range("sheet1.ReturnFileSec"), Range(Target.Address)) Is Nothing Then
'Ayush_25-26
    ID_Selection_ReturnFileSec
    
    
    LockUnlock1398A_Gen
'Change-28.11.2022.103.08
    Tax_LockUnlockTRP
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg"), Sheet1.Range(Target.Address)) Is Nothing Then
    Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Range("clauseiv7provisio139iFlg").value, 1, 1) = "N") Or Range("clauseiv7provisio139iFlg").value = "(Select)" Then
        Application.EnableEvents = False
        
        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
       ' 'Application.EnableEvents = True
    ElseIf (Mid(Range("clauseiv7provisio139iFlg").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_1").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_2").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_3").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_4").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)
        ''Application.EnableEvents = True
    End If
    
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_1"), Sheet1.Range(Target.Address)) Is Nothing Then
    Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "N") Then
    Application.EnableEvents = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)
    ''Application.EnableEvents = True
    ElseIf (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "Y") Then
    Application.EnableEvents = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)
    ''Application.EnableEvents = True
    End If
    Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_2"), Sheet1.Range(Target.Address)) Is Nothing Then
    Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "N") Then
    Application.EnableEvents = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)
    'Application.EnableEvents = True
    ElseIf (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "Y") Then
    Application.EnableEvents = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)
    'Application.EnableEvents = True
    End If
    Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_3"), Sheet1.Range(Target.Address)) Is Nothing Then
    Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "N") Then
    Application.EnableEvents = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
    'Application.EnableEvents = True
    ElseIf (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y") Then
    Application.EnableEvents = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)
    'Application.EnableEvents = True
    End If
    Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_4"), Sheet1.Range(Target.Address)) Is Nothing Then
    Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
    ElseIf (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)
    End If
    Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_1"), Sheet1.Range(Target.Address)) Is Nothing Then
    If (Sheet1.Range("clauseiv7provisio139iAmount_1").value < 6000000 Or Sheet1.Range("clauseiv7provisio139iAmount_1").value = 6000000 Or Sheet1.Range("clauseiv7provisio139iAmount_1").value = "") Then
        fmsgbox "* Please Select 'No' if total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year" & Chr(10) & Chr(13)
        Sheet1.Range("clauseiv7provisio139iAmount_1").value = ""
    End If
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_2"), Sheet1.Range(Target.Address)) Is Nothing Then
    If (Sheet1.Range("clauseiv7provisio139iAmount_2").value < 1000000 Or Sheet1.Range("clauseiv7provisio139iAmount_2").value = 1000000 Or Sheet1.Range("clauseiv7provisio139iAmount_2").value = "") Then
        fmsgbox "* Please Select 'No' if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year" & Chr(10) & Chr(13)
        Sheet1.Range("clauseiv7provisio139iAmount_2").value = ""
    End If
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_3"), Sheet1.Range(Target.Address)) Is Nothing Then
    '        If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 25000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
    '            fmsgbox "* Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees" & Chr(10) & Chr(13)
    '            Sheet1.Range("clauseiv7provisio139iAmount_3").value = " "
    '        End If
    
    age = mdCommon.calculateAge(Sheet1.Range("sheet1.DOB").value)
    If age >= 60 Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 50000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
        '        PAG_E29 AY 2023-24 Old
        
'            fmsgbox "* Please Select        'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)

            'PAG_C29 AY 2024-25 Change
            
            fmsgbox "* Please Select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for Senior citizen ) " & Chr(10) & Chr(13)
            Application.EnableEvents = False
            Sheet1.Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
            'Application.EnableEvents = True
        End If
    Else
        If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 25000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
'        PAG_E29 AY 2023-24 Old

'            fmsgbox "* Please Select        'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)

            'PAG_C29 AY 2024-25 Change
            fmsgbox "* Please Select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for Senior citizen and super senior citizen) " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_3").value = ""
        End If
    End If
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_4"), Sheet1.Range(Target.Address)) Is Nothing Then
    If (Sheet1.Range("clauseiv7provisio139iAmount_4").value < 5000000 Or Sheet1.Range("clauseiv7provisio139iAmount_4").value = "") Then
'Ayush_Legacy14
'        fmsgbox "* Please Select        'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(10) & Chr(13)
        fmsgbox "* Please Select 'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(10) & Chr(13)
        Sheet1.Range("clauseiv7provisio139iAmount_4").value = ""
    End If
End If

'dpk
'Ayush_15/01/2026_V0.2_commented
'If Not Application.Intersect(Sheet1.Range("Increliefus89A"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'    If Sheet1.Range("Increliefus89A").value > Sheet1.Range("IncomeNotified89A").value Then
'    'Ayush_Legacy11
''        fmsgbox "* Relief u/s 89A  cannot be claimed more than amount entered at        'income from retirement benefit account maintained in a notified country u/s 89A'"
'        fmsgbox "* Relief u/s 89A cannot be claimed more than amount entered at 'income from retirement benefit account maintained in a notified country u/s 89A'"
'        Sheet1.Range("Increliefus89A").value = 0
'    End If
'    If Range("Increliefus89A").value > 0 Then
'        fmsgbox "* Please ensure To submit Form 10EE."
'    End If
'End If
'Ayush_commented_28/01
'If Not Application.Intersect(Sheet1.Range("OSIncreliefus89A"), Sheet1.Range(Target.Address)) Is Nothing Then
'    Sheet1.Unprotect Password:=getmsgstate
'    If Sheet1.Range("OSIncreliefus89A").value > Sheet1.Range("OSIncomeNotified89A").value Then
'
'    'Error Message Changed By Jyoti for the SIT-67675 on 25/06/2024
' 'fmsgbox "*  Relief u/s 89A cannot be claimed more than income offered in income from retirement benefit account maintained in a notified country u/s 89A."
'    fmsgbox "*Relief u/s 89A cannot be claimed more than income offered in Income from retirement benefit account maintained in a notified country u/s 89A."
'
'        Sheet1.Range("OSIncreliefus89A").value = 0
'    End If
'    If Range("OSIncreliefus89A").value > 0 Then
'        fmsgbox "* Please ensure To submit Form 10EE."
'    End If
'End If
Sheet1.Protect Password:=getmsgstate

If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" And Sheet8.Range("bacValue").value = 1 Then
    
    'Sheet1.Range(Target.Address).value = "(Select)"
    
    fmsgboxsmall "* In Case of Firm, New tax regime u/s 115BAC Is Not applicable."
    Sheet8.Unprotect Password:=getmsgstate
    Sheet8.Range("bacValue").Clear
    Sheet8.Protect Password:=getmsgstate
    Sheet1.Shapes("BacYes").OLEFormat.Object.value = 0
    Sheet1.Shapes("BacNO").OLEFormat.Object.value = 0
    
End If

'For 115BAC(6) 2024-25 Change New Code Development


'PAG_C22 2024-25 Change

'Commented by Bindu as on 25th Jan 2025 as per DE V2 below logic changed

'  If Not Application.Intersect(Sheet1.Range("Sheet1.OptOutNewTaxRegime"), Sheet1.Range(target.Address)) Is Nothing Then
'
'
'       'Ayush_25-26
''       If (Target.value = "No") Then
'         If (target.value = "New Tax Regime") Then
'
'
'            Sheet8.Unprotect Password:=getmsgstate
'            Sheet8.Range("oldbacValue").value = 1
'            Sheet8.Protect Password:=getmsgstate
'    '        reset10IEAy23_24
'
'            NewBacYesValueChange
'
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.value = ""
'            Sheet1.Unprotect Password:=getmsgstate
'
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.value = ""
'
'            Sheet1.Protect Password:=getmsgstate
'
'
''    ElseIf Target.value = "Yes-within the due date" Then
'         ElseIf target.value = "Old Tax Regime" Then
'
'        Sheet8.Unprotect Password:=getmsgstate
'        Sheet8.Range("oldbacValue").value = 2
'        Sheet8.Protect Password:=getmsgstate
''        reset10IEAy23_24
'        NewBacNoValueChange_2425
'
'       LockUnlock80DU
'       LockUnlock80D
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Interior.Color = "&HCCFFCC"
'
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Interior.Color = "&HCCFFCC"
'
'        Sheet1.Protect Password:=getmsgstate
'
''    ElseIf Target.value = "Yes-but beyond the due date" Then
'    ElseIf target.value = "New Tax Regime" Then
'
'        Sheet8.Unprotect Password:=getmsgstate
'        Sheet8.Range("oldbacValue").value = 1
'        Sheet8.Protect Password:=getmsgstate
''       reset10IEAy23_24
'
'        NewBacYesValueChange
'
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Interior.Color = "&HCCFFCC"
'
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Interior.Color = "&HCCFFCC"
'
'        Sheet1.Protect Password:=getmsgstate
'
'    End If
''ID_Nature_Employment_New Target
'  End If

'-------------------------Today---------------------------------------------------

'Commented by Bindu on 25th Jan 2025 as Continue to opt/Opting in now - NA in Ay 25-26 DE sheet
'If Not Application.Intersect(Sheet1.Range("OptingNewTaxRegime"), Sheet1.Range(target.Address)) Is Nothing Then
'    '     If Sheet8.Range("oldbacValue") = "" Then
'    '        Target.value = "(Select)"
'    '     End If
'
'    If (target.value = "Continue to opt") Or (target.value = "Opting in now") Then
'
'        NewBacYesValueChange
'    Else
'
'        NewBacNoValueChange
'    End If
'
''Change-09.12.2022.103.36
''    ID_Nature_Employment Target
'
'End If

If Not Application.Intersect(Sheet1.Range("Sheet1.Nature"), Sheet1.Range(Target.Address)) Is Nothing Then
    Dim nOfCountnature  As Long
    nOfCountnature = 0
    Dim noi
    
    For Each noi In Sheet1.Range("Sheet1.Nature").Cells
        If (Trim(CStr(noi.value)) = CStr(Trim(Target.value))) And (Target.value <> "(Select)") And (Target.value <> "") And (CStr(Trim(Target.value)) <> "Any Other") Then
            nOfCountnature = nOfCountnature + 1
            
        End If
    Next noi
    
    If nOfCountnature > 1 Then
        '        'MsgBox "Drop down has already been selected in Sheet: Income Details", vbOKOnly, "Error(s)!"
        
        fmsgbox "* Nature of Exempt Allowance has already been selected in Sheet Income Details"
        Target.value = "(Select)"
    End If
    ' nOfCountnature = 0
    If (Sheet1.Range("Count4").value > 1) Then
        fmsgbox " *You can only Claimed one of the 3 Exempt allowances from u/s 10(10B)-First proviso, 10(10B)-Second Proviso And 10(10C) "
        Target.value = "(Select)"
    End If
    

End If

'Change-08.12.2022.103.37
If Not Application.Intersect(Sheet1.Range("Sheet1.Amount"), Sheet1.Range(Target.Address)) Is Nothing Then
ID_Allowance_Amount Target
'    If (Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(13A)-Allowance To meet expenditure incurred On house rent" Then
'        Sheet1.Unprotect Password:=getmsgstate
'        If Sheet1.Range(Target.Address).value > (0.5 * Sheet1.Range("IncD.IncomeFromSal_ii").value) Then
'            fmsgbox ("*Sec 10(13A)-Allowance To meet expenditure incurred On house rent can Not exceed 50% of Salary 17(1)")
'            Sheet1.Range(Target.Address).value = ""
'            Sheet1.Range(Target.Address).Select
'        End If
'
'    End If
End If
'AYUSH---------------------------
'If Not Application.Intersect(Sheet1.Range("Sheet1.Amount"), Sheet1.Range(Target.Address)) Is Nothing Then
'
' If Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then '
' If Target.value >= 2500000 And (Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'
'  fmsgbox " *Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is only upto Rs 25 lakh."
'
' End If
' End If
' End If

'-----------------------------------
'Change-25.01.2023.103.ID.37
If Not Application.Intersect(Sheet1.Range("Sheet1.Amount1"), Sheet1.Range(Target.Address)) Is Nothing Then
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Family pension" Then
        Worksheet_Change Sheet1.Range("IncD.LessDeduction57")
    End If
End If
'End Change

endline:

End Sub

Sub PartAGenChangeSub2(ByVal Target As Range)

Application.EnableEvents = False
On Error GoTo endline

If Not Application.Intersect(Sheet1.Range("Sheet1.Amount"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    Sheet1.Unprotect Password:=getmsgstate
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
        
        If Sheet1.Range(Target.Address).value > 500000 Then
            
            fmsgbox ("*For Sec 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can Not exceed Rs.5 lakhs")
            
            Sheet1.Range(Target.Address).value = ""
            
            Sheet1.Range(Target.Address).Select
            
        End If
        
    End If
    
    'Newly added by Bindu 10(10B) Sencond Provison
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10B) Second proviso - Compensation limit notified by CG in the Official Gazette" Then
        
        If Sheet1.Range(Target.Address).value > 500000 Then
            
            fmsgbox ("*For Sec 10(10B)-Second Proviso- Compensation under scheme approved by the Central Government can Not exceed Rs.5 lakhs")
            
            Sheet1.Range(Target.Address).value = ""
            
            Sheet1.Range(Target.Address).Select
            
        End If
        
    End If
    
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(14)(ii)-Transport allowance granted To certain physically handicapped assessee" Then
        
        If Sheet8.Range("bacValue").value = 1 Then
            
            If Sheet1.Range(Target.Address).value > 38400 Then
                
                fmsgbox ("*For Section 10(14)(ii) -  Transport allowance granted To certain physically handicapped assessee Is selected than amount should Not exceed 38400")
                
                Sheet1.Range(Target.Address).value = ""
                
                Sheet1.Range(Target.Address).Select
                
            End If
            
        End If
        
    End If
    
'    If Range("Sheet1.NatureofEmployment").value <> "State Government" And Range("Sheet1.NatureofEmployment").value <> "Central Government" And Range("Sheet1.NatureofEmployment").value <> "Pensioners-SG" And Range("Sheet1.NatureofEmployment").value <> "Pensioners-CG" Or Range("Sheet1.NatureofEmployment").value <> "Pensioners-PSU" Or Range("Sheet1.NatureofEmployment").value <> "Pensioners-Other" Then
'
'        If (Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'
'            'If Sheet8.Range("bacValue").value = 1 Then
'
'            'Newly changed by Bindu
'            'If Sheet1.Range(Target.Address).value > 300000 Then
'            If Sheet1.Range(Target.Address).value > 2500000 Then
'                'fmsgbox ("*Deduction allowable For Employees other than Central Or State Governments CG - Pensioners Or SG - Pensioners Is only Rs 3 lakh")
'                fmsgbox ("*Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is only Rs  25 lakh")
'                Sheet1.Range(Target.Address).value = ""
'
'                Sheet1.Range(Target.Address).Select
'
'            End If
'
'        End If
'
'
'    End If
    
    Sheet1.Protect Password:=getmsgstate
    
End If

If Target.Validation.Type = 3 Then
    
    If Not Intersect(Range("sheet1.NatureofEmployment"), Range(Target.Address)) Is Nothing Then
    
        'Change-06.04.2023.103.02
            AnyOtherDeductions80CCH Target
        'End Change
        
        LockUnlockDed
        'Change-29.11.2022.103.28
        IC_LockUnlockGrossIncome Target
        
'       ID_Nature_Employment Target
        
        '        Change.27.01.2023.102.CDS02
'            Sheet1.Range("IncD.IncomeFromSal_ii").MergeArea.ClearContents
'            Sheet1.Range("IncD.IncomeFromSal_iii").MergeArea.ClearContents
'            Sheet1.Range("IncD.IncomeFromSal_iv").MergeArea.ClearContents
'            Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.ClearContents
'            Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.ClearContents
'            Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.ClearContents
'            Sheet1.Range("IncomeNotifiedOther89A").MergeArea.ClearContents
'            Sheet1.Range("Sheet1.Nature").value = "(Select)"
'            Sheet1.Range("Sheet1.Description").value = "Not Applicable"
'            Sheet1.Range("Sheet1.Amount").ClearContents
'            Sheet1.Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
'           Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.ClearContents
'          End Change


'Ayush_28/11/2024
      resetNatureexempt
           
   End If

End If

'Ayush_25/06**********************************
If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("sheet1.NatureofEmployment"), Range(Target.Address)) Is Nothing Then
    Set tttNOI = Range("Sheet1.Nature")
    rrCount = 0
            For Each rr In tttNOI.Rows
                    If Trim(Range("I" & rr.row).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
                    rrCount = rrCount + Range("AA" & rr.row).value
                            If Sheet1.Range("sheet1.NatureofEmployment").value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then
                                    If rrCount > 2000000 Then
                                    fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 20 Lakhs where nature of employment is selected as  Public Sector Undertaking,PSU - Pensioners, others - Pensioners, Others" & Chr(13))
                                    Range("AA" & rr.row).value = ""
                                    End If
                          'By Bindu as per DE V0.4 on 24th feb 26
                           ElseIf Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG" Then
                                   'Ayush_11/02/2026_V0.3
                           ' ElseIf Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958" Then
                                    If rrCount > 2500000 Then
                                     fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 25 Lakhs where nature of employment is selected as CG, SG, SG-Pensioners, CG-Pensioners" & Chr(13))
                                    'By Bindu as per DE V0.4 on 24th Feb 26
                                    'fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 25 Lakhs where nature of employment is selected as CG, ""Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958"", SG, SG-Pensioners, CG-Pensioners" & Chr(13))
                                    Range("AA" & rr.row).value = ""
                                    End If
                            End If
                    End If
            Next
End If

'********************************************

'Commented by Ayush_26/06


'If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Then
'
'If Sheet1.Range("sheet1.NatureofEmployment").value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then
'
'    If Trim(Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
'
'        If Sheet1.Range(Target.Address).value > 2000000 Then
'
'            fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should Not exeed Rs.20 Lakhs " & Chr(13))
'            fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 20 Lakhs where nature of employment is selected as  Public Sector Undertaking,PSU - Pensioners, others - Pensioners, Others" & Chr(13))
'           Amount u/s 10(10) can not exceed Rs. 20 Lakhs where nature of employment is selected as  Public Sector Undertaking,PSU - Pensioners, others - Pensioners, Others
'            Range(Target.Address).value = ""
'
'            Range(Target.Address).Select
'
'
'
'        Else
'
'            Set tttNOI = Range("Sheet1.Nature")
'
'            rrCount = 0
'
'            For Each rr In tttNOI.Rows
'
'                If Trim(Range("I" & rr.row).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
'
'                    rrCount = rrCount + Range("AO" & rr.row).value
'
'                End If
'
'            Next
'
'            If rrCount > 2000000 Then
'
'                fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should Not exeed Rs.20 Lakhs " & Chr(13))
'                fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 20 Lakhs where nature of employment is selected as  Public Sector Undertaking,PSU - Pensioners, others - Pensioners, Others" & Chr(13))
'
'                Range(Target.Address).value = ""
'
'                Range(Target.Address).Select
'
'            End If
'
'        End If
'
'    End If
'
'
'  End If
'
'
'End If
'
'
'Ayush_17/06
'If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Then
'
'If Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG" Then
'
'    If Trim(Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
'
'        If Sheet1.Range(Target.Address).value > 2500000 Then
'
'            fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should Not exeed Rs.25 Lakhs " & Chr(13))
'            fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 25 Lakhs where nature of employment is selected as CG, SG, SG-Pensioners, CG-Pensioners" & Chr(13))
'
'            Range(Target.Address).value = ""
'
'            Range(Target.Address).Select
'
'
'
'        Else
'
'            Set tttNOI = Range("Sheet1.Nature")
'
'            rrCount = 0
'
'            For Each rr In tttNOI.Rows
'
'                If Trim(Range("I" & rr.row).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
'
'                    rrCount = rrCount + Range("AO" & rr.row).value
'
'                End If
'
'            Next
'
'            If rrCount > 2500000 Then
'
'                fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should Not exeed Rs.25 Lakhs " & Chr(13))
'                fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 25 Lakhs where nature of employment is selected as CG, SG, SG-Pensioners, CG-Pensioners" & Chr(13))
'                Range(Target.Address).value = ""
'
'                Range(Target.Address).Select
'
'            End If
'
'        End If
'
'    End If
'
'End If
'
'
'End If
'
'Change-06.04.2023.103.03
'If Not Application.Intersect(Sheet1.Range("IncD.AnyOther"), Sheet1.Range(Target.Address)) Is Nothing Then
'Sheet1.Unprotect Password:=getmsgstate
'    If Target.value = "80CCH-Contribution to Agnipath Scheme" Then
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.value = 0
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = False
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = (&HCCFFCC)
'    Else
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.value = 0
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = True
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = (&HD8D8D8)
'    End If
'Sheet1.Protect Password:=getmsgstate
'End If
'End Change

If Target.Validation.Type = 1 Then
    
    If Target.name.name = "IncD.Section80C" Or Target.name.name = "IncD.Section80CCC" Or _
       Target.name.name = "IncD.Section80CCD_SE" Or Target.name.name = "IncD.Section80CCD1B_SE" Or _
       Target.name.name = "IncD.Section80CCD" Or _
       Target.name.name = "IncD.Section80DDB" Or Target.name.name = "IncD.Section80E" Or _
       Target.name.name = "IncD.Section80EE" Or Target.name.name = "IncD.Section80EEA" Or _
       Target.name.name = "IncD.Section80EEB" Or Target.name.name = "IncD.Section80GG" Then
    
End If

End If

If Not Intersect(Range("Sheet1.Description"), Range(Target.Address)) Is Nothing Then
    
    If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
        
        fmsgboxsmall ("*Description should Not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
        
        Range(Target.Address).value = ""
        
        Range(Target.Address).Select
        
    End If
    
End If

If Not Intersect(Range("Sheet1.Description1"), Range(Target.Address)) Is Nothing Then
    
    If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
        
        fmsgboxsmall ("*Description should Not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
        
        Range(Target.Address).value = ""
        
        Range(Target.Address).Select
        
    End If
    
End If

If Target.Validation.Type = 1 Then
    
    If Target.name.name = "IncD.Section80C" Or Target.name.name = "IncD.Section80CCC" Or _
       Target.name.name = "IncD.Section80CCD_SE" Or Target.name.name = "IncD.Section80CCD1B_SE" Or _
       Target.name.name = "IncD.Section80CCD" Or _
       Target.name.name = "IncD.Section80DDB" Or Target.name.name = "IncD.Section80E" Or _
       Target.name.name = "IncD.Section80EE" Or Target.name.name = "IncD.Section80EEA" Or _
       Target.name.name = "IncD.Section80EEB" Or Target.name.name = "IncD.Section80GG" Then
    
End If

End If

If Not Intersect(Range("Sheet1.Description"), Range(Target.Address)) Is Nothing Then
    
    If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
        
        fmsgboxsmall ("*Description should Not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
        
        Range(Target.Address).value = ""
        
        Range(Target.Address).Select
        
    End If
    
End If

If Not Intersect(Range("Sheet1.Description1"), Range(Target.Address)) Is Nothing Then
    
    If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
        
        fmsgboxsmall ("*Description should Not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
        
        Range(Target.Address).value = ""
        
        Range(Target.Address).Select
        
    End If
    
End If

If (Range("IncD.IncomeFromSal_iva").value) > (Range("sheet1.minvalue").value) Then
'Ayush_27/01
'If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Then
If Sheet8.Range("bacValue").value = 2 Then


    
    fmsgboxsmall "* Deduction u/s 16(ia) can Not exceed Rs. 50,000 Or Net Salary whichever Is lower in Sheet Income Details"
    
    'PAG_C35 AY 2023-24 Old
    'Range("IncD.IncomeFromSal_iva").Formula = "=IF(bacValue=1,0,MIN(IncD.IncomeFromSal_net,50000))"
    
    'PAG_C35 AY 2024-25 Change
'    Range("IncD.IncomeFromSal_iva").Formula = "=MIN(IncD.IncomeFromSal_net,50000)"
'Ayush_27/01
    Range("IncD.IncomeFromSal_iva").Formula = "=IF(Return_Dropdown_Regime=1,MIN(IncD.IncomeFromSal_net,75000),IF(Return_Dropdown_Regime=2,MIN(IncD.IncomeFromSal_net,50000),0))"
  
    '=IF(Return_Dropdown_Regime=1,MIN(IncD.IncomeFromSal_net,75000),IF(Return_Dropdown_Regime=2,MIN(IncD.IncomeFromSal_net,50000),0))
    
    'Ayush_25-26
'    ElseIf Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime" Then
'Ayush_27/01
    ElseIf Sheet8.Range("bacValue").value = 1 Then
    fmsgboxsmall "* Deduction u/s 16(ia) can Not exceed Rs. 75,000 Or Net Salary whichever Is lower in Sheet Income Details"
    Range("IncD.IncomeFromSal_iva").Formula = "=IF(Return_Dropdown_Regime=1,MIN(IncD.IncomeFromSal_net,75000),IF(Return_Dropdown_Regime=2,MIN(IncD.IncomeFromSal_net,50000),0))"
End If
End If

'Ayush_25-26
'If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Then
'Ayush_27/01
If Sheet8.Range("bacValue").value = 2 Then

'If (Range("IncD.LessDeduction57").value) > 15000 Or (Range("IncD.LessDeduction57").value) > (Range("temp57").value) / 3 Then
If (Range("IncD.LessDeduction57").value) > (Range("IncD.LessDeduction57New").value) Then
    
    fmsgboxsmall "* ""Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension Or Rs. 15,000 in Sheet Income Details"""
    
    Range("IncD.LessDeduction57").value = Range("IncD.LessDeduction57New").value
    Range("IncD.LessDeduction57").Select
    
    End If
    
'    ElseIf Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime" Then
'Ayush_27/01
    ElseIf Sheet8.Range("bacValue").value = 1 Then
    If (Range("IncD.LessDeduction57").value) > (Range("IncD.LessDeduction57New").value) Then
    
    fmsgboxsmall "* ""Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension Or Rs. 25,000 in Sheet Income Details"""
    
    Range("IncD.LessDeduction57").value = Range("IncD.LessDeduction57New").value
    Range("IncD.LessDeduction57").Select
    End If
    
    
    

End If


If Not Application.Intersect(Sheet1.Range("Sheet1.Nature"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    Sheet1.Range("AA" & Target.row & ":AV" & Target.row).value = ""
    
    Dim ChangeAddress As Variant
    
    sPassword = getmsgstate
    
    ChangeAddress = Target.MergeArea.Address
    
    ChangeAddress = Replace(Target.MergeArea.Address, "I", "Z")
    
    ChangeAddress = Replace(ChangeAddress, "Y", "AN")
    
    If Sheet1.Range(Target.Address).value = "Any Other" Then
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet1.Range(ChangeAddress).ClearContents
        
        Sheet1.Range(ChangeAddress).Interior.Color = (&HCCFFCC)        ', "AU"
        
        Sheet1.Range(ChangeAddress).Locked = False
        
        Sheet1.Protect Password:=sPassword
        
    Else
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet1.Range(ChangeAddress).value = "Not Applicable"
        
        Sheet1.Range(ChangeAddress).Interior.Color = (&HD8D8D8)
        
        Sheet1.Range(ChangeAddress).Locked = True
        
        Sheet1.Protect Password:=sPassword
        
    End If
    
    If Mid(Sheet1.Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then
        
        Sheet1.Unprotect Password:=sPassword
        
        '                Sheet1.Range("IncD.Section80GG").MergeArea.value = ""
        
        Sheet1.Range("IncD.Section80GG").Interior.Color = (&HCCFFCC)        ', "AU"
        
        Sheet1.Range("IncD.Section80GG").MergeArea.Locked = False
        
        Sheet1.Protect Password:=sPassword
        
    Else
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet1.Range("IncD.Section80GG").MergeArea.value = ""
        
        Sheet1.Range("IncD.Section80GG").MergeArea.Interior.Color = (&HCCFFCC)        ', "AU"
        
        Sheet1.Range("IncD.Section80GG").MergeArea.Locked = False
        
        Sheet1.Protect Password:=sPassword
        
    End If
    
End If


If Not Application.Intersect(Sheet1.Range("Sheet1.Nature1"), Sheet1.Range(Target.Address)) Is Nothing Then

     
     
'    If Sheet8.Range("BacValue").value = 2 Then 'PAG_E41 AY 2023-24 Old

        If Range("temp5").value > 0 And Not isdropdownblank(Range(Target.Address).value) Then

            Sheet1.Unprotect Password:=getmsgstate

            Range("IncD.LessDeduction57").MergeArea.Locked = False

            Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Protect Password:=getmsgstate

        Else

            Sheet1.Unprotect Password:=getmsgstate

            'DPKPnch

            Range("IncD.LessDeduction57").MergeArea.ClearContents

            Range("IncD.LessDeduction57").MergeArea.Locked = True

            Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Protect Password:=getmsgstate

'        End If
    End If

 

 
    sPassword = getmsgstate
    
    ChangeAddress = Target.MergeArea.Address
    
    ChangeAddress = Replace(Target.MergeArea.Address, "I", "Z")
    
    ChangeAddress = Replace(ChangeAddress, "Y", "AN")
    
    If Sheet1.Range(Target.Address).value = "Any Other" Then
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet1.Range(ChangeAddress).value = ""
        
        Sheet1.Range(ChangeAddress).Interior.Color = (&HCCFFCC)        ', "AU"
        
        Sheet1.Range(ChangeAddress).Locked = False
        
        Sheet1.Protect Password:=sPassword
        
    Else
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet1.Range(ChangeAddress).value = "Not Applicable"
        
        Sheet1.Range(ChangeAddress).Interior.Color = (&HD8D8D8)
        
        Sheet1.Range(ChangeAddress).Locked = True
        
        Sheet1.Protect Password:=sPassword
        
    End If
    
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
        
        If Mid(Sheet1.Range(Target.Address).value, 1, 1) = "F" Then
            
            fmsgboxsmall "*  Family pension should Not allowed To be selected by a Firm(Other than LLP)in Sheet Income Details"        ', vbOKOnly, "Error(s)!"
            
            Target.value = "(Select)"
            
        Else
            
        End If
        
    End If
    
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
        
        If Mid(Sheet1.Range(Target.Address).value, 1, 1) = "F" Then
            
            fmsgboxsmall "*  Family pension should Not allowed To be selected by a HUF in Sheet Income Details"        ', vbOKOnly, "Error(s)!"
            
            Target.value = "(Select)"
            
        End If
        
    End If
    
End If

Dim nOfCount1       As Long

nOfCount1 = CountDropdown1(Trim(Target.value), Sheet1.Range("Sheet1.Nature1").Cells, Sheet1.Range("Sheet1.Nature1").count)

If currCnt22 > 1 Or currCnt33 > 1 Or currCnt44 > 1 Or currCnt66 > 1 Or currCnt67 > 1 Or currCnt68 > 1 Or currCnt69 > 1 Or currCnt70 > 1 Then
    
    fmsgboxsmall "* Drop down has already been selected in Sheet Income Details"        ', vbOKOnly, "Error(s)!"
    
    Target.value = "(Select)"
    
End If
'Commented as per the Test case on 02-01-2024
'If (Sheet1.Range("CountOS5").value > 1) Then
'    'fmsgbox " *You can only Claimed one of the 4 Exempt allowances from u/s 10(10B)-First proviso, 10(10B)-Second Proviso and 10(10C) "
'    fmsgbox " * Please Select one of the income out of Income u/s 10(11) - First proviso Or Income u/s 10(11) - Second proviso Or Income u/s 10(12) - First proviso Or Income u/s 10(12) - Second proviso"
'    Target.value = "(Select)"
'End If

If Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234A"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "IncD.IntrstPayUs234A" Then
    
    CalculateTaxCheckFlag = True
    
End If

Sheet1.Unprotect Password:=getmsgstate

If Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234B"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "IncD.IntrstPayUs234B" Then
    
    CalculateTaxCheckFlag = True
    
End If
Sheet1.Protect Password:=getmsgstate

Sheet1.Unprotect Password:=getmsgstate

If Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234C"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "IncD.IntrstPayUs234C" Then
    
    CalculateTaxCheckFlag = True
    
End If
Sheet1.Protect Password:=getmsgstate

Sheet1.Unprotect Password:=getmsgstate
If Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234F"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "IncD.IntrstPayUs234F" Then
    
    CalculateTaxCheckFlag = True
    
End If
Sheet1.Protect Password:=getmsgstate


endline:
End Sub


Sub PartAGenChangeSub3(ByVal Target As Range)

On Error GoTo endline

'HP income Validation on Data Entry
'Ayush_commented_27/01
'If (Target.name.name) = "IncD.IncomeFromHP_v" Then
'
'    If Range("IncD.TypeOfHP").value = "Let Out" Then
'
'        If Range("LossCheckHP").value < -200000 Then
'            'Error Message Changed By Ayush for the SIT-56341
''            fmsgboxsmall "* Note : Maximum Loss from House Property that can be set-off Is INR 2,00,000. in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
'             fmsgboxsmall "* Note : The amount cannot exceed Rs. 2 Lakhs" ', vbOKOnly, mdCommon.vbMessageCaption
'
'        End If
'
'    End If
'
'End If
'Ayush_commented_27/01
'If (Target.name.name) = "IncD.IncomeFromHP_v" Then
'
'    Dim hpInterest, salIncome As Variant
'
'    hpInterest = Range("IncD.IncomeFromHP_v").value
'
'    If Range("IncD.TypeOfHP").value = "Self Occupied" Then
'
'        mdIncomeDetails.CheckHousePropertyIncome (hpInterest)
'
'    End If
'
'End If

'If (Target.name.name) = "sheet1.NatureofEmployment" Then
'
'    Sheet1.Unprotect Password:=getmsgstate
'
'    If Range("sheet1.status").value = "I - INDIVIDUAL" And Sheet8.Range("bacValue").value = 2 Then
'
'        If Range("Sheet1.NatureofEmployment").value = "State Government" Or Range("Sheet1.NatureofEmployment").value = "Central Government" Or Range("Sheet1.NatureofEmployment").value = "Public Sector Undertaking" Then
'
'            Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
'
'            Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Range("IncD.Section80CCD").MergeArea.Locked = False
'
'            Range("IncD.Section80CCD").MergeArea.Interior.Color = (&HCCFFCC)
'
'        Else
'
'            Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
'
'            Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
'
'            Range("IncD.Section80CCD").MergeArea.ClearContents
'
'            Range("IncD.Section80CCD").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Range("IncD.Section80CCD").MergeArea.Locked = True
'
'        End If
'
'    End If
'
'
'    Sheet1.Protect Password:=getmsgstate
'
'End If

If (Target.name.name) = "Sheet1.DepositAmount" Then
    
    If Range("Sheet1.DepositAmount").value <= 10000000 Then
        
        fmsgboxsmall "* Amount entered cannot be less than Or equal To Rs. 1 Crore" & Chr(13)
        Application.EnableEvents = False
        Range("Sheet1.DepositAmount").value = ""
        
    End If
    
End If

If (Target.name.name) = "Sheet1.AggrigateAmount" Then
    
    If Range("Sheet1.AggrigateAmount").value <= 200000 Then
        
        fmsgboxsmall "* Amount entered cannot be less than Or equal To Rs. 2 Lakhs" & Chr(13)
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmount").value = ""
        
    End If
    
End If

If (Target.name.name) = "Sheet1.AggrigateAmount1" Then
    
    If Range("Sheet1.AggrigateAmount1").value <= 100000 Then
        
        fmsgboxsmall "* Amount entered cannot be less than Or equal To Rs. 1 Lakh" & Chr(13)
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmount1").value = ""
        
    End If
    
End If

If (Target.name.name) = "sheet1.SeventhProvisoFlag" Then
    
    Sheet1.Unprotect Password:=getmsgstate
    
    If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then
        
        Application.EnableEvents = False
        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
        
        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
        
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
        Application.EnableEvents = True
        
    ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then
        
        Application.EnableEvents = False
        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = False
        
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("Sheet1.DepositAmountFlag").MergeArea.value = "(Select)"
        
        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
        Application.EnableEvents = True
    End If
    
    If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
        
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
        
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
        Application.EnableEvents = True
    ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
        
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("Sheet1.AggrigateAmountFlag").MergeArea.value = "(Select)"
        
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
        Application.EnableEvents = True
    End If
    
    If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
        
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
        
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
        Application.EnableEvents = True
    ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
        
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.value = "(Select)"
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
        Application.EnableEvents = True
    End If
    
    If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then
        Application.EnableEvents = False
        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
        
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True
        
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
        
        Application.EnableEvents = True
    ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = False
        
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("clauseiv7provisio139iFlg").MergeArea.value = "(Select)"
        Application.EnableEvents = True
    End If
    
    Sheet1.Protect Password:=getmsgstate
    
End If

If (Target.name.name) = "Sheet1.DepositAmountFlag" Then
    
    Sheet1.Unprotect Password:=getmsgstate
    
    If (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "N") Then
        Application.EnableEvents = False
        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
        Application.EnableEvents = True
    ElseIf (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
        Application.EnableEvents = True
    End If
    
    Sheet1.Protect Password:=getmsgstate
    
End If

If (Target.name.name) = "Sheet1.AggrigateAmountFlag" Then
    
    Sheet1.Unprotect Password:=getmsgstate
    
    If (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "N") Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
        Application.EnableEvents = True
    ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
        Application.EnableEvents = True
    End If
    
    Sheet1.Protect Password:=getmsgstate
    
End If

If (Target.name.name) = "Sheet1.AggrigateAmountFlag1" Then
    
    Sheet1.Unprotect Password:=getmsgstate
    
    If (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "N") Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
        Application.EnableEvents = True
    ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "Y") Then
        Application.EnableEvents = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
        Application.EnableEvents = True
    End If
    
    Sheet1.Protect Password:=getmsgstate
    
End If

'new insertion

If (Target.name.name) = "IncD.IncomeFromSal_ii" Then

'Ayush_04/02/2025
    fmsgbox ("""Nature of Exempt Allowance(s) will get refreshed if the field ""Salary as per Section 17(1)"" is modified.""")
    Sheet1.Range("Sheet1.Amount").ClearContents
'---------------------------
    
    If (Range("IncD.IncomeFromSal_v").value) > 5000 Or (Range("IncD.IncomeFromSal_v").value) > (Range("IncD.IncomeFromSal_ii").value) / 5 Then
        
        fmsgboxsmall "* Deduction of Entertainment allowance u/s 16(ii) can Not exceed 1/5th of Salary As per section 17(1) Or Rs. 5,000 whichever Is lower in Sheet Income Details"
        Application.EnableEvents = False
        Range("IncD.IncomeFromSal_v").value = ""
        
        Range("IncD.IncomeFromSal_v").Select
        Application.EnableEvents = True
    End If
    
End If

If (Target.name.name) = "IncD.IncomeFromSal_v" Then
    
    If (Range("IncD.IncomeFromSal_v").value) > 5000 Or (Range("IncD.IncomeFromSal_v").value) > (Range("IncD.IncomeFromSal_ii").value) / 5 Then
        
        fmsgboxsmall "* Deduction of Entertainment allowance u/s 16(ii) can Not exceed 1/5th of Salary As per section 17(1) Or Rs. 5,000 whichever Is lower in Sheet Income Details"
        Application.EnableEvents = False
        Range("IncD.IncomeFromSal_v").value = ""
        
        Range("IncD.IncomeFromSal_v").Select
        Application.EnableEvents = True
    End If
    
End If

If (Target.name.name) = "IncD.IncomeFromSal_ivc" Then
    
    If (Range("IncD.IncomeFromSal_ivc").value) > 5000 Then
        
'        fmsgboxsmall "* Deduction of Professional tax u/s 16(iii) cannot exceed Rs. 5,000/ in Sheet Income Details"
        'SIT -56341
        fmsgboxsmall "* Maximum deduction of Professional tax u/s 16(iii) can be claimed of Rs. 5,000"
        Application.EnableEvents = False
        Range("IncD.IncomeFromSal_ivc").value = ""
        
        Range("IncD.IncomeFromSal_ivc").Select
        Application.EnableEvents = True
    End If
    
End If

endline:
End Sub

Sub PartAGenChangeSub4(ByVal Target As Range)

On Error GoTo endline





'Malli>>>>>>>>>>>>>>>>
'If Target.Validation.Type = 3 Or Target.Validation.Type = 6 Then
'If (Target.name.name) = "sheet1.StateCode2" And UCase(Range("sheet1.StateCode2").value) <> "99-FOREIGN" Or _
   (Target.name.name) = "sheet1.PinCode2" Then

   If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("sheet1.PinCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
              
 ' If UCase(Range("sheet1.StateCode2").value) = "99-FOREIGN" Then
  If UCase(Range("sheet1.StateCode2").value) <> "99-FOREIGN" Then
        
        Application.EnableEvents = False
        Range("sheet1.CountryCode2").value = ""
        
        Range("sheet1.CountryCode2").value = "91-INDIA"
        
        'Range("sheet1.PinCode").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        
       ' Sheet1.Range("sheet1.mobileCountryCode2").value = "91"
        
        '
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
        
    'sudha    'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("HASZIP2").MergeArea.value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.Locked = True
        
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
      End If
      
    End If
    
  '  If (Target.name.name) = "sheet1.StateCode2" And UCase(Range("sheet1.StateCode2").value) = "99-FOREIGN" Then
    If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
  If UCase(Range("sheet1.StateCode2").value) = "99-FOREIGN" Then
        Application.EnableEvents = False
        Range("sheet1.CountryCode2").value = ""
        
        'Range("sheet1.PinCode").value = "999999"
        
        Range("sheet1.CountryCode2").value = "(Select)"
        
        Range("sheet1.CountryCode2").Select
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
        
       ' Sheet1.Range("sheet1.mobileCountryCode").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
        
       ' Sheet1.Range("sheet1.mobileCountryCode2").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        
       Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("HASZIP2").MergeArea.value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.Locked = False
        
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
        End If
    End If
    
  '  If (Target.name.name) = "sheet1.StateCode2" And UCase(Range("sheet1.StateCode2").value) = "(SELECT)" Then
    If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
    If UCase(Range("sheet1.StateCode2").value) = "(SELECT)" Then
     
        Application.EnableEvents = False
        Range("sheet1.CountryCode2").value = ""
        
        'Range("sheet1.PinCode").value = ""
        
        Range("sheet1.CountryCode2").value = "(Select)"
        
        Range("sheet1.CountryCode2").Select
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
        
       ' Sheet1.Range("sheet1.mobileCountryCode").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
        
      '  Sheet1.Range("sheet1.mobileCountryCode2").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
        
    'sudha    'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("HASZIP2").MergeArea.value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.Locked = True
        
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
       End If
    End If
'End If 'Ayush_03/04/2026
'<<<<<<<<<<<<<<<<<<<<<
'Ayush_03/04/2026
'If Target.Validation.Type = 3 Then
    
   ' If (Target.name.name) = "sheet1.RepAssessee" Then
            If Not Application.Intersect(Sheet1.Range("sheet1.RepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        '------------------Un Protect--------------------
        
        sPassword = mdCommon.getmsgstate
        
        Sheet1.Unprotect Password:=sPassword
        
        '------------------Un Protect Ends--------------------
        
        If Range("sheet1.RepAssessee").value = "No" Then
            
            Sheet1.Unprotect Password:=sPassword
            Application.EnableEvents = False
            Sheet1.Range("sheet1.NameRepAssessee").value = ""
            
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            
            'Ayush_12/01/2026_V0.2
            Sheet1.Range("sheet1.EmailRepAssessee").value = ""
            
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Range("sheet1.CountryCodeRepAssessee").value = ""
            
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Range("sheet1.MobileRepAssessee").value = ""
            
            Sheet1.Range("sheet1.MobileRepAssessee").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.MobileRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            
            '---------------------
            
            'Ayush_12/01/2026_V0.2_commented
            
'            Sheet1.Range("sheet1.CapacityRep").value = ""
'
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Range("sheet1.AddressRepAssessee").value = ""
'
'            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Range("sheet1.PANRepAssessee").value = ""
'
'            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
'
'            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            Application.EnableEvents = True
        Else
            Application.EnableEvents = False
            Sheet1.Range("sheet1.NameRepAssessee").value = ""
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Locked = False
            
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
            
            'Ayush_12/01/2026_V0.2
            Sheet1.Range("sheet1.EmailRepAssessee").value = ""
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Locked = False
            
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
            
            
            Sheet1.Range("sheet1.CountryCodeRepAssessee").value = ""
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Locked = False
            
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Range("sheet1.MobileRepAssessee").value = ""
            Sheet1.Range("sheet1.MobileRepAssessee").MergeArea.Locked = False
            
            Sheet1.Range("sheet1.MobileRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
            
            '---------------------
            
            
            'Ayush_12/01/2026_V0.2_commented
            
'            Sheet1.Range("sheet1.CapacityRep").value = "(Select)"
'
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
            Application.EnableEvents = True
        End If
        
        '----------------Lock Password-------------------START---
        
        Sheet17.Protect Password:=sPassword
        
        Sheet1.Protect Password:=sPassword
        
        '----------------Lock Password-------------------END-----
        
    End If
    
 '   If (Target.name.name) = "sheet1.Status" Then
         If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        'Change-06.04.2023.ID.03
        AnyOtherDeductions80CCH Target
        'End Change
        LockUnlock80D
'        LockUnlock115BAC

        LockUnlockClauseiv
        
        
        LockUnlock80DU
        
        'Ayush_13/04
    '    LockUnlock80C
        
        
        If Sheet1.Range("temp5").value > 0 Then
            
            If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
                
                fmsgboxsmall "*  Family pension should Not allowed To be selected by a Firm(Other than LLP)in Sheet Income Details"        ', vbOKOnly, "Error(s)!"
                Application.EnableEvents = False
                Sheet1.Range("sheet1.Status").value = "(Select)"
                Application.EnableEvents = True
            End If
            
        End If
        
        If Sheet1.Range("temp5").value > 0 Then
            
            If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
                
                fmsgboxsmall "*  Family pension should Not allowed To be selected by a HUF in Sheet Income Details"        ', vbOKOnly, "Error(s)!"
                
                Sheet1.Range("sheet1.Status").value = "(Select)"
                
            End If
            
        End If
        
    If Sheet8.Range("oldbacValue").value = 2 Then
    
        If Range("sheet1.status").value = "F - FIRM(Other than LLP)" Then
        
                     
            Sheet14.Unprotect Password:=getmsgstate
            ThisWorkbook.Unprotect Password:=getmsgstate
            Sheet14.Visible = xlSheetHidden
            Sheet14.Protect Password:=getmsgstate
            ThisWorkbook.Protect Password:=getmsgstate
            
            'Ayush
         Else
         'Ayush_25-26
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value <> "No" Then
            'If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value <> "New Tax Regime" Then
            'Ayush_27/01
'            If UCase(Sheet1.Range("Sheet1.OptOutNewTaxRegime").value) = UCase("Old Tax Regime") Then
            If Sheet8.Range("bacValue").value = 2 Then
                Sheet14.Unprotect Password:=getmsgstate
                ThisWorkbook.Unprotect Password:=getmsgstate
                Sheet14.Visible = xlSheetVisible
                Sheet14.Protect Password:=getmsgstate
                ThisWorkbook.Protect Password:=getmsgstate
            End If
        End If
    End If
        
               
        
        If Range("sheet1.status").value = "I - INDIVIDUAL" Or Range("sheet1.status").value = "F - FIRM(Other than LLP)" Then
            
            If Trim(Range("sheet1.DOB").value) = "01/01/0001" Then
                
                fmsgboxsmall "*Status in Sheet Income Details Please enter valid DOB."        ', vbOKOnly, mdCommon.vbMessageCaption
                
                Sheet1.Range("sheet1.DOB").Select
                
                Sheet1.Range("sheet1.DOB").value = ""
                
            End If
            
        End If
        
        '-----------------Un Protect--------------------
        
        sPassword = mdCommon.getmsgstate
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet4.Unprotect Password:=sPassword
        
        Sheet3.Unprotect Password:=sPassword
        
        '------------------Un Protect Ends--------------------
        
        LockNatureofEmployment
        
        LockUnlock80TTB
        
        LockUnlockQVIA
        
        LockUnlock80DDB
        
        LockUnlockAadhaar
        
        
        
        If Mid(Trim(Range("sheet1.Status")), 1, 1) = "F" Or Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
            
            LOCKPI
            
            LOCKTDS1
            
                    If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
                        
                        LOCKEE
                        
                    Else
                
            End If
            
        ElseIf Mid(Trim(Range("sheet1.Status")), 1, 1) = "I" Then 'Newly changed by Bindu on 13/12/2024
            
            UNLOCKPI
            
            UNLOCKTDS1
            
        End If
        
        Sheet1.Unprotect Password:=sPassword
        
        Sheet4.Unprotect Password:=sPassword
        Sheet3.Unprotect Password:=sPassword
        
        If Mid(Trim(Range("sheet1.Status")), 1, 1) = "F" Then
            
            Sheet3.Range("BP_E3_2").Locked = False
            
            
            Sheet3.Range("BP_E3_2").Interior.Color = (&HCCFFCC)
            
            Sheet1.Unprotect Password:=sPassword
            
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.SeventhProvisoFlag").Interior.ColorIndex = 15
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.ClearContents
            
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
            
            Sheet1.Range("Sheet1.DepositAmountFlag").Interior.ColorIndex = 15
            
            Sheet1.Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
            
            Sheet1.Unprotect Password:=sPassword
            
            Sheet1.Range("Sheet1.DepositAmount").MergeArea.Locked = True
            
            Sheet1.Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.DepositAmount").MergeArea.ClearContents
            
            Sheet1.Unprotect Password:=sPassword
            
            Sheet1.Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
            
            Sheet1.Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
            
            Sheet1.Unprotect Password:=sPassword
            
            Sheet1.Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
            Sheet1.Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
            Sheet1.Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
            Sheet1.Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
            Sheet1.Range("IncD.Section80C").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section80C").MergeArea.ClearContents
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("IncD.Section89").MergeArea.Locked = True
            Sheet1.Range("IncD.Section89").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section89").MergeArea.ClearContents
            
        ElseIf Sheet8.Range("bacValue").value = 2 Then
            Sheet1.Unprotect Password:=sPassword
            
            Sheet1.Range("IncD.Section80C").MergeArea.Locked = False
            Sheet1.Range("IncD.Section80C").MergeArea.Interior.Color = (&HCCFFCC)
            'Sheet1.Range("IncD.Section80C").ClearContents
            
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = False
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=sPassword
            Range("Sheet1.DepositAmountFlag").MergeArea.Locked = False
            
            Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
            
            Range("Sheet1.DepositAmountFlag").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=sPassword
            
            'Range("Sheet1.DepositAmount").MergeArea.ClearContents
            
            Range("Sheet1.DepositAmount").MergeArea.Locked = False
            
            Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
            
            '                Range("Sheet1.DepositAmount").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=sPassword
            
            Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
            Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
            Range("Sheet1.AggrigateAmountFlag").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=sPassword
            
            'Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
            
            Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
            
            Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Unprotect Password:=sPassword
            
            Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
            Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)
            Range("Sheet1.AggrigateAmountFlag1").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=sPassword
            
            'Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
            
            Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
            
            Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Unprotect Password:=sPassword
            
            Sheet1.Range("IncD.Section89").MergeArea.Locked = False
            
            'Sheet1.Range("IncD.Section89").MergeArea.ClearContents
            
            Sheet1.Range("IncD.Section89").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Unprotect Password:=sPassword
            
            Sheet3.Range("BP_E3_2").Locked = True
            
            'Sheet3.Range("BP_E3_2").ClearContents
            
            Sheet3.Range("BP_E3_2").Interior.Color = (&HD8D8D8)
            
        End If
        
        'Ayush_19/02
        If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Or Mid(Trim(Range("sheet1.Status")), 1, 1) = "I" Then
            Sheet1.Unprotect Password:=sPassword
              Sheet3.Range("BP_E3_2").Locked = True
             'Sheet3.Range("BP_E3_2").ClearContents
 
              Sheet3.Range("BP_E3_2").Interior.Color = (&HD8D8D8)
              Sheet1.Protect Password:=sPassword
        End If
        
        If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
            
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("IncD.Section89").MergeArea.Locked = True
            Sheet1.Range("IncD.Section89").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("IncD.Section89").MergeArea.ClearContents
            Sheet1.Protect Password:=sPassword
        End If
        
        If (Mid(Range("sheet1.Status").value, 1, 1) = "H") Then
            
            Sheet5.Range("Ver.AssesseeVerName").value = ""
            
            Sheet5.Range("Ver.PAN").value = ""
            
            mdCommon.Lock_TDS
            
        ElseIf (Mid(Range("sheet1.Status").value, 1, 1) = "F") Then
            
            '
            
            Sheet5.Range("Ver.AssesseeVerName").value = ""
            
            Sheet5.Range("Ver.PAN").value = ""
            
            '
            
            mdCommon.Lock_TDS
            
        Else
            
            If Mid(UCase(Trim(Sheet1.Range("sheet1.Status").value)), 1, 1) = "I" And Mid(UCase(Trim(Sheet1.Range("sheet1.RepAssessee").value)), 1, 1) = "N" Then
                
                Sheet5.Range("Ver.PAN").value = UCase(Sheet1.Range("sheet1.PAN").value)
                
            End If
            
        End If
        
        '----------------Lock Password-------------------START---
        
        Sheet1.Protect Password:=sPassword
        
        Sheet4.Protect Password:=sPassword
        
        Sheet3.Protect Password:=sPassword
        
        '----------------Lock Password-------------------END-----
        
    End If
    
    'StateCode & Country Validation: ON data entry
    
'    If (Target.name.name) = "sheet1.StateCode1" And UCase(Range("sheet1.StateCode1").value) <> "99-FOREIGN" Then
    If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
       If UCase(Range("sheet1.StateCode1").value) <> "99-FOREIGN" Then
        Application.EnableEvents = False
        Range("sheet1.CountryCode").value = ""
        
        Range("sheet1.CountryCode").value = "91-INDIA"
        
        'Range("sheet1.PinCode").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("sheet1.mobileCountryCode").value = "91"
        
        '
        
        Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
        
    'sudha    'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("HASZIP").MergeArea.value = ""
        
        Sheet1.Range("HASZIP").MergeArea.Locked = True
        
        Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
      End If
    End If
    
  '  If (Target.name.name) = "sheet1.StateCode1" And UCase(Range("sheet1.StateCode1").value) = "99-FOREIGN" Then
   If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
       If UCase(Range("sheet1.StateCode1").value) = "99-FOREIGN" Then
        Application.EnableEvents = False
        Range("sheet1.CountryCode").value = ""
        
        'Range("sheet1.PinCode").value = "999999"
        
        Range("sheet1.CountryCode").value = "(Select)"
        
        Range("sheet1.CountryCode").Select
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
        
        Sheet1.Range("sheet1.mobileCountryCode").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
        
        Sheet1.Range("sheet1.mobileCountryCode2").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
        
     'sudha   'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("HASZIP").MergeArea.value = ""
        
        Sheet1.Range("HASZIP").MergeArea.Locked = False
        
        Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
        End If
        
    End If
    
'    If (Target.name.name) = "sheet1.StateCode1" And UCase(Range("sheet1.StateCode1").value) = "(SELECT)" Then
If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
       If UCase(Range("sheet1.StateCode1").value) = "(SELECT)" Then
        
        Application.EnableEvents = False
        Range("sheet1.CountryCode").value = ""
        
        'Range("sheet1.PinCode").value = ""
        
        Range("sheet1.CountryCode").value = "(Select)"
        
        Range("sheet1.CountryCode").Select
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
        
        Sheet1.Range("sheet1.mobileCountryCode").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
        
        Sheet1.Range("sheet1.mobileCountryCode2").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
        
    'sudha    'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("HASZIP").MergeArea.value = ""
        
        Sheet1.Range("HASZIP").MergeArea.Locked = True
        
        Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
     End If
    End If
    
    'Country Validation: On Data Entry
    
'Change-21.11.2022.103.03
 '   If (Target.name.name) = "sheet1.CountryCode" Then
 If Not Application.Intersect(Sheet1.Range("sheet1.CountryCode"), Sheet1.Range(Target.Address)) Is Nothing Then

    
    ID_Country
        
'        If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
'
'            Sheet1.Range("sheet1.StateCode1") = "(Select)"
'
'            Sheet1.Range("sheet1.PinCode") = ""
'
'            If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-FOREIGN" Then
'
'                MsgBox "* Country Should Not be India For Selected State in Sheet Income Details" & Chr(13)
'
'                Sheet1.Range("sheet1.CountryCode").value = "(Select)"
'
'                Sheet1.Range("sheet1.CountryCode").Select
'
'                'Sheet1.Range("sheet1.StateCode1").value = ""
'
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'
'            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = "91"
'
'            '
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("HASZIP").MergeArea.value = ""
'
'            Sheet1.Range("HASZIP").MergeArea.Locked = True
'
'            Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Protect Password:=getmsgstate
'
'        ElseIf Sheet1.Range("sheet1.CountryCode").value = "(Select)" Then
'
'            Sheet1.Range("sheet1.StateCode1") = "(Select)"
'
'            Sheet1.Unprotect Password:=getmsgstate
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
'
'            Sheet1.Range("sheet1.mobileCountryCode2").value = ""
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("HASZIP").MergeArea.value = ""
'
'            Sheet1.Range("HASZIP").MergeArea.Locked = True
'
'            Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Protect Password:=getmsgstate
'
'        ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
'
'            If UCase(Sheet1.Range("sheet1.StateCode1").value) <> "99-FOREIGN" And Sheet1.Range("sheet1.StateCode1").value <> "(Select)" Then
'
'                MsgBox "* Country Should be India For Selected State in Sheet Income Details" & Chr(13)
'
'                Sheet1.Range("sheet1.CountryCode").value = ""
'
'            End If
'
'            Sheet1.Unprotect Password:=getmsgstate
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
'
'            Sheet1.Range("sheet1.mobileCountryCode2").value = ""
'
'            '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
'
'            Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Sheet1.Range("HASZIP").MergeArea.value = ""
'
'            Sheet1.Range("HASZIP").MergeArea.Locked = False
'
'            Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
'
'            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Sheet1.Protect Password:=getmsgstate
'
'        End If
        
    End If
    
'    If (Target.name.name) = "HASZIP" Then
 If Not Application.Intersect(Sheet1.Range("HASZIP"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        '            Target.value = UCase(Target.value)
        
        If Mid(Trim(Sheet1.Range("HASZIP").value), 1, 1) = "Y" Then
            
            Sheet1.Unprotect Password:=getmsgstate
            Application.EnableEvents = False
            
            Sheet1.Range("sheet1.ZipCode").MergeArea.value = "XXXXXX"
            
            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Protect Password:=getmsgstate
            
            Application.EnableEvents = True
        Else
            Application.EnableEvents = False
            Sheet1.Unprotect Password:=getmsgstate
            
            Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
            
            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
            
            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Protect Password:=getmsgstate
            Application.EnableEvents = True
        End If
        
    End If
    
    'Ayush_13/01/2026_V0.2
       
  '  If (Target.name.name) = "sheet1.StateCode2" And UCase(Range("sheet1.StateCode2").value) <> "99-FOREIGN" Then
  If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
'Ayush_23/03/2026
'       If UCase(Range("sheet1.StateCode1").value) <> "99-FOREIGN" Then
        If UCase(Range("sheet1.StateCode2").value) <> "99-FOREIGN" Then
       
        Application.EnableEvents = False
        Range("sheet1.CountryCode2").value = ""
        
        Range("sheet1.CountryCode2").value = "91-INDIA"
        
        'Range("sheet1.PinCode").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        
       ' Sheet1.Range("sheet1.mobileCountryCode2").value = "91"
        
        '
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
        
    'sudha    'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("HASZIP2").MergeArea.value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.Locked = True
        
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
     End If
     
    End If
    
  '  If (Target.name.name) = "sheet1.StateCode2" And UCase(Range("sheet1.StateCode2").value) = "99-FOREIGN" Then
   If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
'Ayush_23/03/2026
'       If UCase(Range("sheet1.StateCode1").value) = "99-FOREIGN" Then
        If UCase(Range("sheet1.StateCode2").value) = "99-FOREIGN" Then
        Application.EnableEvents = False
        Range("sheet1.CountryCode2").value = ""
        
        'Range("sheet1.PinCode").value = "999999"
        
        Range("sheet1.CountryCode2").value = "(Select)"
        
        Range("sheet1.CountryCode2").Select
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
        
       ' Sheet1.Range("sheet1.mobileCountryCode").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
        
       ' Sheet1.Range("sheet1.mobileCountryCode2").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        
       Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("HASZIP2").MergeArea.value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.Locked = False
        
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
        End If
        
    End If
    
'    If (Target.name.name) = "sheet1.StateCode2" And UCase(Range("sheet1.StateCode2").value) = "(SELECT)" Then

If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
       If UCase(Range("sheet1.StateCode1").value) = "(SELECT)" Then
        
        Application.EnableEvents = False
        Range("sheet1.CountryCode2").value = ""
        
        'Range("sheet1.PinCode").value = ""
        
        Range("sheet1.CountryCode2").value = "(Select)"
        
        Range("sheet1.CountryCode2").Select
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
        
       ' Sheet1.Range("sheet1.mobileCountryCode").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
        
      '  Sheet1.Range("sheet1.mobileCountryCode2").value = ""
        
        '                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
        
    'sudha    'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)
        
        Sheet1.Range("HASZIP2").MergeArea.value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.Locked = True
        
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
        
        Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Protect Password:=getmsgstate
        
        Application.EnableEvents = True
    End If
    
        
    End If
    
    'Country Validation: On Data Entry
    

   ' If (Target.name.name) = "sheet1.CountryCode2" Then
   If Not Application.Intersect(Sheet1.Range("sheet1.CountryCode2"), Sheet1.Range(Target.Address)) Is Nothing Then

    
    ID_Country2
        

        
    End If
    
  '  If (Target.name.name) = "HASZIP2" Then
     If Not Application.Intersect(Sheet1.Range("HASZIP2"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        
        
        If Mid(Trim(Sheet1.Range("HASZIP2").value), 1, 1) = "Y" Then
            
            Sheet1.Unprotect Password:=getmsgstate
            Application.EnableEvents = False
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.value = "XXXXXX"
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Protect Password:=getmsgstate
            
            Application.EnableEvents = True
        Else
            Application.EnableEvents = False
            Sheet1.Unprotect Password:=getmsgstate
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Protect Password:=getmsgstate
            Application.EnableEvents = True
        End If
        
    End If
    
    '-----------------------
    
'    If (Target.name.name = "sheet1.ResidentialStatus1") Then
     If Not Application.Intersect(Sheet1.Range("sheet1.ResidentialStatus1"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        LockUnlock80TTB
        
        LockUnlock80DDB
        
        LockUnlockQVIA
        
        'LockUnlockIBANSELECT
        
        If (Mid(Target.value, 1, 2) = "NR") Then
            
            Lock44AD
            
        Else
            
            Unlock44AD
            
        End If
        
    End If
    
'    If (Target.name.name) = "SELECT80DB" Then
      If Not Application.Intersect(Sheet1.Range("SELECT80DB"), Sheet1.Range(Target.Address)) Is Nothing Then
       
        sPassword = mdCommon.getmsgstate
        
        Sheet1.Unprotect Password:=sPassword
        
        If Mid(Range("SELECT80DB").value, 1, 1) = "1" Then
            
            '
            
            If Mid(Range("sheet1.Status").value, 1, 1) <> "I" And Mid(Range("sheet1.Status").value, 1, 1) <> "(" Then
                
                fmsgboxsmall "* Please Select a valid Option from the dropdown of Part B Sec.80D (Medical Expenditure) under Chapter VIA."
                Application.EnableEvents = False
                Range("SELECT80DB").value = "(Select)"
                
                Range("SELECT80DB").Select
                Application.EnableEvents = True
            End If
            
            If Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
                
                fmsgboxsmall "* Please Select a valid Option from the dropdown of Part B Sec.80D (Medical Expenditure) under Chapter VIA."
                Application.EnableEvents = False
                Range("SELECT80DB").value = "(Select)"
                
                Range("SELECT80DB").Select
                Application.EnableEvents = True
            End If
            
        End If
        
        Sheet1.Protect Password:=sPassword
        
    End If
    
'    If (Target.name.name) = "SELECT80D" Then
       If Not Application.Intersect(Sheet1.Range("SELECT80D"), Sheet1.Range(Target.Address)) Is Nothing Then
       
        sPassword = mdCommon.getmsgstate
        
        Sheet1.Unprotect Password:=sPassword
        
        If Mid(Range("SELECT80D").value, 1, 1) = "2" Or Mid(Range("SELECT80D").value, 1, 1) = "7" Then
            
            If Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
                
                fmsgboxsmall "* Please Select a valid Option from the dropdown of Part A Sec.80D under Chapter VIA."
                Application.EnableEvents = False
                Range("SELECT80D").value = "(Select)"
                
                Range("SELECT80D").Select
                Application.EnableEvents = True
            End If
            
        End If
        
        Sheet1.Protect Password:=sPassword
        
    End If
    
'    If (Target.name.name) = "IncD.TypeOfHP" Then
       If Not Application.Intersect(Sheet1.Range("IncD.TypeOfHP"), Sheet1.Range(Target.Address)) Is Nothing Then

    
    
        
        '------------------Un Protect--------------------
        
        sPassword = mdCommon.getmsgstate
        
        Sheet1.Unprotect Password:=sPassword
        
        'LockUnlock24b
        
        '------------------Un Protect Ends--------------------
 'Ayush_commented_27/01
'        If Mid(Range("IncD.TypeOfHP").value, 1, 1) = "S" Then
'
'            Dim hpInterest1, salIncome1 As Variant
'
'            hpInterest1 = Range("IncD.IncomeFromHP_v").value
'
'            If Range("IncD.TypeOfHP").value = "Self Occupied" Then
'
'                mdIncomeDetails.CheckHousePropertyIncome (hpInterest1)
'
'            End If
'
'            '
'            Application.EnableEvents = False
'            Range("IncD.IncomeFromHP_i").value = ""
'
'            Range("IncD.IncomeFromHP_i").MergeArea.Locked = True
'
'            Range("IncD.IncomeFromHP_i").MergeArea.Interior.Color = (&HD8D8D8)
'
'            Range("IncD.IncomeFromHP_ii").value = ""
'
'            Range("IncD.IncomeFromHP_ii").MergeArea.Locked = True
'
'            Range("IncD.IncomeFromHP_ii").MergeArea.Interior.Color = (&HD8D8D8)
'            Application.EnableEvents = True
'        Else
'            Application.EnableEvents = False
'            Range("IncD.IncomeFromHP_i").MergeArea.Locked = False
'
'            Range("IncD.IncomeFromHP_i").MergeArea.Interior.Color = (&HCCFFCC)
'
'            Range("IncD.IncomeFromHP_ii").MergeArea.Locked = False
'
'            Range("IncD.IncomeFromHP_ii").MergeArea.Interior.Color = (&HCCFFCC)
'            Application.EnableEvents = True
'        End If
        
        'Ayush_27/01
        'LockUnlock24b 'Updated by Bindu on 23rd April 25
        
'If Mid(Range("IncD.TypeOfHP").value, 1, 1) = "S" Or Mid(Range("IncD.TypeOfHP").value, 1, 1) = "(" And Sheet8.Range("bacvalue").value = 1 Then
'    'If Sheet8.Range("bacvalue").value = 2 Then
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet21.Visible = xlSheetHidden
'    Sheet21.Visible = xlSheetHidden
'    ThisWorkbook.Protect Password:=getmsgstate
''    ThisWorkbook.Unprotect Password:=getmsgstate
''    Sheet21.Visible = xlSheetVisible
''    Sheet21.Visible = xlSheetVisible
''    ThisWorkbook.Protect Password:=getmsgstate
'    Else
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet21.Visible = xlSheetVisible
'    Sheet21.Visible = xlSheetVisible
'    ThisWorkbook.Protect Password:=getmsgstate
''     ThisWorkbook.Unprotect Password:=getmsgstate
''    Sheet21.Visible = xlSheetHidden
''    Sheet21.Visible = xlSheetHidden
''    ThisWorkbook.Protect Password:=getmsgstate
''    ThisWorkbook.Unprotect Password:=getmsgstate
''    Sheet20.Visible = xlSheetVisible
''    Sheet20.Visible = xlSheetVisible
''    ThisWorkbook.Protect Password:=getmsgstate
'    End If
        
        
        
        
        '----------------Lock Password-------------------START---
        
        Sheet1.Protect Password:=sPassword
        
        '----------------Lock Password-------------------END-----
        
    End If
    
  '  If (Target.name.name) = "sheet1.ReturnFileSec" Then
         If Not Application.Intersect(Sheet1.Range("sheet1.ReturnFileSec"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        '------------------Un Protect--------------------
        
        sPassword = mdCommon.getmsgstate
        
        Sheet1.Unprotect Password:=sPassword
        
        '------------------Un Protect Ends--------------------
'Change-08.12.2022.103.07
        
        ID_Selection_ReturnFileSec
         
        If Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
            Dim vdate As Variant
            vdate = Sheet5.Range("Ver.Date").value
            
'            If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2022-08-01") Then
'                fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2022 Or extended due date"
'                Range("sheet1.ReturnFileSec").value = "(Select)"
'            End If
            
'            If Sheet8.Range("bacValue").value = 1 Then
'
'                'fmsgbox "* Please Select ""Have you opted  For New tax regime u/s 115BAC And filed Form 10IE in AY 2021-22 ? ""  As n since you have filed Return later." dpk1201
'                fmsgbox "* Please Select ""Have you opted  For New tax regime u/s 115BAC And filed Form 10IE in AY 2022-23 ? ""  As n since you have filed Return later."
'
'                Range("sheet1.ReturnFileSec").value = "(Select)"
'
'            End If
            
        End If
        
        If Range("sheet1.ReturnFileSec").value = "(Select)" And Range("sheet1.ReturnFileSec2").value = "(Select)" Then
            
            Range("sheet1.ReturnFileSec1").value = "(Select)"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
            
            Range("sheet1.ReturnFileSec1").value = "11 - 139(1)-On or before due date"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
            
            Range("sheet1.ReturnFileSec1").value = "12 - 139(4)-After due date"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "142(1)" Then
            
            Range("sheet1.ReturnFileSec1").value = "13 - 142(1)"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "148" Then
            
            Range("sheet1.ReturnFileSec1").value = "14 - 148"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "153A" Then
            
            Range("sheet1.ReturnFileSec1").value = "15 - 153A"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "153C" Then
            
            Range("sheet1.ReturnFileSec1").value = "16 - 153C"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "139(5)-Revised Return" Then
            
            Range("sheet1.ReturnFileSec1").value = "17 - 139(5)-Revised Return"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "139(9)" Then
            
            Range("sheet1.ReturnFileSec1").value = "18 - 139(9)"
            
         'J200924
'      ElseIf Range("sheet1.ReturnFileSec").value = "119(2)(b)- After condonation of delay" Then
'
'            Range("sheet1.ReturnFileSec1").value = "20 - 119(2)(b)- After condonation of delay"
            
'             ElseIf Range("sheet1.ReturnFileSec").value = "139(9A) - After condonation of delay u/s 119(2)(b)" Then
'
'            Range("sheet1.ReturnFileSec1").value = "20 - 139(9A) - After condonation of delay u/s 119(2)(b)"
  ElseIf Range("sheet1.ReturnFileSec").value = "119(2)(b)- After condonation of delay" Then
         
          Range("sheet1.ReturnFileSec1").value = "20 - 119(2)(b)- After condonation of delay"
            
        ElseIf Range("sheet1.ReturnFileSec").value = "139(8A)" Then
            
            Range("sheet1.ReturnFileSec1").value = "21 - 139(8A)"
        End If
        
        If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = "20" Then
            'Newly changed by Bindu
            'fmsgboxsmall "* 119(2)(b)- after condonation of delay should be filed after expiry of due Date mentioned in 139(1)/139(4) in Sheet Income Details"        ', vbInformation, "ITR4"
           'J200924
            ' fmsgboxsmall "* 119(2)(b)- after condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)"        ', vbInformation, "ITR4"
  
  'fmsgboxsmall "*119(2)(b)- After condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)"        ', vbInformation, "ITR4"
  fmsgboxsmall "*""119(2)(b)- After condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)."""        ', vbInformation, "ITR4"

        End If
        
        If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = "17" Then
            If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) > "2027-03-31" Then
                fmsgboxsmall "*""Selected Filing section 139(5) cannot be selected post 31st March, 2027"""
                Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
            End If
        End If
        
        If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = "11" Then
            If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) > "2027-03-31" Then
                fmsgboxsmall "*""Selected Filing section 139(1) cannot be selected post 31st March, 2027"""
               ' Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
            End If
        End If
        
        
     
        
        
        LockUnlockReturnfilesec
        
        '----------------Lock Password-------------------START---
        
        Sheet1.Protect Password:=sPassword
        
        '----------------Lock Password-------------------END-----
        
    End If
    
'End If 'Ayush_03/04/2026


'Ayush 31/03/2026
If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceNo"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.ResidenceName"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.RoadOrStreet"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.LocalityOrArea"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.CityOrTownOrDistrict"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("HASZIP"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.CountryCode"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.ZipCode"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then


Call mdIncomeDetails.AutoPopulateSecondaryAddress
End If

endline:

End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo endline
 Application.ScreenUpdating = False

 Application.EnableEvents = False

 CalculateTaxCheckFlag = False
 
 

 Sheet1.Unprotect Password:=getmsgstate
 

 
 
 
 Rows("28").RowHeight = 80
 
 'Ayush_09/05
 
 If Sheet8.Range("bacvalue").value = 2 Then
        If Mid(Range("sheet1.NatureofEmployment").value, 1, 3) <> "Not" Then
            ThisWorkbook.Unprotect Password:=getmsgstate
            Sheet15.Visible = xlSheetVisible
            ThisWorkbook.Protect Password:=getmsgstate
        Else
            ThisWorkbook.Unprotect Password:=getmsgstate
            Sheet15.Visible = xlSheetHidden
            ThisWorkbook.Protect Password:=getmsgstate
        End If
End If
 
 
 
' If Not Intersect(Target, Range("IncD.Section80CCC")) Is Nothing Then
'        UpdateButtons_Sec80CCC
'        LockUnlock_80CCC
'    End If

 
 
'Bindu
'Ayush_enrol========================
'If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" And _
'       Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "F" Then
'
''Added by Ayush_06/09/2024
'   ' If Len(Sheet1.Range("sheet1.Aadhaar").value) >= 12 Then
'
'   If Len(Sheet1.Range("sheet1.Aadhaar").value) >= 12 Or Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") >= Dformat(Trim("01/10/2024"), "yyyy-mm-dd") Then
'
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = True
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HD8D8D8
'        Sheet1.Protect Password:=getmsgstate
'
'    Else
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'        Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HCCFFCC
'        Sheet1.Protect Password:=getmsgstate
'
'
'
'    End If
'
'End If
'===================

sPassword = getmsgstate




Dim DOB_6 As Variant

Dim validateType As String

Dim result As Variant

'Ayush_13/01/2026_V0.2_commented
'If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("OSIncomeNotified89A_AmountUS")) Is Nothing Or Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("OSIncomeNotified89A_AmountUK")) Is Nothing Or Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("OSIncomeNotified89A_AmountCan")) Is Nothing Or Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("OSIncreliefus89A")) Is Nothing Then
'
'
'
'    If (Sheet1.Range("OSIncomeNotified89A").value <> "") Or (Sheet1.Range("OSIncreliefus89A").value <> "" And Sheet1.Range("OSIncreliefus89A").value <> 0) Then
'        If Sheet1.Range("OSIncreliefus89A").value > Sheet1.Range("OSIncomeNotified89A").value Then
'          Sheet1.Range("IncD_q1OS1").value = Sheet1.Range("OSIncomeNotified89A")
'        Else
'            Sheet1.Range("IncD_q1OS1").value = (Sheet1.Range("OSIncomeNotified89A").value - Sheet1.Range("OSIncreliefus89A").value)
'        End If
'    End If
'
'End If
'------------------------------

'Change.02.03.2023.102.IDS.75
If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.acknowledgementNumber_Opt")) Is Nothing Or Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.acknowledgementNumber_OptOut")) Is Nothing Or Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.acknowledgementNumber")) Is Nothing Then
    If Not IsNumeric(Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value) Then
        MsgBox ("Characters are not allowed only Numbers!!")
    End If
    
    If Not IsNumeric(Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value) Then
        MsgBox ("Characters are not allowed only Numbers!!")
    End If
    If Not IsNumeric(Sheet1.Range("Sheet1.acknowledgementNumber").value) Then
        MsgBox ("Characters are not allowed only Numbers!!")
    End If
End If
'End Change IDS.75

'ID_Nature_Employment_New

'PAG_C24 AY 2024-25 Change New Starts----------------------
If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Dateoffiling")) Is Nothing Then
    If Trim(Range("Sheet1.Dateoffiling").value) <> "00/00/0000" And Range("Sheet1.Dateoffiling").value <> "" Then
            If Not FormatNCheckDate(Trim(Range("Sheet1.Dateoffiling").value)) Then
                fmsgboxStatus "*Please enter date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
                Sheet1.Range("Sheet1.Dateoffiling").Select
                Sheet1.Range("Sheet1.Dateoffiling").value = ""
                
            End If
    End If
    
   
End If



'Ayush_25-26
If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling"), Sheet1.Range(Target.Address)) Is Nothing Then

    Dim Datefiling_BAC115 As Variant
    Datefiling_BAC115 = Dformat(Trim(Target.value), "")
  'Ayush_09/06
'    If Dformat(Trim(Target.value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'  '  If Not ChkMaxDate_filing_2425(Trim(Target.value), "31-07-2024") Then
'       ' fmsgbox "*""Please select ‘No’, if Form 10IEA was filed after due date for AY 2024-25."""
''       fmsgbox "*""Date of filing of Form 10IEA for AY 2024-25 can not be more than system Date in Income Details"""
'        fmsgbox "* Date of filing Form 10IEA shall not be beyond system date."
'        Range("BAC115.DateOfFiling").value = ""      'added by sai on 27/01/2025 for AY 2025-26
'    End If


'Ayush_25/07
'If Dformat(Trim(Target.value), "yyyy-mm-dd") > Dformat(Trim("31/12/2025"), "yyyy-mm-dd") Then
'       fmsgbox "* ""Since form 10IEA is filed after due date please select appropriate option"""
'        Range("BAC115.DateOfFiling").value = ""

If Dformat(Trim(Target.value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then

      fmsgbox "* Date of filing Form 10IEA shall not be beyond system date."

        Range("BAC115.DateOfFiling").value = ""
        
End If

End If



'PAG_C24 AY 2024-25 Change End
'--------------------------------------------------------

'Ayush_25-26
'If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Dateoffiling")) Is Nothing Then
'    If Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Then
'            If Sheet1.Range("Sheet1.Dateoffiling") > "31/07/2025" Then
'                fmsgboxStatus "*Since form 10IEA is filed after due date option for current assessment year cannot be Old tax regime"
'                Sheet1.Range("Sheet1.Dateoffiling").Select
'                Sheet1.Range("Sheet1.Dateoffiling").value = ""
'
'            End If
'    End If
'
'
'End If

'Commented by Bindu on 25th Jan 2025 as per DE V2 these fields releted to some other logic as well
'If Not Application.Intersect(Sheet1.Range("Sheet1.Dateoffiling"), Sheet1.Range(target.Address)) Is Nothing Then
'
'
'
'    Dim Datefiling_BAC As Variant
'    Datefiling_BAC = Dformat(Trim(target.value), "")
'
'
'    If Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Then
'        If Not ChkMaxDate_filing(Trim(target.value), "31-07-2025") Then
'        fmsgbox "*""Since form 10IEA is filed after due date option for current assessment year cannot be Old tax regime"""
'        target.ClearContents
'    End If
'    End If
'    '=================
'
'
'End If

'PAG_C61 2024-25 Change Starts-------------
If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("IncD.AnyOtherDeductions")) Is Nothing Then
    If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government") And (Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) = "I") Then
       If Target.value > (0.462 * (Sheet1.Range("IncD.IncomeFromSal_ii").value)) Then
           fmsgboxStatus "*80CCH can not be greater than Prescribed limit as per section or 2,88,000 which ever is lower"

'Chandru
            'Target.value = 0
       End If
    End If
End If
'PAG_C61 2024-25 Change End---------------



'Change.02.03.2023.102.IDS.78
'If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Amount")) Is Nothing Then
'    If Target.value > 500000 And Range(Target.Address).Offset(0, -32).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette " Then
'        MsgBox (" Amount Under Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette Can not be more than 5 Lakhs")
'        Target.MergeArea.ClearContents
'        Target.MergeArea.Select
'
'    End If
'End If
'
''Ayush_Legacy10
'If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Amount")) Is Nothing Then
'    If Target.value > 500000 And Range(Target.Address).Offset(0, -32).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'        MsgBox (" Amount Under Sec 10(10B)) Second proviso - Compensation under scheme approved by the Central Government can not be more than 5 Lakhs")
'        Target.MergeArea.ClearContents
'        Target.MergeArea.Select
'
'    End If
'End If
'End Change IDS.78

'Ayush_13/01/2026_V0.2
If Not Application.Intersect(Sheet1.Range("Secondary_Address"), Sheet1.Range(Target.Address)) Is Nothing Then
    AutoPopulateSecondaryAddress
End If
'----------------------


validateType = ""



vbMessageCaption = "ITR-4"

DOB_6 = Sheet1.Range("sheet1.DOB").value

'Target.value = UCase(Target.value)

PartAGenChangeSub1 Target

PartAGenChangeSub2 Target

'Change-01.12.2022.103.12

'CurrentYear10IE_23_24 Target AY 2023-24 Old Change

'Change-01.12.2022.103.16B
'CurrentYear10IEDateAckAY_23_24 Target

'CurrentYear10IEDateAckAY_24_25 target 'PAG_C24 AY 2024-25 Change Starts
'Updated by Bindu on 25th Jan 25 as per DE V2

If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then
    Taxregime_25_26_Status
End If

'A23
If Not Application.Intersect(Sheet1.Range("sheet1.115BAC"), Sheet1.Range(Target.Address)) Is Nothing Then
    Taxregime_25_26
End If

'A23(a)
If Not Application.Intersect(Sheet1.Range("sheet1.OptOutNewTaxRegime_New_Yes"), Sheet1.Range(Target.Address)) Is Nothing Then
    Taxregime_25_26_A
End If

'A23(b)
If Not Application.Intersect(Sheet1.Range("sheet1.OptOutNewTaxRegime_New_No"), Sheet1.Range(Target.Address)) Is Nothing Then
    Taxregime_25_26_B
End If


'A23(c)
'If Not Application.Intersect(Sheet1.Range("sheet1.OptOutNewTaxRegime_New_NA"), Sheet1.Range(Target.Address)) Is Nothing Then
If Not Application.Intersect(Sheet1.Range("BAC115.A23b"), Sheet1.Range(Target.Address)) Is Nothing Then
    Taxregime_25_26_C
End If
'Ayush_regime
If Not Application.Intersect(Sheet1.Range("BAC115.A23ii_No"), Sheet1.Range(Target.Address)) Is Nothing Then
   Taxregime_26_27_A23ii_b
End If
'---Need to call

If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling"), Sheet1.Range(Target.Address)) Is Nothing Then
  msgError = ""
  If Not ID_Opt_Date10IEA_New Then fmsgbox msgError       'For Table A23(a) by Ayush
End If
'added by sai on 27/01/2025 for AY 2025-26==========================


If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling_New_Yes"), Sheet1.Range(Target.Address)) Is Nothing Then
    msgError = ""
    If Not ID_Opt_Date10IEA_New_Ai Then fmsgbox msgError  'For Table A23(a)(i) 'modified by sai on 11/02/2025
End If

If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling_New_No"), Sheet1.Range(Target.Address)) Is Nothing Then
'  ID_Opt_Date10IEA_New_Bi   'For Table A23(b)(i)
   msgError = ""
   If Not ID_Opt_Date10IEA_New_Bi Then fmsgbox msgError  'For Table A23(b)(i)  'modified by sai on 11/02/2025
End If

If Not Application.Intersect(Sheet1.Range("BAC115.DateOfFiling_New_NA"), Sheet1.Range(Target.Address)) Is Nothing Then
'   ID_Opt_Date10IEA_New_Ci   'For Table A23(c)(i)
   msgError = ""
   If Not ID_Opt_Date10IEA_New_Ci Then fmsgbox msgError  'For Table A23(c)(i)   'modified by sai on 11/02/2025
End If
'===================================================================

'ID_Nature_Employment_New  'Ayush_13/02/2026


'BAC115_24_25 target 'Ayush_25-26 'Commented by Bindu on 25th Jan 25 Same updated in this function Taxregime_25_26_Status
'BAC115filing_24_25 target 'Ayush_25-26'Commented by Bindu on 25th Jan 25 As per DE V2 these fields releted to some other logic as well
'ID_Opt_Date10IEA_New target 'Ayush_25-26 'Commented by Bindu on 25th Jan 25 As per DE V2 these fields releted to some other logic as well
'BAC115filing_25_26 target 'Ayush_25-26'Commented by Bindu on 25th Jan 25 As per DE V2 these fields releted to some other logic as well

'Change-06.12.2022.103.10
'ID_Opt_Date10IE Target AY 2023-24 OLD

' ID_Opt_Date10IE_New target 'PAG_C25 2024-25 Change Starts ''Commented by Bindu on 25th Jan 25 As per DE V2 these fields releted to some other logic as well

'Change-06.12.2022.103.14
'ID_OptOut_Date10IE Target 'PAG 2023-24 OLD

'Change-06.12.2022.103.18
'ID_Date10IE Target PAG_E24 2023-24 old

'Newly moved from Above her by Bindu
'Ayush_25 -26
'
'If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Dateoffiling")) Is Nothing Then
'
'If Target.value <> "" Then
'    If Dformat(Trim(Range("Sheet1.Dateoffiling").value), "yyyy-mm-dd") <= Dformat(Trim("31/07/2024"), "yyyy-mm-dd") Then
'        If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date" Then
'            MsgBox ("If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly")
'
'           Sheet1.Range("Sheet1.Dateoffiling").Select
'            Sheet1.Range("Sheet1.Dateoffiling").value = ""
'            GoTo Endline
'        End If
'    Else
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date" Then
'                MsgBox ("If option is selected as " & Chr(34) & "Yes, within due date" & Chr(34) & ", then date of filing of Form should be within due date u/s 139(1). Please enter date of filing correctly")
'                Target.MergeArea.ClearContents
'                Target.MergeArea.Select
'                GoTo Endline
'            End If
'      End If
'  End If
'End If


If Not Application.Intersect(Sheet1.Range("sheet1.PANRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "sheet1.PANRepAssessee" Then
    
    Dim TargetPAN1  As String
    
    TargetPAN1 = ""
    
    TargetPAN1 = Target.name.name
    
    'Target.value = UCase(Target.value)
    
    If Range(TargetPAN1).value <> "" Then
        
        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
        If Not mdIncomeDetails.CheckPAN(UCase(Trim(Range(TargetPAN1).value))) Then
            'Ayush_PAN
            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
            fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
            
            Sheet1.Range(TargetPAN1).Select
            
            Sheet1.Range(TargetPAN1).value = ""
            
            GoTo endline
            
        End If
        
        Sheet1.Range(TargetPAN1).value = UCase(Range(TargetPAN1).value)
        
        If Range(TargetPAN1).value = Sheet1.Range("sheet1.PAN").value Then
            
            fmsgboxsmall "* PAN of the representative assessee And assessee (under personal information) cannot be same in Sheet Income Details"        ', vbOKCancel, mdCommon.vbMessageCaption
            
            Sheet1.Range(TargetPAN1).Select
            
            Sheet1.Range(TargetPAN1).value = ""
            
            GoTo endline
            
        End If
        
    End If
    
End If

If Not checkfieldSuperSpecialcharacter(Target.value) Then
    
    fmsgboxsmall ("* Please Do Not enter Special characters < , > in Sheet Income Details" & Chr(13))
    
    Target.value = ""
    
End If

If Not Application.Intersect(Sheet1.Range("sheet1.AadhaarRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "sheet1.AadhaarRepAssessee" Then
    
    'If (Target.name.name) = "sheet1.Aadhaar" Then
    
    If Len(Sheet1.Range("sheet1.AadhaarRepAssessee").value) >= 12 Then
        
        If chkNumeric(Sheet1.Range("sheet1.AadhaarRepAssessee").value) Then
            
            If Sheet1.Range("sheet1.AadhaarRepAssessee").value <> "000000000000" Then
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").value = Sheet1.Range("sheet1.AadhaarRepAssessee").value
                
            Else
                
                fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                
            End If
            
            If Sheet1.Range("sheet1.AadhaarRepAssessee").value < 0 Then
                
                fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                
            End If
            
            If Sheet1.Range("sheet1.AadhaarRepAssessee").value <> "111111111111" Then
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").value = Sheet1.Range("sheet1.AadhaarRepAssessee").value
                
            Else
                
                fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                
                Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                
            End If
            
        Else
            
            fmsgboxStatus "* Invalid Aadhaar Number of the representative"
            
            Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
            
            Sheet1.Range("sheet1.AadhaarRepAssessee").Select
            
        End If
        
    Else
        
        Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
        
    End If
    
    '            End If
    
    If Sheet1.Range("sheet1.AadhaarRepAssessee").value <> "" Then
        
        If Sheet1.Range("sheet1.AadhaarRepAssessee").value = Sheet1.Range("sheet1.Aadhaar").value Then
            
            fmsgboxsmall "* Aadhaar no. entered in Part A -Gen And Aadhaar no. of the representative cannot be same in Sheet Income Details"        ', vbOKCancel, mdCommon.vbMessageCaption
            
            Sheet1.Range("sheet1.AadhaarRepAssessee").Select
            
            Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
            
            GoTo endline
            
        End If
        
    End If
    
End If


'PAG_C7 AY 2024-25 Change

        If Not Application.Intersect(Sheet1.Range("sheet1.NatureofEmployment"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.NatureofEmployment").value = "Not Applicable (eg. Family pension etc)" Then
               MsgBox ("Since Nature of employment selected as ""Not Applicable"", Schedule ""Salary"" is greyed off. If you have Salary income, kindly select correct dropdown in Nature of employment")
            End If
        End If
        





'Ayush_13/01/2026_V0.2
If Not Application.Intersect(Sheet1.Range("sheet1.PinCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
 'If (Target.name.name) = "sheet1.PinCode2" Then
        
        If Len(Range("sheet1.PinCode2").value) = 6 Then
            
            If Range("sheet1.PinCode2").value = 0 Then
                
                fmsgboxStatus ("* Pincode must be exactly 6 digits And should Not start With zero in Sheet Income Details")
                
                Range("sheet1.PinCode2").MergeArea.ClearContents
                
                Range("sheet1.PinCode2").Select
                
            End If
            
            If Len(Range("sheet1.PinCode2").value) = 6 Then
                
                If Not ((IsNumeric(Range("sheet1.PinCode2").value)) And checkspecialcharacterForPinCode(Range("sheet1.PinCode2").value)) Then
                    
                    fmsgboxStatus ("* Invalid Pincode :Pincode cannot have Alphabets/Special characters And should be 6 digits Numeric Value in Sheet Income Details")
                    
                    Application.EnableEvents = False
                    Range("sheet1.PinCode2").MergeArea.ClearContents
                    
                    Range("sheet1.PinCode2").Select
                    
                End If
                
            End If
            
            If Mid(Range("sheet1.PinCode2").value, 1, 1) = 0 Then
                
                fmsgboxStatus ("* Pincode cannot begin With zero")
                Application.EnableEvents = False
                Range("sheet1.PinCode2").MergeArea.ClearContents
                
                Range("sheet1.PinCode2").Select
               
            End If
            
        End If
        
        If Range("sheet1.CountryCode2") = "" Or Range("sheet1.CountryCode2") = "(Select)" Then
            Application.EnableEvents = False
            Range("sheet1.CountryCode2") = "91-INDIA"
            
        End If
        
        Dim PinCode2 As Range
        
     
        For Each PinCode2 In Sheet8.Range("All_Pincode_list")
            
            Dim code2 As String
            
            code2 = Left(CStr(Sheet1.Range("sheet1.PinCode2").value), 6)
            
            If CStr(PinCode2.value) = CStr(code2) Then
                
                Application.EnableEvents = False
                Sheet1.Range("sheet1.StateCode2").value = Sheet8.Range("EQ" & PinCode2.row).value
                
                Exit For
                
            End If
            
        Next PinCode2
        
    End If
    
    Dim pincode3 As Range
        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.PinCode2").value <> "" Then
                'if state and pincode doesn't match, empty the pincode.
                Dim state2 As Variant
                state2 = Application.VLookup(Sheet1.Range("sheet1.PinCode2").value, Sheet8.Range("All_PinCode_V"), 2, False)
                If Not (state2 = UCase(Mid(Sheet1.Range("sheet1.StateCode2").value, 4)) Or StateMatchesPin(Target.value)) Then
                    'Ayush_01/01/2025
                     'Newly added SIT-85381---
                        Application.EnableEvents = False
                        Sheet1.Unprotect Password:=getmsgstate
                        '---ends
                        Sheet1.Range("sheet1.PinCode2").value = ""
                       
                End If
            End If
        End If
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
         Dim DupFlag2 As Boolean
         DupFlag2 = False
                 'If the pincode is not in the duplicate list, populate, else don't populate---
                 For Each pincode3 In Sheet8.Range("Duplicate_Pincode_list")
                    If Sheet1.Range("sheet1.PinCode2").value = pincode3 Then
                        DupFlag2 = True
                        Exit For
                    End If
                 Next pincode3
                 
                
                If DupFlag2 = False Then
                    Dim StateName2
                    Dim StateCode2
                    
                    StateName2 = Application.VLookup(Sheet1.Range("sheet1.PinCode2").value, Sheet8.Range("All_PinCode_V"), 2, False)
                    If IsError(StateName2) Then
                        'MsgBox ("Invalid Pincode")
                        'Sheet1.Range("sheet1.StateCode1").value = "(Select)"
                       Application.EnableEvents = False
                        Sheet1.Range("sheet1.PinCode2").value = ""
                     Application.EnableEvents = True
                        
                    Else
                        StateCode2 = Application.VLookup(StateName2, Sheet8.Range("EW:EX"), 2, False)
                        Sheet1.Range("sheet1.StateCode2").value = StateCode2
                        Sheet1.Range("sheet1.CountryCode2").value = "91-INDIA"
                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
                    End If
                 End If
                 
                 If DupFlag2 = True Then
                 Application.EnableEvents = False
                 
                    If Sheet1.Range("sheet1.StateCode2").value <> "" Then
                        If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode2").value) Then
                            Sheet1.Range("sheet1.StateCode2").value = "(Select)"
                         
                            
                        End If
                    End If
                 End If

         End If

'
'-------------------


'Call mdIncomeDetails.AutoPopulateSecondaryAddress

'HP income Validation on Data Entry

PartAGenChangeSub3 Target


PartAGenChangeSub4 Target

'Ayush_13/01/2026
'If (Target.name.name) = "sheet1.DOB" Then
        If Not Application.Intersect(Sheet1.Range("sheet1.DOB"), Sheet1.Range(Target.Address)) Is Nothing Then

    
    'Ayush_Legacy45
 '   If Range("sheet1.status").value = "I - INDIVIDUAL" Then
        If Range("sheet1.status").value = "I - INDIVIDUAL" Or Range("sheet1.status").value = "F - FIRM(Other than LLP)" Then
            
            If Trim(Range("sheet1.DOB").value) = "01/01/0001" Then
'            "Ayush_10/02
'                fmsgboxStatus "* Date of Birth in Sheet Income Details Please enter a valid date.in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
                fmsgboxStatus """Date of Birth/Formation - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
                     
                
                Sheet1.Range("sheet1.DOB").Select
                
                Sheet1.Range("sheet1.DOB").value = ""
                
            End If
            
        End If
        
        'Change-06.04.2023.103.PAG.01
            AnyOtherDeductions80CCH Target
        'End Change
        LockUnlock80TTB
        
        Sheet1.Unprotect Password:=getmsgstate
        
        '
        
        Sheet1.Protect Password:=getmsgstate
        
    End If
    
    
    
    'DOB Validation: On Data Entry
    
 '   If (Target.name.name) = "sheet1.DOB" Then
        If Not Application.Intersect(Sheet1.Range("sheet1.DOB"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        Dim HUFReq  As Long
        
        If Range("sheet1.status").value = "H - HUF" Then HUFReq = 1 Else HUFReq = 0
        
        If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
            
            result = CheckDate(Trim(Range("sheet1.DOB").value), HUFReq)
            
            If result = 1 Then
  'Ayush_10/02
'                fmsgboxStatus "* Date of Birth in Sheet Income Details Please enter a valid Date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
                 fmsgboxStatus """Date of Birth/Formation - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
                
                Sheet1.Range("sheet1.DOB").Select
                
                Sheet1.Range("sheet1.DOB").value = ""
                
            ElseIf result = 2 Then

'Change-22.11.2023.103.02
'Ayush_26-27_C9
Dim cutoff As Date
cutoff = CDate(Sheet8.Range("DOB_1").value)
                
                'fmsgboxStatus "* Date of Birth in Sheet Income Details  should Not be after 01/04/2022"        ', vbOKOnly, mdCommon.vbMessageCaption
                
'                fmsgboxStatus "* Date of Birth in Sheet Income Details should not be after 01/04/2023"

            'PAG_C1 AY 2024-25 Changes
            'Ayush_25-26
            'AYush_Legacy44
            
            
'                fmsgboxStatus "* Date of Birth in Sheet Income Details should not be after 01/04/2024"
               ' fmsgboxStatus "* Date of Birth in Sheet Income Details should not be after 01/04/2025"
                 fmsgboxStatus "* Date of birth in sheet income details should not be on or after " & Dformat1(cutoff, "dd/mm/yyyy")

                
                Sheet1.Range("sheet1.DOB").Select
                
                Sheet1.Range("sheet1.DOB").value = ""
                
            End If
            
        End If
        
        GoTo endline
        
    End If
    
    'Ayush_13/01/2026
 '   If (Target.name.name) = "sheet1.MobileRepAssessee" Then
       If Not Application.Intersect(Sheet1.Range("sheet1.MobileRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If (Range("sheet1.MobileRepAssessee")) <> "" Then
            
            If Sheet1.Range("sheet1.CountryCodeRepAssessee").value = "91" Then
                
                If Len(Range("sheet1.MobileRepAssessee")) <> 10 Then
                    
                    fmsgboxStatus "* Mobile No. in contact No. of representative assessee should be 10 digits in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    
                    
                    
                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
                    
                End If
                
            Else
                
                If Len(Range("sheet1.MobileRepAssessee")) < 5 Or Len(Range("sheet1.MobileRepAssessee")) > 10 Then
                    
                    fmsgboxStatus "* Mobile No. in contact No. of representative assessee is invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    
                    'Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
                    
                End If
                
            End If
            
            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
                
                If Not IsNumeric(Trim(Range("sheet1.MobileRepAssessee").value)) Then
                    
                    fmsgboxStatus "* Mobile No. in contact No. of representative assessee is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    
                   ' Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
                    
                End If
                
            End If
            
            
            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
              If Not (ValidateMobileNumber_Rep()) Then
                 Sheet1.Range("sheet1.MobileRepAssessee").value = ""
                End If
            End If
            
             
            
            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
                
                If Not CheckSpecialCharacter(Trim(Range("sheet1.MobileRepAssessee").value)) Then
                    
                    fmsgboxStatus "* Mobile No. in contact No. of representative assessee is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    
                   ' Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
                    
                End If
                
            End If
            
            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
                
                If Mid(Trim(Range("sheet1.MobileRepAssessee").value), 1, 1) = 0 Then
                    
                    fmsgboxStatus "* Mobile No. in contact No. of representative assessee should not begin with 0"        ', vbOKOnly, vbMessageCaption
                    
                    'Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
                    
                End If
                
            End If
            
            If Len(Range("sheet1.MobileRepAssessee")) + Len(Range("sheet1.CountryCodeRepAssessee")) > 15 Then
                
                fmsgboxStatus "* Mobile No. And Country Code in contact No. of representative assessee combination should not be more than 15 digits in Sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                'Range("sheet1.Mobileno").Select
                
                Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
                
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.MobileRepAssessee").value = ""
                
            End If
            
        End If
        
    End If
    '---------------
'-----------



If Target.Validation.Type = 6 Then
    
'    If (Target.name.name) = "sheet1.Aadhaar" Then
      If Not Application.Intersect(Sheet1.Range("sheet1.Aadhaar"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Len(Sheet1.Range("sheet1.Aadhaar").value) >= 12 Then
            
            If chkNumeric(Sheet1.Range("sheet1.Aadhaar").value) Then
                
                If Sheet1.Range("sheet1.Aadhaar").value <> "000000000000" Then
           
                    'Sheet1.Range("sheet1.Aadhaar").value = Sheet1.Range("sheet1.Aadhaar").value
                    
                Else
                    
                    fmsgboxStatus "* Invalid Aadhaar Number"
                    
                    Sheet1.Range("sheet1.Aadhaar").value = ""
                    
                    Sheet1.Range("sheet1.Aadhaar").Select
                    'Newly added by Bindu
                    'Ayush_enrol
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HCCFFCC
                    
                End If
                
                If Sheet1.Range("sheet1.Aadhaar").value < 0 Then
                    
                    fmsgboxStatus "* Invalid Aadhaar Number"
                    
                    Sheet1.Range("sheet1.Aadhaar").value = ""
                    
                    Sheet1.Range("sheet1.Aadhaar").Select
                    'Newly added
                    'Ayush_enrol
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HCCFFCC
                    
                End If
                
                If Sheet1.Range("sheet1.Aadhaar").value <> "111111111111" Then
                    
                   ' Sheet1.Range("sheet1.Aadhaar").value = Sheet1.Range("sheet1.Aadhaar").value
                    
                Else
                    
                    fmsgboxStatus "* Please enter the Aadhaar number in correct format"
                    
                    Sheet1.Range("sheet1.Aadhaar").value = ""
                    
                    Sheet1.Range("sheet1.Aadhaar").Select
                    'Newly added
                    'Ayush_enrol
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HCCFFCC
                    
                End If
                
            Else
                
                fmsgboxStatus "* Please enter the Aadhaar number in correct format"
                
                Sheet1.Range("sheet1.Aadhaar").MergeArea.ClearContents
                
                Sheet1.Range("sheet1.Aadhaar").Select
                'Newly added
                'Ayush_enrol
'                Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'                Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HCCFFCC
                
            End If
            
        Else
            'Ayush_enrol
'            Sheet1.Unprotect Password:=getmsgstate
'            'Sheet1.Range("sheet1.Aadhaar").MergeArea.ClearContents
'            Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'            Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HCCFFCC
'            Sheet1.Unprotect Password:=getmsgstate
        End If
        'Newly added
        'Ayush_enrol
'        If Len(Sheet1.Range("sheet1.Aadhaar").value) >= 12 Then
'                Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'                Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = True
'                Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Interior.Color = &HD8D8D8
'                Sheet1.Protect Password:=getmsgstate
'        End If
    End If
    
    'new adhaar enrol
 'Ayush_enrol============================
'    If (Target.name.name) = "sheet1.AadhaarEnrol" Then
'
'        If Len(Sheet1.Range("sheet1.AadhaarEnrol").value) >= 12 Then
'
'            If chkNumeric(Sheet1.Range("sheet1.AadhaarEnrol").value) Then
'
'                If Sheet1.Range("sheet1.AadhaarEnrol").value <> "0000000000000000000000000000" Then
'
'                    'Sheet1.Range("sheet1.AadhaarEnrol").value = Sheet1.Range("sheet1.AadhaarEnrol").value
'
'                Else
'
'                    fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'
'                    Sheet1.Range("sheet1.AadhaarEnrol").Select
'
'                End If
'
'                If Sheet1.Range("sheet1.AadhaarEnrol").value < 0 Then
'
'                    fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'
'                    Sheet1.Range("sheet1.AadhaarEnrol").Select
'
'                End If
'
'                If Sheet1.Range("sheet1.AadhaarEnrol").value <> "1111111111111111111111111111" Then
'
'                    'Sheet1.Range("sheet1.AadhaarEnrol").value = Sheet1.Range("sheet1.AadhaarEnrol").value
'
'                Else
'
'                    fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'
'                    Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'
'                    Sheet1.Range("sheet1.AadhaarEnrol").Select
'
'                End If
'
'            Else
'
'                fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'
'                Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'
'                Sheet1.Range("sheet1.AadhaarEnrol").Select
'
'            End If
'
'        Else
'
'            Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.ClearContents
'
'        End If
'
'    End If
'==========================
    
  '  If (Target.name.name) = "sheet1.PinCode" Then
      If Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Len(Range("sheet1.PinCode").value) = 6 Then
            
            If Range("sheet1.PinCode").value = 0 Then
                
                fmsgboxStatus ("* Pincode must be exactly 6 digits And should Not start With zero in Sheet Income Details")
                
                Range("sheet1.PinCode").MergeArea.ClearContents
                
                Range("sheet1.PinCode").Select
                
            End If
            
            If Len(Range("sheet1.PinCode").value) = 6 Then
                
                If Not ((IsNumeric(Range("sheet1.PinCode").value)) And checkspecialcharacterForPinCode(Range("sheet1.PinCode").value)) Then
                    
                    fmsgboxStatus ("* Invalid Pincode :Pincode cannot have Alphabets/Special characters And should be 6 digits Numeric Value in Sheet Income Details")
                    
                    Application.EnableEvents = False
                    Range("sheet1.PinCode").MergeArea.ClearContents
                    
                    Range("sheet1.PinCode").Select
                    
                End If
                
            End If
            
            If Mid(Range("sheet1.PinCode").value, 1, 1) = 0 Then
                
                fmsgboxStatus ("* Pincode cannot begin With zero")
                Application.EnableEvents = False
                Range("sheet1.PinCode").MergeArea.ClearContents
                
                Range("sheet1.PinCode").Select
               
            End If
            
        End If
        
        If Range("sheet1.CountryCode") = "" Or Range("sheet1.CountryCode") = "(Select)" Then
            Application.EnableEvents = False
            Range("sheet1.CountryCode") = "91-INDIA"
            
        End If
        
        Dim PinCode As Range
        
     'sudha change
        For Each PinCode In Sheet8.Range("All_Pincode_list")
            
            Dim code As String
            
            code = Left(CStr(Sheet1.Range("sheet1.PinCode").value), 6)
            
            If CStr(PinCode.value) = CStr(code) Then
                
                Application.EnableEvents = False
                Sheet1.Range("sheet1.StateCode1").value = Sheet8.Range("EQ" & PinCode.row).value
                
                Exit For
                
            End If
            
        Next PinCode
        
    End If
    
   ' If (Target.name.name) = "sheet1.Mobileno" Then
    If Not Application.Intersect(Sheet1.Range("sheet1.Mobileno"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If (Range("sheet1.Mobileno")) <> "" Then
            
            If Sheet1.Range("sheet1.mobileCountryCode").value = "91" Then
                
                If Len(Range("sheet1.Mobileno")) <> 10 Then
                   'Ayush_28/01/2026
'                    fmsgboxStatus "* Mobile number 1 should be 10 digits in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    fmsgboxStatus "* Primary Mobile No. should be 10 digits in sheet Income Details"
                    
                    'Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                    
                End If
                
            Else
                
                If Len(Range("sheet1.Mobileno")) < 5 Or Len(Range("sheet1.Mobileno")) > 10 Then
                    
'                    fmsgboxStatus "* Mobile number 1 Is invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    fmsgboxStatus "* Primary Mobile No. Is invalid in sheet Income Details"
                    'Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                    
                End If
                
            End If
            
            If Len(Range("sheet1.Mobileno")) > 0 Then
                
                If Not IsNumeric(Trim(Range("sheet1.Mobileno").value)) Then
                    
'                    fmsgboxStatus "* Mobile number 1 Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    fmsgboxStatus "* Primary Mobile No. Is Invalid in sheet Income Details"
                   ' Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                    
                End If
                
            End If
            
            'Ayush_03/06/2025
            If Len(Range("sheet1.Mobileno")) > 0 Then
              If Not (ValidateMobileNumber()) Then
                 Sheet1.Range("sheet1.Mobileno").value = ""
                End If
            End If
            
               ' GoTo endline
            
            If Len(Range("sheet1.Mobileno")) > 0 Then
                
                If Not CheckSpecialCharacter(Trim(Range("sheet1.Mobileno").value)) Then
                    
'                    fmsgboxStatus "* Mobile number 1 Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                    fmsgboxStatus "* Primary Mobile No. Is Invalid in sheet Income Details"
                   ' Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                    
                End If
                
            End If
            
            If Len(Range("sheet1.Mobileno")) > 0 Then
                
                If Mid(Trim(Range("sheet1.Mobileno").value), 1, 1) = 0 Then
                    
'                    fmsgboxStatus "* Mobile number 1 should Not begin With 0"        ', vbOKOnly, vbMessageCaption
                    fmsgboxStatus "* Primary Mobile No. cannot begin With 0"
                    'Range("sheet1.Mobileno").Select
                    
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                    
                End If
                
            End If
            
            If Len(Range("sheet1.Mobileno")) + Len(Range("sheet1.mobileCountryCode")) > 15 Then
                
'                fmsgboxStatus "* Mobile number And country code combination should Not be more than 15 digits.in Sheet Income Details"        ', vbOKOnly, vbMessageCaption
                fmsgboxStatus "* Primary Mobile No. And country code combination should Not be more than 15 digits.in Sheet Income Details"        ', vbOKOnly, vbMessageCaption
                'Range("sheet1.Mobileno").Select
                
                Range("sheet1.Mobileno").MergeArea.ClearContents
                
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.mobileCountryCode").value = ""
                
            End If
            
        End If
        
    End If
    
 '   If (Target.name.name) = "sheet1.mobileCountryCode" Then
           If Not Application.Intersect(Sheet1.Range("sheet1.mobileCountryCode"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Len(Range("sheet1.mobileCountryCode")) <= 5 Then
            
            If Not IsNumeric(Trim(Range("sheet1.mobileCountryCode").value)) Then
                
'                fmsgboxStatus "* Country code For Mobile number 1 Is Invalid in sheet Income Details"        '' , vbOKOnly, vbMessageCaption
                fmsgboxStatus "* Country code in Primary Mobile No. Is Invalid in sheet Income Details"
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.mobileCountryCode").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.mobileCountryCode")) <= 5 Then
            
            If Not CheckSpecialCharacter(Trim(Range("sheet1.mobileCountryCode").value)) Then
                
'                fmsgboxStatus "* Country code For Mobile number 1 Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                fmsgboxStatus "* Country code in Primary Mobile No. Is Invalid in sheet Income Details"
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.mobileCountryCode").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.mobileCountryCode")) <= 5 Then
            
            If Not chkNumeric(Trim(Range("sheet1.mobileCountryCode").value)) Then
                
                fmsgboxStatus "* Country code in Primary Mobile No. Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.mobileCountryCode").value = ""
                
            End If
            
        End If
        
    End If
    
 '   If (Target.name.name) = "sheet1.MobileNoSec" Then
          If Not Application.Intersect(Sheet1.Range("sheet1.MobileNoSec"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Sheet1.Range("sheet1.mobileCountryCode2").value = "91" Then
            
            If Len(Range("sheet1.MobileNoSec")) <> 10 Then
                
                fmsgboxStatus "* Secondary Mobile No. should be 10 digits in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                'Range("sheet1.MobileNoSec").Select
                
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
                
            End If
            
        Else
            
            If Len(Range("sheet1.MobileNoSec")) < 5 Or Len(Range("sheet1.MobileNoSec")) > 10 Then
                
                fmsgboxStatus "* Secondary Mobile No. Is invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
               ' Range("sheet1.MobileNoSec").Select
                
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
                
            End If
            
        End If
        
        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            
            If Not IsNumeric(Trim(Range("sheet1.MobileNoSec").value)) Then
                
                fmsgboxStatus "* Secondary Mobile No. Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
               ' Range("sheet1.MobileNoSec").Select
                
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
                
            End If
            
        End If
        
        'Ayush_03/06/2025
            If Len(Range("sheet1.MobileNoSec")) > 0 Then
              If Not (ValidateMobileNumberSec()) Then
                 Sheet1.Range("sheet1.MobileNoSec").value = ""
                End If
            End If
        
        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            
            If Not CheckSpecialCharacter(Trim(Range("sheet1.MobileNoSec").value)) Then
                
                fmsgboxStatus "* Secondary Mobile No. Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                'Range("sheet1.MobileNoSec").Select
                
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
                
            End If
            
        End If
        
        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            
            If Mid(Trim(Range("sheet1.MobileNoSec").value), 1, 1) = 0 Then
                
                fmsgboxStatus "* Secondary Mobile No. should Not begin With 0"        ' , vbOKOnly, vbMessageCaption
                
              '  Range("sheet1.MobileNoSec").Select
                
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
                
            End If
            
        End If
        
        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            
            If (Range("sheet1.mobileCountryCode2").value) = "" Then
                
                fmsgboxStatus "* Please enter country code For Secondary Mobile No."        ', vbOKOnly, vbMessageCaption
                
            End If
            
        End If
        
        If Len(Range("sheet1.MobileNoSec")) + Len(Range("sheet1.mobileCountryCode2")) > 15 Then
            
            fmsgboxStatus "* Secondary Mobile No. And country code combination should Not be more than 15 digits.in Sheet Income Details"        ', vbOKOnly, vbMessageCaption
            
            'Range("sheet1.MobileNoSec").Select
            
           Range("sheet1.MobileNoSec").MergeArea.ClearContents
            
           ' Range("sheet1.mobileCountryCode2").Select
            
            Range("sheet1.mobileCountryCode2").value = ""
            
        End If
        
    End If
    
  '  If (Target.name.name) = "sheet1.mobileCountryCode2" Then
         If Not Application.Intersect(Sheet1.Range("sheet1.mobileCountryCode2"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Len(Range("sheet1.mobileCountryCode2")) <= 5 Then
            
            If Not IsNumeric(Trim(Range("sheet1.mobileCountryCode2").value)) Then
                
                fmsgboxStatus "* Country code in Secondary Mobile No. Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
              '  Range("sheet1.mobileCountryCode2").Select
                
                Range("sheet1.mobileCountryCode2").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.mobileCountryCode2")) <= 5 Then
            
            If Not CheckSpecialCharacter(Trim(Range("sheet1.mobileCountryCode2").value)) Then
                
                fmsgboxStatus "* Country code in Secondary Mobile No. Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                Range("sheet1.mobileCountryCode2").Select
                
                Range("sheet1.mobileCountryCode2").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.mobileCountryCode2")) <= 5 Then
            
            If Not chkNumeric(Trim(Range("sheet1.mobileCountryCode2").value)) Then
                
                fmsgboxStatus "* Country code in Secondary Mobile No. Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                Range("sheet1.mobileCountryCode2").Select
                
                Range("sheet1.mobileCountryCode2").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.mobileCountryCode2")) > 0 Then
            
            If (Range("sheet1.MobileNoSec").value) = "" Then
                
                fmsgboxStatus "* Please enter Secondary Mobile No."        ', vbOKOnly, vbMessageCaption
                
            End If
            
        End If
        
    End If
    
    'PAN Validation on data entry
    
   ' If (Target.name.name) = "sheet1.PAN" Then
       If Not Application.Intersect(Sheet1.Range("sheet1.PAN"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        Dim TargetPAN As String
        
        TargetPAN = ""
        
        TargetPAN = Target.name.name
        
        If Not mdIncomeDetails.CheckPAN(UCase(Trim(Range(TargetPAN).value))) Then
            
            fmsgboxStatus "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet.in Sheet Income Details"        ', vbOKCancel, mdCommon.vbMessageCaption
            
            Sheet1.Range(TargetPAN).Select
            
            Sheet1.Range(TargetPAN).value = ""
            
            GoTo endline
            
        End If
        Application.EnableEvents = False
        Sheet1.Range(TargetPAN).value = UCase(Range(TargetPAN).value)
        
        'PANStatusChange
        
        '                If Mid(UCase(Trim(Sheet1.Range("sheet1.Status").value)), 1, 1) = "I" Then
        '
        Sheet5.Range("Ver.PAN").value = UCase(Sheet1.Range("sheet1.PAN").value)
        '
        '                End If
        'Application.EnableEvents = True
        GoTo endline
        
    End If
    
    
'    Call mdIncomeDetails.AutoPopulateSecondaryAddress
    Application.EnableEvents = True

    
'Call mdIncomeDetails.AutoPopulateSecondaryAddress
'    If (target.name.name) = "Sheet1.date10IE" Then
'
'        If Trim(Range("Sheet1.date10IE").value) <> "00/00/0000" And Range("Sheet1.date10IE").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE").value)) Then
'
'                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter a valid Date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
'
'                Sheet1.Range("Sheet1.date10IE").Select
'
'                Sheet1.Range("Sheet1.date10IE").value = ""
'
'            Else
'
'                Sheet8.Unprotect Password:=getmsgstate
'                If (Sheet8.Range("oldbacValue").value = 1 And Sheet8.Range("bacValue").value = 1) Then
'
'                    If ChkMaxDatePrev(Trim(Range("Sheet1.date10IE").value), "") Then
'
'                        'fmsgboxStatus "* Date of filing of Form 10IE should not be prior 01/04/2021 in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
'                        fmsgboxStatus "* Date of filing cannot be prior To 1st April 2021 in Sheet Income Details"
'                        Sheet1.Range("Sheet1.date10IE").Select
'
'                        Sheet1.Range("Sheet1.date10IE").value = ""
'                    End If
'                Else
'
'                    If ChkMaxDate(Trim(Range("Sheet1.date10IE").value), "") Then
'
'                        fmsgboxStatus "* Date of filing cannot be prior To 1st April 2022 in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
'
'                        Sheet1.Range("Sheet1.date10IE").Select
'
'                        Sheet1.Range("Sheet1.date10IE").value = ""
'
'                    End If
'                End If
'
'                If (Sheet8.Range("oldbacValue").value = 1 And Sheet8.Range("bacValue").value = 1) Then
'                    If Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") >= Dformat(Trim("01/04/2022"), "yyyy-mm-dd") Then
'                        fmsgbox "* Date of filing of Form 10IE cannot be beyond 31st March, 2022 in Income Details"
'                        Range("Sheet1.date10IE").Select
'                        Range("Sheet1.date10IE").value = ""
'                    End If
'                Else
'                    If Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'                        fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("Sheet1.date10IE").Select
'                        Range("Sheet1.date10IE").value = ""
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'    End If

''Ayush_13/01/2026_V0.2
'
' If (Target.name.name) = "sheet1.PinCode2" Then
'
'        If Len(Range("sheet1.PinCode2").value) = 6 Then
'
'            If Range("sheet1.PinCode2").value = 0 Then
'
'                fmsgboxStatus ("* Pincode must be exactly 6 digits And should Not start With zero in Sheet Income Details")
'
'                Range("sheet1.PinCode2").MergeArea.ClearContents
'
'                Range("sheet1.PinCode2").Select
'
'            End If
'
'            If Len(Range("sheet1.PinCode2").value) = 6 Then
'
'                If Not ((IsNumeric(Range("sheet1.PinCode2").value)) And checkspecialcharacterForPinCode(Range("sheet1.PinCode2").value)) Then
'
'                    fmsgboxStatus ("* Invalid Pincode :Pincode cannot have Alphabets/Special characters And should be 6 digits Numeric Value in Sheet Income Details")
'
'                    Application.EnableEvents = False
'                    Range("sheet1.PinCode2").MergeArea.ClearContents
'
'                    Range("sheet1.PinCode2").Select
'
'                End If
'
'            End If
'
'            If Mid(Range("sheet1.PinCode2").value, 1, 1) = 0 Then
'
'                fmsgboxStatus ("* Pincode cannot begin With zero")
'                Application.EnableEvents = False
'                Range("sheet1.PinCode2").MergeArea.ClearContents
'
'                Range("sheet1.PinCode2").Select
'
'            End If
'
'        End If
'
'        If Range("sheet1.CountryCode2") = "" Or Range("sheet1.CountryCode2") = "(Select)" Then
'            Application.EnableEvents = False
'            Range("sheet1.CountryCode2") = "91-INDIA"
'
'        End If
'
'        Dim PinCode2 As Range
'
'
'        For Each PinCode2 In Sheet8.Range("All_Pincode_list")
'
'            Dim code2 As String
'
'            code2 = Left(CStr(Sheet1.Range("sheet1.PinCode2").value), 6)
'
'            If CStr(PinCode2.value) = CStr(code2) Then
'
'                Application.EnableEvents = False
'                Sheet1.Range("sheet1.StateCode2").value = Sheet8.Range("EQ" & PinCode2.row).value
'
'                Exit For
'
'            End If
'
'        Next PinCode2
'
'    End If
'
'    Dim pincode3 As Range
'        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
'            If Sheet1.Range("sheet1.PinCode2").value <> "" Then
'                'if state and pincode doesn't match, empty the pincode.
'                Dim state2 As Variant
'                state2 = Application.VLookup(Sheet1.Range("sheet1.PinCode2").value, Sheet8.Range("All_PinCode_V"), 2, False)
'                If Not (state2 = UCase(Mid(Sheet1.Range("sheet1.StateCode2").value, 4)) Or StateMatchesPin(Target.value)) Then
'                    'Ayush_01/01/2025
'                     'Newly added SIT-85381---
'                        Application.EnableEvents = False
'                        Sheet1.Unprotect Password:=getmsgstate
'                        '---ends
'                        Sheet1.Range("sheet1.PinCode2").value = ""
'
'                End If
'            End If
'        End If
'        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
'         Dim DupFlag2 As Boolean
'         DupFlag2 = False
'                 'If the pincode is not in the duplicate list, populate, else don't populate---
'                 For Each pincode3 In Sheet8.Range("Duplicate_Pincode_list")
'                    If Sheet1.Range("sheet1.PinCode2").value = pincode3 Then
'                        DupFlag2 = True
'                        Exit For
'                    End If
'                 Next pincode3
'
'
'                If DupFlag2 = False Then
'                    Dim StateName2
'                    Dim StateCode2
'
'                    StateName2 = Application.VLookup(Sheet1.Range("sheet1.PinCode2").value, Sheet8.Range("All_PinCode_V"), 2, False)
'                    If IsError(StateName2) Then
'                        'MsgBox ("Invalid Pincode")
'                        'Sheet1.Range("sheet1.StateCode1").value = "(Select)"
'                       Application.EnableEvents = False
'                        Sheet1.Range("sheet1.PinCode2").value = ""
'                     Application.EnableEvents = True
'
'                    Else
'                        StateCode2 = Application.VLookup(StateName2, Sheet8.Range("EW:EX"), 2, False)
'                        Sheet1.Range("sheet1.StateCode2").value = StateCode2
'                        Sheet1.Range("sheet1.CountryCode2").value = "91-INDIA"
'                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
'                    End If
'                 End If
'
'                 If DupFlag2 = True Then
'                 Application.EnableEvents = False
'
'                    If Sheet1.Range("sheet1.StateCode2").value <> "" Then
'                        If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode2").value) Then
'                            Sheet1.Range("sheet1.StateCode2").value = "(Select)"
'
'
'                        End If
'                    End If
'                 End If
'
'         End If
'
'
''-------------------
    
'    If (Target.name.name) = "sheet1.DOB" Then
'
'    'Ayush_Legacy45
' '   If Range("sheet1.status").value = "I - INDIVIDUAL" Then
'        If Range("sheet1.status").value = "I - INDIVIDUAL" Or Range("sheet1.status").value = "F - FIRM(Other than LLP)" Then
'
'            If Trim(Range("sheet1.DOB").value) = "01/01/0001" Then
''            "Ayush_10/02
''                fmsgboxStatus "* Date of Birth in Sheet Income Details Please enter a valid date.in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
'                fmsgboxStatus """Date of Birth/Formation - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
'
'
'                Sheet1.Range("sheet1.DOB").Select
'
'                Sheet1.Range("sheet1.DOB").value = ""
'
'            End If
'
'        End If
'
'        'Change-06.04.2023.103.PAG.01
'            AnyOtherDeductions80CCH Target
'        'End Change
'        LockUnlock80TTB
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'        '
'
'        Sheet1.Protect Password:=getmsgstate
'
'    End If
'
'
'
'    'DOB Validation: On Data Entry
'
'    If (Target.name.name) = "sheet1.DOB" Then
'
'        Dim HUFReq  As Long
'
'        If Range("sheet1.status").value = "H - HUF" Then HUFReq = 1 Else HUFReq = 0
'
'        If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
'
'            result = CheckDate(Trim(Range("sheet1.DOB").value), HUFReq)
'
'            If result = 1 Then
'  'Ayush_10/02
''                fmsgboxStatus "* Date of Birth in Sheet Income Details Please enter a valid Date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
'                 fmsgboxStatus """Date of Birth/Formation - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
'
'                Sheet1.Range("sheet1.DOB").Select
'
'                Sheet1.Range("sheet1.DOB").value = ""
'
'            ElseIf result = 2 Then
'
''Change-22.11.2023.103.02
''Ayush_26-27_C9
'Dim cutoff As Date
'cutoff = CDate(Sheet8.Range("DOB_1").value)
'
'                'fmsgboxStatus "* Date of Birth in Sheet Income Details  should Not be after 01/04/2022"        ', vbOKOnly, mdCommon.vbMessageCaption
'
''                fmsgboxStatus "* Date of Birth in Sheet Income Details should not be after 01/04/2023"
'
'            'PAG_C1 AY 2024-25 Changes
'            'Ayush_25-26
'            'AYush_Legacy44
'
'
''                fmsgboxStatus "* Date of Birth in Sheet Income Details should not be after 01/04/2024"
'               ' fmsgboxStatus "* Date of Birth in Sheet Income Details should not be after 01/04/2025"
'                 fmsgboxStatus "* Date of birth in sheet income details should not be on or after " & Dformat1(cutoff, "dd/mm/yyyy")
'
'
'                Sheet1.Range("sheet1.DOB").Select
'
'                Sheet1.Range("sheet1.DOB").value = ""
'
'            End If
'
'        End If
'
'        GoTo endline
'
'    End If
    
    'Date of filling orignal return Validation: On Data Entry
    
  '  If (Target.name.name) = "sheet1.OrigRetFiledDate" Then
    If Not Application.Intersect(Sheet1.Range("sheet1.OrigRetFiledDate"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Trim(Range("sheet1.OrigRetFiledDate").value) <> "00/00/0000" And Range("sheet1.OrigRetFiledDate").value <> "" Then
            
            If Not FormatNCheckDate(Trim(Range("sheet1.OrigRetFiledDate").value)) Then
  'Ayush_10/02
'                fmsgboxStatus "* Date Of filing Original Return Is Invalid in Sheet Income Details. Please enter a valid Date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
                fmsgboxStatus "* ""Date of filing of original Return - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
                
                Sheet1.Range("sheet1.OrigRetFiledDate").Select
                
                Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                
            Else
             'Change-13.12.2022.103.25
'                If ChkMaxDatePrev(Trim(Range("sheet1.OrigRetFiledDate").value), "") Then
'
'                    fmsgboxStatus "* Date Of  filing Orignal Return should Not be prior 01/04/2022 in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
  'Ayush_26-27_C57
  
  cutoff1 = CDate(Sheet8.Range("DOB_1").value)
  



                If ChkMaxDate(Trim(Range("sheet1.OrigRetFiledDate").value), "") Then
                
                'PAG_C30 AY 2024-25 Change
                'Ayush_25-26
                   ' fmsgboxStatus "* Date Of filing Orignal Return should Not be prior 01/04/2024 in Sheet Income Details"
                  '  fmsgboxStatus "* Date Of filing Original Return should not be prior 01/04/2025 in Sheet Income Details"
'                     fmsgbox "* Date Of filing Original Return should not be prior " & _
'          Dformat1(cutoff1, "mm/dd/yyyy") & " in Sheet Income Details"
                     fmsgbox "* ""Date of filing of original return cannot be prior to " & _
          Dformat1(cutoff1, "mm/dd/yyyy") & " of AY 26-27"""
'                     fmsgboxStatus "* ""Date of filling cannot be prior to 01/04/2025 in Part A General 139(8A).""" '                    fmsgboxStatus "* Filing section u/s. 139(4) cannot be selected till 31st July, 2024 of such AY or extended due date"
                     
'                    Sheet1.Range("sheet1.OrigRetFiledDate").Select
                    
                    Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                    
                End If
                
            End If
            
        End If
        
 
        If Dformat(Trim(Range("sheet1.OrigRetFiledDate").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
            
'            fmsgbox "* Date of filing of original Return can Not be more than system Date in Income Details"
                fmsgbox "* Date of filing of original return shall not be beyond system date."
            'Date of filing of original return shall not be beyond system date.
            
            Range("sheet1.OrigRetFiledDate").Select
            
            Range("sheet1.OrigRetFiledDate").value = ""
            
        End If
        
        GoTo endline
        
    End If
    
    'Date of filling in response to notice Validation: On Data Entry
    
 '   If (Target.name.name) = "sheet1.NoticeDate" Then
         If Not Application.Intersect(Sheet1.Range("sheet1.NoticeDate"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Trim(Range("sheet1.NoticeDate").value) <> "00/00/0000" Then
            
            If Not FormatNCheckDate(Range("sheet1.NoticeDate").value) Then
                'PAG_E31 AY 2023-24
'                fmsgboxStatus "* Please enter a valid Date in dd/mm/yyyy format in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
                
               'PAG_C31 AY 2024-25 Change
               'Ayush_10/02
'                fmsgboxStatus "* Please enter a date in ""Date of notice /order"" field in dd/mm/yyyy format"        ', vbOKOnly, mdCommon.vbMessageCaption
                 fmsgboxStatus "* ""Date of such Notice or Order - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
                
                Sheet1.Range("sheet1.NoticeDate").Select
                
                Sheet1.Range("sheet1.NoticeDate").value = ""
                
            Else
               
                
                
                If Range("sheet1.NoticeDate").value <> "" Then
                    
                    If ChkMaxDate(Range("sheet1.NoticeDate").value, "") Then
                'Ayush_26-27_C60
'  Dim cutoff2 As Date
'  cutoff2 = CDate(Sheet8.Range("DOB_1").value)
                    'Change-08.12.2023.103.27
                        'fmsgboxStatus "*Date of notice Or Order cannot be prior To 01/04/2022 in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
                        
                        
'                        PAG_E31 AY 2023-24 Old
'                        fmsgboxStatus "*Date of notice Or Order cannot be prior To 01/04/2023 in Sheet Income Details"

                        'PAG_C31 AY 2024-25 Change
                        'Ayush_25-26
                      '  fmsgboxStatus "*Date of Notice/Order cannot be prior to 01/04/2024"
'                        fmsgboxStatus "*Date of Notice/Order cannot be prior to 01/04/2025"
                        fmsgboxStatus "*Date of Notice/Order cannot be prior to " & Dformat1(cutoff1, "mm/dd/yyyy")
                        
                        Sheet1.Range("sheet1.NoticeDate").Select
                        
                        Sheet1.Range("sheet1.NoticeDate").value = ""
                        
                    End If
                    
                End If
                
            End If
            
            result = ""
            
        End If
        

        
        If Dformat(Trim(Range("sheet1.NoticeDate").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
            
            fmsgbox "* Date of such notice or order cannot be more than system Date in Income Details"
            
            Range("sheet1.NoticeDate").Select
            
            Range("sheet1.NoticeDate").value = ""
            
        End If
        
        GoTo endline
        
    End If
    
    'EMAIL Validation : On data entry
    
'    If (Target.name.name) = "sheet1.EmailAddress" Then
        If Not Application.Intersect(Sheet1.Range("sheet1.EmailAddress"), Sheet1.Range(Target.Address)) Is Nothing Then
      
        If Range("sheet1.EmailAddress").value <> "" Then
            
            If Not mdIncomeDetails.CheckEmailAddress(Trim(Range("sheet1.EmailAddress").value), validateType) Then
                
                If validateType = "1" Then fmsgboxStatus "* Invalid Primary Email ID of the taxpayer Please retry."        ', vbOKOnly, mdCommon.vbMessageCaption
                
                If validateType = "2" Then fmsgboxStatus "* Primary Email ID of the taxpayer should Not exceed 125 characters in Sheet Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
                'Ayush_Legacy39
'                If validateType = "3" Then fmsgboxStatus "* Email ID (Self) Is invalid in Sheet  Income Details"        ', vbOKOnly, mdCommon.vbMessageCaption
                
                 If validateType = "3" Then fmsgboxStatus "* ""Invalid Primary Email ID of the taxpayer Please retry."""
                If validateType = "4" Then fmsgboxStatus "* Primary Email ID of the taxpayer Is mandatory in tab Part A General Information"        ', vbOKOnly, mdCommon.vbMessageCaption
                
                Range("sheet1.EmailAddress").Select
                
                Range("sheet1.EmailAddress").value = ""
                
            End If
            
            GoTo endline
            
        End If
        
    End If
    
  '  If (Target.name.name) = "sheet1.EmailAddress2" Then
      If Not Application.Intersect(Sheet1.Range("sheet1.EmailAddress2"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Range("sheet1.EmailAddress2").value <> "" Then
            
            If Not mdIncomeDetails.CheckEmailAddress(Trim(Range("sheet1.EmailAddress2").value), validateType) Then
                
                If validateType = "1" Then fmsgboxStatus "* Invalid Secondary Email ID Please retry"   ', vbOKOnly, mdCommon.vbMessageCaption
                
                If validateType = "2" Then fmsgboxStatus "* Secondary Email ID cannot be more than 125 characters"  ', vbOKOnly, mdCommon.vbMessageCaption
                
              'Error Message Changed By Jyoti for the SIT-67667 on 25/06/2024
                
              ' If validateType = "3" Then fmsgboxStatus "* E-mail ID-2 Is invalid in Sheet Income Details"    ', vbOKOnly, mdCommon.vbMessageCaption
                If validateType = "3" Then fmsgboxStatus "* ""Invalid Secondary Email ID Please retry."""    ', vbOKOnly, mdCommon.vbMessageCaption
              'end change
                
                'If validateType = "4" Then MsgBox "Email Address in Sheet Income Details  is Mandatory", vbOKOnly, mdCommon.vbMessageCaption
                
                Range("sheet1.EmailAddress2").Select
                
                Range("sheet1.EmailAddress2").value = ""
                
            End If
            
            GoTo endline
            
        End If
        
    End If
    
    'Std & phone" On data Entry
    
  '  If (Target.name.name) = "sheet1.STDcode" Then
    If Not Application.Intersect(Sheet1.Range("sheet1.STDcode"), Sheet1.Range(Target.Address)) Is Nothing Then


        Dim val     As Variant

        val = mdIncomeDetails.ValidateSTDcode(Trim(Range("sheet1.STDcode").value))

        If val <> "0" Then
        

            fmsgboxStatus "* STD/ISD Code " & val & " in Sheet Income Details "        ', vbOKOnly, mdCommon.vbMessageCaption

            Sheet1.Range("sheet1.STDcode").value = ""

            Sheet1.Range("sheet1.STDcode").Select

            GoTo endline

        End If

        result = mdIncomeDetails.ValidatePhoneNo(Trim(Range("sheet1.STDcode").value), Trim(Range("sheet1.PhoneNo").value))

        If result <> "0" Then
       

            fmsgboxStatus "* Phone Number " & result & " in Sheet  Income Details "        ', vbOKOnly, vbMessageCaption

            Sheet1.Range("sheet1.PhoneNo").value = ""

            GoTo endline

        End If

        stDcode = Sheet1.Range("sheet1.STDcode").value

        GoTo endline

       End If
    
    'Phone NO" On data Entry
    
 '   If (Target.name.name) = "sheet1.PhoneNo" Then
  If Not Application.Intersect(Sheet1.Range("sheet1.PhoneNo"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Sheet1.Range("sheet1.STDcode").value <> "" Then
            
            If mdIncomeDetails.ValidateSTDcode(Trim(Range("sheet1.STDcode").value)) = 0 Then
                
                result = mdIncomeDetails.ValidatePhoneNo(Trim(Range("sheet1.STDcode").value), Trim(Range("sheet1.PhoneNo").value))
                
                If result <> "0" Then
                    
                    fmsgboxStatus "* Phone Number " & result & " in Sheet Income Details "        ', vbOKOnly, vbMessageCaption
                    
                    Sheet1.Range("sheet1.PhoneNo").value = ""
                    
                    GoTo endline
                    
                End If
                
            Else
                
                Sheet1.Range("sheet1.PhoneNo").value = ""
                
                Sheet1.Range("sheet1.STDcode").Select
                
                GoTo endline
                
            End If
            
        End If
        
        phoneNo = Sheet1.Range("sheet1.PhoneNo").value
        
        GoTo endline
        
    End If
    
'    If (Target.name.name) = "sheet1.ReceiptNo" Then
 If Not Application.Intersect(Sheet1.Range("sheet1.ReceiptNo"), Sheet1.Range(Target.Address)) Is Nothing Then

        
        If Range("Sheet1.ReceiptNo").value <> "" Then
            
            If Not chkNumeric(Range("Sheet1.ReceiptNo").value) Then
                
                fmsgboxStatus ("* Reciept Number Is Invalid in Sheet Income Details")
                
                Range("Sheet1.ReceiptNo").value = ""
                
                Range("Sheet1.ReceiptNo").Select
                
            End If
            
            If Range("Sheet1.ReceiptNo").value = "000000000000000" Then
                
                fmsgboxStatus ("* Reciept Number Is Invalid in Sheet Income Details")
                
                Range("Sheet1.ReceiptNo").value = ""
                
                Range("Sheet1.ReceiptNo").Select
                
            End If
            
        End If
        
    End If
    
End If

'If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("sheet1.Status")) Is Nothing Then
'    Sheet201.Unprotect Password:=getmsgstate
'    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
'        Sheet201.Range("U_AadhaarEnrolmentId").MergeArea.ClearContents
'        Sheet201.Unprotect Password:=getmsgstate
'        Sheet201.Range("U_AadhaarCardNo").MergeArea.ClearContents
'        Sheet201.Unprotect Password:=getmsgstate
'        Sheet201.Range("U_AadhaarEnrolmentId").MergeArea.Locked = True
'        Sheet201.Range("U_AadhaarEnrolmentId").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_AadhaarCardNo").MergeArea.Locked = True
'        Sheet201.Range("U_AadhaarCardNo").MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Unprotect Password:=getmsgstate
'        Sheet201.Range("U_AadhaarEnrolmentId").MergeArea.Locked = False
'        Sheet201.Range("U_AadhaarEnrolmentId").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_AadhaarCardNo").MergeArea.Locked = False
'        Sheet201.Range("U_AadhaarCardNo").MergeArea.Interior.Color = "&HCCFFCC"
'    End If
'    Sheet201.Protect Password:=getmsgstate
'End If

'Ayush_12/01/2026_V0.2
'If (Target.name.name) = "sheet1.EmailRepAssessee" Then
 If Not Application.Intersect(Sheet1.Range("sheet1.EmailRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then
        
        If Range("sheet1.EmailRepAssessee").value <> "" Then
            
            If Not mdIncomeDetails.CheckEmailAddress_rep(Trim(Range("sheet1.EmailRepAssessee").value), validateType) Then
                
                If validateType = "1" Then fmsgboxStatus "* ""Email Address of the representative entered is invalid"""        ', vbOKOnly, mdCommon.vbMessageCaption
                
                If validateType = "2" Then fmsgboxStatus "* ""Email Address of the representative cannot be more than 125 characters"""        ', vbOKOnly, mdCommon.vbMessageCaption
     
                If validateType = "3" Then fmsgboxStatus "* ""Email Address of the representative entered is invalid"""
                
                If validateType = "4" Then fmsgboxStatus "* Please enter the Email Address of representative in Sheet Income details"
                
                Range("sheet1.EmailRepAssessee").Select
                
                Range("sheet1.EmailRepAssessee").value = ""
                
            End If
            
            GoTo endline
            
        End If
        
    End If
    
    
    

 '   If (Target.name.name) = "sheet1.CountryCodeRepAssessee" Then
 If Not Application.Intersect(Sheet1.Range("sheet1.CountryCodeRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then
        
        If Len(Range("sheet1.CountryCodeRepAssessee")) <= 5 Then
            
            If Not IsNumeric(Trim(Range("sheet1.CountryCodeRepAssessee").value)) Then
                
                fmsgboxStatus "* Country code in contact No. of representative assessee Is Invalid in sheet Income Details"        '' , vbOKOnly, vbMessageCaption
                
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.CountryCodeRepAssessee").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.CountryCodeRepAssessee")) <= 5 Then
            
            If Not CheckSpecialCharacter(Trim(Range("sheet1.CountryCodeRepAssessee").value)) Then
                
                fmsgboxStatus "* Country code in contact No. of representative assessee Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.CountryCodeRepAssessee").value = ""
                
            End If
            
        End If
        
        If Len(Range("sheet1.CountryCodeRepAssessee")) <= 5 Then
            
            If Not chkNumeric(Trim(Range("sheet1.CountryCodeRepAssessee").value)) Then
                
                fmsgboxStatus "* Country code in contact No. of representative assessee Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
                
                'Range("sheet1.mobileCountryCode").Select
                
                Range("sheet1.CountryCodeRepAssessee").value = ""
                
            End If
            
        End If
        
    End If
    
    
'    If (Target.name.name) = "sheet1.MobileRepAssessee" Then
'
'        If (Range("sheet1.MobileRepAssessee")) <> "" Then
'
'            If Sheet1.Range("sheet1.CountryCodeRepAssessee").value = "91" Then
'
'                If Len(Range("sheet1.MobileRepAssessee")) <> 10 Then
'
'                    fmsgboxStatus "* Mobile number 1 should be 10 digits in sheet Income Details"        ', vbOKOnly, vbMessageCaption
'
'
'
'                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
'
'                End If
'
'            Else
'
'                If Len(Range("sheet1.MobileRepAssessee")) < 5 Or Len(Range("sheet1.MobileRepAssessee")) > 10 Then
'
'                    fmsgboxStatus "* Mobile number 1 Is invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
'
'                    'Range("sheet1.Mobileno").Select
'
'                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
'
'                End If
'
'            End If
'
'            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
'
'                If Not IsNumeric(Trim(Range("sheet1.MobileRepAssessee").value)) Then
'
'                    fmsgboxStatus "* Mobile number 1 Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
'
'                   ' Range("sheet1.Mobileno").Select
'
'                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
'
'                End If
'
'            End If
'
'            'Ayush_03/06/2025
'            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
'              If Not (ValidateMobileNumber()) Then
'                 Sheet1.Range("sheet1.MobileRepAssessee").value = ""
'                End If
'            End If
'
'               ' GoTo endline
'
'            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
'
'                If Not CheckSpecialCharacter(Trim(Range("sheet1.MobileRepAssessee").value)) Then
'
'                    fmsgboxStatus "* Mobile number 1 Is Invalid in sheet Income Details"        ', vbOKOnly, vbMessageCaption
'
'                   ' Range("sheet1.Mobileno").Select
'
'                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
'
'                End If
'
'            End If
'
'            If Len(Range("sheet1.MobileRepAssessee")) > 0 Then
'
'                If Mid(Trim(Range("sheet1.MobileRepAssessee").value), 1, 1) = 0 Then
'
'                    fmsgboxStatus "* Mobile number 1 should Not begin With 0"        ', vbOKOnly, vbMessageCaption
'
'                    'Range("sheet1.Mobileno").Select
'
'                    Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
'
'                End If
'
'            End If
'
'            If Len(Range("sheet1.MobileRepAssessee")) + Len(Range("sheet1.mobileCountryCode")) > 15 Then
'
'                fmsgboxStatus "* Mobile number And country code combination should Not be more than 15 digits.in Sheet Income Details"        ', vbOKOnly, vbMessageCaption
'
'                'Range("sheet1.Mobileno").Select
'
'                Range("sheet1.MobileRepAssessee").MergeArea.ClearContents
'
'                'Range("sheet1.mobileCountryCode").Select
'
'                Range("sheet1.MobileRepAssessee").value = ""
'
'            End If
'
'        End If
'
'    End If
    
'--------------------
    
'If (Target.name.name) = "sheet1.Status" Or (Target.name.name) = "sheet1.MiddleName" Or (Target.name.name) = "sheet1.FirstName" Or (Target.name.name) = "sheet1.SurNameOrOrgName" Then
   If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("sheet1.MiddleName"), Sheet1.Range(Target.Address)) Is Nothing Or _
             Not Application.Intersect(Sheet1.Range("sheet1.FirstName"), Sheet1.Range(Target.Address)) Is Nothing Or _
             Not Application.Intersect(Sheet1.Range("sheet1.SurNameOrOrgName"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    If (Mid(Range("sheet1.status").value, 1, 1) = "I") Or (Mid(Range("sheet1.status").value, 1, 1) = "(") Then
        
        sPassword = mdCommon.getmsgstate
        
        Sheet5.Unprotect Password:=sPassword
        
        Sheet5.Range("Ver.AssesseeVerName").value = UCase(Sheet1.Range("sheet1.FirstName").value + " " + Sheet1.Range("sheet1.MiddleName").value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").value)
        
        Sheet5.Protect Password:=sPassword
        
        GoTo endline
        
    End If
    
End If

'Ayush_21/04/2026
If Not Application.Intersect(Sheet1.Range("Sec80CCC_Type"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Target.value = "PRAN" Then
        DPMAddComment_PRAN Target
    Else
        DPMADeleteComment_PRAN Target
    
    End If
End If

'If (Target.name.name) = "IncD.Section80D_Calc" Then
If Not Application.Intersect(Sheet1.Range("IncD.Section80D_Calc"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    If (Range("IncD.Section80D_Calc").value) > 100000 Then
        
        fmsgboxStatus ("* Section80D Value cannot exceed 100000 in Sheet Income Details")
        
    End If
    
End If

'If (Target.name.name) = "IncD.Section80EE" Then
 If Not Application.Intersect(Sheet1.Range("IncD.Section80EE"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    If (Range("IncD.Section80EE").value) > 0 Then
        
        'And (Range("IncD.Section80EE").value) <> 0 Then
        
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
        
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
        
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Protect Password:=getmsgstate
        
    Else
        
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
        
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = False
        
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Protect Password:=getmsgstate
        
    End If
    
    Sheet1.Protect Password:=getmsgstate
    
End If

'DPK

'If (Target.name.name) = "IncD.Section80EEA" Then
 If Not Application.Intersect(Sheet1.Range("IncD.Section80EEA"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    If (Range("IncD.Section80EEA").value) > 0 Then
        
        'And (Range("IncD.Section80EE").value) <> 0 Then
        
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("IncD.Section80EE").MergeArea.ClearContents
        
        Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
        
        Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Protect Password:=getmsgstate
        
    Else
        
        Sheet1.Unprotect Password:=getmsgstate
        
        Sheet1.Range("IncD.Section80EE").MergeArea.ClearContents
        
        Sheet1.Range("IncD.Section80EE").MergeArea.Locked = False
        
        Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Protect Password:=getmsgstate
        
    End If
    
End If

'    If Target.Validation.Type = 1 Then

'-------------------------------------------------------
' 2024-25 Change Newly added by Sai on 08-01-2024
'    Sheet1.Unprotect Password:=getmsgstate
'            If (Target.name.name) = "IncD_dqdiv" Then
'            If Range("IncD_dqdiv").value > 0 Then
'                 Range("IncD_dq1div").value = Range("IncD_dqdiv").value
'            End If
'            End If





  If Not Application.Intersect(Sheet1.Range("IncD_dqdiv"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Range("IncD_dqdiv").value > 0 Then
              Range("IncD_dq1div").value = Range("IncD_dqdiv").value
              fmsgboxStatus ("If the benefit of quarterly computation of interest u/s 234C is sought, same may be claimed by making necessary modifications in below Table")
              
        End If
  End If



 
'If Range("IncD_dqdiv").value <> (Range("IncD_dq1div").value + Range("IncD_dq2div").value + Range("IncD_dq3div").value + Range("IncD_dq4div").value + Range("IncD_dq5div").value) Then
'
'        fmsgboxsmall "* Quarterly breakup of dividend income shall match with total dividend income " ', vbOKOnly, "Error(s)!"
'        End If
'End If



'-------------------------------------------------------------


'If (Target.name.name) = "IncD.Section89" Then
If Not Application.Intersect(Sheet1.Range("IncD.Section89"), Sheet1.Range(Target.Address)) Is Nothing Then

    If Range("IncD.Section89").value > 0 Then
        fmsgbox "* Please ensure To submit Form 10E."
    End If
    
    If (Range("IncD.Section89").value) = "" Then
        Range("IncD.Section89").value = 0
    End If
    
    If (Range("IncD.IncomeFromSal_i").value) = 0 Then
        
    End If
    
End If

'If (Target.name.name) = "IncD.Section89A" Then
If Not Application.Intersect(Sheet1.Range("IncD.Section89A"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Range("IncD.Section89A").value > Range("IncD.Section89ANew").value Then
        Range("IncD.Section89A").value = 0
        MsgBox "Section 89A cannot exceed Rs." & Range("IncD.Section89ANew").value
    End If
End If


'==================================================================
'If (Target.name.name) = "IncD.TypeOfHP" Then
If Not Application.Intersect(Sheet1.Range("IncD.TypeOfHP"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    If Not Range("IncD.TypeOfHP").value = "Self Occupied" Then
        
        If Sheet8.Range("bacValue").value = 1 Then
            
            Sheet1.Unprotect Password:=getmsgstate
            
            'IncD.IncomeFromHP_v this field become auto-populate changes made by Bindu on 23rd April 25
            'Range("IncD.IncomeFromHP_v").MergeArea.Locked = False
            
            'Range("IncD.IncomeFromHP_v").Interior.Color = (&HCCFFCC)        ', "AU"
            
        End If
        
    Else
        
        If Sheet8.Range("bacValue").value = 1 Then
             'PAG_E38 AY 2023-24 Old
'            MsgBox "In Case Type of House Property Is Self Occupied And Option  'opting in Now' or 'continue to opt' is selected for 'Option for current assessment year', deduction for interest payable on borrowed capital is not allowed"

             'PAG_C38 AY 2024-25 Change
            MsgBox "In case type of House Property is Self Occupied and new tax regime is selected,deduction for interest payable on borrowed capital is not allowed", vbOKOnly, "Alert!!"

            Sheet1.Unprotect Password:=getmsgstate
            
            'Commented by Bindu as its auto-populate field
            'Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
            
            'Range("IncD.IncomeFromHP_v").Interior.Color = "&HD8D8D8"   ', "AU"
            
            'Range("IncD.IncomeFromHP_v").value = ""
            
        End If
        
    End If
    
End If


'Added by Shrutika(20/04/2025)NewDev

'If Not Application.Intersect(Sheet1.Range("IncD.TypeOfHP"), Sheet1.Range(Target.Address)) Is Nothing Then

'If Sheet8.Range("Bacvalue").value = 2 Then

'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = vbWhite

'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True

'If Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied" Then
'
'
'Application.EnableEvents = False
'
'Sheet1.Range("IncD.IncomeFromHP_v").Formula = "=IF(TotAmt.24b>200000,200000,MIN(200000,TotAmt.24b))"
'Application.EnableEvents = True
'End If
'
'End If
'
'If Sheet8.Range("Bacvalue").value = 1 Then
'
'If Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied" Then
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = vbWhite
'
'Sheet1.Range("IncD.IncomeFromHP_v").value = 0
'
'Else
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = vbWhite
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'
'Application.EnableEvents = False
'
'Sheet1.Range("IncD.IncomeFromHP_v").Formula = "=IF(TotAmt.24b>200000,200000,MIN(200000,TotAmt.24b))"
'
' Application.EnableEvents = True
'End If
'
'End If
'
''End If
'
'If Not Application.Intersect(Sheet1.Range("IncD.IncomeFromHP_i"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'If Sheet1.Range("IncD.GrossRentRecieved").value = 0 Then
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = vbWhite
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'
'Sheet1.Range("IncD.IncomeFromHP_v").value = 0
'
'Else
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = vbWhite
'
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'Application.EnableEvents = False
'Sheet1.Range("IncD.IncomeFromHP_v").Formula = "=IF(TotAmt.24b>200000,200000,MIN(200000,TotAmt.24b))"
'Application.EnableEvents = True
'End If
'
'End If

'------------------------------------------------
 
'===============================================

'Change-04.01.2023.103.51

'ID_234ABCF_LockUnlock 'For PAG_E67,PAG_E68,PAG_E69 AY 2023-24 Old

ID_234ABCF_LockUnlock_New 'For PAG_C67,PAG_C68,PAG_C69 AY 2024-25 Change

'If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18) Then
'    Sheet1.Unprotect Password:=sPassword
'    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Interior.ColorIndex = 2
'    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Interior.ColorIndex = 2
'    Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Interior.ColorIndex = 2
'    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = 2
'    Sheet1.Protect Password:=sPassword
'Else
'    Sheet1.Unprotect Password:=sPassword
'    Range("IncD.IntrstPayUs234A").MergeArea.Locked = False
'    Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = RGB(255, 255, 204)
'    Range("IncD.IntrstPayUs234B").MergeArea.Locked = False
'    Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = RGB(255, 255, 204)
'    Sheet1.Protect Password:=sPassword
'End If


'Newly added by Chandru

'
'If Sheet1.Range("sheet1.Aadhaar").value <> "" Then
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = True
'Sheet1.Range("sheet1.AadhaarEnrol").Interior.Color = "&HD8D8D8"
'End If
'
'If Sheet1.Range("sheet1.Aadhaar").value = "" Then
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("sheet1.AadhaarEnrol").MergeArea.Locked = False
'Sheet1.Range("sheet1.AadhaarEnrol").Interior.Color = (&HCCFFCC)
'
'End If
'Jyoti
Dim PinCode1 As Range
        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.PinCode").value <> "" Then
                'if state and pincode doesn't match, empty the pincode.
                Dim state1 As Variant
                state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, Sheet8.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Sheet1.Range("sheet1.StateCode1").value, 4)) Or StateMatchesPin(Target.value)) Then
                    'Ayush_01/01/2025
                     'Newly added SIT-85381---
                        Application.EnableEvents = False
                        Sheet1.Unprotect Password:=getmsgstate
                        '---ends
                        Sheet1.Range("sheet1.PinCode").value = ""
                       
                End If
            End If
        End If
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then
         Dim DupFlag As Boolean
         DupFlag = False
                 'If the pincode is not in the duplicate list, populate, else don't populate---
                 For Each PinCode1 In Sheet8.Range("Duplicate_Pincode_list")
                    If Sheet1.Range("sheet1.PinCode").value = PinCode1 Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode1
                 
                
                If DupFlag = False Then
                    Dim StateName
                    Dim StateCode
                    
                    StateName = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, Sheet8.Range("All_PinCode_V"), 2, False)
                    If IsError(StateName) Then
                        'MsgBox ("Invalid Pincode")
                        'Sheet1.Range("sheet1.StateCode1").value = "(Select)"
                       Application.EnableEvents = False
                        Sheet1.Range("sheet1.PinCode").value = ""
                     Application.EnableEvents = True
                        
                    Else
                        StateCode = Application.VLookup(StateName, Sheet8.Range("EW:EX"), 2, False)
                        Sheet1.Range("sheet1.StateCode1").value = StateCode
                        Sheet1.Range("sheet1.CountryCode").value = "91-INDIA"
                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
                    End If
                 End If
                 
                 If DupFlag = True Then
                 Application.EnableEvents = False
                 
                    If Sheet1.Range("sheet1.StateCode1").value <> "" Then
                        If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode1").value) Then
                            Sheet1.Range("sheet1.StateCode1").value = "(Select)"
                         
                            
                        End If
                    End If
                 End If

         End If

''Malli 31/03/2026
'If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceNo"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.ResidenceName"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.RoadOrStreet"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.LocalityOrArea"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.CityOrTownOrDistrict"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("HASZIP"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.ZipCode"), Sheet1.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'
'Call mdIncomeDetails.AutoPopulateSecondaryAddress
'End If
'================================
'Ayush_commented_26-27_changes
'Ayush_14/05
'Malli_SIT-118203 27/04/2026

'If Range("IncD.Section80CCD_SE").value > 0 Or Range("IncD.Section80CCD1B_SE").value > 0 Then

If (Range("IncD.Section80CCD_SE").value > 0 Or Range("IncD.Section80CCD1B_SE").value > 0) And Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
'---------------------------

Sheet1.Unprotect Password:=sPassword
'Sheet1.Range("Pran_Number").MergeArea.Locked = False
Sheet1.Range("pran_new").Locked = False
Sheet1.Range("pran_new").Interior.Color = (&HCCFFCC)


Else
Sheet1.Unprotect Password:=sPassword

Sheet1.Range("pran_new").ClearContents
Sheet1.Range("pran_new").Interior.Color = (&HD8D8D8)
Sheet1.Range("pran_new").Locked = True


End If

'Ayush_26-27_C125
'If Not Application.Intersect(Sheet1.Range("IncD.Section80CCC"), Sheet1.Range(Target.Address)) Is Nothing Then
'If Target.value > 0 Then
'Sheet1.Unprotect Password:=sPassword
'Application.EnableEvents = False
'
'
'Range("Sec80CCC_Type").Locked = False
'Range("Sec80CCC_Type").Interior.Color = (&HCCFFCC)
'
'Range("Sec80CCC_Name").Locked = False
'Range("Sec80CCC_Name").Interior.Color = (&HCCFFCC)
'
'Range("Sec80CCC_Amount").Locked = False
'Range("Sec80CCC_Amount").Interior.Color = (&HCCFFCC)
'
'Else
'
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_Type").Locked = True
'Range("Sec80CCC_Type").ClearContents
'Range("Sec80CCC_Type").Interior.Color = (&HD8D8D8)
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_Name").Locked = True
'Range("Sec80CCC_Name").ClearContents
'Range("Sec80CCC_Name").Interior.Color = (&HD8D8D8)
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_Amount").Locked = True
'Range("Sec80CCC_Amount").ClearContents
'Range("Sec80CCC_Amount").Interior.Color = (&HD8D8D8)
'
'End If
'
'Application.EnableEvents = True
'End If
'---------------------------------------------------------

'Ayush_26-27_C131
'If Not Application.Intersect(Sheet1.Range("IncD.Section80CCD_SE"), Sheet1.Range(Target.Address)) Is Nothing Then
'If Target.value > 0 Then
'Sheet1.Unprotect Password:=sPassword
'Application.EnableEvents = False
'
'
'Range("Sec80CCC_1_Type").Locked = False
'Range("Sec80CCC_1_Type").Interior.Color = (&HCCFFCC)
'
'Range("Sec80CCC_1_Name").Locked = False
'Range("Sec80CCC_1_Name").Interior.Color = (&HCCFFCC)
'
'Range("Sec80CCC_1_Amount").Locked = False
'Range("Sec80CCC_1_Amount").Interior.Color = (&HCCFFCC)
'
'Else
'
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_1_Type").Locked = True
'Range("Sec80CCC_1_Type").ClearContents
'Range("Sec80CCC_1_Type").Interior.Color = (&HD8D8D8)
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_1_Name").Locked = True
'Range("Sec80CCC_1_Name").ClearContents
'Range("Sec80CCC_1_Name").Interior.Color = (&HD8D8D8)
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_1_Amount").Locked = True
'Range("Sec80CCC_1_Amount").ClearContents
'Range("Sec80CCC_1_Amount").Interior.Color = (&HD8D8D8)
'
'End If
'
'Application.EnableEvents = True
'End If
''-------------------------------------------------

'Ayush_26-27_C137
'If Not Application.Intersect(Sheet1.Range("IncD.Section80CCD1B_SE"), Sheet1.Range(Target.Address)) Is Nothing Then
'If Target.value > 0 Then
'Sheet1.Unprotect Password:=sPassword
'Application.EnableEvents = False
'
'
'Range("Sec80CCC_1b_Type").Locked = False
'Range("Sec80CCC_1b_Type").Interior.Color = (&HCCFFCC)
'
'Range("Sec80CCC_1b_Name").Locked = False
'Range("Sec80CCC_1b_Name").Interior.Color = (&HCCFFCC)
'
'Range("Sec80CCC_1b_Amount").Locked = False
'Range("Sec80CCC_1b_Amount").Interior.Color = (&HCCFFCC)
'
'Else
'
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_1b_Type").Locked = True
'Range("Sec80CCC_1b_Type").ClearContents
'Range("Sec80CCC_1b_Type").Interior.Color = (&HD8D8D8)
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_1b_Name").Locked = True
'Range("Sec80CCC_1b_Name").ClearContents
'Range("Sec80CCC_1b_Name").Interior.Color = (&HD8D8D8)
'
'Sheet1.Unprotect Password:=sPassword
'Range("Sec80CCC_1b_Amount").Locked = True
'Range("Sec80CCC_1b_Amount").ClearContents
'Range("Sec80CCC_1b_Amount").Interior.Color = (&HD8D8D8)
'
'End If
'
'Application.EnableEvents = True
'End If
'-------------------------------------------------




If Not Application.Intersect(Sheet1.Range("IncD.Section80GG"), Sheet1.Range(Target.Address)) Is Nothing Then


'Ayush_25/06**********************************
    If (Sheet1.Range("Sheet1.HRA").value > 0 And Target.value > 55000) Then
        fmsgbox ("* ""Deduction u/s 10(13A) & 80GG cannot be claimed for the same period""")
'        Sheet1.Unprotect Password:=sPassword
'        Sheet1.Range("IncD.Section80GG").MergeArea.ClearContents
    End If
'*********************************************


If Target.value >= 0 Then

Sheet1.Unprotect Password:=sPassword
Sheets("Income Details").Activate
Range("Sheet1.AckNum").MergeArea.Locked = False
'Range("Sheet1.AckNum").MergeArea.ClearContents
Sheet1.Range("Sheet1.AckNum").MergeArea.Interior.Color = (&HCCFFCC)

Else

Sheet1.Unprotect Password:=sPassword
Range("Sheet1.AckNum").MergeArea.Locked = True
'Range("Sheet1.AckNum").MergeArea.ClearContents
Range("Sheet1.AckNum").MergeArea.Interior.Color = (&HD8D8D8)
End If
'--------------------------

End If


'If Not Application.Intersect(Sheet1.Range("IncD.TypeOfHP"), Sheet1.Range(Target.Address)) Is Nothing Then


'If Sheet8.Range("Bacvalue").value = 2 Then
'If Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied" Then
'Sheet1.Range("IncD.IncomeFromHP_v").Formula = "=IF(TotAmt.24b>200000,200000,MIN(0,TotAmt.24b))"
'
'End If
'
'End If
'
'If Sheet8.Range("Bacvalue").value = 1 Then
'If Sheet1.Range("IncD.TypeOfHP").value = "Self Occupied" Then
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = (&HD8D8D8)
'Sheet1.Range("IncD.IncomeFromHP_v").value = 0
'Else
'Sheet1.Range("IncD.IncomeFromHP_v").MergeArea.Interior.Color = vbWhite
'Sheet1.Range("IncD.IncomeFromHP_v").Formula = "=IF(TotAmt.24b>200000,200000,MIN(0,TotAmt.24b))"
'
'
'End If
'End If
'End If


'=================================

ID_Nature_Employment_New  'Ayush_13/02/2026

Sheet1.Protect Password:=mdCommon.getmsgstate

endline:

Application.EnableEvents = True

Application.ScreenUpdating = True
End Sub
Private Sub Worksheet_Changeold1(ByVal Target As Range)
On Error GoTo endline

If (Target.name.name) = "sheet1.NatureofEmployment" Then
   Sheet1.Unprotect Password:=getmsgstate
 
      
If Range("sheet1.status").value = "I - INDIVIDUAL" And Sheet8.Range("bacValue").value = 2 Then

If Range("Sheet1.NatureofEmployment").value = "State Government" Or Range("Sheet1.NatureofEmployment").value = "Central Government" Or Range("Sheet1.NatureofEmployment").value = "Public Sector Undertaking" Then
MsgBox "0"
        Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)



Else
MsgBox "1"
        
        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HD8D8D8)
        Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
        Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
        
End If
End If
End If

Sheet1.Protect Password:=getmsgstate

If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" And Sheet8.Range("bacValue").value = 1 Then
 Sheet1.Range(Target.Address).value = "(Select)"
      fmsgboxsmall "* In case of Firm, New tax regime u/s 115BAC is not applicable."
 End If
CalculateTaxCheckFlag = False

Dim DOB_6 As Variant
Dim validateType As String
Dim result As Variant
validateType = ""
Application.ScreenUpdating = False
Application.EnableEvents = False

vbMessageCaption = "ITR-4"
DOB_6 = Sheet1.Range("sheet1.DOB").value
'Target.value = UCase(Target.value)


If Not Application.Intersect(Sheet1.Range("Sheet1.Amount"), Sheet1.Range(Target.Address)) Is Nothing Then

Sheet1.Unprotect Password:=getmsgstate

     If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
        If Sheet1.Range(Target.Address).value > 500000 Then
        fmsgbox ("*For Sec 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.5 lakhs")
        Sheet1.Range(Target.Address).value = ""
        Sheet1.Range(Target.Address).Select
        End If
     End If
     
'     If Range("Sheet1.NatureofEmployment").value <> "State Government" And Range("Sheet1.NatureofEmployment").value <> "Central Government" And Range("Sheet1.NatureofEmployment").value <> "Pensioners-SG" And Range("Sheet1.NatureofEmployment").value <> "Pensioners-CG" Then
      If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    
        If Sheet8.Range("bacValue").value = 1 Then
        If Sheet1.Range(Target.Address).value > 38400 Then
        fmsgbox ("*For Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee is selected than amount should not exceed 38400")
        Sheet1.Range(Target.Address).value = ""
        Sheet1.Range(Target.Address).Select
        End If
        End If
'        End If
End If
     
     If Range("Sheet1.NatureofEmployment").value <> "State Government" And Range("Sheet1.NatureofEmployment").value <> "Central Government" And Range("Sheet1.NatureofEmployment").value <> "Pensioners-SG" And Range("Sheet1.NatureofEmployment").value <> "Pensioners-CG" Then
      If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
    
        'If Sheet8.Range("bacValue").value = 1 Then
        If Sheet1.Range(Target.Address).value > 300000 Then
        fmsgbox ("*Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners is only Rs 3 lakh")
        Sheet1.Range(Target.Address).value = ""
        Sheet1.Range(Target.Address).Select
        
        End If
        End If
End If
     
     
Sheet1.Protect Password:=getmsgstate
End If


'Sec 10(10) ------ Natuure of income ------
'Dim rr, rrCount As Variant
'Dim tttNOI As Range
'
'
    If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Then
            If Trim(Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
                If Sheet1.Range(Target.Address).value > 2000000 Then
                    fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should not exeed Rs.20 Lakhs " & Chr(13))
                    Range(Target.Address).value = ""
                    Range(Target.Address).Select
                Else

                    Set tttNOI = Range("Sheet1.Nature")
                    rrCount = 0

                    For Each rr In tttNOI.Rows

                        If Trim(Range("I" & rr.row).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
                            rrCount = rrCount + Range("AO" & rr.row).value
                        End If

                    Next

                    If rrCount > 2000000 Then
                        fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should not exeed Rs.20 Lakhs " & Chr(13))
                        Range(Target.Address).value = ""
                        Range(Target.Address).Select
                    End If
                End If
            End If
    End If
    'dpk
    If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Then
            If Trim(Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
                If Sheet1.Range(Target.Address).value > 500000 Then
                    fmsgboxsmall ("*For Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette, amount should not exeed Rs.5 Lakhs " & Chr(13))
                    Range(Target.Address).value = ""
                    Range(Target.Address).Select
                Else

                    Set tttNOI = Range("Sheet1.Nature")
                    rrCount = 0

                    For Each rr In tttNOI.Rows

                        If Trim(Range("I" & rr.row).value) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
                            rrCount = rrCount + Range("AO" & rr.row).value
                        End If

                    Next

                    If rrCount > 500000 Then
                        fmsgboxsmall ("*For Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette, amount should not exeed Rs.20 Lakhs " & Chr(13))
                        Range(Target.Address).value = ""
                        Range(Target.Address).Select
                    End If
                End If
            End If
    End If
    'dpk
'
'    If Not Intersect(Range("Sheet1.Nature"), Range(Target.Address)) Is Nothing Then
'            If Trim(Sheet1.Range(Target.Address).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
'                If Sheet1.Range(Replace(Target.Address, "I", "AO")).value > 2000000 Then
'                    fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should not exeed Rs.20 Lakhs " & Chr(13))
'                    Sheet1.Range(Replace(Target.Address, "I", "AO")).value = ""
'                    Range(Target.Address).Select
'
'                Else
'
'                    Set tttNOI = Range("Sheet1.Nature")
'                    rrCount = 0
'
'                    For Each rr In tttNOI.Rows
'
'                        If Trim(Range("I" & rr.Row).value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
'                            rrCount = rrCount + Range("AO" & rr.Row).value
'                        End If
'
'                    Next
'
'                    If rrCount > 2000000 Then
'                        fmsgboxsmall ("*For Sec 10(10)-Death-cum-retirement gratuity received, amount should not exeed Rs.20 Lakhs " & Chr(13))
'                        Sheet1.Range(Replace(Target.Address, "I", "AO")).value = ""
'                        Range(Target.Address).Select
'                    End If
'                End If
'
'            End If
'    End If

'Sec 10(10AA) ------ Nature of income ------
'
'    If Not Intersect(Range("Sheet1.Amount"), Range(Target.Address)) Is Nothing Then
'            If Range("Sheet1.NatureofEmployment").value <> "State Government" And Range("Sheet1.NatureofEmployment").value <> "Central Government" Then
'                    If Sheet1.Range(Target.Address).value > 300000 Then
'                            If Trim(Sheet1.Range(Replace(Target.Address, "AO", "I")).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'                                     fmsgboxsmall ("The maximum limit of exemption u/s 10(10AA) cannot exceed Rs 3Lakhs for " & Range("Sheet1.NatureofEmployment").value & Chr(13) & "Please Enter Value Again")
'                                     'Sheet1.Range(Replace(Target.Address, "AO", "I")).value = "(Select)"
'                                     Range(Target.Address).value = ""
'                                     Range(Target.Address).Select
'                            End If
'                    Else
'
'                        Set tttNOI = Range("Sheet1.Nature")
'                        rrCount = 0
'
'                        For Each rr In tttNOI.Rows
'
'                            If Trim(Range("I" & rr.Row).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'                                rrCount = rrCount + Range("AO" & rr.Row).value
'                            End If
'
'                        Next
'
'                        If rrCount > 300000 Then
'                            fmsgboxsmall ("The maximum limit of exemption u/s 10(10AA) cannot exceed Rs 3Lakhs for " & Range("Sheet1.NatureofEmployment").value & Chr(13) & "Please Enter Value Again")
'                            Range(Target.Address).value = ""
'                            Range(Target.Address).Select
'                        End If
'                    End If
'
'            End If
'    End If
'
'
'    If Not Intersect(Range("Sheet1.Nature"), Range(Target.Address)) Is Nothing Then
'            If Range("Sheet1.NatureofEmployment").value <> "State Government" And Range("Sheet1.NatureofEmployment").value <> "Central Government" Then
'                    If Sheet1.Range(Target.Address).value = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'                            If Sheet1.Range(Replace(Target.Address, "I", "AO")).value > 300000 Then
'                                     fmsgboxsmall ("The maximum limit of exemption u/s 10(10AA) cannot exceed Rs 3Lakhs for " & Range("Sheet1.NatureofEmployment").value & Chr(13) & "Please Enter Value Again")
'                                     'Range(Target.Address).value = "(Select)"
'                                     'Sheet1.Range (Replace(Target.Address, "I", "AO"))
'                                     Sheet1.Range(Replace(Target.Address, "I", "AO")).value = ""
'                                     Range(Target.Address).Select
'                            Else
'
'                                    Set tttNOI = Range("Sheet1.Nature")
'                                    rrCount = 0
'
'                                    For Each rr In tttNOI.Rows
'
'                                        If Trim(Range("I" & rr.Row).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
'                                            rrCount = rrCount + Range("AO" & rr.Row).value
'                                        End If
'
'                                    Next
'
'                                    If rrCount > 300000 Then
'                                        fmsgboxsmall ("The maximum limit of exemption u/s 10(10AA) cannot exceed Rs 3Lakhs for " & Range("Sheet1.NatureofEmployment").value & Chr(13) & "Please Enter Value Again")
'                                        Sheet1.Range(Replace(Target.Address, "I", "AO")).value = ""
'                                        Range(Target.Address).Select
'                                    End If
'                            End If
'                    End If
'            End If
'    End If
'
'
'    If Not Intersect(Range("Sheet1.NatureofEmployment"), Range(Target.Address)) Is Nothing Then
'           Range("Sheet1.Nature").value = "(Select)"
'           Range("Sheet1.Amount").value = ""
'    End If

If Target.Validation.Type = 1 Then
          If Target.name.name = "IncD.Section80C" Or Target.name.name = "IncD.Section80CCC" Or _
                 Target.name.name = "IncD.Section80CCD_SE" Or Target.name.name = "IncD.Section80CCD1B_SE" Or _
                Target.name.name = "IncD.Section80CCD" Or Target.name.name = "IncD.Section80DD" Or _
                Target.name.name = "IncD.Section80DDB" Or Target.name.name = "IncD.Section80E" Or _
                Target.name.name = "IncD.Section80EE" Or Target.name.name = "IncD.Section80EEA" Or _
                Target.name.name = "IncD.Section80EEB" Or Target.name.name = "IncD.Section80GG" Then

             
             
            End If

End If


    If Not Intersect(Range("Sheet1.Description"), Range(Target.Address)) Is Nothing Then
            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
                fmsgboxsmall ("*Description should not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
               Range(Target.Address).value = ""
               Range(Target.Address).Select
            End If
            End If

        If Not Intersect(Range("Sheet1.Description1"), Range(Target.Address)) Is Nothing Then
            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
                fmsgboxsmall ("*Description should not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
               Range(Target.Address).value = ""
               Range(Target.Address).Select
            End If
            End If


If Target.Validation.Type = 1 Then
          If Target.name.name = "IncD.Section80C" Or Target.name.name = "IncD.Section80CCC" Or _
                 Target.name.name = "IncD.Section80CCD_SE" Or Target.name.name = "IncD.Section80CCD1B_SE" Or _
                Target.name.name = "IncD.Section80CCD" Or Target.name.name = "IncD.Section80DD" Or _
                Target.name.name = "IncD.Section80DDB" Or Target.name.name = "IncD.Section80E" Or _
                Target.name.name = "IncD.Section80EE" Or Target.name.name = "IncD.Section80EEA" Or _
                Target.name.name = "IncD.Section80EEB" Or Target.name.name = "IncD.Section80GG" Then

             
             
            End If

End If


    If Not Intersect(Range("Sheet1.Description"), Range(Target.Address)) Is Nothing Then
            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
                fmsgboxsmall ("*Description should not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
               Range(Target.Address).value = ""
               Range(Target.Address).Select
            End If
            End If

        If Not Intersect(Range("Sheet1.Description1"), Range(Target.Address)) Is Nothing Then
            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
                fmsgboxsmall ("*Description should not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
               Range(Target.Address).value = ""
               Range(Target.Address).Select
            End If
            End If

           'new
'            If Not Intersect(Range("Sheet1.Description1"), Range(Target.Address)) Is Nothing Then
'            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
'                fmsgboxsmall ("*Description should not contain Special characters < , > " & ",& " & "in Sheet Income Details" & Chr(13))
'               Range(Target.Address).value = ""
'               Range(Target.Address).Select
'            End If
'            End If


'        If Not Intersect(Range("IncD.Section80D_Calc"), Range(Target.Address)) Is Nothing Then
'                 Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
''         End If
'
'         If Not Intersect(Range("IncD.Section80D"), Range(Target.Address)) Is Nothing Then
'             Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'         End If
'
'         If Not Intersect(Range("IncD.Section80DB"), Range(Target.Address)) Is Nothing Then
'             Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'         End If
'
'         If Not Intersect(Range("IncD.Section80DC"), Range(Target.Address)) Is Nothing Then
'             Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'         End If



  If (Range("IncD.IncomeFromSal_iva").value) > (Range("sheet1.minvalue").value) Then
            fmsgboxsmall "* Deduction u/s 16(ia) can not exceed Rs. 50,000 or Net Salary whichever is lower in Sheet Income Details"
            Range("IncD.IncomeFromSal_iva").Formula = "=IF(bacValue=1,0,MIN(IncD.IncomeFromSal_net,50000))"
        End If


If (Range("IncD.LessDeduction57").value) > 15000 Or (Range("IncD.LessDeduction57").value) > (Range("temp57").value) / 3 Then
           fmsgboxsmall "* Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension or Rs. 15,000 in Sheet Income Details"
           Range("IncD.LessDeduction57").value = ""
           Range("IncD.LessDeduction57").Select
           End If


If Not Application.Intersect(Sheet1.Range("Sheet1.Nature"), Sheet1.Range(Target.Address)) Is Nothing Then


'        If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
'                MsgBox ("Nature of income should not contain Special characters < , >" & ",&" & Chr(13))
'               Range(Target.Address).value = ""
'        End If
        Dim ChangeAddress As Variant
        sPassword = getmsgstate
        ChangeAddress = Target.MergeArea.Address
        ChangeAddress = Replace(Target.MergeArea.Address, "I", "Z")
        ChangeAddress = Replace(ChangeAddress, "Y", "AN")

        If Sheet1.Range(Target.Address).value = "Any Other" Then
            Sheet1.Unprotect Password:=sPassword
                Sheet1.Range(ChangeAddress).value = ""
                Sheet1.Range(ChangeAddress).Interior.Color = (&HCCFFCC)  ', "AU"
                Sheet1.Range(ChangeAddress).Locked = False
            Sheet1.Protect Password:=sPassword
        Else
            Sheet1.Unprotect Password:=sPassword
                Sheet1.Range(ChangeAddress).value = "Not Applicable"
                Sheet1.Range(ChangeAddress).Interior.Color = (&HD8D8D8)
                Sheet1.Range(ChangeAddress).Locked = True
            Sheet1.Protect Password:=sPassword
    End If




    If Mid(Sheet1.Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then
    Sheet1.Unprotect Password:=sPassword
'                Sheet1.Range("IncD.Section80GG").MergeArea.value = ""
                Sheet1.Range("IncD.Section80GG").Interior.Color = (&HCCFFCC)  ', "AU"
                Sheet1.Range("IncD.Section80GG").MergeArea.Locked = False
                Sheet1.Protect Password:=sPassword
    Else
    Sheet1.Unprotect Password:=sPassword
                Sheet1.Range("IncD.Section80GG").MergeArea.value = ""
                Sheet1.Range("IncD.Section80GG").MergeArea.Interior.Color = (&HCCFFCC)  ', "AU"
                Sheet1.Range("IncD.Section80GG").MergeArea.Locked = False
                Sheet1.Protect Password:=sPassword
    End If

End If


If Not Application.Intersect(Sheet1.Range("Sheet1.Nature1"), Sheet1.Range(Target.Address)) Is Nothing Then


'        If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
'                MsgBox ("Nature of income should not contain Special characters < , >" & ",&" & Chr(13))
'               Range(Target.Address).value = ""
'        End If
      'If Sheet8.Range("BacValue").value = 2 Then
      If Range("temp5").value > 0 And Not isdropdownblank(Range(Target.Address).value) Then
          Sheet1.Unprotect Password:=getmsgstate
            Range("IncD.LessDeduction57").MergeArea.Locked = False
            Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
        Else
        Sheet1.Unprotect Password:=getmsgstate
        'DPKPnch
            Range("IncD.LessDeduction57").MergeArea.ClearContents
            Range("IncD.LessDeduction57").MergeArea.Locked = True
            Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate

    End If
    ' End If



        sPassword = getmsgstate
        ChangeAddress = Target.MergeArea.Address
        ChangeAddress = Replace(Target.MergeArea.Address, "I", "Z")
        ChangeAddress = Replace(ChangeAddress, "Y", "AN")

        If Sheet1.Range(Target.Address).value = "Any Other" Then
            Sheet1.Unprotect Password:=sPassword
                Sheet1.Range(ChangeAddress).value = ""
                Sheet1.Range(ChangeAddress).Interior.Color = (&HCCFFCC)  ', "AU"
                Sheet1.Range(ChangeAddress).Locked = False
            Sheet1.Protect Password:=sPassword
        Else
            Sheet1.Unprotect Password:=sPassword
                Sheet1.Range(ChangeAddress).value = "Not Applicable"
                Sheet1.Range(ChangeAddress).Interior.Color = (&HD8D8D8)
                Sheet1.Range(ChangeAddress).Locked = True
            Sheet1.Protect Password:=sPassword
    End If

MsgBox "new "
   If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
    If Mid(Sheet1.Range(Target.Address).value, 1, 1) = "F" Then
    MsgBox "Hell"
    fmsgboxsmall "*  Family pension should not allowed to be selected by a Firm(Other than LLP)in Sheet Income Details" ', vbOKOnly, "Error(s)!"
    Target.value = "(Select)"
    Else
    MsgBox "nohello"
    End If
    End If
    
     If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
    If Mid(Sheet1.Range(Target.Address).value, 1, 1) = "F" Then
    fmsgboxsmall "*  Family pension should not allowed to be selected by a HUF in Sheet Income Details" ', vbOKOnly, "Error(s)!"
    Target.value = "(Select)"
    End If
    End If
    

End If

Dim nOfCount2  As Long
                nOfCount2 = CountDropdown1(Trim(Target.value), Sheet1.Range("Sheet1.Nature1").Cells, Sheet1.Range("Sheet1.Nature1").count)
                If currCnt22 > 1 Or currCnt33 > 1 Or currCnt44 > 1 Or currCnt66 > 1 Or currCnt67 > 1 Or currCnt68 > 1 Or currCnt69 > 1 Or currCnt70 > 1 Then
                    fmsgboxsmall "* Drop down has already been selected in Sheet Income Details" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
                End If

'
'   If Not Application.Intersect(Sheet1.Range("Sheet1.Nature1"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'       If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
'                MsgBox ("Nature of income should not contain Special characters < , >" & ",&" & Chr(13))
'               Range(Target.Address).value = ""
'        End If
'
'
'
'        If Sheet1.Range(Target.Address).value = "Any Other" Then
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range(Replace(Target.Address, "I", "Z")).value = ""
'            Sheet1.Range(Replace(Target.Address, "I", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
'            Sheet1.Range(Replace(Target.Address, "I", "Z")).Locked = False
'
'            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
'
'            Sheet1.Protect Password:=getmsgstate
'        Else
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range(Replace(Target.Address, "I", "Z")).value = "NA"
'            Sheet1.Range(Replace(Target.Address, "I", "Z")).Interior.Color = (&HD8D8D8)
'            Sheet1.Range(Replace(Target.Address, "I", "Z")).Locked = True
'            Sheet1.Protect Password:=getmsgstate
'    End If
'    End If




Sheet1.Unprotect Password:=getmsgstate
If (Target.name.name) = "IncD.IntrstPayUs234A" Then
   CalculateTaxCheckFlag = True
End If
Sheet1.Protect Password:=getmsgstate
Sheet1.Unprotect Password:=getmsgstate
If (Target.name.name) = "IncD.IntrstPayUs234B" Then
   CalculateTaxCheckFlag = True
End If
Sheet1.Protect Password:=getmsgstate
Sheet1.Unprotect Password:=getmsgstate
If (Target.name.name) = "IncD.IntrstPayUs234C" Then
   CalculateTaxCheckFlag = True
End If
Sheet1.Protect Password:=getmsgstate
Sheet1.Unprotect Password:=getmsgstate
If (Target.name.name) = "IncD.IntrstPayUs234F" Then
   CalculateTaxCheckFlag = True
End If
Sheet1.Protect Password:=getmsgstate
'If (Target.name.name) = "sheet1.MiddleName" Then
'    Target.value = UCase(Target.value)
'End If
'
'If (Target.name.name) = "sheet1.FirstName" Then
'    Target.value = UCase(Target.value)
'End If
'
'If (Target.name.name) = "sheet1.SurNameOrOrgName" Then
'    Target.value = UCase(Target.value)
'End If

If (Target.name.name) = "sheet1.PANRepAssessee" Then
                Dim TargetPAN1 As String
                TargetPAN1 = ""
                TargetPAN1 = Target.name.name
    'Target.value = UCase(Target.value)
   If Range(TargetPAN1).value <> "" Then

    If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
    'Ayush_Pan
                  '  fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be 'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
                    fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(10) & Chr(13), vbOKCancel, mdCommon.vbMessageCaption
                    Sheet1.Range(TargetPAN1).Select
                    Sheet1.Range(TargetPAN1).value = ""
                    GoTo endline
   End If


   Sheet1.Range(TargetPAN1).value = UCase(Range(TargetPAN1).value)


   If Range(TargetPAN1).value = Sheet1.Range("sheet1.PAN").value Then
                    fmsgboxsmall "* PAN of the representative assessee and assessee (under personal information) cannot be same in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
                    Sheet1.Range(TargetPAN1).Select
                    Sheet1.Range(TargetPAN1).value = ""
                    GoTo endline
   End If
   End If



End If


If Not checkfieldSuperSpecialcharacter(Target.value) Then
        fmsgboxsmall ("* Please do not enter Special characters < , > in Sheet Income Details" & Chr(13))
        Target.value = ""
End If

If (Target.name.name) = "sheet1.AadhaarRepAssessee" Then

'If (Target.name.name) = "sheet1.Aadhaar" Then
                If Len(Sheet1.Range("sheet1.AadhaarRepAssessee").value) >= 12 Then
                    If chkNumeric(Sheet1.Range("sheet1.AadhaarRepAssessee").value) Then
                       If Sheet1.Range("sheet1.AadhaarRepAssessee").value <> "000000000000" Then
                         Sheet1.Range("sheet1.AadhaarRepAssessee").value = Sheet1.Range("sheet1.AadhaarRepAssessee").value
                       Else
                         fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                         Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                         Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                       End If


                       If Sheet1.Range("sheet1.AadhaarRepAssessee").value < 0 Then
                            fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                         Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                         Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                        End If


                       If Sheet1.Range("sheet1.AadhaarRepAssessee").value <> "111111111111" Then
                         Sheet1.Range("sheet1.AadhaarRepAssessee").value = Sheet1.Range("sheet1.AadhaarRepAssessee").value
                       Else
                         fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                         Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                         Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                       End If

                    Else
                      fmsgboxStatus "* Invalid Aadhaar Number of the representative"
                      Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                      Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                    End If
                Else
                    Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                End If
'            End If
        If Sheet1.Range("sheet1.AadhaarRepAssessee").value <> "" Then
        If Sheet1.Range("sheet1.AadhaarRepAssessee").value = Sheet1.Range("sheet1.Aadhaar").value Then
                    fmsgboxsmall "* Aadhaar no. entered in Part A -Gen and Aadhaar no. of the representative cannot be same in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
                    Sheet1.Range("sheet1.AadhaarRepAssessee").Select
                    Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
                    GoTo endline
        End If
        End If
End If



'If (Target.name.name) = "IncD.IncomeFromSal_i" Or (Target.name.name) = "IncD.IncomeFromSal_ii" Or (Target.name.name) = "IncD.IncomeFromSal_iii" Or (Target.name.name) = "IncD.IncomeFromSal_iv" Then
'Dim temp16aa As Variant
'            'temp16a = Sheet1.Range("IncD.IncomeFromSal_va").value
'                temp16aa = (Sheet1.Range("IncD.IncomeFromSal_i").value + Sheet1.Range("IncD.IncomeFromSal_ii").value + Sheet1.Range("IncD.IncomeFromSal_iii").value + Sheet1.Range("IncD.IncomeFromSal_iv").value)
'
'Sheet1.Range("IncD.IncomeFromSal_va").value = WorksheetFunction.Min(40000, temp16aa)
'
'End If







'HP income Validation on Data Entry
        If (Target.name.name) = "IncD.IncomeFromHP_v" Then

            If Range("IncD.TypeOfHP").value = "Let Out" Then
            If Range("LossCheckHP").value < -200000 Then
              fmsgboxsmall "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
            End If
            End If
        End If

' If (Target.name.name) = "IncD.IncomeFromSal_v" Then
'
'          If Range("IncD.IncomeFromSal_v").value > 10000 Then
'          MsgBox ("Deduction u/s 16(ii) and / or 16(iii) should not exceed 10,000." & Chr(13))
'           Target.value = ""
'          End If
'   End If

 If (Target.name.name) = "IncD.IncomeFromHP_v" Then
         Dim hpInterest, salIncome As Variant
           hpInterest = Range("IncD.IncomeFromHP_v").value
          If Range("IncD.TypeOfHP").value = "Self Occupied" Then
          mdIncomeDetails.CheckHousePropertyIncome (hpInterest)
          End If
   End If
'Natue Of Employment

'if IND and BACN
'If CGOV Or SGOV Or PSU Then
'green
'Else
'grayed out
'End If




If (Target.name.name) = "sheet1.NatureofEmployment" Then
   Sheet1.Unprotect Password:=getmsgstate
 
      
If Range("sheet1.status").value = "I - INDIVIDUAL" And Sheet8.Range("bacValue").value = 2 Then

If Range("Sheet1.NatureofEmployment").value = "State Government" Or Range("Sheet1.NatureofEmployment").value = "Central Government" Or Range("Sheet1.NatureofEmployment").value = "Public Sector Undertaking" Then

        Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)



Else

        Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HD8D8D8)
        Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
End If
End If

Sheet1.Protect Password:=getmsgstate



   If (Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Oth" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Pen" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Not") Then
'        Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_v").MergeArea.Locked = True
'        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HD8D8D8)

    Else

'        Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)

    End If

    If (Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Pen" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Not") Then
'        Range("IncD.IncomeFromSal_ivc").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = True
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HD8D8D8)

    Else

'        Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HCCFFCC)

    End If


'ElseIf (Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Gov" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Pub") Then
''        Range("IncD.Deduction16").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)
'
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HCCFFCC)
'End If
'
'If (Range("sheet1.NatureofEmployment").value = "(Select)") Then
'        Range("IncD.IncomeFromSal_v").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_v").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_v").MergeArea.Interior.Color = (&HCCFFCC)
'
'End If
'
'
'If (Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Pen" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Not") Then
'        Range("IncD.IncomeFromSal_ivc").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = True
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HD8D8D8)
'End If
'
'If (Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Gov" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Pub" Or Mid(Range("sheet1.NatureofEmployment").value, 1, 3) = "Oth") Then
'        'Range("IncD.Deduction16ic").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HCCFFCC)
'End If
'
'If (Range("sheet1.NatureofEmployment").value = "(Select)") Then
'        Range("IncD.IncomeFromSal_ivc").MergeArea.ClearContents
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = False
'        Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = (&HCCFFCC)
'
'End If


Sheet1.Protect Password:=getmsgstate
End If

If (Target.name.name) = "Sheet1.DepositAmount" Then
If Range("Sheet1.DepositAmount").value <= 10000000 Then
    fmsgboxsmall "* Amount entered cannot be less than or equal to Rs. 1 Crore" & Chr(13)
    Range("Sheet1.DepositAmount").value = ""
    End If
End If

If (Target.name.name) = "Sheet1.AggrigateAmount" Then
    If Range("Sheet1.AggrigateAmount").value <= 200000 Then
    fmsgboxsmall "* Amount entered cannot be less than or equal to Rs. 2 Lakhs" & Chr(13)
    Range("Sheet1.AggrigateAmount").value = ""
    End If
End If

If (Target.name.name) = "Sheet1.AggrigateAmount1" Then
    If Range("Sheet1.AggrigateAmount1").value <= 100000 Then
    fmsgboxsmall "* Amount entered cannot be less than or equal to Rs. 1 Lakh" & Chr(13)
    Range("Sheet1.AggrigateAmount1").value = ""
    End If
End If

If (Target.name.name) = "sheet1.SeventhProvisoFlag" Then
   Sheet1.Unprotect Password:=getmsgstate

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = False
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.DepositAmountFlag").MergeArea.value = "(Select)"

        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If



If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag").MergeArea.value = "(Select)"

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.value = "(Select)"

        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If



If (Target.name.name) = "Sheet1.DepositAmountFlag" Then
   Sheet1.Unprotect Password:=getmsgstate

If (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "N") Then
        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If


If (Target.name.name) = "Sheet1.AggrigateAmountFlag" Then
   Sheet1.Unprotect Password:=getmsgstate

If (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If


If (Target.name.name) = "Sheet1.AggrigateAmountFlag1" Then
   Sheet1.Unprotect Password:=getmsgstate

If (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If

'new insertion

If (Target.name.name) = "IncD.IncomeFromSal_ii" Then


        If (Range("IncD.IncomeFromSal_v").value) > 5000 Or (Range("IncD.IncomeFromSal_v").value) > (Range("IncD.IncomeFromSal_ii").value) / 5 Then
            fmsgboxsmall "* Deduction of Entertainment allowance u/s 16(ii) can not exceed 1/5th of Salary as per section 17(1) or Rs. 5,000 whichever is lower in Sheet Income Details"
            Range("IncD.IncomeFromSal_v").value = ""
            Range("IncD.IncomeFromSal_v").Select
        End If
End If

'If (Target.name.name) = "IncD.IncomeFromSal_v" Then
'       If (Range("IncD.IncomeFromSal_v").value) > 5000 Or (Range("IncD.IncomeFromSal_v").value) > (Range("IncD.IncomeFromSal_ii").value) / 5 Then
'           fmsgboxsmall "* Deduction of Entertainment allowance u/s 16(ii) can not exceed 1/5th of Salary as per section 17(1) or Rs. 5,000 whichever is lower in Sheet Income Details"
'           Range("IncD.IncomeFromSal_v").value = ""
'           Range("IncD.IncomeFromSal_v").Select
'      End If
'End If

If (Target.name.name) = "IncD.IncomeFromSal_ivc" Then
        If (Range("IncD.IncomeFromSal_ivc").value) > 5000 Then
            fmsgboxsmall "* Deduction of Professional tax u/s 16(iii) cannot exceed Rs. 5,000/ in Sheet Income Details"
            Range("IncD.IncomeFromSal_ivc").value = ""
            Range("IncD.IncomeFromSal_ivc").Select
        End If
End If



' If (Target.name.name) = "IncD.IncomeFromSal_va" Then
'             If (Range("IncD.IncomeFromSal_va").value) > 40000 Then
'            MsgBox "Deduction16(ia) should not exceed 40,000."
'            Range("IncD.IncomeFromSal_va").value = ""
'             End If
'
'             Dim temp16a, temp16aa As Variant
'            temp16a = Sheet1.Range("IncD.IncomeFromSal_va").value
'            temp16aa = (Sheet1.Range("IncD.IncomeFromSal_i").value + Sheet1.Range("IncD.IncomeFromSal_ii").value + Sheet1.Range("IncD.IncomeFromSal_iii").value + Sheet1.Range("IncD.IncomeFromSal_iv").value)
'
'
'        If temp16a <> "" Then
'        If Not Sheet1.Range("IncD.IncomeFromSal_va").value < temp16aa Then
'               MsgBox "Maximum deduction u/s 16(ia) is '40000' or sum of B1 [(i)+(ii)+(iii)+(iv)] in Part B whichever is lower."
'               Range("IncD.IncomeFromSal_va").value = ""
''               If temp16aa > 40000 Then
''               temp16aa = 40000
''               Else
''
''               temp16aa = temp16aa
''               End If
''
''               If temp16aa < 0 Then
''               temp16aa = 0
''               Else
''               temp16aa = temp16aa
''               End If
'
'        End If
'        End If
'
'End If

'end




If Target.Validation.Type = 3 Then

'    If Not Intersect(Range("SELECT80D"), Range(Target.Address)) Is Nothing Then
'        Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'    End If
'
'    If Not Intersect(Range("SELECT80DB"), Range(Target.Address)) Is Nothing Then
'        Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'    End If
'
'    If Not Intersect(Range("SELECT80DC"), Range(Target.Address)) Is Nothing Then
'        Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'    End If






    If (Target.name.name) = "sheet1.RepAssessee" Then
        '------------------Un Protect--------------------
                sPassword = mdCommon.getmsgstate
                Sheet1.Unprotect Password:=sPassword
        '------------------Un Protect Ends--------------------
        If Range("sheet1.RepAssessee").value = "No" Then
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("sheet1.NameRepAssessee").value = ""
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Range("sheet1.CapacityRep").value = ""
            Sheet1.Range("sheet1.CapacityRep").MergeArea.Locked = True
            Sheet1.Range("sheet1.CapacityRep").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.AddressRepAssessee").value = ""
            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.PANRepAssessee").value = ""
            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.AadhaarRepAssessee").value = ""
            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
        Else
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.CapacityRep").value = "(Select)"
            Sheet1.Range("sheet1.CapacityRep").MergeArea.Locked = False
            Sheet1.Range("sheet1.CapacityRep").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.AddressRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.PANRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)
        End If

        '----------------Lock Password-------------------START---
                Sheet17.Protect Password:=sPassword
                Sheet1.Protect Password:=sPassword
        '----------------Lock Password-------------------END-----

    End If



'If Target.Validation.Type = 3 Then
'
'    If (Target.name.name) = "sheet1.RepAssessee" Then
'        '------------------Un Protect--------------------
'                sPassword = mdCommon.getmsgstate
'                Sheet1.Unprotect Password:=sPassword
'        '------------------Un Protect Ends--------------------
'
'    If Range("sheet1.RepAssessee").value = "No" Then
'            Sheet1.Unprotect Password:=sPassword
'            Sheet1.Range("sheet1.CapacityRep").value = ""
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Locked = True
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Interior.Color = (&HD8D8D8)
'Else
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Locked = False
'            Sheet1.Range("sheet1.CapacityRep").MergeArea.Interior.Color = (&HCCFFCC)
'    End If
'    End If
'
'      '----------------Lock Password-------------------START---
'                Sheet17.Protect Password:=sPassword
'                Sheet1.Protect Password:=sPassword
'        '----------------Lock Password-------------------END-----
'    End If





'
'    If (Target.name.name) = "sheet1.PortugeseCC5A" Then
'        '------------------Un Protect--------------------
'                sPassword = mdCommon.getmsgstate
'                Sheet1.Unprotect Password:=sPassword
'                Sheet17.Unprotect Password:=sPassword
'        '------------------Un Protect Ends--------------------
'
'        If Range("sheet1.PortugeseCC5A").value = "No" Then
'            mdCommon.Lock_TDS
'            Sheet1.Unprotect Password:=sPassword
'            Sheet1.Range("sheet1.PANOFSPOUSE").value = ""
'            Sheet1.Range("sheet1.PANOFSPOUSE").MergeArea.Locked = True
'            Sheet1.Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HD8D8D8)
'
''            Sheet4.Range("TDSoth.TD_PAN").value = ""
''            Sheet4.Range("TDSoth.TD_PAN").Locked = True
''            Sheet4.Range("TDSoth.TD_PAN").Interior.Color = (&HD8D8D8)
''
''            Sheet4.Range("TDSoth.TDS_PAN").value = ""
''            Sheet4.Range("TDSoth.TDS_PAN").Locked = True
''            Sheet4.Range("TDSoth.TDS_PAN").Interior.Color = (&HD8D8D8)
''
''             Sheet4.Range("TDSoth2.TD_PAN").value = ""
''            Sheet4.Range("TDSoth2.TD_PAN").Locked = True
''            Sheet4.Range("TDSoth2.TD_PAN").Interior.Color = (&HD8D8D8)
''
''            Sheet4.Range("TDSoth2.TDS_PAN").value = ""
''            Sheet4.Range("TDSoth2.TDS_PAN").Locked = True
''            Sheet4.Range("TDSoth2.TDS_PAN").Interior.Color = (&HD8D8D8)
'        Else
'            mdCommon.UnLock_TDS
'            Sheet1.Unprotect Password:=sPassword
'            Sheet1.Range("sheet1.PANOFSPOUSE").MergeArea.Locked = False
'            Sheet1.Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HCCFFCC)
'            Sheet1.Protect Password:=sPassword
''            Sheet4.Range("TDSoth.TD_PAN").Locked = False
''            Sheet4.Range("TDSoth.TD_PAN").Interior.Color = (&HCCFFCC)
''            Sheet4.Range("TDSoth.TDS_PAN").Locked = False
''            Sheet4.Range("TDSoth.TDS_PAN").Interior.Color = (&HCCFFCC)
''            Sheet4.Range("TDSoth2.TD_PAN").Locked = False
''            Sheet4.Range("TDSoth2.TD_PAN").Interior.Color = (&HCCFFCC)
''            Sheet4.Range("TDSoth2.TDS_PAN").Locked = False
''            Sheet4.Range("TDSoth2.TDS_PAN").Interior.Color = (&HCCFFCC)
'        End If
'
'        '----------------Lock Password-------------------START---
'                Sheet17.Protect Password:=sPassword
'                Sheet1.Protect Password:=sPassword
'        '----------------Lock Password-------------------END-----
'    End If

'PortugeseCC5A
'    If (Target.name.name) = "sheet1.PortugeseCC5A" Then
'        Dim TargetCC5A As String
'        TargetCC5A = ""
'        TargetCC5A = Range("sheet1.PortugeseCC5A").value
'        If UCase(TargetCC5A) = "YES" Then
'            MsgBox "You have selected to be governed by Sec 5A. Please enter only your share of ""Income from Business or profession"", ""Income from House Property"" and ""Income from Other Sources"" Refer to instructions (A202) for further clarification.", vbOKOnly, mdCommon.vbMessageCaption
'        End If
'    End If

    'Status&Gender Validation: on data entry
    
    If (Target.name.name) = "sheet1.Status" Then
    
    LockUnlock80D

    If Sheet1.Range("temp5").value > 0 Then
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "F" Then
    fmsgboxsmall "*  Family pension should not allowed to be selected by a Firm(Other than LLP)in Sheet Income Details" ', vbOKOnly, "Error(s)!"
    Sheet1.Range("sheet1.Status").value = "(Select)"
    End If
    End If
    
       If Sheet1.Range("temp5").value > 0 Then
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
    fmsgboxsmall "*  Family pension should not allowed to be selected by a HUF in Sheet Income Details" ', vbOKOnly, "Error(s)!"
    Sheet1.Range("sheet1.Status").value = "(Select)"
    End If
    End If


            If Range("sheet1.status").value = "I - INDIVIDUAL" Or Range("sheet1.status").value = "F - FIRM(Other than LLP)" Then
            If Trim(Range("sheet1.DOB").value) = "01/01/0001" Then
            fmsgboxsmall "*Status in Sheet Income Details Please enter valid DOB." ', vbOKOnly, mdCommon.vbMessageCaption
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").value = ""
            End If
            End If


        '-----------------Un Protect--------------------
            sPassword = mdCommon.getmsgstate
            Sheet1.Unprotect Password:=sPassword
            Sheet4.Unprotect Password:=sPassword
            Sheet3.Unprotect Password:=sPassword
        '------------------Un Protect Ends--------------------





            LockNatureofEmployment
            LockUnlock80TTB
            LockUnlockQVIA
            LockUnlock80DDB
            

            LockUnlockAadhaar
            If Mid(Trim(Range("sheet1.Status")), 1, 1) = "F" Or Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
                LOCKPI
                LOCKTDS1
            If Mid(Trim(Range("sheet1.Status")), 1, 1) = "H" Then
                LOCKEE
            Else
            
            End If
                
'                LOCKAL
'                LOCKAL_1
'                LOCKAL_2



            Else
                UNLOCKPI
                UNLOCKTDS1
                
                'UNLOCKAL
'                UNLOCKAL
'                UNLOCKAL_1
'                UNLOCKAL_2
'                Sheet3.Range("BP_E3_2").Locked = True
'                Sheet3.Range("BP_E3_2").ClearContents
'                Sheet3.Range("BP_E3_2").Interior.Color = (&HD8D8D8)
            End If

            Sheet1.Unprotect Password:=sPassword
            Sheet4.Unprotect Password:=sPassword
            
            

            If Mid(Trim(Range("sheet1.Status")), 1, 1) = "F" Then

                Sheet3.Range("BP_E3_2").Locked = False
                Sheet3.Range("BP_E3_2").Interior.Color = (&HCCFFCC)

                Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = True
                Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.ClearContents
                Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
                Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
                Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.DepositAmount").MergeArea.ClearContents
                Range("Sheet1.DepositAmount").MergeArea.Locked = True
                Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
                Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
                Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
                Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
                Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
                Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
                Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
                Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
                Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)

                Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
                Sheet1.Range("IncD.Section80C").MergeArea.ClearContents
                Sheet1.Range("IncD.Section80C").MergeArea.Interior.Color = (&HD8D8D8)
                
                

            Else
                Sheet1.Range("IncD.Section80C").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80C").ClearContents
                Sheet1.Range("IncD.Section80C").MergeArea.Interior.Color = (&HCCFFCC)

                Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = False
                Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.value = "(Select)"
                Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = (&HCCFFCC)

                Range("Sheet1.DepositAmountFlag").MergeArea.Locked = False
                Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
                Range("Sheet1.DepositAmountFlag").MergeArea.value = "(Select)"

                Range("Sheet1.DepositAmount").MergeArea.ClearContents
                Range("Sheet1.DepositAmount").MergeArea.Locked = False
                Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
'                Range("Sheet1.DepositAmount").MergeArea.value = "(Select)"

                Range("Sheet1.AggrigateAmountFlag").MergeArea.value = "(Select)"
                Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
                Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)

                Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
                Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
                Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)

                Range("Sheet1.AggrigateAmountFlag1").MergeArea.value = "(Select)"
                Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
                Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)

                Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
                Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
                Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
                
                

                Sheet3.Range("BP_E3_2").Locked = True
                Sheet3.Range("BP_E3_2").ClearContents
                Sheet3.Range("BP_E3_2").Interior.Color = (&HD8D8D8)

            End If






          If (Mid(Range("sheet1.Status").value, 1, 1) = "H") Then

'                Range("sheet1.Gender1").value = "Not applicable"
'                Range("sheet1.Gender1").MergeArea.Locked = True
'                Range("sheet1.Gender1").MergeArea.Interior.Color = (&HD8D8D8)

                Sheet5.Range("Ver.AssesseeVerName").value = ""
                Sheet5.Range("Ver.PAN").value = ""

'                Range("sheet1.PortugeseCC5A").value = "No"
'                Range("sheet1.PortugeseCC5A").MergeArea.Locked = True
'                Range("sheet1.PortugeseCC5A").MergeArea.Interior.Color = (&HD8D8D8)
'
'                'Range("sheet1.aadharFlag").value = "NA"
'
'                'Range("sheet1.ResidentialStatus1").value = "(Select)"
'
'                Range("sheet1.PANOFSPOUSE").value = ""
'                Range("sheet1.PANOFSPOUSE").MergeArea.Locked = True
'                Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HD8D8D8)
                mdCommon.Lock_TDS
          ElseIf (Mid(Range("sheet1.Status").value, 1, 1) = "F") Then

'                Range("sheet1.Gender1").value = "Not applicable"
'                Range("sheet1.Gender1").MergeArea.Locked = True
'                Range("sheet1.Gender1").MergeArea.Interior.Color = (&HD8D8D8)

                Sheet5.Range("Ver.AssesseeVerName").value = ""
                Sheet5.Range("Ver.PAN").value = ""

'                Range("sheet1.PortugeseCC5A").value = "(Select)"
'                Range("sheet1.PortugeseCC5A").MergeArea.Locked = True
'                Range("sheet1.PortugeseCC5A").MergeArea.Interior.Color = (&HD8D8D8)
'
'                'Range("sheet1.aadharFlag").value = "NA"
'
'                'Range("sheet1.ResidentialStatus1").value = "(Select)"
'
'                Range("sheet1.PANOFSPOUSE").value = ""
'                Range("sheet1.PANOFSPOUSE").MergeArea.Locked = True
'                Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HD8D8D8)
                mdCommon.Lock_TDS
            Else
                'Range("sheet1.aadharFlag").value = "(Select)"
'                 Sheet1.Unprotect Password:=sPassword
'                Range("sheet1.PortugeseCC5A").value = "No"
'                Range("sheet1.PortugeseCC5A").MergeArea.Locked = False
'                Range("sheet1.PortugeseCC5A").MergeArea.Interior.Color = (&HCCFFCC)
'
'                'Range("sheet1.ResidentialStatus1").value = "(Select)"
'
'                Range("sheet1.PANOFSPOUSE").MergeArea.Locked = False
'                Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HCCFFCC)

'                Range("sheet1.TaxStatus").MergeArea.Locked = False
'                Range("sheet1.TaxStatus").MergeArea.Interior.Color = (&HCCFFCC)

'                Range("sheet1.Gender1").MergeArea.value = "(Select)"
'                Range("sheet1.Gender1").MergeArea.Locked = False
'                Range("sheet1.Gender1").MergeArea.Interior.Color = (&HCCFFCC)



'                If Range("sheet1.PortugeseCC5A").value = "No" Then
'                        mdCommon.Lock_TDS
'                        Sheet1.Unprotect Password:=sPassword
'                        Range("sheet1.PANOFSPOUSE").value = ""
'                        Range("sheet1.PANOFSPOUSE").MergeArea.Locked = True
'                        Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HD8D8D8)
'
''                        Sheet4.Range("TDSoth.TD_PAN").value = ""
''                        Sheet4.Range("TDSoth.TD_PAN").Locked = True
''                        Sheet4.Range("TDSoth.TD_PAN").Interior.Color = (&HD8D8D8)
''
''                        Sheet4.Range("TDSoth.TDS_PAN").value = ""
''                        Sheet4.Range("TDSoth.TDS_PAN").Locked = True
''                        Sheet4.Range("TDSoth.TDS_PAN").Interior.Color = (&HD8D8D8)
''
''                        Sheet4.Range("TDSoth2.TD_PAN").value = ""
''                        Sheet4.Range("TDSoth2.TD_PAN").Locked = True
''                        Sheet4.Range("TDSoth2.TD_PAN").Interior.Color = (&HD8D8D8)
''
''                        Sheet4.Range("TDSoth2.TDS_PAN").value = ""
''                        Sheet4.Range("TDSoth2.TDS_PAN").Locked = True
''                        Sheet4.Range("TDSoth2.TDS_PAN").Interior.Color = (&HD8D8D8)
'                Else
'                    mdCommon.UnLock_TDS
'                    Sheet1.Unprotect Password:=sPassword
'                    Range("sheet1.PANOFSPOUSE").MergeArea.Locked = False
'                    Range("sheet1.PANOFSPOUSE").MergeArea.Interior.Color = (&HCCFFCC)
'                End If

                If Mid(UCase(Trim(Sheet1.Range("sheet1.Status").value)), 1, 1) = "I" And Mid(UCase(Trim(Sheet1.Range("sheet1.RepAssessee").value)), 1, 1) = "N" Then
                Sheet5.Range("Ver.PAN").value = UCase(Sheet1.Range("sheet1.PAN").value)
            End If
        End If

        '----------------Lock Password-------------------START---
            Sheet1.Protect Password:=sPassword
            Sheet4.Protect Password:=sPassword
            Sheet3.Protect Password:=sPassword
        '----------------Lock Password-------------------END-----





    End If

     'StateCode & Country Validation: ON data entry
        If (Target.name.name) = "sheet1.StateCode1" And UCase(Range("sheet1.StateCode1").value) <> "99-FOREIGN" Then
            'Range("sheet1.CountryCode").value = ""
            Range("sheet1.CountryCode").value = "91-INDIA"
            'Range("sheet1.PinCode").value = ""
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.mobileCountryCode").value = "91"
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = True
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HD8D8D8)

'                   Sheet1.Range("sheet1.mobileCountryCode2").value = "91"
'                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = True
'                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HD8D8D8)

                   Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
   'sudha                'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

                   Sheet1.Range("HASZIP").MergeArea.value = ""
                   Sheet1.Range("HASZIP").MergeArea.Locked = True
                   Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
        End If

        If (Target.name.name) = "sheet1.StateCode1" And UCase(Range("sheet1.StateCode1").value) = "99-FOREIGN" Then
            Range("sheet1.CountryCode").value = ""
            'Range("sheet1.PinCode").value = "999999"
            Range("sheet1.CountryCode").value = "(Select)"
            Range("sheet1.CountryCode").Select

            Sheet1.Unprotect Password:=getmsgstate

'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)

'                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode2").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)

               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = False
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate

        End If

        If (Target.name.name) = "sheet1.StateCode1" And UCase(Range("sheet1.StateCode1").value) = "(SELECT)" Then
            Range("sheet1.CountryCode").value = ""
            'Range("sheet1.PinCode").value = ""
            Range("sheet1.CountryCode").value = "(Select)"
            Range("sheet1.CountryCode").Select
                Sheet1.Unprotect Password:=getmsgstate
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)

'                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode2").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)


                   Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
         'sudha          'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

                   Sheet1.Range("HASZIP").MergeArea.value = ""
                   Sheet1.Range("HASZIP").MergeArea.Locked = True
                   Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
        End If


        'If (Target.name.name) = "sheet1.CountryCode" And UCase(Range("sheet1.CountryCode").value) <> "91-INDIA" Then
        '    Range("sheet1.StateCode1").value = "99-FOREIGN"
        '    Range("sheet1.PinCode").value = "999999"
        'End If

        'If (Target.name.name) = "sheet1.CountryCode" And UCase(Range("sheet1.CountryCode").value) = "91-INDIA" Then
        '    Range("sheet1.StateCode1").value = "(Select)"
        '    Range("sheet1.PinCode").value = ""
        '    Range("sheet1.StateCode1").Select
        'End If

        'If (Target.name.name) = "sheet1.CountryCode" And UCase(Range("sheet1.CountryCode").value) = "(SELECT)" Then
        '    Range("sheet1.StateCode1").value = "(Select)"
        '    Range("sheet1.PinCode").value = ""
        '    Range("sheet1.StateCode1").Select
        'End If

        'Country Validation: On Data Entry
        If (Target.name.name) = "sheet1.CountryCode" Then
          If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
                Sheet1.Range("sheet1.StateCode1") = "(Select)"
    'sudha            'Sheet1.Range("sheet1.PinCode") = ""
                If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-FOREIGN" Then
                    MsgBox "* Country Should not be India for Selected State in Sheet Income Details" & Chr(13)
                    Sheet1.Range("sheet1.CountryCode").value = "(Select)"
                    Sheet1.Range("sheet1.CountryCode").Select
                    'Sheet1.Range("sheet1.StateCode1").value = ""
                End If

                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.mobileCountryCode").value = "91"
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = True
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HD8D8D8)

'                   Sheet1.Range("sheet1.mobileCountryCode2").value = "91"
'                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = True
'                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HD8D8D8)

                   Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
         'sudha          'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

                   Sheet1.Range("HASZIP").MergeArea.value = ""
                   Sheet1.Range("HASZIP").MergeArea.Locked = True
                   Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
          ElseIf Sheet1.Range("sheet1.CountryCode").value = "(Select)" Then
                Sheet1.Range("sheet1.StateCode1") = "(Select)"

                Sheet1.Unprotect Password:=getmsgstate


'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)

'                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode2").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)

                   Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
       'sudha            'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

                   Sheet1.Range("HASZIP").MergeArea.value = ""
                   Sheet1.Range("HASZIP").MergeArea.Locked = True
                   Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
          ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
                If UCase(Sheet1.Range("sheet1.StateCode1").value) <> "99-FOREIGN" And Sheet1.Range("sheet1.StateCode1").value <> "(Select)" Then
                    MsgBox "* Country Should be India for Selected State in Sheet Income Details" & Chr(13)
                    Sheet1.Range("sheet1.CountryCode").value = ""
                End If
                Sheet1.Unprotect Password:=getmsgstate

'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode").MergeArea.Interior.Color = (&HCCFFCC)

'                   Sheet1.Range("sheet1.mobileCountryCode2").Locked = False
                   Sheet1.Range("sheet1.mobileCountryCode2").value = ""
'                   Sheet1.Range("sheet1.mobileCountryCode2").Interior.Color = (&HCCFFCC)

                   Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
           'sudha        'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)

                   Sheet1.Range("HASZIP").MergeArea.value = ""
                   Sheet1.Range("HASZIP").MergeArea.Locked = False
                   Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)

                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
                Sheet1.Protect Password:=getmsgstate
          End If
        End If

       If (Target.name.name) = "HASZIP" Then
'            Target.value = UCase(Target.value)
            If Mid(Trim(Sheet1.Range("HASZIP").value), 1, 1) = "Y" Then
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = "XXXXXX"
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
            Else
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
                Sheet1.Protect Password:=getmsgstate
            End If
        End If

        If (Target.name.name = "sheet1.ResidentialStatus1") Then
           LockUnlock80TTB
            LockUnlock80DDB
            LockUnlockQVIA

            'LockUnlockIBANSELECT
            If (Mid(Target.value, 1, 2) = "NR") Then
                Lock44AD


            Else
                Unlock44AD


            End If
        End If




  If (Target.name.name) = "SELECT80DB" Then
      sPassword = mdCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword


      If Mid(Range("SELECT80DB").value, 1, 1) = "1" Then
'       If calculateAge(DOB_6) <= 59 Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
'          MsgBox "Please select a valid option from the dropdown of Part B Sec.80D (Medical Expenditure) under Chapter VIA."
'          Range("SELECT80DB").value = "(Select)"
'          Range("SELECT80DB").Select
'        End If
        If Mid(Range("sheet1.Status").value, 1, 1) <> "I" And Mid(Range("sheet1.Status").value, 1, 1) <> "(" Then
          fmsgboxsmall "* Please select a valid option from the dropdown of Part B Sec.80D (Medical Expenditure) under Chapter VIA."
          Range("SELECT80DB").value = "(Select)"
          Range("SELECT80DB").Select
        End If

'        If Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" And Mid(Range("SELECT80DB").value, 1, 1) = "2" Or Mid(Range("SELECT80DB").value, 1, 1) = "7" Then
'        MsgBox "Please select a valid option from the dropdown of Part B Sec.80D (Medical Expenditure) under Chapter VIA."
'          Range("SELECT80DB").value = "(Select)"
'         End If

        If Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
          fmsgboxsmall "* Please select a valid option from the dropdown of Part B Sec.80D (Medical Expenditure) under Chapter VIA."
          Range("SELECT80DB").value = "(Select)"
          Range("SELECT80DB").Select
        End If

      End If

      Sheet1.Protect Password:=sPassword

    End If


    If (Target.name.name) = "SELECT80D" Then
      sPassword = mdCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword


      If Mid(Range("SELECT80D").value, 1, 1) = "2" Or Mid(Range("SELECT80D").value, 1, 1) = "7" Then
       If Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
          fmsgboxsmall "* Please select a valid option from the dropdown of Part A Sec.80D under Chapter VIA."
          Range("SELECT80D").value = "(Select)"
          Range("SELECT80D").Select
        End If

      End If


      Sheet1.Protect Password:=sPassword

    End If



        If (Target.name.name) = "IncD.TypeOfHP" Then
        '------------------Un Protect--------------------
                    sPassword = mdCommon.getmsgstate
                    Sheet1.Unprotect Password:=sPassword
            '------------------Un Protect Ends--------------------





        If Mid(Range("IncD.TypeOfHP").value, 1, 1) = "S" Then


         Dim hpInterest1, salIncome1 As Variant
           hpInterest1 = Range("IncD.IncomeFromHP_v").value
          If Range("IncD.TypeOfHP").value = "Self Occupied" Then
          mdIncomeDetails.CheckHousePropertyIncome (hpInterest1)
          End If



'                Range("IncD.IncomeFromHP_ia").value = ""
'                Range("IncD.IncomeFromHP_ia").MergeArea.Locked = True
'                Range("IncD.IncomeFromHP_ia").MergeArea.Interior.Color = (&HD8D8D8)
'
'                Range("IncD.IncomeFromHP_ib").value = ""
'                Range("IncD.IncomeFromHP_ib").MergeArea.Locked = True
'                Range("IncD.IncomeFromHP_ib").MergeArea.Interior.Color = (&HD8D8D8)

                Range("IncD.IncomeFromHP_i").value = ""
                Range("IncD.IncomeFromHP_i").MergeArea.Locked = True
                Range("IncD.IncomeFromHP_i").MergeArea.Interior.Color = (&HD8D8D8)

                Range("IncD.IncomeFromHP_ii").value = ""
                Range("IncD.IncomeFromHP_ii").MergeArea.Locked = True
                Range("IncD.IncomeFromHP_ii").MergeArea.Interior.Color = (&HD8D8D8)

        Else
'                Range("IncD.IncomeFromHP_ia").MergeArea.Locked = False
'                Range("IncD.IncomeFromHP_ia").MergeArea.Interior.Color = (&HCCFFCC)
'
'                 Range("IncD.IncomeFromHP_ib").MergeArea.Locked = False
'                Range("IncD.IncomeFromHP_ib").MergeArea.Interior.Color = (&HCCFFCC)

                Range("IncD.IncomeFromHP_i").MergeArea.Locked = False
                Range("IncD.IncomeFromHP_i").MergeArea.Interior.Color = (&HCCFFCC)

                Range("IncD.IncomeFromHP_ii").MergeArea.Locked = False
                Range("IncD.IncomeFromHP_ii").MergeArea.Interior.Color = (&HCCFFCC)

         End If

         '----------------Lock Password-------------------START---
            Sheet1.Protect Password:=sPassword
        '----------------Lock Password-------------------END-----

         End If





        If (Target.name.name) = "sheet1.ReturnFileSec" Then
   '  If Not Application.Intersect(Range("sheet1.ReturnFileSec"), Range(Target.Address)) Is Nothing Then
            '------------------Un Protect--------------------
                    sPassword = mdCommon.getmsgstate
                    Sheet1.Unprotect Password:=sPassword
            '------------------Un Protect Ends--------------------

'If Range("sheet1.ReturnFileSec2").value <> "(Select)" And Range("sheet1.ReturnFileSec2").value <> "" Then
'      MsgBox "Either ""Filed u/s"" or ""Filed in response to notice u/s"" must have to be selected in tab: Part A General Information", vbOKOnly, "ITR4"
'      Range("sheet1.ReturnFileSec").value = "(Select)"
'      Range("sheet1.ReturnFileSec2").value = "(Select)"
'End If
'AYush_27/02/2026
   

    
    
If Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
            If Sheet8.Range("bacValue").value = 1 Then
                 'fmsgbox "* Please select ""Have you opted  for new tax regime u/s 115BAC and filed Form 10IE in AY 2021-22 ? ""  as n since you have filed return later." dpk1201
                 fmsgbox "* Please select ""Have you opted  for new tax regime u/s 115BAC and filed Form 10IE in AY 2022-23 ? ""  as n since you have filed return later."
                 Range("sheet1.ReturnFileSec").value = ""
            End If
        End If


        If Range("sheet1.ReturnFileSec").value = "(Select)" And Range("sheet1.ReturnFileSec2").value = "(Select)" Then
           Range("sheet1.ReturnFileSec1").value = "(Select)"
        ElseIf Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
           Range("sheet1.ReturnFileSec1").value = "11 - 139(1)-On or before due date"
        ElseIf Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
           Range("sheet1.ReturnFileSec1").value = "12 - 139(4)-After due date"
        ElseIf Range("sheet1.ReturnFileSec").value = "142(1)" Then
           Range("sheet1.ReturnFileSec1").value = "13 - 142(1)"
        ElseIf Range("sheet1.ReturnFileSec").value = "148" Then
           Range("sheet1.ReturnFileSec1").value = "14 - 148"
        ElseIf Range("sheet1.ReturnFileSec").value = "153A" Then
           Range("sheet1.ReturnFileSec1").value = "15 - 153A"
        ElseIf Range("sheet1.ReturnFileSec").value = "153C" Then
           Range("sheet1.ReturnFileSec1").value = "16 - 153C"
        ElseIf Range("sheet1.ReturnFileSec").value = "139(5)-Revised Return" Then
           Range("sheet1.ReturnFileSec1").value = "17 - 139(5)-Revised Return"
        ElseIf Range("sheet1.ReturnFileSec").value = "139(9)" Then
           Range("sheet1.ReturnFileSec1").value = "18 - 139(9)"
           
           'J200924
       ' ElseIf Range("sheet1.ReturnFileSec").value = "119(2)(b)- after condonation of delay" Then
         '  Range("sheet1.ReturnFileSec1").value = "20 - 119(2)(b)- after condonation of delay"
           
            ElseIf Range("sheet1.ReturnFileSec").value = "139(9A) - After condonation of delay u/s 119(2)(b)" Then
           Range("sheet1.ReturnFileSec1").value = "20 - 139(9A) - After condonation of delay u/s 119(2)(b)"
       
        End If


            If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = "20" Then
            
            'J200924
           ' fmsgboxsmall "* 119(2)(b)- after condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)" ', vbInformation, "ITR4"
            fmsgboxsmall "* 139(9A) - After condonation of delay u/s 119(2)(b) should be filed after expiry of due date mentioned in 139(1)/139(4)" ', vbInformation, "ITR4"
             End If



            LockUnlockReturnfilesec


        '----------------Lock Password-------------------START---
            Sheet1.Protect Password:=sPassword
        '----------------Lock Password-------------------END-----

    End If
'
'     If (Target.name.name) = "sheet1.ReturnFileSec2" Then
'            '------------------Un Protect--------------------
'                    sPassword = mdCommon.getmsgstate
'                    Sheet1.Unprotect Password:=sPassword
'            '------------------Un Protect Ends--------------------
'
'If Range("sheet1.ReturnFileSec").value <> "(Select)" And Range("sheet1.ReturnFileSec").value <> "" Then
'      MsgBox "Either ""Filed u/s"" or ""Filed in response to notice u/s"" must have to be selected in tab: Part A General Information", vbOKOnly, "ITR4"
'      Range("sheet1.ReturnFileSec2").value = "(Select)"
'      Range("sheet1.ReturnFileSec").value = "(Select)"
'End If
'
'
'        If Range("sheet1.ReturnFileSec2").value = "(Select)" And Range("sheet1.ReturnFileSec").value = "(Select)" Then
'           Range("sheet1.ReturnFileSec1").value = "(Select)"
''        ElseIf Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
'''           Range("sheet1.ReturnFileSec1").value = "11 - 139(1)-On or before due date"
''        ElseIf Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
''           Range("sheet1.ReturnFileSec1").value = "12 - 139(4)-After due date"
'        ElseIf Range("sheet1.ReturnFileSec2").value = "142(1)" Then
'           Range("sheet1.ReturnFileSec1").value = "13 - 142(1)"
'        ElseIf Range("sheet1.ReturnFileSec2").value = "148" Then
'           Range("sheet1.ReturnFileSec1").value = "14 - 148"
'        ElseIf Range("sheet1.ReturnFileSec2").value = "153A" Then
'           Range("sheet1.ReturnFileSec1").value = "15 - 153A"
'        ElseIf Range("sheet1.ReturnFileSec2").value = "153C" Then
'           Range("sheet1.ReturnFileSec1").value = "16 - 153C"
''        ElseIf Range("sheet1.ReturnFileSec").value = "139(5)-Revised Return" Then
''           Range("sheet1.ReturnFileSec1").value = "17 - 139(5)-Revised Return"
'        ElseIf Range("sheet1.ReturnFileSec2").value = "139(9)" Then
'           Range("sheet1.ReturnFileSec1").value = "18 - 139(9)"
''        ElseIf Range("sheet1.ReturnFileSec").value = "119(2)(b)- after condonation of delay" Then
''           Range("sheet1.ReturnFileSec1").value = "20 - 119(2)(b)- after condonation of delay"
'        End If
'
'
''            If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = "20" Then
''                MsgBox " Sec139 read with section 119(2)(b) should be filed after expiry of due date mentioned in 139(1)/139(4).", vbInformation, "ITR4"
''             End If
'
'
'
'            LockUnlockReturnfilesec
'
'
'        '----------------Lock Password-------------------START---
'            Sheet1.Protect Password:=sPassword
'        '----------------Lock Password-------------------END-----
'
'    End If
'







End If


'If Target.name.name = "sheet1.ReturnType1" Then
'
'LockUnlockReturnfilesec
'                    sPassword = mdCommon.getmsgstate
'                    Sheet1.Unprotect Password:=sPassword
'
'        If Mid(Range("sheet1.ReturnType1").value, 1, 1) = "(" Then
'            Range("sheet1.ReturnFileSec").value = "(Select)"
'            Range("sheet1.ReturnFileSec1").value = "(Select)"
'        End If
'         If Range("sheet1.ReturnFileSec").value = "(Select)" Then
'           Range("sheet1.ReturnFileSec1").value = "(Select)"
'        ElseIf Range("sheet1.ReturnFileSec").value = "Voluntarily on or before the due date u/s 139(1)." Then
'           Range("sheet1.ReturnFileSec1").value = "11 - Voluntarily on or before the due date u/s 139(1)."
'        ElseIf Range("sheet1.ReturnFileSec").value = "Voluntarily after the due date u/s 139(4)." Then
'           Range("sheet1.ReturnFileSec1").value = "12 - Voluntarily after the due date u/s 139(4)."
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice u/s 142(1)." Then
'           Range("sheet1.ReturnFileSec1").value = "13 - In response to notice u/s 142(1)."
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice u/s 148." Then
'           Range("sheet1.ReturnFileSec1").value = "14 - In response to notice u/s 148."
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice u/s 153A." Then
'           Range("sheet1.ReturnFileSec1").value = "15 - In response to notice u/s 153A."
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to notice u/s 153C read with section 153A." Then
'           Range("sheet1.ReturnFileSec1").value = "16 - In response to notice u/s 153C read with section 153A."
'        ElseIf Range("sheet1.ReturnFileSec").value = "Revised return u/s 139(5)." Then
'           Range("sheet1.ReturnFileSec1").value = "17 - Revised return u/s 139(5)."
'        ElseIf Range("sheet1.ReturnFileSec").value = "In response to a notice u/s 139(9)" Then
'           Range("sheet1.ReturnFileSec1").value = "18 - In response to a notice u/s 139(9)"
''        ElseIf Range("sheet1.ReturnFileSec").value = "section 139 read with section 92 CD" Then
''           Range("sheet1.ReturnFileSec1").value = "19 - section 139 read with section 92 CD"
'        ElseIf Range("sheet1.ReturnFileSec").value = "section 139 read with section 119(2)(b)." Then
'           Range("sheet1.ReturnFileSec1").value = "20 - section 139 read with section 119(2)(b)."
'        End If
'
'
'    If Mid(Range("sheet1.ReturnType1").value, 1, 3) = "Rev" Then
'            Range("sheet1.ReturnFileSec").value = "Revised return u/s 139(5)."
'           ' Range("sheet1.ReturnFileSec1").value = "17 - Revised return u/s 139(5)."
'
'    ElseIf Mid(Range("sheet1.ReturnType1").value, 1, 1) = "I" Then
'            Range("sheet1.ReturnFileSec").value = "In response to a notice u/s 139(9)"
'           'Range("sheet1.ReturnFileSec1").value = "18 - In response to a notice u/s 139(9) where the original return filed was a Defective return."
'    End If
'
'        If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 3) = "Rev" Then
'            If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) <> "17" Then
'                        MsgBox "'Filed u/s' field in 'Part A-General information' can only be 'Revised return under section 139(5)' for Filing type 'Revised'"
'                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
'                        Sheet1.Range("sheet1.ReturnType1").Select
'            End If
'        End If
'
'        If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 3) = "Rec" Then
'            If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "18" Then
'                        MsgBox "'Filed u/s' field in 'Part A-General information' cannot be 'In response to notice under section 139(9)' for Filing type 'Rectification'"
'                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
'                        Sheet1.Range("sheet1.ReturnType1").Select
'            End If
'        End If
'
'         If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 1) = "I" Then
'            If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) <> "18" Then
'                        MsgBox "'Filed u/s' field in 'Part A-General information' can be only 'In response to notice under section 139(9)' for Filing type 'In response to a notice u/s 139(9)'"
'                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
'                        Sheet1.Range("sheet1.ReturnType1").Select
'            End If
'        End If
'
'        If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 1) = "O" Then
'            If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "18" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "17" Then
'                        MsgBox "'Filed u/s' field in 'Part A-General information' cannot be 'In response to notice under section 139(9)' / 'Revised return under section 139(5)' for Filing type 'Original'"
'                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
'                        Sheet1.Range("sheet1.ReturnType1").Select
'            End If
'        End If
'LockUnlockReturnfilesec
''       If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "11" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "12" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "13" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "14" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "15" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "16" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "20" Or Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "19" Then
''           If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 1) = "I" Then
''                        MsgBox "Section field in Part A-General information can be only "" In response to notice under section 139(9)"" for filing type ""In response to a notice u/s 139(9)"" ."
''                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                        Sheet1.Range("sheet1.ReturnType1").Select
''            End If
''
''         If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 3) = "Rev" Then
''                        MsgBox "Section field in Part A-General information can only be ""Revised return under section 139(5)"" for filing type ""Revised""."
''                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                        Sheet1.Range("sheet1.ReturnType1").Select
''            End If
''
''        End If
''
''        If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "17" Then
''        If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 1) = "O" Then
''                MsgBox "Section field in Part A-General information cannot be ""In response to notice under section 139(9)"" / ""revised return under section 139(5)"" for filing type ""Original""."
''                Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                Sheet1.Range("sheet1.ReturnType1").Select
''        End If
''
''        If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 1) = "I" Then
''                MsgBox "Section field in Part A-General information can be only ""In response to notice under section 139(9)"" for filing type ""In response to a notice u/s 139(9)""."
''                Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                Sheet1.Range("sheet1.ReturnType1").Select
''        End If
''
''        End If
''
''        If Mid(Trim(Range("sheet1.ReturnFileSec1").value), 1, 2) = "18" Then
''        If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 1) = "O" Then
''                MsgBox "Section field in Part A-General information cannot be ""In response to notice under section 139(9)"" / ""revised return under section 139(5)"" for filing type ""Original""."
''                Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                Sheet1.Range("sheet1.ReturnType1").Select
''        End If
''
''            If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 3) = "Rev" Then
''                        MsgBox "Section field in Part A-General information can only be ""Revised return under section 139(5)"" for filing type ""Revised""."
''                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                        Sheet1.Range("sheet1.ReturnType1").Select
''            End If
''
''            If Mid(Trim(Range("sheet1.ReturnType1").value), 1, 3) = "Rec" Then
''                        MsgBox "Section field in Part A-General information cannot be ""In response to notice under section 139(9)"" for filing type ""Rectification""."
''                        Sheet1.Range("sheet1.ReturnType1").value = "(Select)"
''                        Sheet1.Range("sheet1.ReturnType1").Select
''            End If
''
''        End If
'        Sheet1.Protect Password:=sPassword
'End If


If Target.Validation.Type = 6 Then

 If (Target.name.name) = "sheet1.Aadhaar" Then
                If Len(Sheet1.Range("sheet1.Aadhaar").value) >= 12 Then
                    If chkNumeric(Sheet1.Range("sheet1.Aadhaar").value) Then
                       If Sheet1.Range("sheet1.Aadhaar").value <> "000000000000" Then
                         Sheet1.Range("sheet1.Aadhaar").value = Sheet1.Range("sheet1.Aadhaar").value
                       Else
                         fmsgboxStatus "* Invalid Aadhaar Number"
                         Sheet1.Range("sheet1.Aadhaar").value = ""
                         Sheet1.Range("sheet1.Aadhaar").Select
                       End If


                       If Sheet1.Range("sheet1.Aadhaar").value < 0 Then
                            fmsgboxStatus "* Invalid Aadhaar Number"
                         Sheet1.Range("sheet1.Aadhaar").value = ""
                         Sheet1.Range("sheet1.Aadhaar").Select
                        End If


                       If Sheet1.Range("sheet1.Aadhaar").value <> "111111111111" Then
                         Sheet1.Range("sheet1.Aadhaar").value = Sheet1.Range("sheet1.Aadhaar").value
                       Else
                         fmsgboxStatus "* Invalid Aadhaar Number"
                         Sheet1.Range("sheet1.Aadhaar").value = ""
                         Sheet1.Range("sheet1.Aadhaar").Select
                       End If

                    Else
                      fmsgboxStatus "* Invalid Aadhaar Number"
                      Sheet1.Range("sheet1.Aadhaar").value = ""
                      Sheet1.Range("sheet1.Aadhaar").Select
                    End If
                Else
                    Sheet1.Range("sheet1.Aadhaar").value = ""
                End If
            End If


            'new adhaar enrol
            'Ayush_enrol=====================
'             If (Target.name.name) = "sheet1.AadhaarEnrol" Then
'                If Len(Sheet1.Range("sheet1.AadhaarEnrol").value) >= 12 Then
'                    If chkNumeric(Sheet1.Range("sheet1.AadhaarEnrol").value) Then
'                       If Sheet1.Range("sheet1.AadhaarEnrol").value <> "0000000000000000000000000000" Then
'                         Sheet1.Range("sheet1.AadhaarEnrol").value = Sheet1.Range("sheet1.AadhaarEnrol").value
'                       Else
'                         fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'                         Sheet1.Range("sheet1.AadhaarEnrol").value = ""
'                         Sheet1.Range("sheet1.AadhaarEnrol").Select
'                       End If
'
'
'                       If Sheet1.Range("sheet1.AadhaarEnrol").value < 0 Then
'                       fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'                         Sheet1.Range("sheet1.AadhaarEnrol").value = ""
'                         Sheet1.Range("sheet1.AadhaarEnrol").Select
'                         End If
'
'                       If Sheet1.Range("sheet1.AadhaarEnrol").value <> "1111111111111111111111111111" Then
'                         Sheet1.Range("sheet1.AadhaarEnrol").value = Sheet1.Range("sheet1.AadhaarEnrol").value
'                       Else
'                         fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'                         Sheet1.Range("sheet1.AadhaarEnrol").value = ""
'                         Sheet1.Range("sheet1.AadhaarEnrol").Select
'                       End If
'
'                    Else
'                      fmsgboxStatus "* Please enter the valid Aadhaar Enrolment ID"
'                      Sheet1.Range("sheet1.AadhaarEnrol").value = ""
'                      Sheet1.Range("sheet1.AadhaarEnrol").Select
'                    End If
'                Else
'                    Sheet1.Range("sheet1.AadhaarEnrol").value = ""
'                End If
'            End If
'===================


    If (Target.name.name) = "sheet1.PinCode" Then
        If Len(Range("sheet1.PinCode").value) = 6 Then
            If Range("sheet1.PinCode").value = 0 Then
                fmsgboxStatus ("* Pincode must be exactly 6 digits and should not start with zero in Sheet Income Details")
                Range("sheet1.PinCode").value = ""
                Range("sheet1.PinCode").Select
            End If

            If Len(Range("sheet1.PinCode").value) = 6 Then
                If Not ((IsNumeric(Range("sheet1.PinCode").value)) And checkspecialcharacterForPinCode(Range("sheet1.PinCode").value)) Then
                    fmsgboxStatus ("* Invalid Pincode :Pincode cannot have Alphabets/Special characters and should be 6 digits Numeric Value in Sheet Income Details")
                    Range("sheet1.PinCode").value = ""
                    Range("sheet1.PinCode").Select
                End If
            End If
            If Mid(Range("sheet1.PinCode").value, 1, 1) = 0 Then
                fmsgboxStatus ("* Pincode cannot begin with zero")
                Range("sheet1.PinCode").value = ""
                Range("sheet1.PinCode").Select
            End If
        End If
                If Range("sheet1.CountryCode") = "" Or Range("sheet1.CountryCode") = "(Select)" Then
                     Range("sheet1.CountryCode") = "91-INDIA"
                End If
                 Dim PinCode As Range
                 For Each PinCode In Sheet8.Range("All_Pincode_list")
                      Dim code As String
                    code = Left(CStr(Sheet1.Range("sheet1.PinCode").value), 6)
                    If CStr(PinCode.value) = CStr(code) Then
                        Sheet1.Range("sheet1.StateCode1").value = Sheet8.Range("EQ" & PinCode.row).value
                        Exit For
                    End If
                Next PinCode

    End If

      If (Target.name.name) = "sheet1.Mobileno" Then
        If (Range("sheet1.Mobileno")) <> "" Then
      If Sheet1.Range("sheet1.mobileCountryCode").value = "91" Then
            If Len(Range("sheet1.Mobileno")) <> 10 Then
                fmsgboxStatus "* Mobile number 1 should be 10 digits in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.Mobileno").Select
                Range("sheet1.Mobileno").MergeArea.ClearContents
            End If
            Else
            If Len(Range("sheet1.Mobileno")) < 5 Or Len(Range("sheet1.Mobileno")) > 10 Then
                fmsgboxStatus "* Mobile number 1 is invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.Mobileno").Select
                Range("sheet1.Mobileno").MergeArea.ClearContents
            End If
        End If

        If Len(Range("sheet1.Mobileno")) > 0 Then
            If Not IsNumeric(Trim(Range("sheet1.Mobileno").value)) Then
                fmsgboxStatus "* Mobile number 1 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.Mobileno").Select
                Range("sheet1.Mobileno").MergeArea.ClearContents
            End If
        End If

        If Len(Range("sheet1.Mobileno")) > 0 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.Mobileno").value)) Then
                fmsgboxStatus "* Mobile number 1 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                 Range("sheet1.Mobileno").Select
                Range("sheet1.Mobileno").MergeArea.ClearContents
            End If
        End If

        If Len(Range("sheet1.Mobileno")) > 0 Then
            If Mid(Trim(Range("sheet1.Mobileno").value), 1, 1) = 0 Then
                fmsgboxStatus "* Mobile number 1 should not begin with 0" ', vbOKOnly, vbMessageCaption
                Range("sheet1.Mobileno").Select
                Range("sheet1.Mobileno").MergeArea.ClearContents

            End If
        End If
    

'        If Len(Range("sheet1.Mobileno")) > 0 Then
'            If Not CheckMobileNo(Trim(Range("sheet1.Mobileno").value)) Then
'                MsgBox "Mobile number 1 should not begin with 0", vbOKOnly, vbMessageCaption
'                Range("sheet1.Mobileno").Select
'                Range("sheet1.Mobileno").MergeArea.ClearContents
'            End If
'        End If

        If Len(Range("sheet1.Mobileno")) + Len(Range("sheet1.mobileCountryCode")) > 15 Then
                fmsgboxStatus "* Mobile number and country code combination should not be more than 15 digits.in Sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.Mobileno").Select
                Range("sheet1.Mobileno").MergeArea.ClearContents
                Range("sheet1.mobileCountryCode").Select
                Range("sheet1.mobileCountryCode").value = ""

        End If
        End If
    End If

    If (Target.name.name) = "sheet1.mobileCountryCode" Then
        If Len(Range("sheet1.mobileCountryCode")) <= 5 Then
            If Not IsNumeric(Trim(Range("sheet1.mobileCountryCode").value)) Then
                fmsgboxStatus "* Country code for Mobile number 1 is Invalid in sheet Income Details" '' , vbOKOnly, vbMessageCaption
                Range("sheet1.mobileCountryCode").Select
                Range("sheet1.mobileCountryCode").value = ""

            End If
        End If

        If Len(Range("sheet1.mobileCountryCode")) <= 5 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.mobileCountryCode").value)) Then
                fmsgboxStatus "* Country code for Mobile number 1 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                 Range("sheet1.mobileCountryCode").Select
                Range("sheet1.mobileCountryCode").value = ""

            End If
        End If

        If Len(Range("sheet1.mobileCountryCode")) <= 5 Then
            If Not chkNumeric(Trim(Range("sheet1.mobileCountryCode").value)) Then
                fmsgboxStatus "* Country code for Mobile number 1 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.mobileCountryCode").Select
                Range("sheet1.mobileCountryCode").value = ""

            End If
        End If

    End If


    If (Target.name.name) = "sheet1.MobileNoSec" Then
      If Sheet1.Range("sheet1.mobileCountryCode2").value = "91" Then
            If Len(Range("sheet1.MobileNoSec")) <> 10 Then
                fmsgboxStatus "* Mobile Number 2 should be 10 digits in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.MobileNoSec").Select
               Range("sheet1.MobileNoSec").MergeArea.ClearContents

            End If
            Else
            If Len(Range("sheet1.MobileNoSec")) < 5 Or Len(Range("sheet1.MobileNoSec")) > 10 Then
                fmsgboxStatus "* Mobile Number 2 is invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.MobileNoSec").Select
               Range("sheet1.MobileNoSec").MergeArea.ClearContents

            End If
        End If

        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            If Not IsNumeric(Trim(Range("sheet1.MobileNoSec").value)) Then
                fmsgboxStatus "* Mobile Number 2 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.MobileNoSec").Select
               Range("sheet1.MobileNoSec").MergeArea.ClearContents

            End If
        End If


        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.MobileNoSec").value)) Then
                fmsgboxStatus "* Mobile number 2 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.MobileNoSec").Select
               Range("sheet1.MobileNoSec").MergeArea.ClearContents

            End If
        End If
        If Len(Range("sheet1.MobileNoSec")) > 0 Then
            If Mid(Trim(Range("sheet1.MobileNoSec").value), 1, 1) = 0 Then
                fmsgboxStatus "* Mobile number 2 should not begin with 0" ' , vbOKOnly, vbMessageCaption
                Range("sheet1.MobileNoSec").Select
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
            End If
        End If

        If Len(Range("sheet1.MobileNoSec")) > 0 Then
        If (Range("sheet1.mobileCountryCode2").value) = "" Then
        
        'Error Message Changed By Jyoti for the SIT-67665 on 25/06/2024
        
       ' fmsgboxStatus "* Please enter country code for Mobile Number 2." ', vbOKOnly, vbMessageCaption
       
         fmsgboxStatus "*""Country Code in Mobile Number 2 in Sheet: Personal Information is mandatory.""" ', vbOKOnly, vbMessageCaption
         
            End If
        End If

        If Len(Range("sheet1.MobileNoSec")) + Len(Range("sheet1.mobileCountryCode2")) > 15 Then
                fmsgboxStatus "* Mobile number and country code combination should not be more than 15 digits.in Sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.MobileNoSec").Select
               Range("sheet1.MobileNoSec").MergeArea.ClearContents
                Range("sheet1.mobileCountryCode2").Select
                Range("sheet1.mobileCountryCode2").value = ""

            End If
    End If


    If (Target.name.name) = "sheet1.mobileCountryCode2" Then
        If Len(Range("sheet1.mobileCountryCode2")) <= 5 Then
            If Not IsNumeric(Trim(Range("sheet1.mobileCountryCode2").value)) Then
                fmsgboxStatus "* Country code for Mobile number 2 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.mobileCountryCode2").Select
                Range("sheet1.mobileCountryCode2").value = ""

            End If
        End If

        If Len(Range("sheet1.mobileCountryCode2")) <= 5 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.mobileCountryCode2").value)) Then
                fmsgboxStatus "* Country code for Mobile number 2 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                 Range("sheet1.mobileCountryCode2").Select
                Range("sheet1.mobileCountryCode2").value = ""

            End If
        End If

        If Len(Range("sheet1.mobileCountryCode2")) <= 5 Then
            If Not chkNumeric(Trim(Range("sheet1.mobileCountryCode2").value)) Then
                fmsgboxStatus "* Country code for Mobile number 2 is Invalid in sheet Income Details" ', vbOKOnly, vbMessageCaption
                Range("sheet1.mobileCountryCode2").Select
                Range("sheet1.mobileCountryCode2").value = ""

            End If
        End If

        If Len(Range("sheet1.mobileCountryCode2")) > 0 Then
        If (Range("sheet1.MobileNoSec").value) = "" Then
        fmsgboxStatus "* Please enter Mobile Number 2." ', vbOKOnly, vbMessageCaption
            End If
        End If

    End If

'    If (Target.name.name) = "sheet1.aadharValue" Then
'        If Sheet1.Range("sheet1.aadharValue").value <> "" Then
'            If Sheet1.Range("sheet1.aadharValue").value = 0 Then
'                MsgBox "Aadhaar Number is Invalid in Sheet: Income Details", vbOKOnly
'                Sheet1.Range("sheet1.aadharValue").value = ""
'                GoTo endline
'            End If
'            If Not IsNumeric(Sheet1.Range("sheet1.aadharValue").value) Then
'                    MsgBox "Aadhaar Number is Invalid in Sheet: Income Details", vbOKOnly
'                    Sheet1.Range("sheet1.aadharValue").value = ""
'            End If
'        End If
'            GoTo endline
'    End If


'PAN Validation on data entry
If (Target.name.name) = "sheet1.PAN" Then
                Dim TargetPAN As String
                TargetPAN = ""
                TargetPAN = Target.name.name
                If Not mdIncomeDetails.CheckPAN(UCase(Trim(Range(TargetPAN).value))) Then
                    fmsgboxStatus "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
                    Sheet1.Range(TargetPAN).Select
                    Sheet1.Range(TargetPAN).value = ""
                    GoTo endline
                End If
                Sheet1.Range(TargetPAN).value = UCase(Range(TargetPAN).value)
                'PANStatusChange
                If Mid(UCase(Trim(Sheet1.Range("sheet1.Status").value)), 1, 1) = "I" Then
                    Sheet5.Range("Ver.PAN").value = UCase(Sheet1.Range("sheet1.PAN").value)
                End If
                GoTo endline
            End If

'Spouse PAN Validation on data entry
'If (Target.name.name) = "sheet1.PANOFSPOUSE" Then
'                Dim TargetSpousePAN As String
'                TargetSpousePAN = ""
'                TargetSpousePAN = Target.name.name
'                If Not mdIncomeDetails.CheckPANIncomeDetails_Spouse(UCase(Trim(Range(TargetSpousePAN).value))) Then
'                    MsgBox "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet. 4th character must be 'P' only.", vbOKCancel, mdCommon.vbMessageCaption
'                    Sheet1.Range(TargetSpousePAN).Select
'                    Sheet1.Range(TargetSpousePAN).value = ""
'                    GoTo endline
'                End If
'
'                If (UCase(Range(TargetSpousePAN).value) = UCase(Sheet1.Range("sheet1.PAN").value)) Then
'                 MsgBox "PAN of the spouse cannot be same as assessee PAN .", vbOKCancel, mdCommon.vbMessageCaption
'                    Sheet1.Range(TargetSpousePAN).Select
'                    Sheet1.Range(TargetSpousePAN).value = ""
'                    GoTo endline
'                End If
'
'
'                Sheet1.Range(TargetSpousePAN).value = UCase(Range(TargetSpousePAN).value)
'                GoTo endline
'            End If

'        If (Target.name.name) = "Sheet1.date10IE" Then
'            If Trim(Range("Sheet1.date10IE").value) <> "00/00/0000" And Range("Sheet1.date10IE").value <> "" Then
'                    If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE").value)) Then
'                            fmsgboxStatus "* Date of filing of Form 10IE is Invalid in Sheet Income Details. Please enter a valid date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
'                            Sheet1.Range("Sheet1.date10IE").Select
'                            Sheet1.Range("Sheet1.date10IE").value = ""
'                     Else
'                        If ChkMaxDate(Trim(Range("Sheet1.date10IE").value), "") Then
'                            fmsgboxStatus "* Date of filing of Form 10IE should not be prior 01/04/2022 in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
'                            Sheet1.Range("Sheet1.date10IE").Select
'                            Sheet1.Range("Sheet1.date10IE").value = ""
'                    End If
'                    If Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'                            fmsgbox "* Date of filing of Form 10IE can not be more than system date in Income Details"
'                            Range("Sheet1.date10IE").Select
'                            Range("Sheet1.date10IE").value = ""
'                        End If
'                     End If
'                End If
'        End If

        If (Target.name.name) = "sheet1.DOB" Then
            If Range("sheet1.status").value = "I - INDIVIDUAL" Then
            If Trim(Range("sheet1.DOB").value) = "01/01/0001" Then
            fmsgboxStatus "* Date of Birth in Sheet Income Details Please enter a valid date.in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").value = ""
            End If
            End If

            LockUnlock80TTB
            Sheet1.Unprotect Password:=getmsgstate




'            If calculateAge(DOB_6) <= 59 Then
'
'                Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
'                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
'                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HD8D8D8"
'
'                Sheet1.Range("IncD.Section80TTA").MergeArea.ClearContents
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HCCFFCC"
'
'            ElseIf calculateAge(DOB_6) > 59 And calculateAge(DOB_6) <> 2019 Then
'
'                Sheet1.Range("IncD.Section80TTA").MergeArea.ClearContents
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HD8D8D8"
'
'                Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
'                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HCCFFCC"
'
'            ElseIf calculateAge(DOB_6) = 2019 Then
'
'                Sheet1.Range("IncD.Section80TTA").MergeArea.ClearContents
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80TTA").MergeArea.Interior.Color = "&HCCFFCC"
'
'                Sheet1.Range("IncD.Section80TTB").MergeArea.ClearContents
'                Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
'                Sheet1.Range("IncD.Section80TTB").MergeArea.Interior.Color = "&HCCFFCC"
'
'            End If


           Sheet1.Protect Password:=getmsgstate


            End If



'DOB Validation: On Data Entry
            If (Target.name.name) = "sheet1.DOB" Then
            Dim HUFReq As Long
            If Range("sheet1.status").value = "H - HUF" Then HUFReq = 1 Else HUFReq = 0
                If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then

                    result = CheckDate(Trim(Range("sheet1.DOB").value), HUFReq)
                    If result = 1 Then
                        fmsgboxStatus "* Date of Birth in Sheet Income Details Please enter a valid date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").value = ""
                    ElseIf result = 2 Then
                        fmsgboxStatus "* Date of Birth in Sheet Income Details  should not be after 01/04/2022" ', vbOKOnly, mdCommon.vbMessageCaption
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").value = ""
                    End If
                End If
                GoTo endline
            End If

'Date of filling orignal return Validation: On Data Entry
            If (Target.name.name) = "sheet1.OrigRetFiledDate" Then
                If Trim(Range("sheet1.OrigRetFiledDate").value) <> "00/00/0000" And Range("sheet1.OrigRetFiledDate").value <> "" Then
                    If Not FormatNCheckDate(Trim(Range("sheet1.OrigRetFiledDate").value)) Then
                            fmsgboxStatus "* Date Of filing Orignal Return is Invalid in Sheet Income Details. Please enter a valid date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
                            Sheet1.Range("sheet1.OrigRetFiledDate").Select
                            Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                     Else
                        If ChkMaxDate(Trim(Range("sheet1.OrigRetFiledDate").value), "") Then
                            fmsgboxStatus "* Date Of  filing Orignal Return should not be prior 01/04/2022 in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                            Sheet1.Range("sheet1.OrigRetFiledDate").Select
                            Sheet1.Range("sheet1.OrigRetFiledDate").value = ""
                        End If
                     End If
                End If

                If Dformat(Trim(Range("sheet1.OrigRetFiledDate").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                 fmsgbox "* Date of filing of original return can not be more than system date in Income Details"
                        Range("sheet1.OrigRetFiledDate").Select
                        Range("sheet1.OrigRetFiledDate").value = ""

                End If


                GoTo endline
             End If

'             If (Target.name.name) = "sheet1.OrigRetFiledDate1" Then
'                If Trim(Range("sheet1.OrigRetFiledDate1").value) <> "00/00/0000" And Range("sheet1.OrigRetFiledDate1").value <> "" Then
'                    If Not FormatNCheckDate(Trim(Range("sheet1.OrigRetFiledDate1").value)) Then
'                            MsgBox "Date Of filing Orignal Return is Invalid in Sheet: Income Details. Please enter a valid date in dd/mm/yyyy format.", vbOKOnly, mdCommon.vbMessageCaption
'                            Sheet1.Range("sheet1.OrigRetFiledDate1").Select
'                            Sheet1.Range("sheet1.OrigRetFiledDate1").value = ""
'                     Else
'                        If ChkMaxDate(Trim(Range("sheet1.OrigRetFiledDate1").value), "") Then
'                            MsgBox "Date Of  filing Orignal Return should not be prior 01/04/2019 in Sheet: Income Details", vbOKOnly, mdCommon.vbMessageCaption
'                            Sheet1.Range("sheet1.OrigRetFiledDate1").Select
'                            Sheet1.Range("sheet1.OrigRetFiledDate1").value = ""
'                        End If
'                     End If
'                End If
'                GoTo endline
'             End If

'Date of filling in response to notice Validation: On Data Entry
            If (Target.name.name) = "sheet1.NoticeDate" Then
                If Trim(Range("sheet1.NoticeDate").value) <> "00/00/0000" Then
                    If Not FormatNCheckDate(Range("sheet1.NoticeDate").value) Then
                        fmsgboxStatus "* Please enter a valid date in dd/mm/yyyy format in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                        Sheet1.Range("sheet1.NoticeDate").Select
                        Sheet1.Range("sheet1.NoticeDate").value = ""
                    Else
                        If Range("sheet1.NoticeDate").value <> "" Then
                            If ChkMaxDate(Range("sheet1.NoticeDate").value, "") Then
                                fmsgboxStatus "*Date of notice or Order cannot be prior to 01/04/2022 in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                                Sheet1.Range("sheet1.NoticeDate").Select
                                Sheet1.Range("sheet1.NoticeDate").value = ""
                            End If
                        End If

                    End If
                    result = ""
                End If
                
                    If Dformat(Trim(Range("sheet1.NoticeDate").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                     fmsgbox "* Date of such notice or order cannot be more than system date in Income Details"
                     Range("sheet1.NoticeDate").Select
                     Range("sheet1.NoticeDate").value = ""

                    End If
                
                GoTo endline
            End If

'EMAIL Validation : On data entry
        If (Target.name.name) = "sheet1.EmailAddress" Then
         If Range("sheet1.EmailAddress").value <> "" Then
            If Not mdIncomeDetails.CheckEmailAddress(Trim(Range("sheet1.EmailAddress").value), validateType) Then
                If validateType = "1" Then fmsgboxStatus "* Invalid E-mail ID (Self) Please retry." ', vbOKOnly, mdCommon.vbMessageCaption
                If validateType = "2" Then fmsgboxStatus "* Email ID (Self) should not exceed 125 characters in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                If validateType = "3" Then fmsgboxStatus "* Email ID (Self) is invalid in Sheet  Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                If validateType = "4" Then fmsgboxStatus "* Email ID (Self) is mandatory in tab Part A General Information" ', vbOKOnly, mdCommon.vbMessageCaption
                Range("sheet1.EmailAddress").Select
                Range("sheet1.EmailAddress").value = ""
            End If
          GoTo endline
          End If
        End If

        If (Target.name.name) = "sheet1.EmailAddress2" Then
            If Range("sheet1.EmailAddress2").value <> "" Then
            If Not mdIncomeDetails.CheckEmailAddress(Trim(Range("sheet1.EmailAddress2").value), validateType) Then
                If validateType = "1" Then fmsgboxStatus "* Invalid E-mail ID-2 Please retry" ', vbOKOnly, mdCommon.vbMessageCaption
                If validateType = "2" Then fmsgboxStatus "* E-mail ID-2 should not exceed 125 characters in Sheet  Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                If validateType = "3" Then fmsgboxStatus "* E-mail ID-2 is invalid in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                'If validateType = "4" Then MsgBox "Email Address in Sheet Income Details  is Mandatory", vbOKOnly, mdCommon.vbMessageCaption
                Range("sheet1.EmailAddress2").Select
                Range("sheet1.EmailAddress2").value = ""
            End If
          GoTo endline
          End If
        End If

 'Std & phone" On data Entry
            If (Target.name.name) = "sheet1.STDcode" Then
                Dim val As Variant
               val = mdIncomeDetails.ValidateSTDcode(Trim(Range("sheet1.STDcode").value))
                If val <> "0" Then
                    fmsgboxStatus "* STD/ISD Code " & val & " in Sheet Income Details " ', vbOKOnly, mdCommon.vbMessageCaption
                     Sheet1.Range("sheet1.STDcode").value = ""
                     Sheet1.Range("sheet1.STDcode").Select
                    GoTo endline
                 End If
                 result = mdIncomeDetails.ValidatePhoneNo(Trim(Range("sheet1.STDcode").value), Trim(Range("sheet1.PhoneNo").value))
                            If result <> "0" Then
                                fmsgboxStatus "* Phone Number " & result & " in Sheet  Income Details " ', vbOKOnly, vbMessageCaption
                                Sheet1.Range("sheet1.PhoneNo").value = ""
                                GoTo endline
                            End If
                    stDcode = Sheet1.Range("sheet1.STDcode").value
                    GoTo endline
             End If

 'Phone NO" On data Entry
             If (Target.name.name) = "sheet1.PhoneNo" Then
                 If Sheet1.Range("sheet1.STDcode").value <> "" Then
                     If mdIncomeDetails.ValidateSTDcode(Trim(Range("sheet1.STDcode").value)) = 0 Then
                        result = mdIncomeDetails.ValidatePhoneNo(Trim(Range("sheet1.STDcode").value), Trim(Range("sheet1.PhoneNo").value))
                            If result <> "0" Then
                                fmsgboxStatus "* Phone Number " & result & " in Sheet Income Details " ', vbOKOnly, vbMessageCaption
                                Sheet1.Range("sheet1.PhoneNo").value = ""
                                GoTo endline
                            End If
                    Else
                        Sheet1.Range("sheet1.PhoneNo").value = ""
                        Sheet1.Range("sheet1.STDcode").Select
                        GoTo endline
                    End If
                 End If
                    phoneNo = Sheet1.Range("sheet1.PhoneNo").value
                    GoTo endline
            End If
    If (Target.name.name) = "sheet1.ReceiptNo" Then
        If Range("Sheet1.ReceiptNo").value <> "" Then

            If Not chkNumeric(Range("Sheet1.ReceiptNo").value) Then
            fmsgboxStatus ("* Reciept Number is Invalid in Sheet Income Details")
                Range("Sheet1.ReceiptNo").value = ""
                Range("Sheet1.ReceiptNo").Select
            End If

            If Range("Sheet1.ReceiptNo").value = "000000000000000" Then
                fmsgboxStatus ("* Reciept Number is Invalid in Sheet Income Details")
                Range("Sheet1.ReceiptNo").value = ""
                Range("Sheet1.ReceiptNo").Select
            End If


        End If
    End If

'    If (Target.name.name) = "sheet1.ReceiptNo1" Then
'        If Range("Sheet1.ReceiptNo1").value <> "" Then
'
'            If Not chkNumeric(Range("Sheet1.ReceiptNo1").value) Then
'            MsgBox ("Reciept Number is Invalid in Sheet : Income Details")
'                Range("Sheet1.ReceiptNo1").value = ""
'                Range("Sheet1.ReceiptNo1").Select
'            End If
'
'            If Range("Sheet1.ReceiptNo1").value = "000000000000000" Then
'                MsgBox ("Reciept Number is Invalid in Sheet : Income Details")
'                Range("Sheet1.ReceiptNo1").value = ""
'                Range("Sheet1.ReceiptNo1").Select
'            End If
'
'
'        End If
'    End If

End If


If (Target.name.name) = "sheet1.Status" Or (Target.name.name) = "sheet1.MiddleName" Or (Target.name.name) = "sheet1.FirstName" Or (Target.name.name) = "sheet1.SurNameOrOrgName" Then
        If (Mid(Range("sheet1.status").value, 1, 1) = "I") Or (Mid(Range("sheet1.status").value, 1, 1) = "(") Then
            sPassword = mdCommon.getmsgstate
            Sheet5.Unprotect Password:=sPassword
            Sheet5.Range("Ver.AssesseeVerName").value = UCase(Sheet1.Range("sheet1.FirstName").value + " " + Sheet1.Range("sheet1.MiddleName").value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").value)
            Sheet5.Protect Password:=sPassword
            GoTo endline
        End If
    End If

'    If (Target.name.name) = "IncD.Section80D" Or (Target.name.name) = "IncD.Section80DB" Or (Target.name.name) = "IncD.Section80DC" Then
'    Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'    End If
'
'    If (Target.name.name) = "IncD.Section80D_Calc" Then
'    Sheet1.Range("IncD.Section80D_Calc").Formula = "=IF(IS_HUF,MIN(50000,SUM(IF(IS_FIRM,0,IF(ISERROR(IncdSection80D),0,IncdSection80D)),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))),MIN(100000,SUM(IF(IS_FIRM,0,TempNew80D),IF(IS_FIRM,0,IF(ISERROR(IncdSection80DB),0,IncdSection80DB)))))"
'    End If

    If (Target.name.name) = "IncD.Section80D_Calc" Then
    If (Range("IncD.Section80D_Calc").value) > 100000 Then
    fmsgboxStatus ("* Section80D Value cannot exceed 100000 in Sheet Income Details")
    End If
    End If

    If (Target.name.name) = "IncD.Section80EE" Then
    If (Range("IncD.Section80EE").value) > 0 Then
    'And (Range("IncD.Section80EE").value) <> 0 Then
    Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Protect Password:=getmsgstate
        
    Else
    
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Protect Password:=getmsgstate
    
    End If
        Sheet1.Protect Password:=getmsgstate
    End If
    'DPK
    If (Target.name.name) = "IncD.Section80EEA" Then
    If (Range("IncD.Section80EEA").value) > 0 Then
    'And (Range("IncD.Section80EE").value) <> 0 Then
    
    Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("IncD.Section80EE").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Protect Password:=getmsgstate
    Else
    
    
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("IncD.Section80EE").MergeArea.ClearContents
        Sheet1.Range("IncD.Section80EE").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80EE").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Protect Password:=getmsgstate
    
    End If
    End If

'    If Target.Validation.Type = 1 Then
'
'    End If


    If (Target.name.name) = "IncD.Section89" Then
    If (Range("IncD.Section89").value) = "" Then
    Range("IncD.Section89").value = 0
    End If
MsgBox "df"
    If (Range("IncD.IncomeFromSal_i").value) = 0 Then
    fmsgboxStatus ("* If Salary is not more than zero then Relief u/s 89 cannot be more than zero in Sheet Income Details")
    Range("IncD.Section89").value = 0
    End If

    End If
    
'    If Target.Validation.Type = 1 Then
'    If Sheet1.Range("IncD.IncomeFromSal_allow").value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
'    fmsgboxStatus "* Total of allowances shall not exceed total of salary as per 17(1)"
'    End If
'    End If
    
 If (Target.name.name) = "IncD.TypeOfHP" Then
    If Not Range("IncD.TypeOfHP").value = "Self Occupied" Then
        If Sheet8.Range("bacValue").value = 1 Then
        Sheet1.Unprotect Password:=getmsgstate
         Range("IncD.IncomeFromHP_v").MergeArea.Locked = False
             Range("IncD.IncomeFromHP_v").Interior.Color = (&HCCFFCC)  ', "AU"
       
        End If
    Else
         If Sheet8.Range("bacValue").value = 1 Then
         
        Sheet1.Unprotect Password:=getmsgstate
         Range("IncD.IncomeFromHP_v").MergeArea.Locked = True
             Range("IncD.IncomeFromHP_v").Interior.Color = "&HD8D8D8"  ', "AU"
             Range("IncD.IncomeFromHP_v").value = ""
       
        End If
    End If
End If

'ID_Nature_Employment_New


Sheet1.Protect Password:=mdCommon.getmsgstate
endline:
 Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
Function CountDropdown1(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long


     currCnt22 = 0
      currCnt33 = 0
      currCnt44 = 0
'      currCnt55 = 0
      currCnt66 = 0
      currCnt67 = 0
      currCnt68 = 0
      currCnt69 = 0
      currCnt70 = 0

    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then


       If Mid(Trim(sDropdownValue), 1, 15) = "Interest from S" Then
       currCnt22 = currCnt22 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 15) = "Interest from D" Then
       currCnt33 = currCnt33 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 15) = "Interest from I" Then
       currCnt44 = currCnt44 + 1
       End If

'        If Mid(Trim(sDropdownValue), 1, 11) = "Any Other I" Then
'       currCnt55 = currCnt55 + 1
'       End If

        If Mid(Trim(sDropdownValue), 1, 1) = "F" Then
       currCnt66 = currCnt66 + 1
       End If
        
       If Mid(Trim(sDropdownValue), 1, 112) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)" Then
       currCnt67 = currCnt67 + 1
       End If
       
       If Mid(Trim(sDropdownValue), 1, 113) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)" Then
       currCnt68 = currCnt68 + 1
       End If
       
       If Mid(Trim(sDropdownValue), 1, 112) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)" Then
       currCnt69 = currCnt69 + 1
       End If
       
       If Mid(Trim(sDropdownValue), 1, 113) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)" Then
       currCnt70 = currCnt70 + 1
       End If
        



       End If
      Next
    End If
    CountDropdown = currCnt22
End Function

Private Sub Worksheet_Deactivate()

End Sub

'DropDown_ValueOf_FamilyM_80D
'DropDown_ValueOf_SC_80D
Sub LockUnlock80D()
' Mid(Trim(Range("sheet1.Status")), 1, 1) = "H"
Application.EnableEvents = False
            Sheet16.Unprotect Password:=getmsgstate
    If Mid(Trim(Range("sheet1.Status")), 1, 1) = "F" Then


        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").ClearContents
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = True
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("DropDown_ValueOf_SC_80D").ClearContents
        Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = True
        Sheet16.Range("DropDown_ValueOf_SC_80D").Interior.Color = (&HD8D8D8)

       Sheet16.Range("Self_And_Family_80D").ClearContents
        Sheet16.Range("Self_And_Family_80D").Locked = True
        Sheet16.Range("Self_And_Family_80D").Interior.Color = (&HD8D8D8)

'        Sheet16.Range("Health_Insurance_80D").ClearContents
'        Sheet16.Range("Health_Insurance_80D").Locked = True
'        Sheet16.Range("Health_Insurance_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health_80D").ClearContents
        Sheet16.Range("Preventive_Health_80D").Locked = True
        Sheet16.Range("Preventive_Health_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Senior_Citizen_80D").Locked = True
        Sheet16.Range("Senior_Citizen_80D").Interior.Color = (&HD8D8D8)
        Sheet16.Range("Senior_Citizen_80D").value = ""

'
'        Sheet16.Range("Health_InsuranceSC_80D").value = ""
'        Sheet16.Range("Health_InsuranceSC_80D").Locked = True
'        Sheet16.Range("Health_InsuranceSC_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health_SC_80D").value = ""
        Sheet16.Range("Preventive_Health_SC_80D").Locked = True
        Sheet16.Range("Preventive_Health_SC_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Medical_Expenditure_SC_80D").value = ""
        Sheet16.Range("Medical_Expenditure_SC_80D").Locked = True
        Sheet16.Range("Medical_Expenditure_SC_80D").Interior.Color = (&HD8D8D8)


        '2nd
        Sheet16.Range("Parents_80D").ClearContents
        Sheet16.Range("Parents_80D").Locked = True
        Sheet16.Range("Parents_80D").Interior.Color = (&HD8D8D8)

'        Sheet16.Range("Health_Insurance2_80D").ClearContents
'        Sheet16.Range("Health_Insurance2_80D").Locked = True
'        Sheet16.Range("Health_Insurance2_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health2_80D").ClearContents
        Sheet16.Range("Preventive_Health2_80D").Locked = True
        Sheet16.Range("Preventive_Health2_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Parents_SC_80D").ClearContents
        Sheet16.Range("Parents_SC_80D").Locked = True
        Sheet16.Range("Parents_SC_80D").Interior.Color = (&HD8D8D8)
'
'        Sheet16.Range("Health_Insurance3_80D").ClearContents
'        Sheet16.Range("Health_Insurance3_80D").Locked = True
'        Sheet16.Range("Health_Insurance3_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health3_80D").ClearContents
        Sheet16.Range("Preventive_Health3_80D").Locked = True
        Sheet16.Range("Preventive_Health3_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Medical_Expenditure2_80D").ClearContents
        Sheet16.Range("Medical_Expenditure2_80D").Locked = True
        Sheet16.Range("Medical_Expenditure2_80D").Interior.Color = (&HD8D8D8)

  ElseIf Mid(Trim(Sheet1.Range("sheet1.Status")), 1, 1) = "H" Then

        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "(Select)"
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = False
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("DropDown_ValueOf_SC_80D").ClearContents
        Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = True
        Sheet16.Range("DropDown_ValueOf_SC_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Senior_Citizen_80D").Locked = True
        Sheet16.Range("Senior_Citizen_80D").Interior.Color = 13434879
'        Sheet16.range("Senior_Citizen_80D").Font.Color = RGB(0, 0, 0)
        Sheet16.Range("Senior_Citizen_80D").Formula = "=MIN(50000,SUM(Health_InsuranceSC_80D,Preventive_Health_SC_80D,Medical_Expenditure_SC_80D))"

'        Sheet16.Range("Health_InsuranceSC_80D").Locked = False
'        Sheet16.Range("Health_InsuranceSC_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Preventive_Health_SC_80D").Locked = False
        Sheet16.Range("Preventive_Health_SC_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Medical_Expenditure_SC_80D").Locked = False
        Sheet16.Range("Medical_Expenditure_SC_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Self_And_Family_80D").Locked = True
        Sheet16.Range("Self_And_Family_80D").Interior.Color = 13434879
'        Sheet16.range("Self_And_Family_80D").Font.Color = RGB(0, 0, 0)
        Sheet16.Range("Self_And_Family_80D").Formula = "=MIN(25000,SUM(Health_Insurance_80D,Preventive_Health_80D))"

'        Sheet16.Range("Health_Insurance_80D").Locked = False
'        Sheet16.Range("Health_Insurance_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Preventive_Health_80D").Locked = False
        Sheet16.Range("Preventive_Health_80D").Interior.Color = (&HCCFFCC)


        Sheet16.Range("Parents_80D").ClearContents
        Sheet16.Range("Parents_80D").Locked = True
        Sheet16.Range("Parents_80D").Interior.Color = (&HD8D8D8)
        
        Sheet16.Range("NameInsurerA2.80D").ClearContents
        Sheet16.Range("NameInsurerA2.80D").Locked = True
        Sheet16.Range("NameInsurerA2.80D").Interior.Color = (&HD8D8D8)
        
        Sheet16.Range("PolicyNumA2.80D").ClearContents
        Sheet16.Range("PolicyNumA2.80D").Locked = True
        Sheet16.Range("PolicyNumA2.80D").Interior.Color = (&HD8D8D8)
        
        Sheet16.Range("AmtA2.80D").ClearContents
        Sheet16.Range("AmtA2.80D").Locked = True
        Sheet16.Range("AmtA2.80D").Interior.Color = (&HD8D8D8)

'        Sheet16.Range("Health_Insurance2_80D").ClearContents
'        Sheet16.Range("Health_Insurance2_80D").Locked = True
'        Sheet16.Range("Health_Insurance2_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health2_80D").ClearContents
        Sheet16.Range("Preventive_Health2_80D").Locked = True
        Sheet16.Range("Preventive_Health2_80D").Interior.Color = (&HD8D8D8)

         Sheet16.Range("Parents_SC_80D").ClearContents
        Sheet16.Range("Parents_SC_80D").Locked = True
        Sheet16.Range("Parents_SC_80D").Interior.Color = (&HD8D8D8)

'        Sheet16.Range("Health_Insurance3_80D").ClearContents
'        Sheet16.Range("Health_Insurance3_80D").Locked = True
'        Sheet16.Range("Health_Insurance3_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health3_80D").ClearContents
        Sheet16.Range("Preventive_Health3_80D").Locked = True
        Sheet16.Range("Preventive_Health3_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Medical_Expenditure2_80D").ClearContents
        Sheet16.Range("Medical_Expenditure2_80D").Locked = True
        Sheet16.Range("Medical_Expenditure2_80D").Interior.Color = (&HD8D8D8)
    
  ElseIf Sheet1.Range("sheet1.Status").value = "(Select)" Then
     
     
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = False
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Interior.Color = (&HCCFFCC)
        
        
        

       Sheet16.Range("Self_And_Family_80D").ClearContents
        Sheet16.Range("Self_And_Family_80D").Locked = True
        Sheet16.Range("Self_And_Family_80D").Interior.Color = (&HD8D8D8)
        
        
        
        
'        Sheet16.Range("Health_Insurance_80D").ClearContents
'        Sheet16.Range("Health_Insurance_80D").Locked = True
'        Sheet16.Range("Health_Insurance_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health_80D").ClearContents
        Sheet16.Range("Preventive_Health_80D").Locked = True
        Sheet16.Range("Preventive_Health_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Senior_Citizen_80D").Locked = True
        Sheet16.Range("Senior_Citizen_80D").Interior.Color = (&HD8D8D8)
        Sheet16.Range("Senior_Citizen_80D").value = ""


'        Sheet16.Range("Health_InsuranceSC_80D").value = ""
'        Sheet16.Range("Health_InsuranceSC_80D").Locked = True
'        Sheet16.Range("Health_InsuranceSC_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health_SC_80D").value = ""
        Sheet16.Range("Preventive_Health_SC_80D").Locked = True
        Sheet16.Range("Preventive_Health_SC_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Medical_Expenditure_SC_80D").value = ""
        Sheet16.Range("Medical_Expenditure_SC_80D").Locked = True
        Sheet16.Range("Medical_Expenditure_SC_80D").Interior.Color = (&HD8D8D8)

    '2nd
    
        Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
        Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = False
        Sheet16.Range("DropDown_ValueOf_SC_80D").Interior.Color = (&HCCFFCC)
        
        Sheet16.Range("Parents_80D").ClearContents
        Sheet16.Range("Parents_80D").Locked = True
        Sheet16.Range("Parents_80D").Interior.Color = (&HD8D8D8)
'        Range("Parents_80D").Font.Color = RGB(0, 0, 0)
        
'        Sheet16.Range("Health_Insurance2_80D").ClearContents
'        Sheet16.Range("Health_Insurance2_80D").Locked = True
'        Sheet16.Range("Health_Insurance2_80D").Interior.Color = (&HD8D8D8)
         
        Sheet16.Range("Preventive_Health2_80D").ClearContents
        Sheet16.Range("Preventive_Health2_80D").Locked = False
        Sheet16.Range("Preventive_Health2_80D").Interior.Color = (&HD8D8D8)


        Sheet16.Range("Parents_SC_80D").ClearContents
        Sheet16.Range("Parents_SC_80D").Locked = True
        Sheet16.Range("Parents_SC_80D").Interior.Color = (&HD8D8D8)

'        Sheet16.Range("Health_Insurance3_80D").ClearContents
'        Sheet16.Range("Health_Insurance3_80D").Locked = True
'        Sheet16.Range("Health_Insurance3_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Preventive_Health3_80D").ClearContents
        Sheet16.Range("Preventive_Health3_80D").Locked = True
        Sheet16.Range("Preventive_Health3_80D").Interior.Color = (&HD8D8D8)

        Sheet16.Range("Medical_Expenditure2_80D").ClearContents
        Sheet16.Range("Medical_Expenditure2_80D").Locked = True
        Sheet16.Range("Medical_Expenditure2_80D").Interior.Color = (&HD8D8D8)
        
        


  Else
'Ayush_Legacy9
'         Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "(Select)"
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Locked = False
        Sheet16.Range("DropDown_ValueOf_FamilyM_80D").Interior.Color = (&HCCFFCC)
        
'Ayush_Legacy9
'        Sheet16.Range("DropDown_ValueOf_SC_80D").value = "(Select)"
        Sheet16.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
        Sheet16.Range("DropDown_ValueOf_SC_80D").Locked = False
        Sheet16.Range("DropDown_ValueOf_SC_80D").Interior.Color = (&HCCFFCC)

          Sheet16.Range("Self_And_Family_80D").Locked = True
        Sheet16.Range("Self_And_Family_80D").Interior.Color = 13434879
'        sheet16.range("Self_And_Family_80D").Font.Color = RGB(0, 0, 0)
        Sheet16.Range("Self_And_Family_80D").Formula = "=MIN(25000,SUM(Health_Insurance2_80D,Preventive_Health2_80D))"

'        Sheet16.Range("Health_Insurance_80D").Locked = False
'        Sheet16.Range("Health_Insurance_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Preventive_Health_80D").Locked = False
        Sheet16.Range("Preventive_Health_80D").Interior.Color = (&HCCFFCC)


        Sheet16.Range("Senior_Citizen_80D").Locked = True
        Sheet16.Range("Senior_Citizen_80D").Interior.Color = 13434879
'        sheet16.range("Senior_Citizen_80D").Font.Color = RGB(0, 0, 0)
        Sheet16.Range("Senior_Citizen_80D").Formula = "=MIN(50000,SUM(Health_Insurance3_80D,Preventive_Health3_80D,Medical_Expenditure2_80D))"

'        Sheet16.Range("Health_InsuranceSC_80D").Locked = False
'        Sheet16.Range("Health_InsuranceSC_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Preventive_Health_SC_80D").Locked = False
        Sheet16.Range("Preventive_Health_SC_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Medical_Expenditure_SC_80D").Locked = False
        Sheet16.Range("Medical_Expenditure_SC_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Parents_80D").Locked = True
        Sheet16.Range("Parents_80D").Interior.Color = 13434879
'        sheet16.range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Sheet16.Range("Parents_80D").Formula = "=MIN(25000,SUM(Health_Insurance2_80D,Preventive_Health2_80D))"

'        Sheet16.Range("Health_Insurance2_80D").Locked = False
'        Sheet16.Range("Health_Insurance2_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Preventive_Health2_80D").Locked = False
        Sheet16.Range("Preventive_Health2_80D").Interior.Color = (&HCCFFCC)


        Sheet16.Range("Parents_SC_80D").Locked = False
        Sheet16.Range("Parents_SC_80D").Interior.Color = 13434879
'        sheet16.range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Sheet16.Range("Parents_SC_80D").Formula = "=MIN(50000,SUM(Health_Insurance3_80D,Preventive_Health3_80D,Medical_Expenditure2_80D))"

'        Sheet16.Range("Health_Insurance3_80D").Locked = False
'        Sheet16.Range("Health_Insurance3_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Preventive_Health3_80D").Locked = False
        Sheet16.Range("Preventive_Health3_80D").Interior.Color = (&HCCFFCC)

        Sheet16.Range("Medical_Expenditure2_80D").Locked = False
        Sheet16.Range("Medical_Expenditure2_80D").Interior.Color = (&HCCFFCC)
        
   
        
        
        
   End If
        

  
            Sheet16.Protect Password:=getmsgstate
            Application.EnableEvents = True
End Sub
'-----------------------------------------------------------------

Sub tesst()
Application.EnableEvents = True
End Sub


Sub LockUnlock80DU()
On Error Resume Next
Application.EnableEvents = False
Sheet14.Unprotect Password:=getmsgstate


If Sheet8.Range("oldbacValue").value = 2 Then
    If Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) = "H" Then
    
            Sheet14.Range("NatureDisability_80U").ClearContents
            Sheet14.Range("NatureDisability_80U").Locked = True
            Sheet14.Range("NatureDisability_80U").Interior.Color = "&HD8D8D8"
            
            'Ayush_14/05 SIT-92905
            Sheet14.Range("Disability_80U").ClearContents
            Sheet14.Range("Disability_80U").Locked = True
            Sheet14.Range("Disability_80U").Interior.Color = "&HD8D8D8"
           
    
    'Commented by Ayush_06/05/2025 for 0.7 changes
'            Sheet14.Range("DateFiling10IA_80U").ClearContents
'            Sheet14.Range("DateFiling10IA_80U").Locked = True
'            Sheet14.Range("DateFiling10IA_80U").Interior.Color = "&HD8D8D8"
'            Sheet14.Range("DateFiling10IA_80U").value = ""
'
    
            Sheet14.Range("AckNumber10IA_80U").ClearContents
            Sheet14.Range("AckNumber10IA_80U").Locked = True
            Sheet14.Range("AckNumber10IA_80U").Interior.Color = "&HD8D8D8"
            Sheet14.Range("AckNumber10IA_80U").value = ""
    
    
            Sheet14.Range("UDIDNum_80U").ClearContents
            Sheet14.Range("UDIDNum_80U").Locked = True
            Sheet14.Range("UDIDNum_80U").Interior.Color = "&HD8D8D8"
            Sheet14.Range("UDIDNum_80U").value = ""
    
    ElseIf Range("sheet1.status").value <> "F - FIRM(Other than LLP)" Then
    
    
            
            Sheet14.Range("NatureDisability_80U").Locked = False
            Sheet14.Range("NatureDisability_80U").Interior.Color = "&HCCFFCC"
            
            
            'Ayush_14/05 SIT-92905
             Sheet14.Range("Disability_80U").Locked = False
            Sheet14.Range("Disability_80U").Interior.Color = "&HCCFFCC"
    
  'Commented by Ayush_06/05/2025 for 0.7 changes
'            Sheet14.Range("DateFiling10IA_80U").Locked = False
'            Sheet14.Range("DateFiling10IA_80U").Interior.Color = "&HCCFFCC"
'
    
            Sheet14.Range("AckNumber10IA_80U").Locked = False
            Sheet14.Range("AckNumber10IA_80U").Interior.Color = "&HCCFFCC"
    
            
            Sheet14.Range("UDIDNum_80U").Locked = False
            Sheet14.Range("UDIDNum_80U").Interior.Color = "&HCCFFCC"
            
    
    
    End If

End If


Sheet14.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub



'Ayush_13/04
Sub LockUnlock80C()
On Error Resume Next
Application.EnableEvents = False
Sheet19.Unprotect Password:=getmsgstate


If Sheet8.Range("oldbacValue").value = 2 Then
    
   ' If Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) = "H" Or Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) = "F" Then
         ' If Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) <> "I" Then
'            Sheet19.Range("NaturePayment.80C").ClearContents
'            Sheet19.Range("NaturePayment.80C").Locked = True
'            Sheet19.Range("NaturePayment.80C").Interior.Color = "&HD8D8D8"
'
'
'
'            Sheet19.Range("Amount.80C").ClearContents
'            Sheet19.Range("Amount.80C").Locked = True
'            Sheet19.Range("Amount.80C").Interior.Color = "&HD8D8D8"
'            Sheet19.Range("Amount.80C").value = ""
'
'
'            Sheet19.Range("Identification_Number.80C").ClearContents
'            Sheet19.Range("Identification_Number.80C").Locked = True
'            Sheet19.Range("Identification_Number.80C").Interior.Color = "&HD8D8D8"
'            Sheet19.Range("Identification_Number.80C").value = ""



           If Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) = "F" Then
 'Commented by Ayush_06/05/2025 for 0.7 changes
'            Sheet19.Unprotect Password:=getmsgstate
'            Sheet19.Range("NaturePayment.80C").Locked = True
'            Sheet19.Range("NaturePayment.80C").Interior.Color = "&HD8D8D8"
    
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Amount.80C").Locked = True
            Sheet19.Range("Amount.80C").Interior.Color = "&HD8D8D8"
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Identification_Number.80C").Locked = True
            Sheet19.Range("Identification_Number.80C").Interior.Color = "&HD8D8D8"
    
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Name_of_insurer.").ClearContents
            Sheet19.Range("Name_of_insurer.").Locked = True
            Sheet19.Range("Name_of_insurer.").Interior.Color = "&HD8D8D8"
            Sheet19.Range("Name_of_insurer.").value = ""
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Policy_Document_Number.").ClearContents
            Sheet19.Range("Policy_Document_Number.").Locked = True
            Sheet19.Range("Policy_Document_Number.").Interior.Color = "&HD8D8D8"
            Sheet19.Range("Policy_Document_Number.").value = ""
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Amount.").ClearContents
            Sheet19.Range("Amount.").Locked = True
            Sheet19.Range("Amount.").Interior.Color = "&HD8D8D8"
            Sheet19.Range("Amount.").value = ""
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Visible = xlSheetHidden
            Sheet19.Protect Password:=getmsgstate
    
            
             
           ElseIf Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) = "H" Then
           
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Visible = xlSheetVisible
            Sheet19.Protect Password:=getmsgstate
            
            'Commented by Ayush_05/05/2025 for 0.7 changes
             
'            Sheet19.Unprotect Password:=getmsgstate
'            Sheet19.Range("NaturePayment.80C").Locked = False
'            Sheet19.Range("NaturePayment.80C").Interior.Color = "&HCCFFCC"
    
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Amount.80C").Locked = False
            Sheet19.Range("Amount.80C").Interior.Color = "&HCCFFCC"
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Identification_Number.80C").Locked = False
            Sheet19.Range("Identification_Number.80C").Interior.Color = "&HCCFFCC"
    
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Name_of_insurer.").ClearContents
            Sheet19.Range("Name_of_insurer.").Locked = True
            Sheet19.Range("Name_of_insurer.").Interior.Color = "&HD8D8D8"
            Sheet19.Range("Name_of_insurer.").value = ""
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Policy_Document_Number.").ClearContents
            Sheet19.Range("Policy_Document_Number.").Locked = True
            Sheet19.Range("Policy_Document_Number.").Interior.Color = "&HD8D8D8"
            Sheet19.Range("Policy_Document_Number.").value = ""

            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Amount.").ClearContents
            Sheet19.Range("Amount.").Locked = True
            Sheet19.Range("Amount.").Interior.Color = "&HD8D8D8"
            Sheet19.Range("Amount.").value = ""
            
            
            
           ' End If
             
    
    'ElseIf Range("sheet1.status").value <> "F - FIRM(Other than LLP)" Then
    ElseIf Mid(Sheet1.Range("Sheet1.Status").value, 1, 1) = "I" Then
    
    'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet19.Unprotect Password:=getmsgstate
'            Sheet19.Range("NaturePayment.80C").Locked = False
'            Sheet19.Range("NaturePayment.80C").Interior.Color = "&HCCFFCC"
    
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Amount.80C").Locked = False
            Sheet19.Range("Amount.80C").Interior.Color = "&HCCFFCC"
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Identification_Number.80C").Locked = False
            Sheet19.Range("Identification_Number.80C").Interior.Color = "&HCCFFCC"
    
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Name_of_insurer.").Locked = False
            Sheet19.Range("Name_of_insurer.").Interior.Color = "&HCCFFCC"
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Policy_Document_Number.").Locked = False
            Sheet19.Range("Policy_Document_Number.").Interior.Color = "&HCCFFCC"
            
            Sheet19.Unprotect Password:=getmsgstate
            Sheet19.Range("Amount.").Locked = False
            Sheet19.Range("Amount.").Interior.Color = "&HCCFFCC"
            
End If

            

    
    
    End If

'End If


Sheet19.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
Sub tesst1()
Application.EnableEvents = True
End Sub



Function StateMatchesPin(StateCode As Variant) As Boolean
    StateMatchesPin = False
    Dim state1, state2 As Variant
    
    
    state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, Sheet8.Range("Duplicate_V"), 2, False)
    state2 = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, Sheet8.Range("Duplicate_V"), 3, False)

    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin = False
    Exit Function
    End If
    
    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = state2 Then
    StateMatchesPin = True
    End If
    
End Function
'Ayush_25-26
Public Function ChkMaxDate_filing(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_filing = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        If Year > 2025 Then
            ChkMaxDate_filing = False
            Exit Function
        Else
            If Year = 2025 Then
                If month > 8 Then
                    ChkMaxDate_filing = False
                    Exit Function
                Else
                    If month = 8 Then
                        If dat > 1 Then
                           ChkMaxDate_filing = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_filing = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


'Ayush_25-26
Public Function ChkMaxDate_filing_2425(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

    ChkMaxDate_filing_2425 = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        If Year > 2024 Then
            ChkMaxDate_filing_2425 = False
            Exit Function
        Else
            If Year = 2024 Then
                If month > 8 Then
                    ChkMaxDate_filing_2425 = False
                    Exit Function
                Else
                    If month = 8 Then
                        If dat > 1 Then
                           ChkMaxDate_filing_2425 = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_filing_2425 = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


'Ayush_19
Sub LockUnlock80E()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" And Sheet8.Range("bacvalue").value = 2 Then
     ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet20.Visible = xlSheetHidden
    Sheet20.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate

    End If
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockUnlock24b()
On Error Resume Next
'Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    
    
 'Updated by Bindu on 23rd April 25
 If Sheet8.Range("bacvalue").value = 1 And UCase(Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1)) = "S" Then
    
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet21.Visible = xlSheetHidden
    'Sheet21.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    
ElseIf Sheet8.Range("bacvalue").value = 1 And UCase(Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1)) = "(" Then
    
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet21.Visible = xlSheetHidden
    'Sheet21.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    
 ElseIf Sheet8.Range("bacvalue").value = 1 And Sheet1.Range("IncD.TypeOfHP").value = "" Then
 
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet21.Visible = xlSheetHidden
    'Sheet21.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    

    
ElseIf Sheet8.Range("bacvalue").value = 2 And Sheet1.Range("IncD.TypeOfHP").value = "" Then
 
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet21.Visible = xlSheetHidden
    'Sheet21.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    
  ElseIf Sheet8.Range("bacvalue").value = 2 And UCase(Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1)) = "(" Then
    
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet21.Visible = xlSheetHidden
    'Sheet21.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
 
Else
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet21.Visible = xlSheetVisible
    'Sheet21.Visible = xlSheetVisible
    ThisWorkbook.Protect Password:=getmsgstate
 
 End If

'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet21.Visible = xlSheetVisible
'    'Sheet21.Visible = xlSheetVisible
'    ThisWorkbook.Protect Password:=getmsgstate
 



' ElseIf Sheet8.Range("bacvalue").value = 2 And UCase(Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1)) = "(" Or Sheet1.Range("IncD.TypeOfHP").value = "" Then
'
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet21.Visible = xlSheetHidden
'    'Sheet21.Visible = xlSheetHidden
'    ThisWorkbook.Protect Password:=getmsgstate
'
'Else
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet21.Visible = xlSheetVisible
'    'Sheet21.Visible = xlSheetVisible
'    ThisWorkbook.Protect Password:=getmsgstate
'
' End If
'     End If
'    If Mid(Range("IncD.TypeOfHP").value, 1, 1) <> "S" Then
'        If Sheet8.Range("bacvalue").value = 2 Then
'            'If Sheet8.Range("bacvalue").value = 2 Then
'            ThisWorkbook.Unprotect Password:=getmsgstate
'            Sheet21.Visible = xlSheetVisible
'            Sheet21.Visible = xlSheetVisible
'            ThisWorkbook.Protect Password:=getmsgstate
'        '    ThisWorkbook.Unprotect Password:=getmsgstate
'        '    Sheet21.Visible = xlSheetVisible
'        '    Sheet21.Visible = xlSheetVisible
'        '    ThisWorkbook.Protect Password:=getmsgstate
'        Else
'            ThisWorkbook.Unprotect Password:=getmsgstate
'            Sheet21.Visible = xlSheetHidden
'            Sheet21.Visible = xlSheetHidden
'            ThisWorkbook.Protect Password:=getmsgstate
'        '     ThisWorkbook.Unprotect Password:=getmsgstate
'        '    Sheet21.Visible = xlSheetHidden
'        '    Sheet21.Visible = xlSheetHidden
'        '    ThisWorkbook.Protect Password:=getmsgstate
'        '    ThisWorkbook.Unprotect Password:=getmsgstate
'        '    Sheet20.Visible = xlSheetVisible
'        '    Sheet20.Visible = xlSheetVisible
'        '    ThisWorkbook.Protect Password:=getmsgstate
'    End If
'    End If
    'Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub








Private Sub UpdateButtons_Sec80CCC()
    Dim v As Variant
    On Error Resume Next
    v = Range("IncD.Section80CCC").value
    On Error GoTo 0

    Dim shouldShow As Boolean
    shouldShow = (val(v) > 0)

    Dim btns As Collection
    Set btns = FindFormButtonsByMacro(Sheet1, "AddRows_Sec80CCC")

    If btns Is Nothing Or btns.count = 0 Then
        ' Nothing matched; exit quietly
        Exit Sub
    End If

    Dim shp As Shape
    For Each shp In btns
        If shouldShow Then
            shp.Visible = msoTrue
            shp.Locked = False
           Sheet1.Unprotect Password:=getmsgstate
        Else
            shp.Visible = msoFalse
            shp.Locked = True
            Sheet1.Protect Password:=getmsgstate
        End If
    Next shp
End Sub


Sub LockUnlock_80CCC()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
     
     
     If Sheet1.Range("IncD.Section80CCC").value = 0 Then
     
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Rows("158:158").Hidden = True
        
    Else
    
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Rows("158:158").Hidden = False
     
     End If
     
'    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
'
'    'Ayush_UR
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Rows("25:25").Hidden = False
'
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet201.Visible = xlSheetVisible
'    Sheet202.Visible = xlSheetVisible
'    ThisWorkbook.Protect Password:=getmsgstate
'    Else
'    'Ayush_UR
'    Sheet1.Unprotect Password:=getmsgstate
'     Sheet1.Rows("25:25").Hidden = True
'
'    ThisWorkbook.Unprotect Password:=getmsgstate
'    Sheet201.Visible = xlSheetHidden
'    Sheet202.Visible = xlSheetHidden
'    ThisWorkbook.Protect Password:=getmsgstate
'    End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockUnlock_Secondary_Address()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
     
     
     If Sheet1.Range("Secondary_Address").value = "Yes" Then
     
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = False
            Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HCCFFCC"
            
             
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = False
            Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = False
            Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = False
            Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = False
            Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("HASZIP2").MergeArea.Locked = False
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HCCFFCC"
        
    Else
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.ClearContents
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = True
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.ResidenceNo2").value = ""
        
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.ClearContents
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = True
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.ResidenceName2").value = ""
        
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.ClearContents
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = True
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.RoadOrStreet2").value = ""
        
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.ClearContents
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = True
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.LocalityOrArea2").value = ""
        
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.ClearContents
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = True
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").value = ""
        
        Sheet1.Range("sheet1.StateCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.StateCode2").value = ""
        
        Sheet1.Range("sheet1.CountryCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.CountryCode2").value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.PinCode2").value = ""
        
        Sheet1.Range("HASZIP2").MergeArea.ClearContents
        Sheet1.Range("HASZIP2").MergeArea.Locked = True
        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("HASZIP2").value = ""
        
        Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.ZipCode2").value = ""
        
        
     End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub


Sub DPMAddComment_PRAN(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="PRAN (Permanent Retirement Account Number) is a unique 12-digit number allotted to every subscriber of the National Pension System (NPS)."
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub

Sub DPMADeleteComment_PRAN(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next

Sheet1.Unprotect Password:=getmsgstate
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
Sheet1.Protect Password:=getmsgstate
End Sub

Sub ttt()
Application.EnableEvents = True
End Sub










