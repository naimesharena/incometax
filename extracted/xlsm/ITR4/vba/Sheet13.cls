Attribute VB_Name = "Sheet13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Private Sub Worksheet_Change(ByVal Target As Range)
'
'
'If Not Application.Intersect(Sheet13.Range("SchAL.A.Address_State"), Sheet13.Range(Target.Address)) Is Nothing Then
'
'    Dim PIN_targetadd, state_targetadd As String
'      'state_targetadd = Target.address
'      state_targetadd = Replace(Target.Address, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "K", "M")
'
'
'     If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet13.Range(Replace(state_targetadd, "K", "M")).value = ""
'
'End If
'
'If Not Application.Intersect(Sheet13.Range("SchAL.A.Address_Pin"), Sheet13.Range(Target.Address)) Is Nothing Then
'
'     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
'     Dim ws As Worksheet
'      'pinrng_chk = Target.address
'      pinrng_chk = Replace(Target.Address, "$", "")
'      staterng_chk = Replace(pinrng_chk, "M", "K")
'      cntryrng_chk = Replace(pinrng_chk, "M", "L")
'      ZIPrng_chk = Replace(pinrng_chk, "M", "N")
'
'      Set ws = Worksheets("AL")
'
'      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
'End If
'
'End Sub

'Private Sub Worksheet_Activate()
'On Error GoTo endline
'Application.EnableEvents = False
'    ActiveSheet.Shapes("Button 1").Left = Range("SchAL.Button1").Left
'    ActiveSheet.Shapes("Button 1").Top = Range("SchAL.Button1").Top
'
'    ActiveSheet.Shapes("Button 3").Left = Range("SchAL.Button3").Left
'    ActiveSheet.Shapes("Button 3").Top = Range("SchAL.Button3").Top
'
'endline:
'Application.EnableEvents = True
'End Sub
'
'Private Sub Worksheet_Change(ByVal Target As Range)
'On Error GoTo endline
'Dim rangenm As Variant
'    Application.EnableEvents = False
'
'    CalculateTaxCheckFlag = False
'
'
'    If Not checkfieldSuperSpecialcharacter(Target.value) Then
'        MsgBox ("Please do not enter Special characters < , >" & Chr(13))
'        Target.value = ""
'    End If
'
'
'    If Target.Validation.Type <> 3 Then
'        Target.value = UCase(Target.value)
'    End If
'
''
''        '----------------Unlock Password-------------------START---
''   sPassword = mdCommon.getmsgstate
''   ActiveSheet.Unprotect Password:=sPassword
'''----------------Unlock Password-------------------END-----
'
''Target.value = UCase(Target.value)
'   If Target.Validation.Type = 3 Then
'
'   If Not Intersect(Range("SchAL.A.Address_State"), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=mdCommon.getmsgstate
'            If isdropdownblank(Range(Target.Address).value) Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "M")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "N")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "K", "M")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Locked = True
'
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "N")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "91-INDIA"
'
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "M")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "K", "N")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Locked = True
'            End If
'            Sheet13.Protect Password:=mdCommon.getmsgstate
'        End If
'
'        If Not Intersect(Range("SchAL.C.Address_State"), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=mdCommon.getmsgstate
'            If isdropdownblank(Range(Target.Address).value) Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "M")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "N")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "K", "M")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Locked = True
'
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "N")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Locked = False
'            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "91-INDIA"
'
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HCCFFCC)
'                Sheet13.Range(Replace(Target.Address, "K", "M")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "M")).Locked = False
'
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HD8D8D8)
'                Sheet13.Range(Replace(Target.Address, "K", "N")).value = ""
'                Sheet13.Range(Replace(Target.Address, "K", "N")).Locked = True
'            End If
'            Sheet13.Protect Password:=mdCommon.getmsgstate
'        End If
'
'    If Not Intersect(Range("SchAL.A.Address_State"), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=mdCommon.getmsgstate
'            If isdropdownblank(Range(Target.Address).value) Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
'                 Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'
'            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "91-INDIA"
'            End If
'            Sheet13.Protect Password:=mdCommon.getmsgstate
'            GoTo endline
'        End If
'
'        If Not Intersect(Range("SchAL.C.Address_State"), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=mdCommon.getmsgstate
'            If isdropdownblank(Range(Target.Address).value) Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'            ElseIf Mid(Range(Target.Address).value, 1, 2) = "99" Then
'                 Sheet13.Range(Replace(Target.Address, "K", "L")).value = "(Select)"
'
'            ElseIf Mid(Range(Target.Address).value, 1, 2) <> "99" Then
'                Sheet13.Range(Replace(Target.Address, "K", "L")).value = "91-INDIA"
'            End If
'            Sheet13.Protect Password:=mdCommon.getmsgstate
'            GoTo endline
'        End If
'
'        If Not Intersect(Range("SchAL.A.Address_Country"), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=mdCommon.getmsgstate
'            If Not isdropdownblank(Range(Target.Address).value) Then
'                If Range(Target.Address).value = "91-INDIA" Then
'                    If Sheet13.Range(Replace(Target.Address, "L", "K")).value = "99-Foreign" Then
'                        Sheet13.Range(Target.Address).value = "(Select)"
'                        MsgBox "Personal Information cannot be outside India, as you have selected country Ð India."
'                     End If
'                End If
'            End If
'
'            If Not isdropdownblank(Range(Target.Address).value) Then
'                If (Sheet13.Range(Replace(Target.Address, "L", "K")).value <> "99-Foreign" And Sheet13.Range(Replace(Target.Address, "L", "K")).value <> "(Select)" And Sheet13.Range(Replace(Target.Address, "L", "K")).value <> "") Then
'                    If Sheet13.Range(Target.Address).value <> "91-INDIA" And Sheet13.Range(Target.Address).value <> "(Select)" Then
'                        Sheet13.Range(Target.Address).value = "(Select)"
'                        MsgBox "Country cannot be other than India for the selected State"
'                    End If
'                End If
'            End If
'            Sheet13.Protect Password:=mdCommon.getmsgstate
'            GoTo endline
'         End If
'
'         If Not Intersect(Range("SchAL.C.Address_Country"), Range(Target.Address)) Is Nothing Then
'            Sheet13.Unprotect Password:=mdCommon.getmsgstate
'            If Not isdropdownblank(Range(Target.Address).value) Then
'                If Range(Target.Address).value = "91-INDIA" Then
'                    If Sheet13.Range(Replace(Target.Address, "L", "K")).value = "99-Foreign" Then
'                        Sheet13.Range(Target.Address).value = "(Select)"
'                        MsgBox "Personal Information cannot be outside India, as you have selected country Ð India."
'                    End If
'                End If
'            End If
'
'            If Not isdropdownblank(Range(Target.Address).value) Then
'                If (Sheet13.Range(Replace(Target.Address, "L", "K")).value <> "99-Foreign" And Sheet13.Range(Replace(Target.Address, "L", "K")).value <> "(Select)" And Sheet13.Range(Replace(Target.Address, "L", "K")).value <> "") Then
'                    If Sheet13.Range(Target.Address).value <> "91-INDIA" And Sheet13.Range(Target.Address).value <> "(Select)" Then
'                        Sheet13.Range(Target.Address).value = "(Select)"
'                        MsgBox "Country cannot be other than India for the selected State"
'                    End If
'                End If
'            End If
'            Sheet13.Protect Password:=mdCommon.getmsgstate
'            GoTo endline
'         End If
'
'    End If
'
'
'
'    If (Target.name.name) = "IsImmovableAsset" Then
'        If Mid(Trim(Range("IsImmovableAsset")), 1, 1) = "N" Then
'            LOCKAL
'        ElseIf Mid(Trim(Range("IsImmovableAsset")), 1, 2) = "(S" Then
'            UNLOCKAL
'        Else
'        UNLOCKAL1
'        End If
'    End If
'
'
'    If (Target.name.name) = "IsInterestHeld" Then
'        If Mid(Trim(Range("IsInterestHeld")), 1, 1) = "N" Then
'            LOCKAL_1
'        Else
'            UNLOCKAL_1
'        End If
'    End If
'
'
'
'
'endline:
'    Application.EnableEvents = True
'End Sub
'
'
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
