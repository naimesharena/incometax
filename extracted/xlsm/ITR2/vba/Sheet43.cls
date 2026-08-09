Attribute VB_Name = "Sheet43"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False
    ActiveSheet.Shapes("Button 1").Left = Range("SchAL.Button1").Left
    ActiveSheet.Shapes("Button 1").Top = Range("SchAL.Button1").Top

    ActiveSheet.Shapes("Button 3").Left = Range("SchAL.Button3").Left
    ActiveSheet.Shapes("Button 3").Top = Range("SchAL.Button3").Top
    
endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False
 'Malli----------------
            Dim PinCode As Range
            Dim DupFlag As Boolean
            Dim StateName
            Dim StateCode
            Dim state1 As Variant
'-----------------------------
Dim rangenm As Variant
    Application.EnableEvents = False
    Dim IsImmovableAssetState, IsImmovableAssetStateCountry As Range
    Dim IsInterestHeldState, IsInterestHeldCountry As Range
    Dim TargtVal, targetadd As Variant
    
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
    
    Set IsImmovableAssetState = Range("SchAL.A.Address_State")
    Set IsImmovableAssetStateCountry = Range("SchAL.A.Address_Country")
    
    Set IsInterestHeldState = Range("SchAL.C.Address_State")
    Set IsInterestHeldCountry = Range("SchAL.C.Address_Country")
    
    If Not checkfieldSuperSpecialcharacter(Target.value) Then
        fmsgboxStatus ("Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
    End If
    
    
    If Target.Validation.Type <> 3 Then
        Target.value = UCase(Target.value)
    End If
    
    If Target.Validation.Type = 3 Then
            
    If Not Application.Intersect(IsImmovableAssetState, Range(Target.Address)) Is Nothing Then
           ' Call SetStateCountryImmovableAsset(TargtVal, targetadd, "S")
            Call SetStateCountryImmovableAsset_M(TargtVal, targetadd, "S")

'Malli-------
If Not Application.Intersect(Sheet43.Range("SchAL.A.Address_State"), Sheet43.Range(Target.Address)) Is Nothing Then
    
'    Dim PIN_targetadd, state_targetadd As Variant
'      state_targetadd = Target.Address
'      state_targetadd = Replace(state_targetadd, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "K", "M")
'
'    Call state_Validation(PIN_targetadd, state_targetadd)

Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "K", "M")
      
      Dim ws1 As Worksheet
                             Set ws1 = Worksheets("AL")  '23/09/24
                            
    
    If state_Validation(ws1, PIN_targetadd, state_targetadd) = False Then Sheet43.Range(Replace(state_targetadd, "K", "M")).value = ""
    
End If
'-------------

            GoTo endline
        End If
        
        If Not Application.Intersect(IsImmovableAssetStateCountry, Range(Target.Address)) Is Nothing Then
            Call SetStateCountryImmovableAsset(TargtVal, targetadd, "C")
            GoTo endline
        End If
    
        If Not Application.Intersect(IsInterestHeldState, Range(Target.Address)) Is Nothing Then
             Call SetStateCountryInterestHeldState(TargtVal, targetadd, "S")
             GoTo endline
        End If
        If Not Application.Intersect(IsInterestHeldCountry, Range(Target.Address)) Is Nothing Then
             Call SetStateCountryInterestHeldState(TargtVal, targetadd, "C")
             GoTo endline
        End If
            
        If (Target.name.name) = "IsImmovableAsset" Then
            If Mid(Trim(Range("IsImmovableAsset")), 1, 1) = "N" Then
                LOCKAL
            Else
                UNLOCKAL
            End If
        End If
        
        
'        If (Target.name.name) = "IsInterestHeld" Then
'            If Mid(Trim(Range("IsInterestHeld")), 1, 1) = "N" Then
'                LOCKAL_1
'            Else
'                UNLOCKAL_1
'            End If
'        End If
 
    End If
    
    'Malli comented--------------------------------
    
'  If Not Application.Intersect(Sheet43.Range("SchAL.A.Address_Pin"), Sheet43.Range(Target.Address)) Is Nothing Then
'         Dim PinCode1 As Range
'         Dim DupFlag As Boolean
'         DupFlag = False
'                 For Each PinCode1 In SheetALL.Range("Duplicate_Pincode_list")
'                    If Target.value = PinCode1 Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next PinCode1
'                 If DupFlag = True Then
'                    If Sheet43.Range(Replace(Target.Address, "M", "K")).value <> "" Then
'                   ' MsgBox Sheet43.Range(Replace(Target.Address, "M", "K")).value
'                   Dim Na1 As Variant
'                   Na1 = Target.value
'                        If Not StateMatchesPin1(Sheet43.Range(Replace(Target.Address, "M", "K")).value, Na1) Then
'                            Sheet43.Range(Replace(Target.Address, "M", "K")).value = "(Select)"
'                        End If
'                    End If
'                 End If
'                 If DupFlag = False Then
'                    Dim StateName
'                    Dim StateCode
'                    StateName = Application.VLookup(Sheet43.Range(Target.Address).value, SheetALL.Range("All_PinCode_V"), 2, False)
'                   'Added by Shrutika
'                   If Target.value <> "" Then
'                    If IsError(StateName) Then
'                        MsgBox ("Invalid Pincode")
'                        Application.EnableEvents = False
'                        Target.value = ""
'                        Application.EnableEvents = True
'                    Else
'                          StateCode = Application.VLookup(StateName, SheetALL.Range("EB:EC"), 2, False)
'                          Sheet43.Range(Replace(Target.Address, "M", "K")).value = StateCode
'                          Sheet43.Range(Replace(Target.Address, "M", "L")).value = "91-INDIA"
'
'                    End If
'                 End If
'          End If
'   End If
'
   
   
 
 
   
   
    'Malli--------------------------------------------
'If Not Application.Intersect(Sheet43.Range("SchAL.A.Address_State"), Sheet43.Range(Target.Address)) Is Nothing Then
'
'    Dim PIN_targetadd, state_targetadd As Variant
'      state_targetadd = Target.Address
'      state_targetadd = Replace(state_targetadd, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "K", "M")
'
'    Call state_Validation(PIN_targetadd, state_targetadd)
'
'End If

'------------------------------------------------------------------
   
   'Malli---------------------------
 
If Not Application.Intersect(Sheet43.Range("SchAL.A.Address_Pin"), Sheet43.Range(Target.Address)) Is Nothing Then
'
'     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As Variant
'
'      pinrng_chk = Target.Address
'      pinrng_chk = Replace(pinrng_chk, "$", "")
'      staterng_chk = Replace(pinrng_chk, "M", "K")
'      cntryrng_chk = Replace(pinrng_chk, "M", "L")
'      ZIPrng_chk = Replace(pinrng_chk, "M", "N")
'
'      Call PinState_codes_validation(pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)

 Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
      'pinrng_chk = Target.address
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "M", "K")
      cntryrng_chk = Replace(pinrng_chk, "M", "L")
      ZIPrng_chk = Replace(pinrng_chk, "M", "N")
      
      Set ws = Worksheets("AL")
      
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
End If

'----------------------------------------------------

   
 '---------------------------------------------------------------------
endline:
    Application.EnableEvents = True
End Sub

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
 Function StateMatchesPin1(StateCode As Variant, cd As Variant) As Boolean
 
 
    StateMatchesPin = False
    Dim state1, state2 As Variant

    state1 = Application.VLookup(cd, SheetALL.Range("Duplicate_V"), 2, False)
    state2 = Application.VLookup(cd, SheetALL.Range("Duplicate_V"), 3, False)
 
    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin = False
    Exit Function
    End If
    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = state2 Then
    StateMatchesPin = True
    End If
    End Function


