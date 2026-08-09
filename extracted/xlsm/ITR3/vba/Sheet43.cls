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
'Konda updated on 14-04-2026--SIT-117436
'        target.value = UCase(target.value)
        Target.value = WorksheetFunction.Clean(UCase(Target.value))
    End If
    
    If Target.Validation.Type = 3 Then
            
    If Not Application.Intersect(IsImmovableAssetState, Range(Target.Address)) Is Nothing Then
            'Call SetStateCountryImmovableAsset(TargtVal, TargetAdd, "S")
             Call SetStateCountryImmovableAsset_M(TargtVal, targetadd, "S")
             
'By sai on 17/09/2024=================================================================================
If Not Application.Intersect(Sheet43.Range("SchAL.A.Address_State"), Sheet43.Range(Target.Address)) Is Nothing Then


Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "K", "M")
      
      
    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet43.Range(Replace(state_targetadd, "K", "M")).value = ""
    
End If
'-------------

            GoTo endline
        End If
        
        If Not Application.Intersect(IsImmovableAssetStateCountry, Range(Target.Address)) Is Nothing Then
            Call SetStateCountryImmovableAsset(TargtVal, targetadd, "C")
            GoTo endline
        End If
    
        If Not Application.Intersect(IsInterestHeldState, Range(Target.Address)) Is Nothing Then
             
             'Call SetStateCountryInterestHeldState(TargtVal, TargetAdd, "S")
              Call SetStateCountryImmovableAsset_M(TargtVal, targetadd, "S")

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
        
        
        If (Target.name.name) = "IsInterestHeld" Then
            If Mid(Trim(Range("IsInterestHeld")), 1, 1) = "N" Then
                LOCKAL_1
            Else
                UNLOCKAL_1
            End If
        End If
 
    End If
    
  
 
If Not Application.Intersect(Sheet43.Range("SchAL.A.Address_Pin"), Sheet43.Range(Target.Address)) Is Nothing Then


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

'===================================================================================================
    
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

