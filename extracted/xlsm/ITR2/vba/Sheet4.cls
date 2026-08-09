Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String


Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet4.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet4.Shapes("Button 13").Left = Range("PANButton1").Left + 2
Sheet4.Shapes("Button 13").Top = Range("PANButton1").Top + 6

Sheet4.Shapes("Button 98").Left = Range("Button44AE").Left + 2
Sheet4.Shapes("Button 98").Top = Range("Button44AE").Top + 6

Sheet4.Shapes("Button 46").Left = Range("NatureButton2").Left + 2
Sheet4.Shapes("Button 46").Top = Range("NatureButton2").Top + 6

Sheet4.Shapes("Button 47").Left = Range("NatureButton3").Left + 2
Sheet4.Shapes("Button 47").Top = Range("NatureButton3").Top + 6

Sheet4.Shapes("Button 69").Left = Range("PL.Button1").Left + 2
Sheet4.Shapes("Button 69").Top = Range("PL.Button1").Top + 6

 Sheet4.Protect Password:=sPassword
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
On Error GoTo endline

CalculateTaxCheckFlag = False

        If Not Intersect(Target, Range("Sec44AE.RegNo")) Is Nothing Then
             Dim var As Integer
             var = Application.CountIf(Range("Sec44AE.RegNo"), Target.value)
             If 1 < var Then
                'fmsgbox "* Duplicate Registration No. of goods carriage is not allowed"
                fmsgbox "* Same registration no. cannot be entered more than one time."
                Range(Target.Address).value = ""
            End If
         End If

Dim IsImmovableAssetState, IsImmovableAssetStateCountry As Range
        Dim TargtVal, targetadd As Variant
    
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
    
    Set IsImmovableAssetState = Range("PL_Address_State")
    Set IsImmovableAssetStateCountry = Range("PL_Address_Country")
    
    If Target.Validation.Type = 3 Then
            
    If Not Application.Intersect(IsImmovableAssetState, Range(Target.Address)) Is Nothing Then
            Call SetStateCountryImmovableAsset1(TargtVal, targetadd, "S")
            GoTo endline
        End If
        
        If Not Application.Intersect(IsImmovableAssetStateCountry, Range(Target.Address)) Is Nothing Then
            Call SetStateCountryImmovableAsset1(TargtVal, targetadd, "C")
            GoTo endline
        End If
  End If


 If (Target.name.name) = "PL.AnyCompPaidToNonRes" Then Call LockUnlock("PL.AnyCompPaidToNonRes", "PL.AmtPaidToNonRes")


 '-------------------------------------
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant
    rangenamestring = "PLBD.PAN||Nature_Name3||Nature_Name2"

    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet4.Range(rangenamestring(i))
            chcell.value = UCase(chcell.value)
            'chCell.Interior.Color = (&HD8D8D8)
            Next
        Next
        
        

endline:
Application.EnableEvents = True


End Sub





Sub LockUnlock(Src As String, Dst As String)

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----


    If Range(Src) = "No" Then
        Range(Dst).Select
        Selection.ClearContents
        Selection.Locked = True
        Selection.Interior.Color = (&HD8D8D8)

    Else
        Range(Dst).Select
        Selection.ClearContents
        Selection.Locked = False
        Selection.Interior.Color = (&HCCFFCC)
    End If

'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
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

