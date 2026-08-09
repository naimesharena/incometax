Attribute VB_Name = "Sheet38"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Img_Home_Click()
    Sheet35.Activate
End Sub
Private Sub index_Click()
    Sheet35.Activate
End Sub

Private Sub Cmd_GenerateXML_Click()
Generate_XML.Gen_XML
End Sub



Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet38.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----


    Sheet38.Shapes("Button 1").Left = Range("BA.Button1").Left + 1
    Sheet38.Shapes("Button 1").Top = Range("BA.Button1").Top + 1
    
    Sheet38.Protect Password:=sPassword
endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo endline
    
    CalculateTaxCheckFlag = False
    
    Dim rangenm As Variant
    Dim TargtVal, targetadd As Variant
    Dim BalAmount As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
    
    If Target.Validation.Type = 1 Then
        Target.value = UCase(Target.value)
        For Each BalAmount In Range("SchBA.AccntBalance")
            If BalAmount < 0 Then
                If Len(BalAmount) > 15 Or Not IsNumeric(BalAmount) Then BalAmount = ""
                Else:
                    If Len(BalAmount) > 14 Or Not IsNumeric(BalAmount) Then BalAmount = ""
                End If
        Next
        
        If (Target.name.name) = "NumBankAccountsHeld" Then
            If Not IsNumeric(Target.value) Then
                fmsgboxStatus "Number of Bank Accounts held should be Numeric Value"
            End If
    
            If Target.value < 0 Then
                fmsgboxStatus "Number of Bank Accounts held cannot be negative Value"
            End If
        End If
    End If
    
    If Target.Validation.Type = 3 Then
        Dim E_id As String
        Dim Bankname As Range
        For Each Bankname In Range("SchBA.BankName")
            If Bankname = "Others" Then
                E_id = Application.InputBox("Enter the Bank Name")
                If E_id = "False" Then
                    Bankname = ""
                Else
                    Bankname = "Others-" & E_id
                End If
            End If
        Next
        
        Dim bn As Range
        For Each bn In Range("SchBA.BankName")
            If Len(bn) > 125 Then
               fmsgboxStatus ("Bank Name should not exceeds 125 characters")
                bn = ""
            End If
        Next
    
        Dim SourceAmount As Range
        Set SourceAmount = Range("SchBA.AcntStatus")
        If Not Application.Intersect(SourceAmount, Range(Target.Address)) Is Nothing Then
            Call CheckAccountStatus(TargtVal, targetadd)
            GoTo endline
        End If
    End If
    
    If Target.Validation.Type = 6 Then
        Target.value = UCase(Target.value)
        Dim BalAmunt As Range
        Set BalAmunt = Range("SchBA.AccntBalance")
                
        If Not Application.Intersect(BalAmunt, Range(Target.Address)) Is Nothing Then
            If ((Range(targetadd).Offset(0, -1).value = "Closed") And (TargtVal <> 0)) Then
                fmsgboxStatus "Please Recheck Your Account Status"
            End If
            GoTo endline
        End If
    End If
    
'----------------Bank Code---------


Dim BankNameFromIFSC As Variant
Dim Address As Variant
Dim rangecell As Range

If Not Application.Intersect(Range("SchBA.IFSC"), Range(Target.Address)) Is Nothing Then
    For Each rangecell In Sheet38.Range("SchBA.IFSC")
        If rangecell.value <> "" Then
            'If Not checkListIFSC(rangecell.value) Then
            If Not ValidateIFSCList(rangecell.value) Then
                fmsgboxStatus "* IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                Address = Replace(rangecell.Address, "B", "C")
                Sheet38.Range(Address).value = ""
                rangecell.value = ""
                GoTo endline
            Else
                Address = rangecell.Address
                Address = Replace(Address, "B", "C")
                BankNameFromIFSC = GetBankName(Mid(rangecell, 1, 4))
                Sheet38.Range(Address).value = BankNameFromIFSC
            End If
        End If
    Next
End If
endline:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
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

