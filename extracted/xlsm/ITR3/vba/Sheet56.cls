Attribute VB_Name = "Sheet56"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
On Error Resume Next
'    ActiveSheet.Shapes("80GGA").Left = Range("Button_80GGA").Left + 5
'    ActiveSheet.Shapes("80GGA").Top = Range("Button_80GGA").Top + 5

    'ThisWorkbook.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo exit1
sPassword = EfilingCommon.getmsgstate
'Sheet56.Unprotect Password:=sPassword
'Sheet56.Range("taxcheck").value = ""
'Sheet56.Protect Password:=sPassword
    Application.EnableEvents = False
'    If Target.Validation.Type = 3 Then
'         GoTo exit1
'    End If

'    Target.value = UCase(Target.value)
Dim var As Integer


' If Target.Column = 10 Then
'
'        Dim PinCode As Range
'        For Each PinCode In SheetALL.Range("PincodeList")
'            Dim code As String
'            code = Left(CStr(Target.value), 3)
'            If CStr(PinCode.value) = CStr(code) Then
'                Range("I" & Target.row).value = SheetALL.Range("CX" & PinCode.row).value
'                Exit For
'            End If
'         Next PinCode
'    End If
    
    

    
    
    


 'By sai on 17/09/2024=================================================================
If Not Application.Intersect(Sheet56.Range("State_Code_80GGANew"), Sheet56.Range(Target.Address)) Is Nothing Then
    
    Dim PIN_targetadd, state_targetadd As String
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "I", "J")
    
    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet56.Range(Replace(state_targetadd, "I", "J")).value = ""
    
End If
 '==========================================================================================

'By sai on 17/09/2024=======================================================================
If Not Application.Intersect(Sheet56.Range("Pincode_80GGANew"), Sheet56.Range(Target.Address)) Is Nothing Then
 
     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
     
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "J", "I")
      Set ws = Worksheets("80GGA")
      
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
End If
    
 '============================================================================================



    
    
    If Target.Column = 11 Then
        If 1 < Application.CountIf(Range("PAN_of_donee_80GGANew"), Range("K" & Target.row).value) Then
                'MsgBox "* Duplicate Pan value is not allowed"
                'ramya 70568
                MsgBox "* ""Please enter aggregate amount of donation made in cash to same Donee once"""
                Range("K" & Target.row).value = ""
                Range("L" & Target.row).value = ""
                Range("M" & Target.row).value = ""
        End If
    End If

    If Target.Column = 12 Then
        If Not Range("K" & Target.row).value = "" Then
        var = Application.CountIf(Range("L1:L" & Rows.count), Range("L" & Target.row).value)
           If Range("L" & Target.row).value > 2000 Then
                  MsgBox "* If Donation in cash is more than Rs. 2000, 80GGA Deduction is not allowed"
            End If
        End If
    End If
     
 Sheet56.Protect Password:=sPassword
exit1:
Application.EnableEvents = True
Sheet56.Activate
End Sub
