Attribute VB_Name = "mdESR"
Option Explicit

Public AmtDebPL_ESR1i As Variant
Public AmtDebPL_ESR1iia As Variant
Public AmtUs35Allowable_ESR1iia As Variant
Public ExcessAmtOverDebPL_ESR1iia As Variant
Public AmtDebPL_ESR35CCC As Variant
Public AmtUs35Allowable_ESR35CCC As Variant
Public ExcessAmtOverDebPL_ESR35CCC As Variant
Public AmtDebPL_ESR35CCD As Variant
Public AmtUs35Allowable_ESR35CCD As Variant
Public ExcessAmtOverDebPL_ESR35CCD As Variant

Public AmtUs35Allowable_ESR1i As Variant
Public ExcessAmtOverDebPL_ESR1i As Variant
Public AmtDebPL_ESR1ii As Variant
Public AmtUs35Allowable_ESR1ii As Variant
Public ExcessAmtOverDebPL_ESR1ii As Variant
Public AmtDebPL_ESR1iii As Variant
Public AmtUs35Allowable_ESR1iii As Variant
Public ExcessAmtOverDebPL_ESR1iii As Variant
Public AmtDebPL_ESR1iv As Variant
Public AmtUs35Allowable_ESR1iv As Variant
Public ExcessAmtOverDebPL_ESR1iv As Variant
Public AmtDebPL_ESR2AA As Variant
Public AmtUs35Allowable_ESR2AA As Variant
Public ExcessAmtOverDebPL_ESR2AA As Variant
Public AmtDebPL_ESR2AB As Variant
Public AmtUs35Allowable_ESR2AB As Variant
Public ExcessAmtOverDebPL_ESR2AB As Variant
Public AmtDebPL_ESRTOT As Variant
Public AmtUs35Allowable_ESRTOT As Variant
Public ExcessAmtOverDebPL_ESRTOT As Variant

Dim msgValidateSheetESR1i As Variant

Sub validateESRClick()
ValidateSheetESR
fmsgboxStatus "Sheet ESR is OK"
End Sub

Sub ValidateSheetESR()
    If Not ValidatesheetESR1i Then
        Sheet12.Activate
        fmsgboxsmall msgValidateSheetESR1i ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub

Function ValidatesheetESR1i() As Boolean
     ValidatesheetESR1i = True
     
     If Not ValidateAmtDebPL_ESR1i() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESR1i() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESR1i() Then ValidatesheetESR1i = False
     If Not ValidateAmtDebPL_ESR1ii() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESR1ii() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESR1ii() Then ValidatesheetESR1i = False
     If Not ValidateAmtDebPL_ESR1iii() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESR1iii() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESR1iii() Then ValidatesheetESR1i = False
     If Not ValidateAmtDebPL_ESR1iv() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESR1iv() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESR1iv() Then ValidatesheetESR1i = False
     If Not ValidateAmtDebPL_ESR2AA() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESR2AA() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESR2AA() Then ValidatesheetESR1i = False
     If Not ValidateAmtDebPL_ESR2AB() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESR2AB() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESR2AB() Then ValidatesheetESR1i = False
     If Not ValidateAmtDebPL_ESRTOT() Then ValidatesheetESR1i = False
     If Not ValidateAmtUs35Allowable_ESRTOT() Then ValidatesheetESR1i = False
     If Not ValidateExcessAmtOverDebPL_ESRTOT() Then ValidatesheetESR1i = False
End Function

Function ValidateAmtDebPL_ESR1i() As Boolean
    ValidateAmtDebPL_ESR1i = True
    
    AmtDebPL_ESR1i = Sheet12.Range("ESR1i.AmtDebPL").value
        
    AmtDebPL_ESR1iia = Sheet12.Range("ESR1iia.AmtDebPL").value
    AmtUs35Allowable_ESR1iia = Sheet12.Range("ESR1iia.AmtUs35Allowable").value
    ExcessAmtOverDebPL_ESR1iia = Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").value
    AmtDebPL_ESR35CCC = Sheet12.Range("ESR35CCC.AmtDebPL").value
    AmtUs35Allowable_ESR35CCC = Sheet12.Range("ESR35CCC.AmtUs35Allowable").value
    ExcessAmtOverDebPL_ESR35CCC = Sheet12.Range("ESR35CCC.ExcessAmtOverDebPL").value
    AmtDebPL_ESR35CCD = Sheet12.Range("ESR35CCD.AmtDebPL").value
    AmtUs35Allowable_ESR35CCD = Sheet12.Range("ESR35CCD.AmtUs35Allowable").value
    ExcessAmtOverDebPL_ESR35CCD = Sheet12.Range("ESR35CCD.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtUs35Allowable_ESR1i() As Boolean
    ValidateAmtUs35Allowable_ESR1i = True
    AmtUs35Allowable_ESR1i = Sheet12.Range("ESR1i.AmtUs35Allowable").value
End Function

Function ValidateExcessAmtOverDebPL_ESR1i() As Boolean
    ValidateExcessAmtOverDebPL_ESR1i = True
    ExcessAmtOverDebPL_ESR1i = Sheet12.Range("ESR1i.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtDebPL_ESR1ii() As Boolean
    ValidateAmtDebPL_ESR1ii = True
    AmtDebPL_ESR1ii = Sheet12.Range("ESR1ii.AmtDebPL").value
End Function

Function ValidateAmtUs35Allowable_ESR1ii() As Boolean
    ValidateAmtUs35Allowable_ESR1ii = True
    AmtUs35Allowable_ESR1ii = Sheet12.Range("ESR1ii.AmtUs35Allowable").value
End Function

Function ValidateExcessAmtOverDebPL_ESR1ii() As Boolean
    ValidateExcessAmtOverDebPL_ESR1ii = True
    ExcessAmtOverDebPL_ESR1ii = Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtDebPL_ESR1iii() As Boolean
    ValidateAmtDebPL_ESR1iii = True
    AmtDebPL_ESR1iii = Sheet12.Range("ESR1iii.AmtDebPL").value
End Function

Function ValidateAmtUs35Allowable_ESR1iii() As Boolean
    ValidateAmtUs35Allowable_ESR1iii = True
    AmtUs35Allowable_ESR1iii = Sheet12.Range("ESR1iii.AmtUs35Allowable").value
End Function

Function ValidateExcessAmtOverDebPL_ESR1iii() As Boolean
    ValidateExcessAmtOverDebPL_ESR1iii = True
    ExcessAmtOverDebPL_ESR1iii = Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtDebPL_ESR1iv() As Boolean
    ValidateAmtDebPL_ESR1iv = True
    AmtDebPL_ESR1iv = Sheet12.Range("ESR1iv.AmtDebPL").value
End Function

Function ValidateAmtUs35Allowable_ESR1iv() As Boolean
    ValidateAmtUs35Allowable_ESR1iv = True
    AmtUs35Allowable_ESR1iv = Sheet12.Range("ESR1iv.AmtUs35Allowable").value
End Function

Function ValidateExcessAmtOverDebPL_ESR1iv() As Boolean
    ValidateExcessAmtOverDebPL_ESR1iv = True
    ExcessAmtOverDebPL_ESR1iv = Sheet12.Range("ESR1iv.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtDebPL_ESR2AA() As Boolean
    ValidateAmtDebPL_ESR2AA = True
    AmtDebPL_ESR2AA = Sheet12.Range("ESR2AA.AmtDebPL").value
End Function

Function ValidateAmtUs35Allowable_ESR2AA() As Boolean
    ValidateAmtUs35Allowable_ESR2AA = True
    AmtUs35Allowable_ESR2AA = Sheet12.Range("ESR2AA.AmtUs35Allowable").value
End Function

Function ValidateExcessAmtOverDebPL_ESR2AA() As Boolean
    ValidateExcessAmtOverDebPL_ESR2AA = True
    ExcessAmtOverDebPL_ESR2AA = Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtDebPL_ESR2AB() As Boolean
    ValidateAmtDebPL_ESR2AB = True
    AmtDebPL_ESR2AB = Sheet12.Range("ESR2AB.AmtDebPL").value
End Function

Function ValidateAmtUs35Allowable_ESR2AB() As Boolean
    ValidateAmtUs35Allowable_ESR2AB = True
    AmtUs35Allowable_ESR2AB = Sheet12.Range("ESR2AB.AmtUs35Allowable").value
End Function

Function ValidateExcessAmtOverDebPL_ESR2AB() As Boolean
    ValidateExcessAmtOverDebPL_ESR2AB = True
    ExcessAmtOverDebPL_ESR2AB = Sheet12.Range("ESR2AB.ExcessAmtOverDebPL").value
End Function

Function ValidateAmtDebPL_ESRTOT() As Boolean
    ValidateAmtDebPL_ESRTOT = True
    AmtDebPL_ESRTOT = Sheet12.Range("ESRTOT.AmtDebPL").value
    
    If Len(AmtDebPL_ESRTOT) > 14 Then msgValidateSheetESR1i = msgValidateSheetESR1i & "* Total cannot exceed 14 digits" & Chr(13)
    
End Function

Function ValidateAmtUs35Allowable_ESRTOT() As Boolean
    ValidateAmtUs35Allowable_ESRTOT = True
    AmtUs35Allowable_ESRTOT = Sheet12.Range("ESRTOT.AmtUs35Allowable").value
    
    If Len(AmtUs35Allowable_ESRTOT) > 14 Then msgValidateSheetESR1i = msgValidateSheetESR1i & "* Total cannot exceed 14 digits" & Chr(13)
    
End Function

Function ValidateExcessAmtOverDebPL_ESRTOT() As Boolean
    ValidateExcessAmtOverDebPL_ESRTOT = True
    ExcessAmtOverDebPL_ESRTOT = Sheet12.Range("ESRTOT.ExcessAmtOverDebPL").value
    If Len(ExcessAmtOverDebPL_ESRTOT) > 14 Then
        msgValidateSheetESR1i = "* Total Amount in Sch ESR cannot exceed 14 digits"
        ValidateExcessAmtOverDebPL_ESRTOT = False
        Exit Function
    End If
End Function

Function DefaultAmtDebPL_ESRTOT() As String
    DefaultAmtDebPL_ESRTOT = "0"
End Function

Function DefaultAmtUs35Allowable_ESRTOT() As String
    DefaultAmtUs35Allowable_ESRTOT = "0"
End Function

Function DefaultExcessAmtOverDebPL_ESRTOT() As String
    DefaultExcessAmtOverDebPL_ESRTOT = "0"
End Function

Function DefaultAmtDebPL_ESR1i() As String
DefaultAmtDebPL_ESR1i = "0"
End Function
Function DefaultAmtUs35Allowable_ESR1i() As String
DefaultAmtUs35Allowable_ESR1i = "0"
End Function
Function DefaultExcessAmtOverDebPL_ESR1i() As String
DefaultExcessAmtOverDebPL_ESR1i = "0"
End Function
Function DefaultAmtDebPL_ESR1ii() As String
DefaultAmtDebPL_ESR1ii = "0"
End Function
Function DefaultAmtUs35Allowable_ESR1ii() As String
DefaultAmtUs35Allowable_ESR1ii = "0"
End Function
Function DefaultExcessAmtOverDebPL_ESR1ii() As String
DefaultExcessAmtOverDebPL_ESR1ii = "0"
End Function
Function DefaultAmtDebPL_ESR1iii() As String
DefaultAmtDebPL_ESR1iii = "0"
End Function
Function DefaultAmtUs35Allowable_ESR1iii() As String
DefaultAmtUs35Allowable_ESR1iii = "0"
End Function
Function DefaultExcessAmtOverDebPL_ESR1iii() As String
DefaultExcessAmtOverDebPL_ESR1iii = "0"
End Function
Function DefaultAmtDebPL_ESR1iv() As String
DefaultAmtDebPL_ESR1iv = "0"
End Function
Function DefaultAmtUs35Allowable_ESR1iv() As String
DefaultAmtUs35Allowable_ESR1iv = "0"
End Function
Function DefaultExcessAmtOverDebPL_ESR1iv() As String
DefaultExcessAmtOverDebPL_ESR1iv = "0"
End Function
Function DefaultAmtDebPL_ESR2AA() As String
DefaultAmtDebPL_ESR2AA = "0"
End Function
Function DefaultAmtUs35Allowable_ESR2AA() As String
DefaultAmtUs35Allowable_ESR2AA = "0"
End Function
Function DefaultExcessAmtOverDebPL_ESR2AA() As String
DefaultExcessAmtOverDebPL_ESR2AA = "0"
End Function
Function DefaultAmtDebPL_ESR2AB() As String
DefaultAmtDebPL_ESR2AB = "0"
End Function
Function DefaultAmtUs35Allowable_ESR2AB() As String
DefaultAmtUs35Allowable_ESR2AB = "0"
End Function
Function DefaultExcessAmtOverDebPL_ESR2AB() As String
DefaultExcessAmtOverDebPL_ESR2AB = "0"
End Function


