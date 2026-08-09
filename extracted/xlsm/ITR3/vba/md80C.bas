Attribute VB_Name = "md80C"
'module addded by Chetan C M for AY 2025-26 on 26/05/2025
'start--
Public MsgBox_80C As Variant
Dim end_80C As Variant
'Dim end_80CNaturePayment As Variant
Dim end_80CIdentification_Number As Variant
Dim end_80CAmount As Variant
Sub ValidateSheet80C_Click()
Dim vbMessgaeCaption As String
'Dim vbMessgaeCaption As String
vbMessgaeCaption = "ITR 3: AY: 2025-26"
Validate80C_All
'MsgBox "Sheet 80C is OK", vbOKOnly, vbMessgaeCaption
fmsgboxStatus "Sheet 80C is OK"
End Sub
Sub Validate80C_All()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80C")
If Not Validate_80C Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80C
    CloseMsg
End If
End Sub
Function Validate_80C()
Validate_80C = True
'TableA1
'setTblinfo_80CNaturePayment
setTblinfo_80CAmount
setTblinfo_80CIdentification_Number
'end_80C = WorksheetFunction.Max(0, end_80C, end_80CNaturePayment, end_80CIdentification_Number, end_80CAmount)
end_80C = WorksheetFunction.Max(0, end_80C, end_80CIdentification_Number, end_80CAmount)
'    If Not ValidateNaturePayment_80C Then Validate_80C = False
    If Not ValidateAmount_80C Then Validate_80C = False
    If Not ValidateIdentification_Number_80C Then Validate_80C = False
    If Not Validategreater_80C Then Validate_80C = False
    '---------------
End Function

'Sub setTblinfo_80CNaturePayment()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("NaturePayment.80C").count
'    Set rangecells = Range("NaturePayment.80C").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80CNaturePayment = ccount
'End Sub

Sub setTblinfo_80CAmount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Amount.80C").count
    Set rangecells = Range("Amount.80C").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CAmount = ccount
End Sub
Sub setTblinfo_80CIdentification_Number()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Identification_Number.80C").count
    Set rangecells = Range("Identification_Number.80C").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CIdentification_Number = ccount
End Sub



Function ValidateAmount_80C() As Boolean
    ValidateAmount_80C = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet61.Range("Amount.80C").Cells
    ReDim Amount_80C(end_80C)
    For i = 1 To end_80C
        Amount_80C(i) = rangecells.item(i).value
        If Not chkCompulsory(Amount_80C(i)) Then
            'changed by Chetan C M for AY 2025-26
            'start--
'             MsgBox_80C = MsgBox_80C + "* ""Amount eligible for deduction u/s 80C is mandatory in schedule 80C"" at Sr. No " & i & "" & Chr(13)
                MsgBox_80C = MsgBox_80C + "*""Amount eligible for deduction u/s 80C is mandatory in schedule 80C"""
            '--end
            ValidateAmount_80C = False
            Exit Function
        End If
'         If Len(Amount_80C(i)) > 14 Then
'            MsgBox_80C = MsgBox_80C + "* Amount eligible for deduction at Sr. No " & i & " in Sheet 80C cannot be more than 14 characters." & Chr(13)
'            ValidateAmount_80C = False
'            Exit Function
'         End If
         
         If Not IsNumeric(Amount_80C(i)) Then
            MsgBox_80E = MsgBox_80E & "* Amount eligible for deduction at Sr. No  " & i & "  in Sheet 80C should be Numeric value" & Chr(13)
            ValidateAmount_80C = False
            Exit Function
        End If
        
        If Amount_80C(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Amount eligible for deduction at Sr. No  " & i & "  in Sheet 80C cannot exceed 14 digits" & Chr(13)
            ValidateAmount_80C = False
            Exit Function
        End If
Next
End Function


Function ValidateIdentification_Number_80C() As Boolean
    ValidateIdentification_Number_80C = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet61.Range("Identification_Number.80C").Cells
    ReDim Identification_Number_80C(end_80C)
    For i = 1 To end_80C
        Identification_Number_80C(i) = rangecells.item(i).value
        If Not chkCompulsory(Identification_Number_80C(i)) Then
            
                MsgBox_80C = MsgBox_80C + "* ""Policy number or Document Identification number is mandatory in schedule 80C"""
         
            ValidateIdentification_Number_80C = False
            Exit Function
        End If
         If Len(Identification_Number_80C(i)) > 50 Then
            MsgBox_80C = MsgBox_80C + "* Policy number or Document Identification number at Sr. No " & i & " cannot be more than 50 characters." & Chr(13)
            ValidateIdentification_Number_80C = False
            Exit Function
         End If
         
         If Not checkfieldSuperSpecialcharacter(Identification_Number_80C(i)) Then
             MsgBox_80C = MsgBox_80C + "* Policy number or Document Identification number in schedule 80C at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateIdentification_Number_80C = False
            Exit Function
        End If
Next
End Function

Function Validategreater_80C() As Boolean
    Validategreater_80C = True

         If (Len(Sheet61.Range("TotAmount.80C").value) > 14) Then
            MsgBox_80C = MsgBox_80C + "*  Total of deduction u/s 80C amount cannot exceed 14 Digits." & Chr(13)
            Validategreater_80C = False
            Exit Function
         End If
         
End Function
'--end

