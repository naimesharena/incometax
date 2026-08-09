Attribute VB_Name = "md80C"
 
 
Public MsgBox_80C As Variant
Dim end_80C As Variant
Dim end_80CNaturePayment As Variant
Dim end_80CIdentification_Number As Variant
Dim end_80CAmount As Variant


'Ayush_23/04
'Ayush_23

Sub cmd_80CPrev_click()
If Sheet1.Range("sheet1.Status").value = "H - HUF" And Sheet8.Range("bacvalue") = 2 Then
Sheet11.Activate

Else
Sheet20.Activate

End If



End Sub

Sub ValidateSheet80C__Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 4: AY: 2025-26"
Validate80C__All
'MsgBox "Sheet 80G is OK", vbOKOnly, vbMessgaeCaption
fmsgboxStatus "Schedule 80C is OK"
End Sub
Sub Validate80C__All()
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
'Ayush-------------
'TableA1
'Commented by Ayush_06/05/2025 for 0.7 changes
'setTblinfo_80CNaturePayment
setTblinfo_80CAmount
setTblinfo_80CIdentification_Number

'end_80C = WorksheetFunction.Max(0, end_80C, end_80CNaturePayment, end_80CIdentification_Number, end_80CAmount)
end_80C = WorksheetFunction.Max(0, end_80CIdentification_Number, end_80CAmount)
  '  If Not ValidateNaturePayment_80C Then Validate_80C = False
    If Not ValidateAmount_80C Then Validate_80C = False
    If Not ValidateIdentification_Number_80C Then Validate_80C = False
    
    If Len(Sheet19.Range("TotAmount.80C")) > 14 Then
            MsgBox_80C = MsgBox_80C + "* Total of Deduction u/s 80C at Sr. No " & i & " in Schedule 80C cannot be more than 14 characters." & Chr(13)
            Validate_80C = False
            
         End If
    
    
    
    '---------------
End Function
'Commented by Ayush_06/05/2025 for 0.7 changes
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
'Commented by Ayush_06/05/2025 for 0.7 changes
'Function ValidateNaturePayment_80C() As Boolean
'    ValidateNaturePayment_80C = True
''    setTblinfo_TCS
''setTblinfo_80DNameA1
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("NaturePayment.80C").Cells
'    ReDim Nature_80C(end_80C)
'    For i = 1 To end_80C
'        Nature_80C(i) = rangecells.item(i).value
'        If Not chkCompulsory(Nature_80C(i)) Then
'             MsgBox_80C = MsgBox_80C + "* ""Nature of payment is mandatory in schedule 80C at Sr. No " & i & """" & Chr(13)
'            ValidateNaturePayment_80C = False
'            Exit Function
'        End If
'         If Len(Nature_80C(i)) > 100 Then
'            MsgBox_80C = MsgBox_80C + "* Nature of payment Sr. No " & i & " in Schedule 80C cannot be more than 100 characters." & Chr(13)
'            ValidateNaturePayment_80C = False
'            Exit Function
'         End If
'
'         If Not checkfieldSuperSpecialcharacter(Nature_80C(i)) Then
'             MsgBox_80C = MsgBox_80C + "* Nature of payment in schedule 80C at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
'            ValidateNaturePayment_80C = False
'            Exit Function
'        End If
'Next
'End Function

Function ValidateAmount_80C() As Boolean
    ValidateAmount_80C = True
'    setTblinfo_TCS
'setTblinfo_80DNameA1
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet19.Range("Amount.80C").Cells
    ReDim Amount_80C(end_80C)
    For i = 1 To end_80C
        Amount_80C(i) = rangecells.item(i).value
        If Not chkCompulsory(Amount_80C(i)) Then
             MsgBox_80C = MsgBox_80C + "* ""Amount eligible for deduction u/s 80C is mandatory in schedule 80C at Sr. No " & i & """" & Chr(13)
            ValidateAmount_80C = False
            Exit Function
        End If
'         If Len(Amount_80C(i)) > 14 Then
'            MsgBox_80C = MsgBox_80C + "* Amount eligible for deduction at Sr. No " & i & " in Sheet 80C cannot be more than 14 characters." & Chr(13)
'            ValidateAmount_80C = False
'            Exit Function
'         End If
         
         If Not IsNumeric(Amount_80C(i)) Then
            MsgBox_80E = MsgBox_80E & "* Amount eligible for deduction at Sr. No  " & i & "  in Schedule 80C should be Numeric value" & Chr(13)
            ValidateAmount_80C = False
            Exit Function
        End If
        
        If Amount_80C(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Amount eligible for deduction at Sr. No  " & i & "  in Schedule 80C cannot exceed 14 digits" & Chr(13)
            ValidateAmount_80C = False
            Exit Function
        End If
Next
End Function


Function ValidateIdentification_Number_80C() As Boolean
    ValidateIdentification_Number_80C = True
'    setTblinfo_TCS
'setTblinfo_80DNameA1
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet19.Range("Identification_Number.80C").Cells
    ReDim Identification_Number_80C(end_80C)
    For i = 1 To end_80C
        Identification_Number_80C(i) = rangecells.item(i).value
        If Not chkCompulsory(Identification_Number_80C(i)) Then
             MsgBox_80C = MsgBox_80C + "* ""Policy number or Document Identification number is mandatory in schedule 80C"" at Sr. No " & i & """" & Chr(13)
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

 
 
