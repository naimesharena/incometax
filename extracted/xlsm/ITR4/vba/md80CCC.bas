Attribute VB_Name = "md80CCC"
''Commented by Ayush_06/05/2025 for 0.7 changes
'Public MsgBox_ As Variant
'Dim end_ As Variant
'Dim end_NameInsurer As Variant
'Dim end_Policy_Number As Variant
'Dim end_Amount As Variant
'
'
'Sub ValidateSheet80C_2_Click()
'Dim vbMessgaeCaption As String
'
'vbMessgaeCaption = "ITR 4: AY: 2025-26"
'Validate80C_2_All
''MsgBox "Sheet 80G is OK", vbOKOnly, vbMessgaeCaption
'fmsgboxStatus "Schedule  is OK"
'End Sub
'Sub Validate80C_2_All()
'Dim vbMessgaeCaption As String
'vbMessgaeCaption = "Error"
'   Dim sourceSheet As Worksheet
'    Set sourceSheet = ThisWorkbook.Sheets("80C_")
'If Not Validate_2 Then
'    sourceSheet.Activate
'   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
'    fmsgboxStatus MsgBox_
'    CloseMsg
'End If
'End Sub
'Function Validate_2()
'Validate_2 = True
''Ayush-------------
''TableA1
'setTblinfo_NameInsurer
'
'setTblinfo_Policy_Number
'setTblinfo_Amount
'end_ = WorksheetFunction.Max(0, end_, end_NameInsurer, end_Policy_Number, end_Amount)
'    If Not ValidateNameInsurer_ Then Validate_2 = False
'    If Not ValidateAmount_ Then Validate_2 = False
'    If Not ValidatePolicy_Number_ Then Validate_2 = False
'
'    If Len(Sheet19.Range("TotAmount.")) > 14 Then
'            MsgBox_ = MsgBox_ + "* Total of Deduction u/s  at Sr. No " & i & " in Schedule  cannot be more than 14 characters." & Chr(13)
'            Validate_2 = False
'
'         End If
'
'    '---------------
'End Function
'Sub setTblinfo_NameInsurer()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Name_of_insurer.").count
'    Set rangecells = Range("Name_of_insurer.").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_NameInsurer = ccount
'End Sub
'Sub setTblinfo_Amount()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Amount.").count
'    Set rangecells = Range("Amount.").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_Amount = ccount
'End Sub
'Sub setTblinfo_Policy_Number()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Policy_Document_Number.").count
'    Set rangecells = Range("Policy_Document_Number.").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_Policy_Number = ccount
'End Sub
'
'Function ValidateNameInsurer_() As Boolean
'    ValidateNameInsurer_ = True
''    setTblinfo_TCS
''setTblinfo_80DNameA1
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Name_of_insurer.").Cells
'    ReDim Nature_(end_)
'    For i = 1 To end_
'        Nature_(i) = rangecells.item(i).value
'        If Not chkCompulsory(Nature_(i)) Then
'             MsgBox_ = MsgBox_ + "* ""Name of the insurer"" is mandatory in schedule  at Sr. No " & i & "" & Chr(13)
'            ValidateNameInsurer_ = False
'            Exit Function
'        End If
'         If Len(Nature_(i)) > 75 Then
'            MsgBox_ = MsgBox_ + "* Name Of insurer at  Sr. No " & i & " in Schedule  cannot be more than 75 characters." & Chr(13)
'            ValidateNameInsurer_ = False
'            Exit Function
'         End If
'
'         If Not checkfieldSuperSpecialcharacter(Nature_(i)) Then
'             MsgBox_ = MsgBox_ + "* Name Of insurer in schedule  at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
'            ValidateNameInsurer_ = False
'            Exit Function
'        End If
'Next
'End Function
'
'Function ValidateAmount_() As Boolean
'    ValidateAmount_ = True
''    setTblinfo_TCS
''setTblinfo_80DNameA1
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Amount.").Cells
'    ReDim Amount_(end_)
'    For i = 1 To end_
'        Amount_(i) = rangecells.item(i).value
'        If Not chkCompulsory(Amount_(i)) Then
'             MsgBox_ = MsgBox_ + "* ""Deduction u/s   is mandatory in schedule  Sr. No " & i & """" & Chr(13)
'            ValidateAmount_ = False
'            Exit Function
'        End If
''         If Len(Amount_(i)) > 125 Then
''            MsgBox_ = MsgBox_ + "* Amount at A1as Sr. No " & i & " in Sheet 80D cannot be more than 125 characters." & Chr(13)
''            ValidateAmount_ = False
''            Exit Function
''         End If
'
'          If Not IsNumeric(Amount_(i)) Then
'            MsgBox_ = MsgBox_ & "* Deduction u/s  at Sr. No  " & i & "  in Schedule  should be Numeric value" & Chr(13)
'            ValidateAmount_ = False
'            Exit Function
'        End If
'
'        If Amount_(i) > 99999999999999# Then
'            MsgBox_ = MsgBox_ & "* Deduction u/s at Sr. No  " & i & "  in Schedule  cannot exceed 14 digits" & Chr(13)
'            ValidateAmount_ = False
'            Exit Function
'        End If
'Next
'End Function
'
'
'Function ValidatePolicy_Number_() As Boolean
'    ValidatePolicy_Number_ = True
''    setTblinfo_TCS
''setTblinfo_80DNameA1
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet19.Range("Policy_Document_Number.").Cells
'    ReDim Identification_Number_(end_)
'    For i = 1 To end_
'        Identification_Number_(i) = rangecells.item(i).value
'        If Not chkCompulsory(Identification_Number_(i)) Then
'             MsgBox_ = MsgBox_ + "* ""Policy document number"" is mandatory in schedule  at Sr. No " & i & "" & Chr(13)
'            ValidatePolicy_Number_ = False
'            Exit Function
'        End If
'         If Len(Identification_Number_(i)) > 100 Then
'            MsgBox_ = MsgBox_ + "* Policy document number at Sr. No " & i & " in Schedule  cannot be more than 100 characters." & Chr(13)
'            ValidatePolicy_Number_ = False
'            Exit Function
'         End If
'
'         If Not checkfieldSuperSpecialcharacter(Identification_Number_(i)) Then
'             MsgBox_ = MsgBox_ + "* Policy document number in schedule  at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
'            ValidatePolicy_Number_ = False
'            Exit Function
'        End If
'Next
'End Function
'
'
'
