Attribute VB_Name = "Sch80U_80DD"

Option Explicit

Public MsgBox_80U, MsgBox_80DD As String

 
Public Nature_of_disability_80U As Variant
Public Amount_of_deduction_80U As Variant
Public Date_filingFm10IA_80U As Variant
Public AckNo_ofForm10IAfiled_80U As Variant
Public UDID_Num_80U As Variant
Public rngname_80U As Variant
Public end80U, end80U1, end80U2, end80U3, end80U4  As Long

'added by Chetan C M for AY 2025-26
Public end80U5, end80U6 As Long
Public Type_of_disability_80U As Variant
Public AckNo_ofFormRule11A_80U As Variant

'--------------------80DD
Public end80DD7, end80DD6, end80DD5, end80DD4, end80DD3, end80DD2, end80DD1, end80DD  As Long

Public Type_dependent_80DD As Variant
Public Aadhaar_dependent_80DD As Variant
Public PAN_dependent_80DD  As Variant
Public UDID_Num_80DD As Variant
Public AckNo_ofForm10IAfiled_80DD As Variant
Public Date_filingFm10IA_80DD As Variant
Public Amount_of_deduction_80DD As Variant
Public Nature_of_disability_80DD As Variant

'added by Chetan C M for AY 2025-26
Public end80DD8, end80DD9    As Long
Public Type_of_disability_80DD As Variant
Public AckNo_ofFormRule11A_80DD As Variant

Sub ValidateSheet80U_Click()
 Dim vbMessgaeCaption As String
 'vbMessgaeCaption = "ITR 2: AY: 2024-25"
  vbMessgaeCaption = "ITR 2: AY: 2025-26" 'Shrutika(27-02-25)
 
  Validate80U
  'MsgBox "Sheet 80U is OK", vbOKOnly, vbMessgaeCaption
   fmsgbox "Sheet 80U is OK"
End Sub
Sub ValidateSheet80DD_Click()
 Dim vbMessgaeCaption As String
 'vbMessgaeCaption = "ITR 2: AY: 2024-25"
 vbMessgaeCaption = "ITR 2: AY: 2025-26" 'Shrutika(27-02-25)
 
Validate80DD
'MsgBox "Sheet 80U is OK", vbOKOnly, vbMessgaeCaption
fmsgbox "Sheet 80DD is OK"
End Sub

'---------------------------------------------------------------------------------------------
Sub Validate80U()
    If Not Validate80U_1 Then
        Sheet58.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80U)
        CloseMsg
    End If
End Sub



Function Validate80U_1() As Boolean
    Validate80U_1 = True
    MsgBox_80U = "Schedule 80U : " & Chr(10)
         
setTableInfo80U
'setTableInfo80U1
setTableInfo80U2
setTableInfo80U3
setTableInfo80U4

'added by Chetan C M for AY 2025-26
setTableInfo80U5
setTableInfo80U6

    end80U = WorksheetFunction.Max(0, end80U, end80U2, end80U3, end80U4, end80U5, end80U6)


'    end80U = WorksheetFunction.Max(0, end80U, end80U2, end80U3, end80U4)

 If end80U > 0 Then

       If Not ValidateNature_disability_80U Then Validate80U_1 = False
       If Not ValidateAmount_of_deduction_80U Then Validate80U_1 = False
       If Not ValidateType_disability_80U Then Validate80U_1 = False    'added by Chetan C M for AY 2025-26
       If Not ValidateDate_of_filingofForm10IA_80U Then Validate80U_1 = False
       If Not ValidateAckNoFm10IAfiled_80U Then Validate80U_1 = False
       If Not ValidateAckNo_ofFormRule11A_80U Then Validate80U_1 = False    'added by Chetan C M for AY 2025-26
       If Not ValidateUDIDNum_80U Then Validate80U_1 = False
End If
        

        
End Function


Sub setTableInfo80U()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Naturedisability_80U").Cells
    mIntCells = Sheet58.Range("Naturedisability_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80U = ccount
    'rngname_80U = "Naturedisability_80U;Amtdeduction_80U;DatefilingFm10IA_80U;AckNoFm10IAfiled_80U;UDIDNum_80U;"
End Sub

'added by Chetan C M for AY 2025-26
'start--
'Disability_80U
Sub setTableInfo80U5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Disability_80U").Cells
    mIntCells = Sheet58.Range("Disability_80U").count
    For mIntCtr = 1 To mIntCells
        'If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
        ccount = ccount + 1
        End If
    Next
    end80U5 = ccount
End Sub
'--end

Sub setTableInfo80U1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Amtdeduction_80U").Cells
    mIntCells = Sheet58.Range("Amtdeduction_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U1 = ccount
End Sub

Sub setTableInfo80U2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("DatefilingFm10IA_80U").Cells
    mIntCells = Sheet58.Range("DatefilingFm10IA_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U2 = ccount
End Sub

Sub setTableInfo80U3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("AckNoFm10IAfiled_80U").Cells
    mIntCells = Sheet58.Range("AckNoFm10IAfiled_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U3 = ccount
End Sub

'added by Chetan C M for AY 2025-26
'start--
Sub setTableInfo80U6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("AcknowledgeNum11A2_80U").Cells
    mIntCells = Sheet58.Range("AcknowledgeNum11A2_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U6 = ccount
End Sub
'--end

Sub setTableInfo80U4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("UDIDNum_80U").Cells
    mIntCells = Sheet58.Range("UDIDNum_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80U4 = ccount
End Sub

Function ValidateNature_disability_80U() As Boolean
   ValidateNature_disability_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Naturedisability_80U").Cells
    Dim i As Long
    ReDim Nature_of_disability_80U(end80U)
    
    For i = 1 To end80U

        Nature_of_disability_80U(i) = rangecells.item(i).value
        
       If isdropdownblank(Nature_of_disability_80U(i)) Or UCase(Nature_of_disability_80U(i)) = UCase("Select") Then
            MsgBox_80U = MsgBox_80U + "* Please select one of the dropdown in 'Nature of disability' in schedule 80U" & Chr(13)
            ValidateNature_disability_80U = False
            Exit Function
            UpdateProgressBar
        
       End If
    Next
End Function
 
'added by Chetan C M for AY 2025-26
'start--
Function ValidateType_disability_80U() As Boolean
   ValidateType_disability_80U = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Disability_80U").Cells
    Dim i As Long
    ReDim Type_of_disability_80U(end80U)
    
    For i = 1 To end80U

        Type_of_disability_80U(i) = rangecells.item(i).value
        
        'If isdropdownblank(Type_of_disability_80U(i)) Or UCase(Type_of_disability_80U(i)) = UCase("Select") Then
         'Malli-------
        If Type_of_disability_80U(i) = "" Or Type_of_disability_80U(i) = "(Select)" Then
           MsgBox_80U = MsgBox_80U + "* Selection of ""Type of disability"" in schedule 80U is mandatory." & Chr(13)
           ValidateType_disability_80U = False
            Exit Function
         
            UpdateProgressBar
            
'        Else
'
'        If Sheet58.Range("AckNoFm10IAfiled_80DD").value = Sheet58.Range("AckNoFm10IAfiled_80U").value Then
'        If Sheet58.Range("AckNoFm10IAfiled_80DD").MergeArea.Locked = False Then
'            MsgBox "Ack. No. of Form 10IA filed in Schedule 80U and Schedule 80DD should not be same"
'        End If
'        End If
        
        
        End If
    Next
End Function
'--end
 
 
Function ValidateAmount_of_deduction_80U() As Boolean
    ValidateAmount_of_deduction_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Amtdeduction_80U").Cells
    Dim i As Long
    ReDim Amount_of_deduction_80U(end80U)
    
    For i = 1 To end80U
        Amount_of_deduction_80U(i) = rangecells.item(i).value
        
        If Not chkCompulsory(Amount_of_deduction_80U(i)) Then
            MsgBox_80U = MsgBox_80U + "* Please enter Amount of deduction in Schedule 80U" & Chr(13)
            ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
        UpdateProgressBar

        If Not IsNumeric(Amount_of_deduction_80U(i)) Then
           MsgBox_80U = MsgBox_80U + "* Amount of deduction Schedule 80U should be Numeric value" & Chr(13)
            ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
        
        If Amount_of_deduction_80U(i) > 99999999999999# Then
           MsgBox_80U = MsgBox_80U + "* Amount of deduction in Schedule 80U cannot exceed 14 digits" & Chr(13)
           ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
Next

End Function

 
Function ValidateDate_of_filingofForm10IA_80U() As Boolean
    
ValidateDate_of_filingofForm10IA_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("DatefilingFm10IA_80U").Cells
    Dim i As Long
    ReDim Date_filingFm10IA_80U(end80U)
    
    For i = 1 To end80U
        Date_filingFm10IA_80U(i) = rangecells.item(i).value
        
        'added by Chetan C M for AY 2025-26
        'start--
        
        If Sheet58.Range("DatefilingFm10IA_80U").Locked = False Then
        If Date_filingFm10IA_80U(i) = "" Or Date_filingFm10IA_80U(i) = "(Select)" Then
          If Sheet58.Range("DatefilingFm10IA_80U").MergeArea.Locked = False Then ' Added by Aavula Naresh SIT-93764
            MsgBox_80U = MsgBox_80U + "* Please enter date of filing of Form 10IA" & Chr(13)
            ValidateDate_of_filingofForm10IA_80U = False
            Exit Function
        End If
        End If
        
        '--end

        If Not CheckDateddmmyyyy(Date_filingFm10IA_80U(i)) Then
            ValidateDate_of_filingofForm10IA_80U = False
            MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA must be a valid dd/mm/yyyy format in Schedule 80U" & Chr(13)
            Exit Function
        End If
            UpdateProgressBar
    
        If Len(Date_filingFm10IA_80U(i)) > 10 Then
            ValidateDate_of_filingofForm10IA_80U = False
            MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA cannot exceed 10 digits in Schedule 80U" & Chr(13)
            Exit Function
        End If
    End If
    
Next
End Function

 

Function ValidateAckNoFm10IAfiled_80U() As Boolean
    
ValidateAckNoFm10IAfiled_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("AckNoFm10IAfiled_80U").Cells
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80U(end80U)
    
    For i = 1 To end80U
        AckNo_ofForm10IAfiled_80U(i) = rangecells.item(i).value
        
      'added by Chetan C M for AY 2025-26
        'start--
        If Sheet58.Range("AckNoFm10IAfiled_80U").Locked = False Then
        If AckNo_ofForm10IAfiled_80U(i) = "" Or AckNo_ofForm10IAfiled_80U(i) = "(Select)" Then
          If Sheet58.Range("AckNoFm10IAfiled_80U").MergeArea.Locked = False Then
'            MsgBox_80U = MsgBox_80U + "* Please enter Ack. No. of Form 10IA in Schedule 80U" & Chr(13)
            MsgBox_80U = MsgBox_80U + "* Please enter the Acknowledgement number of Form 10IA" & Chr(13) ' added by Chetan C M on 18/06/2025 for AY 2025-26
            ValidateAckNoFm10IAfiled_80U = False
            Exit Function
          End If
        End If
        
        '--end
     
      If Not checkfieldspecialcharacter1(AckNo_ofForm10IAfiled_80U(i)) Then
'             MsgBox_80U = MsgBox_80U + "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80U." & Chr(13)
            MsgBox_80U = MsgBox_80U + "* Please enter a valid acknowledgement number of Form 10IA" & Chr(13)   ' added by Chetan C M on 26/06/2025 for AY 2025-26 based in V0.6 DE sheet changes
             ValidateAckNoFm10IAfiled_80U = False
             Exit Function
        End If
     
        'UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80U(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80U = False
        MsgBox_80U = MsgBox_80U + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
    End If
    
Next
End Function
 
'added by Chetan C M for AY 2025-26
'start--
Function ValidateAckNo_ofFormRule11A_80U() As Boolean

ValidateAckNo_ofFormRule11A_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("AcknowledgeNum11A2_80U").Cells
    Dim i As Long
    ReDim Acknowledge_Num11A2_80U(end80U)
    
    For i = 1 To end80U
        Acknowledge_Num11A2_80U(i) = rangecells.item(i).value
     
     'added by Chetan C M for AY 2025-26
        'start--
        ' If Sheet58.Range("Disability_80U").value <> "(i)autism, cerebral palsy, or multiple disabilities" Then
        
        If Sheet58.Range("Disability_80U").value = "(ii) others" Then  'Add by Sudhakar

        If Acknowledge_Num11A2_80U(i) = "" Or Acknowledge_Num11A2_80U(i) = "(Select)" Then
       
            MsgBox_80U = MsgBox_80U + "* Please enter the Acknowledgement Number of Form as per Rule 11A(2)(ii)" & Chr(13)
            ValidateAckNo_ofFormRule11A_80U = False
            Exit Function
        End If
       
        '--end
     
'Sudha_2025-26

'      If Not checkfieldspecialcharacter1(Acknowledge_Num11A2_80U(i)) Then
'             MsgBox_80U = MsgBox_80U + "* Ack no. of Form as per Rule 11A(2) field cannot contain special characters in Schedule 80U." & Chr(13)
'             ValidateAckNo_ofFormRule11A_80U = False
'             Exit Function
'        End If
        
        
              If Not checkfieldspecialcharacter101(Acknowledge_Num11A2_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* Ack no. of Form as per Rule 11A(2) field cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateAckNo_ofFormRule11A_80U = False
             Exit Function
        End If
 '===========
 
        'UpdateProgressBar

    'If Len(Acknowledge_Num11A2_80U(i)) > 15 Then
    
'SIT-95387 by sai on 26/06/2025
     If Len(Acknowledge_Num11A2_80U(i)) > 50 Then
        ValidateAckNo_ofFormRule11A_80U = False
        MsgBox_80U = MsgBox_80U + "*  Ack no. of Form as per Rule 11A(2) field cannot exceed 50 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
End If
     
Next
End Function
'--end
 
Function ValidateUDIDNum_80U() As Boolean
    
ValidateUDIDNum_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("UDIDNum_80U").Cells
    Dim i As Long
    ReDim UDID_Num_80U(end80U)
    
    For i = 1 To end80U
        UDID_Num_80U(i) = rangecells.item(i).value
     
      If Not checkallfieldspecialcharacter_80U_80DD(UDID_Num_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* UDID Number cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateUDIDNum_80U = False
             Exit Function
        End If
     
        'UpdateProgressBar

    If Len(UDID_Num_80U(i)) > 18 Then   'Chandru
        ValidateUDIDNum_80U = False
        MsgBox_80U = MsgBox_80U + "*UDID Number cannot exceed 18 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
Next
End Function

'__________________________________________________80DD
Sub setTableInfo80DD()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Naturedisability_80DD").Cells
    mIntCells = Sheet58.Range("Naturedisability_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80DD = ccount
    'rngname_80GGA = "RelevantClauseClaimed_80GGA;Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_total_80GGA;Donation_Eligible_80GGA;"
End Sub

'added by Chetan C M for AY 2025-26
'start--
Sub setTableInfo80DD8() 'TypeDisability80DD(1)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Disability_80DD").Cells
    mIntCells = Sheet58.Range("Disability_80DD").count
    For mIntCtr = 1 To mIntCells
         'If Not isdropdownblank(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
         'Debug.Print rangecells.item(mIntCtr).Value
'          If rangecells.item(mIntCtr).Value <> "" And UCase(rangecells.item(mIntCtr).Value) <> UCase("(Select)") Then    'Malli
          If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "(Select)" Then
'         If Not isdropdownblankDD(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
          ccount = ccount + 1
        End If
    Next
    end80DD8 = ccount
'    rngname_80DD = "NatureDisability_80DD;AmtDeduction_80DD;Dependent_80DD;PanDependent_80DD;AadhaarDependent_80DD;Form10IA_80DD;AcknowledgeNum_80DD;UDIDNumber_80DD"
End Sub
'--end

 
Sub setTableInfo80DD1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Amtdeduction_80DD").Cells
    mIntCells = Sheet58.Range("Amtdeduction_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" And Not rangecells.item(mIntCtr).value = "0" Then
        ccount = ccount + 1
        End If
    Next
    end80DD1 = ccount
End Sub

Sub setTableInfo80DD2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Typedependent_80DD").Cells
    mIntCells = Sheet58.Range("Typedependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80DD2 = ccount
End Sub

Sub setTableInfo80DD3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("PANdependent_80DD").Cells
    mIntCells = Sheet58.Range("PANdependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80DD3 = ccount
End Sub

Sub setTableInfo80DD4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Aadhaardependent_80DD").Cells
    mIntCells = Sheet58.Range("Aadhaardependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD4 = ccount
End Sub

Sub setTableInfo80DD5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("DatefilingFm10IA_80DD").Cells
    mIntCells = Sheet58.Range("DatefilingFm10IA_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD5 = ccount
End Sub

Sub setTableInfo80DD6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("AckNoFm10IAfiled_80DD").Cells
    mIntCells = Sheet58.Range("AckNoFm10IAfiled_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD6 = ccount
End Sub

'added by Chetan C M for AY 2025-26
'start--
Sub setTableInfo80DD9()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("AcknowledgeNum11A2_80DD").Cells
    mIntCells = Sheet58.Range("AcknowledgeNum11A2_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD9 = ccount
End Sub
'--end


Sub setTableInfo80DD7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("UDIDNum_80DD").Cells
    mIntCells = Sheet58.Range("UDIDNum_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD7 = ccount
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------

'-------------------------------------
Function ValidateNature_disability_80DD() As Boolean
   ValidateNature_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Naturedisability_80DD").Cells
    Dim i As Long
    ReDim Nature_of_disability_80DD(end80DD)
    
    For i = 1 To end80DD

        Nature_of_disability_80DD(i) = rangecells.item(i).value
        
        If isdropdownblank(Nature_of_disability_80DD(i)) Or UCase(Nature_of_disability_80DD(i)) = UCase("Select") Then
               
         
            MsgBox_80DD = MsgBox_80DD + "* Please select one of the dropdown in 'Nature of disability' in Schedule 80DD" & Chr(13)
           ValidateNature_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function

'added by Chetan C M for AY 2025-26
'start--
Function ValidateType_disability_80DD() As Boolean
   ValidateType_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Disability_80DD").Cells
    Dim i As Long
    ReDim Type_of_disability_80DD(end80DD)
    
    For i = 1 To end80DD

        Type_of_disability_80DD(i) = rangecells.item(i).value
        
       ' If isdropdownblank(Type_of_disability_80DD(i)) Or UCase(Type_of_disability_80DD(i)) = UCase("Select") Then
        
         If Type_of_disability_80DD(i) = "" Or Type_of_disability_80DD(i) = "(Select)" Then  'Malli
           MsgBox_80DD = MsgBox_80DD + "*Selection of ""Type of disability"" in schedule 80DD is mandatory." & Chr(13)
           ValidateType_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
'        Else
'
'        If Sheet58.Range("AckNoFm10IAfiled_80U").value = Sheet58.Range("AckNoFm10IAfiled_80DD").value Then
'        If Sheet58.Range("AckNoFm10IAfiled_80U").MergeArea.Locked = False Then
'            MsgBox "Ack. No. of Form 10IA filed in Schedule 80DD and Schedule 80U should not be same"
'        End If
'        End If
        
        End If
    Next
End Function
'--end


'-----------------------------------------------------
Function ValidateAmount_of_deduction_80DD() As Boolean
    ValidateAmount_of_deduction_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Amtdeduction_80DD").Cells
    Dim i As Long
    ReDim Amount_of_deduction_80DD(end80DD)
    
    For i = 1 To end80DD
        
     Amount_of_deduction_80DD(i) = rangecells.item(i).value
        
If Amount_of_deduction_80DD(i) <> 0 Then
        
If Not chkCompulsory(Amount_of_deduction_80DD(i)) Then
            MsgBox_80DD = MsgBox_80DD + "* Please enter Amount of deduction in Schedule 80DD" & Chr(13)
            ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        UpdateProgressBar

        If Not IsNumeric(Amount_of_deduction_80DD(i)) Then
           MsgBox_80DD = MsgBox_80DD + "* Amount of deduction in  Schedule 80DD should be Numeric value" & Chr(13)
            ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        
        If Amount_of_deduction_80DD(i) > 99999999999999# Then
           MsgBox_80DD = MsgBox_80DD + "* Amount of deduction in Schedule 80DD cannot exceed 14 digits" & Chr(13)
           ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        
   End If
Next

End Function

'--------------------------------------------------------------------------------------------------------------
Function ValidateDate_of_filingofForm10IA_80DD() As Boolean
    
ValidateDate_of_filingofForm10IA_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("DatefilingFm10IA_80DD").Cells
    Dim i As Long
    ReDim Date_filingFm10IA_80DD(end80DD)
    
    For i = 1 To end80DD
        Date_filingFm10IA_80DD(i) = rangecells.item(i).value
        
        'added by Chetan C M for AY 2025-26
    'start--
    If Sheet58.Range("DatefilingFm10IA_80DD").Locked = False Then
    If Date_filingFm10IA_80DD(i) = "" Or Date_filingFm10IA_80DD(i) = "(Select)" Then
      If Sheet58.Range("DatefilingFm10IA_80DD").MergeArea.Locked = False Then
           MsgBox_80DD = MsgBox_80DD + "* Please enter the date of filing of Form 10IA" & Chr(13)
            ValidateDate_of_filingofForm10IA_80DD = False
            Exit Function
      End If
    End If
    
    '--end
        
    If Not CheckDateddmmyyyy(Date_filingFm10IA_80DD(i)) Then
        ValidateDate_of_filingofForm10IA_80DD = False
       MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA must be a valid dd/mm/yyyy format in Schedule 80DD" & Chr(13)
        Exit Function
    End If
        UpdateProgressBar

    If Len(Date_filingFm10IA_80DD(i)) > 10 Then
        ValidateDate_of_filingofForm10IA_80DD = False
        MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA cannot exceed 10 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
    End If
    
Next
End Function

'---------------------------------------------------------------------

Function ValidateAckNoFm10IAfiled_80DD() As Boolean
    
ValidateAckNoFm10IAfiled_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("AckNoFm10IAfiled_80DD").Cells
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80DD(end80DD)
    
    For i = 1 To end80DD
        AckNo_ofForm10IAfiled_80DD(i) = rangecells.item(i).value
        
        'added by Chetan C M for AY 2025-26
    'start--
    If Sheet58.Range("AckNoFm10IAfiled_80DD").Locked = False Then
        If AckNo_ofForm10IAfiled_80DD(i) = "" Or AckNo_ofForm10IAfiled_80DD(i) = "(Select)" Then
         If Sheet58.Range("AckNoFm10IAfiled_80DD").MergeArea.Locked = False Then
'           MsgBox_80DD = MsgBox_80DD + "* Please enter the Ack. No. of Form 10IA filed" & Chr(13)
            MsgBox_80DD = MsgBox_80DD + "* Please enter the Acknowledgement number of Form 10IA" & Chr(13)  'added by Chetan C M on 18/06/2025 for AY 2025-26
            ValidateAckNoFm10IAfiled_80DD = False
            Exit Function
         End If
        End If
        
    '--end
     
      If Not checkallfieldspecialcharacter_80U_80DD(AckNo_ofForm10IAfiled_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80DD." & Chr(13)
             MsgBox_80DD = MsgBox_80DD + "* Please enter a valid acknowledgement number of Form 10IA" & Chr(13) ' added by Chetan C M on 26/06/2025 for AY 2025-26 based in V0.6 DE sheet changes
             ValidateAckNoFm10IAfiled_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80DD(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
    End If
    
Next
End Function

'added by Chetan C M for AY 2025-26
'start--
Function ValidateAckNo_ofFormRule11A_80DD() As Boolean
    
ValidateAckNo_ofFormRule11A_80DD = True
 
    Dim rangecells, rangecells2 As Range
    Set rangecells = Sheet58.Range("AcknowledgeNum11A2_80DD").Cells
    Set rangecells2 = Sheet58.Range("Disability_80DD").Cells
    Dim i As Long
    ReDim Acknowledge_Num11A2_80DD(end80DD)
    ReDim Disability_80DD(end80DD)
    For i = 1 To end80DD
        Acknowledge_Num11A2_80DD(i) = rangecells.item(i).value
        Disability_80DD(i) = rangecells2.item(i).value
        
    'Added by Aavula Naresh
      If Sheet58.Range("Disability_80DD").value = "(ii) others" Then 'Added by Shrutika(SIT-93775)
       If Acknowledge_Num11A2_80DD(i) = "" Or Acknowledge_Num11A2_80DD(i) = "(Select)" Then
           MsgBox_80DD = MsgBox_80DD + "* Please enter the Acknowledgement Number of Form as per Rule 11A(2)(ii)" & Chr(13)
            ValidateAckNo_ofFormRule11A_80DD = False
            Exit Function
        End If
        End If
     '''''''''
    
'sudha_2025-26

'      If Not checkallfieldspecialcharacter_80U_80DD(Acknowledge_Num11A2_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "* Ack no. of Form as per Rule 11A(2) filed cannot contain special characters in Schedule 80DD." & Chr(13)
'             ValidateAckNo_ofFormRule11A_80DD = False
'             Exit Function
'        End If
        
          If Not checkfieldspecialcharacter101(Acknowledge_Num11A2_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Ack no. of Form as per Rule 11A(2) filed cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateAckNo_ofFormRule11A_80DD = False
             Exit Function
        End If
        
     
        UpdateProgressBar

    'If Len(Acknowledge_Num11A2_80DD(i)) > 15 Then
'SIT-95387 by sai on 26/06/2025
    If Len(Acknowledge_Num11A2_80DD(i)) > 15 Then
        ValidateAckNo_ofFormRule11A_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*  Ack no. of Form as per Rule 11A(2) filed cannot exceed 15 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     
Next
End Function
'--end

'------------------------------------------------------------------------------------------------------------------------------
Function ValidateUDIDNum_80DD() As Boolean
    
ValidateUDIDNum_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("UDIDNum_80DD").Cells
    Dim i As Long
    ReDim UDID_Num_80DD(end80DD)
    
    For i = 1 To end80DD
        UDID_Num_80DD(i) = rangecells.item(i).value
     
      If Not checkallfieldspecialcharacter_80U_80DD(UDID_Num_80DD(i)) Then
            MsgBox_80DD = MsgBox_80DD + "* UDID Number cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateUDIDNum_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(UDID_Num_80DD(i)) > 18 Then 'Chandru
        ValidateUDIDNum_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*UDID Number cannot exceed 18 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     
Next
End Function

Function ValidatePANdependent_80DD() As Boolean
    
ValidatePANdependent_80DD = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("PANdependent_80DD").Cells
    Dim i As Long
    ReDim PAN_dependent_80DD(end80DD)
    
    For i = 1 To end80DD
        PAN_dependent_80DD(i) = rangecells.item(i).value
        
If Trim(PAN_dependent_80DD(i)) <> "" Then
     
      If Not checkfieldspecialcharacter(PAN_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "*Invalid PAN in Schedule 80DD." & Chr(13)
            ValidatePANdependent_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Not CheckPAB_80DD(UCase(PAN_dependent_80DD(i))) Then
             MsgBox_80DD = MsgBox_80DD + "*Invalid PAN in Schedule 80DD." & Chr(13)
             ValidatePANdependent_80DD = False
             Exit Function
        End If

'If PAN_dependent_80DD(i) <> o Then
        If ((UCase(PAN_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(PAN_dependent_80DD(i)) = UCase(Sheet52.Range("sheet9.PAN").value))) Then
        MsgBox_80DD = MsgBox_80DD + "* PAN of the dependent cannot be same as assessee PAN in Part-A General Information." & Chr(13)
         ValidatePANdependent_80DD = False
        Exit Function
        End If
'End If
 
 End If
 
Next
End Function


Function ValidateAadhaardependent_80DD() As Boolean
    
ValidateAadhaardependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Aadhaardependent_80DD").Cells
    Dim i As Long
    ReDim Aadhaar_dependent_80DD(end80DD)
    
    For i = 1 To end80DD
        Aadhaar_dependent_80DD(i) = rangecells.item(i).value
     

If Trim(Aadhaar_dependent_80DD(i)) <> "" Then


      If Not checkfieldspecialcharacter(Aadhaar_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "*Invalid Aadhaar in Schedule 80DD." & Chr(13)
            ValidateAadhaardependent_80DD = False
             Exit Function
        End If


      If UCase(Aadhaar_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.adhaarno").value) Then
        MsgBox_80DD = MsgBox_80DD + "* Aadhaar of the dependent cannot be same as assessee Aadhaar in Part-A General Information." & Chr(13)
         ValidateAadhaardependent_80DD = False
        Exit Function
        End If
     
        If Not IsNumeric(Aadhaar_dependent_80DD(i)) Then
            
           MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
           ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Len(Aadhaar_dependent_80DD(i)) <> 12 Then
       MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
           ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "000000000000" Then
            MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
            ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "111111111111" Then
            MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
            ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
    End If
     
Next

UpdateProgressBar
End Function

Function ValidateTypedependent_80DD() As Boolean
  ValidateTypedependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("Typedependent_80DD").Cells
    Dim i As Long
    end80DD = 1
    ReDim Type_dependent_80DD(end80DD)
    
    For i = 1 To end80DD

        Type_dependent_80DD(i) = rangecells.item(i).value
        
        If isdropdownblank(Type_dependent_80DD(i)) Or UCase(Type_dependent_80DD(i)) = UCase("Select") Then
           MsgBox_80DD = MsgBox_80DD + "*Please select one of the dropdown in 'Type of dependent' in schedule 80DD" & Chr(13)
           ValidateTypedependent_80DD = False
           Exit Function
        
        End If
        UpdateProgressBar
    Next
End Function
Sub Validate80DD()
    If Not Validate80DD_1 Then
        Sheet58.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80DD)
        CloseMsg
    End If
End Sub



Function Validate80DD_1() As Boolean
    Validate80DD_1 = True
    MsgBox_80DD = "Schedule 80DD : " & Chr(10)


setTableInfo80DD
'setTableInfo80DD1
setTableInfo80DD2
setTableInfo80DD3
setTableInfo80DD4
setTableInfo80DD5
setTableInfo80DD6
setTableInfo80DD7

'added by Chetan C M for AY 2025-26
setTableInfo80DD8
setTableInfo80DD9

end80DD = WorksheetFunction.Max(0, end80DD, end80DD2, end80DD3, end80DD4, end80DD5, end80DD6, end80DD7, end80DD8, end80DD9)

If end80DD > 0 Then
If Not ValidateNature_disability_80DD Then Validate80DD_1 = False
If Not ValidateType_disability_80DD Then Validate80DD_1 = False 'added by Chetan C M for AY 2025-26
'If Not ValidateAmount_of_deduction_80DD Then Validate80DD_1 = False
If Not ValidateTypedependent_80DD Then Validate80DD_1 = False 'SIT-93744 updated by Shrutika
If Not ValidateDate_of_filingofForm10IA_80DD Then Validate80DD_1 = False
If Not ValidateAckNoFm10IAfiled_80DD Then Validate80DD_1 = False
If Not ValidateAckNo_ofFormRule11A_80DD Then Validate80DD_1 = False 'added by Chetan C M for AY 2025-26
If Not ValidateUDIDNum_80DD Then Validate80DD_1 = False
If Not ValidatePANdependent_80DD Then Validate80DD_1 = False
If Not ValidateAadhaardependent_80DD Then Validate80DD_1 = False

End If
End Function


Function ChkAdhaarNo_80DD() As Boolean
On Error Resume Next
    ChkAdhaarNo_80DD = True
    Dim AadharNumber As Variant
    
    
    AadharNumber = Trim(Sheet58.Range("Aadhaardependent_80DD"))
    
    
    
    If AadharNumber <> "" Then
        If Not IsNumeric(AadharNumber) Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo_80DD = False
            Exit Function
        End If
    
        If AadharNumber = "000000000000" Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo_80DD = False
            Exit Function
        End If
        
        If AadharNumber < 0 Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo_80DD = False
            Exit Function
        End If
        
        If AadharNumber = "111111111111" Then
            'errmsgAdhaar = "is invalid"
            ChkAdhaarNo_80DD = False
            Exit Function
        End If
        
    
        If Len(AadharNumber) <> 12 Then
            errmsgAdhaar = "For Aadhaar Number- 12 digit."
            ChkAdhaarNo_80DD = False
            Exit Function
        End If
    End If
End Function
