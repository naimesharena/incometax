Attribute VB_Name = "Sch80U_80DD"

Option Explicit

Public MsgBox_80U, MsgBox_80DD As String

 
Public Nature_of_disability_80U As Variant

Public Type_of_disability_80U As Variant

Public Amount_of_deduction_80U As Variant
Public Date_filingFm10IA_80U As Variant
Public AckNo_ofForm10IAfiled_80U As Variant
Public AckNo_ofForm11IAfiled_80U As Variant
Public UDID_Num_80U As Variant
Public rngname_80U As Variant
Public end80U, end80U1, end80U2, end80U3, end80U4, end80U5 As Long


'--------------------80DD
Public end80DD8, end80DD7, end80DD6, end80DD5, end80DD4, end80DD3, end80DD2, end80DD1, end80DD   As Long

Public Type_of_disability_80DD As Variant

Public Type_dependent_80DD As Variant
Public Aadhaar_dependent_80DD As Variant
Public PAN_dependent_80DD  As Variant
Public UDID_Num_80DD As Variant
Public AckNo_ofForm10IAfiled_80DD As Variant
Public Date_filingFm10IA_80DD As Variant
Public Date_filingFm11IA_80DD As Variant
Public Amount_of_deduction_80DD As Variant
Public Nature_of_disability_80DD As Variant

Sub ValidateSheet80U_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 3: AY: 2025-26"
  Validate80U
  'MsgBox "Sheet 80U is OK", vbOKOnly, vbMessgaeCaption
   fmsgbox "Sheet 80U is OK"
End Sub
Sub ValidateSheet80DD_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 3: AY: 2025-26"
Validate80DD
'MsgBox "Sheet 80U is OK", vbOKOnly, vbMessgaeCaption
fmsgbox "Sheet 80DD is OK"
End Sub

'---------------------------------------------------------------------------------------------
Sub Validate80U()
    If Not Validate80U_1 Then
        Sheet59.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80U)
        CloseMsg
    End If
End Sub



Function Validate80U_1() As Boolean
    Validate80U_1 = True
    MsgBox_80U = "Schedule 80U : " & Chr(10)
         
setTableInfo80U
setTableInfo80U1
setTableInfo80U2
setTableInfo80U3
setTableInfo80U4
setTableInfo80U5


    end80U = WorksheetFunction.Max(0, end80U, end80U1, end80U2, end80U3, end80U4, end80U5)

 If end80U > 0 Then

       If Not ValidateNature_disability_80U Then Validate80U_1 = False
       
       If Not ValidateType_disability_80U Then Validate80U_1 = False 'Added by Aavula Naresh for AY 2025-26
       
       'If Not ValidateAmount_of_deduction_80U Then Validate80U_1 = False
       
       If Not ValidateDate_of_filingofForm10IA_80U Then Validate80U_1 = False
       If Not ValidateAckNoFm10IAfiled_80U Then Validate80U_1 = False
       
       If Not ValidateFormAckNum11A_80U_80U Then Validate80U_1 = False 'Added by Aavula Naresh for AY 2025-26

       If Not ValidateUDIDNum_80U Then Validate80U_1 = False
End If
        
End Function


Sub setTableInfo80U()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("Naturedisability_80U").Cells
    mIntCells = Sheet59.Range("Naturedisability_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80U = ccount
    'rngname_80U = "Naturedisability_80U;Amtdeduction_80U;DatefilingFm10IA_80U;AckNoFm10IAfiled_80U;UDIDNum_80U;"
End Sub
Sub setTableInfo80U1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("Disability_80U").Cells
    mIntCells = Sheet59.Range("Disability_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80U1 = ccount
End Sub

'Sub setTableInfo80U1()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet59.Range("Amtdeduction_80U").Cells
'    mIntCells = Sheet59.Range("Amtdeduction_80U").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80U1 = ccount
'End Sub


Sub setTableInfo80U2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("DatefilingFm10IA_80U").Cells
    mIntCells = Sheet59.Range("DatefilingFm10IA_80U").count
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
    Set rangecells = Sheet59.Range("AckNoFm10IAfiled_80U").Cells
    mIntCells = Sheet59.Range("AckNoFm10IAfiled_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U3 = ccount
End Sub
Sub setTableInfo80U4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("FormAckNum11A_80U").Cells
    mIntCells = Sheet59.Range("FormAckNum11A_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U4 = ccount
End Sub

Sub setTableInfo80U5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("UDIDNum_80U").Cells
    mIntCells = Sheet59.Range("UDIDNum_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80U5 = ccount
End Sub

Function ValidateNature_disability_80U() As Boolean
   ValidateNature_disability_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Naturedisability_80U").Cells
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
Function ValidateType_disability_80U() As Boolean
   ValidateType_disability_80U = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Disability_80U").Cells
    Dim i As Long
    ReDim Type_of_disability_80U(end80U)
    For i = 1 To end80U

        Type_of_disability_80U(i) = rangecells.item(i).value
        If isdropdownblankDD(Type_of_disability_80U(i)) Or UCase(Type_of_disability_80U(i)) = UCase("(Select)") Then
            MsgBox_80U = MsgBox_80U + "* Selection of ""Type of disability"" in schedule 80U is mandatory." & Chr(13)
            
           ValidateType_disability_80U = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function
 
Function ValidateAmount_of_deduction_80U() As Boolean
    ValidateAmount_of_deduction_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Amtdeduction_80U").Cells
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
    Set rangecells = Sheet59.Range("DatefilingFm10IA_80U").Cells
    Dim i As Long
    ReDim Date_filingFm10IA_80U(end80U)
    
    For i = 1 To end80U
        Date_filingFm10IA_80U(i) = rangecells.item(i).value
        
        'Added by Aavula Naresh for AY 2025-26 New Enhancement
            If Not chkCompulsory(Date_filingFm10IA_80U(i)) Then
             If Sheet59.Range("DatefilingFm10IA_80U").MergeArea.Locked = False Then
            MsgBox_80U = MsgBox_80U + "* ""Please enter the date of filing of Form 10IA""" & Chr(13)
            ValidateDate_of_filingofForm10IA_80U = False
            Exit Function
        End If
        End If
        ''''''''
        UpdateProgressBar

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
     
Next
End Function

 

Function ValidateAckNoFm10IAfiled_80U() As Boolean
    
ValidateAckNoFm10IAfiled_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("AckNoFm10IAfiled_80U").Cells
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80U(end80U)
    
    For i = 1 To end80U
        AckNo_ofForm10IAfiled_80U(i) = rangecells.item(i).value
        
        'Added by Aavula Naresh for AY 2025-26 New Enhancement
            If Not chkCompulsory(AckNo_ofForm10IAfiled_80U(i)) Then
             If Sheet59.Range("AckNoFm10IAfiled_80U").MergeArea.Locked = False Then
            MsgBox_80U = MsgBox_80U + "* ""Please enter the Acknowledgement number of Form 10IA""" & Chr(13)
            ValidateAckNoFm10IAfiled_80U = False
            Exit Function
        End If
        End If
        ''''''''
     
      If Not checkallfieldspecialcharacter_80U_80DD(AckNo_ofForm10IAfiled_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateAckNoFm10IAfiled_80U = False
             Exit Function
        End If
     
        'UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80U(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80U = False
        MsgBox_80U = MsgBox_80U + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
Next
End Function

'Added by Aavula Naresh for AY 2025-26 New Enhancement

Function ValidateFormAckNum11A_80U_80U() As Boolean
    
ValidateFormAckNum11A_80U_80U = True
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("FormAckNum11A_80U").Cells
    Dim i As Long
    ReDim AckNo_ofForm11IAfiled_80U(end80U)
    
    For i = 1 To end80U
        AckNo_ofForm11IAfiled_80U(i) = rangecells.item(i).value
        
        'Added by Aavula Naresh for AY 2025-26 New Enhancement
            If Not chkCompulsory(AckNo_ofForm11IAfiled_80U(i)) Then
             If Sheet59.Range("FormAckNum11A_80U").MergeArea.Locked = False Then
            MsgBox_80U = MsgBox_80U + "* ""Please enter the Acknowledgement Number of Form as per Rule 11A(2)(ii)""" & Chr(13)
            ValidateFormAckNum11A_80U_80U = False
            Exit Function
        End If
        End If
        ''''''''
     
      If Not checkallfieldspecialcharacter_80U_80DD(AckNo_ofForm11IAfiled_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* Ack. No. of Form 11IA filed cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateFormAckNum11A_80U_80U = False
             Exit Function
        End If

    If Len(AckNo_ofForm11IAfiled_80U(i)) > 15 Then
        ValidateFormAckNum11A_80U_80U = False
        MsgBox_80U = MsgBox_80U + "*  Ack. No. of Form 11IA filed cannot exceed 15 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
Next
End Function
 
Function ValidateUDIDNum_80U() As Boolean
    
ValidateUDIDNum_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("UDIDNum_80U").Cells
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

    If Len(UDID_Num_80U(i)) > 18 Then
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
    Set rangecells = Sheet59.Range("Naturedisability_80DD").Cells
    mIntCells = Sheet59.Range("Naturedisability_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80DD = ccount
    'rngname_80GGA = "RelevantClauseClaimed_80GGA;Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_total_80GGA;Donation_Eligible_80GGA;"
End Sub

 
Sub setTableInfo80DD1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("Disability_80DD").Cells
    mIntCells = Sheet59.Range("Disability_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
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
    Set rangecells = Sheet59.Range("Typedependent_80DD").Cells
    mIntCells = Sheet59.Range("Typedependent_80DD").count
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
    Set rangecells = Sheet59.Range("PANdependent_80DD").Cells
    mIntCells = Sheet59.Range("PANdependent_80DD").count
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
    Set rangecells = Sheet59.Range("Aadhaardependent_80DD").Cells
    mIntCells = Sheet59.Range("Aadhaardependent_80DD").count
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
    Set rangecells = Sheet59.Range("DatefilingFm10IA_80DD").Cells
    mIntCells = Sheet59.Range("DatefilingFm10IA_80DD").count
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
    Set rangecells = Sheet59.Range("AckNoFm10IAfiled_80DD").Cells
    mIntCells = Sheet59.Range("AckNoFm10IAfiled_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD6 = ccount
End Sub
Sub setTableInfo80DD7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("FormAckNum11A_80DD").Cells
    mIntCells = Sheet59.Range("FormAckNum11A_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD7 = ccount
End Sub

Sub setTableInfo80DD8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet59.Range("UDIDNum_80DD").Cells
    mIntCells = Sheet59.Range("UDIDNum_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD8 = ccount
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------

'-------------------------------------
Function ValidateNature_disability_80DD() As Boolean
   ValidateNature_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Naturedisability_80DD").Cells
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

Function ValidateType_disability_80DD() As Boolean
   ValidateType_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Disability_80DD").Cells
    Dim i As Long
    ReDim Type_of_disability_80DD(end80DD)
    For i = 1 To end80DD

        Type_of_disability_80DD(i) = rangecells.item(i).value
        If isdropdownblankDD(Type_of_disability_80DD(i)) Or UCase(Type_of_disability_80DD(i)) = UCase("(Select)") Then
            MsgBox_80DD = MsgBox_80DD + "* Selection of ""Type of disability"" in schedule 80DD is mandatory." & Chr(13)
            
           ValidateType_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function
'-----------------------------------------------------
Function ValidateAmount_of_deduction_80DD() As Boolean
    ValidateAmount_of_deduction_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Amtdeduction_80DD").Cells
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
    Set rangecells = Sheet59.Range("DatefilingFm10IA_80DD").Cells
    Dim i As Long
    ReDim Date_filingFm10IA_80DD(end80DD)
    
    For i = 1 To end80DD
        Date_filingFm10IA_80DD(i) = rangecells.item(i).value
        
             'Added by Aavula Naresh for AY 2025-26 New Enhancement
            If Not chkCompulsory(Date_filingFm10IA_80DD(i)) Then
             If Sheet59.Range("DatefilingFm10IA_80DD").MergeArea.Locked = False Then
            MsgBox_80DD = MsgBox_80DD + "* ""Please enter the date of filing of Form 10IA""" & Chr(13)
            ValidateDate_of_filingofForm10IA_80DD = False
            Exit Function
        End If
        End If
        ''''''''
        
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
     
Next
End Function

'---------------------------------------------------------------------

Function ValidateAckNoFm10IAfiled_80DD() As Boolean
    
ValidateAckNoFm10IAfiled_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("AckNoFm10IAfiled_80DD").Cells
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80DD(end80DD)
    
    For i = 1 To end80DD
        AckNo_ofForm10IAfiled_80DD(i) = rangecells.item(i).value
        
              'Added by Aavula Naresh for AY 2025-26 New Enhancement
            If Not chkCompulsory(AckNo_ofForm10IAfiled_80DD(i)) Then
             If Sheet59.Range("AckNoFm10IAfiled_80DD").MergeArea.Locked = False Then
            MsgBox_80DD = MsgBox_80DD + "* ""Please enter the Acknowledgement number of Form 10IA""" & Chr(13)
            ValidateAckNoFm10IAfiled_80DD = False
            Exit Function
        End If
        End If
        ''''''''
     
      If Not checkallfieldspecialcharacter_80U_80DD(AckNo_ofForm10IAfiled_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateAckNoFm10IAfiled_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80DD(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     
Next
End Function

'Added by Aavula Naresh for AY 2025-26 New Enhancement

Function ValidateFormAckNum11A_80D_80DD() As Boolean
    
ValidateFormAckNum11A_80D_80DD = True
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("FormAckNum11A_80DD").Cells
    Dim i As Long
    ReDim AckNo_ofForm11IAfiled_80DD(end80DD)
    
    For i = 1 To end80DD
        AckNo_ofForm11IAfiled_80DD(i) = rangecells.item(i).value
        
        'Added by Aavula Naresh for AY 2025-26 New Enhancement
            If Not chkCompulsory(AckNo_ofForm11IAfiled_80DD(i)) Then
             If Sheet59.Range("FormAckNum11A_80DD").MergeArea.Locked = False Then
            MsgBox_80DD = MsgBox_80DD + "* ""Please enter the Acknowledgement Number of Form as per Rule 11A(2)(ii)""" & Chr(13)
            ValidateFormAckNum11A_80D_80DD = False
            Exit Function
        End If
        End If
        ''''''''
     
      If Not checkallfieldspecialcharacter_80U_80DD(AckNo_ofForm11IAfiled_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Ack. No. of Form 11IA filed cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateFormAckNum11A_80D_80DD = False
             Exit Function
        End If

    If Len(AckNo_ofForm11IAfiled_80DD(i)) > 15 Then
        ValidateFormAckNum11A_80D_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*  Ack. No. of Form 11IA filed cannot exceed 15 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     
Next
End Function
'------------------------------------------------------------------------------------------------------------------------------
Function ValidateUDIDNum_80DD() As Boolean
    
ValidateUDIDNum_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("UDIDNum_80DD").Cells
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

    If Len(UDID_Num_80DD(i)) > 18 Then
        ValidateUDIDNum_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*UDID Number cannot exceed 18 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     
Next
End Function

'Function ValidatePANdependent_80DD() As Boolean
'
'ValidatePANdependent_80DD = True
''    setTableInfo80GGA
'    Dim rangecells As Range
'    Set rangecells = Sheet59.Range("PANdependent_80DD").Cells
'    Dim i As Long
'    ReDim PAN_dependent_80DD(end80DD)
'
'    For i = 1 To end80DD
'        PAN_dependent_80DD(i) = rangecells.item(i).value
'
''        If Not chkCompulsory(PAN_dependent_80DD(i)) Then
''            MsgBox_80DD = MsgBox_80DD + "* Please enter pan in Schedule 80DD" & Chr(13)
''            ValidatePANdependent_80DD = False
''            Exit Function
''        End If
''        UpdateProgressBar
'
'If Trim(PAN_dependent_80DD(i)) <> "" Then
'
'      If Not checkfieldspecialcharacter(PAN_dependent_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "*Invalid PAN in Schedule 80DD." & Chr(13)
'            ValidatePANdependent_80DD = False
'             Exit Function
'        End If
'
'        UpdateProgressBar
'
'    If Not CheckPAB_80DD(UCase(PAN_dependent_80DD(i))) Then
'             MsgBox_80DD = MsgBox_80DD + "*Invalid PAN in Schedule 80DD." & Chr(13)
'             ValidatePANdependent_80DD = False
'             Exit Function
'        End If
'
''If PAN_dependent_80DD(i) <> o Then
'        If ((UCase(PAN_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(PAN_dependent_80DD(i)) = UCase(Sheet52.Range("sheet9.PAN").value))) Then
'        MsgBox_80DD = MsgBox_80DD + "* PAN of the dependent cannot be same as assessee PAN in Part-A General Information." & Chr(13)
'         ValidatePANdependent_80DD = False
'        Exit Function
'        End If
''End If
'
' End If
'
'Next
'End Function
'
'
'Function ValidateAadhaardependent_80DD() As Boolean
'
'ValidateAadhaardependent_80DD = True
'
'    Dim rangecells As Range
'    Set rangecells = Sheet59.Range("Aadhaardependent_80DD").Cells
'    Dim i As Long
'    ReDim Aadhaar_dependent_80DD(end80DD)
'
'    For i = 1 To end80DD
'        Aadhaar_dependent_80DD(i) = rangecells.item(i).value
'
'
'If Trim(Aadhaar_dependent_80DD(i)) <> "" Then
'
'
'      If Not checkfieldspecialcharacter(Aadhaar_dependent_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "*Invalid Aadhaar in Schedule 80DD." & Chr(13)
'            ValidateAadhaardependent_80DD = False
'             Exit Function
'        End If
'
'
'      If UCase(Aadhaar_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.adhaarno").value) Then
'        MsgBox_80DD = MsgBox_80DD + "* Aadhaar of the dependent cannot be same as assessee Aadhaar in Part-A General Information." & Chr(13)
'         ValidateAadhaardependent_80DD = False
'        Exit Function
'        End If
'
'        If Not IsNumeric(Aadhaar_dependent_80DD(i)) Then
'
'           MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
'           ValidateAadhaardependent_80DD = False
'            Exit Function
'        End If
'
'        If Len(Aadhaar_dependent_80DD(i)) <> 12 Then
'       MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
'           ValidateAadhaardependent_80DD = False
'            Exit Function
'        End If
'
'        If Aadhaar_dependent_80DD(i) = "000000000000" Then
'            MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
'            ValidateAadhaardependent_80DD = False
'            Exit Function
'        End If
'
'        If Aadhaar_dependent_80DD(i) = "111111111111" Then
'            MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
'            ValidateAadhaardependent_80DD = False
'            Exit Function
'        End If
'
'    End If
'
'Next
'
'UpdateProgressBar
'End Function
Function ValidatePan_Aadhar_80DD() As Boolean
ValidatePan_Aadhar_80DD = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet59.Range("PANdependent_80DD").Cells
    Set rangecells1 = Sheet59.Range("Aadhaardependent_80DD").Cells
    Dim i As Long
    ReDim PAN_dependent_80DD(end80DD)
    ReDim Aadhaar_dependent_80DD(end80DD)
    
    For i = 1 To end80DD
        PAN_dependent_80DD(i) = rangecells.item(i).value
        Aadhaar_dependent_80DD(i) = rangecells1.item(i).value

            If (PAN_dependent_80DD(i) = "" And Aadhaar_dependent_80DD(i) = "") Then
                MsgBox_80DD = MsgBox_80DD + "* Please Enter Either ""PAN"" or ""Aadhaar"" in schedule 80DD"" " & Chr(13)
                ValidatePan_Aadhar_80DD = False
                Exit Function
            End If

       If Trim(PAN_dependent_80DD(i)) <> "" Then
     
      If Not checkfieldspecialcharacter(PAN_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "*Invalid PAN in Schedule 80DD." & Chr(13)
            ValidatePan_Aadhar_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

 If Not CheckPAB_80DD(UCase(PAN_dependent_80DD(i))) Then
             MsgBox_80DD = MsgBox_80DD + "*Invalid PAN in Schedule 80DD." & Chr(13)
             ValidatePan_Aadhar_80DD = False
             Exit Function
        End If


       If ((UCase(PAN_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(PAN_dependent_80DD(i)) = UCase(Sheet52.Range("sheet9.PAN").value))) Then
        MsgBox_80DD = MsgBox_80DD + "* PAN of the dependent cannot be same as assessee PAN in Part-A General Information." & Chr(13)
        ValidatePan_Aadhar_80DD = False
        Exit Function
        End If

If Trim(Aadhaar_dependent_80DD(i)) <> "" Then


      If Not checkfieldspecialcharacter(Aadhaar_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "*Invalid Aadhaar in Schedule 80DD." & Chr(13)
            ValidatePan_Aadhar_80DD = False
             Exit Function
        End If


      If UCase(Aadhaar_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.adhaarno").value) Then
        MsgBox_80DD = MsgBox_80DD + "* Aadhaar of the dependent cannot be same as assessee Aadhaar in Part-A General Information." & Chr(13)
         ValidatePan_Aadhar_80DD = False
        Exit Function
        End If
     
        If Not IsNumeric(Aadhaar_dependent_80DD(i)) Then
            
           MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
           ValidatePan_Aadhar_80DD = False
            Exit Function
        End If
        
        If Len(Aadhaar_dependent_80DD(i)) <> 12 Then
       MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
           ValidatePan_Aadhar_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "000000000000" Then
            MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
            ValidatePan_Aadhar_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "111111111111" Then
            MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD" & Chr(13)
            ValidatePan_Aadhar_80DD = False
            Exit Function
        End If
        
    End If
End If
    Next
End Function


Function ValidateTypedependent_80DD() As Boolean
  ValidateTypedependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet59.Range("Typedependent_80DD").Cells
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
        Sheet59.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80DD)
        CloseMsg
    End If
End Sub



Function Validate80DD_1() As Boolean
    Validate80DD_1 = True
    MsgBox_80DD = "Schedule 80DD : " & Chr(10)


setTableInfo80DD
setTableInfo80DD1
setTableInfo80DD2
setTableInfo80DD3
setTableInfo80DD4
setTableInfo80DD5
setTableInfo80DD6
setTableInfo80DD7
setTableInfo80DD8

end80DD = WorksheetFunction.Max(0, end80DD, end80DD1, end80DD2, end80DD3, end80DD4, end80DD5, end80DD6, end80DD7, end80DD8)

If end80DD > 0 Then
If Not ValidateNature_disability_80DD Then Validate80DD_1 = False
If Not ValidateType_disability_80DD Then Validate80DD_1 = False 'Added by Aavula Naresh for AY 2025-26
'If Not ValidateAmount_of_deduction_80DD Then Validate80DD_1 = False
If Not ValidateTypedependent_80DD Then Validate80DD_1 = False
If Not ValidateDate_of_filingofForm10IA_80DD Then Validate80DD_1 = False
If Not ValidateAckNoFm10IAfiled_80DD Then Validate80DD_1 = False

If Not ValidateFormAckNum11A_80D_80DD Then Validate80DD_1 = False 'Added by Aavula Naresh for AY 2025-26

If Not ValidateUDIDNum_80DD Then Validate80DD_1 = False
'If Not ValidatePANdependent_80DD Then Validate80DD_1 = False
'If Not ValidateAadhaardependent_80DD Then Validate80DD_1 = False

If Not ValidatePan_Aadhar_80DD Then Validate80DD_1 = False



End If
End Function


Function ChkAdhaarNo_80DD() As Boolean
On Error Resume Next
    ChkAdhaarNo_80DD = True
    Dim AadharNumber As Variant
    
    
    AadharNumber = Trim(Sheet59.Range("Aadhaardependent_80DD"))
    
    
    
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

