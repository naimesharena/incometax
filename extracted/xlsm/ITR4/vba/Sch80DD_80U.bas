Attribute VB_Name = "Sch80DD_80U"
Option Explicit

 

Public Nature_of_disability_80DD As Variant        'Nature80DD(1)
Public Amount_of_deduction_80DD As Variant         'Deduction_80DD(2)
Public Type_dependent_80DD As Variant               'Dependent80DD(3
Public Pan_Dependent_80DD As Variant               'PAN80DD(4)
Public Aadhaar_dependent_80DD As Variant           'Aadhaar80DD(5)
Public Date_filingFm10IA_80DD As Variant           'Date10IA80DD(6)
Public AckNo_ofForm10IAfiled_80DD As Variant       'Acknowledge_No80DD(7)
Public UDID_Num_80DD As Variant                    'UID_Num80DD(8)


Public end80DD, end80DD1, end80DD2, end80DD3, end80DD4, end80DD5, end80DD6, end80DD7, end80DD8 As Variant
'---------------80DD--------------------

'--------------80U----------------------------
Public MsgBox_80U, MsgBox_80DD As String

 
Public Nature_of_disability_80U As Variant
Public Amount_of_deduction_80U As Variant
Public Date_filingFm10IA_80U As Variant
Public AckNo_ofForm10IAfiled_80U As Variant
Public UDID_Num_80U As Variant
Public rngname_80U As Variant
Public end80U, end80U1, end80U2, end80U3, end80U4, end80U5 As Long
Sub ValidateSheet80DD_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2025-26"
    Validate80DD

fmsgboxoK "Sheet 80DD is OK"
End Sub
Sub ValidateSheet80U_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2025-26"
    Validate80U

fmsgboxoK "Sheet 80U is OK"
End Sub
Sub Validate80DD()
    If Not Validate80DD_1 Then
        Sheet14.Activate
        fmsgbox (MsgBox_80DD)
        CloseMsg
    End If
    
  If Sheet14.Range("AcknowledgeNum_80DD").value <> "" Or Sheet14.Range("AckNumber10IA_80U").value <> "" Then
    
    If Sheet14.Range("AcknowledgeNum_80DD").value = Sheet14.Range("AckNumber10IA_80U").value Then
        MsgBox_80DD = MsgBox_80DD & "* Acknowledgement number of Form 10IA filed for self and Dependent can't be same. Please provide proper acknowledgement number" & Chr(13)
       fmsgbox MsgBox_80DD
       CloseMsg
    End If
  End If
    
End Sub

Function Validate80DD_1() As Boolean
    Validate80DD_1 = True
    
    MsgBox_80DD = "Schedule 80DD : " & Chr(10)
        setTableInfo80DD
       ' setTableInfo80DD1
        setTableInfo80DD2
        setTableInfo80DD3
        setTableInfo80DD4
        'Commented by Ayush_06/05/2025 for 0.7 changes
     '   setTableInfo80DD5
        setTableInfo80DD6
        setTableInfo80DD7
        setTableInfo80DD8
        
     'Commented by Ayush_06/05/2025 for 0.7 changes
  '  end80DD = WorksheetFunction.Max(0, end80DD, end80DD2, end80DD3, end80DD4, end80DD5, end80DD6, end80DD7, end80DD8)
    end80DD = WorksheetFunction.Max(0, end80DD, end80DD2, end80DD3, end80DD4, end80DD6, end80DD7, end80DD8)

        If end80DD > 0 Then
        
            If Not ValidateNature_disability_80DD Then Validate80DD_1 = False
            'Nature80DD(1)
            'Ayush
            
             If Not ValidateType_disability_80DD Then Validate80DD_1 = False
            
            If Not ValidateAmount_of_deduction_80DD Then Validate80DD_1 = False         'Amount_Deduction80DD(2)
            
            If Not ValidateTypedependent_80DD Then Validate80DD_1 = False               'Dependent80DD(3)
            
            If Not ValidatePan1_80DD Then Validate80DD_1 = False                         'PAN80DD(4)
            
            If Not ValidateAadhaardependent_80DD Then Validate80DD_1 = False             'Aadhaar80DD(5)
            
           ' If Not ValidateDate_of_filingofForm10IA_80DD Then Validate80DD_1 = False     'Date10IA80DD(6)
            
            If Not ValidateAckNoFm10IAfiled_80DD Then Validate80DD_1 = False             'Acknowledge_No80DD(7)
            
            'ValidateAckNoFm11A2filed_80DD
            'Ayush_20
            'Commented by Ayush_06/05/2025 for 0.7 changes
            
          '  If Not ValidateAckNoFm11A2filed_80DD Then Validate80DD_1 = False
            
            If Not ValidateUDIDNum_80DD Then Validate80DD_1 = False                      'UID_Num80DD(8)

        End If
        
End Function
Sub setTableInfo80DD() 'Nature80DD(1)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("NatureDisability_80DD").Cells
    mIntCells = Sheet14.Range("NatureDisability_80DD").count
    For mIntCtr = 1 To mIntCells
         If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
          ccount = ccount + 1
        End If
    Next
    end80DD = ccount
'    rngname_80DD = "NatureDisability_80DD;AmtDeduction_80DD;Dependent_80DD;PanDependent_80DD;AadhaarDependent_80DD;Form10IA_80DD;AcknowledgeNum_80DD;UDIDNumber_80DD"
End Sub
Sub setTableInfo80DD1() 'Amount_Deduction80DD(2)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AmtDeduction_80DD").Cells
    mIntCells = Sheet14.Range("AmtDeduction_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "0" Then
        ccount = ccount + 1
        End If
    Next
    end80DD1 = ccount
End Sub
Sub setTableInfo80DD2() 'Dependent80DD(3)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Dependent_80DD").Cells
    mIntCells = Sheet14.Range("Dependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80DD2 = ccount
End Sub
Sub setTableInfo80DD3() 'PAN80DD(4)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("PanDependent_80DD").Cells
    mIntCells = Sheet14.Range("PanDependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80DD3 = ccount
End Sub
Sub setTableInfo80DD4() 'Aadhaar80DD(5)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AadhaarDependent_80DD").Cells
    mIntCells = Sheet14.Range("AadhaarDependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80DD4 = ccount
End Sub
'Commented by Ayush_06/05/2025 for 0.7 changes
'Sub setTableInfo80DD5() 'Date10IA80DD(6)
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet14.Range("Form10IA_80DD").Cells
'    mIntCells = Sheet14.Range("Form10IA_80DD").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80DD5 = ccount
'End Sub
Sub setTableInfo80DD6() 'Acknowledge_No80DD(7)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AcknowledgeNum_80DD").Cells
    mIntCells = Sheet14.Range("AcknowledgeNum_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80DD6 = ccount
End Sub
Sub setTableInfo80DD7() 'UID_Num80DD(8)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("UDIDNumber_80DD").Cells
    mIntCells = Sheet14.Range("UDIDNumber_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80DD7 = ccount
End Sub
'Ayush
Sub setTableInfo80DD8() 'TypeDisability80DD(1)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Disability_80DD").Cells
    mIntCells = Sheet14.Range("Disability_80DD").count
    For mIntCtr = 1 To mIntCells
         If Not isdropdownblankDD(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
          ccount = ccount + 1
        End If
    Next
    end80DD8 = ccount
'    rngname_80DD = "NatureDisability_80DD;AmtDeduction_80DD;Dependent_80DD;PanDependent_80DD;AadhaarDependent_80DD;Form10IA_80DD;AcknowledgeNum_80DD;UDIDNumber_80DD"
End Sub
'Nature80DD(1)
Function ValidateNature_disability_80DD() As Boolean
   ValidateNature_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("NatureDisability_80DD").Cells
    Dim i As Long
    ReDim Nature_of_disability_80DD(end80DD)
    
    For i = 1 To end80DD

        Nature_of_disability_80DD(i) = rangecells.item(i).value
        
        If isdropdownblank(Nature_of_disability_80DD(i)) Or UCase(Nature_of_disability_80DD(i)) = UCase("Select") Then
               
         
            MsgBox_80DD = MsgBox_80DD + "*Nature of Disability in schedule 80DD is mandatory " & Chr(13)
           ValidateNature_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function

Function ValidateType_disability_80DD() As Boolean
   ValidateType_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Disability_80DD").Cells
    Dim i As Long
    ReDim Type_of_disability_80DD(end80DD)
    
    For i = 1 To end80DD

        Type_of_disability_80DD(i) = rangecells.item(i).value
        
        If isdropdownblankDD(Type_of_disability_80DD(i)) Or UCase(Type_of_disability_80DD(i)) = UCase("(Select)") Then
               
         
            MsgBox_80DD = MsgBox_80DD + "*Selection of ""Type of disability"" in schedule 80DD is mandatory." & Chr(13)
           ValidateType_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function
'Amount_Deduction80DD(2)
Function ValidateAmount_of_deduction_80DD() As Boolean
    ValidateAmount_of_deduction_80DD = True

    Dim rangecells As Range
    Set rangecells = Sheet14.Range("AmtDeduction_80DD").Cells
    Dim i As Long
    ReDim Amount_of_deduction_80DD(end80DD)
    For i = 1 To end80DD
        Amount_of_deduction_80DD(i) = rangecells.item(i).value
        If Not chkCompulsory(Amount_of_deduction_80DD(i)) Then

            MsgBox_80DD = MsgBox_80DD + "* Amount of deduction in schedule 80DD is mandatory" & Chr(13)
            ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        UpdateProgressBar
 
        If Not IsNumeric(Amount_of_deduction_80DD(i)) Then
            MsgBox_80DD = MsgBox_80DD + "*  Amount of deduction in schedule 80DD should be Numeric value" & Chr(13)
            ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        If Amount_of_deduction_80DD(i) > 99999999999999# Then
           MsgBox_80DD = MsgBox_80DD + "*  Amount of deduction in schedule 80DD cannot exceed 14 digits" & Chr(13)
           ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
 
Next
 
End Function
'Dependent80DD(3)
Function ValidateTypedependent_80DD() As Boolean
  ValidateTypedependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Dependent_80DD").Cells
    Dim i As Long
    end80DD = 1
    ReDim Type_dependent_80DD(end80DD)
    
    For i = 1 To end80DD

        Type_dependent_80DD(i) = rangecells.item(i).value
        
        If isdropdownblank(Type_dependent_80DD(i)) Or UCase(Type_dependent_80DD(i)) = UCase("Select") Then
           MsgBox_80DD = MsgBox_80DD + "*Type of dependent in schedule 80DD is mandatory" & Chr(13)
           ValidateTypedependent_80DD = False
           Exit Function
        
        End If
        
        
        'Ayush_legacy1
        If Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Or Sheet1.Range("sheet1.Status").value = "F - FIRM(Other than LLP)" Then
        If Type_dependent_80DD(i) = "8-Member of the HUF (in case of HUF)" Then
        MsgBox_80DD = MsgBox_80DD + "*Only HUF can select  Type of dependent as ""Member of the HUF"". Please select applicable dropdown in ""Type of Dependent""." & Chr(13)
           ValidateTypedependent_80DD = False
           Exit Function
        
        End If
        End If
        
        
        UpdateProgressBar
    Next
End Function
'PAN80DD(4)
Function ValidatePan1_80DD() As Boolean
ValidatePan1_80DD = True
   
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("PanDependent_80DD").Cells
    Dim i As Long
    ReDim Pan_Dependent_80DD(end80DD)
    
    Dim rangecells1 As Range
    Set rangecells1 = Sheet14.Range("AadhaarDependent_80DD").Cells
    ReDim Aadhaar_dependent_80DD(end80DD)
    
    
    For i = 1 To end80DD
        Pan_Dependent_80DD(i) = rangecells.item(i).value
        
        Aadhaar_dependent_80DD(i) = rangecells1.item(i).value
        
       If Not chkCompulsory(Aadhaar_dependent_80DD(i)) Then
        If Not chkCompulsory(Pan_Dependent_80DD(i)) Then
'              MsgBox_80DD = MsgBox_80DD + "* PAN is mandatory in schedule 80DD" & Chr(13)
                MsgBox_80DD = MsgBox_80DD + "* Please enter either PAN or Aadhaar of dependent in schedule 80DD." & Chr(13)
            ValidatePan1_80DD = False
            Exit Function
        End If
    End If
     If Not chkCompulsory(Pan_Dependent_80DD(i)) Then
        If Not chkCompulsory(Aadhaar_dependent_80DD(i)) Then
            MsgBox_80DD = MsgBox_80DD + "* Aadhaar is mandatory in schedule 80DD" & Chr(13)
            ValidatePan1_80DD = False
            Exit Function
        End If
    End If

       
    If Not checkfieldspecialcharacter(Pan_Dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Invalid PAN in Schedule 80DD" & Chr(13)
             ValidatePan1_80DD = False
             Exit Function
    End If
        If Not Check80DDPAN(UCase(Pan_Dependent_80DD(i))) Then
             MsgBox_80DD = MsgBox_80DD + "* Invalid PAN in Schedule 80DD" & Chr(13)
             ValidatePan1_80DD = False
             Exit Function
        End If
     If chkCompulsory(Pan_Dependent_80DD(i)) Then
        If ((UCase(Pan_Dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(Pan_Dependent_80DD(i)) = UCase(Sheet5.Range("Ver.PAN").value))) Then
         MsgBox_80DD = MsgBox_80DD + "* PAN of the dependent cannot be same as assessee PAN in Part-A General Information" & Chr(13)
         ValidatePan1_80DD = False
         Exit Function
        End If
    End If
    Next
End Function
'Aadhaar80DD (5)
Function ValidateAadhaardependent_80DD() As Boolean
    ValidateAadhaardependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("AadhaarDependent_80DD").Cells
    Dim i As Long
    ReDim Aadhaar_dependent_80DD(end80DD)
    
    Dim rangecells1 As Range
    Set rangecells1 = Sheet14.Range("PanDependent_80DD").Cells
    ReDim Pan_Dependent_80DD(end80DD)

    
    
    For i = 1 To end80DD
        Aadhaar_dependent_80DD(i) = rangecells.item(i).value
        Pan_Dependent_80DD(i) = rangecells1.item(i).value
        
        
    If Trim(Aadhaar_dependent_80DD(i)) <> "" Then
      If Not checkfieldspecialcharacter(Aadhaar_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
            ValidateAadhaardependent_80DD = False
             Exit Function
        End If


      If UCase(Aadhaar_dependent_80DD(i)) = UCase(Sheet1.Range("Sheet1.Aadhaar").value) Then
        MsgBox_80DD = MsgBox_80DD + "* Aadhaar of the dependent cannot be same as assessee Aadhaar in Part-A General Information." & Chr(13)
         ValidateAadhaardependent_80DD = False
        Exit Function
        End If
     
        If Not IsNumeric(Aadhaar_dependent_80DD(i)) Then
            
           MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
           ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Len(Aadhaar_dependent_80DD(i)) <> 12 Then
       MsgBox_80DD = MsgBox_80DD + "* Please enter the valid Aadhaar number in Schedule 80DD." & Chr(13)
           ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "000000000000" Then
        MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
         
            ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "111111111111" Then
       MsgBox_80DD = MsgBox_80DD + "* Invalid Aadhaar in Schedule 80DD." & Chr(13)
            
            ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
    End If

Next
UpdateProgressBar
End Function
'Date10IA80DD(6)
'Commented by Ayush_06/05/2025 for 0.7 changes
'Function ValidateDate_of_filingofForm10IA_80DD() As Boolean
'
'ValidateDate_of_filingofForm10IA_80DD = True
'
'    Dim rangecells As Range
'    Set rangecells = Sheet14.Range("Form10IA_80DD").Cells
'    Dim rangecells1 As Range
'    Set rangecells1 = Sheet14.Range("Disability_80DD").Cells
'    Dim i As Long
'    ReDim Date_filingFm10IA_80DD(end80DD)
'    ReDim Type_Disability_80DD(end80DD)
'
'
'    For i = 1 To end80DD
'        Date_filingFm10IA_80DD(i) = rangecells.item(i).value
'
''         For i = 1 To end80DD
'        Type_Disability_80DD(i) = rangecells1.item(i).value
'      'Ayush_20
'    If Type_Disability_80DD(i) = "(x) multiple disability" Or Type_Disability_80DD(i) = "(ix) cerebral palsy" Or Type_Disability_80DD(i) = "(viii) autism" Then
'    If Not chkCompulsory(Date_filingFm10IA_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "*  Please provide date of filing of Form 10IA in schedule 80DD" & Chr(13)
'            ValidateDate_of_filingofForm10IA_80DD = False
'            Exit Function
'        End If
'        End If
'
'
'    If Not CheckDateddmmyyyy(Date_filingFm10IA_80DD(i)) Then
'        ValidateDate_of_filingofForm10IA_80DD = False
'        'Ayush_10/02
''       MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA must be a valid dd/mm/yyyy format in Schedule 80DD" & Chr(13)
'        MsgBox_80DD = MsgBox_80DD + "* ""Date of filing of Form 10IA - Please enter valid date in valid format.""" & Chr(13)
'        Exit Function
'    End If
'        UpdateProgressBar
'
'    If Len(Date_filingFm10IA_80DD(i)) > 10 Then
'        ValidateDate_of_filingofForm10IA_80DD = False
'        MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA cannot exceed 10 digits in Schedule 80DD" & Chr(13)
'        Exit Function
'    End If
'    'Ayush
''    If Not ChkMinInclusiveDate(Date_filingFm10IA_80DD(i), "2025-04-01") Then
''                MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA in Schedule 80DD should not be less than 01/04/2025" & Chr(13)
''                ValidateDate_of_filingofForm10IA_80DD = False
''         Exit Function
''    End If
'
'
'    If Not ChkMinInclusiveDate(Trim(Dformat(Date_filingFm10IA_80DD(i), "yyyy-mm-dd")), "2025-04-01") Then
'           MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA cannot before 01/04/2025 " & Chr(13)
'            ValidateDate_of_filingofForm10IA_80DD = False
'
'
'          Exit Function
'        End If
'
'Next
''Next
'End Function
'Acknowledge_No80DD(7)
Function ValidateAckNoFm10IAfiled_80DD() As Boolean
    
ValidateAckNoFm10IAfiled_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("AcknowledgeNum_80DD").Cells
    Dim rangecells1 As Range
    Set rangecells1 = Sheet14.Range("Disability_80DD").Cells
    
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80DD(end80DD)
    ReDim Type_Disability_80DD(end80DD)
    
    For i = 1 To end80DD
        AckNo_ofForm10IAfiled_80DD(i) = rangecells.item(i).value
        Type_Disability_80DD(i) = rangecells1.item(i).value
     'Ayush_20
'     If Type_Disability_80DD(i) = "(x) multiple disability" Or Type_Disability_80DD(i) = "(ix) cerebral palsy" Or Type_Disability_80DD(i) = "(viii) autism" Then
        If Type_Disability_80DD(i) = "(i) autism, cerebral palsy, or multiple disabilities" Then
        If Not chkCompulsory(AckNo_ofForm10IAfiled_80DD(i)) Then 'Modified by sai on 25/04/2025
        'SIT-91940 by sai on 25/04/2025
            MsgBox_80DD = MsgBox_80DD + "* Please provide Acknowledgement of Form 10IA in schedule 80DD" & Chr(13)
             ValidateAckNoFm10IAfiled_80DD = False
             Exit Function
         End If
     End If
     
      If Not checkfieldspecialcharacter(AckNo_ofForm10IAfiled_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Ack. No. of Form 10IA filed at Sr. No  " & i & " cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateAckNoFm10IAfiled_80DD = False
             Exit Function
        End If
        
        
     
        UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80DD(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80DD = False
        MsgBox_80DD = MsgBox_80DD + "* Ack. No. of Form 10IA filed cannot exceed 15 digits at Sr. No " & i & " in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     
Next
End Function

'AYUSH_20
'Commented by Ayush_06/05/2025 for 0.7 changes
'Function ValidateAckNoFm11A2filed_80DD() As Boolean
'
'ValidateAckNoFm11A2filed_80DD = True
'
'    Dim rangecells As Range
'    Set rangecells = Sheet14.Range("AcknowledgeNum11A2_80DD").Cells
'
'
'    Dim i As Long
'    ReDim AckNo_ofForm11A2filed_80DD(end80DD)
'
'
'    For i = 1 To end80DD
'        AckNo_ofForm11A2filed_80DD(i) = rangecells.item(i).value
'        'Type_Disability_80DD(i) = rangecells1.item(i).value
'     'Ayush_20
'
'
'      If Not checkfieldspecialcharacter(AckNo_ofForm11A2filed_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "* Ack no. of Form as per Rule 11A(2)(ii)  at Sr. No  " & i & " cannot contain special characters in Schedule 80DD." & Chr(13)
'             ValidateAckNoFm11A2filed_80DD = False
'             Exit Function
'        End If
'
'
'
'        UpdateProgressBar
'
'    If Len(AckNo_ofForm11A2filed_80DD(i)) > 50 Then
'        ValidateAckNoFm11A2filed_80DD = False
'        MsgBox_80DD = MsgBox_80DD + "* Ack no. of Form as per Rule 11A(2)(ii)  cannot exceed 50 digits at Sr. No " & i & " in Schedule 80DD" & Chr(13)
'        Exit Function
'    End If
'
'Next
'End Function
'AYUSH_20

'Commented by Ayush_06/05/2025 for 0.7 changes
'Function ValidateAckNoFm11A2filed_80U() As Boolean
'
'ValidateAckNoFm11A2filed_80U = True
'
'    Dim rangecells As Range
'    Set rangecells = Sheet14.Range("AcknowledgeNum11A2_80U").Cells
'
'
'    Dim i As Long
'    ReDim AckNo_ofForm11A2filed_80U(end80U)
'
'
'    For i = 1 To end80U
'        AckNo_ofForm11A2filed_80U(i) = rangecells.item(i).value
'        'Type_Disability_80DD(i) = rangecells1.item(i).value
'     'Ayush_20
'
'
'      If Not checkfieldspecialcharacter(AckNo_ofForm11A2filed_80U(i)) Then
'             MsgBox_80U = MsgBox_80U + "* Ack no. of Form as per Rule 11A(2)(ii)  at Sr. No  " & i & " cannot contain special characters in Schedule 80U." & Chr(13)
'             ValidateAckNoFm11A2filed_80U = False
'             Exit Function
'        End If
'
'
'
'        UpdateProgressBar
'
'    If Len(AckNo_ofForm11A2filed_80U(i)) > 50 Then
'        ValidateAckNoFm11A2filed_80U = False
'        MsgBox_80U = MsgBox_80U + "* Ack no. of Form as per Rule 11A(2)(ii)  cannot exceed 50 digits at Sr. No " & i & " in Schedule 80U" & Chr(13)
'        Exit Function
'    End If
'
'Next
'End Function


'UID_Num80DD(8)
Function ValidateUDIDNum_80DD() As Boolean
ValidateUDIDNum_80DD = True

    Dim rangecells As Range
    Set rangecells = Sheet14.Range("UDIDNumber_80DD").Cells
    Dim i As Long
    ReDim UDID_Num_80DD(end80DD)
    For i = 1 To end80DD
        UDID_Num_80DD(i) = rangecells.item(i).value
        'Ayush_Legacy3
      If Not checkfieldspecialcharacter80DD_80U(UDID_Num_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* UDID Number cannot contain special characters in Schedule 80DD" & Chr(13)
             ValidateUDIDNum_80DD = False
             Exit Function
        End If
        UpdateProgressBar
 
    If Len(UDID_Num_80DD(i)) > 18 Then 'Chandru
        ValidateUDIDNum_80DD = False
        MsgBox_80DD = MsgBox_80DD + "* UDID Number cannot exceed 18 digits in Schedule 80DD" & Chr(13)
        Exit Function
    End If
Next
End Function
Function Check80DDPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 1 is for individuals .So,4th character of PAN should be "P"

    Check80DDPAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            Check80DDPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            Check80DDPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            Check80DDPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            Check80DDPAN = False
            Exit Function
        End If
        'Ayush_25-26
          
'        If Not UCase(Mid(PAN, 4, 1)) = UCase("P") Then
'            Check80DDPAN = False
'            Exit Function
'        End If
'
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            Check80DDPAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            Check80DDPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            Check80DDPAN = False
            Exit Function
        End If
    End If
End Function



'80U Schedule Starts---------------------------


Sub Validate80U()
    If Not Validate80U_1 Then
        Sheet14.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80U)
        CloseMsg
    End If
    
    
    If Sheet14.Range("AcknowledgeNum_80DD").value <> "" Or Sheet14.Range("AckNumber10IA_80U").value <> "" Then
    If Sheet14.Range("AcknowledgeNum_80DD").value = Sheet14.Range("AckNumber10IA_80U").value Then
        MsgBox_80U = MsgBox_80U & "* Acknowledgement number of Form 10IA filed for self and Dependent can't be same. Please provide proper acknowledgement number" & Chr(13)
       fmsgbox MsgBox_80U
       CloseMsg
    End If
    End If
    
End Sub

Function Validate80U_1() As Boolean
    Validate80U_1 = True
    MsgBox_80U = "Schedule 80U : " & Chr(10)
         
setTableInfo80U
'setTableInfo80U1
'Commented by Ayush_06/05/2025 for 0.7 changes
'setTableInfo80U2
setTableInfo80U3
setTableInfo80U4
setTableInfo80U5


'    end80U = WorksheetFunction.Max(0, end80U, end80U1, end80U2, end80U3, end80U4, end80U5)
    end80U = WorksheetFunction.Max(0, end80U, end80U1, end80U3, end80U4, end80U5)


 If end80U > 0 Then

       If Not ValidateNature_disability_80U Then Validate80U_1 = False
       If Not ValidateType_disability_80U Then Validate80U_1 = False
       If Not ValidateAmount_of_deduction_80U Then Validate80U_1 = False
       'Commented by Ayush_06/05/2025 for 0.7 changes
      ' If Not ValidateDate_of_filingofForm10IA_80U Then Validate80U_1 = False
       If Not ValidateAckNumber10IA_80U Then Validate80U_1 = False
       'AYUSH_20
       'Commented by Ayush_06/05/2025 for 0.7 changes
      ' If Not ValidateAckNoFm11A2filed_80U Then Validate80U_1 = False
       'ValidateAckNoFm11A2filed_80U
       
       If Not ValidateUDIDNum_80U Then Validate80U_1 = False
End If
        
End Function


Sub setTableInfo80U()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("NatureDisability_80U").Cells
    mIntCells = Sheet14.Range("NatureDisability_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) And Not UCase(rangecells.item(mIntCtr).value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80U = ccount
    'rngname_80U = "NatureDisability_80U;AmountDeduction_80U;DateFiling10IA_80U;AckNumber10IA_80U;UDIDNum_80U;"
End Sub
Sub setTableInfo80U1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AmountDeduction_80U").Cells
    mIntCells = Sheet14.Range("AmountDeduction_80U").count
    For mIntCtr = 1 To mIntCells
    MsgBox rangecells.item(mIntCtr).value
        If Not rangecells.item(mIntCtr).value = "0" Or Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U1 = ccount
End Sub
'Commented by Ayush_06/05/2025 for 0.7 changes
'Sub setTableInfo80U2()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet14.Range("DateFiling10IA_80U").Cells
'    mIntCells = Sheet14.Range("DateFiling10IA_80U").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80U2 = ccount
'End Sub

Sub setTableInfo80U3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AckNumber10IA_80U").Cells
    mIntCells = Sheet14.Range("AckNumber10IA_80U").count
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
    Set rangecells = Sheet14.Range("UDIDNum_80U").Cells
    mIntCells = Sheet14.Range("UDIDNum_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80U4 = ccount
End Sub
'Disability_80U
Sub setTableInfo80U5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Disability_80U").Cells
    mIntCells = Sheet14.Range("Disability_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblankDD(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80U5 = ccount
End Sub

Function ValidateNature_disability_80U() As Boolean
   ValidateNature_disability_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("NatureDisability_80U").Cells
    Dim i As Long
    ReDim Nature_of_disability_80U(end80U)
    
    For i = 1 To end80U

        Nature_of_disability_80U(i) = rangecells.item(i).value
        
        If isdropdownblank(Nature_of_disability_80U(i)) Or UCase(Nature_of_disability_80U(i)) = UCase("Select") Then
        
           MsgBox_80U = MsgBox_80U + "* Nature of  Disability in schedule 80U is mandatory" & Chr(13)
            ValidateNature_disability_80U = False
            Exit Function
            UpdateProgressBar
       End If
    Next
End Function

Function ValidateType_disability_80U() As Boolean
   ValidateType_disability_80U = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Disability_80U").Cells
    Dim i As Long
    ReDim Type_of_disability_80U(end80U)
    
    For i = 1 To end80U

        Type_of_disability_80U(i) = rangecells.item(i).value
        'Ayush_06/05/2025
        If isdropdownblankDD(Type_of_disability_80U(i)) Or UCase(Type_of_disability_80U(i)) = UCase("(Select)") Then
              'SIT-91922 Fixed by Bindu
         
            MsgBox_80U = MsgBox_80U + "*Selection of ""Type of disability"" in schedule 80U is mandatory." & Chr(13)
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
    Set rangecells = Sheet14.Range("AmountDeduction_80U").Cells
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
           MsgBox_80U = MsgBox_80U + "* Amount of deduction in  Schedule 80U should be Numeric value" & Chr(13)
            ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
        
        If Amount_of_deduction_80U(i) > 99999999999999# Then
           MsgBox_80U = MsgBox_80U + "* Amount of deduction  in Schedule 80U cannot exceed 14 digits" & Chr(13)
           ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
Next

End Function

 'Commented by Ayush_06/05/2025 for 0.7 changes
'Function ValidateDate_of_filingofForm10IA_80U() As Boolean
'
'ValidateDate_of_filingofForm10IA_80U = True
''    setTableInfo80GGA
'    Dim rangecells As Range
'    Set rangecells = Sheet14.Range("DateFiling10IA_80U").Cells
'
'    Dim rangecells1 As Range
'    Set rangecells1 = Sheet14.Range("Disability_80U").Cells
'    Dim i As Long
'    ReDim Date_filingFm10IA_80U(end80U)
'    ReDim Type_Disability_80U(end80U)
'
'
'
'
'
'
'    For i = 1 To end80U
'        Date_filingFm10IA_80U(i) = rangecells.item(i).value
'        Type_Disability_80U(i) = rangecells1.item(i).value
'
'        If Type_Disability_80U(i) = "(x) multiple disability" Or Type_Disability_80U(i) = "(ix) cerebral palsy" Or Type_Disability_80U(i) = "(viii) autism" Then
'        If Not chkCompulsory(Date_filingFm10IA_80U(i)) Then
'             MsgBox_80U = MsgBox_80U + "*  Please provide date of filing of Form 10IA in schedule 80U " & Chr(13)
'
'            ValidateDate_of_filingofForm10IA_80U = False
'            Exit Function
'        End If
'
'
'        End If
'
'    If Not CheckDateddmmyyyy(Date_filingFm10IA_80U(i)) Then
'        ValidateDate_of_filingofForm10IA_80U = False
'        'Ayush_10/02
''        MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA must be a valid dd/mm/yyyy format in Schedule 80U" & Chr(13)
'MsgBox_80U = MsgBox_80U + "* ""Date of filing of Form 10IA - Please enter valid date in valid format.""" & Chr(13)
'        Exit Function
'    End If
'        UpdateProgressBar
'
'    If Len(Date_filingFm10IA_80U(i)) > 10 Then
'        ValidateDate_of_filingofForm10IA_80U = False
'        MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA cannot exceed 10 digits in Schedule 80U" & Chr(13)
'        Exit Function
'    End If
'
'
'     If Not ChkMinInclusiveDate(Trim(Dformat(Date_filingFm10IA_80U(i), "yyyy-mm-dd")), "2025-04-01") Then
'           MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA cannot before 01/04/2025 in schedule 80U" & Chr(13)
'            ValidateDate_of_filingofForm10IA_80U = False
'
'
'          Exit Function
'        End If
'
'Next
'End Function

 

Function ValidateAckNumber10IA_80U() As Boolean
    
ValidateAckNumber10IA_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("AckNumber10IA_80U").Cells
    'AYUSH_20
    Dim rangecells1 As Range
    Set rangecells1 = Sheet14.Range("Disability_80U").Cells
    
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80U(end80U)
    'AYUSH_20
    ReDim Type_Disability_80U(end80U)
    
    For i = 1 To end80U
        AckNo_ofForm10IAfiled_80U(i) = rangecells.item(i).value
        'AYUSH_20
         Type_Disability_80U(i) = rangecells1.item(i).value
     'Ayush_06/05/2025
'     If Type_Disability_80U(i) = "(x) multiple disability" Or Type_Disability_80U(i) = "(ix) cerebral palsy" Or Type_Disability_80U(i) = "(viii) autism" Then
     If Type_Disability_80U(i) = "(i) autism, cerebral palsy, or multiple disabilities" Then
     If Not chkCompulsory(AckNo_ofForm10IAfiled_80U(i)) Then
            
'sudha_2025-26

            'MsgBox_80U = MsgBox_80U + "* Please enter a valid acknowledgement number of Form 10IA" & Chr(13)
             MsgBox_80U = MsgBox_80U + "* Please provide Acknowledgement of Form 10IA in schedule 80U" & Chr(13)
             
 '===============================
             ValidateAckNumber10IA_80U = False
             Exit Function
     End If
     
     End If
     
      If Not checkfieldspecialcharacter(AckNo_ofForm10IAfiled_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* Ack. No. of Form 10IA filed at Sr. No  " & i & " cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateAckNumber10IA_80U = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80U(i)) > 15 Then
        ValidateAckNumber10IA_80U = False
        MsgBox_80U = MsgBox_80U + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
Next
End Function
 
Function ValidateUDIDNum_80U() As Boolean
    
ValidateUDIDNum_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("UDIDNum_80U").Cells
    Dim i As Long
    ReDim UDID_Num_80U(end80U)
    
    For i = 1 To end80U
        UDID_Num_80U(i) = rangecells.item(i).value
     'Ayush_Legacy3
      If Not checkfieldspecialcharacter80DD_80U(UDID_Num_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* UDID Number at Sr. No " & i & " cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateUDIDNum_80U = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(UDID_Num_80U(i)) > 18 Then  'Chandru
        ValidateUDIDNum_80U = False
        MsgBox_80U = MsgBox_80U + "*UDID Number cannot exceed 18 digits in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
Next
End Function





