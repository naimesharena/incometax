Attribute VB_Name = "mdBA"
'Import_PARTA_OI_XML
Option Explicit

Public Msgbox_BA As String
Public ColCount1_BA, ColCount2_BA, ColCount3_BA, ColCount4_BA, ColCount5_BA, ColCount1_BA_IFSC, ColCount6_BA, ColCount55_BA, tempxml As Variant
Public rngname_BankTable As Variant
'Malli
Public BankAccnType_BA As Variant
Public rowcount_PartB As Variant
Public countIFSC_PartB As Variant

Public ColCount1_BA5, ColCount55_BA6_Other, ColCount1_typeaccnt_BA5 As Variant 'newly added by sai on 26/03/2025 for AY 2025-26

Public BankName_BA, BankNameCode, BankAddr_BA, BankJointHolder_BA, BankAccntnum_BA, CashDeposited_BA, BankAccntBalanc_BA, BankIFSC_BA, BankAccountStatus_BA As Variant


Sub ValidateSheetBA()
    Application.ScreenUpdating = False
    Dim vbMessgaeCaption As String
    vbMessgaeCaption = "ITR 3: AY: 2025-26" 'Year Changed by sai
    ValidateBA
    fmsgboxStatus ("Part B TI TTI is OK ") ', vbOKOnly, vbMessgaeCaption
    Application.ScreenUpdating = True
End Sub

Sub ValidateBA()
    Application.ScreenUpdating = False
    Dim vbMessgaeCaption As String
    vbMessgaeCaption = "ITR 3: AY: 2025-26" 'Year Changed by sai
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Part B - TI TTI")

           
    If Not ValidateSchBA Then
        sourceSheet.Activate
        fmsgboxsmall (Msgbox_BA) ', vbOKOnly, "Error(s!)"
        Application.ScreenUpdating = True
        CloseMsg
    End If

End Sub
Function ValidateSchBA() As Boolean
    ValidateSchBA = True
    Msgbox_BA = ""
    
     setTableInfo_BA_IFSC
     setTableInfo_BA
     setTableInfo_BA3
     
     setTableInfo_BA44 'uncommented by sai on 26/03/2025
     setTableInfo_BA6  'uncommented by sai on 26/03/2025
     
     setTableInfo_BA66  'newly added by sai on 26/03/2025 for AY 2025-26
     
     
     'Malli
     'ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA_IFSC, ColCount1_BA, ColCount4_BA, ColCount6_BA, ColCount55_BA)
'      ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA_IFSC, ColCount1_BA, ColCount4_BA) ' ColCount6_BA, ColCount55_BA)

    ' Modified by sai on 17/03/2025
      ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA_IFSC, ColCount1_BA, ColCount4_BA, ColCount6_BA, ColCount55_BA, ColCount1_typeaccnt_BA5)
    
            If Not ValidateCheckBox_BA Then ValidateSchBA = False
            If Not ValidateIFSC Then ValidateSchBA = False
            
            
           

End Function
Function ValidateCheckBox_BA() As Boolean
ValidateCheckBox_BA = True
Dim rangecells As Range
Dim rangecells1 As Range
Set rangecells = Range("Sheet9.OthMICRCode").Cells
'Set rangecells1 = Range("tempxml").Cells
'Dim i As Long
'setTableInfo_BA44

'ReDim BankIFSC_BA(ColCount55_BA)

'For i = 1 To ColCount55_BA
   ' BankIFSC_BA(i) = rangecells.item(i).value
   
'If isdropdownblank(BankIFSC_BA(i)) Then
'        Msgbox_BA = Msgbox_BA + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet Part B - TI TTI is mandatory." & Chr(13)
'        ValidateCheckBox_BA = False
'        Exit Function
'
'    End If
'
'    If Len(BankIFSC_BA(i)) > 11 Then
'            Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Part B TI TTI cannot exceed 11 characters" & Chr(13)
'            ValidateCheckBox_BA = False
'            Exit Function
'        End If
'
'    If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_BA(i))) Then
'                Msgbox_BA = Msgbox_BA + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
'                ValidateCheckBox_BA = False
'                Exit Function
'            End If
'

   ' Next
 
End Function
Function ValidateIFSC() As Boolean
    ValidateIFSC = True
'    setTableInfo_BA_IFSC
    Dim rangecells, rangecells1, rangecells2, rangecells3 As Range
    Set rangecells = Range("Sheet9.OthMICRCode").Cells
    Set rangecells1 = Range("Sheet9.OthBankName").Cells
    Set rangecells2 = Range("Sheet9.OthBankAccountNumber").Cells
    'Set rangecells3 = Range("tempxml").Cells
    
    Set rangecells3 = Range("Sheet9.TypeOfAccount").Cells
    
    Dim i As Long
    ReDim BankIFSC_BA(ColCount1_BA_IFSC)
    ReDim BankName_BA(ColCount1_BA_IFSC)
    ReDim BankAccntnum_BA(ColCount1_BA_IFSC)
     ReDim BankAccnType_BA(ColCount1_BA_IFSC)
    'ReDim tempxml(ColCount1_BA_IFSC)
    
'     If IBANSELECT = "Y" Then
'    If Not ColCount6_BA > 0 Then
'        Msgbox_BA = Msgbox_BA + "* Please select atleast one account in which you prefer to get your refund  in Sheet Part B - TI TTI is mandatory" & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'    End If
'    End If
    
    'Malli added
    
    If ColCount1_BA_IFSC > 0 Then
    
    
     If IBANSELECT = "Y" Then
      If Not ColCount6_BA > 0 Then
        Msgbox_BA = Msgbox_BA + "* Please select atleast one account in which you prefer to get your refund  in Sheet Part B - TI TTI is mandatory" & Chr(13)
        ValidateIFSC = False
        Exit Function
        End If
    End If
    

    If Not ColCount55_BA6_Other > 0 Then
          Msgbox_BA = Msgbox_BA + "*""Please select atleast one saving/current/cash credit/overdraft/non resident bank account in which you prefer to get your refund""" & Chr(13)
          ValidateIFSC = False
          Exit Function
    End If


     
'***********************************************************************************


    For i = 1 To ColCount1_BA_IFSC
        BankIFSC_BA(i) = rangecells.item(i).value
        BankName_BA(i) = rangecells1.item(i).value
        BankAccntnum_BA(i) = rangecells2.item(i).value
        BankAccnType_BA(i) = rangecells3.item(i).value
        
        'tempxml(i) = rangecells3.item(i).value
        
'        If isdropdownblank(BankIFSC_BA(i)) Then
'        Msgbox_BA = Msgbox_BA + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet Part B - TI TTI is mandatory." & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'        End If

        'Malli===================
        If isdropdownblank(BankIFSC_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* ""IFSC code is mandatory in Pt 13 i.e. bank account details in Schedule Part B TTI in point 13.""" & Chr(13)
        ValidateIFSC = False
        Exit Function
        End If
        
        If Len(BankIFSC_BA(i)) > 11 Then
            Msgbox_BA = Msgbox_BA + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets." & Chr(13)
            ValidateIFSC = False
            Exit Function
        End If
        
        If Not ValidateIFSCList(UCase(BankIFSC_BA(i))) Then
                'Msgbox_BA = Msgbox_BA + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
'                Msgbox_BA = Msgbox_BA + "* ""Please enter a valid IFSC in Schedule Part B TTI in point 13"" & Chr(13)"
                Msgbox_BA = Msgbox_BA + "* ""Please ensure the IFSC entered in Schedule Part B TTI in point 13 is valid""" & Chr(13)
                ValidateIFSC = False
                Exit Function
        End If
        
        If Not CheckIFSC(UCase(BankIFSC_BA(i))) Then
        Msgbox_BA = Msgbox_BA + "* ""IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."" & Chr(13)"
           ValidateIFSC = False
        Exit Function
        End If
        
        '===========================
'        If Len(BankIFSC_BA(i)) > 11 Then
'            Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Part B TI TTI cannot exceed 11 characters" & Chr(13)
'            ValidateIFSC = False
'            Exit Function
'        End If
'
'
'            If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_BA(i))) Then
'                Msgbox_BA = Msgbox_BA + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
'                ValidateIFSC = False
'                Exit Function
'            End If
        
        If BankName_BA(i) = "" Then
            Msgbox_BA = Msgbox_BA + "* ""Please enter the Bank name at Sr.No " & i & " in Part B TI TTI""" & Chr(13)
            ValidateIFSC = False
            Exit Function
        End If
         
        
        If Len(BankName_BA(i)) > 125 Then
            Msgbox_BA = Msgbox_BA + "* Bank Name at Sr.No " & i & " in Part B TI TTI cannot exceed 125 characters" & Chr(13)
            ValidateIFSC = False
            Exit Function
        End If

    
            
            If Not chkCompulsory(BankAccntnum_BA(i)) Then
            'Msgbox_BA = Msgbox_BA + "* ""Please enter the Account Number in Schedule Part B TTI in point 18"" " & Chr(13)
            Msgbox_BA = Msgbox_BA + "*""Please enter the Account Number at Sr.No " & i & " in Part B TI TTI""" & Chr(13)
            ValidateIFSC = False
            Exit Function
           End If
           
           If Not ValidateBankAccountNumber_BA(BankAccntnum_BA(i), i) Then
            ValidateIFSC = False
            Exit Function
        End If
        
        
        If isdropdownblank(BankAccnType_BA(i)) Or UCase(BankAccnType_BA(i)) = UCase("Select") Then
       
        Msgbox_BA = Msgbox_BA + "* ""please select type of account at Sr.No " & i & " in Schedule Part B TTI""" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If

        
       
    Next
Else

'       Msgbox_BA = Msgbox_BA + "* ""Please Enter Bank Account details in Schedule Part B TTI""" & Chr(13)
'       ValidateIFSC = False
'       Exit Function
       
    'Newly added by sai
       If Sheet7.Range("Sheet9.IBANSELECT").value = "Yes" Then
       Msgbox_BA = Msgbox_BA + "* ""Please Enter Bank Account details in Schedule Part B TTI""" & Chr(13)
       ValidateIFSC = False
       Exit Function
       End If
       
       
       

End If
       
    'Next
End Function

Function BankCode() As Boolean
    BankCode = True
    Dim rangecells As Range
    Set rangecells = Range("BankNameWithCode").Cells
    Dim i As Long
    ReDim BankNameCode(ColCount1_BA_IFSC)
    For i = 1 To ColCount1_BA_IFSC
        BankNameCode(i) = rangecells.item(i).value
        If IsNumeric(BankName_BA(i)) Then
            BankCode = False
            Exit Function
        End If
    Next
End Function
Function ValidateNameofHolders_BA() As Boolean
    ValidateNameofHolders_BA = True
    Dim rangecells As Range
    Dim i, UTNintCount As Long
    Set rangecells = Range("SchBA.JointHoderName").Cells
    UTNintCount = Range("SchBA.JointHoderName").count
    
    ReDim BankJointHolder_BA(UTNintCount)
    For i = 1 To UTNintCount
        BankJointHolder_BA(i) = rangecells.item(i).value
        If Not rangecells.item(i).value = "" Then
        End If
        
        If BankJointHolder_BA(i) <> "" Then
            If Len(BankJointHolder_BA(i)) > 125 Then
                Msgbox_BA = Msgbox_BA + "* Name of the Joint Account Holder at Sr.No " & i + 1 & " in Part B TI TTI cannot exceed 125 characters" & Chr(13)
                ValidateNameofHolders_BA = False
                Exit Function
            End If
        End If
    Next
End Function

Function ValidateAccntStatus_BA() As Boolean
    ValidateAccntStatus_BA = True
    Dim rangecells As Range
    Dim rangecellbal As Range
    Set rangecells = Range("Sheet9.OthBankAccountType1").Cells
    'Set rangecellbal = Range("SchBA.AccntBalance").Cells
    Dim i As Long
    ReDim BankAccountStatus_BA(ColCount1_BA_IFSC)
    'ReDim BankAccntBalanc_BA(ColCount1_BA_IFSC)
    For i = 1 To ColCount1_BA_IFSC
        BankAccountStatus_BA(i) = UCase(Mid(rangecells.item(i).value, 1, 3))
        If BankAccountStatus_BA(i) = "CAS" Then BankAccountStatus_BA(i) = "CCT"
        If isdropdownblank(BankAccountStatus_BA(i)) Then
            Msgbox_BA = Msgbox_BA + "* Account Status at Sr.No " & i + 1 & " in Part B TI TTI is mandatory" & Chr(13)
            ValidateAccntStatus_BA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAccntBalance_BA() As Boolean
    ValidateAccntBalance_BA = True
    Dim rangecells As Range
    Set rangecells = Range("SchBA.AccntBalance").Cells
    Dim i As Long
    ReDim BankAccntBalanc_BA(ColCount1_BA_IFSC)
    For i = 1 To ColCount1_BA_IFSC
        BankAccntBalanc_BA(i) = rangecells.item(i).value
          
        If Not chkCompulsory(BankAccntBalanc_BA(i)) Then
            Msgbox_BA = Msgbox_BA + "* Account Balance at Sr.No " & i + 1 & " in Part B TI TTI is mandatory" & Chr(13)
            ValidateAccntBalance_BA = False
            Exit Function
        End If
        
        If Not IsNumeric(BankAccntBalanc_BA(i)) Then
            Msgbox_BA = Msgbox_BA + "* Account Balance at Sr.No " & i + 1 & " Should be a numeric value" & Chr(13)
            ValidateAccntBalance_BA = False
            Exit Function
        End If
        
        If ((BankAccntBalanc_BA(i) > 99999999999999#) Or (BankAccntBalanc_BA(i) < -99999999999999#)) Then
            Msgbox_BA = Msgbox_BA + "* Account Balance at Sr.No " & i + 1 & " cannot exceed 14 digits" & Chr(13)
            ValidateAccntBalance_BA = False
            Exit Function
        End If
    
    Next
End Function

Sub setTableInfo_BA_IFSC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet9.OthMICRCode").count
    Set rangecells = Range("Sheet9.OthMICRCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not Trim(rangecells.item(mIntCtr).value) = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA_IFSC = ccount
    rngname_BankTable = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.OthCashDeposited"
End Sub

Sub setTableInfo_BA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet9.OthBankName").count
    Set rangecells = Range("Sheet9.OthBankName").Cells
    For mIntCtr = 1 To mIntCells
            If Not ((Trim(rangecells.item(mIntCtr).value) = "") Or (Trim(rangecells.item(mIntCtr).value) = "(Select)")) Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA = ccount
End Sub
Sub setTableInfo_BA44()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim j As Long
     Dim b As Long
    
    ccount = 0
  
    mIntCells = Sheet7.Range("tempxml").count
    Set rangecells = Sheet7.Range("tempxml").Cells
    For mIntCtr = 1 To mIntCells
      'j = 1
      b = 0
         For j = mIntCtr To mIntCells
         
        If Not (Trim(rangecells.item(mIntCtr).value) = False Or Trim(rangecells.item(mIntCtr).value) = "") Or Not (Trim(rangecells.item(j).value) = False Or Trim(rangecells.item(j).value) = "") Then
            b = 1
            
          End If
        Next
        If b = 1 Then
        ccount = ccount + 1
        End If
    Next
    ColCount55_BA = ccount
    'rngname_AuditInfo2 = "Audit.Act;Audit.Act_Description;Audit.Sections;Audit.Date;"
End Sub
Sub setTableInfo_BA3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet9.OthBankAccountNumber").count
    Set rangecells = Range("Sheet9.OthBankAccountNumber").Cells
    For mIntCtr = 1 To mIntCells
            If Not Trim(rangecells.item(mIntCtr).value) = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount4_BA = ccount
End Sub
Sub setTableInfo_BA6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Range("tempxml").count
    Set rangecells = Range("tempxml").Cells
    For mIntCtr = 1 To mIntCells
             If rangecells.item(mIntCtr).value = True Then
               ccount = ccount + 1
           End If
    Next
    ColCount6_BA = ccount
End Sub
Sub setTableInfo_BA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.JointHoderName").count
    Set rangecells = Range("SchBA.JointHoderName").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
        
    If ccount <= ColCount1_BA_IFSC Then
    ColCount3_BA = ColCount1_BA_IFSC
    Else:
    ColCount3_BA = ccount
    End If
End Sub
Sub setTableInfo_BA4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.AccntBalance").count
    Set rangecells = Range("SchBA.AccntBalance").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount5_BA = ccount
End Sub


Sub setTableInfo_BA5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet9.OthBankAccountType1").count
    Set rangecells = Range("Sheet9.OthBankAccountType1").Cells
    For mIntCtr = 1 To mIntCells
            If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
               ccount = ccount + 1
           End If
    Next
    ColCount6_BA = ccount
End Sub
Function ValidateBankAccountNumber_BA(BankAccountNumber As Variant, CC As Long) As Boolean
    ValidateBankAccountNumber_BA = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB() As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter_Bank(BankAccountNumber) Then
            Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & "is invalid in sheet Part B TI TTI, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI " & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If
    
    End If
    
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is mandatory in Part B TI TTI" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is mandatory in Part B TI TTI" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "/-") > 0) Or (InStr(BankAccountNumber, "-/") > 0)) Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI  " & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
    

    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI " & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    ReDim myB(Len(BankAccountNumber) - 1)
    For i = 1 To Len(BankAccountNumber)
        myB(i - 1) = Mid(BankAccountNumber, i, 1)
    Next

    For i = LBound(myB) To UBound(myB)
        If IsNumeric(myB(i)) Then
            countnum = countnum + 1
        End If

        If i > LBound(myB) And i < UBound(myB) Then
            If myB(i) = 0 Then
                If myB(i - 1) = 0 Then
                    zeroCount = zeroCount + 1
                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
                Else
                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
                End If
            End If
        End If

    Next

    If BeforeZero <> "" And AfterZero <> "" Then
        If zeroCount > 1 Then
            Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI." & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "Bank Account Number at Sr.No " & CC & " in Part B TI TTI is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
End Function
Function ValidateBankAccountNumber_FA(BankAccountNumber As Variant, CC As Long) As Boolean
    ValidateBankAccountNumber_FA = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB() As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter_Bank(BankAccountNumber) Then
            msgbox9 ("Bank Account Number at Sr.No " & CC & " is invalid in sheet Part B TI TTI, Only "" / "" and "" - "" special characters are allowed.")
            ValidateBankAccountNumber_FA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            msgbox9 ("Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI ")
            ValidateBankAccountNumber_FA = False
            Exit Function
        End If
    
    End If
    
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 40) Then
        msgbox9 ("Bank Account Number at Sr.No " & CC & " is mandatory in Part B TI TTI")
        ValidateBankAccountNumber_FA = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        msgbox9 ("Bank Account Number at Sr.No " & CC & " is mandatory in Part B TI TTI")
        ValidateBankAccountNumber_FA = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        msgbox9 ("Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI  ")
        ValidateBankAccountNumber_FA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        msgbox9 ("Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI")
        ValidateBankAccountNumber_FA = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        msgbox9 ("Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI ")
        ValidateBankAccountNumber_FA = False
        Exit Function
    End If

'    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
'        msgbox9 ("Bank Account Number at Sr.No " & cc & " is invalid in Part B TI TTI")
'        ValidateBankAccountNumber_FA = False
'        Exit Function
'    End If

    ReDim myB(Len(BankAccountNumber) - 1)
    For i = 1 To Len(BankAccountNumber)
        myB(i - 1) = Mid(BankAccountNumber, i, 1)
    Next

    For i = LBound(myB) To UBound(myB)
        If IsNumeric(myB(i)) Then
            countnum = countnum + 1
        End If

        If i > LBound(myB) And i < UBound(myB) Then
            If myB(i) = 0 Then
                If myB(i - 1) = 0 Then
                    zeroCount = zeroCount + 1
                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
                Else
                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
                End If
            End If
        End If

    Next

    If BeforeZero <> "" And AfterZero <> "" Then
        If zeroCount > 1 Then
            msgbox9 ("Bank Account Number at Sr.No " & CC & " is invalid in Part B TI TTI.")
            ValidateBankAccountNumber_FA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        msgbox9 ("Bank Account Number at Sr.No " & CC & " in Part B TI TTI is invalid. Account Number should have a minimum of 1 numeric digit (1-9)")
        ValidateBankAccountNumber_FA = False
        Exit Function
    End If
End Function

Function checkforNonMandatoryCol() As Boolean
    checkforNonMandatoryCol = True
    Dim temp As Double
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.JointHoderName").count
    Set rangecells = Range("SchBA.JointHoderName").Cells
    temp = ColCount1_BA_IFSC
    
    For mIntCtr = (temp + 1) To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            Msgbox_BA = Msgbox_BA + "* Please Enter all the Mandatory fiels at Sr.No " & mIntCtr & " in PART - TI - TTI" & Chr(13)
            checkforNonMandatoryCol = False
            Exit Function
        End If
    Next
End Function

Function ValidateMismatchSchBA()
    On Error Resume Next
    ValidateMismatchSchBA = True
    Msgbox_BA = ""
    If ColCount1_BA_IFSC + 1 <> (Sheet7.Range("Sheet9.NumofBankAccounts").value) Then
        ValidateMismatchSchBA = False
        Msgbox_BA = Msgbox_BA + "* Number of bank accounts held by you at any time (including opened/closed) during the previous year should be equal to the number of rows filled in the table."
        Sheet7.Range("Sheet9.NumofBankAccounts").Activate
    End If
End Function


Sub AddRows_BA()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    
    Set sourceSheet = ThisWorkbook.Sheets("BA")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.JointHoderName||SchBA.AcntNo||SchBA.AcntStatus||SchBA.AccntBalance"
    ActiveCellRange = EfilingCommon.searchLastRow("SchBA.IFSC")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub



Function CheckAccountStatus(Tval As Variant, Tadd As Variant)
On Error Resume Next
    Application.EnableEvents = False
    Dim rangecellsBA, rangecells1BA As Range
    Dim mIntcellsBA, q As Long
    Dim PrevVal As Variant
   
    sPassword = EfilingCommon.getmsgstate
    Sheet7.Unprotect Password:=sPassword

    Set rangecellsBA = Range("SchBA.AcntStatus").Cells
    Set rangecells1BA = Range("SchBA.AccntBalance").Cells

    mIntcellsBA = Range("SchBA.AcntStatus").count

    If Tval = "Existing" Then
        Sheet7.Range(Tadd).Select
        Sheet7.Range(Tadd).Offset(0, 1) = ""
        Sheet7.Range(Tadd).Offset(0, 1).Locked = False
        Sheet7.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        GoTo endfd
        Exit Function
    End If

    For q = 1 To mIntcellsBA
        If rangecellsBA.item(q).value = "(Select)" Then
            rangecells1BA.item(q).value = ""
        End If
                
        If rangecellsBA.item(q).value = "Closed" Then
            rangecells1BA.item(q).value = 0
            rangecells1BA.item(q).Locked = True
            rangecells1BA.item(q).Interior.Color = (&HD8D8D8)
        Else
            rangecells1BA.item(q).Locked = False
            rangecells1BA.item(q).Interior.Color = (&HCCFFCC)
        End If
    Next

endfd:
    Sheet7.Protect Password:=sPassword
    Application.EnableEvents = True
End Function
'Malli
'Malli
Function ValidateIFSCList(IFSC As String) As Boolean
    ValidateIFSCList = False
    Dim IFSCCodeSet1 As Range
    Dim IFSCCodeSet2 As Range
    Dim IFSCCodeSet3 As Range
    Dim IFSCCodeSet4 As Range
    Dim IFSCCodeSet5 As Range
    
    Dim IFSC1 As Range
    Dim IFSC2 As Range
    Dim IFSC3 As Range
    Dim IFSC4 As Range
    Dim IFSC5 As Range

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Set IFSCCodeSet1 = Range("IFSCCodeSet1")
    Set IFSCCodeSet2 = Range("IFSCCodeSet2")
    Set IFSCCodeSet3 = Range("IFSCCodeSet3")
    Set IFSCCodeSet4 = Range("IFSCCodeSet4")
    Set IFSCCodeSet5 = Range("IFSCCodeSet5")
    
    For Each IFSC1 In IFSCCodeSet1
        If IFSC1.value = IFSC Then
        ValidateIFSCList = True
        Exit For
        End If
    Next

    If Not ValidateIFSCList Then
        For Each IFSC2 In IFSCCodeSet2
            If IFSC2.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    If Not ValidateIFSCList Then
        For Each IFSC3 In IFSCCodeSet3
            If IFSC3.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    If Not ValidateIFSCList Then
        For Each IFSC4 In IFSCCodeSet4
            If IFSC4.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    
    If Not ValidateIFSCList Then
        For Each IFSC5 In IFSCCodeSet5
            If IFSC5.value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    ValidateIFSCList = True
   '----------------Lock Password-------------------START---
   Sheet8.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
 
End Function
Function CheckIFSC(IFSC As String) As Boolean
On Error Resume Next
'IFSC Code should be exactly 11 characters,
'First 4 characters should be alphabets,
'5th character must be zero (0) and remaining 6 should be either numeric or alphabets
    CheckIFSC = True
    If Len(IFSC) > 0 Then
        If Not ChkAlphabet(Mid(IFSC, 1, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 2, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 3, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 4, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not Mid(IFSC, 5, 1) = 0 Then
            CheckIFSC = False
            Exit Function
        End If
        
       ' If IsNumeric(Mid(IFSC, 6, 6)) Then
       '     CheckIFSC = True
       '     Exit Function
       ' Else
           
            If Not IsNumeric(Mid(IFSC, 6, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 6, 1)) Then
            CheckIFSC = False
            Exit Function
            End If
            End If

            
            If Not IsNumeric(Mid(IFSC, 7, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 7, 1)) Then
            CheckIFSC = False
           Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 8, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 8, 1)) Then
            CheckIFSC = False
            Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 9, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 9, 1)) Then
            CheckIFSC = False
            Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 10, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 10, 1)) Then
            CheckIFSC = False
           Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 11, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 11, 1)) Then
            CheckIFSC = False
           Exit Function
            End If
            End If
If Not checkfieldSuperSpecialcharacterDot(IFSC) Then
CheckIFSC = False
End If
          

'End If
End If
        
    
End Function
Function checkfieldSuperSpecialcharacterDot(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharacterDot = True
    Dim arr As Variant
    arr = Array(".")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharacterDot = False
            Exit Function
        End If
        Next
    Next
End Function
'Malli newly ADDED
'Malli
Sub BankImport_CG()

Dim count54Acc, count54IFSC, count54BAcc, count54BIFSC, count54FAcc, count54FIFSC
Dim TotalDiffRow

Dim Acc54FColumn, IFSC54FColumn, Acc54Column, Acc54BColumn, IFSC54Column, IFSC54BColumn
Dim cChell As Variant
Dim rowcount54F, i, rowcount54, rowcount54B
Dim rowcount54F_1, rowcount54_1, rowcount54B_1


Dim IFSCColumn, AccColumn, TypeAccCoulmn
Dim ii As Long


'Part B TTI
IFSCColumn = Sheet7.Range("Sheet9.OthMICRCode").Column
AccColumn = Sheet7.Range("Sheet9.OthBankAccountNumber").Column
TypeAccCoulmn = Sheet7.Range("Sheet9.TypeOfAccount").Column


countIFSC_PartB = Sheet7.Range("Sheet9.OthMICRCode").count

'CG sheet
count54Acc = Sheet13.Range("Count54Acc").value
count54IFSC = Sheet13.Range("Count54IFSC").value


Call PartB_rows

If count54Acc = count54IFSC Then
   If count54IFSC > countIFSC_PartB Then
       TotalDiffRow = count54IFSC - countIFSC_PartB
       
        For ii = 1 To TotalDiffRow
           Cmd_AddRowsBank_Click
        Next ii
   End If
End If


If count54Acc > 0 Then

    If count54Acc = count54IFSC Then
      '54 Column CG sheet
      rowcount54 = getRowNo(Sheet13.Range("Deduction.54.AcctNum").name)
      Acc54Column = Sheet13.Range("Deduction.54.AcctNum").Column
      IFSC54Column = Sheet13.Range("Deduction.54.IFSC").Column
      rowcount54_1 = Sheet13.Range("Deduction.54.IFSC").count
       
      For i = 1 To rowcount54_1
          If Sheet13.Cells(rowcount54, IFSC54Column).value <> "" And Sheet13.Cells(rowcount54, Acc54Column).value <> "" Then
            Sheet7.Cells(rowcount_PartB, IFSCColumn).value = Sheet13.Cells(rowcount54, IFSC54Column).value
            Sheet7.Cells(rowcount_PartB, AccColumn).value = Sheet13.Cells(rowcount54, Acc54Column).value
            Sheet7.Cells(rowcount_PartB, TypeAccCoulmn).value = "Capital Gain Account Scheme"
            rowcount_PartB = rowcount_PartB + 1
          End If
          rowcount54 = rowcount54 + 1
      Next
       
    End If

End If


'CG 54B sheet
count54BAcc = Sheet13.Range("Count54BAcc").value
count54BIFSC = Sheet13.Range("Count54BIFSC").value


countIFSC_PartB = Sheet7.Range("Sheet9.OthMICRCode").count


Call PartB_rows

If count54BAcc = count54BIFSC Then
   If count54BIFSC > countIFSC_PartB Then
       TotalDiffRow = count54BIFSC - countIFSC_PartB
       
        For ii = 1 To TotalDiffRow
           Cmd_AddRowsBank_Click
        Next ii
   End If
End If


If count54BAcc > 0 Then

    If count54BAcc = count54BIFSC Then
      '54 Column CG sheet
      rowcount54B = getRowNo(Sheet13.Range("Deduction.54B.AcctNum").name)
      Acc54BColumn = Sheet13.Range("Deduction.54B.AcctNum").Column
      IFSC54BColumn = Sheet13.Range("Deduction.54B.IFSC").Column
      rowcount54B_1 = Sheet13.Range("Deduction.54B.IFSC").count
       
      For i = 1 To rowcount54B_1
         If Sheet13.Cells(rowcount54B, IFSC54BColumn).value <> "" And Sheet13.Cells(rowcount54B, Acc54BColumn).value <> "" Then
            Sheet7.Cells(rowcount_PartB, IFSCColumn).value = Sheet13.Cells(rowcount54B, IFSC54BColumn).value
            Sheet7.Cells(rowcount_PartB, AccColumn).value = Sheet13.Cells(rowcount54B, Acc54BColumn).value
            Sheet7.Cells(rowcount_PartB, TypeAccCoulmn).value = "Capital Gain Account Scheme"
            rowcount_PartB = rowcount_PartB + 1
         End If
          rowcount54B = rowcount54B + 1
      Next
       
    End If

End If


'CG 54F
count54FAcc = Sheet13.Range("Count54FAcc").value
count54FIFSC = Sheet13.Range("Count54FIFSC").value


countIFSC_PartB = Sheet7.Range("Sheet9.OthMICRCode").count


Call PartB_rows

If count54FAcc = count54FIFSC Then
   If count54FIFSC > countIFSC_PartB Then
       TotalDiffRow = count54FIFSC - countIFSC_PartB
       
        For ii = 1 To TotalDiffRow
           Cmd_AddRowsBank_Click
        Next ii
   End If
End If


If count54FAcc > 0 Then

    If count54FAcc = count54FIFSC Then
      '54 Column CG sheet
      rowcount54F = getRowNo(Sheet13.Range("Deduction.54F.AcctNum").name)
      Acc54FColumn = Sheet13.Range("Deduction.54F.AcctNum").Column
      IFSC54FColumn = Sheet13.Range("Deduction.54F.IFSC").Column
      rowcount54F_1 = Sheet13.Range("Deduction.54F.IFSC").count
       
      For i = 1 To rowcount54F_1
        If Sheet13.Cells(rowcount54F, IFSC54FColumn).value <> "" And Sheet13.Cells(rowcount54F, Acc54FColumn).value <> "" Then
            Sheet7.Cells(rowcount_PartB, IFSCColumn).value = Sheet13.Cells(rowcount54F, IFSC54FColumn).value
            Sheet7.Cells(rowcount_PartB, AccColumn).value = Sheet13.Cells(rowcount54F, Acc54FColumn).value
            Sheet7.Cells(rowcount_PartB, TypeAccCoulmn).value = "Capital Gain Account Scheme"
            rowcount_PartB = rowcount_PartB + 1
         End If
         rowcount54F = rowcount54F + 1
      Next
       
    End If

End If

End Sub
Sub PartB_rows()

Dim mIntCells
Dim mIntCtr
Dim rangecells As Range

mIntCells = Range("Sheet9.OthMICRCode").count
Set rangecells = Range("Sheet9.OthMICRCode").Cells
rowcount_PartB = getRowNo(Sheet7.Range("Sheet9.OthMICRCode").name)

For mIntCtr = 1 To mIntCells
    If Not Trim(rangecells.item(mIntCtr).value) = "" Then
           rowcount_PartB = rowcount_PartB + 1
           countIFSC_PartB = countIFSC_PartB - 1
       End If
Next

End Sub

'Newly added by sai on 26/03/2025 for AY 2025-26

Sub setTableInfo_BA66()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount, ccount2 As Long

    Dim ccount1 As Long
    Dim rangecells1 As Range
    Dim mIntCells1 As Long
    Sheet7.Activate
    mIntCells1 = Range("Sheet9.TypeOfAccount").count
    Set rangecells1 = Range("Sheet9.TypeOfAccount").Cells

    ccount1 = 0
    ccount = 0
    ccount2 = 0
    mIntCells = Range("tempxml").count
    Set rangecells = Range("tempxml").Cells
    For mIntCtr = 1 To mIntCells


             If rangecells.item(mIntCtr).value = True Then
               ccount = ccount + 1

               If Trim(rangecells1.item(mIntCtr).value) <> Trim("Other") Then
                     ccount1 = ccount1 + 1
               End If
           End If



    Next
    ColCount1_BA5 = ccount
    ColCount55_BA6_Other = ccount1
    ColCount1_typeaccnt_BA5 = ccount2
End Sub




