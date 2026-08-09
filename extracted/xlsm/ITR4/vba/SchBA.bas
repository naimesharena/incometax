Attribute VB_Name = "SchBA"
Option Explicit

Public Msgbox_BA As String
Public ColCount1_BA, ColCount2_BA, ColCount3_BA, ColCount4_BA, ColCount5_BA, ColCount1_BA_IFSC, ColCount1_BA_Type_of_Account, ColCountAccntStat_BA, ColCount1_BA5, ColCount55_BA6_Other, ColCount55_BA As Variant
Dim sPassword As Variant
Public BankName_BA, BankAddr_BA, BankJointHolder_BA, BankAccntnum_BA, BankAccntBalanc_BA, AccountType_BA, CashDeposited_BA, BankIFSC_BA, Type_of_Account_BA, BankAccntStatus, BankNameCode_BA, BankAccntStatus_BA, tempXML As Variant


Public IBANSELECT As Variant
Public IBAN As Variant
Public IBANNameOfBank As Variant
Public IBANCOUNTRY As Variant
Public IBANACCNO As Variant

Public end_IBAN, end_IBANACCNO, end_IBANNameOfBank, end_IBANCOUNTRY As Variant
Public rngname_IBAN, rngname_IBANACCNO, rngname_IBANNameOfBank, rngname_IBANCOUNTRY As Variant

Sub ValidateSheetBA()
Dim vbMessgaeCaption As String
'vbMessgaeCaption = "ITR 4: AY: 2022-23" dpk1201
'vbMessgaeCaption = "ITR 4: AY: 2023-24"
vbMessgaeCaption = "ITR 4: AY: 2024-25" 'AY CHANGED
ValidateBA
fmsgboxStatus ("Sheet BA is OK ") ', vbOKOnly, vbMessgaeCaption
End Sub


Sub PrevBA_Click()
Sheet5.Activate
End Sub

Sub cmdNext_Click_BA()
Sheet6.Activate
End Sub
Sub AddRows_BA()
Dim vRows As Long
Dim sourceSheet As Worksheet

Set sourceSheet = ThisWorkbook.Sheets("Taxes Paid and Verification")
sourceSheet.Activate
'TAX DETAILS-E6 Old
'mdCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.CheckBox||tempXML" '||SchBA.DepositedCash"
'TAX DETAILS-C6 AY_2024_25 Change
'mdCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.Type_of_Account" '||SchBA.DepositedCash"

'Malli_20/03/2025
mdCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.Type_of_Account||SchBA.CheckBox||tempXML"
ActiveCellRange = mdCommon.searchLastRow("SchBA.IFSC")
vRows = mdCommon.insertRowUnderSectionWithFormulaOne
LinkCheckBoxes  'Malli_20/03/2025
CheckAccountStatus1
End Sub


Sub ValidateBA()
Dim vbMessgaeCaption As String
'vbMessgaeCaption = "ITR 4: AY: 2022-23" dpk1201
'vbMessgaeCaption = "ITR 4: AY: 2023-24"
vbMessgaeCaption = "ITR 4: AY: 2024-25"
Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Taxes Paid and Verification")
    

            
'    If Not ValidateIBANInfo Then
'        sourceSheet.Activate
'        MsgBox (Msgbox_BA), vbOKOnly, "Error(s)!"
'        Application.ScreenUpdating = True
'        CloseMsg
'    End If

        
    If Not ValidateSchBA Then
        sourceSheet.Activate
        fmsgboxsmall (Msgbox_BA) ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If

    
End Sub

Function ValidateSchBA() As Boolean

ValidateSchBA = True
Msgbox_BA = ""

setTableInfo_BA_IFSC
setTableInfo_BA
setTableInfo_BA3
setTableInfo_BA_Type_of_Account  'Added by Malli 'TAX DETAILS-C7 AY_2024_25 Added
'setTableInfo_BA44
'setTableInfo_BA5

        'ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA, ColCount1_BA_IFSC, ColCount1_BA_Type_of_Account, ColCount4_BA, ColCount1_BA5)
    ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA, ColCount1_BA_IFSC, ColCount1_BA_Type_of_Account, ColCount4_BA, ColCount1_BA5)
    
    If ColCount1_BA_IFSC > 0 Then
        
        'If Not ValidateCheckBox_BA Then ValidateSchBA = False
        If Not ValidateIFSC Then ValidateSchBA = False
'        If Not ValidateBankName_BA Then ValidateSchBA = False
        If Not ValidateAccntNumber_BA Then ValidateSchBA = False
         
        'ValidateType_of_Account
        'TAX DETAILS-C7 AY_2024_25 Added
        If Not ValidateType_of_Account Then ValidateSchBA = False 'Malli
  Else
      
     Msgbox_BA = Msgbox_BA + "* Please enter Bank Account Details in Sheet Taxes Paid and Verification" & Chr(13)
     ValidateSchBA = False
     Exit Function
     
  End If
  

End Function

Function AccountType() As Boolean
AccountType = True
'setTableInfo_BA3
Dim rangecells As Range
Set rangecells = Range("SchBA.AccountType").Cells
Dim i As Long
ReDim AccountType_BA(ColCount1_BA_IFSC)

For i = 1 To ColCount1_BA_IFSC
    AccountType_BA(i) = rangecells.item(i).value
   
   
    If Not chkCompulsory(AccountType_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Type of Account  at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        AccountType = False
        Exit Function
    End If
   
   
        
    If AccountType_BA(i) = "(Select)" Then
        Msgbox_BA = Msgbox_BA + "* Type of Account  at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        AccountType = False
        Exit Function
    End If
    
    
    Next
End Function


Function ValidateDepositCash() As Boolean
    ValidateDepositCash = True
    'setTableInfo_BA3
    Dim rangecells As Range
    Set rangecells = Range("SchBA.DepositedCash").Cells
    Dim i, j As Long
    ReDim CashDeposited_BA(ColCount1_BA_IFSC)

    For i = 1 To ColCount1_BA_IFSC
        CashDeposited_BA(i) = rangecells.item(i).value

        If Trim(CashDeposited_BA(i)) <> "" Then
            For j = 1 To Len(CashDeposited_BA(i))
                If Not IsNumeric(Mid(CashDeposited_BA(i), j, 1)) Then
                    Msgbox_BA = Msgbox_BA + "* Cash deposited in Bank at Sr.No " & i & " in Sheet Taxes Paid and Verification must contain only digits from 0 to 9" & Chr(13)
                    ValidateDepositCash = False
                    Exit Function
                End If
            Next
        End If
    Next
End Function
Function ValidateIFSC() As Boolean
ValidateIFSC = True
'setTableInfo_BA_IFSC
setTableInfo_BA5
setTableInfo_BA6
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim rangecells4 As Range

Dim rangecells5 As Range 'Malli



Set rangecells1 = Range("SchBA.BankName").Cells
Set rangecells2 = Range("SchBA.AcntNo").Cells
Set rangecells3 = Range("SchBA.IFSC").Cells
Set rangecells4 = Range("tempXML").Cells

Set rangecells5 = Range("SchBA.Type_of_Account").Cells

Dim i As Long

ReDim BankIFSC_BA(ColCount1_BA_IFSC)
ReDim BankName_BA(ColCount1_BA_IFSC)
ReDim BankAccntnum_BA(ColCount1_BA_IFSC)
ReDim Type_of_Account_BA(ColCount1_BA_IFSC) 'malli
ReDim tempXML(ColCount1_BA_IFSC)

'  If Not ColCount1_BA5 > 0 Then
'        Msgbox_BA = Msgbox_BA + "Please select atleast one account in which you prefer to get your refund  in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'    End If
'

'Ayush

  If Not ColCount1_BA5 > 0 Then
       Msgbox_BA = Msgbox_BA + "* ""Please select atleast one account in which you prefer to get your refund""." & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If

'Ayush_28/04
If ColCount55_BA6_Other > 0 Or Not ColCount1_BA5 > 0 Then
'If Not ColCount55_BA6_Other > 0 Then  'Malli_20/03/2025
'          Msgbox_BA = Msgbox_BA + "*""Please select atleast one saving/current/cash credit/overdraft/non resident bank acount  in which you prefer to get your refund""." & Chr(13)
'Newly modified by sai on 24/04/2025
          Msgbox_BA = Msgbox_BA + "*""Bank account with type “Others” cannot be selected for refund credit as its not eligible for refund.""" & Chr(13)
          
          ValidateIFSC = False
          Exit Function
    End If

For i = 1 To ColCount1_BA_IFSC

    BankName_BA(i) = rangecells1.item(i).value
    BankAccntnum_BA(i) = rangecells2.item(i).value
    BankIFSC_BA(i) = rangecells3.item(i).value
    tempXML(i) = rangecells4.item(i).value
    Type_of_Account_BA(i) = rangecells5.item(i).value 'Malli
    
    If BankName_BA(i) = "" Then
            Msgbox_BA = Msgbox_BA + "* Please enter the Bank name in Bank Details at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
            ValidateIFSC = False
            Exit Function
    End If
    
    If Len(BankName_BA(i)) > 125 Then
        Msgbox_BA = Msgbox_BA + "* Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 125 characters" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    
'    If isdropdownblank(BankIFSC_BA(i)) Then
'        Msgbox_BA = Msgbox_BA + "* Please enter the IFS Code of the Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'
'    End If
    'commeneted by Sai---------
    'Ayush_21
    If Not chkCompulsory(BankIFSC_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "IFS Code of the Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory." & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    If Len(BankIFSC_BA(i)) > 11 Then
        Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 11 characters" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    'If UCase(BankIFSC_BA(i)) <> "NNNN0NNNNNN" Then
        If Not ValidateIFSCList(UCase(BankIFSC_BA(i))) Then
            Msgbox_BA = Msgbox_BA + "* IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details at Sr.No " & i & " in Sheet Taxes Paid and verification" & Chr(13)
            ValidateIFSC = False
            Exit Function
        End If
    'End If
    
'    If Not chkCompulsory(BankAccntnum_BA(i)) Then
'        Msgbox_BA = Msgbox_BA + "* Please enter the Account Number in Bank Details at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'    End If
       
       'Malli
'       If (Type_of_Account_BA(i) = "" Or Type_of_Account_BA(i) = "(Select)") Then
'        Msgbox_BA = Msgbox_BA + "* Please select the  Type of Account  at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'       End If
       
'        If Not ValidateAccntNumber_BA(BankAccntnum_BA(i), i) Then
'        ValidateIFSC = False
'        Exit Function
'        End If
    
    
    Next
End Function

'
'Function ValidateBankName_BA() As Boolean
'ValidateBankName_BA = True
'Dim rangecells As Range
'Set rangecells = Range("SchBA.BankName").Cells
'Dim a As Long
'
''a = Range("SchBA.BankName").Count
'Dim i As Long
'ReDim BankName_BA(ColCount1_BA_IFSC)
'For i = 1 To ColCount1_BA_IFSC
'      'BankName_BA(i) = rangecells.item(i).value
'
''        If BankName_BA(i) = "" Then
''            Msgbox_BA = Msgbox_BA + "Name of the Bank at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
''            ValidateBankName_BA = False
''            Exit Function
''        End If
'
'    Next
'End Function
Function ValidateAccntNumber_BA() As Boolean
ValidateAccntNumber_BA = True

Dim rangecells As Range
Set rangecells = Range("SchBA.AcntNo").Cells
Dim i As Long
ReDim BankAccntnum_BA(ColCount1_BA_IFSC)
For i = 1 To ColCount1_BA_IFSC
    BankAccntnum_BA(i) = rangecells.item(i).value
'
'    If Not chkCompulsory(BankAccntnum_BA(i)) Then
'        Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
'        ValidateAccntNumber_BA = False
'        Exit Function
'    End If
'
    If Not ValidateBankAccountNumber_BA(BankAccntnum_BA(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNumber_BA = False
        Exit Function
    End If
    
    
    
    
    Next
End Function
Function ValidateCheckBox_BA() As Boolean
ValidateCheckBox_BA = True

Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range 'Malli

Set rangecells1 = Range("SchBA.IFSC").Cells
Set rangecells = Range("tempXML").Cells
Set rangecells2 = Range("SchBA.Type_of_Account").Cells 'Malli


Dim i As Long
setTableInfo_BA44
ReDim tempXML(ColCount55_BA)
ReDim BankIFSC_BA(ColCount55_BA)
ReDim Type_of_Account_BA(ColCount55_BA)
For i = 1 To ColCount55_BA
    tempXML(i) = rangecells.item(i).value
    BankIFSC_BA(i) = rangecells1.item(i).value
    Type_of_Account_BA(i) = rangecells2.item(i).value  'Malli
    
     
     If isdropdownblank(BankIFSC_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Please enter the IFS Code of the Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateCheckBox_BA = False
        Exit Function

    End If
    
    If Len(BankIFSC_BA(i)) > 11 Then
        Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 11 characters" & Chr(13)
        ValidateCheckBox_BA = False
        Exit Function
    End If
    
    
Next
    
    
    If Not ColCount1_BA5 > 0 Then
        
        
        'TAX DETAILS-E7 AY_2023_24  Old
        'Msgbox_BA = Msgbox_BA + "* Please select atleast one account in which you prefer to get your refund  in Sheet Taxes Paid and Verification" & Chr(13)
        
        'added by malli
        'TAX DETAILS-C7 AY_2024_25 Change
        Msgbox_BA = Msgbox_BA + "* Please select atleast one saving/current/cash credit/overdraft/non resident bank acount  in which you prefer to get your refund in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateCheckBox_BA = False
        Exit Function
    End If
    
    
     
     
     
End Function


Function ValidateNameofHolders_BA() As Boolean
ValidateNameofHolders_BA = True
'setTableInfo_BA2


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
        Msgbox_BA = Msgbox_BA + "* Name of the Joint Account Holder at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 125 characters" & Chr(13)
        ValidateNameofHolders_BA = False
        Exit Function
    End If
End If
    Next
End Function
Function ValidateAccntStatus_BA() As Boolean
ValidateAccntStatus_BA = True
Dim rangecells As Range
Set rangecells = Range("SchBA.AccntStatus").Cells
Dim i As Long
ReDim BankAccntStatus_BA(ColCount1_BA_IFSC)
For i = 1 To ColCount1_BA_IFSC
    BankAccntStatus_BA(i) = rangecells.item(i).value
        
        If BankAccntStatus_BA(i) = "" Or BankAccntStatus_BA(i) = "(Select)" Then
            Msgbox_BA = Msgbox_BA + "* Account Status at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
            ValidateAccntStatus_BA = False
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
    mIntCells = Range("SchBA.IFSC").count
    Set rangecells = Range("SchBA.IFSC").Cells
    For mIntCtr = 1 To mIntCells
            If Not Trim(rangecells.item(mIntCtr).value) = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA_IFSC = ccount
End Sub

Sub setTableInfo_BA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.BankName").count
    Set rangecells = Range("SchBA.BankName").Cells
    For mIntCtr = 1 To mIntCells
            If Not ((Trim(rangecells.item(mIntCtr).value) = "")) Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA = ccount
End Sub
Sub setTableInfo_BA3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.AcntNo").count
    Set rangecells = Range("SchBA.AcntNo").Cells
    For mIntCtr = 1 To mIntCells
            If Not Trim(rangecells.item(mIntCtr).value) = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount4_BA = ccount
End Sub
Sub setTableInfo_BA5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("tempXML").count
    Set rangecells = Range("tempXML").Cells
    For mIntCtr = 1 To mIntCells
            If (rangecells.item(mIntCtr).value) = True Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA5 = ccount
End Sub

'Ayush_21
Sub setTableInfo_BA6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    
    
    'INC_C43 2024-25 Bindu
    Dim ccount1 As Long
    Dim rangecells1 As Range
    Dim mIntCells1 As Long
    
    mIntCells1 = Range("SchBA.Type_of_Account").count
    Set rangecells1 = Range("SchBA.Type_of_Account").Cells
    
    ccount1 = 0
    ccount = 0
    mIntCells = Range("tempxml").count
    Set rangecells = Range("tempxml").Cells
    For mIntCtr = 1 To mIntCells
            If rangecells.item(mIntCtr).value = True Then
               ccount = ccount + 1
               'Malli comented
'               If Trim(rangecells1.item(mIntCtr).value) = Trim("Other") Then
'                     ccount1 = ccount1 + 1
'               End If
               
               'Malli_20/03/2025
               If Trim(rangecells1.item(mIntCtr).value) = Trim("Other") Then
                     ccount1 = ccount1 + 1
               End If
           End If
    Next
    ColCount1_BA5 = ccount
    ColCount55_BA6_Other = ccount1
End Sub
Sub setTableInfo_BA44()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim j As Long
     Dim b As Long

    ccount = 0

    mIntCells = Sheet5.Range("tempXML").count
    Set rangecells = Sheet5.Range("tempXML").Cells
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
Sub setTableInfo_AccntStat_BA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.AccntStatus").count
    Set rangecells = Range("SchBA.AccntStatus").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
               ccount = ccount + 1
           End If
    Next
    ColCountAccntStat_BA = ccount
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
            If Not (rangecells.item(mIntCtr).value = "") Then
               ccount = ccount + 1
           End If
    Next
    ColCount5_BA = ccount
End Sub

Function ValidateBankAccountNumber_BARow1(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_BARow1 = True
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
    
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_BARow1 = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in sheet Taxes Paid and Verification " & Chr(13)
            ValidateBankAccountNumber_BARow1 = False
            Exit Function
        End If
    
    End If
    
    
        If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification  " & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        Msgbox_BA = Msgbox_BA & "Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
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
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification." & Chr(13)
            ValidateBankAccountNumber_BARow1 = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & cc & " in Sheet Taxes Paid and Verification is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If
End Function
Function ValidateBankAccountNumber_BA1(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_BA1 = True
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
    
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_BA1 = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in sheet Taxes Paid and Verification " & Chr(13)
            ValidateBankAccountNumber_BA1 = False
            Exit Function
        End If
    
    End If
    
    
        If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 40) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA1 = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA1 = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification  " & Chr(13)
        ValidateBankAccountNumber_BA1 = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA1 = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA1 = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
        ValidateBankAccountNumber_BA1 = False
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
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification." & Chr(13)
            ValidateBankAccountNumber_BA1 = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " in Sheet Taxes Paid and Verification is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_BA1 = False
        Exit Function
    End If
End Function

Function ValidateBankAccountNumber_BA(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_BA = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB(), ValidateIFSC As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in sheet Taxes Paid and Verification " & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        Msgbox_BA = Msgbox_BA & "*  Please enter the Account Number in Bank Details at Sr.No " & cc & " in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification  " & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
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
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc + 1 & " is invalid in Sheet Taxes Paid and Verification." & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & cc + 1 & " in Sheet Taxes Paid and Verification is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_BA = False
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
                Msgbox_BA = Msgbox_BA + "* Please Enter all the Mandatory fiels at Sr.No " & mIntCtr & " in Schedule Taxes Paid and Verification" & Chr(13)
                checkforNonMandatoryCol = False
                Exit Function
           End If
    Next
End Function


Function ValidateMismatchSchBA()
On Error Resume Next
ValidateMismatchSchBA = True
Msgbox_BA = ""
If (ColCount1_BA_IFSC + 1) <> (Sheet5.Range("NumBankAccountsHeld").value) Then
   ValidateMismatchSchBA = False
   Msgbox_BA = Msgbox_BA + "* Sheet Taxes Paid and Verification: No. of bank accounts held during the P.Y. should be equal to the number of rows filled in Bank Details."
   Sheet5.Activate
   Sheet5.Range("NumBankAccountsHeld").Activate
End If
    
End Function




Function CheckAccountStatus(Tval As Variant, Tadd As Variant)
On Error Resume Next
Application.EnableEvents = False


'------------------Un Protect--------------------
   sPassword = mdCommon.getmsgstate
   Sheet5.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------

Dim rangecellsBA, rangecells1BA As Range
    Dim mIntcellsBA, q As Long
    Dim PrevVal As Variant

Set rangecellsBA = Range("SchBA.AccntStatus").Cells
Set rangecells1BA = Range("SchBA.AccntBalance").Cells


mIntcellsBA = Range("SchBA.AccntStatus").count

'If Sheet5.Range(Tadd).Offset(0, 1) <> "" Then
'Else

If Tval = "Existing" Then
Sheet5.Range(Tadd).Select
Sheet5.Range(Tadd).Offset(0, 1) = ""
Sheet5.Range(Tadd).Offset(0, 1).Locked = False
Sheet5.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
GoTo endfd
Exit Function
End If
'End If


For q = 1 To mIntcellsBA
    
    If rangecellsBA.item(q).value = "(Select)" Then
        rangecells1BA.item(q).value = ""
    End If
    
 
    If rangecellsBA.item(q).value = "Closed" Then
        rangecells1BA.item(q).value = 0
        rangecells1BA.item(q).Locked = True
        rangecells1BA.item(q).Interior.Color = (&HD8D8D8)
    Else
       'rangecells1BA.item(q).value = ""
       rangecells1BA.item(q).Locked = False
       rangecells1BA.item(q).Interior.Color = (&HCCFFCC)
    End If

Next

endfd:
'----------------Lock Password-------------------START---
  Sheet5.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True

End Function




Sub CheckAccountStatus1()
On Error Resume Next
Application.EnableEvents = False

'------------------Un Protect--------------------
'   sPassword = mdCommon.getmsgstate
'   Sheet5.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------

Dim rangecellsBA, rangecells1BA, Banknamerng As Range
    Dim mIntcellsBA, q As Long
Set rangecellsBA = Range("SchBA.AccntStatus").Cells
Set rangecells1BA = Range("SchBA.AccntBalance").Cells
Set Banknamerng = Range("SchBA.BankName").Cells

mIntcellsBA = Range("SchBA.AccntStatus").count


For q = 1 To mIntcellsBA
    If rangecellsBA.item(q).value = "" Then rangecellsBA.item(q).value = "(Select)"
    'If Banknamerng.item(q).value = "" Then Banknamerng.item(q).value = "(Select)"
Next


'----------------Lock Password-------------------START---
 ' Sheet5.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Sub

Function CheckIFSC1(Tval As Variant, Tadd As Variant)
On Error Resume Next
Application.EnableEvents = False


If Tval <> "" Then

      
       If Not CheckIFSC(UCase(Tval)) Then
           fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
           Sheet5.Range(Tadd) = ""
           Sheet5.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If

      
     
     If Not ValidateIFSCList(UCase(Tval)) Then
            fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
            Sheet5.Range(Tadd) = ""
            Sheet5.Range(Tadd).Offset(0, 1) = ""
            GoTo endfd
      End If
    Sheet5.Range(Tadd).Select
    Sheet5.Range(Tadd).Offset(0, 1) = UCase(Sheet5.Range(Tadd).Offset(0, 9))
End If
If Tval = "" Then
    Sheet5.Range(Tadd).Offset(0, 1) = ""
End If

endfd:

Application.EnableEvents = True

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


'End If
End If
        
    
End Function


'added by malli


'TAX DETAILS-C7 AY_2024_25 Added
Sub setTableInfo_BA_Type_of_Account()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.Type_of_Account").count
    Set rangecells = Range("SchBA.Type_of_Account").Cells
    For mIntCtr = 1 To mIntCells
            If Not (Trim(rangecells.item(mIntCtr).value) = "" Or Trim(rangecells.item(mIntCtr).value) = "(Select)") Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA_Type_of_Account = ccount
End Sub

'malli
'TAX DETAILS-C7 AY_2024_25 Added
Function ValidateType_of_Account() As Boolean
ValidateType_of_Account = True
'setTableInfo_BA_IFSC
 

Dim rangecells5 As Range 'Malli

Set rangecells5 = Range("SchBA.Type_of_Account").Cells

Dim i As Long

 
ReDim Type_of_Account_BA(ColCount1_BA_IFSC) 'malli
 
 setTableInfo_BA44

For i = 1 To ColCount1_BA_IFSC

     
    Type_of_Account_BA(i) = rangecells5.item(i).value 'Malli
    
     
    
     If (Type_of_Account_BA(i) = "" Or Type_of_Account_BA(i) = "(Select)") Then
        Msgbox_BA = Msgbox_BA + "* Please select the ""Type of Account"" at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateType_of_Account = False
        Exit Function
    End If
    
'    'added by malli
'
'    'TAX DETAILS-C7 AY_2024_25 Added
'
'    If (Mid(Type_of_Account_BA(i), 1, 1) = "O" And ColCount1_BA5 > 0) Then
'        Msgbox_BA = Msgbox_BA & "* Please select atleast one saving/current/cash credit/overdraft/non resident bank acount  in which you prefer to get your refund in Sheet Taxes Paid and Verification " & Chr(13)
'        ValidateType_of_Account = False
'        Exit Function
'    End If
    
'     If (ColCount1_BA5 > 0 And Type_of_Account_BA(i) = "Other") Then
'        Msgbox_BA = Msgbox_BA + "* malli Please select the  Type of Account  at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
'        ValidateType_of_Account = False
'        Exit Function
'    End If
'
'    If ColCount1_BA5 > 0 Then
'    If Type_of_Account_BA(i) = "Other" Then
'        Msgbox_BA = Msgbox_BA + "* Malli Please select the  Type of Account  at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
'        ValidateType_of_Account = False
'        Exit Function
'    End If
'    End If
    
    
Next




End Function







