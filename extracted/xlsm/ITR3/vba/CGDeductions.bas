Attribute VB_Name = "CGDeductions"
Option Explicit

Public msgboxSHEETCG, msgboxNewSHEETCG As String
Public rngname_Ded54B, rngname_Ded54D, rngname_Ded54EC, rngname_Ded54EE, rngname_Ded54GA, rngname_Ded54G, rngname_Ded54F, rngname_Ded54, rngname_Ded115F As Variant
Public end_Ded54B_1, end_Ded54B_2, end_Ded54B_3, end_Ded54B_4, end_Ded54B_5  As Long
Public end_Ded54EC_1, end_Ded54EC_2, end_Ded54EC_3, end_Ded54EC_4  As Long
Public end_Ded54EE_1, end_Ded54EE_2, end_Ded54EE_3, end_Ded54EE_4  As Long
Public end_Ded54GA_1, end_Ded54GA_2, end_Ded54GA_3, end_Ded54GA_4, end_Ded54GA_5 As Long
Public end_Ded54G_1, end_Ded54G_2, end_Ded54G_3, end_Ded54G_4, end_Ded54G_5  As Long
Public end_Ded54D_1, end_Ded54D_2, end_Ded54D_3, end_Ded54D_4, end_Ded54D_5 As Long
Public end_Ded54F_1, end_Ded54F_2, end_Ded54F_3, end_Ded54F_4, end_Ded54F_5 As Long
Public end_Ded54_1, end_Ded54_2, end_Ded54_3, end_Ded54_4, end_Ded54_5  As Long
Public end_Ded115F_1, end_Ded115F_2, end_Ded115F_3, end_Ded115F_4, end_Ded115F_5 As Long
Public DateOfTransfer, DateOfTransferB, DateOfTransferD, DateOfTransferE, DateOfTransferEE, DateOfTransferF, DateOfTransferG, DateOfTransferGA, DateOfTransfer5F As Variant
Public CostOfLandB, CostOfLandD As Variant
Public CostOfAssetG, CostOfAssetGA As Variant
Public CostOfHouse, CostOfHouseF As Variant
Public AmountInvestedE, AmountInvestedEE, AmountInvested5F As Variant
Public AmountDepositedF As Variant
Public DateOfPurchase, DateOfPurchaseB, DateOfPurchaseD, DateOfPurchaseE, DateOfPurchaseEE, DateOfPurchaseF, DateOfPurchaseG, DateOfPurchaseGA, DateOfPurchase5F As Variant
Public CashDeposited, CashDepositedB, CashDepositedD, CashDepositedG, CashDepositedGA As Variant
Public AmountClaimed, AmountClaimedB, AmountClaimedD, AmountClaimedE, AmountClaimedEE, AmountClaimedF, AmountClaimedG, AmountClaimedGA, AmountClaimed5F As Variant
'Newly added for 54
Public BankIFSC, BankAccntnum, dateOfdeposit
'Newly added for 54B
Public DateofDepositB, BankIFSCB, BankAccntnumB
'Newly added for 54D
Public DateofDepositD, BankIFSCD, BankAccntnumD
'Newly added for 54F
Public DateofDepositF, BankIFSCF, BankAccntnumF
'Newly added for 54G
Public DateofDepositG, BankIFSCG, BankAccntnumG
'Newly added for 54GA
Public DateofDepositGA, BankIFSCGA, BankAccntnumGA
Sub setTableInfo_Ded54BDateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54B.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54B.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54B_1 = ccount
    'Newly changed by Bindu as per DE 0.2
    rngname_Ded54B = "Deduction.54B.DateOfTransfer;Deduction.54B.CostOfAgriLand;Deduction.54B.DateOfPurLand;Deduction.54B.AmtdepCG;Deduction.54B.DateofDeposit;Deduction.54B.AcctNum;Deduction.54B.IFSC;Deduction.54B.AmtDedClaimed;"
End Sub

Sub setTableInfo_Ded54BCostOfLand()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54B.CostOfAgriLand").count
    Set rangecells = Sheet13.Range("Deduction.54B.CostOfAgriLand").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54B_2 = ccount
End Sub

Sub setTableInfo_Ded54BDatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54B.DateOfPurLand").count
    Set rangecells = Sheet13.Range("Deduction.54B.DateOfPurLand").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54B_3 = ccount
End Sub

Sub setTableInfo_Ded54BCashDeposited()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54B.AmtdepCG").count
    Set rangecells = Sheet13.Range("Deduction.54B.AmtdepCG").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54B_4 = ccount
End Sub

Sub setTableInfo_Ded54BAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54B.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54B.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54B_5 = ccount
End Sub

Function ValidateDed54BDateTransfer() As Boolean
ValidateDed54BDateTransfer = True
setTableInfo_Ded54BDateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54B.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransferB(end_Ded54B_1)
For i = 1 To end_Ded54B_1
DateOfTransferB(i) = rangecells.item(i).value

If Len(DateOfTransferB(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransferB(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
    ValidateDed54BDateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransferB(i), " D1b ", "Date of transfer of original asset ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54BDateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransferB(i)) Then
        ValidateDed54BDateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54BCostOfLand() As Boolean
     ValidateDed54BCostOfLand = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54B.CostOfAgriLand").Cells
    ReDim CostOfLandB(end_Ded54B_1)
    For i = 1 To end_Ded54B_1
        CostOfLandB(i) = rangecells.item(i).value
        
'     If Not chkMandatory(CostOfLandB(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cost of Land at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54BCostOfLand = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CostOfLandB(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost of Land at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54BCostOfLand = False
        Exit Function
    End If
    
    If (Len(CostOfLandB(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost of Land at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54BCostOfLand = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54BDateOfPurchase() As Boolean
ValidateDed54BDateOfPurchase = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54B.DateOfPurLand").Cells
Dim i As Long
ReDim DateOfPurchaseB(end_Ded54B_1)
For i = 1 To end_Ded54B_1
DateOfPurchaseB(i) = rangecells.item(i).value

If Len(DateOfPurchaseB(i)) = 0 Then
End If

'If Not chkMandatory(DateOfPurchaseB(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54BDateOfPurchase = False
'    Exit Function
'End If

If Not CheckDateddmmyyyy(DateOfPurchaseB(i)) Then
        ValidateDed54BDateOfPurchase = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54BCashDeposited() As Boolean
     ValidateDed54BCashDeposited = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54B.AmtdepCG").Cells
    ReDim CashDepositedB(end_Ded54B_1)
    For i = 1 To end_Ded54B_1
        CashDepositedB(i) = rangecells.item(i).value

'     If Not chkMandatory(CashDepositedB(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cash Deposited at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54BCashDeposited = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CashDepositedB(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54BCashDeposited = False
        Exit Function
    End If
    
    If (Len(CashDepositedB(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54BCashDeposited = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54BAmountClaimed() As Boolean
     ValidateDed54BAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54B.AmtDedClaimed").Cells
    ReDim AmountClaimedB(end_Ded54B_1)
    For i = 1 To end_Ded54B_1
        AmountClaimedB(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimedB(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54BAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimedB(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54BAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimedB(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54BAmountClaimed = False
        Exit Function
    End If
 Next
End Function


Function ValidateDednTable54B_CG() As Boolean
    Dim i As Long
    ValidateDednTable54B_CG = True
    setTableInfo_Ded54BDateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54B:" & Chr(13) & Chr(13)
    If Not ValidateDed54BDateTransfer() Then ValidateDednTable54B_CG = False
    If (Len(Sheet13.Range("Deduction.54B.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54BCostOfLand() Then ValidateDednTable54B_CG = False
        If Not ValidateDed54BDateOfPurchase() Then ValidateDednTable54B_CG = False
        If Not ValidateDed54BCashDeposited() Then ValidateDednTable54B_CG = False
        If Not ValidateDed54BAmountClaimed() Then ValidateDednTable54B_CG = False
        
        'Newly added by Bindu as per DE 0.2
        If Not ValidateDed54BDateofDeposit() Then ValidateDednTable54B_CG = False
        If Not ValidateDedB54IFSC() Then ValidateDednTable54B_CG = False
    End If
    
    setTableInfo_Ded54BCostOfLand
    setTableInfo_Ded54BDatePurchase
    setTableInfo_Ded54BCashDeposited
    setTableInfo_Ded54BAmountClaimed
    
    If (end_Ded54B_1 <> end_Ded54B_5) Then '(end_Ded54B_1 <> end_Ded54B_2) Or (end_Ded54B_1 <> end_Ded54B_3) Or (end_Ded54B_1 <> end_Ded54B_4) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54B." & Chr(13)
             ValidateDednTable54B_CG = False
             Exit Function
    
    End If
    
If end_Ded54B_1 > 0 Then


For i = 1 To end_Ded54B_1

        If CostOfLandB(i) <> "" Or DateOfPurchaseB(i) <> "" Then

         If (CostOfLandB(i) <> "") Then ' Or Ded54DateOfPurchase(i) <> "") Then
         If DateOfPurchaseB(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54B Date Of Purchase at the row " & i & " ." & Chr(13)
            ValidateDednTable54B_CG = False
            Exit Function
         End If
         End If


         If DateOfPurchaseB(i) <> "" Then
         If (CostOfLandB(i) = "") Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54B Cost Of Land at the row " & i & " ." & Chr(13)
            ValidateDednTable54B_CG = False
            Exit Function
         End If
         End If

        Else

         If CashDepositedB(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill either Ded54B Date Of Purchase and Ded54B Cost Of Land or Ded54B Cash Deposited at the row " & i & " ." & Chr(13)
            ValidateDednTable54B_CG = False
            Exit Function
         End If


        End If
 Next

 End If
    
    
    
    
    
End Function

Sub setTableInfo_Ded54DDateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54D.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54D.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54D_1 = ccount
    rngname_Ded54D = "Deduction.54D.DateOfTransfer;Deduction.54D.CostOfPurcConst;Deduction.54B.DateOfPurLandB;Deduction.54D.AmtdepCG;Deduction.54D.AmtDedClaimed;"
End Sub


Sub setTableInfo_Ded54DCost()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54D.CostOfPurcConst").count
    Set rangecells = Sheet13.Range("Deduction.54D.CostOfPurcConst").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54D_2 = ccount
End Sub

Sub setTableInfo_Ded54DDatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54B.DateOfPurLandB").count
    Set rangecells = Sheet13.Range("Deduction.54B.DateOfPurLandB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54D_3 = ccount
End Sub

Sub setTableInfo_Ded54DCashDeposited()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54D.AmtdepCG").count
    Set rangecells = Sheet13.Range("Deduction.54D.AmtdepCG").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54D_4 = ccount
End Sub

Sub setTableInfo_Ded54DAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54D.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54D.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54D_5 = ccount
End Sub

Function ValidateDed54DDateTransfer() As Boolean
ValidateDed54DDateTransfer = True
setTableInfo_Ded54DDateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54D.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransferD(end_Ded54D_1)
For i = 1 To end_Ded54D_1
DateOfTransferD(i) = rangecells.item(i).value

If Len(DateOfTransferD(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransferD(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
    ValidateDed54DDateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransferD(i), " D1c ", "Date of acquisition of original asset ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54DDateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransferD(i)) Then
        ValidateDed54DDateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54DCostOfLand() As Boolean
     ValidateDed54DCostOfLand = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54D.CostOfPurcConst").Cells
    ReDim CostOfLandD(end_Ded54D_1)
    For i = 1 To end_Ded54D_1
        CostOfLandD(i) = rangecells.item(i).value
        
'      If Not chkMandatory(CostOfLandD(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cost of Land at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54DCostOfLand = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CostOfLandD(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost of Land at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54DCostOfLand = False
        Exit Function
    End If
    
    If (Len(CostOfLandD(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost of Land at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54DCostOfLand = False
        Exit Function

    End If
 Next
End Function

Function ValidateDed54DDateOfPurchase() As Boolean
ValidateDed54DDateOfPurchase = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54B.DateOfPurLandB").Cells
Dim i As Long
ReDim DateOfPurchaseD(end_Ded54D_1)
For i = 1 To end_Ded54D_1
DateOfPurchaseD(i) = rangecells.item(i).value

If Len(DateOfPurchaseD(i)) = 0 Then
End If

'If Not chkMandatory(DateOfPurchaseD(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54DDateOfPurchase = False
'    Exit Function
'End If


If Not CheckDateddmmyyyy(DateOfPurchaseD(i)) Then
        ValidateDed54DDateOfPurchase = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54DCashDeposited() As Boolean
     ValidateDed54DCashDeposited = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54D.AmtdepCG").Cells
    ReDim CashDepositedD(end_Ded54D_1)
    For i = 1 To end_Ded54D_1
        CashDepositedD(i) = rangecells.item(i).value
        
'     If Not chkMandatory(CashDepositedD(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cash Deposited at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54DCashDeposited = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CashDepositedD(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54DCashDeposited = False
        Exit Function
    End If
    
    If (Len(CashDepositedD(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54DCashDeposited = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54DAmountClaimed() As Boolean
     ValidateDed54DAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54D.AmtDedClaimed").Cells
    ReDim AmountClaimedD(end_Ded54D_1)
    For i = 1 To end_Ded54D_1
        AmountClaimedD(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimedD(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54DAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimedD(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54DAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimedD(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54DAmountClaimed = False
        Exit Function
    End If
 Next
End Function

Function ValidateDednTable54D_CG() As Boolean
    Dim i As Long
    ValidateDednTable54D_CG = True
    setTableInfo_Ded54DDateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54:" & Chr(13) & Chr(13)
    If Not ValidateDed54DDateTransfer() Then ValidateDednTable54D_CG = False
    If (Len(Sheet13.Range("Deduction.54D.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54DCostOfLand() Then ValidateDednTable54D_CG = False
        If Not ValidateDed54DDateOfPurchase() Then ValidateDednTable54D_CG = False
        If Not ValidateDed54DCashDeposited() Then ValidateDednTable54D_CG = False
        If Not ValidateDed54DAmountClaimed() Then ValidateDednTable54D_CG = False
        'Newly added by Bindu as per DE 0.2
        If Not ValidateDed54DDateofDeposit() Then ValidateDednTable54D_CG = False
        If Not ValidateDedD54IFSC() Then ValidateDednTable54D_CG = False
        
    End If
    
    setTableInfo_Ded54DCost
    setTableInfo_Ded54DDatePurchase
    setTableInfo_Ded54DCashDeposited
    setTableInfo_Ded54DAmountClaimed
    
    If (end_Ded54D_1 <> end_Ded54D_5) Then '(end_Ded54D_1 <> end_Ded54D_2) Or (end_Ded54D_1 <> end_Ded54D_3) Or (end_Ded54D_1 <> end_Ded54D_4) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54D." & Chr(13)
             ValidateDednTable54D_CG = False
             Exit Function
    
    End If
    
If end_Ded54D_1 > 0 Then

For i = 1 To end_Ded54D_1

        If CostOfLandD(i) <> "" Or DateOfPurchaseD(i) <> "" Then

         If CostOfLandD(i) <> "" Then ' Or Ded54DateOfPurchase(i) <> "") Then
         If DateOfPurchaseD(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54D Date Of Purchase at the row " & i & " ." & Chr(13)
            ValidateDednTable54D_CG = False
            Exit Function
         End If
         End If


         If DateOfPurchaseD(i) <> "" Then
         If CostOfLandD(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54D Cost Of Land at the row " & i & " ." & Chr(13)
            ValidateDednTable54D_CG = False
            Exit Function
         End If
         End If

        Else

         If CashDepositedD(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill either Ded54D Date Of Purchase and Ded54D Cost Of Land or Ded54D Cash Deposited at the row " & i & " ." & Chr(13)
            ValidateDednTable54D_CG = False
            Exit Function
         End If


        End If
Next

End If
    
    
    
End Function

Sub setTableInfo_Ded54ECDateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54EC.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54EC.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54EC_1 = ccount
    rngname_Ded54EC = "Deduction.54EC.DateOfTransfer;Deduction.54EC.AmountInvestd;Deduction.54EC.DateOfInvestment;Deduction.54EC.AmtDedClaimed"
End Sub

Sub setTableInfo_Ded54ECAmountInvested()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54EC.AmountInvestd").count
    Set rangecells = Sheet13.Range("Deduction.54EC.AmountInvestd").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54EC_2 = ccount
End Sub

Sub setTableInfo_Ded54ECDatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54EC.DateOfInvestment").count
    Set rangecells = Sheet13.Range("Deduction.54EC.DateOfInvestment").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54EC_3 = ccount
End Sub

Sub setTableInfo_Ded54ECAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54EC.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54EC.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54EC_4 = ccount
End Sub

Function ValidateDed54ECDateTransfer() As Boolean
ValidateDed54ECDateTransfer = True
setTableInfo_Ded54ECDateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54EC.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransferE(end_Ded54EC_1)
For i = 1 To end_Ded54EC_1
DateOfTransferE(i) = rangecells.item(i).value

If Len(DateOfTransferE(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransferE(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
    ValidateDed54ECDateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransferE(i), " D1d ", "Date of transfer of original asset ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54ECDateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransferE(i)) Then
        ValidateDed54ECDateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54ECAmountInvested() As Boolean
     ValidateDed54ECAmountInvested = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54EC.AmountInvestd").Cells
    ReDim AmountInvestedE(end_Ded54EC_1)
    For i = 1 To end_Ded54EC_1
        AmountInvestedE(i) = rangecells.item(i).value
        
'     If Not chkMandatory(AmountInvestedE(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54ECAmountInvested = False
'         Exit Function
'     End If
     
     If Not IsNumeric(AmountInvestedE(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54ECAmountInvested = False
        Exit Function
    End If
    
    If (Len(AmountInvestedE(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54ECAmountInvested = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54ECDateOfInvestment() As Boolean
ValidateDed54ECDateOfInvestment = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54EC.DateOfInvestment").Cells
Dim i As Long
ReDim DateOfPurchaseE(end_Ded54EC_1)
For i = 1 To end_Ded54EC_1
DateOfPurchaseE(i) = rangecells.item(i).value

If Len(DateOfPurchaseE(i)) = 0 Then
End If

'If Not chkMandatory(DateOfPurchaseE(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Investment at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54ECDateOfInvestment = False
'    Exit Function
'End If

If Not CheckDateddmmyyyy(DateOfPurchaseE(i)) Then
        ValidateDed54ECDateOfInvestment = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Investment at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54ECAmountClaimed() As Boolean
     ValidateDed54ECAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54EC.AmtDedClaimed").Cells
    ReDim AmountClaimedE(end_Ded54EC_1)
    For i = 1 To end_Ded54EC_1
        AmountClaimedE(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimedE(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54ECAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimedE(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54ECAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimedE(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54ECAmountClaimed = False
        Exit Function
    End If
 Next
End Function


Function ValidateDednTable54EC_CG() As Boolean
    ValidateDednTable54EC_CG = True
    setTableInfo_Ded54ECDateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54EC:" & Chr(13) & Chr(13)
    If Not ValidateDed54ECDateTransfer() Then ValidateDednTable54EC_CG = False
    If (Len(Sheet13.Range("Deduction.54EC.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54ECAmountInvested() Then ValidateDednTable54EC_CG = False
        If Not ValidateDed54ECDateOfInvestment() Then ValidateDednTable54EC_CG = False
        If Not ValidateDed54ECAmountClaimed() Then ValidateDednTable54EC_CG = False
    End If
    
    setTableInfo_Ded54ECAmountInvested
    setTableInfo_Ded54ECDatePurchase
    setTableInfo_Ded54ECAmountClaimed
    
    If (end_Ded54EC_1 <> end_Ded54EC_4) Then  '(end_Ded54EC_1 <> end_Ded54EC_2) Or (end_Ded54EC_1 <> end_Ded54EC_3) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54EC." & Chr(13)
             ValidateDednTable54EC_CG = False
             Exit Function
    
    End If
    
End Function

'Sub setTableInfo_Ded54EEDateTransfer()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    mIntCells = Sheet13.Range("Deduction.54EE.DateOfTransfer").Count
'    Set rangecells = Sheet13.Range("Deduction.54EE.DateOfTransfer").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'               ccount = ccount + 1
'           End If
'    Next
'    end_Ded54EE_1 = ccount
'    rngname_Ded54EE = "Deduction.54EE.DateOfTransfer;Deduction.54EE.AmountInvestd;Deduction.54EE.DateOfInvestment;Deduction.54EE.AmtDedClaimed"
'End Sub
'
'Sub setTableInfo_Ded54EEAmountInvested()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    mIntCells = Sheet13.Range("Deduction.54EE.AmountInvestd").Count
'    Set rangecells = Sheet13.Range("Deduction.54EE.AmountInvestd").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'               ccount = ccount + 1
'           End If
'    Next
'    end_Ded54EE_2 = ccount
'End Sub
'
'Sub setTableInfo_Ded54EEDatePurchase()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    mIntCells = Sheet13.Range("Deduction.54EE.DateOfInvestment").Count
'    Set rangecells = Sheet13.Range("Deduction.54EE.DateOfInvestment").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'               ccount = ccount + 1
'           End If
'    Next
'    end_Ded54EE_3 = ccount
'End Sub
'
'Sub setTableInfo_Ded54EEAmountClaimed()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    mIntCells = Sheet13.Range("Deduction.54EE.AmtDedClaimed").Count
'    Set rangecells = Sheet13.Range("Deduction.54EE.AmtDedClaimed").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'               ccount = ccount + 1
'           End If
'    Next
'    end_Ded54EE_4 = ccount
'End Sub
'
'Function ValidateDed54EEDateTransfer() As Boolean
'ValidateDed54EEDateTransfer = True
'setTableInfo_Ded54EEDateTransfer
'Dim rangecells As Range
'Set rangecells = Sheet13.Range("Deduction.54EE.DateOfTransfer").Cells
'Dim i As Long
'ReDim DateOfTransferEE(end_Ded54EE_1)
'For i = 1 To end_Ded54EE_1
'DateOfTransferEE(i) = rangecells.item(i).value
'
'If Len(DateOfTransferEE(i)) = 0 Then
'End If
'
'If Not chkMandatory(DateOfTransferEE(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54EEDateTransfer = False
'    Exit Function
'End If
'
'If Not FormatNCheckDate(DateOfTransferEE(i)) Then
'        ValidateDed54EEDateTransfer = False
'        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
'        Exit Function
'End If
'
'Next
'End Function
'
'Function ValidateDed54EEAmountInvested() As Boolean
'     ValidateDed54EEAmountInvested = True
'    Dim rangecells As Range
'       Dim i As Long
'    Set rangecells = Sheet13.Range("Deduction.54EE.AmountInvestd").Cells
'    ReDim AmountInvestedEE(end_Ded54EE_1)
'    For i = 1 To end_Ded54EE_1
'        AmountInvestedEE(i) = rangecells.item(i).value
'
''     If Not chkMandatory(AmountInvestedEE(i)) Then
''         msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr. No  " & i & " is Mandatory." & Chr(13)
''         ValidateDed54EEAmountInvested = False
''         Exit Function
''     End If
'
'     If Not IsNumeric(AmountInvestedEE(i)) Then
'        msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr.No " & i & " Should be a numeric value." & Chr(13)
'        ValidateDed54EEAmountInvested = False
'        Exit Function
'    End If
'
'    If (Len(AmountInvestedEE(i)) > 14) Then
'        msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
'        ValidateDed54EEAmountInvested = False
'        Exit Function
'    End If
' Next
'End Function
'
'Function ValidateDed54EEDateOfInvestment() As Boolean
'ValidateDed54EEDateOfInvestment = True
'Dim rangecells As Range
'Set rangecells = Sheet13.Range("Deduction.54EE.DateOfInvestment").Cells
'Dim i As Long
'ReDim DateOfPurchaseEE(end_Ded54EE_1)
'For i = 1 To end_Ded54EE_1
'DateOfPurchaseEE(i) = rangecells.item(i).value
'
'If Len(DateOfPurchaseEE(i)) = 0 Then
'End If
'
''If Not chkMandatory(DateOfPurchaseEE(i)) Then
''    msgboxSHEETCG = msgboxSHEETCG + "* Date of Investment at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
''    ValidateDed54EEDateOfInvestment = False
''    Exit Function
''End If
'
'If Not FormatNCheckDate(DateOfPurchaseEE(i)) Then
'        ValidateDed54EEDateOfInvestment = False
'        msgboxSHEETCG = msgboxSHEETCG + "* Date of Investment at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
'        Exit Function
'End If
'
'Next
'End Function
'
'Function ValidateDed54EEAmountClaimed() As Boolean
'     ValidateDed54EEAmountClaimed = True
'    Dim rangecells As Range
'       Dim i As Long
'    Set rangecells = Sheet13.Range("Deduction.54EE.AmtDedClaimed").Cells
'    ReDim AmountClaimedEE(end_Ded54EE_1)
'    For i = 1 To end_Ded54EE_1
'        AmountClaimedEE(i) = rangecells.item(i).value
'
'     If Not chkMandatory(AmountClaimedEE(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54EEAmountClaimed = False
'         Exit Function
'     End If
'
'     If Not IsNumeric(AmountClaimedEE(i)) Then
'        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
'        ValidateDed54EEAmountClaimed = False
'        Exit Function
'    End If
'
'    If (Len(AmountClaimedEE(i)) > 14) Then
'        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
'        ValidateDed54EEAmountClaimed = False
'        Exit Function
'    End If
' Next
'End Function
'
'
'Function ValidateDednTable54EE_CG() As Boolean
'    ValidateDednTable54EE_CG = True
'    setTableInfo_Ded54EEDateTransfer
'
''    msgboxSHEETCG = ""
''    msgboxSHEETCG = "Deduction u/s 54EE:" & Chr(13) & Chr(13)
'    If Not ValidateDed54EEDateTransfer() Then ValidateDednTable54EE_CG = False
'    If (Len(Sheet13.Range("Deduction.54EE.DateOfTransfer").item(0).value) > 0) Then
'        If Not ValidateDed54EEAmountInvested() Then ValidateDednTable54EE_CG = False
'        If Not ValidateDed54EEDateOfInvestment() Then ValidateDednTable54EE_CG = False
'        If Not ValidateDed54EEAmountClaimed() Then ValidateDednTable54EE_CG = False
'    End If
'
'    setTableInfo_Ded54EEAmountInvested
'    setTableInfo_Ded54EEDatePurchase
'    setTableInfo_Ded54EEAmountClaimed
'
'    If (end_Ded54EE_1 <> end_Ded54EE_4) Then '(end_Ded54EE_1 <> end_Ded54EE_2) Or (end_Ded54EE_1 <> end_Ded54EE_3) Or
'
'             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54EE." & Chr(13)
'             ValidateDednTable54EE_CG = False
'             Exit Function
'
'    End If
'
'End Function

Sub setTableInfo_Ded54GDateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54G.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54G.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54G_1 = ccount
    rngname_Ded54G = "Deduction.54G.DateOfTransfer;Deduction.54G.CostOfExpenses;Deduction.54G.DateOfPurc;Deduction.54G.AmtdepCG;Deduction.54G.DateofDeposit;Deduction.54G.AcctNum;Deduction.54G.IFSC;Deduction.54G.AmtDedClaimed"
End Sub

Sub setTableInfo_Ded54GCostOfAsset()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54G.CostOfExpenses").count
    Set rangecells = Sheet13.Range("Deduction.54G.CostOfExpenses").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54G_2 = ccount
End Sub

Sub setTableInfo_Ded54GDatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54G.DateOfPurc").count
    Set rangecells = Sheet13.Range("Deduction.54G.DateOfPurc").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54G_3 = ccount
End Sub

Sub setTableInfo_Ded54GCashDeposited()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54G.AmtdepCG").count
    Set rangecells = Sheet13.Range("Deduction.54G.AmtdepCG").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54G_4 = ccount
End Sub

Sub setTableInfo_Ded54GAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54G.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54G.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54G_5 = ccount
End Sub

Function ValidateDed54GDateTransfer() As Boolean
ValidateDed54GDateTransfer = True
setTableInfo_Ded54GDateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54G.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransferG(end_Ded54G_1)
For i = 1 To end_Ded54G_1
DateOfTransferG(i) = rangecells.item(i).value

If Len(DateOfTransferG(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransferG(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
    ValidateDed54GDateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransferG(i), " D1f ", "Date of transfer of original asset from urban area ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54GDateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransferG(i)) Then
        ValidateDed54GDateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54GCostOfAsset() As Boolean
     ValidateDed54GCostOfAsset = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54G.CostOfExpenses").Cells
    ReDim CostOfAssetG(end_Ded54G_1)
    For i = 1 To end_Ded54G_1
        CostOfAssetG(i) = rangecells.item(i).value
        
'    If Not chkMandatory(CostOfAssetG(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cost Of Asset at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54GCostOfAsset = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CostOfAssetG(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost Of Asset at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54GCostOfAsset = False
        Exit Function
    End If
    
    If (Len(CostOfAssetG(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost Of Asset at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54GCostOfAsset = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54GDateOfPurchase() As Boolean
ValidateDed54GDateOfPurchase = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54G.DateOfPurc").Cells
Dim i As Long
ReDim DateOfPurchaseG(end_Ded54G_1)
For i = 1 To end_Ded54G_1
DateOfPurchaseG(i) = rangecells.item(i).value

If Len(DateOfPurchaseG(i)) = 0 Then
End If

'If Not chkMandatory(DateOfPurchaseG(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54GDateOfPurchase = False
'    Exit Function
'End If

If Not CheckDateddmmyyyy(DateOfPurchaseG(i)) Then
        ValidateDed54GDateOfPurchase = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54GCashDeposited() As Boolean
     ValidateDed54GCashDeposited = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54G.AmtdepCG").Cells
    ReDim CashDepositedG(end_Ded54G_1)
    For i = 1 To end_Ded54G_1
        CashDepositedG(i) = rangecells.item(i).value
        
'      If Not chkMandatory(CashDepositedG(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54GCashDeposited = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CashDepositedG(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54GCashDeposited = False
        Exit Function
    End If
    
    If (Len(CashDepositedG(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54GCashDeposited = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54GAmountClaimed() As Boolean
     ValidateDed54GAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54G.AmtDedClaimed").Cells
    ReDim AmountClaimedG(end_Ded54G_1)
    For i = 1 To end_Ded54G_1
        AmountClaimedG(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimedG(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54GAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimedG(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54GAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimedG(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54GAmountClaimed = False
        Exit Function
    End If
 Next
End Function


Function ValidateDednTable54G_CG() As Boolean
    Dim i As Long
    ValidateDednTable54G_CG = True
    setTableInfo_Ded54GDateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54G:" & Chr(13) & Chr(13)
    If Not ValidateDed54GDateTransfer() Then ValidateDednTable54G_CG = False
    If (Len(Sheet13.Range("Deduction.54G.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54GCostOfAsset() Then ValidateDednTable54G_CG = False
        If Not ValidateDed54GDateOfPurchase() Then ValidateDednTable54G_CG = False
        If Not ValidateDed54GCashDeposited() Then ValidateDednTable54G_CG = False
        If Not ValidateDed54GAmountClaimed() Then ValidateDednTable54G_CG = False
        
        'Newly added by Bindu as per DE 0.2
        If Not ValidateDed54GDateofDeposit() Then ValidateDednTable54G_CG = False
        If Not ValidateDedG54IFSC() Then ValidateDednTable54G_CG = False

    End If
    
    setTableInfo_Ded54GCostOfAsset
    setTableInfo_Ded54GDatePurchase
    setTableInfo_Ded54GCashDeposited
    setTableInfo_Ded54GAmountClaimed
    
    If (end_Ded54G_1 <> end_Ded54G_5) Then '(end_Ded54G_1 <> end_Ded54G_2) Or (end_Ded54G_1 <> end_Ded54G_3) Or (end_Ded54G_1 <> end_Ded54G_4) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54G." & Chr(13)
             ValidateDednTable54G_CG = False
             Exit Function
    
    End If
    
    
   If end_Ded54G_1 > 0 Then


For i = 1 To end_Ded54G_1

        If CostOfAssetG(i) <> "" Or DateOfPurchaseG(i) <> "" Then

         If CostOfAssetG(i) <> "" Then ' Or Ded54DateOfPurchase(i) <> "") Then
         If DateOfPurchaseG(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54G Date Of Purchase at the row " & i & " ." & Chr(13)
            ValidateDednTable54G_CG = False
            Exit Function
         End If
         End If


         If DateOfPurchaseG(i) <> "" Then
         If CostOfAssetG(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54G Cost Of Asset at the row " & i & " ." & Chr(13)
            ValidateDednTable54G_CG = False
            Exit Function
         End If
         End If

        Else

         If CashDepositedG(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill either Ded54G Date Of Purchase and Ded54G Cost Of Asset or Ded54G Cash Deposited at the row " & i & " ." & Chr(13)
            ValidateDednTable54G_CG = False
            Exit Function
         End If


        End If
Next

End If
    
 
    
End Function

Sub setTableInfo_Ded54GADateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54GA.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54GA.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54GA_1 = ccount
    rngname_Ded54GA = "Deduction.54GA.DateOfTransfer;Deduction.54GA.CostOfExpenses;Deduction.54GA.DateOfPurc;Deduction.54GA.AmtdepCG;Deduction.54GA.DateofDeposit;Deduction.54GA.AcctNum;Deduction.54GA.IFSC;Deduction.54GA.AmtDedClaimed"
End Sub

Sub setTableInfo_Ded54GACostOfAsset()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54GA.CostOfExpenses").count
    Set rangecells = Sheet13.Range("Deduction.54GA.CostOfExpenses").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54GA_2 = ccount
End Sub

Sub setTableInfo_Ded54GADatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54GA.DateOfPurc").count
    Set rangecells = Sheet13.Range("Deduction.54GA.DateOfPurc").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54GA_3 = ccount
End Sub

Sub setTableInfo_Ded54GACashDeposited()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54GA.AmtdepCG").count
    Set rangecells = Sheet13.Range("Deduction.54GA.AmtdepCG").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54GA_4 = ccount
End Sub

Sub setTableInfo_Ded54GAAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54GA.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54GA.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54GA_5 = ccount
End Sub

Function ValidateDed54GADateTransfer() As Boolean
ValidateDed54GADateTransfer = True
setTableInfo_Ded54GADateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54GA.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransferGA(end_Ded54GA_1)
For i = 1 To end_Ded54GA_1
DateOfTransferGA(i) = rangecells.item(i).value

If Len(DateOfTransferGA(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransferGA(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
    ValidateDed54GADateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransferGA(i), " D1g ", "Date of transfer of original asset from urban area ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54GADateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransferGA(i)) Then
        ValidateDed54GADateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54GACostOfAsset() As Boolean
     ValidateDed54GACostOfAsset = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54GA.CostOfExpenses").Cells
    ReDim CostOfAssetGA(end_Ded54GA_1)
    For i = 1 To end_Ded54GA_1
        CostOfAssetGA(i) = rangecells.item(i).value
        
'     If Not chkMandatory(CostOfAssetGA(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cost Of Asset at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54GACostOfAsset = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CostOfAssetGA(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost Of Asset at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54GACostOfAsset = False
        Exit Function
    End If
    
    If (Len(CostOfAssetGA(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost Of Asset at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54GACostOfAsset = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54GADateOfPurchase() As Boolean
ValidateDed54GADateOfPurchase = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54GA.DateOfPurc").Cells
Dim i As Long
ReDim DateOfPurchaseGA(end_Ded54GA_1)
For i = 1 To end_Ded54GA_1
DateOfPurchaseGA(i) = rangecells.item(i).value

If Len(DateOfPurchaseGA(i)) = 0 Then
End If


'If Not chkMandatory(DateOfPurchaseGA(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54GADateOfPurchase = False
'    Exit Function
'End If

If Not CheckDateddmmyyyy(DateOfPurchaseGA(i)) Then
        ValidateDed54GADateOfPurchase = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54GACashDeposited() As Boolean
     ValidateDed54GACashDeposited = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54GA.AmtdepCG").Cells
    ReDim CashDepositedGA(end_Ded54GA_1)
    For i = 1 To end_Ded54GA_1
        CashDepositedGA(i) = rangecells.item(i).value
'     If Not chkMandatory(CashDepositedGA(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cash Deposited at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54GACashDeposited = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CashDepositedGA(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54GACashDeposited = False
        Exit Function
    End If
    
    If (Len(CashDepositedGA(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54GACashDeposited = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54GAAmountClaimed() As Boolean
     ValidateDed54GAAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54GA.AmtDedClaimed").Cells
    ReDim AmountClaimedGA(end_Ded54GA_1)
    For i = 1 To end_Ded54GA_1
        AmountClaimedGA(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimedGA(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54GAAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimedGA(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54GAAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimedGA(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54GAAmountClaimed = False
        Exit Function
    End If
 Next
End Function


Function ValidateDednTable54GA_CG() As Boolean
    Dim i As Long
    ValidateDednTable54GA_CG = True
    setTableInfo_Ded54GADateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54GA:" & Chr(13) & Chr(13)
    If Not ValidateDed54GADateTransfer() Then ValidateDednTable54GA_CG = False
    If (Len(Sheet13.Range("Deduction.54GA.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54GACostOfAsset() Then ValidateDednTable54GA_CG = False
        If Not ValidateDed54GADateOfPurchase() Then ValidateDednTable54GA_CG = False
        If Not ValidateDed54GACashDeposited() Then ValidateDednTable54GA_CG = False
        If Not ValidateDed54GAAmountClaimed() Then ValidateDednTable54GA_CG = False
        
        'Newly added by Bindu as per DE 0.2
        If Not ValidateDed54GADateofDeposit() Then ValidateDednTable54GA_CG = False
        If Not ValidateDedGA54IFSC() Then ValidateDednTable54GA_CG = False

    End If
    
    setTableInfo_Ded54GACostOfAsset
    setTableInfo_Ded54GADatePurchase
    setTableInfo_Ded54GACashDeposited
    setTableInfo_Ded54GAAmountClaimed
    
    If (end_Ded54GA_1 <> end_Ded54GA_5) Then '(end_Ded54GA_1 <> end_Ded54GA_1) Or (end_Ded54GA_1 <> end_Ded54GA_2) Or (end_Ded54GA_1 <> end_Ded54GA_3) Or (end_Ded54GA_1 <> end_Ded54GA_4) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54GA." & Chr(13)
             ValidateDednTable54GA_CG = False
             Exit Function
    
    End If
    
    
If end_Ded54GA_1 > 0 Then

For i = 1 To end_Ded54GA_1

        If CostOfAssetGA(i) <> "" Or DateOfPurchaseGA(i) <> "" Then

         If CostOfAssetGA(i) <> "" Then ' Or Ded54DateOfPurchase(i) <> "" Then
         If DateOfPurchaseGA(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54GA Date Of Purchase at the row " & i & " ." & Chr(13)
            ValidateDednTable54GA_CG = False
            Exit Function
         End If
         End If


         If DateOfPurchaseGA(i) <> "" Then
         If CostOfAssetGA(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54GA Cost Of Asset at the row " & i & " ." & Chr(13)
            ValidateDednTable54GA_CG = False
            Exit Function
         End If
         End If

        Else

         If CashDepositedGA(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill either Ded54GA Date Of Purchase and Ded54GA Cost Of Asset or Ded54GA Cash Deposited at the row " & i & " ." & Chr(13)
            ValidateDednTable54GA_CG = False
            Exit Function
         End If


        End If
Next

End If
    
End Function


Function chkMandatory(field As Variant) As Boolean
chkMandatory = True
If Len(Trim(field)) <= 0 Then
    chkMandatory = False
End If
If IsEmpty(field) Then
chkMandatory = False
End If
End Function

Sub setTableInfo_Ded115FDateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.115F.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.115F.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded115F_1 = ccount
    rngname_Ded115F = "Deduction.115F.DateOfTransfer;Deduction.115F.AmtInvt;Deduction.115F.DateOfInvest;Deduction.115F.AmtDedClaimed"
End Sub

Sub setTableInfo_Ded115FAmountInvested()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.115F.AmtInvt").count
    Set rangecells = Sheet13.Range("Deduction.115F.AmtInvt").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded115F_2 = ccount
End Sub

Sub setTableInfo_Ded115FDatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.115F.DateOfInvest").count
    Set rangecells = Sheet13.Range("Deduction.115F.DateOfInvest").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded115F_3 = ccount
End Sub

Sub setTableInfo_Ded115FAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.115F.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.115F.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded115F_4 = ccount
End Sub

Function ValidateDed115FDateTransfer() As Boolean
ValidateDed115FDateTransfer = True
setTableInfo_Ded115FDateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.115F.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransfer5F(end_Ded115F_1)
For i = 1 To end_Ded115F_1
DateOfTransfer5F(i) = rangecells.item(i).value

If Len(DateOfTransfer5F(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransfer5F(i)) Then

    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet IT." & Chr(13)
    ValidateDed115FDateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransfer5F(i), " D1h ", "Date of transfer of original foreign exchange asset  ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed115FDateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransfer5F(i)) Then

        ValidateDed115FDateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet IT." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed115FAmountInvested() As Boolean
     ValidateDed115FAmountInvested = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.115F.AmtInvt").Cells
    ReDim AmountInvested5F(end_Ded115F_1)
    For i = 1 To end_Ded115F_1
        AmountInvested5F(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountInvested5F(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed115FAmountInvested = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountInvested5F(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed115FAmountInvested = False
        Exit Function
    End If
    
    If (Len(AmountInvested5F(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Invested at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed115FAmountInvested = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed115FDateOfInvestment() As Boolean
ValidateDed115FDateOfInvestment = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.115F.DateOfInvest").Cells
Dim i As Long
ReDim DateOfPurchase5F(end_Ded115F_1)
For i = 1 To end_Ded115F_1
DateOfPurchase5F(i) = rangecells.item(i).value

If Len(DateOfPurchase5F(i)) = 0 Then
End If

If Not chkMandatory(DateOfPurchase5F(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Investment at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
    ValidateDed115FDateOfInvestment = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfPurchase5F(i)) Then
        ValidateDed115FDateOfInvestment = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Investment at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed115FAmountClaimed() As Boolean
     ValidateDed115FAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.115F.AmtDedClaimed").Cells
    ReDim AmountClaimed5F(end_Ded115F_1)
    For i = 1 To end_Ded115F_1
        AmountClaimed5F(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimed5F(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed115FAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimed5F(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed115FAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimed5F(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed115FAmountClaimed = False
        Exit Function
    End If
 Next
End Function

Function ValidateDednTable115F_CG() As Boolean
    ValidateDednTable115F_CG = True
    setTableInfo_Ded115FDateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 115F:" & Chr(13) & Chr(13)
    If Not ValidateDed115FDateTransfer() Then ValidateDednTable115F_CG = False
    If (Len(Sheet13.Range("Deduction.115F.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed115FAmountInvested() Then ValidateDednTable115F_CG = False
        If Not ValidateDed115FDateOfInvestment() Then ValidateDednTable115F_CG = False
        If Not ValidateDed115FAmountClaimed() Then ValidateDednTable115F_CG = False
    End If
    
    setTableInfo_Ded115FAmountInvested
    setTableInfo_Ded115FDatePurchase
    setTableInfo_Ded115FAmountClaimed
    
    If (end_Ded115F_1 <> end_Ded115F_2) Or (end_Ded115F_1 <> end_Ded115F_3) Or (end_Ded115F_1 <> end_Ded115F_4) Then
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 115F." & Chr(13)
             ValidateDednTable115F_CG = False
             Exit Function
    
    End If
    
End Function


Sub setTableInfo_Ded54FDateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54F.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54F.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54F_1 = ccount
    rngname_Ded54F = "Deduction.54F.DateOfTransfer;Deduction.54F.CostOfHouse;Deduction.54F.DateOfPurc;Deduction.54F.AmtdepCG;Deduction.54F.DateofDeposit;Deduction.54F.AcctNum;Deduction.54F.IFSC;Deduction.54F.AmtDedClaimed"
End Sub

Sub setTableInfo_Ded54FCostOfHouse()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54F.CostOfHouse").count
    Set rangecells = Sheet13.Range("Deduction.54F.CostOfHouse").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54F_2 = ccount
End Sub

Sub setTableInfo_Ded54FDatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54F.DateOfPurc").count
    Set rangecells = Sheet13.Range("Deduction.54F.DateOfPurc").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54F_3 = ccount
End Sub

Sub setTableInfo_Ded54FAmountDeposited()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54F.AmtdepCG").count
    Set rangecells = Sheet13.Range("Deduction.54F.AmtdepCG").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54F_4 = ccount
End Sub

Sub setTableInfo_Ded54FAmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54F.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54F.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54F_5 = ccount
End Sub

Function ValidateDed54FDateTransfer() As Boolean
ValidateDed54FDateTransfer = True
setTableInfo_Ded54ECDateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54F.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransferF(end_Ded54F_1)
For i = 1 To end_Ded54F_1
DateOfTransferF(i) = rangecells.item(i).value

If Len(DateOfTransferF(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransferF(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet IT." & Chr(13)
    ValidateDed54FDateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransferF(i), " D1e ", "Date of transfer of original asset ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54FDateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransferF(i)) Then
        ValidateDed54FDateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet IT." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54FCostOfHouse() As Boolean
     ValidateDed54FCostOfHouse = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54F.CostOfHouse").Cells
    ReDim CostOfHouseF(end_Ded54F_1)
    For i = 1 To end_Ded54F_1
        CostOfHouseF(i) = rangecells.item(i).value

'    If Not chkMandatory(CostOfHouseF(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cost Of House at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54FCostOfHouse = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CostOfHouseF(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost Of House at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54FCostOfHouse = False
        Exit Function
    End If
    
    If (Len(CostOfHouseF(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost Of House at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54FCostOfHouse = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54FDateOfPurchase() As Boolean
ValidateDed54FDateOfPurchase = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54F.DateOfPurc").Cells
Dim i As Long
ReDim DateOfPurchaseF(end_Ded54F_1)
For i = 1 To end_Ded54F_1
DateOfPurchaseF(i) = rangecells.item(i).value

If Len(DateOfPurchaseF(i)) = 0 Then
End If


'If Not chkMandatory(DateOfPurchaseF(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
'    ValidateDed54FDateOfPurchase = False
'    Exit Function
'End If

If Not CheckDateddmmyyyy(DateOfPurchaseF(i)) Then
        ValidateDed54FDateOfPurchase = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54FAmountDeposited() As Boolean
     ValidateDed54FAmountDeposited = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54F.AmtdepCG").Cells
    ReDim AmountDepositedF(end_Ded54F_1)
    For i = 1 To end_Ded54F_1
        AmountDepositedF(i) = rangecells.item(i).value
        
'     If Not chkMandatory(AmountDepositedF(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54FAmountDeposited = False
'         Exit Function
'     End If
     
     If Not IsNumeric(AmountDepositedF(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54FAmountDeposited = False
        Exit Function
    End If
    
    If (Len(AmountDepositedF(i)) > 9) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Deposited at Sr.No " & i & " cannot exceed 9 digits." & Chr(13)
        ValidateDed54FAmountDeposited = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54FAmountClaimed() As Boolean
     ValidateDed54FAmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54F.AmtDedClaimed").Cells
    ReDim AmountClaimedF(end_Ded54F_1)
    For i = 1 To end_Ded54F_1
        AmountClaimedF(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimedF(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54FAmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimedF(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54FAmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimedF(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54FAmountClaimed = False
        Exit Function
    End If
 Next
End Function

Function ValidateDednTable54F_CG() As Boolean
Dim i As Long
    ValidateDednTable54F_CG = True
    setTableInfo_Ded54FDateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54F:" & Chr(13) & Chr(13)
    If Not ValidateDed54FDateTransfer() Then ValidateDednTable54F_CG = False
    If (Len(Sheet13.Range("Deduction.54F.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54FCostOfHouse() Then ValidateDednTable54F_CG = False
        If Not ValidateDed54FDateOfPurchase() Then ValidateDednTable54F_CG = False
        If Not ValidateDed54FAmountDeposited() Then ValidateDednTable54F_CG = False
        If Not ValidateDed54FAmountClaimed() Then ValidateDednTable54F_CG = False
        
        'Newly added by Bindu as per DE 0.2
        If Not ValidateDed54FDateofDeposit() Then ValidateDednTable54F_CG = False
        If Not ValidateDedF54IFSC() Then ValidateDednTable54F_CG = False

    End If
    
    setTableInfo_Ded54FCostOfHouse
    setTableInfo_Ded54FDatePurchase
    setTableInfo_Ded54FAmountDeposited
    setTableInfo_Ded54FAmountClaimed
    
    If (end_Ded54F_1 <> end_Ded54F_5) Then '(end_Ded54F_1 <> end_Ded54F_2) Or (end_Ded54F_1 <> end_Ded54F_3) Or (end_Ded54F_1 <> end_Ded54F_4) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54F." & Chr(13)
             ValidateDednTable54F_CG = False
             Exit Function
    
    End If
    
    If end_Ded54F_1 > 0 Then

For i = 1 To end_Ded54F_1

        If CostOfHouseF(i) <> "" Or DateOfPurchaseF(i) <> "" Then

         If CostOfHouseF(i) <> "" Then ' Or Ded54DateOfPurchase(i) <> "" Then
         If DateOfPurchaseF(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54F Date Of Purchase at the row " & i & " ." & Chr(13)
            ValidateDednTable54F_CG = False
            Exit Function
         End If
         End If


         If DateOfPurchaseF(i) <> "" Then
         If CostOfHouseF(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54F Cost Of House at the row " & i & " ." & Chr(13)
            ValidateDednTable54F_CG = False
            Exit Function
         End If
         End If

        Else

         If AmountDepositedF(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill either Ded54F Date Of Purchase and Ded54F Cost Of House or Ded54F Amount Deposited at the row " & i & " ." & Chr(13)
            ValidateDednTable54F_CG = False
            Exit Function
         End If


        End If
Next

End If
    
    
    
End Function


Sub setTableInfo_Ded54DateTransfer()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54.DateOfTransfer").count
    Set rangecells = Sheet13.Range("Deduction.54.DateOfTransfer").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54_1 = ccount
    'Newly changed as per DE 0.2 by Bindu
    rngname_Ded54 = "Deduction.54.DateOfTransfer;Deduction.54.CostOfHouse;Deduction.54.DateOfPursConst;Deduction.54.AmtdepCG;Deduction.54.DateofDeposit;Deduction.54.AcctNum;Deduction.54.IFSC;Deduction.54.AmtDedClaimed;"
End Sub


Sub setTableInfo_Ded54Cost()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54.CostOfHouse").count
    Set rangecells = Sheet13.Range("Deduction.54.CostOfHouse").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54_2 = ccount
End Sub

Sub setTableInfo_Ded54DatePurchase()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54.DateOfPursConst").count
    Set rangecells = Sheet13.Range("Deduction.54.DateOfPursConst").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54_3 = ccount
End Sub

Sub setTableInfo_Ded54CashDeposited()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54.AmtdepCG").count
    Set rangecells = Sheet13.Range("Deduction.54.AmtdepCG").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54_4 = ccount
End Sub

Sub setTableInfo_Ded54AmountClaimed()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    mIntCells = Sheet13.Range("Deduction.54.AmtDedClaimed").count
    Set rangecells = Sheet13.Range("Deduction.54.AmtDedClaimed").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    end_Ded54_5 = ccount
End Sub
Function msgauddate(auddate, Tableno, DateNo) As String
    msgauddate = ""
    'auddate = Dformatexisting(auddate)
    Dim verDate As Variant
    verDate = Sheet52.Range("sheet9.Date").value
    verDate = Dformat(verDate, "yyyy-mm-dd")
    If Dformat(auddate, "yyyy-mm-dd") > verDate Then
    msgauddate = "In Schedule CG," & DateNo & " cannot be a Future Date in Table " & Tableno & " "
    Exit Function
    End If
End Function
Function ValidateDed54DateTransfer() As Boolean
ValidateDed54DateTransfer = True
setTableInfo_Ded54DateTransfer
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54.DateOfTransfer").Cells
Dim i As Long
ReDim DateOfTransfer(end_Ded54_1)
For i = 1 To end_Ded54_1
DateOfTransfer(i) = rangecells.item(i).value

If Len(DateOfTransfer(i)) = 0 Then
End If

If Not chkMandatory(DateOfTransfer(i)) Then
    msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " is Mandatory in Sheet IT." & Chr(13)
    ValidateDed54DateTransfer = False
    Exit Function
End If

Dim msg1 As Variant
msg1 = msgauddate(DateOfTransfer(i), " D1a ", " Date of transfer of original asset ")
If (msg1 <> "") Then
    msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
    ValidateDed54DateTransfer = False
    Exit Function
End If

If Not CheckDateddmmyyyy(DateOfTransfer(i)) Then
        ValidateDed54DateTransfer = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Transfer at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet IT." & Chr(13)
        Exit Function
End If
  
  
Next
End Function

Function ValidateDed54CostOfHouse() As Boolean
     ValidateDed54CostOfHouse = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54.CostOfHouse").Cells
    ReDim CostOfHouse(end_Ded54_1)
    For i = 1 To end_Ded54_1
        CostOfHouse(i) = rangecells.item(i).value
        
'     If Not chkMandatory(CostOfHouse(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cost of new residential house at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54CostOfHouse = False
'         Exit Function
'     End If
     
     If Not IsNumeric(CostOfHouse(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost of new residential house at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54CostOfHouse = False
        Exit Function
    End If
    
    If (Len(CostOfHouse(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cost of new residential house at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54CostOfHouse = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54DateOfPurchase() As Boolean
ValidateDed54DateOfPurchase = True
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54.DateOfPursConst").Cells
Dim i As Long
ReDim DateOfPurchase(end_Ded54_1)
For i = 1 To end_Ded54_1
DateOfPurchase(i) = rangecells.item(i).value

If Len(DateOfPurchase(i)) = 0 Then
End If

'If Not chkMandatory(DateOfPurchase(i)) Then
'    msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " is Mandatory in Sheet IT." & Chr(13)
'    ValidateDed54DateOfPurchase = False
'    Exit Function
'End If

If Not CheckDateddmmyyyy(DateOfPurchase(i)) Then
        ValidateDed54DateOfPurchase = False
        msgboxSHEETCG = msgboxSHEETCG + "* Date of Purchase at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet IT." & Chr(13)
        Exit Function
End If
  
Next
End Function

Function ValidateDed54CashDeposited() As Boolean
     ValidateDed54CashDeposited = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54.AmtdepCG").Cells
    ReDim CashDeposited(end_Ded54_1)
    For i = 1 To end_Ded54_1
        CashDeposited(i) = rangecells.item(i).value
        
'          If Not chkMandatory(CashDeposited(i)) Then
'         msgboxSHEETCG = msgboxSHEETCG + "* Cash Deposited at Sr. No  " & i & " is Mandatory." & Chr(13)
'         ValidateDed54CashDeposited = False
'         Exit Function
'     End If

     
     If Not IsNumeric(CashDeposited(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cash Deposited at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54CashDeposited = False
        Exit Function
    End If
    
    If (Len(CashDeposited(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Cash Deposited at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54CashDeposited = False
        Exit Function
    End If
 Next
End Function

Function ValidateDed54AmountClaimed() As Boolean
     ValidateDed54AmountClaimed = True
    Dim rangecells As Range
       Dim i As Long
    Set rangecells = Sheet13.Range("Deduction.54.AmtDedClaimed").Cells
    ReDim AmountClaimed(end_Ded54_1)
    For i = 1 To end_Ded54_1
        AmountClaimed(i) = rangecells.item(i).value
        
     If Not chkMandatory(AmountClaimed(i)) Then
         msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr. No  " & i & " is Mandatory." & Chr(13)
         ValidateDed54AmountClaimed = False
         Exit Function
     End If
     
     If Not IsNumeric(AmountClaimed(i)) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " Should be a numeric value." & Chr(13)
        ValidateDed54AmountClaimed = False
        Exit Function
    End If
    
    If (Len(AmountClaimed(i)) > 14) Then
        msgboxSHEETCG = msgboxSHEETCG + "* Amount Claimed at Sr.No " & i & " cannot exceed 14 digits." & Chr(13)
        ValidateDed54AmountClaimed = False
        Exit Function
    End If
 Next
End Function

Function ValidateDednTable54_CG() As Boolean
    Dim i As Long
    ValidateDednTable54_CG = True
    setTableInfo_Ded54DateTransfer
    
'    msgboxSHEETCG = ""
'    msgboxSHEETCG = "Deduction u/s 54:" & Chr(13) & Chr(13)
    If Not ValidateDed54DateTransfer() Then ValidateDednTable54_CG = False
    If (Len(Sheet13.Range("Deduction.54.DateOfTransfer").item(0).value) > 0) Then
        If Not ValidateDed54CostOfHouse() Then ValidateDednTable54_CG = False
        If Not ValidateDed54DateOfPurchase() Then ValidateDednTable54_CG = False
        If Not ValidateDed54CashDeposited() Then ValidateDednTable54_CG = False
        If Not ValidateDed54AmountClaimed() Then ValidateDednTable54_CG = False
        'Newly added below 2 functions by Bindu as per DE 0.2
        If Not ValidateDed54DateofDeposit() Then ValidateDednTable54_CG = False
        If Not ValidateDed54IFSC() Then ValidateDednTable54_CG = False
    End If
    
    setTableInfo_Ded54Cost
    setTableInfo_Ded54DatePurchase
    setTableInfo_Ded54CashDeposited
    setTableInfo_Ded54AmountClaimed
    
    If (end_Ded54_1 <> end_Ded54_5) Then '(end_Ded54_1 <> end_Ded54_2) Or (end_Ded54_1 <> end_Ded54_3) Or (end_Ded54_1 <> end_Ded54_4) Or
        
             msgboxSHEETCG = msgboxSHEETCG + "* Enter All Mandatory Fields in Deduction table 54." & Chr(13)
             ValidateDednTable54_CG = False
             Exit Function
    
    End If
    
If end_Ded54_1 > 0 Then

For i = 1 To end_Ded54_1

        If CostOfHouse(i) <> "" Or DateOfPurchase(i) <> "" Then

         If CostOfHouse(i) <> "" Then ' Or Ded54DateOfPurchase(i) <> "" Then
         If DateOfPurchase(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54 Date Of Purchase at the row " & i & " ." & Chr(13)
            ValidateDednTable54_CG = False
            Exit Function
         End If
         End If


         If DateOfPurchase(i) <> "" Then
         If CostOfHouse(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill Ded54 Cost Of House at the row " & i & " ." & Chr(13)
            ValidateDednTable54_CG = False
            Exit Function
         End If
         End If

        Else

         If CashDeposited(i) = "" Then
            msgboxSHEETCG = msgboxSHEETCG & "Please fill either Ded54 Date Of Purchase and Ded54 Cost Of House or Ded54 Cash Deposited at the row " & i & " ." & Chr(13)
            ValidateDednTable54_CG = False
            Exit Function
         End If


        End If
Next

End If
    
    
End Function

Function ValidateDed54DateofDeposit() As Boolean
ValidateDed54DateofDeposit = True
'setTableInfo_Ded54DateofDeposit
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54.DateofDeposit").Cells
Dim i As Long
ReDim dateOfdeposit(end_Ded54_1)

Dim rangecells1 As Range
    Set rangecells1 = Sheet13.Range("Deduction.54.AmtdepCG").Cells
    ReDim CashDeposited(end_Ded54_1)

For i = 1 To end_Ded54_1
    dateOfdeposit(i) = rangecells.item(i).value

    CashDeposited(i) = rangecells1.item(i).value

    If Len(CashDeposited(i)) > 0 And CashDeposited(i) <> 0 Then
        If Not chkMandatory(dateOfdeposit(i)) Then
         'Added by Aavula
            msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Date of deposit"" at Sr. No " & i & "." & Chr(13)
            ValidateDed54DateofDeposit = False
            Exit Function
        End If
    End If
    
    Dim msg1 As Variant
    msg1 = msgauddate(dateOfdeposit(i), " D1a ", " Date of Deposit")
    If (msg1 <> "") Then
        msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
        ValidateDed54DateofDeposit = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(dateOfdeposit(i)) Then
            ValidateDed54DateofDeposit = False
            msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
            Exit Function
    End If
  
  
Next
End Function
Function ValidateDed54IFSC() As Boolean
    ValidateDed54IFSC = True

    Dim rangecells, rangecells1, rangecells2 As Range
    Set rangecells = Range("Deduction.54.IFSC").Cells
    Set rangecells1 = Range("Deduction.54.AmtdepCG").Cells
    Set rangecells2 = Range("Deduction.54.AcctNum").Cells
    
    
    Dim i As Long
    ReDim BankIFSC(end_Ded54_1)
    ReDim CashDeposited(end_Ded54_1)
    ReDim BankAccntnum(end_Ded54_1)
    
   
    
    For i = 1 To end_Ded54_1
        BankIFSC(i) = rangecells.item(i).value
        CashDeposited(i) = rangecells1.item(i).value
        BankAccntnum(i) = rangecells2.item(i).value
        
     If Len(CashDeposited(i)) > 0 And CashDeposited(i) <> 0 Then
       If isdropdownblank(BankIFSC(i)) Then
       
        'Added by Aavula
         msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory"" at Sr.No " & i & ". " & Chr(13)
         ValidateDed54IFSC = False
         'Exit Function
       End If
    End If
    
    If Len(CashDeposited(i)) > 0 And CashDeposited(i) <> 0 Then
            If Not chkCompulsory(BankAccntnum(i)) Then
                msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Account Number"" at Sr.No " & i & " " & Chr(13)
                ValidateDed54IFSC = False
                'Exit Function
           End If
        End If
           
           If Not ValidateBankAccountNumber_CG(BankAccntnum(i), i) Then
            ValidateDed54IFSC = False
            'Exit Function
        End If
    
        
        If Len(BankIFSC(i)) > 11 Then
            'msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets. at Sr.No " & i & " in Sheet CG" & Chr(13)
             
             'Added by Aavula
             msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets at Sr.No " & i & "" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code at Sr.No " & i & " in CG cannot exceed 11 characters" & Chr(13)
            ValidateDed54IFSC = False
            'Exit Function
        End If
        
         'Added by Aavula
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSC(i))) Then
            msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFS code at Sr.No " & i & "" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateDed54IFSC = False
            'Exit Function
        End If
        
    Next
End Function
Function ValidateBankAccountNumber_CG(BankAccountNumber As Variant, CC As Long) As Boolean
    ValidateBankAccountNumber_CG = True
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
            msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & CC & " is invalid in sheet CG, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_CG = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & CC & " is invalid in CG " & Chr(13)
            ValidateBankAccountNumber_CG = False
            Exit Function
        End If
    
    
    
    
'    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
'        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & cc & " is mandatory in CG" & Chr(13)
'        ValidateBankAccountNumber_CG = False
'        Exit Function
'    End If
'
'    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
'        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & cc & " is mandatory in CG" & Chr(13)
'        ValidateBankAccountNumber_CG = False
'        Exit Function
'    End If

'    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
'        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & cc & " is invalid in CG  " & Chr(13)
'        ValidateBankAccountNumber_CG = False
'        Exit Function
'    End If

'    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
'        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & cc & " is invalid in CG" & Chr(13)
'        ValidateBankAccountNumber_CG = False
'        Exit Function
'    End If
'

'    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
'        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & cc & " is invalid in CG " & Chr(13)
'        ValidateBankAccountNumber_CG = False
'        Exit Function
'    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & CC & " is invalid in CG" & Chr(13)
        ValidateBankAccountNumber_CG = False
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
            msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & CC & " is invalid in CG." & Chr(13)
            ValidateBankAccountNumber_CG = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        msgboxSHEETCG = msgboxSHEETCG & "* Bank Account Number at Sr.No " & CC & " in CG is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_CG = False
        Exit Function
    End If
    
End If
End Function

Function ValidateDed54BDateofDeposit() As Boolean
ValidateDed54BDateofDeposit = True
'setTableInfo_Ded54DateofDepositB
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54B.DateofDeposit").Cells
Dim i As Long
ReDim DateofDepositB(end_Ded54B_1)

Dim rangecells1 As Range
    Set rangecells1 = Sheet13.Range("Deduction.54B.AmtdepCG").Cells
    ReDim CashDepositedB(end_Ded54B_1)

For i = 1 To end_Ded54B_1
    DateofDepositB(i) = rangecells.item(i).value

    CashDepositedB(i) = rangecells1.item(i).value

    If Len(CashDepositedB(i)) > 0 And CashDepositedB(i) <> 0 Then
        If Not chkMandatory(DateofDepositB(i)) Then
        
           'Added by Aavula
            msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Date of deposit"" at Sr. No " & i & " ." & Chr(13)
            ValidateDed54BDateofDeposit = False
            Exit Function
        End If
    End If
    
    Dim msg1 As Variant
    msg1 = msgauddate(DateofDepositB(i), " D1B ", " Date of Deposit")
    If (msg1 <> "") Then
        msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
        ValidateDed54BDateofDeposit = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(DateofDepositB(i)) Then
            ValidateDed54BDateofDeposit = False
            msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
            Exit Function
    End If
  
  
Next
End Function
Function ValidateDedB54IFSC() As Boolean
    ValidateDedB54IFSC = True

    Dim rangecells, rangecells1, rangecells2 As Range
    Set rangecells = Range("Deduction.54B.IFSC").Cells
    Set rangecells1 = Range("Deduction.54B.AmtdepCG").Cells
    Set rangecells2 = Range("Deduction.54B.AcctNum").Cells
    
    
    Dim i As Long
    ReDim BankIFSCB(end_Ded54B_1)
    ReDim CashDepositedB(end_Ded54B_1)
    ReDim BankAccntnumB(end_Ded54B_1)
    
   
    
    For i = 1 To end_Ded54B_1
        BankIFSCB(i) = rangecells.item(i).value
        CashDepositedB(i) = rangecells1.item(i).value
        BankAccntnumB(i) = rangecells2.item(i).value
        
     If Len(CashDepositedB(i)) > 0 And CashDepositedB(i) <> 0 Then
       If isdropdownblank(BankIFSCB(i)) Then
       
          'Added by Aavula
          msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory"" at Sr.No " & i & " ." & Chr(13)

       
         'msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory in Pt 14 i.e. bank account details in Schedule Part B TTI in point 18."" at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         ValidateDedB54IFSC = False
         'Exit Function
       End If
    End If
    
    If Len(CashDepositedB(i)) > 0 And CashDepositedB(i) <> 0 Then
            If Not chkCompulsory(BankAccntnumB(i)) Then
                msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Account Number"" at Sr.No " & i & " " & Chr(13)

                'msgboxSHEETCG = msgboxSHEETCG + "* Account Number at Sr.No " & i & " in CG is mandatory" & Chr(13)
                ValidateDedB54IFSC = False
                'Exit Function
           End If
        End If
           
           If Not ValidateBankAccountNumber_CG(BankAccntnumB(i), i) Then
            ValidateDedB54IFSC = False
            'Exit Function
        End If
        
        If Len(BankIFSCB(i)) > 11 Then
        
            msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets. at Sr.No " & i & " " & Chr(13)

            'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code at Sr.No " & i & " in CG cannot exceed 11 characters" & Chr(13)
            ValidateDedB54IFSC = False
            'Exit Function
        End If
        
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSCB(i))) Then
            
           msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFS code at Sr.No " & i & " " & Chr(13)

            'msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFSC in Schedule Part B TTI in point 18 at Sr.No " & i & " in Sheet CG" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateDedB54IFSC = False
            'Exit Function
        End If
            
        
        
       
    Next
End Function
Function ValidateDed54DDateofDeposit() As Boolean
ValidateDed54DDateofDeposit = True
'setTableInfo_Ded54DateofDepositD
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54D.DateofDeposit").Cells
Dim i As Long
ReDim DateofDepositD(end_Ded54D_1)

Dim rangecells1 As Range
    Set rangecells1 = Sheet13.Range("Deduction.54D.AmtdepCG").Cells
    ReDim CashDepositedD(end_Ded54D_1)

For i = 1 To end_Ded54D_1
    DateofDepositD(i) = rangecells.item(i).value

    CashDepositedD(i) = rangecells1.item(i).value

    If Len(CashDepositedD(i)) > 0 And CashDepositedD(i) <> 0 Then
        If Not chkMandatory(DateofDepositD(i)) Then
            'msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
             'Added by Aavula
           msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Date of deposit"" at Sr. No " & i & "." & Chr(13)

            ValidateDed54DDateofDeposit = False
            Exit Function
        End If
    End If
    
    Dim msg1 As Variant
    msg1 = msgauddate(DateofDepositD(i), " D1c ", " Date of Deposit")
    If (msg1 <> "") Then
        msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
        ValidateDed54DDateofDeposit = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(DateofDepositD(i)) Then
            ValidateDed54DDateofDeposit = False
            msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
            Exit Function
    End If
  
  
Next
End Function
Function ValidateDedD54IFSC() As Boolean
    ValidateDedD54IFSC = True

    Dim rangecells, rangecells1, rangecells2 As Range
    Set rangecells = Range("Deduction.54D.IFSC").Cells
    Set rangecells1 = Range("Deduction.54D.AmtdepCG").Cells
    Set rangecells2 = Range("Deduction.54D.AcctNum").Cells
    
    
    Dim i As Long
    ReDim BankIFSCD(end_Ded54D_1)
    ReDim CashDepositedD(end_Ded54D_1)
    ReDim BankAccntnumD(end_Ded54D_1)
    
   
    
    For i = 1 To end_Ded54D_1
        BankIFSCD(i) = rangecells.item(i).value
        CashDepositedD(i) = rangecells1.item(i).value
        BankAccntnumD(i) = rangecells2.item(i).value
        
     If Len(CashDepositedD(i)) > 0 And CashDepositedD(i) <> 0 Then
       If isdropdownblank(BankIFSCD(i)) Then
       
        'Added by Aavula
          msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory "" at Sr.No " & i & "." & Chr(13)

         'msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory in Pt 14 i.e. bank account details in Schedule Part B TTI in point 18."" at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         ValidateDedD54IFSC = False
         'Exit Function
       End If
    End If
    
    If Len(CashDepositedD(i)) > 0 And CashDepositedD(i) <> 0 Then
            If Not chkCompulsory(BankAccntnumD(i)) Then
                msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Account Number"" at Sr.No " & i & "" & Chr(13)

                'msgboxSHEETCG = msgboxSHEETCG + "* Account Number at Sr.No " & i & " in CG is mandatory" & Chr(13)
                ValidateDedD54IFSC = False
                'Exit Function
           End If
        End If
           
           If Not ValidateBankAccountNumber_CG(BankAccntnumD(i), i) Then
            ValidateDedD54IFSC = False
            'Exit Function
        End If
        
        If Len(BankIFSCD(i)) > 11 Then
            msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets. at Sr.No " & i & "" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code at Sr.No " & i & " in CG cannot exceed 11 characters" & Chr(13)
            ValidateDedD54IFSC = False
            'Exit Function
        End If
        
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSCD(i))) Then
            
            'msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFSC in Schedule Part B TTI in point 18 at Sr.No " & i & " in Sheet CG" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
             'Added by Aavula
            msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFS code at Sr.No " & i & "" & Chr(13)

            ValidateDedD54IFSC = False
            'Exit Function
        End If
            
        
        
       
    Next
End Function
Function ValidateDed54FDateofDeposit() As Boolean
ValidateDed54FDateofDeposit = True
'setTableInfo_Ded54DateofDepositF
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54F.DateofDeposit").Cells
Dim i As Long
ReDim DateofDepositF(end_Ded54F_1)
 
Dim rangecells1 As Range
    Set rangecells1 = Sheet13.Range("Deduction.54F.AmtdepCG").Cells
    ReDim CashDepositedF(end_Ded54F_1)
 
For i = 1 To end_Ded54F_1
    DateofDepositF(i) = rangecells.item(i).value
 
    CashDepositedF(i) = rangecells1.item(i).value
 
    If Len(CashDepositedF(i)) > 0 And CashDepositedF(i) <> 0 Then
        If Not chkMandatory(DateofDepositF(i)) Then
           ' msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
           
             'Added by Aavula
            msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Date of deposit"" at Sr. No " & i & "." & Chr(13)

            ValidateDed54FDateofDeposit = False
            Exit Function
        End If
    End If
    Dim msg1 As Variant
    msg1 = msgauddate(DateofDepositF(i), " D1e ", " Date of Deposit")
    If (msg1 <> "") Then
        msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
        ValidateDed54FDateofDeposit = False
        Exit Function
    End If
    If Not CheckDateddmmyyyy(DateofDepositF(i)) Then
            ValidateDed54FDateofDeposit = False
            msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
            Exit Function
    End If

Next
End Function
Function ValidateDedF54IFSC() As Boolean
    ValidateDedF54IFSC = True
 
    Dim rangecells, rangecells1, rangecells2 As Range
    Set rangecells = Range("Deduction.54F.IFSC").Cells
    Set rangecells1 = Range("Deduction.54F.AmtdepCG").Cells
    Set rangecells2 = Range("Deduction.54F.AcctNum").Cells

    Dim i As Long
    ReDim BankIFSCF(end_Ded54F_1)
    ReDim CashDepositedF(end_Ded54F_1)
    ReDim BankAccntnumF(end_Ded54F_1)

    For i = 1 To end_Ded54F_1
        BankIFSCF(i) = rangecells.item(i).value
        CashDepositedF(i) = rangecells1.item(i).value
        BankAccntnumF(i) = rangecells2.item(i).value
     If Len(CashDepositedF(i)) > 0 And CashDepositedF(i) <> 0 Then
       If isdropdownblank(BankIFSCF(i)) Then
       
       'Added by Aavula
          msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory"" at Sr.No " & i & "." & Chr(13)
          
         'msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory in Pt 14 i.e. bank account details in Schedule Part B TTI in point 18."" at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         ValidateDedF54IFSC = False
         'Exit Function
       End If
    End If
    If Len(CashDepositedF(i)) > 0 And CashDepositedF(i) <> 0 Then
            If Not chkCompulsory(BankAccntnumF(i)) Then
                msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Account Number"" at Sr.No " & i & "" & Chr(13)
 
                'msgboxSHEETCG = msgboxSHEETCG + "* Account Number at Sr.No " & i & " in CG is mandatory" & Chr(13)
                ValidateDedF54IFSC = False
                'Exit Function
           End If
        End If
           If Not ValidateBankAccountNumber_CG(BankAccntnumF(i), i) Then
            ValidateDedF54IFSC = False
            'Exit Function
        End If
        If Len(BankIFSCF(i)) > 11 Then
            msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets. at Sr.No " & i & "" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code at Sr.No " & i & " in CG cannot exceed 11 characters" & Chr(13)
            ValidateDedF54IFSC = False
            'Exit Function
        End If
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSCF(i))) Then
        
         'Added by Aavula
            msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFS code at Sr.No " & i & "" & Chr(13)

            'msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFSC in Schedule Part B TTI in point 18 at Sr.No " & i & " in Sheet CG" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateDedF54IFSC = False
            'Exit Function
        End If


    Next
End Function
Function ValidateDed54GDateofDeposit() As Boolean
ValidateDed54GDateofDeposit = True
'setTableInfo_Ded54DateofDepositG
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54G.DateofDeposit").Cells
Dim i As Long
ReDim DateofDepositG(end_Ded54G_1)
 
Dim rangecells1 As Range
    Set rangecells1 = Sheet13.Range("Deduction.54G.AmtdepCG").Cells
    ReDim CashDepositedG(end_Ded54G_1)
 
For i = 1 To end_Ded54G_1
    DateofDepositG(i) = rangecells.item(i).value
 
    CashDepositedG(i) = rangecells1.item(i).value
 
    If Len(CashDepositedG(i)) > 0 And CashDepositedG(i) <> 0 Then
        If Not chkMandatory(DateofDepositG(i)) Then
            msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Date of deposit"" at Sr. No " & i & "." & Chr(13)
            ValidateDed54GDateofDeposit = False
            Exit Function
        End If
    End If
    Dim msg1 As Variant
    msg1 = msgauddate(DateofDepositG(i), " D1f ", " Date of Deposit")
    If (msg1 <> "") Then
        msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
        ValidateDed54GDateofDeposit = False
        Exit Function
    End If
    If Not CheckDateddmmyyyy(DateofDepositG(i)) Then
            ValidateDed54GDateofDeposit = False
            msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
            Exit Function
    End If

Next
End Function
Function ValidateDedG54IFSC() As Boolean
    ValidateDedG54IFSC = True
 
    Dim rangecells, rangecells1, rangecells2 As Range
    Set rangecells = Range("Deduction.54G.IFSC").Cells
    Set rangecells1 = Range("Deduction.54G.AmtdepCG").Cells
    Set rangecells2 = Range("Deduction.54G.AcctNum").Cells

    Dim i As Long
    ReDim BankIFSCG(end_Ded54G_1)
    ReDim CashDepositedG(end_Ded54G_1)
    ReDim BankAccntnumG(end_Ded54G_1)

    For i = 1 To end_Ded54G_1
        BankIFSCG(i) = rangecells.item(i).value
        CashDepositedG(i) = rangecells1.item(i).value
        BankAccntnumG(i) = rangecells2.item(i).value
     If Len(CashDepositedG(i)) > 0 And CashDepositedG(i) <> 0 Then
       If isdropdownblank(BankIFSCG(i)) Then
       
        'Added by Aavula
         msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory"" at Sr.No " & i & "." & Chr(13)

         'msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory in Pt 14 i.e. bank account details in Schedule Part B TTI in point 18."" at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         ValidateDedG54IFSC = False
         'Exit Function
       End If
    End If
    If Len(CashDepositedG(i)) > 0 And CashDepositedG(i) <> 0 Then
            If Not chkCompulsory(BankAccntnumG(i)) Then
                msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Account Number"" at Sr.No " & i & " " & Chr(13)
 
                'msgboxSHEETCG = msgboxSHEETCG + "* Account Number at Sr.No " & i & " in CG is mandatory" & Chr(13)
                ValidateDedG54IFSC = False
                'Exit Function
           End If
        End If
           If Not ValidateBankAccountNumber_CG(BankAccntnumG(i), i) Then
            ValidateDedG54IFSC = False
            'Exit Function
        End If
        If Len(BankIFSCG(i)) > 11 Then
            msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets. at Sr.No " & i & " " & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code at Sr.No " & i & " in CG cannot exceed 11 characters" & Chr(13)
            ValidateDedG54IFSC = False
            'Exit Function
        End If
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSCG(i))) Then
         
          'Added by Aavula
         msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFS code at Sr.No " & i & " " & Chr(13)
         
            'msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFSC in Schedule Part B TTI in point 18 at Sr.No " & i & " in Sheet CG" & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateDedG54IFSC = False
            'Exit Function
        End If


    Next
End Function

Function ValidateDed54GADateofDeposit() As Boolean
ValidateDed54GADateofDeposit = True
'setTableInfo_Ded54DateofDepositGA
Dim rangecells As Range
Set rangecells = Sheet13.Range("Deduction.54GA.DateofDeposit").Cells
Dim i As Long
ReDim DateofDepositGA(end_Ded54GA_1)
 
Dim rangecells1 As Range
    Set rangecells1 = Sheet13.Range("Deduction.54GA.AmtdepCG").Cells
    ReDim CashDepositedGA(end_Ded54GA_1)
 
For i = 1 To end_Ded54GA_1
    DateofDepositGA(i) = rangecells.item(i).value
 
    CashDepositedGA(i) = rangecells1.item(i).value
 
    If Len(CashDepositedGA(i)) > 0 And CashDepositedGA(i) <> 0 Then
        If Not chkMandatory(DateofDepositGA(i)) Then
        
           'Added by Aavula
           msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Date of deposit"" at Sr. No " & i & " ." & Chr(13)
         
            'msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " is Mandatory in Sheet CG." & Chr(13)
            ValidateDed54GADateofDeposit = False
            Exit Function
        End If
    End If
    Dim msg1 As Variant
    msg1 = msgauddate(DateofDepositGA(i), " D1g ", " Date of Deposit")
    If (msg1 <> "") Then
        msgboxSHEETCG = msgboxSHEETCG + msg1 & Chr(13)
        ValidateDed54GADateofDeposit = False
        Exit Function
    End If
    If Not CheckDateddmmyyyy(DateofDepositGA(i)) Then
            ValidateDed54GADateofDeposit = False
            msgboxSHEETCG = msgboxSHEETCG + "* Date of Deposit at Sr. No " & i & " must be a valid dd/mm/yyyy format in Sheet CG." & Chr(13)
            Exit Function
    End If

Next
End Function
Function ValidateDedGA54IFSC() As Boolean
    ValidateDedGA54IFSC = True
 
    Dim rangecells, rangecells1, rangecells2 As Range
    Set rangecells = Range("Deduction.54GA.IFSC").Cells
    Set rangecells1 = Range("Deduction.54GA.AmtdepCG").Cells
    Set rangecells2 = Range("Deduction.54GA.AcctNum").Cells

    Dim i As Long
    ReDim BankIFSCGA(end_Ded54GA_1)
    ReDim CashDepositedGA(end_Ded54GA_1)
    ReDim BankAccntnumGA(end_Ded54GA_1)

    For i = 1 To end_Ded54GA_1
        BankIFSCGA(i) = rangecells.item(i).value
        CashDepositedGA(i) = rangecells1.item(i).value
        BankAccntnumGA(i) = rangecells2.item(i).value
     If Len(CashDepositedGA(i)) > 0 And CashDepositedGA(i) <> 0 Then
       If isdropdownblank(BankIFSCGA(i)) Then
       
         'Added by Aavula
         msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory."" at Sr.No " & i & " ." & Chr(13)

         'msgboxSHEETCG = msgboxSHEETCG + "* ""IFSC code is mandatory in Pt 14 i.e. bank account details in Schedule Part B TTI in point 18."" at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet CG is mandatory." & Chr(13)
         ValidateDedGA54IFSC = False
         'Exit Function
       End If
    End If
    If Len(CashDepositedGA(i)) > 0 And CashDepositedGA(i) <> 0 Then
            If Not chkCompulsory(BankAccntnumGA(i)) Then
                msgboxSHEETCG = msgboxSHEETCG + "* ""Please enter the Account Number"" at Sr.No " & i & " " & Chr(13)
 
                'msgboxSHEETCG = msgboxSHEETCG + "* Account Number at Sr.No " & i & " in CG is mandatory" & Chr(13)
                ValidateDedGA54IFSC = False
                'Exit Function
           End If
        End If
           If Not ValidateBankAccountNumber_CG(BankAccntnumGA(i), i) Then
            ValidateDedGA54IFSC = False
            'Exit Function
        End If
        If Len(BankIFSCGA(i)) > 11 Then
            msgboxSHEETCG = msgboxSHEETCG + "*IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets. at Sr.No " & i & " " & Chr(13)
            'msgboxSHEETCG = msgboxSHEETCG + "* IFS Code at Sr.No " & i & " in CG cannot exceed 11 characters" & Chr(13)
            ValidateDedGA54IFSC = False
            'Exit Function
        End If
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSCGA(i))) Then
        
             'Added by Aavula
            msgboxSHEETCG = msgboxSHEETCG + "* Please enter a valid IFS code at Sr.No " & i & " " & Chr(13)

            'msgboxSHEETCG = msgboxSHEETCG + "*Please enter a valid IFSC in Schedule Part B TTI in point 18 at Sr.No " & i & " in Sheet CG" & Chr(13)
 
            'msgboxSHEETCG = msgboxSHEETCG + "* Invalid IFSC Code at Sr.No " & i & " Refer to your bank for valid ifsc codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateDedGA54IFSC = False
            'Exit Function
        End If


    Next
End Function
