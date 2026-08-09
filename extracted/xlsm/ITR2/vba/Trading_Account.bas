Attribute VB_Name = "Trading_Account"
Option Explicit
Public TradingNature1, TradingAmount1, TradingNature2, TradingAmount2, MsgBx_TradingNature1, MsgBx_TradingNature2, MsgBx_TradingNature3 As Variant
Public ColCount_TradingAmount1, ColCount_TradingNature1, ColCount_TradingAmount2, ColCount_TradingNature2 As Long


Sub AddRows_TradingAcc1()
    Sheets("Trading Account").Activate
    EfilingCommon.DefinedgridNameRange = "TradingAcc_RevenueNature||TradingAcc_RevenueAmount"
    ActiveCellRange = EfilingCommon.searchLastRow("TradingAcc_RevenueNature")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows_TradingAcc2()
    Sheets("Trading Account").Activate
    EfilingCommon.DefinedgridNameRange = "TradingAcc_NatureOfDirectExpenses||TradingAcc_OtherDirectAmount"
    ActiveCellRange = EfilingCommon.searchLastRow("TradingAcc_NatureOfDirectExpenses")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub CmdValidate_TradingAccount_Click()
    ValidateSheet_Trading
    fmsgboxStatus "Sheet Trading Account is ok"
CloseMsg
End Sub
Function ValidateSheet_Trading()
    If Not ValidateNatureAmt1 Then
    Sheet49.Activate
    fmsgboxsmall (MsgBx_TradingNature1)
    CloseMsg
    End If
    
    If Not ValidateNatureAmt2 Then
    Sheet49.Activate
    fmsgboxsmall (MsgBx_TradingNature2)
    CloseMsg
    End If
    
    If Not Validate14DigitsCheck Then
    Sheet49.Activate
    fmsgboxsmall (MsgBx_TradingNature3)
    CloseMsg
    End If
    
    
End Function
Function Validate14DigitsCheck() As Boolean
Validate14DigitsCheck = True
'MsgBox Sheet49.Range("TradingAcc_RevenueTotal")
'If Len(Sheet49.Range("TradingAcc_RevenueTotal").value) > 14 Then
'MsgBx_TradingNature3 = "*Revenue Total should not exceed 14 digits in Trading Account" & Chr(13)
'Validate14DigitsCheck = False
'Exit Function
'End If

If Len(Sheet49.Range("TradingAcc_TotalOfSalesOrGross").value) > 14 Then
MsgBx_TradingNature3 = "*Total Of Sales Or Gross should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If

If Len(Sheet49.Range("TradingAcc_TotalDutiesTaxesCess").value) > 14 Then
MsgBx_TradingNature3 = "*Total Duties Taxes Cess should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If

If Len(Sheet49.Range("TradingAcc_TotalRevenueFromOperations").value) > 14 Then
MsgBx_TradingNature3 = "*Total Revenue From Operations should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If

If Len(Sheet49.Range("TradingAcc_TotalOfCreditsToTradingAccount").value) > 14 Then
MsgBx_TradingNature3 = "*Total Of Credits To Trading Account should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If

If Len(Sheet49.Range("TradingAcc_TotalDirectExpenses").value) > 14 Then
MsgBx_TradingNature3 = "*Total Direct Expenses should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If

If Len(Sheet49.Range("TradingAcc_TotalDuties").value) > 14 Then
MsgBx_TradingNature3 = "*Total Duties should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If

If Sheet49.Range("TradingAcc_CostOfGoods").value > 0 Then
If Len(Sheet49.Range("TradingAcc_CostOfGoods").value) > 14 Then
MsgBx_TradingNature3 = "*Cost Of Goods should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If
End If


If Sheet49.Range("TradingAcc_CostOfGoods").value < 0 Then
If Len(Sheet49.Range("TradingAcc_CostOfGoods").value) > 15 Then
MsgBx_TradingNature3 = "*Cost Of Goods should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If
End If


If Sheet49.Range("TradingAcc_GrossProfitOrLoss").value > 0 Then
If Len(Sheet49.Range("TradingAcc_GrossProfitOrLoss").value) > 14 Then
MsgBx_TradingNature3 = "*Gross Profit Or Loss should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If
End If

If Sheet49.Range("TradingAcc_GrossProfitOrLoss").value < 0 Then
If Len(Sheet49.Range("TradingAcc_GrossProfitOrLoss").value) > 15 Then
MsgBx_TradingNature3 = "*Gross Profit Or Loss should not exceed 14 digits in Trading Account" & Chr(13)
Validate14DigitsCheck = False
Exit Function
End If
End If

End Function
Function ValidateNatureAmt1() As Boolean
ValidateNatureAmt1 = True
    countfilledCol_TradingAmt1
    countfilledCol_TradingNature1
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet49.Range("TradingAcc_RevenueNature").Cells
    Set rangecells1 = Sheet49.Range("TradingAcc_RevenueAmount").Cells
    Dim i As Long
    ReDim TradingNature1(ColCount_TradingNature1)
    ReDim TradingAmount1(ColCount_TradingNature1)
    MsgBx_TradingNature1 = ""

        
For i = 1 To ColCount_TradingNature1
TradingNature1(i) = rangecells.item(i).value
TradingAmount1(i) = rangecells1.item(i).value
If Not chkCompulsory(TradingAmount1(i)) Then
            MsgBx_TradingNature1 = "* Please enter the amount in other operating revenue at Sr. No  " & i & "  in Trading Account" & Chr(13)
            ValidateNatureAmt1 = False
            Exit Function
        End If
        
If Not chkCompulsory(TradingNature1(i)) Then
            MsgBx_TradingNature1 = "* Please specify the nature in other operating revenue in Trading Accoun at Sr. No  " & i & " in Trading Account" & Chr(13)
            ValidateNatureAmt1 = False
            Exit Function
        End If

Next

If (ColCount_TradingNature1 <> ColCount_TradingAmount1) Then
            MsgBx_TradingNature1 = "* Enter all mandatory Fields in Table Other Operating Revenues Sheet Trading Account" & Chr(13)
            ValidateNatureAmt1 = False
            Exit Function
End If

End Function

Sub countfilledCol_TradingAmt1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet49.Range("TradingAcc_RevenueAmount").count
    Set rangecells = Sheet49.Range("TradingAcc_RevenueAmount").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_TradingAmount1 = ccount
End Sub


Sub countfilledCol_TradingNature1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet49.Range("TradingAcc_RevenueNature").count
    Set rangecells = Sheet49.Range("TradingAcc_RevenueNature").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_TradingNature1 = ccount
End Sub

Function ValidateNatureAmt2() As Boolean
ValidateNatureAmt2 = True
    countfilledCol_TradingAmt2
    countfilledCol_TradingNature2
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet49.Range("TradingAcc_NatureOfDirectExpenses").Cells
    Set rangecells1 = Sheet49.Range("TradingAcc_OtherDirectAmount").Cells
    Dim i As Long
    ReDim TradingNature2(ColCount_TradingNature2)
    ReDim TradingAmount2(ColCount_TradingNature2)
    MsgBx_TradingNature2 = ""
For i = 1 To ColCount_TradingNature2
TradingNature2(i) = rangecells.item(i).value
TradingAmount2(i) = rangecells1.item(i).value
            If Not chkCompulsory(TradingAmount2(i)) Then
                        MsgBx_TradingNature2 = "* Please enter amount of other direct expenses at Sr. No  " & i & "  in Trading Account" & Chr(13)
                        ValidateNatureAmt2 = False
                        Exit Function
            End If
                    
            If Not chkCompulsory(TradingNature2(i)) Then
                        MsgBx_TradingNature2 = "* Please enter nature of direct expenses  at Sr. No  " & i & "  in Trading Account" & Chr(13)
                        ValidateNatureAmt2 = False
                        Exit Function
                End If
            
            Next

            If (ColCount_TradingNature2 <> ColCount_TradingAmount2) Then
                        MsgBx_TradingNature2 = "* Enter all mandatory Fields in Table Other Direct Expenses in Trading Account  " & Chr(13)
                        ValidateNatureAmt2 = False
                        Exit Function
            End If
            
            If Len(Sheet49.Range("TD.OthDE").value) > 14 Then
                MsgBx_TradingNature2 = "* Total at point 9 cannot be more than 14 digits in Trading Account  " & Chr(13)
                        ValidateNatureAmt2 = False
                        Exit Function
            End If

End Function

Sub countfilledCol_TradingAmt2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet49.Range("TradingAcc_OtherDirectAmount").count
    Set rangecells = Sheet49.Range("TradingAcc_OtherDirectAmount").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_TradingAmount2 = ccount
End Sub


Sub countfilledCol_TradingNature2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet49.Range("TradingAcc_NatureOfDirectExpenses").count
    Set rangecells = Sheet49.Range("TradingAcc_NatureOfDirectExpenses").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_TradingNature2 = ccount
End Sub

