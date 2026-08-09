Attribute VB_Name = "Nature_of_business"
Option Explicit

Public MsgBx_NOB As String

Public end_NOBCode As Long

'Set Values
Public Code_NOB, TradeName_NOB, TradeName1_NOB, TradeName2_NOB As Variant

Sub CmdValidate_NOB_Click()
ValidateSheetNOB_ALL
fmsgboxStatus "Sheet Nature of Business is OK"
CloseMsg
End Sub

Sub ValidateSheetNOB_ALL()
    If Not ValidateNOBGrid Then
        Sheet2.Activate
        fmsgboxsmall (MsgBx_NOB) ', vbOKOnly, "Error(s!)"
        CloseMsg
    End If
End Sub


Function ValidateNOBGrid() As Boolean
ValidateNOBGrid = True
    MsgBx_NOB = ""
    MsgBx_NOB = "* Nature of Business" & Chr(10)
    subProcCaption = "Validating NOB"
    If Not ValidateCode_NOB Then ValidateNOBGrid = False
    If Not ValidateTradeName_NOB Then ValidateNOBGrid = False
    If Not ValidateTradeName1_NOB Then ValidateNOBGrid = False
    If Not ValidateTradeName2_NOB Then ValidateNOBGrid = False
    

End Function

Function ValidateCode_NOB() As Boolean
ValidateCode_NOB = True
    setTableCode_NOB
    Dim rangecells As Range
    Set rangecells = Sheet2.Range("NOB.Code").Cells
    Dim i As Long

    ReDim Code_NOB(end_NOBCode)
    noOfProcessSub = end_NOBCode
    For i = 1 To end_NOBCode
         If InStr(rangecells.item(i).value, "Other") <> 0 Then
            If Range("NOB.TradeNameOrDescription2").item(i).value = "" Then
                 MsgBx_NOB = MsgBx_NOB + "* Description is compulsory: schedule Nature of business"
                 ValidateCode_NOB = False
                 
                 Exit Function
            End If
          End If
        Code_NOB(i) = rangecells.item(i).value
        UpdateProgressBar
    Next
'End If
End Function

Function ValidateTradeName_NOB() As Boolean
ValidateTradeName_NOB = True
    Dim rangecells As Range
    Set rangecells = Sheet2.Range("NOB.TradeNameOrDescription").Cells
    Dim i As Long
    ReDim TradeName_NOB(end_NOBCode)
    For i = 1 To end_NOBCode
    TradeName_NOB(i) = rangecells.item(i).value
Next
End Function

Function ValidateTradeName1_NOB() As Boolean
ValidateTradeName1_NOB = True
    Dim rangecells As Range
    Set rangecells = Sheet2.Range("NOB.TradeNameOrDescription2").Cells
    Dim i As Long
    ReDim TradeName1_NOB(end_NOBCode)
    For i = 1 To end_NOBCode
    TradeName1_NOB(i) = rangecells.item(i).value
Next
End Function

Function ValidateTradeName2_NOB() As Boolean
ValidateTradeName2_NOB = True
    Dim rangecells As Range
    Set rangecells = Sheet2.Range("NOB.TradeNameOrDescription3").Cells
    Dim i As Long
    ReDim TradeName2_NOB(end_NOBCode)
    For i = 1 To end_NOBCode
    TradeName2_NOB(i) = rangecells.item(i).value
Next
End Function

Sub setTableCode_NOB()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet2.Range("NOB.Code").count
    Set rangecells = Sheet2.Range("NOB.Code").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
                ccount = ccount + 1
            Else
                Exit For
            End If
    Next
 end_NOBCode = ccount
End Sub

Sub AddRows_NatureOfBusiness()
    Dim vRows As Long
    Sheet2.Activate
    EfilingCommon.DefinedgridNameRange = "NOB.Code||NOB.TradeNameOrDescription||NOB.TradeNameOrDescription2"
    ActiveCellRange = EfilingCommon.searchLastRow("NOB.Code")
    vRows = EfilingCommon.insertRowNatureOfBusiness
End Sub
