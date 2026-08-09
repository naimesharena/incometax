Attribute VB_Name = "Quantitative_Details"
Option Explicit

Public ItemName_QDTradingConcern As Variant
Public UnitOfMeasure_QDTradingConcern As Variant
Public OpeningStock_QDTradingConcern As Variant
Public PurchaseQty_QDTradingConcern As Variant
Public SaleQty_QDTradingConcern As Variant
Public ClgStock_QDTradingConcern As Variant
Public AnyShortExces_QDTradingConcern As Variant
Public ItemName_QDRawMaterial As Variant
Public UnitOfMeasure_QDRawMaterial As Variant
Public OpeningStock_QDRawMaterial As Variant
Public PurchaseQty_QDRawMaterial As Variant
Public SaleQty_QDRawMaterial As Variant
Public ClgStock_QDRawMaterial As Variant
Public PrevYrConsum_QDRawMaterial As Variant
Public yldFinisProd_QDRawMaterial As Variant
Public PercentYld_QDRawMaterial As Variant
Public AnyShortExces_QDRawMaterial As Variant
Public ItemName_QDFinishrByProd As Variant
Public UnitOfMeasure_QDFinishrByProd As Variant
Public OpeningStock_QDFinishrByProd As Variant
Public PurchaseQty_QDFinishrByProd As Variant
Public SaleQty_QDFinishrByProd As Variant
Public ClgStock_QDFinishrByProd As Variant
Public PrevYrConsum_QDFinishrByProd As Variant
Public AnyShortExces_QDFinishrByProd As Variant

Public rngname_QDTradingConcern As Variant
Public end_QDTradingConcern As Variant
Dim msgValidateSheetQDTradingConcern As String

Public rngname_QDRawMaterial As Variant
Public end_QDRawMaterial As Variant
Dim msgValidateSheetQDRawMaterial As String

Public rngname_QDFinishrByProd As Variant
Public end_QDFinishrByProd As Variant
Dim msgValidateSheet6 As Variant
Dim msgValidateSheetQDFinishrByProd As String
Dim msgSheetQDError As Variant
Sub AddRows_Trdng_QD()
    Dim vRows As Long
    Sheet6.Activate
    EfilingCommon.DefinedgridNameRange = "QDTradingConcern.ItemName||QDTradingConcern.UnitOfMeasure||QDTradingConcern.OpeningStock||QDTradingConcern.PurchaseQty||QDTradingConcern.SaleQty||QDTradingConcern.ClgStock||QDTradingConcern.AnyShortExces"
    ActiveCellRange = EfilingCommon.searchLastRow("QDTradingConcern.ItemName")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub AddRows_Rawmtrl_QD()
    Dim vRows As Long
    Sheet6.Activate
    EfilingCommon.DefinedgridNameRange = "QDRawMaterial.ItemName||QDRawMaterial.UnitOfMeasure||QDRawMaterial.OpeningStock||QDRawMaterial.PurchaseQty||QDRawMaterial.PrevYrConsum||QDRawMaterial.SaleQty||QDRawMaterial.ClgStock||QDRawMaterial.yldFinisProd||QDRawMaterial.PercentYld||QDRawMaterial.AnyShortExces"
    ActiveCellRange = EfilingCommon.searchLastRow("QDRawMaterial.ItemName")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows_Finished_QD()
    Dim vRows As Long
    Sheet6.Activate
    EfilingCommon.DefinedgridNameRange = "QDFinishrByProd.ItemName||QDFinishrByProd.UnitOfMeasure||QDFinishrByProd.OpeningStock||QDFinishrByProd.PurchaseQty||QDFinishrByProd.PrevYrConsum||QDFinishrByProd.SaleQty||QDFinishrByProd.ClgStock||QDFinishrByProd.AnyShortExces"
    ActiveCellRange = EfilingCommon.searchLastRow("QDFinishrByProd.ItemName")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub ValidateSheetQDClick()
ValidateSheetQD
fmsgboxStatus "Sheet Quantitative Details (QD) is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheetQD()
    
    If Not ValidatesheetQDTradingConcern Then
        Sheet6.Activate
        fmsgboxsmall "A" & Chr(13) & msgValidateSheetQDTradingConcern ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    '    MsgBox "Sheet is ok", vbOKOnly, "ITR-6"
    If Not ValidatesheetQDRawMaterial Then
        Sheet6.Activate
        fmsgboxsmall "B" & Chr(13) & msgValidateSheetQDRawMaterial ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    '    MsgBox "Sheet is ok ", vbOKOnly, "ITR-6"
       
    If Not ValidatesheetQDFinishrByProd Then
        Sheet6.Activate
        fmsgboxsmall "C" & Chr(13) & msgValidateSheetQDFinishrByProd ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    If Not CheckMandatoryVAlidate() Then
        Sheet6.Activate
        fmsgboxsmall msgSheetQDError ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub

Function ValidatesheetQDTradingConcern() As Boolean
     ValidatesheetQDTradingConcern = True
     subProcCaption = "Validating QD 1"
     If Not ValidateItemName_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
     If Not ValidateUnitOfMeasure_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
     If Not ValidateOpeningStock_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
     If Not ValidatePurchaseQty_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
     If Not ValidateSaleQty_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
     If Not ValidateClgStock_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
     If Not ValidateAnyShortExces_QDTradingConcern() Then ValidatesheetQDTradingConcern = False
End Function

Sub setTblinfo_QDTradingConcern()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet6.Range("QDTradingConcern.ItemName").count
    Set rangecells = Sheet6.Range("QDTradingConcern.ItemName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_QDTradingConcern = ccount
    rngname_QDTradingConcern = "QDTradingConcern.ItemName;QDTradingConcern.UnitOfMeasure;QDTradingConcern.OpeningStock;QDTradingConcern.PurchaseQty;QDTradingConcern.SaleQty;QDTradingConcern.ClgStock;QDTradingConcern.AnyShortExces;"
End Sub



Function ValidatesheetQDRawMaterial() As Boolean
     ValidatesheetQDRawMaterial = True
     subProcCaption = "Validating QD 2"
     If Not ValidateItemName_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidateUnitOfMeasure_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidateOpeningStock_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidatePurchaseQty_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidateSaleQty_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidateClgStock_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidatePrevYrConsum_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidateyldFinisProd_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidatePercentYld_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
     If Not ValidateAnyShortExces_QDRawMaterial() Then ValidatesheetQDRawMaterial = False
End Function

Sub setTblinfo_QDRawMaterial()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet6.Range("QDRawMaterial.ItemName").count
    Set rangecells = Sheet6.Range("QDRawMaterial.ItemName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_QDRawMaterial = ccount
    rngname_QDRawMaterial = "QDRawMaterial.ItemName;QDRawMaterial.UnitOfMeasure;QDRawMaterial.OpeningStock;QDRawMaterial.PurchaseQty;QDRawMaterial.SaleQty;QDRawMaterial.ClgStock;QDRawMaterial.PrevYrConsum;QDRawMaterial.yldFinisProd;QDRawMaterial.PercentYld;QDRawMaterial.AnyShortExces;"
End Sub

Function ValidatesheetQDFinishrByProd() As Boolean
     ValidatesheetQDFinishrByProd = True
     subProcCaption = "Validating QD 3"
     If Not ValidateItemName_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidateUnitOfMeasure_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidateOpeningStock_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidatePurchaseQty_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidateSaleQty_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidateClgStock_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidatePrevYrConsum_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
     If Not ValidateAnyShortExces_QDFinishrByProd() Then ValidatesheetQDFinishrByProd = False
'Change-6, ANK, 27.05.2022
     If Not ValidateItemName_FinishedProduct() Then ValidatesheetQDFinishrByProd = False
    
'---end change---
End Function

Sub setTblinfo_QDFinishrByProd()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    
    mIntCells = Sheet6.Range("QDFinishrByProd.ItemName").count
    Set rangecells = Sheet6.Range("QDFinishrByProd.ItemName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_QDFinishrByProd = ccount
    rngname_QDFinishrByProd = "QDFinishrByProd.ItemName;QDFinishrByProd.UnitOfMeasure;QDFinishrByProd.OpeningStock;QDFinishrByProd.PurchaseQty;QDFinishrByProd.SaleQty;QDFinishrByProd.ClgStock;QDFinishrByProd.PrevYrConsum;QDFinishrByProd.AnyShortExces;"
End Sub

Function ValidateItemName_QDTradingConcern() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateItemName_QDTradingConcern = True
    setTblinfo_QDTradingConcern
    
    Set rangecells = Sheet6.Range("QDTradingConcern.ItemName").Cells
    ReDim ItemName_QDTradingConcern(end_QDTradingConcern)
    noOfProcessSub = end_QDTradingConcern
    For i = 1 To end_QDTradingConcern
        ItemName_QDTradingConcern(i) = rangecells.item(i).value
        If Not chkCompulsory(ItemName_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please enter Item name  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateItemName_QDTradingConcern = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(ItemName_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* ItemName at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule Quantitative Details")
            ValidateItemName_QDTradingConcern = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
 
Function ValidateUnitOfMeasure_QDTradingConcern() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateUnitOfMeasure_QDTradingConcern = True
    setTblinfo_QDTradingConcern

    Set rangecells = Sheet6.Range("QDTradingConcern.UnitOfMeasure").Cells
    ReDim UnitOfMeasure_QDTradingConcern(end_QDTradingConcern)
    For i = 1 To end_QDTradingConcern
        UnitOfMeasure_QDTradingConcern(i) = rangecells.item(i).value
        UnitOfMeasure_QDTradingConcern(i) = Mid(UnitOfMeasure_QDTradingConcern(i), 1, 3)
        
        If UnitOfMeasure_QDTradingConcern(i) = "(Se" Then
            msgbox_QDTradingConcern ("* Please select Unit of measure  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateUnitOfMeasure_QDTradingConcern = False
            Exit Function
        End If
        
        If Not chkCompulsory(UnitOfMeasure_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please select Unit of measure at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateUnitOfMeasure_QDTradingConcern = False
            Exit Function
        End If
    Next
End Function

Function ValidateOpeningStock_QDTradingConcern() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateOpeningStock_QDTradingConcern = True
    setTblinfo_QDTradingConcern

    Set rangecells = Sheet6.Range("QDTradingConcern.OpeningStock").Cells
    ReDim OpeningStock_QDTradingConcern(end_QDTradingConcern)
    For i = 1 To end_QDTradingConcern
        OpeningStock_QDTradingConcern(i) = rangecells.item(i).value
        If Not chkCompulsory(OpeningStock_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please enter Opening stock at Sr. No  " & i & "  in schedule Quantitative Details")
            ValidateOpeningStock_QDTradingConcern = False
            Exit Function
        End If
    Next
End Function

Function ValidatePurchaseQty_QDTradingConcern() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePurchaseQty_QDTradingConcern = True
    setTblinfo_QDTradingConcern
    
    Set rangecells = Sheet6.Range("QDTradingConcern.PurchaseQty").Cells
    ReDim PurchaseQty_QDTradingConcern(end_QDTradingConcern)
    For i = 1 To end_QDTradingConcern
        PurchaseQty_QDTradingConcern(i) = rangecells.item(i).value
        If Not chkCompulsory(PurchaseQty_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please enter Purchase at Sr. No  " & i & " in schedule Quantitative Details")
            ValidatePurchaseQty_QDTradingConcern = False
            Exit Function
        End If
    Next
End Function

Function ValidateSaleQty_QDTradingConcern() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateSaleQty_QDTradingConcern = True
    setTblinfo_QDTradingConcern
    
    Set rangecells = Sheet6.Range("QDTradingConcern.SaleQty").Cells
    ReDim SaleQty_QDTradingConcern(end_QDTradingConcern)
    For i = 1 To end_QDTradingConcern
        SaleQty_QDTradingConcern(i) = rangecells.item(i).value
        If Not chkCompulsory(SaleQty_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please enter Sales  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateSaleQty_QDTradingConcern = False
            Exit Function
        End If
    Next
End Function

Function ValidateClgStock_QDTradingConcern() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateClgStock_QDTradingConcern = True
    setTblinfo_QDTradingConcern
    
    Set rangecells = Sheet6.Range("QDTradingConcern.ClgStock").Cells
    ReDim ClgStock_QDTradingConcern(end_QDTradingConcern)
    For i = 1 To end_QDTradingConcern
        ClgStock_QDTradingConcern(i) = rangecells.item(i).value
        If Not chkCompulsory(ClgStock_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please enter Closing stock at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateClgStock_QDTradingConcern = False
            Exit Function
        End If
    Next
End Function

Function ValidateAnyShortExces_QDTradingConcern() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAnyShortExces_QDTradingConcern = True
    setTblinfo_QDTradingConcern

    Set rangecells = Sheet6.Range("QDTradingConcern.AnyShortExces").Cells
    ReDim AnyShortExces_QDTradingConcern(end_QDTradingConcern)
    For i = 1 To end_QDTradingConcern
        AnyShortExces_QDTradingConcern(i) = rangecells.item(i).value
        If Not chkCompulsory(AnyShortExces_QDTradingConcern(i)) Then
            msgbox_QDTradingConcern ("* Please enter Shortage/ excess at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateAnyShortExces_QDTradingConcern = False
            Exit Function
        End If
    Next
End Function

Function ValidateItemName_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateItemName_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.ItemName").Cells
    ReDim ItemName_QDRawMaterial(end_QDRawMaterial)
    noOfProcessSub = end_QDRawMaterial
    For i = 1 To end_QDRawMaterial
        ItemName_QDRawMaterial(i) = rangecells.item(i).value
        If Len(ItemName_QDRawMaterial(i)) > 0 Then
            If Not checkfieldspecialcharacter(ItemName_QDRawMaterial(i)) Then
                msgbox_QDRawMaterial ("* Please enter Item name at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule Quantitative Details")
                ValidateItemName_QDRawMaterial = False
                Exit Function
            End If
        End If
        UpdateProgressBar
    Next
End Function
 
Function ValidateUnitOfMeasure_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateUnitOfMeasure_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.UnitOfMeasure").Cells
    ReDim UnitOfMeasure_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        UnitOfMeasure_QDRawMaterial(i) = rangecells.item(i).value
        UnitOfMeasure_QDRawMaterial(i) = Mid(UnitOfMeasure_QDRawMaterial(i), 1, 3)
        If UnitOfMeasure_QDRawMaterial(i) = "(Se" Then
            msgbox_QDRawMaterial ("* Please select Unit of measure  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateUnitOfMeasure_QDRawMaterial = False
            Exit Function
        End If
        If Not chkCompulsory(UnitOfMeasure_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please select Unit of measure at Sr. No  " & i & "  in schedule Quantitative Details")
            ValidateUnitOfMeasure_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidateOpeningStock_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateOpeningStock_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.OpeningStock").Cells
    ReDim OpeningStock_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        OpeningStock_QDRawMaterial(i) = rangecells.item(i).value
        If Not chkCompulsory(OpeningStock_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please enter Opening stock at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateOpeningStock_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidatePurchaseQty_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePurchaseQty_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.PurchaseQty").Cells
    ReDim PurchaseQty_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        PurchaseQty_QDRawMaterial(i) = rangecells.item(i).value
        If Not chkCompulsory(PurchaseQty_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please enter Purchase  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidatePurchaseQty_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidateSaleQty_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateSaleQty_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.SaleQty").Cells
    ReDim SaleQty_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        SaleQty_QDRawMaterial(i) = rangecells.item(i).value
        If Not chkCompulsory(SaleQty_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please enter Consumption  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateSaleQty_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidateClgStock_QDRawMaterial() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateClgStock_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.ClgStock").Cells
    ReDim ClgStock_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        ClgStock_QDRawMaterial(i) = rangecells.item(i).value
        If Not chkCompulsory(ClgStock_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please enter Closing stock  at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateClgStock_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidatePrevYrConsum_QDRawMaterial() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidatePrevYrConsum_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.PrevYrConsum").Cells
    ReDim PrevYrConsum_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        PrevYrConsum_QDRawMaterial(i) = rangecells.item(i).value
        If Not chkCompulsory(PrevYrConsum_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please enter Consumption. at Sr. No  " & i & " in schedule Quantitative Details")
            ValidatePrevYrConsum_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidateyldFinisProd_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateyldFinisProd_QDRawMaterial = True
    setTblinfo_QDRawMaterial
    
    Set rangecells = Sheet6.Range("QDRawMaterial.yldFinisProd").Cells
    ReDim yldFinisProd_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        yldFinisProd_QDRawMaterial(i) = rangecells.item(i).value
        
    Next
End Function

Function ValidatePercentYld_QDRawMaterial() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePercentYld_QDRawMaterial = True
    setTblinfo_QDRawMaterial

    Set rangecells = Sheet6.Range("QDRawMaterial.PercentYld").Cells
    ReDim PercentYld_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        PercentYld_QDRawMaterial(i) = rangecells.item(i).text
        
    Next
End Function

Function ValidateAnyShortExces_QDRawMaterial() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateAnyShortExces_QDRawMaterial = True
    setTblinfo_QDRawMaterial

    Set rangecells = Sheet6.Range("QDRawMaterial.AnyShortExces").Cells
    
    ReDim AnyShortExces_QDRawMaterial(end_QDRawMaterial)
    For i = 1 To end_QDRawMaterial
        AnyShortExces_QDRawMaterial(i) = rangecells.item(i).value
        If Not chkCompulsory(AnyShortExces_QDRawMaterial(i)) Then
            msgbox_QDRawMaterial ("* Please enter Shortage/ excessif any at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateAnyShortExces_QDRawMaterial = False
            Exit Function
        End If
    Next
End Function

Function ValidateItemName_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateItemName_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd
    
    Set rangecells = Sheet6.Range("QDFinishrByProd.ItemName").Cells
    ReDim ItemName_QDFinishrByProd(end_QDFinishrByProd)
    noOfProcessSub = end_QDFinishrByProd
    If end_QDFinishrByProd >= 1 Then
        If end_QDFinishrByProd = 0 Then
            msgbox_QDFinishrByProd ("*  Please enter Finished goods entry missing in Sheet Quantitative Details in schedule Quantitative Details")
            ValidateItemName_QDFinishrByProd = False
            Exit Function
        End If
    End If
    
    If end_QDFinishrByProd >= 1 Then
        If end_QDRawMaterial = 0 Then
            msgbox_QDFinishrByProd ("* Please enter Raw materials entry missing in Sheet Quantitative Details in schedule Quantitative Details")
            ValidateItemName_QDFinishrByProd = False
            Exit Function
        End If
    End If
    
    For i = 1 To end_QDFinishrByProd
        ItemName_QDFinishrByProd(i) = rangecells.item(i).value
        If Len(ItemName_QDFinishrByProd(i)) > 0 Then
            If Not checkfieldspecialcharacter(ItemName_QDFinishrByProd(i)) Then
                msgbox_QDFinishrByProd ("* Item Name at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule Quantitative Details")
                ValidateItemName_QDFinishrByProd = False
                Exit Function
            End If
        End If
        UpdateProgressBar
    Next
End Function
 
Function ValidateUnitOfMeasure_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateUnitOfMeasure_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd
    
    Set rangecells = Sheet6.Range("QDFinishrByProd.UnitOfMeasure").Cells
    ReDim UnitOfMeasure_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        UnitOfMeasure_QDFinishrByProd(i) = rangecells.item(i).value
        UnitOfMeasure_QDFinishrByProd(i) = Mid(UnitOfMeasure_QDFinishrByProd(i), 1, 3)
        If UnitOfMeasure_QDFinishrByProd(i) = "(Se" Then
            msgbox_QDFinishrByProd ("* Plase enter Unit Of Measure at Sr. No  " & i & " in schedule Quantitative Details")
            ValidateUnitOfMeasure_QDFinishrByProd = False
            Exit Function
        End If
        If Not chkCompulsory(UnitOfMeasure_QDFinishrByProd(i)) Then
            msgbox_QDFinishrByProd ("* Please enter Unit Of Measure at Sr. No  " & i & "  in schedule Quantitative Details")
            ValidateUnitOfMeasure_QDFinishrByProd = False
            Exit Function
        End If
    Next
End Function

Function ValidateOpeningStock_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateOpeningStock_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd
    
    Set rangecells = Sheet6.Range("QDFinishrByProd.OpeningStock").Cells
    ReDim OpeningStock_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        OpeningStock_QDFinishrByProd(i) = rangecells.item(i).value
        If Not chkCompulsory(OpeningStock_QDFinishrByProd(i)) Then
            msgbox_QDFinishrByProd ("* Please enter Opening Stock at Sr. No  " & i & "  in schedule Quantitative Details")
            ValidateOpeningStock_QDFinishrByProd = False
            Exit Function
        End If
    Next
End Function

Function ValidatePurchaseQty_QDFinishrByProd() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidatePurchaseQty_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd

    Set rangecells = Sheet6.Range("QDFinishrByProd.PurchaseQty").Cells
    ReDim PurchaseQty_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        PurchaseQty_QDFinishrByProd(i) = rangecells.item(i).value
        If Not chkCompulsory(PurchaseQty_QDFinishrByProd(i)) Then
            msgbox_QDFinishrByProd ("* Please enter Purchase at Sr. No  " & i & " in schedule Quantitative Details")
            ValidatePurchaseQty_QDFinishrByProd = False
            Exit Function
        End If
    Next
End Function

Function ValidateSaleQty_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateSaleQty_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd
    
    Set rangecells = Sheet6.Range("QDFinishrByProd.SaleQty").Cells
    ReDim SaleQty_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        SaleQty_QDFinishrByProd(i) = rangecells.item(i).value
        If Not chkCompulsory(SaleQty_QDFinishrByProd(i)) Then
            msgbox_QDFinishrByProd ("* Please enter Sales Sr. No  " & i & "in schedule Quantitative Details")
            ValidateSaleQty_QDFinishrByProd = False
            Exit Function
        End If
    Next
End Function

Function ValidateClgStock_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
        
    ValidateClgStock_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd

    Set rangecells = Sheet6.Range("QDFinishrByProd.ClgStock").Cells
    ReDim ClgStock_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        ClgStock_QDFinishrByProd(i) = rangecells.item(i).value
        If Not chkCompulsory(ClgStock_QDFinishrByProd(i)) Then
            msgbox_QDFinishrByProd ("* Please enter Closing Stock Sr. No  " & i & "in schedule Quantitative Details")
            ValidateClgStock_QDFinishrByProd = False
            Exit Function
        End If
    Next
End Function

Function ValidatePrevYrConsum_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePrevYrConsum_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd
    
    Set rangecells = Sheet6.Range("QDFinishrByProd.PrevYrConsum").Cells
    ReDim PrevYrConsum_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        PrevYrConsum_QDFinishrByProd(i) = rangecells.item(i).value
    Next
End Function

Function ValidateAnyShortExces_QDFinishrByProd() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAnyShortExces_QDFinishrByProd = True
    setTblinfo_QDFinishrByProd
    
    Set rangecells = Sheet6.Range("QDFinishrByProd.AnyShortExces").Cells
    ReDim AnyShortExces_QDFinishrByProd(end_QDFinishrByProd)
    For i = 1 To end_QDFinishrByProd
        AnyShortExces_QDFinishrByProd(i) = rangecells.item(i).value
        If Not chkCompulsory(AnyShortExces_QDFinishrByProd(i)) Then
            msgbox_QDFinishrByProd ("* Please enter Shortage/excess Sr. No  " & i & " in schedule Quantitative Details")
            ValidateAnyShortExces_QDFinishrByProd = False
            Exit Function
        End If
    Next
End Function

'Change-5, ANK, 27.05.2022
Function ValidateItemName_FinishedProduct() As Boolean
ValidateItemName_FinishedProduct = True
'if item is blank but any other field is filled then throw an error
Dim FinishedProductCombinedRanges As String
Dim FinishedProductRangeArray() As String
FinishedProductCombinedRanges = "QDFinishrByProd.ItemName;QDFinishrByProd.UnitOfMeasure;QDFinishrByProd.OpeningStock;QDFinishrByProd.PurchaseQty;QDFinishrByProd.SaleQty;QDFinishrByProd.ClgStock;QDFinishrByProd.PrevYrConsum;QDFinishrByProd.AnyShortExces;"
FinishedProductRangeArray = Split(FinishedProductCombinedRanges, ";")

Dim tableRange As Variant
Dim ItemNameEntries As Long
Dim maxEntries As Long
Dim rangecells As Range
Dim i As Long
Dim mIntCells As Long
Dim mIntCtr As Long
Dim ccount As Long

maxEntries = 0
ccount = 0
mIntCells = Sheet6.Range("QDFinishrByProd.ItemName").count
Set rangecells = Sheet6.Range("QDFinishrByProd.ItemName").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
ItemNameEntries = ccount

Dim maxIndex As Long
maxIndex = UBound(FinishedProductRangeArray)

For i = 0 To (maxIndex - 1)
    ccount = 0
    mIntCells = Sheet6.Range(FinishedProductRangeArray(i)).count
    Set rangecells = Sheet6.Range(FinishedProductRangeArray(i)).Cells
    For mIntCtr = 1 To mIntCells
        If (Not rangecells.item(mIntCtr).value = "") And (Not rangecells.item(mIntCtr).value = "(Select)") Then
            ccount = ccount + 1
        End If
    Next
    
    If ccount > maxEntries Then
        maxEntries = ccount
    End If
    
Next i

If ItemNameEntries < maxEntries Then
    msgbox_QDFinishrByProd ("* Please enter Item name in schedule Quantitative Details")
    ValidateItemName_FinishedProduct = False
End If


End Function
'---end change---

Function CheckMandatoryVAlidate() As Boolean
    CheckMandatoryVAlidate = True
   If UBound(ItemName_QDRawMaterial) > 0 Or UBound(ItemName_QDFinishrByProd) > 0 Then
    
    If UBound(ItemName_QDFinishrByProd) = 0 Then
        msgSheetQDError = msgSheetQDError + "Please enter Finished Goods Table is mandatory in schedule Quantitative Details" & Chr(13)
        CheckMandatoryVAlidate = False
    End If
    
    If UBound(ItemName_QDRawMaterial) = 0 Then
        msgSheetQDError = msgSheetQDError + "Please enter Raw materials Table is mandatory in schedule Quantitative Details" & Chr(13)
        CheckMandatoryVAlidate = False
    End If
    
   End If

End Function

Function DefaultOpeningStock_QDTradingConcern() As String
    DefaultOpeningStock_QDTradingConcern = "0"
End Function
Function DefaultPurchaseQty_QDTradingConcern() As String
    DefaultPurchaseQty_QDTradingConcern = "0"
End Function
Function DefaultSaleQty_QDTradingConcern() As String
    DefaultSaleQty_QDTradingConcern = "0"
End Function
Function DefaultClgStock_QDTradingConcern() As String
    DefaultClgStock_QDTradingConcern = "0"
End Function
Function DefaultAnyShortExces_QDTradingConcern() As String
    DefaultAnyShortExces_QDTradingConcern = "0"
End Function
Function DefaultItemName_QDRawMaterial() As String
    DefaultItemName_QDRawMaterial = "0"
End Function
Function DefaultUnitOfMeasure_QDRawMaterial() As String
    DefaultUnitOfMeasure_QDRawMaterial = "0"
End Function
Function DefaultOpeningStock_QDRawMaterial() As String
    DefaultOpeningStock_QDRawMaterial = "0"
End Function
Function DefaultPurchaseQty_QDRawMaterial() As String
    DefaultPurchaseQty_QDRawMaterial = "0"
End Function
Function DefaultSaleQty_QDRawMaterial() As String
    DefaultSaleQty_QDRawMaterial = "0"
End Function
Function DefaultClgStock_QDRawMaterial() As String
    DefaultClgStock_QDRawMaterial = "0"
End Function
Function DefaultPrevYrConsum_QDRawMaterial() As String
    DefaultPrevYrConsum_QDRawMaterial = "0"
End Function
Function DefaultyldFinisProd_QDRawMaterial() As String
    DefaultyldFinisProd_QDRawMaterial = "0"
End Function
Function DefaultPercentYld_QDRawMaterial() As String
    DefaultPercentYld_QDRawMaterial = "0"
End Function
Function DefaultAnyShortExces_QDRawMaterial() As String
    DefaultAnyShortExces_QDRawMaterial = "0"
End Function
Function DefaultItemName_QDFinishrByProd() As String
    DefaultItemName_QDFinishrByProd = "0"
End Function
Function DefaultUnitOfMeasure_QDFinishrByProd() As String
    DefaultUnitOfMeasure_QDFinishrByProd = "0"
End Function
Function DefaultOpeningStock_QDFinishrByProd() As String
    DefaultOpeningStock_QDFinishrByProd = "0"
End Function
Function DefaultPurchaseQty_QDFinishrByProd() As String
    DefaultPurchaseQty_QDFinishrByProd = "0"
End Function
Function DefaultSaleQty_QDFinishrByProd() As String
    DefaultSaleQty_QDFinishrByProd = "0"
End Function
Function DefaultClgStock_QDFinishrByProd() As String
    DefaultClgStock_QDFinishrByProd = "0"
End Function
Function DefaultPrevYrConsum_QDFinishrByProd() As String
    DefaultPrevYrConsum_QDFinishrByProd = "0"
End Function
Function DefaultAnyShortExces_QDFinishrByProd() As String
    DefaultAnyShortExces_QDFinishrByProd = "0"
End Function

Function msgbox5(strmsg As String) As String
    msgValidateSheet6 = msgValidateSheet6 & strmsg & Chr(13)
End Function

Function msgbox_QDTradingConcern(strmsg As String) As String
     msgValidateSheetQDTradingConcern = msgValidateSheetQDTradingConcern & strmsg & Chr(13)
End Function

Function msgbox_QDRawMaterial(strmsg As String) As String
     msgValidateSheetQDRawMaterial = msgValidateSheetQDRawMaterial & strmsg & Chr(13)
End Function

Function msgbox_QDFinishrByProd(strmsg As String) As String
     msgValidateSheetQDFinishrByProd = msgValidateSheetQDFinishrByProd & strmsg & Chr(13)
End Function

