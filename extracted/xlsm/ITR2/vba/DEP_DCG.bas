Attribute VB_Name = "DEP_DCG"
Option Explicit

Public DeprBlockTot15Percent_DEPP As Variant
Public DeprBlockTot30Percent_DEPP As Variant
Public DeprBlockTot40Percent_DEPP As Variant

Public TotPlntMach_DEPP As Variant
Public DeprBlockTot5Percent_DEPB As Variant
Public DeprBlockTot10Percent_DEPB As Variant
Public DeprBlockTot100Percent_DEPB As Variant
Public TotBuildng_DEPB As Variant
Public FurnitureSummary_DEP As Variant
Public IntangibleAssetSummary_DEP As Variant
Public ShipsSummary_DEP As Variant
Public TotalDepreciation_DEP As Variant
Public DeprBlockTot15Percent_DCGP As Variant
Public DeprBlockTot30Percent_DCGP As Variant
Public DeprBlockTot40Percent_DCGP As Variant

Public TotPlntMach_DCGP As Variant
Public DeprBlockTot5Percent_DCGB As Variant
Public DeprBlockTot10Percent_DCGB As Variant
Public DeprBlockTot100Percent_DCGB As Variant
Public TotBuildng_DCGB As Variant
Public FurnitureSummary_DCG As Variant
Public IntangibleAssetSummary_DCG As Variant
Public ShipsSummary_DCG As Variant
Public TotalDepreciation_DCG As Variant

Dim msgValidateSheetDEP As Variant
Dim msgValidateSheetDCG As Variant

Sub validateDEPDCGClick()
ValidateSheetDEP_DCG
fmsgboxStatus "Sheet DEP_DCG is OK" ', vbOKOnly
End Sub

Sub ValidateSheetDEP_DCG()
    If Not ValidatesheetDEP Then
        Sheet9.Activate
        fmsgboxsmall msgValidateSheetDEP ', vbOKOnly, "Error(s)!"
        CloseMsg
    'Else
        'MsgBox "Sheet is ok",vbOKOnly,"ITR-3"
    End If
    
    If Not ValidatesheetDCG Then
        Sheet9.Activate
        fmsgboxsmall msgValidateSheetDCG ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub

Function ValidatesheetDEP() As Boolean
     ValidatesheetDEP = True
     If Not ValidateDeprBlockTot15Percent_DEPP() Then ValidatesheetDEP = False
     If Not ValidateDeprBlockTot30Percent_DEPP() Then ValidatesheetDEP = False
     If Not ValidateDeprBlockTot40Percent_DEPP() Then ValidatesheetDEP = False

     If Not ValidateTotPlntMach_DEPP() Then ValidatesheetDEP = False
     If Not ValidateDeprBlockTot5Percent_DEPB() Then ValidatesheetDEP = False
     If Not ValidateDeprBlockTot10Percent_DEPB() Then ValidatesheetDEP = False
     If Not ValidateDeprBlockTot100Percent_DEPB() Then ValidatesheetDEP = False
     If Not ValidateTotBuildng_DEPB() Then ValidatesheetDEP = False
     If Not ValidateFurnitureSummary_DEP() Then ValidatesheetDEP = False
     If Not ValidateIntangibleAssetSummary_DEP() Then ValidatesheetDEP = False
     If Not ValidateShipsSummary_DEP() Then ValidatesheetDEP = False
     If Not ValidateTotalDepreciation_DEP() Then ValidatesheetDEP = False
End Function

Function ValidatesheetDCG() As Boolean
     ValidatesheetDCG = True
     If Not ValidateDeprBlockTot15Percent_DCGP() Then ValidatesheetDCG = False
     If Not ValidateDeprBlockTot30Percent_DCGP() Then ValidatesheetDCG = False
     If Not ValidateDeprBlockTot40Percent_DCGP() Then ValidatesheetDCG = False

     If Not ValidateTotPlntMach_DCGP() Then ValidatesheetDCG = False
     If Not ValidateDeprBlockTot5Percent_DCGB() Then ValidatesheetDCG = False
     If Not ValidateDeprBlockTot10Percent_DCGB() Then ValidatesheetDCG = False
     If Not ValidateDeprBlockTot100Percent_DCGB() Then ValidatesheetDCG = False
     If Not ValidateTotBuildng_DCGB() Then ValidatesheetDCG = False
     If Not ValidateFurnitureSummary_DCG() Then ValidatesheetDCG = False
     If Not ValidateIntangibleAssetSummary_DCG() Then ValidatesheetDCG = False
     If Not ValidateShipsSummary_DCG() Then ValidatesheetDCG = False
     If Not ValidateTotalDepreciation_DCG() Then ValidatesheetDCG = False
End Function

Function ValidateDeprBlockTot15Percent_DEPP() As Boolean
    ValidateDeprBlockTot15Percent_DEPP = True
    DeprBlockTot15Percent_DEPP = Sheet11.Range("DEPP.DeprBlockTot15Percent").text
End Function

Function ValidateDeprBlockTot30Percent_DEPP() As Boolean
    ValidateDeprBlockTot30Percent_DEPP = True
    DeprBlockTot30Percent_DEPP = Sheet11.Range("DEPP.DeprBlockTot30Percent").text
End Function

Function ValidateDeprBlockTot40Percent_DEPP() As Boolean
    ValidateDeprBlockTot40Percent_DEPP = True
    DeprBlockTot40Percent_DEPP = Sheet11.Range("DEPP.DeprBlockTot40Percent").text
End Function

Function ValidateTotPlntMach_DEPP() As Boolean
    ValidateTotPlntMach_DEPP = True
    TotPlntMach_DEPP = Sheet11.Range("DEPP.TotPlntMach").text
    
    If Len(TotPlntMach_DEPP) > 14 Then
    msgValidateSheetDEP = msgValidateSheetDEP & "* Value at Total depreciation on plant and machinery  (1a + 1b + 1c +1d)  should be at most 14 digits " & Chr(13)
    ValidateTotPlntMach_DEPP = False
    End If
End Function

Function ValidateDeprBlockTot5Percent_DEPB() As Boolean
    ValidateDeprBlockTot5Percent_DEPB = True
    DeprBlockTot5Percent_DEPB = Sheet11.Range("DEPB.DeprBlockTot5Percent").text
End Function

Function ValidateDeprBlockTot10Percent_DEPB() As Boolean
    ValidateDeprBlockTot10Percent_DEPB = True
    DeprBlockTot10Percent_DEPB = Sheet11.Range("DEPB.DeprBlockTot10Percent").text
End Function

Function ValidateDeprBlockTot100Percent_DEPB() As Boolean
    ValidateDeprBlockTot100Percent_DEPB = True
    DeprBlockTot100Percent_DEPB = Sheet11.Range("DEPB.DeprBlockTot100Percent").text
End Function

Function ValidateTotBuildng_DEPB() As Boolean
    ValidateTotBuildng_DEPB = True
    TotBuildng_DEPB = Sheet11.Range("DEPB.TotBuildng").text
    If Len(TotBuildng_DEPB) > 14 Then
    msgValidateSheetDEP = msgValidateSheetDEP & "* Value at Total depreciation on building (2a+2b+2c) should be at most 14 digits " & Chr(13)
    ValidateTotBuildng_DEPB = False
    End If
End Function

Function ValidateFurnitureSummary_DEP() As Boolean
    ValidateFurnitureSummary_DEP = True
    FurnitureSummary_DEP = Sheet11.Range("DEP.FurnitureSummary").text
End Function

Function ValidateIntangibleAssetSummary_DEP() As Boolean
    ValidateIntangibleAssetSummary_DEP = True
    IntangibleAssetSummary_DEP = Sheet11.Range("DEP.IntangibleAssetSummary").text
End Function

Function ValidateShipsSummary_DEP() As Boolean
    ValidateShipsSummary_DEP = True
    ShipsSummary_DEP = Sheet11.Range("DEP.ShipsSummary").text
End Function

Function ValidateTotalDepreciation_DEP() As Boolean
    ValidateTotalDepreciation_DEP = True
    TotalDepreciation_DEP = Sheet11.Range("DEP.TotalDepreciation").text
    If Len(TotalDepreciation_DEP) > 14 Then
    msgValidateSheetDEP = msgValidateSheetDEP & "* Value at Total depreciation ( 1e+2d+3+4+5) should be at most 14 digits " & Chr(13)
    ValidateTotalDepreciation_DEP = False
    End If
    
End Function

Function ValidateDeprBlockTot15Percent_DCGP() As Boolean
    ValidateDeprBlockTot15Percent_DCGP = True
    DeprBlockTot15Percent_DCGP = Sheet11.Range("DCGP.DeprBlockTot15Percent").text
End Function

Function ValidateDeprBlockTot30Percent_DCGP() As Boolean
    ValidateDeprBlockTot30Percent_DCGP = True
    DeprBlockTot30Percent_DCGP = Sheet11.Range("DCGP.DeprBlockTot30Percent").text
End Function

Function ValidateDeprBlockTot40Percent_DCGP() As Boolean
    ValidateDeprBlockTot40Percent_DCGP = True
    DeprBlockTot40Percent_DCGP = Sheet11.Range("DCGP.DeprBlockTot40Percent").text
End Function

Function ValidateTotPlntMach_DCGP() As Boolean
    ValidateTotPlntMach_DCGP = True
    TotPlntMach_DCGP = Sheet11.Range("DCGP.TotPlntMach").text
End Function

Function ValidateDeprBlockTot5Percent_DCGB() As Boolean
    ValidateDeprBlockTot5Percent_DCGB = True
    DeprBlockTot5Percent_DCGB = Sheet11.Range("DCGB.DeprBlockTot5Percent").text
End Function

Function ValidateDeprBlockTot10Percent_DCGB() As Boolean
    ValidateDeprBlockTot10Percent_DCGB = True
    DeprBlockTot10Percent_DCGB = Sheet11.Range("DCGB.DeprBlockTot10Percent").text
End Function

Function ValidateDeprBlockTot100Percent_DCGB() As Boolean
    ValidateDeprBlockTot100Percent_DCGB = True
    DeprBlockTot100Percent_DCGB = Sheet11.Range("DCGB.DeprBlockTot100Percent").text
End Function

Function ValidateTotBuildng_DCGB() As Boolean
    ValidateTotBuildng_DCGB = True
    TotBuildng_DCGB = Sheet11.Range("DCGB.TotBuildng").text
End Function

Function ValidateFurnitureSummary_DCG() As Boolean
    ValidateFurnitureSummary_DCG = True
    FurnitureSummary_DCG = Sheet11.Range("DCG.FurnitureSummary").text
End Function

Function ValidateIntangibleAssetSummary_DCG() As Boolean
    ValidateIntangibleAssetSummary_DCG = True
    IntangibleAssetSummary_DCG = Sheet11.Range("DCG.IntangibleAssetSummary").text
End Function

Function ValidateShipsSummary_DCG() As Boolean
    ValidateShipsSummary_DCG = True
    ShipsSummary_DCG = Sheet11.Range("DCG.ShipsSummary").text
End Function

Function ValidateTotalDepreciation_DCG() As Boolean
    ValidateTotalDepreciation_DCG = True
    TotalDepreciation_DCG = Sheet11.Range("DCG.TotalDepreciation").text
End Function

Function DefaultDeprBlockTot15Percent_DEPP() As String
DefaultDeprBlockTot15Percent_DEPP = "0"
End Function
Function DefaultDeprBlockTot30Percent_DEPP() As String
DefaultDeprBlockTot30Percent_DEPP = "0"
End Function
Function DefaultDeprBlockTot40Percent_DEPP() As String
DefaultDeprBlockTot40Percent_DEPP = "0"
End Function

Function DefaultTotPlntMach_DEPP() As String
DefaultTotPlntMach_DEPP = "0"
End Function
Function DefaultDeprBlockTot5Percent_DEPB() As String
DefaultDeprBlockTot5Percent_DEPB = "0"
End Function
Function DefaultDeprBlockTot10Percent_DEPB() As String
DefaultDeprBlockTot10Percent_DEPB = "0"
End Function
Function DefaultDeprBlockTot100Percent_DEPB() As String
DefaultDeprBlockTot100Percent_DEPB = "0"
End Function
Function DefaultTotBuildng_DEPB() As String
DefaultTotBuildng_DEPB = "0"
End Function
Function DefaultFurnitureSummary_DEP() As String
DefaultFurnitureSummary_DEP = "0"
End Function
Function DefaultIntangibleAssetSummary_DEP() As String
DefaultIntangibleAssetSummary_DEP = "0"
End Function
Function DefaultShipsSummary_DEP() As String
DefaultShipsSummary_DEP = "0"
End Function
Function DefaultTotalDepreciation_DEP() As String
DefaultTotalDepreciation_DEP = "0"
End Function
Function DefaultDeprBlockTot15Percent_DCGP() As String
DefaultDeprBlockTot15Percent_DCGP = "0"
End Function
Function DefaultDeprBlockTot30Percent_DCGP() As String
DefaultDeprBlockTot30Percent_DCGP = "0"
End Function
Function DefaultDeprBlockTot40Percent_DCGP() As String
DefaultDeprBlockTot40Percent_DCGP = "0"
End Function

Function DefaultTotPlntMach_DCGP() As String
DefaultTotPlntMach_DCGP = "0"
End Function
Function DefaultDeprBlockTot5Percent_DCGB() As String
DefaultDeprBlockTot5Percent_DCGB = "0"
End Function
Function DefaultDeprBlockTot10Percent_DCGB() As String
DefaultDeprBlockTot10Percent_DCGB = "0"
End Function
Function DefaultDeprBlockTot100Percent_DCGB() As String
DefaultDeprBlockTot100Percent_DCGB = "0"
End Function
Function DefaultTotBuildng_DCGB() As String
DefaultTotBuildng_DCGB = "0"
End Function
Function DefaultFurnitureSummary_DCG() As String
DefaultFurnitureSummary_DCG = "0"
End Function
Function DefaultIntangibleAssetSummary_DCG() As String
DefaultIntangibleAssetSummary_DCG = "0"
End Function
Function DefaultShipsSummary_DCG() As String
DefaultShipsSummary_DCG = "0"
End Function
Function DefaultTotalDepreciation_DCG() As String
DefaultTotalDepreciation_DCG = "0"
End Function



