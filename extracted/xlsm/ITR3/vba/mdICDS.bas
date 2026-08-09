Attribute VB_Name = "mdICDS"
'Malli-----------------
Public ICDS_InAccPolicies As Variant
Public ICDS_InValuationInv As Variant
Public ICDS_InConsContracts As Variant
Public ICDS_InRevenueRecog As Variant
Public ICDS_InTangibleFixAssests As Variant
Public ICDS_InChngRates As Variant
Public ICDS_InGovgrants As Variant
Public ICDS_InSecurities As Variant
Public ICDS_InBorrowingCosts As Variant
Public ICDS_InProvLiability As Variant
Public ICDS_InTotal, ICDS_DeTotal1 As Variant


Public ICDS_DeAccPolicies As Variant
Public ICDS_DeValuationInv As Variant
Public ICDS_DeConsContracts As Variant
Public ICDS_DeRevenueRecog As Variant
Public ICDS_DeTangibleFixAssests As Variant
Public ICDS_DeChngRates As Variant
Public ICDS_DeGovgrants As Variant
Public ICDS_DeSecurities As Variant
Public ICDS_DeBorrowingCosts As Variant
Public ICDS_DeProvLiability As Variant
 
'Malli-----------------





Public ICDS_AccPolicies As Variant
Public ICDS_ValuationInv As Variant
Public ICDS_ConsContracts As Variant
Public ICDS_RevenueRecog As Variant
Public ICDS_TangibleFixAssests As Variant
Public ICDS_ChngRates As Variant
Public ICDS_Govgrants As Variant
Public ICDS_Securities As Variant
Public ICDS_BorrowingCosts As Variant
Public ICDS_ProvLiability As Variant
Public ICDS_Total, ICDS_Totalb As Variant


Sub ValidateICDS_Click()
    ValidateICDS
    fmsgboxStatus "Sheet ICDS is Ok" ', vbOKOnly, "ITR-3"
    CloseMsg
End Sub

Sub ValidateICDS()
 ICDS_AccPolicies = Sheet42.Range("ICDS.AccPolicies").value
 ICDS_ValuationInv = Sheet42.Range("ICDS.ValuationInv").value
 ICDS_ConsContracts = Sheet42.Range("ICDS.ConsContracts").value
 ICDS_RevenueRecog = Sheet42.Range("ICDS.RevenueRecog").value
 ICDS_TangibleFixAssests = Sheet42.Range("ICDS.TangibleFixAssests").value
 ICDS_ChngRates = Sheet42.Range("ICDS.ChngRates").value
 ICDS_Govgrants = Sheet42.Range("ICDS.Govgrants").value
 ICDS_Securities = Sheet42.Range("ICDS.Securities").value
 ICDS_BorrowingCosts = Sheet42.Range("ICDS.BorrowingCosts").value
 ICDS_ProvLiability = Sheet42.Range("ICDS.ProvLiability").value

'ICDS_E6_AY_2023_24 OLd
' ICDS_Total = Sheet42.Range("ICDS.Total").value
' ICDS_Totalb = Sheet42.Range("ICDS.Totalb").value
 'Malli-----------------
 
 'ICDS_C3_AY2024_25 Added
 
 ICDS_InAccPolicies = Sheet42.Range("ICDS.InAccPolicies").value
 ICDS_InValuationInv = Sheet42.Range("ICDS.InValuationInv").value
 ICDS_InConsContracts = Sheet42.Range("ICDS.InConsContracts").value
 ICDS_InRevenueRecog = Sheet42.Range("ICDS.InRevenueRecog").value
 ICDS_InTangibleFixAssests = Sheet42.Range("ICDS.InTangibleFixAssests").value
 ICDS_InChngRates = Sheet42.Range("ICDS.InChngRates").value
 ICDS_InGovgrants = Sheet42.Range("ICDS.InGovgrants").value
 ICDS_InSecurities = Sheet42.Range("ICDS.InSecurities").value
 ICDS_InBorrowingCosts = Sheet42.Range("ICDS.InBorrowingCosts").value
 ICDS_InProvLiability = Sheet42.Range("ICDS.InProvLiability").value
 
  
 
 'ICDS_C4_AY_2024_25 Added
 ICDS_DeAccPolicies = Sheet42.Range("ICDS.DeAccPolicies").value
 ICDS_DeValuationInv = Sheet42.Range("ICDS.DeValuationInv").value
 ICDS_DeConsContracts = Sheet42.Range("ICDS.DeConsContracts").value
 ICDS_DeRevenueRecog = Sheet42.Range("ICDS.DeRevenueRecog").value
 ICDS_DeTangibleFixAssests = Sheet42.Range("ICDS.DeTangibleFixAssests").value
 ICDS_DeChngRates = Sheet42.Range("ICDS.DeChngRates").value
 ICDS_DeGovgrants = Sheet42.Range("ICDS.DeGovgrants").value
 ICDS_DeSecurities = Sheet42.Range("ICDS.DeSecurities").value
 ICDS_DeBorrowingCosts = Sheet42.Range("ICDS.DeBorrowingCosts").value
 ICDS_DeProvLiability = Sheet42.Range("ICDS.DeProvLiability").value
 
 ICDS_InTotal = Sheet42.Range("ICDS.Total").value
 ICDS_DeTotal1 = Sheet42.Range("ICDS.DeTotal1").value
 
 
' If Len(Abs(ICDS_Total)) > 14 Then
'    Sheet42.Activate
'    fmsgboxStatus "Schedule ICDS. Total Net effect Cannot be greater than 14 digits"
'    CloseMsg
' End If
'
'
' If Len((ICDS_Totalb)) > 14 Then
'    Sheet42.Activate
'    fmsgboxStatus "Schedule ICDS. Total Negative Net effect Cannot be greater than 14 digits"
'    CloseMsg
' End If

'ICDS_C6_AY_2024_25  Change
 If Len(Abs(ICDS_InTotal)) > 14 Then
    Sheet42.Activate
    fmsgboxStatus "Schedule ICDS. Total Net effect Cannot be greater than 14 digits"
    CloseMsg
 End If
 
  If Len(Abs(ICDS_DeTotal1)) > 14 Then
    Sheet42.Activate
     fmsgboxStatus "Schedule ICDS. Total Net effect Cannot be greater than 14 digits"
   CloseMsg
 End If

End Sub
