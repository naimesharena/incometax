Attribute VB_Name = "mdICDS"
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
 ICDS_Total = Sheet42.Range("ICDS.Total").value
 ICDS_Totalb = Sheet42.Range("ICDS.Totalb").value
 
 If Len(Abs(ICDS_Total)) > 14 Then
    Sheet42.Activate
    fmsgboxStatus "Schedule ICDS. Total Net effect Cannot be greater than 14 digits"
    CloseMsg
 End If
 
 
 If Len((ICDS_Totalb)) > 14 Then
    Sheet42.Activate
    fmsgboxStatus "Schedule ICDS. Total Negative Net effect Cannot be greater than 14 digits"
    CloseMsg
 End If
End Sub
