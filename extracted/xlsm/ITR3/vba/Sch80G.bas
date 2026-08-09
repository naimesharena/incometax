Attribute VB_Name = "Sch80G"
Option Explicit

Dim tempPan11, tempPan111 As Variant

Public TotDon50PercentNoApprReqd_PerYES10080G As Variant
Public TotElig50PercentNoApprReqd_PerYES10080G As Variant

Public TotDon50PercentNoApprReqd_PerNO5080G As Variant
Public TotElig50PercentNoApprReqd_PerNO5080G As Variant

Public TotDon50PercentApprReqd_Per5080G As Variant
Public TotElig50PercentApprReqd_Per5080G As Variant

Public errmsg80G As Variant

Public DoneeName_Per10080G As Variant
Public AddrDetail_Per10080G As Variant
Public CityOrTownOrDistrict_Per10080G As Variant
Public StateCode_Per10080G As Variant
Public PinCode_Per10080G As Variant
Public DoneePAN_Per10080G As Variant
Public DonationAmt_Per10080G As Variant
Public EligibleAmt_Per10080G As Variant
Public TotDon100Percent_Per10080G As Variant
Public TotElig100Percent_Per10080G As Variant
Public DonationAmt_Per10080G_InCash As Variant

Public DoneeName_PerNO5080G As Variant
Public AddrDetail_PerNO5080G As Variant
Public CityOrTownOrDistrict_PerNO5080G As Variant
Public StateCode_PerNO5080G As Variant
Public PinCode_PerNO5080G As Variant
Public DoneePAN_PerNO5080G As Variant
Public DonationAmt_PerNO5080G As Variant
Public EligibleAmt_PerNO5080G As Variant
Public TotDon100Percent_PerNO5080G As Variant
Public TotElig100Percent_PerNO5080G As Variant
Public DonationAmt_PerNO5080G_InCash As Variant


Public DoneeName_PerYES10080G As Variant
Public AddrDetail_PerYES10080G As Variant
Public CityOrTownOrDistrict_PerYES10080G As Variant
Public StateCode_PerYES10080G As Variant
Public PinCode_PerYES10080G As Variant
Public DoneePAN_PerYES10080G As Variant
Public DonationAmt_PerYES10080G As Variant
Public EligibleAmt_PerYES10080G As Variant
Public TotDon100Percent_PerYES10080G As Variant
Public TotElig100Percent_PerYES10080G As Variant
Public DonationAmt_PerYES10080G_InCash As Variant


Public DoneeName_Per5080G As Variant
Public AddrDetail_Per5080G As Variant
Public CityOrTownOrDistrict_Per5080G As Variant
Public StateCode_Per5080G As Variant
Public PinCode_Per5080G As Variant
Public DoneePAN_Per5080G As Variant
Public DonationAmt_Per5080G As Variant
Public EligibleAmt_Per5080G As Variant
Public TotDon100Percent_Per5080G As Variant
Public TotElig100Percent_Per5080G As Variant
Public DonationAmt_Per5080G_InCash As Variant

Public ChequeNumber_80GA As Variant
Public BankIFSC_80GA As Variant
Public Donation_other_80GA As Variant

Public ChequeNumber_80GB As Variant
Public BankIFSC_80GB As Variant
Public Donation_other_80GB As Variant

Public ChequeNumber_80GC As Variant
Public BankIFSC_80GC As Variant
Public Donation_other_80GC As Variant

Public ChequeNumber_80GD As Variant
Public BankIFSC_80GD As Variant
Public Donation_other_80GD As Variant

Public rngname_Per10080G2 As Variant
Public end_Per10080G2, end_Per10080G3, end_Per10080G4, end_Per10080G5, end_Per10080G6, end_Per10080G7, end_Per10080G8, end_Per10080G9, end_Per10080G10 As Variant

Public rngname_Per10080G As Variant
Public end_Per10080G As Variant
Dim msgValidateSheetPer10080G As String

Public rngname_PerNO5080G2 As Variant
Public end_PerNO5080G2, end_PerNO5080G3, end_PerNO5080G4, end_PerNO5080G5, end_PerNO5080G6, end_PerNO5080G7, end_PerNO5080G8, end_PerNO5080G9, end_PerNO5080G10 As Variant

Public rngname_PerNO5080G As Variant
Public end_PerNO5080G As Variant
Dim msgValidateSheetPerNO5080G As String

Public rngname_PerYES10080G2 As Variant
Public end_PerYES10080G2, end_PerYES10080G3, end_PerYES10080G4, end_PerYES10080G5, end_PerYES10080G6, end_PerYES10080G7, end_PerYES10080G8, end_PerYES10080G9, end_PerYES10080G10 As Variant

Public rngname_PerYES10080G As Variant
Public end_PerYES10080G As Variant
Dim msgValidateSheetPerYES10080G As String

Public rngname_Per5080G2 As Variant
Public end_Per5080G2, end_Per5080G3, end_Per5080G4, end_Per5080G5, end_Per5080G6, end_Per5080G7, end_Per5080G8, end_Per5080G9, end_Per5080G10 As Variant

Public rngname_Per5080G As Variant
Public end_Per5080G As Variant
Dim msgValidateSheetPer5080G As String

Public TotalDonationsUs80G_Per5080G As Variant
Public TotalEligibleDonationsUs80G_Per5080G As Variant

'added by Chetan C M on 14/10/2025 SIT-SIT-91306 (TRC-628156)
Public Total_DonationAmt_Per10080G As Variant
Public Total_EligibleAmt_Per10080G As Variant
Public Total_TotalDonationAmt_80G_A As Variant
Public Total_EligibleAmt_80G_A As Variant

Sub AddRows_Per10080G()
    Dim numberofrows As Long
    setTblinfo_Per10080G
    SelectLastRow ("Per10080G.DoneeName")
    numberofrows = InsertRowsAndFillFormulas80G()
    Call ExendRangeNameToTable(numberofrows, rngname_Per10080G)
End Sub

Sub setTblinfo_Per10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.DoneeName").count
    Set rangecells = Sheet20.Range("Per10080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G = ccount
'    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationInCash;Per10080G.DonationAmt;Per10080G.TotalDonationAmt;Per10080G.EligibleAmt;"
    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationInCash;Per10080G.DonationAmt;Per10080G.TotalDonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"

End Sub

Sub setTblinfo_Per10080G2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.DoneePAN").count
    Set rangecells = Sheet20.Range("Per10080G.DoneePAN").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G2 = ccount
'    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
     rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"
End Sub

'new set table for table 80GA
Sub setTblinfo_Per10080G3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.AddrDetail").count
    Set rangecells = Sheet20.Range("Per10080G.AddrDetail").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G3 = ccount
'    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationInCash;Per10080G.DonationAmt;Per10080G.TotalDonationAmt;Per10080G.EligibleAmt;"
    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationInCash;Per10080G.DonationAmt;Per10080G.TotalDonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"
End Sub


Sub setTblinfo_Per10080G4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.CityOrTownOrDistrict").count
    Set rangecells = Sheet20.Range("Per10080G.CityOrTownOrDistrict").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G4 = ccount
'    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
     rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"

End Sub

Sub setTblinfo_Per10080G5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.StateCode").count
    Set rangecells = Sheet20.Range("Per10080G.StateCode").Cells
    
    For mIntCtr = 1 To mIntCells
         If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G5 = ccount
'    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
     rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"

End Sub

Sub setTblinfo_Per10080G6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.PinCode").count
    Set rangecells = Sheet20.Range("Per10080G.PinCode").Cells
    
    For mIntCtr = 1 To mIntCells
         If Not (rangecells.item(mIntCtr).value = "") Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G6 = ccount
'    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"
End Sub

Sub setTblinfo_Per10080G7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.DonationInCash").count
    Set rangecells = Sheet20.Range("Per10080G.DonationInCash").Cells
    
    For mIntCtr = 1 To mIntCells
'         If Not (rangecells.item(mIntCtr).value = "") Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026
            ccount = ccount + 1
        End If
    Next
    end_Per10080G7 = ccount
'    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"

End Sub

Sub setTblinfo_Per10080G8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per10080G.DonationAmt").count
    Set rangecells = Sheet20.Range("Per10080G.DonationAmt").Cells
    
    For mIntCtr = 1 To mIntCells
'         If Not (rangecells.item(mIntCtr).value = "") Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026
            ccount = ccount + 1
        End If
    Next
    end_Per10080G8 = ccount
'    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;"
    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"
End Sub
'end

Sub setTblinfo_Per10080G9()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Chequeno_80GA").count
    Set rangecells = Sheet20.Range("Chequeno_80GA").Cells
    
    For mIntCtr = 1 To mIntCells
         If Not (rangecells.item(mIntCtr).value = "") Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G9 = ccount
    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"
End Sub
Sub setTblinfo_Per10080G10()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("IFSC_80GA").count
    Set rangecells = Sheet20.Range("IFSC_80GA").Cells
    
    For mIntCtr = 1 To mIntCells
         If Not (rangecells.item(mIntCtr).value = "") Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G10 = ccount
    rngname_Per10080G2 = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.EligibleAmt;Chequeno_80GA;IFSC_80GA;"
End Sub

Sub AddRows_PerNO5080G()
    Dim numberofrows As Long
    setTblinfo_PerNO5080G
    SelectLastRow ("PerNO5080G.DoneeName")
    numberofrows = InsertRowsAndFillFormulas80G()
    Call ExendRangeNameToTable(numberofrows, rngname_PerNO5080G)
End Sub

Sub setTblinfo_PerNO5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.DoneeName").count
    Set rangecells = Sheet20.Range("PerNO5080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G = ccount
'    rngname_PerNO5080G = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationInCash;PerNO5080G.DonationAmt;PerNO5080G.TotalDonationAmt;PerNO5080G.EligibleAmt;"
    rngname_PerNO5080G = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationInCash;PerNO5080G.DonationAmt;PerNO5080G.TotalDonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB"

End Sub

Sub setTblinfo_PerNO5080G2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.DoneePAN").count
    Set rangecells = Sheet20.Range("PerNO5080G.DoneePAN").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G2 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub
' new set table for Table 80G B

Sub setTblinfo_PerNO5080G3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.AddrDetail").count
    Set rangecells = Sheet20.Range("PerNO5080G.AddrDetail").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G3 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

Sub setTblinfo_PerNO5080G4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.CityOrTownOrDistrict").count
    Set rangecells = Sheet20.Range("PerNO5080G.CityOrTownOrDistrict").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G4 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

Sub setTblinfo_PerNO5080G5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.StateCode").count
    Set rangecells = Sheet20.Range("PerNO5080G.StateCode").Cells
    
    For mIntCtr = 1 To mIntCells
       If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G5 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

Sub setTblinfo_PerNO5080G6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.PinCode").count
    Set rangecells = Sheet20.Range("PerNO5080G.PinCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G6 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

Sub setTblinfo_PerNO5080G7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.DonationInCash").count
    Set rangecells = Sheet20.Range("PerNO5080G.DonationInCash").Cells
    
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026

            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G7 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

Sub setTblinfo_PerNO5080G8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerNO5080G.DonationAmt").count
    Set rangecells = Sheet20.Range("PerNO5080G.DonationAmt").Cells
    
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026

            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G8 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub
Sub setTblinfo_PerNO5080G9()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Chequeno_80GB").count
    Set rangecells = Sheet20.Range("Chequeno_80GB").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G9 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

Sub setTblinfo_PerNO5080G10()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("IFSC_80GB").count
    Set rangecells = Sheet20.Range("IFSC_80GB").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G8 = ccount
    rngname_PerNO5080G2 = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.EligibleAmt;Chequeno_80GB;IFSC_80GB;"
End Sub

'end
Sub AddRows_Per5080G()
    Dim numberofrows As Long
    setTblinfo_Per5080G
    SelectLastRow ("Per5080G.DoneeWithPanName")
    numberofrows = InsertRowsAndFillFormulas80G()
    Call ExendRangeNameToTable(numberofrows, rngname_Per5080G)
End Sub

Sub setTblinfo_Per5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.DoneeWithPanName").count
    Set rangecells = Sheet20.Range("Per5080G.DoneeWithPanName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G = ccount
    rngname_Per5080G = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationInCash;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;Per5080G.TotalDonationAmt;Per5080G.DonationInCash1;"
End Sub

Sub setTblinfo_Per5080G2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.DoneePAN").count
    Set rangecells = Sheet20.Range("Per5080G.DoneePAN").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G2 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub
' set table for new

Sub setTblinfo_Per5080G3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.AddrDetail").count
    Set rangecells = Sheet20.Range("Per5080G.AddrDetail").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G3 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub

Sub setTblinfo_Per5080G4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.CityOrTownOrDistrict").count
    Set rangecells = Sheet20.Range("Per5080G.CityOrTownOrDistrict").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G4 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub

Sub setTblinfo_Per5080G5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.StateCode").count
    Set rangecells = Sheet20.Range("Per5080G.StateCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G5 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub

Sub setTblinfo_Per5080G6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.PinCode").count
    Set rangecells = Sheet20.Range("Per5080G.PinCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G6 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub

Sub setTblinfo_Per5080G7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.DonationInCash").count
    Set rangecells = Sheet20.Range("Per5080G.DonationInCash").Cells
    
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026

            ccount = ccount + 1
        End If
    Next
    end_Per5080G7 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub

Sub setTblinfo_Per5080G8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Per5080G.DonationAmt").count
    Set rangecells = Sheet20.Range("Per5080G.DonationAmt").Cells
    
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026

            ccount = ccount + 1
        End If
    Next
    end_Per5080G8 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub
'Added by Aavula for AY 2026-27
Sub setTblinfo_Per5080G9()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Chequeno_80GD").count
    Set rangecells = Sheet20.Range("Chequeno_80GD").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G9 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub

'Added by Aavula for AY 2026-27
Sub setTblinfo_Per5080G10()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("IFSC_80GD").count
    Set rangecells = Sheet20.Range("IFSC_80GD").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G10 = ccount
    rngname_Per5080G2 = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.EligibleAmt;Chequeno_80GD;IFSC_80GD;"
End Sub
'end
Sub AddRows_PerYES10080G()
    Dim numberofrows As Long
    setTblinfo_PerYES10080G
    SelectLastRow ("PerYES10080G.DoneeWithPanName")
    numberofrows = InsertRowsAndFillFormulas80G()
    Call ExendRangeNameToTable(numberofrows, rngname_PerYES10080G)
End Sub

Sub setTblinfo_PerYES10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.DoneeWithPanName").count
    Set rangecells = Sheet20.Range("PerYES10080G.DoneeWithPanName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G = ccount
'    rngname_PerYES10080G = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationInCash;PerYES10080G.DonationAmt;PerYES10080G.TotalDonationAmt;PerYES10080G.EligibleAmt;"
    rngname_PerYES10080G = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationInCash;PerYES10080G.DonationAmt;PerYES10080G.TotalDonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"

End Sub

Sub setTblinfo_PerYES10080G2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.DoneePAN").count
    Set rangecells = Sheet20.Range("PerYES10080G.DoneePAN").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G2 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
 
 'set new table for Table c
 Sub setTblinfo_PerYES10080G3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.AddrDetail").count
    Set rangecells = Sheet20.Range("PerYES10080G.AddrDetail").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G3 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
 Sub setTblinfo_PerYES10080G4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.CityOrTownOrDistrict").count
    Set rangecells = Sheet20.Range("PerYES10080G.CityOrTownOrDistrict").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G4 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
  Sub setTblinfo_PerYES10080G5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.StateCode").count
    Set rangecells = Sheet20.Range("PerYES10080G.StateCode").Cells
    
    For mIntCtr = 1 To mIntCells
       If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G5 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
 Sub setTblinfo_PerYES10080G6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.PinCode").count
    Set rangecells = Sheet20.Range("PerYES10080G.PinCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G6 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
 Sub setTblinfo_PerYES10080G7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.DonationInCash").count
    Set rangecells = Sheet20.Range("PerYES10080G.DonationInCash").Cells
    
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026

            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G7 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
  Sub setTblinfo_PerYES10080G8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("PerYES10080G.DonationAmt").count
    Set rangecells = Sheet20.Range("PerYES10080G.DonationAmt").Cells
    
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
         If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then 'Added by Naresh on 23/06/2026

            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G8 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 'Added by Aavula for AY 2026-27
   Sub setTblinfo_PerYES10080G9()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("Chequeno_80GC").count
    Set rangecells = Sheet20.Range("Chequeno_80GC").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G9 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
    'Added by Aavula for AY 2026-27
   Sub setTblinfo_PerYES10080G10()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet20.Range("IFSC_80GC").count
    Set rangecells = Sheet20.Range("IFSC_80GC").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G10 = ccount
    rngname_PerYES10080G2 = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.EligibleAmt;Chequeno_80GC;IFSC_80GC;"
 End Sub
 
 'end
 
Sub ValidateSheet80GClick()
ValidateSheet80G
fmsgboxStatus "Sheet 80G is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheet80G()
    
    tempPan11 = Sheet1.Range("sheet1.PAN").value
    tempPan111 = Sheet52.Range("sheet9.PAN").value
    
    subProcCaption = "Validating 80G A"
    If Not ValidatesheetPer10080G Then
        Sheet20.Activate
        fmsgboxsmall "Sch 80G-A:" & Chr(13) & msgValidateSheetPer10080G ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    subProcCaption = "Validating 80G C"
    If Not ValidatesheetPerYES10080G Then
        Sheet20.Activate
        fmsgboxsmall "Sch 80G-C:" & Chr(13) & msgValidateSheetPerYES10080G ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    subProcCaption = "Validating 80G B"
    If Not ValidatesheetPerNO5080G Then
        Sheet20.Activate
        fmsgboxsmall "Sch 80G-B:" & Chr(13) & msgValidateSheetPerNO5080G ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    subProcCaption = "Validating 80G D"
    If Not ValidatesheetPer5080G Then
        Sheet20.Activate
        fmsgboxsmall "Sch 80G-D:" & Chr(13) & msgValidateSheetPer5080G ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If

End Sub

Function ValidatesheetPer10080G() As Boolean
    ValidatesheetPer10080G = True
    setTblinfo_Per10080G
    setTblinfo_Per10080G2
    setTblinfo_Per10080G3
    setTblinfo_Per10080G4
    setTblinfo_Per10080G5
    setTblinfo_Per10080G6
    setTblinfo_Per10080G7
    setTblinfo_Per10080G8
    
    setTblinfo_Per10080G9 'Added by Aavula for AY 2026-27
    setTblinfo_Per10080G10 'Added by Aavula for AY 2026-27
    
'    end_Per10080G = WorksheetFunction.Max(0, end_Per10080G, end_Per10080G2, end_Per10080G3, end_Per10080G4, end_Per10080G5, end_Per10080G6, end_Per10080G7, end_Per10080G8)
    end_Per10080G = WorksheetFunction.Max(0, end_Per10080G, end_Per10080G2, end_Per10080G3, end_Per10080G4, end_Per10080G5, end_Per10080G6, end_Per10080G7, end_Per10080G8, end_Per10080G9, end_Per10080G10)
  
    If Not ValidateDoneeName_Per10080G() Then ValidatesheetPer10080G = False

        If Not ValidateAddrDetail_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateCityOrTownOrDistrict_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateStateCode_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidatePinCode_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateDoneePAN_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateDonationAmt_Per10080G() Then ValidatesheetPer10080G = False
        
         'Added by Aavula for AY 2026-27
        If Not ValidateChequeNumber_80G() Then ValidatesheetPer10080G = False
        If Not ValidateIFSC_80G Then ValidatesheetPer10080G = False
        ''''''''''''''''
        If Not ValidateEligibleAmt_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateTotDon100Percent_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateTotElig100Percent_Per10080G() Then ValidatesheetPer10080G = False
        If Not ValidateNonEditableFields_Per10080G() Then ValidatesheetPer10080G = False    'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147)
        If Not ValidateTotalFields_80G() Then ValidatesheetPer10080G = False    'added by Chetan C M on 14/10/2025 SIT-SIT-91306 (TRC-628156)
        
        
        
End Function

'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147)
'start--
Function ValidateNonEditableFields_Per10080G() As Boolean
    ValidateNonEditableFields_Per10080G = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    
    Set rangecells = Sheet20.Range("Per10080G.TotalDonationAmt").Cells
    Set rangecells1 = Sheet20.Range("Per10080G.EligibleAmt").Cells
    
    Dim i As Long
    ReDim TotalDonationAmt_80G_A(end_Per10080G)
    ReDim EligibleAmt_80G_A(end_Per10080G)
   
    
    For i = 1 To end_Per10080G
    
        TotalDonationAmt_80G_A(i) = rangecells.item(i).value
        EligibleAmt_80G_A(i) = rangecells1.item(i).value
        
        If Len(TotalDonationAmt_80G_A(i)) > 14 Then
           msgbox_Per10080G "* Total Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_Per10080G = False
        End If
        
        If Len(EligibleAmt_80G_A(i)) > 14 Then
           msgbox_Per10080G "* Eligible Amount of Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_Per10080G = False
        End If
        
    Next

End Function
'--end

 'added by Chetan C M on 14/10/2025 SIT-SIT-91306 (TRC-628156)
'start--
Function ValidateTotalFields_80G() As Boolean
 ValidateTotalFields_80G = True

 Total_DonationAmt_Per10080G = Sheet20.Range("Per10080G.TotDon100Percent_InCash").value
 Total_EligibleAmt_Per10080G = Sheet20.Range("Per10080G.TotDon100Percent").value
 Total_TotalDonationAmt_80G_A = Sheet20.Range("Per10080G.TotDon100Percent_Total").value
 Total_EligibleAmt_80G_A = Sheet20.Range("Per10080G.TotElig100Percent").value

 
   
    If Len(Total_DonationAmt_Per10080G) > 14 Then
    
       msgbox_Per10080G "* Total of Donation in cash cannot be greater than 14 digits in Schedule 80G" & Chr(13)
    
       ValidateTotalFields_80G = False
       
    End If
    
    If Len(Total_EligibleAmt_Per10080G) > 14 Then
    
       msgbox_Per10080G "* Total of Donation in other mode cannot be greater than 14 digits in Schedule 80G" & Chr(13)
    
       ValidateTotalFields_80G = False
       
    End If
    
    If Len(Total_TotalDonationAmt_80G_A) > 14 Then
    
       msgbox_Per10080G "* Total of Total Donation cannot be greater than 14 digits in Schedule 80G" & Chr(13)
    
       ValidateTotalFields_80G = False
       
    End If
    
    If Len(Total_EligibleAmt_80G_A) > 14 Then
    
       msgbox_Per10080G "* Total of Eligible Amount of Donation cannot be greater than 14 digits in Schedule 80G" & Chr(13)
    
       ValidateTotalFields_80G = False
       
    End If
    
End Function
'--end

Function ValidatesheetPerYES10080G() As Boolean
    ValidatesheetPerYES10080G = True
    Dim i As Long
    
    setTblinfo_PerYES10080G
    setTblinfo_PerYES10080G2
    setTblinfo_PerYES10080G3
    setTblinfo_PerYES10080G4
    setTblinfo_PerYES10080G5
    setTblinfo_PerYES10080G6
    setTblinfo_PerYES10080G7
    setTblinfo_PerYES10080G8
    
   'Added by Aavula for AY 2026-27
    setTblinfo_PerYES10080G9
    setTblinfo_PerYES10080G10
    '''''
    
    end_PerYES10080G = WorksheetFunction.Max(0, end_PerYES10080G, end_PerYES10080G2, end_PerYES10080G3, end_PerYES10080G4, end_PerYES10080G5, end_PerYES10080G6, end_PerYES10080G7, end_PerYES10080G8, end_PerYES10080G9, end_PerYES10080G10)
        
        If Not ValidateDoneeName_PerYES10080G() Then ValidatesheetPerYES10080G = False
        
        If Not ValidateAddrDetail_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateCityOrTownOrDistrict_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateStateCode_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidatePinCode_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateDoneePAN_PerYES10080G() Then ValidatesheetPerYES10080G = False
        
         'Added by Aavula for AY 2026-27
        If Not ValidateChequeNumber_80GC() Then ValidatesheetPerYES10080G = False
        If Not ValidateIFSC_80GC Then ValidatesheetPerYES10080G = False
        '''''''''''
        
        If Not ValidateDonationAmt_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateEligibleAmt_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateTotDon50PercentNoApprReqd_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateTotElig50PercentNoApprReqd_PerYES10080G() Then ValidatesheetPerYES10080G = False
        If Not ValidateNonEditableFields_PerYES10080G() Then ValidatesheetPerYES10080G = False    'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147,TRC-628156)

End Function

'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147)
'start--
Function ValidateNonEditableFields_PerYES10080G() As Boolean
    ValidateNonEditableFields_PerYES10080G = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    
    Set rangecells = Sheet20.Range("PerYES10080G.TotalDonationAmt").Cells
    Set rangecells1 = Sheet20.Range("PerYES10080G.EligibleAmt").Cells
    
    
    Dim i As Long
    ReDim TotalDonationAmt_80G_C(end_PerYES10080G)
    ReDim EligibleAmt_80G_C(end_PerYES10080G)
   
    
    For i = 1 To end_PerYES10080G
    
        TotalDonationAmt_80G_C(i) = rangecells.item(i).value
        EligibleAmt_80G_C(i) = rangecells1.item(i).value
        
        If Len(TotalDonationAmt_80G_C(i)) > 14 Then
           msgbox_PerYES10080G "* Total Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_PerYES10080G = False
        End If
        
        If Len(EligibleAmt_80G_C(i)) > 14 Then
           msgbox_PerYES10080G "* Eligible Amount of Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_PerYES10080G = False
        End If
        
    Next

End Function
'--end

Function ValidatesheetPerNO5080G() As Boolean
    ValidatesheetPerNO5080G = True
    
    setTblinfo_PerNO5080G
    setTblinfo_PerNO5080G2
    setTblinfo_PerNO5080G3
    setTblinfo_PerNO5080G4
    setTblinfo_PerNO5080G5
    setTblinfo_PerNO5080G6
    setTblinfo_PerNO5080G7
    setTblinfo_PerNO5080G8
    
    setTblinfo_PerNO5080G9  'Added by Aavula for AY 2026-27
    setTblinfo_PerNO5080G10 'Added by Aavula for AY 2026-27
    
    end_PerNO5080G = WorksheetFunction.Max(0, end_PerNO5080G, end_PerNO5080G2, end_PerNO5080G3, end_PerNO5080G4, end_PerNO5080G5, end_PerNO5080G6, end_PerNO5080G7, end_PerNO5080G8, end_PerNO5080G9, end_PerNO5080G10)
    
    If Not ValidateDoneeName_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateAddrDetail_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateCityOrTownOrDistrict_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateStateCode_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidatePinCode_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateDoneePAN_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateDonationAmt_PerNO5080G() Then ValidatesheetPerNO5080G = False
     'Added by Aavula for AY 2026-27
    If Not ValidateChequeNumber_80GB() Then ValidatesheetPerNO5080G = False
    If Not ValidateIFSC_80GB Then ValidatesheetPerNO5080G = False
    '''''''''''''
    If Not ValidateEligibleAmt_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateTotDon50PercentNoApprReqd_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateTotElig50PercentNoApprReqd_PerNO5080G() Then ValidatesheetPerNO5080G = False
    If Not ValidateNonEditableFields_PerNO5080G() Then ValidatesheetPerNO5080G = False    'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147,TRC-628156)
   
End Function

'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147)
'start--
Function ValidateNonEditableFields_PerNO5080G() As Boolean
    ValidateNonEditableFields_PerNO5080G = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    
    Set rangecells = Sheet20.Range("PerNO5080G.TotalDonationAmt").Cells
    Set rangecells1 = Sheet20.Range("PerNO5080G.EligibleAmt").Cells
    
    
    Dim i As Long
    ReDim TotalDonationAmt_80G_B(end_PerNO5080G)
    ReDim EligibleAmt_80G_B(end_PerNO5080G)
   
    
    For i = 1 To end_PerNO5080G
    
        TotalDonationAmt_80G_B(i) = rangecells.item(i).value
        EligibleAmt_80G_B(i) = rangecells1.item(i).value
        
        If Len(TotalDonationAmt_80G_B(i)) > 14 Then
           msgbox_PerNO5080G "* Total Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_PerNO5080G = False
        End If
        
        If Len(EligibleAmt_80G_B(i)) > 14 Then
           msgbox_PerNO5080G "* Eligible Amount of Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_PerNO5080G = False
        End If
        
    Next

End Function
'--end

Function ValidatesheetPer5080G() As Boolean
    ValidatesheetPer5080G = True
    setTblinfo_Per5080G
    setTblinfo_Per5080G2
    setTblinfo_Per5080G3
    setTblinfo_Per5080G4
    setTblinfo_Per5080G5
    setTblinfo_Per5080G6
    setTblinfo_Per5080G7
    setTblinfo_Per5080G8
    
    setTblinfo_Per5080G9 'Added by Aavula for AY 2026-27
    setTblinfo_Per5080G10 'Added by Aavula for AY 2026-27
    
    end_Per5080G = WorksheetFunction.Max(0, end_Per5080G, end_Per5080G2, end_Per5080G3, end_Per5080G4, end_Per5080G5, end_Per5080G6, end_Per5080G7, end_Per5080G8, end_Per5080G9, end_Per5080G10)
        If Not ValidateDoneeWithPanName_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateDoneePAN_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateAddrDetail_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateCityOrTownOrDistrict_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateStateCode_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidatePinCode_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateDonationAmt_Per5080G() Then ValidatesheetPer5080G = False
        
         'Added by Aavula for AY 2026-27
        If Not ValidateChequeNumber_80GD() Then ValidatesheetPer5080G = False
        If Not ValidateIFSC_80GD() Then ValidatesheetPer5080G = False
        '''''''''''
        If Not ValidateEligibleAmt_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateTotDon50PercentApprReqd_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateTotElig50PercentApprReqd_Per5080G() Then ValidatesheetPer5080G = False
'Change-27.01.2023.101.IDS.2B
        If Not ValidateARNPer5080G() Then ValidatesheetPer5080G = False
'---
'Change-16.02.2023.103.80G.
        If Not ValidatePAN_ARN_80GD() Then ValidatesheetPer5080G = False
            
        If Not ValidateTotalDonationsUs80G_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateTotalEligibleDonationsUs80G_Per5080G() Then ValidatesheetPer5080G = False
        If Not ValidateNonEditableFields_Per5080G() Then ValidatesheetPer5080G = False    'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147,TRC-628156)
    
End Function

'added by Chetan C M on 24/09/2025 SIT-SIT-91306 (TRC-628147)
'start--
Function ValidateNonEditableFields_Per5080G() As Boolean
    ValidateNonEditableFields_Per5080G = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    
    Set rangecells = Sheet20.Range("Per5080G.TotalDonationAmt").Cells
    Set rangecells1 = Sheet20.Range("Per5080G.EligibleAmt").Cells
    
    
    Dim i As Long
    ReDim TotalDonationAmt_80G_D(end_Per5080G)
    ReDim EligibleAmt_80G_D(end_Per5080G)
   
    
    For i = 1 To end_Per5080G
    
        TotalDonationAmt_80G_D(i) = rangecells.item(i).value
        EligibleAmt_80G_D(i) = rangecells1.item(i).value
        
        If Len(TotalDonationAmt_80G_D(i)) > 14 Then
           msgbox_Per5080G "* Total Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_Per5080G = False
        End If
        
        If Len(EligibleAmt_80G_D(i)) > 14 Then
           msgbox_Per5080G "* Eligible Amount of Donation at Sr. No " & i & " cannot exceed 14 digits in Schedule 80G" & Chr(13)
        ValidateNonEditableFields_Per5080G = False
        End If
        
    Next

End Function
'--end

Function ValidateDoneeName_Per10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateDoneeName_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.DoneeName").Cells
    ReDim DoneeName_Per10080G(end_Per10080G)
    noOfProcessSub = end_Per10080G
    For i = 1 To end_Per10080G
        DoneeName_Per10080G(i) = rangecells.item(i).value

        If Not chkCompulsory(DoneeName_Per10080G(i)) Then
'            msgbox_Per10080G ("* Please enter DoneeName at Sr. No  " & i & " in schedule 80G Table A")
            'Jyoti A
            'SIT-69597 Changed by sai
            msgbox_Per10080G ("*""Please enter name of donee in schedule 80G Table A"" at Sr. No " & i & "")
            
            ValidateDoneeName_Per10080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(DoneeName_Per10080G(i)) Then
            msgbox_Per10080G ("* DoneeName at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in schedule 80G Table A")
            ValidateDoneeName_Per10080G = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
 
Function ValidateAddrDetail_Per10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateAddrDetail_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.AddrDetail").Cells
    ReDim AddrDetail_Per10080G(end_Per10080G)
    
    For i = 1 To end_Per10080G

        AddrDetail_Per10080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(AddrDetail_Per10080G(i)) Then
            'msgbox_Per10080G ("* Please enter Address Detail at Sr. No  " & i & "  in schedule 80G Table A")
            
            'Added by Aavula
            msgbox_Per10080G ("* Please enter address of donee in schedule 80G Table A at Sr. No  " & i & "")
            ValidateAddrDetail_Per10080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(AddrDetail_Per10080G(i)) Then
            msgbox_Per10080G ("* Address Detail at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed  in schedule 80G Table A")
            ValidateAddrDetail_Per10080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateCityOrTownOrDistrict_Per10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateCityOrTownOrDistrict_Per10080G = True
   ' setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.CityOrTownOrDistrict").Cells
    ReDim CityOrTownOrDistrict_Per10080G(end_Per10080G)
    For i = 1 To end_Per10080G

        CityOrTownOrDistrict_Per10080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(CityOrTownOrDistrict_Per10080G(i)) Then
            msgbox_Per10080G ("* Please enter CityOrTownOrDistrict at Sr. No  " & i & "   in schedule 80G Table A")
            ValidateCityOrTownOrDistrict_Per10080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CityOrTownOrDistrict_Per10080G(i)) Then
            msgbox_Per10080G ("* CityOrTownOrDistrict at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in schedule 80G Table A")
            ValidateCityOrTownOrDistrict_Per10080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateStateCode_Per10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateStateCode_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.StateCode").Cells
    ReDim StateCode_Per10080G(end_Per10080G)
    For i = 1 To end_Per10080G

        StateCode_Per10080G(i) = rangecells.item(i).value
        StateCode_Per10080G(i) = Mid(StateCode_Per10080G(i), 1, 2)
        If isdropdownblank(StateCode_Per10080G(i)) Then
            msgbox_Per10080G ("* Please select StateCode from the DropDown at Sr. No  " & i & "  in schedule 80G Table A")
            ValidateStateCode_Per10080G = False
            Exit Function
        End If
        
    'Changed By sai on 16/10/2024=============================================================
        If StateCode_Per10080G(i) <> "" Then
        
        Sheets("80G").Activate
        Dim PIN_targetadd, state_targetadd As String
        
        Dim ws_1 As Worksheet
        Set ws_1 = Worksheets("80G")
        
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")

            If Not state_Validation_Level(ws_1, PIN_targetadd, state_targetadd) Then
                     msgbox_hprptfrm ("* Invalid Pincode")
                     Sheet20.Range(Replace(state_targetadd, "H", "I")).value = ""
                     
                   ValidateStateCode_Per10080G = False
                 End If
            End If
        
     '=================================================================================

        
        If (StateCode_Per10080G(i) = "99") Then
            msgbox_Per10080G ("* StateCode at Sr. No  " & i & " is Invaild in schedule 80G Table A")
            ValidateStateCode_Per10080G = False
            Exit Function
        End If
        
    Next
End Function

Function ValidatePinCode_Per10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePinCode_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.PinCode").Cells
    ReDim PinCode_Per10080G(end_Per10080G)
    For i = 1 To end_Per10080G
'        If Len(DoneeName_Per10080G(i)) = 0 Then
'            Exit For
'        End If
        
        PinCode_Per10080G(i) = rangecells.item(i).value
        If Not chkCompulsory(PinCode_Per10080G(i)) Then
            msgbox_Per10080G ("* Please enter PinCode at Sr. No  " & i & "   in schedule 80G Table A")
            ValidatePinCode_Per10080G = False
            Exit Function
        End If
        
        If Not chkNumeric(PinCode_Per10080G(i)) Then
            msgbox_Per10080G ("* PinCode at Sr. No  " & i & "   only digits 0 to 9 allowed  in schedule 80G Table A")
            ValidatePinCode_Per10080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateDoneePAN_Per10080G() As Boolean
    Dim rangecells As Range
    Dim tempPan As String
    Dim i As Long
    
    ValidateDoneePAN_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.DoneePAN").Cells
    ReDim DoneePAN_Per10080G(end_Per10080G)
    For i = 1 To end_Per10080G
'        If Len(DoneeName_Per10080G(i)) = 0 Then
'            Exit For
'        End If
        DoneePAN_Per10080G(i) = rangecells.item(i).value
        
        tempPan = DoneePAN_Per10080G(i)
        If Not chkCompulsory(DoneePAN_Per10080G(i)) Then
'            msgbox_Per10080G ("* Please enter DoneePAN at Sr. No  " & i & "  in schedule 80G Table A")
            msgbox_Per10080G ("* ""Please enter PAN of donee at Sr. No " & i & " in schedule 80G Table A""")
            
            ValidateDoneePAN_Per10080G = False
            Exit Function
        End If
        
        If Not ValidatePAN(tempPan) Then
'            msgbox_Per10080G ("* DoneePAN at Sr. No  " & i & "   is invalid. First 5 alphabets, next 4 digits, then alphabet  in schedule 80G Table A")
            msgbox_Per10080G ("* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr. No " & i & " in schedule 80G Table A.""")

            ValidateDoneePAN_Per10080G = False
            Exit Function
        End If
        
        If tempPan = tempPan11 Or tempPan = tempPan111 Then
            msgbox_Per10080G ("* DoneePAN at Sr. No  " & i & "  Please enter the PAN to whom donation is made  in schedule 80G Table A")
            ValidateDoneePAN_Per10080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonationAmt_Per10080G() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    ValidateDonationAmt_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.DonationAmt").Cells
    Set rangecells1 = Sheet20.Range("Per10080G.DonationInCash").Cells
    ReDim DonationAmt_Per10080G(end_Per10080G)
    ReDim DonationAmt_Per10080G_InCash(end_Per10080G)
    
    For i = 1 To end_Per10080G
'        If Len(DoneeName_Per10080G(i)) = 0 Then
'            Exit For
'        End If
        DonationAmt_Per10080G(i) = rangecells.item(i).value
        DonationAmt_Per10080G_InCash(i) = rangecells1.item(i).value
        
        If (DonationAmt_Per10080G(i) = "" And DonationAmt_Per10080G_InCash(i) = "") Then
           'changed by Chetan C M on 11/09/2025
        'start--
            'msgbox_Per10080G ("* Please enter Donation in Cash or Donation in Other mode at Sr. No  " & i & "   in schedule 80G Table A")
            msgbox_Per10080G ("* 'Enter the amount of donation either in field Donation in cash' or 'Donation in other mode  in schedule 80G Table A'")
        '--end
        ValidateDonationAmt_Per10080G = False
            Exit Function
        End If
        If (DonationAmt_Per10080G(i) < 0) Then
            msgbox_Per10080G ("* 80G.A Amount at Sr. No  " & i & " should be Non negative, no decimal, upto 99,999,999,999,999  in schedule 80G Table A")
            ValidateDonationAmt_Per10080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateEligibleAmt_Per10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateEligibleAmt_Per10080G = True
    'setTblinfo_Per10080G
    
    Set rangecells = Sheet20.Range("Per10080G.EligibleAmt").Cells
    ReDim EligibleAmt_Per10080G(end_Per10080G)
    For i = 1 To end_Per10080G
'        If Len(DoneeName_Per10080G(i)) = 0 Then
'            Exit For
'        End If
        EligibleAmt_Per10080G(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotDon100Percent_Per10080G() As Boolean
ValidateTotDon100Percent_Per10080G = True
TotDon100Percent_Per10080G = Sheet20.Range("Per10080G.TotDon100Percent").value
End Function

Function ValidateTotElig100Percent_Per10080G() As Boolean
ValidateTotElig100Percent_Per10080G = True
TotElig100Percent_Per10080G = Sheet20.Range("Per10080G.TotElig100Percent").value
End Function

Function ValidateDoneeName_PerNO5080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateDoneeName_PerNO5080G = True
    'setTblinfo_PerNO5080G
    Set rangecells = Sheet20.Range("PerNO5080G.DoneeName").Cells
    ReDim DoneeName_PerNO5080G(end_PerNO5080G)
    noOfProcessSub = end_PerNO5080G
    For i = 1 To end_PerNO5080G
        DoneeName_PerNO5080G(i) = rangecells.item(i).value
  
        If Not chkCompulsory(DoneeName_PerNO5080G(i)) Then
'            msgbox_PerNO5080G ("* Please enter DoneeName at Sr. No  " & i & " in schedule 80G Table B")
            'Jyoti B
         'SIT-69597  Changed by sai
            msgbox_PerNO5080G ("*""Please enter name of donee in schedule 80G Table B"" at Sr. No " & i & "")
            
            ValidateDoneeName_PerNO5080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(DoneeName_PerNO5080G(i)) Then
            msgbox_PerNO5080G ("* DoneeName at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in schedule 80G Table B")
            ValidateDoneeName_PerNO5080G = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
 
Function ValidateAddrDetail_PerNO5080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAddrDetail_PerNO5080G = True
    'setTblinfo_PerNO5080G
    
    Set rangecells = Sheet20.Range("PerNO5080G.AddrDetail").Cells
    ReDim AddrDetail_PerNO5080G(end_PerNO5080G)
    
    For i = 1 To end_PerNO5080G

        AddrDetail_PerNO5080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(AddrDetail_PerNO5080G(i)) Then
            'msgbox_PerNO5080G ("* Please enter Address Detail at Sr. No  " & i & "  in schedule 80G Table B")
            
            'Added by Aavula SIT-69596
            msgbox_PerNO5080G ("* Please enter address of donee in schedule 80G Table B at Sr. No  " & i & "")

            ValidateAddrDetail_PerNO5080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(AddrDetail_PerNO5080G(i)) Then
            msgbox_PerNO5080G ("* Address Detail at Sr. No  " & i & "  < > & ' " & Chr(34) & " are not allowed  in schedule 80G Table B")
            ValidateAddrDetail_PerNO5080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateCityOrTownOrDistrict_PerNO5080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateCityOrTownOrDistrict_PerNO5080G = True
    'setTblinfo_PerNO5080G
    
    Set rangecells = Sheet20.Range("PerNO5080G.CityOrTownOrDistrict").Cells
    ReDim CityOrTownOrDistrict_PerNO5080G(end_PerNO5080G)
    
    For i = 1 To end_PerNO5080G
   
        CityOrTownOrDistrict_PerNO5080G(i) = rangecells.item(i).value
        If Not chkCompulsory(CityOrTownOrDistrict_PerNO5080G(i)) Then
'            msgbox_PerNO5080G ("* Please enter CityOrTownOrDistrict at Sr. No  " & i & "   in schedule 80G Table B")
            'Added by Aavula
            msgbox_PerNO5080G ("* Please enter city/town/district of donee in schedule 80G Table B at Sr. No  " & i & "")
            
            
            ValidateCityOrTownOrDistrict_PerNO5080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CityOrTownOrDistrict_PerNO5080G(i)) Then
            msgbox_PerNO5080G ("* CityOrTownOrDistrict at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed  in schedule 80G Table B")
            ValidateCityOrTownOrDistrict_PerNO5080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateStateCode_PerNO5080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateStateCode_PerNO5080G = True
    'setTblinfo_PerNO5080G
    
    Set rangecells = Sheet20.Range("PerNO5080G.StateCode").Cells
    ReDim StateCode_PerNO5080G(end_PerNO5080G)
    
    For i = 1 To end_PerNO5080G
     
        StateCode_PerNO5080G(i) = rangecells.item(i).value
        StateCode_PerNO5080G(i) = Mid(StateCode_PerNO5080G(i), 1, 2)
        If isdropdownblank(StateCode_PerNO5080G(i)) Then
           'changed by Chetan C M on 11/09/2025
        'start--
            'msgbox_PerNO5080G ("* Please select StateCode from the DropDown at Sr. No  " & i & " in schedule 80G Table B")
            msgbox_PerNO5080G ("* Please select state of donee in schedule 80G Table B")
        '--end
        ValidateStateCode_PerNO5080G = False
            Exit Function
        End If
        
        
    'Changed By sai on 16/10/2024====================================================================
        If StateCode_PerNO5080G(i) <> "" Then
        Sheets("80G").Activate
        
        Dim ws_2 As Worksheet
        Set ws_2 = Worksheets("80G")

          Dim PIN_targetadd, state_targetadd As String
      
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
         PIN_targetadd = Replace(state_targetadd, "H", "I")

          
          If Not state_Validation_Level(ws_2, PIN_targetadd, state_targetadd) Then
                    msgbox_PerNO5080G ("* Invalid Pincode")
                     Sheet20.Range(Replace(state_targetadd, "H", "I")).value = ""
                     
                   ValidateStateCode_PerNO5080G = False
                 End If
            End If




     '========================================================================================
        
        
        
        
        
        If (StateCode_PerNO5080G(i) = "99") Then
            msgbox_PerNO5080G ("* StateCode at Sr. No  " & i & "is Invalid  in schedule 80G Table B")
            ValidateStateCode_PerNO5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidatePinCode_PerNO5080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePinCode_PerNO5080G = True
    'setTblinfo_PerNO5080G
    
    Set rangecells = Sheet20.Range("PerNO5080G.PinCode").Cells
    ReDim PinCode_PerNO5080G(end_PerNO5080G)
    For i = 1 To end_PerNO5080G

        PinCode_PerNO5080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(PinCode_PerNO5080G(i)) Then
'            msgbox_PerNO5080G ("* Please enter PinCode at Sr. No  " & i & "   in schedule 80G Table B")
           'SIT-69611 Changed by sai
             msgbox_PerNO5080G ("*""Please enter pin code of donee in schedule 80G Table B"" at Sr. No " & i & "")
             
            ValidatePinCode_PerNO5080G = False
            Exit Function
        End If
        
        If Not chkNumeric(PinCode_PerNO5080G(i)) Then
'            msgbox_PerNO5080G ("* PinCode at Sr. No  " & i & " only digits 0 to 9 allowed in schedule 80G Table B")
            
        'SIT-69611 Changed by sai
             msgbox_PerNO5080G ("*""Please enter a 6 digit valid Pin Code in schedule 80G Table B"" at Sr. No " & i & "")
            
            ValidatePinCode_PerNO5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateDoneePAN_PerNO5080G() As Boolean
    Dim rangecells As Range
    Dim tempPan As String
    Dim i As Long
    
    ValidateDoneePAN_PerNO5080G = True
    'setTblinfo_PerNO5080G
    Set rangecells = Sheet20.Range("PerNO5080G.DoneePAN").Cells
    ReDim DoneePAN_PerNO5080G(end_PerNO5080G)
    
    For i = 1 To end_PerNO5080G

        DoneePAN_PerNO5080G(i) = rangecells.item(i).value
        
        tempPan = DoneePAN_PerNO5080G(i)
        
        If Not chkCompulsory(DoneePAN_PerNO5080G(i)) Then
'            msgbox_PerNO5080G ("* Please enter DoneePAN at Sr. No  " & i & "  in schedule 80G Table B")


          'SIT -69608 Changed by sai
             msgbox_PerNO5080G ("*""Please enter PAN of donee in schedule 80G Table B"" at Sr. No  " & i & "")

            ValidateDoneePAN_PerNO5080G = False
            Exit Function
        End If
        
        If Not ValidatePAN(tempPan) Then
'            msgbox_PerNO5080G ("* DoneePAN at Sr. No  " & i & "    is invalid. First 5 alphabets, next 4 digits, then alphabet in schedule 80G Table B")
            
            msgbox_PerNO5080G ("* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule 80G Table B."" at Sr. No  " & i & "")
            
            ValidateDoneePAN_PerNO5080G = False
            Exit Function
        End If
        
        If tempPan = tempPan11 Or tempPan = tempPan111 Then
            msgbox_PerNO5080G ("* DoneePAN at Sr. No  " & i & " Please enter the PAN to whom donation is made in schedule 80G Table B")
            ValidateDoneePAN_PerNO5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonationAmt_PerNO5080G() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    ValidateDonationAmt_PerNO5080G = True
    'setTblinfo_PerNO5080G
    Set rangecells = Sheet20.Range("PerNO5080G.DonationAmt").Cells
    Set rangecells1 = Sheet20.Range("PerNO5080G.DonationInCash").Cells
    ReDim DonationAmt_PerNO5080G(end_PerNO5080G)
    ReDim DonationAmt_PerNO5080G_InCash(end_PerNO5080G)
    
    For i = 1 To end_PerNO5080G
        DonationAmt_PerNO5080G(i) = rangecells.item(i).value
        DonationAmt_PerNO5080G_InCash(i) = rangecells1.item(i).value
        
        If (DonationAmt_PerNO5080G(i) = "" And DonationAmt_PerNO5080G_InCash(i) = "") Then
'            msgbox_PerNO5080G ("* Please enter Donation in Cash or Donation in other mode at Sr. No  " & i & "  in schedule 80G Table B")
                
             'Jyoti SIT-79698
                
             'Added by Aavula 02/08/2024 SIT-69614
             msgbox_PerNO5080G ("* Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode"" in schedule 80G Table B at Sr. No  " & i & "")
            ValidateDonationAmt_PerNO5080G = False
            Exit Function
        End If
        
        If (DonationAmt_PerNO5080G(i) < 0) Then
            msgbox_PerNO5080G ("* Amount at Sr. No  " & i & "   should be Non negative, no decimal, upto 99,999,999,999,999 in schedule 80G Table B")
            ValidateDonationAmt_PerNO5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateEligibleAmt_PerNO5080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateEligibleAmt_PerNO5080G = True
    setTblinfo_PerNO5080G

    Set rangecells = Sheet20.Range("PerNO5080G.EligibleAmt").Cells
    ReDim EligibleAmt_PerNO5080G(end_PerNO5080G)
    
    For i = 1 To end_PerNO5080G
        EligibleAmt_PerNO5080G(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotDon50PercentNoApprReqd_PerNO5080G() As Boolean
    ValidateTotDon50PercentNoApprReqd_PerNO5080G = True
    TotDon50PercentNoApprReqd_PerNO5080G = Sheet20.Range("PerNO5080G.TotDon100Percent").value
End Function

Function ValidateTotElig50PercentNoApprReqd_PerNO5080G() As Boolean
    ValidateTotElig50PercentNoApprReqd_PerNO5080G = True
    TotElig50PercentNoApprReqd_PerNO5080G = Sheet20.Range("PerNO5080G.TotElig100Percent").value
End Function

Function ValidateDoneeName_PerYES10080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
 
    ValidateDoneeName_PerYES10080G = True
'    setTblinfo_PerYES10080G
    
    Set rangecells = Sheet20.Range("PerYES10080G.DoneeWithPanName").Cells
    ReDim DoneeName_PerYES10080G(end_PerYES10080G)
    noOfProcessSub = end_PerYES10080G
    For i = 1 To end_PerYES10080G
        DoneeName_PerYES10080G(i) = rangecells.item(i).value
'        If Len(DoneeName_PerYES10080G(i)) = 0 Then
'            Exit For
'        End If
'
        If Not chkCompulsory(DoneeName_PerYES10080G(i)) Then

' msgbox_PerYES10080G ("* Please enter DoneeName at Sr. No  " & i & "  in schedule 80G Table C")

        'Added by Aavula
            msgbox_PerYES10080G ("* ""Please enter name of donee in schedule 80G Table C"" at Sr. No " & i & "")

            ValidateDoneeName_PerYES10080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(DoneeName_PerYES10080G(i)) Then
            msgbox_PerYES10080G ("* DoneeName at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in schedule 80G Table C")
            ValidateDoneeName_PerYES10080G = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
 
Function ValidateAddrDetail_PerYES10080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
 
    ValidateAddrDetail_PerYES10080G = True
    'setTblinfo_PerYES10080G
    Set rangecells = Sheet20.Range("PerYES10080G.AddrDetail").Cells
    ReDim AddrDetail_PerYES10080G(end_PerYES10080G)
    
    For i = 1 To end_PerYES10080G
'        If Len(DoneeName_PerYES10080G(i)) = 0 Then
'            Exit For
'        End If
        
        AddrDetail_PerYES10080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(AddrDetail_PerYES10080G(i)) Then
'            msgbox_PerYES10080G ("* Please enter Address Detail at Sr. No  " & i & "  in schedule 80G Table C")


        'SIT-69596  Changed by sai
            msgbox_PerYES10080G ("* ""Please enter address of donee in schedule 80G Table C"" at Sr. No  " & i & "")

            ValidateAddrDetail_PerYES10080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(AddrDetail_PerYES10080G(i)) Then
            msgbox_PerYES10080G ("* Address Detail at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in schedule 80G Table C")
            ValidateAddrDetail_PerYES10080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateCityOrTownOrDistrict_PerYES10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateCityOrTownOrDistrict_PerYES10080G = True
    'setTblinfo_PerYES10080G
 
    Set rangecells = Sheet20.Range("PerYES10080G.CityOrTownOrDistrict").Cells
    ReDim CityOrTownOrDistrict_PerYES10080G(end_PerYES10080G)
    
    For i = 1 To end_PerYES10080G
'        If Len(DoneeName_PerYES10080G(i)) = 0 Then
'            Exit For
'        End If
        
        CityOrTownOrDistrict_PerYES10080G(i) = rangecells.item(i).value
'
        If Not chkCompulsory(CityOrTownOrDistrict_PerYES10080G(i)) Then
            msgbox_PerYES10080G ("* ""Please enter CityOrTownOrDistrict in schedule 80G Table C"" at Sr. No  " & i & "")
            ValidateCityOrTownOrDistrict_PerYES10080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CityOrTownOrDistrict_PerYES10080G(i)) Then
            msgbox_PerYES10080G ("* CityOrTownOrDistrict at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in schedule 80G Table C")
            ValidateCityOrTownOrDistrict_PerYES10080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateStateCode_PerYES10080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateStateCode_PerYES10080G = True
    'setTblinfo_PerYES10080G
    
    Set rangecells = Sheet20.Range("PerYES10080G.StateCode").Cells
    ReDim StateCode_PerYES10080G(end_PerYES10080G)
    
    For i = 1 To end_PerYES10080G
'        If Len(DoneeName_PerYES10080G(i)) = 0 Then
'            Exit For
'        End If
        
        StateCode_PerYES10080G(i) = rangecells.item(i).value
        StateCode_PerYES10080G(i) = Mid(StateCode_PerYES10080G(i), 1, 2)
        If isdropdownblank(StateCode_PerYES10080G(i)) Then
'            msgbox_PerYES10080G ("* Please StateCode select DropDown at Sr. No  " & i & "  in schedule 80G Table C")

        'SIT-69594 Changed by sai
            msgbox_PerYES10080G ("* ""Please select state of donee in schedule 80G Table C"" at Sr. No " & i & "")

            ValidateStateCode_PerYES10080G = False
            Exit Function
        End If
        
        
    'Changed By sai on 16/10/2024=======================================================
    If StateCode_PerYES10080G(i) <> "" Then
    Sheets("80G").Activate
    
        Dim ws_3 As Worksheet
        Set ws_3 = Worksheets("80G")
        
          Dim PIN_targetadd, state_targetadd As String
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")

         If Not state_Validation_Level(ws_3, PIN_targetadd, state_targetadd) Then
                 msgbox_hprptfrm ("* Invalid Pincode")
                 Sheet20.Range(Replace(state_targetadd, "H", "I")).value = ""
                 
               ValidateStateCode_PerYES10080G = False
             End If
        End If


        
     '============================================================================

        
        If (StateCode_PerYES10080G(i) = "99") Then
            msgbox_PerYES10080G ("* StateCode at Sr. No  " & i & " is Invalid in schedule 80G Table C")
            ValidateStateCode_PerYES10080G = False
            Exit Function
        End If
    Next
End Function

Function ValidatePinCode_PerYES10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePinCode_PerYES10080G = True
    'setTblinfo_PerYES10080G
    
    Set rangecells = Sheet20.Range("PerYES10080G.PinCode").Cells
    ReDim PinCode_PerYES10080G(end_PerYES10080G)
    
    For i = 1 To end_PerYES10080G
'        If Len(DoneeName_PerYES10080G(i)) = 0 Then
'            Exit For
'        End If
        
        PinCode_PerYES10080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(PinCode_PerYES10080G(i)) Then
'            msgbox_PerYES10080G ("* Please enter PinCode at Sr. No  " & i & "  in schedule 80G Table C")

            msgbox_PerYES10080G ("* ""Please enter PinCode in schedule 80G Table C"" at Sr. No  " & i & "")

            ValidatePinCode_PerYES10080G = False
            Exit Function
        End If
        
        If Not chkNumeric(PinCode_PerYES10080G(i)) Then
            msgbox_PerYES10080G ("* PinCode at Sr. No  " & i & "  only digits 0 to 9 allowed in schedule 80G Table C")
            ValidatePinCode_PerYES10080G = False
            Exit Function
        End If
    Next
End Function


Function ValidateDoneePAN_PerYES10080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim tempPan As String
    
    ValidateDoneePAN_PerYES10080G = True
    'setTblinfo_PerYES10080G
    
    Set rangecells = Sheet20.Range("PerYES10080G.DoneePAN").Cells
    ReDim DoneePAN_PerYES10080G(end_PerYES10080G)
    
    For i = 1 To end_PerYES10080G
'        If Len(DoneeName_PerYES10080G(i)) = 0 Then
'            Exit For
'        End If
'
        DoneePAN_PerYES10080G(i) = rangecells.item(i).value
        
        tempPan = DoneePAN_PerYES10080G(i)
        'ramya 70567
      If Not chkCompulsory(DoneePAN_PerYES10080G(i)) Then
           ' msgbox_PerYES10080G ("* Please enter DoneePAN at Sr. No  " & i & "  in schedule 80G Table C")
            ' SIT-70567 Changed by sai
            msgbox_PerYES10080G ("*""Please enter PAN of donee in schedule 80G Table C"" at Sr. No  " & i & "")
            ValidateDoneePAN_PerYES10080G = False
            Exit Function
        End If
        
        If Not ValidatePAN(tempPan) Then
'            msgbox_PerYES10080G ("* DoneePAN at Sr. No  " & i & "   is invalid. First 5 alphabets, next 4 digits, then alphabet in schedule 80G Table C")
            'SIT-70567
            msgbox_PerYES10080G ("* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule 80G Table C at Sr. No " & i & "")
            
            
            ValidateDoneePAN_PerYES10080G = False
            Exit Function
        End If
        
        If tempPan = tempPan11 Or tempPan = tempPan111 Then
            msgbox_PerYES10080G ("* DoneePAN at Sr. No  " & i & " Please enter the PAN to whom donation is made in schedule 80G Table C")
            ValidateDoneePAN_PerYES10080G = False
            Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt_PerYES10080G() As Boolean
    Dim i As Long
    Dim rangecells, rangecells1 As Range
    
    ValidateDonationAmt_PerYES10080G = True
    'setTblinfo_PerYES10080G

    Set rangecells = Sheet20.Range("PerYES10080G.DonationAmt").Cells
    Set rangecells1 = Sheet20.Range("PerYES10080G.DonationInCash").Cells
    ReDim DonationAmt_PerYES10080G(end_PerYES10080G)
    ReDim DonationAmt_PerYES10080G_InCash(end_PerYES10080G)
    For i = 1 To end_PerYES10080G
        DonationAmt_PerYES10080G(i) = rangecells.item(i).value
        DonationAmt_PerYES10080G_InCash(i) = rangecells1.item(i).value
        If (DonationAmt_PerYES10080G(i) = "" And DonationAmt_PerYES10080G_InCash(i) = "") Then
'            msgbox_PerYES10080G ("* Please enter Donation in Cash or Donation in other mode at Sr. No  " & i & "  in Sheet 80G is mandatory")
      'SIT-69614 changed by sai
            msgbox_PerYES10080G ("*Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode"" in schedule 80G Table C at Sr. No " & i & "")

            ValidateDonationAmt_PerYES10080G = False
            Exit Function
        End If
        
        If (DonationAmt_PerYES10080G(i) < 0) Then
            msgbox_PerYES10080G ("* 80G.C Amount at Sr. No  " & i & "   should be Non negative, no decimal, upto 99,999,999,999,999 in schedule 80G Table C")
            ValidateDonationAmt_PerYES10080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateEligibleAmt_PerYES10080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateEligibleAmt_PerYES10080G = True
    setTblinfo_PerYES10080G
    Set rangecells = Sheet20.Range("PerYES10080G.EligibleAmt").Cells
    ReDim EligibleAmt_PerYES10080G(end_PerYES10080G)
    
    For i = 1 To end_PerYES10080G
        EligibleAmt_PerYES10080G(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotDon50PercentNoApprReqd_PerYES10080G() As Boolean
    ValidateTotDon50PercentNoApprReqd_PerYES10080G = True
    TotDon50PercentNoApprReqd_PerYES10080G = Sheet20.Range("PerYES10080G.TotDon100Percent").value
End Function

Function ValidateTotElig50PercentNoApprReqd_PerYES10080G() As Boolean
    ValidateTotElig50PercentNoApprReqd_PerYES10080G = True
    TotElig50PercentNoApprReqd_PerYES10080G = Sheet20.Range("PerYES10080G.TotElig100Percent").value
End Function

Function ValidateDoneeWithPanName_Per5080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateDoneeWithPanName_Per5080G = True
    'setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.DoneeWithPanName").Cells
    ReDim DoneeName_Per5080G(end_Per5080G)
    noOfProcessSub = end_Per5080G
    For i = 1 To end_Per5080G
        DoneeName_Per5080G(i) = rangecells.item(i).value
        
'        If Len(DoneeName_Per5080G(i)) = 0 Then
'            Exit For
'        End If
        
        If Not chkCompulsory(DoneeName_Per5080G(i)) Then
'            msgbox_Per5080G ("* Please enter DoneeWithPanName at Sr. No  " & i & "  in schedule 80G Table D")
       'Jyoti D
           'SIT-69597  Changed by sai
            msgbox_Per5080G ("*""Please enter name of donee in schedule 80G Table D"" at Sr. No " & i & "")
            
            ValidateDoneeWithPanName_Per5080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(DoneeName_Per5080G(i)) Then
            msgbox_Per5080G ("* DoneeWithPanName at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in schedule 80G Table D")
            ValidateDoneeWithPanName_Per5080G = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
  
Function ValidateDoneePAN_Per5080G() As Boolean
    Dim rangecells As Range
    Dim tempPan As String
    Dim i As Long
    
    ValidateDoneePAN_Per5080G = True
    'setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.DoneePAN").Cells
    ReDim DoneePAN_Per5080G(end_Per5080G)
    
    For i = 1 To end_Per5080G
'        If Len(DoneeName_Per5080G(i)) = 0 Then
'            Exit For
'        End If
        DoneePAN_Per5080G(i) = rangecells.item(i).value
        
        tempPan = DoneePAN_Per5080G(i)
        
        If Not chkCompulsory(DoneePAN_Per5080G(i)) Then
'            msgbox_Per5080G ("* PLease enter DoneePAN at Sr. No  " & i & "  in schedule 80G Table D")
          
'SIT-69327 changed by sai
              msgbox_Per5080G ("* ""Please enter PAN of donee in schedule 80G Table D"" at Sr. No  " & i & "")


            ValidateDoneePAN_Per5080G = False
            Exit Function
        End If
        
        If Not ValidatePAN(tempPan) Then
'            msgbox_Per5080G ("* DoneePAN at Sr. No  " & i & " is invalid. First 5 alphabets, next 4 digits, then alphabet in schedule 80G Table D")
    'SIT-69327 changed by sai
            msgbox_Per5080G ("* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule 80G Table D"" at Sr. No  " & i & "")
            
            ValidateDoneePAN_Per5080G = False
            Exit Function
        End If
        
        If tempPan = tempPan11 Or tempPan = tempPan111 Then
            msgbox_Per5080G ("* DoneePAN at Sr. No  " & i & " Please enter the PAN to whom donation is made in schedule 80G Table D")
            ValidateDoneePAN_Per5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateAddrDetail_Per5080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateAddrDetail_Per5080G = True
    'setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.AddrDetail").Cells
    ReDim AddrDetail_Per5080G(end_Per5080G)
    
    For i = 1 To end_Per5080G
'        If Len(DoneeName_Per5080G(i)) = 0 Then
'            Exit For
'        End If
    
        AddrDetail_Per5080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(AddrDetail_Per5080G(i)) Then
           ' msgbox_Per5080G ("* Please enter Address Detail at Sr. No  " & i & "  in schedule 80G Table D")
           
           'Added by Aavula
           msgbox_Per5080G ("* Please enter address of donee in schedule 80G Table D at Sr. No  " & i & "")
            ValidateAddrDetail_Per5080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(AddrDetail_Per5080G(i)) Then
            msgbox_Per5080G ("* Address Detail at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in schedule 80G Table D")
            ValidateAddrDetail_Per5080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateCityOrTownOrDistrict_Per5080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateCityOrTownOrDistrict_Per5080G = True
    'setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.CityOrTownOrDistrict").Cells
    ReDim CityOrTownOrDistrict_Per5080G(end_Per5080G)
    
    For i = 1 To end_Per5080G
'        If Len(DoneeName_Per5080G(i)) = 0 Then
'            Exit For
'        End If
        
        CityOrTownOrDistrict_Per5080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(CityOrTownOrDistrict_Per5080G(i)) Then
            msgbox_Per5080G ("* Please enter CityOrTownOrDistrict at Sr. No  " & i & "  in schedule 80G Table D")
            ValidateCityOrTownOrDistrict_Per5080G = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CityOrTownOrDistrict_Per5080G(i)) Then
            msgbox_Per5080G ("* CityOrTownOrDistrict at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule 80G Table D")
            ValidateCityOrTownOrDistrict_Per5080G = False
            Exit Function
        End If
    Next
End Function
 
Function ValidateStateCode_Per5080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateStateCode_Per5080G = True
    'setTblinfo_Per5080G
    Set rangecells = Sheet20.Range("Per5080G.StateCode").Cells
    ReDim StateCode_Per5080G(end_Per5080G)
    
    For i = 1 To end_Per5080G
'        If Len(DoneeName_Per5080G(i)) = 0 Then
'            Exit For
'        End If
        StateCode_Per5080G(i) = rangecells.item(i).value
        StateCode_Per5080G(i) = Mid(StateCode_Per5080G(i), 1, 2)
        
        If isdropdownblank(StateCode_Per5080G(i)) Then
            msgbox_Per5080G ("* Please select StateCode DropDown at Sr. No  " & i & "  in schedule 80G Table D")
            ValidateStateCode_Per5080G = False
            Exit Function
        End If
        
        If (StateCode_Per5080G(i) = "99") Then
            msgbox_Per5080G ("* StateCode at Sr. No  " & i & " is Invalid in schedule 80G Table D")
            ValidateStateCode_Per5080G = False
            Exit Function
        End If
        
    'By sai on 17/09/2024
    If StateCode_Per5080G(i) <> "" Then
    Sheets("80G").Activate
    
        Dim ws_4 As Worksheet
        Set ws_4 = Worksheets("80G")

          Dim PIN_targetadd, state_targetadd As String
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")
      
        If Not state_Validation_Level(ws_4, PIN_targetadd, state_targetadd) Then
                     msgbox_hprptfrm ("* Invalid Pincode")
                     Sheet20.Range(Replace(state_targetadd, "H", "I")).value = ""
                     
                   ValidateStateCode_Per5080G = False
                 End If
     End If
'==========================================================================================================
    Next
End Function

Function ValidatePinCode_Per5080G() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidatePinCode_Per5080G = True
    'setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.PinCode").Cells
    ReDim PinCode_Per5080G(end_Per5080G)
    
    For i = 1 To end_Per5080G
'        If Len(DoneeName_Per5080G(i)) = 0 Then
'            Exit For
'        End If
        
        PinCode_Per5080G(i) = rangecells.item(i).value
        
        If Not chkCompulsory(PinCode_Per5080G(i)) Then
'            msgbox_Per5080G ("* Please enter PinCode at Sr. No  " & i & "  in schedule 80G Table D")

          'SIT-69611 Changed by sai
            msgbox_Per5080G ("*""Please enter pin code of donee in schedule 80G Table D"" at Sr. No " & i & "")
            
            ValidatePinCode_Per5080G = False
            Exit Function
        End If
        
        If Not chkNumeric(PinCode_Per5080G(i)) Then
'            msgbox_Per5080G ("* PinCode at Sr. No  " & i & "  only digits 0 to 9 allowed in schedule 80G Table D")
            
            msgbox_Per5080G ("*""Please enter a 6 digit valid Pin Code in schedule 80G Table D"" at Sr. No " & i & "")
            
            ValidatePinCode_Per5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonationAmt_Per5080G() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    ValidateDonationAmt_Per5080G = True
    'setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.DonationAmt").Cells
    Set rangecells1 = Sheet20.Range("Per5080G.DonationInCash").Cells
    ReDim DonationAmt_Per5080G(end_Per5080G)
    ReDim DonationAmt_Per5080G_InCash(end_Per5080G)
    For i = 1 To end_Per5080G
        DonationAmt_Per5080G(i) = rangecells.item(i).value
        DonationAmt_Per5080G_InCash(i) = rangecells1.item(i).value
        
        If (DonationAmt_Per5080G(i) = "" And DonationAmt_Per5080G_InCash(i) = "") Then
'            msgbox_Per5080G ("* Please enter Donation in Cash or Donation in Other mode at Sr. No  " & i & "  in schedule 80G Table D")
        'SIT-69327 changed by sai
            msgbox_Per5080G ("* Enter the amount of donation either in field ""Donation in cash"" Or ""Donation in other mode in schedule 80G Table D at Sr. No  " & i & "")
            
            
            ValidateDonationAmt_Per5080G = False
            Exit Function
        End If
        
        If (DonationAmt_Per5080G(i) < 0) Then
            msgbox_Per5080G ("* 80G.D Amount at Sr. No  " & i & "   should be Non negative, no decimal, upto 99,999,999,999,999 in schedule 80G Table D")
            ValidateDonationAmt_Per5080G = False
            Exit Function
        End If
    Next
End Function

Function ValidateEligibleAmt_Per5080G() As Boolean
    Dim i As Long
    Dim rangecells As Range
    
    ValidateEligibleAmt_Per5080G = True
    setTblinfo_Per5080G
    
    Set rangecells = Sheet20.Range("Per5080G.EligibleAmt").Cells
    ReDim EligibleAmt_Per5080G(end_Per5080G)
    For i = 1 To end_Per5080G
        EligibleAmt_Per5080G(i) = rangecells.item(i).value
    Next
End Function

Function ValidateTotDon50PercentApprReqd_Per5080G() As Boolean
    ValidateTotDon50PercentApprReqd_Per5080G = True
    TotDon50PercentApprReqd_Per5080G = Sheet20.Range("Per5080G.TotDon100Percent").value
End Function

Function ValidateTotElig50PercentApprReqd_Per5080G() As Boolean
    ValidateTotElig50PercentApprReqd_Per5080G = True
    TotElig50PercentApprReqd_Per5080G = Sheet20.Range("Per5080G.TotElig100Percent").value
End Function

Function ValidateTotalDonationsUs80G_Per5080G() As Boolean
    ValidateTotalDonationsUs80G_Per5080G = True
    TotalDonationsUs80G_Per5080G = Sheet20.Range("Per5080G.TotalDonationsUs80G").value
End Function

Function ValidateTotalEligibleDonationsUs80G_Per5080G() As Boolean
    ValidateTotalEligibleDonationsUs80G_Per5080G = True
    TotalEligibleDonationsUs80G_Per5080G = Sheet20.Range("Per5080G.TotalEligibleDonationsUs80G").value
End Function

Function msgbox_Per10080G(strmsg As String) As String
     msgValidateSheetPer10080G = msgValidateSheetPer10080G & strmsg & Chr(13)
End Function

Function msgbox_PerNO5080G(strmsg As String) As String
     msgValidateSheetPerNO5080G = msgValidateSheetPerNO5080G & strmsg & Chr(13)
End Function

Function msgbox_Per5080G(strmsg As String) As String
     msgValidateSheetPer5080G = msgValidateSheetPer5080G & strmsg & Chr(13)
End Function

Function msgbox_PerYES10080G(strmsg As String) As String
     msgValidateSheetPerYES10080G = msgValidateSheetPerYES10080G & strmsg & Chr(13)
End Function


Function ValidateChequeNumber_80G() As Boolean
    ValidateChequeNumber_80G = True
    
    Dim rangecells As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("Per10080G.DonationAmt").Cells
    ReDim Donation_other_80GA(end_Per10080G)
    
    Set rangecells = Sheet20.Range("Chequeno_80GA").Cells
    Dim i As Long
    ReDim ChequeNumber_80GA(end_Per10080G)
    
    
    For i = 1 To end_Per10080G
        ChequeNumber_80GA(i) = rangecells.item(i).value
        Donation_other_80GA(i) = rangecells2.item(i).value
        
        
        
        
    If rangecells2.item(i).Locked = False Then

           If Len(Donation_other_80GA(i)) > 0 And Donation_other_80GA(i) <> 0 Then
            
            If Not chkCompulsory(ChequeNumber_80GA(i)) Then
            msgbox_Per10080G ("* please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction")
                ValidateChequeNumber_80G = False
                Exit Function
            End If
            
            If Len(ChequeNumber_80GA(i)) > 0 Then
              If Not checkfieldspecialcharacter_Bank(ChequeNumber_80GA(i)) Then
                  msgbox_Per10080G ("* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.No " & i & " in Table A""") & Chr(13)
                  ValidateChequeNumber_80G = False
                  Exit Function
              End If
            End If
            
            If Len(ChequeNumber_80GA(i)) > 50 Then
                msgbox_Per10080G ("* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction at Sr.No " & i & " in Table A cannot exceed 50 characters")
                ValidateChequeNumber_80G = False
                Exit Function
            End If
     End If
     End If
     
        UpdateProgressBar
    Next
End Function

Function ValidateIFSC_80G() As Boolean
    ValidateIFSC_80G = True
    
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("Per10080G.DonationAmt").Cells
    ReDim Donation_other_80GA(end_Per10080G)
    
    Dim rangecells3 As Range
    Set rangecells3 = Range("IFSC_80GA").Cells
    ReDim BankIFSC_80GA(end_Per10080G)
        Dim i As Long

    
    For i = 1 To end_Per10080G
        BankIFSC_80GA(i) = rangecells3.item(i).value
        Donation_other_80GA(i) = rangecells2.item(i).value

    If rangecells2.item(i).Locked = False Then
        If Len(Donation_other_80GA(i)) > 0 And Donation_other_80GA(i) <> 0 Then
         If isdropdownblank(BankIFSC_80GA(i)) Then
           msgbox_Per10080G ("* please enter ""your bank IFSC from which contribution is made""")
           ValidateIFSC_80G = False
           Exit Function
         End If
        End If
         
          
        If Len(BankIFSC_80GA(i)) > 11 Then
           msgbox_Per10080G ("* IFS Code at Sr.No " & i & " in Table A cannot exceed 11 characters")
           ValidateIFSC_80G = False
           Exit Function
        End If
            
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_80GA(i))) Then
           msgbox_Per10080G ("* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)")
           ValidateIFSC_80G = False
           Exit Function
        End If
    End If
        Next
 End Function


Function ValidateChequeNumber_80GC() As Boolean
    ValidateChequeNumber_80GC = True
    
    Dim rangecells As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("PerYES10080G.DonationAmt").Cells
    ReDim Donation_other_80GC(end_PerYES10080G)
    
    Set rangecells = Sheet20.Range("Chequeno_80GC").Cells
    Dim i As Long
    ReDim ChequeNumber_80GC(end_PerYES10080G)
    
    
    For i = 1 To end_PerYES10080G
        ChequeNumber_80GC(i) = rangecells.item(i).value
        Donation_other_80GC(i) = rangecells2.item(i).value
        
        
        
    If rangecells2.item(i).Locked = False Then

           If Len(Donation_other_80GC(i)) > 0 And Donation_other_80GC(i) <> 0 Then
            
            If Not chkCompulsory(ChequeNumber_80GC(i)) Then
            msgbox_PerYES10080G ("* ""please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction")
   
                ValidateChequeNumber_80GC = False
                Exit Function
            End If
            
             If Len(ChequeNumber_80GC(i)) > 0 Then
              If Not checkfieldspecialcharacter_Bank(ChequeNumber_80GC(i)) Then
                  msgbox_PerYES10080G ("* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.No " & i & " in Table C""") & Chr(13)
                  ValidateChequeNumber_80GC = False
                  Exit Function
              End If
            End If
            
            If Len(ChequeNumber_80GC(i)) > 50 Then
                msgbox_PerYES10080G ("* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction at Sr.No " & i & " in Table C cannot exceed 50 characters")
  
                ValidateChequeNumber_80GC = False
                Exit Function
            End If
     End If
     End If
     
        UpdateProgressBar
    Next
End Function

Function ValidateIFSC_80GC() As Boolean
    ValidateIFSC_80GC = True
    
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("PerYES10080G.DonationAmt").Cells
    ReDim Donation_other_80GC(end_PerYES10080G)
    
    Dim rangecells3 As Range
    Set rangecells3 = Range("IFSC_80GC").Cells
    ReDim BankIFSC_80GC(end_PerYES10080G)
        Dim i As Long

    
    For i = 1 To end_PerYES10080G
        BankIFSC_80GC(i) = rangecells3.item(i).value
        Donation_other_80GC(i) = rangecells2.item(i).value

    If rangecells2.item(i).Locked = False Then
        If Len(Donation_other_80GC(i)) > 0 And Donation_other_80GC(i) <> 0 Then
         If isdropdownblank(BankIFSC_80GC(i)) Then
           msgbox_PerYES10080G ("* please enter ""your bank IFSC from which contribution is made""")

           ValidateIFSC_80GC = False
           Exit Function
         End If
        End If
         
          
        If Len(BankIFSC_80GC(i)) > 11 Then
           msgbox_PerYES10080G ("* IFS Code at Sr.No " & i & " in Table A cannot exceed 11 characters")
           ValidateIFSC_80GC = False
           Exit Function
        End If
            
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_80GC(i))) Then
          msgbox_PerYES10080G ("* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric")
           ValidateIFSC_80GC = False
           Exit Function
        End If
    End If
        Next
 End Function


Function ValidateChequeNumber_80GB() As Boolean
    ValidateChequeNumber_80GB = True
    
    Dim rangecells As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("PerNO5080G.DonationAmt").Cells
    ReDim Donation_other_80GB(end_PerNO5080G)
    
    Set rangecells = Sheet20.Range("Chequeno_80GB").Cells
    Dim i As Long
    ReDim ChequeNumber_80GB(end_PerNO5080G)
    
    
    For i = 1 To end_PerNO5080G
        ChequeNumber_80GB(i) = rangecells.item(i).value
        Donation_other_80GB(i) = rangecells2.item(i).value
        
        
        
    If rangecells2.item(i).Locked = False Then

           If Len(Donation_other_80GB(i)) > 0 And Donation_other_80GB(i) <> 0 Then
            
            If Not chkCompulsory(ChequeNumber_80GB(i)) Then
             msgbox_PerNO5080G ("* ""please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction""")
                ValidateChequeNumber_80GB = False
                Exit Function
            End If
            
             If Len(ChequeNumber_80GB(i)) > 0 Then
              If Not checkfieldspecialcharacter_Bank(ChequeNumber_80GB(i)) Then
                  msgbox_PerNO5080G ("* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.No " & i & " in Table B""") & Chr(13)
                  ValidateChequeNumber_80GB = False
                  Exit Function
              End If
            End If
            
            If Len(ChequeNumber_80GB(i)) > 50 Then
                   msgbox_PerNO5080G ("* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction at Sr.No " & i & " in Table B cannot exceed 50 characters""")
                ValidateChequeNumber_80GB = False
                Exit Function
            End If
     End If
     End If
     
        UpdateProgressBar
    Next
End Function

Function ValidateIFSC_80GB() As Boolean
    ValidateIFSC_80GB = True
    
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("PerNO5080G.DonationAmt").Cells
    ReDim Donation_other_80GB(end_PerNO5080G)
    
    Dim rangecells3 As Range
    Set rangecells3 = Range("IFSC_80GB").Cells
    ReDim BankIFSC_80GB(end_PerNO5080G)
        Dim i As Long

    
    For i = 1 To end_PerNO5080G
        BankIFSC_80GB(i) = rangecells3.item(i).value
        Donation_other_80GB(i) = rangecells2.item(i).value

    If rangecells2.item(i).Locked = False Then
        If Len(Donation_other_80GB(i)) > 0 And Donation_other_80GB(i) <> 0 Then
         If isdropdownblank(BankIFSC_80GB(i)) Then
             msgbox_PerNO5080G ("* ""please enter ""your bank IFSC from which contribution is made""")
           ValidateIFSC_80GB = False
           Exit Function
         End If
        End If
         
          
        If Len(BankIFSC_80GB(i)) > 11 Then
           msgbox_PerNO5080G ("* ""IFS Code at Sr.No " & i & " in Table A cannot exceed 11 characters""")

           ValidateIFSC_80GB = False
           Exit Function
        End If
            
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_80GB(i))) Then
          msgbox_PerNO5080G ("* ""Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric""")
   
           ValidateIFSC_80GB = False
           Exit Function
        End If
    End If
        Next
 End Function
Function ValidateChequeNumber_80GD() As Boolean
    ValidateChequeNumber_80GD = True
    
    Dim rangecells As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("Per5080G.DonationAmt").Cells
    ReDim Donation_other_80GD(end_Per5080G)
    
    Set rangecells = Sheet20.Range("Chequeno_80GD").Cells
    Dim i As Long
    ReDim ChequeNumber_80GD(end_Per5080G)
    
    
    For i = 1 To end_Per5080G
        ChequeNumber_80GD(i) = rangecells.item(i).value
        Donation_other_80GD(i) = rangecells2.item(i).value
        
        
        
    If rangecells2.item(i).Locked = False Then

           If Len(Donation_other_80GD(i)) > 0 And Donation_other_80GD(i) <> 0 Then
            
            If Not chkCompulsory(ChequeNumber_80GD(i)) Then
             msgbox_Per5080G ("* ""please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction""")
                ValidateChequeNumber_80GD = False
                Exit Function
            End If
            
            If Len(ChequeNumber_80GD(i)) > 0 Then
              If Not checkfieldspecialcharacter_Bank(ChequeNumber_80GD(i)) Then
                  msgbox_Per5080G ("* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.No " & i & " in Table D""") & Chr(13)
                  ValidateChequeNumber_80GD = False
                  Exit Function
              End If
            End If
            
            If Len(ChequeNumber_80GD(i)) > 50 Then
                msgbox_Per5080G ("* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction at Sr.No " & i & " in Table D cannot exceed 50 characters""")
                ValidateChequeNumber_80GD = False
                Exit Function
            End If
     End If
     End If
     
        UpdateProgressBar
    Next
End Function

Function ValidateIFSC_80GD() As Boolean
    ValidateIFSC_80GD = True
    
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet20.Range("Per5080G.DonationAmt").Cells
    ReDim Donation_other_80GD(end_Per5080G)
    
    Dim rangecells3 As Range
    Set rangecells3 = Range("IFSC_80GD").Cells
    ReDim BankIFSC_80GD(end_Per5080G)
        Dim i As Long

    
    For i = 1 To end_Per5080G
        BankIFSC_80GD(i) = rangecells3.item(i).value
        Donation_other_80GD(i) = rangecells2.item(i).value

    If rangecells2.item(i).Locked = False Then
        If Len(Donation_other_80GD(i)) > 0 And Donation_other_80GD(i) <> 0 Then
         If isdropdownblank(BankIFSC_80GD(i)) Then
             msgbox_Per5080G ("* ""please enter ""your bank IFSC from which contribution is made""")
           ValidateIFSC_80GD = False
           Exit Function
         End If
        End If
         
          
        If Len(BankIFSC_80GD(i)) > 11 Then
           msgbox_Per5080G ("* ""IFS Code at Sr.No " & i & " in Table A cannot exceed 11 characters""")
           ValidateIFSC_80GD = False
           Exit Function
        End If
            
        If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_80GD(i))) Then
          msgbox_Per5080G ("* ""Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric""")
   
           ValidateIFSC_80GD = False
           Exit Function
        End If
    End If
        Next
 End Function
