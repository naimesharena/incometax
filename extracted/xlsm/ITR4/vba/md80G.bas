Attribute VB_Name = "md80G"
Option Explicit

Public GA80DoneeName As Variant
Public GA80DoneeAdrDet As Variant
Public GA80DoneeCityTownn As Variant
Public GA80DoneeState As Variant
Public GA80DoneePinCode As Variant
Public GA80DoneePAN As Variant
Public GA80DoneeAmnt As Variant
Public GA80DoneeInCash As Variant
Public GA80DoneeTot_InCash As Variant
Public GA80DoneeTotalDonation As Variant
Public GA80DoneeTot_TotalDonation As Variant
Public GA80DoneeEligibleAmnt As Variant
Public GA80DoneeTotAmnt As Variant
Public GA80DoneeTotEligAmnt As Variant
Public GA80DoneeTraref As Variant
Public GA80DoneeIFSC As Variant

Public GB80DoneeName As Variant
Public GB80DoneeAdrDet As Variant
Public GB80DoneeCityTownn As Variant
Public GB80DoneeState As Variant
Public GB80DoneePinCode As Variant
Public GB80DoneePAN As Variant
Public GB80DoneeAmnt As Variant
Public GB80DoneeInCash As Variant
Public GB80DoneeTot_InCash As Variant
Public GB80DoneeTotalDonation As Variant
Public GB80DoneeTot_TotalDonation As Variant
Public GB80DoneeEligibleAmnt As Variant
Public GB80DoneeTotAmnt As Variant
Public GB80DoneeTotEligAmnt As Variant
Public GB80DoneeTraref As Variant
Public GB80DoneeIFSC As Variant

Public GC80DoneeName As Variant
Public GC80DoneeAdrDet As Variant
Public GC80DoneeCityTownn As Variant
Public GC80DoneeState As Variant
Public GC80DoneePinCode As Variant
Public GC80DoneePAN As Variant
Public GC80DoneeAmnt As Variant
Public GC80DoneeInCash As Variant
Public GC80DoneeTot_InCash As Variant
Public GC80DoneeTotalDonation As Variant
Public GC80DoneeTot_TotalDonation As Variant
Public GC80DoneeEligibleAmnt As Variant
Public GC80DoneeTotAmnt As Variant
Public GC80DoneeTotEligAmnt As Variant
Public GC80DoneeTraref As Variant
Public GC80DoneeIFSC As Variant


Public GD80DoneeName As Variant
Public GD80DoneeAdrDet As Variant
Public GD80DoneeCityTownn As Variant
Public GD80DoneeState As Variant
Public GD80DoneePinCode As Variant
Public GD80DoneePAN As Variant
Public GD80DoneeARN As Variant
Public GD80DoneeAmnt As Variant
Public GD80DoneeInCash As Variant
Public GD80DoneeTot_InCash As Variant
Public GD80DoneeTotalDonation As Variant
Public GD80DoneeTot_TotalDonation As Variant
Public GD80DoneeEligibleAmnt As Variant
Public GD80DoneeTotAmnt As Variant
Public GD80DoneeTotEligAmnt As Variant
Public GD80DoneeTraref As Variant
Public GD80DoneeIFSC As Variant

Public G80DoneeTotAmnt As Variant
Public G80DoneeTotEligAmnt As Variant
Public G80DoneeTotCash As Variant
Public G80DoneeTotTotalDonation As Variant

'Local variable
Dim tablename, tablecode As Variant
Dim errmsg80G As Variant
Dim end_norows, end_noofvalues As Variant
Dim tempDoneeName As Variant
Dim tempDoneeAddr As Variant
Dim tempDoneeCity As Variant
Dim tempDoneeState As Variant
Dim tempDoneePinCode As Variant
Dim tempDoneePAN As Variant
Dim tempDoneeARN As Variant
Dim tempDoneeTraref As Variant
Dim tempDoneeIFSC As Variant
Dim tempDoneeAmnt, tempDoneeCash As Variant
Public msg14Dig80G As String

Sub cmdValidate_Click_80G()
ValidateSheet80G
fmsgboxStatus "Sheet 80G is Ok" ', vbOKOnly, vbMessageCaption
End Sub
Sub cmdPrev_Click_80G()
Sheet16.Activate
End Sub
'
Sub cmdNext_Click_80G()
Sheet16.Activate
End Sub
 
Sub AddRows_Per10080GA()
Sheets("80G").Activate
mdCommon.DefinedgridNameRange = "Per10080G.DoneeName||Per10080G.AddrDetail||Per10080G.CityOrTownOrDistrict||Per10080G.StateCode||Per10080G.PinCode||Per10080G.DoneePAN||Per10080G.Donationcash||Per10080G.DonationAmt||Per10080G.Traref||Per10080G.IFSC||Per10080G.TotDonation||Per10080G.EligibleAmt||Data_80G_A||comb_80G_A||Comb_donation_80G_A||Comb_PAN_80G_A"
ActiveCellRange = mdCommon.searchLastRow("Per10080G.DoneeName")
mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows_Per10080GB()
Sheets("80G").Activate
mdCommon.DefinedgridNameRange = "PerNO5080G.DoneeName||PerNO5080G.AddrDetail||PerNO5080G.CityOrTownOrDistrict||PerNO5080G.StateCode||PerNO5080G.PinCode||PerNO5080G.DoneePAN||PerNO5080G.DonationAmt||PerNO5080G.DonationCash||PerNO5080G.Traref||PerNO5080G.IFSC||PerNO5080G.TotDonation||PerNO5080G.EligibleAmt||Data_80G_B||comb_80G_B||Comb_donation_80G_B||Comb_PAN_80G_B"
ActiveCellRange = mdCommon.searchLastRow("PerNO5080G.DoneeName")
mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows_Per10080GC()
Sheets("80G").Activate
mdCommon.DefinedgridNameRange = "PerYES10080G.DoneeName||PerYES10080G.AddrDetail||PerYES10080G.CityOrTownOrDistrict||PerYES10080G.StateCode||PerYES10080G.PinCode||PerYES10080G.DoneePAN||PerYES10080G.DonationAmt||PerYES10080G.DonationCash||PerYES10080G.Traref||PerYES10080G.IFSC||PerYES10080G.TotDonation||PerYES10080G.EligibleAmt||Data_80G_C||comb_80G_C||Comb_donation_80G_C||Comb_PAN_80G_C"
ActiveCellRange = mdCommon.searchLastRow("PerYES10080G.DoneeName")
mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows_Per10080GD()
Sheets("80G").Activate
mdCommon.DefinedgridNameRange = "Per5080G.DoneeName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DoneeARN||Per5080G.DonationAmt||Per5080G.DonationCash||Per5080G.TotDonation||Per5080G.Traref||Per5080G.IFSC||Per5080G.EligibleAmt||Per5080G.DonationCashTemp||Data_80G_D||comb_80G_D||Comb_donation_80G_D||Comb_PAN_80G_D"
ActiveCellRange = mdCommon.searchLastRow("Per5080G.DoneeName")
mdCommon.insertRowUnderSectionWithFormula
End Sub


Sub ValidateSheet80G()
msgError = ""

If Not validate80G Then
'MessageBox.LMessagebox.Height = 600
'MessageBox.Height = 700
'MessageBox.CommandButton1.Top = 640

Sheet6.Activate
fmsgbox msgError ', vbOKOnly, "Error(s)!"
CloseMsg
End If


If Not validate14DIgits80G Then

Sheet6.Activate
fmsgboxsmall msg14Dig80G ', vbOKOnly, "Error(s)!"
CloseMsg
End If

End Sub

Function validate14DIgits80G() As Boolean
validate14DIgits80G = True
'-------------A---------------
If Sheet6.Range("Per10080G.TotDon100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in other mode cannot exceed 14 digits in 80G A" & Chr(13)
End If

If Sheet6.Range("Per10080G.TotElig100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Eligible Amount of Donation cannot exceed 14 digits in 80G A" & Chr(13)
End If

If Sheet6.Range("Per10080G.TotDonCash100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G A" & Chr(13)
End If

If Sheet6.Range("Per10080G.TotTotalDonation100Percent").value > 99999999999999# Then

'If Sheet6.Range("Per10080G.TotTotalDonation100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G A" & Chr(13)
End If


'-------------B---------------

If Sheet6.Range("PerNO5080G.TotDon100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.80GB.1
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation cannot exceed 14 digits in 80G B" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in Other Mode cannot exceed 14 digits in 80G B" & Chr(13)
    'End change
End If


If Sheet6.Range("PerNO5080G.TotElig100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Eligible Amount of Donation cannot exceed 14 digits in 80G B" & Chr(13)
End If

If Sheet6.Range("PerNO5080G.TotDoncash100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.56
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G B" & Chr(13)
    'End Change IDS.56
End If

If Sheet6.Range("PerNO5080G.TotTotalDonation100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.80GB.2
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G B" & Chr(13)
    'End change
End If


'-------------C---------------

If Sheet6.Range("PerYES10080G.TotDon100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.63
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation cannot exceed 14 digits in 80G C" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in Other mode cannot exceed 14 digits in 80G C" & Chr(13)
    'End Change IDS.63
End If

If Sheet6.Range("PerYES10080G.TotElig100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Eligible Amount of Donation cannot exceed 14 digits in 80G C" & Chr(13)
End If

If Sheet6.Range("PerYES10080G.TotDoncash100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.62
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G C" & Chr(13)
    'End ChangeIDS.62
    
End If

If Sheet6.Range("PerYES10080G.TotTotalDonation100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.80GC.1
'    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G C" & Chr(13)
    'End Change
End If
'-------------D---------------

If Sheet6.Range("Per5080G.TotDon100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.67
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation cannot exceed 14 digits in 80G D" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in Other Mode cannot exceed 14 digits in 80G D" & Chr(13)
    'End Change IDS.67
End If


If Sheet6.Range("Per5080G.TotElig100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    msg14Dig80G = msg14Dig80G + "* Total Eligible Amount of Donation cannot exceed 14 digits in 80G D" & Chr(13)
End If

If Sheet6.Range("Per5080G.TotDoncash100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.69
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G D" & Chr(13)
    'End Change IDS.69
End If

If Sheet6.Range("Per5080G.TotTotalDonation100Percent").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.68
'    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G D" & Chr(13)
    'End Change IDS.68
End If
'-------------Total---------------
If Sheet6.Range("Per5080G.TotalDonationsUs80G").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.80GD.1
    'msg14Dig80G = msg14Dig80G + "* Total Donation (A+B+C+D) cannot exceed 14 digits" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Donation in Other Mode cannot exceed 14 digits in 80G E" & Chr(13)
    'End Change
End If


If Sheet6.Range("Per5080G.TotalEligibleDonationsUs80G").value > 99999999999999# Then
    validate14DIgits80G = False
    'Change.01.03.2023.102.IDS.70
    'msg14Dig80G = msg14Dig80G + "* Total Eligible Amt of Donation (A+B+C+D) cannot exceed 14 digits " & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Eligible Amt of Donation (A+B+C+D) cannot exceed 14 digits in 80G E " & Chr(13)
    'End Change IDS.70
End If

If Sheet6.Range("Per5080G.TotalCashDonationsUs80G").value > 99999999999999# Then
    validate14DIgits80G = False
'    Change.01.03.2023.102.Total80G
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Donation in cash cannot exceed 14 digits in 80G E" & Chr(13)
End If

If Sheet6.Range("Per5080G.TotalOfTotalDonationsUs80G").value > 99999999999999# Then
    validate14DIgits80G = False
    'msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G A" & Chr(13)
    msg14Dig80G = msg14Dig80G + "* Total Amount of Total Donation cannot exceed 14 digits in 80G E" & Chr(13)
'    End Change
End If
End Function


Sub settableinfo80G(Optional tabname As Variant = "")

Dim intCellCount, intcntr As Long

Dim rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8, rangecells9, rangecells10, rangecells11 As Range
Dim ccount As Variant
Dim TbCod As String
Dim x As Long

end_noofvalues = 0
end_norows = 0
If tabname <> "" Then
    Set rangecells1 = Sheet6.Range(tabname & ".DoneeName").Cells
    Set rangecells2 = Sheet6.Range(tabname & ".AddrDetail").Cells
    Set rangecells3 = Sheet6.Range(tabname & ".CityOrTownOrDistrict").Cells
    Set rangecells4 = Sheet6.Range(tabname & ".StateCode").Cells
    Set rangecells5 = Sheet6.Range(tabname & ".PinCode").Cells
    Set rangecells6 = Sheet6.Range(tabname & ".DoneePAN").Cells
    Set rangecells7 = Sheet6.Range(tabname & ".DonationAmt").Cells
    Set rangecells8 = Sheet6.Range(tabname & ".Donationcash").Cells
    Set rangecells9 = Sheet6.Range(tabname & ".TotDonation").Cells
    Set rangecells10 = Sheet6.Range(tabname & ".Traref").Cells
    Set rangecells11 = Sheet6.Range(tabname & ".IFSC").Cells

    intCellCount = Sheet6.Range(tabname & ".DoneeName").count
Else
    Set rangecells1 = Sheet6.Range(tabname & ".DoneeName").Cells
    Set rangecells2 = Sheet6.Range(tabname & ".AddrDetail").Cells
    Set rangecells3 = Sheet6.Range(tabname & ".CityOrTownOrDistrict").Cells
    Set rangecells4 = Sheet6.Range(tabname & ".StateCode").Cells
    Set rangecells5 = Sheet6.Range(tabname & ".PinCode").Cells
    Set rangecells6 = Sheet6.Range(tabname & ".DoneePAN").Cells
    Set rangecells7 = Sheet6.Range(tabname & ".Donationcash").Cells
    Set rangecells9 = Sheet6.Range(tabname & ".TotDonation").Cells
    Set rangecells10 = Sheet6.Range(tabname & ".Traref").Cells
    Set rangecells11 = Sheet6.Range(tabname & ".IFSC").Cells
    
    intCellCount = Sheet6.Range(tabname & ".DoneeName").count
End If
ccount = 0

For intcntr = 1 To intCellCount
Sheet6.Activate


'Konda updated on 20-03-2026--SIT-115316
'If Not ((rangecells1.item(intcntr).value = "") Or (rangecells2.item(intcntr).value = "") Or (rangecells3.item(intcntr).value = "") Or (rangecells4.item(intcntr).value = "" Or rangecells4.item(intcntr).value = "(Select)") Or (rangecells5.item(intcntr).value = "") Or (rangecells6.item(intcntr).value = "") Or (rangecells9.item(intcntr).value = 0) Or (rangecells10.item(intcntr).value = "") Or (rangecells11.item(intcntr).value = "")) Then ccount = ccount + 1
If Not ((rangecells1.item(intcntr).value = "") Or (rangecells2.item(intcntr).value = "") Or (rangecells3.item(intcntr).value = "") Or (rangecells4.item(intcntr).value = "" Or rangecells4.item(intcntr).value = "(Select)") Or (rangecells5.item(intcntr).value = "") Or (rangecells6.item(intcntr).value = "") Or (rangecells9.item(intcntr).value = 0)) Then ccount = ccount + 1
'======================

'Malli---------------
If rangecells4.item(intcntr).value <> "" And Not rangecells4.item(intcntr).value = "(Select)" Then
Dim PIN_targetadd, state_targetadd As String
      
      state_targetadd = Replace(rangecells4.item(intcntr).Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "H", "I")
      
      
    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet6.Range(Replace(state_targetadd, "H", "I")).value = ""
  End If

'--------------------

Next

For intcntr = 1 To intCellCount
If tabname = "Per10080G" Then TbCod = "80G:A"
If tabname = "PerNO5080G" Then TbCod = "80G:B"
If tabname = "PerYES10080G" Then TbCod = "80G:C"
If tabname = "Per5080G" Then TbCod = "80G:D"
If Mid(rangecells4.item(intcntr).value, 1, 2) = "99" Then
    fmsgbox TbCod & Chr(10) & "* Select valid option At Sr No. " & intcntr & " in State code in " & TbCod   ', "Error(s)!"
    CloseMsg
End If
Next

end_noofvalues = ccount
    
end_norows = intCellCount

End Sub

Function validate80G() As Boolean
On Error Resume Next

validate80G = True

subProcCaption = "Validating 80GA"
Dim cntr1, cntr2 As Long
errmsg80G = ""

tablename = "Per10080G"
tablecode = ": Table A"

settableinfo80G (tablename)

noOfProcessSub = end_norows

ReDim GA80DoneeName(end_noofvalues)
ReDim GA80DoneeAdrDet(end_noofvalues)
ReDim GA80DoneeCityTownn(end_noofvalues)
ReDim GA80DoneeState(end_noofvalues)
ReDim GA80DoneePinCode(end_noofvalues)
ReDim GA80DoneePAN(end_noofvalues)
ReDim GA80DoneeAmnt(end_noofvalues)
ReDim GA80DoneeInCash(end_noofvalues)
ReDim GA80DoneeTotalDonation(end_noofvalues)
ReDim GA80DoneeEligibleAmnt(end_noofvalues)
ReDim GA80DoneeTraref(end_noofvalues)
ReDim GA80DoneeIFSC(end_noofvalues)
cntr2 = 0

For cntr1 = 1 To end_norows
    If Not ifanyEmpty80G(cntr1) Then
    
        If errmsg80G = "" Then
        GA80DoneeName(cntr2) = tempDoneeName
        GA80DoneeAdrDet(cntr2) = tempDoneeAddr
        GA80DoneeCityTownn(cntr2) = tempDoneeCity
        GA80DoneeState(cntr2) = tempDoneeState
        GA80DoneePinCode(cntr2) = tempDoneePinCode
        GA80DoneePAN(cntr2) = tempDoneePAN
        GA80DoneeAmnt(cntr2) = tempDoneeAmnt
        GA80DoneeInCash(cntr2) = tempDoneeCash
        GA80DoneeTraref(cntr2) = tempDoneeTraref
        GA80DoneeIFSC(cntr2) = tempDoneeIFSC
        GA80DoneeTotalDonation(cntr2) = Sheet6.Range(tablename & ".TotDonation").item(cntr1).value
        GA80DoneeEligibleAmnt(cntr2) = Sheet6.Range(tablename & ".EligibleAmt").item(cntr1).value
        cntr2 = cntr2 + 1
        End If
    Else
        msgError = msgError & errmsg80G & Chr(13)
    End If
    UpdateProgressBar
Next
GA80DoneeTot_InCash = Sheet6.Range("Per10080G.TotDonCash100Percent").value
GA80DoneeTotAmnt = Sheet6.Range("Per10080G.TotDon100Percent").value
GA80DoneeTot_TotalDonation = Sheet6.Range("Per10080G.TotTotalDonation100Percent").value
GA80DoneeTotEligAmnt = Sheet6.Range("Per10080G.TotElig100Percent").value

tablename = "PerNO5080G"
tablecode = ": Table B"
subProcCaption = "Validating 80GB"
settableinfo80G (tablename)
noOfProcessSub = end_norows
ReDim GB80DoneeName(end_noofvalues)
ReDim GB80DoneeAdrDet(end_noofvalues)
ReDim GB80DoneeCityTownn(end_noofvalues)
ReDim GB80DoneeState(end_noofvalues)
ReDim GB80DoneePinCode(end_noofvalues)
ReDim GB80DoneePAN(end_noofvalues)
ReDim GB80DoneeAmnt(end_noofvalues)
ReDim GB80DoneeInCash(end_noofvalues)
ReDim GB80DoneeTotalDonation(end_noofvalues)
ReDim GB80DoneeEligibleAmnt(end_noofvalues)
ReDim GB80DoneeTraref(end_noofvalues)
ReDim GB80DoneeIFSC(end_noofvalues)
cntr2 = 0

For cntr1 = 1 To end_norows
    If Not ifanyEmpty80G(cntr1) Then
    If errmsg80G = "" Then

        GB80DoneeName(cntr2) = tempDoneeName
        GB80DoneeAdrDet(cntr2) = tempDoneeAddr
        GB80DoneeCityTownn(cntr2) = tempDoneeCity
        GB80DoneeState(cntr2) = tempDoneeState
        GB80DoneePinCode(cntr2) = tempDoneePinCode
        GB80DoneePAN(cntr2) = tempDoneePAN
        GB80DoneeAmnt(cntr2) = tempDoneeAmnt
        GB80DoneeInCash(cntr2) = tempDoneeCash
        GB80DoneeTraref(cntr2) = tempDoneeTraref
        GB80DoneeIFSC(cntr2) = tempDoneeIFSC
        GB80DoneeTotalDonation(cntr2) = Sheet6.Range(tablename & ".TotDonation").item(cntr1).value
        GB80DoneeEligibleAmnt(cntr2) = Sheet6.Range(tablename & ".EligibleAmt").item(cntr1).value
        cntr2 = cntr2 + 1
    End If
    Else
        msgError = msgError & errmsg80G & Chr(13)
    End If
    UpdateProgressBar
Next
GB80DoneeTot_InCash = Sheet6.Range("PerNO5080G.TotDoncash100Percent").value
GB80DoneeTotAmnt = Sheet6.Range("PerNO5080G.TotDon100Percent").value
GB80DoneeTot_TotalDonation = Sheet6.Range("PerNO5080G.TotTotalDonation100Percent").value
GB80DoneeTotEligAmnt = Sheet6.Range("PerNO5080G.TotElig100Percent").value


tablename = "PerYES10080G"
tablecode = ": Table C"
subProcCaption = "Validating 80GC"
settableinfo80G (tablename)
noOfProcessSub = end_norows
ReDim GC80DoneeName(end_noofvalues)
ReDim GC80DoneeAdrDet(end_noofvalues)
ReDim GC80DoneeCityTownn(end_noofvalues)
ReDim GC80DoneeState(end_noofvalues)
ReDim GC80DoneePinCode(end_noofvalues)
ReDim GC80DoneePAN(end_noofvalues)
ReDim GC80DoneeAmnt(end_noofvalues)
ReDim GC80DoneeInCash(end_noofvalues)
ReDim GC80DoneeTotalDonation(end_noofvalues)
ReDim GC80DoneeEligibleAmnt(end_noofvalues)
ReDim GC80DoneeTraref(end_noofvalues)
ReDim GC80DoneeIFSC(end_noofvalues)
cntr2 = 0

For cntr1 = 1 To end_norows
    If Not ifanyEmpty80G(cntr1) Then
        If errmsg80G = "" Then
        GC80DoneeName(cntr2) = tempDoneeName
        GC80DoneeAdrDet(cntr2) = tempDoneeAddr
        GC80DoneeCityTownn(cntr2) = tempDoneeCity
        GC80DoneeState(cntr2) = tempDoneeState
        GC80DoneePinCode(cntr2) = tempDoneePinCode
        GC80DoneePAN(cntr2) = tempDoneePAN
        GC80DoneeAmnt(cntr2) = tempDoneeAmnt
        GC80DoneeInCash(cntr2) = tempDoneeCash
        GC80DoneeTraref(cntr2) = tempDoneeTraref
        GC80DoneeIFSC(cntr2) = tempDoneeIFSC
        GC80DoneeTotalDonation(cntr2) = Sheet6.Range(tablename & ".TotDonation").item(cntr1).value
        GC80DoneeEligibleAmnt(cntr2) = Sheet6.Range(tablename & ".EligibleAmt").item(cntr1).value
        cntr2 = cntr2 + 1
        End If
    Else
        msgError = msgError & errmsg80G & Chr(13)
    End If
    UpdateProgressBar
Next
GC80DoneeTot_InCash = Sheet6.Range("PerYES10080G.TotDoncash100Percent").value
GC80DoneeTotAmnt = Sheet6.Range("PerYES10080G.TotDon100Percent").value
GC80DoneeTot_TotalDonation = Sheet6.Range("PerYES10080G.TotTotalDonation100Percent").value
GC80DoneeTotEligAmnt = Sheet6.Range("PerYES10080G.TotElig100Percent").value


tablename = "Per5080G"
tablecode = ": Table D"
subProcCaption = "Validating 80GD"
settableinfo80G (tablename)
noOfProcessSub = end_norows
ReDim GD80DoneeName(end_noofvalues)
ReDim GD80DoneeAdrDet(end_noofvalues)
ReDim GD80DoneeCityTownn(end_noofvalues)
ReDim GD80DoneeState(end_noofvalues)
ReDim GD80DoneePinCode(end_noofvalues)
ReDim GD80DoneePAN(end_noofvalues)
ReDim GD80DoneeARN(end_noofvalues)
ReDim GD80DoneeAmnt(end_noofvalues)
ReDim GD80DoneeInCash(end_noofvalues)
ReDim GD80DoneeTotalDonation(end_noofvalues)
ReDim GD80DoneeEligibleAmnt(end_noofvalues)
ReDim GD80DoneeTraref(end_noofvalues)
ReDim GD80DoneeIFSC(end_noofvalues)
cntr2 = 0

For cntr1 = 1 To end_norows
    If Not ifanyEmpty80G(cntr1) Then
    If errmsg80G = "" Then
        GD80DoneeName(cntr2) = tempDoneeName
        GD80DoneeAdrDet(cntr2) = tempDoneeAddr
        GD80DoneeCityTownn(cntr2) = tempDoneeCity
        GD80DoneeState(cntr2) = tempDoneeState
        GD80DoneePinCode(cntr2) = tempDoneePinCode
        GD80DoneePAN(cntr2) = tempDoneePAN
        GD80DoneeARN(cntr2) = tempDoneeARN
        GD80DoneeAmnt(cntr2) = tempDoneeAmnt
        GD80DoneeInCash(cntr2) = tempDoneeCash
        GD80DoneeTraref(cntr2) = tempDoneeTraref
        GD80DoneeIFSC(cntr2) = tempDoneeIFSC
        GD80DoneeTotalDonation(cntr2) = Sheet6.Range(tablename & ".TotDonation").item(cntr1).value
        GD80DoneeEligibleAmnt(cntr2) = Sheet6.Range(tablename & ".EligibleAmt").item(cntr1).value
        cntr2 = cntr2 + 1
        End If
    Else
        msgError = msgError & errmsg80G & Chr(13)
    End If
    UpdateProgressBar
Next
'Change-ANK
If Not ValidateARNnumber_80GD Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If
If Not ValidatePAN_ARN_80GD Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If

'If Not ValidateComb_80GDA Then
'If errmsg80G <> "" Then
'msgError = msgError & errmsg80G & Chr(13)
'End If
'End If

'Ayush_15/03/2026========================
If Not ValidateTransaction_80GD Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If

If Not ValidateTransaction_80GDB Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If

If Not ValidateTransaction_80GDC Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If

If Not ValidateTransaction_80GDD Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If

'Ayush_03/04/2026
If Not ValidatePAN_Comb90 Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If

If Not ValidatePAN_CashRestriction Then
If errmsg80G <> "" Then
msgError = msgError & errmsg80G & Chr(13)
End If
End If
'--------------------


'========================================

GD80DoneeTot_InCash = Sheet6.Range("Per5080G.TotDoncash100Percent").value
GD80DoneeTotAmnt = Sheet6.Range("Per5080G.TotDon100Percent").value
GD80DoneeTot_TotalDonation = Sheet6.Range("Per5080G.TotTotalDonation100Percent").value
GD80DoneeTotEligAmnt = Sheet6.Range("Per5080G.TotElig100Percent").value

G80DoneeTotAmnt = Sheet6.Range("Per5080G.TotalDonationsUs80G").value
G80DoneeTotEligAmnt = Sheet6.Range("Per5080G.TotalEligibleDonationsUs80G").value
G80DoneeTotCash = Sheet6.Range("Per5080G.TotalCashDonationsUs80G").value
G80DoneeTotTotalDonation = Sheet6.Range("Per5080G.TotalOfTotalDonationsUs80G").value

If msgError <> "" Then validate80G = False

End Function
'Change-07.02.2023.101.80G.02A
Function ValidatePAN_ARN_80GD() As Boolean
errmsg80G = ""
ValidatePAN_ARN_80GD = True
'if pan and arn are same, then show error
Dim i, j As Long
For i = 1 To Range("Per5080G.DoneeARN").count
    'If Range("Per5080G.DoneePAN").Cells(i, 1).value <> "" And Range("Per5080G.DoneeARN").Cells(i, 1).value <> "" Then
    If Range("Per5080G.DoneePAN").Cells(i, 1).value <> "" Then
    For j = i + 1 To Range("Per5080G.DoneeARN").count
        If Range("Per5080G.DoneePAN").Cells(i, 1).value = Range("Per5080G.DoneePAN").Cells(j, 1).value And _
            Range("Per5080G.DoneeARN").Cells(i, 1).value = Range("Per5080G.DoneeARN").Cells(j, 1).value Then
'            Change.02.03.2023.102.IDS.82
'            errmsg80G = errmsg80G & "*Donation to same donee cannot be entered more than once, please enter different ARN, if applicable."
            errmsg80G = errmsg80G & "*Donation to same donee cannot be entered more than once, kindly enter gross amount of donation or enter different ARN, if applicable."
'            End Change.IDS.82
            ValidatePAN_ARN_80GD = False
            Exit Function
        End If
    Next
    End If
Next
End Function
'Change-07.02.2023.101.80G.02B
Function ValidateARNnumber_80GD() As Boolean
errmsg80G = ""
    ValidateARNnumber_80GD = True

    Dim i, j As Long
    Dim PANValue As Variant
    For i = 1 To Range("Per5080G.DoneeARN").count
            If (Range("Per5080G.DoneeARN").Cells(i, 1).value <> "") Then
                If Not checkfieldspecialcharacter1(Range("Per5080G.DoneeARN").Cells(i, 1).value) Then
                    errmsg80G = errmsg80G & "* Special Characters are not allowed in ARN Number at Sr.NO " & i & " in Schedule 80G_D" & Chr(13)
                    ValidateARNnumber_80GD = False
                End If
            End If
    Next
    
End Function
'Ayush_15/03/2026=================
Function ValidateTransaction_80GD() As Boolean
errmsg80G = ""
    ValidateTransaction_80GD = True

    Dim i, j As Long
  '  Dim PANValue As Variant
    For i = 1 To Range("Per10080G.Traref").count
            If (Range("Per10080G.Traref").Cells(i, 1).value <> "") Then
                If Not checkfieldspecialcharacter_Transaction(Range("Per10080G.Traref").Cells(i, 1).value) Then
                    errmsg80G = errmsg80G & "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.NO " & i & " in Schedule 80G_A" & Chr(13)
                    ValidateTransaction_80GD = False
                End If
            End If
    Next
    
End Function
'Ayush_30/03
'Ayush_15/03/2026=================
'Function ValidateComb_80GDA() As Boolean
'errmsg80G = ""
'    ValidateComb_80GDA = True
'
'    Dim i, j As Long
'  '  Dim PANValue As Variant
'    For i = 1 To Range("Per10080G.Donationcash").count
'            If (Range("Per10080G.Donationcash").Cells(i, 1).value <> "") Then
'                'If Not checkfieldspecialcharacter_Transaction(Range("Per10080G.Traref").Cells(i, 1).value) Then
'                    If Sheet6.Range("comb_90").value = 2 Then
'                    errmsg80G = errmsg80G & "* sjfhfsdhjsfdjh" & Chr(13)
'                    ValidateComb_80GDA = False
'                End If
'            End If
'    Next
'
'End Function
'---------
'=================================
'Ayush_15/03/2026=================
Function ValidateTransaction_80GDB() As Boolean
errmsg80G = ""
    ValidateTransaction_80GDB = True

    Dim i, j As Long
  '  Dim PANValue As Variant
    For i = 1 To Range("PerNO5080G.Traref").count
            If (Range("PerNO5080G.Traref").Cells(i, 1).value <> "") Then
                If Not checkfieldspecialcharacter_Transaction(Range("PerNO5080G.Traref").Cells(i, 1).value) Then
                errmsg80G = errmsg80G & "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.NO " & i & " in Schedule 80G_B" & Chr(13)
                ValidateTransaction_80GDB = False
                End If
            End If
    Next
    
End Function
'=================================
'Ayush_15/03/2026=================
Function ValidateTransaction_80GDC() As Boolean
errmsg80G = ""
    ValidateTransaction_80GDC = True

    Dim i, j As Long
  '  Dim PANValue As Variant
    For i = 1 To Range("PerYES10080G.Traref").count
            If (Range("PerYES10080G.Traref").Cells(i, 1).value <> "") Then
                If Not checkfieldspecialcharacter_Transaction(Range("PerYES10080G.Traref").Cells(i, 1).value) Then
                    errmsg80G = errmsg80G & "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.NO " & i & " in Schedule 80G_C" & Chr(13)
                    ValidateTransaction_80GDC = False
                End If
            End If
    Next
    
End Function
'=================================
'Ayush_15/03/2026=================
Function ValidateTransaction_80GDD() As Boolean
errmsg80G = ""
    ValidateTransaction_80GDD = True

    Dim i, j As Long
  '  Dim PANValue As Variant
    For i = 1 To Range("Per5080G.Traref").count
            If (Range("Per5080G.Traref").Cells(i, 1).value <> "") Then
                If Not checkfieldspecialcharacter_Transaction(Range("Per5080G.Traref").Cells(i, 1).value) Then
                    errmsg80G = errmsg80G & "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction only special characters / or - are allowed at Sr.NO " & i & " in Schedule 80G_D" & Chr(13)
                    ValidateTransaction_80GDD = False
                End If
            End If
    Next
    
End Function
'=================================


Function ifanyEmpty80G(counter As Variant) As Boolean
On Error Resume Next
ifanyEmpty80G = False
errmsg80G = ""

tempDoneeName = Trim(Sheet6.Range(tablename & ".DoneeName").item(counter).value)
tempDoneeAddr = Trim(Sheet6.Range(tablename & ".AddrDetail").item(counter).value)
tempDoneeCity = Trim(Sheet6.Range(tablename & ".CityOrTownOrDistrict").item(counter).value)
tempDoneeState = Trim(Sheet6.Range(tablename & ".StateCode").item(counter).value)
tempDoneePinCode = Trim(Sheet6.Range(tablename & ".PinCode").item(counter).value)
tempDoneePAN = Trim(Sheet6.Range(tablename & ".DoneePAN").item(counter).value)
tempDoneeARN = Trim(Sheet6.Range(tablename & ".DoneeARN").item(counter).value)
tempDoneeAmnt = Trim(Sheet6.Range(tablename & ".DonationAmt").item(counter).value)
tempDoneeCash = Trim(Sheet6.Range(tablename & ".Donationcash").item(counter).value)
tempDoneeTraref = Trim(Sheet6.Range(tablename & ".Traref").item(counter).value)
tempDoneeIFSC = Trim(Sheet6.Range(tablename & ".IFSC").item(counter).value)

ifanyEmpty80G = True
If Len(tempDoneeName) > 125 Then
    errmsg80G = errmsg80G & "* Please enter name of donee at Sr no " & counter & " cannot exceed 125 characters in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneeAddr) > 200 Then
    errmsg80G = errmsg80G & "* Please enter address of donee at Sr no " & counter & " cannot exceed 200 characters in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneeCity) > 50 Then
    errmsg80G = errmsg80G & "* Please enter city/town/district at Sr no " & counter & " cannot exceed 50 characters in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneePinCode) > 6 Then
    errmsg80G = errmsg80G & "* Please enter pin code of donee at Sr no " & counter & " cannot exceed 6 digits in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneePAN) > 10 Then
    errmsg80G = errmsg80G & "* Please enter Donee PAN at Sr no " & counter & " cannot exceed 10 characters in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneeTraref) > 50 Then
    errmsg80G = errmsg80G & "* please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution  transaction at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneeIFSC) > 11 Then
    errmsg80G = errmsg80G & "* Please enter ""your bank IFSC from which Contribution  is made"" at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
End If

If Len(tempDoneeAmnt) > 0 Then
    If ((Not IsNumeric(tempDoneeAmnt)) Or (tempDoneeAmnt < 0)) Then
        errmsg80G = errmsg80G & "* Donation Amount at Sr no " & counter & " should be Numeric, Non Negative, not exceeding 14 digits in Sheet 80G" & tablecode & Chr(13)
        GoTo below
    End If
    If (tempDoneeAmnt) > 99999999999999# Then
        errmsg80G = errmsg80G & "* Donation Amount at Sr no " & counter & " is cannot exceed 14 digits in Sheet 80G" & tablecode & Chr(13)
        GoTo below
    End If
End If

below:
If errmsg80G = "" Then
    ifanyEmpty80G = False
End If

'If Not (tempDoneeName = "" And tempDoneeAddr = "" And tempDoneeCity = "" And CheckEmptyDDList(tempDoneeState) And tempDoneePinCode = "" And tempDoneePAN = "" And (tempDoneeAmnt = "" Or tempDoneeCash = "")) Then
If Not (tempDoneeName = "" And tempDoneeAddr = "" And tempDoneeCity = "" And CheckEmptyDDList(tempDoneeState) And tempDoneePinCode = "" And tempDoneePAN = "" And tempDoneeAmnt = "" And tempDoneeCash = "" And tempDoneeTraref = "" And tempDoneeIFSC = "") Then
    ifanyEmpty80G = True
    If tempDoneeName = "" Then errmsg80G = errmsg80G & "* Please enter Name of Donee at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    If tempDoneeAddr = "" Then errmsg80G = errmsg80G & "* Please enter Address of Donee at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    If tempDoneeCity = "" Then errmsg80G = errmsg80G & "* Please enter City/Town/District of Donee at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    If mdCommon.CheckEmptyDDList(tempDoneeState) Then errmsg80G = errmsg80G & "* Please select the State of Donee at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    If tempDoneePinCode = "" Then errmsg80G = errmsg80G & "* Please enter Pincode of Donee at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    'If Len("" & tempDoneePinCode) > 6 Then
     If (tempDoneePinCode <> "" And Not IsNumeric(tempDoneePinCode)) Then errmsg80G = errmsg80G & "* Donee Pincode at Sr no " & counter & "is invalid in Sheet 80G" & tablecode & Chr(13)
    'End If
   ' If tempDoneeAmnt > 0 AndAlso tempDoneeAmnt = CInt(tempDoneeAmnt) Then
    If tempDoneeAmnt > 0 And tempDoneeAmnt <> "" Then
   
'If IsNumeric(tempDoneeAmnt) _
'   And CDbl(tempDoneeAmnt) > 0 _
'   And CDbl(tempDoneeAmnt) = Fix(CDbl(tempDoneeAmnt)) Then

    If tempDoneeTraref = "" Then errmsg80G = errmsg80G & "* Please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution  transaction"" at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    If tempDoneeIFSC = "" Then errmsg80G = errmsg80G & "* Please enter ""your bank IFSC from which Contribution  is made"" at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    End If
    If tempDoneePAN = "" Then
        errmsg80G = errmsg80G & "* Please enter the PAN of Donee at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    Else
'        If Not CheckPAN(tempDoneePAN) Then errmsg80G = errmsg80G & "* Donee PAN at Sr no " & counter & " is not Valid in Sheet 80G" & tablecode & Chr(13)
'       SIT-56341
        If Not CheckPAN(tempDoneePAN) Then errmsg80G = errmsg80G & "* Donee PAN at Sr no " & counter & " Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Sheet 80G" & tablecode & Chr(13)
        If Not CheckPanWithAssessPAN(tempDoneePAN) Then errmsg80G = errmsg80G & "* Donee PAN cannot be same as assesse PAN or verification PAN at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    End If
    


    If ((tempDoneeAmnt = "" Or tempDoneeAmnt = 0) And (tempDoneeCash = "" Or tempDoneeCash = 0)) Then errmsg80G = errmsg80G & "* Please enter amount of donation either in field ""Donation in cash"" or ""Donation in other mode""  at Sr no " & counter & " in Sheet 80G" & tablecode & Chr(13)
    If errmsg80G = "" Then ifanyEmpty80G = False
    
End If
End Function


Function CheckPanWithAssessPAN(DoneePan As Variant) As Boolean
CheckPanWithAssessPAN = True

If ((UCase(DoneePan) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(DoneePan) = UCase(Sheet5.Range("Ver.PAN").value))) Then
    CheckPanWithAssessPAN = False
End If

End Function

'Ayush_03/04/2026
Function ValidatePAN_Comb90() As Boolean

    errmsg80G = ""
    ValidatePAN_Comb90 = True

    Dim i As Long
    Dim comb90Val As Long
    Dim panVal As String
    Dim lastrow As Long

    'Loop through PAN column (assumed column J = 10)
    lastrow = Sheet6.Cells(Sheet6.Rows.count, "J").End(xlUp).row
    'For i = 1 To Sheet6.Range("J:J").Rows.count
    For i = 1 To lastrow

        panVal = Trim(Sheet6.Cells(i, 10).value)  'PAN column
        comb90Val = Sheet6.Cells(i, "X").value    'comb_90 column

        If panVal <> "" Then
            If comb90Val > 0 Then
'                errmsg80G = errmsg80G & _
'                    "* Same PAN cannot be entered in other block at Sr.NO " & _
'                    i & Chr(13)
                 errmsg80G = errmsg80G & _
                    "* Same PAN cannot be entered in other block." & Chr(13)
                ValidatePAN_Comb90 = False
            End If
        End If

    Next i

End Function
'Ayush_03/04/2026
Function ValidatePAN_CashRestriction() As Boolean

    Dim panCell As Range
    Dim panRanges As Range
    Dim cntOnes As Long
    Dim panVal As String
    Dim lastRow1 As Long

    errmsg80G = ""
    ValidatePAN_CashRestriction = True

    '-----------------------------------
    ' Combine all PAN named ranges
    '-----------------------------------
    Set panRanges = Union( _
        Sheet6.Range("Per10080G.DoneePAN"), _
        Sheet6.Range("PerNO5080G.DoneePAN"), _
        Sheet6.Range("Per5080G.DoneePAN"), _
        Sheet6.Range("PerYES10080G.DoneePAN") _
    )

    '-----------------------------------
    ' Check helper column V
    '-----------------------------------
    lastRow1 = Sheet6.Cells(Sheet6.Rows.count, "V").End(xlUp).row
    cntOnes = Application.CountIf(Sheet6.Range("V2:V" & lastRow1), 1)
 '   cntOnes = Application.CountIf(Sheet6.Range("V:V"), 1)
    If cntOnes <= 1 Then Exit Function

    '-----------------------------------
    ' Loop through PAN cells
    '-----------------------------------
    For Each panCell In panRanges.Cells

        panVal = Trim(panCell.value)

        If panVal <> "" Then

           

            'Rule 2 : Donation in cash restriction
'            errmsg80G = errmsg80G & _
'                "* Same PAN cannot be entered for ""Donation in cash"" at Row " & _
'                panCell.row & Chr(13)
            errmsg80G = errmsg80G & _
                "* Same PAN cannot be entered for ""Donation in cash""" & Chr(13)

            ValidatePAN_CashRestriction = False
            Exit Function

        End If
    Next panCell

End Function

