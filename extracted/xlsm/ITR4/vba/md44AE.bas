Attribute VB_Name = "md44AE"
Option Explicit

Dim end_noofrow As Variant
Dim end_noofvalues As Variant
Dim errorMsg_44AE As Variant
Dim tablename As Variant

Public HVHoldingPeriod As Variant
Public HVIncome As Variant
Public HVDeemedIncome As Variant
Public HVPresumptiveIncome As Variant


Sub cmdValidate44AE_Click()
ValidateSheet44AE
fmsgboxStatus "Sheet is Ok" ', vbOKOnly, vbMessageCaption
End Sub

'Sub cmdPrev_Click_44AE()
'Sheet1.Activate
'End Sub

'Sub cmdNext_Click_44AE()
'Sheet3.Activate
'End Sub

Sub CommandButton21_Click()
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheets("Help").Visible = xlSheetVisible
    Sheets("Help").Activate
    ThisWorkbook.Protect Password:=getmsgstate
End Sub


Sub ValidateSheet44AE()
msgError = ""
If Not validate44AE() Then
Sheet2.Activate
fmsgbox msgError ', vbOKOnly, "Error(s)!"
CloseMsg
End If

If Range("HVPresumptiveInc").value > 99999999999999# Then
    fmsgboxStatus "Presumptive Income from Goods Carriage cannot exceed 14 digits" & Chr(13) ', "Error(s)!"
    CloseMsg
End If


End Sub

Sub tablesinfo44AE()

Dim rangecells1 As Range
Dim rangecells2 As Range
Dim mIntCells As Long
Dim mIntCtr As Long
Dim ccount As Long
ccount = 0

Set rangecells1 = Sheet2.Range(tablename & "V_HoldingPeriod").Cells
Set rangecells2 = Sheet2.Range(tablename & "V_Income").Cells
mIntCells = Sheet2.Range(tablename & "V_HoldingPeriod").count

For mIntCtr = 1 To mIntCells
    If Not IsEmpty(rangecells1.item(mIntCtr).value) Or Not IsEmpty(rangecells2.item(mIntCtr).value) Then ccount = ccount + 1
Next
end_noofvalues = ccount
End Sub

Function validate44AE() As Boolean
On Error Resume Next
validate44AE = True


Dim cntr1, cntr2 As Variant
Dim holdingTemp, incomeTemp, deemedincomeTemp As Variant
Dim noofHeavyVehicle, noofLightVehicle As Variant
end_noofvalues = 0
end_noofrow = 10

tablename = "H"
tablesinfo44AE
ReDim HVHoldingPeriod(end_noofvalues)
ReDim HVIncome(end_noofvalues)
ReDim HVDeemedIncome(end_noofvalues)
cntr2 = 0

subProcCaption = "Valiadting 44AE"
noOfProcessSub = end_noofrow

For cntr1 = 1 To end_noofrow
    holdingTemp = Trim(Range("HV_HoldingPeriod").item(cntr1).value)
    incomeTemp = Trim(Range("HV_Income").item(cntr1).value)
    deemedincomeTemp = Trim(Range("AE44.HEE3").item(cntr1).value)

    If (holdingTemp <> "" And Not incomeTemp <> "") Or (incomeTemp <> "" And Not holdingTemp <> "") Then
        If holdingTemp <> "" Then msgError = msgError & "* Income Per Vehicle at Sr. no " & cntr1 & " is Mandatory for Goods Carriage" & Chr(13)
        If incomeTemp <> "" Then msgError = msgError & "* Holding Per Vehicle at Sr. no " & cntr1 & " is Mandatory for Goods Carriage" & Chr(13)

    Else
    If ((holdingTemp <> "") And (incomeTemp <> "")) Then
        HVHoldingPeriod(cntr2) = holdingTemp
        HVIncome(cntr2) = incomeTemp
        HVDeemedIncome(cntr2) = deemedincomeTemp
        If HVHoldingPeriod(cntr2) > 12 Or HVHoldingPeriod(cntr2) < 1 Then msgError = msgError & "* Holding Period as Sr. no " & cntr1 & " cannot be more than 12" & Chr(13)
        If HVIncome(cntr2) < 7500 Or HVIncome(cntr2) < 0 Then msgError = msgError & "* Income Per vehicle at Sr. no " & cntr1 & " cannot be less than 5000" & Chr(13)
        cntr2 = cntr2 + 1
        End If
    End If

 UpdateProgressBar
Next
noofHeavyVehicle = cntr2
If noofHeavyVehicle > 10 Then msgError = msgError & "* No of vehicles cannot be more than 10. For more than 10 vehicles, pl fill Form ITR 3"
HVPresumptiveIncome = Trim(Range("HVPresumptiveInc").value)
If msgError <> "" Then validate44AE = False

End Function

