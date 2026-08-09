Attribute VB_Name = "ImportPrefill"

Option Explicit
'03/01/2021
Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    
    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options
Dim rngname_TCS_Prfl, rngname_TDS2_Prfl, end_TDS2oth

'Malli
Public incDeductionsOthIncCPC_TAX_chk As Boolean
'-----------------
'Malli_02/06/2026
Public CG_POPUPPfl As Boolean
'-----------------

Sub ImportPrefill()
Dim filePath, jsonText, bacVal As String

MsgBox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."

With Application.FileDialog(msoFileDialogFilePicker)
    'Makes sure the user can select only one file
    .AllowMultiSelect = False
    .Title = "Please select a Json file."
    'Filter to just the following types of files to narrow down selection options
    .Filters.add "Json File", "*.json", 1
    'Show the dialog box
    If .Show = True Then
        filePath = .SelectedItems.item(1)
    Else
        Exit Sub
    End If
    On Error Resume Next
End With

Open filePath For Input As #1
jsonText = Input$(LOF(1), 1)
Close #1

ImportPersonalInfo_pfl (jsonText)
ImportPartA_139_8A (jsonText)
ImportScheduleIT (jsonText)
'ImportPL_pfl (jsonText)
'ImportNatureOfBusiness_pfl (jsonText)
ImportFilingStatus_pfl (jsonText)
ImportRefund_pfl (jsonText)

'ScheduleTCSImport_pfl (jsonText)
ScheduleTCSImport_pfl1 (jsonText)
ImportScheduleTDS2_pfl (jsonText)
ImportTDSonSalary_pfl (jsonText)
ImportPartB_TTI (jsonText)
'ImportScheduleS (jsonText)
ImportScheduleS_23_24 (jsonText)
ScheduleCFLImport (jsonText)
ImportScheduleHP (jsonText)
ImportScheduleOS (jsonText)
AMTCImport (jsonText)
ImportScheduleFA (jsonText)
ImportScheduleAL (jsonText)
ImportScheduleFSI (jsonText)
ImportScheduleEI (jsonText)
'ImportSchedule80D (jsonText)
ImportScheduleSPI (jsonText)

ScheduleVIAImport (jsonText)

ImportSchedulePTI (jsonText)
ImportSchedule5A (jsonText)
ScheduleTDS1Import (jsonText)
ScheduleTDS3Import (jsonText)
'ImportScheduleTCS (jsonText)
'ImportTradingAccount (jsonText)
'ImportManufacturingAccount (jsonText)
'ImportPARTA_OI (jsonText)
'ImportPARTA_QD (jsonText)
'ImportITR3ScheduleBP (jsonText)
ScheduleDPMImport (jsonText)
ESRImport (jsonText)
ImportScheduleCG_POPUPPfl (jsonText)   'Malli_AY_2026_27  02/06/2026
'ImportITR3ScheduleUD (jsonText)
'ScheduleICDSImport (jsonText)
'ScheduleTPSAImport (jsonText)
'ScheduleGSTImport (jsonText)
'ImportScheduleIF_pfl (jsonText)
'ImportSchedule80 (jsonText)
ImportScheduleESOP (jsonText)

ImportVerification_pfl (jsonText)





 




'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

'ImportSchedule80GD (jsonText)
'ImportSchedule80D (jsonText)

'If (SheetALL.Range("bacValue").value) = 2 Then
'    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" Then
'        ImportSchedule_80E_Pfl (jsonText)
'        ImportSchedule_80EE_pfl (jsonText)
'        ImportSchedule_80EEA_pfl (jsonText)
'        ImportSchedule_80EEB_pfl (jsonText)
'    End If
'End If

'Commented by malli as per AY_2026-27 prefill posibility removed
'Malli added
'If SheetALL.Range("bacValue").value = 2 Then
'ImportSchedule80DD_80U_pfl (jsonText)
'End If
'--------------------------------------

'----------------------------------------------

End Sub

Function ImportScheduleIF_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

Dim jsonDictionary, init As Object
Dim Node, Nodelist As Object
Dim TotalProfitShareAmt, rowcount2, TotalFirmCapBalOn31Mar As Variant
Dim NameOfFirmColNo, PANColNo, FirmNameColNo, FirmPANColNo, IsLiableToAuditColNo, Is92EApplicableColNo, ProfitSharePercentColNo, ProfitShareAmtColNo As Variant
Dim FirmCapBalOn31MarColNo As Variant
Dim TotalXMLRow, rowcount, cnt, RecFirm  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long



Set init = jsonObject("lastFiledITR")("scheduleIF")


If init <> Null And init <> Empty Then
If init.exists("partnerFirmDetails") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleIF")("partnerFirmDetails")

FirmNameColNo = Sheet21.Range("IF.FirmName").Column
FirmPANColNo = Sheet21.Range("IF.FirmPAN").Column
IsLiableToAuditColNo = Sheet21.Range("IF.IsLiableToAudit").Column
Is92EApplicableColNo = Sheet21.Range("IF.Is92EApplicable").Column
ProfitSharePercentColNo = Sheet21.Range("IF.ProfitSharePercent").Column
ProfitShareAmtColNo = Sheet21.Range("IF.ProfitShareAmt").Column
FirmCapBalOn31MarColNo = Sheet21.Range("IF.FirmCapBalOn31Mar").Column
NameOfFirmColNo = Sheet1.Range("Sheet1.NameOfFirm").Column
PANColNo = Sheet1.Range("Sheet1.FirmPAN").Column

TotalExRow = Range("IF.FirmName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet21.Range("IF.FirmName").Locked = False Then
            Sheet21.Range("IF.FirmName").ClearContents
        End If
        If Sheet21.Range("IF.FirmPAN").Locked = False Then
            Sheet21.Range("IF.FirmPAN").ClearContents
        End If
        If Sheet21.Range("IF.IsLiableToAudit").Locked = False Then
            Sheet21.Range("IF.IsLiableToAudit").ClearContents
        End If
        If Sheet21.Range("IF.Is92EApplicable").Locked = False Then
            Sheet21.Range("IF.Is92EApplicable").ClearContents
        End If
        If Sheet21.Range("IF.ProfitSharePercent").Locked = False Then
            Sheet21.Range("IF.ProfitSharePercent").ClearContents
        End If
        If Sheet21.Range("IF.ProfitShareAmt").Locked = False Then
            Sheet21.Range("IF.ProfitShareAmt").ClearContents
        End If
        If Sheet21.Range("IF.FirmCapBalOn31Mar").Locked = False Then
            Sheet21.Range("IF.FirmCapBalOn31Mar").ClearContents
        End If
        If Sheet1.Range("Sheet1.NameOfFirm").Locked = False Then
            Sheet1.Range("Sheet1.NameOfFirm").ClearContents
        End If
        If Sheet1.Range("Sheet1.FirmPAN").Locked = False Then
            Sheet1.Range("Sheet1.FirmPAN").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_FIRM (TotalDiffRow)
        AddDiffRows_FirmDetails (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet21.Range("IF.FirmName").name)
    rowcount2 = getRowNo(Sheet1.Range("Sheet1.NameOfFirm").name) - 1
    rowcount = rowcount - 1
    cnt = 0
    Dim IsLiableToAudit
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        rowcount2 = rowcount2 + 1
        If Sheet21.Cells(rowcount, FirmNameColNo).Locked = False Then
            Sheet21.Cells(rowcount, FirmNameColNo).value = Node("firmName")
        End If
        If Sheet21.Cells(rowcount, FirmPANColNo).Locked = False Then
            Sheet21.Cells(rowcount, FirmPANColNo).value = Node("firmPAN")
        End If
        If Sheet21.Cells(rowcount, IsLiableToAuditColNo).Locked = False Then
            Sheet21.Cells(rowcount, IsLiableToAuditColNo).value = Node("isLiableToAudit")
        End If
        If Sheet21.Cells(rowcount, Is92EApplicableColNo).Locked = False Then
            Sheet21.Cells(rowcount, Is92EApplicableColNo).value = Node("sec92EFirmFlag")
        End If
        If Sheet21.Cells(rowcount, ProfitSharePercentColNo).Locked = False Then
            Sheet21.Cells(rowcount, ProfitSharePercentColNo).value = Node("profitSharePercent")
        End If
        
        If Sheet1.Cells(rowcount2, NameOfFirmColNo).Locked = False Then
            Sheet1.Cells(rowcount2, NameOfFirmColNo).value = Node("firmName")
        End If
        If Sheet1.Cells(rowcount2, PANColNo).Locked = False Then
            Sheet1.Cells(rowcount2, PANColNo).value = Node("firmPAN")
        End If
'        If Sheet21.Cells(rowcount, ProfitShareAmtColNo).Locked = False Then
'            Sheet21.Cells(rowcount, ProfitShareAmtColNo).value = Node("ProfitShareAmt")
'        End If
'        If Sheet21.Cells(rowcount, FirmCapBalOn31MarColNo).Locked = False Then
'            Sheet21.Cells(rowcount, FirmCapBalOn31MarColNo).value = Node("FirmCapBalOn31Mar")
'        End If
        
        cnt = cnt + 1
        
    Next Node
    RecFirm = cnt
 End If
 End If
 
'If TotalProfitShareAmt <> "" and Sheet21.Range("IF.TotalProfitShareAmt").locked=false Then
'    Sheet21.Range("IF.TotalProfitShareAmt").value = TotalProfitShareAmt
'End If
'
'If TotalFirmCapBalOn31Mar <> "" and Sheet21.Range("IF.TotalFirmCapBalOn31Mar").locked=false Then
'    Sheet21.Range("IF.TotalFirmCapBalOn31Mar").value = TotalFirmCapBalOn31Mar
'End If

End Function
Function ImportPL_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

Dim init, Nodelist, Node, cnt, TotalExRow, TotalXMLRow, TotalDiffRow, rowcount
Dim iNob As Variant
Dim sNob As Variant
Dim iAEOLE As Variant
Dim sAEOLE As Variant

If Sheet4.Range("PL.ProfitOnConvOfInventary").Locked = False Then
    Sheet4.Range("PL.ProfitOnConvOfInventary").value = jsonObject("form3CD")("PartAPL")("CreditsToPL")("OthIncome")("ProfitOnCnvInvntryToCapAsst")
End If


Set init = jsonObject("lastFiledITR")
If init.exists("natOfBus44AD") Then

    Dim NOBCodeADColNo, NOBNameADColNo, NOBDescriptionADColNo, RecNatOfBus44AD As Variant

    Set Nodelist = jsonObject("lastFiledITR")("natOfBus44AD")
    
    NOBNameADColNo = Range("Sec44AD.NOBName").Column
    NOBCodeADColNo = Range("Sec44AD.NOBCode").Column
    NOBDescriptionADColNo = Range("Sec44AD.NOBDescription").Column

    TotalExRow = Range("Sec44AD.NOBName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet4.Range("Sec44AD.NOBName").Locked = False Then
            Sheet4.Range("Sec44AD.NOBName").ClearContents
        End If
        If Sheet4.Range("Sec44AD.NOBCode").Locked = False Then
            Sheet4.Range("Sec44AD.NOBCode").ClearContents
        End If
        If Sheet4.Range("Sec44AD.NOBDescription").Locked = False Then
            Sheet4.Range("Sec44AD.NOBDescription").ClearContents
        End If
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("Sec44AD.NOBName").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
             
            
            
            iNob = Node("codeAD")
            sNob = Findtext(iNob, "NOB44AD")
        
        If Sheet4.Cells(rowcount, NOBNameADColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBNameADColNo).value = Node("nameOfBusiness")
        End If
        If Sheet4.Cells(rowcount, NOBCodeADColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBCodeADColNo).value = sNob
        End If
        If Sheet4.Cells(rowcount, NOBDescriptionADColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBDescriptionADColNo).value = Node("description")
        End If
        cnt = cnt + 1
        
    Next Node
     
    RecNatOfBus44AD = cnt
End If



'"NatOfBus44ADA" start
Set init = jsonObject("lastFiledITR")
If init.exists("natOfBus44ADA") Then

    Dim NOBName1ColNo, NOBCode1ColNo, NOBDescription1ColNo, RecNatOfBus44ADA As Variant

    Set Nodelist = jsonObject("lastFiledITR")("natOfBus44ADA")
    
    NOBName1ColNo = Range("Sec44AD.NOBName1").Column
    NOBCode1ColNo = Range("Sec44AD.NOBCode1").Column
    NOBDescription1ColNo = Range("Sec44AD.NOBDescription1").Column

    TotalExRow = Range("Sec44AD.NOBName1").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet4.Range("Sec44AD.NOBName1").Locked = False Then
            Sheet4.Range("Sec44AD.NOBName1").ClearContents
        End If
        If Sheet4.Range("Sec44AD.NOBCode1").Locked = False Then
            Sheet4.Range("Sec44AD.NOBCode1").ClearContents
        End If
        If Sheet4.Range("Sec44AD.NOBDescription1").Locked = False Then
            Sheet4.Range("Sec44AD.NOBDescription1").ClearContents
        End If
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("Sec44AD.NOBName1").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
        iNob = Node("codeADA")
        sNob = Findtext(iNob, "NOB44ADA")
        
        If Sheet4.Cells(rowcount, NOBName1ColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBName1ColNo).value = Node("nameOfBusiness")
        End If
        If Sheet4.Cells(rowcount, NOBCode1ColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBCode1ColNo).value = sNob
        End If
        If Sheet4.Cells(rowcount, NOBDescription1ColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBDescription1ColNo).value = Node("description")
        End If
        
        cnt = cnt + 1
        
    Next Node
     
    RecNatOfBus44ADA = cnt
End If



'"NatOfBus44AE" start
Set init = jsonObject("lastFiledITR")
If init.exists("natOfBus44AE") Then

    Dim NOBName2ColNo, NOBCode2ColNo, NOBDescription2ColNo, RecNatOfBus44AE As Variant

    Set Nodelist = jsonObject("lastFiledITR")("natOfBus44AE")
    
    NOBName2ColNo = Range("Sec44AD.NOBName2").Column
    NOBCode2ColNo = Range("Sec44AD.NOBCode2").Column
    NOBDescription2ColNo = Range("Sec44AD.NOBDescription2").Column

    TotalExRow = Range("Sec44AD.NOBName2").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet4.Range("Sec44AD.NOBName2").Locked = False Then
            Sheet4.Range("Sec44AD.NOBName2").ClearContents
        End If
        If Sheet4.Range("Sec44AD.NOBCode2").Locked = False Then
            Sheet4.Range("Sec44AD.NOBCode2").ClearContents
        End If
        If Sheet4.Range("Sec44AD.NOBDescription2").Locked = False Then
            Sheet4.Range("Sec44AD.NOBDescription2").ClearContents
        End If
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("Sec44AD.NOBName2").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
        iNob = Node("codeAE")
        sNob = Findtext(iNob, "NOB")
    
        If Sheet4.Cells(rowcount, NOBName2ColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBName2ColNo).value = Node("nameOfBusiness")
        End If
        If Sheet4.Cells(rowcount, NOBCode2ColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBCode2ColNo).value = sNob
        End If
        If Sheet4.Cells(rowcount, NOBDescription2ColNo).Locked = False Then
            Sheet4.Cells(rowcount, NOBDescription2ColNo).value = Node("description")
        End If
        cnt = cnt + 1
        
    Next Node
     
    RecNatOfBus44AE = cnt
End If




'"GoodsDtlsUs44AE" start
Set init = jsonObject("lastFiledITR")
If init.exists("goodsDtlsUs44AE") Then
    Dim RegNoColNo, OLHColNo, TonnageCpctColNo, NoOfMonthsColNo, PresumptiveIncomeColNo As Variant

    Set Nodelist = jsonObject("lastFiledITR")("goodsDtlsUs44AE")
    
    RegNoColNo = Sheet4.Range("Sec44AE.RegNo").Column
    OLHColNo = Sheet4.Range("Sec44AE.OLH").Column
    TonnageCpctColNo = Sheet4.Range("Sec44AE.TonnageCpct").Column
    NoOfMonthsColNo = Sheet4.Range("Sheet44AE.NoOfMonths").Column
    PresumptiveIncomeColNo = Sheet4.Range("Sheet44AE.PresumptiveIncome").Column

    TotalExRow = Range("Sec44AE.RegNo").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet4.Range("Sec44AE.RegNo").Locked = False Then
            Sheet4.Range("Sec44AE.RegNo").ClearContents
        End If
        
        If Sheet4.Range("Sec44AE.OLH").Locked = False Then
            Sheet4.Range("Sec44AE.OLH").ClearContents
        End If
        
        If Sheet4.Range("Sec44AE.TonnageCpct").Locked = False Then
            Sheet4.Range("Sec44AE.TonnageCpct").ClearContents
        End If
        
        If Sheet4.Range("Sheet44AE.NoOfMonths").Locked = False Then
            Sheet4.Range("Sheet44AE.NoOfMonths").ClearContents
        End If
        
        If Sheet4.Range("Sheet44AE.PresumptiveIncome").Locked = False Then
            Sheet4.Range("Sheet44AE.PresumptiveIncome").ClearContents
        End If
    End If
    
    rowcount = getRowNo(Sheet4.Range("Sec44AE.RegNo").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
             
      
        
        iAEOLE = Node("ownedLeasedHiredFlag")
          
        If iAEOLE = "OWN" Then
        sAEOLE = "Owned"
        ElseIf iAEOLE = "LEASE" Then
        sAEOLE = "Leased"
        ElseIf iAEOLE = "HIRED" Then
        sAEOLE = "Hired"
        End If
    
        If Sheet4.Cells(rowcount, RegNoColNo).Locked = False Then
            Sheet4.Cells(rowcount, RegNoColNo).value = Node("regNumberGoodsCarriage")
        End If
        
        If Sheet4.Cells(rowcount, OLHColNo).Locked = False Then
            Sheet4.Cells(rowcount, OLHColNo).value = sAEOLE
        End If
        
        If Sheet4.Cells(rowcount, TonnageCpctColNo).Locked = False Then
            Sheet4.Cells(rowcount, TonnageCpctColNo).value = Node("tonnageCapacity")
        End If
        
        If Sheet4.Cells(rowcount, NoOfMonthsColNo).Locked = False Then
            Sheet4.Cells(rowcount, NoOfMonthsColNo).value = Node("holdingPeriod")
        End If
        
'        If Sheet4.Cells(rowcount, PresumptiveIncomeColNo).Locked = False Then
'            Sheet4.Cells(rowcount, PresumptiveIncomeColNo).value = Node("PresumptiveIncome")
'        End If
        
        cnt = cnt + 1
        
    Next Node
     
    RecNatOfBus44AE = cnt
End If

End Function
Function ImportNatureOfBusiness_pfl(jsonText As String)

On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
Dim iNob, sNob, Nodelist, Node, NOB, Trade1, Trade2, RecTDS1, cnt, rowcount, init, TotalExRow, TotalXMLRow, TotalDiffRow

Set init = jsonObject("lastFiledITR")("natOfBus")
If init.exists("NatureOfBusiness") Then
    Set Nodelist = New Collection
    Set Nodelist = jsonObject("lastFiledITR")("natOfBus")("NatureOfBusiness")
    NOB = Sheet2.Range("NOB.Code").Column
    Trade1 = Sheet2.Range("NOB.TradeNameOrDescription").Column
    Trade2 = Sheet2.Range("NOB.TradeNameOrDescription2").Column
    TotalExRow = Range("NOB.Code").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        If Sheet2.Range("NOB.Code").Locked = False Then
            Sheet2.Range("NOB.Code").ClearContents
        End If
        If Sheet2.Range("NOB.TradeNameOrDescription").Locked = False Then
            Sheet2.Range("NOB.TradeNameOrDescription").ClearContents
        End If
        If Sheet2.Range("NOB.TradeNameOrDescription2").Locked = False Then
            Sheet2.Range("NOB.TradeNameOrDescription2").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
         AddDiffRows_NOBP (TotalDiffRow)
    End If
        
    rowcount = getRowNo(Sheet2.Range("NOB.Code").name)
    rowcount = rowcount - 1
    cnt = 0
    Dim tradeName
    For Each Node In Nodelist
        rowcount = rowcount + 1

        iNob = Node("Code")
        sNob = Findtext(iNob, "Nature_of_Business")
        
        If Sheet2.Cells(rowcount, NOB).Locked = False Then
            Sheet2.Cells(rowcount, NOB).value = sNob
        End If
        
        tradeName = ""
        If Node.exists("TradeName1") Then
            tradeName = Node("TradeName1")
            If Node.exists("TradeName2") Then
                tradeName = tradeName & ", " & Node("TradeName2")
            End If
            If Node.exists("TradeName3") Then
                tradeName = tradeName & ", " & Node("TradeName3")
            End If
        ElseIf Node.exists("TradeName2") Then
            tradeName = Node("TradeName2")
            If Node.exists("TradeName3") Then
                tradeName = tradeName & ", " & Node("TradeName3")
            End If
        ElseIf Node.exists("TradeName3") Then
            tradeName = Node("TradeName3")
        End If
        
       
        If Sheet2.Cells(rowcount, Trade1).Locked = False Then
            Sheet2.Cells(rowcount, Trade1).value = tradeName
        End If
        If Sheet2.Cells(rowcount, Trade2).Locked = False Then
            Sheet2.Cells(rowcount, Trade2).value = Node("Description")
        End If
        
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If

End Function
Function ScheduleGSTImport(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

Dim TotalXMLRow As Long
Dim TotalDiffRow As Long
Dim RecTDS1, rowcount, cnt As Long
Dim Nodelist As Object
Dim Node, init As Object
Dim strdate As String
Dim YYYY, MM, DD As String
Dim TotalExRow, GSTColNo, GSTAmtColNo As Long
Dim AssYr As Variant


Set init = jsonObject("lastFiledITR")
If init.exists("turnoverGrsRcptForGSTIN") Then

Set Nodelist = jsonObject("lastFiledITR")("turnoverGrsRcptForGSTIN")
    
    GSTColNo = Sheet51.Range("GST_GSTRno").Column
    GSTAmtColNo = Sheet51.Range("GST_GSTRAmount").Column
    
    TotalExRow = Sheet51.Range("GST_GSTRno").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet51.Range("GST_GSTRno").Locked = False Then
            Sheet51.Range("GST_GSTRno").ClearContents
        End If
        If Sheet51.Range("GST_GSTRAmount").Locked = False Then
            Sheet51.Range("GST_GSTRAmount").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_GST (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet51.Range("GST_GSTRno").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
            If Sheet51.Range("F" & rowcount).Locked = False Then
                Sheet51.Range("F" & rowcount).value = Node("gstinNo")
            End If
            If Sheet51.Range("G" & rowcount).Locked = False Then
                Sheet51.Range("G" & rowcount).value = Node("amtTurnGrossRcptGSTIN")
            End If
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If
End Function

Function ScheduleTPSAImport(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

Dim RecTDS1, TotalExRow, TotalDiffRow, TotalXMLRow, RecTPSA As Long
Dim rowcount, cnt As Long
Dim Nodelist As Object
Dim Node, init As Object
Dim strdate As String
Dim YYYY, MM, DD As String
Dim BSRCodeColNo, BankBranchNameColNo, DateDepColNo, SrlNoOfChalnColNo, AmountColNo, FinYrColNo, TPSAAmountColNo As Long
Set init = jsonObject("form26as")("ScheduleTPSA")
If init.exists("DtlsTaxesPaid") Then
 
Set Nodelist = jsonObject("form26as")("ScheduleTPSA")("DtlsTaxesPaid")
    
    
    If Nodelist.count > 0 Then
        Sheet5.Range("sheet7.Section92CE_Flag").value = "Yes"
    End If
        
   BSRCodeColNo = Sheet57.Range("TPSC_BSRCode").Column
   BankBranchNameColNo = Sheet57.Range("TPSC_BankName").Column
   DateDepColNo = Sheet57.Range("TPSC_DateDep").Column
   SrlNoOfChalnColNo = Sheet57.Range("TPSC_SrlNoOfChaln").Column
   AmountColNo = Sheet57.Range("TPSC_Amt").Column
      
   TotalExRow = Range("TPSC_BSRCode").Rows.count
    
   TotalXMLRow = Nodelist.count
   TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet57.Range("TPSC_BSRCode").Locked = False Then
            Sheet57.Range("TPSC_BSRCode").ClearContents
        End If
        If Sheet57.Range("TPSC_BankName").Locked = False Then
            Sheet57.Range("TPSC_BankName").ClearContents
        End If
        If Sheet57.Range("TPSC_DateDep").Locked = False Then
            Sheet57.Range("TPSC_DateDep").ClearContents
        End If
        If Sheet57.Range("TPSC_SrlNoOfChaln").Locked = False Then
            Sheet57.Range("TPSC_SrlNoOfChaln").ClearContents
        End If
        If Sheet57.Range("TPSC_Amt").Locked = False Then
            Sheet57.Range("TPSC_Amt").ClearContents
        End If
      
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Item5 (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet57.Range("TPSC_BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
            If Sheet57.Cells(rowcount, BSRCodeColNo).Locked = False Then
                Sheet57.Cells(rowcount, BSRCodeColNo).value = Node("BSRCode")
            End If
            If Sheet57.Cells(rowcount, BankBranchNameColNo).Locked = False Then
                Sheet57.Cells(rowcount, BankBranchNameColNo).value = Node("BankBranchName")
            End If
            If Sheet57.Cells(rowcount, DateDepColNo).Locked = False Then
                YYYY = Mid(Node("DateDep"), 1, 4)
                MM = Mid(Node("DateDep"), 6, 2)
                DD = Mid(Node("DateDep"), 9, 2)
                strdate = DD & "/" & MM & "/" & YYYY
                Sheet57.Cells(rowcount, DateDepColNo).value = strdate
            End If
            If Sheet57.Cells(rowcount, SrlNoOfChalnColNo).Locked = False Then
                Sheet57.Cells(rowcount, SrlNoOfChalnColNo).value = Node("SrlNoOfChaln")
            End If
            If Sheet57.Cells(rowcount, AmountColNo).Locked = False Then
                Sheet57.Cells(rowcount, AmountColNo).value = Node("Amount")
            End If
    cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
 End If
     End Function

Function ScheduleICDSImport(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)


    If Sheet42.Range("ICDS.AccPolicies").Locked = False Then
        Sheet42.Range("ICDS.AccPolicies").value = jsonObject("form3CD")("ScheduleICDS")("AccPolicyAmt")
    End If
    If Sheet42.Range("ICDS.ValuationInv").Locked = False Then
        Sheet42.Range("ICDS.ValuationInv").value = jsonObject("form3CD")("ScheduleICDS")("InventoriesValue")
    End If
    If Sheet42.Range("ICDS.ConsContracts").Locked = False Then
        Sheet42.Range("ICDS.ConsContracts").value = jsonObject("form3CD")("ScheduleICDS")("ConstContractsAmt")
    End If
    If Sheet42.Range("ICDS.RevenueRecog").Locked = False Then
        Sheet42.Range("ICDS.RevenueRecog").value = jsonObject("form3CD")("ScheduleICDS")("RevenueRcgAmt")
    End If
    If Sheet42.Range("ICDS.TangibleFixAssests").Locked = False Then
'Change-1, ANK-126, 26.05.2022
        Sheet42.Range("ICDS.TangibleFixAssests").value = jsonObject("form3CD")("ScheduleICDS")("TangiableFixedAsset")
'        Sheet42.Range("ICDS.TangibleFixAssests").value = jsonObject("form3CD")("ScheduleICDS")("TangibleFixedAsset")
'---
    End If
    If Sheet42.Range("ICDS.ChngRates").Locked = False Then
        Sheet42.Range("ICDS.ChngRates").value = jsonObject("form3CD")("ScheduleICDS")("ForeignExgRates")
    End If
    If Sheet42.Range("ICDS.Govgrants").Locked = False Then
        Sheet42.Range("ICDS.Govgrants").value = jsonObject("form3CD")("ScheduleICDS")("GovtGrants")
    End If
    If Sheet42.Range("ICDS.Securities").Locked = False Then
        Sheet42.Range("ICDS.Securities").value = jsonObject("form3CD")("ScheduleICDS")("Securities")
    End If
    If Sheet42.Range("ICDS.BorrowingCosts").Locked = False Then
        Sheet42.Range("ICDS.BorrowingCosts").value = jsonObject("form3CD")("ScheduleICDS")("BorrowingCosts")
    End If
    If Sheet42.Range("ICDS.ProvLiability").Locked = False Then
        Sheet42.Range("ICDS.ProvLiability").value = jsonObject("form3CD")("ScheduleICDS")("ProvAssets")
    End If



End Function

Function ImportITR3ScheduleUD(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

Dim jsonDictionary, init As Object
Dim Node, Nodelist As Object
Dim CurrAssYr, CurBalCFNY, CurAllowBalCFNY, AssessmentYearColNo, BFColNo, SetoffColNo, BalanceColNo, Ud115BACColNo As Variant
Dim UD2BFColNo, UD2SetoffColNo, UD2BalanceColNo, TotBFUDepritAmt, TotCurYrdepritSetoffInc, TotBFUAllowAmt As Variant
Dim TotCurYrAllowSetoffInc, TotalBalCFNY As Variant
Dim TotalXMLRow, RecSchdulUD, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long




Set init = jsonObject("lastFiledITR")
If init.exists("scheduleUD") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleUD")


        AssessmentYearColNo = Sheet15.Range("UD.AssessmentYear").Column
        BFColNo = Sheet15.Range("UD.BF").Column
        UD2BFColNo = Sheet15.Range("UD2.BF").Column
        
        
    TotalExRow = Range("UD.AssessmentYear").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet15.Range("UD.AssessmentYear").Locked = False Then
            Sheet15.Range("UD.AssessmentYear").ClearContents
        End If
        If Sheet15.Range("UD.BF").Locked = False Then
            Sheet15.Range("UD.BF").ClearContents
        End If
        If Sheet15.Range("UD2.BF").Locked = False Then
            Sheet15.Range("UD2.BF").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_SchdlUD (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Range("UD.AssessmentYear").name)
'    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        
        If Sheet15.Cells(rowcount, AssessmentYearColNo).Locked = False Then
            Sheet15.Cells(rowcount, AssessmentYearColNo).value = Node("AssYr")
        End If
        If Sheet15.Cells(rowcount, BFColNo).Locked = False Then
            Sheet15.Cells(rowcount, BFColNo).value = Node("AmtBFUD")
        End If
        
        If Sheet15.Cells(rowcount, UD2BFColNo).Locked = False Then
            Sheet15.Cells(rowcount, UD2BFColNo).value = Node("AmtBFUAllow")
        End If
        
        rowcount = rowcount + 1
    cnt = cnt + 1
    Next Node
    RecSchdulUD = cnt
End If
End Function

Function ImportScheduleCGFor23(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
    Dim XpathOfSPI As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    
    Dim Name_of_district, SchEI_Pincode, SchEI_AgriculturalLand, SchEI_OwedOrLease, SchEI_IrrigatedOrRainFed As Variant
    Dim AgriLandOwnedFlag, NatureDesc, NatureDesc1 As Variant
    
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow, rowcount, cnt As Long
    Dim jsonDictionary, init As Object
    Dim Node, Nodelist, Node1, nodeList1 As Object
    Dim num, StCount, i As Variant
    num = 0
    
    Set init = jsonObject("form26AS")("ShortTermCapGainFor23")
    If init.exists("SaleofLandBuildDtls") Then
    Set Nodelist = jsonObject("form26AS")("ShortTermCapGainFor23")("SaleofLandBuildDtls")
    
    StCount = Range("PropertySectionCount_STCG").value
    
    TotalXMLRow = Nodelist.count
    
    TotalDiffRow = TotalXMLRow - StCount
    Sheet13.Activate
    For i = 1 To TotalDiffRow
            AddBlockCall_STCGrptfrm
    Next
   
    For Each Node In Nodelist
    
        num = num + 1

            If Sheet13.Range("STCG.Aai_FullConsideration" & num).Locked = False Then
                Sheet13.Range("STCG.Aai_FullConsideration" & num).value = Node("FullConsideration")
            End If
            If Sheet13.Range("STCG.Aaii_PropertyValuation" & num).Locked = False Then
                Sheet13.Range("STCG.Aaii_PropertyValuation" & num).value = Node("PropertyValuation")
            End If
        
        
        Next Node1
        End If
        Set init = Node("TrnsfImmblPrprty")
        If init.exists("TrnsfImmblPrprtyDtls") Then
        Dim CG_A_NameOfBuyer, CG_A_PanOfBuyer, CG_A_AadhaarOfBuyer, CG_A_PercentageShare, CG_A_Amount, CG_A_AddressOfBuyer, CG_A_PinCodeOfBuyer, CG_A_StateCodeOfBuyer, CG_A_CountryCodeOfBuyer, CG_A_ZipCodeOfBuyer
        Set nodeList1 = Node("TrnsfImmblPrprty")("TrnsfImmblPrprtyDtls")
        
        CG_A_NameOfBuyer = Sheet13.Range("CG_A_NameOfBuyer" & num).Column
        CG_A_PanOfBuyer = Sheet13.Range("CG_A_PanOfBuyer" & num).Column
        CG_A_AadhaarOfBuyer = Sheet13.Range("CG_A_AadhaarOfBuyer" & num).Column
        CG_A_PercentageShare = Sheet13.Range("CG_A_PercentageShare" & num).Column
        CG_A_Amount = Sheet13.Range("CG_A_Amount" & num).Column
        CG_A_AddressOfBuyer = Sheet13.Range("CG_A_AddressOfBuyer" & num).Column
        CG_A_StateCodeOfBuyer = Sheet13.Range("CG_A_StateCodeOfBuyer" & num).Column
        CG_A_PinCodeOfBuyer = Sheet13.Range("CG_A_PinCodeOfBuyer" & num).Column
        CG_A_CountryCodeOfBuyer = Sheet13.Range("CG_A_CountryOfBuyer" & num).Column
        CG_A_ZipCodeOfBuyer = Sheet13.Range("CG_A_ZipCodeOfBuyer" & num).Column
        
        TotalExRow = Range("CG_A_NameOfBuyer" & num).Rows.count
        
        TotalXMLRow = nodeList1.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet13.Range("CG_A_NameOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_NameOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_PanOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_PanOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_AadhaarOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_AadhaarOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_PercentageShare" & num).Locked = False Then
                Sheet13.Range("CG_A_PercentageShare" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_Amount" & num).Locked = False Then
                Sheet13.Range("CG_A_Amount" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_AddressOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_AddressOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_StateCodeOfBuyer" & num).Locked = False Then
                Dim cell
                For Each cell In Sheet13.Range("CG_A_StateCodeOfBuyer" & num).Cells
                    cell.value = "(Select)"
                Next
            End If
            If Sheet13.Range("CG_A_PinCodeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_PinCodeOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_CountryCodeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_CountryCodeOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_A_ZipCodeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_A_ZipCodeOfBuyer" & num).ClearContents
            End If
        End If
        
       
        
        If (TotalDiffRow > 0) Then
            AddSTCGA1 (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet13.Range("CG_A_NameOfBuyer" & num).name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node1 In nodeList1
            rowcount = rowcount + 1
            
            If Sheet13.Cells(rowcount, CG_A_NameOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_NameOfBuyer).value = Node1("NameOfBuyer")
            End If
            If Sheet13.Cells(rowcount, CG_A_PanOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_PanOfBuyer).value = Node1("PANofBuyer")
            End If
            If Sheet13.Cells(rowcount, CG_A_AadhaarOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_AadhaarOfBuyer).value = Node1("AaadhaarOfBuyer")
            End If
            If Sheet13.Cells(rowcount, CG_A_PercentageShare).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_PercentageShare).value = Node1("PercentageShare")
            End If
            If Sheet13.Cells(rowcount, CG_A_Amount).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_Amount).value = Node1("Amount")
            End If
            If Sheet13.Cells(rowcount, CG_A_AddressOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_AddressOfBuyer).value = Node1("AddressOfProperty")
            End If
            
            If Sheet13.Cells(rowcount, CG_A_StateCodeOfBuyer).Locked = False Then
                Dim iState, sState
                iState = Node1("StateCode")
                sState = Findtext(CStr(iState), "State")
                Sheet13.Cells(rowcount, CG_A_StateCodeOfBuyer).value = sState
            End If
            
                
            If Sheet13.Cells(rowcount, CG_A_PinCodeOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_PinCodeOfBuyer).value = Node1("PinCode")
            End If
            
            If Sheet13.Cells(rowcount, CG_A_CountryCodeOfBuyer).Locked = False Then
                Dim iCountry, sCountry
                iCountry = Node1("CountryCode")
                sCountry = Findtext(CStr(iCountry), "Country")
                Sheet13.Cells(rowcount, CG_A_CountryCodeOfBuyer).value = sCountry
            End If
            If Sheet13.Cells(rowcount, CG_A_ZipCodeOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_A_ZipCodeOfBuyer).value = Node1("ZipCode")
            End If
            
            cnt = cnt + 1
            
        Next Node1
        End If
    Next Node
   End If
    
'    If Sheet13.Range("STCG.A20ia_FullConsideration").Locked = False Then
'        Sheet13.Range("STCG.A20ia_FullConsideration").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SlumpSaleInStcg")("FullConsideration")
'    End If
'    If Sheet13.Range("STCG.A20ib_NetWorthOfDivision").Locked = False Then
'        Sheet13.Range("STCG.A20ib_NetWorthOfDivision").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SlumpSaleInStcg")("NetWorthOfDivision")
'    End If
    Set init = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")
    If init.exists("EquityMFonSTT") Then
    Set Nodelist = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("EquityMFonSTT")
    
     For Each Node In Nodelist
     
        If Node("MFSectionCode") = "1A" Then
        
            If Sheet13.Range("STCG.A2ia_FullConsideration").Locked = False Then
                Sheet13.Range("STCG.A2ia_FullConsideration").value = Node("EquityMFonSTTDtls")("FullConsideration")
            End If
            
            If Sheet13.Range("STCG.A2ibi_AquisitCost").Locked = False Then
                Sheet13.Range("STCG.A2ibi_AquisitCost").value = Node("EquityMFonSTTDtls")("DeductSec48")("AquisitCost")
            End If
            If Sheet13.Range("STCG.A2ibii_ImproveCost").Locked = False Then
                Sheet13.Range("STCG.A2ibii_ImproveCost").value = Node("EquityMFonSTTDtls")("DeductSec48")("ImproveCost")
            End If
            If Sheet13.Range("STCG.A2ibiii_ExpOnTrans").Locked = False Then
                Sheet13.Range("STCG.A2ibiii_ExpOnTrans").value = Node("EquityMFonSTTDtls")("DeductSec48")("ExpOnTrans")
            End If
            If Sheet13.Range("STCG.A2id_LossSec94of7Or94of8").Locked = False Then
                Sheet13.Range("STCG.A2id_LossSec94of7Or94of8").value = Node("EquityMFonSTTDtls")("LossSec94of7Or94of8")
            End If
                
        End If
        
        If Node("MFSectionCode") = "5AD1biip" Then
        
            If Sheet13.Range("STCG.A2iia_FullConsideration").Locked = False Then
                Sheet13.Range("STCG.A2iia_FullConsideration").value = Node("EquityMFonSTTDtls")("FullConsideration")
            End If
'
            If Sheet13.Range("STCG.A2iibi_AquisitCost").Locked = False Then
                Sheet13.Range("STCG.A2iibi_AquisitCost").value = Node("EquityMFonSTTDtls")("DeductSec48")("AquisitCost")
            End If
            If Sheet13.Range("STCG.A2iibii_ImproveCost").Locked = False Then
                Sheet13.Range("STCG.A2iibii_ImproveCost").value = Node("EquityMFonSTTDtls")("DeductSec48")("ImproveCost")
            End If
            If Sheet13.Range("STCG.A2iibiii_ExpOnTrans").Locked = False Then
                Sheet13.Range("STCG.A2iibiii_ExpOnTrans").value = Node("EquityMFonSTTDtls")("DeductSec48")("ExpOnTrans")
            End If
            If Sheet13.Range("STCG.A2iid_LossSec94of7Or94of8").Locked = False Then
                Sheet13.Range("STCG.A2iid_LossSec94of7Or94of8").value = Node("EquityMFonSTTDtls")("LossSec94of7Or94of8")
            End If
        
        End If
  
    Next Node
    End If
        Set Nodelist = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRITransacSec48Dtl")
    
       'Malli CG_51_AY_2025_26 CHANGE this field is autopopulated
'        If Sheet13.Range("STCG.A3a_NRItaxSTTPaid").Locked = False Then
'            Sheet13.Range("STCG.A3a_NRItaxSTTPaid").value = Nodelist("NRItaxSTTPaid")
'        End If
        '-----------------------
        If Sheet13.Range("STCG.A3b_NRItaxSTTNotPaid").Locked = False Then
            Sheet13.Range("STCG.A3b_NRItaxSTTNotPaid").value = Nodelist("NRItaxSTTNotPaid")
        End If
        
    '    "NRISecur115AD" start
            If Sheet13.Range("STCG.A4aia_FullConsideration").Locked = False Then
                Sheet13.Range("STCG.A4aia_FullConsideration").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("FullValueConsdRecvUnqshr")
            End If
            If Sheet13.Range("STCG.A4aib_FullMarketvalue").Locked = False Then
                Sheet13.Range("STCG.A4aib_FullMarketvalue").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("FairMrktValueUnqshr")
            End If
            If Sheet13.Range("STCG.A4aic_FullConsideration50CA").Locked = False Then
                Sheet13.Range("STCG.A4aic_FullConsideration50CA").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("FullValueConsdSec50CA")
            End If
            If Sheet13.Range("STCG.A4aii_FullConsiderationS").Locked = False Then
                Sheet13.Range("STCG.A4aii_FullConsiderationS").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("FullValueConsdOthUnqshr")
            End If
            If Sheet13.Range("STCG.A4a_FullConsideration").Locked = False Then
                Sheet13.Range("STCG.A4a_FullConsideration").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("FullConsideration")
            End If
            If Sheet13.Range("STCG.A4bi_AquisitCost").Locked = False Then
                Sheet13.Range("STCG.A4bi_AquisitCost").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("DeductSec48")("AquisitCost")
            End If
            If Sheet13.Range("STCG.A4bii_ImproveCost").Locked = False Then
                Sheet13.Range("STCG.A4bii_ImproveCost").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("DeductSec48")("ImproveCost")
            End If
            If Sheet13.Range("STCG.A4biii_ExpOnTrans").Locked = False Then
                Sheet13.Range("STCG.A4biii_ExpOnTrans").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("DeductSec48")("ExpOnTrans")
            End If
            If Sheet13.Range("STCG.A4biv_TotalDedn").Locked = False Then
                Sheet13.Range("STCG.A4biv_TotalDedn").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("DeductSec48")("TotalDedn")
            End If
            If Sheet13.Range("STCG.A4c_BalanceCG").Locked = False Then
                Sheet13.Range("STCG.A4c_BalanceCG").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("BalanceCG")
            End If
            If Sheet13.Range("STCG.A4d_LossSec94of7Or94of8").Locked = False Then
                Sheet13.Range("STCG.A4d_LossSec94of7Or94of8").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("LossSec94of7Or94of8")
            End If
            If Sheet13.Range("STCG.A4e_CapgainonAssets").Locked = False Then
                Sheet13.Range("STCG.A4e_CapgainonAssets").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRISecur115AD")("CapgainonAssets")
            End If
            If Sheet13.Range("STCG.A5aia_FullConsideration").Locked = False Then
                Sheet13.Range("STCG.A5aia_FullConsideration").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("FullValueConsdRecvUnqshr")
            End If
            If Sheet13.Range("STCG.A5aib_Fullmarketvalue").Locked = False Then
                Sheet13.Range("STCG.A5aib_Fullmarketvalue").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("FairMrktValueUnqshr")
            End If
            If Sheet13.Range("STCG.A5aii_FullConsideration").Locked = False Then
                Sheet13.Range("STCG.A5aii_FullConsideration").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("FullValueConsdOthUnqshr")
            End If
            If Sheet13.Range("STCG.A5bi_AquisitCost").Locked = False Then
                Sheet13.Range("STCG.A5bi_AquisitCost").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("DeductSec48")("AquisitCost")
            End If
            If Sheet13.Range("STCG.A5bii_ImproveCost").Locked = False Then
                Sheet13.Range("STCG.A5bii_ImproveCost").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("DeductSec48")("ImproveCost")
            End If
            If Sheet13.Range("STCG.A5biii_ExpOnTrans").Locked = False Then
                Sheet13.Range("STCG.A5biii_ExpOnTrans").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("DeductSec48")("ExpOnTrans")
            End If
            If Sheet13.Range("STCG.A5d_LossSec94of7Or94of8").Locked = False Then
                Sheet13.Range("STCG.A5d_LossSec94of7Or94of8").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("LossSec94of7Or94of8")
            End If
        ExemptionSecCode = ""
        
        Set init = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("ExemptionOrDednUs54")
        If init.exists("ExemptionOrDednUs54Dtls") Then
        
        Set nodeList1 = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("SaleOnOtherAssets")("ExemptionOrDednUs54")("ExemptionOrDednUs54Dtls")
        
        For Each Node1 In nodeList1
        
            ExemptionSecCode = Node1("ExemptionSecCode")
            
            If ExemptionSecCode = "54D" And Sheet13.Range("STCG.A5ei_ExemptionAmount").Locked = False Then
               Sheet13.Range("STCG.A5ei_ExemptionAmount").value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54G" And Sheet13.Range("STCG.A5eii_ExemptionAmount").Locked = False Then
               Sheet13.Range("STCG.A5eii_ExemptionAmount").value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54GA" And Sheet13.Range("STCG.A5eiii_ExemptionAmount").Locked = False Then
               Sheet13.Range("STCG.A5eiii_ExemptionAmount").value = Node1("ExemptionAmount")
            End If
        
        Next Node1
        End If
   
    Dim UnutilizedStcgFlag
    UnutilizedStcgFlag = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("UnutilizedStcgFlag")
    
    If UnutilizedStcgFlag = "Y" Then
        Sheet13.Range("STCG.A6_StatusUnutilizedCG").value = "Yes"
    ElseIf UnutilizedStcgFlag = "N" Then
        Sheet13.Range("STCG.A6_StatusUnutilizedCG").value = "No"
    Else
        Sheet13.Range("STCG.A6_StatusUnutilizedCG").value = "Not Applicable"
    End If
    
    If UnutilizedStcgFlag = "Y" Then
    
        Dim A7_PreviousYer_1, A7_DeductionSection_1, A7_AsstAcquiredYer_1, A7_AmtUtilised_1, A7_AmtNotUsed_1 As Variant
        
        Set init = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("UnutilizedCg")
        If init.exists("UnutilizedCgPrvYrDtls") Then
        
        Set nodeList1 = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("UnutilizedCg")("UnutilizedCgPrvYrDtls")
        
        A7_PreviousYer_1 = Sheet13.Range("STCG.A7_PreviousYer_1").Column
        A7_DeductionSection_1 = Sheet13.Range("STCG.A7_DeductionSection_1").Column
        A7_AsstAcquiredYer_1 = Sheet13.Range("STCG.A7_AsstAcquiredYer_1").Column
        A7_AmtUtilised_1 = Sheet13.Range("STCG.A7_AmtUtilised_1").Column
        A7_AmtNotUsed_1 = Sheet13.Range("STCG.A7_AmtNotUsed_1").Column
        
        
        TotalExRow = Sheet13.Range("STCG.A7_PreviousYer_1").Rows.count
        
        TotalXMLRow = nodeList1.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet13.Range("STCG.A7_PreviousYer_1").Locked = False Then
                Sheet13.Range("STCG.A7_PreviousYer_1").ClearContents
            End If
            
            If Sheet13.Range("STCG.A7_DeductionSection_1").Locked = False Then
                Sheet13.Range("STCG.A7_DeductionSection_1").ClearContents
            End If
            If Sheet13.Range("STCG.A7_AsstAcquiredYer_1").Locked = False Then
                Sheet13.Range("STCG.A7_AsstAcquiredYer_1").ClearContents
            End If
            If Sheet13.Range("STCG.A7_AmtUtilised_1").Locked = False Then
                Sheet13.Range("STCG.A7_AmtUtilised_1").ClearContents
            End If
            If Sheet13.Range("STCG.A7_AmtNotUsed_1").Locked = False Then
                Sheet13.Range("STCG.A7_AmtNotUsed_1").ClearContents
            End If
        End If
        If (TotalDiffRow > 0) Then
            AddRowsSchCG_STCG (TotalDiffRow)
        End If
        
        rowcount = Sheet13.Range("STCG.A7_PreviousYer_1").row
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node1 In nodeList1
            rowcount = rowcount + 1
            
            If Sheet13.Cells(rowcount, A7_PreviousYer_1).Locked = False Then
                Sheet13.Cells(rowcount, A7_PreviousYer_1).value = Node1("PrvYrInWhichAsstTrnsfrd")
            End If
            
            If Sheet13.Cells(rowcount, A7_DeductionSection_1).Locked = False Then
                Sheet13.Cells(rowcount, A7_DeductionSection_1).value = Node1("SectionClmd")
            End If
            If Sheet13.Cells(rowcount, A7_AsstAcquiredYer_1).Locked = False Then
                If Node1("YrInWhichAssetAcq") = "2017" Then
                    Sheet13.Cells(rowcount, A7_AsstAcquiredYer_1).value = "2017-18"
                ElseIf Node1("YrInWhichAssetAcq") = "2018" Then
                    Sheet13.Cells(rowcount, A7_AsstAcquiredYer_1).value = "2018-19"
                ElseIf Node1("YrInWhichAssetAcq") = "2019" Then
                    Sheet13.Cells(rowcount, A7_AsstAcquiredYer_1).value = "2019-20"
                ElseIf Node1("YrInWhichAssetAcq") = "2020" Then
                    Sheet13.Cells(rowcount, A7_AsstAcquiredYer_1).value = "2020-21"
                End If
            End If
            If Sheet13.Cells(rowcount, A7_AmtUtilised_1).Locked = False Then
                Sheet13.Cells(rowcount, A7_AmtUtilised_1).value = Node1("AmtUtilized")
            End If
            If Sheet13.Cells(rowcount, A7_AmtNotUsed_1).Locked = False Then
                Sheet13.Cells(rowcount, A7_AmtNotUsed_1).value = Node1("AmtUnutilized")
            End If
            
            cnt = cnt + 1
        Next Node1
        Sheet13.Protect Password:=getmsgstate
        End If
    End If
    
        If Sheet13.Range("STCG.A6_AmountDeemedOth").Locked = False Then
            Sheet13.Range("STCG.A6_AmountDeemedOth").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("AmtDeemedStcg")
        End If
        If Sheet13.Range("STCG.A8_PassThroughIncome15").Locked = False Then
            Sheet13.Range("STCG.A8_PassThroughIncome15").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("PassThrIncNatureSTCG15Per")
        End If
        If Sheet13.Range("STCG.A8_PassThroughIncome30").Locked = False Then
            Sheet13.Range("STCG.A8_PassThroughIncome30").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("PassThrIncNatureSTCG30Per")
        End If
        If Sheet13.Range("STCG.A8_PassThroughIncome_ApplicableRate").Locked = False Then
            Sheet13.Range("STCG.A8_PassThroughIncome_ApplicableRate").value = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("PassThrIncNatureSTCGAppRate")
        End If
    
    '"NRIDTAADtls" start
    Dim A8_AmtI, A8_IA1tA6, A8_Country, A8_DTAA_Article, A8_RatePTreaty, A8_CertiStatus, A8_SecIT, A8_RateIT, A8_ApplRate As Variant
        
    Set init = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")
    If init.exists("NRICgDTAA") Then
    Set init = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRICgDTAA")
    If init.exists("NRIDTAADtls") Then

    Set nodeList1 = jsonObject("ScheduleCGFor23")("ShortTermCapGainFor23")("NRICgDTAA")("NRIDTAADtls")
    
    A8_AmtI = Sheet13.Range("STCG.A8_AmtI").Column
    A8_IA1tA6 = Sheet13.Range("STCG.A8_IA1tA6").Column
    A8_Country = Sheet13.Range("STCG.A8_Country").Column
    A8_DTAA_Article = Sheet13.Range("STCG.A8_DTAA_Article").Column
    A8_RatePTreaty = Sheet13.Range("STCG.A8_RatePTreaty").Column
    A8_CertiStatus = Sheet13.Range("STCG.A8_CertiStatus").Column
    A8_SecIT = Sheet13.Range("STCG.A8_SecIT").Column
    A8_RateIT = Sheet13.Range("STCG.A8_RateIT").Column
    
    TotalExRow = Range("STCG.A8_AmtI").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet13.Range("STCG.A8_AmtI").Locked = False Then
            Sheet13.Range("STCG.A8_AmtI").ClearContents
        End If
        If Sheet13.Range("STCG.A8_IA1tA6").Locked = False Then
            Sheet13.Range("STCG.A8_IA1tA6").ClearContents
        End If
        If Sheet13.Range("STCG.A8_Country").Locked = False Then
            Sheet13.Range("STCG.A8_Country").ClearContents
        End If
        If Sheet13.Range("STCG.A8_DTAA_Article").Locked = False Then
            Sheet13.Range("STCG.A8_DTAA_Article").ClearContents
        End If
        If Sheet13.Range("STCG.A8_RatePTreaty").Locked = False Then
            Sheet13.Range("STCG.A8_RatePTreaty").ClearContents
        End If
        If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" And Sheet13.Range("STCG.A8_CertiStatus").Locked = False Then
            Sheet13.Range("STCG.A8_CertiStatus").ClearContents
        End If
        If Sheet13.Range("STCG.A8_SecIT").Locked = False Then
            Sheet13.Range("STCG.A8_SecIT").ClearContents
        End If
        If Sheet13.Range("STCG.A8_RateIT").Locked = False Then
            Sheet13.Range("STCG.A8_RateIT").ClearContents
        End If
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_NRIDTAADtls (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("STCG.A8_AmtI").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        If Sheet13.Cells(rowcount, A8_AmtI).Locked = False Then
            Sheet13.Cells(rowcount, A8_AmtI).value = Node1("DTAAamt")
        End If
        If Sheet13.Cells(rowcount, A8_IA1tA6).Locked = False Then
            Sheet13.Cells(rowcount, A8_IA1tA6).value = Node1("ItemNoincl")
        End If
        
        Dim iCountry2, sCountry2, TaxRescertifiedFlag
        
        iCountry2 = Node1("CountryCode")
        sCountry2 = Findtext(iCountry2, "Country")
        
        If Sheet13.Cells(rowcount, A8_Country).Locked = False Then
            Sheet13.Cells(rowcount, A8_Country).value = sCountry2
        End If
        If Sheet13.Cells(rowcount, A8_DTAA_Article).Locked = False Then
            Sheet13.Cells(rowcount, A8_DTAA_Article).value = Node1("DTAAarticle")
        End If
        If Sheet13.Cells(rowcount, A8_RatePTreaty).Locked = False Then
            Sheet13.Cells(rowcount, A8_RatePTreaty).value = Node1("RateAsPerTreaty")
        End If
        
        TaxRescertifiedFlag = Node1("TaxRescertifiedFlag")
        
        If TaxRescertifiedFlag = "Y" Then
            TaxRescertifiedFlag = "Yes"
        ElseIf TaxRescertifiedFlag = "N" Then
            TaxRescertifiedFlag = "No"
        Else
            TaxRescertifiedFlag = "(Select)"
        End If
        
        If Sheet13.Cells(rowcount, A8_CertiStatus).Locked = False Then
            Sheet13.Cells(rowcount, A8_CertiStatus).value = TaxRescertifiedFlag
        End If
        
            If Sheet13.Cells(rowcount, A8_SecIT).Locked = False Then
                Sheet13.Cells(rowcount, A8_SecIT).value = Node1("SecITAct")
            End If
            If Sheet13.Cells(rowcount, A8_RateIT).Locked = False Then
                Sheet13.Cells(rowcount, A8_RateIT).value = Node1("RateAsPerITAct")
            End If
        
        cnt = cnt + 1
        
    Next Node1
    End If
    End If

    Dim LtCount As Variant
    
    Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofLandBuild")
    If init.exists("SaleofLandBuildDtls") Then
    
    Set Nodelist = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofLandBuild")("SaleofLandBuildDtls")
    num = 0
    
    LtCount = Range("PropertySectionCount_LTCG").value
    
    TotalXMLRow = Nodelist.count
    
    TotalDiffRow = TotalXMLRow - LtCount
    
    Sheet13.Activate
    For i = 1 To TotalDiffRow
            AddBlockCall_LTCGrptfrm
    Next
   
    
    For Each Node In Nodelist
    
        num = num + 1
        If Sheet13.Range("LTCG.Bai_FullConsideration" & num).Locked = False Then
            Sheet13.Range("LTCG.Bai_FullConsideration" & num).value = Node("FullConsideration")
        End If
        If Sheet13.Range("LTCG.Baii_PropertyValuation" & num).Locked = False Then
            Sheet13.Range("LTCG.Baii_PropertyValuation" & num).value = Node("PropertyValuation")
        End If
        If Sheet13.Range("LTCG.Bbi_AquisitCost" & num).Locked = False Then
            Sheet13.Range("LTCG.Bbi_AquisitCost" & num).value = Node("AquisitCost")
        End If
        If Sheet13.Range("LTCG.Bbii_ImproveCost" & num).Locked = False Then
            Sheet13.Range("LTCG.Bbii_ImproveCost" & num).value = Node("ImproveCost")
        End If
        If Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & num).Locked = False Then
            Sheet13.Range("LTCG.Bbiii_ExpOnTrans" & num).value = Node("ExpOnTrans")
        End If
        ExemptionSecCode = ""
        Set init = Node("ExemptionOrDednUs54")
        If init.exists("ExemptionOrDednUs54Dtls") Then

        Set nodeList1 = Node("ExemptionOrDednUs54")("ExemptionOrDednUs54Dtls")
        
        For Each Node1 In nodeList1
        
            ExemptionSecCode = Node1("ExemptionSecCode")
        
            If ExemptionSecCode = "54" And Sheet13.Range("LTCG.Bdi_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdi_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54B" And Sheet13.Range("LTCG.Bdii_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdii_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54D" And Sheet13.Range("LTCG.Bdiii_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdiii_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54EC" And Sheet13.Range("LTCG.Bdiv_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdiv_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54F" And Sheet13.Range("LTCG.Bdv_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdv_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54G" And Sheet13.Range("LTCG.Bdvi_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdvi_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54GA" And Sheet13.Range("LTCG.Bdvii_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdvii_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
            If ExemptionSecCode = "54GB" And Sheet13.Range("LTCG.Bdviii_ExemptionAmount" & num).Locked = False Then
               Sheet13.Range("LTCG.Bdviii_ExemptionAmount" & num).value = Node1("ExemptionAmount")
            End If
            
        Next Node1
        End If
        Set init = Node("TrnsfImmblPrprty")
        If init.exists("TrnsfImmblPrprtyDtls") Then
        Dim CG_B_NameOfBuyer, CG_B_PanOfBuyer, CG_B_AadhaaeOfBuyer, CG_B_PercentageShare, CG_B_Amount, CG_B_AddressOfBuyer, CG_B_PinCodeOfBuyer, CG_B_StateCodeOfBuyer, CG_B_CountryCodeOfBuyer, CG_B_ZipCodeOfBuyer
        Set nodeList1 = Node("TrnsfImmblPrprty")("TrnsfImmblPrprtyDtls")
        
        CG_B_NameOfBuyer = Sheet13.Range("CG_B_NameOfBuyer" & num).Column
        CG_B_PanOfBuyer = Sheet13.Range("CG_B_PanOfBuyer" & num).Column
        CG_B_AadhaaeOfBuyer = Sheet13.Range("CG_B_AadhaaeOfBuyer" & num).Column
        CG_B_PercentageShare = Sheet13.Range("CG_B_PercentageShare" & num).Column
        CG_B_Amount = Sheet13.Range("CG_B_Amount" & num).Column
        CG_B_AddressOfBuyer = Sheet13.Range("CG_B_AddressOfBuyer" & num).Column
        CG_B_PinCodeOfBuyer = Sheet13.Range("CG_B_PinCodeOfBuyer" & num).Column
        CG_B_StateCodeOfBuyer = Sheet13.Range("CG_B_StateCodeOfBuyer" & num).Column
        CG_B_CountryCodeOfBuyer = Sheet13.Range("CG_B_CountryOfBuyer" & num).Column
        CG_B_ZipCodeOfBuyer = Sheet13.Range("CG_B_ZipCodeOfBuyer" & num).Column
        
        TotalExRow = Range("CG_B_NameOfBuyer" & num).Rows.count
        
        TotalXMLRow = nodeList1.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet13.Range("CG_B_NameOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_NameOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_PanOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_PanOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_AadhaaeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_AadhaaeOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_PercentageShare" & num).Locked = False Then
                Sheet13.Range("CG_B_PercentageShare" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_Amount" & num).Locked = False Then
                Sheet13.Range("CG_B_Amount" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_AddressOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_AddressOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_PinCodeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_PinCodeOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_StateCodeOfBuyer" & num).Locked = False Then
                  For Each cell In Sheet13.Range("CG_B_StateCodeOfBuyer" & num).Cells
                    cell.value = "(Select)"
                Next
            End If
            If Sheet13.Range("CG_B_CountryCodeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_CountryCodeOfBuyer" & num).ClearContents
            End If
            If Sheet13.Range("CG_B_ZipCodeOfBuyer" & num).Locked = False Then
                Sheet13.Range("CG_B_ZipCodeOfBuyer" & num).ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
            AddLTCGB1 (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet13.Range("CG_B_NameOfBuyer" & num).name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node1 In nodeList1
            rowcount = rowcount + 1
            
            If Sheet13.Cells(rowcount, CG_B_NameOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_NameOfBuyer).value = Node1("NameOfBuyer")
            End If
            If Sheet13.Cells(rowcount, CG_B_PanOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_PanOfBuyer).value = Node1("PANofBuyer")
            End If
            If Sheet13.Cells(rowcount, CG_B_AadhaaeOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_AadhaaeOfBuyer).value = Node1("AaadhaarOfBuyer")
            End If
            If Sheet13.Cells(rowcount, CG_B_PercentageShare).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_PercentageShare).value = Node1("PercentageShare")
            End If
            If Sheet13.Cells(rowcount, CG_B_Amount).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_Amount).value = Node1("Amount")
            End If
            If Sheet13.Cells(rowcount, CG_B_AddressOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_AddressOfBuyer).value = Node1("AddressOfProperty")
            End If
          
            If Sheet13.Cells(rowcount, CG_B_StateCodeOfBuyer).Locked = False Then
                Dim iState1, sState1
                iState1 = Node1("StateCode")
                sState1 = Findtext(CStr(iState1), "State")
                Sheet13.Cells(rowcount, CG_B_StateCodeOfBuyer).value = sState1
            End If
            If Sheet13.Cells(rowcount, CG_B_CountryCodeOfBuyer).Locked = False Then
                Dim iCountry1, sCountry1
                iCountry1 = Node1("CountryCode")
                sCountry1 = Findtext(CStr(iCountry1), "Country")
                Sheet13.Cells(rowcount, CG_B_CountryCodeOfBuyer).value = sCountry1
            End If
              If Sheet13.Cells(rowcount, CG_B_PinCodeOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_PinCodeOfBuyer).value = Node1("PinCode")
            End If
            If Sheet13.Cells(rowcount, CG_B_ZipCodeOfBuyer).Locked = False Then
                Sheet13.Cells(rowcount, CG_B_ZipCodeOfBuyer).value = Node1("ZipCode")
            End If
            cnt = cnt + 1
            
        Next Node1
        End If
    Next Node
    End If
    Set Node = Nothing
    Set Node = CreateObject("Scripting.Dictionary")
   
    Set Node1 = Nothing
    Set Node1 = CreateObject("Scripting.Dictionary")
    
    Set Nodelist = Nothing
    Set Nodelist = CreateObject("Scripting.Dictionary")
   
    Set nodeList1 = Nothing
    Set nodeList1 = CreateObject("Scripting.Dictionary")
   
    
'    If Sheet13.Range("LTCG.B02a_FullConsideration").Locked = False Then
'        Sheet13.Range("LTCG.B02a_FullConsideration").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SlumpSaleInLtcg")("FullConsideration")
'    End If
'    If Sheet13.Range("LTCG.B02b_NetWorthOfDivision").Locked = False Then
'        Sheet13.Range("LTCG.B02b_NetWorthOfDivision").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SlumpSaleInLtcg")("NetWorthOfDivision")
'    End If
    
    Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SlumpSaleInLtcg")("ExemptionOrDednUs54")
    If init.exists("ExemptionOrDednUs54Dtls") Then
    Set Nodelist = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SlumpSaleInLtcg")("ExemptionOrDednUs54")("ExemptionOrDednUs54Dtls")
    
    For Each Node In Nodelist
    
        ExemptionSecCode = Node("ExemptionSecCode")
                
'        If ExemptionSecCode = "54EC" And Sheet13.Range("LTCG.B02di_ExemptionAmount").Locked = False Then
'           Sheet13.Range("LTCG.B02di_ExemptionAmount").value = Node("ExemptionAmount")
'        End If
        
'        If ExemptionSecCode = "54F" And Sheet13.Range("LTCG.B02dii_ExemptionAmount").Locked = False Then
'           Sheet13.Range("LTCG.B02dii_ExemptionAmount").value = Node("ExemptionAmount")
'        End If
        
    Next Node
    End If
    
        If Sheet13.Range("LTCG.B2a_FullConsideration").Locked = False Then
            Sheet13.Range("LTCG.B2a_FullConsideration").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofBondsDebntr")("FullConsideration")
        End If
        If Sheet13.Range("LTCG.B2bi_AquisitCost").Locked = False Then
            Sheet13.Range("LTCG.B2bi_AquisitCost").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofBondsDebntr")("DeductSec48")("AquisitCost")
        End If
        If Sheet13.Range("LTCG.B2bii_ImproveCost").Locked = False Then
            Sheet13.Range("LTCG.B2bii_ImproveCost").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofBondsDebntr")("DeductSec48")("ImproveCost")
        End If
        If Sheet13.Range("LTCG.B2biii_ExpOnTrans").Locked = False Then
            Sheet13.Range("LTCG.B2biii_ExpOnTrans").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofBondsDebntr")("DeductSec48")("ExpOnTrans")
        End If
        If Sheet13.Range("LTCG.B03d_ExemptionGrandTotal").Locked = False Then
            Sheet13.Range("LTCG.B03d_ExemptionGrandTotal").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofBondsDebntr")("DeductionUs54F")
        End If
    Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")
    If init.exists("Proviso112Applicable") Then
    
    Set Nodelist = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("Proviso112Applicable")
    Dim Proviso112SectionCode
    
    For Each Node In Nodelist
    
        Proviso112SectionCode = Node("Proviso112SectionCode")
        If Proviso112SectionCode = "22" Then
        
            If Sheet13.Range("LTCG.B3a_FullConsideration").Locked = False Then
                Sheet13.Range("LTCG.B3a_FullConsideration").value = Node("Proviso112Applicabledtls")("FullConsideration")
            End If
            If Sheet13.Range("LTCG.B3bi_AquisitCost").Locked = False Then
                Sheet13.Range("LTCG.B3bi_AquisitCost").value = Node("Proviso112Applicabledtls")("DeductSec48")("AquisitCost")
            End If
            If Sheet13.Range("LTCG.B3bii_ImproveCost").Locked = False Then
                Sheet13.Range("LTCG.B3bii_ImproveCost").value = Node("Proviso112Applicabledtls")("DeductSec48")("ImproveCost")
            End If
            If Sheet13.Range("LTCG.B3biii_ExpOnTrans").Locked = False Then
                Sheet13.Range("LTCG.B3biii_ExpOnTrans").value = Node("Proviso112Applicabledtls")("DeductSec48")("ExpOnTrans")
            End If
            If Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").Locked = False Then
                Sheet13.Range("LTCG.B04d_ExemptionGrandTotal").value = Node("Proviso112Applicabledtls")("DeductionUs54F")
            End If
        
        ElseIf Proviso112SectionCode = "5ACA1b" Then
        
            If Sheet13.Range("LTCG.Bii3a_FullConsideration").Locked = False Then
                Sheet13.Range("LTCG.Bii3a_FullConsideration").value = Node("Proviso112Applicabledtls")("FullConsideration")
            End If
            If Sheet13.Range("LTCG.Bii3bi_AquisitCost").Locked = False Then
                Sheet13.Range("LTCG.Bii3bi_AquisitCost").value = Node("Proviso112Applicabledtls")("DeductSec48")("AquisitCost")
            End If
            If Sheet13.Range("LTCG.Bii3bii_ImproveCost").Locked = False Then
                Sheet13.Range("LTCG.Bii3bii_ImproveCost").value = Node("Proviso112Applicabledtls")("DeductSec48")("ImproveCost")
            End If
            If Sheet13.Range("LTCG.Bii3biii_ExpOnTrans").Locked = False Then
                Sheet13.Range("LTCG.Bii3biii_ExpOnTrans").value = Node("Proviso112Applicabledtls")("DeductSec48")("ExpOnTrans")
            End If
            If Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").Locked = False Then
                Sheet13.Range("LTCG.B04di_ExemptionGrandTotal").value = Node("Proviso112Applicabledtls")("DeductionUs54F")
            End If
        
        End If
        
    Next Node
    End If
    If Sheet13.Range("B5_DeductionTotal").Locked = False Then
        Sheet13.Range("B5_DeductionTotal").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleOfEquityShareUs112A")("DeductionUs54F")
    End If
   
        If Sheet13.Range("LTCG.B4a_LTCGWithoutBenefit").Locked = False Then
            Sheet13.Range("LTCG.B4a_LTCGWithoutBenefit").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRIProvisoSec48")("LTCGWithoutBenefit")
        End If
        If Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").Locked = False Then
            Sheet13.Range("LTCG.B06d_ExemptionGrandTotal").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRIProvisoSec48")("DeductionUs54F")
        End If
        If Sheet13.Range("LTCG.B4e_LTCGOnUnlistedSecurity").Locked = False Then
            Sheet13.Range("LTCG.B4e_LTCGOnUnlistedSecurity").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRIProvisoSec48")("BalanceCG")
        End If
        
        Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRIOnSec112and115")
        If init.exists("NRIOnSec112and115Dtls") Then
        '"NRIOnSec112and115" start
        Set Nodelist = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRIOnSec112and115")("NRIOnSec112and115Dtls")
        Dim sectioncode
        
        For Each Node In Nodelist
        
            sectioncode = Node("SectionCode")
            
            If sectioncode = "21ciii" Then
                
                If Sheet13.Range("LTCG.B6aia_FullConsideration").Locked = False Then
                    Sheet13.Range("LTCG.B6aia_FullConsideration").value = Node("FullValueConsdRecvUnqshr")
                End If
                If Sheet13.Range("LTCG.B6aib_FullMarketValue").Locked = False Then
                    Sheet13.Range("LTCG.B6aib_FullMarketValue").value = Node("FairMrktValueUnqshr")
                End If
                If Sheet13.Range("LTCG.B6aic_FullConsideration50CA").Locked = False Then
                    Sheet13.Range("LTCG.B6aic_FullConsideration50CA").value = Node("FullValueConsdSec50CA")
                End If
                If Sheet13.Range("LTCG.B6aii_FullConsiderationS").Locked = False Then
                    Sheet13.Range("LTCG.B6aii_FullConsiderationS").value = Node("FullValueConsdOthUnqshr")
                End If
                If Sheet13.Range("LTCG.Bi5a_FullConsideration").Locked = False Then
                    Sheet13.Range("LTCG.Bi5a_FullConsideration").value = Node("FullConsideration")
                End If
                If Sheet13.Range("LTCG.Bi5bi_AquisitCost").Locked = False Then
                    Sheet13.Range("LTCG.Bi5bi_AquisitCost").value = Node("DeductSec48")("AquisitCost")
                End If
                If Sheet13.Range("LTCG.Bi5bii_ImproveCost").Locked = False Then
                    Sheet13.Range("LTCG.Bi5bii_ImproveCost").value = Node("DeductSec48")("ImproveCost")
                End If
                If Sheet13.Range("LTCG.Bi5biii_ExpOnTrans").Locked = False Then
                    Sheet13.Range("LTCG.Bi5biii_ExpOnTrans").value = Node("DeductSec48")("ExpOnTrans")
                End If
                If Sheet13.Range("LTCG.Bi5biv_TotalDedn").Locked = False Then
                    Sheet13.Range("LTCG.Bi5biv_TotalDedn").value = Node("DeductSec48")("TotalDedn")
                End If
                If Sheet13.Range("LTCG.Bi5c_BalanceCG").Locked = False Then
                    Sheet13.Range("LTCG.Bi5c_BalanceCG").value = Node("BalanceCG")
                End If
                If Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").Locked = False Then
                    Sheet13.Range("LTCG.B07dii_ExemptionGrandTotal").value = Node("DeductionUs54F")
                End If
                If Sheet13.Range("LTCG.Bi5e_CapgainonAssets").Locked = False Then
                    Sheet13.Range("LTCG.Bi5e_CapgainonAssets").value = Node("CapgainonAssets")
                End If
            
            ElseIf sectioncode = "5AC1c" Then
            
                If Sheet13.Range("LTCG.B6ibia_FullConsideration").Locked = False Then
                    Sheet13.Range("LTCG.B6ibia_FullConsideration").value = Node("FullValueConsdRecvUnqshr")
                End If
                If Sheet13.Range("LTCG.B6ibib_FullMarketValue").Locked = False Then
                    Sheet13.Range("LTCG.B6ibib_FullMarketValue").value = Node("FairMrktValueUnqshr")
                End If
                If Sheet13.Range("LTCG.B6ibii_FullConsideration50CA").Locked = False Then
                    Sheet13.Range("LTCG.B6ibii_FullConsideration50CA").value = Node("FullValueConsdSec50CA")
                End If
                If Sheet13.Range("LTCG.Bii5a_FullConsideration").Locked = False Then
                    Sheet13.Range("LTCG.Bii5a_FullConsideration").value = Node("FullConsideration")
                End If
                If Sheet13.Range("LTCG.Bii5bi_AquisitCost").Locked = False Then
                    Sheet13.Range("LTCG.Bii5bi_AquisitCost").value = Node("DeductSec48")("AquisitCost")
                End If
                If Sheet13.Range("LTCG.Bii5bii_ImproveCost").Locked = False Then
                    Sheet13.Range("LTCG.Bii5bii_ImproveCost").value = Node("DeductSec48")("ImproveCost")
                End If
                If Sheet13.Range("LTCG.Bii5biii_ExpOnTrans").Locked = False Then
                    Sheet13.Range("LTCG.Bii5biii_ExpOnTrans").value = Node("DeductSec48")("ExpOnTrans")
                End If
                If Sheet13.Range("LTCG.Bii5biv_TotalDedn").Locked = False Then
                    Sheet13.Range("LTCG.Bii5biv_TotalDedn").value = Node("DeductSec48")("TotalDedn")
                End If
                If Sheet13.Range("LTCG.Bii5c_BalanceCG").Locked = False Then
                    Sheet13.Range("LTCG.Bii5c_BalanceCG").value = Node("BalanceCG")
                End If
                If Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").Locked = False Then
                    Sheet13.Range("LTCG.B07div_ExemptionGrandTotal").value = Node("DeductionUs54F")
                End If
                If Sheet13.Range("LTCG.Bii5e_CapgainonAssets").Locked = False Then
                    Sheet13.Range("LTCG.Bii5e_CapgainonAssets").value = Node("CapgainonAssets")
                End If
                
            ElseIf sectioncode = "5ADiii" Then
            
                If Sheet13.Range("LTCG.B6iiiaia_FullConsideration").Locked = False Then
                    Sheet13.Range("LTCG.B6iiiaia_FullConsideration").value = Node("FullValueConsdRecvUnqshr")
                End If
                If Sheet13.Range("LTCG.B6iiiaib_FullMarketValue").Locked = False Then
                    Sheet13.Range("LTCG.B6iiiaib_FullMarketValue").value = Node("FairMrktValueUnqshr")
                End If
                If Sheet13.Range("LTCG.B6iiiaic_FullConsideration50CA").Locked = False Then
                    Sheet13.Range("LTCG.B6iiiaic_FullConsideration50CA").value = Node("FullValueConsdSec50CA")
                End If
                If Sheet13.Range("LTCG.B6iiiaii_FullConsiderationS").Locked = False Then
                    Sheet13.Range("LTCG.B6iiiaii_FullConsiderationS").value = Node("FullValueConsdOthUnqshr")
                End If
                If Sheet13.Range("LTCG.Biii5a_FullConsideration").Locked = False Then
                    Sheet13.Range("LTCG.Biii5a_FullConsideration").value = Node("FullConsideration")
                End If
                If Sheet13.Range("LTCG.Biii5bi_AquisitCost").Locked = False Then
                    Sheet13.Range("LTCG.Biii5bi_AquisitCost").value = Node("DeductSec48")("AquisitCost")
                End If
                If Sheet13.Range("LTCG.Biii5bii_ImproveCost").Locked = False Then
                    Sheet13.Range("LTCG.Biii5bii_ImproveCost").value = Node("DeductSec48")("ImproveCost")
                End If
                If Sheet13.Range("LTCG.Biii5biii_ExpOnTrans").Locked = False Then
                    Sheet13.Range("LTCG.Biii5biii_ExpOnTrans").value = Node("DeductSec48")("ExpOnTrans")
                End If
                If Sheet13.Range("LTCG.Biii5biv_TotalDedn").Locked = False Then
                    Sheet13.Range("LTCG.Biii5biv_TotalDedn").value = Node("DeductSec48")("TotalDedn")
                End If
                If Sheet13.Range("LTCG.Biii5c_BalanceCG").Locked = False Then
                    Sheet13.Range("LTCG.Biii5c_BalanceCG").value = Node("BalanceCG")
                End If
                If Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").Locked = False Then
                    Sheet13.Range("LTCG.B07dv_ExemptionGrandTotal").value = Node("DeductionUs54F")
                End If
                If Sheet13.Range("LTCG.Biii5e_CapgainonAssets").Locked = False Then
                    Sheet13.Range("LTCG.Biii5e_CapgainonAssets").value = Node("CapgainonAssets")
                End If
            
            End If
            
        Next Node
        
        If Sheet13.Range("LTCG.8.LTCG_AfterThreshold").Locked = False Then
            Sheet13.Range("LTCG.8.LTCG_AfterThreshold").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleOfEquityShareUs112A")("BalanceCG")
        End If
        If Sheet13.Range("LTCG.8.DeductionUnder54F").Locked = False Then
            Sheet13.Range("LTCG.8.DeductionUnder54F").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleOfEquityShareUs112A")("DeductionUs54F")
        End If
        If Sheet13.Range("LTCG.5.LTCG_AssetsB8Above").Locked = False Then
            Sheet13.Range("LTCG.5.LTCG_AssetsB8Above").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleOfEquityShareUs112A")("CapgainonAssets")
        End If
        
            If Sheet13.Range("LTCG.B6a_SaleonSpecAsset").Locked = False Then
                Sheet13.Range("LTCG.B6a_SaleonSpecAsset").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleofForeignAsset")("SaleonSpecAsset")
            End If
            If Sheet13.Range("LTCG.B6b_DednSpecAssetus115").Locked = False Then
                Sheet13.Range("LTCG.B6b_DednSpecAssetus115").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleofForeignAsset")("DednSpecAssetus115")
            End If
            If Sheet13.Range("LTCG.B6c_BalonSpeciAsset").Locked = False Then
                Sheet13.Range("LTCG.B6c_BalonSpeciAsset").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleofForeignAsset")("BalonSpeciAsset")
            End If
'Commented by Riyaz on 17/03/2025
'            If Sheet13.Range("LTCG.B6d_SaleOtherSpecAsset").Locked = False Then
'                Sheet13.Range("LTCG.B6d_SaleOtherSpecAsset").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleofForeignAsset")("SaleOtherSpecAsset")
'            End If
            If Sheet13.Range("LTCG.B6e_DednOtherSpecAssetus115").Locked = False Then
                Sheet13.Range("LTCG.B6e_DednOtherSpecAssetus115").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleofForeignAsset")("DednOtherSpecAssetus115")
            End If
            If Sheet13.Range("LTCG.B6f_BalOtherthanSpecAsset").Locked = False Then
                Sheet13.Range("LTCG.B6f_BalOtherthanSpecAsset").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRISaleofForeignAsset")("BalOtherthanSpecAsset")
            End If
    End If
    
    If Sheet13.Range("LTCG.B8aia_FullConsideration").Locked = False Then
        Sheet13.Range("LTCG.B8aia_FullConsideration").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("FullValueConsdRecvUnqshr")
    End If
    If Sheet13.Range("LTCG.B8aib_FullMarketValue").Locked = False Then
       Sheet13.Range("LTCG.B8aib_FullMarketValue").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("FairMrktValueUnqshr")
    End If
    If Sheet13.Range("LTCG.B8aii_FullConsiderationO").Locked = False Then
        Sheet13.Range("LTCG.B8aii_FullConsiderationO").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("FullValueConsdOthUnqshr")
    End If

    If Sheet13.Range("LTCG.B7bi_AquisitCost").Locked = False Then
        Sheet13.Range("LTCG.B7bi_AquisitCost").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("DeductSec48")("AquisitCost")
    End If
    If Sheet13.Range("LTCG.B7bii_ImproveCost").Locked = False Then
        Sheet13.Range("LTCG.B7bii_ImproveCost").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("DeductSec48")("ImproveCost")
    End If
    If Sheet13.Range("LTCG.B7biii_ExpOnTrans").Locked = False Then
        Sheet13.Range("LTCG.B7biii_ExpOnTrans").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("DeductSec48")("ExpOnTrans")
    End If
    Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("ExemptionOrDednUs54")
    If init.exists("ExemptionOrDednUs54Dtls") Then
        
    Set Nodelist = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("SaleofAssetNA")("ExemptionOrDednUs54")("ExemptionOrDednUs54Dtls")
    
    For Each Node In Nodelist
    
        ExemptionSecCode = Node("ExemptionSecCode")
        
        If ExemptionSecCode = "54D" And Sheet13.Range("LTCG.B7di_ExemptionAmount0").Locked = False Then
           Sheet13.Range("LTCG.B7di_ExemptionAmount0").value = Node("ExemptionAmount")
        End If
        
        If ExemptionSecCode = "54F" And Sheet13.Range("LTCG.B7dii_ExemptionAmount").Locked = False Then
           Sheet13.Range("LTCG.B7dii_ExemptionAmount").value = Node("ExemptionAmount")
        End If
        
        If ExemptionSecCode = "54G" And Sheet13.Range("LTCG.B7diii_ExemptionAmount").Locked = False Then
           Sheet13.Range("LTCG.B7diii_ExemptionAmount").value = Node("ExemptionAmount")
        End If
        
        If ExemptionSecCode = "54GA" And Sheet13.Range("LTCG.B7div_ExemptionAmount").Locked = False Then
           Sheet13.Range("LTCG.B7div_ExemptionAmount").value = Node("ExemptionAmount")
        End If
        
    Next Node
    End If
    
    Dim UnutilizedLtcgFlag
    UnutilizedLtcgFlag = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("UnutilizedLtcgFlag")
    
    If UnutilizedLtcgFlag = "Y" Then
        Sheet13.Range("LTCG.A9_StatusUnutilizedCG").value = "Yes"
    ElseIf UnutilizedLtcgFlag = "N" Then
        Sheet13.Range("LTCG.A9_StatusUnutilizedCG").value = "No"
    Else
        Sheet13.Range("LTCG.A9_StatusUnutilizedCG").value = "Not Applicable"
    End If
    
    If UnutilizedLtcgFlag = "Y" Then
    
        '"UnutilizedCg" start
        Dim B9_PreviousYer_1, B9_DeductionSection_1, B9_AsstAcquiredYer_1, B9_AmtUtilised_1, B9_AmtNotUsed_1 As Variant
        
        Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("UnutilizedCg")
        If init.exists("UnutilizedCgPrvYrDtls") Then
        Set nodeList1 = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("UnutilizedCg")("UnutilizedCgPrvYrDtls")
        
        B9_PreviousYer_1 = Sheet13.Range("LTCG.B9_PreviousYer_1").Column
        B9_DeductionSection_1 = Sheet13.Range("LTCG.B9_DeductionSection_1").Column
        B9_AsstAcquiredYer_1 = Sheet13.Range("LTCG.B9_AsstAcquiredYer_1").Column
        B9_AmtUtilised_1 = Sheet13.Range("LTCG.B9_AmtUtilised_1").Column
        B9_AmtNotUsed_1 = Sheet13.Range("LTCG.B9_AmtNotUsed_1").Column
        
        TotalExRow = Range("LTCG.B9_PreviousYer_1").Rows.count
        
        TotalXMLRow = nodeList1.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
        
            If Sheet13.Range("LTCG.B9_PreviousYer_1").Locked = False Then
                Sheet13.Range("LTCG.B9_PreviousYer_1").ClearContents
            End If
            If Sheet13.Range("LTCG.B9_DeductionSection_1").Locked = False Then
                Sheet13.Range("LTCG.B9_DeductionSection_1").ClearContents
            End If
            If Sheet13.Range("LTCG.B9_AsstAcquiredYer_1").Locked = False Then
                Sheet13.Range("LTCG.B9_AsstAcquiredYer_1").ClearContents
            End If
            If Sheet13.Range("LTCG.B9_AmtUtilised_1").Locked = False Then
                Sheet13.Range("LTCG.B9_AmtUtilised_1").ClearContents
            End If
            If Sheet13.Range("LTCG.B9_AmtNotUsed_1").Locked = False Then
                Sheet13.Range("LTCG.B9_AmtNotUsed_1").ClearContents
            End If
           
        End If
        
        If (TotalDiffRow > 0) Then
            AddDiffRows_UnutilizedLtcg (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet13.Range("LTCG.B9_PreviousYer_1").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node1 In nodeList1
            rowcount = rowcount + 1
            
            If Sheet13.Cells(rowcount, B9_PreviousYer_1).Locked = False Then
                Sheet13.Cells(rowcount, B9_PreviousYer_1).value = Node1("PrvYrInWhichAsstTrnsfrd")
            End If
            If Sheet13.Cells(rowcount, B9_DeductionSection_1).Locked = False Then
                Sheet13.Cells(rowcount, B9_DeductionSection_1).value = Node1("SectionClmd")
            End If
            If Sheet13.Cells(rowcount, B9_AsstAcquiredYer_1).Locked = False Then
                If Node1("YrInWhichAssetAcq") = "2017" Then
                    Sheet13.Cells(rowcount, B9_AsstAcquiredYer_1).value = "2017-18"
                ElseIf Node1("YrInWhichAssetAcq") = "2018" Then
                    Sheet13.Cells(rowcount, B9_AsstAcquiredYer_1).value = "2018-19"
                ElseIf Node1("YrInWhichAssetAcq") = "2019" Then
                    Sheet13.Cells(rowcount, B9_AsstAcquiredYer_1).value = "2019-20"
                ElseIf Node1("YrInWhichAssetAcq") = "2020" Then
                    Sheet13.Cells(rowcount, B9_AsstAcquiredYer_1).value = "2020-21"
                End If
            End If
            If Sheet13.Cells(rowcount, B9_AmtUtilised_1).Locked = False Then
                Sheet13.Cells(rowcount, B9_AmtUtilised_1).value = Node1("AmtUtilized")
            End If
            If Sheet13.Cells(rowcount, B9_AmtNotUsed_1).Locked = False Then
                Sheet13.Cells(rowcount, B9_AmtNotUsed_1).value = Node1("AmtUnutilized")
            End If
            
            cnt = cnt + 1
            
        Next Node1
        End If
    End If
    
    If Sheet13.Range("LTCG.B9_AmountDeemedOth").Locked = False Then
        Sheet13.Range("LTCG.B9_AmountDeemedOth").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("AmtDeemedLtcg")
    End If
        If Sheet13.Range("LTCG.B12_PassThroughIncome10_112A").Locked = False Then
            Sheet13.Range("LTCG.B12_PassThroughIncome10_112A").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("PassThrIncNatureLTCGUs112A")
        End If
        If Sheet13.Range("LTCG.B12_PassThroughIncome10").Locked = False Then
            Sheet13.Range("LTCG.B12_PassThroughIncome10").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("PassThrIncNatureLTCG10Per")
        End If
        If Sheet13.Range("LTCG.B12_PassThroughIncome20").Locked = False Then
            Sheet13.Range("LTCG.B12_PassThroughIncome20").value = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("PassThrIncNatureLTCG20Per")
        End If
    
    Dim B10_LtcgAmt, B10_ItemNo, B10_Country, B10_DTAA_Article, B10_RatepTreaty, B10_CertiStatus, B10_SecIT, B10_RatePITAct, B10_ApplRate As Variant
    Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")
    If init.exists("NRICgDTAA") Then
    Set init = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRICgDTAA")
    If init.exists("NRIDTAADtls") Then
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("LongTermCapGain23")("NRICgDTAA")("NRIDTAADtls")
    
    B10_LtcgAmt = Sheet13.Range("LTCG.B10_LtcgAmt").Column
    B10_ItemNo = Sheet13.Range("LTCG.B10_ItemNo").Column
    B10_Country = Sheet13.Range("LTCG.B10_Country").Column
    B10_DTAA_Article = Sheet13.Range("LTCG.B10_DTAA_Article").Column
    B10_RatepTreaty = Sheet13.Range("LTCG.B10_RatepTreaty").Column
    B10_CertiStatus = Sheet13.Range("LTCG.B10_CertiStatus").Column
    B10_SecIT = Sheet13.Range("LTCG.B10_SecIT").Column
    B10_RatePITAct = Sheet13.Range("LTCG.B10_RatePITAct").Column
    
    TotalExRow = Range("LTCG.B10_LtcgAmt").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        
        If Sheet13.Range("LTCG.B10_LtcgAmt").Locked = False Then
            Sheet13.Range("LTCG.B10_LtcgAmt").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_ItemNo").Locked = False Then
            Sheet13.Range("LTCG.B10_ItemNo").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_Country").Locked = False Then
            Sheet13.Range("LTCG.B10_Country").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_DTAA_Article").Locked = False Then
            Sheet13.Range("LTCG.B10_DTAA_Article").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_RatepTreaty").Locked = False Then
            Sheet13.Range("LTCG.B10_RatepTreaty").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_CertiStatus").Locked = False Then
            Sheet13.Range("LTCG.B10_CertiStatus").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_SecIT").Locked = False Then
            Sheet13.Range("LTCG.B10_SecIT").ClearContents
        End If
        If Sheet13.Range("LTCG.B10_RatePITAct").Locked = False Then
            Sheet13.Range("LTCG.B10_RatePITAct").ClearContents
        End If
        
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_Ltcg_NRIDTAADtls (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("LTCG.B10_LtcgAmt").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
       
        If Sheet13.Cells(rowcount, B10_LtcgAmt).Locked = False Then
            Sheet13.Cells(rowcount, B10_LtcgAmt).value = Node1("DTAAamt")
        End If
        If Sheet13.Cells(rowcount, B10_ItemNo).Locked = False Then
            Sheet13.Cells(rowcount, B10_ItemNo).value = Node1("ItemNoincl")
        End If
        
        iCountry = Node1("CountryCode")
        sCountry = Findtext(iCountry, "Country")
            
        If Sheet13.Cells(rowcount, B10_Country).Locked = False Then
            Sheet13.Cells(rowcount, B10_Country).value = sCountry
        End If
        
        If Sheet13.Cells(rowcount, B10_DTAA_Article).Locked = False Then
            Sheet13.Cells(rowcount, B10_DTAA_Article).value = Node1("DTAAarticle")
        End If
        If Sheet13.Cells(rowcount, B10_RatepTreaty).Locked = False Then
            Sheet13.Cells(rowcount, B10_RatepTreaty).value = Node1("RateAsPerTreaty")
        End If
        
        TaxRescertifiedFlag = Node1("TaxRescertifiedFlag")
        
        If TaxRescertifiedFlag = "Y" Then
            TaxRescertifiedFlag = "YES"
        ElseIf TaxRescertifiedFlag = "N" Then
            TaxRescertifiedFlag = "NO"
        Else
            TaxRescertifiedFlag = "(Select)"
        End If
        If Sheet13.Cells(rowcount, B10_CertiStatus).Locked = False Then
            Sheet13.Cells(rowcount, B10_CertiStatus).value = TaxRescertifiedFlag
        End If
        If Sheet13.Cells(rowcount, B10_SecIT).Locked = False Then
            Sheet13.Cells(rowcount, B10_SecIT).value = Node1("SecITAct")
        End If
        If Sheet13.Cells(rowcount, B10_RatePITAct).Locked = False Then
            Sheet13.Cells(rowcount, B10_RatePITAct).value = Node1("RateAsPerITAct")
        End If
        
        cnt = cnt + 1
        
    Next Node1
    End If
    End If
    Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs54") Then
    
    Dim DateOfTransfer54, CostOfHouse54, DateOfPursConst54, AmtdepCG54, AmtDedClaimed54 As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54")
    
    DateOfTransfer54 = Sheet13.Range("Deduction.54.DateOfTransfer").Column
    CostOfHouse54 = Sheet13.Range("Deduction.54.CostOfHouse").Column
    DateOfPursConst54 = Sheet13.Range("Deduction.54.DateOfPursConst").Column
    AmtdepCG54 = Sheet13.Range("Deduction.54.AmtdepCG").Column
    AmtDedClaimed54 = Sheet13.Range("Deduction.54.AmtDedClaimed").Column
    
    TotalExRow = Range("Deduction.54.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet13.Range("Deduction.54.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.54.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.54.CostOfHouse").Locked = False Then
            Sheet13.Range("Deduction.54.CostOfHouse").ClearContents
        End If
        If Sheet13.Range("Deduction.54.DateOfPursConst").Locked = False Then
            Sheet13.Range("Deduction.54.DateOfPursConst").ClearContents
        End If
        If Sheet13.Range("Deduction.54.AmtdepCG").Locked = False Then
            Sheet13.Range("Deduction.54.AmtdepCG").ClearContents
        End If
        If Sheet13.Range("Deduction.54.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.54.AmtDedClaimed").ClearContents
        End If
        
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs54 (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.54.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    Dim dob, DOB1, strdate1 As Variant
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofTransfer")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        If Sheet13.Cells(rowcount, DateOfTransfer54).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer54).value = strdate
        End If
        If Sheet13.Cells(rowcount, CostOfHouse54).Locked = False Then
            Sheet13.Cells(rowcount, CostOfHouse54).value = Node1("CostofNewResHouse")
        End If
        
        DOB1 = Node1("DateofPurchase")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        If Sheet13.Cells(rowcount, DateOfPursConst54).Locked = False Then
            Sheet13.Cells(rowcount, DateOfPursConst54).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtdepCG54).Locked = False Then
            Sheet13.Cells(rowcount, AmtdepCG54).value = Node1("AmtDeposited")
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed54).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed54).value = Node1("AmtDeducted")
        End If
             
        cnt = cnt + 1
        
    Next Node1
    End If
    
    Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs54B") Then
    
    '"DeducClaimDtlsUs54B" start
    Dim DateOfTransfer54B, CostOfAgriLand54B, DateOfPurLand54B, AmtdepCG54B, AmtDedClaimed54B As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54B")
    
    DateOfTransfer54B = Sheet13.Range("Deduction.54B.DateOfTransfer").Column
    CostOfAgriLand54B = Sheet13.Range("Deduction.54B.CostOfAgriLand").Column
    DateOfPurLand54B = Sheet13.Range("Deduction.54B.DateOfPurLand").Column
    AmtdepCG54B = Sheet13.Range("Deduction.54B.AmtdepCG").Column
    AmtDedClaimed54B = Sheet13.Range("Deduction.54B.AmtDedClaimed").Column
    
    TotalExRow = Range("Deduction.54B.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet13.Range("Deduction.54B.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.54B.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.54B.CostOfAgriLand").Locked = False Then
            Sheet13.Range("Deduction.54B.CostOfAgriLand").ClearContents
        End If
        If Sheet13.Range("Deduction.54B.DateOfPurLand").Locked = False Then
            Sheet13.Range("Deduction.54B.DateOfPurLand").ClearContents
        End If
        If Sheet13.Range("Deduction.54B.AmtdepCG").Locked = False Then
            Sheet13.Range("Deduction.54B.AmtdepCG").ClearContents
        End If
        If Sheet13.Range("Deduction.54B.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.54B.AmtDedClaimed").ClearContents
        End If
        
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs54B (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.54B.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofTransfer")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfTransfer54B).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer54B).value = strdate
        End If
        If Sheet13.Cells(rowcount, CostOfAgriLand54B).Locked = False Then
            Sheet13.Cells(rowcount, CostOfAgriLand54B).value = Node1("CostofNewAgriLand")
        End If
        
        DOB1 = Node1("DateofPurchase")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        If Sheet13.Cells(rowcount, DateOfPurLand54B).Locked = False Then
            Sheet13.Cells(rowcount, DateOfPurLand54B).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtdepCG54B).Locked = False Then
            Sheet13.Cells(rowcount, AmtdepCG54B).value = Node1("AmtDeposited")
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed54B).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed54B).value = Node1("AmtDeducted")
        End If
             
        cnt = cnt + 1
        
    Next Node1
    End If
    
    Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs54D") Then
    '"DeducClaimDtlsUs54D" start
    Dim DateOfTransfer54D, CostOfPurcConst54D, DateOfPurLandB54D, AmtdepCG54D, AmtDedClaimed54D As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54D")
    
    DateOfTransfer54D = Sheet13.Range("Deduction.54D.DateOfTransfer").Column
    CostOfPurcConst54D = Sheet13.Range("Deduction.54D.CostOfPurcConst").Column
    DateOfPurLandB54D = Sheet13.Range("Deduction.54B.DateOfPurLandB").Column
    AmtdepCG54D = Sheet13.Range("Deduction.54D.AmtdepCG").Column
    AmtDedClaimed54D = Sheet13.Range("Deduction.54D.AmtDedClaimed").Column
    
    
    TotalExRow = Range("Deduction.54D.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet13.Range("Deduction.54D.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.54D.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.54D.CostOfPurcConst").Locked = False Then
            Sheet13.Range("Deduction.54D.CostOfPurcConst").ClearContents
        End If
        If Sheet13.Range("Deduction.54B.DateOfPurLandB").Locked = False Then
            Sheet13.Range("Deduction.54B.DateOfPurLandB").ClearContents
        End If
        If Sheet13.Range("Deduction.54D.AmtdepCG").Locked = False Then
            Sheet13.Range("Deduction.54D.AmtdepCG").ClearContents
        End If
        If Sheet13.Range("Deduction.54D.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.54D.AmtDedClaimed").ClearContents
        End If
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs54D (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.54D.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofAcquisition")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        If Sheet13.Cells(rowcount, DateOfTransfer54D).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer54D).value = strdate
        End If
        If Sheet13.Cells(rowcount, CostOfPurcConst54D).Locked = False Then
            Sheet13.Cells(rowcount, CostOfPurcConst54D).value = Node1("CostofNewLandBuilding")
        End If
        
        DOB1 = Node1("DateofPurchase")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        If Sheet13.Cells(rowcount, DateOfPurLandB54D).Locked = False Then
            Sheet13.Cells(rowcount, DateOfPurLandB54D).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtdepCG54D).Locked = False Then
            Sheet13.Cells(rowcount, AmtdepCG54D).value = Node1("AmtDeposited")
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed54D).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed54D).value = Node1("AmtDeducted")
        End If
             
        cnt = cnt + 1
        
    Next Node1
   End If
   
   Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs54EC") Then
   '"DeducClaimDtlsUs54EC" start
    Dim DateOfTransfer54EC, AmountInvestd54EC, DateOfInvestment54EC, AmtDedClaimed54EC As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54EC")
    
    DateOfTransfer54EC = Sheet13.Range("Deduction.54EC.DateOfTransfer").Column
    AmountInvestd54EC = Sheet13.Range("Deduction.54EC.AmountInvestd").Column
    DateOfInvestment54EC = Sheet13.Range("Deduction.54EC.DateOfInvestment").Column
    AmtDedClaimed54EC = Sheet13.Range("Deduction.54EC.AmtDedClaimed").Column
    
    TotalExRow = Range("Deduction.54EC.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet13.Range("Deduction.54EC.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.54EC.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.54EC.AmountInvestd").Locked = False Then
            Sheet13.Range("Deduction.54EC.AmountInvestd").ClearContents
        End If
        If Sheet13.Range("Deduction.54EC.DateOfInvestment").Locked = False Then
            Sheet13.Range("Deduction.54EC.DateOfInvestment").ClearContents
        End If
        If Sheet13.Range("Deduction.54EC.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.54EC.AmtDedClaimed").ClearContents
        End If
        
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs54EC (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.54EC.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofTransfer")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfTransfer54EC).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer54EC).value = strdate
        End If
        If Sheet13.Cells(rowcount, AmountInvestd54EC).Locked = False Then
            Sheet13.Cells(rowcount, AmountInvestd54EC).value = Node1("AmtInvested")
        End If
        
        DOB1 = Node1("DateofInvestment")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        If Sheet13.Cells(rowcount, DateOfInvestment54EC).Locked = False Then
            Sheet13.Cells(rowcount, DateOfInvestment54EC).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed54EC).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed54EC).value = Node1("AmtDeducted")
        End If
             
        cnt = cnt + 1
        
    Next Node1
   End If
   
   Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs54F") Then
   '"DeducClaimDtlsUs54F" start
    Dim DateOfTransfer54F, CostOfHouse54F, DateOfPurc54F, AmtdepCG54F, AmtDedClaimed54F As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54F")
    
    DateOfTransfer54F = Sheet13.Range("Deduction.54F.DateOfTransfer").Column
    CostOfHouse54F = Sheet13.Range("Deduction.54F.CostOfHouse").Column
    DateOfPurc54F = Sheet13.Range("Deduction.54F.DateOfPurc").Column
    AmtdepCG54F = Sheet13.Range("Deduction.54F.AmtdepCG").Column
    AmtDedClaimed54F = Sheet13.Range("Deduction.54F.AmtDedClaimed").Column

    TotalExRow = Range("Deduction.54F.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet13.Range("Deduction.54F.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.54F.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.54F.CostOfHouse").Locked = False Then
            Sheet13.Range("Deduction.54F.CostOfHouse").ClearContents
        End If
        If Sheet13.Range("Deduction.54F.DateOfPurc").Locked = False Then
            Sheet13.Range("Deduction.54F.DateOfPurc").ClearContents
        End If
        If Sheet13.Range("Deduction.54F.AmtdepCG").Locked = False Then
            Sheet13.Range("Deduction.54F.AmtdepCG").ClearContents
        End If
        If Sheet13.Range("Deduction.54F.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.54F.AmtDedClaimed").ClearContents
        End If
        
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs54F (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.54F.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofTransfer")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfTransfer54F).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer54F).value = strdate
        End If
        If Sheet13.Cells(rowcount, CostOfHouse54F).Locked = False Then
            Sheet13.Cells(rowcount, CostOfHouse54F).value = Node1("CostofNewResHouse")
        End If
        
        DOB1 = Node1("DateofPurchase")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfPurc54F).Locked = False Then
            Sheet13.Cells(rowcount, DateOfPurc54F).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtdepCG54F).Locked = False Then
            Sheet13.Cells(rowcount, AmtdepCG54F).value = Node1("AmtDeposited")
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed54F).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed54F).value = Node1("AmtDeducted")
        End If
        
        cnt = cnt + 1
        
    Next Node1
   End If
   
   Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs54G") Then
   '"DeducClaimDtlsUs54G" start
    Dim DateOfTransfer54G, CostOfExpenses54G, DateOfPurc54G, AmtdepCG54G, AmtDedClaimed54G As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54G")
    
    DateOfTransfer54G = Sheet13.Range("Deduction.54G.DateOfTransfer").Column
    CostOfExpenses54G = Sheet13.Range("Deduction.54G.CostOfExpenses").Column
    DateOfPurc54G = Sheet13.Range("Deduction.54G.DateOfPurc").Column
    AmtdepCG54G = Sheet13.Range("Deduction.54G.AmtdepCG").Column
    AmtDedClaimed54G = Sheet13.Range("Deduction.54G.AmtDedClaimed").Column
    
    TotalExRow = Range("Deduction.54G.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        
        If Sheet13.Range("Deduction.54G.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.54G.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.54G.CostOfExpenses").Locked = False Then
            Sheet13.Range("Deduction.54G.CostOfExpenses").ClearContents
        End If
        If Sheet13.Range("Deduction.54G.DateOfPurc").Locked = False Then
            Sheet13.Range("Deduction.54G.DateOfPurc").ClearContents
        End If
        If Sheet13.Range("Deduction.54G.AmtdepCG").Locked = False Then
            Sheet13.Range("Deduction.54G.AmtdepCG").ClearContents
        End If
        If Sheet13.Range("Deduction.54G.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.54G.AmtDedClaimed").ClearContents
        End If
    
    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs54G (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.54G.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofTransfer")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfTransfer54G).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer54G).value = strdate
        End If
        If Sheet13.Cells(rowcount, CostOfExpenses54G).Locked = False Then
            Sheet13.Cells(rowcount, CostOfExpenses54G).value = Node1("CostofNewAsset")
        End If
        
        DOB1 = Node1("DateofPurchase")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfPurc54G).Locked = False Then
            Sheet13.Cells(rowcount, DateOfPurc54G).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtdepCG54G).Locked = False Then
            Sheet13.Cells(rowcount, AmtdepCG54G).value = Node1("AmtDeposited")
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed54G).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed54G).value = Node1("AmtDeducted")
        End If
        
        cnt = cnt + 1
        
    Next Node1
   End If
   
'   Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
'    If init.exists("DeducClaimDtlsUs54GA") Then
'   '"DeducClaimDtlsUs54GA" start
'    Dim DateOfTransfer54GA, CostOfExpenses54GA, DateOfPurc54GA, AmtdepCG54GA, AmtDedClaimed54GA As Variant
'
'    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs54GA")
'
'    DateOfTransfer54GA = Sheet13.Range("Deduction.54GA.DateOfTransfer").Column
'    CostOfExpenses54GA = Sheet13.Range("Deduction.54GA.CostOfExpenses").Column
'    DateOfPurc54GA = Sheet13.Range("Deduction.54GA.DateOfPurc").Column
'    AmtdepCG54GA = Sheet13.Range("Deduction.54GA.AmtdepCG").Column
'    AmtDedClaimed54GA = Sheet13.Range("Deduction.54GA.AmtDedClaimed").Column
'
'    TotalExRow = Range("Deduction.54GA.DateOfTransfer").Rows.count
'
'    TotalXMLRow = nodeList1.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'
'        If Sheet13.Range("Deduction.54GA.DateOfTransfer").Locked = False Then
'            Sheet13.Range("Deduction.54GA.DateOfTransfer").ClearContents
'        End If
'        If Sheet13.Range("Deduction.54GA.CostOfExpenses").Locked = False Then
'            Sheet13.Range("Deduction.54GA.CostOfExpenses").ClearContents
'        End If
'        If Sheet13.Range("Deduction.54GA.DateOfPurc").Locked = False Then
'            Sheet13.Range("Deduction.54GA.DateOfPurc").ClearContents
'        End If
'        If Sheet13.Range("Deduction.54GA.AmtdepCG").Locked = False Then
'            Sheet13.Range("Deduction.54GA.AmtdepCG").ClearContents
'        End If
'        If Sheet13.Range("Deduction.54GA.AmtDedClaimed").Locked = False Then
'            Sheet13.Range("Deduction.54GA.AmtDedClaimed").ClearContents
'        End If
'
'    End If
'
'    If (TotalDiffRow > 0) Then
'        AddDiffRows_DeducClaimDtlsUs54GA (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet13.Range("Deduction.54GA.DateOfTransfer").name)
'    rowcount = rowcount - 1
'    cnt = 0
'
'    For Each Node1 In nodeList1
'        rowcount = rowcount + 1
'
'        dob = Node1("DateofTransfer")
'        If dob <> "" Then
'            YYYY = Mid(dob, 1, 4)
'            MM = Mid(dob, 6, 2)
'            DD = Mid(dob, 9, 2)
'            strdate = DD & "/" & MM & "/" & YYYY
'        End If
'
'        If Sheet13.Cells(rowcount, DateOfTransfer54GA).Locked = False Then
'            Sheet13.Cells(rowcount, DateOfTransfer54GA).value = strdate
'        End If
'        If Sheet13.Cells(rowcount, CostOfExpenses54GA).Locked = False Then
'            Sheet13.Cells(rowcount, CostOfExpenses54GA).value = Node1("CostofNewAsset")
'        End If
'
'        DOB1 = Node1("DateofPurchase")
'        If DOB1 <> "" Then
'            YYYY = Mid(DOB1, 1, 4)
'            MM = Mid(DOB1, 6, 2)
'            DD = Mid(DOB1, 9, 2)
'            strdate1 = DD & "/" & MM & "/" & YYYY
'        End If
'
'        If Sheet13.Cells(rowcount, DateOfPurc54GA).Locked = False Then
'            Sheet13.Cells(rowcount, DateOfPurc54GA).value = strdate1
'        End If
'        If Sheet13.Cells(rowcount, AmtdepCG54GA).Locked = False Then
'            Sheet13.Cells(rowcount, AmtdepCG54GA).value = Node1("AmtDeposited")
'        End If
'        If Sheet13.Cells(rowcount, AmtDedClaimed54GA).Locked = False Then
'            Sheet13.Cells(rowcount, AmtDedClaimed54GA).value = Node1("AmtDeducted")
'        End If
'
'        cnt = cnt + 1
'
'    Next Node1
'   End If
   
   Set init = jsonObject("ScheduleCGFor23")("DeducClaimInfo")
    If init.exists("DeducClaimDtlsUs115F") Then
   '"DeducClaimDtlsUs115F" start
    Dim DateOfTransfer115F, AmtInvt115F, DateOfInvest115F, AmtDedClaimed115F As Variant
    
    Set nodeList1 = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeducClaimDtlsUs115F")
    
    DateOfTransfer115F = Sheet13.Range("Deduction.115F.DateOfTransfer").Column
    AmtInvt115F = Sheet13.Range("Deduction.115F.AmtInvt").Column
    DateOfInvest115F = Sheet13.Range("Deduction.115F.DateOfInvest").Column
    AmtDedClaimed115F = Sheet13.Range("Deduction.115F.AmtDedClaimed").Column

    TotalExRow = Range("Deduction.115F.DateOfTransfer").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
                
        If Sheet13.Range("Deduction.115F.DateOfTransfer").Locked = False Then
            Sheet13.Range("Deduction.115F.DateOfTransfer").ClearContents
        End If
        If Sheet13.Range("Deduction.115F.AmtInvt").Locked = False Then
            Sheet13.Range("Deduction.115F.AmtInvt").ClearContents
        End If
        If Sheet13.Range("Deduction.115F.DateOfInvest").Locked = False Then
            Sheet13.Range("Deduction.115F.DateOfInvest").ClearContents
        End If
        If Sheet13.Range("Deduction.115F.AmtDedClaimed").Locked = False Then
            Sheet13.Range("Deduction.115F.AmtDedClaimed").ClearContents
        End If

    End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_DeducClaimDtlsUs115F (TotalDiffRow)
        End If
    
    rowcount = getRowNo(Sheet13.Range("Deduction.115F.DateOfTransfer").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
        
        dob = Node1("DateofTransfer")
        If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfTransfer115F).Locked = False Then
            Sheet13.Cells(rowcount, DateOfTransfer115F).value = strdate
        End If
        If Sheet13.Cells(rowcount, AmtInvt115F).Locked = False Then
            Sheet13.Cells(rowcount, AmtInvt115F).value = Node1("AmtInvested")
        End If
        
        DOB1 = Node1("DateofInvestment")
        If DOB1 <> "" Then
            YYYY = Mid(DOB1, 1, 4)
            MM = Mid(DOB1, 6, 2)
            DD = Mid(DOB1, 9, 2)
            strdate1 = DD & "/" & MM & "/" & YYYY
        End If
        
        If Sheet13.Cells(rowcount, DateOfInvest115F).Locked = False Then
            Sheet13.Cells(rowcount, DateOfInvest115F).value = strdate1
        End If
        If Sheet13.Cells(rowcount, AmtDedClaimed115F).Locked = False Then
            Sheet13.Cells(rowcount, AmtDedClaimed115F).value = Node1("AmtDeducted")
        End If
        
        cnt = cnt + 1
        
    Next Node1
   End If
    'Sheet13.Range("DEDN.1c_TotDeductClaim").value = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("TotDeductClaim")
    Sheet13.Range("DEDN.D2_DeductPAN").value = jsonObject("ScheduleCGFor23")("DeducClaimInfo")("DeductPAN")
    
If Sheet13.Range("AccSTCG.Upto15Of6").Locked = False Then
    Sheet13.Range("AccSTCG.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder15Per")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccSTCG.Upto15Of9").Locked = False Then
    Sheet13.Range("AccSTCG.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder15Per")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccSTCG.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccSTCG.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder15Per")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccSTCG.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccSTCG.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder15Per")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccSTCG.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccSTCG.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder15Per")("DateRange")("Up16Of3To31Of3")
End If


If Sheet13.Range("AccSTCG30.Upto15Of6").Locked = False Then
    Sheet13.Range("AccSTCG30.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder30Per")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccSTCG30.Upto15Of9").Locked = False Then
    Sheet13.Range("AccSTCG30.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder30Per")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("ACCSTCG30.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("ACCSTCG30.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder30Per")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("ACCSTCG30.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("ACCSTCG30.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder30Per")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("ACCSTCG30.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("ACCSTCG30.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnder30Per")("DateRange")("Up16Of3To31Of3")
End If


If Sheet13.Range("AccSTCGOTH.Upto15Of6").Locked = False Then
    Sheet13.Range("AccSTCGOTH.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderAppRate")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccSTCGOTH.Upto15Of9").Locked = False Then
    Sheet13.Range("AccSTCGOTH.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderAppRate")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccSTCGOTH.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccSTCGOTH.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderAppRate")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccSTCGOTH.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccSTCGOTH.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderAppRate")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccSTCGOTH.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccSTCGOTH.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderAppRate")("DateRange")("Up16Of3To31Of3")
End If


If Sheet13.Range("AccSTCGDTAA.Upto15Of6").Locked = False Then
    Sheet13.Range("AccSTCGDTAA.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderDTAARate")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccSTCGDTAA.Upto15Of9").Locked = False Then
    Sheet13.Range("AccSTCGDTAA.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderDTAARate")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccSTCGDTAA.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccSTCGDTAA.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderDTAARate")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccSTCGDTAA.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccSTCGDTAA.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderDTAARate")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccSTCGDTAA.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccSTCGDTAA.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("ShortTermUnderDTAARate")("DateRange")("Up16Of3To31Of3")
End If


If Sheet13.Range("AccLTCG.Upto15Of6").Locked = False Then
    Sheet13.Range("AccLTCG.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder10Per")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccLTCG.Upto15Of9").Locked = False Then
    Sheet13.Range("AccLTCG.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder10Per")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccLTCG.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccLTCG.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder10Per")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccLTCG.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccLTCG.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder10Per")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccLTCG.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccLTCG.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder10Per")("DateRange")("Up16Of3To31Of3")
End If


If Sheet13.Range("AccLTCGNP.Upto15Of6").Locked = False Then
    Sheet13.Range("AccLTCGNP.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder20Per")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccLTCGNP.Upto15Of9").Locked = False Then
    Sheet13.Range("AccLTCGNP.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder20Per")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccLTCGNP.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccLTCGNP.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder20Per")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccLTCGNP.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccLTCGNP.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder20Per")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccLTCGNP.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccLTCGNP.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnder20Per")("DateRange")("Up16Of3To31Of3")
End If


If Sheet13.Range("AccLTCGDTAA.Upto15Of6").Locked = False Then
    Sheet13.Range("AccLTCGDTAA.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnderDTAARate")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccLTCGDTAA.Upto15Of9").Locked = False Then
    Sheet13.Range("AccLTCGDTAA.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnderDTAARate")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccLTCGDTAA.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccLTCGDTAA.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnderDTAARate")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccLTCGDTAA.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccLTCGDTAA.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnderDTAARate")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccLTCGDTAA.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccLTCGDTAA.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("LongTermUnderDTAARate")("DateRange")("Up16Of3To31Of3")
End If
    

If Sheet13.Range("AccVDA.Upto15Of6").Locked = False Then
    Sheet13.Range("AccVDA.Upto15Of6").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("VDATrnsfGainsUnder30Per")("DateRange")("Upto15Of6")
End If
If Sheet13.Range("AccVDA.Upto15Of9").Locked = False Then
    Sheet13.Range("AccVDA.Upto15Of9").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("VDATrnsfGainsUnder30Per")("DateRange")("Upto15Of9")
End If
If Sheet13.Range("AccVDA.Up16Of9To15Of12").Locked = False Then
    Sheet13.Range("AccVDA.Up16Of9To15Of12").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("VDATrnsfGainsUnder30Per")("DateRange")("Up16Of9To15Of12")
End If
If Sheet13.Range("AccVDA.Up16Of12To15Of3").Locked = False Then
    Sheet13.Range("AccVDA.Up16Of12To15Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("VDATrnsfGainsUnder30Per")("DateRange")("Up16Of12To15Of3")
End If
If Sheet13.Range("AccVDA.Up16Of3To31Of3").Locked = False Then
    Sheet13.Range("AccVDA.Up16Of3To31Of3").value = jsonObject("ScheduleCGFor23")("AccruOrRecOfCG")("VDATrnsfGainsUnder30Per")("DateRange")("Up16Of3To31Of3")
End If
        
End Function

Function ESRImport(jsonText As String)
On Error Resume Next

Dim jsonDictionary, jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
If Sheet12.Range("ESR1i.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR1i.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351I")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1ii.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR1ii.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Ii")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1iia.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR1iia.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Iia")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1iii.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR1iii.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Iii")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1iv.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR1iv.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Iv")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR2AA.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR2AA.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section352AA")("deductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR2AB.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR2AB.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section352AB")("deductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR35CCC.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR35CCC.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section35CCC")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR35CCD.AmtDebPL").Locked = False Then
    Sheet12.Range("ESR35CCD.AmtDebPL").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section35CCD")("DeductUs35")("AmtDebPL")
End If


If Sheet12.Range("ESR1i.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR1i.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351I")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1ii.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR1ii.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Ii")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1iia.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR1iia.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Iia")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1iii.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR1iii.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Iii")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR1iv.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR1iv.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section351Iv")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR2AA.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR2AA.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section352AA")("deductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR2AB.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR2AB.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section352AB")("deductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR35CCC.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR35CCC.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section35CCC")("DeductUs35")("AmtDebPL")
End If
If Sheet12.Range("ESR35CCD.AmtUs35Allowable").Locked = False Then
    Sheet12.Range("ESR35CCD.AmtUs35Allowable").value = jsonObject("form3CD")("ScheduleESR")("DeductionUs35")("Section35CCD")("DeductUs35")("AmtDebPL")
End If


End Function
Function ScheduleDPMImport(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
If Sheet10.Range("DPM15.WDVFirstDayBAC").Locked = False Then
    Sheet10.Range("DPM15.WDVFirstDayBAC").value = jsonObject("form3CD")("PlantMachinery")("ScheduleDPM")("Rate15")("DepreciationDetail")("AdjustAccTax115BACAmt")
End If
If Sheet10.Range("DPM30.WDVFirstDayBAC").Locked = False Then
    Sheet10.Range("DPM30.WDVFirstDayBAC").value = jsonObject("form3CD")("PlantMachinery")("ScheduleDPM")("Rate30")("DepreciationDetail")("AdjustAccTax115BACAmt")
End If

If Sheet10.Range("DPM40.WDVFirstDayBAC").Locked = False Then
    Sheet10.Range("DPM40.WDVFirstDayBAC").value = jsonObject("form3CD")("PlantMachinery")("ScheduleDPM")("Rate40")("DepreciationDetail")("AdjustAccTax115BACAmt")
End If

If Sheet10.Range("DPM45.WDVFirstDayBAC").Locked = False Then
    Sheet10.Range("DPM45.WDVFirstDayBAC").value = jsonObject("form3CD")("PlantMachinery")("ScheduleDPM")("Rate45")("DepreciationDetail")("AdjustAccTax115BACAmt")
End If

If Sheet10.Range("DPM15.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DPM15.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate15")("DepreciationDetail")("AdditionsGrThan180Days")
End If

If Sheet10.Range("DPM30.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DPM30.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate30")("DepreciationDetail")("AdditionsGrThan180Days")
End If

If Sheet10.Range("DPM40.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DPM40.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate40")("DepreciationDetail")("AdditionsGrThan180Days")
End If

If Sheet10.Range("DPM45.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DPM45.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate45")("DepreciationDetail")("AdditionsGrThan180Days")
End If

If Sheet10.Range("DPM15.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DPM15.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate15")("DepreciationDetail")("AdditionsLessThan180Days")
End If

If Sheet10.Range("DPM30.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DPM30.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate30")("DepreciationDetail")("AdditionsLessThan180Days")
End If
If Sheet10.Range("DPM40.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DPM40.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate40")("DepreciationDetail")("AdditionsLessThan180Days")
End If

If Sheet10.Range("DPM45.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DPM45.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDPM")("PlantMachinery")("Rate45")("DepreciationDetail")("AdditionsLessThan180Days")
End If



If Sheet10.Range("RealizationTotalPeriod").Locked = False Then
    Sheet10.Range("RealizationTotalPeriod").value = jsonObject("lastFiledITR")("ScheduleDPM")("PlantMachinery")("Rate15")("DepreciationDetail")("WdvFirstDay")
End If


'change-23 26.05.2022

If Sheet10.Range("DPM15.WDVFirstDay").Locked = False Then
    Sheet10.Range("DPM15.WDVFirstDay").value = jsonObject("lastFiledITR")("scheduleDPM")("PlantMachinery")("Rate15")("DepreciationDetail")("WdvfirstDay")
End If

If Sheet10.Range("DPM30.WDVFirstDay").Locked = False Then
    Sheet10.Range("DPM30.WDVFirstDay").value = jsonObject("lastFiledITR")("scheduleDPM")("PlantMachinery")("Rate30")("DepreciationDetail")("WdvfirstDay")
End If
If Sheet10.Range("DPM40.WDVFirstDay").Locked = False Then
    Sheet10.Range("DPM40.WDVFirstDay").value = jsonObject("lastFiledITR")("scheduleDPM")("PlantMachinery")("Rate40")("DepreciationDetail")("WdvfirstDay")
End If

If Sheet10.Range("DPM45.WDVFirstDay3").Locked = False Then
    Sheet10.Range("DPM45.WDVFirstDay3").value = jsonObject("lastFiledITR")("scheduleDPM")("PlantMachinery")("Rate45")("depreciationDetail")("wdvfirstDay")
End If

'Change-10, Test-019, Remark - Adding lastFiledITR values
If Sheet10.Range("DAOB5.WDVFirstDay").Locked = False Then
    Sheet10.Range("DAOB5.WDVFirstDay").value = jsonObject("lastFiledITR")("scheduleDOA")("building")("Rate5")("DepreciationDetail")("WdvfirstDay")
End If
If Sheet10.Range("DAOB10.WDVFirstDay").Locked = False Then
    Sheet10.Range("DAOB10.WDVFirstDay").value = jsonObject("lastFiledITR")("scheduleDOA")("building")("Rate10")("DepreciationDetail")("WdvfirstDay")
End If
If Sheet10.Range("DAOB100.WDVFirstDay").Locked = False Then
    Sheet10.Range("DAOB100.WDVFirstDay").value = jsonObject("lastFiledITR")("scheduleDOA")("building")("Rate40")("DepreciationDetail")("WdvfirstDay")
End If

'
'If Sheet10.Range("DAOB5.WDVFirstDay").Locked = False Then
'    Sheet10.Range("DAOB5.WDVFirstDay").value = jsonObject("lastFiledITR")("ScheduleDOA")("building")("Rate5")("DepreciationDetail")("WdvfirstDay")
'End If
'
'If Sheet10.Range("DAOB10.WDVFirstDay").Locked = False Then
'    Sheet10.Range("DAOB10.WDVFirstDay").value = jsonObject("lastFiledITR")("ScheduleDOA")("building")("Rate10")("DepreciationDetail")("WdvfirstDay")
'End If
'If Sheet10.Range("DAOB40.WDVFirstDay").Locked = False Then
'    Sheet10.Range("DAOB40.WDVFirstDay").value = jsonObject("lastFiledITR")("ScheduleDOA")("building")("Rate40")("DepreciationDetail")("WdvfirstDay")
'End If
If Sheet10.Range("DAOB5.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DAOB5.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDOA")("building")("Rate5")("DepreciationDetail")("AdditionsGrThan180Days")
End If
If Sheet10.Range("DAOB10.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DAOB10.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDOA")("building")("Rate10")("DepreciationDetail")("AdditionsGrThan180Days")
End If
If Sheet10.Range("DAOB100.AdditionsGrThan180Days").Locked = False Then
    Sheet10.Range("DAOB100.AdditionsGrThan180Days").value = jsonObject("form3CD")("ScheduleDOA")("building")("Rate40")("DepreciationDetail")("AdditionsGrThan180Days")
End If
If Sheet10.Range("DAOB5.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DAOB5.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDOA")("building")("Rate5")("DepreciationDetail")("AdditionsLessThan180Days")
End If

If Sheet10.Range("DAOB10.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DAOB10.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDOA")("building")("Rate10")("DepreciationDetail")("AdditionsLessThan180Days")
End If
If Sheet10.Range("DAOB100.AdditionsLessThan180Days").Locked = False Then
    Sheet10.Range("DAOB100.AdditionsLessThan180Days").value = jsonObject("form3CD")("ScheduleDOA")("building")("Rate40")("DepreciationDetail")("AdditionsLessThan180Days")
End If



 End Function

Function ImportITR3ScheduleBP(jsonText As String)
On Error Resume Next
Dim init As Object
Dim jsonDictionary As Object
Dim Node, Nodelist, jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
Dim InterestDisAllowUs23SMEAct, DeemIncUs32AD, DeemIncUs33AB, DeemIncUs33ABA, DeemIncUs35ABA, DeemIncUs35ABB, DeemIncUs40A3A, DeemIncUs72A, DeemIncUs80HHD
Dim DeemIncUs80IA
InterestDisAllowUs23SMEAct = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("InterestDisAllowUs23SMEAct")
If InterestDisAllowUs23SMEAct <> "" And Sheet9.Range("sheet11.InterestDisAllowUs23SMEAct").Locked = False Then
    Sheet9.Range("sheet11.InterestDisAllowUs23SMEAct").value = UCase(InterestDisAllowUs23SMEAct)
End If
DeemIncUs32AD = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs32AD")
If DeemIncUs32AD <> "" And Sheet9.Range("sheet11.DeemIncUs32AD").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs32AD").value = UCase(DeemIncUs32AD)
End If
DeemIncUs33AB = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs33AB")
If DeemIncUs33AB <> "" And Sheet9.Range("sheet11.DeemIncUs33AB").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs33AB").value = UCase(DeemIncUs33AB)
End If
DeemIncUs33ABA = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs33ABA")
If DeemIncUs33ABA <> "" And Sheet9.Range("sheet11.DeemIncUs33ABA").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs33ABA").value = UCase(DeemIncUs33ABA)
End If
DeemIncUs35ABA = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs35ABA")
If DeemIncUs35ABA <> "" And Sheet9.Range("sheet11.DeemIncUs35ABA").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs35ABA").value = UCase(DeemIncUs35ABA)
End If
DeemIncUs35ABB = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs35ABB")
If DeemIncUs35ABB <> "" And Sheet9.Range("sheet11.DeemIncUs35ABB").Locked = False Then
     Sheet9.Range("sheet11.DeemIncUs35ABB").value = UCase(DeemIncUs35ABB)
End If
DeemIncUs40A3A = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs40A3A")
If DeemIncUs40A3A <> "" And Sheet9.Range("sheet11.DeemIncUs40A").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs40A").value = UCase(DeemIncUs40A3A)
End If
DeemIncUs72A = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs72A")
If DeemIncUs72A <> "" And Sheet9.Range("sheet11.DeemIncUs72A").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs72A").value = UCase(DeemIncUs72A)
End If
DeemIncUs80HHD = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs80HHD")
If DeemIncUs80HHD <> "" And Sheet9.Range("sheet11.DeemIncUs80HHD").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs80HHD").value = UCase(DeemIncUs80HHD)
End If
DeemIncUs80IA = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs80IA")
If DeemIncUs80IA <> "" And Sheet9.Range("sheet11.DeemIncUs80IA").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs80IA").value = UCase(DeemIncUs80IA)
End If
DeemIncUs32AD = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs32AD")
If DeemIncUs32AD <> "" And Sheet9.Range("sheet11.DeemIncUs32AD").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs32AD").value = UCase(DeemIncUs32AD)
End If
DeemIncUs33AB = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs33AB")
If DeemIncUs33AB <> "" And Sheet9.Range("sheet11.DeemIncUs33AB").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs33AB").value = UCase(DeemIncUs33AB)
End If
DeemIncUs33ABA = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs33ABA")
If DeemIncUs33ABA <> "" And Sheet9.Range("sheet11.DeemIncUs33ABA").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs33ABA").value = UCase(DeemIncUs33ABA)
End If
DeemIncUs35ABA = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs35ABA")
If DeemIncUs35ABA <> "" And Sheet9.Range("sheet11.DeemIncUs35ABA").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs35ABA").value = UCase(DeemIncUs35ABA)
End If
DeemIncUs35ABB = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs35ABB")
If DeemIncUs35ABB <> "" And Sheet9.Range("sheet11.DeemIncUs35ABB").Locked = False Then
     Sheet9.Range("sheet11.DeemIncUs35ABB").value = UCase(DeemIncUs35ABB)
End If
DeemIncUs40A3A = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs40A3A")
If DeemIncUs40A3A <> "" And Sheet9.Range("sheet11.DeemIncUs40A").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs40A").value = UCase(DeemIncUs40A3A)
End If
DeemIncUs72A = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs72A")
If DeemIncUs72A <> "" And Sheet9.Range("sheet11.DeemIncUs72A").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs72A").value = UCase(DeemIncUs72A)
End If
DeemIncUs80HHD = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs80HHD")
If DeemIncUs80HHD <> "" And Sheet9.Range("sheet11.DeemIncUs80HHD").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs80HHD").value = UCase(DeemIncUs80HHD)
End If
DeemIncUs80IA = jsonObject("form3CD")("ScheduleBP")("BusinessIncOthThanSpec")("DeemIncUs80IA")
If DeemIncUs80IA <> "" And Sheet9.Range("sheet11.DeemIncUs80IA").Locked = False Then
    Sheet9.Range("sheet11.DeemIncUs80IA").value = UCase(DeemIncUs80IA)
End If
Dim Section44B
Section44B = jsonObject("form3CD")("ScheduleBP")("DeemedProfitBusUs")("Section44B")
If Section44B <> "" And Sheet9.Range("sheet11.Section44B").Locked = False Then
    Sheet9.Range("sheet11.Section44B").value = UCase(Section44B)
End If
Dim Section44BBA
Section44BBA = jsonObject("form3CD")("ScheduleBP")("DeemedProfitBusUs")("Section44BBA")
If Section44BBA <> "" And Sheet9.Range("sheet11.Section44BBA").Locked = False Then
    Sheet9.Range("sheet11.Section44BBA").value = UCase(Section44BBA)
End If
Dim Section44BB
Section44BB = jsonObject("form3CD")("ScheduleBP")("DeemedProfitBusUs")("Section44BB")
If Section44BB <> "" And Sheet9.Range("Sheet11.Section44BB").Locked = False Then
    Sheet9.Range("Sheet11.Section44BB").value = UCase(Section44BB)
End If
Dim Section44DA
Section44DA = jsonObject("form3CD")("ScheduleBP")("DeemedProfitBusUs")("Section44DA")
If Section44DA <> "" And Sheet9.Range("sheet12.Section44DA").Locked = False Then
    Sheet9.Range("sheet12.Section44DA").value = UCase(Section44DA)
End If


End Function
Function ImportPARTA_QD(jsonText As String)
On Error Resume Next
    Dim RecTDS1, TotalExRow, TotalDiffRow, TotalXMLRow As Long
    Dim Nodelist, Node, init, jsonObject As Object
    Dim UnitOfMeasureColNo, OpeningStockColNo, PurchaseQtyColNo, SaleQtyColNo, ClgStockColNo, AnyShortExcesColNo, yldFinisProdColNo, PercentYldColNo, PrevyrManfactColNo, PrevYrConsumColNo As Long
    Dim ItemNameColNo, rowcount, cnt
     Set jsonObject = ParseJson.ParseJson(jsonText)
    
Set init = jsonObject("form3CD")("PartAQD")("TradingConcern")
If init.exists("QuantitDet") Then

    Set Nodelist = jsonObject("form3CD")("PartAQD")("TradingConcern")("QuantitDet")
    
    

    ItemNameColNo = Sheet6.Range("QDTradingConcern.ItemName").Column
    UnitOfMeasureColNo = Sheet6.Range("QDTradingConcern.UnitOfMeasure").Column
    OpeningStockColNo = Sheet6.Range("QDTradingConcern.OpeningStock").Column
    PurchaseQtyColNo = Sheet6.Range("QDTradingConcern.PurchaseQty").Column
    SaleQtyColNo = Sheet6.Range("QDTradingConcern.SaleQty").Column
    ClgStockColNo = Sheet6.Range("QDTradingConcern.ClgStock").Column
    AnyShortExcesColNo = Sheet6.Range("QDTradingConcern.AnyShortExces").Column
    
    TotalExRow = Range("QDTradingConcern.ItemName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("QDTradingConcern.ItemName").Locked = False Then
            Sheet6.Range("QDTradingConcern.ItemName").ClearContents
        End If
        If Sheet6.Range("QDTradingConcern.UnitOfMeasure").Locked = False Then
            Sheet6.Range("QDTradingConcern.UnitOfMeasure").ClearContents
        End If
        If Sheet6.Range("QDTradingConcern.OpeningStock").Locked = False Then
            Sheet6.Range("QDTradingConcern.OpeningStock").ClearContents
        End If
        If Sheet6.Range("QDTradingConcern.PurchaseQty").Locked = False Then
            Sheet6.Range("QDTradingConcern.PurchaseQty").ClearContents
        End If
        If Sheet6.Range("QDTradingConcern.SaleQty").Locked = False Then
            Sheet6.Range("QDTradingConcern.SaleQty").ClearContents
        End If
        If Sheet6.Range("QDTradingConcern.ClgStock").Locked = False Then
            Sheet6.Range("QDTradingConcern.ClgStock").ClearContents
        End If
        If Sheet6.Range("QDTradingConcern.AnyShortExces").Locked = False Then
            Sheet6.Range("QDTradingConcern.AnyShortExces").ClearContents
        End If
    End If
    
'    If (TotalDiffRow > 0) Then
'     AddDiffRows_Item (TotalDiffRow)
'    End If

    rowcount = getRowNo(Sheet6.Range("QDTradingConcern.ItemName").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            If Sheet6.Cells(rowcount, ItemNameColNo).Locked = False Then
                Sheet6.Cells(rowcount, ItemNameColNo).value = Node("ItemName")
            End If
            If Sheet6.Cells(rowcount, UnitOfMeasureColNo).Locked = False Then
                 Dim val
                val = Findtext(CStr(Node("UnitOfMeasure")), "Unit")
                Sheet6.Cells(rowcount, UnitOfMeasureColNo).value = val
            End If
            If Sheet6.Cells(rowcount, OpeningStockColNo).Locked = False Then
                Sheet6.Cells(rowcount, OpeningStockColNo).value = Node("OpeningStock")
            End If
            If Sheet6.Cells(rowcount, PurchaseQtyColNo).Locked = False Then
                Sheet6.Cells(rowcount, PurchaseQtyColNo).value = Node("PurchaseQty")
            End If
            If Sheet6.Cells(rowcount, SaleQtyColNo).Locked = False Then
                Sheet6.Cells(rowcount, SaleQtyColNo).value = Node("SaleQty")
            End If
            If Sheet6.Cells(rowcount, ClgStockColNo).Locked = False Then
                Sheet6.Cells(rowcount, ClgStockColNo).value = Node("ClgStock")
            End If
            If Sheet6.Cells(rowcount, AnyShortExcesColNo).Locked = False Then
                Sheet6.Cells(rowcount, AnyShortExcesColNo).value = Node("AnyShortExces")
            End If
            
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If

Set init = jsonObject("form3CD")("PartAQD")("ManfactrConcern")("RawMaterial")
If init.exists("QuantitDet") Then

    Set Nodelist = jsonObject("form3CD")("PartAQD")("ManfactrConcern")("RawMaterial")("QuantitDet")
      
    ItemNameColNo = Sheet6.Range("QDRawMaterial.ItemName").Column
    UnitOfMeasureColNo = Sheet6.Range("QDRawMaterial.UnitOfMeasure").Column
    OpeningStockColNo = Sheet6.Range("QDRawMaterial.OpeningStock").Column
    PurchaseQtyColNo = Sheet6.Range("QDRawMaterial.PurchaseQty").Column
    PrevYrConsumColNo = Sheet6.Range("QDRawMaterial.PrevYrConsum").Column
    SaleQtyColNo = Sheet6.Range("QDRawMaterial.SaleQty").Column
    ClgStockColNo = Sheet6.Range("QDRawMaterial.ClgStock").Column
    yldFinisProdColNo = Sheet6.Range("QDRawMaterial.yldFinisProd").Column
    PercentYldColNo = Sheet6.Range("QDRawMaterial.PercentYld").Column
    AnyShortExcesColNo = Sheet6.Range("QDRawMaterial.AnyShortExces").Column
           
    TotalExRow = Range("QDRawMaterial.ItemName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("QDRawMaterial.ItemName").Locked = False Then
            Sheet6.Range("QDRawMaterial.ItemName").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.UnitOfMeasure").Locked = False Then
            Sheet6.Range("QDRawMaterial.UnitOfMeasure").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.OpeningStock").Locked = False Then
            Sheet6.Range("QDRawMaterial.OpeningStock").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.PurchaseQty").Locked = False Then
            Sheet6.Range("QDRawMaterial.PurchaseQty").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.PrevYrConsum").Locked = False Then
            Sheet6.Range("QDRawMaterial.PrevYrConsum").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.SaleQty").Locked = False Then
            Sheet6.Range("QDRawMaterial.SaleQty").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.ClgStock").Locked = False Then
            Sheet6.Range("QDRawMaterial.ClgStock").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.yldFinisProd").Locked = False Then
            Sheet6.Range("QDRawMaterial.yldFinisProd").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.PercentYld").Locked = False Then
            Sheet6.Range("QDRawMaterial.PercentYld").ClearContents
        End If
        If Sheet6.Range("QDRawMaterial.AnyShortExces").Locked = False Then
            Sheet6.Range("QDRawMaterial.AnyShortExces").ClearContents
        End If
    End If
    
'    If (TotalDiffRow > 0) Then
'       AddDiffRows_Item1 (TotalDiffRow)
'    End If

    rowcount = getRowNo(Sheet6.Range("QDRawMaterial.ItemName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
            If Sheet6.Cells(rowcount, ItemNameColNo).Locked = False Then
                Sheet6.Cells(rowcount, ItemNameColNo).value = UCase(Node("ItemName"))
            End If
            If Sheet6.Cells(rowcount, UnitOfMeasureColNo).Locked = False Then
                 Dim valQDRawMaterial
                valQDRawMaterial = Findtext(CStr(Node("UnitOfMeasure")), "Unit")
                Sheet6.Cells(rowcount, UnitOfMeasureColNo).value = valQDRawMaterial
            End If
            If Sheet6.Cells(rowcount, OpeningStockColNo).Locked = False Then
                Sheet6.Cells(rowcount, OpeningStockColNo).value = Node("OpeningStock")
            End If
            If Sheet6.Cells(rowcount, PurchaseQtyColNo).Locked = False Then
                Sheet6.Cells(rowcount, PurchaseQtyColNo).value = Node("PurchaseQty")
            End If
            If Sheet6.Cells(rowcount, PrevYrConsumColNo).Locked = False Then
                Sheet6.Cells(rowcount, PrevYrConsumColNo).value = Node("PrevYrConsum")
            End If
            If Sheet6.Cells(rowcount, SaleQtyColNo).Locked = False Then
                Sheet6.Cells(rowcount, SaleQtyColNo).value = Node("SaleQty")
            End If
            If Sheet6.Cells(rowcount, ClgStockColNo).Locked = False Then
                Sheet6.Cells(rowcount, ClgStockColNo).value = Node("ClgStock")
            End If
            If Sheet6.Cells(rowcount, yldFinisProdColNo).Locked = False Then
'Change-3, ANK-130, 26.05.2022
                Sheet6.Cells(rowcount, yldFinisProdColNo).value = Node("YldFinisProd")
'                Sheet6.Cells(rowcount, yldFinisProdColNo).value = node("yldFinisProd")
'---end change---
            End If
            If Sheet6.Cells(rowcount, PercentYldColNo).Locked = False Then
                Sheet6.Cells(rowcount, PercentYldColNo).value = Node("PercentYld")
            End If
            If Sheet6.Cells(rowcount, AnyShortExcesColNo).Locked = False Then
                Sheet6.Cells(rowcount, AnyShortExcesColNo).value = Node("AnyShortExces")
            End If
            
            
       cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If

Set init = jsonObject("form3CD")("PartAQD")("ManfactrConcern")("FinishrByProd")
If init.exists("QuantitDet") Then

    Set Nodelist = jsonObject("form3CD")("PartAQD")("ManfactrConcern")("FinishrByProd")("QuantitDet")
    
    ItemNameColNo = Sheet6.Range("QDFinishrByProd.ItemName").Column
    UnitOfMeasureColNo = Sheet6.Range("QDFinishrByProd.UnitOfMeasure").Column
    OpeningStockColNo = Sheet6.Range("QDFinishrByProd.OpeningStock").Column
    PurchaseQtyColNo = Sheet6.Range("QDFinishrByProd.PurchaseQty").Column
    PrevyrManfactColNo = Sheet6.Range("QDFinishrByProd.PrevYrConsum").Column
    SaleQtyColNo = Sheet6.Range("QDFinishrByProd.SaleQty").Column
    ClgStockColNo = Sheet6.Range("QDFinishrByProd.ClgStock").Column
    AnyShortExcesColNo = Sheet6.Range("QDFinishrByProd.AnyShortExces").Column
    
      
    TotalExRow = Range("QDFinishrByProd.ItemName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet6.Range("QDFinishrByProd.ItemName").Locked = False Then
            Sheet6.Range("QDFinishrByProd.ItemName").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.UnitOfMeasure").Locked = False Then
            Sheet6.Range("QDFinishrByProd.UnitOfMeasure").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.OpeningStock").Locked = False Then
            Sheet6.Range("QDFinishrByProd.OpeningStock").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.PurchaseQty").Locked = False Then
            Sheet6.Range("QDFinishrByProd.PurchaseQty").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.PrevYrConsum").Locked = False Then
            Sheet6.Range("QDFinishrByProd.PrevYrConsum").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.SaleQty").Locked = False Then
            Sheet6.Range("QDFinishrByProd.SaleQty").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.ClgStock").Locked = False Then
            Sheet6.Range("QDFinishrByProd.ClgStock").ClearContents
        End If
        If Sheet6.Range("QDFinishrByProd.AnyShortExces").Locked = False Then
            Sheet6.Range("QDFinishrByProd.AnyShortExces").ClearContents
        End If
    End If
    
'    If (TotalDiffRow > 0) Then
'      AddDiffRows_Item2 (TotalDiffRow)
'    End If

    rowcount = getRowNo(Sheet6.Range("QDFinishrByProd.ItemName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
            If Sheet6.Cells(rowcount, ItemNameColNo).Locked = False Then
                Sheet6.Cells(rowcount, ItemNameColNo).value = UCase(Node("ItemName"))
            End If
            If Sheet6.Cells(rowcount, UnitOfMeasureColNo).Locked = False Then
                 Dim valQDFinishrByProd
                valQDFinishrByProd = Findtext(CStr(Node("UnitOfMeasure")), "Unit")
                Sheet6.Cells(rowcount, UnitOfMeasureColNo).value = valQDFinishrByProd
            End If
            If Sheet6.Cells(rowcount, OpeningStockColNo).Locked = False Then
                Sheet6.Cells(rowcount, OpeningStockColNo).value = Node("OpeningStock")
            End If
            If Sheet6.Cells(rowcount, PurchaseQtyColNo).Locked = False Then
                Sheet6.Cells(rowcount, PurchaseQtyColNo).value = Node("PurchaseQty")
            End If
            If Sheet6.Cells(rowcount, PrevYrConsumColNo).Locked = False Then
                Sheet6.Cells(rowcount, PrevYrConsumColNo).value = Node("PrevyrManfact")
            End If
            If Sheet6.Cells(rowcount, SaleQtyColNo).Locked = False Then
                Sheet6.Cells(rowcount, SaleQtyColNo).value = Node("SaleQty")
            End If
            If Sheet6.Cells(rowcount, ClgStockColNo).Locked = False Then
                Sheet6.Cells(rowcount, ClgStockColNo).value = Node("ClgStock")
            End If
            If Sheet6.Cells(rowcount, AnyShortExcesColNo).Locked = False Then
                Sheet6.Cells(rowcount, AnyShortExcesColNo).value = Node("AnyShortExces")
            End If
            
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If
End Function

Function ImportPARTA_OI(jsonText As String)
On Error Resume Next
Dim jsonObject
Dim MethodOfAcct
Dim ChangeInAcctMethFlg, EffectOnPL, DecProOrIncLossUs145_A, ValRawMaterial, ValFinishedGoods, ChngStockValMetFlg As Variant
    Set jsonObject = ParseJson.ParseJson(jsonText)

'issue method of acct
MethodOfAcct = jsonObject("lastFiledITR")("partAOI")("MethodOfAcct")
If MethodOfAcct = "CASH" Then
    Sheet5.Range("Sheet5.MethodOfAcct") = "Cash"
Else
    Sheet5.Range("Sheet5.MethodOfAcct") = "Mercantile"
End If
'end



ChangeInAcctMethFlg = jsonObject("lastFiledITR")("partAOI")("ChangeInAcctMethFlg")
If ChangeInAcctMethFlg = "Y" Then
    ChangeInAcctMethFlg = "Yes"
ElseIf ChangeInAcctMethFlg = "N" Then
    ChangeInAcctMethFlg = "No"
Else
    ChangeInAcctMethFlg = "(Select)"
End If

Sheet5.Range("sheet5.ChangeInAcctMethFlg").value = ChangeInAcctMethFlg

ValRawMaterial = jsonObject("lastFiledITR")("partAOI")("MethodOfValClgStk")("ValRawMaterial")
If ValRawMaterial = "1" Then
    ValRawMaterial = "1 - Cost or market rate , whichever is less"
ElseIf ValRawMaterial = "2" Then
    ValRawMaterial = "2 - At cost"
ElseIf ValRawMaterial = "3" Then
    ValRawMaterial = "3 - At Market rate"
Else
    ValRawMaterial = "(Select)"
End If

Sheet5.Range("sheet5.ValRawMaterial") = ValRawMaterial

ValFinishedGoods = jsonObject("lastFiledITR")("partAOI")("MethodOfValClgStk")("ValFinishedGoods")
If ValFinishedGoods = "1" Then
    ValFinishedGoods = "1 - Cost or market rate , whichever is less"
ElseIf ValFinishedGoods = "2" Then
    ValFinishedGoods = "2 - At cost"
ElseIf ValFinishedGoods = "3" Then
    ValFinishedGoods = "3 - At Market rate"
Else
    ValFinishedGoods = "(Select)"
End If

Sheet5.Range("sheet5.ValFinishedGoods") = ValFinishedGoods

ChngStockValMetFlg = jsonObject("form3CD")("PartAOI")("MethodOfValClgStk")("ChngStockValMetFlg")
If ChngStockValMetFlg = "Y" Then
    ChngStockValMetFlg = "Yes"
ElseIf ChngStockValMetFlg = "N" Then
    ChngStockValMetFlg = "No"
Else
    ChngStockValMetFlg = "(Select)"
End If


        
        If Sheet5.Range("sheet6.ChngStockValMetFlg").Locked = False Then
        Sheet5.Activate
            Sheet5.Range("sheet6.ChngStockValMetFlg").value = ChngStockValMetFlg
        End If
        
EffectOnPL = jsonObject("form3CD")("PartAOI")("MethodOfValClgStk")("EffectOnPL")
        If Sheet5.Range("sheet6.EffectOnPL").Locked = False Then
            Sheet5.Range("sheet6.EffectOnPL").value = EffectOnPL
        End If
        
DecProOrIncLossUs145_A = jsonObject("form3CD")("PartAOI")("MethodOfValClgStk")("DecProOrIncLossUs145A")
        If Sheet5.Range("sheet6.EffectOnPL4e").Locked = False Then
            Sheet5.Range("sheet6.EffectOnPL4e").value = DecProOrIncLossUs145_A
        End If
        
        If Sheet5.Range("sheet6.Section28Items").Locked = False Then
            Sheet5.Range("sheet6.Section28Items").value = jsonObject("form3CD")("PartAOI")("NoCredToPLAmt")("Section28Items")
        End If
        If Sheet5.Range("sheet6.ProformaCreditsDue").Locked = False Then
            Sheet5.Range("sheet6.ProformaCreditsDue").value = jsonObject("form3CD")("PartAOI")("NoCredToPLAmt")("ProformaCreditsDue")
        End If
        If Sheet5.Range("sheet6.PrevYrEscalClaim").Locked = False Then
            Sheet5.Range("sheet6.PrevYrEscalClaim").value = jsonObject("form3CD")("PartAOI")("NoCredToPLAmt")("PrevYrEscalClaim")
        End If
        If Sheet5.Range("sheet6.OthItemInc").Locked = False Then
            Sheet5.Range("sheet6.OthItemInc").value = jsonObject("form3CD")("PartAOI")("NoCredToPLAmt")("OthItemInc")
        End If
        If Sheet5.Range("sheet6.CapReceipt").Locked = False Then
            Sheet5.Range("sheet6.CapReceipt").value = jsonObject("form3CD")("PartAOI")("NoCredToPLAmt")("CapReceipt")
        End If
        
''        If Sheet5.Range("sheet6.StkInsurPrem").Locked = False Then
''            Sheet5.Range("sheet6.StkInsurPrem").value = jsonObject("PARTA_OI")("AmtDisallUs36")("StkInsurPrem")
''        End If
''        If Sheet5.Range("sheet6.EmpHealthInsurPrem").Locked = False Then
''            Sheet5.Range("sheet6.EmpHealthInsurPrem").value = jsonObject("PARTA_OI")("AmtDisallUs36")("EmpHealthInsurPrem")
''        End If
        If Sheet5.Range("sheet6.EmpBonusCommSum").Locked = False Then
            Sheet5.Range("sheet6.EmpBonusCommSum").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs36")("EmpBonusCommSum")
        End If
'        If Sheet5.Range("sheet6.IntOnBorrCap").Locked = False Then
'            Sheet5.Range("sheet6.IntOnBorrCap").value = jsonObject("PARTA_OI")("AmtDisallUs36")("IntOnBorrCap")
'        End If
'        If Sheet5.Range("sheet6.ZeroCoupBondDisc").Locked = False Then
'            Sheet5.Range("sheet6.ZeroCoupBondDisc").value = jsonObject("PARTA_OI")("AmtDisallUs36")("ZeroCoupBondDisc")
'        End If
'        If Sheet5.Range("sheet6.RecogPFContribAmt").Locked = False Then
'            Sheet5.Range("sheet6.RecogPFContribAmt").value = jsonObject("PARTA_OI")("AmtDisallUs36")("RecogPFContribAmt")
'        End If
'        If Sheet5.Range("sheet6.AppSuperAnnFundAmt").Locked = False Then
'            Sheet5.Range("sheet6.AppSuperAnnFundAmt").value = jsonObject("PARTA_OI")("AmtDisallUs36")("AppSuperAnnFundAmt")
'        End If
'        If Sheet5.Range("Sheet6.PensionSchemeSec80CCD").Locked = False Then
'            Sheet5.Range("Sheet6.PensionSchemeSec80CCD").value = jsonObject("PARTA_OI")("AmtDisallUs36")("PensionSchemeSec80CCD")
'        End If
'        If Sheet5.Range("sheet6.AppGratFundAmt").Locked = False Then
'            Sheet5.Range("sheet6.AppGratFundAmt").value = jsonObject("PARTA_OI")("AmtDisallUs36")("AppGratFundAmt")
'        End If
'        If Sheet5.Range("sheet6.OthFundAmt").Locked = False Then
'            Sheet5.Range("sheet6.OthFundAmt").value = jsonObject("PARTA_OI")("AmtDisallUs36")("OthFundAmt")
'        End If
        If Sheet5.Range("sheet6.EmpContributionCredits").Locked = False Then
            Sheet5.Range("sheet6.EmpContributionCredits").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs36")("EmpContributionCredits")
        End If
'        If Sheet5.Range("sheet6.BadDebtDoubtAmt").Locked = False Then
'            Sheet5.Range("sheet6.BadDebtDoubtAmt").value = jsonObject("PARTA_OI")("AmtDisallUs36")("BadDebtDoubtAmt")
'        End If
'        If Sheet5.Range("sheet6.BadDebtDoubtProvn").Locked = False Then
'            Sheet5.Range("sheet6.BadDebtDoubtProvn").value = jsonObject("PARTA_OI")("AmtDisallUs36")("BadDebtDoubtProvn")
'        End If
'        If Sheet5.Range("sheet6.SpecResrvTranfr").Locked = False Then
'            Sheet5.Range("sheet6.SpecResrvTranfr").value = jsonObject("PARTA_OI")("AmtDisallUs36")("SpecResrvTranfr")
'        End If
'        If Sheet5.Range("sheet6.FamPlanPromoExp").Locked = False Then
'            Sheet5.Range("sheet6.FamPlanPromoExp").value = jsonObject("PARTA_OI")("AmtDisallUs36")("FamPlanPromoExp")
'        End If
'        If Sheet5.Range("sheet6.SecuritiesPaidAmt").Locked = False Then
'            Sheet5.Range("sheet6.SecuritiesPaidAmt").value = jsonObject("PARTA_OI")("AmtDisallUs36")("SecuritiesPaidAmt")
'        End If
'        If Sheet5.Range("sheet6.MarketLoss").Locked = False Then
'            Sheet5.Range("sheet6.MarketLoss").value = jsonObject("PARTA_OI")("AmtDisallUs36")("MrktLossOthExpLossICDS")
'        End If
'        If Sheet5.Range("sheet6.OthDisallowances").Locked = False Then
'            Sheet5.Range("sheet6.OthDisallowances").value = jsonObject("PARTA_OI")("AmtDisallUs36")("OthDisallowances")
'        End If
        
        If Sheet5.Range("sheet6.CapitalNatureExp").Locked = False Then
            Sheet5.Range("sheet6.CapitalNatureExp").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("CapitalNatureExp")
        End If
        If Sheet5.Range("sheet6.PersonalExp").Locked = False Then
            Sheet5.Range("sheet6.PersonalExp").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("PersonalExp")
        End If
'        If Sheet5.Range("sheet6.BusOrProfessnExp").Locked = False Then
'            Sheet5.Range("sheet6.BusOrProfessnExp").value = jsonObject("PARTA_OI")("AmtDisallUs37")("BusOrProfessnExp")
'        End If
        If Sheet5.Range("sheet6.PoliticPartyExp").Locked = False Then
            Sheet5.Range("sheet6.PoliticPartyExp").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("PoliticPartyExp")
        End If
        If Sheet5.Range("sheet6.LawVoilatPenalExp").Locked = False Then
            Sheet5.Range("sheet6.LawVoilatPenalExp").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("LawVoilatPenalExp")
        End If
        If Sheet5.Range("sheet6.OthPenalFineExp").Locked = False Then
            Sheet5.Range("sheet6.OthPenalFineExp").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("OthPenalFineExp")
        End If
        If Sheet5.Range("sheet6.OffenceExp").Locked = False Then
            Sheet5.Range("sheet6.OffenceExp").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("OffenceExp")
        End If
        If Sheet5.Range("sheet6.ContigentLiability").Locked = False Then
            Sheet5.Range("sheet6.ContigentLiability").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs37")("ContigentLiability")
        End If
'        If Sheet5.Range("sheet6.OthAmtNotAllowUs37").Locked = False Then
'            Sheet5.Range("sheet6.OthAmtNotAllowUs37").value = jsonObject("PARTA_OI")("AmtDisallUs37")("OthAmtNotAllowUs37")
'        End If
        If Sheet5.Range("sheet6.NonCompChapXVIIBAmt").Locked = False Then
            Sheet5.Range("sheet6.NonCompChapXVIIBAmt").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40")("NonCompChapXVIIBAmt")
        End If
'        If Sheet5.Range("shee6.NonComp40aiiChapXVIIBAmt").Locked = False Then
'            Sheet5.Range("shee6.NonComp40aiiChapXVIIBAmt").value = jsonObject("PARTA_OI")("AmtDisallUs40")("NonComp40aiiChapXVIIBAmt")
'        End If
'Change-7, Test-044, Remark - Two typos corrected, following code is commented
         If Sheet5.Range("sheet6.NonComp40aibChapXVIIBAmt").Locked = False Then
            Sheet5.Range("sheet6.NonComp40aibChapXVIIBAmt").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40")("NonComp40AibChapXVIIBAmt")
        End If
'        If Sheet5.Range("sheet6.NonComp40aibChapXVIIBAmt").Locked = False Then
'            Sheet5.Range("sheet6.NonComp40aibChapXVIIBAmt").value = jsonObject("form3CD")("PARTAOI")("AmtDisallUs40")("NonComp40aibChapXVIIBAmt")
'        End If
        If Sheet5.Range("sheet6.NonComp40aiiiChapXVIIBAmt").Locked = False Then
            Sheet5.Range("sheet6.NonComp40aiiiChapXVIIBAmt").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40")("NonComp40AiiiChapXVIIBAmt")
        End If
'        If Sheet5.Range("sheet6.TaxAmtOnProfits").Locked = False Then
'            Sheet5.Range("sheet6.TaxAmtOnProfits").value = jsonObject("PARTA_OI")("AmtDisallUs40")("TaxAmtOnProfits")
'        End If
        If Sheet5.Range("sheet7.WTAmt").Locked = False Then
            Sheet5.Range("sheet7.WTAmt").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40")("Wtamt")
        End If
        If Sheet5.Range("sheet7.RolyatyOrServiceFee").Locked = False Then
            Sheet5.Range("sheet7.RolyatyOrServiceFee").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40")("RolyatyOrServiceFee")
        End If
        If Sheet5.Range("sheet7.IntSalBonPartner").Locked = False Then
            Sheet5.Range("sheet7.IntSalBonPartner").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40")("IntSalBonPartner")
        End If
'        If Sheet5.Range("sheet7.OthDisallow").Locked = False Then
'            Sheet5.Range("sheet7.OthDisallow").value = jsonObject("PARTA_OI")("AmtDisallUs40")("OthDisallow")
'        End If
'        If Sheet5.Range("sheet7.AmtDisallUs40PyNowAll").Locked = False Then
'            Sheet5.Range("sheet7.AmtDisallUs40PyNowAll").value = jsonObject("PARTA_OI")("AmtDisallUs40")("AmtDisallUs40PyNowAll")
'        End If
'        If Sheet5.Range("sheet7.AmtPaidUs40A2b").Locked = False Then
'            Sheet5.Range("sheet7.AmtPaidUs40A2b").value = jsonObject("PARTA_OI")("AmtDisallUs40A")("AmtPaidUs40A2b")
'        End If
        If Sheet5.Range("sheet7.AmtGT20kCash").Locked = False Then
            Sheet5.Range("sheet7.AmtGT20kCash").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40A")("AmtPaidUs40A3")
        End If
        If Sheet5.Range("sheet7.ProvPmtGrat").Locked = False Then
            Sheet5.Range("sheet7.ProvPmtGrat").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40A")("ProvPmtGrat")
        End If
        If Sheet5.Range("sheet7.ContToSetupTrust").Locked = False Then
            Sheet5.Range("sheet7.ContToSetupTrust").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs40A")("ContToSetupTrust")
        End If
'        If Sheet5.Range("sheet7.OthDisallow1").Locked = False Then
'            Sheet5.Range("sheet7.OthDisallow1").value = jsonObject("PARTA_OI")("AmtDisallUs40A")("OthDisallow")
'        End If
        If Sheet5.Range("sheet7.TaxDutyCesAmt").Locked = False Then
            Sheet5.Range("sheet7.TaxDutyCesAmt").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("TaxDutyCesAmt")
        End If
        If Sheet5.Range("sheet7.ContToEmpPFSFGF").Locked = False Then
            Sheet5.Range("sheet7.ContToEmpPFSFGF").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("ContToEmpPFSFGF")
        End If
        If Sheet5.Range("sheet7.EmpBonusComm").Locked = False Then
            Sheet5.Range("sheet7.EmpBonusComm").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("EmpBonusComm")
        End If
        If Sheet5.Range("sheet7.IntPayaleToFI").Locked = False Then
            Sheet5.Range("sheet7.IntPayaleToFI").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("IntPayaleToFI")
        End If
        If Sheet5.Range("sheet7.IntPayaleToFISchBank").Locked = False Then
            Sheet5.Range("sheet7.IntPayaleToFISchBank").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("IntPayaleToFISchBank")
        End If
        If Sheet5.Range("sheet7.LeaveEncashPayable").Locked = False Then
            Sheet5.Range("sheet7.LeaveEncashPayable").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("LeaveEncashPayable")
        End If
        If Sheet5.Range("sheet7.RailwayAssets").Locked = False Then
            Sheet5.Range("sheet7.RailwayAssets").value = jsonObject("form3CD")("PartAOI")("AmtDisallUs43BPyNowAll")("AmtUs43B")("RailwayAssetsPayable")
        End If
        
        If Sheet5.Range("sheet7.TaxDutyCesAmt1").Locked = False Then
            Sheet5.Range("sheet7.TaxDutyCesAmt1").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("TaxDutyCesAmt")
        End If
        If Sheet5.Range("sheet7.ContToEmpPFSFGF1").Locked = False Then
            Sheet5.Range("sheet7.ContToEmpPFSFGF1").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("ContToEmpPFSFGF")
        End If
        If Sheet5.Range("sheet7.EmpBonusComm1").Locked = False Then
            Sheet5.Range("sheet7.EmpBonusComm1").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("EmpBonusComm")
        End If
        If Sheet5.Range("sheet7.IntPayaleToFI1").Locked = False Then
            Sheet5.Range("sheet7.IntPayaleToFI1").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("IntPayaleToFI")
        End If
        If Sheet5.Range("sheet7.IntPayaleToFI1_11da").Locked = False Then
            Sheet5.Range("sheet7.IntPayaleToFI1_11da").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("IntPayaleToFINBFC")
        End If
        If Sheet5.Range("sheet7.IntPayaleToFISchBank1").Locked = False Then
            Sheet5.Range("sheet7.IntPayaleToFISchBank1").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("IntPayaleToFISchBank")
        End If
        If Sheet5.Range("sheet7.LeaveEncashPayable1").Locked = False Then
            Sheet5.Range("sheet7.LeaveEncashPayable1").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("LeaveEncashPayable")
        End If
        If Sheet5.Range("sheet7.RailwayAssetsPayable").Locked = False Then
            Sheet5.Range("sheet7.RailwayAssetsPayable").value = jsonObject("form3CD")("PartAOI")("AmtDisall43B")("AmtUs43B")("RailwayAssetsPayable")
        End If
'        If Sheet5.Range("sheet7.UnionExciseDuty").Locked = False Then
'            Sheet5.Range("sheet7.UnionExciseDuty").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("UnionExciseDuty")
'        End If
'        If Sheet5.Range("sheet7.ServiceTax").Locked = False Then
'            Sheet5.Range("sheet7.ServiceTax").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("ServiceTax")
'        End If
'        If Sheet5.Range("sheet7.VATorSaleTax").Locked = False Then
'            Sheet5.Range("sheet7.VATorSaleTax").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("VATorSaleTax")
'        End If
'        If Sheet5.Range("sheet7.OthDutyTaxCess").Locked = False Then
'            Sheet5.Range("sheet7.OthDutyTaxCess").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("Cess")
'        End If
'        If Sheet5.Range("Sheet7.CGST").Locked = False Then
'            Sheet5.Range("Sheet7.CGST").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("CentralGoodServiceTax")
'        End If
'        If Sheet5.Range("Sheet7.SGST").Locked = False Then
'            Sheet5.Range("Sheet7.SGST").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("StateGoodServiceTax")
'        End If
'        If Sheet5.Range("Sheet7.IGST").Locked = False Then
'            Sheet5.Range("Sheet7.IGST").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("IntegratedGoodServiceTax")
'        End If
'        If Sheet5.Range("Sheet7.UTGST").Locked = False Then
'            Sheet5.Range("Sheet7.UTGST").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("UnionTerrGoodServiceTax")
'        End If
'        If Sheet5.Range("sheet7.OthDutyTaxCess").Locked = False Then
'            Sheet5.Range("sheet7.OthDutyTaxCess").value = jsonObject("PARTA_OI")("AmtExciseCustomsVATOutstanding")("ExciseCustomsVAT")("OthDutyTaxCess")
'        End If
        If Sheet5.Range("sheet7.DeemedProfUs33AB").Locked = False Then
            Sheet5.Range("sheet7.DeemedProfUs33AB").value = jsonObject("form3CD")("PartAOI")("DeemedProfUs33AB")
        End If
        If Sheet5.Range("sheet7.DeemedProfUs33ABA").Locked = False Then
            Sheet5.Range("sheet7.DeemedProfUs33ABA").value = jsonObject("form3CD")("PartAOI")("DeemedProfUs33ABA")
        End If
        If Sheet5.Range("sheet7.ProfTaxAmtUs41").Locked = False Then
            Sheet5.Range("sheet7.ProfTaxAmtUs41").value = jsonObject("form3CD")("PartAOI")("ProfTaxAmtUs41")
        End If
        If Sheet5.Range("sheet7.PriorAmtIncCrDrPL").Locked = False Then
            Sheet5.Range("sheet7.PriorAmtIncCrDrPL").value = jsonObject("form3CD")("PartAOI")("PriorAmtIncCrDrPL")
        End If
        If Sheet5.Range("sheet7.AmountOfExpenditure14A").Locked = False Then
            Sheet5.Range("sheet7.AmountOfExpenditure14A").value = jsonObject("form3CD")("PartAOI")("AmountOfExpDisAllwUs14A")
        End If
'Change-2, ANK, 01.06.2022
        If Sheet5.Range("sheet7.Section92CE_Flag").Locked = False Then
            Sheet5.Range("sheet7.Section92CE_Flag").value = jsonObject("form3CD")("PartAOI")("Exercise92CE2AFlg")
        End If
'---

End Function
Function ImportManufacturingAccount(jsonText As String)
On Error Resume Next

    Dim jsonObject, OpenStockRawMaterial, OpenStockWorkProgress
    
    Set jsonObject = ParseJson.ParseJson(jsonText)
    If jsonObject("lastFiledITR").exists("manufacturingAccount") Then
        OpenStockRawMaterial = jsonObject("lastFiledITR")("manufacturingAccount")("openingInventory")("OpngStckRawMat")
        If Sheet48.Range("ManuFactureAcc_OpenStockRawMaterial").Locked = False And OpenStockRawMaterial <> "" Then
            Sheet48.Range("ManuFactureAcc_OpenStockRawMaterial").value = OpenStockRawMaterial
        End If
        OpenStockWorkProgress = jsonObject("lastFiledITR")("manufacturingAccount")("openingInventory")("OpngStckWrkinPrgrs")
        If Sheet48.Range("ManuFactureAcc_OpenStockWorkProgress").Locked = False And OpenStockWorkProgress <> "" Then
            Sheet48.Range("ManuFactureAcc_OpenStockWorkProgress").value = OpenStockWorkProgress
        End If
       
    End If
    
End Function
Function ImportTradingAccount(jsonText As String)
On Error Resume Next

    Dim TotalXMLRow, TotalExRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    Dim rowcount, cnt As Long
    
    Dim Nodelist, Node, init, jsonObject As Object
    Set jsonObject = ParseJson.ParseJson(jsonText)
    If Sheet49.Range("TradingAcc_SalesOfGoods").Locked = False Then
        Sheet49.Range("TradingAcc_SalesOfGoods").value = jsonObject("form3CD")("TradingAccount")("SaleOfGoods")
    End If
    
    If Sheet49.Range("TradingAcc_ClosingStockOfFinishedStocks").Locked = False Then
        Sheet49.Range("TradingAcc_ClosingStockOfFinishedStocks").value = jsonObject("form3CD")("TradingAccount")("ClsngStckOfFinishedStcks")
    End If
    
    
    If Sheet49.Range("TradingAcc_OpeningStockOfFinishedGoods").Locked = False Then
        Sheet49.Range("TradingAcc_OpeningStockOfFinishedGoods").value = jsonObject("form3CD")("TradingAccount")("OpngStckOfFinishedStcks")
    End If
    
    
End Function


Function ImportScheduleTCS(jsonText As String)
On Error Resume Next
Dim jsonDictionary, jsonObject As Object
Dim Node, Nodelist, init As Object

Dim TANColNo, DEDNameColNo, DeductedYearColNo, BroughtFwdTDSAmtColNo, TotalTCSColNo, ClaimColNo, AmtCarriedFwdColNo, TotalSchTCS As Variant
Dim TotalXMLRow, rowcount, cnt, RecTCS  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Set jsonObject = ParseJson.ParseJson(jsonText)

TotalSchTCS = jsonObject("form26as")("scheduleTCS")("totalSchTCS")
  

Set init = jsonObject("form26as")("scheduleTCS")
If init.exists("tcs") Then
  
Set Nodelist = jsonObject("form26as")("ScheduleTCS")("tcs")

    TANColNo = Sheet28.Range("TCS.TAN").Column
    DEDNameColNo = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Column
'    DeductedYearColNo = Sheet28.Range("TCS.DeductedYear").Column
'    BroughtFwdTDSAmtColNo = Sheet28.Range("TCS.BroughtFwdTDSAmt").Column
    TotalTCSColNo = Sheet28.Range("TCS.TotalTCS").Column
    ClaimColNo = Sheet28.Range("TCS.AmtTCSClaimedThisYear").Column
    AmtCarriedFwdColNo = Sheet28.Range("TCS.AmtCarriedFwd").Column
    
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet28.Range("TCS.TAN").Locked = False Then
            Sheet28.Range("TCS.TAN").ClearContents
        End If
        If Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Locked = False Then
            Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
        End If
'        If Sheet28.Range("TCS.DeductedYear").Locked = False Then
'            Sheet28.Range("TCS.DeductedYear").ClearContents
'        End If
'        If Sheet28.Range("TCS.BroughtFwdTDSAmt").Locked = False Then
'            Sheet28.Range("TCS.BroughtFwdTDSAmt").ClearContents
'        End If
        If Sheet28.Range("TCS.TotalTCS").Locked = False Then
            Sheet28.Range("TCS.TotalTCS").ClearContents
        End If
        If Sheet28.Range("TCS.AmtTCSClaimedThisYear").Locked = False Then
            Sheet28.Range("TCS.AmtTCSClaimedThisYear").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
'     AddDiffRows_TCS (TotalDiffRow)
     AddDiffRows_ScheduleTCS (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
       
        If Sheet28.Cells(rowcount, TANColNo).Locked = False Then
            Sheet28.Cells(rowcount, TANColNo).value = Node("EmployerOrDeductorOrCollectDetl")("tan")
        End If
        If Sheet28.Cells(rowcount, DEDNameColNo).Locked = False Then
            Sheet28.Cells(rowcount, DEDNameColNo).value = Node("EmployerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
        End If
'        If Sheet28.Cells(rowcount, DeductedYearColNo).Locked = False Then
'            Sheet28.Cells(rowcount, DeductedYearColNo).value = Node("deductedYr")
'        End If
        If Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).Locked = False And Node("broughtFwdTDSAmt") <> 0 Then
            Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).value = Node("broughtFwdTDSAmt")
        End If
        If Sheet28.Cells(rowcount, TotalTCSColNo).Locked = False And Node("totalTCS") <> 0 Then
            Sheet28.Cells(rowcount, TotalTCSColNo).value = Node("deductedYear")
        End If
        If Sheet28.Cells(rowcount, ClaimColNo).Locked = False Then
            Sheet28.Cells(rowcount, ClaimColNo).value = Node("amtTaxCollected")
        End If
              
    cnt = cnt + 1
    Next Node
RecTCS = cnt
End If
If TotalSchTCS <> "" Then
    Sheet28.Range("TCS.Sum").value = TotalSchTCS
End If

End Function
Function ScheduleTDS3Import(jsonText As String)
On Error Resume Next

    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, Node, init, jsonObject As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow, rowcount, cnt As Long
    Dim TDSCreditName_TDS3, HeadIncome_TDS3
    Dim TDSCreditName, PanOtherPerson, Aadhar_TDS3, PAN, Aadhaar, DeductedYear, BroughtFwdTDSAmt, SectionTDS_3_ColNo As Variant
    Dim DeductedInOwnHands, DeductedIncome5A, DeductedTDS5A, ClaimedInOwnHands, ClaimedIncome5A, ClaimedTDS5A
    Dim ClaimedPAN5A, ClaimedAadhar5A_TDS3, GrossAmount, HeadOfIncome
    Dim Nodelist2
    
Set jsonObject = ParseJson.ParseJson(jsonText)
'Change-6(b), Test-033, Remark - following "Set" and "If" statement uncommented, along with "End If"
'one more exists condition added for "tds3Details"
'form26as.scheduleTDS3Dtls.tds3Details.TDSCreditName

Set init = jsonObject("form26as")
If init.exists("scheduleTDS3Dtls") Then
If init("scheduleTDS3Dtls").exists("tds3Details") Then

'Change-6(c), Test-033, Remark - The array is in tds3Details, nodeList set accordingly
'following statement is commented
    Set Nodelist = jsonObject("form26as")("scheduleTDS3Dtls")("tds3Details")
'    Set nodeList = jsonObject("form26as")("scheduleTDS3Dtls")

    'Malli comented-----------------
'    If jsonObject("form26as").exists("scheduleTDS3Dtls") Then
'    Set Nodelist2 = jsonObject("lastFiledITR")("tdsOnOthThanSals")("tdSonOthThanSal")
'    For Each Node In Nodelist2
'        Nodelist.add Node
'    Next Node
'    Else
'    Set Nodelist = jsonObject("lastFiledITR")("tdsOnOthThanSals")("tdSonOthThanSal")
'    End If
    '-------------------------------
    'lastFiledITR.scheduleTDS3Dtls.tds3Details.broughtFwdTDSAmt
    TDSCreditName = Sheet28.Range("TDS3.TDSCreditName").Column
    PanOtherPerson = Sheet28.Range("TDS3.PanOtherPerson").Column
    Aadhar_TDS3 = Sheet28.Range("Aadhar_TDS3").Column
    PAN = Sheet28.Range("TDS3.PAN").Column
    Aadhaar = Sheet28.Range("TDS3.Aadhaar").Column
    DeductedYear = Sheet28.Range("TDS3.DeductedYear").Column
    BroughtFwdTDSAmt = Sheet28.Range("TDS3.BroughtFwdTDSAmt").Column
    DeductedInOwnHands = Sheet28.Range("TDS3.DeductedInOwnHands").Column
    DeductedIncome5A = Sheet28.Range("TDS3.DeductedIncome5A").Column
    DeductedTDS5A = Sheet28.Range("TDS3.DeductedTDS5A").Column
    ClaimedInOwnHands = Sheet28.Range("TDS3.ClaimedInOwnHands").Column
    ClaimedIncome5A = Sheet28.Range("TDS3.ClaimedIncome5A").Column
    ClaimedTDS5A = Sheet28.Range("TDS3.ClaimedTDS5A").Column
    ClaimedPAN5A = Sheet28.Range("TDS3.ClaimedPAN5A").Column
    ClaimedAadhar5A_TDS3 = Sheet28.Range("ClaimedAadhar5A_TDS3").Column
    GrossAmount = Sheet28.Range("TDS3.GrossAmount").Column
    HeadOfIncome = Sheet28.Range("TDS3.HeadOfIncome").Column
    
    SectionTDS_3_ColNo = Sheet28.Range("TDS3.TdsIsDeducted").Column
    
    TotalExRow = Range("TDS3.PAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet28.Range("TDS3.TDSCreditName").Locked = False Then
            Sheet28.Range("TDS3.TDSCreditName").ClearContents
        End If
        If Sheet28.Range("TDS3.PanOtherPerson").Locked = False Then
            Sheet28.Range("TDS3.PanOtherPerson").ClearContents
        End If
        If Sheet28.Range("Aadhar_TDS3").Locked = False Then
            Sheet28.Range("Aadhar_TDS3").ClearContents
        End If
        If Sheet28.Range("TDS3.PAN").Locked = False Then
            Sheet28.Range("TDS3.PAN").ClearContents
        End If
        If Sheet28.Range("TDS3.Aadhaar").Locked = False Then
            Sheet28.Range("TDS3.Aadhaar").ClearContents
        End If
        'Malli------------
        If Sheet28.Range("TDS3.TdsIsDeducted").Locked = False Then
            Sheet28.Range("TDS3.TdsIsDeducted").ClearContents
        End If
        '----------------
        If Sheet28.Range("TDS3.DeductedYear").Locked = False Then
            Sheet28.Range("TDS3.DeductedYear").ClearContents
        End If
        If Sheet28.Range("TDS3.BroughtFwdTDSAmt").Locked = False Then
            Sheet28.Range("TDS3.BroughtFwdTDSAmt").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedInOwnHands").Locked = False Then
            Sheet28.Range("TDS3.DeductedInOwnHands").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedIncome5A").Locked = False Then
            Sheet28.Range("TDS3.DeductedIncome5A").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedTDS5A").Locked = False Then
            Sheet28.Range("TDS3.DeductedTDS5A").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedInOwnHands").Locked = False Then
            Sheet28.Range("TDS3.ClaimedInOwnHands").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedIncome5A").Locked = False Then
            Sheet28.Range("TDS3.ClaimedIncome5A").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedTDS5A").Locked = False Then
            Sheet28.Range("TDS3.ClaimedTDS5A").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedPAN5A").Locked = False Then
            Sheet28.Range("TDS3.ClaimedPAN5A").ClearContents
        End If
        If Sheet28.Range("ClaimedAadhar5A_TDS3").Locked = False Then
            Sheet28.Range("ClaimedAadhar5A_TDS3").ClearContents
        End If
        If Sheet28.Range("TDS3.GrossAmount").Locked = False Then
            Sheet28.Range("TDS3.GrossAmount").ClearContents
        End If
        If Sheet28.Range("TDS3.HeadOfIncome").Locked = False Then
            Sheet28.Range("TDS3.HeadOfIncome").ClearContents
        End If
    
    End If
    
    If (TotalDiffRow > 0) Then
    Sheet28.Activate   'Malli_SIT_93890
     AddDiffRows_TDS3oth (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TDS3.PAN").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
 
          'Malli---------------------------
          'AY_2023_24 OLd
          'TDSCreditName_TDS3 = Node("tDSCreditName")
          'AY_2024_25 change
          'TDSCreditName_TDS3 = Nodelist("tDSCreditName")
          TDSCreditName_TDS3 = Node("TDSCreditName")
        '----------------------------------------
        If UCase(TDSCreditName_TDS3) = "S" Then
            TDSCreditName_TDS3 = "Self"
        End If
        If UCase(TDSCreditName_TDS3) = "O" Then
            TDSCreditName_TDS3 = "Other Person"
        End If
        'Malli--------------
        'HeadOfIncome
        'AY_2023_24 OLd
       ' HeadIncome_TDS3 = UCase(Node("headOfIncome"))
        'AY_2024_25 Change
        HeadIncome_TDS3 = UCase(Node("HeadOfIncome"))
        '---------------------
        'HeadIncome_TDS3 = UCase(Node("headOfIncome"))
         If HeadIncome_TDS3 = "HP" Then
           HeadIncome_TDS3 = "Income from House Property"
         ElseIf HeadIncome_TDS3 = "BP" Then
            HeadIncome_TDS3 = "Income from Business & Profession"
         ElseIf HeadIncome_TDS3 = "CG" Then
            HeadIncome_TDS3 = "Income from Capital Gains"
         ElseIf HeadIncome_TDS3 = "OS" Then
            HeadIncome_TDS3 = "Income from Other Sources"
         ElseIf HeadIncome_TDS3 = "EI" Then
            HeadIncome_TDS3 = "Exempt Income"
         ElseIf HeadIncome_TDS3 = "NA" Then
            HeadIncome_TDS3 = "Not applicable ( only in case TDS is deducted u/s 194N)"
         Else
            HeadIncome_TDS3 = "(Select)"
        End If
 '----------------------------'Malli
'AY_2023_24 OLD
    'Change-6(d)(i), Test-033, Remark - Unnecessary fields commented
    
'        If Sheet28.Cells(rowcount, TDSCreditName).Locked = False Then
'            Sheet28.Cells(rowcount, TDSCreditName).value = TDSCreditName_TDS3
'        End If
'AY_2024_25 Change

        If Sheet28.Cells(rowcount, TDSCreditName).Locked = False Then
            Sheet28.Cells(rowcount, TDSCreditName).value = TDSCreditName_TDS3
        End If
'--------------------------------------------
'        If Sheet28.Cells(rowcount, PanOtherPerson).Locked = False Then
'            Sheet28.Cells(rowcount, PanOtherPerson).value = Node("panOfTenant")
'        End If
'        If Sheet28.Cells(rowcount, Aadhar_TDS3).Locked = False Then
'            Sheet28.Cells(rowcount, Aadhar_TDS3).value = Node("aadhaarOfOtherPerson")
'        End If
        'Change-6(a), Test-033, Remark - "pANOfBuyerTenant" change to "panOfTenant"
        'following statements are commented
        If Sheet28.Cells(rowcount, PAN).Locked = False Then
            Sheet28.Cells(rowcount, PAN).value = Node("panOfTenant")
        End If
        Application.EnableEvents = True
'        If Sheet28.Cells(rowcount, PAN).Locked = False Then
'            Sheet28.Cells(rowcount, PAN).value = Node("pANOfBuyerTenant")
'        End If
    'Change-6(d)(ii), Test-033, Remark - Unnecessary fields commented
    
'        If Sheet28.Cells(rowcount, Aadhaar).Locked = False Then
'            Sheet28.Cells(rowcount, Aadhaar).value = Node("aadhaarOfBuyerTenant")
'        End If


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 Dim SectionTDS, TDSSection_Gen
 SectionTDS = Node("sectionCode")
 If SectionTDS <> "" Then

                    If UCase(SectionTDS) = UCase("92A") Then
                    TDSSection_Gen = "192-Salary-Payment to Government employees other than Indian Government employees"

                    ElseIf UCase(SectionTDS) = UCase("92B") Then
                    TDSSection_Gen = "192-Salary-Payment to employees other than Government employees"

                    ElseIf UCase(SectionTDS) = UCase("92C") Then
                    TDSSection_Gen = "192-Salary-Payment to Indian Government employees"

                    'Malli----------------------
                    'ElseIf TDSSection_Gen_E = "192A" Then
                    ElseIf UCase(SectionTDS) = UCase("2AA") Then
                    TDSSection_Gen = "192A-TDS on PF withdrawal"
                    '------------------------------------

                    ElseIf UCase(SectionTDS) = UCase("193") Then
                    TDSSection_Gen = "193-Interest on Securities"

                    ElseIf UCase(SectionTDS) = UCase("194") Then
                    TDSSection_Gen = "194-Dividends"

                    ElseIf UCase(SectionTDS) = UCase("94A") Then
                    TDSSection_Gen = "194A-Interest other than 'Interest on securities'"

                    ElseIf UCase(SectionTDS) = UCase("94B") Then
                    TDSSection_Gen = "194B-Winning from lottery or crossword puzzle"

                    'Malli---------
                    'ElseIf TDSSection_Gen_E = "94BA" Then
                      ElseIf UCase(SectionTDS) = UCase("9BA") Then
                    TDSSection_Gen = "194BA-Winnings from online games"
                    '------------------------

                    ElseIf UCase(SectionTDS) = UCase("4BB") Then
                    TDSSection_Gen = "194BB-Winning from horse race"

                    ElseIf UCase(SectionTDS) = UCase("94C") Then
                    TDSSection_Gen = "194C-Payments to contractors and sub-contractors"

                    ElseIf UCase(SectionTDS) = UCase("94D") Then
                    TDSSection_Gen = "194D-Insurance commission"

                    ElseIf UCase(SectionTDS) = UCase("4DA") Then
                    TDSSection_Gen = "194DA-Payment in respect of life insurance policy"

                    ElseIf UCase(SectionTDS) = UCase("94E") Then
                    TDSSection_Gen = "194E-Payments to non-resident sportsmen or sports associations"

                    ElseIf UCase(SectionTDS) = UCase("4EE") Then
                    TDSSection_Gen = "194EE-Payments in respect of deposits under National Savings"
'Malli-----------------------------------------------------------------
                    'ElseIf TDSSection_Gen_E = "4F" Then
                     ElseIf UCase(SectionTDS) = UCase("94F") Then
                    TDSSection_Gen = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"

                    'ElseIf TDSSection_Gen_E = "4G" Then
                     ElseIf UCase(SectionTDS) = UCase("94G") Then
                    TDSSection_Gen = "194G-Commission, price, etc. on sale of lottery tickets"

                    'ElseIf TDSSection_Gen_E = "4H" Then
                      ElseIf UCase(SectionTDS) = UCase("94H") Then
                    TDSSection_Gen = "194H-Commission or brokerage"

                    'ElseIf TDSSection_Gen_E = "4-IA" Then
                     ElseIf UCase(SectionTDS) = UCase("4IA") Then
                      'SIT_94076
                    'TDSSection_Gen = "194I(a)-Rent on hiring of plant and machinery"
                    TDSSection_Gen = "194IA-TDS on Sale of immovable property"

                    'ElseIf TDSSection_Gen_E = "4-IB" Then
                     ElseIf UCase(SectionTDS) = UCase("4IB") Then
                     'SIT_94076
                    'TDSSection_Gen = "194I(b)-Rent on other than plant and machinery"
                    TDSSection_Gen = "194IB-Payment of rent by certain individuals or Hindu undivided"

                    'ElseIf TDSSection_Gen_E = "4IA" Then
                    ElseIf UCase(SectionTDS) = UCase("9IA") Then
                    TDSSection_Gen = "194IA-TDS on Sale of immovable property"

                   ' ElseIf TDSSection_Gen_E = "4IB" Then
                    ElseIf UCase(SectionTDS) = UCase("9IB") Then
                    TDSSection_Gen = "194IB-Payment of rent by certain individuals or Hindu undivided"

                    ElseIf UCase(SectionTDS) = UCase("4IC") Then
                    TDSSection_Gen = "194IC-Payment under specified agreement"

                    'ElseIf TDSSection_Gen_E = "94J-A" Then
                    ElseIf UCase(SectionTDS) = UCase("4JA") Then
                    TDSSection_Gen = "194J(a)-Fees for technical services"

                    'ElseIf TDSSection_Gen_E = "94J-B" Then
                    ElseIf UCase(SectionTDS) = UCase("4JB") Then
                    TDSSection_Gen = "194J(b)-Fees for professional  services or royalty etc"

                    ElseIf UCase(SectionTDS) = UCase("94K") Then
                    TDSSection_Gen = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"

                    ElseIf UCase(SectionTDS) = UCase("4LA") Then
                    TDSSection_Gen = "194LA-Payment of compensation on acquisition of certain immovable"

                    ElseIf UCase(SectionTDS) = UCase("4LB") Then
                    TDSSection_Gen = "194LB-Income by way of Interest from Infrastructure Debt fund"

                    'ElseIf TDSSection_Gen_E = "4LC1" Then
                    ElseIf UCase(SectionTDS) = UCase("LC1") Then
                    TDSSection_Gen = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"

                    'ElseIf TDSSection_Gen_E = "4LC2" Then
                    ElseIf UCase(SectionTDS) = UCase("LC2") Then
                    TDSSection_Gen = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"

                    'ElseIf TDSSection_Gen_E = "4LC3" Then
                     ElseIf UCase(SectionTDS) = UCase("Lc3") Then
                    TDSSection_Gen = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
'Malli------------------------------------------------------------------
'                    ElseIf TDSSection_Gen_E = "4BA1" Then
'                    TDSSection_Gen = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'
'                    ElseIf TDSSection_Gen_E = "4BA2" Then
'                    TDSSection_Gen = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'
'                    ElseIf TDSSection_Gen_E = "LBA1" Then
'                    TDSSection_Gen = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
'
'                    ElseIf TDSSection_Gen_E = "LBA2" Then
'                    TDSSection_Gen = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"

'Malli_SIT-117189  13/04/2026

                    ElseIf UCase(SectionTDS) = UCase("BA1") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then
                    TDSSection_Gen = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"

                    ElseIf UCase(SectionTDS) = UCase("BA2") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then
                    TDSSection_Gen = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"

                    ElseIf UCase(SectionTDS) = UCase("BA1") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
                    TDSSection_Gen = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"

                    ElseIf UCase(SectionTDS) = UCase("BA2") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
                    TDSSection_Gen = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"



'------------------------------------------------------------------------
                    'ElseIf TDSSection_Gen_E = "LBA3" Then
                    ElseIf UCase(SectionTDS) = UCase("BA3") Then
                    TDSSection_Gen = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"

                    ElseIf UCase(SectionTDS) = UCase("LBB") Then
                    TDSSection_Gen = "194LBB-Income in respect of units of investment fund"

                    ElseIf UCase(SectionTDS) = UCase("94R") Then
                    TDSSection_Gen = "194R-Benefits or perquisites of business or profession"

                    ElseIf UCase(SectionTDS) = UCase("94S") Then
                    TDSSection_Gen = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"

                    'ElseIf TDSSection_Gen_E = "94B-P" Then
                     ElseIf UCase(SectionTDS) = UCase("4BP") Then
                    TDSSection_Gen = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"

                    'ElseIf TDSSection_Gen_E = "94R-P" Then
                    ElseIf UCase(SectionTDS) = UCase("4RP") Then
                    TDSSection_Gen = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"

                    'ElseIf TDSSection_Gen_E = "94S-P" Then
                    ElseIf UCase(SectionTDS) = UCase("4SP") Then
                    TDSSection_Gen = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"

                    ElseIf UCase(SectionTDS) = UCase("LBC") Then
                    TDSSection_Gen = "194LBC-Income in respect of investment in securitization trust"

                    ElseIf UCase(SectionTDS) = UCase("4LD") Then
                    TDSSection_Gen = "194LD-TDS on interest on bonds / government securities"

                    ElseIf UCase(SectionTDS) = UCase("94M") Then
                    TDSSection_Gen = "194M-Payment of certain sums by certain individuals or HUF"

                    ElseIf UCase(SectionTDS) = UCase("94N") Then
                    TDSSection_Gen = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"

                    'ElseIf TDSSection_Gen_E = "94N-F" Then
                    ElseIf UCase(SectionTDS) = UCase("4NF") Then
                    TDSSection_Gen = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"

                    'ElseIf TDSSection_Gen_E = "94N-C" Then
                    ElseIf UCase(SectionTDS) = UCase("4NC") Then
                    TDSSection_Gen = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"

                    'ElseIf TDSSection_Gen_E = UCase("94N-FT" Then
                    ElseIf UCase(SectionTDS) = UCase("NFT") Then
                    TDSSection_Gen = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"

                    ElseIf UCase(SectionTDS) = UCase("94O") Then
                    TDSSection_Gen = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."

                    ElseIf UCase(SectionTDS) = UCase("94P") Then
                    TDSSection_Gen = "194P-Deduction of tax in case of specified senior citizen"

                    ElseIf UCase(SectionTDS) = UCase("94Q") Then
                    TDSSection_Gen = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"

                    ElseIf UCase(SectionTDS) = UCase("195") Then
                    TDSSection_Gen = "195-Other sums payable to a non-resident"

                    ElseIf UCase(SectionTDS) = UCase("96A") Then
                    TDSSection_Gen = "196A-Income in respect of units of non-residents"

                    ElseIf UCase(SectionTDS) = UCase("96B") Then
                    TDSSection_Gen = "196B-Payments in respect of units to an offshore fund"

                    ElseIf UCase(SectionTDS) = UCase("96C") Then
                    TDSSection_Gen = "196C-Income from foreign currency bonds or shares of Indian"

                    ElseIf UCase(SectionTDS) = UCase("96D") Then
                    TDSSection_Gen = "196D-Income of foreign institutional investors from securities"

                    'ElseIf TDSSection_Gen_E = "96DA" Then
                     ElseIf UCase(SectionTDS) = UCase("6DA") Then
                    TDSSection_Gen = "196D(1A)-Income of specified fund from securities"

                    'ElseIf TDSSection_Gen_E = "94BA-P" Then
                     ElseIf UCase(SectionTDS) = UCase("BAP") Then
                    TDSSection_Gen = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
                    Else
                        TDSSection_Gen = "(Select)"
                    End If

            End If
            If TDSSection_Gen <> "" Then
            Sheet28.Cells(rowcount, SectionTDS_3_ColNo).value = TDSSection_Gen
                    SectionTDS = ""
             
        End If

        
        
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



        If Sheet28.Cells(rowcount, DeductedYear).Locked = False Then
            Sheet28.Cells(rowcount, DeductedYear).value = Node("deductedYr")
        End If
        If Sheet28.Cells(rowcount, BroughtFwdTDSAmt).Locked = False And Node("broughtFwdTDSAmt") <> 0 Then
            Sheet28.Cells(rowcount, BroughtFwdTDSAmt).value = Node("broughtFwdTDSAmt")
        End If
        If Sheet28.Cells(rowcount, DeductedInOwnHands).Locked = False And Node("taxDeductCreditDtls")("taxDeductedOwnHands") <> 0 Then
            Sheet28.Cells(rowcount, DeductedInOwnHands).value = Node("taxDeductCreditDtls")("taxDeductedOwnHands")
        End If
'        If Sheet28.Cells(rowcount, DeductedIncome5A).Locked = False Then
'            Sheet28.Cells(rowcount, DeductedIncome5A).value = Node("taxDeductCreditDtls")("taxDeductedIncome")
'        End If
'        If Sheet28.Cells(rowcount, DeductedTDS5A).Locked = False Then
'            Sheet28.Cells(rowcount, DeductedTDS5A).value = Node("taxDeductCreditDtls")("taxDeductedTDS")
'        End If
        If Sheet28.Cells(rowcount, ClaimedInOwnHands).Locked = False Then
            Sheet28.Cells(rowcount, ClaimedInOwnHands).value = Node("taxDeductCreditDtls")("taxClaimedOwnHands")
        End If
'        If Sheet28.Cells(rowcount, ClaimedIncome5A).Locked = False Then
'            Sheet28.Cells(rowcount, ClaimedIncome5A).value = Node("taxDeductCreditDtls")("taxClaimedIncome")
'        End If
'        If Sheet28.Cells(rowcount, ClaimedTDS5A).Locked = False Then
'            Sheet28.Cells(rowcount, ClaimedTDS5A).value = Node("taxDeductCreditDtls")("taxClaimedTDS")
'        End If
'        If Sheet28.Cells(rowcount, ClaimedPAN5A).Locked = False Then
'            Sheet28.Cells(rowcount, ClaimedPAN5A).value = Node("taxDeductCreditDtls")("taxClaimedSpouseOthPrsnPAN")
'        End If
'        If Sheet28.Cells(rowcount, ClaimedAadhar5A_TDS3).Locked = False Then
'            Sheet28.Cells(rowcount, ClaimedAadhar5A_TDS3).value = Node("taxDeductCreditDtls")("spouseOthPrsnAadhaar")
'        End If
        If Sheet28.Cells(rowcount, GrossAmount).Locked = False Then
            Sheet28.Cells(rowcount, GrossAmount).value = Node("grossAmount")
        End If
        If Sheet28.Cells(rowcount, HeadOfIncome).Locked = False Then
            Sheet28.Cells(rowcount, HeadOfIncome).value = HeadIncome_TDS3
        End If
     
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If
End If
End Function
 

Function ScheduleTDS1Import(jsonText As String)
On Error Resume Next
   
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, Node, init, jsonObject As Object
    
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow, rowcount, cnt As Long

 Set jsonObject = ParseJson.ParseJson(jsonText)

Set init = jsonObject("form26as")
If init.exists("tdsOnSalaries") Then
  
    Set Nodelist = jsonObject("tdsOnSalaries")
    
    TANNoEmployer = Sheet28.Range("TDS1.TAN").Column
    TDSNameOfEmployer = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet28.Range("TDS1.IncChrgSal").Column
    TDSTotalTax = Sheet28.Range("TDS1.TotalTDSSal").Column
    
    TotalExRow = Range("TDS1.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet28.Range("TDS1.TAN").Locked = False Then
            Sheet28.Range("TDS1.TAN").ClearContents
        End If
        If Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Locked = False Then
            Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").ClearContents
        End If
        If Sheet28.Range("TDS1.IncChrgSal").Locked = False Then
            Sheet28.Range("TDS1.IncChrgSal").ClearContents
        End If
        If Sheet28.Range("TDS1.TotalTDSSal").Locked = False Then
            Sheet28.Range("TDS1.TotalTDSSal").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TDS1.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            If Sheet28.Cells(rowcount, TANNoEmployer).Locked = False Then
                Sheet28.Cells(rowcount, TANNoEmployer).value = Node("EmployerOrDeductorOrCollectDetl")("tan")
            End If
            If Sheet28.Cells(rowcount, TDSNameOfEmployer).Locked = False Then
                Sheet28.Cells(rowcount, TDSNameOfEmployer).value = Node("EmployerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
            If Sheet28.Cells(rowcount, TDSIncomeCharge).Locked = False Then
                Sheet28.Cells(rowcount, TDSIncomeCharge).value = Node("incChrgSal")
            End If
            If Sheet28.Cells(rowcount, TDSTotalTax).Locked = False Then
                Sheet28.Cells(rowcount, TDSTotalTax).value = Node("totalTDSSal")
            End If
        cnt = cnt + 1
        
    Next Node
    RecTDS1 = cnt
End If
End Function

Function ImportSchedule5A(jsonText As String)
On Error Resume Next
Dim NameOfSpouse, PANOfSpouse, AadhaarOfSpouse, jsonObject
Set jsonObject = ParseJson.ParseJson(jsonText)

NameOfSpouse = jsonObject("lastFiledITR")("schedule5A2014")("nameOfSpouse")
PANOfSpouse = jsonObject("lastFiledITR")("schedule5A2014")("panOfSpouse")
'Newly added by Bindu
AadhaarOfSpouse = jsonObject("lastFiledITR")("schedule5A2014")("aadharOfSpouse")
 
'NameOfSpouse = jsonObject("Schedule5A2014")("NameOfSpouse")
'PANOfSpouse = jsonObject("Schedule5A2014")("PANOfSpouse")
If NameOfSpouse <> "" And Sheet29.Range("S5A_NameOfSpouse").Locked = False Then
    Sheet29.Range("S5A_NameOfSpouse").value = NameOfSpouse
End If

If PANOfSpouse <> "" And Sheet29.Range("S5A_PANOfSpouse").Locked = False Then
    Sheet29.Range("S5A_PANOfSpouse").value = PANOfSpouse
End If
'Newly added by Bindu

If AadhaarOfSpouse <> "" And Sheet29.Range("S5A_AadhaarOfSpouse").Locked = False Then
    Sheet29.Range("S5A_AadhaarOfSpouse").value = AadhaarOfSpouse
End If




End Function

Function ImportSchedulePTI_Old(jsonText As String)
On Error Resume Next
    Dim XpathOfPTI As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As Object
    Dim Node As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim rowcount, cnt, i As Long
    Dim init, jsonObject As Object
   
    
'Change-28, prefill-tracker, D
Set jsonObject = ParseJson.ParseJson(jsonText)
Set init = jsonObject("form64A")("schedulePTI")
If init.exists("schedulePTIDtls") Then
   
    Set Nodelist = jsonObject("form64A")("SchedulePTI")("SchedulePTIDtls")

    TotalExRow = Sheet41.Range("cntr.Ptirptfrm").value
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If TotalDiffRow > 0 Then
        For i = 1 To TotalDiffRow
            Sheet41.Activate
            AddRows_PTI
        Next
    End If
    Dim InvstmntCvrdUs115UA115UB As Variant
    
    rowcount = 0
    cnt = 0
    For Each Node In Nodelist
    
            rowcount = rowcount + 1
            InvstmntCvrdUs115UA115UB = UCase(Node(("InvstmntCvrdUs115UA115UB")))
                
                If InvstmntCvrdUs115UA115UB = "A" Then
                   InvstmntCvrdUs115UA115UB = "SECTION 115UA"
                ElseIf InvstmntCvrdUs115UA115UB = "B" Then
                   InvstmntCvrdUs115UA115UB = "SECTION 115UB"
                Else
                    InvstmntCvrdUs115UA115UB = "(Select)"
                End If
                
        If Sheet41.Range("PTI_Investment_UA_UB" & rowcount).Locked = False Then
            Sheet41.Range("PTI_Investment_UA_UB" & rowcount).value = InvstmntCvrdUs115UA115UB
        End If
        If Sheet41.Range("PTI_NameOfBusiness" & rowcount).Locked = False Then
            Sheet41.Range("PTI_NameOfBusiness" & rowcount).value = UCase(Node("BusinessName"))
        End If
        If Sheet41.Range("PTI_PAN" & rowcount).Locked = False Then
            Sheet41.Range("PTI_PAN" & rowcount).value = UCase(Node("BusinessPAN"))
        End If
        If Sheet41.Range("PTI_HPIncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_HPIncomeAmt" & rowcount).value = Node("IncFromHP")("AmountOfInc")
        End If
        If Sheet41.Range("PTI_HP_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_HP_ShareAmt" & rowcount).value = Node("IncFromHP")("CurrYrLossShareByInvstFund")
        End If
'        If Sheet41.Range("PTI_HP_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_HP_NetAmt" & rowcount).value = Node("IncFromHP")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_HPTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_HPTDSAmt" & rowcount).value = Node("IncFromHP")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_CG_STCGIncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCGIncomeAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_NetAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_STCGTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCGTDSAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & rowcount).value = Node("CapitalGainsPTI")("STCG_Sec111A")("AmountOfInc")
        End If
'        If Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Sec111A")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_SectionA_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_SectionA_NetAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Sec111A")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Sec111A")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & rowcount).value = Node("capitalGainsPTI")("IncOthSrc")("AmountOfInc")
        End If
        If Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("IncOthSrc")("CurrYrLossShareByInvstFund")
        End If
'        If Sheet41.Range("PTI_CG_STCG_Others_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_Others_NetAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Others")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Others")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_CG_LTCGIncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCGIncomeAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_NetAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_LTCGTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCGTDSAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & rowcount).value = Node("capitalGainsPTI")("ltcg_Sec112A")("amountOfInc")
        End If
'        If Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Sec112A")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_SectionA_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_SectionA_NetAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Sec112A")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Sec112A")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & rowcount).value = Node("capitalGainsPTI")("ltcg_Others")("amountOfInc")
        End If
'        If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Others")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_Others_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_Others_NetAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Others")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Others")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_Us_iv_IncomeAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_iv_IncomeAmta" & rowcount).value = Node("IncClmdPTI")("TotalSec23FBB")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_iv_NetAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_iv_NetAmta" & rowcount).value = Node("IncClmdPTI")("TotalSec23FBB")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_Us_iv_TdsAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_iv_TdsAmta" & rowcount).value = Node("IncClmdPTI")("TotalSec23FBB")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_UsIncomeAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsIncomeAmta" & rowcount).value = Node("IncClmdPTI")("Sec23FBB")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_NetAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_NetAmta" & rowcount).value = Node("IncClmdPTI")("Sec23FBB")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_UsTDSAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsTDSAmta" & rowcount).value = Node("IncClmdPTI")("Sec23FBB")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_Usb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Usb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SectionCode")
'        End If
'        If Sheet41.Range("PTI_UsIncomeAmtb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsIncomeAmtb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SecBCIncExmptDtl")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_NetAmtb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_NetAmtb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SecBCIncExmptDtl")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_UsTDSAmtb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsTDSAmtb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SecBCIncExmptDtl")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_Usc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Usc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SectionCode")
'        End If
'        If Sheet41.Range("PTI_UsIncomeAmtc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsIncomeAmtc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SecBCIncExmptDtl")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_NetAmtc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_NetAmtc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SecBCIncExmptDtl")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_UsTDSAmtc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsTDSAmtc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SecBCIncExmptDtl")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_OSIncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_OSIncomeAmt" & rowcount).value = Node("os_Others")("amountOfInc")
        End If
'        If Sheet41.Range("PTI_OS_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_NetAmt" & rowcount).value = Node("OS_Others")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_OSTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OSTDSAmt" & rowcount).value = Node("OS_Others")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & rowcount).value = Node("OS_Dividend")("AmountOfInc")
        End If
'        If Sheet41.Range("PTI_OS_Dividend_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Dividend_NetAmt" & rowcount).value = Node("OS_Dividend")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_OS_Dividend_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Dividend_TdsAmt" & rowcount).value = Node("OS_Dividend")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_OS_Others_IncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Others_IncomeAmt" & rowcount).value = Node("IncOthSrc")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_OS_Others_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Others_NetAmt" & rowcount).value = Node("IncOthSrc")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_OS_Others_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Others_TdsAmt" & rowcount).value = Node("IncOthSrc")("TDSAmount")
'        End If
                
            cnt = cnt + 1
    Next Node
    
    RecTDS1 = cnt
End If
End Function

Function ImportSchedule80D(jsonText As String)
On Error Resume Next
Dim Node, Nodelist, jsonObject, init As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

Set init = jsonObject("lastFiledITR")

If Not init.exists("schedule80D") Then
    Exit Function
End If

If Not init("schedule80D").exists("Sec80DSelfFamSrCtznHealth") Then
    Exit Function
End If
Set Node = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")

 
        
        If Trim(Node("SeniorCitizenFlag")) <> "" Then
            If Node("SeniorCitizenFlag") = "Y" Then
                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "Yes"
            ElseIf Node("SeniorCitizenFlag") = "N" Then
                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "No"
            ElseIf Node("SeniorCitizenFlag") = "S" Then
            'Change-4, Test-063, Remark- A spelling correction required, following statement is commented
            '"No claiming for Self/Family" to "Not claiming for Self/Family"
                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "Not Claiming for Self/Family"
'                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "No Claiming for Self/Family"
            End If
        End If
        
        If Trim(Node("ParentSeniorCitizenFlag")) <> "" Then
            If Node("ParentSeniorCitizenFlag") = "Y" Then
                Sheet55.Range("DropDown_ValueOf_SC_80D").value = "Yes"
            ElseIf Node("ParentSeniorCitizenFlag") = "N" Then
                Sheet55.Range("DropDown_ValueOf_SC_80D").value = "No"
            ElseIf Node("ParentSeniorCitizenFlag") = "P" Then
                Sheet55.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
            End If
        Else
            Sheet55.Range("DropDown_ValueOf_SC_80D").value = "(Select)"
        End If
        
End Function

Function ImportScheduleSPI(jsonText As String)
On Error Resume Next
    Dim XpathOfSPI As String
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    
    Dim SpecifiedPersonName, PANofSpecPerson, ReltnShip, NatureOfInc, AmtIncluded, AadhaarofSpecPerson, IncmNature_SPI As Variant
    
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim jsonDictionary, init, jsonObject As Object
    Dim Node, Nodelist As Object
    Dim rowcount As Variant
    Set jsonObject = ParseJson.ParseJson(jsonText)

'lastFiledITR.scheduleSPI.specifiedPerson.aadhaarOfSpecPerson
Set init = jsonObject("lastFiledITR")("scheduleSPI")
If init.exists("specifiedPerson") Then
    Set Nodelist = jsonObject("lastFiledITR")("scheduleSPI")("specifiedPerson")
    
    SpecifiedPersonName = Sheet21.Range("SPI.SpecifiedPersonName").Column
    PANofSpecPerson = Sheet21.Range("SPI.PANofSpecPerson").Column
    ReltnShip = Sheet21.Range("SPI.ReltnShip").Column
    NatureOfInc = Sheet21.Range("SPI.NatureOfInc").Column
    AmtIncluded = Sheet21.Range("SPI.AmtIncluded").Column
    AadhaarofSpecPerson = Sheet21.Range("SPI.AadhaarofSpecPerson").Column
    
    TotalExRow = Sheet21.Range("SPI.SpecifiedPersonName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet21.Range("SPI.SpecifiedPersonName").Locked = False Then
            Sheet21.Range("SPI.SpecifiedPersonName").ClearContents
        End If
        If Sheet21.Range("SPI.PANofSpecPerson").Locked = False Then
            Sheet21.Range("SPI.PANofSpecPerson").ClearContents
        End If
        If Sheet21.Range("SPI.ReltnShip").Locked = False Then
            Sheet21.Range("SPI.ReltnShip").ClearContents
        End If
        If Sheet21.Range("SPI.NatureOfInc").Locked = False Then
            Sheet21.Range("SPI.NatureOfInc").ClearContents
        End If
        If Sheet21.Range("SPI.AmtIncluded").Locked = False Then
            Sheet21.Range("SPI.AmtIncluded").ClearContents
        End If
        If Sheet21.Range("SPI.AadhaarofSpecPerson").Locked = False Then
            Sheet21.Range("SPI.AadhaarofSpecPerson").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_SPI (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet21.Range("SPI.SpecifiedPersonName").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
            IncmNature_SPI = UCase(Node("headIncIncluded"))
            
            If IncmNature_SPI = "BP" Then
               IncmNature_SPI = "Business/Profession"
            ElseIf IncmNature_SPI = "OS" Then
               IncmNature_SPI = "Other sources"
            ElseIf IncmNature_SPI = "EI" Then
               IncmNature_SPI = "Exempt Income"
            ElseIf IncmNature_SPI = "HP" Then
               IncmNature_SPI = "House Property"
            ElseIf IncmNature_SPI = "CG" Then
               IncmNature_SPI = "Capital Gains"
            ElseIf IncmNature_SPI = "SA" Then
               IncmNature_SPI = "Salary"
            Else
                IncmNature_SPI = "(Select)"
            End If
        
        If Sheet21.Cells(rowcount, SpecifiedPersonName).Locked = False Then
            Sheet21.Cells(rowcount, SpecifiedPersonName).value = UCase(Node("specifiedPersonName"))
        End If
        If Sheet21.Cells(rowcount, PANofSpecPerson).Locked = False Then
            Sheet21.Cells(rowcount, PANofSpecPerson).value = UCase(Node("paNofSpecPerson"))
        End If
        If Sheet21.Cells(rowcount, ReltnShip).Locked = False Then
            Sheet21.Cells(rowcount, ReltnShip).value = UCase(Node("reltnShip"))
        End If
        If Sheet21.Cells(rowcount, NatureOfInc).Locked = False Then
            Sheet21.Cells(rowcount, NatureOfInc).value = IncmNature_SPI
        End If
        If Sheet21.Cells(rowcount, AmtIncluded).Locked = False Then
            Sheet21.Cells(rowcount, AmtIncluded).value = Node("amtIncluded")
        End If
        If Sheet21.Cells(rowcount, AadhaarofSpecPerson).Locked = False Then
            Sheet21.Cells(rowcount, AadhaarofSpecPerson).value = UCase(Node("aadhaarOfSpecPerson"))
        End If
            
    Next Node
 End If
End Function

'Function ImportSchedule80D(jsonText As String)
'On Error Resume Next
'Dim Node, nodeList, jsonObject As Object
'Set jsonObject = ParseJson.ParseJson(jsonText)
'Set Node = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")
'
'        If Trim(Node("SeniorCitizenFlag")) <> "" Then
'            If Node("SeniorCitizenFlag") = "Y" Then
'                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "Yes"
'            ElseIf Node("SeniorCitizenFlag") = "N" Then
'                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "No"
'            ElseIf Node("SeniorCitizenFlag") = "S" Then
'                Sheet55.Range("DropDown_ValueOf_FamilyM_80D").value = "No Claiming for Self/Family"
'            End If
'        End If
'
'        If Trim(Node("HlthInsPremSlfFamSrCtzn")) <> "" And Sheet55.Range("Health_InsuranceSC_80D").Locked = False Then
'            Sheet55.Range("Health_InsuranceSC_80D").value = Node("HlthInsPremSlfFamSrCtzn")
'        End If
'        If Trim(Node("PrevHlthChckUpSlfFamSrCtzn")) <> "" And Sheet55.Range("Preventive_Health_SC_80D").Locked = False Then
'            Sheet55.Range("Preventive_Health_SC_80D").value = Node("PrevHlthChckUpSlfFamSrCtzn")
'        End If
'        If Trim(Node("MedicalExpSlfFamSrCtzn")) <> "" And Sheet55.Range("Medical_Expenditure_SC_80D").Locked = False Then
'            Sheet55.Range("Medical_Expenditure_SC_80D").value = Node("MedicalExpSlfFamSrCtzn")
'        End If
'
'        If Trim(Node("ParentSeniorCitizenFlag")) <> "" Then
'            If Node("ParentSeniorCitizenFlag") = "Y" Then
'                Sheet55.Range("DropDown_ValueOf_SC_80D").value = "Yes"
'            ElseIf Node("ParentSeniorCitizenFlag") = "N" Then
'                Sheet55.Range("DropDown_ValueOf_SC_80D").value = "No"
'            ElseIf Node("ParentSeniorCitizenFlag") = "P" Then
'                Sheet55.Range("DropDown_ValueOf_SC_80D").value = "Not claiming for Parents"
'            End If
'        Else
'            Sheet55.Range("DropDown_ValueOf_SC_80D").value = "(Select)"
'        End If
'
'        If Sheet55.Range("DropDown_ValueOf_SC_80D").value = "Yes" Then
'
'            If Trim(Node("HlthInsPremParentsSrCtzn")) <> "" And Sheet55.Range("Health_Insurance3_80D").Locked = False Then
'            Sheet55.Range("Health_Insurance3_80D").value = CStr(Node("HlthInsPremParentsSrCtzn"))
'            End If
'            If Trim(Node("PrevHlthChckUpParentsSrCtzn")) <> "" And Sheet55.Range("Preventive_Health3_80D").Locked = False Then
'                Sheet55.Range("Preventive_Health3_80D").value = CStr(Node("PrevHlthChckUpParentsSrCtzn"))
'            End If
'            If Trim(Node("MedicalExpParentsSrCtzn")) <> "" And Sheet55.Range("Medical_Expenditure2_80D").Locked = False Then
'                Sheet55.Range("Medical_Expenditure2_80D").value = CStr(Node("MedicalExpParentsSrCtzn"))
'            End If
'
'        ElseIf Sheet55.Range("DropDown_ValueOf_SC_80D").value = "No" Then
'
'            If Trim(Node("HlthInsPremParents")) <> "" And Sheet55.Range("Health_Insurance2_80D").Locked = False Then
'                Sheet55.Range("Health_Insurance2_80D").value = CStr(Node("HlthInsPremParents"))
'            End If
'            If Trim(Node("PrevHlthChckUpParents")) <> "" And Sheet55.Range("Preventive_Health2_80D").Locked = False Then
'                Sheet55.Range("Preventive_Health2_80D").value = CStr(Node("PrevHlthChckUpParents"))
'            End If
'
'        ElseIf Sheet55.Range("DropDown_ValueOf_SC_80D").value = "(Select)" Then
'
'            If Trim(Node("HlthInsPremParents")) <> "" And Sheet55.Range("Health_Insurance2_80D").Locked = False Then
'                Sheet55.Range("Health_Insurance2_80D").value = CStr(Node("HlthInsPremParents"))
'            End If
'            If Trim(Node("PrevHlthChckUpParents")) <> "" And Sheet55.Range("Preventive_Health2_80D").Locked = False Then
'                Sheet55.Range("Preventive_Health2_80D").value = CStr(Node("PrevHlthChckUpParents"))
'            End If
'             If Trim(Node("HlthInsPremParentsSrCtzn")) <> "" And Sheet55.Range("Health_Insurance3_80D").Locked = False Then
'                Sheet55.Range("Health_Insurance3_80D").value = CStr(Node("HlthInsPremParentsSrCtzn"))
'            End If
'            If Trim(Node("PrevHlthChckUpParentsSrCtzn")) <> "" And Sheet55.Range("Preventive_Health3_80D").Locked = False Then
'                Sheet55.Range("Preventive_Health3_80D").value = CStr(Node("PrevHlthChckUpParentsSrCtzn"))
'            End If
'            If Trim(Node("MedicalExpParentsSrCtzn")) <> "" And Sheet55.Range("Medical_Expenditure2_80D").Locked = False Then
'                Sheet55.Range("Medical_Expenditure2_80D").value = CStr(Node("MedicalExpParentsSrCtzn"))
'            End If
'
'        End If
'
'End Function

Function ImportScheduleEI(jsonText As String)
On Error Resume Next
    Dim XpathOfSPI As String
    Dim TotalXMLRow, rowcount As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1  As Long
    
    Dim Name_of_district, SchEI_Pincode, SchEI_AgriculturalLand, SchEI_OwedOrLease, SchEI_IrrigatedOrRainFed As Variant
    Dim AgriLandOwnedFlag, NatureDesc, NatureDesc1, AgriLandIrrigatedFlag As Variant
    
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow  As Long
    Dim jsonDictionary, jsonObject As Object
    Dim Node, Nodelist, nodeList1, init As Object
Set jsonObject = ParseJson.ParseJson(jsonText)

If Not IsNull(jsonObject("ScheduleEI")) Then    'For AY 2025-26_Konda_23_07_2025 SIT-97244
  
    Set Nodelist = jsonObject("lastFiledITR")("scheduleEI")
    
        If Sheet24.Range("Sheet20.scei.InterestInc").Locked = False Then
            Sheet24.Range("Sheet20.scei.InterestInc").value = Nodelist("interestInc")
        End If
        If Sheet24.Range("Sheet20.scei.NetAgriIncOrOthrIncRule7").Locked = False Then
            Sheet24.Range("Sheet20.scei.NetAgriIncOrOthrIncRule7").value = Nodelist("agriIncRule7and8")
        End If
        If Sheet24.Range("Sheet20.ExpenditureOnAgriculture").Locked = False Then
            Sheet24.Range("Sheet20.ExpenditureOnAgriculture").value = Nodelist("expIncAgri")
        End If
        If Sheet24.Range("Sheet20.UnabsorbedAgriculturalloss").Locked = False Then
            Sheet24.Range("Sheet20.UnabsorbedAgriculturalloss").value = Nodelist("unabAgriLossPrev8")
        End If
Set init = jsonObject("lastFiledITR")("scheduleEI")("excNetAgriInc")
If init.exists("excNetAgriIncDtls") Then
    
    Set Nodelist = jsonObject("lastFiledITR")("scheduleEI")("excNetAgriInc")("excNetAgriIncDtls")
     
    Name_of_district = Sheet24.Range("Name_of_district").Column
    SchEI_Pincode = Sheet24.Range("SchEI_Pincode").Column
    SchEI_AgriculturalLand = Sheet24.Range("SchEI_AgriculturalLand").Column
    SchEI_OwedOrLease = Sheet24.Range("SchEI_OwedOrLease").Column
    SchEI_IrrigatedOrRainFed = Sheet24.Range("SchEI_IrrigatedOrRainFed").Column

    TotalExRow = Sheet24.Range("Name_of_district").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    
    If (TotalXMLRow > 0) Then
        If Sheet24.Range("Name_of_district").Locked = False Then
            Sheet24.Range("Name_of_district").ClearContents
        End If
        If Sheet24.Range("SchEI_Pincode").Locked = False Then
            Sheet24.Range("SchEI_Pincode").ClearContents
        End If
        If Sheet24.Range("SchEI_AgriculturalLand").Locked = False Then
            Sheet24.Range("SchEI_AgriculturalLand").ClearContents
        End If
        If Sheet24.Range("SchEI_OwedOrLease").Locked = False Then
            Sheet24.Range("SchEI_OwedOrLease").ClearContents
        End If
        If Sheet24.Range("SchEI_IrrigatedOrRainFed").Locked = False Then
            Sheet24.Range("SchEI_IrrigatedOrRainFed").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_ExcNetAgriInc (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet24.Range("Name_of_district").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            AgriLandOwnedFlag = UCase(Node("agriLandOwnedFlag"))
            
            
            If AgriLandOwnedFlag = "O" Then
                AgriLandOwnedFlag = "Owned"
            ElseIf AgriLandOwnedFlag = "H" Then
                AgriLandOwnedFlag = "Held on Lease"
            End If
            
            AgriLandIrrigatedFlag = UCase(Node("agriLandIrrigatedFlag"))
            
            If AgriLandIrrigatedFlag = "IRG" Then
                AgriLandIrrigatedFlag = "Irrigated "
            ElseIf AgriLandIrrigatedFlag = "RF" Then
                AgriLandIrrigatedFlag = "Rain-fed"
            End If
            
            If Sheet24.Cells(rowcount, Name_of_district).Locked = False Then
                Sheet24.Cells(rowcount, Name_of_district).value = Node("nameOfDistrict")
            End If
            If Sheet24.Cells(rowcount, SchEI_Pincode).Locked = False Then
                Sheet24.Cells(rowcount, SchEI_Pincode).value = Node("pinCode")
            End If
            If Sheet24.Cells(rowcount, SchEI_AgriculturalLand).Locked = False Then
                Sheet24.Cells(rowcount, SchEI_AgriculturalLand).value = Node("measurementOfLand")
            End If
            
            'Malli
'            If Sheet24.Cells(rowcount, SchEI_OwedOrLease).Locked = False Then
'                Sheet24.Cells(rowcount, SchEI_OwedOrLease).value = AgriLandOwnedFlag
'            End If
'            If Sheet24.Cells(rowcount, SchEI_IrrigatedOrRainFed).Locked = False Then
'                Sheet24.Cells(rowcount, SchEI_IrrigatedOrRainFed).value = AgriLandIrrigatedFlag
'            End If

            If Sheet24.Cells(rowcount, SchEI_OwedOrLease).Locked = False Then
                Sheet24.Cells(rowcount, SchEI_OwedOrLease).value = Node("agriLandOwnedFlag")
            End If
            If Sheet24.Cells(rowcount, SchEI_IrrigatedOrRainFed).Locked = False Then
                Sheet24.Cells(rowcount, SchEI_IrrigatedOrRainFed).value = Node("agriLandIrrigatedFlag")
            End If
            '---------------------------
            
    Next Node
End If


'commented by Malli_23/04/2026 Vo.3 changes De sheet updated
'Set init = jsonObject("lastFiledITR")("scheduleEI")("othersInc")
'If init.exists("othersIncDtls") Then
'
'
'    Set Nodelist = jsonObject("lastFiledITR")("scheduleEI")("othersInc")("othersIncDtls")
'    Dim NatureOfIncome, Description, Amount As Variant
'
'    NatureOfIncome = Sheet24.Range("EI.NatureOfIncome").Column
'    Description = Sheet24.Range("EI.Description").Column
'    Amount = Sheet24.Range("EI.Amount").Column
'
'    TotalExRow = Sheet24.Range("EI.NatureOfIncome").Rows.count
'
'    TotalXMLRow = Nodelist.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'
'    If (TotalXMLRow > 0) Then
'        If Sheet24.Range("EI.NatureOfIncome").Locked = False Then
'            Sheet24.Range("EI.NatureOfIncome").ClearContents
'        End If
'        If Sheet24.Range("EI.Description").Locked = False Then
'            Sheet24.Range("EI.Description").ClearContents
'        End If
'        If Sheet24.Range("EI.Amount").Locked = False Then
'            Sheet24.Range("EI.Amount").ClearContents
'        End If
'    End If
'
'    If (TotalDiffRow > 0) Then
'        AddDiffRows_OthersInc (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet24.Range("EI.NatureOfIncome").name)
'    rowcount = rowcount - 1
'
'    For Each Node In Nodelist
'        rowcount = rowcount + 1
'
'            'pending name of database value for finding "NatureOfIncome"
'            NatureDesc = Node("NatureDesc")
'            NatureDesc1 = Findtext("Sec " & CStr(NatureDesc), "EI_SectionDropdown1")
'
'            If Sheet24.Cells(rowcount, NatureOfIncome).Locked = False Then
'    'Konda updated IPIP-72359 on 28-082025
'                If Node("NatureDesc") = "DMDP" Then
'                    Sheet24.Cells(rowcount, NatureOfIncome).value = "Defense Medical Disability Pension"
'    '--------------IPIP-72359 on 28-082025
'                ElseIf Node("NatureDesc") = "OTH" Then
'                    Sheet24.Cells(rowcount, NatureOfIncome).value = "Any other"
'                Else
'                    Sheet24.Cells(rowcount, NatureOfIncome).value = NatureDesc1
'                End If
'            End If
'            If Sheet24.Cells(rowcount, Description).Locked = False Then
'               Sheet24.Cells(rowcount, Description).value = Node("othNatOfInc")
'            End If
'            If Sheet24.Cells(rowcount, Amount).Locked = False Then
'                Sheet24.Cells(rowcount, Amount).value = Node("othAmount")
'            End If
'
'    Next Node
'End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Set init = jsonObject("lastFiledITR")("scheduleEI")("incNotChrgblAsPerDTAA")
If init.exists("incNotChrgblAsPerDTAADtls") Then
    Dim SCHEI_DTAA_AmountOfIncome, SCHEI_DTAA_NatureOfIncome, SCHEI_DTAA_CountryCode, SCHEI_DTAA_Article, SCHEI_DTAA_HeadOfIncome, SCHEI_TRCObtained, iCountry, sCountry As Variant
    Dim HeadOfIncome_EIDTAA, TRC_EIDTA As Variant
    
    Set Nodelist = jsonObject("lastFiledITR")("scheduleEI")("incNotChrgblAsPerDTAA")("incNotChrgblAsPerDTAADtls")
    
    SCHEI_DTAA_AmountOfIncome = Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Column
    SCHEI_DTAA_NatureOfIncome = Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Column
    SCHEI_DTAA_CountryCode = Sheet24.Range("SCHEI_DTAA_CountryCode").Column
    SCHEI_DTAA_Article = Sheet24.Range("SCHEI_DTAA_Article").Column
    SCHEI_DTAA_HeadOfIncome = Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Column
    SCHEI_TRCObtained = Sheet24.Range("SCHEI_TRCObtained").Column
    
    TotalExRow = Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    
    If (TotalXMLRow > 0) Then
        If Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Locked = False Then
            Sheet24.Range("SCHEI_DTAA_AmountOfIncome").ClearContents
        End If
        If Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Locked = False Then
            Sheet24.Range("SCHEI_DTAA_NatureOfIncome").ClearContents
        End If
        If Sheet24.Range("SCHEI_DTAA_CountryCode").Locked = False Then
            Sheet24.Range("SCHEI_DTAA_CountryCode").ClearContents
        End If
        If Sheet24.Range("SCHEI_DTAA_Article").Locked = False Then
            Sheet24.Range("SCHEI_DTAA_Article").ClearContents
        End If
        If Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Locked = False Then
            Sheet24.Range("SCHEI_DTAA_HeadOfIncome").ClearContents
        End If
        If Sheet24.Range("SCHEI_TRCObtained").Locked = False Then
            Sheet24.Range("SCHEI_TRCObtained").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_IncNotChrgblAsPerDTAA (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet24.Range("SCHEI_DTAA_AmountOfIncome").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
            iCountry = Node("countryCode")
            sCountry = Findtext(iCountry, "Country")
            
            HeadOfIncome_EIDTAA = Node("headOfIncome")
            
            If HeadOfIncome_EIDTAA = "SA" Then
                HeadOfIncome_EIDTAA = "Salary"
            ElseIf HeadOfIncome_EIDTAA = "HP" Then
                HeadOfIncome_EIDTAA = "House Property"
            ElseIf HeadOfIncome_EIDTAA = "PG" Then
               HeadOfIncome_EIDTAA = "Profits & Gains from Business & Profession"
            ElseIf HeadOfIncome_EIDTAA = "CG" Then
                HeadOfIncome_EIDTAA = "Capital Gain"
            ElseIf HeadOfIncome_EIDTAA = "OS" Then
                HeadOfIncome_EIDTAA = "Income from Other sources"
            Else
                HeadOfIncome_EIDTAA = "(Select)"
            End If
            

            TRC_EIDTA = Node("tRCFlag")
            If TRC_EIDTA = "Y" Then
                TRC_EIDTA = "Yes"
            ElseIf TRC_EIDTA = "N" Then
                TRC_EIDTA = "No"
            Else
                TRC_EIDTA = "(Select)"
            End If
        
            If Sheet24.Cells(rowcount, SCHEI_DTAA_AmountOfIncome).Locked = False Then
                Sheet24.Cells(rowcount, SCHEI_DTAA_AmountOfIncome).value = Node("amountOfIncome")
            End If
            If Sheet24.Cells(rowcount, SCHEI_DTAA_NatureOfIncome).Locked = False Then
                Sheet24.Cells(rowcount, SCHEI_DTAA_NatureOfIncome).value = Node("natureOfIncome")
            End If
            If Sheet24.Cells(rowcount, SCHEI_DTAA_CountryCode).Locked = False Then
                Sheet24.Cells(rowcount, SCHEI_DTAA_CountryCode).value = sCountry
            End If
            If Sheet24.Cells(rowcount, SCHEI_DTAA_Article).Locked = False Then
                Sheet24.Cells(rowcount, SCHEI_DTAA_Article).value = Node("articleOfDTAA")
            End If
            If Sheet24.Cells(rowcount, SCHEI_DTAA_HeadOfIncome).Locked = False Then
                Sheet24.Cells(rowcount, SCHEI_DTAA_HeadOfIncome).value = HeadOfIncome_EIDTAA
            End If
            If Sheet24.Cells(rowcount, SCHEI_TRCObtained).Locked = False Then
                Sheet24.Cells(rowcount, SCHEI_TRCObtained).value = TRC_EIDTA
            End If
            
    Next Node
End If
End If 'For AY 2025-26_Konda_23_07_2025 SIT-97244
    If Sheet24.Range("SCHEI_Pass_Through_Income").Locked = False Then
        Sheet24.Range("SCHEI_Pass_Through_Income").value = jsonObject("lastFiledITR")("scheduleEI")("passThrIncNotChrgblTax")
    End If

End Function

Function ImportScheduleAL(jsonText As String)
On Error Resume Next

Dim jsonDictionary, init As Object
Dim Node, Nodelist, jsonObject As Object
Dim DepositsInBank, SharesAndSecurities, InsurancePolicies, LoansAndAdvancesGiven, CashInHand, JewelleryBullionEtc As Variant
Dim ArchCollDrawPaintSulpArt, VehiclYachtsBoatsAircrafts, InterstAOPFlag, ImmovableAssetFlag, LiabilityInRelatAssets As Variant
Dim DescriptionColNo, Address_FlatColNo, Address_VillageColNo, Address_RoadColNo, Address_AreaColNo, Address_CityColNo As Variant
Dim Address_StateColNo, Address_CountryColNo, Address_PinColNo, Address_ZipColNo, AmountColNo As Variant
Dim NameColNo, FlatColNo, VillageColNo, RoadColNo, AreaColNo, CityColNo, StateColNo, CountryColNo, PinColNo, ZipColNo As Variant
Dim PANColNo, InvestmentColNo As Variant
Dim TotalXMLRow, rowcount, cnt, RecAdd, RecFirm  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim sCountry, iCountry, sState, iState As Variant
Dim s1Country, i1Country, s1State, i1State As Variant
Set jsonObject = ParseJson.ParseJson(jsonText)

'Change-1, ANK, 11.07.2022
If Not jsonObject("lastFiledITR").exists("scheduleAL") Then
    Exit Function
End If
'---end change---



DepositsInBank = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("DepositsInBank")
SharesAndSecurities = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("SharesAndSecurities")
InsurancePolicies = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("InsurancePolicies")
LoansAndAdvancesGiven = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("LoansAndAdvancesGiven")
CashInHand = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("CashInHand")

'AY_2024_25 'Malli
JewelleryBullionEtc = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("JewelleryBullionEtc")
ArchCollDrawPaintSulpArt = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("ArchCollDrawPaintSulpArt")
VehiclYachtsBoatsAircrafts = jsonObject("lastFiledITR")("scheduleAL")("MovableAsset")("VehiclYachtsBoatsAircrafts")
'-----------------

ImmovableAssetFlag = jsonObject("lastFiledITR")("scheduleAL")("ImmovableAssetFlag")
InterstAOPFlag = jsonObject("lastFiledITR")("scheduleAL")("InterstAOPFlag")
LiabilityInRelatAssets = jsonObject("lastFiledITR")("scheduleAL")("LiabilityInRelatAssets")

If ImmovableAssetFlag <> "" Then
        If ImmovableAssetFlag = "Y" Then
            ImmovableAssetFlag = "Yes"
        ElseIf ImmovableAssetFlag = "N" Then
            ImmovableAssetFlag = "No"
        End If
Else
    ImmovableAssetFlag = "(Select)"
End If
    
    If Sheet43.Range("IsImmovableAsset").Locked = False Then
        Sheet43.Range("IsImmovableAsset").value = ImmovableAssetFlag
    End If
        
Set init = jsonObject("lastFiledITR")("scheduleAL")
If init.exists("ImmovableDetails") Then
  If Sheet43.Range("IsImmovableAsset").Locked = False Then
        Sheet43.Range("IsImmovableAsset").value = "Yes"
    End If
Set Nodelist = jsonObject("lastFiledITR")("scheduleAL")("ImmovableDetails")

DescriptionColNo = Sheet43.Range("SchAL.A.Description").Column
Address_FlatColNo = Sheet43.Range("SchAL.A.Address_Flat").Column
Address_VillageColNo = Sheet43.Range("SchAL.A.Address_Village").Column
Address_RoadColNo = Sheet43.Range("SchAL.A.Address_Road").Column
Address_AreaColNo = Sheet43.Range("SchAL.A.Address_Area").Column
Address_CityColNo = Sheet43.Range("SchAL.A.Address_City").Column
Address_StateColNo = Sheet43.Range("SchAL.A.Address_State").Column
Address_CountryColNo = Sheet43.Range("SchAL.A.Address_Country").Column
Address_PinColNo = Sheet43.Range("SchAL.A.Address_Pin").Column
Address_ZipColNo = Sheet43.Range("SchAL.A.Address_Zip").Column
AmountColNo = Sheet43.Range("SchAL.A.Amount").Column

TotalExRow = Range("SchAL.A.Description").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        
        If Sheet43.Range("SchAL.A.Description").Locked = False Then
            Sheet43.Range("SchAL.A.Description").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_Flat").Locked = False Then
            Sheet43.Range("SchAL.A.Address_Flat").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_Village").Locked = False Then
            Sheet43.Range("SchAL.A.Address_Village").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_Road").Locked = False Then
            Sheet43.Range("SchAL.A.Address_Road").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_Area").Locked = False Then
            Sheet43.Range("SchAL.A.Address_Area").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_City").Locked = False Then
            Sheet43.Range("SchAL.A.Address_City").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_State").Locked = False Then
            Sheet43.Range("SchAL.A.Address_State").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_Country").Locked = False Then
            Sheet43.Range("SchAL.A.Address_Country").ClearContents
        End If
        If Sheet43.Range("SchAL.A.Address_Pin").Locked = False Then
            Sheet43.Range("SchAL.A.Address_Pin").ClearContents
        End If
        'if         Sheet43.Range("SchAL.A.Address_Zip").locked=false then
        '        Sheet43.Range("SchAL.A.Address_Zip").ClearContents
        'end if
        If Sheet43.Range("SchAL.A.Amount").Locked = False Then
            Sheet43.Range("SchAL.A.Amount").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_ImmvblDetails (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet43.Range("SchAL.A.Description").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            
        iCountry = Node("AddressAL")("CountryCode")
        sCountry = Findtext(CStr(iCountry), "Country")
        
        iState = Node("AddressAL")("StateCode")
        sState = Findtext(CStr(iState), "State")
            
            If Sheet43.Cells(rowcount, DescriptionColNo).Locked = False Then
                Sheet43.Cells(rowcount, DescriptionColNo).value = Node("Description")
            End If
            If Sheet43.Cells(rowcount, Address_FlatColNo).Locked = False Then
                Sheet43.Cells(rowcount, Address_FlatColNo).value = Node("AddressAL")("ResidenceNo")
            End If
            If Sheet43.Cells(rowcount, Address_VillageColNo).Locked = False Then
                Sheet43.Cells(rowcount, Address_VillageColNo).value = Node("AddressAL")("ResidenceName")
            End If
            If Sheet43.Cells(rowcount, Address_RoadColNo).Locked = False Then
                Sheet43.Cells(rowcount, Address_RoadColNo).value = Node("AddressAL")("RoadOrStreet")
            End If
            If Sheet43.Cells(rowcount, Address_AreaColNo).Locked = False Then
                Sheet43.Cells(rowcount, Address_AreaColNo).value = Node("AddressAL")("LocalityOrArea")
            End If
            If Sheet43.Cells(rowcount, Address_CityColNo).Locked = False Then
                Sheet43.Cells(rowcount, Address_CityColNo).value = Node("AddressAL")("CityOrTownOrDistrict")
            End If
        If iCountry = "91" Then
            If sCountry <> "" Then
                Sheet43.Cells(rowcount, Address_CountryColNo).value = sCountry
                iCountry = ""
                sCountry = ""
            End If
            If sState <> "" Then
                Sheet43.Cells(rowcount, Address_StateColNo).value = sState
              iState = ""
              sState = ""
            End If
        Else
            If sState <> "" Then
                Sheet43.Cells(rowcount, Address_StateColNo).value = sState
              iState = ""
              sState = ""
            End If
            If sCountry <> "" Then
                Sheet43.Cells(rowcount, Address_CountryColNo).value = sCountry
                iCountry = ""
                sCountry = ""
            End If
            
        End If
        
        
        
        If Sheet43.Cells(rowcount, Address_PinColNo).Locked = False Then
            Sheet43.Cells(rowcount, Address_PinColNo).value = Node("AddressAL")("PinCode")
        End If
        
        If (Node("AddressAL")("ZipCode")) <> "0" And Sheet43.Cells(rowcount, Address_ZipColNo).Locked = False Then
            Sheet43.Cells(rowcount, Address_ZipColNo).value = Node("AddressAL")("ZipCode")
        End If
        
        If Sheet43.Cells(rowcount, AmountColNo).Locked = False Then
            Sheet43.Cells(rowcount, AmountColNo).value = Node("Amount")
        End If
        
cnt = cnt + 1
Next Node
RecAdd = cnt
End If
'Change-2, ANK-129, 26.05.2022
        If DepositsInBank <> "" And Sheet43.Range("SchAL.B.Bank").Locked = False Then
            'UCase(DepositsInBank_AL) = DepositsInBank
'            Range("SchAL.B.Bank").value = DepositsInBank
            Sheet43.Range("SchAL.B.Bank").value = DepositsInBank
        End If
        
        If SharesAndSecurities <> "" And Sheet43.Range("SchAL.B.Share").Locked = False Then
'            Range("SchAL.B.Share").value = SharesAndSecurities
            Sheet43.Range("SchAL.B.Share").value = SharesAndSecurities
        End If
        
        If InsurancePolicies <> "" And Sheet43.Range("SchAL.B.Insurance").Locked = False Then
'            Range("SchAL.B.Insurance").value = InsurancePolicies
            Sheet43.Range("SchAL.B.Insurance").value = InsurancePolicies
        End If
        
        If LoansAndAdvancesGiven <> "" And Sheet43.Range("SchAL.b.Loans").Locked = False Then
'             Range("SchAL.b.Loans").value = LoansAndAdvancesGiven
            Sheet43.Range("SchAL.B.Loans").value = LoansAndAdvancesGiven
        End If
        
        If CashInHand <> "" And Sheet43.Range("SchAL.b.Cash").Locked = False Then
'            Range("SchAL.b.Cash").value = CashInHand
            Sheet43.Range("SchAL.B.Cash").value = CashInHand
        End If
        
       ' AY_2024_25 Malli
        If JewelleryBullionEtc <> "" And Sheet43.Range("SchAL.B.Jwellery").Locked = False Then
'            Range("SchAL.B.Jwellery").value = JewelleryBullionEtc
            Sheet43.Range("SchAL.B.Jwellery").value = JewelleryBullionEtc
        End If

        If ArchCollDrawPaintSulpArt <> "" And Sheet43.Range("SchAL.B.Art").Locked = False Then
'            Range("SchAL.B.Art").value = ArchCollDrawPaintSulpArt
            Sheet43.Range("SchAL.B.Art").value = ArchCollDrawPaintSulpArt
        End If

        If VehiclYachtsBoatsAircrafts <> "" And Sheet43.Range("SchAL.B.Vehicles").Locked = False Then
'            Range("SchAL.B.Vehicles").value = VehiclYachtsBoatsAircrafts
            Sheet43.Range("SchAL.B.Vehicles").value = VehiclYachtsBoatsAircrafts
        End If
        
        '---------------------------------------
'---end change---
        
'If InterstAOPFlag <> "" Then
'        If InterstAOPFlag = "Y" Then
'            InterstAOPFlag = "Yes"
'        ElseIf InterstAOPFlag = "N" Then
'            InterstAOPFlag = "No"
'        End If
'Else
'    InterstAOPFlag = "(Select)"
'End If
'If Sheet43.Range("IsInterestHeld").Locked = False Then
'    Sheet43.Range("IsInterestHeld").value = InterstAOPFlag
'End If
'Set init = jsonObject("lastFiledITR")("scheduleAL")
'If init.exists("InterestHeldInaAsset") Then
'
'Set NodeList = jsonObject("lastFiledITR")("scheduleAL")("InterestHeldInaAsset")
'
'NameColNo = Sheet43.Range("SchAL.C.Name").Column
'FlatColNo = Sheet43.Range("SchAL.C.Address_Flat").Column
'VillageColNo = Sheet43.Range("SchAL.C.Address_Village").Column
'RoadColNo = Sheet43.Range("SchAL.C.Address_Road").Column
'AreaColNo = Sheet43.Range("SchAL.C.Address_Area").Column
'CityColNo = Sheet43.Range("SchAL.C.Address_City").Column
'StateColNo = Sheet43.Range("SchAL.C.Address_State").Column
'CountryColNo = Sheet43.Range("SchAL.C.Address_Country").Column
'PinColNo = Sheet43.Range("SchAL.C.Address_Pin").Column
'ZipColNo = Sheet43.Range("SchAL.C.Address_Zip").Column
'
'PANColNo = Sheet43.Range("SchAL.C.PAN").Column
'InvestmentColNo = Sheet43.Range("SchAL.C.Investment").Column
'
'
'TotalExRow = Range("SchAL.C.Name").Rows.count
'
'    TotalXMLRow = NodeList.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'
'        If Sheet43.Range("SchAL.C.Name").Locked = False Then
'            Sheet43.Range("SchAL.C.Name").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_Flat").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_Flat").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_Village").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_Village").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_Road").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_Road").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_Area").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_Area").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_City").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_City").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_State").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_State").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_Country").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_Country").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Address_Pin").Locked = False Then
'            Sheet43.Range("SchAL.C.Address_Pin").ClearContents
'        End If
'        'if Sheet43.Range("SchAL.C.Address_Zip").locked=false then
'        '    Sheet43.Range("SchAL.C.Address_Zip").ClearContents
'        'end if
'        If Sheet43.Range("SchAL.C.PAN").Locked = False Then
'            Sheet43.Range("SchAL.C.PAN").ClearContents
'        End If
'        If Sheet43.Range("SchAL.C.Investment").Locked = False Then
'            Sheet43.Range("SchAL.C.Investment").ClearContents
'        End If
'
'    End If
'
'    If (TotalDiffRow > 0) Then
'        AddDiffRows_IntrstInAsst (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet43.Range("SchAL.C.Name").name)
'    rowcount = rowcount - 1
'    cnt = 0
'
'    For Each node In NodeList
'        rowcount = rowcount + 1
'
'         i1Country = node("AddressAL")("CountryCode")
'         s1Country = Findtext(CStr(i1Country), "Country")
'
'         i1State = node("AddressAL")("StateCode")
'         s1State = Findtext(CStr(i1State), "State")
'
'            If Sheet43.Cells(rowcount, NameColNo).Locked = False Then
'                Sheet43.Cells(rowcount, NameColNo).value = node("NameOfFirm")
'            End If
'            If Sheet43.Cells(rowcount, FlatColNo).Locked = False Then
'                Sheet43.Cells(rowcount, FlatColNo).value = node("AddressAL")("ResidenceNo")
'            End If
'            If Sheet43.Cells(rowcount, VillageColNo).Locked = False Then
'                Sheet43.Cells(rowcount, VillageColNo).value = node("AddressAL")("ResidenceName")
'            End If
'            If Sheet43.Cells(rowcount, RoadColNo).Locked = False Then
'                Sheet43.Cells(rowcount, RoadColNo).value = node("AddressAL")("RoadOrStreet")
'            End If
'            If Sheet43.Cells(rowcount, AreaColNo).Locked = False Then
'                Sheet43.Cells(rowcount, AreaColNo).value = node("AddressAL")("LocalityOrArea")
'            End If
'            If Sheet43.Cells(rowcount, CityColNo).Locked = False Then
'                Sheet43.Cells(rowcount, CityColNo).value = node("AddressAL")("CityOrTownOrDistrict")
'            End If
'            If i1Country = "91" Then
'                If s1Country <> "" Then
'                    Sheet43.Cells(rowcount, CountryColNo).value = s1Country
'                End If
'                If s1State <> "" Then
'                    Sheet43.Cells(rowcount, StateColNo).value = s1State
'                End If
'            Else
'                If s1State <> "" Then
'                    Sheet43.Cells(rowcount, StateColNo).value = s1State
'                End If
'                If s1Country <> "" Then
'                    Sheet43.Cells(rowcount, CountryColNo).value = s1Country
'                End If
'
'            End If
'
'        If Sheet43.Cells(rowcount, PinColNo).Locked = False Then
'            Sheet43.Cells(rowcount, PinColNo).value = node("AddressAL")("PinCode")
'        End If
'
'        If (node("AddressAL")("ZipCode")) <> 0 And Sheet43.Cells(rowcount, ZipColNo).Locked = False Then
'            Sheet43.Cells(rowcount, ZipColNo).value = node("AddressAL")("ZipCode")
'        End If
'
'        If Sheet43.Cells(rowcount, PANColNo).Locked = False Then
'            Sheet43.Cells(rowcount, PANColNo).value = node("PanOfFirm")
'        End If
'        If Sheet43.Cells(rowcount, InvestmentColNo).Locked = False Then
'            Sheet43.Cells(rowcount, InvestmentColNo).value = node("AssesseInvestment")
'        End If
'cnt = cnt + 1
'Next node
'RecFirm = cnt
'End If
If LiabilityInRelatAssets <> "" And Range("SchAL.Liability").Locked = False Then
    Range("SchAL.Liability").value = LiabilityInRelatAssets
End If

End Function
Function ImportScheduleFSI(jsonText As String)

On Error Resume Next

Dim jsonDictionary, jsonObject As Object
Dim Node, Nodelist, init As Object

Dim MobileNoSec, emailAddress, emailAddressSec, dob, Status1, aadhaarCardNo, AadhaarEnrolmentId, Country As Variant
Dim sCountry, iCountry As Variant
Dim sState, iState As Variant
Dim YYYY, MM, DD, strdate, TotalExRow As String
Dim TotalXMLRow, RecIncOut, rowcount, cnt, TotalDiffRow, i As Long
i = 0
Set jsonObject = ParseJson.ParseJson(jsonText)

Set init = jsonObject("form67")("scheduleFSI")
If init.exists("scheduleFSIDtls") Then
 
        Set Nodelist = jsonObject("form67")("scheduleFSI")("scheduleFSIDtls")
    
        TotalExRow = Range("FSI_Count").value
        
        TotalXMLRow = Nodelist.count
        
        TotalDiffRow = TotalXMLRow - TotalExRow
        Sheet40.Activate
'Change-2, ANK, 02.06.2022
        If (TotalDiffRow > 0) Then
            Dim k As Variant
            For k = 1 To TotalDiffRow
                AddBlockCall_FSIfrm
            Next k
        End If
'---
    
        cnt = 0
        For Each Node In Nodelist
'Change-5, ANK-133, 26.05.2022
                i = i + 1
                iCountry = Node("countryCode")
                sCountry = Findtext(CStr(iCountry), "Country")
                Country = Mid(sCountry, InStr(1, sCountry, "-") + 1) & ":" & Node("countryCode")
'                Country = node("countryCode") & ":" & node("countryCode")
'---end change---
             If Sheet40.Range("FSI_CountryCode" & i).Locked = False Then
                Sheet40.Range("FSI_CountryCode" & i).value = Country
            End If
                               
'Change-6, ANK-134, ANK-135, ANK-136, ANK-137, ANK-138,, 26.05.2022
'            If Sheet40.Range("FSI_IncFromSal" & i).Locked = False Then
'                Sheet40.Range("FSI_IncFromSal" & i).value = node("incFromSal")
'            End If
'            If Sheet40.Range("FSI_IncFromHP" & i).Locked = False Then
'                Sheet40.Range("FSI_IncFromHP" & i).value = node("incFromHP")
'            End If
'             If Sheet40.Range("FSI_IncFromBP" & i).Locked = False Then
'                Sheet40.Range("FSI_IncFromBP" & i).value = node("incFromBusiness")
'            End If
'            If Sheet40.Range("FSI_IncCapGain" & i).Locked = False Then
'                Sheet40.Range("FSI_IncCapGain" & i).value = node("incCapGain")
'            End If
'            If Sheet40.Range("FSI_IncOthSrc" & i).Locked = False Then
'                Sheet40.Range("FSI_IncOthSrc" & i).value = node("incOthSrc")
'            End If


            If Sheet40.Range("FSI_IncFromSal" & i).Locked = False Then
                Sheet40.Range("FSI_IncFromSal" & i).value = Node("incFromSal")("incFrmOutsideInd")
            End If
            If Sheet40.Range("FSI_IncFromHP" & i).Locked = False Then
                Sheet40.Range("FSI_IncFromHP" & i).value = Node("incFromHP")("incFrmOutsideInd")
            End If
'             If Sheet40.Range("FSI_IncFromBP" & i).Locked = False Then
'                Sheet40.Range("FSI_IncFromBP" & i).value = node("incFromBusiness")("incFrmOutsideInd")
'            End If
            If Sheet40.Range("FSI_IncCapGain" & i).Locked = False Then
                Sheet40.Range("FSI_IncCapGain" & i).value = Node("incCapGain")("incFrmOutsideInd")
            End If
            If Sheet40.Range("FSI_IncOthSrc" & i).Locked = False Then
                Sheet40.Range("FSI_IncOthSrc" & i).value = Node("incOthSrc")("incFrmOutsideInd")
            End If
            
            If Sheet40.Range("FSI_TXNOSAL" & i).Locked = False Then
                Sheet40.Range("FSI_TXNOSAL" & i).value = Node("incFromSal")("taxPaidOutsideInd")
            End If
            If Sheet40.Range("FSI_TXNOHP" & i).Locked = False Then
                Sheet40.Range("FSI_TXNOHP" & i).value = Node("incFromHP")("taxPaidOutsideInd")
            End If
'             If Sheet40.Range("FSI_TXNOBP" & i).Locked = False Then
'                Sheet40.Range("FSI_TXNOBP" & i).value = node("incFromBusiness")("taxPaidOutsideInd")
'            End If
            If Sheet40.Range("FSI_TXNOCapGain" & i).Locked = False Then
                Sheet40.Range("FSI_TXNOCapGain" & i).value = Node("incCapGain")("taxPaidOutsideInd")
            End If
            If Sheet40.Range("FSI_TXNOOthSrc" & i).Locked = False Then
                Sheet40.Range("FSI_TXNOOthSrc" & i).value = Node("incOthSrc")("taxPaidOutsideInd")
            End If
            
            If Sheet40.Range("FSI_TXNISAL" & i).Locked = False Then
                Sheet40.Range("FSI_TXNISAL" & i).value = Node("incFromSal")("taxPayableinInd")
            End If
            If Sheet40.Range("FSI_TXNIHP" & i).Locked = False Then
                Sheet40.Range("FSI_TXNIHP" & i).value = Node("incFromHP")("taxPayableinInd")
            End If
'             If Sheet40.Range("FSI_TXNIBP" & i).Locked = False Then
'                Sheet40.Range("FSI_TXNIBP" & i).value = node("incFromBusiness")("taxPayableinInd")
'            End If
            If Sheet40.Range("FSI_TXnICapGain" & i).Locked = False Then
                Sheet40.Range("FSI_TXnICapGain" & i).value = Node("incCapGain")("taxPayableinInd")
            End If
            If Sheet40.Range("FSI_TXNIOthSrc" & i).Locked = False Then
                Sheet40.Range("FSI_TXNIOthSrc" & i).value = Node("incOthSrc")("taxPayableinInd")
            End If
            
            If Sheet40.Range("FSI_DTAASAL" & i).Locked = False Then
                Sheet40.Range("FSI_DTAASAL" & i).value = Node("incFromSal")("dtaaReliefUs90Or90A")
            End If
            If Sheet40.Range("FSI_DTAAHP" & i).Locked = False Then
                Sheet40.Range("FSI_DTAAHP" & i).value = Node("incFromHP")("dtaaReliefUs90Or90A")
            End If
'             If Sheet40.Range("FSI_DTAABP" & i).Locked = False Then
'                Sheet40.Range("FSI_DTAABP" & i).value = node("incFromBusiness")("dtaaReliefUs90Or90A")
'            End If
            If Sheet40.Range("FSI_DTAACapGain" & i).Locked = False Then
                Sheet40.Range("FSI_DTAACapGain" & i).value = Node("incCapGain")("dtaaReliefUs90Or90A")
            End If
            If Sheet40.Range("FSI_DTAAOthSrc" & i).Locked = False Then
                Sheet40.Range("FSI_DTAAOthSrc" & i).value = Node("incOthSrc")("dtaaReliefUs90Or90A")
            End If
            
'             If Sheet40.Range("FSI_IncFromSal" & i).Locked = False Then
'                Sheet40.Range("FSI_IncFromSal" & i).value = node("incFromSal")("taxPaidOutsideInd")
'            End If
'            If Sheet40.Range("FSI_TXNISAL" & i).Locked = False Then
'                Sheet40.Range("FSI_TXNISAL" & i).value = node("incFromSal")("taxPayableinInd")
'            End If
'            If Sheet40.Range("FSI_DTAASAL" & i).Locked = False Then
'                Sheet40.Range("FSI_DTAASAL" & i).value = node("incFromSal")("dtaaReliefUs90Or90A")
'            End If
'---end change---
            cnt = cnt + 1
        
        Next Node
    RecIncOut = cnt
End If
End Function

Function ImportScheduleOS(jsonText As String)
On Error Resume Next
Dim jsonObject, init, Nodelist, Node As Object

Dim othDescColNo, OtherSectionsColNo, TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, cnt, RecOthInc As Variant
Set jsonObject = ParseJson.ParseJson(jsonText)

Set init = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("othersInc")
If init.exists("othersIncDtls") Then

Set Nodelist = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("othersInc")("othersIncDtls")

    othDescColNo = Sheet14.Range("os.othDesc").Column
    OtherSectionsColNo = Sheet14.Range("os.OtherSections").Column

TotalExRow = Sheet14.Range("os.OtherSections").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        If Sheet14.Range("os.othDesc").Locked = False Then
            Sheet14.Range("os.othDesc").ClearContents
        End If
        If Sheet14.Range("os.OtherSections").Locked = False Then
            Sheet14.Range("os.OtherSections").ClearContents
        End If
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_OthInc (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet14.Range("os.othDesc").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each Node In Nodelist
        rowcount = rowcount + 1

    If Sheet14.Cells(rowcount, othDescColNo).Locked = False Then
        Sheet14.Cells(rowcount, othDescColNo).value = Node("othNatOfInc")
    End If
    If Sheet14.Cells(rowcount, OtherSectionsColNo).Locked = False Then
        Sheet14.Cells(rowcount, OtherSectionsColNo).value = Node("othAmount")
    End If

cnt = cnt + 1
Next Node
RecOthInc = cnt
End If

Set init = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")
If init.exists("OthersGrossDtls") Then

Set Nodelist = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("OthersGrossDtls")

    othDescColNo = Sheet14.Range("os.SourceDescription").Column
    OtherSectionsColNo = Sheet14.Range("os.SourceAmount").Column

TotalExRow = Sheet14.Range("os.SourceDescription").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        If Sheet14.Range("os.SourceDescription").Locked = False Then
            Sheet14.Range("os.SourceDescription").ClearContents
        End If
        If Sheet14.Range("os.SourceAmount").Locked = False Then
            Sheet14.Range("os.SourceAmount").ClearContents
        End If
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_OthInc (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet14.Range("os.SourceDescription").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each Node In Nodelist
        rowcount = rowcount + 1

   'added by Chetan C M on 22/09/2025 for SIT-94128(issue2)
    'start--
        Dim iNature As Variant
        
        iNature = Node("SourceDescription")
                
        If iNature = "5A1ai" Then
            iNature = "115A(1)(a)(i)- Dividends interest and income from units purchase in foreign currency"
        ElseIf iNature = "5A1aA" Then
            iNature = "115A(1)(a)(A)- Dividend in the case of non-resident received from a unit in an International Financial  Services Centre"
        ElseIf iNature = "5A1aii" Then
            iNature = "115A(1)(a)(ii)- Interest received from govt/Indian Concerns received in Foreign Currency"
        ElseIf iNature = "5A1aiia" Then
            iNature = "115A(1) (a)(iia) -Interest from Infrastructure Debt Fund"
        
        'Malli_24/04/2026
        ElseIf iNature = "5A1aiiaa" Then
            'iNature = "115A(1) (a)(iiaa) -Interest as per Sec. 194LC(1)"
            iNature = "5A1aiiaa - 115A(1) (a)(iiaa) - Interest as referred in proviso to section 194LC(1)"
         '----------------------------------
         
        ElseIf iNature = "5A1aiiab" Then
            iNature = "115A(1) (a)(iiab) -Interest as per Sec. 194LD"
        ElseIf iNature = "5A1aiiac" Then
            iNature = "115A(1) (a)(iiac) -Interest as per Sec. 194LBA"
        ElseIf iNature = "5A1aiii" Then
            iNature = "115A(1) (a)(iii) -Income received in respect of units of UTI purchased in foreign currency"
        ElseIf iNature = "5A1bA" Then
            iNature = "115A(1)(b)(A) & 115A(1)(b)(B)- Income from royalty or fees for technical services received from Government or Indian concern"
        ElseIf iNature = "5AC1ab" Then
            iNature = "115AC(1)(a) - Income by way of interest on bonds purchased in foreign currency - non-resident"
        ElseIf iNature = "5AC1abD" Then
            iNature = "115AC(1)(b) - Income by way of Dividend on GDRs purchased in foreign currency - non-resident"
        ElseIf iNature = "5ACA1a" Then
            iNature = "115ACA(1)(a) - Income from GDR purchased in foreign currency -resident"
        ElseIf iNature = "5AD1i" Then
            iNature = "115AD(1)(i) -Income (other than Dividend) received by an FII in respect of securities (other than units as per Sec 115AB)"
        ElseIf iNature = "5AD1iP" Then
            iNature = "115AD(1)(i) -Income received by an FII in respect of bonds or government securities as per Sec 194LD"
        ElseIf iNature = "5BBA" Then
            iNature = "115BBA - Tax on non-residents sportsmen or sports associations"
        ElseIf iNature = "5BBC" Then
            iNature = "115BBC - Anonymous donations"
        ElseIf iNature = "5BBF" Then
            iNature = "115BBF - Tax on income from patent"
        ElseIf iNature = "5BBG" Then
            iNature = "115BBG - Tax on income from transfer of carbon credits"
        ElseIf iNature = "5Ea" Then
            iNature = "115E(a) - Investment income"
            
          'Malli_24/04/2026
        ElseIf iNature = "5A1aiiaaP" Then
            'iNature = "115A(1) (a)(iiaa) -Interest as referred in proviso to section 194LC(1)"
            iNature = "5A1aiiaaP-Income received by non-resident as referred in proviso to section 194LC(1)"
            '----------------------
            
        'Malli_24/04/2026  V0.3 changes DE sheet Updated
        ElseIf iNature = "5A1aiiaa2P" Then
            iNature = "5A1aiiaa2P-Income received by non-resident as referred in second proviso to section 194LC(1)"
        '-----------------------------------------------
        
        ElseIf iNature = "5AD1iDiv" Then
            iNature = "115AD(1)(i) - Income (being dividend) received by an FII in respect of securities (other than units referred to in section 115AB)"

        Else
            iNature = "(Select)"
        End If
    
    '--end
    
    If Sheet14.Cells(rowcount, othDescColNo).Locked = False Then
    'changed by Chetan C M on 22/09/2025 for SIT-94128(issue2)
    'start--
'        Sheet14.Cells(rowcount, othDescColNo).value = node("SourceDescription")
         Sheet14.Cells(rowcount, othDescColNo).value = iNature
    '--end
    End If
    
    If Sheet14.Cells(rowcount, OtherSectionsColNo).Locked = False Then
        Sheet14.Cells(rowcount, OtherSectionsColNo).value = Node("SourceAmount")
    End If

cnt = cnt + 1
Next Node
RecOthInc = cnt
End If

'If Sheet14.Range("os.DividendGross").Locked = False And jsonObject("form24q")("scheduleOS")("incOthThanOwnRaceHorse")("dividendGross") <> "" Then
'          Sheet14.Range("os.DividendGross").value = jsonObject("form24q")("scheduleOS")("incOthThanOwnRaceHorse")("dividendGross")
'          End If
          If Sheet14.Range("os.DividendGrossaii").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("DividendOthThan22e") <> "" Then
          Sheet14.Range("os.DividendGrossaii").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("DividendOthThan22e")
          End If
          If Sheet14.Range("os.SavingBank").Locked = False And jsonObject("form24q")("intrstFrmSavingBank") <> "" Then
          Sheet14.Range("os.SavingBank").value = jsonObject("form24q")("intrstFrmSavingBank")
          End If
          If Sheet14.Range("os.TermDeposit").Locked = False And jsonObject("form26as")("intrstFrmTermDeposit") <> "" Then
          Sheet14.Range("os.TermDeposit").value = jsonObject("form26as")("intrstFrmTermDeposit")
          End If
          
     'Malli commented----------SIT_93728
'          If Sheet14.Range("os.IncomeTaxRefund").Locked = False And jsonObject("incDeductionsOthIncCPC")("othSrcOthAmount") <> "" Then
'          Sheet14.Range("os.IncomeTaxRefund").value = jsonObject("incDeductionsOthIncCPC")("othSrcOthAmount")
'          End If
          
           'Malli<<<<<<<<<<<<<<<<<
        If Sheet14.Range("os.IncomeTaxRefund").Locked = False Then
           incDeductionsOthIncCPC_TAX_chk = False
           
            Set init = jsonObject("form26as")    '//*Primary source//
            If init <> Empty Then
            If init.exists("incomeDeductionsOthersInc") Then
                 Set Nodelist = jsonObject("form26as")("incomeDeductionsOthersInc")
                            For Each Node In Nodelist
                                    If UCase(Node("othSrcNatureDesc")) = "TAX" Then
                                         Sheet14.Range("os.IncomeTaxRefund").value = Node("othSrcOthAmount")
                                         incDeductionsOthIncCPC_TAX_chk = True
                                    End If
                                
                            Next Node
            End If
            End If
           
           
            If incDeductionsOthIncCPC_TAX_chk <> True Then
                     Set init = jsonObject("incDeductionsOthIncCPC")   '//*secoundary source//
                         If init <> Empty Then
                             Set Nodelist = jsonObject("incDeductionsOthIncCPC")
                             For Each Node In Nodelist
                                     If UCase(Node("othSrcNatureDesc")) = "TAX" Then
                                          Sheet14.Range("os.IncomeTaxRefund").value = Node("othSrcOthAmount")
                                     End If
                                  
                             Next Node
                         End If
            End If
        End If
   
          '>>>>>>>>>>>>>>>>>>>>>>
          
          '------------------------------------------------------
          
          
          If Sheet14.Range("os.Others").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("intrstFrmOthers") <> "" Then
          Sheet14.Range("os.Others").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("intrstFrmOthers")
          End If
          If Sheet14.Range("os.RentFromMachPlantBldgs").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("rentFromMachPlantBldgs") <> "" Then
          Sheet14.Range("os.RentFromMachPlantBldgs").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("rentFromMachPlantBldgs")
          End If
          If Sheet14.Range("os.WinLottRacePuzz").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("ltryPzzlChrgblUs115BB") <> "" Then
          Sheet14.Range("os.WinLottRacePuzz").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("ltryPzzlChrgblUs115BB")
          End If
          If Sheet14.Range("os.Receipts").Locked = False And jsonObject("form26as")("scheduleOS")("incFromOwnHorse")("receipts") <> "" Then
          Sheet14.Range("os.Receipts").value = jsonObject("form26as")("scheduleOS")("incFromOwnHorse")("receipts")
          End If


'
'''AY_2024_25 Add Malli
'          If Sheet14.Range("os.WinOnlineGame").Locked = False And jsonObject("form26as")("scheduleOS")("IncOthThanOwnRaceHorse")("IncChrgblUs115BBJ") <> "" Then
'          Sheet14.Range("os.WinOnlineGame").value = jsonObject("form26as")("scheduleOS")("IncOthThanOwnRaceHorse")("IncChrgblUs115BBJ")
'          End If

''Newly added by Malli 'AY_2025_26 OS_DE036 this field is autopopulated
'AY_2024_25 Add Malli
'          If Sheet14.Range("os.WinOnlineGame").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("IncChrgblUs115BBJ") <> "" Then
'          'Sheet14.Range("os.WinOnlineGame").value = jsonObject("form26as")("scheduleOS")("IncOthThanOwnRaceHorse")("IncChrgblUs115BBJ")
'          Sheet14.Range("os.WinOnlineGame").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("IncChrgblUs115BBJ")
'          End If

'form26as.ScheduleOS.IncOthThanOwnRaceHorse.IncChrgblUs115BBJWin
'AY_2025_26 Malli_OS_C34_Prefill
'          If Sheet14.Range("os.WinOnlineGame2aii_i").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("IncChrgblUs115BBJWin") <> "" Then
'                Sheet14.Range("os.WinOnlineGame2aii_i").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("IncChrgblUs115BBJWin")
'          End If
'-------------------------------
'--------------------------------

'Malli----------AY_2025_26_ITR-2_PrefillSRS_V0.3_29/05/2025
 'Shameema confirmed for this year , the tag was changed  IncChrgblUs115BBJ to IncChrgblUs115BBJWin is the latest
          If Sheet14.Range("os.WinOnlineGame").Locked = False And jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("IncChrgblUs115BBJWin") <> "" Then
                Sheet14.Range("os.WinOnlineGame").value = jsonObject("form26as")("scheduleOS")("incOthThanOwnRaceHorse")("IncChrgblUs115BBJWin")
          End If
'-------------------------------

End Function
Function AMTCImport(jsonText As String)
On Error Resume Next
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, cnt As Long
    Dim Nodelist, jsonObject, init As Object
    Dim Node As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim AssYr As Variant
    Dim TaxSection115JC, TaxOthProvisions, AmtTaxCreditAvailable
Set jsonObject = ParseJson.ParseJson(jsonText)
    'TaxSection115JC = jsonObject("ScheduleAMTC")("TaxSection115JC")
    'TaxOthProvisions = jsonObject("ScheduleAMTC")("TaxOthProvisions")
    'AmtTaxCreditAvailable = jsonObject("ScheduleAMTC")("AmtTaxCreditAvailable")

Set init = jsonObject("lastFiledITR")("scheduleAMTC")
If init.exists("scheduleAMTCDtls") Then

    Set Nodelist = jsonObject("lastFiledITR")("scheduleAMTC")("scheduleAMTCDtls")
    cnt = 0
    For Each Node In Nodelist
            AssYr = Node("assYr")
            
            '<<<Malli_SIT-112543 13/04/2026
            '* the respective column AMT Credit Utilised during the Current Assessment Year for prefill posibility ia No, so code comented
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            If AssYr = "2013-14" Then
                If Sheet23.Range("AMTC.AmtCreditFwd1").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd1").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd1").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy1").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy1").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized1").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized1").value = Node("gross")
'                End If
            End If
            If AssYr = "2014-15" Then
                If Sheet23.Range("AMTC.AmtCreditFwd2").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd2").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd2").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy2").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy2").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized2").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized2").value = Node("gross")
'                End If
            End If
            If AssYr = "2015-16" Then
                If Sheet23.Range("AMTC.AmtCreditFwd3").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd3").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd3").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy3").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy3").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized3").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized3").value = Node("gross")
'                End If
            End If
            If AssYr = "2016-17" Then
                If Sheet23.Range("AMTC.AmtCreditFwd").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized").value = Node("gross")
'                End If
            End If
            If AssYr = "2017-18" Then
                If Sheet23.Range("AMTC.AmtCreditFwd5").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd5").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd5").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy5").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy5").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized5").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized5").value = Node("gross")
'                End If
            End If
            If AssYr = "2018-19" Then
                If Sheet23.Range("AMTC.AmtCreditFwd4").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd4").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd4").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy4").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy4").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized4").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized4").value = Node("gross")
'                End If
            End If
            If AssYr = "2019-20" Then
                If Sheet23.Range("AMTC.AmtCreditFwd6").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd6").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd6").value = Node("gross")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized6").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized6").value = Node("amtCreditUtilized")
'                End If

                If Sheet23.Range("AMTC.AmtCreditSetOfEy6").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy6").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized6").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized6").value = Node("gross")
'                End If
            End If
            
            If AssYr = "2020-21" Then
                If Sheet23.Range("AMTC.AmtCreditFwd7").Locked = False Then
                   ' Sheet23.Range("AMTC.AmtCreditFwd7").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd7").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy7").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditUtilized7").value = Node("amtCreditUtilized")
                    Sheet23.Range("AMTC.AmtCreditSetOfEy7").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized7").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized7").value = Node("gross")
'                End If
            End If
            If AssYr = "2021-22" Then
                If Sheet23.Range("AMTC.AmtCreditFwd8").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditFwd8").value = Node("amtCreditFwd")
                    Sheet23.Range("AMTC.AmtCreditFwd8").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy8").Locked = False Then
                    'Sheet23.Range("AMTC.AmtCreditUtilized8").value = Node("amtCreditUtilized")
                    Sheet23.Range("AMTC.AmtCreditSetOfEy8").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized8").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized8").value = Node("gross")
'                End If
            End If
            
            'AY_2024_25 Malli
            
            If AssYr = "2022-23" Then
                If Sheet23.Range("AMTC.AmtCreditFwd9").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditFwd9").value = Node("gross")
                End If
                If Sheet23.Range("AMTC.AmtCreditSetOfEy9").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy9").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized9").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized9").value = Node("amtCreditFwd")
'                End If
            End If
            
            
            If AssYr = "2023-24" Then
                If Sheet23.Range("AMTC.AmtCreditFwd10").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditFwd10").value = Node("gross")
                End If
                'MalliAY_2025-26
                If Sheet23.Range("AMTC.AmtCreditSetOfEy10").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy10").value = Node("amtCreditSetOfEy")
                End If
'                If Sheet23.Range("AMTC.AmtCreditUtilized10").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized10").value = Node("amtCreditFwd")
'                End If
            End If
            
            'Malli_AY_2025_26
            
            If AssYr = "2024-25" Then
                If Sheet23.Range("AMTC.AmtCreditFwd11").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditFwd11").value = Node("gross")
                End If
                'Malli AMTC_C78_AY_2026_27
                If Sheet23.Range("AMTC.AmtCreditSetOfEy11").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditSetOfEy11").value = Node("amtCreditSetOfEy")
                End If
                
'                If Sheet23.Range("AMTC.AmtCreditUtilized11").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized11").value = Node("amtCreditFwd")
'                End If
            End If
            
            
            'Malli_03/02/2026  AMTC_C83,84,86_AY_2026_27
            
            If AssYr = "2025-26" Then
                If Sheet23.Range("AMTC.AmtCreditFwd12").Locked = False Then
                    Sheet23.Range("AMTC.AmtCreditFwd12").value = Node("gross")
                End If
'                If Sheet23.Range("AMTC.AmtCreditSetOfEy12").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditSetOfEy12").value = Node("amtCreditSetOfEy")
'                End If

'                If Sheet23.Range("AMTC.AmtCreditUtilized12").Locked = False Then
'                    Sheet23.Range("AMTC.AmtCreditUtilized12").value = Node("amtCreditFwd")
'                End If
            End If
            
            '-------------------------------------------
            
            
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End If
End Function
Function ImportScheduleHP(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2, Nodelist3, init As Object
    Dim Node, Node1, Node2, Node3 As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant
    Dim iTotal24BRow, sTotal24BRow, TotalDiff24BRow As Variant
    Dim rowcount, cnt As Long
    
    
    '03/02/2026
 Dim ais_Tenchk_pfl As Boolean
 ais_Tenchk_pfl = False
 '--------------------
    
    
    
Set jsonObject = ParseJson.ParseJson(jsonText)

'Malli------------------------------------------------------
Dim init_1 As Object
Set init_1 = jsonObject("form26as")("scheduleHP")
If init_1 <> Empty Then
Dim Node_1, Node_1M, Nodelist_2, Node_2, Node_1MM, Node_3, Nodelist_3 As Object
Dim panOfTenant_count As Integer

Set Node_1M = init_1("propertyDetails")
     panOfTenant_count = 0
     For Each Node_1 In Node_1M
     sTotalCoRow = 0
  
       Set Node_1MM = Node_1("tenantDetails")
          For Each Node_2 In Node_1MM
          Dim malli As Variant
            malli = Node_2("panOfTenant")
                 If malli <> "" Then
                 panOfTenant_count = panOfTenant_count + 1
                 End If
       Next Node_2
  Next Node_1


End If


If panOfTenant_count <= 0 Then

'-------------------------------------------------------------

Set init = jsonObject("lastFiledITR")("scheduleHP")
If init <> Empty Then
If init.exists("propertyDetails") Then

    Set Nodelist = jsonObject("lastFiledITR")("scheduleHP")("propertyDetails")
    TotalExRow = Sheet8.Range("PropertySectionCOunt").value
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    Dim i As Long
    For i = 1 To TotalDiffRow
        Sheet8.Activate
        AddBlockCall_hprptfrm
    Next

    rowcount = 0
    cnt = 0

    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        'Co-Owner
        iTotalCoRow = Sheet8.Range("HP.Co.Name" & rowcount).Rows.count
         
        sTotalCoRow = Node("coOwners").count
        TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
        If TotalDiffCoRow > 0 Then
            Sheet8.Activate
            AddPropertyCoOWners (TotalDiffCoRow)
        End If
        
        iTotalTenRow = Sheet8.Range("HP.NameofTenant" & rowcount).Rows.count
        sTotalTenRow = Node("tenantDetails").count
        TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
        If TotalDiffTenRow > 0 Then
            Sheet8.Activate
            AddPropertyTenant (TotalDiffTenRow)
        End If
            
            If Sheet8.Range("HP.AddrDetail" & rowcount).Locked = False Then
                Sheet8.Range("HP.AddrDetail" & rowcount).value = UCase(Node("addressDetailWithZipCode")("addrDetail"))
            End If
            If Sheet8.Range("HP.CityOrTownOrDistrict" & rowcount).Locked = False Then
                Sheet8.Range("HP.CityOrTownOrDistrict" & rowcount).value = UCase(Node("addressDetailWithZipCode")("cityOrTownOrDistrict"))
            End If
            
            iState = UCase(Node("addressDetailWithZipCode")("stateCode"))
            sState = Findtext(iState, "State")
            Sheet8.Range("HP.StateCode" & rowcount).value = sState
            iCountry = UCase(Node("addressDetailWithZipCode")("countryCode"))
            sCountry = Findtext(iCountry, "Country")
            
            If Sheet8.Range("HP.CountryCode" & rowcount).Locked = False Then
                Sheet8.Range("HP.CountryCode" & rowcount).value = sCountry
            End If
            If Sheet8.Range("HP.PinCode" & rowcount).Locked = False Then
                Sheet8.Range("HP.PinCode" & rowcount).value = UCase(Node("addressDetailWithZipCode")("pinCode"))
            End If
            If Sheet8.Range("HP.ZipCode" & rowcount).Locked = False Then
                Sheet8.Range("HP.ZipCode" & rowcount).value = UCase(Node("addressDetailWithZipCode")("zipCode"))
            End If
            
            Dim OwnerProperty_HP
            
            OwnerProperty_HP = UCase(Node("propertyOwner"))
            
            If OwnerProperty_HP = "SE" Then
               OwnerProperty_HP = "Self"
            ElseIf OwnerProperty_HP = "MI" Then
               OwnerProperty_HP = "Minor"
            ElseIf OwnerProperty_HP = "SP" Then
               OwnerProperty_HP = "Spouse"
            ElseIf OwnerProperty_HP = "OT" Then
               OwnerProperty_HP = "Others"
            Else
                
            End If
            
            If Sheet8.Range("HP.OwnerProperty" & rowcount).Locked = False Then
                Sheet8.Range("HP.OwnerProperty" & rowcount).value = OwnerProperty_HP
            End If
            
            If Sheet8.Range("HP.OwnerPropertyDescription" & rowcount).Locked = False Then
                Sheet8.Range("HP.OwnerPropertyDescription" & rowcount).value = Node("propertyOwnerOther")
            End If
            
            Dim Co_Ownedflag As Variant
            
            Co_Ownedflag = UCase(Node("propCoOwnedFlg"))
            If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
            Co_Ownedflag = "Yes"
            ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
            Co_Ownedflag = "No"
            Else
            Co_Ownedflag = "(Select)"
            End If
            
            Sheet8.Range("HP.CoOwnedYN" & rowcount).value = Co_Ownedflag
             
            Dim hpshare As Variant
            hpshare = UCase(Node("asseseeShareProperty"))
            
            If Sheet8.Range("HP.SharePercent" & rowcount).Locked = False Then
                Sheet8.Range("HP.SharePercent" & rowcount).value = hpshare
            End If
            
    If Node.exists("coOwners") Then
            Set nodeList1 = Node("coOwners")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.Co.Name" & rowcount).name)
            cnt = cnt - 1
            
            For Each Node1 In nodeList1
                    cnt = cnt + 1
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Name" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Name" & rowcount).Column).value = UCase(Node1("nameCoOwner"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Pan" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Pan" & rowcount).Column).value = UCase(Node1("panCoOwner"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Aadhaar" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Aadhaar" & rowcount).Column).value = UCase(Node1("aadhaarCoOwner"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Share" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Share" & rowcount).Column).value = UCase(Node1("percentShareProperty"))
                End If
                    
            Next Node1
            End If
            
            iLetOut = UCase(Node("ifLetOut"))
            
            
            If iLetOut = "Y" Or iLetOut = "L" Then
                sLetOut = "Let Out"
            ElseIf iLetOut = "D" Then
                sLetOut = "Deemed Let Out"
            ElseIf iLetOut = "N" Or iLetOut = "S" Then
                sLetOut = "Self Occupied"
            Else
                sLetOut = "(Select)"
            End If
            
            
            Sheet8.Range("HP.ifLetOut" & rowcount).value = sLetOut
    
            If Node.exists("tenantDetails") Then
            
            Set Nodelist2 = Node("tenantDetails")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.NameofTenant" & rowcount).name)
            cnt = cnt - 1
            For Each Node2 In Nodelist2
                cnt = cnt + 1
                If Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node2("nameOfTenant"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node2("panOfTenant"))
                    ais_Tenchk_pfl = True
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node2("panTANofTenant"))
                    ais_Tenchk_pfl = True
                End If
                    If Node2.Exsist("AadhaarofTenant") Then
'                        Sheet8.Cells(cnt, Sheet8.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node2("aadhaarofTenant"))
                         Sheet8.Cells(cnt, Sheet8.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node2("AadhaarofTenant"))
                    End If
            Next Node2
            End If
            
                If Sheet8.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                    Sheet8.Range("HP.AnnualLetableValue" & rowcount).value = UCase(Node("Rentdetails")("AnnualLetableValue"))
                End If
'Malli comented------------
'                If Sheet8.Range("HP.RentNotRealized" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.RentNotRealized" & rowcount).value = UCase(Node("Rentdetails")("RentNotRealized"))
'                End If
'                If Sheet8.Range("HP.LocalTaxes" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.LocalTaxes" & rowcount).value = UCase(Node("Rentdetails")("LocalTaxes"))
'                End If
'                If Sheet8.Range("HP.IntOnBorwCap" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.IntOnBorwCap" & rowcount).value = UCase(Node("Rentdetails")("IntOnBorwCap"))
'                End If
'                If Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).value = UCase(Node("Rentdetails")("ArrearsUnrealizedRentRcvd"))
'                End If
'------------------------------

                 'Malli_03/02/2026  HP_C_AY_2026_27>>>>>>>>>>>>>>

                  If Node.exists("rentdetails") Then
                      
                            'Schedule24(B)
                            iTotal24BRow = Sheet8.Range("LoanfrmBankOrInstitute.24b" & rowcount).Rows.count
                            sTotal24BRow = Node("rentdetails")("section24B")("section24BDtls").count
                            TotalDiff24BRow = WorksheetFunction.Max((sTotal24BRow - iTotal24BRow), 0)
                            
                            If TotalDiff24BRow > 0 Then
                                Sheet8.Activate
                                EfilingCommon.AddSection24b (TotalDiff24BRow)
                            End If
                             
                          
                            Set Nodelist3 = Node("rentdetails")("section24B")("section24BDtls")
                                cnt = 0
                                cnt = getRowNo(Sheet8.Range("LoanfrmBankOrInstitute.24b" & rowcount).name)
                                cnt = cnt - 1
                                
                                For Each Node3 In Nodelist3
                                    cnt = cnt + 1
                                    
                                                If Sheet8.Cells(cnt, Sheet8.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Locked = False Then
                                                    Dim LoanTknFrom_24B, LoanTknFrom_24B_pfl
                                                    
                                                    LoanTknFrom_24B_pfl = Node3("loanTknFrom")
                                                        If UCase(LoanTknFrom_24B_pfl) = UCase("B") Then
                                                                LoanTknFrom_24B = "Bank "
                                                        ElseIf UCase(LoanTknFrom_24B_pfl) = UCase("I") Then
                                                                LoanTknFrom_24B = "Other than Bank"
                                                        Else
                                                                LoanTknFrom_24B = "(Select)"
                                                        End If
                                                    
                                                    Sheet8.Cells(cnt, Sheet8.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).value = LoanTknFrom_24B
                                                End If
                                                
                                                If Sheet8.Cells(cnt, Sheet8.Range("bankName.24b" & rowcount).Column).Locked = False Then
                                                    Sheet8.Cells(cnt, Sheet8.Range("bankName.24b" & rowcount).Column).value = UCase(Node3("bankOrInstnName"))
                                                End If
                                                
                                                If Sheet8.Cells(cnt, Sheet8.Range("loanAccNum.24b" & rowcount).Column).Locked = False Then
                                                    Sheet8.Cells(cnt, Sheet8.Range("loanAccNum.24b" & rowcount).Column).value = UCase(Node3("loanAccNoOfBankOrInstnRefNo"))
                                                End If
                                                
                                                If Sheet8.Cells(cnt, Sheet8.Range("loanDate.24b" & rowcount).Column).Locked = False Then
                                                    Dim str24BDate As Variant
                                                    str24BDate = Node3("dateofLoan")
                                                    If str24BDate <> "" Then
                                                            str24BDate = Mid(str24BDate, 9, 2) & "/" & Mid(str24BDate, 6, 2) & "/" & Mid(str24BDate, 1, 4)
                                                            Sheet8.Cells(cnt, Sheet8.Range("loanDate.24b" & rowcount).Column).value = str24BDate
                                                               str24BDate = ""
                                                    End If
                                                End If
                                                
                                                If Sheet8.Cells(cnt, Sheet8.Range("loanAmt.24b" & rowcount).Column).Locked = False Then
                                                    Sheet8.Cells(cnt, Sheet8.Range("loanAmt.24b" & rowcount).Column).value = UCase(Node3("totalLoanAmt"))
                                                End If
                                                
                                                
                                                 
                                Next Node3
                            End If

                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                
                
            cnt = cnt + 1
        Next Node
End If
End If
'Malli------------------------------------

Else:
'----------------------------------------------------------
    TotalExRow = Sheet8.Range("PropertySectionCOunt").value
    TotalXMLRow = Node_1M.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    Dim j As Long
    For j = 1 To TotalDiffRow
        Sheet8.Activate
        AddBlockCall_hprptfrm
    Next
'-----------------------------------------------
 


rowcount = 0
  
For Each Node_1 In Node_1M
       
                Dim typeOfHP_tre As Boolean
                typeOfHP_tre = True
                rowcount = rowcount + 1
                'Malli---------------
                
                iTotalTenRow = Sheet8.Range("HP.NameofTenant" & rowcount).Rows.count
                sTotalTenRow = Node_1("tenantDetails").count
                TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
                        
                        If TotalDiffTenRow > 0 Then
                        Sheet8.Activate
                        AddPropertyTenant (TotalDiffTenRow)
                        End If
                '-------------------
                cnt = 0
                cnt = getRowNo(Sheet8.Range("HP.NameofTenant" & rowcount).name)
                cnt = cnt - 1
            
Set Nodelist_2 = Node_1("tenantDetails")
For Each Node_2 In Nodelist_2
              
     If UCase(Node_2("panOfTenant")) <> "" Then
            
            Dim M1 As Boolean
            M1 = True
            ' If isdropdownblank(Sheet8.Range("HP.ifLetOut" & rowcount)) Then
                        If typeOfHP_tre = True Then
                               iLetOut = Node_1("typeOfHP")
                               
                               If iLetOut = "Y" Or iLetOut = "L" Then
                                   sLetOut = "Let Out"
                               ElseIf iLetOut = "D" Then
                                   sLetOut = "Deemed Let Out"
                               ElseIf iLetOut = "N" Or iLetOut = "S" Then
                                   sLetOut = "Self Occupied"
                               Else
                                   sLetOut = "(Select)"
                               End If
                                
                               If Sheet8.Range("HP.ifLetOut" & rowcount).Locked = False Then
                                   Sheet8.Range("HP.ifLetOut" & rowcount).value = sLetOut
                               End If
                               
                                typeOfHP_tre = False
                    
                         End If
    
                cnt = cnt + 1

                    If Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                        Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node_2("nameOfTenant"))
                    End If
                    If Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                        Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node_2("panOfTenant"))
                        ais_Tenchk_pfl = True
                    End If
                    If Sheet8.Cells(cnt, Sheet8.Range("HP.AadharofTenant" & rowcount).Column).Locked = False Then
                        Sheet8.Cells(cnt, Sheet8.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node_2("aadhaarofTenant"))
                    End If
 
                    If Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                        Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node_2("panTANofTenant"))
                        ais_Tenchk_pfl = True
                    End If
                    
            
             End If
             
Next Node_2
                
           If M1 = True Then
                If Sheet8.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                Sheet8.Range("HP.AnnualLetableValue" & rowcount).value = UCase(Node_1("grossRent"))
                End If
                End If
                
Next Node_1
 
End If


'Malli_03/02/2026
Dim init_ais As Object
Dim Node_1ais, Node_ais, Node_2ais, Node_1MMais As Object
Dim iTotalTenRow_ais, sTotalTenRow_ais, TotalDiffTenRow_ais, panOfTenant_counts_ais As Variant

If ais_Tenchk_pfl <> True Then

        If jsonObject("ais") <> Null Then
        Set init_ais = jsonObject("ais")("ScheduleHP")
        
                If init_ais <> Empty Then
                Set Node_ais = init_ais("PropertyDetails")
                
                     panOfTenant_counts_ais = 0
                     
                  For Each Node_1ais In Node_ais
                     sTotalCoRow = 0
                  
                       Set Node_1MMais = Node_1ais("TenantDetails")
                          For Each Node_2ais In Node_1MMais
                            Dim ais_panOfTenant_Nodechk, ais_PANTANofTenant_Nodechk As Variant
                            
                            ais_panOfTenant_Nodechk = Node_2ais("PANOfTenant")
                            ais_PANTANofTenant_Nodechk = Node_2ais("PANTANofTenant")
                            
                                 If ais_panOfTenant_Nodechk <> "" Or ais_PANTANofTenant_Nodechk <> "" Then
                                 panOfTenant_counts_ais = panOfTenant_counts_ais + 1
                                 End If
                                 
                       Next Node_2ais
                  Next Node_1ais
                  
                End If
                
                
                If panOfTenant_counts_ais > 0 Then
                
                Dim TotalDiff_aisRow, Total_aisExRow, Total_aisXMLRow As Variant
                
                    Total_aisExRow = Sheet8.Range("PropertySectionCOunt").value
                    Total_aisXMLRow = Node_ais.count
                    TotalDiff_aisRow = Total_aisXMLRow - Total_aisExRow
                
                    Dim i_ais As Long
                    For i_ais = 1 To TotalDiff_aisRow
                        Sheet8.Activate
                        AddBlockCall_hprptfrm
                    Next
                
                rowcount = 0
                
                For Each Node_1ais In Node_ais
                rowcount = rowcount + 1
                
                        iTotalTenRow_ais = Sheet8.Range("HP.NameofTenant" & rowcount).Rows.count
                        sTotalTenRow_ais = Node_1ais("TenantDetails").count
                        TotalDiffTenRow_ais = WorksheetFunction.Max((sTotalTenRow_ais - iTotalTenRow_ais), 0)
        
                                If TotalDiffTenRow_ais > 0 Then
                                Sheet8.Activate
                                AddPropertyTenant (TotalDiffTenRow_ais)
                                End If
                       
                        cnt = 0
                        cnt = getRowNo(Sheet8.Range("HP.NameofTenant" & rowcount).name)
                        cnt = cnt - 1
                
                        If Sheet8.Range("HP.ifLetOut" & rowcount).Locked = False Then
                        Sheet8.Range("HP.ifLetOut" & rowcount).value = "Let Out"
                        End If
                
                
                            Set Node_1MMais = Node_1ais("TenantDetails")
                                  For Each Node_2ais In Node_1MMais
                                        cnt = cnt + 1

                                                If Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node_2ais("PANOfTenant"))
                                                End If
                                                
                                                If Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node_2ais("PANTANofTenant"))
                                                End If
 
                                  Next Node_2ais
                
                
                      cnt = cnt + 1
                Next Node_1ais
                
                End If
                
                
        End If
End If
'----------------

End Function

Function ImportScheduleHP_Old(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2, init As Object
    Dim Nodelist_1, Node_1 As Object
    Dim Nodelist_2, Node_2 As Object
    Dim Node, Node1, Node2 As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant

    Dim rowcount, cnt As Long
Set jsonObject = ParseJson.ParseJson(jsonText)
Set init = jsonObject("lastFiledITR")("scheduleHP")
If init <> Empty Then
If init.exists("propertyDetails") Then

    Set Nodelist = jsonObject("lastFiledITR")("scheduleHP")("propertyDetails")
    TotalExRow = Sheet8.Range("PropertySectionCOunt").value
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    Dim i As Long
    For i = 1 To TotalDiffRow
        Sheet8.Activate
        AddBlockCall_hprptfrm
    Next

    rowcount = 0
    cnt = 0

    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        'Co-Owner
        iTotalCoRow = Sheet8.Range("HP.Co.Name" & rowcount).Rows.count
         
        sTotalCoRow = Node("coOwners").count
        TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
        If TotalDiffCoRow > 0 Then
            Sheet8.Activate
            AddPropertyCoOWners (TotalDiffCoRow)
        End If
        
        iTotalTenRow = Sheet8.Range("HP.NameofTenant" & rowcount).Rows.count
        sTotalTenRow = Node("tenantDetails").count
        TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
        If TotalDiffTenRow > 0 Then
            Sheet8.Activate
            AddPropertyTenant (TotalDiffTenRow)
        End If
            
            If Sheet8.Range("HP.AddrDetail" & rowcount).Locked = False Then
                Sheet8.Range("HP.AddrDetail" & rowcount).value = UCase(Node("addressDetailWithZipCode")("addrDetail"))
            End If
            If Sheet8.Range("HP.CityOrTownOrDistrict" & rowcount).Locked = False Then
                Sheet8.Range("HP.CityOrTownOrDistrict" & rowcount).value = UCase(Node("addressDetailWithZipCode")("cityOrTownOrDistrict"))
            End If
            
            iState = UCase(Node("addressDetailWithZipCode")("stateCode"))
            sState = Findtext(iState, "State")
            Sheet8.Range("HP.StateCode" & rowcount).value = sState
            iCountry = UCase(Node("addressDetailWithZipCode")("countryCode"))
            sCountry = Findtext(iCountry, "Country")
            
            If Sheet8.Range("HP.CountryCode" & rowcount).Locked = False Then
                Sheet8.Range("HP.CountryCode" & rowcount).value = sCountry
            End If
            If Sheet8.Range("HP.PinCode" & rowcount).Locked = False Then
                Sheet8.Range("HP.PinCode" & rowcount).value = UCase(Node("addressDetailWithZipCode")("pinCode"))
            End If
            If Sheet8.Range("HP.ZipCode" & rowcount).Locked = False Then
                Sheet8.Range("HP.ZipCode" & rowcount).value = UCase(Node("addressDetailWithZipCode")("zipCode"))
            End If
            
            Dim OwnerProperty_HP
            
            OwnerProperty_HP = UCase(Node("propertyOwner"))
            
            If OwnerProperty_HP = "SE" Then
               OwnerProperty_HP = "Self"
            ElseIf OwnerProperty_HP = "MI" Then
               OwnerProperty_HP = "Minor"
            ElseIf OwnerProperty_HP = "SP" Then
               OwnerProperty_HP = "Spouse"
            ElseIf OwnerProperty_HP = "OT" Then
               OwnerProperty_HP = "Others"
            Else
                
            End If
            
            If Sheet8.Range("HP.OwnerProperty" & rowcount).Locked = False Then
                Sheet8.Range("HP.OwnerProperty" & rowcount).value = OwnerProperty_HP
            End If
            
            If Sheet8.Range("HP.OwnerPropertyDescription" & rowcount).Locked = False Then
                Sheet8.Range("HP.OwnerPropertyDescription" & rowcount).value = Node("propertyOwnerOther")
            End If
            
            Dim Co_Ownedflag As Variant
            
            Co_Ownedflag = UCase(Node("propCoOwnedFlg"))
            If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
            Co_Ownedflag = "Yes"
            ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
            Co_Ownedflag = "No"
            Else
            Co_Ownedflag = "(Select)"
            End If
            
            Sheet8.Range("HP.CoOwnedYN" & rowcount).value = Co_Ownedflag
             
            Dim hpshare As Variant
            hpshare = UCase(Node("asseseeShareProperty"))
            
            If Sheet8.Range("HP.SharePercent" & rowcount).Locked = False Then
                Sheet8.Range("HP.SharePercent" & rowcount).value = hpshare
            End If
            
            If Node.exists("coOwners") Then
            Set nodeList1 = Node("coOwners")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.Co.Name" & rowcount).name)
            cnt = cnt - 1
            
            For Each Node1 In nodeList1
                    cnt = cnt + 1
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Name" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Name" & rowcount).Column).value = UCase(Node1("nameCoOwner"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Pan" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Pan" & rowcount).Column).value = UCase(Node1("panCoOwner"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Aadhaar" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Aadhaar" & rowcount).Column).value = UCase(Node1("aadhaarCoOwner"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Share" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.Co.Share" & rowcount).Column).value = UCase(Node1("percentShareProperty"))
                End If
                    
            Next Node1
            End If
    Next Node
End If
End If
            
            
    Dim init_1
    Set init_1 = jsonObject("form26as")("scheduleHP")
    If init_1 <> Empty Then
    
        'Set Nodelist_1 = jsonObject("form26as")("scheduleHP")("propertyDetails")
        iLetOut = jsonObject("form26as")("scheduleHP")("propertyDetails")("typeOfHP")
       
         If Node_1("typeOfHP") = "" Then
            iLetOut = UCase(Node("ifLetOut"))
        End If
            
            If iLetOut = "Y" Then
                sLetOut = "Let Out"
            ElseIf iLetOut = "D" Then
                sLetOut = "Deemed Let Out"
            ElseIf iLetOut = "N" Then
                sLetOut = "Self Occupied"
            Else
                sLetOut = "(Select)"
            End If
            
            
            Sheet8.Range("HP.ifLetOut" & rowcount).value = sLetOut
          
          'Newly added by Bindu Pre-fill SRS - Primary Tag
          If Node_1.exists("tenantDetails") Then
            Set Nodelist_2 = Node("tenantDetails")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.NameofTenant" & rowcount).name)
            cnt = cnt - 1
            For Each Node_2 In Nodelist_2
                cnt = cnt + 1
                If Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node_2("nameOfTenant"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node_2("panOfTenant"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node_2("panTANofTenant"))
                End If
                    If Node2.Exsist("AadhaarofTenant") Then
                        Sheet8.Cells(cnt, Sheet8.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node_2("aadhaarofTenant"))
                    End If
            Next Node_2
          
    'Secondary tag
     Else
            If Node.exists("tenantDetails") Then
            Set Nodelist2 = Node("tenantDetails")
            cnt = 0
            cnt = getRowNo(Sheet8.Range("HP.NameofTenant" & rowcount).name)
            cnt = cnt - 1
            For Each Node2 In Nodelist2
                cnt = cnt + 1
                If Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.NameofTenant" & rowcount).Column).value = UCase(Node2("nameOfTenant"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.PANofTenant" & rowcount).Column).value = UCase(Node2("panOfTenant"))
                End If
                If Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                    Sheet8.Cells(cnt, Sheet8.Range("HP.TANofTenant" & rowcount).Column).value = UCase(Node2("panTANofTenant"))
                End If
                    If Node2.Exsist("AadhaarofTenant") Then
                        Sheet8.Cells(cnt, Sheet8.Range("HP.AadharofTenant" & rowcount).Column).value = UCase(Node2("AadhaarofTenant"))
                    End If
            Next Node2
            End If
            
        End If
            'Newly changed the Tag by Bindu as per Pre-fill SRS
                If Sheet8.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                    Sheet8.Range("HP.AnnualLetableValue" & rowcount).value = UCase(Node_1("grossRent"))
                End If
                
'                If Sheet8.Range("HP.RentNotRealized" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.RentNotRealized" & rowcount).value = UCase(Node("Rentdetails")("RentNotRealized"))
'                End If
'                If Sheet8.Range("HP.LocalTaxes" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.LocalTaxes" & rowcount).value = UCase(Node("Rentdetails")("LocalTaxes"))
'                End If
'                If Sheet8.Range("HP.IntOnBorwCap" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.IntOnBorwCap" & rowcount).value = UCase(Node("Rentdetails")("IntOnBorwCap"))
'                End If
'                If Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).Locked = False Then
'                    Sheet8.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).value = UCase(Node("Rentdetails")("ArrearsUnrealizedRentRcvd"))
'                End If
            cnt = cnt + 1
End If


'    Set nodeList = jsonObject("ScheduleHP")
'    rowcount = 0
'
'    If ActiveWorkbook.Sheets("House Property").Visible = xlSheetVisible Then
'
'        For Each Node In nodeList
'        rowcount = rowcount + 1
'        If Sheet8.Range("HP.PassTroughIncome").Locked = False Then
'            Sheet8.Range("HP.PassTroughIncome").value = UCase(Node("PassThroghIncome"))
'        End If
'        Next Node
'    End If

End Function
Private Function DecodeBase64(ByVal strData As String) As Byte()

 

    Dim objXML As Object
    Dim objNode As Object
   
    ' help from MSXML
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = strData
    DecodeBase64 = objNode.nodeTypedValue
   
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Function ImportPersonalInfo_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary, init, Nodelist As Object
Dim sCountry, iCountry As Variant
Dim sState, iState, Status As Variant
Dim YYYY, MM, DD, strdate As String

Dim DateOFFormOrIncorp, LocalityOrArea, CountryCodeMobileNoSec, MobileNoSec
Dim ResidenceNo, CityOrTownOrDistrict, CountryCodeMobile, MobileNo, emailAddress, emailAddressSec, STDcode, PhoneNo, RoadOrStreet, PinCode, ResidenceName, ZipCode
Dim ResidentialStatus, firstName, SurNameOrOrgName, middleName
Dim dob, aadhaarCardNo, PAN, portugeseCC5A, BenefitUs115HFlg

Set jsonObject = ParseJson.ParseJson(jsonText)

'MsgBox jsonObject
'MsgBox jsonText


'DateOFFormOrIncorp = jsonObject("personalInfo")("orgFirmInfo")("DateOFFormOrIncorp")


Set init = jsonObject("personalInfo")
If init.exists("assesseeName") Then
    firstName = jsonObject("personalInfo")("assesseeName")("firstName")
    SurNameOrOrgName = jsonObject("personalInfo")("assesseeName")("surNameOrOrgName")
    middleName = jsonObject("personalInfo")("assesseeName")("middleName")
End If

PAN = jsonObject("personalInfo")("pan")
portugeseCC5A = jsonObject("personalInfo")("portugeseCC5A")

Set init = jsonObject("personalInfo")
If init.exists("address") Then
    Dim addressObject As Object
    Set addressObject = init("address")
    ResidenceNo = addressObject("residenceNo")
    ResidenceName = addressObject("residenceName")
    RoadOrStreet = addressObject("roadOrStreet")
    LocalityOrArea = addressObject("localityOrArea")
    CityOrTownOrDistrict = addressObject("cityOrTownOrDistrict")
    iState = addressObject("stateCode")
    sState = Findtext(CStr(iState), "State")
    iCountry = addressObject("countryCode")
    sCountry = Findtext(CStr(iCountry), "Country")
    PinCode = addressObject("pinCode")
    ZipCode = addressObject("zipCode")
    If addressObject.exists("phone") Then
        Dim phoneobject As Object
        Set phoneobject = addressObject("phone")
        STDcode = phoneobject("stDcode")
        PhoneNo = phoneobject("phoneNo")
    End If
    CountryCodeMobile = addressObject("countryCodeMobile")
    MobileNo = addressObject("mobileNo")
    CountryCodeMobileNoSec = addressObject("countryCodeMobileNoSec")
    MobileNoSec = addressObject("mobileNoSec")
    emailAddress = addressObject("emailAddress")
    emailAddressSec = addressObject("emailAddressSecondary")
End If

If jsonObject.exists("personalInfo") Then
    Dim OrgFirmInfo
    Set OrgFirmInfo = jsonObject("personalInfo")
    If OrgFirmInfo.exists("dob") Then
        dob = jsonObject("personalInfo")("dob")
    End If
    If dob <> "" Then
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            strdate = DD & "/" & MM & "/" & YYYY
        End If
    
        If strdate <> "" And Sheet1.Range("sheet1.DOB").Locked = False Then
            Sheet1.Range("sheet1.DOB").value = strdate
        End If
End If
    
'Change-28, prefill-tracker, D
    Status = init("status")
    aadhaarCardNo = init("aadhaarCardNo")
    aadhaarCardNo = StrConv(DecodeBase64(aadhaarCardNo), vbUnicode)
    
    If jsonObject("lastFiledITR").exists("filingStatus") Then
        If jsonObject("lastFiledITR")("filingStatus").exists("ResidentialStatus") Then
            ResidentialStatus = jsonObject("lastFiledITR")("filingStatus")("ResidentialStatus")
             If ResidentialStatus = "NRI" Then
                'ResidentialStatus = "NRI-Non-Resident"  'Malli_sit_93940_Issue-1
                ResidentialStatus = "NRI - Non Resident"
                If Sheet1.Range("sheet1.ResidentialStatus1").Locked = False Then
                    Sheet1.Range("sheet1.ResidentialStatus1").value = ResidentialStatus
                End If
            End If
            If ResidentialStatus = "RES" Then
                'ResidentialStatus = "RES-Resident"
                ResidentialStatus = "RES - Resident"
                If Sheet1.Range("sheet1.ResidentialStatus1").Locked = False Then
                    Sheet1.Range("sheet1.ResidentialStatus1").value = ResidentialStatus
                End If
            End If
            If ResidentialStatus = "NOR" Then
                'ResidentialStatus = "NOR - Resident but not Ordinarily Resident"
                ResidentialStatus = "NOR - Resident but not Ordinarily Resident"
                If Sheet1.Range("sheet1.ResidentialStatus1").Locked = False Then
                    Sheet1.Range("sheet1.ResidentialStatus1").value = ResidentialStatus
                End If
            End If
        End If
     End If
'    Status = init("status")
'    aadhaarCardNo = init("aadhaarCardNo")
'    aadhaarCardNo = StrConv(DecodeBase64(aadhaarCardNo), vbUnicode)
'    If init.exists("filingStatus") Then
'        Dim fillingStatuObject As Object
'        Set fillingStatuObject = init("filingStatus")
'        ResidentialStatus = fillingStatuObject("residentialStatus")
''        ResidentialStatus = Findtext(ResidentialStatus, "ResStatus")
'         If ResidentialStatus = "NRI" Then
'            ResidentialStatus = "NRI-Non-Resident"
'            If Sheet1.Range("sheet1.ResidentialStatus1").Locked = False Then
'                Sheet1.Range("sheet1.ResidentialStatus1").value = ResidentialStatus
'            End If
'        End If
'        If ResidentialStatus = "RES" Then
'            ResidentialStatus = "RES-Resident"
'            If Sheet1.Range("sheet1.ResidentialStatus1").Locked = False Then
'                Sheet1.Range("sheet1.ResidentialStatus1").value = ResidentialStatus
'            End If
'        End If
'        If Sheet1.Range("sheet1.ReturnFileSec").Locked = False Then
'                Sheet1.Range("sheet1.ReturnFileSec").value = fillingStatuObject("returnFileSec")
'            End If
'    End If
    
    
    Sheet1.Unprotect Password:=getmsgstate
    'Change-12, Test-005, Remark - Status is coming from personalInfo.status
    Status = jsonObject("personalInfo")("status")
    
'    If Status <> Empty Then
'           'Status = orgFirmInfolastFiledITR("StatusOrCompanyType")
'           If Status = "I" Or UCase(Status) = "INDIVIDUAL" Then
'                Status = "I-INDIVIDUAL"
'            ElseIf Status = "H" Or UCase(Status) = "HUF" Then
'                Status = "H-HUF"
'            Else
'                Status = "(Select)"
'            End If
'        End If
        If Status <> Empty Then
           'Status = orgFirmInfolastFiledITR("StatusOrCompanyType")
           If Status = "IND" Or UCase(Status) = "I" Then
                Status = "I-INDIVIDUAL"
            ElseIf Status = "HUF" Or UCase(Status) = "H" Then
                Status = "H-HUF"
            Else
                Status = "(Select)"
            End If
        End If
        
'Change-3, ANK, 01.06.2022
        Sheet1.Range("sheet1.Status").value = Status
'---
    If firstName <> "" And Sheet1.Range("sheet1.FirstName").Locked = False Then
        Sheet1.Range("sheet1.FirstName").value = firstName
    End If
    
    If jsonObject.exists("form10IF") And jsonObject("form10IF") <> Null Then
    
    'AY_2024_25 Change 'Malli
    
'    If jsonObject("form10IF").exists("newTaxRegime") Then
'    If jsonObject("form10IF")("newTaxRegime") = "Y" Then
'        Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
'        BacYesValueChange
'    ElseIf jsonObject("form10IF")("newTaxRegime") = "N" Then
'        Sheet1.Shapes("BacNo").OLEFormat.Object.value = 1
'        BacNoValueChange
'    End If
'    End If
    
    
  
  '\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//
  
    
'If jsonObject("form10IF").exists("newTaxRegime") Then
'    If jsonObject("form10IF")("newTaxRegime") = "Y" Then
'        Sheet1.Shapes("BacYes").OLEFormat.Object.value = 1
'        BacNoValueChange
'    ElseIf jsonObject("form10IF")("newTaxRegime") = "N" Then
'        Sheet1.Shapes("BacNo").OLEFormat.Object.value = 1
'        BacYesValueChange
'    End If
'    End If
'End If
    
    
        If jsonObject("form10IF").exists("newTaxRegime") Then
                If jsonObject("form10IF")("newTaxRegime") = "Y" Or _
                   jsonObject("form10IF")("newTaxRegime") = "N" Then
                        Sheet1.Shapes("BacNo").OLEFormat.Object.value = 1
                        BacYesValueChange
                End If
        End If
        
    Else:
                Sheet1.Shapes("BacNo").OLEFormat.Object.value = 1
                        BacYesValueChange
End If
'----------------------------------------------------
    
    '----------------------------------------------------------
'    If jsonObject.exists("filingStatus") And jsonObject("filingStatus") <> Null Then
'    If jsonObject("filingStatus").exists("OptingNewTaxRegime") Then
'    If jsonObject("filingStatus")("OptingNewTaxRegime") = 1 Then
'        Sheet1.Range("OptingNewTaxRegime").value = "Opting in now"
'    ElseIf jsonObject("filingStatus")("OptingNewTaxRegime") = 2 Then
'
'        Sheet1.Range("OptingNewTaxRegime").value = "Not Opting"
'    ElseIf jsonObject("filingStatus")("OptingNewTaxRegime") = 3 Then
'
'        Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt"
'    ElseIf jsonObject("filingStatus")("OptingNewTaxRegime") = 4 Then
'
'        Sheet1.Range("OptingNewTaxRegime").value = "Opt out"
'    End If
'    End If
'    End If
'
    
'    If Sheet1.Range("acknowledgmentnumber").Locked = False Then
'        Sheet1.Range("acknowledgmentnumber").value = jsonObject("form10IE")("AcknNumber")
'    End If
    'Change-1, Test-005 Remark-Date format to be changed, following 3 lines code is commented
'    If Sheet1.Range("sheet1.capacityassessee").Locked = False Then
'        Dim DateOfFilingReqdFormat As String
'        DateOfFilingReqdFormat = Mid(jsonObject("form10IE")("DateOfFiling"), 9, 2) & "/" & Mid(jsonObject("form10IE")("DateOfFiling"), 6, 2) & "/" & Mid(jsonObject("form10IE")("DateOfFiling"), 1, 4)
'        Sheet1.Range("sheet1.capacityassessee").value = DateOfFilingReqdFormat
'    End If
    
'    If Sheet1.Range("sheet1.capacityassessee").Locked = False Then
'        Sheet1.Range("sheet1.capacityassessee").value = jsonObject("form10IE")("DateOfFiling")
'    End If
    
    If middleName <> "" And Sheet1.Range("sheet1.MiddleName").Locked = False Then
        Sheet1.Range("sheet1.MiddleName").value = middleName
        
    End If
    If SurNameOrOrgName <> "" And Sheet1.Range("sheet1.SurNameOrOrgName").Locked = False Then
        Sheet1.Range("sheet1.SurNameOrOrgName").value = SurNameOrOrgName
        
    End If
    If portugeseCC5A <> "" And Sheet1.Range("sheet1.PortugeseCC5A").Locked = False Then
    If portugeseCC5A = "T" Or portugeseCC5A = "True" Or portugeseCC5A = "Yes" Or portugeseCC5A = "Y" Then
        Sheet1.Range("sheet1.PortugeseCC5A").value = "Yes"
    End If
    If portugeseCC5A = "F" Or portugeseCC5A = "False" Or portugeseCC5A = "No" Or portugeseCC5A = "N" Then
        Sheet1.Range("sheet1.PortugeseCC5A").value = "No"
    End If
    End If
    
    
    Dim RepName, Repcapacity, RepAddress, RepPAN, RepAadhaar, CompDirectorPrvYrFlg
    RepName = jsonObject("assesseeRep")("repName")
    
    'AY_2024_25 ADD 'Malli
        
        'Y - Yes  'N -No  'sheet1.AsseseeRepFlg
        If RepName <> "" And Sheet1.Range("sheet1.RepName").Locked = False Then
            Sheet1.Range("sheet1.AsseseeRepFlg").value = "Y - Yes"
            ElseIf RepName = "" And Sheet1.Range("sheet1.RepName").Locked = False Then
            Sheet1.Range("sheet1.AsseseeRepFlg").value = "N -No"
        End If
         '-------------------------------------------
    
        If RepName <> "" And Sheet1.Range("sheet1.RepName").Locked = False Then
            Sheet1.Range("sheet1.RepName").value = RepName
        End If
        
        
   'Malli_03/02/2025   PAG_C72to75_AY_2026_27

'        Repcapacity = jsonObject("assesseeRep")("repCapacity")
'        If Repcapacity <> "" Then
'              If Repcapacity = "L" Then
'                    Repcapacity = "Legal Heir"
'               ElseIf Repcapacity = "M" Then
'                    Repcapacity = "Manager"
'               ElseIf Repcapacity = "G" Then
'                    Repcapacity = "Guardian"
'               ElseIf Repcapacity = "O" Then
'                    Repcapacity = "Other"
'               End If
'     Else
'        Repcapacity = "(Select)"
'     End If
'
'     If Repcapacity <> "" And Sheet1.Range("sheet1.RepCapacity").Locked = False Then
'            Sheet1.Range("sheet1.RepCapacity").value = Repcapacity
'     End If
'     RepAddress = jsonObject("assesseeRep")("repAddress")
'    If RepAddress <> "" And Sheet1.Range("sheet1.RepAddress").Locked = False Then
'            Sheet1.Range("sheet1.RepAddress").value = UCase(RepAddress)
'     End If
'     RepPAN = jsonObject("assesseeRep")("repPAN")
'     If RepPAN <> "" And Sheet1.Range("sheet1.RepPAN").Locked = False Then
'            Sheet1.Range("sheet1.RepPAN").value = UCase(RepPAN)
'     End If
'     RepAadhaar = jsonObject("assesseeRep")("aadhaarCardNo")
'     If RepAadhaar <> "" And Sheet1.Range("Sheet1.RepAadhar").Locked = False Then
'        Sheet1.Range("Sheet1.RepAadhar").value = UCase(RepAadhaar)
'     End If
     
     
'Malli_03/02/2025   PAG_C72to75_AY_2026_27
'assesseeRep.repEmailID

        Dim repEmailID_pfl As Variant
        repEmailID_pfl = jsonObject("assesseeRep")("repEmailID")

     If repEmailID_pfl <> "" And Sheet1.Range("sheet1.Email_Representative").Locked = False Then
        Sheet1.Range("sheet1.Email_Representative").value = UCase(repEmailID_pfl)
     End If
     
      Dim repMobileNo_pfl, countryCodeRepMobileNo_pfl As Variant
    
      countryCodeRepMobileNo_pfl = jsonObject("assesseeRep")("countryCodeRepMobileNo")
      repMobileNo_pfl = jsonObject("assesseeRep")("repMobileNo")

     If countryCodeRepMobileNo_pfl <> "" And Sheet1.Range("sheet1.Repcountrycode").Locked = False Then
        Sheet1.Range("sheet1.Repcountrycode").value = UCase(countryCodeRepMobileNo_pfl)
     End If

     If repMobileNo_pfl <> "" And Sheet1.Range("sheet1.Contact_Representative").Locked = False Then
        Sheet1.Range("sheet1.Contact_Representative").value = UCase(repMobileNo_pfl)
     End If

     
 '-----------------------------------------------------------------------------------------
     CompDirectorPrvYrFlg = jsonObject("lastFiledITR")("compDirectorPrvYrFlg")
    
    If CompDirectorPrvYrFlg <> "" Then
            If CompDirectorPrvYrFlg = "Y" Then
                CompDirectorPrvYrFlg = "Yes"
            ElseIf CompDirectorPrvYrFlg = "N" Then
                CompDirectorPrvYrFlg = "No"
            End If
    Else
        CompDirectorPrvYrFlg = "(Select)"
    End If
    
    Sheet1.Range("IsDirectorCompany").value = CompDirectorPrvYrFlg
    If PAN <> "" And Sheet1.Range("sheet1.PAN").Locked = False Then
        Sheet1.Range("sheet1.PAN").value = PAN
    End If
    Dim CompanyNameColNo, CompanyTypeColNo, CompPANColNo, SharesColNo, DINColNo, TotalExRow, TotalDiffRow, TotalXMLRow, rowcount, cnt, Node, RecCompDtls, CB_CostColNo
    Set init = jsonObject("lastFiledITR")("compDirectorPrvYr")
If init.exists("compDirectorPrvYrDtls") Then
   Set Nodelist = jsonObject("lastFiledITR")("compDirectorPrvYr")("compDirectorPrvYrDtls")
   
       CompanyNameColNo = Sheet1.Range("Sheet1.CompanyName").Column
       CompanyTypeColNo = Sheet1.Range("Sheet1.CompanyType").Column
       CompPANColNo = Sheet1.Range("Sheet1.CompPAN").Column
       SharesColNo = Sheet1.Range("Sheet1.Shares").Column
       DINColNo = Sheet1.Range("Sheet1.DIN").Column
       
    TotalExRow = Range("Sheet1.CompanyName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet1.Range("Sheet1.CompanyName").Locked = False Then
            Sheet1.Range("Sheet1.CompanyName").ClearContents
        End If
        If Sheet1.Range("Sheet1.CompanyType").Locked = False Then
            Sheet1.Range("Sheet1.CompanyType").ClearContents
        End If
        If Sheet1.Range("Sheet1.CompPAN").Locked = False Then
            Sheet1.Range("Sheet1.CompPAN").ClearContents
        End If
        If Sheet1.Range("Sheet1.Shares").Locked = False Then
            Sheet1.Range("Sheet1.Shares").ClearContents
        End If
        If Sheet1.Range("Sheet1.DIN").Locked = False Then
            Sheet1.Range("Sheet1.DIN").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_Companydetails (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet1.Range("Sheet1.CompanyName").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
      
        If Sheet1.Cells(rowcount, CompanyNameColNo).Locked = False Then
            Sheet1.Cells(rowcount, CompanyNameColNo).value = Node("nameOfCompany")
        End If
        
'        CompanyType = jsonObject("PartA_GEN1")("FilingStatus")("CompDirectorPrvYr")("CompDirectorPrvYrDtls")
'        If CompanyType <> "" Then
'            If CompanyType = "D" Then
'                CompanyType = "Domestic"
'            ElseIf CompanyType = "F" Then
'                CompanyType = "Foreign"
'            ElseIf CompanyType = "" Then
'                CompanyType = "(Select)"
'            End If
'        End If
        
        If Node("companyType") <> "" Then
            If Node("companyType") = "D" Then
                Sheet1.Cells(rowcount, CompanyTypeColNo).value = "Domestic"
            ElseIf Node("companyType") = "F" Then
                Sheet1.Cells(rowcount, CompanyTypeColNo).value = "Foreign"
            End If
        Else
                Sheet1.Cells(rowcount, CompanyTypeColNo).value = "(Select)"
        End If
        
'        If Sheet1.Cells(rowcount, CompanyNameColNo).Locked = False Then
'            Sheet1.Cells(rowcount, CompanyTypeColNo).value = Node("CompanyType")
'        End If
        If Sheet1.Cells(rowcount, CompPANColNo).Locked = False Then
            Sheet1.Cells(rowcount, CompPANColNo).value = Node("pan")
        End If
        
        If Node("sharesTypes") <> "" Then
            If Node("sharesTypes") = "L" Then
                Sheet1.Cells(rowcount, SharesColNo).value = "Listed"
            ElseIf Node("sharesTypes") = "U" Then
                Sheet1.Cells(rowcount, SharesColNo).value = "Unlisted"
            End If
        Else
                Sheet1.Cells(rowcount, SharesColNo).value = "(Select)"
        End If
        
'        If Sheet1.Cells(rowcount, SharesColNo).Locked = False Then
'            Sheet1.Cells(rowcount, SharesColNo).value = Node("SharesTypes")
'        End If
        
        If Sheet1.Cells(rowcount, DINColNo).Locked = False Then
            Sheet1.Cells(rowcount, DINColNo).value = Node("din")
        End If
    cnt = cnt + 1
    Next Node
    RecCompDtls = cnt
End If
Dim HeldUnlistedEqShrPrYrFlg
 HeldUnlistedEqShrPrYrFlg = jsonObject("lastFiledITR")("heldUnlistedEqShrPrYrFlg")
    
    If HeldUnlistedEqShrPrYrFlg <> "" Then
            If HeldUnlistedEqShrPrYrFlg = "Y" Then
                HeldUnlistedEqShrPrYrFlg = "Yes"
            ElseIf HeldUnlistedEqShrPrYrFlg = "N" Then
                HeldUnlistedEqShrPrYrFlg = "No"
            End If
    Else
        HeldUnlistedEqShrPrYrFlg = "(Select)"
    End If
    
    If HeldUnlistedEqShrPrYrFlg <> "" And Sheet1.Range("IsShareOwner").Locked = False Then
            Sheet1.Range("IsShareOwner").value = HeldUnlistedEqShrPrYrFlg
    End If
'    Dim PartnerInFirmFlg
'    PartnerInFirmFlg = jsonObject("lastFiledITR")("PartnerInFirmFlg")
'    If PartnerInFirmFlg <> "" Then
'            If PartnerInFirmFlg = "Y" Then
'                PartnerInFirmFlg = "Yes"
'            ElseIf PartnerInFirmFlg = "N" Then
'                PartnerInFirmFlg = "No"
'            End If
'    Else
'        PartnerInFirmFlg = "(Select)"
'    End If
'
'
'
'
'
'    Sheet1.Range("IsPartnerFirm").value = PartnerInFirmFlg
'    Dim NriPEinIndia
'     NriPEinIndia = jsonObject("lastFiledITR")("NriPEinIndia")
'     If NriPEinIndia <> "" Then
'            If NriPEinIndia = "Y" Then
'                NriPEinIndia = "Yes"
'            ElseIf NriPEinIndia = "N" Then
'                NriPEinIndia = "No"
'            End If
'    Else
'        NriPEinIndia = "(Select)"
'    End If
'
'    If Sheet1.Range("sheet1.NRI_PE").Locked = False Then
'        Sheet1.Range("sheet1.NRI_PE").value = NriPEinIndia
'    End If
'    Dim ForeignExchangeFlag
'    ForeignExchangeFlag = jsonObject("lastFiledITR")("ForeignExchangeFlag")
'
'
'If ForeignExchangeFlag <> "" Then
'            If ForeignExchangeFlag = "Y" Then
'                ForeignExchangeFlag = "Yes"
'            ElseIf ForeignExchangeFlag = "N" Then
'                ForeignExchangeFlag = "No"
'            End If
'Else
'    ForeignExchangeFlag = "(Select)"
'End If
'
'    If Sheet1.Range("sheet1.Foreign_Exchange").Locked = False Then
'        Sheet1.Range("sheet1.Foreign_Exchange").value = ForeignExchangeFlag
'    End If
'    Dim LiableSec44AAflg
'     LiableSec44AAflg = jsonObject("lastFiledITR")("AuditInfo")("LiableSec44AAflg")
'    If LiableSec44AAflg = "Y" Then
'        Sheet1.Range("sheet1.LiableSec44AAflg").value = "Yes"
'    ElseIf LiableSec44AAflg = "N" Then
'        Sheet1.Range("sheet1.LiableSec44AAflg").value = "No"
'    ElseIf LiableSec44AAflg = "" Then
'        Sheet1.Range("sheet1.LiableSec44AAflg").value = "(Select)"
'    End If
'    Dim LiableSec44AEflg
'    LiableSec44AEflg = jsonObject("lastFiledITR")("AuditInfo")("IncDclrdUs")
'    If LiableSec44AEflg = "Y" Then
'        Sheet1.Range("sheet1.LiableSec44AEflg").value = "Yes"
'    ElseIf LiableSec44AEflg = "N" Then
'        Sheet1.Range("sheet1.LiableSec44AEflg").value = "No"
'    ElseIf LiableSec44AEflg = "" Then
'        Sheet1.Range("sheet1.LiableSec44AEflg").value = "(Select)"
'    End If
'    Dim LiableSec5cflg
'    LiableSec5cflg = jsonObject("lastFiledITR")("AuditInfo")("TotalSalesExcOneCr")
'    If LiableSec5cflg = "Y" Then
'        Sheet1.Range("sheet1.LiableSec5cflg").value = "Yes"
'    ElseIf LiableSec5cflg = "N" Then
'        Sheet1.Range("sheet1.LiableSec5cflg").value = "No"
'    End If
'
'    Dim LiableSecAggrigateflg
'    LiableSecAggrigateflg = jsonObject("lastFiledITR")("AuditInfo")("AgrOfAllAmtsRcvd")
'    If LiableSecAggrigateflg <> Empty Then
'    If LiableSecAggrigateflg = "Y" And Sheet1.Range("sheet1.LiableSecAggrigateflg").Locked = False Then
'        Sheet1.Range("sheet1.LiableSecAggrigateflg").value = "Yes"
'    ElseIf LiableSecAggrigateflg = "N" And Sheet1.Range("sheet1.LiableSecAggrigateflg").Locked = False Then
'        Sheet1.Range("sheet1.LiableSecAggrigateflg").value = "No"
'    End If
'    End If
'
'    'Change-2 Test-028 Remark-typo in the range field, the vba code following is commented
'    ' "sheet1.LiableSecAggrigateflg1" changed to "sheet1.LiableSecAggrigate1flg"
'    Dim LiableSecAggrigateflg1
'    LiableSecAggrigateflg1 = jsonObject("lastFiledITR")("AuditInfo")("AgrOfAllPayMade")
'    If LiableSecAggrigateflg1 <> Empty Then
'    If LiableSecAggrigateflg1 = "Y" And Sheet1.Range("sheet1.LiableSecAggrigate1flg").Locked = False Then
'        Sheet1.Range("sheet1.LiableSecAggrigate1flg").value = "Yes"
'    ElseIf LiableSecAggrigateflg1 = "N" And Sheet1.Range("sheet1.LiableSecAggrigate1flg").Locked = False Then
'        Sheet1.Range("sheet1.LiableSecAggrigate1flg").value = "No"
'    End If
'    End If
'
''    Dim LiableSecAggrigateflg1
''    LiableSecAggrigateflg1 = jsonObject("lastFiledITR")("AuditInfo")("AgrOfAllPayMade")
''    If LiableSecAggrigateflg1 <> Empty Then
''    If LiableSecAggrigateflg1 = "Y" And Sheet1.Range("sheet1.LiableSecAggrigateflg1").Locked = False Then
''        Sheet1.Range("sheet1.LiableSecAggrigateflg1").value = "Yes"
''    ElseIf LiableSecAggrigateflg1 = "N" And Sheet1.Range("sheet1.LiableSecAggrigateflg1").Locked = False Then
''        Sheet1.Range("sheet1.LiableSecAggrigateflg1").value = "No"
''    End If
''    End If
'
'
''change-22, prefill-tracker, D32
'    Dim LiableSec44ABflg
'    LiableSec44ABflg = jsonObject("form3CD")("AuditInfo")("LiableSec44ABflg")
'    If LiableSec44ABflg = "Y" Then
'        Sheet1.Range("sheet1.LiableSec44ABflg").value = "Y-Yes"
'    ElseIf LiableSec44ABflg = "N" Then
'        Sheet1.Range("sheet1.LiableSec44ABflg").value = "N-No"
'    ElseIf LiableSec44ABflg = "" Then
'        Sheet1.Range("sheet1.LiableSec44ABflg").value = "(Select)"
'    End If
'
'    Dim AuditAccountantFlg
'    AuditAccountantFlg = jsonObject("form3CD")("AuditInfo")("AuditAccountantFlg")
'    If AuditAccountantFlg = "Y" Then
'        Sheet1.Range("sheet1.AuditAccountantFlg").value = "Yes"
'    ElseIf AuditAccountantFlg = "N" Then
'        Sheet1.Range("sheet1.AuditAccountantFlg").value = "No"
'    End If
'
'    Dim AuditReportFurnishDate
'    AuditReportFurnishDate = jsonObject("form3CD")("AuditInfo")("AuditReportFurnishDate")
'    If AuditReportFurnishDate <> "" Then
'        YYYY = Mid(AuditReportFurnishDate, 1, 4)
'        MM = Mid(AuditReportFurnishDate, 6, 2)
'        DD = Mid(AuditReportFurnishDate, 9, 2)
'        Sheet1.Range("sheet1.AuditReportFurnishDate").value = DD & "/" & MM & "/" & YYYY
'    End If
'    Sheet1.Range("sheet1.AuditorName").value = jsonObject("form3CD")("AuditInfo")("AuditorName")
'    Sheet1.Range("sheet1.AuditorMemNo").value = jsonObject("form3CD")("AuditInfo")("AuditorMemNo")
'    Sheet1.Range("sheet1.AudFrmName").value = jsonObject("form3CD")("AuditInfo")("AudFrmName")
'    Sheet1.Range("sheet1.AudFrmRegNo").value = jsonObject("form3CD")("AuditInfo")("AudFrmRegNo")
'    Sheet1.Range("sheet1.AudFrmPAN").value = jsonObject("form3CD")("AuditInfo")("AudFrmPAN")
'    Sheet1.Range("sheet1.AudFrmAadhaar").value = jsonObject("form3CD")("AuditInfo")("AudFrmAadhaar")
'    Dim strAuditDate, LiableSec92Eflg, LiableSec92Eflgdia, DateOfAudit, iAuditAct As Variant
'
'    strAuditDate = jsonObject("form3CD")("AuditInfo")("AuditDate")
'    If strAuditDate <> "" Then
'        YYYY = Mid(strAuditDate, 1, 4)
'        MM = Mid(strAuditDate, 6, 2)
'        DD = Mid(strAuditDate, 9, 2)
'        Sheet1.Range("sheet1.AuditDate").value = DD & "/" & MM & "/" & YYYY
'    End If
'    YYYY = ""
'    MM = ""
'    DD = ""
'    LiableSec92Eflg = jsonObject("form3CEB")("AuditInfo")("LiableSec92Eflg")
'    If LiableSec92Eflg = "Y" Then
'        Sheet1.Range("sheet1.LiableSec92Eflg").value = "Y-Yes"
'    ElseIf LiableSec92Eflg = "N" Then
'        Sheet1.Range("sheet1.LiableSec92Eflg").value = "N-No"
'    End If
    'Change-3, Test-052, Remark-typo in jsonObject "AccountAuditflag", following vba code is commented
    '"AccountAuditFlag" changed to "AuditAccountantFlg"
'    LiableSec92Eflgdia = jsonObject("form3CEB")("AuditInfo")("AuditAccountantFlg")
'    LiableSec92Eflgdia = jsonObject("form3CEB")("AuditInfo")("AccountAuditFlag")
'    If LiableSec92Eflgdia = "Y" Then
'        Sheet1.Range("sheet1.LiableSec92Eflgdia").value = "Yes"
'    ElseIf LiableSec92Eflgdia = "N" Then
'        Sheet1.Range("sheet1.LiableSec92Eflgdia").value = "No"
'    End If
'    DateOfAudit = jsonObject("form3CEB")("AuditInfo")("AuditDetails92E")("DateOfAudit")
'    If DateOfAudit <> "" Then
'        YYYY = Mid(DateOfAudit, 1, 4)
'        MM = Mid(DateOfAudit, 6, 2)
'        DD = Mid(DateOfAudit, 9, 2)
'        Sheet1.Range("sheet1.AuditDateSec92E").value = DD & "/" & MM & "/" & YYYY
'    End If
'    Dim AuditedSectionColNo, AuditYNColNo, DateOfAuditColNo
'    Set init = jsonObject("auditInfo")
'Dim auditFlag, AuditDt, AuditDate, RecAuditDtls As Variant
'If init.exists("AuditDetails") Then
'
'    Set NodeList = jsonObject("auditInfo")("AuditDetails")
'
'    AuditedSectionColNo = Sheet1.Range("sheet1.AuditedSection").Column
'    AuditYNColNo = Sheet1.Range("sheet1.AuditYN").Column
'    DateOfAuditColNo = Sheet1.Range("sheet1.DateOfAudit").Column
'
'    TotalExRow = Range("Sheet1.AuditedSection").Rows.count
'
'    TotalXMLRow = NodeList.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'        If Sheet1.Range("sheet1.AuditedSection").Locked = False Then
'            Sheet1.Range("sheet1.AuditedSection").ClearContents
'        End If
'        If Sheet1.Range("Sheet1.AuditYN").Locked = False Then
'            Sheet1.Range("Sheet1.AuditYN").ClearContents
'        End If
'        If Sheet1.Range("Sheet1.DateOfAudit").Locked = False Then
'            Sheet1.Range("Sheet1.DateOfAudit").ClearContents
'        End If
'    End If
'
'    If (TotalDiffRow > 0) Then
'        AddDiffRows_AuditDetails (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet1.Range("Sheet1.AuditedSection").name)
'    rowcount = rowcount - 1
'    cnt = 0
'
'
'
'    For Each node In NodeList
'        rowcount = rowcount + 1
'
'        auditFlag = node("AuditFlag")
'
'        If auditFlag = "Y" Then
'            auditFlag = "Yes"
'        ElseIf auditFlag = "N" Then
'            auditFlag = "No"
'        ElseIf auditFlag = "" Then
'            auditFlag = "(Select)"
'        End If
'
'        AuditDt = node("DateOfAudit")
'        If AuditDt <> "" Then
'            YYYY = Mid(AuditDt, 1, 4)
'            MM = Mid(AuditDt, 6, 2)
'            DD = Mid(AuditDt, 9, 2)
'            AuditDate = DD & "/" & MM & "/" & YYYY
'
'            Sheet1.Cells(rowcount, DateOfAuditColNo).value = AuditDate
'
'            AuditDt = ""
'            YYYY = ""
'            MM = ""
'            DD = ""
'        End If
'
'        If Sheet1.Cells(rowcount, AuditedSectionColNo).Locked = False Then
'            Sheet1.Cells(rowcount, AuditedSectionColNo).value = node("AuditedSection")
'        End If
'        If Sheet1.Cells(rowcount, AuditYNColNo).Locked = False Then
'            Sheet1.Cells(rowcount, AuditYNColNo).value = auditFlag
'        End If
'
'        cnt = cnt + 1
'
'    Next node
'
'    RecAuditDtls = cnt
'End If
'
'Change-3, ANK, 11.07.2022
'Dim PriSecflag
'PriSecflag = False
'
'If jsonObject("lastFiledITR") <> Null Then
'    If jsonObject("lastFiledITR").exists("AuditInfo") Then
'        Set init = jsonObject("lastFiledITR")("AuditInfo")
'        If init.exists("AuditReportDetails") Then
'            PriSecflag = True
'            Set NodeList = jsonObject("lastFiledITR")("AuditInfo")("AuditReportDetails")
'        End If
'    End If
'End If
'
'If PriSecflag = False Then
'    If jsonObject("form3CD") <> Null Then
'        If jsonObject("form3CD").exists("PartA_GEN2") Then
'            If jsonObject("form3CD")("PartA_GEN2").exists("AuditReportDetails") Then
'                PriSecflag = True
'                Set NodeList = jsonObject("form3CD")("PartA_GEN2")("AuditReportDetails")
'            End If
'        End If
'    End If
'End If
'---end change---

'If PriSecflag = True Then
'    Dim ActColNo, OthersColNo, SectionsColNo, ACTYNColNo, DateColNo, iACTYN
'
''    AuditedSectionColNo = Sheet1.Range("Audit.Act").Column
''    DateOfAuditColNo = Sheet1.Range("Audit.Sections").Column
'
'
'    ActColNo = Sheet1.Range("Audit.Act").Column
'    OthersColNo = Sheet1.Range("Audit.Others").Column
'    SectionsColNo = Sheet1.Range("Audit.Sections").Column
'    ACTYNColNo = Sheet1.Range("Audit.ACTYN").Column
'    DateColNo = Sheet1.Range("Audit.Date").Column
'
'
'    TotalExRow = Range("Audit.Act").Rows.count
'
'    TotalXMLRow = NodeList.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'        If Sheet1.Range("Audit.Act").Locked = False Then
'            Sheet1.Range("Audit.Act").ClearContents
'        End If
''        If Sheet1.Range("Audit.Others").Locked = False Then
''            Sheet1.Range("Audit.Others").ClearContents
''        End If
'        If Sheet1.Range("Audit.Sections").Locked = False Then
'            Sheet1.Range("Audit.Sections").ClearContents
'        End If
'        If Sheet1.Range("Audit.ACTYN").Locked = False Then
'            Sheet1.Range("Audit.ACTYN").ClearContents
'        End If
'        If Sheet1.Range("Audit.Date").Locked = False Then
'            Sheet1.Range("Audit.Date").ClearContents
'        End If
'
'    End If
'
'    If (TotalDiffRow > 0) Then
'       AddDiffRows_AuditReportDTL (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet1.Range("Audit.Act").name)
'    rowcount = rowcount - 1
'    cnt = 0
'
'
'
'    For Each node In NodeList
'        rowcount = rowcount + 1
'
'        iACTYN = node("AuditReportSection")
'
'        Dim iDate, sDate
'        iDate = node("DateOfAudit")
'        If iDate <> "" Then
'            YYYY = Mid(iDate, 1, 4)
'            MM = Mid(iDate, 6, 2)
'            DD = Mid(iDate, 9, 2)
'            sDate = DD & "/" & MM & "/" & YYYY
'        End If
'
'        If iACTYN = "Y" Then
'            iACTYN = "Yes"
'        ElseIf iACTYN = "N" Then
'            iACTYN = "No"
'        ElseIf iACTYN = "" Then
'            iACTYN = "(Select)"
'        End If
'        iAuditAct = node("AuditReportAct")
'
'        If iAuditAct = "1" Then
'            iAuditAct = "Banking Regulation Act, 1949"
'        ElseIf iAuditAct = "2" Then
'            iAuditAct = "Central Excise Act,1944"
'        ElseIf iAuditAct = "3" Then
'            iAuditAct = "Central Sales Tax Act, 1956"
'        ElseIf iAuditAct = "4" Then
'            iAuditAct = "Central Goods and Services Tax Act, 2017"
'        ElseIf iAuditAct = "5" Then
'            iAuditAct = "Charitable And Religious Trusts Act, 1920"
'        ElseIf iAuditAct = "6" Then
'            iAuditAct = "Companies Act, 2013"
'        ElseIf iAuditAct = "7" Then
'            iAuditAct = "Electricity Act, 2003"
'        ElseIf iAuditAct = "8" Then
'            iAuditAct = "Employees Provident Fund and Miscellaneous Provisions Act, 1952"
'        ElseIf iAuditAct = "9" Then
'            iAuditAct = "Foreign Exchange Management Act, 1999"
'        ElseIf iAuditAct = "10" Then
'            iAuditAct = "Government Superannuation Fund Act, 1956"
'        ElseIf iAuditAct = "11" Then
'            iAuditAct = "Indian Trusts Act, 1882"
'        ElseIf iAuditAct = "12" Then
'            iAuditAct = "Integrated Goods and Services Tax Act, 2017"
'        ElseIf iAuditAct = "13" Then
'            iAuditAct = "Limited Liability Partnership Act, 2008"
'        ElseIf iAuditAct = "14" Then
'            iAuditAct = "Payment of Gratuity Act, 1972"
'        ElseIf iAuditAct = "15" Then
'            iAuditAct = "SEBI Act, 1992"
'        ElseIf iAuditAct = "16" Then
'            iAuditAct = "Securities Contract (Regulation) Act, 1956"
'        ElseIf iAuditAct = "17" Then
'            iAuditAct = "State Goods and Services Tax Act, 2017"
'        ElseIf iAuditAct = "18" Then
'            iAuditAct = "Union Territories Goods and Services Tax Act, 2017"
'        ElseIf iAuditAct = "19" Then
'            iAuditAct = "Others"
'        Else
'            iAuditAct = "(Select)"
'        End If
'
''        If Sheet1.Cells(rowcount, AuditedSectionColNo).Locked = False Then
''            Sheet1.Cells(rowcount, AuditedSectionColNo).value = iAuditAct
''        End If
'
'        Sheet1.Cells(rowcount, ActColNo).value = iAuditAct
'        If iAuditAct = "Others" Then
'            If Sheet1.Cells(rowcount, OthersColNo).Locked = False Then
'                Sheet1.Cells(rowcount, OthersColNo).value = node("AuditReportActOthers")
'            End If
'        End If
'
'        If Sheet1.Cells(rowcount, SectionsColNo).Locked = False Then
'            Sheet1.Cells(rowcount, SectionsColNo).value = node("AuditedSection")
'        End If
'
'        If Sheet1.Cells(rowcount, ACTYNColNo).Locked = False Then
'            Sheet1.Cells(rowcount, ACTYNColNo).value = iACTYN
'        End If
'         If Sheet1.Cells(rowcount, DateColNo).Locked = False Then
'            Sheet1.Cells(rowcount, DateColNo).value = sDate
'        End If
'        cnt = cnt + 1
'
'    Next node
'
'    RecAuditDtls = cnt
'End If
'


    Dim NameOfCompanyColNo, TypeofCompanyColNo, PANColNo, OP_CountColNo, OP_CostColNo, Acquired_CountColNo, Acquired_DateColNo, Acquired_FaceValueColNo
    Dim Acquired_IssueColNo, Acquired_CostColNo, Transferred_CountColNo, Transferred_SaleColNo, CB_CountColNo, PurchaseDt, PurchaseDate, RecSHARE, RecJRES
    Set init = jsonObject("lastFiledITR")("heldUnlistedEqShrPrYr")
If init.exists("heldUnlistedEqShrPrYrDtls") Then
    Set Nodelist = jsonObject("lastFiledITR")("heldUnlistedEqShrPrYr")("heldUnlistedEqShrPrYrDtls")
    
    'lastFiledITR.heldUnlistedEqShrPrYr.heldUnlistedEqShrPrYrDtls.opngBalNumberOfShares
    'lastFiledITR.heldUnlistedEqShrPrYr.heldUnlistedEqShrPrYrDtls.opngBalCostOfAcquisition
    
    NameOfCompanyColNo = Sheet1.Range("Share_NameofCompany").Column
    TypeofCompanyColNo = Sheet1.Range("Share_TypeofCompany").Column
    PANColNo = Sheet1.Range("Share_PAN").Column
    OP_CountColNo = Sheet1.Range("Share_OP_Count").Column
    OP_CostColNo = Sheet1.Range("Share_OP_Cost").Column
    Acquired_CountColNo = Sheet1.Range("Share_Acquired_Count").Column
    Acquired_DateColNo = Sheet1.Range("Share_Acquired_Date").Column
    Acquired_FaceValueColNo = Sheet1.Range("Share_Acquired_FaceValue").Column
    Acquired_IssueColNo = Sheet1.Range("Share_Acquired_Issue").Column
    Acquired_CostColNo = Sheet1.Range("Share_Acquired_Cost").Column
    Transferred_CountColNo = Sheet1.Range("Share_Transferred_Count").Column
    Transferred_SaleColNo = Sheet1.Range("Share_Transferred_Sale").Column
    CB_CountColNo = Sheet1.Range("Share_CB_Count").Column
    CB_CostColNo = Sheet1.Range("Share_CB_Cost").Column
    
    TotalExRow = Range("Share_NameofCompany").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet1.Range("Share_NameofCompany").Locked = False Then
            Sheet1.Range("Share_NameofCompany").ClearContents
        End If

        If Sheet1.Range("Share_TypeofCompany").Locked = False Then
            Sheet1.Range("Share_TypeofCompany").ClearContents
        End If
        
        If Sheet1.Range("Share_PAN").Locked = False Then
            Sheet1.Range("Share_PAN").ClearContents
        End If
        If Sheet1.Range("Share_OP_Count").Locked = False Then
            Sheet1.Range("Share_OP_Count").ClearContents
        End If
        If Sheet1.Range("Share_OP_Cost").Locked = False Then
            Sheet1.Range("Share_OP_Cost").ClearContents
        End If
        If Sheet1.Range("Share_Acquired_Count").Locked = False Then
            Sheet1.Range("Share_Acquired_Count").ClearContents
        End If
        If Sheet1.Range("Share_Acquired_Date").Locked = False Then
            Sheet1.Range("Share_Acquired_Date").ClearContents
        End If
        If Sheet1.Range("Share_Acquired_FaceValue").Locked = False Then
            Sheet1.Range("Share_Acquired_FaceValue").ClearContents
        End If
        If Sheet1.Range("Share_Acquired_Issue").Locked = False Then
            Sheet1.Range("Share_Acquired_Issue").ClearContents
        End If
        If Sheet1.Range("Share_Acquired_Cost").Locked = False Then
            Sheet1.Range("Share_Acquired_Cost").ClearContent
        End If
        If Sheet1.Range("Share_Transferred_Count").Locked = False Then
            Sheet1.Range("Share_Transferred_Count").ClearContents
        End If
        If Sheet1.Range("Share_Transferred_Sale").Locked = False Then
            Sheet1.Range("Share_Transferred_Sale").ClearContents
        End If
        If Sheet1.Range("Share_CB_Count").Locked = False Then
            Sheet1.Range("Share_CB_Count").ClearContents
        End If
        If Sheet1.Range("Share_CB_Cost").Locked = False Then
            Sheet1.Range("Share_CB_Cost").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_ShareDetails (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet1.Range("Share_NameofCompany").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
        PurchaseDt = Acquired_DateColNo
            YYYY = Mid(PurchaseDt, 1, 4)
            MM = Mid(PurchaseDt, 6, 2)
            DD = Mid(PurchaseDt, 9, 2)
        PurchaseDate = DD & "/" & MM & "/" & YYYY
        
    If Sheet1.Cells(rowcount, NameOfCompanyColNo).Locked = False Then
        Sheet1.Cells(rowcount, NameOfCompanyColNo).value = Node("nameOfCompany")
    End If

        If Node("companyType") <> "" Then
            If Node("companyType") = "D" Then
                Sheet1.Cells(rowcount, TypeofCompanyColNo).value = "Domestic"
            ElseIf Node("companyType") = "F" Then
                Sheet1.Cells(rowcount, TypeofCompanyColNo).value = "Foreign"
            End If
        Else
                Sheet1.Cells(rowcount, TypeofCompanyColNo).value = "(Select)"
        End If
    
'    If Sheet1.Cells(rowcount, TypeofCompanyColNo).Locked = False Then
'        Sheet1.Cells(rowcount, TypeofCompanyColNo).value = Node("CompanyType")
'    End If
    
    If Sheet1.Cells(rowcount, PANColNo).Locked = False Then
        Sheet1.Cells(rowcount, PANColNo).value = Node("pan")
    End If
    If Sheet1.Cells(rowcount, OP_CountColNo).Locked = False Then
        Sheet1.Cells(rowcount, OP_CountColNo).value = Node("opngBalNumberOfShares")
    End If
    If Sheet1.Cells(rowcount, OP_CostColNo).Locked = False Then
        Sheet1.Cells(rowcount, OP_CostColNo).value = Node("opngBalCostOfAcquisition")
    End If
    If Sheet1.Cells(rowcount, Acquired_CountColNo).Locked = False Then
        Sheet1.Cells(rowcount, Acquired_CountColNo).value = Node("shrAcqDurYrNumberOfShares")
    End If
    
'        AuditDt = node("dateOfSubscrPurchase")
'        If AuditDt <> "" Then
'            YYYY = Mid(AuditDt, 1, 4)
'            MM = Mid(AuditDt, 6, 2)
'            DD = Mid(AuditDt, 9, 2)
'            AuditDate = DD & "/" & MM & "/" & YYYY
'
'            If Sheet1.Cells(rowcount, Acquired_DateColNo).Locked = False Then
'                Sheet1.Cells(rowcount, Acquired_DateColNo).value = AuditDate
'            End If
'
'            AuditDt = ""
'            YYYY = ""
'            MM = ""
'            DD = ""
'        End If
    
    
'    If Sheet1.Cells(rowcount, PurchaseDate).Locked = False Then
'        Sheet1.Cells(rowcount, PurchaseDate).value = Node("DateOfSubscrPurchase")
'    End If
    If Sheet1.Cells(rowcount, Acquired_FaceValueColNo).Locked = False Then
        Sheet1.Cells(rowcount, Acquired_FaceValueColNo).value = Node("faceValuePerShare")
    End If
    If Sheet1.Cells(rowcount, Acquired_IssueColNo).Locked = False And Node("faceValuePerShare") <> "" And Node("issuePricePerShare") <> 0 Then
        Sheet1.Cells(rowcount, Acquired_IssueColNo).value = Node("issuePricePerShare")
    End If
    If Sheet1.Cells(rowcount, Acquired_CostColNo).Locked = False And Node("purchasePricePerShare") <> "" And Node("purchasePricePerShare") <> 0 Then
        Sheet1.Cells(rowcount, Acquired_CostColNo).value = Node("purchasePricePerShare")
    End If
    If Sheet1.Cells(rowcount, Transferred_CountColNo).Locked = False Then
        Sheet1.Cells(rowcount, Transferred_CountColNo).value = Node("shrTrnfNumberOfShares")
    End If
    If Sheet1.Cells(rowcount, Transferred_SaleColNo).Locked = False Then
        Sheet1.Cells(rowcount, Transferred_SaleColNo).value = Node("shrTrnfSaleConsideration")
    End If
    If Sheet1.Cells(rowcount, CB_CountColNo).Locked = False Then
        Sheet1.Cells(rowcount, CB_CountColNo).value = Node("clsngBalNumberOfShares")
    End If
    If Sheet1.Cells(rowcount, CB_CostColNo).Locked = False Then
        Sheet1.Cells(rowcount, CB_CostColNo).value = Node("clsngBalCostOfAcquisition")
    End If
     
    cnt = cnt + 1
    Next Node
    RecSHARE = cnt
End If

    
    If ResidenceNo <> "" And Sheet1.Range("sheet1.ResidenceNo").Locked = False Then
        Sheet1.Range("sheet1.ResidenceNo").value = ResidenceNo
    End If
    If ResidenceName <> "" And Sheet1.Range("sheet1.ResidenceName").Locked = False Then
        Sheet1.Range("sheet1.ResidenceName").value = ResidenceName
    End If
    If RoadOrStreet <> "" And Sheet1.Range("sheet1.RoadOrStreet").Locked = False Then
        Sheet1.Range("sheet1.RoadOrStreet").value = RoadOrStreet
    End If
    If LocalityOrArea <> "" And Sheet1.Range("sheet1.LocalityOrArea").Locked = False Then
        Sheet1.Range("sheet1.LocalityOrArea").value = LocalityOrArea
    End If
    If CityOrTownOrDistrict <> "" And Sheet1.Range("sheet1.CityOrTownOrDistrict").Locked = False Then
        Sheet1.Range("sheet1.CityOrTownOrDistrict").value = CityOrTownOrDistrict
    End If
    If sState <> "" And Sheet1.Range("sheet1.StateCode1").Locked = False Then
        Sheet1.Range("sheet1.StateCode1").value = sState
    End If
    If sCountry <> "" And Sheet1.Range("sheet1.CountryCode").Locked = False Then
        Sheet1.Range("sheet1.CountryCode").value = sCountry
    End If
    If PinCode <> "" And Sheet1.Range("sheet1.PinCode").Locked = False Then
        Sheet1.Range("sheet1.PinCode").value = PinCode
    End If
    If ZipCode <> "" And Sheet1.Range("sheet1.ZipCode").Locked = False Then
        Sheet1.Range("sheet1.ZipCode").value = ZipCode
    End If
    If STDcode <> "" And Sheet1.Range("sheet1.STDcode").Locked = False Then
        Sheet1.Range("sheet1.STDcode").value = STDcode
    End If
    If PhoneNo <> "" And Sheet1.Range("sheet1.PhoneNo").Locked = False Then
        Sheet1.Range("sheet1.PhoneNo").value = PhoneNo
    End If
    If CountryCodeMobile <> "" And Sheet1.Range("Countrycode1").Locked = False Then
        Sheet1.Range("Countrycode1").value = CountryCodeMobile
    End If
    If MobileNo <> "" And Sheet1.Range("sheet1.MobileNo").Locked = False Then
        Sheet1.Range("sheet1.MobileNo").value = MobileNo
    End If
    If CountryCodeMobileNoSec <> "" And Sheet1.Range("Countrycode2").Locked = False Then
        Sheet1.Range("Countrycode2").value = CountryCodeMobileNoSec
    End If
    If MobileNoSec <> "" And Sheet1.Range("sheet1.MobileNoSec").Locked = False Then
        Sheet1.Range("sheet1.MobileNoSec").value = MobileNoSec
    End If
    If emailAddress <> "" And Sheet1.Range("sheet1.EmailAddress").Locked = False Then
        Sheet1.Range("sheet1.EmailAddress").value = emailAddress
    End If
    If emailAddressSec <> "" And Sheet1.Range("sheet1.EmailAddressSec").Locked = False Then
        Sheet1.Range("sheet1.EmailAddressSec").value = emailAddressSec
    End If
    
    If aadhaarCardNo <> "" And Sheet1.Range("sheet1.adhaarno").Locked = False Then
        Sheet1.Range("sheet1.adhaarno").value = aadhaarCardNo
    End If
    'Change-9(b), Test-032, Remark - placed after returnFileSec
'    If jsonObject("filingStatus")("receiptNo") <> "" And Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
'        Sheet1.Range("sheet1.ReceiptNo").value = jsonObject("filingStatus")("receiptNo")
'    End If
    
    
    Dim iReturnFile, ReturnFileSec, sReturnFile
    ReturnFileSec = jsonObject("filingStatus")("returnFileSec")
    
    
    If ReturnFileSec <> "" Then
        iReturnFile = ReturnFileSec
'Change-9(a), Test-032, Remark - typo corrected and 19 added
'following code Is commented
        If iReturnFile = "11" Then
           sReturnFile = "139(1)- On or Before due date"
        ElseIf iReturnFile = "12" Then
           sReturnFile = "139(4)- After due date"
        ElseIf iReturnFile = "13" Then
           sReturnFile = "142(1)"
        ElseIf iReturnFile = "14" Then
           sReturnFile = "148"
        ElseIf iReturnFile = "15" Then
           sReturnFile = "153A"
        ElseIf iReturnFile = "16" Then
           sReturnFile = "153C"
        ElseIf iReturnFile = "17" Then
           sReturnFile = "139(5)- Revised Return"
        ElseIf iReturnFile = "18" Then
           sReturnFile = "139(9)"
        ElseIf iReturnFile = "19" Then
           sReturnFile = "92CD-Modified return"
        ElseIf iReturnFile = "20" Then
          'Malli----
           'AY_2023_24 OLD
           sReturnFile = "119(2)(b)- after condonation of delay"
          'AY_2024_25 Change
           'sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
           '--------------------
        End If
'        If iReturnFile = "11" Then
'           sReturnFile = "139(1)-On or before due date"
'        ElseIf iReturnFile = "12" Then
'           sReturnFile = "139(4)-After due date"
'        ElseIf iReturnFile = "13" Then
'           sReturnFile = "142(1)"
'        ElseIf iReturnFile = "14" Then
'           sReturnFile = "148"
'        ElseIf iReturnFile = "15" Then
'           sReturnFile = "153A"
'        ElseIf iReturnFile = "16" Then
'           sReturnFile = "153C"
'        ElseIf iReturnFile = "17" Then
'           sReturnFile = "139(5)-Revised Return"
'        ElseIf iReturnFile = "18" Then
'           sReturnFile = "139(9)"
'        ElseIf iReturnFile = "20" Then
'           sReturnFile = "119(2)(b)- after condonation of delay"
'        End If
       If sReturnFile <> "" Then ' Add by Malli
        Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
      End If
      End If
    If jsonObject("filingStatus")("receiptNo") <> "" And Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
        Sheet1.Range("sheet1.ReceiptNo").value = jsonObject("filingStatus")("receiptNo")
    End If
    
    Dim OrigRetFiledDt, OrigRetFiledDate, NoticeNo
    OrigRetFiledDt = jsonObject("filingStatus")("origRetFiledDate")
    If OrigRetFiledDt <> "" Then
    YYYY = Mid(OrigRetFiledDt, 1, 4)
    MM = Mid(OrigRetFiledDt, 6, 2)
    DD = Mid(OrigRetFiledDt, 9, 2)
    OrigRetFiledDate = DD & "/" & MM & "/" & YYYY
    End If

    
    If OrigRetFiledDate <> "" And Sheet1.Range("Sheet1.OrigRetFiledDate").Locked = False Then
        Sheet1.Range("Sheet1.OrigRetFiledDate").value = OrigRetFiledDate
     End If
     NoticeNo = jsonObject("filingStatus")("uniqueNo")
     If NoticeNo <> "" And Sheet1.Range("Sheet1.UniqueNo").Locked = False Then
        Sheet1.Range("Sheet1.UniqueNo").value = NoticeNo
     End If
     Dim NoticeDateUnderSec
     NoticeDateUnderSec = jsonObject("filingStatus")("noticeDateUnderSec")
     strdate = ""
           If NoticeDateUnderSec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
                 YYYY = Mid(NoticeDateUnderSec, 1, 4)
                MM = Mid(NoticeDateUnderSec, 6, 2)
                DD = Mid(NoticeDateUnderSec, 9, 2)
                strdate = DD & "/" & MM & "/" & YYYY
                Sheet1.Range("sheet1.NoticeDate").value = strdate
            End If
     
     Set init = jsonObject("lastFiledITR")("jurisdictionResPrevYr")
Dim JurisdictionResidenceColNo, TINColNo

If init.exists("jurisdictionResPrevYrDtls") Then
Set Nodelist = jsonObject("lastFiledITR")("jurisdictionResPrevYr")("jurisdictionResPrevYrDtls")
        
        JurisdictionResidenceColNo = Sheet1.Range("NRI_Juridiction").Column
        TINColNo = Sheet1.Range("NRI_Taxpayer").Column

        TotalExRow = Range("NRI_Juridiction").Rows.count
    
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
            
    If (TotalXMLRow > 0) Then
        If Sheet1.Range("NRI_Juridiction").Locked = False Then
            Sheet1.Range("NRI_Juridiction").ClearContents
        End If
        If Sheet1.Range("NRI_Taxpayer").Locked = False Then
            Sheet1.Range("NRI_Taxpayer").ClearContents
        End If
    End If
    
     If (TotalDiffRow > 0) Then
        AddDiffRows_JurisdictionRES (TotalDiffRow)
     End If
            
    rowcount = getRowNo(Sheet1.Range("NRI_Juridiction").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        If Sheet1.Cells(rowcount, JurisdictionResidenceColNo).Locked = False Then
            Dim iJurisdictionResidence, sJurisdictionResidence
            iJurisdictionResidence = Node("jurisdictionResidence")
            sJurisdictionResidence = Findtext(CStr(iJurisdictionResidence), "Country_Juridiction")
            Sheet1.Cells(rowcount, JurisdictionResidenceColNo).value = sJurisdictionResidence
        End If
        If Sheet1.Cells(rowcount, TINColNo).Locked = False Then
            Sheet1.Cells(rowcount, TINColNo).value = Node("tin")
        End If
            
    cnt = cnt + 1
    Next Node
    RecJRES = cnt
 End If


    
    BenefitUs115HFlg = jsonObject("lastFiledITR")("benefitUs115HFlg")
    
    If BenefitUs115HFlg <> "" Then
            If BenefitUs115HFlg = "Y" Then
                BenefitUs115HFlg = "Yes"
            ElseIf BenefitUs115HFlg = "N" Then
                BenefitUs115HFlg = "No"
            End If
    Else
        BenefitUs115HFlg = "(Select)"
    End If
        
    If Sheet1.Range("Sheet1.115H").Locked = False Then
        Sheet1.Range("Sheet1.115H").value = BenefitUs115HFlg
    End If
    
'    Dim lastFiledITR
'    Set lastFiledITR = jsonObject("lastFiledITR")
'    If lastFiledITR.exists("orgFirmInfo") Then
'        Dim orgFirmInfolastFiledITR
'        Set orgFirmInfolastFiledITR = lastFiledITR("orgFirmInfo")

    'End If
    

    
End Function


Function ScheduleCFLImport(jsonText As String)
On Error Resume Next
Dim YYYY, MM, DD, Date_1, dname, CFL_HP, CFL_STCG, CFL_LTCG As String
Dim jsonObject, Nodelist, Node As Object
Set jsonObject = ParseJson.ParseJson(jsonText)


Set Nodelist = jsonObject("scheduleCFL")("CarryFwdLossDetail")
For Each Node In Nodelist
If Node("DateOfFiling") <> "" Then
    Date_1 = Node("DateOfFiling")
    YYYY = Mid(Date_1, 1, 4)
    MM = Mid(Date_1, 6, 2)
    DD = Mid(Date_1, 9, 2)
    Date_1 = DD & "/" & MM & "/" & YYYY
    
'    If YYYY = "2010" Then
'      If Date_1 <> "" And Sheet17.Range("yr1999.DateOfFiling").Locked = False Then
'    Sheet17.Range(dname).value = Date_1
'    End If
'    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2010").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2010").value = Node("HpLossCF")
'    End If
'    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2010").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2010").value = Node("StcgLossCF")
'    End If
'    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2010").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2010").value = Node("LtcgLossCF")
'    End If
'    End If
'    If YYYY = "2011" Then
'      If Date_1 <> "" And Sheet17.Range("yr2000.DateOfFiling1").Locked = False Then
'    Sheet17.Range(dname).value = Date_1
'    End If
'    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2011").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2011").value = Node("HpLossCF")
'    End If
'    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2011").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2011").value = Node("StcgLossCF")
'    End If
'    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2011").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2011").value = Node("LtcgLossCF")
'    End If
'    End If
'    If YYYY = "2012" Then
'      If Date_1 <> "" And Sheet17.Range("yr2001.DateOfFiling2").Locked = False Then
'    Sheet17.Range(dname).value = Date_1
'    End If
'    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2012").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2012").value = Node("HpLossCF")
'    End If
'    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2012").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2012").value = Node("StcgLossCF")
'    End If
'    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2012").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2012").value = Node("LtcgLossCF")
'    End If
'    End If


 'Malli_commented_AY_2026_27<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 
'     If YYYY = "2013" Then
'      If Date_1 <> "" And Sheet17.Range("yr2002.DateOfFiling3").Locked = False Then
'    Sheet17.Range("yr2002.DateOfFiling3").value = Date_1
'    End If
'    If node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2013").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2013").value = node("HpLossCF")
'    End If
'    If node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2013").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2013").value = node("StcgLossCF")
'    End If
'    If node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2013").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2013").value = node("LtcgLossCF")
'    End If
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("U6").Locked = False Then
'    Sheet17.Range("U6").value = node("OthSrcLossRaceHorseCF")
'    End If
'
'    If node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2003.LossFrmSpecifiedBusCF3").Locked = False Then
'    Sheet17.Range("yr2003.LossFrmSpecifiedBusCF3").value = node("LossFrmSpecifiedBusCF")
'    End If
'
'    End If
'
'
'     If YYYY = "2014" Then
'      If Date_1 <> "" And Sheet17.Range("yr2003.DateOfFiling4").Locked = False Then
'    Sheet17.Range("yr2003.DateOfFiling4").value = Date_1
'    End If
'    If node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2014").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2014").value = node("HpLossCF")
'    End If
'    If node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2014").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2014").value = node("StcgLossCF")
'    End If
'    If node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2014").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2014").value = node("LtcgLossCF")
'    End If
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2003.OthSrcLossRaceHorseCF4").Locked = False Then
'    Sheet17.Range("yr2003.OthSrcLossRaceHorseCF4").value = node("OthSrcLossRaceHorseCF")
'    End If
'
''    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4a").Locked = False Then
''    Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4a").value = Node("BroughtFrwrdBusLoss")
''    End If
'
'    If node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2003.LossFrmSpecifiedBusCF4").Locked = False Then
'    Sheet17.Range("yr2003.LossFrmSpecifiedBusCF4").value = node("LossFrmSpecifiedBusCF")
'    End If
'
'    End If
'
'
'    If YYYY = "2015" Then
'      If Date_1 <> "" And Sheet17.Range("yr2004.DateOfFiling5").Locked = False Then
'    Sheet17.Range("yr2004.DateOfFiling5").value = Date_1
'    End If
'    If node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2015").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2015").value = node("HpLossCF")
'    End If
'    If node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2015").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2015").value = node("StcgLossCF")
'    End If
'    If node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2015").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2015").value = node("LtcgLossCF")
'    End If
'
''    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5a").Locked = False Then
''    Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5a").value = Node("BroughtFrwrdBusLoss")
''    End If
'
''    If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2004.LossFrmSpecifiedBusCF5").Locked = False Then
''    Sheet17.Range("yr2004.LossFrmSpecifiedBusCF5").value = Node("LossFrmSpecifiedBusCF")
''    End If
'
'    End If
'
'
'     If YYYY = "2016" Then
'      If Date_1 <> "" And Sheet17.Range("yr2005.DateOfFiling6").Locked = False Then
'    Sheet17.Range("yr2005.DateOfFiling6").value = Date_1
'    End If
'    If node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2016").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2016").value = node("HpLossCF")
'    End If
'    If node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2016").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2016").value = node("StcgLossCF")
'    End If
'    If node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2016").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2016").value = node("LtcgLossCF")
'    End If
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2005.OthSrcLossRaceHorseCF6").Locked = False Then
'    Sheet17.Range("yr2005.OthSrcLossRaceHorseCF6").value = node("OthSrcLossRaceHorseCF")
'    End If
'
''    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6a").Locked = False Then
''    Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6a").value = Node("BroughtFrwrdBusLoss")
''    End If
'
''    If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2005.LossFrmSpecifiedBusCF6").Locked = False Then
''    Sheet17.Range("yr2005.LossFrmSpecifiedBusCF6").value = Node("LossFrmSpecifiedBusCF")
''    End If
'
'    End If
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    'Malli_commented_AY_2026_27
'    If YYYY = "2017" Then
'      If Date_1 <> "" And Sheet17.Range("yr2006.DateOfFiling7").Locked = False Then
'    Sheet17.Range("yr2006.DateOfFiling7").value = Date_1
'    End If
'    If node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2017").Locked = False Then
'    Sheet17.Range("CFL_HP_Normal_2017").value = node("HpLossCF")
'    End If
'    If node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2017").Locked = False Then
'    Sheet17.Range("CFL_STCG_Normal_2017").value = node("StcgLossCF")
'    End If
'    If node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2017").Locked = False Then
'    Sheet17.Range("CFL_LTCG_Normal_2017").value = node("LtcgLossCF")
'    End If
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2006.OthSrcLossRaceHorseCF7").Locked = False Then
'    Sheet17.Range("yr2006.OthSrcLossRaceHorseCF7").value = node("OthSrcLossRaceHorseCF")
'    End If
'
''    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7a").Locked = False Then
''    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7a").value = Node("BroughtFrwrdBusLoss")
''    End If
'
''     If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2006.LossFrmSpecifiedBusCF7").Locked = False Then
''    Sheet17.Range("yr2006.LossFrmSpecifiedBusCF7").value = Node("LossFrmSpecifiedBusCF")
''    End If
'
'    End If
'    '------------------------------------------
    
    
    If YYYY = "2018" Then
      If Date_1 <> "" And Sheet17.Range("yr2006.DateOfFiling9").Locked = False Then
    Sheet17.Range("yr2006.DateOfFiling9").value = Date_1
    End If
    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2018").Locked = False Then
    Sheet17.Range("CFL_HP_Normal_2018").value = Node("HpLossCF")
    End If
    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2018").Locked = False Then
    Sheet17.Range("CFL_STCG_Normal_2018").value = Node("StcgLossCF")
    End If
    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2018").Locked = False Then
    Sheet17.Range("CFL_LTCG_Normal_2018").value = Node("LtcgLossCF")
    End If
    
    'Malli_Commented AY_2026-27
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2006.OthSrcLossRaceHorseCF9").Locked = False Then
'    Sheet17.Range("yr2006.OthSrcLossRaceHorseCF9").value = node("OthSrcLossRaceHorseCF")
'    End If
    '--------------------------
    
'    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9a").Locked = False Then
'    Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9a").value = Node("BroughtFrwrdBusLoss")
'    End If
    
'    If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2006.LossFrmSpecifiedBusCF9").Locked = False Then
'    Sheet17.Range("yr2006.LossFrmSpecifiedBusCF9").value = Node("LossFrmSpecifiedBusCF")
'    End If
    
'    If Node("LossFrmSpecBusCF") <> "" And Sheet17.Range("yr2006.LossFrmSpecBusCF9").Locked = False Then
'    Sheet17.Range("yr2006.LossFrmSpecBusCF9").value = Node("LossFrmSpecBusCF")
'    End If
    
    End If
    
    
    If YYYY = "2019" Then
      If Date_1 <> "" And Sheet17.Range("yr2007.DateOfFiling9").Locked = False Then
    Sheet17.Range("yr2007.DateOfFiling9").value = Date_1
    End If
    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2019").Locked = False Then
    Sheet17.Range("CFL_HP_Normal_2019").value = Node("HpLossCF")
    End If
    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2019").Locked = False Then
    Sheet17.Range("CFL_STCG_Normal_2019").value = Node("StcgLossCF")
    End If
    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2019").Locked = False Then
    Sheet17.Range("CFL_LTCG_Normal_2019").value = Node("LtcgLossCF")
    End If
    
    'Malli_Commented AY_2026-27
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2007.OthSrcLossRaceHorseCF9").Locked = False Then
'    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF9").value = node("OthSrcLossRaceHorseCF")
'    End If
    '--------------------------
    
'    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9a").Locked = False Then
'    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9a").value = Node("BroughtFrwrdBusLoss")
'    End If
    
'    If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2007.LossFrmSpecifiedBusCF9").Locked = False Then
'    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF9").value = Node("LossFrmSpecifiedBusCF")
'    End If
    
'    If Node("LossFrmSpecBusCF") <> "" And Sheet17.Range("yr2007.LossFrmSpecBusCF9").Locked = False Then
'    Sheet17.Range("yr2007.LossFrmSpecBusCF9").value = Node("LossFrmSpecBusCF")
'    End If
    
    End If
    
    
    If YYYY = "2020" Then
      If Date_1 <> "" And Sheet17.Range("yr2020.DateOfFiling10").Locked = False Then
    Sheet17.Range("yr2020.DateOfFiling10").value = Date_1
    End If
    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2020").Locked = False Then
    Sheet17.Range("CFL_HP_Normal_2020").value = Node("HpLossCF")
    End If
    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2020").Locked = False Then
    Sheet17.Range("CFL_STCG_Normal_2020").value = Node("StcgLossCF")
    End If
    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2020").Locked = False Then
    Sheet17.Range("CFL_LTCG_Normal_2020").value = Node("LtcgLossCF")
    End If
    
    'Malli_Commented AY_2026-27
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2020.OthSrcLossRaceHorseCF10").Locked = False Then
'    Sheet17.Range("yr2020.OthSrcLossRaceHorseCF10").value = node("OthSrcLossRaceHorseCF")
'    End If
    '--------------------
    
'    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10a").Locked = False Then
'    Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10a").value = Node("BroughtFrwrdBusLoss")
'    End If
    
'    If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2020.LossFrmSpecifiedBusCF10").Locked = False Then
'    Sheet17.Range("yr2020.LossFrmSpecifiedBusCF10").value = Node("LossFrmSpecifiedBusCF")
'    End If
    
'    If Node("LossFrmSpecBusCF") <> "" And Sheet17.Range("yr2020.LossFrmSpecBusCF10").Locked = False Then
'    Sheet17.Range("yr2020.LossFrmSpecBusCF10").value = Node("LossFrmSpecBusCF")
'    End If
    
    End If
    
    
    
    If YYYY = "2021" Then
      If Date_1 <> "" And Sheet17.Range("yr2021.DateOfFiling11").Locked = False Then
    Sheet17.Range("yr2021.DateOfFiling11").value = Date_1
    End If
    If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2021").Locked = False Then
    Sheet17.Range("CFL_HP_Normal_2021").value = Node("HpLossCF")
    End If
    If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2021").Locked = False Then
    Sheet17.Range("CFL_STCG_Normal_2021").value = Node("StcgLossCF")
    End If
    If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2021").Locked = False Then
    Sheet17.Range("CFL_LTCG_Normal_2021").value = Node("LtcgLossCF")
    End If
    
    'Malli_Commented AY_2026-27
'    If node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2021.OthSrcLossRaceHorseCF11").Locked = False Then
'    Sheet17.Range("yr2021.OthSrcLossRaceHorseCF11").value = node("OthSrcLossRaceHorseCF")
'    End If
    '--------------
    
'    If Node("BroughtFrwrdBusLoss") <> "" And Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11a").Locked = False Then
'    Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11a").value = Node("BroughtFrwrdBusLoss")
'    End If
    
'    If Node("LossFrmSpecifiedBusCF") <> "" And Sheet17.Range("yr2021.LossFrmSpecifiedBusCF11").Locked = False Then
'    Sheet17.Range("yr2021.LossFrmSpecifiedBusCF11").value = Node("LossFrmSpecifiedBusCF")
'    End If
    
'    If Node("LossFrmSpecBusCF") <> "" And Sheet17.Range("yr2021.LossFrmSpecBusCF11").Locked = False Then
'    Sheet17.Range("yr2021.LossFrmSpecBusCF11").value = Node("LossFrmSpecBusCF")
'    End If
    
    End If
    'Malli
    
    

        If YYYY = "2022" Then
        
            If Date_1 <> "" And Sheet17.Range("yr2022.DateOfFiling11").Locked = False Then
            Sheet17.Range("yr2022.DateOfFiling11").value = Date_1
            End If
            If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2022").Locked = False Then
            Sheet17.Range("CFL_HP_Normal_2022").value = Node("HpLossCF")
            End If
            If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2022").Locked = False Then
            Sheet17.Range("CFL_STCG_Normal_2022").value = Node("StcgLossCF")
            End If
            If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2022").Locked = False Then
            Sheet17.Range("CFL_LTCG_Normal_2022").value = Node("LtcgLossCF")
            End If
            If Node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2022.OthSrcLossRaceHorseCF11").Locked = False Then
            Sheet17.Range("yr2022.OthSrcLossRaceHorseCF11").value = Node("OthSrcLossRaceHorseCF")
            End If
        
        End If
    
        If YYYY = "2023" Then
          
            If Date_1 <> "" And Sheet17.Range("yr2023.DateOfFiling12").Locked = False Then
            Sheet17.Range("yr2023.DateOfFiling12").value = Date_1
            End If
            If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2023").Locked = False Then
            Sheet17.Range("CFL_HP_Normal_2023").value = Node("HpLossCF")
            End If
            If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2023").Locked = False Then
            Sheet17.Range("CFL_STCG_Normal_2023").value = Node("StcgLossCF")
            End If
            If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2023").Locked = False Then
            Sheet17.Range("CFL_LTCG_Normal_2023").value = Node("LtcgLossCF")
            End If
            If Node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2023.OthSrcLossRaceHorseCF12").Locked = False Then
            Sheet17.Range("yr2023.OthSrcLossRaceHorseCF12").value = Node("OthSrcLossRaceHorseCF")
            End If
    
        End If
    
    
 'Malli_AY_2026_27
         If YYYY = "2024" Then
              
            If Date_1 <> "" And Sheet17.Range("yr2024.DateOfFiling13").Locked = False Then
            Sheet17.Range("yr2024.DateOfFiling13").value = Date_1
            End If
            
            If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2024").Locked = False Then
            Sheet17.Range("CFL_HP_Normal_2024").value = Node("HpLossCF")
            End If
            
            If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2024").Locked = False Then
            Sheet17.Range("CFL_STCG_Normal_2024").value = Node("StcgLossCF")
            End If
            
            If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2024").Locked = False Then
            Sheet17.Range("CFL_LTCG_Normal_2024").value = Node("LtcgLossCF")
            End If
            
            If Node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2024.OthSrcLossRaceHorseCF13").Locked = False Then
            Sheet17.Range("yr2024.OthSrcLossRaceHorseCF13").value = Node("OthSrcLossRaceHorseCF")
            End If
        
            
        End If
        
        
        If YYYY = "2025" Then
              
            If Date_1 <> "" And Sheet17.Range("yr2025.DateOfFiling14").Locked = False Then
            Sheet17.Range("yr2025.DateOfFiling14").value = Date_1
            End If
            
            If Node("HpLossCF") <> "" And Sheet17.Range("CFL_HP_Normal_2025").Locked = False Then
            Sheet17.Range("CFL_HP_Normal_2025").value = Node("HpLossCF")
            End If
            
            If Node("StcgLossCF") <> "" And Sheet17.Range("CFL_STCG_Normal_2025").Locked = False Then
            Sheet17.Range("CFL_STCG_Normal_2025").value = Node("StcgLossCF")
            End If
            
            If Node("LtcgLossCF") <> "" And Sheet17.Range("CFL_LTCG_Normal_2025").Locked = False Then
            Sheet17.Range("CFL_LTCG_Normal_2025").value = Node("LtcgLossCF")
            End If
            
            If Node("OthSrcLossRaceHorseCF") <> "" And Sheet17.Range("yr2025.OthSrcLossRaceHorseCF14").Locked = False Then
            Sheet17.Range("yr2025.OthSrcLossRaceHorseCF14").value = Node("OthSrcLossRaceHorseCF")
            End If
        
            
        End If
 
 '----------------
    
    End If

    Next Node

 
End Function
Function ImportScheduleFA(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
Dim jsonDictionary As Object
Dim Node, Nodelist, init As Object
Dim CountryColNo, BankColNo, AddressOfBankColNo, ZipCodeColNo, ForeignAccountNumberColNo, OwnerStatusColNo As Variant
Dim AccOpenDateColNo, PeakBalColNo, ClosingBalanceColNo, IntrstAccuredColNo As Variant
Dim CustodialAcc_CountryColNo, CustodialAcc_BanknameColNo, CustodialAcc_AddressOfBankColNo, CustodialAcc_ZipCodeColNo As Variant
Dim CustodialAcc_ForeignAccountNumberColNo, CustodialAcc_OwnerStatusColNo, CustodialAcc_AccOpenDateColNo As Variant
Dim CustodialAcc_PeakBalColNo, CustodialAcc_ClosingBalColNo, CustodialAcc_Grossinterest_NatureColNo, CustodialAcc_GrossinterestColNo As Variant
Dim FEquity_CountryColNo, FEquity_BankNameColNo, FEquity_AddressOfBankColNo, FEquity_ZipCodeColNo As Variant
Dim FEquity_NatureOfEntityColNo, FEquity_AccOpenDateColNo, FEquity_InitialvalueColNo, FEquity_PeakBalColNo As Variant
Dim FEquity_ClosingBalColNo, FEquity_TotalgrossamountColNo, FEquity_TotalgrosproceedsColNo As Variant
Dim FA_A4_CountryColNo, FA_A4_BankNameColNo, FA_A4_BankAddColNo, FA_A4_ZipCodeColNo, FA_A4_AccOpeningDateColNo As Variant
Dim FA_A4_the_cashvalue, FA_A4_Total_grossamount As Variant
Dim FA_C_CountryColNo, FA_C_ZipCodeColNo, FA_C_PropertyAddColNo, FA_C_OwnershipBeneficiaryColNo As Variant
Dim FA_C_DateOfAcquisitionColNo, FA_C_TotalInvColNo, FA_C_IncomeFromPropertyColNo, FA_C_NatureOfIncomeColNo As Variant
Dim FA_C_AmountColNo, FA_C_ScheduleOfferedColNo, FA_C_ItemNumOfScheduleColNo As Variant
Dim FA_D_CountryColNo, FA_D_ZipCodeColNo, FA_D_AssetNatureColNo, FA_D_OwnershipBeneficiaryColNo As Variant
Dim FA_D_DateOfAcquisitionColNo, FA_D_TotalInvColNo, FA_D_IncomeFromPropertyColNo, FA_D_NatureOfIncomeColNo As Variant
Dim FA_D_AmountColNo, FA_D_ScheduleOfferedColNo, FA_D_ItemNumOfScheduleColNo As Variant
Dim FA_E_InstName, FA_E_InstAdd, FA_E_Country, FA_E_ZipCode, FA_E_AccountName, FA_E_InstitutionAccountNumber As Variant
Dim FA_E_PeakBalInv, FA_E_IncomeAccured, FA_E_IncomeAccuredAmount, FA_E_Amount, FA_E_Scheduleoffered, FA_E_ItemNumberSchedule As Variant
Dim FA_F_Country, FA_F_ZipCode, FA_F_NameOfTrust, FA_F_AddressOfTrust, FA_F_NameOfOtherTrustees As Variant
Dim FA_F_AddressOfOtherTrustees, FA_F_NameOfSettlor, FA_F_AddressOfSettlor, FA_F_NameOfBeneficiaries As Variant
Dim FA_F_AddressOfBeneficiaries, FA_F_DateSinceHeld, FA_F_IncomeDerivedIsTaxable, FA_F_IncomeDerivedAmount As Variant
Dim FA_F_Amount, FA_F_Scheduleoffered, FA_F_ItemNumberSchedule
Dim FA_G_Country, FA_G_ZipCode, FA_G_NameOfPerson, FA_G_AddressOfPerson, FA_G_IncomeDerived, FA_G_NatureOfIncome As Variant
Dim FA_G_WhetherTaxable, FA_G_Amount, FA_G_Scheduleoffered, FA_G_ItemNumberSchedule As Variant
Dim TotalXMLRow, RecFAcc, rowcount, cnt, RecFCAcc, RecFEqAcc, RecFCashIns, RecImProp, RecOthAssts, RecActSgnAuth, RecTrustOIT, RecSrcIncOI As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim iCountry, sCountry, iStatus, sStatus As Variant
Dim YYYY, MM, DD, Date_1 As String
Dim IncTaxSch

Dim NodeList_FA_A1
Set NodeList_FA_A1 = jsonObject("lastFiledITR")
If NodeList_FA_A1 <> Null Then
If NodeList_FA_A1.exists("scheduleFA") Then

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsForiegnBank") Then
Set Nodelist = init("detailsForiegnBank")

    CountryColNo = Sheet27.Range("FA_A1_Country").Column
    BankColNo = Sheet27.Range("FA_A1_BankName").Column
    AddressOfBankColNo = Sheet27.Range("FA_A1_BankAdd").Column
    ZipCodeColNo = Sheet27.Range("FA_A1_ZipCode").Column
    ForeignAccountNumberColNo = Sheet27.Range("FA_A1_ForeignAccountNumber").Column
    OwnerStatusColNo = Sheet27.Range("FA_A1_StatusBeneficiary").Column
    AccOpenDateColNo = Sheet27.Range("FA_A1_AccOpeningDate").Column
    PeakBalColNo = Sheet27.Range("FA_A1_PeakBal").Column
    ClosingBalanceColNo = Sheet27.Range("FA_A1_ClosingBal").Column
    IntrstAccuredColNo = Sheet27.Range("FA_A1_Grossinterest").Column
    
    TotalExRow = Sheet27.Range("FA_A1_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_A1_Country").Locked = False Then
            Sheet27.Range("FA_A1_Country").ClearContents
        End If
        If Sheet27.Range("FA_A1_BankName").Locked = False Then
            Sheet27.Range("FA_A1_BankName").ClearContents
        End If
        If Sheet27.Range("FA_A1_BankAdd").Locked = False Then
            Sheet27.Range("FA_A1_BankAdd").ClearContents
        End If
        If Sheet27.Range("FA_A1_ZipCode").Locked = False Then
            Sheet27.Range("FA_A1_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_A1_ForeignAccountNumber").Locked = False Then
            Sheet27.Range("FA_A1_ForeignAccountNumber").ClearContents
        End If
        If Sheet27.Range("FA_A1_StatusBeneficiary").Locked = False Then
            Sheet27.Range("FA_A1_StatusBeneficiary").ClearContents
        End If
        If Sheet27.Range("FA_A1_AccOpeningDate").Locked = False Then
            Sheet27.Range("FA_A1_AccOpeningDate").ClearContents
        End If
        If Sheet27.Range("FA_A1_PeakBal").Locked = False Then
            Sheet27.Range("FA_A1_PeakBal").ClearContents
        End If
        If Sheet27.Range("FA_A1_ClosingBal").Locked = False Then
            Sheet27.Range("FA_A1_ClosingBal").ClearContents
        End If
        If Sheet27.Range("FA_A1_Grossinterest").Locked = False Then
            Sheet27.Range("FA_A1_Grossinterest").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A1_FA1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A1_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
         
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
            
            If Sheet27.Cells(rowcount, CountryColNo).Locked = False Then
                Sheet27.Cells(rowcount, CountryColNo).value = sCountry
            End If
            If Sheet27.Cells(rowcount, BankColNo).Locked = False Then
                Sheet27.Cells(rowcount, BankColNo).value = Node("bankname")
            End If
            If Sheet27.Cells(rowcount, AddressOfBankColNo).Locked = False Then
                Sheet27.Cells(rowcount, AddressOfBankColNo).value = Node("addressOfBank")
            End If
            If Sheet27.Cells(rowcount, ZipCodeColNo).Locked = False Then
                Sheet27.Cells(rowcount, ZipCodeColNo).value = Node("zipCode")
            End If
            If Sheet27.Cells(rowcount, ForeignAccountNumberColNo).Locked = False Then
                Sheet27.Cells(rowcount, ForeignAccountNumberColNo).value = Trim(Node("foreignAccountNumber"))
            End If
            
           
            iStatus = Node("ownerStatus")
            If iStatus = "BENIFICIARY" Then
                sStatus = "BENIFICIARY"
            ElseIf iStatus = "BENEFICIAL_OWNER" Then
                 sStatus = "BENEFICIAL_OWNER"
            ElseIf iStatus = "OWNER" Then
                 sStatus = "OWNER"
            Else
                sStatus = "(Select)"
            End If
                 
             Sheet27.Cells(rowcount, OwnerStatusColNo).value = sStatus
            
            Date_1 = Node("accOpenDate")
            
            'Malli
            If Date_1 <> "" Then
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            End If
            
            Sheet27.Cells(rowcount, AccOpenDateColNo) = Date_1
            
            If Sheet27.Cells(rowcount, PeakBalColNo).Locked = False Then
                Sheet27.Cells(rowcount, PeakBalColNo).value = Node("peakBalanceDuringYear")
            End If
            If Sheet27.Cells(rowcount, ClosingBalanceColNo).Locked = False Then
                Sheet27.Cells(rowcount, ClosingBalanceColNo).value = Node("closingBalance")
            End If
            If Sheet27.Cells(rowcount, IntrstAccuredColNo).Locked = False Then
                Sheet27.Cells(rowcount, IntrstAccuredColNo).value = Node("intrstAccured")
            End If
            
cnt = cnt + 1
Next Node
RecFAcc = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("dtlsForeignCustodialAcc") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("dtlsForeignCustodialAcc")
Dim CustodialAcc_FA_A2_GrossinterestColNo
    CustodialAcc_CountryColNo = Sheet27.Range("FA_A2_Country").Column
    CustodialAcc_BanknameColNo = Sheet27.Range("FA_A2_BankName").Column
    CustodialAcc_AddressOfBankColNo = Sheet27.Range("FA_A2_BankAdd").Column
    CustodialAcc_ZipCodeColNo = Sheet27.Range("FA_A2_ZipCode").Column
    CustodialAcc_ForeignAccountNumberColNo = Sheet27.Range("FA_A2_ForeignAccountNumber").Column
    CustodialAcc_OwnerStatusColNo = Sheet27.Range("FA_A2_StatusBeneficiary").Column
    CustodialAcc_AccOpenDateColNo = Sheet27.Range("FA_A2_AccOpeningDate").Column
    CustodialAcc_PeakBalColNo = Sheet27.Range("FA_A2_PeakBal").Column
    CustodialAcc_ClosingBalColNo = Sheet27.Range("FA_A2_ClosingBal").Column
    CustodialAcc_Grossinterest_NatureColNo = Sheet27.Range("FA_A2_Grossinterest_Nature").Column
      CustodialAcc_FA_A2_GrossinterestColNo = Sheet27.Range("FA_A2_Grossinterest").Column
        
    TotalExRow = Sheet27.Range("FA_A2_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_A2_Country").Locked = False Then
            Sheet27.Range("FA_A2_Country").ClearContents
        End If
        If Sheet27.Range("FA_A2_BankName").Locked = False Then
            Sheet27.Range("FA_A2_BankName").ClearContents
        End If
        If Sheet27.Range("FA_A2_BankAdd").Locked = False Then
            Sheet27.Range("FA_A2_BankAdd").ClearContents
        End If
        If Sheet27.Range("FA_A2_ZipCode").Locked = False Then
            Sheet27.Range("FA_A2_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_A2_ForeignAccountNumber").Locked = False Then
            Sheet27.Range("FA_A2_ForeignAccountNumber").ClearContents
        End If
        If Sheet27.Range("FA_A2_StatusBeneficiary").Locked = False Then
            Sheet27.Range("FA_A2_StatusBeneficiary").ClearContents
        End If
        If Sheet27.Range("FA_A2_AccOpeningDate").Locked = False Then
            Sheet27.Range("FA_A2_AccOpeningDate").ClearContents
        End If
        If Sheet27.Range("FA_A2_PeakBal").Locked = False Then
            Sheet27.Range("FA_A2_PeakBal").ClearContents
        End If
        If Sheet27.Range("FA_A2_ClosingBal").Locked = False Then
            Sheet27.Range("FA_A2_ClosingBal").ClearContents
        End If
        If Sheet27.Range("FA_A2_Grossinterest_Nature").Locked = False Then
            Sheet27.Range("FA_A2_Grossinterest_Nature").ClearContents
        End If
         If Sheet27.Range("FA_A2_Grossinterest").Locked = False Then
            Sheet27.Range("FA_A2_Grossinterest").ClearContents
        End If
    End If
    

    If (TotalDiffRow > 0) Then
     AddRows_A2_FA2 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A2_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, CustodialAcc_CountryColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_CountryColNo).value = sCountry
            End If
            
            If Sheet27.Cells(rowcount, CustodialAcc_BanknameColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_BanknameColNo).value = Node("financialInstName")
            End If
            If Sheet27.Cells(rowcount, CustodialAcc_AddressOfBankColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_AddressOfBankColNo).value = Node("financialInstAddress")
            End If
            If Sheet27.Cells(rowcount, CustodialAcc_ZipCodeColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_ZipCodeColNo).value = Node("zipCode")
            End If
            
            
            iStatus = Node("status")
            If iStatus = "BENIFICIARY" Then
                sStatus = "BENIFICIARY"
            ElseIf iStatus = "BENEFICIAL_OWNER" Then
                 sStatus = "BENEFICIAL_OWNER"
            ElseIf iStatus = "OWNER" Then
                 sStatus = "OWNER"
            Else
                sStatus = "(Select)"
            End If
                 
            If Sheet27.Cells(rowcount, CustodialAcc_OwnerStatusColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_OwnerStatusColNo).value = sStatus
            End If
             
            If Sheet27.Cells(rowcount, CustodialAcc_ForeignAccountNumberColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_ForeignAccountNumberColNo).value = Trim(Node("accountNumber"))
            End If
            
            Date_1 = Node("accOpenDate")
            If Date_1 <> "" Then
                YYYY = Mid(Date_1, 1, 4)
                MM = Mid(Date_1, 6, 2)
                DD = Mid(Date_1, 9, 2)
                Date_1 = DD & "/" & MM & "/" & YYYY
                Sheet27.Cells(rowcount, CustodialAcc_AccOpenDateColNo) = Date_1
            End If
'            Change 26.03.2023.102.Commented because of Defect sheet mentioned that prefill possibility is no for this field(defect 20.03.2023.).
'            If Sheet27.Cells(rowcount, CustodialAcc_PeakBalColNo).Locked = False Then
'                Sheet27.Cells(rowcount, CustodialAcc_PeakBalColNo).value = Node("peakBalanceDuringPeriod")
'            End If
            If Sheet27.Cells(rowcount, CustodialAcc_ClosingBalColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_ClosingBalColNo).value = Node("closingBalance")
            End If
           If Sheet27.Cells(rowcount, CustodialAcc_Grossinterest_NatureColNo).Locked = False Then
                Dim valNatureOfAmount
                If Node("NatureOfAmount") = "I" Then
                    valNatureOfAmount = "Interest"
                ElseIf Node("NatureOfAmount") = "D" Then
                    valNatureOfAmount = "Dividend"
                ElseIf Node("NatureOfAmount") = "S" Then
                   valNatureOfAmount = "Proceeds from sale or redemption of financial assets"
                ElseIf Node("NatureOfAmount") = "O" Then
                    valNatureOfAmount = "Other income"
                ElseIf Node("NatureOfAmount") = "N" Then
                    valNatureOfAmount = "No Amount paid/credited"
                End If
                Sheet27.Cells(rowcount, CustodialAcc_Grossinterest_NatureColNo).value = valNatureOfAmount
            End If
             If Sheet27.Cells(rowcount, CustodialAcc_FA_A2_GrossinterestColNo).Locked = False Then
                Sheet27.Cells(rowcount, CustodialAcc_FA_A2_GrossinterestColNo).value = Node("grossAmtPaidCredited")
            End If
            
cnt = cnt + 1
Next Node
RecFCAcc = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("dtlsForeignEquityDebtInterest") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("dtlsForeignEquityDebtInterest")

    FEquity_CountryColNo = Sheet27.Range("FA_A3_Country").Column
    FEquity_BankNameColNo = Sheet27.Range("FA_A3_BankName").Column
    FEquity_AddressOfBankColNo = Sheet27.Range("FA_A3_BankAdd").Column
    FEquity_ZipCodeColNo = Sheet27.Range("FA_A3_ZipCode").Column
    FEquity_NatureOfEntityColNo = Sheet27.Range("FA_A3_NatureOfEntity").Column
    FEquity_AccOpenDateColNo = Sheet27.Range("FA_A3_AccOpeningDate").Column
    FEquity_InitialvalueColNo = Sheet27.Range("FA_A3_initialvalue").Column
    FEquity_PeakBalColNo = Sheet27.Range("FA_A3_PeakBal").Column
    FEquity_ClosingBalColNo = Sheet27.Range("FA_A3_ClosingBal").Column
    FEquity_TotalgrossamountColNo = Sheet27.Range("FA_A3_Totalgrossamount").Column
    FEquity_TotalgrosproceedsColNo = Sheet27.Range("FA_A3_Totalgrosproceeds").Column
       
    
    TotalExRow = Sheet27.Range("FA_A3_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_A3_Country").Locked = False Then
            Sheet27.Range("FA_A3_Country").ClearContents
        End If
        If Sheet27.Range("FA_A3_BankName").Locked = False Then
            Sheet27.Range("FA_A3_BankName").ClearContents
        End If
        If Sheet27.Range("FA_A3_BankAdd").Locked = False Then
            Sheet27.Range("FA_A3_BankAdd").ClearContents
        End If
        If Sheet27.Range("FA_A3_ZipCode").Locked = False Then
            Sheet27.Range("FA_A3_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_A3_NatureOfEntity").Locked = False Then
            Sheet27.Range("FA_A3_NatureOfEntity").ClearContents
        End If
        If Sheet27.Range("FA_A3_initialvalue").Locked = False Then
            Sheet27.Range("FA_A3_initialvalue").ClearContents
        End If
        If Sheet27.Range("FA_A3_AccOpeningDate").Locked = False Then
            Sheet27.Range("FA_A3_AccOpeningDate").ClearContents
        End If
        If Sheet27.Range("FA_A3_PeakBal").Locked = False Then
            Sheet27.Range("FA_A3_PeakBal").ClearContents
        End If
        If Sheet27.Range("FA_A3_ClosingBal").Locked = False Then
            Sheet27.Range("FA_A3_ClosingBal").ClearContents
        End If
        If Sheet27.Range("FA_A3_Totalgrossamount").Locked = False Then
            Sheet27.Range("FA_A3_Totalgrossamount").ClearContents
        End If
        If Sheet27.Range("FA_A3_Totalgrosproceeds").Locked = False Then
            Sheet27.Range("FA_A3_Totalgrosproceeds").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A3_FA3 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A3_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FEquity_CountryColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_CountryColNo).value = sCountry
            End If
            If Sheet27.Cells(rowcount, FEquity_BankNameColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_BankNameColNo).value = Node("nameOfEntity")
            End If
            If Sheet27.Cells(rowcount, FEquity_AddressOfBankColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_AddressOfBankColNo).value = Node("addressOfEntity")
            End If
            If Sheet27.Cells(rowcount, FEquity_ZipCodeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_ZipCodeColNo).value = Node("zipCode")
            End If
            If Sheet27.Cells(rowcount, FEquity_NatureOfEntityColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_NatureOfEntityColNo).value = Node("natureOfEntity")
            End If
            If Sheet27.Cells(rowcount, FEquity_InitialvalueColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_InitialvalueColNo).value = Node("initialValOfInvstmnt")
            End If
            
            Date_1 = Node("interestAcquiringDate")
            If Date_1 <> "" Then
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            End If
            Sheet27.Cells(rowcount, FEquity_AccOpenDateColNo).value = Date_1
'            Change 26.03.2023.102.Commented because of Defect sheet mentioned that prefill possibility is no for this field(defect 20.03.2023.).
'            If Sheet27.Cells(rowcount, FEquity_PeakBalColNo).Locked = False Then
'                Sheet27.Cells(rowcount, FEquity_PeakBalColNo).value = Node("peakBalanceDuringPeriod")
'            End If
            If Sheet27.Cells(rowcount, FEquity_ClosingBalColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_ClosingBalColNo).value = Node("closingBalance")
            End If
            If Sheet27.Cells(rowcount, FEquity_TotalgrossamountColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_TotalgrossamountColNo).value = Node("totGrossAmtPaidCredited")
            End If
            If Sheet27.Cells(rowcount, FEquity_TotalgrosproceedsColNo).Locked = False Then
                Sheet27.Cells(rowcount, FEquity_TotalgrosproceedsColNo).value = Node("totGrossProceeds")
            End If
            
cnt = cnt + 1
Next Node
RecFEqAcc = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("dtlsForeignCashValueInsurance") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("dtlsForeignCashValueInsurance")

    FA_A4_CountryColNo = Sheet27.Range("FA_A4_Country").Column
    FA_A4_BankNameColNo = Sheet27.Range("FA_A4_BankName").Column
    FA_A4_BankAddColNo = Sheet27.Range("FA_A4_BankAdd").Column
    FA_A4_ZipCodeColNo = Sheet27.Range("FA_A4_ZipCode").Column
    FA_A4_AccOpeningDateColNo = Sheet27.Range("FA_A4_AccOpeningDate").Column
    FA_A4_the_cashvalue = Sheet27.Range("FA_A4_the_cashvalue").Column
    FA_A4_Total_grossamount = Sheet27.Range("FA_A4_Total_grossamount").Column
    
    
    TotalExRow = Sheet27.Range("FA_A4_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_A4_Country").Locked = False Then
            Sheet27.Range("FA_A4_Country").ClearContents
        End If
        If Sheet27.Range("FA_A4_BankName").Locked = False Then
            Sheet27.Range("FA_A4_BankName").ClearContents
        End If
        If Sheet27.Range("FA_A4_BankAdd").Locked = False Then
            Sheet27.Range("FA_A4_BankAdd").ClearContents
        End If
        If Sheet27.Range("FA_A4_ZipCode").Locked = False Then
            Sheet27.Range("FA_A4_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_A4_AccOpeningDate").Locked = False Then
            Sheet27.Range("FA_A4_AccOpeningDate").ClearContents
        End If
        If Sheet27.Range("FA_A4_the_cashvalue").Locked = False Then
            Sheet27.Range("FA_A4_the_cashvalue").ClearContents
        End If
        If Sheet27.Range("FA_A4_Total_grossamount").Locked = False Then
            Sheet27.Range("FA_A4_Total_grossamount").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddRows_A4_FA4 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_A4_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
                        
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FA_A4_CountryColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_A4_CountryColNo).value = sCountry
            End If
            If Sheet27.Cells(rowcount, FA_A4_BankNameColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_A4_BankNameColNo).value = Node("financialInstName")
            End If
            If Sheet27.Cells(rowcount, FA_A4_BankAddColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_A4_BankAddColNo).value = Node("financialInstAddress")
            End If
            If Sheet27.Cells(rowcount, FA_A4_ZipCodeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_A4_ZipCodeColNo).value = Node("zipCode")
            End If
            
            
            Date_1 = Node("contractDate")
            If Date_1 <> "" Then
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            End If
            Sheet27.Cells(rowcount, FA_A4_AccOpeningDateColNo).value = Date_1
            
            If Sheet27.Cells(rowcount, FA_A4_the_cashvalue).Locked = False Then
                Sheet27.Cells(rowcount, FA_A4_the_cashvalue).value = Node("cashValOrSurrenderVal")
            End If
            If Sheet27.Cells(rowcount, FA_A4_Total_grossamount).Locked = False Then
                Sheet27.Cells(rowcount, FA_A4_Total_grossamount).value = Node("totGrossAmtPaidCredited")
            End If
            
cnt = cnt + 1
Next Node
RecFCashIns = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsFinancialInterest") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("detailsFinancialInterest")
Dim FA_B_CountryColNo, FA_B_ZipCodeColNo, FA_B_EntityNatureColNo, FA_B_EntityNameColNo, FA_B_EntityAddColNo, FA_B_NatureOfInterestvalueColNo, FA_B_DateSinceHeldColNo, FA_B_TotalInvColNo, FA_B_IncomeInterestColNo
    Dim FA_B_NatureOfIncomeColNo, FA_B_AmountColNo, FA_B_ScheduleOfferedColNo, FA_B_ItemNumOfScheduleColNo
    FA_B_CountryColNo = Sheet27.Range("FA_B_Country").Column
    FA_B_ZipCodeColNo = Sheet27.Range("FA_B_ZipCode").Column
    FA_B_EntityNatureColNo = Sheet27.Range("FA_B_EntityNature").Column
    FA_B_EntityNameColNo = Sheet27.Range("FA_B_EntityName").Column
    FA_B_EntityAddColNo = Sheet27.Range("FA_B_EntityAdd").Column
    FA_B_NatureOfInterestvalueColNo = Sheet27.Range("FA_B_NatureOfInterest").Column
    FA_B_DateSinceHeldColNo = Sheet27.Range("FA_B_DateSinceHeld").Column
    FA_B_TotalInvColNo = Sheet27.Range("FA_B_TotalInv").Column
    FA_B_IncomeInterestColNo = Sheet27.Range("FA_B_IncomeInterest").Column
    FA_B_NatureOfIncomeColNo = Sheet27.Range("FA_B_NatureOfIncome").Column
    FA_B_AmountColNo = Sheet27.Range("FA_B_Amount").Column
    FA_B_ScheduleOfferedColNo = Sheet27.Range("FA_B_ScheduleOffered").Column
    FA_B_ItemNumOfScheduleColNo = Sheet27.Range("FA_B_ItemNumOfSchedule").Column
    
    
    TotalExRow = Sheet27.Range("FA_B_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_B_Country").Locked = False Then
            Sheet27.Range("FA_B_Country").ClearContents
        End If
        If Sheet27.Range("FA_B_ZipCode").Locked = False Then
            Sheet27.Range("FA_B_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_B_EntityNature").Locked = False Then
            Sheet27.Range("FA_B_EntityNature").ClearContents
        End If
        If Sheet27.Range("FA_B_EntityName").Locked = False Then
            Sheet27.Range("FA_B_EntityName").ClearContents
        End If
        If Sheet27.Range("FA_B_EntityAdd").Locked = False Then
            Sheet27.Range("FA_B_EntityAdd").ClearContents
        End If
        If Sheet27.Range("FA_B_NatureOfInterest").Locked = False Then
            Sheet27.Range("FA_B_NatureOfInterest").ClearContents
        End If
        If Sheet27.Range("FA_B_DateSinceHeld").Locked = False Then
            Sheet27.Range("FA_B_DateSinceHeld").ClearContents
        End If
        If Sheet27.Range("FA_B_TotalInv").Locked = False Then
            Sheet27.Range("FA_B_TotalInv").ClearContents
        End If
        If Sheet27.Range("FA_B_IncomeInterest").Locked = False Then
            Sheet27.Range("FA_B_IncomeInterest").ClearContents
        End If
        If Sheet27.Range("FA_B_NatureOfIncome").Locked = False Then
            Sheet27.Range("FA_B_NatureOfIncome").ClearContents
        End If
        If Sheet27.Range("FA_B_Amount").Locked = False Then
            Sheet27.Range("FA_B_Amount").ClearContents
        End If
        If Sheet27.Range("FA_B_ScheduleOffered").Locked = False Then
            Sheet27.Range("FA_B_ScheduleOffered").ClearContents
        End If
        If Sheet27.Range("FA_B_ItemNumOfSchedule").Locked = False Then
            Sheet27.Range("FA_B_ItemNumOfSchedule").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_B1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_B_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
                        
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FA_B_CountryColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_CountryColNo).value = sCountry
            End If
            If Sheet27.Cells(rowcount, FA_B_ZipCodeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_ZipCodeColNo).value = Node("zipCode")
            End If
            If Sheet27.Cells(rowcount, FA_B_EntityNatureColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_EntityNatureColNo).value = Node("natureOfEntity")
            End If
            If Sheet27.Cells(rowcount, FA_B_EntityNameColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_EntityNameColNo).value = Node("nameOfEntity")
            End If
            If Sheet27.Cells(rowcount, FA_B_EntityAddColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_EntityAddColNo).value = Node("addressOfEntity")
            End If
            If Sheet27.Cells(rowcount, FA_B_NatureOfInterestvalueColNo).Locked = False Then
                 iStatus = Node("natureOfInt")
                If iStatus = "DIRECT" Then
                    sStatus = "DIRECT"
                ElseIf iStatus = "BENEFICIAL_OWNER" Then
                     sStatus = "BENEFICIAL OWNER"
'                ElseIf iStatus = "OWNER" Then
'                     sStatus = "OWNER"
     'Added by Shrutika(12/03/2025)SIT 89286
                ElseIf iStatus = "BENIFICIARY" Then
                     sStatus = "BENIFICIARY"
                Else
                    sStatus = "(Select)"
                End If
                Sheet27.Cells(rowcount, FA_B_NatureOfInterestvalueColNo).value = sStatus
            End If
            If Sheet27.Cells(rowcount, FA_B_DateSinceHeldColNo).Locked = False Then
                Date_1 = Node("dateHeld")
                If Date_1 <> "" Then
                YYYY = Mid(Date_1, 1, 4)
                MM = Mid(Date_1, 6, 2)
                DD = Mid(Date_1, 9, 2)
                Date_1 = DD & "/" & MM & "/" & YYYY
                End If
                Sheet27.Cells(rowcount, FA_B_DateSinceHeldColNo).value = Date_1
           End If
            If Sheet27.Cells(rowcount, FA_B_TotalInvColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_TotalInvColNo).value = Node("totalInvestment")
            End If
            If Sheet27.Cells(rowcount, FA_B_IncomeInterestColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_IncomeInterestColNo).value = Node("incFromInt")
            End If
             If Sheet27.Cells(rowcount, FA_B_NatureOfIncomeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_NatureOfIncomeColNo).value = Node("natureOfInc")
            End If
             If Sheet27.Cells(rowcount, FA_B_AmountColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_AmountColNo).value = Node("incTaxAmt")
            End If
             If Sheet27.Cells(rowcount, FA_B_ScheduleOfferedColNo).Locked = False Then
               If Node("incTaxSch") = "OS" Then
                   IncTaxSch = "Other sources"
                ElseIf Node("IncTaxSch") = "BU" Then
                   IncTaxSch = "Business"
                ElseIf Node("IncTaxSch") = "HP" Then
                   IncTaxSch = "House Property"
                ElseIf Node("IncTaxSch") = "CG" Then
                   IncTaxSch = "Capital Gains"
                ElseIf Node("IncTaxSch") = "SA" Then
                   IncTaxSch = "Salary"
                ElseIf Node("IncTaxSch") = "NI" Then
                   IncTaxSch = "No Income during the year"
                ElseIf Node("IncTaxSch") = "EI" Then
                   IncTaxSch = "Exempt Income"
                End If
                Sheet27.Cells(rowcount, FA_B_ScheduleOfferedColNo).value = IncTaxSch
            End If
             If Sheet27.Cells(rowcount, FA_B_ItemNumOfScheduleColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_B_ItemNumOfScheduleColNo).value = Trim(Node("incTaxSchNo"))
            End If
            
cnt = cnt + 1
Next Node
RecFCashIns = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsImmovableProperty") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("detailsImmovableProperty")

    FA_C_CountryColNo = Sheet27.Range("FA_C_Country").Column
    FA_C_ZipCodeColNo = Sheet27.Range("FA_C_ZipCode").Column
    FA_C_PropertyAddColNo = Sheet27.Range("FA_C_PropertyAdd").Column
    FA_C_OwnershipBeneficiaryColNo = Sheet27.Range("FA_C_OwnershipBeneficiary").Column
    FA_C_DateOfAcquisitionColNo = Sheet27.Range("FA_C_DateOfAcquisition").Column
    FA_C_TotalInvColNo = Sheet27.Range("FA_C_TotalInv").Column
    FA_C_IncomeFromPropertyColNo = Sheet27.Range("FA_C_IncomeFromProperty").Column
    FA_C_NatureOfIncomeColNo = Sheet27.Range("FA_C_NatureOfIncome").Column
    FA_C_AmountColNo = Sheet27.Range("FA_C_Amount").Column
    FA_C_ScheduleOfferedColNo = Sheet27.Range("FA_C_ScheduleOffered").Column
    FA_C_ItemNumOfScheduleColNo = Sheet27.Range("FA_C_ItemNumOfSchedule").Column
    
    
    TotalExRow = Sheet27.Range("FA_C_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_C_Country").Locked = False Then
            Sheet27.Range("FA_C_Country").ClearContents
        End If
        If Sheet27.Range("FA_C_ZipCode").Locked = False Then
            Sheet27.Range("FA_C_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_C_PropertyAdd").Locked = False Then
            Sheet27.Range("FA_C_PropertyAdd").ClearContents
        End If
        If Sheet27.Range("FA_C_OwnershipBeneficiary").Locked = False Then
            Sheet27.Range("FA_C_OwnershipBeneficiary").ClearContents
        End If
        If Sheet27.Range("FA_C_DateOfAcquisition").Locked = False Then
            Sheet27.Range("FA_C_DateOfAcquisition").ClearContents
        End If
        If Sheet27.Range("FA_C_TotalInv").Locked = False Then
            Sheet27.Range("FA_C_TotalInv").ClearContents
        End If
        If Sheet27.Range("FA_C_IncomeFromProperty").Locked = False Then
            Sheet27.Range("FA_C_IncomeFromProperty").ClearContents
        End If
        If Sheet27.Range("FA_C_NatureOfIncome").Locked = False Then
            Sheet27.Range("FA_C_NatureOfIncome").ClearContents
        End If
        If Sheet27.Range("FA_C_Amount").Locked = False Then
            Sheet27.Range("FA_C_Amount").ClearContents
        End If
        If Sheet27.Range("FA_C_ScheduleOffered").Locked = False Then
            Sheet27.Range("FA_C_ScheduleOffered").ClearContents
        End If
        If Sheet27.Range("FA_C_ItemNumOfSchedule").Locked = False Then
            Sheet27.Range("FA_C_ItemNumOfSchedule").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_FA_C1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_C_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
                        
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            Sheet27.Cells(rowcount, FA_C_CountryColNo).value = sCountry
            Sheet27.Cells(rowcount, FA_C_PropertyAddColNo).value = Node("addressOfProperty")
            Sheet27.Cells(rowcount, FA_C_ZipCodeColNo).value = Node("zipCode")
            
            iStatus = Node("ownership")
            If iStatus = "BENIFICIARY" Then
                sStatus = "BENIFICIARY"
            ElseIf iStatus = "BENEFICIAL_OWNER" Then
                 sStatus = "BENEFICIAL_OWNER"
            ElseIf iStatus = "DIRECT" Then
                 sStatus = "DIRECT"
            Else
                sStatus = "(Select)"
            End If
                 
            If Sheet27.Cells(rowcount, FA_C_OwnershipBeneficiaryColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_C_OwnershipBeneficiaryColNo).value = sStatus
            End If
            
            Date_1 = Node("dateOfAcq")
            If Date_1 <> "" Then
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            End If
            Sheet27.Cells(rowcount, FA_C_DateOfAcquisitionColNo).value = Date_1
            
            If Sheet27.Cells(rowcount, FA_C_TotalInvColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_C_TotalInvColNo).value = Node("totalInvestment")
            End If
            If Sheet27.Cells(rowcount, FA_C_IncomeFromPropertyColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_C_IncomeFromPropertyColNo).value = Node("incDrvProperty")
            End If
            If Sheet27.Cells(rowcount, FA_C_NatureOfIncomeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_C_NatureOfIncomeColNo).value = Node("natureOfInc")
            End If
            If Sheet27.Cells(rowcount, FA_C_AmountColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_C_AmountColNo).value = Node("incTaxAmt")
            End If
            If Sheet27.Cells(rowcount, FA_C_ScheduleOfferedColNo).Locked = False Then
                 
                If Node("incTaxSch") = "OS" Then
                   IncTaxSch = "Other sources"
                ElseIf Node("IncTaxSch") = "BU" Then
                   IncTaxSch = "Business"
                ElseIf Node("IncTaxSch") = "HP" Then
                   IncTaxSch = "House Property"
                ElseIf Node("IncTaxSch") = "CG" Then
                   IncTaxSch = "Capital Gains"
                ElseIf Node("IncTaxSch") = "SA" Then
                   IncTaxSch = "Salary"
                ElseIf Node("IncTaxSch") = "NI" Then
                   IncTaxSch = "No Income during the year"
                ElseIf Node("IncTaxSch") = "EI" Then
                   IncTaxSch = "Exempt Income"
                End If
                Sheet27.Cells(rowcount, FA_C_ScheduleOfferedColNo).value = IncTaxSch
            End If
            If Sheet27.Cells(rowcount, FA_C_ItemNumOfScheduleColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_C_ItemNumOfScheduleColNo).value = Trim(Node("incTaxSchNo"))
            End If
            
cnt = cnt + 1
Next Node
RecImProp = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsOthAssets") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("detailsOthAssets")


    FA_D_CountryColNo = Sheet27.Range("FA_D_Country").Column
    FA_D_ZipCodeColNo = Sheet27.Range("FA_D_ZipCode").Column
    FA_D_AssetNatureColNo = Sheet27.Range("FA_D_AssetNature").Column
    FA_D_OwnershipBeneficiaryColNo = Sheet27.Range("FA_D_OwnershipBeneficiary").Column
    FA_D_DateOfAcquisitionColNo = Sheet27.Range("FA_D_DateOfAcquisition").Column
    FA_D_TotalInvColNo = Sheet27.Range("FA_D_TotalInv").Column
    FA_D_IncomeFromPropertyColNo = Sheet27.Range("FA_D_IncomeFromProperty").Column
    FA_D_NatureOfIncomeColNo = Sheet27.Range("FA_D_NatureOfIncome").Column
    FA_D_AmountColNo = Sheet27.Range("FA_D_Amount").Column
    FA_D_ScheduleOfferedColNo = Sheet27.Range("FA_D_ScheduleOffered").Column
    FA_D_ItemNumOfScheduleColNo = Sheet27.Range("FA_D_ItemNumOfSchedule").Column
    
    
    
    TotalExRow = Sheet27.Range("FA_D_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_D_Country").Locked = False Then
            Sheet27.Range("FA_D_Country").ClearContents
        End If
        If Sheet27.Range("FA_D_ZipCode").Locked = False Then
            Sheet27.Range("FA_D_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_D_AssetNature").Locked = False Then
            Sheet27.Range("FA_D_AssetNature").ClearContents
        End If
        If Sheet27.Range("FA_D_OwnershipBeneficiary").Locked = False Then
            Sheet27.Range("FA_D_OwnershipBeneficiary").ClearContents
        End If
        If Sheet27.Range("FA_D_DateOfAcquisition").Locked = False Then
            Sheet27.Range("FA_D_DateOfAcquisition").ClearContents
        End If
        If Sheet27.Range("FA_D_TotalInv").Locked = False Then
            Sheet27.Range("FA_D_TotalInv").ClearContents
        End If
        If Sheet27.Range("FA_D_IncomeFromProperty").Locked = False Then
            Sheet27.Range("FA_D_IncomeFromProperty").ClearContents
        End If
        If Sheet27.Range("FA_D_NatureOfIncome").Locked = False Then
            Sheet27.Range("FA_D_NatureOfIncome").ClearContents
        End If
        If Sheet27.Range("FA_D_Amount").Locked = False Then
            Sheet27.Range("FA_D_Amount").ClearContents
        End If
        If Sheet27.Range("FA_D_ScheduleOffered").Locked = False Then
            Sheet27.Range("FA_D_ScheduleOffered").ClearContents
        End If
        If Sheet27.Range("FA_D_ItemNumOfSchedule").Locked = False Then
            Sheet27.Range("FA_D_ItemNumOfSchedule").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_D1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_D_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
            
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FA_D_CountryColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_CountryColNo).value = sCountry
            End If
            If Sheet27.Cells(rowcount, FA_D_AssetNatureColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_AssetNatureColNo).value = Node("natureOfAsset")
            End If
            If Sheet27.Cells(rowcount, FA_D_ZipCodeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_ZipCodeColNo).value = Node("zipCode")
            End If
            
            iStatus = Node("ownership")
            If iStatus = "BENIFICIARY" Then
                sStatus = "BENIFICIARY"
            ElseIf iStatus = "BENEFICIAL_OWNER" Then
                 sStatus = "BENEFICIAL_OWNER"
            ElseIf iStatus = "DIRECT" Then
                 sStatus = "DIRECT"
            Else
                sStatus = "(Select)"
            End If
                 
            If Sheet27.Cells(rowcount, FA_D_OwnershipBeneficiaryColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_OwnershipBeneficiaryColNo).value = sStatus
            End If
            
            
            Date_1 = Node("dateOfAcq")
            If Date_1 <> "" Then
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            End If
            Sheet27.Cells(rowcount, FA_D_DateOfAcquisitionColNo).value = Date_1
            
            If Sheet27.Cells(rowcount, FA_D_TotalInvColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_TotalInvColNo).value = Node("totalInvestment")
            End If
            If Sheet27.Cells(rowcount, FA_D_IncomeFromPropertyColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_IncomeFromPropertyColNo).value = Node("incDrvAsset")
            End If
            If Sheet27.Cells(rowcount, FA_D_NatureOfIncomeColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_NatureOfIncomeColNo).value = Node("natureOfInc")
            End If
            If Sheet27.Cells(rowcount, FA_D_AmountColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_AmountColNo).value = Node("incTaxAmt")
            End If
            If Sheet27.Cells(rowcount, FA_D_ScheduleOfferedColNo).Locked = False Then
               If Node("incTaxSch") = "OS" Then
                   IncTaxSch = "Other sources"
                ElseIf Node("IncTaxSch") = "BU" Then
                   IncTaxSch = "Business"
                ElseIf Node("IncTaxSch") = "HP" Then
                   IncTaxSch = "House Property"
                ElseIf Node("IncTaxSch") = "CG" Then
                   IncTaxSch = "Capital Gains"
                ElseIf Node("IncTaxSch") = "SA" Then
                   IncTaxSch = "Salary"
                ElseIf Node("IncTaxSch") = "NI" Then
                   IncTaxSch = "No Income during the year"
                ElseIf Node("IncTaxSch") = "EI" Then
                   IncTaxSch = "Exempt Income"
                End If
                Sheet27.Cells(rowcount, FA_D_ScheduleOfferedColNo).value = IncTaxSch
            End If
            If Sheet27.Cells(rowcount, FA_D_ItemNumOfScheduleColNo).Locked = False Then
                Sheet27.Cells(rowcount, FA_D_ItemNumOfScheduleColNo).value = Trim(Node("incTaxSchNo"))
            End If
            
cnt = cnt + 1
Next Node
RecOthAssts = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsOfAccntsHvngSigningAuth") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("detailsOfAccntsHvngSigningAuth")

    FA_E_InstName = Sheet27.Range("FA_E_InstName").Column
    FA_E_InstAdd = Sheet27.Range("FA_E_InstAdd").Column
    FA_E_Country = Sheet27.Range("FA_E_Country").Column
    FA_E_ZipCode = Sheet27.Range("FA_E_ZipCode").Column
    FA_E_AccountName = Sheet27.Range("FA_E_AccountName").Column
    FA_E_InstitutionAccountNumber = Sheet27.Range("FA_E_InstitutionAccountNumber").Column
    FA_E_PeakBalInv = Sheet27.Range("FA_E_PeakBalInv").Column
    FA_E_IncomeAccured = Sheet27.Range("FA_E_IncomeAccured").Column
    FA_E_IncomeAccuredAmount = Sheet27.Range("FA_E_IncomeAccuredAmount").Column
    FA_E_Amount = Sheet27.Range("FA_E_Amount").Column
    FA_E_Scheduleoffered = Sheet27.Range("FA_E_Scheduleoffered").Column
    FA_E_ItemNumberSchedule = Sheet27.Range("FA_E_ItemNumberSchedule").Column
    
    TotalExRow = Sheet27.Range("FA_E_InstName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_E_InstName").Locked = False Then
            Sheet27.Range("FA_E_InstName").ClearContents
        End If
        If Sheet27.Range("FA_E_ZipCode").Locked = False Then
            Sheet27.Range("FA_E_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_E_InstAdd").Locked = False Then
            Sheet27.Range("FA_E_InstAdd").ClearContents
        End If
        If Sheet27.Range("FA_E_AccountName").Locked = False Then
            Sheet27.Range("FA_E_AccountName").ClearContents
        End If
        If Sheet27.Range("FA_E_InstitutionAccountNumber").Locked = False Then
            Sheet27.Range("FA_E_InstitutionAccountNumber").ClearContents
        End If
        If Sheet27.Range("FA_E_Country").Locked = False Then
            Sheet27.Range("FA_E_Country").ClearContents
        End If
        If Sheet27.Range("FA_E_PeakBalInv").Locked = False Then
            Sheet27.Range("FA_E_PeakBalInv").ClearContents
        End If
        If Sheet27.Range("FA_E_IncomeAccured").Locked = False Then
            Sheet27.Range("FA_E_IncomeAccured").ClearContents
        End If
        If Sheet27.Range("FA_E_IncomeAccuredAmount").Locked = False Then
            Sheet27.Range("FA_E_IncomeAccuredAmount").ClearContents
        End If
        If Sheet27.Range("FA_E_Amount").Locked = False Then
            Sheet27.Range("FA_E_Amount").ClearContents
        End If
        If Sheet27.Range("FA_E_Scheduleoffered").Locked = False Then
            Sheet27.Range("FA_E_Scheduleoffered").ClearContents
        End If
        If Sheet27.Range("FA_E_ItemNumberSchedule").Locked = False Then
            Sheet27.Range("FA_E_ItemNumberSchedule").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_E1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_E_InstName").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
            
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FA_E_Country).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_Country).value = sCountry
            End If
            If Sheet27.Cells(rowcount, FA_E_InstName).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_InstName).value = Node("nameOfInstitution")
            End If
            If Sheet27.Cells(rowcount, FA_E_ZipCode).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_ZipCode).value = Node("zipCode")
            End If
            If Sheet27.Cells(rowcount, FA_E_InstAdd).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_InstAdd).value = Node("addressOfInstitution")
            End If
            If Sheet27.Cells(rowcount, FA_E_AccountName).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_AccountName).value = Node("nameMentionedInAccnt")
            End If
            If Sheet27.Cells(rowcount, FA_E_InstitutionAccountNumber).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_InstitutionAccountNumber).value = Trim(Node("institutionAccountNumber"))
            End If
            If Sheet27.Cells(rowcount, FA_E_PeakBalInv).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_PeakBalInv).value = Node("peakBalanceOrInvestment")
            End If
            Dim IncAccuredTaxFlag
            IncAccuredTaxFlag = Node("incAccuredTaxFlag")
             If IncAccuredTaxFlag <> "" Then
                    If IncAccuredTaxFlag = "Y" Then
                        IncAccuredTaxFlag = "Yes"
                    ElseIf IncAccuredTaxFlag = "N" Then
                        IncAccuredTaxFlag = "No"
                    End If
            Else
                IncAccuredTaxFlag = "(Select)"
            End If
    
            If Sheet27.Cells(rowcount, FA_E_IncomeAccured).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_IncomeAccured).value = IncAccuredTaxFlag
            End If
            If Sheet27.Cells(rowcount, FA_E_IncomeAccuredAmount).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_IncomeAccuredAmount).value = Node("incAccuredInAcc")
            End If
            If Sheet27.Cells(rowcount, FA_E_Amount).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_Amount).value = Node("incOfferedAmt")
            End If
            If Sheet27.Cells(rowcount, FA_E_Scheduleoffered).Locked = False Then
                If Node("incOfferedSch") = "OS" Then
                   IncTaxSch = "Other sources"
                ElseIf Node("incOfferedSch") = "BU" Then
                   IncTaxSch = "Business"
                ElseIf Node("incOfferedSch") = "HP" Then
                   IncTaxSch = "House Property"
                ElseIf Node("incOfferedSch") = "CG" Then
                   IncTaxSch = "Capital Gains"
                ElseIf Node("incOfferedSch") = "SA" Then
                   IncTaxSch = "Salary"
                ElseIf Node("incOfferedSch") = "NI" Then
                   IncTaxSch = "No Income during the year"
                ElseIf Node("incOfferedSch") = "EI" Then
                   IncTaxSch = "Exempt Income"
                End If
                Sheet27.Cells(rowcount, FA_E_Scheduleoffered).value = IncTaxSch
            End If
            If Sheet27.Cells(rowcount, FA_E_ItemNumberSchedule).Locked = False Then
                Sheet27.Cells(rowcount, FA_E_ItemNumberSchedule).value = Trim(Node("incOfferedSchNo"))
            End If

cnt = cnt + 1
Next Node
RecActSgnAuth = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsOfTrustOutIndiaTrustee") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("detailsOfTrustOutIndiaTrustee")

    FA_F_Country = Sheet27.Range("FA_F_Country").Column
    FA_F_ZipCode = Sheet27.Range("FA_F_ZipCode").Column
    FA_F_NameOfTrust = Sheet27.Range("FA_F_NameOfTrust").Column
    FA_F_AddressOfTrust = Sheet27.Range("FA_F_AddressOfTrust").Column
    FA_F_NameOfOtherTrustees = Sheet27.Range("FA_F_NameOfOtherTrustees").Column
    FA_F_AddressOfOtherTrustees = Sheet27.Range("FA_F_AddressOfOtherTrustees").Column
    FA_F_NameOfSettlor = Sheet27.Range("FA_F_NameOfSettlor").Column
    FA_F_AddressOfSettlor = Sheet27.Range("FA_F_AddressOfSettlor").Column
    FA_F_NameOfBeneficiaries = Sheet27.Range("FA_F_NameOfBeneficiaries").Column
    FA_F_AddressOfBeneficiaries = Sheet27.Range("FA_F_AddressOfBeneficiaries").Column
    FA_F_DateSinceHeld = Sheet27.Range("FA_F_DateSinceHeld").Column
    FA_F_IncomeDerivedIsTaxable = Sheet27.Range("FA_F_IncomeDerivedIsTaxable").Column
    FA_F_IncomeDerivedAmount = Sheet27.Range("FA_F_IncomeDerivedAmount").Column
    FA_F_Amount = Sheet27.Range("FA_F_Amount").Column
    FA_F_Scheduleoffered = Sheet27.Range("FA_F_Scheduleoffered").Column
    FA_F_ItemNumberSchedule = Sheet27.Range("FA_F_ItemNumberSchedule").Column
    
    
    TotalExRow = Sheet27.Range("FA_F_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_F_Country").Locked = False Then
            Sheet27.Range("FA_F_Country").ClearContents
        End If
        If Sheet27.Range("FA_F_ZipCode").Locked = False Then
            Sheet27.Range("FA_F_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_F_NameOfTrust").Locked = False Then
            Sheet27.Range("FA_F_NameOfTrust").ClearContents
        End If
        If Sheet27.Range("FA_F_AddressOfTrust").Locked = False Then
            Sheet27.Range("FA_F_AddressOfTrust").ClearContents
        End If
        If Sheet27.Range("FA_F_NameOfOtherTrustees").Locked = False Then
            Sheet27.Range("FA_F_NameOfOtherTrustees").ClearContents
        End If
        If Sheet27.Range("FA_F_AddressOfOtherTrustees").Locked = False Then
            Sheet27.Range("FA_F_AddressOfOtherTrustees").ClearContents
        End If
        If Sheet27.Range("FA_F_NameOfSettlor").Locked = False Then
            Sheet27.Range("FA_F_NameOfSettlor").ClearContents
        End If
        If Sheet27.Range("FA_F_AddressOfSettlor").Locked = False Then
            Sheet27.Range("FA_F_AddressOfSettlor").ClearContents
        End If
        If Sheet27.Range("FA_F_NameOfBeneficiaries").Locked = False Then
            Sheet27.Range("FA_F_NameOfBeneficiaries").ClearContents
        End If
        If Sheet27.Range("FA_F_AddressOfBeneficiaries").Locked = False Then
            Sheet27.Range("FA_F_AddressOfBeneficiaries").ClearContents
        End If
        If Sheet27.Range("FA_F_DateSinceHeld").Locked = False Then
            Sheet27.Range("FA_F_DateSinceHeld").ClearContents
        End If
        If Sheet27.Range("FA_F_IncomeDerivedIsTaxable").Locked = False Then
            Sheet27.Range("FA_F_IncomeDerivedIsTaxable").ClearContents
        End If
        If Sheet27.Range("FA_F_IncomeDerivedAmount").Locked = False Then
            Sheet27.Range("FA_F_IncomeDerivedAmount").ClearContents
        End If
        If Sheet27.Range("FA_F_Amount").Locked = False Then
            Sheet27.Range("FA_F_Amount").ClearContents
        End If
        If Sheet27.Range("FA_F_Scheduleoffered").Locked = False Then
            Sheet27.Range("FA_F_Scheduleoffered").ClearContents
        End If
        If Sheet27.Range("FA_F_ItemNumberSchedule").Locked = False Then
            Sheet27.Range("FA_F_ItemNumberSchedule").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_F1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_F_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
                    
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FA_F_Country).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_Country).value = sCountry
            End If
            If Sheet27.Cells(rowcount, FA_F_ZipCode).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_ZipCode).value = Node("zipCode")
            End If
            If Sheet27.Cells(rowcount, FA_F_NameOfTrust).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_NameOfTrust).value = Node("nameOfTrust")
            End If
            If Sheet27.Cells(rowcount, FA_F_AddressOfTrust).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_AddressOfTrust).value = Node("addressOfTrust")
            End If
            If Sheet27.Cells(rowcount, FA_F_NameOfOtherTrustees).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_NameOfOtherTrustees).value = Node("nameOfOtherTrustees")
            End If
            If Sheet27.Cells(rowcount, FA_F_AddressOfOtherTrustees).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_AddressOfOtherTrustees).value = Node("addressOfOtherTrustees")
            End If
            If Sheet27.Cells(rowcount, FA_F_NameOfSettlor).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_NameOfSettlor).value = Node("nameOfSettlor")
            End If
            If Sheet27.Cells(rowcount, FA_F_AddressOfSettlor).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_AddressOfSettlor).value = Node("addressOfSettlor")
            End If
            If Sheet27.Cells(rowcount, FA_F_NameOfBeneficiaries).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_NameOfBeneficiaries).value = Node("nameOfBeneficiaries")
            End If
            If Sheet27.Cells(rowcount, FA_F_AddressOfBeneficiaries).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_AddressOfBeneficiaries).value = Node("addressOfBeneficiaries")
            End If
            
            Date_1 = Node("dateHeld")
            If Date_1 <> "" Then
            YYYY = Mid(Date_1, 1, 4)
            MM = Mid(Date_1, 6, 2)
            DD = Mid(Date_1, 9, 2)
            Date_1 = DD & "/" & MM & "/" & YYYY
            End If
            Sheet27.Cells(rowcount, FA_F_DateSinceHeld).value = Date_1
            
             If Node("IncDrvTaxFlag") <> "" Then
                    If Node("IncDrvTaxFlag") = "Y" Then
                        If Sheet27.Cells(rowcount, FA_F_IncomeDerivedIsTaxable).Locked = False Then
                            Sheet27.Cells(rowcount, FA_F_IncomeDerivedIsTaxable).value = "Yes"
                        End If
                    ElseIf Node("IncDrvTaxFlag") = "N" Then
                       If Sheet27.Cells(rowcount, FA_F_IncomeDerivedIsTaxable).Locked = False Then
                            Sheet27.Cells(rowcount, FA_F_IncomeDerivedIsTaxable).value = "No"
                        End If
                    End If
            Else
                If Sheet27.Cells(rowcount, FA_F_IncomeDerivedIsTaxable).Locked = False Then
                            Sheet27.Cells(rowcount, FA_F_IncomeDerivedIsTaxable).value = "(Select)"
                        End If
            End If
            
            
            If Sheet27.Cells(rowcount, FA_F_IncomeDerivedAmount).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_IncomeDerivedAmount).value = Node("incDrvFromTrust")
            End If
            If Sheet27.Cells(rowcount, FA_F_Amount).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_Amount).value = Node("incOfferedAmt")
            End If
            If Sheet27.Cells(rowcount, FA_F_Scheduleoffered).Locked = False Then
               If Node("incOfferedSch") = "OS" Then
                   IncTaxSch = "Other sources"
                ElseIf Node("incOfferedSch") = "BU" Then
                   IncTaxSch = "Business"
                ElseIf Node("incOfferedSch") = "HP" Then
                   IncTaxSch = "House Property"
                ElseIf Node("incOfferedSch") = "CG" Then
                   IncTaxSch = "Capital Gains"
                ElseIf Node("incOfferedSch") = "SA" Then
                   IncTaxSch = "Salary"
                ElseIf Node("incOfferedSch") = "NI" Then
                   IncTaxSch = "No Income during the year"
                ElseIf Node("incOfferedSch") = "EI" Then
                   IncTaxSch = "Exempt Income"
                End If
                Sheet27.Cells(rowcount, FA_F_Scheduleoffered).value = IncTaxSch
            End If
            If Sheet27.Cells(rowcount, FA_F_ItemNumberSchedule).Locked = False Then
                Sheet27.Cells(rowcount, FA_F_ItemNumberSchedule).value = Trim(Node("incOfferedSchNo"))
            End If

cnt = cnt + 1
Next Node
RecTrustOIT = cnt
End If

Set init = jsonObject("lastFiledITR")("scheduleFA")
If init.exists("detailsOfOthSourcesIncOutsideIndia") Then

Set Nodelist = jsonObject("lastFiledITR")("scheduleFA")("detailsOfOthSourcesIncOutsideIndia")


    FA_G_Country = Sheet27.Range("FA_G_Country").Column
    FA_G_ZipCode = Sheet27.Range("FA_G_ZipCode").Column
    FA_G_NameOfPerson = Sheet27.Range("FA_G_NameOfPerson").Column
    FA_G_AddressOfPerson = Sheet27.Range("FA_G_AddressOfPerson").Column
    FA_G_IncomeDerived = Sheet27.Range("FA_G_IncomeDerived").Column
    FA_G_NatureOfIncome = Sheet27.Range("FA_G_NatureOfIncome").Column
    FA_G_WhetherTaxable = Sheet27.Range("FA_G_WhetherTaxable").Column
    FA_G_Amount = Sheet27.Range("FA_G_Amount").Column
    FA_G_Scheduleoffered = Sheet27.Range("FA_G_Scheduleoffered").Column
    FA_G_ItemNumberSchedule = Sheet27.Range("FA_G_ItemNumberSchedule").Column

     
    TotalExRow = Sheet27.Range("FA_G_Country").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet27.Range("FA_G_Country").Locked = False Then
            Sheet27.Range("FA_G_Country").ClearContents
        End If
        If Sheet27.Range("FA_G_ZipCode").Locked = False Then
            Sheet27.Range("FA_G_ZipCode").ClearContents
        End If
        If Sheet27.Range("FA_G_NameOfPerson").Locked = False Then
            Sheet27.Range("FA_G_NameOfPerson").ClearContents
        End If
        If Sheet27.Range("FA_G_AddressOfPerson").Locked = False Then
            Sheet27.Range("FA_G_AddressOfPerson").ClearContents
        End If
        If Sheet27.Range("FA_G_IncomeDerived").Locked = False Then
            Sheet27.Range("FA_G_IncomeDerived").ClearContents
        End If
        If Sheet27.Range("FA_G_NatureOfIncome").Locked = False Then
            Sheet27.Range("FA_G_NatureOfIncome").ClearContents
        End If
        If Sheet27.Range("FA_G_WhetherTaxable").Locked = False Then
            Sheet27.Range("FA_G_WhetherTaxable").ClearContents
        End If
        If Sheet27.Range("FA_G_Amount").Locked = False Then
            Sheet27.Range("FA_G_Amount").ClearContents
        End If
        If Sheet27.Range("FA_G_Scheduleoffered").Locked = False Then
            Sheet27.Range("FA_G_Scheduleoffered").ClearContents
        End If
        If Sheet27.Range("FA_G_ItemNumberSchedule").Locked = False Then
            Sheet27.Range("FA_G_ItemNumberSchedule").ClearContents
        End If
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_FA_G1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet27.Range("FA_G_Country").name)
    rowcount = rowcount - 1
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
            
         
            iCountry = Node("countryCode")
            sCountry = Findtext(CStr(iCountry), "Country")
         
            If Sheet27.Cells(rowcount, FA_G_Country).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_Country).value = sCountry
            End If
'Konda updated on 21-08-2025---SIT-89373
'            If Sheet27.Cells(rowcount, FA_G_ZipCode).Locked = False Then
'                Sheet27.Cells(rowcount, FA_G_ZipCode).value = node("ZipCode")
'            End If
            If Sheet27.Cells(rowcount, FA_G_ZipCode).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_ZipCode).value = Node("zipCode")
            End If
'------------------------------

            If Sheet27.Cells(rowcount, FA_G_NameOfPerson).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_NameOfPerson).value = Node("nameOfPerson")
            End If
            If Sheet27.Cells(rowcount, FA_G_AddressOfPerson).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_AddressOfPerson).value = Node("addressOfPerson")
            End If
            If Sheet27.Cells(rowcount, FA_G_IncomeDerived).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_IncomeDerived).value = Node("incDerived")
            End If
            If Sheet27.Cells(rowcount, FA_G_NatureOfIncome).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_NatureOfIncome).value = Node("natureOfInc")
            End If
            Dim IncDrvTaxFlag
            IncDrvTaxFlag = Node("incDrvTaxFlag")
             If IncDrvTaxFlag <> "" Then
                    If IncDrvTaxFlag = "Y" Then
                        IncDrvTaxFlag = "Yes"
                    ElseIf IncDrvTaxFlag = "N" Then
                        IncDrvTaxFlag = "No"
                    End If
            Else
                IncDrvTaxFlag = "(Select)"
            End If
            
            If Sheet27.Cells(rowcount, FA_G_WhetherTaxable).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_WhetherTaxable).value = IncDrvTaxFlag
            End If
            If Sheet27.Cells(rowcount, FA_G_Amount).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_Amount).value = Node("incOfferedAmt")
            End If
            If Sheet27.Cells(rowcount, FA_G_Scheduleoffered).Locked = False Then
                If Node("incOfferedSch") = "OS" Then
                   IncTaxSch = "Other sources"
                ElseIf Node("incOfferedSch") = "BU" Then
                   IncTaxSch = "Business"
                ElseIf Node("incOfferedSch") = "HP" Then
                   IncTaxSch = "House Property"
                ElseIf Node("incOfferedSch") = "CG" Then
                   IncTaxSch = "Capital Gains"
                ElseIf Node("incOfferedSch") = "SA" Then
                   IncTaxSch = "Salary"
                ElseIf Node("incOfferedSch") = "NI" Then
                   IncTaxSch = "No Income during the year"
                ElseIf Node("incOfferedSch") = "EI" Then
                   IncTaxSch = "Exempt Income"
                End If
                Sheet27.Cells(rowcount, FA_G_Scheduleoffered).value = IncTaxSch
            End If
            If Sheet27.Cells(rowcount, FA_G_ItemNumberSchedule).Locked = False Then
                Sheet27.Cells(rowcount, FA_G_ItemNumberSchedule).value = Trim(Node("incOfferedSchNo"))
            End If

cnt = cnt + 1
Next Node
RecSrcIncOI = cnt
End If
End If
End If
End Function

Function ImportPartB_TTI(jsonText As String)
On Error Resume Next

Dim jsonDictionary, jsonObject As Object
Dim Node, Nodelist, init As Object

Dim TaxDeemedTISec115JC, SurchargeOnAboveCrore, EducationCess, TotalTax, TaxAtNormalRatesOnAggrInc, TaxAtSpecialRates, RebateOnAgriInc As Variant
Dim TaxPayableOnTotInc, Rebate87A, TaxPayableOnRebate, Surcharge25ofSI As Variant
Dim TI_SurchargeOnAboveCrore, TotalSurcharge, TI_EducationCess, GrossTaxLiability As Variant
Dim GrossTaxPayable, CreditUS115JD, TaxPayAfterCreditUs115JD, Section89, Section89_1, Section90, Section91, TotTaxRelief, NetTaxLiability As Variant
Dim intrstPayUs234A, intrstPayUs234B, intrstPayUs234C, LateFilingFee234F, TotalIntrstPay, AggregateTaxInterestLiability As Variant
Dim AdvanceTax, TDS, TCS, SelfAssessmentTax, TotalTaxesPaid, BalTaxPayable, RefundDue, BankDtlsFlag As Variant
Dim IFSCColNo, BankNameColNo, ACCNOColNo, CheckBox, SWIFTColNo, IBANBankColNo, CountryCodeColNo, IBANAccNoColNo, AssetOutIndiaFlag As Variant
Dim sCountry, iCountry As Variant
Dim TotalXMLRow, rowcount, cnt, RecBads1, RecFbads  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long

Set jsonObject = ParseJson.ParseJson(jsonText)

TaxDeemedTISec115JC = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnDeemedTI")("TaxDeemedTISec115JC")
SurchargeOnAboveCrore = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnDeemedTI")("SurchargeOnAboveCrore")
EducationCess = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnDeemedTI")("SurchargeOnAboveCrore")
TotalTax = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnDeemedTI")("TotalTax")
TaxAtNormalRatesOnAggrInc = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("TaxAtNormalRatesOnAggrInc")
TaxAtSpecialRates = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("TaxAtSpecialRates")
RebateOnAgriInc = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("RebateOnAgriInc")
TaxPayableOnTotInc = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("TaxPayableOnTotInc")
Rebate87A = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("Rebate87A")
TaxPayableOnRebate = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("TaxPayableOnRebate")
Surcharge25ofSI = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("Surcharge25ofSI")
TI_SurchargeOnAboveCrore = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("SurchargeOnAboveCrore")
TotalSurcharge = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("TotalSurcharge")
TI_EducationCess = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("EducationCess")
GrossTaxLiability = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayableOnTI")("GrossTaxLiability")
GrossTaxPayable = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("GrossTaxPayable")
CreditUS115JD = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("CreditUS115JD")
TaxPayAfterCreditUs115JD = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxPayAfterCreditUs115JD")
Section89 = jsonObject("form24q")("taxComputation")("section89")
'Newly added by Bindu
Section89_1 = jsonObject("form10E")("TaxComputation")("Section89")

Section90 = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxRelief")("Section90")
Section91 = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxRelief")("Section91")
TotTaxRelief = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("TaxRelief")("TotTaxRelief")
NetTaxLiability = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("NetTaxLiability")
intrstPayUs234A = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("IntrstPay")("IntrstPayUs234A")
intrstPayUs234B = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("IntrstPay")("intrstPayUs234B")
intrstPayUs234C = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("IntrstPay")("IntrstPayUs234C")
LateFilingFee234F = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("IntrstPay")("LateFilingFee234F")
TotalIntrstPay = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("IntrstPay")("TotalIntrstPay")
AggregateTaxInterestLiability = jsonObject("PartB_TTI")("ComputationOfTaxLiability")("AggregateTaxInterestLiability")
AdvanceTax = jsonObject("PartB_TTI")("TaxPaid")("TaxesPaid")("AdvanceTax")
TDS = jsonObject("PartB_TTI")("TaxPaid")("TaxesPaid")("TDS")
TCS = jsonObject("PartB_TTI")("TaxPaid")("TaxesPaid")("TCS")
SelfAssessmentTax = jsonObject("PartB_TTI")("TaxPaid")("TaxesPaid")("SelfAssessmentTax")
TotalTaxesPaid = jsonObject("PartB_TTI")("TaxPaid")("TaxesPaid")("TotalTaxesPaid")
BalTaxPayable = jsonObject("PartB_TTI")("TaxPaid")("BalTaxPayable")
RefundDue = jsonObject("PartB_TTI")("Refund")("RefundDue")

'AY_2024_25 Change 'Malli
'BankDtlsFlag = jsonObject("PartB_TTI")("Refund")("BankAccountDtls")("BankDtlsFlag")

AssetOutIndiaFlag = jsonObject("lastFiledITR")("assetOutIndiaFlag")

'Change-27, prefill-tracker, D89
Dim primaryFlag As Integer
primaryFlag = 0
If Section89 <> "" And Sheet7.Range("Sheet9.Section89").Locked = False Then

Application.EnableEvents = False    'Malli updated on 04-05-2026--SIT-118458
    Sheet7.Range("Sheet9.Section89").value = Section89
Application.EnableEvents = True
    primaryFlag = 1
End If

'Newly added by Bindu

If Section89 = "" Then
 If Section89_1 <> "" And Sheet7.Range("Sheet9.Section89").Locked = False Then
 Application.EnableEvents = False   'Malli updated on 04-05-2026--SIT-118458
    Sheet7.Range("Sheet9.Section89").value = Section89_1
Application.EnableEvents = True
    primaryFlag = 1
  End If
End If




If primaryFlag = 0 Then
    If (Trim(jsonObject("form10E")("TaxComputation")("section89")) <> "Empty") And (Sheet7.Range("Sheet9.Section89").Locked = False) Then
        Application.EnableEvents = False    'Malli updated on 04-05-2026--SIT-118458
        Sheet7.Range("Sheet9.Section89").value = jsonObject("form10E")("TaxComputation")("section89")
        Application.EnableEvents = True
    End If
End If

'If Section89 <> "" And Sheet7.Range("Sheet9.Section89").Locked = False Then
'    Sheet7.Range("Sheet9.Section89").value = Section89
'End If



If AssetOutIndiaFlag <> "" Then
        'Konda updated as for schema  SIT-89156 on 04-03-2025
'        If AssetOutIndiaFlag = "Y" Then
'            AssetOutIndiaFlag = "Yes"
'        ElseIf AssetOutIndiaFlag = "N" Then
'            AssetOutIndiaFlag = "No"
'        End If
        If AssetOutIndiaFlag = "YES" Then
            AssetOutIndiaFlag = "Yes"
        ElseIf AssetOutIndiaFlag = "NO" Then
            AssetOutIndiaFlag = "No"
        End If
'End updated SIT-89156 on 04-03-2025
    Else
        AssetOutIndiaFlag = "(Select)"
    End If

Sheet7.Range("AOIFlag_1").value = AssetOutIndiaFlag

'Change-13, Test-037, Remark - Only 1 object is to be filled even if multiple present
'following code is commented and new is added
If jsonObject.exists("lastFiledITR") Then
    If jsonObject("lastFiledITR").exists("foreignBankDetails") Then
        Set Nodelist = jsonObject("lastFiledITR")("foreignBankDetails")
        If Sheet7.Range("Sheet9.IBAN").Locked = False Then
            Sheet7.Range("Sheet9.IBAN").value = Nodelist(1)("swiftCode")
        End If
        If Sheet7.Range("Sheet9.IBANBANK").Locked = False Then
            Sheet7.Range("Sheet9.IBANBANK").value = Nodelist(1)("bankName")
        End If
        If Sheet7.Range("Sheet9.IBANCOUNTRY").Locked = False Then
'Konda updated 21-08-2025--SIT-89233
           ' Sheet7.Range("Sheet9.IBANCOUNTRY").value = Nodelist(1)("countryCode")
    
                iCountry = Nodelist(1)("countryCode")
                sCountry = Findtext(CStr(iCountry), "Country_NoIndia")

            Sheet7.Range("Sheet9.IBANCOUNTRY").value = sCountry
'---------------SIT-89233

        End If
        If Sheet7.Range("Sheet9.IBANACCNO").Locked = False Then
            Sheet7.Range("Sheet9.IBANACCNO").value = Nodelist(1)("iban")
        End If
    End If
End If
        
   
'    Set init = jsonObject("lastFiledITR")
'    If init.exists("foreignBankDetails") Then
'
'        Set nodeList = jsonObject("lastFiledITR")("foreignBankDetails")
'
'                SWIFTColNo = Range("Sheet9.IBAN").Column
'                IBANBankColNo = Range("Sheet9.IBANBANK").Column
'                CountryCodeColNo = Range("Sheet9.IBANCOUNTRY").Column
'                IBANAccNoColNo = Range("Sheet9.IBANACCNO").Column
'
'
'            TotalExRow = Range("Sheet9.IBAN").Rows.count
'
'            TotalXMLRow = nodeList.count
'            TotalDiffRow = TotalXMLRow - TotalExRow
'
'
'            If (TotalXMLRow > 0) Then
'                If Range("Sheet9.IBAN").Locked = False Then
'                    Range("Sheet9.IBAN").ClearContents
'                End If
'                If Range("Sheet9.IBANBANK").Locked = False Then
'                    Range("Sheet9.IBANBANK").ClearContents
'                End If
'                If Range("Sheet9.IBANCOUNTRY").Locked = False Then
'                    Range("Sheet9.IBANCOUNTRY").ClearContents
'                End If
'                If Range("Sheet9.IBANACCNO").Locked = False Then
'                    Range("Sheet9.IBANACCNO").ClearContents
'                End If
'            End If
'
'            If (TotalDiffRow > 0) Then
'             AddDiffRows_FBANK (TotalDiffRow)
'            End If
'
'            rowcount = getRowNo(Sheet7.Range("Sheet9.IBAN").name)
'            rowcount = rowcount - 1
'            cnt = 0
'
'            For Each Node In nodeList
'                rowcount = rowcount + 1
'
'                    iCountry = CountryCodeColNo
'                    sCountry = Findtext(CStr(iCountry), "Country_NoIndia")
'
'                    Cells(rowcount, SWIFTColNo).value = Node("swiftCode")
'                    Cells(rowcount, IBANBankColNo).value = Node("bankName")
'                    Cells(rowcount, sCountry).value = Node("countryCode")
'                    Cells(rowcount, IBANAccNoColNo).value = Node("iban")
'
'                cnt = cnt + 1
'                Next Node
'        RecFbads = cnt
'        End If
 
End Function
Function ImportScheduleS(jsonText As String)
On Error Resume Next
Dim jsonObject, Nodelist, Node As Object
Dim Section89, iState, sState, NatureColNo, DescriptionColNo, AmountColNo, TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, iNature, sNature, cnt As Variant
Set jsonObject = ParseJson.ParseJson(jsonText)
Set Nodelist = jsonObject("form24q")("salaries")("salary")
If Nodelist <> Empty Then
For Each Node In Nodelist
    If Sheet39.Range("SAL.NameOfEmployer1").Locked = False Then
    Sheet39.Range("SAL.NameOfEmployer1").value = Node("nameOfEmployer")
End If
If Sheet39.Range("SAL.PANofEmployer1").Locked = False Then
    Sheet39.Range("SAL.PANofEmployer1").value = Node("tanOfEmployer")
End If
If Sheet39.Range("SAL.AddrDetail1").Locked = False Then
    Sheet39.Range("SAL.AddrDetail1").value = Node("addressDetail")("addDetail")
End If
    
If Sheet39.Range("SAL.CityOrTownOrDistrict1").Locked = False Then
    Sheet39.Range("SAL.CityOrTownOrDistrict1").value = Node("addressDetail")("cityOrTownOrDistrict")
End If
'Change-5(a), Test-057, Remark - the following commented statements are pasted after filling state
' and indenting is corrected
'If Sheet39.Range("SAL.PinCode1").Locked = False Then
'    Sheet39.Range("SAL.PinCode1").value = Node("addressDetail")("pinCode")
'End If
'If Sheet39.Range("SAL.ZipCode1").Locked = False Then
'    Sheet39.Range("SAL.ZipCode1").value = Node("addressDetail")("zipCode")
'End If

iState = UCase(Node("addressDetail")("stateCode"))

If Len(iState) = 1 Then
iState = "0" & iState
End If

sState = Findtext(iState, "State")
If Sheet39.Range("SAL.StateCode1").Locked = False Then
    Sheet39.Range("SAL.StateCode1").value = sState
End If
'Change-5(b), Test-057 & Test-036, Remark - As specified in above, Indenting is corrected
If Sheet39.Range("SAL.PinCode1").Locked = False Then
    Sheet39.Range("SAL.PinCode1").value = Node("addressDetail")("pinCode")
End If
If Sheet39.Range("SAL.ZipCode1").Locked = False Then
    Sheet39.Range("SAL.ZipCode1").value = Node("addressDetail")("zipCode")
End If

Next Node
End If

'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

'If Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False Then
'    Sheet39.Range("SAL.Entertainment_allowance_1").value = jsonObject("form24q")("incomeDeductions")("entertainmentAlw16Ii")
'End If
'If Sheet39.Range("SAL.DeductionUnderSection1").Locked = False Then
'    Sheet39.Range("SAL.DeductionUnderSection1").value = jsonObject("form24q")("incomeDeductions")("professionalTaxUs16Iii")
'End If
'----------------------------

    Set Nodelist = jsonObject("form24q")("allwncExemptUs10DtlsType")
    
    NatureColNo = Sheet39.Range("Salary.Nature1").Column
    DescriptionColNo = Sheet39.Range("Salary.Description1").Column
    AmountColNo = Sheet39.Range("Salary.Amount1").Column
    
    TotalExRow = Sheet39.Range("Salary.Nature1").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet39.Range("Salary.Nature1").Locked = False Then
            Sheet39.Range("Salary.Nature1").ClearContents
        End If
        If Sheet39.Range("Salary.Description1").Locked = False Then
            Sheet39.Range("Salary.Description1").ClearContents
        End If
        If Sheet39.Range("Salary.Amount1").Locked = False Then
            Sheet39.Range("Salary.Amount1").ClearContents
        End If
    End If
    
    
'If (TotalDiffRow > 0) Then
' AddDiffRows_SalaryAllowance (TotalDiffRow)
'End If
    
    rowcount = getRowNo(Sheet39.Range("Salary.Nature1").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each Node In Nodelist
    
    If Node("salNatureDesc") <> "10(13A)" Then  'Malli added AY_2025_26_V0.6
        rowcount = rowcount + 1
'Change-4, ANK-130, 26.05.2022
            iNature = "Sec " & UCase(Node("salNatureDesc"))
'            iNature = "Sec " & UCase(node("SalNatureDesc"))

            sNature = Findtext(iNature, "Salary.OthersAllowance")
            If Node("salNatureDesc") = "10(10B)(i)" Then
'            If node("SalNatureDesc") = "10(10B)(i)" Then
                sNature = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
            ElseIf Node("salNatureDesc") = "10(10B)(ii)" Then
'            ElseIf node("SalNatureDesc") = "10(10B)(ii)" Then
                sNature = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
            ElseIf Node("salNatureDesc") = "OTH" Then
'            ElseIf node("SalNatureDesc") = "OTH" Then
                sNature = "ANY OTHER"
            End If
            If Sheet39.Cells(rowcount, NatureColNo).Locked = False Then
                Sheet39.Cells(rowcount, NatureColNo).value = sNature
            End If
            If Sheet39.Cells(rowcount, DescriptionColNo).Locked = False Then
                Sheet39.Cells(rowcount, DescriptionColNo).value = UCase(Node("salNatureDesc"))
            End If
            If Sheet39.Cells(rowcount, AmountColNo).Locked = False Then
                Sheet39.Cells(rowcount, AmountColNo).value = Node("salOthAmount")
'                Sheet39.Cells(rowcount, AmountColNo).value = node("SalOthAmount")
            End If
'---end change---
        cnt = cnt + 1
        End If 'Malli  'Malli added AY_2025_26_V0.6
    Next Node
    
    
    
End Function
Function ScheduleVIAImport(jsonText As String)

On Error Resume Next

Dim jsonObject As Object

Dim SELECT80DD, SELECT80DDB, ProvisonFlag, ProvisionFlag
Set jsonObject = ParseJson.ParseJson(jsonText)

'Malli--commented removed Prefill posibility--AY_2025_26_V0.6-----------------
 
'If Sheet19.Range("scvia.Section80C").Locked = False Then
'    Sheet19.Range("scvia.Section80C").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80C")
'End If
'-----------------------------------------------------------------------------
Sheet1.LockUnlockQVIA   'SIT_93940_issue1


'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

'If Sheet19.Range("scvia.Section80CCC").Locked = False Then
'    Sheet19.Range("scvia.Section80CCC").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCC")
'End If

'If Sheet19.Range("scvia.Section80CCD_SE").Locked = False Then
'    Sheet19.Range("scvia.Section80CCD_SE").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployeeOrSE")
'End If

'If Sheet19.Range("scvia.Section80CCD1B_SE").Locked = False Then
'    Sheet19.Range("scvia.Section80CCD1B_SE").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCD1B")
'End If

'If Sheet19.Range("scvia.Section80CCD1B_SE").Locked = False Then
'    Sheet19.Range("scvia.Section80CCD1B_SE").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCD1B")
'End If
'----------------------------------

If Sheet19.Range("scvia.Section80CCD").Locked = False Then
    Sheet19.Range("scvia.Section80CCD").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployer")
End If


If Sheet19.Range("scvia.Section80CCD").Locked = False Then
    Sheet19.Range("scvia.Section80CCD").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployer")
End If


'Malli-------
'AY_2024_25 Change  Malli
'lastFiledITR.usrDeductUndChapVIAType.section80DD
'If Sheet19.Range("scvia.Section80DD").Locked = False Then
'    Sheet19.Range("scvia.Section80DD").value = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DD")
'End If
 
'V1-A_C30_AY_2024_25 OLD  'Malli
'lastFiledITR.usrDeductUndChapVIAType.section80DDB
'If Sheet19.Range("scvia.Section80DDB").Locked = False Then
'    Sheet19.Range("scvia.Section80DDB").value = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDB")
'End If


'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

''V1-A_C30_AY_2024_25  New 'Malli Changed by Bindu as per Pre-fill SRS
''lastFiledITR.usrDeductUndChapVIAType.section80DDBUsrType
'If Sheet19.Range("SELECT80DDB").Locked = False Then
'    SELECT80DDB = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDBUsrType")
'    If SELECT80DDB = "1" Then
'        Sheet19.Range("SELECT80DDB").value = "1-Self or dependent"
'    ElseIf SELECT80DDB = "2" Then
'        Sheet19.Range("SELECT80DDB").value = "2-Self or Dependent(Senior Citizen)"
'    End If
'    'Sheet19.Range("SELECT80DDB").value = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDBUsrType")
'End If


'V1-A_C16_AY_2026-27   Malli_03/02/2026
'lastFiledITR.usrDeductUndChapVIAType.nameOfSpecDisease80DDB
'Dim nameOfSpecDisease80DDB_pfl As Variant
'nameOfSpecDisease80DDB_pfl = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("nameOfSpecDisease80DDB")
'
'If nameOfSpecDisease80DDB_pfl <> "" And Sheet19.Range("Sheet19.Specified_Disease").Locked = False Then
'
''nameOfSpecDisease80DDB_pfl = "o"
'            If UCase(nameOfSpecDisease80DDB_pfl) = UCase("a") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(a) Dementia"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("b") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(b) Dystonia Musculorum Deformans"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("c") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(c) Motor Neuron Disease"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("d") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(d) Ataxia"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("e") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(e) Chorea"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("f") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(f) Hemiballismus"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("g") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(g) Aphasia"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("h") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(h) Parkinsons Disease"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("i") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(i) Malignant Cancers"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("j") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(j) Full Blown Acquired Immuno-Deficiency Syndrome (AIDS)"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("k") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(k) Chronic Renal failure"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("l") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(l) Hematological disorders"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("m") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(m) Hemophilia"
'
'            ElseIf UCase(nameOfSpecDisease80DDB_pfl) = UCase("n") Then
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(n) Thalassaemia"
'
'            Else
'               Sheet19.Range("Sheet19.Specified_Disease").value = "(Select)"
'            End If
'
'
'
'
'End If



'If Sheet19.Range("scvia.Section80TTA").Locked = False Then
'    Sheet19.Range("scvia.Section80TTA").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80TTA")
'End If
'change-21 prefill-tracker, D30
'If Sheet19.Range("scvia.Section80TTB").Locked = False Then
'    Sheet19.Range("scvia.Section80TTB").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80TTB")
'End If

'--------------------------------------
'--------------------------------------
 
'Malli--commented removed Prefill posibility--AY_2025_26_V0.6-----------------
'If Sheet19.Range("scvia.Section80E").Locked = False Then
'    Sheet19.Range("scvia.Section80E").value = jsonObject("form24q")("usrDeductUndChapVIAType")("section80E")
'End If

'----------------------------------------

'---------------------------
'VIA_C33_AY_2024_25 'Malli
'lastFiledITR.usrDeductUndChapVIAType.section80U
'If Sheet19.Range("scvia.Section80U").Locked = False Then
'    Sheet19.Range("scvia.Section80U").value = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80U")
'End If
'----------------------------

'If Sheet19.Range("scvia.Section80JJAA").Locked = False Then
'    Sheet19.Range("scvia.Section80JJAA").value = jsonObject("form10DA")("UsrDeductUndChapVIAType")("Section80JJAA")
'End If
'If Sheet19.Range("scvia.Section80JJA").Locked = False Then
'    Sheet19.Range("scvia.Section80JJA").value = jsonObject("form3CD")("ScheduleVIA")("usrDeductUndChapVIA")("Section80JJA")
'End If
'If Sheet19.Range("scvia.Section80IBA").Locked = False Then
'    Sheet19.Range("scvia.Section80IBA").value = jsonObject("form3CD")("ScheduleVIA")("usrDeductUndChapVIA")("Section80IBA")
'End If
'If Sheet19.Range("scvia.Section80GGC").Locked = False Then
'    Sheet19.Range("scvia.Section80GGC").value = jsonObject("form3CD")("ScheduleVIA")("usrDeductUndChapVIA")("Section80GGC")
'End If
'Change-26, prefill-tracker, D84
'If Sheet19.Range("scvia.Section80IAB").Locked = False Then
'    Sheet19.Range("scvia.Section80IAB").value = jsonObject("form10CCB")("form10CCBAmountDetls")(1)("deductAmountSec80")
'End If


'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

'Malli--AY_2025_26_V0.6--------------------------------

'Dim Form10BAAckNum_Pfl As Variant
'    Form10BAAckNum_Pfl = jsonObject("Form10BA")("Form10BAAckNum")
'    If Form10BAAckNum_Pfl <> "" Then
'        If SheetALL.Range("bacValue").value = 1 Then
'            Sheet19.Unprotect Password:=getmsgstate
'                Sheet19.Range("Sheet19.AckNum").Locked = False
'                Sheet19.Range("Sheet19.AckNum").Interior.Color = "&HCCFFCC"
'                Sheet19.Range("Sheet19.AckNum").value = Form10BAAckNum_Pfl
'                Sheet19.Range("Sheet19.AckNum").Font.Color = (&HD8D8D8)
'                Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HD8D8D8)
'                Sheet19.Range("Sheet19.AckNum").Locked = True
'
'        Else:
'            If Form10BAAckNum_Pfl <> "" And Sheet19.Range("Sheet19.AckNum").MergeArea.Locked = False Then
'                Sheet19.Range("Sheet19.AckNum").value = Form10BAAckNum_Pfl
'            End If
'        End If
'    End If
    
 '------------------------------------------------
    
'If Form10BAAckNum_Pfl <> "" Then
'        If Sheet19.Range("Sheet19.AckNum").Locked = True Then
'            Sheet19.Unprotect Password:=getmsgstate
'            Sheet19.Range("Sheet19.AckNum").MergeArea.Locked = False
'            Sheet19.Range("Sheet19.AckNum").MergeArea.Interior.Color = (&HCCFFCC)
'            Sheet19.Unprotect Password:=getmsgstate
'            Sheet19.Range("Sheet19.AckNum").value = Form10BAAckNum_Pfl
'        ElseIf Sheet1.Range("Sheet1.AckNum").Locked = False Then
'            Sheet19.Range("Sheet19.AckNum").value = Form10BAAckNum_Pfl
'        End If
'End If
'-------------------------------------------------------


'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

''Malli----------AY_2025_26_ITR-2_PrefillSRS_V0.3_29/05/2025
''form10CCD.Form10CCDAckNum
'Dim Form10CCDAckNum_pfl As Variant
'Form10CCDAckNum_pfl = jsonObject("form10CCD")("Form10CCDAckNum")
'
'If Form10CCDAckNum_pfl <> "" Then
'               If SheetALL.Range("bacValue").value = 1 Then
'                    Sheet19.Unprotect Password:=getmsgstate
'                        Sheet19.Range("Sheet19.AckNum10CCD").Locked = False
'                        Sheet19.Range("Sheet19.AckNum10CCD").Interior.Color = "&HCCFFCC"
'                        Sheet19.Range("Sheet19.AckNum10CCD").value = Form10CCDAckNum_pfl
'                        Sheet19.Range("Sheet19.AckNum10CCD").Font.Color = (&HD8D8D8)
'                        Sheet19.Range("Sheet19.AckNum10CCD").Interior.Color = (&HD8D8D8)
'                        Sheet19.Range("Sheet19.AckNum10CCD").Locked = True
'
'                Else:
'                        If Form10CCDAckNum_pfl <> "" And Sheet19.Range("Sheet19.AckNum10CCD").Locked = False Then
'                            Sheet19.Range("Sheet19.AckNum10CCD").value = Form10CCDAckNum_pfl
'                        End If
'                End If
'End If

'---------------------------------------------------------

''form10CCE.Form10CCEAckNum
'Dim Form10CCEAckNum_pfl As Variant
'Form10CCEAckNum_pfl = jsonObject("form10CCE")("Form10CCEAckNum")
'
'If Form10CCEAckNum_pfl <> "" Then
'        If SheetALL.Range("bacValue").value = 1 Then
'                Sheet19.Unprotect Password:=getmsgstate
'                Sheet19.Range("Sheet19.AckNum.10CCE").Locked = False
'                Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = "&HCCFFCC"
'                Sheet19.Range("Sheet19.AckNum.10CCE").value = Form10CCEAckNum_pfl
'                Sheet19.Range("Sheet19.AckNum.10CCE").Font.Color = (&HD8D8D8)
'                Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HD8D8D8)
'                Sheet19.Range("Sheet19.AckNum.10CCE").Locked = True
'
'                Else:
'
'                If Form10CCEAckNum_pfl <> "" And Sheet19.Range("Sheet19.AckNum.10CCE").Locked = False Then
'                Sheet19.Range("Sheet19.AckNum.10CCE").value = Form10CCEAckNum_pfl
'                End If
'        End If
'End If
'------------------------------------------------------------------------

'---------------------------------------------------------

End Function
Function ImportScheduleIT(jsonText As String)
On Error Resume Next
Dim jsonDictionary, init, jsonObject As Object
Dim Node, Nodelist As Object
Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo, TotalTaxPayments, TotalSchTCS As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim DepDate, Depdt As String
Dim YYYY, MM, DD As String
Set jsonObject = ParseJson.ParseJson(jsonText)


Set init = jsonObject("form26as")
If init.exists("taxPayments") Then
     
TotalTaxPayments = init("TotalTaxPayments")
 Set Nodelist = init("taxPayments")("taxPayment")
 

    BSRCodeColNo = Sheet25.Range("IT.BSRCode").Column
    DateDepColNo = Sheet25.Range("IT.DateDep").Column
    SrlNoChallanColNo = Sheet25.Range("IT.SrlNoOfChaln").Column
    AmtColNo = Sheet25.Range("IT.Amt").Column
    
TotalExRow = Range("IT.BSRCode").Rows.count
    
TotalXMLRow = Nodelist.count
TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        If Sheet25.Range("IT.BSRCode").Locked = False Then
            Sheet25.Range("IT.BSRCode").ClearContents
        End If
        If Sheet25.Range("IT.DateDep").Locked = False Then
            Sheet25.Range("IT.DateDep").ClearContents
        End If
        If Sheet25.Range("IT.SrlNoOfChaln").Locked = False Then
            Sheet25.Range("IT.SrlNoOfChaln").ClearContents
        End If
        If Sheet25.Range("IT.Amt").Locked = False Then
            Sheet25.Range("IT.Amt").ClearContents
        End If
    End If
    
If (TotalDiffRow > 0) Then
 AddDiffRows_ScheduleIT (TotalDiffRow)
End If
    
rowcount = getRowNo(Sheet25.Range("IT.BSRCode").name)
rowcount = rowcount - 1
cnt = 0
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            Sheet25.Cells(rowcount, BSRCodeColNo).value = Node("bsrCode")
            
'Change-7, ANK-139, 26.05.2022
'            Depdt = node("dateDep")
'            If Depdt <> "" Then
'                YYYY = Mid(Depdt, 1, 4)
'                MM = Mid(Depdt, 6, 2)
'                DD = Mid(Depdt, 9, 2)
'                DepDate = DD & "/" & MM & "/" & YYYY
                
                If Sheet25.Cells(rowcount, DateDepColNo).Locked = False Then
                    Sheet25.Cells(rowcount, DateDepColNo).value = Mid(Node("dateDep"), 9, 2) & "/" & Mid(Node("dateDep"), 6, 2) & "/" & Mid(Node("dateDep"), 1, 4)
'                    Sheet25.Cells(rowcount, DateDepColNo).value = node("dateDep")
                End If
'            End If
'---end change---
            
            If Sheet25.Cells(rowcount, SrlNoChallanColNo).Locked = False Then
                Sheet25.Cells(rowcount, SrlNoChallanColNo).value = Node("srlNoOfChaln")
            End If
            If Sheet25.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet25.Cells(rowcount, AmtColNo).value = Node("amt")
            End If
      
        cnt = cnt + 1
    Next Node
RecTDS1 = cnt
End If
'If TotalSchTCS <> "" Then
'    Sheet25.Range("IT.Sum").value = TotalSchTCS
'End If


End Function
'change-20 , prefill-tracker
Function ImportFilingStatus_pfl(jsonText As String)
On Error Resume Next
'change-20 , prefill-tracker, D5
Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
Dim ProvisoFlag
ProvisoFlag = jsonObject("filingStatus")("SeventhProvisio139")
    If ProvisoFlag <> "" Then
        If ProvisoFlag = "Y" Then
            ProvisoFlag = "Yes"
        ElseIf ProvisoFlag = "N" Then
            ProvisoFlag = "No"
        End If
    Else
        ProvisoFlag = "(Select)"
    End If
    
    If Sheet1.Range("Sheet1.SeventhProvisoFlag").Locked = False Then
        Sheet1.Range("Sheet1.SeventhProvisoFlag").value = ProvisoFlag
    End If
    
    
    'Malli----------AY_2025_26
'change-20 , prefill-tracker, D6
    Dim DepositAmountFlag, DepositAmount
  'insights.DepAmtAggAmtExcd1CrPrYrFlg
'insights.AmtSeventhProvisio139i
    'comented by Malli prifill tag was changed
    'AY_2024-25 OLD
   ' DepositAmountFlag = jsonObject("filingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
    'PAG_C33_AY_2025_26 CHANGE
    DepositAmountFlag = jsonObject("insights")("DepAmtAggAmtExcd1CrPrYrFlg")
    '----------------------
        If DepositAmountFlag <> "" Then
            If DepositAmountFlag = "Y" Then
                DepositAmountFlag = "Yes"
                'Malli_AY_2024_25 OLD
                'DepositAmount = jsonObject("filingStatus")("AmtSeventhProvisio139i")
                'AY_2025_26
                DepositAmount = jsonObject("insights")("AmtSeventhProvisio139i")
                '--------------------
            ElseIf DepositAmountFlag = "N" Then
                DepositAmountFlag = "No"
            End If
        Else
            DepositAmountFlag = "(Select)"
        End If
    
    If Sheet1.Range("Sheet1.DepositAmountFlag").Locked = False Then
        Sheet1.Range("Sheet1.DepositAmountFlag").value = DepositAmountFlag
    End If
    'Malli--06_11_2024
    'If DepositAmount <> Empty Then
    If DepositAmount <> "" And Sheet1.Range("Sheet1.DepositAmount").Locked = False Then
        Sheet1.Range("Sheet1.DepositAmount").value = DepositAmount
    End If
    
    '---------------------------------------------------
    '-----------------
'change-20 , prefill-tracker, D7
    Dim AggrigateAmountFlag, AggrigateAmount
    AggrigateAmountFlag = jsonObject("form26as")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
        If AggrigateAmountFlag <> "" Then
            If AggrigateAmountFlag = "Y" Then
                AggrigateAmountFlag = "Yes"
                AggrigateAmount = jsonObject("form26as")("AmtSeventhProvisio139ii")
            ElseIf AggrigateAmountFlag = "N" Then
                AggrigateAmountFlag = "No"
            End If
        Else
            AggrigateAmountFlag = "(Select)"
        End If
        
    If Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False Then
        Sheet1.Range("Sheet1.AggrigateAmountFlag").value = AggrigateAmountFlag
    End If
    If AggrigateAmount <> "" And Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
                Sheet1.Range("Sheet1.AggrigateAmount").value = AggrigateAmount
    End If
    Dim init, Node As Object
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
     'Malli_SIT-112685  06/05/2026
    'Set init = jsonObject("FilingStatus")
    Set init = jsonObject("filingStatus")
     
    If init.exists("clauseiv7provisio139i") Then
       
      
      Dim clauseiv7provisio139i_pfl As Variant
      clauseiv7provisio139i_pfl = init("clauseiv7provisio139i")
      
      
         If UCase(Trim(clauseiv7provisio139i_pfl)) = UCase(Trim("Y")) Then
            'If init("clauseiv7provisio139i") = "Y" Then
               'Malli----------06/11/2024
                If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg") = "Yes"
                End If
                '----------
                Sheet1.Range("clauseiv7provisio139iFlg_1") = "No"
                Sheet1.Range("clauseiv7provisio139iFlg_2") = "No"
                'Malli-----------06/11/2024
                 If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_3") = "No"
                End If
                 If Sheet1.Range("clauseiv7provisio139iFlg_4").Locked = False Then
                Sheet1.Range("clauseiv7provisio139iFlg_4") = "No"
                End If
                '----------------------
            For Each Node In init("clauseiv7provisio139iDtls")
            
             Dim clauseiv7provisio139iNature_pfl
                 clauseiv7provisio139iNature_pfl = Node("clauseiv7provisio139iNature")
            
                'If Node("clauseiv7provisio139iNature") = "1" Then
                If clauseiv7provisio139iNature_pfl = "1" Then
                'Malli-----------06/11/2024
                    If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
                    End If
                    If Sheet1.Range("clauseiv7provisio139iAmount_3").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iAmount_3") = Node("clauseiv7provisio139iAmount")
                    End If
                    '------------
                'ElseIf Node("clauseiv7provisio139iNature") = "2" Then
                ElseIf clauseiv7provisio139iNature_pfl = "2" Then
                     If Sheet1.Range("clauseiv7provisio139iFlg_4").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
                    End If
                    If Sheet1.Range("clauseiv7provisio139iAmount_4").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iAmount_4") = Node("clauseiv7provisio139iAmount")
                    End If
                End If
                
            Next Node
        Else
        'Malli-----------06/11/2024
            If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
            Sheet1.Range("clauseiv7provisio139iFlg") = "No"
            End If
        End If
        
        
        
    Else
        
            If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
            Sheet1.Range("clauseiv7provisio139iFlg") = "No"
            End If
        
        
    End If
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
'    'AY_2024_25 'Malli
'
'    Dim LEI_Nbr, Valid_uptodate
'    'filingStatus.LEIDtls.LEINumber
'    'filingStatus.LEIDtls.ValidUptoDate
'
'    LEI_Nbr = jsonObject("filingStatus")("LEIDtls")("LEINumber")
'    Valid_uptodate = jsonObject("filingStatus")("LEIDtls")("ValidUptoDate")
'
'
'    If LEI_Nbr <> "" And Sheet1.Range("LEI_Number").Locked = False Then
'                Sheet1.Range("LEI_Number").value = LEI_Nbr
'    End If
'    '"2024-09-24
'    If Valid_uptodate <> "" And Sheet1.Range("LEI_Date").Locked = False Then
'                'Sheet1.Range("LEI_Date").value = Valid_uptodate
'                Sheet1.Range("LEI_Date").value = Mid(Valid_uptodate, 9, 2) & "/" & Mid(Valid_uptodate, 6, 2) & "/" & Mid(Valid_uptodate, 1, 4)
'    End If
'
'AY_2024_25 'Malli
    Dim LEI_Nbr, Valid_uptodate
    'filingStatus.LEIDtls.LEINumber
    'filingStatus.LEIDtls.ValidUptoDate
    'personalInfo
    'LEI_Nbr = jsonObject("filingStatus")("LEIDtls")("LEINumber")
    'Valid_uptodate = jsonObject("filingStatus")("LEIDtls")("ValidUptoDate")
    LEI_Nbr = jsonObject("personalInfo")("filingStatus")("LEIDtls")("LEINumber")
    Valid_uptodate = jsonObject("personalInfo")("filingStatus")("LEIDtls")("ValidUptoDate")
    If LEI_Nbr <> "" And Sheet1.Range("LEI_Number").Locked = False Then
                Sheet1.Range("LEI_Number").value = LEI_Nbr
    End If
    '"2024-09-24
    If Valid_uptodate <> "" And Sheet1.Range("LEI_Date").Locked = False Then
                'Sheet1.Range("LEI_Date").value = Valid_uptodate
                Sheet1.Range("LEI_Date").value = Mid(Valid_uptodate, 9, 2) & "/" & Mid(Valid_uptodate, 6, 2) & "/" & Mid(Valid_uptodate, 1, 4)
    End If


End Function


'Function ImportFilingStatus_pfl(jsonText As String)
''On Error Resume Next
'
'Dim jsonObject As Object
'Dim jsonDictionary As Object
'Dim ReturnFileSec, AsseseeRepFlg, Node, nodeList, NewTaxRegime, Form10IEDate, Form10IEAckNo, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, RepCapacity, RepAddress, RepPAN, RepAadhaar
'
'Set jsonObject = ParseJson.ParseJson(jsonText)
'
'    Dim Filingtype As Variant
'    Dim sReturnFile, iReturnFile As Variant
'    Dim sPort5A, iPort5A As Variant
'    Dim DateofOriginalfile As Variant
'    Dim NoticeDateussec, UniqueNo, InvstmntFundRefrdSec115UB, ResidentialStatus, Section115BA As Variant
'    Dim iProvisoFlag, sProvisoFlag As Variant
'    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
'    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
'    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
'    Dim YYYY, MM, DD, strdate As String
'
'    If (Not jsonObject.exists("filingStatus")) Then
'        Exit Function
'    End If
'
'    ReturnFileSec = jsonObject("filingStatus")("returnFileSec")
'
'    ReceiptNo = jsonObject("filingStatus")("ReceiptNo")
'
'    OrigRetFiledDate = jsonObject("filingStatus")("origRetFiledDate")
'    NoticeDateUnderSec = jsonObject("filingStatus")("noticeDateUnderSec")
'
'    UniqueNo = jsonObject("filingStatus")("uniqueNo")
'    InvstmntFundRefrdSec115UB = jsonObject("filingStatus")("invstmntFundRefrdSec115UB")
'    ResidentialStatus = jsonObject("filingStatus")("residentialStatus")
'
'
'    If ReturnFileSec <> "" Then
'        iReturnFile = ReturnFileSec
'
'        If iReturnFile = "11" Then
'           sReturnFile = "139(1)-On or before due date"
'        ElseIf iReturnFile = "12" Then
'           sReturnFile = "139(4)-After due date"
'        ElseIf iReturnFile = "13" Then
'           sReturnFile = "142(1)"
'        ElseIf iReturnFile = "14" Then
'           sReturnFile = "148"
'        ElseIf iReturnFile = "15" Then
'           sReturnFile = "153A"
'        ElseIf iReturnFile = "16" Then
'           sReturnFile = "153C"
'        ElseIf iReturnFile = "17" Then
'           sReturnFile = "139(5)-Revised Return"
'        ElseIf iReturnFile = "18" Then
'           sReturnFile = "139(9)"
'        ElseIf iReturnFile = "20" Then
'           sReturnFile = "119(2)(b)- after condonation of delay"
'        End If
'
'        Sheet1.Range("sheet1.ReturnFileSec").value = sReturnFile
'      End If
'
'    If UniqueNo <> "" And Sheet1.Range("Sheet1.UniqueNo").Locked = False Then
'        Sheet1.Range("Sheet1.UniqueNo").value = UniqueNo
'    End If
''    If InvstmntFundRefrdSec115UB <> "" Then
''        Sheet1.Range("sheet1.InvstmntFundRefrdSec115UB").value = InvstmntFundRefrdSec115UB
''    End If
''    If ResidentialStatus <> "" Then
''        Sheet1.Range("sheet1.ResidentialStatus1").value = ResidentialStatus
''    End If
'
'   If Sheet1.Range("sheet1.ReturnFileSec") <> "" Then
'        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec")
'
'        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
'
'        If ReturnFileSec = "17" Then
'              If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
'                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            DateofOriginalfile = OrigRetFiledDate
'            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
'                Sheet1.Range("sheet1.OrigRetFiledDate").value = Mid(DateofOriginalfile, 1, 2) & "/" & Mid(DateofOriginalfile, 4, 2) & "/" & Mid(DateofOriginalfile, 7, 4)
'            End If
'        Else
'            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
'                Sheet1.Range("sheet1.ReceiptNo").value = ReceiptNo
'            End If
'            DateofOriginalfile = OrigRetFiledDate
'            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
'                 YYYY = Mid(DateofOriginalfile, 1, 4)
'                MM = Mid(DateofOriginalfile, 6, 2)
'                DD = Mid(DateofOriginalfile, 9, 2)
'                strdate = DD & "/" & MM & "/" & YYYY
'                    Sheet1.Range("sheet1.OrigRetFiledDate").value = strdate
'            End If
'
'            NoticeDateussec = NoticeDateUnderSec
'           If NoticeDateussec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
'                 YYYY = Mid(DateofOriginalfile, 1, 4)
'                MM = Mid(DateofOriginalfile, 6, 2)
'                DD = Mid(DateofOriginalfile, 9, 2)
'                strdate = DD & "/" & MM & "/" & YYYY
'                Sheet1.Range("sheet1.NoticeDate").value = strdate
'            End If
'        End If
'    End If
'
'
'
'
'End Function

Function ImportVerification_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject As Object
Dim jsonDictionary As Object
Dim FatherName, AssesseeVerName, AssesseeVerPAN, Capacity, Place As Variant

Set jsonObject = ParseJson.ParseJson(jsonText)
 If Not jsonObject.exists("verification") Then
        Exit Function
    End If
    Dim verificationObject As Object
    Set verificationObject = jsonObject("verification")
    If verificationObject.exists("declaration") Then
        Dim declarationObject As Object
        Set declarationObject = verificationObject("declaration")
        FatherName = declarationObject("fatherName")
        AssesseeVerName = declarationObject("assesseeVerName")
        AssesseeVerPAN = declarationObject("assesseeVerPAN")
    End If
    Capacity = verificationObject("capacity")
    
   If Capacity = "Self" Then
      Capacity = "S"
    ElseIf Capacity = "Representative" Then
      Capacity = "R"
    ElseIf Capacity = "Karta" Then
      Capacity = "K"
    ElseIf Capacity = "Authorised Signatory" Then
      Capacity = "A"
    End If
    If Capacity = "S" Then
        Capacity = "Self"
    ElseIf Capacity = "R" Then
        Capacity = "Representative"
    ElseIf Capacity = "K" Then
        Capacity = "Karta"
    ElseIf Capacity = "A" Then
        Capacity = "Authorised Signatory"
    Else
        Capacity = "(Select)"
    End If
        
    If FatherName <> "" And Sheet52.Range("sheet9.FatherName").Locked = False Then
        Sheet52.Range("sheet9.FatherName").value = FatherName
    End If

    If AssesseeVerName <> "" And Sheet52.Range("sheet9.AssesseeVerName").Locked = False Then
        Sheet52.Range("sheet9.AssesseeVerName").value = AssesseeVerName
    End If
    If AssesseeVerPAN <> "" And Sheet52.Range("sheet9.PAN").Locked = False Then
        Sheet52.Range("sheet9.PAN").value = AssesseeVerPAN
    End If
        
    If Capacity <> "" And Sheet52.Range("sheet9.Capacity").Locked = False Then
        Sheet52.Range("Sheet9.Capacity").value = Capacity
    End If
                                                
     Dim Repcapacity
     Repcapacity = jsonObject("assesseeRep")("repCapacity")
        If Repcapacity <> "" Then
              If Repcapacity = "L" Then
                    Repcapacity = "Legal Heir"
               ElseIf Repcapacity = "M" Then
                    Repcapacity = "Manager"
               ElseIf Repcapacity = "G" Then
                    Repcapacity = "Guardian"
               ElseIf Repcapacity = "O" Then
                    Repcapacity = "Other"
               End If
     Else
        Repcapacity = "(Select)"
     End If
     
     If Repcapacity <> "" And Sheet1.Range("sheet1.RepCapacity").Locked = False Then
            Sheet1.Range("sheet1.RepCapacity").value = Repcapacity
     End If
End Function

Function ImportRefund_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Node, Nodelist As Object
Dim IFSCColNo, BankNameColNo, ACCNOColNo, CheckBox, AccountTyp As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim NodeMain, NodeListMain
'Change-8(a), Test-037, Remark - Setting a flag to identify availability of at least 1 Indian bank account
Dim IndianAccntflag
IndianAccntflag = 1

Set jsonObject = ParseJson.ParseJson(jsonText)
'   If Not jsonObject.exists("bankAccountDtls") Then
'        Exit Function
'    End If
'
'    Set NodeListMain = jsonObject("bankAccountDtls")
'
'    For Each NodeMain In NodeListMain
'
'        Set nodeList = NodeMain("addtnlBankDetails")
'

Dim nodeTemp




Set Nodelist = New Collection

If jsonObject.exists("bankAccountDtls") Then
Set NodeListMain = jsonObject("bankAccountDtls")
Dim flag
If NodeListMain <> Empty Then
For Each NodeMain In NodeListMain
    For Each Node In NodeMain("addtnlBankDetails")
        flag = 0
        For Each nodeTemp In Nodelist
            If nodeTemp("bankAccountNo") = Node("bankAccountNo") Then
                flag = 1
            End If
        Next nodeTemp
        If flag = 0 Then
            Nodelist.add Node
        End If
        
    Next Node
Next NodeMain
End If
End If

'Comented by Mallikarjun  as per AY_2025_26 the secoundary Prefill source lastFiledITR was removed
'AY_2025-26
'Set NodeListMain = jsonObject("lastFiledITR")("bankAccountDtls")
'For Each NodeMain In NodeListMain
'    For Each Node In NodeMain("addtnlBankDetails")
'        flag = 0
'        For Each nodeTemp In Nodelist
'            If nodeTemp("bankAccountNo") = Node("bankAccountNo") Then
'                flag = 1
'            End If
'        Next nodeTemp
'        If flag = 0 Then
'            Nodelist.add Node
'        End If
'
'    Next Node
'Next NodeMain
'--------------------------------------------------------------------
'Change-8(b), Test-037, Remark - Setting a flag to identify availability of at least 1 Indian bank account
If Nodelist.count < 1 Then
    IndianAccntflag = 0
End If
'Change-8(c), Test-037, Remark - setting the field as no if IndianAccntflag is 0
'And unlocking the foreign bank account details field
If IndianAccntflag = 0 Then
    Sheet7.Range("Sheet9.IBANSELECT") = "No"
End If


 'Malli_AY_2025_26----------------
Set NodeListMain = jsonObject("bankAccountDtls")
    For Each NodeMain In NodeListMain
        Set Nodelist = NodeMain("addtnlBankDetails")
'--------------------------


        IFSCColNo = Sheet7.Range("Sheet9.OthMICRCode").Column
        BankNameColNo = Sheet7.Range("Sheet9.OthBankName").Column
        ACCNOColNo = Sheet7.Range("Sheet9.OthBankAccountNumber").Column
    
       ' AY_2024_25 Change 'Malli
        CheckBox = Sheet7.Range("tempxml").Column  'Malli_AY_2025_26
         
        AccountTyp = Sheet7.Range("Sheet9.TypeOfAccount").Column  'Malli
        
        TotalExRow = Sheet7.Range("Sheet9.OthMICRCode").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        Sheet7.Activate 'By Bindu SIT-94209
        
        If (TotalXMLRow > 0) Then
            Sheet7.Range("Sheet9.OthMICRCode").ClearContents
            Sheet7.Range("Sheet9.OthBankName").ClearContents
            Sheet7.Range("Sheet9.OthBankAccountNumber").ClearContents
            Sheet7.Range("Sheet9.TypeOfAccount").ClearContents           'Malli
            'Sheet7.Range("Sheet9.CheckBox").ClearContents
            Sheet7.Range("tempxml").ClearContents  'Malli_AY_2025_26
        End If
        
        
        Sheet7.Activate 'By Bindu SIT-94209
        'Malli comented--AY_2025_26
        If (TotalDiffRow > 0) Then
            'AddDiffRows_BANK (TotalDiffRow)
            Dim idi As Long
            idi = 0
            For idi = 1 To TotalDiffRow
                AddDiffRows_BANK (1)
            Next
        End If
        
        
        
        '-----------------------------------
        rowcount = getRowNo(Sheet7.Range("Sheet9.OthMICRCode").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In Nodelist
            rowcount = rowcount + 1
                
                Sheet7.Cells(rowcount, IFSCColNo).value = Node("ifsccode")
                Sheet7.Cells(rowcount, BankNameColNo).value = Node("bankName")
                Sheet7.Cells(rowcount, ACCNOColNo).value = Node("bankAccountNo")
           
           '
            'Malli_AY_2025_26 Uncomented--------------
               ''AY-2024_25 Change
                'AY_2025_26
                If Node("useForRefund") = "true" Then
                    Sheet7.Cells(rowcount, CheckBox).value = "true"
                ElseIf Node("useForRefund") = "false" Then
                    Sheet7.Cells(rowcount, CheckBox).value = "false"
                ElseIf Node("useForRefund") = "" Then
                     Sheet7.Cells(rowcount, CheckBox).value = "false"
                End If
                
                LinkCheckBoxes
           '-----------------------------------------------------------------

                 Dim AcType As Variant
                 AcType = Node("AccountType")
                 
                 
                 If AcType = "SB" Then
                    Sheet7.Cells(rowcount, AccountTyp).value = "Savings Account"

                    ElseIf AcType = "CA" Then
                    Sheet7.Cells(rowcount, AccountTyp).value = "Current Account"

                    ElseIf AcType = "CC" Then
                    Sheet7.Cells(rowcount, AccountTyp).value = "Cash Credit Account"

                    ElseIf AcType = "OD" Then
                    Sheet7.Cells(rowcount, AccountTyp).value = "Over Draft Account"

                    ElseIf AcType = "NRO" Then
                    Sheet7.Cells(rowcount, AccountTyp).value = "Non Resident Account"

                    ElseIf AcType = "OTH" Then
                    Sheet7.Cells(rowcount, AccountTyp).value = "Other"
                End If
                
                
                
                                    
             
            '-----------------------------------------
            
            cnt = cnt + 1
        Next Node
            RecTDS1 = cnt
       Next NodeMain  'Malli_AY_2025_26
End Function
Sub AddDiffRows_BANK(DiffRows As Long)
    setTableInfo_BANK
    Sheet7.Activate
    searchLastRow ("Sheet9.OthMICRCode")
    insertRowUnderSectionWithFormula (DiffRows)
    LinkCheckBoxes
End Sub

Sub setTableInfo_BANK()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet7.Range("Sheet9.OthMICRCode").count
    Set rangecells = Sheet7.Range("Sheet9.OthMICRCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "") Or (rangecells.item(mIntCtr).value = "(Select)")) Then
               ccount = ccount + 1
           End If
    Next
    'DefinedgridNameRange = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.CheckBox||tempxml"
    'AY_2024_26 OLD
'    DefinedgridNameRange = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount"
    'Malli_AY_2025_26 CHANGE
    DefinedgridNameRange = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount||Sheet9.CheckBox||tempxml"
    '------------------------------
End Sub

Function ScheduleTCSImport_pfl(jsonText As String)
On Error Resume Next
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, Node As Object
    Dim strdate As String
    Dim DepDate, Depdt As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow, rowcount, cnt As Long
    Dim jsonObject As Object
    Dim TANColNo, EmployerOrDeductorOrCollecterNameColNo, DeductedYrColNo, BroughtFwdTCSAmtColNo, TotalTCSColNo, AmtTCSClaimedThisYearColNo, AmtCarriedFwdColNo As Variant
'Change-11(a), Test-022, Remark - replace
    Set jsonObject = ParseJson.ParseJson(jsonText)
    Set Nodelist = New Collection
    
'Change-11(b), Test-022, Remark - adding form26as and lastFiledITR in the nodeList
'Dim flag
'    flag = 1
    If jsonObject("form26as").exists("scheduleTCS") Then
        If jsonObject("form26as")("scheduleTCS").exists("tcs") Then
            Set Nodelist = jsonObject("form26as")("scheduleTCS")("tcs")
'        Else
'            flag = 0
        End If
'        Else
'        flag = 0
    End If
    
    
    If jsonObject("lastFiledITR").exists("scheduleTCS") Then
        If jsonObject("lastFiledITR")("scheduleTCS").exists("tcs") Then
            For Each Node In jsonObject("lastFiledITR")("scheduleTCS")("tcs")
                If Node.exists("deductedYr") Or Node.exists("broughtFwdTCSAmt") Then
                Dim nodeTemp
                Set nodeTemp = CreateObject("Scripting.Dictionary")
                nodeTemp("deductedYr") = Node("deductedYr")
                nodeTemp("broughtFwdTCSAmt") = Node("broughtFwdTCSAmt")
                Nodelist.add nodeTemp
                End If
            Next Node
        End If
    End If
'-----------------------------------------------------------------------------

    'Set jsonObject = ParseJson.ParseJson(jsonText)
    'Set nodeList = jsonObject("form26as")("scheduleTCS")("tcs")
    TANColNo = Sheet28.Range("TCS.TAN").Column
    EmployerOrDeductorOrCollecterNameColNo = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Column
'    DeductedYrColNo = Sheet28.Range("TCS.DeductedYear").Column
'    BroughtFwdTCSAmtColNo = Sheet28.Range("TCS.BroughtFwdTDSAmt").Column
    TotalTCSColNo = Sheet28.Range("TCS.TotalTCS").Column
    AmtTCSClaimedThisYearColNo = Sheet28.Range("TCS.AmtTCSClaimedThisYear").Column
    'AmtCarriedFwdColNo = Sheet28.Range("TCS.AmtCarriedFwd").Column

    TotalExRow = Sheet28.Range("TCS.TAN").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet28.Range("TCS.TAN").ClearContents
        Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
'        Sheet28.Range("TCS.DeductedYear").ClearContents
'        Sheet28.Range("TCS.BroughtFwdTDSAmt").ClearContents
        Sheet28.Range("TCS.TotalTCS").ClearContents
        Sheet28.Range("TCS.AmtTCSClaimedThisYear").ClearContents
    End If

    If (TotalDiffRow > 0) Then
        AddRows_TCS_import_Prfl (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet28.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 1
    
    Dim rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6 As Range
    
    Set rangecells1 = Sheet28.Range("TCS.TAN").Cells
    Set rangecells2 = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
'    Set rangecells3 = Sheet28.Range("TCS.DeductedYear").Cells
'    Set rangecells4 = Sheet28.Range("TCS.BroughtFwdTDSAmt").Cells
    Set rangecells5 = Sheet28.Range("TCS.TotalTCS").Cells
    Set rangecells6 = Sheet28.Range("TCS.AmtTCSClaimedThisYear").Cells

    For Each Node In Nodelist
        rowcount = rowcount + 1
        If Node.exists("employerOrDeductorOrCollectDetl") Then
            If Node("employerOrDeductorOrCollectDetl")("tan") <> "" Then
                rangecells1.item(cnt).value = Node("employerOrDeductorOrCollectDetl")("tan")
            End If
            If Node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName") <> "" Then
                rangecells2.item(cnt).value = Node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
        End If
'        If Node.exists("lastFiledITR") Then
'            Dim lastFiledITR, ScheduleTCS, tcs
'            Set lastFiledITR = Node("lastFiledITR")
'            If lastFiledITR.exists("scheduleTCS") Then
'                Set ScheduleTCS = lastFiledITR("scheduleTCS")
'                If ScheduleTCS.exists("tcs") Then
'                    Set tcs = ScheduleTCS("tcs")
'                    If tcs("deductcedYr") <> "" Then
'                        rangecells3.item(cnt).value = tcs("deductcedYr")
'                    End If
'                     If tcs("broughtFwdTCSAmt") <> "" Then
'                        rangecells4.item(cnt).value = tcs("broughtFwdTCSAmt")
'                    End If
'                End If
'            End If
'        End If
'Change-11(c), Test-022, Remark -
'        If node("deductedYr") <> "" Then
'            rangecells3.item(cnt).value = node("deductedYr")
'        End If
'         If node("broughtFwdTCSAmt") <> "" Then
'            rangecells4.item(cnt).value = node("broughtFwdTCSAmt")
'        End If
        
        If Node("amtTCSClaimedThisYear") <> "" Then
            rangecells5.item(cnt).value = Node("amtTCSClaimedThisYear")
            rangecells6.item(cnt).value = Node("amtTCSClaimedThisYear")
        End If
     
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
    
 End Function
 '==========================================================================================
 Function ScheduleTCSImport_pfl1(jsonText As String)
On Error Resume Next
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, Node As Object
    Dim strdate As String
    Dim DepDate, Depdt As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow, rowcount, cnt As Long
    Dim jsonObject As Object
    Dim TANColNo, EmployerOrDeductorOrCollecterNameColNo, DeductedYrColNo, BroughtFwdTCSAmtColNo, TotalTCSColNo, AmtTCSClaimedThisYearColNo, AmtCarriedFwdColNo As Variant
    Set jsonObject = ParseJson.ParseJson(jsonText)
    Set Nodelist = New Collection
    Dim TCS1CreditColNo, TCS1PanOPColNo, TCS1_CollectYearColNo, TCS1_ColctOwnHandsColNo, TCS16iiCollected37iColNo, TCS1_ClmdOwnHandsColNo, TCS1_ClmdOPTCSColNo, TCS1_ClmdOPPANColNo As Variant
    If jsonObject("form26as").exists("scheduleTCS") Then
        If jsonObject("form26as")("scheduleTCS").exists("tcs") Then
            Set Nodelist = jsonObject("form26as")("scheduleTCS")("tcs")
        End If
    End If
    
    
    If jsonObject("lastFiledITR").exists("scheduleTCS") Then
        If jsonObject("lastFiledITR")("scheduleTCS").exists("tcs") Then
            For Each Node In jsonObject("lastFiledITR")("scheduleTCS")("tcs")
                If Node.exists("deductedYr") Or Node.exists("broughtFwdTCSAmt") Then
                Dim nodeTemp
                Set nodeTemp = CreateObject("Scripting.Dictionary")
                nodeTemp("deductedYr") = Node("deductedYr")
                nodeTemp("broughtFwdTCSAmt") = Node("broughtFwdTCSAmt")
                Nodelist.add nodeTemp
                End If
            Next Node
        End If
    End If
'-----------------------------------------------------------------------------
    TCS1CreditColNo = Sheet28.Range("TCS1.CreditName").Column
    TANColNo = Sheet28.Range("TCS1.TAN").Column
    TCS1PanOPColNo = Sheet28.Range("TCS1.PANOP").Column
    TCS1_CollectYearColNo = Sheet28.Range("TCS1.CollectYear").Column
    BroughtFwdTCSAmtColNo = Sheet28.Range("TCS1.UnclaimedBFAmount").Column
    TCS1_ColctOwnHandsColNo = Sheet28.Range("TCS1.CollectedOwnHands").Column
    TCS16iiCollected37iColNo = Sheet28.Range("TCS1.Collected37i").Column
    TCS1_ClmdOwnHandsColNo = Sheet28.Range("TCS1.ClaimedOwnHands").Column
    TCS1_ClmdOPTCSColNo = Sheet28.Range("TCS1.ClaimedOPTCS").Column
    TCS1_ClmdOPPANColNo = Sheet28.Range("TCS1.ClaimedOPPAN").Column

    TotalExRow = Sheet28.Range("TCS1.TAN").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet28.Range("TCS1.CreditName").ClearContents
        Sheet28.Range("TCS1.TAN").ClearContents
        Sheet28.Range("TCS1.PANOP").ClearContents
        Sheet28.Range("TCS1.CollectYear").ClearContents
        Sheet28.Range("TCS1.UnclaimedBFAmount").ClearContents
        Sheet28.Range("TCS1.CollectedOwnHands").ClearContents
        Sheet28.Range("TCS1.Collected37i").ClearContents
        Sheet28.Range("TCS1.ClaimedOwnHands").ClearContents
        Sheet28.Range("TCS1.ClaimedOPTCS").ClearContents
        Sheet28.Range("TCS1.ClaimedOPPAN").ClearContents
    End If

    If (TotalDiffRow > 0) Then
        AddRows_TCS1_import (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet28.Range("TCS1.CreditName").name)
    rowcount = rowcount - 1
    cnt = 1
    
    Dim rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8, rangecells9, rangecells10 As Range
    
       Set rangecells1 = Sheet28.Range("TCS1.CreditName").Cells
       Set rangecells2 = Sheet28.Range("TCS1.TAN").Cells
       Set rangecells3 = Sheet28.Range("TCS1.PANOP").Cells
       Set rangecells4 = Sheet28.Range("TCS1.CollectYear").Cells
       Set rangecells5 = Sheet28.Range("TCS1.UnclaimedBFAmount").Cells
       Set rangecells6 = Sheet28.Range("TCS1.CollectedOwnHands").Cells
       Set rangecells7 = Sheet28.Range("TCS1.Collected37i").Cells
       Set rangecells8 = Sheet28.Range("TCS1.ClaimedOwnHands").Cells
       Set rangecells9 = Sheet28.Range("TCS1.ClaimedOPTCS").Cells
       Set rangecells10 = Sheet28.Range("TCS1.ClaimedOPPAN").Cells
    
    For Each Node In Nodelist
        rowcount = rowcount + 1
        If Node.exists("TCSCreditName") Then
            If Node("TCSCreditName") <> "" Then
            
 'Konda updated on 05-03-2025--SIT-89193
''                Dim TCSCreditName1 As Object
           'Uncomment by Shrutika(12/03/2025)SIT 89286
                Dim TCSCreditNamestr As String
''                TCSCreditName1 = Node("TCSCreditName")
'                If UCase(TCSCreditName1) = "1" Then
'                    TCSCreditNamestr = "Self"
'                'End If
'                ElseIf UCase(TCSCreditName1) = "2" Then
'                    TCSCreditNamestr = "Other Person"
'                End If
'                rangecells1.item(cnt).value = TCSCreditNamestr
'            End If
    'Malli_SIT-89193  enums are changed 1 to s and 2 to O
'            If node("TCSCreditName") = "1" Then
'                TCSCreditNamestr = "Self"
'            rangecells1.item(cnt).value = TCSCreditNamestr
'            ElseIf node("TCSCreditName") = "2" Then
'                TCSCreditNamestr = "Other Person"
'            rangecells1.item(cnt).value = TCSCreditNamestr
'            End If
            Dim TCSCreditNamestr_pfl
            TCSCreditNamestr_pfl = UCase(Node("TCSCreditName"))
            If TCSCreditNamestr_pfl <> "" Then
                    If TCSCreditNamestr_pfl = UCase("S") Then
                        TCSCreditNamestr = "Self"
                    rangecells1.item(cnt).value = TCSCreditNamestr
                    ElseIf TCSCreditNamestr_pfl = UCase("O") Then
                        TCSCreditNamestr = "Other Person"
                    rangecells1.item(cnt).value = TCSCreditNamestr
                    End If
            End If

        End If
'End updated on 05-03-2025----SIT-89193

            If Node("employerOrDeductorOrCollectDetl")("tan") <> "" Then
                rangecells2.item(cnt).value = Node("employerOrDeductorOrCollectDetl")("tan")
            End If
        End If
        
        If Node("amtTCSClaimedThisYear") <> "" Then
            rangecells6.item(cnt).value = Node("amtTCSClaimedThisYear")
            rangecells8.item(cnt).value = Node("amtTCSClaimedThisYear")
        End If
     
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
    
 End Function
 '==========================================================================================
 Sub AddRows_TCS_import_Prfl(DiffRows As Long)
    Sheet28.Activate
    setTblinfo_TCS_Prfl
    SelectLastRow ("TCS.TAN")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_TCS_Prfl)
End Sub
Sub setTblinfo_TCS_Prfl()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet28.Range("TCS.TAN").count
    Set rangecells = Sheet28.Range("TCS.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TCS = ccount
'    rngname_TCS_Prfl = "TCS.TAN;TCS.EmployerOrDeductorOrCollecterName;TCS.DeductedYear;TCS.TotalTCS;TCS.AmtTCSClaimedThisYear;TCS.BroughtFwdTDSAmt;TCS.AmtCarriedFwd;"
    rngname_TCS_Prfl = "TCS.TAN;TCS.EmployerOrDeductorOrCollecterName;TCS.TotalTCS;TCS.AmtTCSClaimedThisYear;TCS.AmtCarriedFwd;"

End Sub

Function ImportScheduleTDS2_pfl(jsonText As String)
    On Error Resume Next

    Dim jsonObject As Object
    Dim jsonDictionary, init As Object
    Dim TotalXMLRow  As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, Node, Nodelist2 As Object
    Dim strdate As String
    Dim DepDate, Depdt As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow, rowcount, cnt As Long

    Dim TDSCreditName_TDS2, HeadIncome_TDS2
    Dim TDSCreditNameColNo, PANofOtherPersonColNo, AadhaarOfOtherPersonColNo, TANOfDeductorColNo, DeductedYrColNo, BroughtFwdTDSAmtColNo
    Dim TaxDeductedOwnHandsColNo, TaxDeductedIncomeColNo, TaxDeductedTDSColNo, TaxClaimedOwnHandsColNo, TaxClaimedIncomeColNo, TaxClaimedTDSColNo, SectionTDS_2_ColNo
    Dim TaxClaimedSpouseOthPrsnPANColNo, SpouseOthPrsnAadhaarColNo, GrossAmountColNo, HeadOfIncomeColNo, AmtCarriedFwdColNo

    Set jsonObject = ParseJson.ParseJson(jsonText)
    
    'Malli uncommented------------SIT_94033_TDS2issue1
    If Not jsonObject.exists("form26as") Then
        Exit Function
    End If

    If Not jsonObject("form26as").exists("tdsOnOthThanSals") Then
        Exit Function
    End If
 '------------------------------------------------------

    
    
    Dim ScheduleTDS2Object As Object
    Set ScheduleTDS2Object = jsonObject("form26as")("tdsOnOthThanSals")
'    If Not ScheduleTDS2Object.exists("tdSonOthThanSal") Then
'         Exit Function
'    End If
    Set Nodelist = ScheduleTDS2Object("tdSonOthThanSal")
    
    'Malli----------
'    If ScheduleTDS2Object.exists("tdSonOthThanSal") Then
'    Set Nodelist2 = jsonObject("lastFiledITR")("tdsOnOthThanSals")("tdSonOthThanSal")
'    For Each Node In Nodelist2
'        Nodelist.add Node
'    Next Node
'    Else
'        Set Nodelist = jsonObject("lastFiledITR")("tdsOnOthThanSals")("tdSonOthThanSal")
'    End If
    '-----------------------
    
    'lastFiledITR.scheduleTDS3Dtls.tds3Details.deductedYr
    TANOfDeductorColNo = Sheet28.Range("TDS2.TAN").Column
    DeductedYrColNo = Sheet28.Range("TDS2.DeductedYear").Column
    BroughtFwdTDSAmtColNo = Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column
    TaxDeductedOwnHandsColNo = Sheet28.Range("TDS2.DeductedInOwnHands").Column
    'TaxDeductedTDSColNo = Sheet28.Range("TDS2.DeductedTDS5A").Column
    TaxClaimedOwnHandsColNo = Sheet28.Range("TDS2.ClaimedInOwnHands").Column
   ' TaxClaimedTDSColNo = Sheet28.Range("TDS2.ClaimedTDS5A").Column
    GrossAmountColNo = Sheet28.Range("TDS2.GrossAmount").Column
    HeadOfIncomeColNo = Sheet28.Range("TDS2.HeadOfIncome").Column
    TDSCreditName_TDS2 = Sheet28.Range("TDS2.TDSCreditName").Column
     
    'Malli---------------AY_2025_26_V0.6--------------------------------
    SectionTDS_2_ColNo = Sheet28.Range("TDS2.TdsDeducted").Column
    '-------------------------------------------------------------------

    TotalExRow = Sheet28.Range("TDS2.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
        
        
    If (TotalDiffRow > 0) Then
    Sheet28.Activate   'Malli_SIT_93890
        AddDiffRows_TDS2_Prfl (TotalDiffRow)
    End If
        
    rowcount = getRowNo(Sheet28.Range("TDS2.DeductedYear").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
     
        rowcount = rowcount + 1
        If Node("TDSCreditName") <> "" Then
            If Node("TDSCreditName") = "S" Then
                Sheet28.Cells(rowcount, TDSCreditName_TDS2).value = "Self"
            ElseIf Node("TDSCreditName") = "O" Then
                Sheet28.Cells(rowcount, TDSCreditName_TDS2).value = "Other Person"
            End If
        End If
        If Node("deductedYr") <> "" Then
            Sheet28.Cells(rowcount, DeductedYrColNo).value = Node("deductedYr")
        End If
        If Node("broughtFwdTDSAmt") <> "" Then
            Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).value = Node("broughtFwdTDSAmt")
        End If
        If Node.exists("employerOrDeductorOrCollectDetl") Then
            If Node("employerOrDeductorOrCollectDetl")("tan") <> "" Then
                Sheet28.Cells(rowcount, TANOfDeductorColNo).value = Node("employerOrDeductorOrCollectDetl")("tan")
            End If
        End If
         Sheet28.Cells(rowcount, DeductedYrColNo) = "(Select)"
        
        
        If Node.exists("taxDeductCreditDtls") Then
            Dim TaxDeductCreditDtls
            Set TaxDeductCreditDtls = Node("taxDeductCreditDtls")
            If TaxDeductCreditDtls.exists("taxDeductedOwnHands") <> "" Then
                Sheet28.Cells(rowcount, TaxDeductedOwnHandsColNo).value = TaxDeductCreditDtls("taxDeductedOwnHands")
            End If
            If TaxDeductCreditDtls.exists("taxClaimedOwnHands") <> "" Then
                Sheet28.Cells(rowcount, TaxClaimedOwnHandsColNo).value = TaxDeductCreditDtls("taxClaimedOwnHands")
            End If
            
        End If
        If Node("grossAmount") <> "" Then
            Sheet28.Cells(rowcount, GrossAmountColNo).value = Node("grossAmount")
        End If
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
          
 
 '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 Dim SectionTDS, TDSSection_Gen
 SectionTDS = Node("sectionCode")
 If SectionTDS <> "" Then

            If UCase(SectionTDS) = UCase("92A") Then
                    TDSSection_Gen = "192-Salary-Payment to Government employees other than Indian Government employees"

                    ElseIf UCase(SectionTDS) = UCase("92B") Then
                    TDSSection_Gen = "192-Salary-Payment to employees other than Government employees"

                    ElseIf UCase(SectionTDS) = UCase("92C") Then
                    TDSSection_Gen = "192-Salary-Payment to Indian Government employees"

                    'Malli----------------------
                    'ElseIf TDSSection_Gen_E = "192A" Then
                    ElseIf UCase(SectionTDS) = UCase("2AA") Then
                    TDSSection_Gen = "192A-TDS on PF withdrawal"
                    '------------------------------------

                    ElseIf UCase(SectionTDS) = UCase("193") Then
                    TDSSection_Gen = "193-Interest on Securities"

                    ElseIf UCase(SectionTDS) = UCase("194") Then
                    TDSSection_Gen = "194-Dividends"

                    ElseIf UCase(SectionTDS) = UCase("94A") Then
                    TDSSection_Gen = "194A-Interest other than 'Interest on securities'"

                    ElseIf UCase(SectionTDS) = UCase("94B") Then
                    TDSSection_Gen = "194B-Winning from lottery or crossword puzzle"

                    'Malli---------
                    'ElseIf TDSSection_Gen_E = "94BA" Then
                      ElseIf UCase(SectionTDS) = UCase("9BA") Then
                    TDSSection_Gen = "194BA-Winnings from online games"
                    '------------------------

                    ElseIf UCase(SectionTDS) = UCase("4BB") Then
                    TDSSection_Gen = "194BB-Winning from horse race"

                    ElseIf UCase(SectionTDS) = UCase("94C") Then
                    TDSSection_Gen = "194C-Payments to contractors and sub-contractors"

                    ElseIf UCase(SectionTDS) = UCase("94D") Then
                    TDSSection_Gen = "194D-Insurance commission"

                    ElseIf UCase(SectionTDS) = UCase("4DA") Then
                    TDSSection_Gen = "194DA-Payment in respect of life insurance policy"

                    ElseIf UCase(SectionTDS) = UCase("94E") Then
                    TDSSection_Gen = "194E-Payments to non-resident sportsmen or sports associations"

                    ElseIf UCase(SectionTDS) = UCase("4EE") Then
                    TDSSection_Gen = "194EE-Payments in respect of deposits under National Savings"
'Malli-----------------------------------------------------------------
                    'ElseIf TDSSection_Gen_E = "4F" Then
                     ElseIf UCase(SectionTDS) = UCase("94F") Then
                    TDSSection_Gen = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"

                    'ElseIf TDSSection_Gen_E = "4G" Then
                     ElseIf UCase(SectionTDS) = UCase("94G") Then
                    TDSSection_Gen = "194G-Commission, price, etc. on sale of lottery tickets"

                    'ElseIf TDSSection_Gen_E = "4H" Then
                      ElseIf UCase(SectionTDS) = UCase("94H") Then
                    TDSSection_Gen = "194H-Commission or brokerage"

                    'ElseIf TDSSection_Gen_E = "4-IA" Then
                     ElseIf UCase(SectionTDS) = UCase("4IA") Then
                    TDSSection_Gen = "194I(a)-Rent on hiring of plant and machinery"

                    'ElseIf TDSSection_Gen_E = "4-IB" Then
                     ElseIf UCase(SectionTDS) = UCase("4IB") Then
                    TDSSection_Gen = "194I(b)-Rent on other than plant and machinery"

                    'ElseIf TDSSection_Gen_E = "4IA" Then
                    ElseIf UCase(SectionTDS) = UCase("9IA") Then
                    TDSSection_Gen = "194IA-TDS on Sale of immovable property"

                   ' ElseIf TDSSection_Gen_E = "4IB" Then
                    ElseIf UCase(SectionTDS) = UCase("9IB") Then
                    TDSSection_Gen = "194IB-Payment of rent by certain individuals or Hindu undivided"

                    ElseIf UCase(SectionTDS) = UCase("4IC") Then
                    TDSSection_Gen = "194IC-Payment under specified agreement"

                    'ElseIf TDSSection_Gen_E = "94J-A" Then
                    ElseIf UCase(SectionTDS) = UCase("4JA") Then
                    TDSSection_Gen = "194J(a)-Fees for technical services"

                    'ElseIf TDSSection_Gen_E = "94J-B" Then
                    ElseIf UCase(SectionTDS) = UCase("4JB") Then
                    TDSSection_Gen = "194J(b)-Fees for professional  services or royalty etc"

                    ElseIf UCase(SectionTDS) = UCase("94K") Then
                    TDSSection_Gen = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"

                    ElseIf UCase(SectionTDS) = UCase("4LA") Then
                    TDSSection_Gen = "194LA-Payment of compensation on acquisition of certain immovable"

                    ElseIf UCase(SectionTDS) = UCase("4LB") Then
                    TDSSection_Gen = "194LB-Income by way of Interest from Infrastructure Debt fund"

                    'ElseIf TDSSection_Gen_E = "4LC1" Then
                    ElseIf UCase(SectionTDS) = UCase("LC1") Then
                    TDSSection_Gen = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"

                    'ElseIf TDSSection_Gen_E = "4LC2" Then
                    ElseIf UCase(SectionTDS) = UCase("LC2") Then
                    TDSSection_Gen = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"

                    'ElseIf TDSSection_Gen_E = "4LC3" Then
                     ElseIf UCase(SectionTDS) = UCase("Lc3") Then
                    TDSSection_Gen = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
'Malli------------------------------------------------------------------
'                    ElseIf TDSSection_Gen_E = "4BA1" Then
'                    TDSSection_Gen = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'
'                    ElseIf TDSSection_Gen_E = "4BA2" Then
'                    TDSSection_Gen = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'
'                    ElseIf TDSSection_Gen_E = "LBA1" Then
'                    TDSSection_Gen = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
'
'                    ElseIf TDSSection_Gen_E = "LBA2" Then
'                    TDSSection_Gen = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"


'Malli_SIT-117189  13/04/2026

                    ElseIf UCase(SectionTDS) = UCase("BA1") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then
                    TDSSection_Gen = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"

                    ElseIf UCase(SectionTDS) = UCase("BA2") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then
                    TDSSection_Gen = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"

                    ElseIf UCase(SectionTDS) = UCase("BA1") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
                    TDSSection_Gen = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"

                    ElseIf UCase(SectionTDS) = UCase("BA2") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
                    TDSSection_Gen = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"



'------------------------------------------------------------------------
                    'ElseIf TDSSection_Gen_E = "LBA3" Then
                    ElseIf UCase(SectionTDS) = UCase("BA3") Then
                    TDSSection_Gen = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"

                    ElseIf UCase(SectionTDS) = UCase("LBB") Then
                    TDSSection_Gen = "194LBB-Income in respect of units of investment fund"

                    ElseIf UCase(SectionTDS) = UCase("94R") Then
                    TDSSection_Gen = "194R-Benefits or perquisites of business or profession"

                    ElseIf UCase(SectionTDS) = UCase("94S") Then
                    TDSSection_Gen = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"

                    'ElseIf TDSSection_Gen_E = "94B-P" Then
                     ElseIf UCase(SectionTDS) = UCase("4BP") Then
                    TDSSection_Gen = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"

                    'ElseIf TDSSection_Gen_E = "94R-P" Then
                    ElseIf UCase(SectionTDS) = UCase("4RP") Then
                    TDSSection_Gen = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"

                    'ElseIf TDSSection_Gen_E = "94S-P" Then
                    ElseIf UCase(SectionTDS) = UCase("4SP") Then
                    TDSSection_Gen = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"

                    ElseIf UCase(SectionTDS) = UCase("LBC") Then
                    TDSSection_Gen = "194LBC-Income in respect of investment in securitization trust"

                    ElseIf UCase(SectionTDS) = UCase("4LD") Then
                    TDSSection_Gen = "194LD-TDS on interest on bonds / government securities"

                    ElseIf UCase(SectionTDS) = UCase("94M") Then
                    TDSSection_Gen = "194M-Payment of certain sums by certain individuals or HUF"

                    ElseIf UCase(SectionTDS) = UCase("94N") Then
                    TDSSection_Gen = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"

                    'ElseIf TDSSection_Gen_E = "94N-F" Then
                    ElseIf UCase(SectionTDS) = UCase("4NF") Then
                    TDSSection_Gen = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"

                    'ElseIf TDSSection_Gen_E = "94N-C" Then
                    ElseIf UCase(SectionTDS) = UCase("4NC") Then
                    TDSSection_Gen = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"

                    'ElseIf TDSSection_Gen_E = UCase("94N-FT" Then
                    ElseIf UCase(SectionTDS) = UCase("NFT") Then
                    TDSSection_Gen = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"

                    ElseIf UCase(SectionTDS) = UCase("94O") Then
                    TDSSection_Gen = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."

                    ElseIf UCase(SectionTDS) = UCase("94P") Then
                    TDSSection_Gen = "194P-Deduction of tax in case of specified senior citizen"

                    ElseIf UCase(SectionTDS) = UCase("94Q") Then
                    TDSSection_Gen = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"

                    ElseIf UCase(SectionTDS) = UCase("195") Then
                    TDSSection_Gen = "195-Other sums payable to a non-resident"

                    ElseIf UCase(SectionTDS) = UCase("96A") Then
                    TDSSection_Gen = "196A-Income in respect of units of non-residents"

                    ElseIf UCase(SectionTDS) = UCase("96B") Then
                    TDSSection_Gen = "196B-Payments in respect of units to an offshore fund"

                    ElseIf UCase(SectionTDS) = UCase("96C") Then
                    TDSSection_Gen = "196C-Income from foreign currency bonds or shares of Indian"

                    ElseIf UCase(SectionTDS) = UCase("96D") Then
                    TDSSection_Gen = "196D-Income of foreign institutional investors from securities"

                    'ElseIf TDSSection_Gen_E = "96DA" Then
                     ElseIf UCase(SectionTDS) = UCase("6DA") Then
                    TDSSection_Gen = "196D(1A)-Income of specified fund from securities"

                    'ElseIf TDSSection_Gen_E = "94BA-P" Then
                     ElseIf UCase(SectionTDS) = UCase("BAP") Then
                    TDSSection_Gen = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
                    Else
                        TDSSection_Gen = "(Select)"
                End If

            
            If TDSSection_Gen <> "" Then
            Sheet28.Cells(rowcount, SectionTDS_2_ColNo).value = TDSSection_Gen
                    SectionTDS = ""
            End If
        End If

        
        
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
   'Malli commented---------AY_2025_26_V0.6--------------------------------
        'Newly added By Bindu as per Pre-fill SRS
        
'        If node("sectionCode") <> "" Then
'
'           Dim sectioncode
'           sectioncode = UCase(node("sectionCode"))
'
'           If sectioncode = "94N" Then
'             HeadIncome_TDS2 = "Not applicable ( only in case TDS is deducted u/s 194N)"
'
'           ElseIf sectioncode = "94I(b)" Or sectioncode = "194IB" Then
'             HeadIncome_TDS2 = "Income from House Property"
'
'           ElseIf sectioncode = "94" Or sectioncode = "94A" Or _
'           sectioncode = "93" Or sectioncode = "94I" Or _
'           sectioncode = "94D" Or sectioncode = "94DA" Or _
'           sectioncode = "94G" Or sectioncode = "94H" Or _
'           sectioncode = "94B" Or sectioncode = "94LB" Or _
'           sectioncode = "94LC" Or sectioncode = "94LD" Or _
'           sectioncode = "94LBL" Or sectioncode = "94E" Or _
'           sectioncode = "94BB" Or sectioncode = "94BBJ" Then
'
'               HeadIncome_TDS2 = "Income from Other Sources"
'
'            ElseIf sectioncode = "94LA" Or sectioncode = "94IA" Then
'              HeadIncome_TDS2 = "Income from Capital Gains"
'            Else
'                HeadIncome_TDS2 = "(Select)"
'
'           End If
'         Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = HeadIncome_TDS2
'
'        End If
        
   '----------------------------------------------------------------------------------------------
   
        'Malli--------Uncommented---------AY_2025_26_V0.6--------------------------------
        If Node("headOfIncome") <> "" Then
            HeadIncome_TDS2 = UCase(Node("headOfIncome"))

             If HeadIncome_TDS2 = "HP" Then
               HeadIncome_TDS2 = "Income from House Property"
             ElseIf HeadIncome_TDS2 = "BP" Then
                HeadIncome_TDS2 = "Income from Business & Profession"
             ElseIf HeadIncome_TDS2 = "CG" Then
                HeadIncome_TDS2 = "Income from Capital Gains"
             ElseIf HeadIncome_TDS2 = "OS" Then
                HeadIncome_TDS2 = "Income from Other Sources"
             ElseIf HeadIncome_TDS2 = "EI" Then
                HeadIncome_TDS2 = "Exempt Income"
             ElseIf HeadIncome_TDS2 = "NA" Then
                HeadIncome_TDS2 = "Not applicable ( only in case TDS is deducted u/s 194N)"
             Else
                HeadIncome_TDS2 = "(Select)"
            End If

            Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = HeadIncome_TDS2
        End If
'--------------------------------------------------------------------------------------------
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
End Function

Sub AddDiffRows_TDS2_Prfl(DiffRows As Long)
    setDiffTblinfo_TDS2oth_prl
    Sheet28.Activate
    SelectLastRow ("TDS2.TDSCreditName")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_TDS2_Prfl)
End Sub

Sub setDiffTblinfo_TDS2oth_prl()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet28.Range("TDS2.TAN").count
    Set rangecells = Sheet28.Range("TDS2.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS2oth = ccount
    'Malli--------
    'AY_2023_24 old
    'rngname_TDS2_Prfl = "TDS2.TDSCreditName;TDS2.PANOP;TDS2.Aadhaarop;TDS2.TAN;TDS2.DeductedYear;TDS2.BroughtFwdTDSAmt;TDS2.DeductedInOwnHands;TDS2.DeductedIncome5A;TDS2.DeductedTDS5A;TDS2.ClaimedInOwnHands;TDS2.ClaimedIncome5A;TDS2.ClaimedTDS5A;TDS2.ClaimedPAN5A;TDS2.ClaimedAadhaar5A;TDS2.AmtCarriedForward;TDS2.GrossAmount;TDS2.HeadOfIncome;"
    'AY_2024_25 change
    'rngname_TDS2_Prfl = "TDS2.TDSCreditName;TDS2.PanOtherPerson;TDSAadharOfOther_TDS2;TDS2.TAN;TDS2.DeductedYear;TDS2.BroughtFwdTDSAmt;TDS2.DeductedInOwnHands;TCS.ClaimedOwnHands;TDS2.DeductedTDS5A;TDS2.ClaimedInOwnHands;TDS2.ClaimedIncome5A;TDS2.ClaimedTDS5A;TDS2.ClaimedPAN5A;ClaimedAadhar5A_TDS2;TDS2.GrossAmount;TDS2.HeadOfIncome;TDS2.AmtCarriedForward;"
'AY_2025_26 change
     rngname_TDS2_Prfl = "TDS2.TDSCreditName;TDS2.PanOtherPerson;TDSAadharOfOther_TDS2;TDS2.TAN;TDS2.TdsDeducted;TDS2.DeductedYear;TDS2.BroughtFwdTDSAmt;TDS2.DeductedInOwnHands;TCS.ClaimedOwnHands;TDS2.DeductedTDS5A;TDS2.ClaimedInOwnHands;TDS2.ClaimedIncome5A;TDS2.ClaimedTDS5A;TDS2.ClaimedPAN5A;ClaimedAadhar5A_TDS2;TDS2.GrossAmount;TDS2.HeadOfIncome;TDS2.AmtCarriedForward;"
   '--------------------
    End Sub

Function ImportTDSonSalary_pfl(jsonText As String)
On Error Resume Next
    
    Dim jsonObject, jsonDictionary As Object
    Dim Node, Nodelist As Object
    Dim rowcount, cnt As Variant
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim TotalExRow As Long

    Set jsonObject = ParseJson.ParseJson(jsonText)
    
    Set Nodelist = jsonObject("form26as")("tdsOnSalaries")("tdsOnSalary")
    
    
    TANNoEmployer = Sheet28.Range("TDS1.TAN").Column
    TDSNameOfEmployer = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet28.Range("TDS1.IncChrgSal").Column
    TDSTotalTax = Sheet28.Range("TDS1.TotalTDSSal").Column
    
    TotalExRow = Sheet28.Range("TDS1.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet28.Range("TDS1.TAN").ClearContents
        Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet28.Range("TDS1.IncChrgSal").ClearContents
        Sheet28.Range("TDS1.TotalTDSSal").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_TDS1_Prfl (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet28.Range("TDS1.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each Node In Nodelist
            rowcount = rowcount + 1
        
            Sheet28.Cells(rowcount, TANNoEmployer).value = Node("employerOrDeductorOrCollectDetl")("tan")
            Sheet28.Cells(rowcount, TDSNameOfEmployer).value = Node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            Sheet28.Cells(rowcount, TDSIncomeCharge).value = Node("incChrgSal")
            Sheet28.Cells(rowcount, TDSTotalTax).value = Node("totalTDSSal")
            
        
        cnt = cnt + 1
    Next Node
    RecTDS1 = cnt
    
End Function

Sub AddDiffRows_TDS1_Prfl(DiffRows As Long)
    setDiffTblinfo_TDS1_Prfl
    Sheet28.Activate
    searchLastRow ("TDS1.TAN")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TaxP)
End Sub

Sub setDiffTblinfo_TDS1_Prfl()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet28.Range("TDS1.TAN").count
    Set rangecells = Sheet28.Range("TDS1.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    EfilingCommon.DefinedgridNameRange = "TDS1.TAN||TDS1.EmployerOrDeductorOrCollecterName||TDS1.IncChrgSal||TDS1.TotalTDSSal"
    'rngname_TaxP = "TDSal.TAN;TDSal.EmployerOrDeductorOrCollecterName;TDSal.IncChrgSal;TDSal.TotalTDSSal;"
End Sub

Function ImportSchedule80(jsonText As String)
On Error Resume Next
Dim init, jsonObject, Node As Object
Dim Nodelist, nodeList1, Nodelist2, Nodelist3, Nodelist4, Nodelist5, nodeList6, nodeList7, nodeList8 As Object
Dim nodeList9, nodeList10, nodeList11, nodeList12, nodeList13, nodeList14, nodeList15 As Object

Set Nodelist = New Collection
Set nodeList1 = New Collection
Set Nodelist2 = New Collection
Set Nodelist3 = New Collection
Set Nodelist4 = New Collection
Set Nodelist5 = New Collection
Set nodeList6 = New Collection
Set nodeList7 = New Collection
Set nodeList8 = New Collection
Set nodeList9 = New Collection
Set nodeList10 = New Collection
Set nodeList11 = New Collection
Set nodeList12 = New Collection
Set nodeList13 = New Collection
Set nodeList14 = New Collection
Set nodeList15 = New Collection

Dim nodeTemp

Dim TotalXMLRow, rowcount, cnt, RecTCS  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim DeductAmountSec80ColNo
Dim DeductSikkimColNo, DeductHimachalColNo, DeductUttaranchalColNo, DeductNagalandColNo, DeductTripuraColNo
Dim DeductAssamColNo, DeductArunachalPradeshColNo, DeductManipurColNo, DeductMizoramColNo, DeductMeghalayaColNo
Dim DeductAmountSec80iColNo, DeductAmountSec80iiColNo, DeductAmountSec80iiiColNo, DeductAmountSec80ivColNo
Dim DeductAmountSec804ColNo, DeductAmountSec807AColNo, DeductAmountSec807BColNo, DeductAmountSec809ColNo, DeductAmountSec8010ColNo
Dim DeductAmountSec80Food11AColNo, DeductAmountSec8011BColNo, DeductAmountSec8011CColNo, DeductAmountSec80Fruit11AColNo


Set jsonObject = ParseJson.ParseJson(jsonText)

If jsonObject.exists("form10CCB") Then
    If jsonObject("form10CCB").exists("form10CCBAmountDetls") Then
        Set Nodelist = jsonObject("form10CCB")("form10CCBAmountDetls")
        
        For Each Node In Nodelist
            If Node.exists("secCd") And Node.exists("subSecCd") Then
                If Node("secCd") = "80IA" And Node("subSecCd") = "form10ccb4iv" Then
                    Set nodeTemp = CreateObject("Scripting.Dictionary")
                    nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                    nodeList1.add nodeTemp
                End If
                If Node("secCd") = "80IB" Then
                    If Node("subSecCd") = "form10ccb9" Then
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        Nodelist2.add nodeTemp
                    End If
                    If Node("subSecCd") = "form10ccb10" Then
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        Nodelist3.add nodeTemp
                    End If
                    If Node("subSecCd") = "form10ccb11Aii" Then
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        Nodelist4.add nodeTemp
                    End If
                    If Node("subSecCd") = "form10ccb11Ai" Then
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        Nodelist5.add nodeTemp
                    End If
                End If
                If Node("secCd") = "80IE" And Node("subSecCd") = "form10ccb80IE1" Then
                    If Node("state") = "28" Then    'Sikkim
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList6.add nodeTemp
                    End If
                    If Node("state") = "04" Then    'Assam
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList7.add nodeTemp
                    End If
                    If Node("state") = "03" Then    'Arunachal Pradesh
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList8.add nodeTemp
                    End If
                    If Node("state") = "20" Then    'Manipur
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList9.add nodeTemp
                    End If
                    If Node("state") = "22" Then    'Mizoram
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList10.add nodeTemp
                    End If
                    If Node("state") = "21" Then    'Meghalaya
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList11.add nodeTemp
                    End If
                    If Node("state") = "23" Then    'Nagaland
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList12.add nodeTemp
                    End If
                    If Node("state") = "30" Then    'Tripura
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList13.add nodeTemp
                    End If
                End If
                If Node("secCd") = "80Ic" And (Node("subSecCd") = "form10ccb2aii" Or Node("subSecCd") = "formccb2bii") Then
                    If Node("state") = "13" Then    'Himachal Pradesh
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList14.add nodeTemp
                    End If
                    If Node("state") = "34" Then    'Uttarakhand denoted as Uttaranchal
                        Set nodeTemp = CreateObject("Scripting.Dictionary")
                        nodeTemp("deductAmountSec80") = Node("deductAmountSec80")
                        nodeList15.add nodeTemp
                    End If
                End If
            End If
        Next Node
        
        '---------
If (nodeList1.count > 0) Then
    
    DeductAmountSec80iiiColNo = Sheet37.Range("IA80.DeductProfUs80_IA_4_iii").Column
    
    TotalExRow = Sheet37.Range("IA80.DeductProfUs80_IA_4_iii").Rows.count
    
    TotalXMLRow = nodeList1.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
            If Sheet37.Range("IA80.DeductProfUs80_IA_4_iii").Locked = False Then
                Sheet37.Range("IA80.DeductProfUs80_IA_4_iii").ClearContents
            End If
    End If

    If (TotalDiffRow > 0) Then
        AddDiffRows_80iii (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet37.Range("IA80.DeductProfUs80_IA_4_iii").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each Node In nodeList1
            rowcount = rowcount + 1
            If Sheet37.Cells(rowcount, DeductAmountSec80iiiColNo).Locked = False Then
                Sheet37.Cells(rowcount, DeductAmountSec80iiiColNo).value = Node("deductAmountSec80")
            End If
        cnt = cnt + 1
    Next Node
    RecTCS = cnt
End If
        '---------
If Nodelist2.count > 0 Then

        DeductAmountSec809ColNo = Sheet37.Range("IB80.DeductMinOilUs80_IB_9").Column
        
        TotalExRow = Sheet37.Range("IB80.DeductMinOilUs80_IB_9").Rows.count
        
        TotalXMLRow = Nodelist2.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IB80.DeductMinOilUs80_IB_9").Locked = False Then
                Sheet37.Range("IB80.DeductMinOilUs80_IB_9").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_809 (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IB80.DeductMinOilUs80_IB_9").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In Nodelist2
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductAmountSec809ColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductAmountSec809ColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
    RecTCS = cnt
End If
        '---------
If Nodelist3.count > 0 Then
   
        DeductAmountSec8010ColNo = Sheet37.Range("IB80.DeductHousUs80_IB_10").Column
        
        TotalExRow = Sheet37.Range("IB80.DeductHousUs80_IB_10").Rows.count
        
        TotalXMLRow = Nodelist3.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IB80.DeductHousUs80_IB_10").Locked = False Then
                Sheet37.Range("IB80.DeductHousUs80_IB_10").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_8010 (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IB80.DeductHousUs80_IB_10").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In Nodelist3
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductAmountSec8010ColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductAmountSec8010ColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If Nodelist4.count > 0 Then
   
        DeductAmountSec80Fruit11AColNo = Sheet37.Range("IB80.DeductFruitVegUs80_IB_11A").Column
        
        TotalExRow = Sheet37.Range("IB80.DeductFruitVegUs80_IB_11A").Rows.count
        
        TotalXMLRow = Nodelist4.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IB80.DeductFruitVegUs80_IB_11A").Locked = False Then
                Sheet37.Range("IB80.DeductFruitVegUs80_IB_11A").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_80Fruit11A (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IB80.DeductFruitVegUs80_IB_11A").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In Nodelist4
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductAmountSec80Fruit11AColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductAmountSec80Fruit11AColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If Nodelist5.count > 0 Then
 
        DeductAmountSec80Food11AColNo = Sheet37.Range("IB80.DeductFoodGrainUs80_IB_11A").Column
        
        TotalExRow = Sheet37.Range("IB80.DeductFoodGrainUs80_IB_11A").Rows.count
        
        TotalXMLRow = Nodelist5.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IB80.DeductFoodGrainUs80_IB_11A").Locked = False Then
                Sheet37.Range("IB80.DeductFoodGrainUs80_IB_11A").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_80Food11A (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IB80.DeductFoodGrainUs80_IB_11A").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In Nodelist5
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductAmountSec80Food11AColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductAmountSec80Food11AColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If nodeList6.count > 0 Then

        DeductSikkimColNo = Sheet37.Range("IC80.DeductInSikkim").Column
        
        TotalExRow = Sheet37.Range("IC80.DeductInSikkim").Rows.count
        
        TotalXMLRow = nodeList6.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.DeductInSikkim").Locked = False Then
                Sheet37.Range("IC80.DeductInSikkim").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
        AddDiffRows_80Sikkim (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.DeductInSikkim").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList6
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductSikkimColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductSikkimColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If nodeList14.count > 0 Then
        DeductHimachalColNo = Sheet37.Range("IC80.DeductInHimachalP").Column
        
        TotalExRow = Sheet37.Range("IC80.DeductInHimachalP").Rows.count
        
        TotalXMLRow = nodeList14.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.DeductInHimachalP").Locked = False Then
                Sheet37.Range("IC80.DeductInHimachalP").ClearContents
            End If
               
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Himachal (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.DeductInHimachalP").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList14
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductHimachalColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductHimachalColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
    RecTCS = cnt
End If
        '---------
If nodeList15.count > 0 Then
        DeductUttaranchalColNo = Sheet37.Range("IC80.DeductInUttaranchal").Column
        
        TotalExRow = Sheet37.Range("IC80.DeductInUttaranchal").Rows.count
        
        TotalXMLRow = nodeList15.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.DeductInUttaranchal").Locked = False Then
                Sheet37.Range("IC80.DeductInUttaranchal").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Uttaranchal (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.DeductInUttaranchal").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList15
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductUttaranchalColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductUttaranchalColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If

        '---------
If nodeList7.count > 0 Then
        DeductAssamColNo = Sheet37.Range("IC80.Assam").Column
        
        TotalExRow = Sheet37.Range("IC80.Assam").Rows.count
        
        TotalXMLRow = nodeList7.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.Assam").Locked = False Then
                Sheet37.Range("IC80.Assam").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Assam (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.Assam").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList7
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductAssamColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductAssamColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If nodeList8.count > 0 Then
        DeductArunachalPradeshColNo = Sheet37.Range("IC80.ArunachalPradesh").Column
        
        TotalExRow = Sheet37.Range("IC80.ArunachalPradesh").Rows.count
        
        TotalXMLRow = nodeList8.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.ArunachalPradesh").Locked = False Then
                Sheet37.Range("IC80.ArunachalPradesh").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_ArunachalPradesh (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.ArunachalPradesh").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList8
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductArunachalPradeshColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductArunachalPradeshColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
    RecTCS = cnt
End If
        '---------
If nodeList9.count > 0 Then
    
        DeductManipurColNo = Sheet37.Range("IC80.Manipur").Column
        
        TotalExRow = Sheet37.Range("IC80.Manipur").Rows.count
        
        TotalXMLRow = nodeList9.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.Manipur").Locked = False Then
                Sheet37.Range("IC80.Manipur").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Manipur (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.Manipur").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList9
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductManipurColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductManipurColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
    RecTCS = cnt
End If
        '---------
If nodeList10.count > 0 Then
    
        DeductMizoramColNo = Sheet37.Range("IC80.Mizoram").Column
        
        TotalExRow = Sheet37.Range("IC80.Mizoram").Rows.count
        
        TotalXMLRow = nodeList10.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.Mizoram").Locked = False Then
                Sheet37.Range("IC80.Mizoram").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Mizoram (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.Mizoram").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList10
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductMizoramColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductMizoramColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If nodeList11.count > 0 Then
    
        DeductMeghalayaColNo = Sheet37.Range("IC80.Meghalaya").Column
        
        TotalExRow = Sheet37.Range("IC80.Meghalaya").Rows.count
        
        TotalXMLRow = nodeList11.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.Meghalaya").Locked = False Then
                Sheet37.Range("IC80.Meghalaya").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Meghalaya (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.Meghalaya").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList11
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductMeghalayaColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductMeghalayaColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If nodeList12.count > 0 Then
    
        DeductNagalandColNo = Sheet37.Range("IC80.Nagaland").Column
        
        TotalExRow = Sheet37.Range("IC80.Nagaland").Rows.count
        
        TotalXMLRow = nodeList12.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.Nagaland").Locked = False Then
                Sheet37.Range("IC80.Nagaland").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Nagaland (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.Nagaland").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList12
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductNagalandColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductNagalandColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
If nodeList13.count > 0 Then
    
        DeductTripuraColNo = Sheet37.Range("IC80.Tripura").Column
        
        TotalExRow = Sheet37.Range("IC80.Tripura").Rows.count
        
        TotalXMLRow = nodeList13.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        If (TotalXMLRow > 0) Then
            If Sheet37.Range("IC80.Tripura").Locked = False Then
                Sheet37.Range("IC80.Tripura").ClearContents
            End If
        End If
        
        If (TotalDiffRow > 0) Then
         AddDiffRows_Tripura (TotalDiffRow)
        End If
        
        rowcount = getRowNo(Sheet37.Range("IC80.Tripura").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each Node In nodeList13
            rowcount = rowcount + 1
                If Sheet37.Cells(rowcount, DeductTripuraColNo).Locked = False Then
                    Sheet37.Cells(rowcount, DeductTripuraColNo).value = Node("deductAmountSec80")
                End If
        cnt = cnt + 1
        Next Node
        RecTCS = cnt
End If
        '---------
        
    End If
End If
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------

'Change-29, prefill-tracker, D85
Function ImportScheduleESOP(jsonText As String)
On Error Resume Next
Dim jsonObject
Set jsonObject = ParseJson.ParseJson(jsonText)

'UNcommented by Mallikarjun AY_2026_27 as per enhancement
'AY_2024_25 ESOP_C1 Malli
If jsonObject("ScheduleESOP").exists("ScheduleESOP2122_Type") Then
    If jsonObject("ScheduleESOP")("ScheduleESOP2122_Type").exists("TaxDeferredBFEarlierAY") Then
        If jsonObject("ScheduleESOP")("ScheduleESOP2122_Type")("AssessmentYear") = "2021-22" And Sheet46.Range("ESOP.TaxPrevAY").Locked = False Then
            Sheet46.Range("ESOP.TaxPrevAY") = jsonObject("ScheduleESOP")("ScheduleESOP2122_Type")("TaxDeferredBFEarlierAY")
        End If
    End If
End If


If jsonObject("ScheduleESOP").exists("ScheduleESOP2223_Type") Then
    If jsonObject("ScheduleESOP")("ScheduleESOP2223_Type").exists("TaxDeferredBFEarlierAY") Then
        If jsonObject("ScheduleESOP")("ScheduleESOP2223_Type")("AssessmentYear") = "2022-23" And Sheet46.Range("ESOP.TaxPrevAY1").Locked = False Then
            Sheet46.Range("ESOP.TaxPrevAY1") = jsonObject("ScheduleESOP")("ScheduleESOP2223_Type")("TaxDeferredBFEarlierAY")
        End If
    End If
End If
 
'--Malli--------------------------------------
'AY_2024_25 ESOP_C1 Malli
 If jsonObject("ScheduleESOP").exists("ScheduleESOP2324_Type") Then
    If jsonObject("ScheduleESOP")("ScheduleESOP2324_Type").exists("TaxDeferredBFEarlierAY") Then
        If jsonObject("ScheduleESOP")("ScheduleESOP2324_Type")("AssessmentYear") = "2023-24" And Sheet46.Range("ESOP.TaxPrevAY2").Locked = False Then
            Sheet46.Range("ESOP.TaxPrevAY2") = jsonObject("ScheduleESOP")("ScheduleESOP2324_Type")("TaxDeferredBFEarlierAY")
        End If
    End If
End If

'----------------------------------------------

'ScheduleESOP.ScheduleESOP2425_Type.TaxDeferredBFEarlierAY
'Malli_AY_2025_26----------------

If jsonObject("ScheduleESOP").exists("ScheduleESOP2425_Type") Then
    If jsonObject("ScheduleESOP")("ScheduleESOP2425_Type").exists("TaxDeferredBFEarlierAY") Then
        If jsonObject("ScheduleESOP")("ScheduleESOP2425_Type")("AssessmentYear") = "2024-25" And Sheet46.Range("ESOP.TaxPrevAY3").Locked = False Then
            Sheet46.Range("ESOP.TaxPrevAY3") = jsonObject("ScheduleESOP")("ScheduleESOP2425_Type")("TaxDeferredBFEarlierAY")
        End If
    End If
End If
'--------------------------------


'Malli_AY_2026_27 03/02/2026
If jsonObject("ScheduleESOP").exists("ScheduleESOP2526_Type") Then
    If jsonObject("ScheduleESOP")("ScheduleESOP2526_Type").exists("TaxDeferredBFEarlierAY") Then
        If jsonObject("ScheduleESOP")("ScheduleESOP2526_Type")("AssessmentYear") = "2025-26" And Sheet46.Range("ESOP.TaxPrevAY4").Locked = False Then
            Sheet46.Range("ESOP.TaxPrevAY4") = jsonObject("ScheduleESOP")("ScheduleESOP2526_Type")("TaxDeferredBFEarlierAY")
        End If
    End If
End If


'------------------------------------------


End Function



Function ImportPartA_139_8A(jsonText As String)

On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
Set jsonObject = jsonObject("personalInfo")
If Sheet201.Range("U_AadhaarCardNo").Locked = False Then
    Sheet201.Range("U_AadhaarCardNo").value = jsonObject("personalInfo")("aadhaarCardNo")
End If


'If Sheet201.Range("U_AadhaarEnrolmentId").Locked = False Then
'    Sheet201.Range("U_AadhaarEnrolmentId").value = jsonObject("personalInfo")("aadhaarCardNo")
'End If


If Sheet201.Range("U_PreviouslyFiledForThisAY").Locked = False Then
    If jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY") = "Y" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY").value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY") = "N" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No"
    End If
End If


If Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Locked = False Then
    If jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY_139_8A") = "1" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "139(1)"
        
    ElseIf jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY_139_8A") = "2" Then
    'Malli--
        'Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "Other"
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "Others"
    End If
End If


If Sheet201.Range("U_ITRForm").Locked = False Then
    Sheet201.Range("U_ITRForm").value = jsonObject("PartA_139_8A")("Applicable_139_8A")("ITRForm")
End If

If Sheet201.Range("U_AcknowledgementNo").Locked = False Then
    Sheet201.Range("U_AcknowledgementNo").value = jsonObject("PartA_139_8A")("Applicable_139_8A")("AcknowledgementNo")
End If

If Sheet201.Range("U_OrigRetFiledDate").Locked = False Then
'Malli_AY_2026-27  SIT-113034  13/04/2026
If jsonObject("filingStatus")("origRetFiledDate") <> "" Then
    Sheet201.Range("U_OrigRetFiledDate").value = jsonObject("filingStatus")("origRetFiledDate")
End If
End If

If Sheet201.Range("U_LaidOutIn_139_8A").Locked = False Then
    If jsonObject("PartA_139_8A")("LaidOutIn_139_8A") = "Y" Then
        Sheet201.Range("U_LaidOutIn_139_8A").value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("LaidOutIn_139_8A") = "N" Then
        Sheet201.Range("U_LaidOutIn_139_8A").value = "No"
    End If
End If


Dim Nodelist, Node, itemp, vRows

If jsonObject("PartA_139_8A").exists("UpdatingInc") Then

    Set Nodelist = jsonObject("PartA_139_8A")("UpdatingInc")("ReasonsForUpdatingIncDtls")
    If Nodelist.count > 0 Then

        itemp = 0
        
        If Nodelist.count > Sheet201.Range("U_ReasonsForUpdatingIncome").Rows.count Then
            Sheet201.Activate
            EfilingCommon.DefinedgridNameRange = ("U_ReasonsForUpdatingIncome")
            ActiveCellRange = EfilingCommon.searchLastRow("U_ReasonsForUpdatingIncome")
            vRows = EfilingCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet201.Range("U_ReasonsForUpdatingIncome").Rows.count)
        End If
        
        For Each Node In Nodelist
        itemp = itemp + 1
        Select Case Node("ReasonsForUpdatingIncome")
        
        Case "1"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Return previously not filed"
        Case "2"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Income not reported correctly"
        Case "3"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Wrong heads of income chosen"
        Case "4"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Reduction of carried forward loss"
        Case "5"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Reduction of unabsorbed depreciation"
        Case "6"
              Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Reduction of tax credit u/s 115JB/115JC"
        Case "7"
              Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Wrong rate of tax"
        Case "OTH"
               Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).value = "Others "
        End Select
        
        
        Next Node
    
    End If
    
End If

If Sheet201.Range("U_UpdatedReturnDuringPeriod").Locked = False Then
    If jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "1" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Up to 12 months from the end of Relevant Assessment Year"
        
    ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "2" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 12 to 24 Months from the end of Relevant Assessment  Year"
    End If
End If


If Sheet201.Range("U_UnabsorbedDepreciation").Locked = False Then
    If jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UnabsorbedDepreciation") = "Y" Then
        Sheet201.Range("U_UnabsorbedDepreciation").value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UnabsorbedDepreciation") = "N" Then
        Sheet201.Range("U_UnabsorbedDepreciation").value = "No"
    End If
End If

Set Nodelist = New Collection


If jsonObject("PartA_139_8A").exists("RetrntoRedCarriedFL") Then
If jsonObject("PartA_139_8A")("RetrntoRedCarriedFL").exists("UDYear") Then
    Set Nodelist = jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UDYear")("UnabsorbedDepreciationYearDtls")
    If Nodelist.count > 0 Then
        itemp = 0
        
        If Nodelist.count > Sheet201.Range("U_UnabsorbedDepreciationYear").Rows.count Then
            Sheet201.Activate
            EfilingCommon.DefinedgridNameRange = ("U_UnabsorbedDepreciationYear")
            ActiveCellRange = EfilingCommon.searchLastRow("U_UnabsorbedDepreciationYear")
            vRows = EfilingCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet201.Range("U_UnabsorbedDepreciationYear").Rows.count)
        End If
        
        For Each Node In Nodelist
        itemp = itemp + 1
            
            Dim YearUD As Variant
            YearUD = Node("UnabsorbedDepreciationYear")
            If YearUD <> "" Then
                'Malli-------30/10/2024
''                If YearUD = "2023" Then
''                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2023-24"
''                ElseIf YearUD = "2024" Then
''                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2024-25"
''                End If
                
                If YearUD = "2025" Then
                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2025-26"
                ElseIf YearUD = "2026" Then
                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).value = "2026-27"
                End If
                
                
                '------------------
            End If
            
            If Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If Node("ReturnFiledEffectFlg") = "Y" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).value = "Yes"

                ElseIf Node("ReturnFiledEffectFlg") = "N" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).value = "No"
                End If
             End If
             
              If Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If Node("ReturnType") = "1" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Orginal Return 139(1)/139(4)"
                ElseIf Node("ReturnType") = "2" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Revised return"
                ElseIf Node("ReturnType") = "3" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).value = "Updated return"
                End If
             End If
        
        Next Node
    
    End If
End If
End If




End Function

Function ImportSchedule80GD(jsonText As String)
On Error Resume Next
Dim Node, Nodelist, jsonObject, init As Object
Dim DoneeWithPanName, DoneePAN, DoneeARN, AddrDetail, CityOrTownOrDistrict, StateCode, PinCode, DonationInCash, DonationAmt, TotalDonationAmt, EligibleAmt As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long
Set jsonObject = ParseJson.ParseJson(jsonText)

Set init = jsonObject("Schedule80G")

If Not init.exists("Don50PercentApprReqd") Then
    Exit Function
End If

If Not init("Don50PercentApprReqd").exists("DoneeWithPan") Then
    Exit Function
End If
Set Nodelist = jsonObject("Schedule80G")("Don50PercentApprReqd")("DoneeWithPan")

    DoneeWithPanName = Sheet20.Range("Per5080G.DoneeWithPanName").Column
    DoneePAN = Sheet20.Range("Per5080G.DoneePAN").Column
    DoneeARN = Sheet20.Range("Per5080G.DonationRefNumber").Column
    AddrDetail = Sheet20.Range("Per5080G.AddrDetail").Column
    CityOrTownOrDistrict = Sheet20.Range("Per5080G.CityOrTownOrDistrict").Column
    StateCode = Sheet20.Range("Per5080G.StateCode").Column
    PinCode = Sheet20.Range("Per5080G.PinCode").Column
    DonationInCash = Sheet20.Range("Per5080G.DonationInCash").Column
    DonationAmt = Sheet20.Range("Per5080G.DonationAmt").Column

    TotalExRow = Sheet20.Range("Per5080G.DoneeWithPanName").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet20.Range("Per5080G.DoneeWithPanName").Locked = False Then
            Sheet20.Range("Per5080G.DoneeWithPanName").ClearContents
        End If
        If Sheet20.Range("Per5080G.DoneePAN").Locked = False Then
            Sheet20.Range("Per5080G.DoneePAN").ClearContents
        End If
        If Sheet20.Range("Per5080G.DonationRefNumber").Locked = False Then
            Sheet20.Range("Per5080G.DonationRefNumber").ClearContents
        End If
        If Sheet20.Range("Per5080G.AddrDetail").Locked = False Then
            Sheet20.Range("Per5080G.AddrDetail").ClearContents
        End If
        If Sheet20.Range("Per5080G.CityOrTownOrDistrict").Locked = False Then
            Sheet20.Range("Per5080G.CityOrTownOrDistrict").ClearContents
        End If
        If Sheet20.Range("Per5080G.StateCode").Locked = False Then
            Sheet20.Range("Per5080G.StateCode").ClearContents
        End If
        If Sheet20.Range("Per5080G.PinCode").Locked = False Then
            Sheet20.Range("Per5080G.PinCode").ClearContents
        End If
        If Sheet20.Range("Per5080G.DonationInCash").Locked = False Then
            Sheet20.Range("Per5080G.DonationInCash").ClearContents
        End If
        If Sheet20.Range("Per5080G.DonationAmt").Locked = False Then
            Sheet20.Range("Per5080G.DonationAmt").ClearContents
        End If
       
    End If
    
    If (TotalDiffRow > 0) Then
        AddRows_Don5080G (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet20.Range("Per5080G.DoneeWithPanName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            If Sheet20.Cells(rowcount, DoneeWithPanName).Locked = False Then
                Sheet20.Cells(rowcount, DoneeWithPanName).value = Node("DoneeWithPanName")
            End If
            If Sheet20.Cells(rowcount, DoneePAN).Locked = False Then
                Sheet20.Cells(rowcount, DoneePAN).value = Node("DoneePAN")
            End If
            If Sheet20.Cells(rowcount, DoneeARN).Locked = False Then
                Sheet20.Cells(rowcount, DoneeARN).value = Node("DoneeARN")
            End If
            If Sheet20.Cells(rowcount, AddrDetail).Locked = False Then
                Sheet20.Cells(rowcount, AddrDetail).value = Node("AddressDetail")("AddrDetail")
            End If
            If Sheet20.Cells(rowcount, CityOrTownOrDistrict).Locked = False Then
                Sheet20.Cells(rowcount, CityOrTownOrDistrict).value = Node("AddressDetail")("CityOrTownOrDistrict")
            End If
            Dim iState
            iState = Node("AddressDetail")("StateCode")
            If Len(iState) = "1" Then
                iState = "0" & iState
            End If

            If iState = "99" Then
                iState = ""
            End If

            If Sheet20.Cells(rowcount, StateCode).Locked = False Then
                Sheet20.Cells(rowcount, StateCode).value = Findtext(iState, "State")
            End If
            If Sheet20.Cells(rowcount, PinCode).Locked = False Then
                Sheet20.Cells(rowcount, PinCode).value = Node("AddressDetail")("PinCode")
            End If
 'Konda updated on 21-08-2025 SIT-89250
'            If Sheet20.Cells(rowcount, DonationInCash).Locked = False Then
'                Sheet20.Cells(rowcount, DonationInCash).value = node("DonationAmt")
'            End If
            
            If Sheet20.Cells(rowcount, DonationInCash).Locked = False Then
                Sheet20.Cells(rowcount, DonationInCash).value = Node("DonationAmtCash")
            End If
'--------------------on 21-08-2025--SIT-89250

            If Sheet20.Cells(rowcount, DonationAmt).Locked = False Then
                Sheet20.Cells(rowcount, DonationAmt).value = Node("DonationAmtOtherMode")
            End If
            
        cnt = cnt + 1
    Next Node
    
    RecTDS1 = cnt
    
End Function
Function ImportScheduleS_23_24(jsonText As String)
On Error Resume Next
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, nodeList1A, nodeList1B, jsonObject As Object
    Dim Node, Node1 As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim init As Object
    Dim rowcount, rowcount1, rowcount2, rowcount3, cnt As Long
    Dim TANofEmployer
    Dim NatureColNo, DescriptionColNo, AmountColNo, iNature, sNature As Variant
    Set jsonObject = ParseJson.ParseJson(jsonText)

Set init = jsonObject("form24q")("salaries")
If init.exists("salary") Then
     
     'Newly updated from Bottom to here by Bindu as per defect
    Set Nodelist = jsonObject("lastFiledITR")("natOfEmployment")
  
    rowcount = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
       
            Dim NatureOfEmployment As String
            NatureOfEmployment = UCase(Node)

            If NatureOfEmployment = "CGOV" Then NatureOfEmployment = "Central Government"
            If NatureOfEmployment = "SGOV" Then NatureOfEmployment = "State Government"
            If NatureOfEmployment = "PE" Then NatureOfEmployment = "CG-Pensioners"
            If NatureOfEmployment = "PESG" Then NatureOfEmployment = "SG-Pensioners"
            If NatureOfEmployment = "PEPS" Then NatureOfEmployment = "PSU-Pensioners"
            If NatureOfEmployment = "PEO" Then NatureOfEmployment = "Others-Pensioners"
            If NatureOfEmployment = "PSU" Then NatureOfEmployment = "Public Sector Undertaking"
            If NatureOfEmployment = "OTH" Then NatureOfEmployment = "OTHERS"

'         'Malli commented-------AY_2025_26_V0.6----Prefill osibility removed----------------------------
'            If Sheet39.Range("sheet.EmployerCategory" & rowcount).Locked = False Then
'                Sheet39.Range("sheet.EmployerCategory" & rowcount).value = NatureOfEmployment
'            End If
'            '--------------------------------------------------
    Next
    '--ends
     
     
     
    Set Nodelist = jsonObject("form24q")("salaries")("salary")
     
    TotalExRow = Sheet39.Range("SalarySectionCount").value
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    

    Dim i As Long
    For i = 1 To TotalDiffRow
      addSalariesBlock
    Next
    
    rowcount = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
            If Sheet39.Range("SAL.NameOfEmployer" & rowcount).Locked = False Then
                Sheet39.Range("SAL.NameOfEmployer" & rowcount).value = UCase(Node("nameOfEmployer"))
            End If
            
'            Dim NatureOfEmployment As String
'            NatureOfEmployment = UCase(Node("natOfEmployment"))
'
'            If NatureOfEmployment = "CGOV" Then NatureOfEmployment = "Central Government"
'            If NatureOfEmployment = "SGOV" Then NatureOfEmployment = "State Government"
'            If NatureOfEmployment = "PE" Then NatureOfEmployment = "CG-Pensioners"
'            If NatureOfEmployment = "PESG" Then NatureOfEmployment = "SG-Pensioners"
'            If NatureOfEmployment = "PEPS" Then NatureOfEmployment = "PSU-Pensioners"
'            If NatureOfEmployment = "PEO" Then NatureOfEmployment = "Others-Pensioners"
'            If NatureOfEmployment = "PSU" Then NatureOfEmployment = "Public Sector Undertaking"
'            If NatureOfEmployment = "OTH" Then NatureOfEmployment = "OTHERS"
            
            If Sheet39.Range("SAL.PANofEmployer" & rowcount).Locked = False Then
                Sheet39.Range("SAL.PANofEmployer" & rowcount).value = Node("tanOfEmployer")
            End If
            
'            If Sheet39.Range("sheet.EmployerCategory" & rowcount).Locked = False Then
'                Sheet39.Range("sheet.EmployerCategory" & rowcount).value = NatureOfEmployment
'            End If
            
            If Sheet39.Range("SAL.AddrDetail" & rowcount).Locked = False Then
                Sheet39.Range("SAL.AddrDetail" & rowcount).value = UCase(Node("addressDetail")("addDetail"))
            End If
            
            'Newly changed by Bindu Tag name as pre-fill SRS
            iState = UCase(Node("addressDetail")("stateCode"))

            If Len(iState) = 1 Then
            iState = "0" & iState
            End If
            
            sState = Findtext(iState, "State")
            If Sheet39.Range("SAL.StateCode" & rowcount).Locked = False Then
                Sheet39.Range("SAL.StateCode" & rowcount).value = sState
            End If
             If Sheet39.Range("SAL.CityOrTownOrDistrict" & rowcount).Locked = False Then
                Sheet39.Range("SAL.CityOrTownOrDistrict" & rowcount).value = UCase(Node("addressDetail")("cityOrTownOrDistrict"))
            End If
            If Sheet39.Range("SAL.PinCode" & rowcount).Locked = False Then
                Sheet39.Range("SAL.PinCode" & rowcount).value = UCase(Node("addressDetail")("pinCode"))
            End If
            If Sheet39.Range("SAL.ZipCode" & rowcount).Locked = False Then
                Sheet39.Range("SAL.ZipCode" & rowcount).value = UCase(Node("addressDetail")("zipCode"))
            End If
            
            
  
            
            
            
            '"NatureOfSalary" start
            Set init = Node("salarys")("natureOfSalary")
            If init.exists("othersIncDtls") Then

            Set nodeList1A = Node("salarys")("natureOfSalary")("othersIncDtls")
            Dim section, salaryDescription, SalaryAmount As Variant
            
            section = Sheet39.Range("Salary.Section" & rowcount).Column
            salaryDescription = Sheet39.Range("Salary.salaryDescription" & rowcount).Column
            SalaryAmount = Sheet39.Range("Salary.SalaryAmount" & rowcount).Column
            
            
            TotalExRow = Range("Salary.Section" & rowcount).Rows.count
            
            TotalXMLRow = nodeList1A.count
            TotalDiffRow = TotalXMLRow - TotalExRow
            
            If (TotalXMLRow > 0) Then
                Sheet39.Range("Salary.Section" & rowcount).ClearContents
                Sheet39.Range("Salary.salaryDescription" & rowcount).ClearContents
                Sheet39.Range("Salary.SalaryAmount" & rowcount).ClearContents
                
            End If
            
            If (TotalDiffRow > 0) Then
                AddSalary17_1 (TotalDiffRow)
            End If
            
            rowcount1 = getRowNo(Sheet39.Range("Salary.Section" & rowcount).name)
            rowcount1 = rowcount1 - 1
            'If IsEmpty(nodeList1A) Then
            If nodeList1A.count <= 0 Then
                Else
            For Each Node1 In nodeList1A
                rowcount1 = rowcount1 + 1
                
                If Sheet39.Cells(rowcount1, section).Locked = False Then
                    Dim valsalarydropdown1
                    If Node1("natureDesc") = "1" Then
                        valsalarydropdown1 = "Basic Salary"
                    ElseIf Node1("natureDesc") = "2" Then
                        valsalarydropdown1 = "Dearness Allowance"
                    ElseIf Node1("natureDesc") = "3" Then
                        valsalarydropdown1 = "Conveyance Allowance"
                    ElseIf Node1("natureDesc") = "4" Then
                        valsalarydropdown1 = "House Rent Allowance"
                    ElseIf Node1("natureDesc") = "5" Then
                        valsalarydropdown1 = "Leave Travel Allowance"
                    ElseIf Node1("natureDesc") = "6" Then
                        valsalarydropdown1 = "Children Education Allowance"
                    ElseIf Node1("natureDesc") = "7" Then
                        valsalarydropdown1 = "Other Allowance"
                    ElseIf Node1("natureDesc") = "8" Then
                        valsalarydropdown1 = "The contribution made  by the Employer  towards  pension scheme as referred u/s 80CCD"
                    
                    'Malli_22/04/2026  V0.3 changes
                    ElseIf Node1("natureDesc") = "9" Then
                       ' valsalarydropdown1 = "Amount deemed to be income under rule 11 of Fourth Schedule"
                    'Newly changed as per DE sheet
                        'valsalarydropdown1 = "Amount deemed to be income under rule 11(4) of Part-A of Fourth Schedule"
                        valsalarydropdown1 = "Amount deemed to be income under rule 6 of Part-A of Fourth Schedule"

                    
                    
                    ElseIf Node1("natureDesc") = "10" Then
                        
                        'valsalarydropdown1 = "Amount deemed to be income under rule 6 of Fourth Schedule"
                        'Newly changed as per DE sheet
                        
'                        valsalarydropdown1 = "Amount deemed to be income under rule 6 of Part-A of Fourth Schedule"
                        valsalarydropdown1 = "Amount deemed to be income under rule 11(4) of Part-A of Fourth Schedule"
                        
                        '---------------------------------------
                    
                    ElseIf Node1("natureDesc") = "11" Then
                        valsalarydropdown1 = "Annuity or pension"
                    ElseIf Node1("natureDesc") = "12" Then
                        valsalarydropdown1 = "Commuted Pension"
                    ElseIf Node1("natureDesc") = "13" Then
                        valsalarydropdown1 = "Gratuity"
                    ElseIf Node1("natureDesc") = "14" Then
                        valsalarydropdown1 = "Fees/ commission"
                    ElseIf Node1("natureDesc") = "15" Then
                        valsalarydropdown1 = "Advance of salary"
                    ElseIf Node1("natureDesc") = "16" Then
                        valsalarydropdown1 = "Leave Encashment"
                    ElseIf Node1("natureDesc") = "OTH" Then
                        valsalarydropdown1 = "Others"
                    End If
                    Sheet39.Cells(rowcount1, section).value = valsalarydropdown1
                End If
'                If Sheet39.Cells(rowcount1, salaryDescription).Locked = False Then
'                    Sheet39.Cells(rowcount1, salaryDescription).value = Node1("othNatOfInc")
'                End If
                If Sheet39.Cells(rowcount1, SalaryAmount).Locked = False Then
                    Sheet39.Cells(rowcount1, SalaryAmount).value = Node1("othAmount")
                End If
                
            Next Node1
             End If
           
          'Malli_AY_2025_26------------
          
          '2nd prefill source
          Else:
                  Dim salary_2ndpfl  As Variant
                   
                   salary_2ndpfl = Node("salarys")("salary")
                  
                  If salary_2ndpfl <> "" Then
                  
                        rowcount1 = getRowNo(Sheet39.Range("Salary.Section" & rowcount).name)
                        
                        section = Sheet39.Range("Salary.Section" & rowcount).Column
                        salaryDescription = Sheet39.Range("Salary.salaryDescription" & rowcount).Column
                        SalaryAmount = Sheet39.Range("Salary.SalaryAmount" & rowcount).Column
                        
                        Sheet39.Range("Salary.Section" & rowcount).ClearContents
                        Sheet39.Range("Salary.salaryDescription" & rowcount).ClearContents
                        Sheet39.Range("Salary.SalaryAmount" & rowcount).ClearContents
                        
                        If Sheet39.Cells(rowcount1, section).Locked = False Then
                                Sheet39.Cells(rowcount1, section).value = "Basic Salary"
                        End If
                        If Sheet39.Cells(rowcount1, SalaryAmount).Locked = False Then
                                Sheet39.Cells(rowcount1, SalaryAmount).value = salary_2ndpfl
                                salary_2ndpfl = ""
                        End If

                   End If
          '-----------------------------
        End If
                      
            '"NatureOfPerquisites" start
            Set init = Node("salarys")("natureOfPerquisites")
            If init.exists("othersIncDtls") Then
            '2
            Set nodeList1B = Node("salarys")("natureOfPerquisites")("othersIncDtls")
            Dim ValueSection, ValueDescription, ValueAmount As Variant
            
            ValueSection = Sheet39.Range("Salary.ValueSection" & rowcount).Column
            ValueDescription = Sheet39.Range("Salary.ValueDescription" & rowcount).Column
            ValueAmount = Sheet39.Range("Salary.ValueAmount" & rowcount).Column
            
            
            TotalExRow = Range("Salary.ValueSection" & rowcount).Rows.count
            
            TotalXMLRow = nodeList1B.count
            TotalDiffRow = TotalXMLRow - TotalExRow
            
            If (TotalXMLRow > 0) Then
                If Sheet39.Range("Salary.ValueSection" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ValueSection" & rowcount).ClearContents
                End If
                If Sheet39.Range("Salary.ValueDescription" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ValueDescription" & rowcount).ClearContents
                End If
                If Sheet39.Range("Salary.ValueAmount" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ValueAmount" & rowcount).ClearContents
                End If
            End If
            
            If (TotalDiffRow > 0) Then
                AddSalary17_2 (TotalDiffRow)
            End If
            
            rowcount2 = getRowNo(Sheet39.Range("Salary.ValueSection" & rowcount).name)
            rowcount2 = rowcount2 - 1
            'If IsEmpty(nodeList1B) Then
            If nodeList1B.count <= 0 Then
                Else
            For Each Node1 In nodeList1B
                rowcount2 = rowcount2 + 1
                
                If Sheet39.Cells(rowcount2, ValueSection).Locked = False Then
                    Dim valsalarydropdown2
                    If Node1("natureDesc") = "1" Then
                        valsalarydropdown2 = "Accommodation"
                    ElseIf Node1("natureDesc") = "2" Then
                        valsalarydropdown2 = "Cars / Other Automotive"
                    ElseIf Node1("natureDesc") = "3" Then
                        valsalarydropdown2 = "Sweeper, gardener, watchman or personal attendant"
                    ElseIf Node1("natureDesc") = "4" Then
                        valsalarydropdown2 = "Gas, electricity, water"
                    ElseIf Node1("natureDesc") = "5" Then
                        valsalarydropdown2 = "Interest free or concessional loans"
                    ElseIf Node1("natureDesc") = "6" Then
                        valsalarydropdown2 = "Holiday expenses"
                    ElseIf Node1("natureDesc") = "7" Then
                        valsalarydropdown2 = "Free or concessional travel"
                    ElseIf Node1("natureDesc") = "8" Then
                        valsalarydropdown2 = "Free meals"
                    ElseIf Node1("natureDesc") = "9" Then
                        valsalarydropdown2 = "Free education"
                    ElseIf Node1("natureDesc") = "10" Then
                        valsalarydropdown2 = "Gifts, vouchers, etc."
                    ElseIf Node1("natureDesc") = "11" Then
                        valsalarydropdown2 = "Credit card expenses"
                    ElseIf Node1("natureDesc") = "12" Then
                        valsalarydropdown2 = "Club expenses"
                    ElseIf Node1("natureDesc") = "13" Then
                        valsalarydropdown2 = "Use of movable assets by employees"
                    ElseIf Node1("natureDesc") = "14" Then
                        valsalarydropdown2 = "Transfer of assets to employee"
                    ElseIf Node1("natureDesc") = "15" Then
                        valsalarydropdown2 = "Value of any other benefit/amenity/service/privilege"
                    ElseIf Node1("natureDesc") = "16" Then
                         valsalarydropdown2 = "Stock options allotted or transferred by employer being an eligible start-up referred to in section 80-IAC-Tax to be deferred"
                    ElseIf Node1("natureDesc") = "17" Then
                        valsalarydropdown2 = "Stock options (non-qualified options) other than ESOP in col 16 above."
                    ElseIf Node1("natureDesc") = "18" Then
                        valsalarydropdown2 = "Contribution by employer to fund and scheme taxable under section 17(2)(vii)"
                    ElseIf Node1("natureDesc") = "19" Then
                        valsalarydropdown2 = "Annual accretion by way of interest, dividend etc. to the balance at the credit of fund and scheme referred to in section 17(2)(vii) and taxable under section 17(2)(viia)"

                    ElseIf Node1("natureDesc") = "21" Then
                        valsalarydropdown2 = "Stock options allotted or transferred by employer being an eligible start-up referred to in section 80-IAC-Tax not to be deferred"
                    ElseIf Node1("natureDesc") = "OTH" Then
                        valsalarydropdown2 = "Other benefits or amenities"
                    End If
                    Sheet39.Cells(rowcount2, ValueSection).value = valsalarydropdown2
                End If
'                If Sheet39.Cells(rowcount2, ValueDescription).Locked = False Then
'                    Sheet39.Cells(rowcount2, ValueDescription).value = Node1("othNatOfInc")
'                End If
                
                If Sheet39.Cells(rowcount2, ValueAmount).Locked = False Then
                
                    Sheet39.Cells(rowcount2, ValueAmount).value = Node1("othAmount")
                
                End If
                
                
            Next Node1
            End If
            
        'Malli_AY_2025_26------------
          '2nd prefill source
          Else:
          
           Dim valueOfPerquisites_2ndpfl  As Variant
           valueOfPerquisites_2ndpfl = Node("salarys")("valueOfPerquisites")
           
           If valueOfPerquisites_2ndpfl <> "" Then
           
                    rowcount2 = getRowNo(Sheet39.Range("Salary.ValueSection" & rowcount).name)
                    
                    ValueSection = Sheet39.Range("Salary.ValueSection" & rowcount).Column
                    ValueDescription = Sheet39.Range("Salary.ValueDescription" & rowcount).Column
                    ValueAmount = Sheet39.Range("Salary.ValueAmount" & rowcount).Column
                    
                    If Sheet39.Range("Salary.ValueSection" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ValueSection" & rowcount).ClearContents
                    End If
                    If Sheet39.Range("Salary.ValueDescription" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ValueDescription" & rowcount).ClearContents
                    End If
                    If Sheet39.Range("Salary.ValueAmount" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ValueAmount" & rowcount).ClearContents
                    End If
                    
                    If Sheet39.Cells(rowcount2, ValueSection).Locked = False Then
                                Sheet39.Cells(rowcount2, ValueSection).value = "Other benefits or amenities"
                    End If
                    If Sheet39.Cells(rowcount2, ValueAmount).Locked = False Then
                                Sheet39.Cells(rowcount2, ValueAmount).value = valueOfPerquisites_2ndpfl
                                valueOfPerquisites_2ndpfl = ""
                    End If
           End If
          
          '----------------------------------------
            End If
            
            '"NatureOfProfitInLieuOfSalary" start
            Set init = Node("salarys")("natureOfProfitInLieuOfSalary")
            If init.exists("othersIncDtls") Then
            '1
            Set nodeList1 = Node("salarys")("natureOfProfitInLieuOfSalary")("othersIncDtls")
            Dim ProfitSection, ProfitDescription, ProfitAmount As Variant
            
            ProfitSection = Sheet39.Range("Salary.ProfitSection" & rowcount).Column
            ProfitDescription = Sheet39.Range("Salary.ProfitDescription" & rowcount).Column
            ProfitAmount = Sheet39.Range("Salary.ProfitAmount" & rowcount).Column

            TotalExRow = Range("Salary.ProfitSection" & rowcount).Rows.count
            
            TotalXMLRow = nodeList1.count
            TotalDiffRow = TotalXMLRow - TotalExRow
            
            If (TotalXMLRow > 0) Then
                If Sheet39.Range("Salary.ProfitSection" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ProfitSection" & rowcount).ClearContents
                End If
                If Sheet39.Range("Salary.ProfitDescription" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ProfitDescription" & rowcount).ClearContents
                End If
                If Sheet39.Range("Salary.ProfitAmount" & rowcount).Locked = False Then
                    Sheet39.Range("Salary.ProfitAmount" & rowcount).ClearContents
                End If
            End If
            
            If (TotalDiffRow > 0) Then
                AddSalary17_3 (TotalDiffRow)
            End If
            
            rowcount3 = getRowNo(Sheet39.Range("Salary.ProfitSection" & rowcount).name)
            rowcount3 = rowcount3 - 1
            If nodeList1.count <= 0 Then
                Else
            For Each Node1 In nodeList1
                rowcount3 = rowcount3 + 1
                
                If Sheet39.Cells(rowcount3, ProfitSection).Locked = False Then
                    Dim valsalarydropdown3
                    If Node1("natureDesc") = "1" Then
'                        valsalarydropdown3 = "Compensation due/received by an assessee from his employer or former employer in connection with the  termination of his employment or modification thereto"
                         valsalarydropdown3 = "Any compensation due or received by an assessee from an employer or former employer in connection with the termination of his employment or modification thereto."
                    ElseIf Node1("natureDesc") = "2" Then
                        valsalarydropdown3 = "Any payment due/received by an assessee from his employer or a former employer or from a provident or other fund, sum received under Keyman Insurance Policy including Bonus thereto"""
                    ElseIf Node1("natureDesc") = "3" Then
                        valsalarydropdown3 = "Any amount due/received by assessee from any person before joining or after cessation of employment with that person"
                    ElseIf Node1("natureDesc") = "OTH" Then
                        valsalarydropdown3 = "Any Other"
                    End If
                    Sheet39.Cells(rowcount3, ProfitSection).value = valsalarydropdown3
                End If
'                If Sheet39.Cells(rowcount3, ProfitDescription).Locked = False Then
'                    Sheet39.Cells(rowcount3, ProfitDescription).value = Node1("othNatOfInc")
'                End If
                If Sheet39.Cells(rowcount3, ProfitAmount).Locked = False Then
                    Sheet39.Cells(rowcount3, ProfitAmount).value = Node1("othAmount")
                End If
                
            Next Node1
            End If
            'Malli_AY_2025_26------------
          '2nd prefill source
          Else:
             Dim profitsinLieuOfSalary_2ndpfl As Variant
              profitsinLieuOfSalary_2ndpfl = Node("salarys")("profitsinLieuOfSalary")
              
              If profitsinLieuOfSalary_2ndpfl <> "" Then
              
                        ProfitSection = Sheet39.Range("Salary.ProfitSection" & rowcount).Column
                        ProfitDescription = Sheet39.Range("Salary.ProfitDescription" & rowcount).Column
                        ProfitAmount = Sheet39.Range("Salary.ProfitAmount" & rowcount).Column
                        
                        rowcount3 = getRowNo(Sheet39.Range("Salary.ProfitSection" & rowcount).name)
                        
                        If Sheet39.Range("Salary.ProfitSection" & rowcount).Locked = False Then
                        Sheet39.Range("Salary.ProfitSection" & rowcount).ClearContents
                        End If
                        If Sheet39.Range("Salary.ProfitDescription" & rowcount).Locked = False Then
                        Sheet39.Range("Salary.ProfitDescription" & rowcount).ClearContents
                        End If
                        If Sheet39.Range("Salary.ProfitAmount" & rowcount).Locked = False Then
                        Sheet39.Range("Salary.ProfitAmount" & rowcount).ClearContents
                        End If
                        
                        If Sheet39.Cells(rowcount3, ProfitSection).Locked = False Then
                                   Sheet39.Cells(rowcount3, ProfitSection).value = "Any Other"
                        End If
                        If Sheet39.Cells(rowcount3, ProfitAmount).Locked = False Then
                                   Sheet39.Cells(rowcount3, ProfitAmount).value = profitsinLieuOfSalary_2ndpfl
                                   profitsinLieuOfSalary_2ndpfl = ""
                        End If
              
              End If
          '-------------------------------
            
            End If
            cnt = cnt + 1
    Next Node
    End If
    
    
    
  '------------------------------------------------------------------------------
    
    
If Not IsNull(jsonObject("form24q")) Then 'SIT-97244 Konda_28_07_2025

Set Nodelist = jsonObject("form24q")("allwncExemptUs10DtlsType")
    
    NatureColNo = Sheet39.Range("Salary.Nature1").Column
    DescriptionColNo = Sheet39.Range("Salary.Description1").Column
    AmountColNo = Sheet39.Range("Salary.Amount1").Column
    
    TotalExRow = Sheet39.Range("Salary.Nature1").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
   ' AddRows_Others_Pre_fill
    
    
    If (TotalDiffRow > 0) Then
                AddRows_Others_Pre_fill (TotalDiffRow)
    End If
    
    If (TotalXMLRow > 0) Then
        If Sheet39.Range("Salary.Nature1").Locked = False Then
            Sheet39.Range("Salary.Nature1").ClearContents
        End If
        If Sheet39.Range("Salary.Description1").Locked = False Then
            Sheet39.Range("Salary.Description1").ClearContents
        End If
        If Sheet39.Range("Salary.Amount1").Locked = False Then
            Sheet39.Range("Salary.Amount1").ClearContents
        End If
    End If
    rowcount = getRowNo(Sheet39.Range("Salary.Nature1").name)
    rowcount = rowcount - 1
    cnt = 0

    For Each Node In Nodelist
        
       '--------------Malli-----------
       'SIT-72538
    'If node("salNatureDesc") = "10(5)" Or node("salNatureDesc") = "10(10)" Or node("salNatureDesc") = "10(10A)" Or node("salNatureDesc") = "10(10AA)" Or node("salNatureDesc") = "10(13A)" Then
    If Node("salNatureDesc") = "10(5)" Or Node("salNatureDesc") = "10(10)" Or Node("salNatureDesc") = "10(10A)" Or Node("salNatureDesc") = "10(10AA)" Then
      
        rowcount = rowcount + 1
    'Change-4, ANK-130, 26.05.2022
            iNature = "Sec " & UCase(Node("salNatureDesc"))
'            iNature = "Sec " & UCase(node("SalNatureDesc"))

            sNature = Findtext(iNature, "Salary.OthersAllowance")
            If Node("salNatureDesc") = "10(10B)(i)" Then
'            If node("SalNatureDesc") = "10(10B)(i)" Then
                sNature = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
            ElseIf Node("salNatureDesc") = "10(10B)(ii)" Then
'            ElseIf node("SalNatureDesc") = "10(10B)(ii)" Then
                sNature = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
            ElseIf Node("salNatureDesc") = "OTH" Then
'            ElseIf node("SalNatureDesc") = "OTH" Then
                sNature = "ANY OTHER"
            End If
            If Sheet39.Cells(rowcount, NatureColNo).Locked = False Then
               If sNature <> "Sec " Then
                Sheet39.Cells(rowcount, NatureColNo).value = sNature
                End If
            End If
            If Sheet39.Cells(rowcount, DescriptionColNo).Locked = False Then
                Sheet39.Cells(rowcount, DescriptionColNo).value = UCase(Node("salNatureDesc"))
            End If
            
            'Malli
        If Sheet39.Cells(rowcount, NatureColNo).value <> "" Then
            If Sheet39.Cells(rowcount, AmountColNo).Locked = False Then
                Sheet39.Cells(rowcount, AmountColNo).value = Node("salOthAmount")
'                Sheet39.Cells(rowcount, AmountColNo).value = node("SalOthAmount")
            End If
        End If
'---end change---
        cnt = cnt + 1
         End If
    'Malli-----------------------------------
    Next Node
    
End If

'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

'If Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False Then
'    Sheet39.Range("SAL.Entertainment_allowance_1").value = jsonObject("form24q")("incomeDeductions")("entertainmentAlw16Ii")
'End If

'If Sheet39.Range("SAL.DeductionUnderSection1").Locked = False Then
'    Sheet39.Range("SAL.DeductionUnderSection1").value = jsonObject("form24q")("incomeDeductions")("professionalTaxUs16Iii")
'End If

'------------------------------------------


'\\* SIT-111718 SIT-111762 Malli_AY_2026-27
'Description:Old Tax dependent fields should not get loaded in utility if we change from New Tax to Old Tax in utility.*//

''Malli_AY_2026-27   03/02/2026
'Dim EA10_13A_placeofwork, EA10_13A_placeofwork_pfl As Variant
'
''Place of Residence
''lastFiledITR.scheduleEA10_13A.placeofwork
'EA10_13A_placeofwork = jsonObject("lastFiledITR")("scheduleEA10_13A")("placeofwork")
'
'If EA10_13A_placeofwork <> "" And Sheet39.Range("Sch10of13A_PlaceofWrk").Locked = False Then
'        If EA10_13A_placeofwork = "1" Then
'               EA10_13A_placeofwork_pfl = "1. Metro"
'        ElseIf EA10_13A_placeofwork = "2" Then
'               EA10_13A_placeofwork_pfl = "2. Non Metro"
'        Else
'               EA10_13A_placeofwork_pfl = "(Select)"
'        End If
'    Sheet39.Range("Sch10of13A_PlaceofWrk").value = EA10_13A_placeofwork_pfl
'End If

 

'-----------------------------
   

End Function


'Malli
'Form10IA.dateOfFiling 'Form10IA.acknNumber

Function ImportSchedule80DD_80U_pfl(jsonText As String)
On Error Resume Next
Dim section80UUsrType As Variant
Dim init, jsonObject As Object
Set jsonObject = ParseJson.ParseJson(jsonText)
 
Dim schedule80DDObject As Object
 

Dim dateOfFiling, acknNumber As Object

dateOfFiling = jsonObject("Form10IA")("dateOfFiling")


If dateOfFiling <> "" Then
Dim dob As Variant
dob = Mid(dateOfFiling, 9, 2) & "/" & Mid(dateOfFiling, 6, 2) & "/" & Mid(dateOfFiling, 1, 4)
If Sheet58.Range("DatefilingFm10IA_80DD").Locked = False Then
   Sheet58.Range("DatefilingFm10IA_80DD").value = dob      'DatefilingFm10IA_80DD
End If

If Sheet58.Range("DatefilingFm10IA_80U").Locked = False Then
   Sheet58.Range("DatefilingFm10IA_80U").value = dob       'DatefilingFm10IA_80U
End If

End If
Dim acknNumber_1 As Variant
acknNumber_1 = jsonObject("Form10IA")("acknNumber")

If acknNumber <> "" And Sheet58.Range("AckNoFm10IAfiled_80DD").Locked = False Then
        Sheet58.Range("AckNoFm10IAfiled_80DD").value = acknNumber_1   'AckNoFm10IAfiled_80DD
End If
If acknNumber <> "" And Sheet58.Range("AckNoFm10IAfiled_80U").Locked = False Then
Sheet58.Range("AckNoFm10IAfiled_80U").value = acknNumber_1            'AckNoFm10IAfiled_80U
End If




section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType")

If section80UUsrType = "1" Then
    section80UUsrType = "1-Self with disability"
ElseIf section80UUsrType = "2" Then
    section80UUsrType = "2-Self with severe disability"
Else
    section80UUsrType = ""
End If

If section80UUsrType <> "" And Sheet1.Range("Naturedisability_80U").Locked = False Then
     Sheet58.Range("Naturedisability_80U").value = section80UUsrType
End If



End Function
Sub AddRows_Others_Pre_fill(nOfRows As Long)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheets("Schedule S").Activate
        
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    'numberofrows = iRows
    'For te = 1 To Sheet39.Range("SalarySectionCount").value
    
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.Nature1")
        EfilingCommon.DefinedgridNameRange = "Salary.Nature1||Salary.Description1||Salary.Amount1"
        'ActiveCellRange = EfilingCommon.searchLastRow("Salary.Nature" & te)
        
       ' If te > 1 Then
             EfilingCommon.insertRowUnderSectionWithFormula (nOfRows)
            
        'Else
          '  numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
       ' End If
    'Next
    
    

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
'    For i = 1 To Sheet39.Range("SalarySectionCount").value
       ' newfrmsize = Sheet39.Range("SalarySectionItem").value
      '  newfrmsize = newfrmsize + numberofrows
        'Sheet39.Range("SalarySectionItem").value = newfrmsize
'    Next


endline:
   
'----------------Lock Password-------------------START---
   Sheet39.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
Function ImportSchedulePTI(jsonText As String)
On Error Resume Next
    Dim XpathOfPTI As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As Object
    Dim Node As Object
    Dim strdate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim rowcount, cnt, i As Long
    Dim init, jsonObject As Object
   
    
'Change-28, prefill-tracker, D
Set jsonObject = ParseJson.ParseJson(jsonText)
Set init = jsonObject("form64A")("schedulePTI")
If init.exists("schedulePTIDtls") Then
   
    Set Nodelist = jsonObject("form64A")("SchedulePTI")("SchedulePTIDtls")

    TotalExRow = Sheet41.Range("cntr.Ptirptfrm").value
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If TotalDiffRow > 0 Then
        For i = 1 To TotalDiffRow
            Sheet41.Activate
            AddRows_PTI
        Next
    End If
    Dim InvstmntCvrdUs115UA115UB As Variant
    
    rowcount = 0
    cnt = 0
    For Each Node In Nodelist
    
            rowcount = rowcount + 1
            InvstmntCvrdUs115UA115UB = UCase(Node(("InvstmntCvrdUs115UA115UB")))
                
                If InvstmntCvrdUs115UA115UB = "A" Then
                   InvstmntCvrdUs115UA115UB = "SECTION 115UA"
                ElseIf InvstmntCvrdUs115UA115UB = "B" Then
                   InvstmntCvrdUs115UA115UB = "SECTION 115UB"
                Else
                    InvstmntCvrdUs115UA115UB = "(Select)"
                End If
                
        If Sheet41.Range("PTI_Investment_UA_UB" & rowcount).Locked = False Then
            Sheet41.Range("PTI_Investment_UA_UB" & rowcount).value = InvstmntCvrdUs115UA115UB
        End If
        If Sheet41.Range("PTI_NameOfBusiness" & rowcount).Locked = False Then
            Sheet41.Range("PTI_NameOfBusiness" & rowcount).value = UCase(Node("BusinessName"))
        End If
        If Sheet41.Range("PTI_PAN" & rowcount).Locked = False Then
            Sheet41.Range("PTI_PAN" & rowcount).value = UCase(Node("BusinessPAN"))
        End If
        If Sheet41.Range("PTI_HPIncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_HPIncomeAmt" & rowcount).value = Node("IncFromHP")("AmountOfInc")
        End If
        If Sheet41.Range("PTI_HP_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_HP_ShareAmt" & rowcount).value = Node("IncFromHP")("CurrYrLossShareByInvstFund")
        End If
'        If Sheet41.Range("PTI_HP_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_HP_NetAmt" & rowcount).value = Node("IncFromHP")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_HPTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_HPTDSAmt" & rowcount).value = Node("IncFromHP")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_CG_STCGIncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCGIncomeAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_NetAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_STCGTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCGTDSAmt" & rowcount).value = Node("capitalGainsPTI")("ShortTermCG")("TDSAmount")
'        End If
'Chandru
        If Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_SectionA_IncomeAmt" & rowcount).value = Node("CapitalGainsPTI")("STCG_Sec111A")("CurrYrLossShareByInvstFund")
        End If
        If Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_SectionA_ShareAmt" & rowcount).value = Node("CapitalGainsPTI")("STCG_Sec111A")("CurrYrLossShareByInvstFund")
        End If
        
        
'        If Sheet41.Range("PTI_CG_STCG_SectionA_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_SectionA_NetAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Sec111A")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_SectionA_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Sec111A")("TDSAmount")
'        End If

'
'        If Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & rowcount).value = Node("capitalGainsPTI")("IncOthSrc")("AmountOfInc")
'        End If

'Chandru
        If Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_Others_IncomeAmt" & rowcount).value = Node("CapitalGainsPTI")("STCG_Others")("AmountOfInc")
        End If
        
        If Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_STCG_Others_ShareAmt" & rowcount).value = Node("CapitalGainsPTI")("STCG_Others")("CurrYrLossShareByInvstFund")
        End If
'        If Sheet41.Range("PTI_CG_STCG_Others_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_Others_NetAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Others")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_STCG_Others_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("stcg_Others")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_CG_LTCGIncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCGIncomeAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_ShareAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_ShareAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("CurrYrLossShareByInvstFund")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_NetAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_LTCGTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCGTDSAmt" & rowcount).value = Node("capitalGainsPTI")("LongTermCG")("TDSAmount")
'        End If
'Chandru - modified
        If Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_LTCG_SectionA_IncomeAmt" & rowcount).value = Node("CapitalGainsPTI")("LTCG_Sec112A")("AmountOfInc")
        End If
        If Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_LTCG_SectionA_ShareAmt" & rowcount).value = Node("CapitalGainsPTI")("LTCG_Sec112A")("CurrYrLossShareByInvstFund")
        End If
        
        
'        If Sheet41.Range("PTI_CG_LTCG_SectionA_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_SectionA_NetAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Sec112A")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_SectionA_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Sec112A")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_LTCG_Others_IncomeAmt" & rowcount).value = Node("CapitalGainsPTI")("LTCG_Others")("AmountOfInc")
        End If
        If Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_CG_LTCG_Others_ShareAmt" & rowcount).value = Node("CapitalGainsPTI")("LTCG_Others")("CurrYrLossShareByInvstFund")
        End If
        
'        If Sheet41.Range("PTI_CG_LTCG_Others_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_Others_NetAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Others")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_CG_LTCG_Others_TdsAmt" & rowcount).value = Node("capitalGainsPTI")("LTCG_Others")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_Us_iv_IncomeAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_iv_IncomeAmta" & rowcount).value = Node("IncClmdPTI")("TotalSec23FBB")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_iv_NetAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_iv_NetAmta" & rowcount).value = Node("IncClmdPTI")("TotalSec23FBB")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_Us_iv_TdsAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_iv_TdsAmta" & rowcount).value = Node("IncClmdPTI")("TotalSec23FBB")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_UsIncomeAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsIncomeAmta" & rowcount).value = Node("IncClmdPTI")("Sec23FBB")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_NetAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_NetAmta" & rowcount).value = Node("IncClmdPTI")("Sec23FBB")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_UsTDSAmta" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsTDSAmta" & rowcount).value = Node("IncClmdPTI")("Sec23FBB")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_Usb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Usb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SectionCode")
'        End If
'        If Sheet41.Range("PTI_UsIncomeAmtb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsIncomeAmtb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SecBCIncExmptDtl")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_NetAmtb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_NetAmtb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SecBCIncExmptDtl")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_UsTDSAmtb" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsTDSAmtb" & rowcount).value = Node("IncClmdPTI")("SecBIncExmptDtl")("SecBCIncExmptDtl")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_Usc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Usc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SectionCode")
'        End If
'        If Sheet41.Range("PTI_UsIncomeAmtc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsIncomeAmtc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SecBCIncExmptDtl")("AmountOfInc")
'        End If
'        If Sheet41.Range("PTI_Us_NetAmtc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_Us_NetAmtc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SecBCIncExmptDtl")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_UsTDSAmtc" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_UsTDSAmtc" & rowcount).value = Node("IncClmdPTI")("SecCIncExmptDtl")("SecBCIncExmptDtl")("TDSAmount")
'        End If
'        If Sheet41.Range("PTI_OSIncomeAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OSIncomeAmt" & rowcount).value = Node("os_Others")("amountOfInc")
'        End If
'        If Sheet41.Range("PTI_OS_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_NetAmt" & rowcount).value = Node("OS_Others")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_OSTDSAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OSTDSAmt" & rowcount).value = Node("OS_Others")("TDSAmount")
'        End If
        If Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_OS_Dividend_IncomeAmt" & rowcount).value = Node("OS_Dividend")("AmountOfInc")
        End If
'        If Sheet41.Range("PTI_OS_Dividend_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Dividend_NetAmt" & rowcount).value = Node("OS_Dividend")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_OS_Dividend_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Dividend_TdsAmt" & rowcount).value = Node("OS_Dividend")("TDSAmount")
'        End If
'Chandru
        If Sheet41.Range("PTI_OS_Others_IncomeAmt" & rowcount).Locked = False Then
            Sheet41.Range("PTI_OS_Others_IncomeAmt" & rowcount).value = Node("IncOthSrc")("AmountOfInc")
        End If
'        If Sheet41.Range("PTI_OS_Others_NetAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Others_NetAmt" & rowcount).value = Node("IncOthSrc")("NetIncomeLoss")
'        End If
'        If Sheet41.Range("PTI_OS_Others_TdsAmt" & rowcount).Locked = False Then
'            Sheet41.Range("PTI_OS_Others_TdsAmt" & rowcount).value = Node("IncOthSrc")("TDSAmount")
'        End If
                
            cnt = cnt + 1
    Next Node
    
    RecTDS1 = cnt
End If
End Function


'Malli_AY_2026-27_03/02/2026
'lastFiledITR.schedule80E.schedule80EDtls

Function ImportSchedule_80E_Pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80E As Object
Dim Node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80E, IFSC_80E, BankName_80E, PANof_80E, LoanAcctNum_80E, LoanDate_80E, TotalLoanAmt_80E, LoanOutStanding_80E, Intrst_80E As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80E, rowcount, cnt As Long
Dim strdate As String
Dim YYYY, MM, DD As String
 
 
    
Set jsonObject = ParseJson.ParseJson(jsonText)
 
Set NodeList80E = jsonObject("lastFiledITR")
 
If NodeList80E.exists("schedule80E") Then
 
Set Nodelist = jsonObject("lastFiledITR")("schedule80E")("schedule80EDtls")

    LoanfrmBankOrInstitute_80E = Sheet60.Range("LoanfrmBankOrInstitute.80E").Column
    BankName_80E = Sheet60.Range("bankName.80E").Column
    LoanAcctNum_80E = Sheet60.Range("loanAccNum.80E").Column
    LoanDate_80E = Sheet60.Range("loanDate.80E").Column
    TotalLoanAmt_80E = Sheet60.Range("loanAmt.80E").Column
     
    
    TotalExRow = Sheet60.Range("LoanfrmBankOrInstitute.80E").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet60.Range("LoanfrmBankOrInstitute.80E").Locked = False Then
            Sheet60.Range("LoanfrmBankOrInstitute.80E").ClearContents
        End If
        
        If Sheet60.Range("bankName.80E").Locked = False Then
            Sheet60.Range("bankName.80E").ClearContents
        End If
        
        If Sheet60.Range("loanAccNum.80E").Locked = False Then
            Sheet60.Range("loanAccNum.80E").ClearContents
        End If
        
        If Sheet60.Range("loanDate.80E").Locked = False Then
            Sheet60.Range("loanDate.80E").ClearContents
        End If
        
        If Sheet60.Range("loanAmt.80E").Locked = False Then
            Sheet60.Range("loanAmt.80E").ClearContents
        End If
        
        If Sheet60.Range("loanOutstanding.80E").Locked = False Then
            Sheet60.Range("loanOutstanding.80E").ClearContents
        End If
        
        If Sheet60.Range("Intrst.80E").Locked = False Then
            Sheet60.Range("Intrst.80E").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80E (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80E").name)
    
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
        If Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80E).Locked = False Then
            Dim LoanTknFrom_80E, LoanTknFrom_80E_pfl
            
            LoanTknFrom_80E_pfl = Node("loanTknFrom")
            If UCase(LoanTknFrom_80E_pfl) = UCase("B") Then
                LoanTknFrom_80E = "Bank"
            ElseIf UCase(LoanTknFrom_80E_pfl) = UCase("I") Then
                LoanTknFrom_80E = "Institution"
            Else
                LoanTknFrom_80E = "(Select)"
            End If
            Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80E).value = LoanTknFrom_80E
        End If
        
        
        If Sheet60.Cells(rowcount, BankName_80E).Locked = False Then
            Sheet60.Cells(rowcount, BankName_80E).value = Node("bankOrInstnName")
        End If
         
        If Sheet60.Cells(rowcount, LoanAcctNum_80E).Locked = False Then
            Sheet60.Cells(rowcount, LoanAcctNum_80E).value = Node("loanAccNoOfBankOrInstnRefNo")
        End If
                
            strdate = Node("dateofLoan")
            If strdate <> "" Then
               strdate = Mid(strdate, 9, 2) & "/" & Mid(strdate, 6, 2) & "/" & Mid(strdate, 1, 4)
            End If
        
        If Sheet60.Cells(rowcount, LoanDate_80E).Locked = False Then
            Sheet60.Cells(rowcount, LoanDate_80E).value = strdate
            strdate = ""
        End If
        
        If Sheet60.Cells(rowcount, TotalLoanAmt_80E).Locked = False Then
            Sheet60.Cells(rowcount, TotalLoanAmt_80E).value = Node("totalLoanAmt")
        End If
        
         
        
        cnt = cnt + 1
    Next Node
    Rec80E = cnt
 End If
End Function

'Malli_AY_2026-27_22/01/2026
'lastFiledITR.Schedule80EE.Schedule80EEDtls
Function ImportSchedule_80EE_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EE As Object
Dim Node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EE, IFSC_80EE, BankName_80EE, PANof_80EE, LoanAcctNum_80EE, LoanDate_80EE, TotalLoanAmt_80EE, LoanOutStanding_80EE, Intrst_80EE As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EE, rowcount, cnt As Long
Dim strdate As String
Dim YYYY, MM, DD As String

 
 
    
Set jsonObject = ParseJson.ParseJson(jsonText)
Set NodeList80EE = jsonObject("lastFiledITR")
 
If NodeList80EE.exists("schedule80EE") Then
 
Set Nodelist = jsonObject("lastFiledITR")("schedule80EE")("schedule80EEDtls")

    LoanfrmBankOrInstitute_80EE = Sheet60.Range("LoanfrmBankOrInstitute.80EE").Column
    BankName_80EE = Sheet60.Range("bankName.80EE").Column
    LoanAcctNum_80EE = Sheet60.Range("loanAccNum.80EE").Column
    LoanDate_80EE = Sheet60.Range("loanDate.80EE").Column
    TotalLoanAmt_80EE = Sheet60.Range("loanAmt.80EE").Column
     
    
    TotalExRow = Sheet60.Range("LoanfrmBankOrInstitute.80EE").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet60.Range("LoanfrmBankOrInstitute.80EE").Locked = False Then
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").ClearContents
        End If
        If Sheet60.Range("IFSC.80EE").Locked = False Then
            Sheet60.Range("IFSC.80EE").ClearContents
        End If
        If Sheet60.Range("bankName.80EE").Locked = False Then
            Sheet60.Range("bankName.80EE").ClearContents
        End If
        If Sheet60.Range("PAN.80EE").Locked = False Then
            Sheet60.Range("PAN.80EE").ClearContents
        End If
        If Sheet60.Range("loanAccNum.80EE").Locked = False Then
            Sheet60.Range("loanAccNum.80EE").ClearContents
        End If
        If Sheet60.Range("loanDate.80EE").Locked = False Then
            Sheet60.Range("loanDate.80EE").ClearContents
        End If
        If Sheet60.Range("loanAmt.80EE").Locked = False Then
            Sheet60.Range("loanAmt.80EE").ClearContents
        End If
        If Sheet60.Range("loanOutstanding.80EE").Locked = False Then
            Sheet60.Range("loanOutstanding.80EE").ClearContents
        End If
        
        If Sheet60.Range("Intrst.80EE").Locked = False Then
            Sheet60.Range("Intrst.80EE").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EE (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80EE").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
        If Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80EE).Locked = False Then
            Dim LoanTknFrom_80EE, LoanTknFrom_80EE_pfl
            
            LoanTknFrom_80EE_pfl = Node("loanTknFrom")
            If UCase(LoanTknFrom_80EE_pfl) = UCase("B") Then
                LoanTknFrom_80EE = "Bank"
            ElseIf UCase(LoanTknFrom_80EE_pfl) = UCase("I") Then
                LoanTknFrom_80EE = "Institution"
            Else
                LoanTknFrom_80EE = "(Select)"
            End If
            Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80EE).value = LoanTknFrom_80EE
        End If
         
        If Sheet60.Cells(rowcount, BankName_80EE).Locked = False Then
            Sheet60.Cells(rowcount, BankName_80EE).value = Node("bankOrInstnName")
        End If
         
        If Sheet60.Cells(rowcount, LoanAcctNum_80EE).Locked = False Then
            Sheet60.Cells(rowcount, LoanAcctNum_80EE).value = Node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strdate = Node("dateofLoan")
            If strdate <> "" Then
                strdate = Mid(strdate, 9, 2) & "/" & Mid(strdate, 6, 2) & "/" & Mid(strdate, 1, 4)
            End If
        
        If Sheet60.Cells(rowcount, LoanDate_80EE).Locked = False Then
            Sheet60.Cells(rowcount, LoanDate_80EE).value = strdate
            strdate = ""
        End If
        If Sheet60.Cells(rowcount, TotalLoanAmt_80EE).Locked = False Then
            Sheet60.Cells(rowcount, TotalLoanAmt_80EE).value = Node("totalLoanAmt")
        End If
        
         
        cnt = cnt + 1
    Next Node
    Rec80EE = cnt
End If
End Function

'Malli_AY_2026-27_03/02/2026
'lastFiledITR.schedule80EEA.schedule80EEADtls
Function ImportSchedule_80EEA_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EEA As Object
Dim Node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEA, IFSC_80EEA, BankName_80EEA, PANof_80EEA, LoanAcctNum_80EEA, LoanDate_80EEA, TotalLoanAmt_80EEA, LoanOutStanding_80EEA, Intrst_80EEA As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEA, rowcount, cnt As Long
Dim strdate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson.ParseJson(jsonText)
Set NodeList80EEA = jsonObject("lastFiledITR")
 
If NodeList80EEA.exists("schedule80EEA") Then
  
        
Dim PropStmpDtyVal_80EEA
PropStmpDtyVal_80EEA = jsonObject("lastFiledITR")("schedule80EEA")("propStmpDtyVal")

If PropStmpDtyVal_80EEA <> "" Then
    If Sheet60.Range("Stampduty.80EEA").Locked = False Then
            Sheet60.Range("Stampduty.80EEA").value = PropStmpDtyVal_80EEA
        End If
End If



Set Nodelist = jsonObject("lastFiledITR")("schedule80EEA")("schedule80EEADtls")

    LoanfrmBankOrInstitute_80EEA = Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Column
    BankName_80EEA = Sheet60.Range("bankName.80EEA").Column
    LoanAcctNum_80EEA = Sheet60.Range("loanAccNum.80EEA").Column
    LoanDate_80EEA = Sheet60.Range("loanDate.80EEA").Column
    TotalLoanAmt_80EEA = Sheet60.Range("loanAmt.80EEA").Column
     
    
    TotalExRow = Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Locked = False Then
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
        End If
        
        If Sheet60.Range("bankName.80EEA").Locked = False Then
            Sheet60.Range("bankName.80EEA").ClearContents
        End If
        
        If Sheet60.Range("loanAccNum.80EEA").Locked = False Then
            Sheet60.Range("loanAccNum.80EEA").ClearContents
        End If
        If Sheet60.Range("loanDate.80EEA").Locked = False Then
            Sheet60.Range("loanDate.80EEA").ClearContents
        End If
        If Sheet60.Range("loanAmt.80EEA").Locked = False Then
            Sheet60.Range("loanAmt.80EEA").ClearContents
        End If
        If Sheet60.Range("loanOutstanding.80EEA").Locked = False Then
            Sheet60.Range("loanOutstanding.80EEA").ClearContents
        End If
        
        If Sheet60.Range("Intrst.80EEA").Locked = False Then
            Sheet60.Range("Intrst.80EEA").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EEA (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80EEA").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        
        If Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).Locked = False Then
            Dim LoanTknFrom_80EEA, LoanTknFrom_80EEA_pfl
            
            LoanTknFrom_80EEA_pfl = Node("loanTknFrom")
            If UCase(LoanTknFrom_80EEA_pfl) = UCase("B") Then
                LoanTknFrom_80EEA = "Bank"
            ElseIf UCase(LoanTknFrom_80EEA_pfl) = UCase("I") Then
                LoanTknFrom_80EEA = "Institution"
            Else
                LoanTknFrom_80EEA = "(Select)"
            End If
            Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).value = LoanTknFrom_80EEA
        End If
        
        If Sheet60.Cells(rowcount, BankName_80EEA).Locked = False Then
            Sheet60.Cells(rowcount, BankName_80EEA).value = Node("bankOrInstnName")
        End If
         
        If Sheet60.Cells(rowcount, LoanAcctNum_80EEA).Locked = False Then
            Sheet60.Cells(rowcount, LoanAcctNum_80EEA).value = Node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strdate = Node("dateofLoan")
            If strdate <> "" Then
                strdate = Mid(strdate, 9, 2) & "/" & Mid(strdate, 6, 2) & "/" & Mid(strdate, 1, 4)
            End If
        
        If Sheet60.Cells(rowcount, LoanDate_80EEA).Locked = False Then
            Sheet60.Cells(rowcount, LoanDate_80EEA).value = strdate
            strdate = ""
        End If
        If Sheet60.Cells(rowcount, TotalLoanAmt_80EEA).Locked = False Then
            Sheet60.Cells(rowcount, TotalLoanAmt_80EEA).value = Node("totalLoanAmt")
        End If
         
        cnt = cnt + 1
    Next Node
    Rec80EEA = cnt
 End If
End Function

'Malli_AY_2026-27_03/02/2026
'lastFiledITR.schedule80EEB.schedule80EEBDtls

Function ImportSchedule_80EEB_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EEB As Object
Dim Node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEB, IFSC_80EEB, BankName_80EEB, Vehicle_value_80EEB, VehicleRegNum_80EEB, PANof_80EEB, LoanAcctNum_80EEB, LoanDate_80EEB, TotalLoanAmt_80EEB, LoanOutStanding_80EEB, Intrst_80EEB As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEB, rowcount, cnt As Long
Dim strdate As String
Dim YYYY, MM, DD As String

 
 
    
Set jsonObject = ParseJson.ParseJson(jsonText)
Set NodeList80EEB = jsonObject("lastFiledITR")
 
If NodeList80EEB.exists("schedule80EEB") Then
  
Set Nodelist = jsonObject("lastFiledITR")("schedule80EEB")("schedule80EEBDtls")

    LoanfrmBankOrInstitute_80EEB = Sheet60.Range("LoanfrmBankOrInstitute.80EEB").Column
    BankName_80EEB = Sheet60.Range("bankName.80EEB").Column
    LoanAcctNum_80EEB = Sheet60.Range("loanAccNum.80EEB").Column
    LoanDate_80EEB = Sheet60.Range("loanDate.80EEB").Column
    TotalLoanAmt_80EEB = Sheet60.Range("loanAmt.80EEB").Column
    VehicleRegNum_80EEB = Sheet60.Range("VehicleRegNum.80EEB").Column
     
    
    
    TotalExRow = Sheet60.Range("LoanfrmBankOrInstitute.80EEB").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet60.Range("LoanfrmBankOrInstitute.80EEB").Locked = False Then
            Sheet60.Range("LoanfrmBankOrInstitute.80EEB").ClearContents
        End If
         
        If Sheet60.Range("bankName.80EEB").Locked = False Then
            Sheet60.Range("bankName.80EEB").ClearContents
        End If
        
        If Sheet60.Range("loanAccNum.80EEB").Locked = False Then
            Sheet60.Range("loanAccNum.80EEB").ClearContents
        End If
        If Sheet60.Range("loanDate.80EEB").Locked = False Then
            Sheet60.Range("loanDate.80EEB").ClearContents
        End If
        If Sheet60.Range("loanAmt.80EEB").Locked = False Then
            Sheet60.Range("loanAmt.80EEB").ClearContents
        End If
        If Sheet60.Range("loanOutstanding.80EEB").Locked = False Then
            Sheet60.Range("loanOutstanding.80EEB").ClearContents
        End If
        
        If Sheet60.Range("VehicleRegNum.80EEB").Locked = False Then
            Sheet60.Range("VehicleRegNum.80EEB").ClearContents
        End If
        If Sheet60.Range("Intrst.80EEB").Locked = False Then
            Sheet60.Range("Intrst.80EEB").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EEB (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80EEB").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each Node In Nodelist
        rowcount = rowcount + 1
        If Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).Locked = False Then
            Dim LoanTknFrom_80EEB, LoanTknFrom_80EEB_pfl
            LoanTknFrom_80EEB_pfl = Node("loanTknFrom")
            
            If UCase(LoanTknFrom_80EEB_pfl) = UCase("B") Then
                LoanTknFrom_80EEB = "Bank"
            ElseIf UCase(LoanTknFrom_80EEB_pfl) = UCase("I") Then
                LoanTknFrom_80EEB = "Institution"
            Else
                LoanTknFrom_80EEB = "(Select)"
            End If
            Sheet60.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).value = LoanTknFrom_80EEB
        End If
          
        If Sheet60.Cells(rowcount, BankName_80EEB).Locked = False Then
            Sheet60.Cells(rowcount, BankName_80EEB).value = Node("bankOrInstnName")
        End If
         
        If Sheet60.Cells(rowcount, LoanAcctNum_80EEB).Locked = False Then
            Sheet60.Cells(rowcount, LoanAcctNum_80EEB).value = Node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strdate = Node("dateofLoan")
            If strdate <> "" Then
                strdate = Trim(Mid(strdate, 9, 2) & "/" & Mid(strdate, 6, 2) & "/" & Mid(strdate, 1, 4))
            End If
        
        If Sheet60.Cells(rowcount, LoanDate_80EEB).Locked = False Then
            Sheet60.Cells(rowcount, LoanDate_80EEB).value = strdate
            strdate = ""
        End If
        If Sheet60.Cells(rowcount, TotalLoanAmt_80EEB).Locked = False Then
            Sheet60.Cells(rowcount, TotalLoanAmt_80EEB).value = Node("totalLoanAmt")
        End If
         
        If Sheet60.Cells(rowcount, VehicleRegNum_80EEB).Locked = False Then
            Sheet60.Cells(rowcount, VehicleRegNum_80EEB).value = Node("vehicleRegNo")
        End If
         
        cnt = cnt + 1
    Next Node
    Rec80EEB = cnt
End If
End Function







'Malli_AY_2026_27  02/06/2026

Function ImportScheduleCG_POPUPPfl(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary, NodelistCG_POPUP As Object
Dim Node, Nodelist As Object
Dim CG_POPUP_slno, CG_POPUP_Description, CG_POPUP_PrptyAdrs, CG_POPUP_TrasAmt As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80E, rowcount, cnt As Long
Dim strdate As String
Dim YYYY, MM, DD As String



Set jsonObject = ParseJson.ParseJson(jsonText)
Set NodelistCG_POPUP = jsonObject("insights")

If NodelistCG_POPUP.exists("capitalGains") Then


       Set Nodelist = jsonObject("insights")("capitalGains")("propertyDetails")
       
 
    CG_POPUP_Description = Sheet62.Range("CG_POPUP_Description").Column
    CG_POPUP_PrptyAdrs = Sheet62.Range("CG_POPUP_PrptyAdrs").Column
    CG_POPUP_TrasAmt = Sheet62.Range("CG_POPUP_TrasAmt").Column
    
    TotalExRow = Sheet62.Range("CG_POPUP_Description").Rows.count
    
            TotalXMLRow = Nodelist.count
            TotalDiffRow = TotalXMLRow - TotalExRow
            
            If (TotalXMLRow > 0) Then
                    If Sheet62.Range("CG_POPUP_Description").Locked = False Then
                        Sheet62.Range("CG_POPUP_Description").ClearContents
                    End If
                    
                    If Sheet62.Range("CG_POPUP_PrptyAdrs").Locked = False Then
                        Sheet62.Range("CG_POPUP_PrptyAdrs").ClearContents
                    End If
                    
                    If Sheet62.Range("CG_POPUP_TrasAmt").Locked = False Then
                        Sheet62.Range("CG_POPUP_TrasAmt").ClearContents
                    End If
            End If
            
            
            rowcount = getRowNo(Sheet62.Range("CG_POPUP_Description").name)
    
            rowcount = rowcount - 1
            cnt = 0
            CG_POPUPPfl = False
            For Each Node In Nodelist
                rowcount = rowcount + 1
                
                If rowcount <= 6 Then
                     Application.EnableEvents = False
                     
                     Dim infoCodepfl, addrDetailofl, transactionAmountpfl As Variant
                     infoCodepfl = Node("infoCode")
                     addrDetailofl = Node("addressDetailWithZipCode")("addrDetail")
                     transactionAmountpfl = Node("transactionAmount")
                     
                        If infoCodepfl <> "" And Sheet62.Cells(rowcount, CG_POPUP_Description).Locked = False Then
                           Sheet62.Cells(rowcount, CG_POPUP_Description).value = Node("infoCode")
                           CG_POPUPPfl = True
                        End If
                        
                        If addrDetailofl <> "" And Sheet62.Cells(rowcount, CG_POPUP_PrptyAdrs).Locked = False Then
                           Sheet62.Cells(rowcount, CG_POPUP_PrptyAdrs).value = Node("addressDetailWithZipCode")("addrDetail")
                           CG_POPUPPfl = True
                        End If
                        
                        If transactionAmountpfl <> "" And Sheet62.Cells(rowcount, CG_POPUP_TrasAmt).Locked = False Then
                           Sheet62.Cells(rowcount, CG_POPUP_TrasAmt).value = Node("transactionAmount")
                           CG_POPUPPfl = True
                        End If
                        
                    Application.EnableEvents = True
                End If
            Next Node
            
            If CG_POPUPPfl <> False Then
            Application.EnableEvents = False
            Sheet13.Activate
              UserForm3.Show
            Application.EnableEvents = True
            End If

End If

 
End Function













