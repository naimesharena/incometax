Attribute VB_Name = "msSI"
Option Explicit

'Public end_os As Long
Public age As Long

Public SecCode_SI As Variant          'added  new
Public SplRatePercent_SI As Variant   'added  new
Public SplRateInc_SI As Variant       'added  new
Public SplRateIncTax_SI As Variant    'added  new
Public TotSplRateIncTax_SI As Variant 'added  new

Public SecCode_SI1 As Variant          'added  new
Public SplRatePercent_SI1 As Variant   'added  new
Public SplRateInc_SI1 As Variant       'added  new
Public SplRateIncTax_SI1 As Variant    'added  new

Public end_SI, end_SI2, rngname_SI, rngname_SI2 As Variant 'added new
Dim msgValidateSheetSI, MsgIFSheet As String 'added new

'added new
Sub ValidateSI()
If Not ValidatesheetSI Then
        Sheet21.Activate
        fmsgboxsmall (msgValidateSheetSI)
        CloseMsg
End If
End Sub

'added new
Function ValidatesheetSI()
 ValidatesheetSI = True
    
    
    
    If Not ValidateSecCode_SI() Then ValidatesheetSI = False
    If (Len(Sheet21.Range("SI.SecCode").item(1).value) > 0) Then
        If Not ValidateSplRatePercent_SI() Then ValidatesheetSI = False
        If Not ValidateSplRateInc_SI() Then ValidatesheetSI = False
        If Not ValidateSplRateIncTax_SI() Then ValidatesheetSI = False
        If Not ValidateTotSplRateIncTax_SI() Then ValidatesheetSI = False
    End If
      '--Newly added by Bindu on 25th May 26-----'
        If Not ValidateSI_Editable() Then ValidatesheetSI = False
    '--End
    
    
    setTblinfo_SI2
    
    If (end_SI <> end_SI2) Then
        msgbox_SI ("* * Section Code in Sch SI is Mandatory")
        ValidatesheetSI = False
        Exit Function
    End If
End Function

'added new

Sub setTblinfo_SI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet21.Range("SI.SecCode").count
    Set rangecells = Sheet21.Range("SI.SecCode").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_SI = ccount
    rngname_SI = "SI.SecCode;SI.SplRatePercent;SI.SplRateInc;SI.SplRateIncCalc;SI.SplRateIncTax;"
End Sub

'added new
Sub setTblinfo_SI2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet21.Range("SI.SplRateIncTax").count
    Set rangecells = Sheet21.Range("SI.SplRateIncTax").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = 0 Then
            ccount = ccount + 1
        End If
    Next
    end_SI2 = ccount
    rngname_SI2 = "SI.SecCode;SI.SplRatePercent;SI.SplRateInc;SI.SplRateIncCalc;SI.SplRateIncTax;"
 End Sub

'added new
Function ValidateSecCode_SI() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    ValidateSecCode_SI = True
    setTblinfo_SI
    
    Set rangecells = Sheet21.Range("SI.SecCode").Cells
    Set rangecells1 = Sheet21.Range("SI.SecCode1").Cells
    ReDim SecCode_SI(end_SI)
    ReDim SecCode_SI1(end_SI)
    
    For i = 1 To end_SI
        SecCode_SI(i) = rangecells.item(i).value
        SecCode_SI(i) = Mid(SecCode_SI(i), 1, 15)
        
        SecCode_SI1(i) = rangecells1.item(i).value
'Konda updated SIT-91200 on 10-04-2025
'        SecCode_SI1(i) = Mid(SecCode_SI1(i), 1, 15)
         SecCode_SI1(i) = Mid(SecCode_SI1(i), 1, 18)
        
        If SecCode_SI(i) = "21cii" Then
            SecCode_SI(i) = "21ciii"
        End If
        
         If SecCode_SI1(i) = "21cii" Then
            SecCode_SI1(i) = "21ciii"
        End If
        
        If Not chkCompulsory(SecCode_SI(i)) Then
            msgbox_SI ("* SecCode at Sr. No  " & i & "  in Sheet SI  is Mandatory")
            ValidateSecCode_SI = False
            Exit Function
        End If
    Next
End Function
'added new
Function ValidateSplRatePercent_SI() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    ValidateSplRatePercent_SI = True
    setTblinfo_SI

    Set rangecells = Sheet21.Range("SI.SplRatePercent").Cells
    Set rangecells1 = Sheet21.Range("SI.SplRatePercent1").Cells
    
    ReDim SplRatePercent_SI(end_SI)
    ReDim SplRatePercent_SI1(end_SI)
    
    For i = 1 To end_SI
        SplRatePercent_SI(i) = rangecells.item(i).value
        SplRatePercent_SI1(i) = rangecells1.item(i).value
        'SplRatePercent_SI(i) = Mid(SplRatePercent_SI(i), 1, 5)
    Next
End Function
'added new
Function ValidateSplRateInc_SI() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    ValidateSplRateInc_SI = True
    setTblinfo_SI
        
    Dim DTAAIncome As Variant
    
    DTAAIncome = Sheet21.Range("DTAA_INCOME_CG").value
    
    If (DTAAIncome) = "" Then
        msgbox_SI ("* DTAA CG Income in Sheet: SI should be Non negative, no decimal,and cannot be greater than Income, please put enter 0 for no  income.")
        ValidateSplRateInc_SI = False
        Exit Function
    End If
    
     DTAAIncome = Sheet21.Range("DTAA_INCOME_OS").value
    
    If (DTAAIncome) = "" Then
        msgbox_SI ("* DTAA OS Income in Sheet: SI should be Non negative, no decimal,and cannot be greater than Income, please put enter 0 for no  income.")
        ValidateSplRateInc_SI = False
        Exit Function
    End If
        
    Set rangecells = Sheet21.Range("SI.SplRateInc").Cells
    Set rangecells1 = Sheet21.Range("SI.SplRateInc1").Cells
    ReDim SplRateInc_SI(end_SI)
    ReDim SplRateInc_SI1(end_SI)
    For i = 1 To end_SI
        SplRateInc_SI(i) = rangecells.item(i).value
        SplRateInc_SI1(i) = rangecells1.item(i).value
        If (SplRateInc_SI(i)) < 0 Then
            msgbox_SI ("* Income for row " & i & " in Sheet: SI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateSplRateInc_SI = False
            Exit Function
        End If
    Next
End Function
'added new
Function ValidateSplRateIncTax_SI() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Dim DTAATax As Variant
    ValidateSplRateIncTax_SI = True
    setTblinfo_SI
    Set rangecells = Sheet21.Range("SI.SplRateIncTax").Cells
    Set rangecells1 = Sheet21.Range("SI.SplRateIncTax1").Cells
    ReDim SplRateIncTax_SI(end_SI)
    ReDim SplRateIncTax_SI1(end_SI)
    
    DTAATax = Sheet21.Range("DTAA_TAX_CG").value
    If DTAATax = "" Then
    msgbox_SI ("* DTAA CG TAX in Sheet: SI should be Non negative, no decimal please put enter 0 for no Tax.")
    ValidateSplRateIncTax_SI = False
    Exit Function
    End If
    
    DTAATax = Sheet21.Range("Os.taxthereon.111").value
    If DTAATax = "" Then
    msgbox_SI ("* System calculated Tax thereon of DTAA in Sheet: SI should be Non negative, no decimal please put enter 0 for no Tax.")
    ValidateSplRateIncTax_SI = False
    Exit Function
    End If
    
    '--Not required
'    DTAATax = Sheet21.Range("DTAA_Inc_OS").value
'    If DTAATax > 0 Then
'    msgbox_SI ("* System calculated Tax thereon of DTAA in Sheet: SI should be Non negative, no decimal please put enter 0 for no Tax.")
'    End If
'
'
'    DTAATax = Sheet21.Range("DTAA_Income_OS_1").value
'    If DTAATax > 0 Then
'    msgbox_SI ("* System calculated Tax thereon of DTAA in Sheet: SI should be Non negative, no decimal please put enter 0 for no Tax.")
'    End If
'
'    DTAATax = Sheet21.Range("DTAA_Income_OS_2").value
'    If DTAATax > 0 Then
'    msgbox_SI ("* System calculated Tax thereon of DTAA in Sheet: SI should be Non negative, no decimal please put enter 0 for no Tax.")
'    End If
    
    For i = 1 To end_SI
        SplRateIncTax_SI(i) = rangecells.item(i).value
        SplRateIncTax_SI1(i) = rangecells1.item(i).value
        
        If (SplRateIncTax_SI(i)) < 0 Then
            msgbox_SI ("* System Calc Tax for row " & i & " in Sheet: SI should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateSplRateIncTax_SI = False
            Exit Function
        End If
    Next
End Function
'added new
Function ValidateTotSplRateIncTax_SI() As Boolean
    ValidateTotSplRateIncTax_SI = True
    TotSplRateIncTax_SI = Sheet21.Range("SI.TotSplRateIncTax").value
End Function



Sub populateSI()

Dim count As Long
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim rangecells4 As Range
Dim sectionDes As String
Dim section As Long
Dim count2 As Long
Dim j As Long



Application.EnableEvents = False
Application.ScreenUpdating = False

   'assiging value to the sheet SI
   
 '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet21.Unprotect Password:=sPassword


Dim exemption As Double
Dim partbSetoffInc As Double

 exemption = getExemption_SI
 partbSetoffInc = Sheet2.Range("Sheet2.TotalIncome").value - Sheet2.Range("Sheet2.IncChargeableTaxSplRates").value
 
'----------------Unlock Password-------------------START---
    Sheet21.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
        
Sheet21.Range("exemptionAmt").value = getExemption_SI
Sheet21.Range("Age").value = age

 '----------------Lock Password-------------------START---
  Sheet21.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    
   If (partbSetoffInc > 0) Then
    If (partbSetoffInc >= exemption) Then
        exemption = 0
    Else
        exemption = exemption - partbSetoffInc
    End If
  End If
  
'Sheet2.Range("exemptionAmt").value = exemption

  If (exemption <= Sheet21.Range("SI.SplRateInc").Cells.item(5).value) Then
        Sheet21.Range("SI.SplRateIncCalc").Cells.item(5).value = Application.WorksheetFunction.Max(0, Sheet21.Range("SI.SplRateInc").Cells.item(5).value - exemption)
        exemption = 0
  Else
       Sheet21.Range("SI.SplRateIncCalc").Cells.item(5).value = 0
        exemption = Application.WorksheetFunction.Max(0, exemption - Sheet21.Range("SI.SplRateInc").Cells.item(5).value)
 End If
    
 If (exemption <= Sheet21.Range("SI.SplRateInc").Cells.item(3).value) Then
       Sheet21.Range("SI.SplRateIncCalc").Cells.item(3).value = Application.WorksheetFunction.Max(0, Sheet21.Range("SI.SplRateInc").Cells.item(3).value - exemption)
        exemption = 0
Else
        Sheet21.Range("SI.SplRateIncCalc").Cells.item(3).value = 0
        exemption = Application.WorksheetFunction.Max(0, exemption - Sheet21.Range("SI.SplRateInc").Cells.item(3).value)
End If
    
If (exemption <= Sheet21.Range("SI.SplRateInc").Cells.item(6).value) Then
        Sheet21.Range("SI.SplRateIncCalc").Cells.item(6).value = Application.WorksheetFunction.Max(0, Sheet21.Range("SI.SplRateInc").Cells.item(6).value - exemption)
        exemption = 0
Else
        Sheet21.Range("SI.SplRateIncCalc").Cells.item(6).value = 0
        exemption = Application.WorksheetFunction.Max(0, exemption - Sheet21.Range("SI.SplRateInc").Cells.item(6).value)
End If
  

 '----------------Lock Password-------------------START---
  Sheet21.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Function getExemption_SI()
Dim resStatus As String
Dim Status As String


age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").value)
 resStatus = Sheet1.Range("sheet1.ResidentialStatus1").value               'RES , NRI
Status = Sheet1.Range("sheet1.Status").value
        If (Mid(Status, 1, 1) = "I" And age > 59 And age < 80 And (Mid(resStatus, 1, 3) = "RES" Or Mid(resStatus, 1, 3) = "NOR")) Then
            getExemption_SI = 300000
        ElseIf (Mid(Status, 1, 1) = "I" And age >= 80 And (Mid(resStatus, 1, 3) = "RES" Or Mid(resStatus, 1, 3) = "NOR")) Then
            getExemption_SI = 500000
        ElseIf ((Mid(Status, 1, 1) = "I" Or Mid(Status, 1, 1) = "H") And (Mid(resStatus, 1, 3) = "RES" Or Mid(resStatus, 1, 3) = "NOR")) Then
            getExemption_SI = 250000
        Else
            getExemption_SI = 0
        End If

End Function
'dpk2
Sub setTblinfo_OS_si()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet14.Range("os.SourceDescription").count
 Set rangecells = Sheet14.Range("os.SourceDescription").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 end_os = ccount
End Sub


'added new
Function msgbox_SI(strmsg As String) As String
     msgValidateSheetSI = msgValidateSheetSI & strmsg & Chr(13)
End Function

'added new

Function DefaultAmtIncluded_SPI() As String
    DefaultAmtIncluded_SPI = "0"
End Function
'added new
Function DefaultSplRatePercent_SI() As String
    DefaultSplRatePercent_SI = "1"
End Function
'added new
Function DefaultSplRateInc_SI() As String
    DefaultSplRateInc_SI = "0"
End Function
'added new
Function DefaultSplRateIncTax_SI() As String
    DefaultSplRateIncTax_SI = "0"
End Function
'added new
Function DefaultTotSplRateIncTax_SI() As String
    DefaultTotSplRateIncTax_SI = "0"
End Function

Function ValidateSI_Editable() As Boolean

ValidateSI_Editable = True

If Sheet21.Range("Editable_111A_20_STCG") = "Err" Then '111A 20%
    msgbox_SI ("**Amount reported at ""111A - STCG on shares units on which STT paid as applicable"" of Schedule SI can not be more than sum of amount at Sl. No. A3ie and A4a of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_PTI_20_STCG") = "Err" Then 'PTI 20%
    msgbox_SI ("**Amount reported at ""Pass Through Income in the nature of Short Term Capital Gain chargeable @ 20% "" of Schedule SI can not be more than amount at Sl. No. A8a of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_115AD_20_STCG") = "Err" Then '115AD 20%
    msgbox_SI ("**Amount reported at ""115AD(1)(b)(ii)-Proviso - Short term capital gains referred to in section 111A rws. 115AD by FII, as applicable u/s 115AD(1)(b)(ii)"" of Schedule SI can not be more than amount at Sl. No. A3iie of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If
    
If Sheet21.Range("Editable_20_STCG_Sum") = "Err" Then 'STCG 20% SUM
    msgbox_SI ("**The amount entered for STCG@20% in Schedule SI is not equal to Sl. No. 5vi of Schedule BFLA. Kindly recheck the amounts entered in Schedule SI")
    ValidateSI_Editable = False
    Exit Function
End If


'STCG_30%======================

If Sheet21.Range("Editable_115AD_30_STCG") = "Err" Then '115AD 30%
    msgbox_SI ("**Amount reported at ""115AD(ii) -Short term capital gains (other than on equity share or equity oriented mutual fund referred to in section 111A) by an FII"" of Schedule SI can not be more than amount at Sl. No. A5e of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_PTI_30_STCG") = "Err" Then 'PTI 30%
    msgbox_SI ("**Amount reported at ""Pass Through Income in the nature of Short Term Capital Gain chargeable @ 30% "" of Schedule SI can not be more than amount at Sl. No. A8b of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_30_STCG_Sum") = "Err" Then 'STCG 30% SUM
    msgbox_SI ("**The amount entered for STCG@30% in Schedule SI is not equal to Sl. No. 5vii of Schedule BFLA. Kindly recheck the amounts entered in Schedule SI")
    ValidateSI_Editable = False
    Exit Function
End If

'LTCG_12.5%======================


If Sheet21.Range("Editable_21b_12_5_LTCG") = "Err" Then '21b--G23 Cell
    msgbox_SI ("**Amount reported at ""112 -LTCG on others"" of Schedule SI can not be more than sum of amount at Sl. No. B1g, B2e, B5c, B9e, B10 of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_22_12_5_LTCG") = "Err" Then '12_5_22--G26 Cell
      msgbox_SI ("**Amount reported at ""112(1) - LTCG on listed securities/ units without indexation"" of Schedule SI can not be more than amount at Sl. No. B3ie of Schedule CG")
      ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_112A_12_5_LTCG") = "Err" Then '112A--G28 Cell
    msgbox_SI ("**Amount reported at ""112A -LTCG on sale of shares /units of equity oriented fund/units of business trust on which STT is paid"" of Schedule SI can not be more than amount at Sl. No. B4c of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_1121C_12_5_LTCG") = "Err" Then '1121c--G30 Cell
      msgbox_SI ("**Amount reported at ""112(1)(c)(iii) - Long term capital gains on transfer of unlisted securities in the case of non-residents"" of Schedule SI can not be more than amount at Sl. No. B6ie of Schedule CG")
     ValidateSI_Editable = False
    Exit Function
End If


If Sheet21.Range("Editable_115AC1_12_5_LTCG") = "Err" Then '115AC1--G50
      msgbox_SI ("**Amount reported at ""115AC(1)(c) -Long term capital gains arising from their transfer of bonds or GDR purchased in foreign currency in case of a non-resident"" of Schedule SI can not be more than amount at Sl. No. B6iie of Schedule CG")
      ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_115ACA1b_12_5_LTCG") = "Err" Then '115ACA1b--G53 Cell
    msgbox_SI ("**Amount reported at ""115ACA(1)(b) - Long term capital gains arising from the transfer of GDR purchased in foreign currency -resident"" of Schedule SI can not be more than amount at Sl. No. B3iie of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_115AD1Biii_FII_12_5_LTCG") = "Err" Then '115AD1Biii_FII--G58 Cell
      msgbox_SI ("**Amount reported at ""115AD(1)(b)(iii) - Long term capital gains (other than on equity share or equity oriented mutual fund referred to in section 112A) by an FII"" of Schedule SI can not be more than amount at Sl. No. B6iiie of Schedule CG")
     ValidateSI_Editable = False
    Exit Function
End If
   
If Sheet21.Range("Editable_115AD1Biii_12_5_LTCG") = "Err" Then '115AD1biii--G59 Cell
    msgbox_SI ("**Amount reported at ""115AD(1)(b)(iii) Proviso -For NON-RESIDENTS - From sale of equity share in a company or unit of equity oriented fund or unit of a business trust on which STT is paid under section 112A"" of Schedule SI can not be more than amount at Sl. No. B7c of Schedule CG")
   ValidateSI_Editable = False
    Exit Function
End If
    
If Sheet21.Range("Editable_115EB_12_5_LTCG") = "Err" Then '115Eb--G72 Cell
    msgbox_SI ("**Amount reported at ""115E(b) - Long term capital gains of a non-resident Indian on any foreign exchange asset "" of Schedule SI can not be more than amount at Sl. No. B8c of Schedule CG")
   ValidateSI_Editable = False
    Exit Function
End If
   
If Sheet21.Range("Editable_PTI12_12_5_LTCG") = "Err" Then 'PTI_LTCG12--G79 Cell
    msgbox_SI ("**Amount reported at ""Pass Through Income in the nature of Long Term Capital Gain chargeable @ 12.5% u/s 112A"" of Schedule SI can not be more than amount at Sl. No. B11a1 of Schedule CG")
    ValidateSI_Editable = False
    Exit Function
End If
    
If Sheet21.Range("Editable_PTIOTH112A_12_5_LTCG") = "Err" Then 'PTIOTH112A_12--G81 Cell
    msgbox_SI ("**Amount reported at ""Pass Through Income in the nature of Long Term Capital Gain chargeable @ 12.5% other than u/s 112A"" of Schedule SI can not be more than amount at Sl. No. B11a2 of Schedule CG")
   ValidateSI_Editable = False
    Exit Function
End If

If Sheet21.Range("Editable_12_5_LTCG_Sum") = "Err" Then 'LTCG 12.5% SUM
    msgbox_SI ("**The amount entered for LTCG@12.5% in Schedule SI is not equal to Sl. No. 5x of Schedule BFLA. Kindly recheck the amounts entered in Schedule SI")
    ValidateSI_Editable = False
    Exit Function
End If


End Function
