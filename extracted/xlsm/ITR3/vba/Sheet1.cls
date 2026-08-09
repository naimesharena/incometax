Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

Public ActiveCellRange As String
Dim DateRangecells As Range
   Dim x, dateCtr As Long
   Dim CreditDate As String
Private Sub Img_Home_Click()
   Sheet35.Activate
   If UCase(ActiveSheet.name) <> UCase("Home") Then
      Sheet35.Activate
    End If
End Sub
Private Sub index_Click()
   Sheet35.Activate
    If UCase(ActiveSheet.name) <> UCase("Home") Then
      Sheet35.Activate
    End If
End Sub

Sub ab()
Application.EnableEvents = False
End Sub

Private Sub Cmd_GenerateXML_Click()
Generate_XML.Gen_XML
End Sub

Sub LockUnlockIBANSELECT()
On Error Resume Next
Application.EnableEvents = False
Sheet7.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR" Then
        Sheet7.Range("Sheet9.IBANSELECT").ClearContents
        Sheet7.Range("Sheet9.IBANSELECT").Locked = True
        Sheet7.Range("Sheet9.IBANSELECT").Interior.Color = "&HD8D8D8"
        Sheet7.Range("Sheet9.IBANSELECT").value = "Yes"
    Else
        Sheet7.Range("Sheet9.IBANSELECT").ClearContents
        Sheet7.Range("Sheet9.IBANSELECT").Locked = False
        Sheet7.Range("Sheet9.IBANSELECT").Interior.Color = "&HCCFFCC"
        Sheet7.Range("Sheet9.IBANSELECT").value = "Yes"
    End If

    Sheet7.LockUnlockBANKDTL
    Application.EnableEvents = False
Sheet7.Protect Password:=getmsgstate
End Sub

'added by Chetan C M on 27/01/2026 for AY 2026-27
'start--
Sub LockUnlockPLSec44()
On Error Resume Next
Application.EnableEvents = False
Sheet4.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
    
        Sheet4.Range("PL.GrossReceipt44B").ClearContents
        Sheet4.Range("PL.GrossReceipt44B").Locked = False
        Sheet4.Range("PL.GrossReceipt44B").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.GrossReceipt44BB").ClearContents
        Sheet4.Range("PL.GrossReceipt44BB").Locked = False
        Sheet4.Range("PL.GrossReceipt44BB").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.GrossReceipt44BBA").ClearContents
        Sheet4.Range("PL.GrossReceipt44BBA").Locked = False
        Sheet4.Range("PL.GrossReceipt44BBA").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.GrossReceipt44BBC").ClearContents
        Sheet4.Range("PL.GrossReceipt44BBC").Locked = False
        Sheet4.Range("PL.GrossReceipt44BBC").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.GrossReceipt44BBD").ClearContents
        Sheet4.Range("PL.GrossReceipt44BBD").Locked = False
        Sheet4.Range("PL.GrossReceipt44BBD").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.NetProfit44B").ClearContents
        Sheet4.Range("PL.NetProfit44B").Locked = False
        Sheet4.Range("PL.NetProfit44B").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.NetProfit44BB").ClearContents
        Sheet4.Range("PL.NetProfit44BB").Locked = False
        Sheet4.Range("PL.NetProfit44BB").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.NetProfit44BBA").ClearContents
        Sheet4.Range("PL.NetProfit44BBA").Locked = False
        Sheet4.Range("PL.NetProfit44BBA").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.NetProfit44BBC").ClearContents
        Sheet4.Range("PL.NetProfit44BBC").Locked = False
        Sheet4.Range("PL.NetProfit44BBC").Interior.Color = "&HCCFFCC"
        
        Sheet4.Range("PL.NetProfit44BBD").ClearContents
        Sheet4.Range("PL.NetProfit44BBD").Locked = False
        Sheet4.Range("PL.NetProfit44BBD").Interior.Color = "&HCCFFCC"
        
    
    Else
        
        Sheet4.Range("PL.GrossReceipt44B").ClearContents
        Sheet4.Range("PL.GrossReceipt44B").Locked = True
        Sheet4.Range("PL.GrossReceipt44B").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.GrossReceipt44BB").ClearContents
        Sheet4.Range("PL.GrossReceipt44BB").Locked = True
        Sheet4.Range("PL.GrossReceipt44BB").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.GrossReceipt44BBA").ClearContents
        Sheet4.Range("PL.GrossReceipt44BBA").Locked = True
        Sheet4.Range("PL.GrossReceipt44BBA").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.GrossReceipt44BBC").ClearContents
        Sheet4.Range("PL.GrossReceipt44BBC").Locked = True
        Sheet4.Range("PL.GrossReceipt44BBC").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.GrossReceipt44BBD").ClearContents
        Sheet4.Range("PL.GrossReceipt44BBD").Locked = True
        Sheet4.Range("PL.GrossReceipt44BBD").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.NetProfit44B").ClearContents
        Sheet4.Range("PL.NetProfit44B").Locked = True
        Sheet4.Range("PL.NetProfit44B").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.NetProfit44BB").ClearContents
        Sheet4.Range("PL.NetProfit44BB").Locked = True
        Sheet4.Range("PL.NetProfit44BB").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.NetProfit44BBA").ClearContents
        Sheet4.Range("PL.NetProfit44BBA").Locked = True
        Sheet4.Range("PL.NetProfit44BBA").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.NetProfit44BBC").ClearContents
        Sheet4.Range("PL.NetProfit44BBC").Locked = True
        Sheet4.Range("PL.NetProfit44BBC").Interior.Color = "&HD8D8D8"
        
        Sheet4.Range("PL.NetProfit44BBD").ClearContents
        Sheet4.Range("PL.NetProfit44BBD").Locked = True
        Sheet4.Range("PL.NetProfit44BBD").Interior.Color = "&HD8D8D8"
        
        
    End If

    Application.EnableEvents = False
Sheet4.Protect Password:=getmsgstate
End Sub
'--end


'added by Chetan C M on 09/02/2026 for AY 2026-27
'start--
Sub LockUnlockEIincome()
On Error Resume Next
Application.EnableEvents = False
Sheet24.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
    
        Sheet24.Range("SCHEI_DTAA_AmountOfIncome").ClearContents
        Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Locked = False
        Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Interior.Color = "&HCCFFCC"
        
        Sheet24.Range("SCHEI_DTAA_NatureOfIncome").ClearContents
        Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Locked = False
        Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Interior.Color = "&HCCFFCC"
        
        Sheet24.Range("SCHEI_DTAA_CountryCode").ClearContents
        Sheet24.Range("SCHEI_DTAA_CountryCode").Locked = False
        Sheet24.Range("SCHEI_DTAA_CountryCode").Interior.Color = "&HCCFFCC"
        
        Sheet24.Range("SCHEI_DTAA_Article").ClearContents
        Sheet24.Range("SCHEI_DTAA_Article").Locked = False
        Sheet24.Range("SCHEI_DTAA_Article").Interior.Color = "&HCCFFCC"
        
        Sheet24.Range("SCHEI_DTAA_HeadOfIncome").ClearContents
        Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Locked = False
        Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Interior.Color = "&HCCFFCC"
                
        Sheet24.Range("SCHEI_TRCObtained").ClearContents
        Sheet24.Range("SCHEI_TRCObtained").Locked = False
        Sheet24.Range("SCHEI_TRCObtained").Interior.Color = "&HCCFFCC"
        
        
    
    Else
    
       Sheet24.Range("SCHEI_DTAA_AmountOfIncome").ClearContents
        Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Locked = True
        Sheet24.Range("SCHEI_DTAA_AmountOfIncome").Interior.Color = "&HD8D8D8"
        
        Sheet24.Range("SCHEI_DTAA_NatureOfIncome").ClearContents
        Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Locked = True
        Sheet24.Range("SCHEI_DTAA_NatureOfIncome").Interior.Color = "&HD8D8D8"
        
        Sheet24.Range("SCHEI_DTAA_CountryCode").ClearContents
        Sheet24.Range("SCHEI_DTAA_CountryCode").Locked = True
        Sheet24.Range("SCHEI_DTAA_CountryCode").Interior.Color = "&HD8D8D8"
        
        Sheet24.Range("SCHEI_DTAA_Article").ClearContents
        Sheet24.Range("SCHEI_DTAA_Article").Locked = True
        Sheet24.Range("SCHEI_DTAA_Article").Interior.Color = "&HD8D8D8"
        
        Sheet24.Range("SCHEI_DTAA_HeadOfIncome").ClearContents
        Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Locked = True
        Sheet24.Range("SCHEI_DTAA_HeadOfIncome").Interior.Color = "&HD8D8D8"
                
        Sheet24.Range("SCHEI_TRCObtained").ClearContents
        Sheet24.Range("SCHEI_TRCObtained").Locked = True
        Sheet24.Range("SCHEI_TRCObtained").Interior.Color = "&HD8D8D8"
        
    End If

    Application.EnableEvents = False
Sheet24.Protect Password:=getmsgstate
End Sub
'--end

Sub UNLOCK5A()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet29.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_5_FA As Variant
Dim prngname_5_FA As String

rngname_5_FA = "S5A_NameOfSpouse||S5A_PANOfSpouse||S5A_AadhaarOfSpouse||BooksSpouse44ABFlg||BooksSpouse92EFlg"
UNLOCKRANGE5A (rngname_5_FA)
prngname_5_FA = "S5AR"
UNLOCKRANGE5A (prngname_5_FA)
'----------------Lock Password-------------------START---
  Sheet29.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub


Sub LOCK5A()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet29.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_5_FA As Variant
rngname_5_FA = "S5A_NameOfSpouse||S5A_PANOfSpouse||S5A_AadhaarOfSpouse||BooksSpouse44ABFlg||BooksSpouse92EFlg"
LOCKRANGE5A (rngname_5_FA)
rngname_5_FA = "S5AR"
LOCKRANGE5A (rngname_5_FA)


'----------------Lock Password-------------------START---
  Sheet29.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGE5A(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet29.Range(rangenamestring(i))
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Next
       Next
End Sub

Sub UNLOCKRANGE5A(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet29.Range(rangenamestring(i))
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Next
       Next
End Sub


Sub StatusBasedDefaultValue()
On Error Resume Next
Sheet1.Unprotect Password:=getmsgstate
Status = Sheet1.Range("sheet1.Status")
If (Mid(Status, 1, 1) = "H") Then

       Range("Sheet1.PortugeseCC5A") = "No"

End If

Sheet1.Protect Password:=getmsgstate

End Sub
Sub LockUnlock1398A_Gen()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.ClearContents
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = True
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.SeventhProvisoFlag").value = ""
'        Sheet201.Visible = xlSheetVisible
'        Sheet202.Visible = xlSheetVisible


'=====================================================
'Newly added by sai on 05/11/2024

        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents

        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True

        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.DepositAmount").MergeArea.ClearContents

        Range("Sheet1.DepositAmount").MergeArea.Locked = True

        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)


        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)


        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents

        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True

        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)



        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents

        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True

        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)


        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)

'End Change by sai on 05/11/2024

'========================================================================================

    Else
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = False
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.SeventhProvisoFlag").value = "(Select)"
'        Sheet201.Visible = xlSheetHidden
'        Sheet202.Visible = xlSheetHidden
    End If
    
    
    
    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
    
    
    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Rows("21:21").Hidden = False
    Sheet1.Rows("31:31").Hidden = False 'Added by Aavula Naresh for AY 2026-27
    
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet201.Visible = xlSheetVisible
    Sheet202.Visible = xlSheetVisible
    ThisWorkbook.Protect Password:=getmsgstate
    Else
    
     Sheet1.Unprotect Password:=getmsgstate
'     Sheet1.Rows("21:21").Hidden = True
     Sheet1.Rows("31:31").Hidden = True 'Added by Aavula Naresh for AY 2026-27
     
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet201.Visible = xlSheetHidden
    Sheet202.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockUnlockQVIA()
On Error Resume Next
Application.EnableEvents = False
Sheet19.Unprotect Password:=getmsgstate

    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
    
        Sheet19.Range("SELECT80D").MergeArea.ClearContents
        Sheet19.Range("SELECT80D").MergeArea.Locked = True
        Sheet19.Range("SELECT80D").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("SELECT80D").value = ""

        Sheet19.Range("SELECT80DB").MergeArea.ClearContents
        Sheet19.Range("SELECT80DB").MergeArea.Locked = True
        Sheet19.Range("SELECT80DB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("SELECT80DB").value = ""

        Sheet19.Range("SELECT80DC").MergeArea.ClearContents
        Sheet19.Range("SELECT80DC").MergeArea.Locked = True
        Sheet19.Range("SELECT80DC").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("SELECT80DC").value = ""
'Malli
'        Sheet19.Range("SELECT80U").MergeArea.ClearContents
'        Sheet19.Range("SELECT80U").MergeArea.Locked = True
'        Sheet19.Range("SELECT80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("SELECT80U").value = ""
'--------------------------

'Commented by sudha_2025-256

'        Sheet11.Range("scvia.Section80EEA").ClearContents
'        Sheet11.Range("scvia.Section80EEA").Interior.Color = (&HD8D8D8)
'        Sheet11.Range("scvia.Section80EEA").Locked = True
'        Sheet11.Range("scvia.Section80EEA").value = ""

'        Sheet11.Range("scvia.Section80EEB").ClearContents
'        Sheet11.Range("scvia.Section80EEB").Locked = True
'        Sheet11.Range("scvia.Section80EEB").Interior.Color = (&HD8D8D8)
'        Sheet11.Range("scvia.Section80EEB").value = ""

'commented by Chetan C M for AY 2026-27-----start

'        Sheet19.Range("scvia.Section80CCC").Interior.ColorIndex = 15   ', "AU"
'        Sheet19.Range("scvia.Section80CCC").MergeArea.Locked = True
'        Sheet19.Range("scvia.Section80CCC").value = ""

'        Sheet19.Range("scvia.Section80CCD_SE").Interior.ColorIndex = 15   ', "AU"
'        Sheet19.Range("scvia.Section80CCD_SE").MergeArea.Locked = True
'        Sheet19.Range("scvia.Section80CCD_SE").value = ""

'        Sheet19.Range("scvia.Section80CCD1B_SE").Interior.ColorIndex = 15   ', "AU"
'        Sheet19.Range("scvia.Section80CCD1B_SE").MergeArea.Locked = True
'        Sheet19.Range("scvia.Section80CCD1B_SE").value = ""
'----end
        
        Sheet19.Range("scvia.Section80CCD").Interior.ColorIndex = 15   ', "AU"
        Sheet19.Range("scvia.Section80CCD").MergeArea.Locked = True
        Sheet19.Range("scvia.Section80CCD").value = ""
        
        Sheet19.Range("scvia.Section80GG").Interior.ColorIndex = 15   ', "AU"
        Sheet19.Range("scvia.Section80GG").MergeArea.Locked = True
        Sheet19.Range("scvia.Section80GG").value = ""
        
'        '''''''''''''''
        Sheet19.Range("Sheet19.AckNum").Locked = True
        Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HD8D8D8)
        Sheet19.Range("Sheet19.AckNum").Font.Color = (&HD8D8D8)

        Sheet19.Range("Sheet19.AckNum.10CCD").MergeArea.Locked = True
        Sheet19.Range("Sheet19.AckNum.10CCD").Interior.Color = (&HD8D8D8)
        Sheet19.Range("Sheet19.AckNum.10CCD").Font.Color = (&HD8D8D8)
        Sheet19.Range("Sheet19.AckNum.10CCD").MergeArea.Locked = True

         Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = True
         Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HD8D8D8)
         Sheet19.Range("Sheet19.AckNum.10CCE").Font.Color = (&HD8D8D8)
         Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = True
'        ''''''''''''''
        
    Else
        Sheet19.Range("SELECT80D").MergeArea.ClearContents
        Sheet19.Range("SELECT80D").MergeArea.Locked = False
        Sheet19.Range("SELECT80D").MergeArea.Interior.Color = "&HCCFFCC"
        'Sheet19.Range("SELECT80D").value = "(Select)"

        Sheet19.Range("SELECT80DB").MergeArea.ClearContents
        Sheet19.Range("SELECT80DB").MergeArea.Locked = False
        Sheet19.Range("SELECT80DB").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet19.Range("SELECT80DC").MergeArea.ClearContents
        Sheet19.Range("SELECT80DC").MergeArea.Locked = False
        Sheet19.Range("SELECT80DC").MergeArea.Interior.Color = "&HCCFFCC"

'Malli
'        Sheet19.Range("SELECT80U").MergeArea.ClearContents
'        Sheet19.Range("SELECT80U").MergeArea.Locked = False
'        Sheet19.Range("SELECT80U").MergeArea.Interior.Color = "&HCCFFCC"
'--------------
'Commented by sudha_2025-26

'        Sheet11.Range("scvia.Section80EEA").ClearContents
'        Sheet11.Range("scvia.Section80EEA").Locked = False
'        Sheet11.Range("scvia.Section80EEA").Interior.Color = "&HCCFFCC"
'
'        Sheet11.Range("scvia.Section80EEB").ClearContents
'        Sheet11.Range("scvia.Section80EEB").Locked = False
'        Sheet11.Range("scvia.Section80EEB").Interior.Color = "&HCCFFCC"

'commented by Chetan C M for AY 2026-27----start--
'        Sheet19.Range("scvia.Section80CCC").MergeArea.Locked = False
'        Sheet19.Range("scvia.Section80CCC").Interior.Color = "&HCCFFCC"
        
'        Sheet19.Range("scvia.Section80CCD_SE").MergeArea.Locked = False
'        Sheet19.Range("scvia.Section80CCD_SE").Interior.Color = "&HCCFFCC"
        
'        Sheet19.Range("scvia.Section80CCD1B_SE").MergeArea.Locked = False
'        Sheet19.Range("scvia.Section80CCD1B_SE").Interior.Color = "&HCCFFCC"
'---end
        
        Sheet19.Range("scvia.Section80CCD").MergeArea.Locked = False
        Sheet19.Range("scvia.Section80CCD").Interior.Color = "&HCCFFCC"
        
        Sheet19.Range("scvia.Section80GG").MergeArea.Locked = False
        Sheet19.Range("scvia.Section80GG").Interior.Color = "&HCCFFCC"
        
    End If
    
    If SheetALL.Range("bacValue").value = 2 Then
    
         Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HCCFFCC)
         Sheet19.Range("Sheet19.AckNum").Font.Color = vbBlack

         Sheet19.Range("Sheet19.AckNum.10CCD").MergeArea.Locked = False
         Sheet19.Range("Sheet19.AckNum.10CCD").Interior.Color = (&HCCFFCC)
         Sheet19.Range("Sheet19.AckNum.10CCD").Font.Color = vbBlack

         Sheet19.Range("Sheet19.AckNum.10CCE").MergeArea.Locked = False
         Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HCCFFCC)
         Sheet19.Range("Sheet19.AckNum.10CCE").Font.Color = vbBlack

        Sheet19.Range("Sheet19.AckNum").Locked = False


    Else

         Sheet19.Range("Sheet19.AckNum").value = ""
         Sheet19.Range("Sheet19.AckNum").Locked = True
         Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HD8D8D8)
        
         Sheet19.Range("Sheet19.AckNum.10CCD").Locked = True
         Sheet19.Range("Sheet19.AckNum.10CCD").Interior.Color = (&HD8D8D8)
         Sheet19.Range("Sheet19.AckNum.10CCD").ClearContents

         Sheet19.Range("Sheet19.AckNum.10CCE").Locked = True
         Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HD8D8D8)
         Sheet19.Range("Sheet19.AckNum.10CCE").ClearContents


    End If
    
     If (SheetALL.Range("bacValue").value = 1) Or (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then
     
'     If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I") And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Then

'Malli
'        Sheet19.Range("SELECT80DD").MergeArea.ClearContents
'        Sheet19.Range("SELECT80DD").MergeArea.Locked = True
'        Sheet19.Range("SELECT80DD").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("SELECT80DD").value = ""
        
        '------------------------------------

        Sheet19.Range("SELECT80DDB").MergeArea.ClearContents
        Sheet19.Range("SELECT80DDB").MergeArea.Locked = True
        Sheet19.Range("SELECT80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("SELECT80DDB").value = ""
        
 'sudha_2025-26
 
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.ClearContents
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Locked = True
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("Sheet19.Specified_Disease").value = ""
        
        
'Malli
'        Sheet19.Range("scvia.Section80DD").MergeArea.ClearContents
'        Sheet19.Range("scvia.Section80DD").MergeArea.Locked = True
'        Sheet19.Range("scvia.Section80DD").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("scvia.Section80DD").value = ""
'----------------------------

        Sheet19.Range("scvia.Section80DDB").MergeArea.ClearContents
        Sheet19.Range("scvia.Section80DDB").MergeArea.Locked = True
        Sheet19.Range("scvia.Section80DDB").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet19.Range("scvia.Section80DDB").value = ""
'Malli
'        Sheet19.Range("SELECT80U").MergeArea.ClearContents
'        Sheet19.Range("SELECT80U").MergeArea.Locked = True
'        Sheet19.Range("SELECT80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("SELECT80U").value = ""

'        Sheet19.Range("scvia.Section80U").MergeArea.ClearContents
'        Sheet19.Range("scvia.Section80U").MergeArea.Locked = True
'        Sheet19.Range("scvia.Section80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("scvia.Section80U").value = ""

'         Sheet19.Range("Sheet19.AckNum.10CCD").Locked = True
'         Sheet19.Range("Sheet19.AckNum.10CCD").Interior.Color = (&HD8D8D8)
'         Sheet19.Range("Sheet19.AckNum.10CCD").ClearContents
'
'         Sheet19.Range("Sheet19.AckNum.10CCE").Locked = True
'         Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HD8D8D8)
'         Sheet19.Range("Sheet19.AckNum.10CCE").ClearContents
         
'--------------------
        ElseIf Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
'Malli
'        Sheet19.Range("SELECT80U").MergeArea.ClearContents
'        Sheet19.Range("SELECT80U").MergeArea.Locked = True
'        Sheet19.Range("SELECT80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("SELECT80U").value = ""

'        Sheet19.Range("scvia.Section80U").MergeArea.ClearContents
'        Sheet19.Range("scvia.Section80U").MergeArea.Locked = True
'        Sheet19.Range("scvia.Section80U").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet19.Range("scvia.Section80U").value = ""

'--------------
         Else
         
'Malli

'        Sheet19.Range("SELECT80DD").MergeArea.ClearContents
'        Sheet19.Range("SELECT80DD").MergeArea.Locked = False
'        Sheet19.Range("SELECT80DD").MergeArea.Interior.Color = "&HCCFFCC"
'----------------------
        Sheet19.Range("SELECT80DDB").MergeArea.ClearContents
        Sheet19.Range("SELECT80DDB").MergeArea.Locked = False
        Sheet19.Range("SELECT80DDB").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.ClearContents
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Locked = False
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Interior.Color = "&HCCFFCC"
        

'         Sheet19.Range("SELECT80U").MergeArea.ClearContents
'        Sheet19.Range("SELECT80U").MergeArea.Locked = False
'        Sheet19.Range("SELECT80U").MergeArea.Interior.Color = "&HCCFFCC"
'Malli
'        Sheet19.Range("scvia.Section80DD").MergeArea.ClearContents
'        Sheet19.Range("scvia.Section80DD").MergeArea.Locked = False
'        Sheet19.Range("scvia.Section80DD").MergeArea.Interior.Color = "&HCCFFCC"
'----------------------

        Sheet19.Range("scvia.Section80DDB").MergeArea.ClearContents
        Sheet19.Range("scvia.Section80DDB").MergeArea.Locked = False
        Sheet19.Range("scvia.Section80DDB").MergeArea.Interior.Color = "&HCCFFCC"
'Malli
'        Sheet19.Range("scvia.Section80U").MergeArea.ClearContents
'        Sheet19.Range("scvia.Section80U").MergeArea.Locked = False
'        Sheet19.Range("scvia.Section80U").MergeArea.Interior.Color = "&HCCFFCC"

'---------------
    End If


'If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" And ((Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "RE") Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NO")) Then
If (SheetALL.Range("bacValue").value = 2) And (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "H" And ((Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "RE") Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) = "NO")) Then
        'Malli
'
'        Sheet19.Range("SELECT80DD").MergeArea.ClearContents
'        Sheet19.Range("SELECT80DD").MergeArea.Locked = False
'        Sheet19.Range("SELECT80DD").MergeArea.Interior.Color = "&HCCFFCC"
'-----------------------
        Sheet19.Range("SELECT80DDB").MergeArea.ClearContents
        Sheet19.Range("SELECT80DDB").MergeArea.Locked = False
        Sheet19.Range("SELECT80DDB").MergeArea.Interior.Color = "&HCCFFCC"
        
 'sudha_2025-26
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.ClearContents
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Locked = False
        Sheet19.Range("Sheet19.Specified_Disease").MergeArea.Interior.Color = "&HCCFFCC"
        
        

'        Sheet19.Range("scvia.Section80DD").MergeArea.ClearContents
'        Sheet19.Range("scvia.Section80DD").MergeArea.Locked = False
'        Sheet19.Range("scvia.Section80DD").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet19.Range("scvia.Section80DDB").MergeArea.ClearContents
        Sheet19.Range("scvia.Section80DDB").MergeArea.Locked = False
        Sheet19.Range("scvia.Section80DDB").MergeArea.Interior.Color = "&HCCFFCC"
        
'         Sheet19.Range("Sheet19.AckNum.10CCD").Locked = False
'         Sheet19.Range("Sheet19.AckNum.10CCD").Interior.Color = (&HCCFFCC)
'
'        Sheet19.Range("Sheet19.AckNum.10CCE").Locked = False
'        Sheet19.Range("Sheet19.AckNum.10CCE").Interior.Color = (&HCCFFCC)
        
End If
Application.EnableEvents = False
Sheet19.Protect Password:=getmsgstate
End Sub

'Sub LockUnlockAadhaar()
'On Error Resume Next
'Application.EnableEvents = False
'
'Sheet1.Unprotect Password:=getmsgstate
'    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
'        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Locked = True
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("sheet1.adhaarno").value = ""
'
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = True
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("sheet1.adhaarenrol").value = ""
'    Else
'        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Locked = False
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Interior.Color = "&HCCFFCC"
'        'Sheet1.Range("sheet1.adhaarno").value = ""
'
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = False
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = "&HCCFFCC"
'        'Sheet1.Range("sheet1.adhaarenrol").value = ""
'
'    End If
'    Application.EnableEvents = True
'Dim validateType As String
'Sheet1.Protect Password:=getmsgstate
'End Sub

'Newly modified by sai on 26/09/2024

Sub LockUnlockAadhaar()
On Error Resume Next
Application.EnableEvents = False
 
Sheet1.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Then
        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
        Sheet1.Range("sheet1.adhaarno").MergeArea.Locked = True
        Sheet1.Range("sheet1.adhaarno").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.adhaarno").value = ""
        
 'Commented by Aavula
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = True
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("sheet1.adhaarenrol").value = ""
    Else
        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
        Sheet1.Range("sheet1.adhaarno").MergeArea.Locked = False
        Sheet1.Range("sheet1.adhaarno").MergeArea.Interior.Color = "&HCCFFCC"
        'Sheet1.Range("sheet1.adhaarno").value = ""

  'Commented by Aavula
'       If Not Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = True Then
'
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = False
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = "&HCCFFCC"
'        'Sheet1.Range("sheet1.adhaarenrol").value = ""
'        End If

    End If
    Application.EnableEvents = True
Dim validateType As String
Sheet1.Protect Password:=getmsgstate
End Sub
Sub LockUnlockClaim115H()
On Error Resume Next
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 2) <> "NR") Then
        Sheet1.Range("Sheet1.115H").MergeArea.ClearContents
        Sheet1.Range("sheet1.115H").MergeArea.Locked = False
        Sheet1.Range("sheet1.115H").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.115H").value = "No"
    Else
        Sheet1.Range("sheet1.115H").MergeArea.ClearContents
        Sheet1.Range("sheet1.115H").MergeArea.Locked = True
        Sheet1.Range("sheet1.115H").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.115H").value = ""

        LockUnlockPartBTTI
        If Trim(Mid(Sheet1.Range("sheet1.ResidentialStatus1"), 1, 2)) <> "NR" Then
            LOCKCG
        Else
            UNLOCKcg

        End If

        'Application.EnableEvents = False 'to keep things working; called multiple times in change event
    End If
'    Application.EnableEvents = True
Sheet1.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Activate()

On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet1.Shapes("Button 401").Left = Range("SchButton.act").Left + 2
Sheet1.Shapes("Button 401").Top = Range("SchButton.act").Top + 2

Sheet1.Shapes("Button 403").Left = Range("Resident_Button").Left + 3
Sheet1.Shapes("Button 403").Top = Range("Resident_Button").Top + 3


Sheet1.Shapes("Button 434").Left = Range("Button_FirmTable").Left + 3
Sheet1.Shapes("Button 434").Top = Range("Button_FirmTable").Top + 3

Sheet1.Protect Password:=sPassword

''Newly added by sai on 06/10/2024====================================================
'     Application.EnableEvents = False
'
'    Dim formula As Variant
'    Dim cellrange As Range
'    Set cellrange = Sheet1.Range("sheet1.ReturnFileSec")
'
'    If Dformat(Sheet52.Range("sheet9.Date").value, "DD/MM/YYYY") > Dformat("31/12/2024", "DD/MM/YYYY") Then
''        formula = "=Returnfiledundersection_new"
'        formula = "=ReturnFileUnderSection1"
'
'      Else
''        formula = "=ReturnFileUnderSection1"
'        formula = "=Returnfiledundersection_new"
'        Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
'    End If
'
'    sPassword = EfilingCommon.getmsgstate
'    Sheet1.Unprotect Password:=sPassword
'
'    With cellrange.Validation
'        .Delete
'        .add Type:=xlValidateList, _
'         AlertStyle:=xlValidAlertStop, _
'         Operator:=xlBetween, _
'         Formula1:=formula
'    End With
'
'    Sheet1.Protect Password:=sPassword
'    Sheet1.Activate
''==============================================================================================
'
'Application.EnableEvents = True

End Sub


Private Sub Worksheet_BeforeDelete()

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub
Sub WClockunlockProviso()

   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg").value, 1, 1) = "N") Or Range("clauseiv7provisio139iFlg").value = "(Select)" Then

        
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)

ElseIf (Mid(Range("clauseiv7provisio139iFlg").value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_1").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_2").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_3").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_4").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)

End If

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)

ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = False
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
'        Range("Sheet1.DepositAmountFlag").value = "(Select)"

        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iFlg").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HCCFFCC)
        'Range("clauseiv7provisio139iFlg").value = "(Select)"

End If

End Sub
Sub WClockunlockone()

Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("sheet1.Status").value = "H-HUF" Then
    
        Sheet4.Unprotect Password:=getmsgstate
        
'        Sheet4.Range("PL_62i").value = ""
'        Sheet4.Range("PL_62i").Interior.Color = (&HD8D8D8) ', "AU"
'        Sheet4.Range("PL_62i").MergeArea.Locked = True
'
        Sheet4.Range("PL_62ii").value = ""
        Sheet4.Range("PL_62ii").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("PL_62ii").MergeArea.Locked = True
        Sheet4.Range("Sec44AD.NOBName1").value = ""
        Sheet4.Range("Sec44AD.NOBName1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("Sec44AD.NOBName1").Locked = True
        Sheet4.Range("Sec44AD.NOBCode1").value = ""
        Sheet4.Range("Sec44AD.NOBCode1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("Sec44AD.NOBCode1").Locked = True
        Sheet4.Range("Sec44AD.NOBDescription1").value = ""
        Sheet4.Range("Sec44AD.NOBDescription1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("Sec44AD.NOBDescription1").Locked = True
        Sheet4.Protect Password:=getmsgstate
        Sheet9.Unprotect Password:=getmsgstate
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").value = ""
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Locked = True
        Sheet9.Protect Password:=getmsgstate
        Sheet4.Protect Password:=getmsgstate
        
 'Commented by sudha_2025_26
        
'        Sheet19.Unprotect Password:=getmsgstate
'        Sheet19.Range("scvia.Section80E").value = ""
'        Sheet19.Range("scvia.Section80E").Interior.Color = (&HD8D8D8) ', "AU"
'        Sheet19.Range("scvia.Section80E").Locked = True
        
'        Sheet19.Range("scvia.Section80EE").value = ""
'        Sheet19.Range("scvia.Section80EE").Interior.Color = (&HD8D8D8) ', "AU"
'        Sheet19.Range("scvia.Section80EE").Locked = True
'        Sheet19.Protect Password:=getmsgstate
        
        LockUnlockESOP
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUS").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUS").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUS").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUK").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUK").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUK").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountCan").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountCan").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountCan").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeTaxablePrevYear89A").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeTaxablePrevYear89A").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeTaxablePrevYear89A").MergeArea.Locked = True
        Sheet14.Protect Password:=getmsgstate
        


    Else

        Sheet4.Unprotect Password:=getmsgstate
'
'        Sheet4.Range("PL_62i").value = ""
'        Sheet4.Range("PL_62i").Interior.Color = (&HCCFFCC)  ', "AU"
'        Sheet4.Range("PL_62i").MergeArea.Locked = False
        
        Sheet4.Range("PL_62ii").value = ""
        Sheet4.Range("PL_62ii").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("PL_62ii").MergeArea.Locked = False
        Sheet4.Range("Sec44AD.NOBName1").value = ""
        Sheet4.Range("Sec44AD.NOBName1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("Sec44AD.NOBName1").Locked = False
         Sheet4.Range("Sec44AD.NOBCode1").value = ""
        Sheet4.Range("Sec44AD.NOBCode1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("Sec44AD.NOBCode1").Locked = False
         Sheet4.Range("Sec44AD.NOBDescription1").value = ""
        Sheet4.Range("Sec44AD.NOBDescription1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("Sec44AD.NOBDescription1").Locked = False
        Sheet4.Protect Password:=getmsgstate
        Sheet9.Unprotect Password:=getmsgstate
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").value = ""
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Locked = False
        Sheet9.Protect Password:=getmsgstate
        
        'Commented by sudha_2025-26
'        Sheet19.Unprotect Password:=getmsgstate
'        Sheet19.Range("scvia.Section80E").value = ""
'        Sheet19.Range("scvia.Section80E").Interior.Color = (&HCCFFCC) ', "AU"
'        Sheet19.Range("scvia.Section80E").Locked = False
'        Sheet19.Range("scvia.Section80EE").value = ""
'        Sheet19.Range("scvia.Section80EE").Interior.Color = (&HCCFFCC) ', "AU"
'        Sheet19.Range("scvia.Section80EE").Locked = False
'        Sheet19.Protect Password:=getmsgstate
        
        LockUnlockESOP
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncreliefus89A").Locked = False
        Sheet14.Range("OSIncomeNotified89A_AmountUS").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeNotified89A_AmountUS").Locked = False
        Sheet14.Range("OSIncomeNotified89A_AmountUK").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeNotified89A_AmountUK").Locked = False
        Sheet14.Range("OSIncomeNotified89A_AmountCan").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeNotified89A_AmountCan").Locked = False
        Sheet14.Range("OSIncomeTaxablePrevYear89A").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeTaxablePrevYear89A").Locked = False
        Sheet14.Protect Password:=getmsgstate
    End If



End Sub
Sub PartAGen_sub1(ByVal Target As Range)

Dim validateType As String

'If Not Application.Intersect(Sheet1.Range("Opt_Out115BAC"), Sheet1.Range(Target.Address)) Is Nothing Then
'    CurrentYear10IE_23_24 Target
'End If

'If Not Application.Intersect(Sheet1.Range("OptingNewTaxRegime"), Sheet1.Range(Target.Address)) Is Nothing Then
'    CurrentYear10IEDateAckAY_23_24 Target
'End If

    'Newly modified by Sai For AY 2024-25
'       If Not Application.Intersect(Sheet1.Range("Sheet1.OptOutNewTaxRegime"), Sheet1.Range(target.Address)) Is Nothing Then
'          CurrentYear10IEDateAckAY_24_25 target
'       End If




'If Not Application.Intersect(Sheet1.Range("Sheet1.date10IE_Opt"), Sheet1.Range(Target.Address)) Is Nothing Then
'    ID_Opt_Date10IE Target
'End If
'
'If Not Application.Intersect(Sheet1.Range("Sheet1.date10IE_OptOut"), Sheet1.Range(Target.Address)) Is Nothing Then
'    ID_OptOut_Date10IE Target
'End If
'
'If Not Application.Intersect(Sheet1.Range("sheet1.capacityassessee"), Sheet1.Range(Target.Address)) Is Nothing Then
'    ID_Date10IE Target
'End If
'end---

If Not Application.Intersect(Sheet1.Range("NRI_Juridiction"), Sheet1.Range(Target.Address)) Is Nothing Then

If 1 < Application.CountIf(Range("NRI_Juridiction"), Target.value) Then
             fmsgboxStatus "* Same dropdown cannot be selected twice "
             Range(Target.Address).value = "(Select)"
        End If
End If

'msgbox"d1"
If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then
WClockunlockone
End If

If Not Application.Intersect(Sheet1.Range("Sheet1.Foreign_exchange"), Sheet1.Range(Target.Address)) Is Nothing Then

   Sheet22.Unprotect Password:=getmsgstate
If Range("Sheet1.Foreign_exchange").value = "No" Then
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").value = 0
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").Interior.Color = (&HD8D8D8)
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").Locked = True
Else
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").MergeArea.Locked = False
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").MergeArea.Interior.Color = (&HCCFFCC)

End If
Sheet1.Unprotect Password:=getmsgstate
Sheet22.Protect Password:=getmsgstate
End If
'dpk
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg"), Sheet1.Range(Target.Address)) Is Nothing Then
    WClockunlockProviso
End If

'Change-20.04.2023.101.PL -> SIT-48882
If (Not Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing) Or (Not Intersect(Sheet1.Range("sheet1.ResidentialStatus1"), Sheet1.Range(Target.Address)) Is Nothing) Then

    If Sheet1.Range("sheet1.Status").value = "H-HUF" Then
        LOCKPL3
    Else
        If Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL" And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
            LOCKPL3
        Else
            UNLOCKPL3
        End If
    End If
    
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_1"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_2"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_3"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

'If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_3"), Sheet1.Range(Target.Address)) Is Nothing Then
'   Sheet1.Unprotect Password:=getmsgstate
'If (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "N") Then
'        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
'        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
'        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
'ElseIf (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 2) = "(S") Then
'        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
'        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)
'End If
'Sheet1.Protect Password:=getmsgstate
'End If


Dim i As Long

'Commented by Aavula

'If Not Application.Intersect(Sheet1.Range("sheet1.adhaarenrol"), Sheet1.Range(Target.Address)) Is Nothing Then
'   Sheet1.Unprotect Password:=getmsgstate
'   If (Not IsNumeric(Target.value)) And ((Target.value <> "")) Then
'    MsgBox "Please enter the valid Aadhaar Enrolment ID in schedule Part A General."
'    Target.value = ""
''Change 01, Satya, 18.10.2022
'   Else
'
'        For i = 1 To Len(Target.value)
'            If Not Mid$(Target.value, i, 1) Like "[0-9]" Then
''                Validated = Validated & Mid$(Target.value, i, 1)
''            Else
'                MsgBox "Please enter the valid Aadhaar Enrolment ID in schedule Part A General."
'                Target.value = ""
'                Exit For
'            End If
'        Next
''End Change----
'   End If
'    Sheet1.Protect Password:=getmsgstate
'End If


'Change-02.11.2023.102.Ver.0.5
'Commented by Aavula
'If Not Application.Intersect(Sheet1.Range("sheet1.adhaarno"), Sheet1.Range(Target.Address)) Is Nothing Then
'    If Sheet1.Range("sheet1.adhaarno").Locked = False Then
'
''        If Len(Sheet1.Range("sheet1.adhaarno").value) > 0 Then
'      If Len(Sheet1.Range("sheet1.adhaarno").value) > 0 Or Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") >= Dformat(Trim("01/10/2024"), "yyyy-mm-dd") Then
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = True
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = (&HD8D8D8)
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
'        Else
'
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = False
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = (&HCCFFCC)
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
'      End If
'    End If
'
'End If

'Newly added by sai on 26/09/2024
'Commented by Aavula
 
'If Not Application.Intersect(Sheet1.Range("sheet1.adhaarno"), Sheet1.Range(Target.Address)) Is Nothing Or _
'   Not Application.Intersect(Sheet1.Range("sheet1.adhaarenrol"), Sheet1.Range(Target.Address)) Is Nothing Or _
'   Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'
''    If Sheet1.Range("sheet1.adhaarno").Locked = False Then
''
'    If Len(Sheet1.Range("sheet1.adhaarno").value) > 0 Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") > Dformat(Trim("30/09/2024"), "yyyy-mm-dd") Then
'
'        '----------------------------------------
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = True
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = (&HD8D8D8)
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
'        Else
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Locked = False
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.Interior.Color = (&HCCFFCC)
''        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
''        End If
'    End If
'End If



'=============================================================================================


'End Change
'Change-02.11.2023.102.Ver.0.5

'Commented by Aavula

'If Not Application.Intersect(Sheet1.Range("sheet1.adhaarenrol"), Sheet1.Range(Target.Address)) Is Nothing Then
'If Sheet1.Range("sheet1.adhaarenrol").Locked = False Then
'        If Len(Sheet1.Range("sheet1.adhaarenrol").value) > 0 Then
'         Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Locked = True
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Interior.Color = (&HD8D8D8)
'        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
'        Else
'
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Locked = False
'        Sheet1.Range("sheet1.adhaarno").MergeArea.Interior.Color = (&HCCFFCC)
'        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
'        End If
'    End If
'End If
'End Change
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_4"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If


End Sub

Sub PartAGen_sub2(ByVal Target As Range)

Dim validateType As String
Dim i
'Change-19.12.2022.102.07A

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_1"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_1").value < 6000000 Or Sheet1.Range("clauseiv7provisio139iAmount_1").value = 6000000 Or Sheet1.Range("clauseiv7provisio139iAmount_1").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if ""total sales, turnover or gross receipts, as the case may be, of the person in the business is less than or equal to sixty lakh rupees during the previous year " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_1").value = ""
        End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_2"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_2").value < 1000000 Or Sheet1.Range("clauseiv7provisio139iAmount_2").value = 1000000 Or Sheet1.Range("clauseiv7provisio139iAmount_2").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year" & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_2").value = ""
        End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_3"), Sheet1.Range(Target.Address)) Is Nothing Then
        age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").value)
       If age >= 60 Then
            If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 50000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
            fmsgbox "* Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_3").value = ""
            End If
        Else
            If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 25000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
            fmsgbox "* Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_3").value = ""
            End If
       End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_4"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_4").value < 5000000 Or Sheet1.Range("clauseiv7provisio139iAmount_4").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_4").value = ""
        End If
End If

'Commented by Sai For AY 2023-24
'If Not Application.Intersect(Sheet1.Range("OptingNewTaxRegime"), Sheet1.Range(Target.Address)) Is Nothing Then
'     If (Target.value = "Continue to opt") Or (Target.value = "Opting in now") Then
     
  'Newly Modified by Sai For AY 2024-25
  
  
  'Commented on 10/02/2025 by Aavula Naresh
  
'  If Not Application.Intersect(Sheet1.Range("Sheet1.OptOutNewTaxRegime"), Sheet1.Range(target.Address)) Is Nothing Then
'       If (target.value = "No") Or (target.value = "Yes-but beyond due date") Then
'         NewBacYesValueChange
'
'           Else
'
'         NewBacNoValueChange
'
'    End If
'  End If



  'Commented by Aavula For AY 2026-27
  
'  'Newly Added by Aavula Naresh on 10/02/2025
'    If Not Application.Intersect(Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime"), Sheet1.Range(target.Address)) Is Nothing Then
'      OptOutNewTaxregime_25_26
'    End If


   '
'    If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime")) Is Nothing Then
'      If Sheet1.Range("Sheet1.MethodofOptOutNewTaxRegime").value = "by exercising the option in the return of income only (form 10IEA is not applicable)" Then
'           fmsgboxStatus "*Option of exercising the Tax regime option in the return of Income is available only for the cases where form 10IEA is not applicable. (i.e., for Taxpayer having no business income)"
'        target.MergeArea.Select
'      End If
'    End If
  '===========================================''''
  

  
'  'Newly Added by Aavula Naresh on 06/02/2025
'
'    If Not Application.Intersect(Sheet1.Range("sheet1.115BAC"), Sheet1.Range(target.Address)) Is Nothing Then
'        Taxregime_25_26
'    End If
'
'    If Not Application.Intersect(Sheet1.Range("sheet1.OptOutNewTaxRegime_New_Yes"), Sheet1.Range(target.Address)) Is Nothing Then
'        Taxregime_25_26_A
'    End If
'
'    If Not Application.Intersect(Sheet1.Range("sheet1.OptOutNewTaxRegime_New_No"), Sheet1.Range(target.Address)) Is Nothing Then
'        Taxregime_25_26_B
'    End If
'
'    If Not Application.Intersect(Sheet1.Range("sheet1.OptOutNewTaxRegime_New_NA"), Sheet1.Range(target.Address)) Is Nothing Then
'        Taxregime_25_26_C
'    End If




'ID_Opt_Date10IEA_New target
'
'ID_Opt_Date10IEA_New_Ai target
'ID_Opt_Date10IEA_New_Bi target
'ID_Opt_Date10IEA_New_Ci target
'
'''===================================================================''''''''''
'
'
'  'Newly Added by Aavula Naresh on 10/02/2025
'
'  If Not Application.Intersect(Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB"), Sheet1.Range(target.Address)) Is Nothing Then
'       If (target.value = "No") Then
'         NewTaxRegime_New
'           Else
'         OldTaxRegime_New
'       End If
'  End If
'''============================================================''''''''''
  ''''''''''''''''''''''''''''''''''''' End''''''''''''''''''''''''''''''''''''''
  
  
   If Not Application.Intersect(Sheet1.Range("sheet1.Form10IEAoldtax"), Sheet1.Range(Target.Address)) Is Nothing Then
        Form10IEAoldtax_26_27
    End If
    
   If Not Application.Intersect(Sheet1.Range("sheet1.Form10IEAoldtax_IA"), Sheet1.Range(Target.Address)) Is Nothing Then
        Form10IEAoldtax_IA_26_27
    End If
   
   If Not Application.Intersect(Sheet1.Range("Sheet1.Form10IEANewTax"), Sheet1.Range(Target.Address)) Is Nothing Then
        Form10IEANewtax_26_27
    End If
    
   If Not Application.Intersect(Sheet1.Range("Sheet1.Form10IEANewTaxIB"), Sheet1.Range(Target.Address)) Is Nothing Then
        Form10IEANewTaxIB_26_27
    End If
  
  
   DateOfFilingform_bi Target
   
   DateOfFilingform_Ibi Target
   
   
If Not Application.Intersect(Sheet1.Range("Sheet1.Form10IEAB"), Sheet1.Range(Target.Address)) Is Nothing Then
       If (Target.value = "No") Then
         NewTaxRegime_New
           Else
         OldTaxRegime_New
       End If
  End If
   
   'Added by Aavula for AY 2026-27 06/04/2026
   
  If Not Application.Intersect(Sheet1.Range("Sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then
         LockUnlock80E_80EE
  End If
  
''''''''''
  

If Not Application.Intersect(Sheet1.Range("sheet1.ReturnFileSec"), Sheet1.Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
Sheet52.Unprotect Password:=getmsgstate
    LockUnlock1398A_Gen
    If Sheet1.Range(Target.Address).value = "139(4)- After due date" Then
        Dim vdate As Variant
        vdate = Sheet52.Range("sheet9.Date").value
    'Change-19.12.2022.102.04A
'    If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2022-08-01") Then
'        fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2022 or extended due date"
'     If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2023-08-01") Then
'        fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2023 or extended due date"

     'Newly modified by Sai For AY 2024-25
'      If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2024-08-01") Then
'        fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2024 or extended due date"

'For AY 2025-26_Sadineni_22_09_2025
    If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), Dformat(Sheet33.Range("NonAudit_DueDate_1").value, "yyyy-mm-dd")) Then
'        fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2025 or extended due date"
'        fmsgbox "Filing section u/s. 139(4) cannot be selected till 15th Sep, 2025"
'         fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2025 or extended due date"
         fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st August, 2026"
    'end---
        Range("sheet1.ReturnFileSec").value = "139(1)- On or Before due date"
    End If
    
    End If
    
 '====================================================================================================
'For AY 2025-26_Konda_19_12_2025
If Sheet1.Range(Target.Address).value = "139(8A)" Then
        Dim vdate1 As Variant
        vdate1 = Sheet52.Range("sheet9.Date").value

    If Not ChkMinInclusiveDate(Dformat(vdate1, "yyyy-mm-dd"), Dformat(Sheet33.Range("Duedate_1398A").value, "yyyy-mm-dd")) Then
'          fmsgbox "Filing section u/s. 139(8) cannot be selected till  31st Dec, 2025 or extended due date"
           fmsgbox "Filing section u/s. 139(8) cannot be selected till  31st Dec, 2026"
         Range("sheet1.ReturnFileSec").value = "139(1)- On or Before due date"
    End If
End If
'===========================================================================================================
    
    
    
    Sheet1.Protect Password:=getmsgstate
    Sheet52.Protect Password:=getmsgstate
End If

'Newly added by Sai on 14/02/2024 For AY 2024-25================================================================
  If Not Application.Intersect(Sheet1.Range("sheet1.AuditYN"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Sheet1.Range(Target.Address).value = "Yes" Then
            Sheet1.Unprotect Password:=getmsgstate
            
            Sheet1.Range(Replace(Target.Address, "N", "O")).value = ""
            Sheet1.Range(Replace(Target.Address, "N", "O")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(Target.Address, "N", "O")).MergeArea.Locked = False
            
            
            Sheet1.Range(Replace(Target.Address, "N", "P")).value = ""
            Sheet1.Range(Replace(Target.Address, "N", "P")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(Target.Address, "N", "P")).MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            
            Sheet1.Range(Replace(Target.Address, "N", "O")).value = ""
            Sheet1.Range(Replace(Target.Address, "N", "O")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(Target.Address, "N", "O")).MergeArea.Locked = True
            
            Sheet1.Range(Replace(Target.Address, "N", "P")).value = ""
            Sheet1.Range(Replace(Target.Address, "N", "P")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(Target.Address, "N", "P")).MergeArea.Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If

'=========================================================================================================
If Not Application.Intersect(Sheet1.Range("Audit.Act"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Sheet1.Range(Target.Address).value = "Others" Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(Target.Address, "F", "O")).value = ""
            Sheet1.Range(Replace(Target.Address, "F", "O")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(Target.Address, "F", "O")).MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(Target.Address, "F", "O")).value = ""
            Sheet1.Range(Replace(Target.Address, "F", "O")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(Target.Address, "F", "O")).MergeArea.Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If
If Not Intersect(Range("NRI_Juridiction"), Range(Target.Address)) Is Nothing Then

   Sheet1.Unprotect Password:=getmsgstate
If (Range(Target.Address).value) = "9998-Not Applicable (Not Resident in any Country)" Then
          fmsgboxStatus "*In field Tax Identification Number,""Please enter Passport No. along with the name of Country in which Passport is issued""" & Chr(13)
End If
End If
If Not Intersect(Range("Sheet1.CompPAN"), Range(Target.Address)) Is Nothing Then
         Target.value = UCase(Target.value)
End If
 If Not Intersect(Range("Sheet1.FirmPAN"), Range(Target.Address)) Is Nothing Then
         Target.value = UCase(Target.value)
End If
If Not Intersect(Range("sheet1.Status"), Range(Target.Address)) Is Nothing Then
            
'            For i = 1 To Sheet24.Range("EI.NatureOfIncome").count
'                If Mid(Sheet1.Range(Target.name.name).value, 1, 1) = "H" Or Mid(Sheet1.Range(Target.name.name).value, 1, 1) = "(" Then
'                If Mid(Sheet24.Range("EI.NatureOfIncome").Cells.item(i).value, 1, 11) = "Sec 10(12A)" Or Mid(Sheet24.Range("EI.NatureOfIncome").Cells.item(i).value, 1, 11) = "Sec 10(12B)" Then
'                Sheet24.Range("EI.NatureOfIncome").Cells.item(i).value = "(Select)"
'                End If
'                End If
'            Next
        Application.EnableEvents = False
                        StatusBasedDefaultValue
                        LockUnlock80D
                        LockUnlockQVIA
            LockUnlockAadhaar
            LockUnlockClaim115H
            HideUnhideCheck
            LockUnlockSalary
            LockUnlockTDS
              NewResidentialwork
            If Mid(Sheet1.Range(Target.name.name).value, 1, 1) = "H" Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet9.Unprotect Password:=getmsgstate
            
'                Sheet1.Range("sheet1.PortugeseCC5A").value = ""
                Sheet1.Unprotect Password:=getmsgstate
                Sheet1.Range("sheet1.PortugeseCC5A").Interior.Color = (&HD8D8D8)
                Sheet1.Range("sheet1.PortugeseCC5A").MergeArea.Locked = True
                
                
                LOCK5A
                LOCK80EE_A_B
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").value = ""
                Sheet9.Unprotect Password:=getmsgstate
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HD8D8D8)
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").MergeArea.Locked = True

                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").value = ""
                Sheet9.Unprotect Password:=getmsgstate
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").Interior.Color = (&HD8D8D8)
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").MergeArea.Locked = True
            Sheet1.Protect Password:=getmsgstate
            Sheet9.Protect Password:=getmsgstate
            Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet9.Unprotect Password:=getmsgstate
                Sheet1.Range("sheet1.PortugeseCC5A").value = "No"
                Sheet1.Unprotect Password:=getmsgstate
                Sheet1.Range("sheet1.PortugeseCC5A").Interior.Color = (&HCCFFCC)
                Sheet1.Range("sheet1.PortugeseCC5A").MergeArea.Locked = False
                LOCK5A
                UNLOCK80EE_A_B

                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").value = ""
                Sheet9.Unprotect Password:=getmsgstate
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HCCFFCC)
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").MergeArea.Locked = False

                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").value = ""
                Sheet9.Unprotect Password:=getmsgstate
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").Interior.Color = (&HCCFFCC)
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate
            Sheet9.Protect Password:=getmsgstate
            End If
            If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "I" Then
                Sheet52.Range("sheet9.PAN").value = ""
            Else
                Sheet52.Range("sheet9.PAN").value = Sheet1.Range("sheet1.PAN").value
            End If
             Application.EnableEvents = True
        End If
If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then

    
'            LockUnlock_FIIFPI
            LockUnlock115AD
            LockUnlockSalary89A
End If

If Not Intersect(Range("IsDirectorCompany"), Range(Target.Address)) Is Nothing Then

            LockUnlockCheck
End If
If Not Intersect(Range("IsPartnerFirm"), Range(Target.Address)) Is Nothing Then
            LockUnlockFirm
End If
If Not Intersect(Range("IsShareOwner"), Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Trim(Sheet1.Range("IsShareOwner").value), 1, 1)) = "Y" Then
        UnLockShares
    Else
        LockShares
    End If
End If


       If Not Intersect(Range("sheet1.FirstName"), Range(Target.Address)) Is Nothing Then
           Sheet1.Range("sheet1.FirstName").value = UCase(Sheet1.Range("sheet1.FirstName").value)
       End If
       If Not Intersect(Range("sheet1.MiddleName"), Range(Target.Address)) Is Nothing Then
           Sheet1.Range("sheet1.MiddleName").value = UCase(Sheet1.Range("sheet1.MiddleName").value)
       End If
       If Not Intersect(Range("sheet1.SurNameOrOrgName"), Range(Target.Address)) Is Nothing Then
           Sheet1.Range("sheet1.SurNameOrOrgName").value = UCase(Sheet1.Range("sheet1.SurNameOrOrgName").value)
       End If

        If Not Intersect(Range("sheet1.PAN"), Range(Target.Address)) Is Nothing Then
        
               Dim TargetPAN As String
               TargetPAN = ""
               TargetPAN = Target.name.name
                Sheet1.Range(TargetPAN).value = UCase(Range(TargetPAN).value)
                   If Not Part_A_General.CheckPAN(UCase(Trim(Range(TargetPAN).value))) Then
                        fmsgboxStatus " Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."
                        Sheet1.Range(TargetPAN).Select
                        Sheet1.Range(TargetPAN).value = ""
                    Else
                       If Mid(Range(TargetPAN).value, 4, 1) = "P" Then
                            Sheet52.Range("sheet9.PAN").value = UCase(Range(TargetPAN).value)
                       End If

                   End If
           End If
           
''           If Not Intersect(Range("sheet1.RepPAN"), Range(target.Address)) Is Nothing Then
''                   Dim TargetRepPAN As String
''                   TargetRepPAN = ""
''                   TargetRepPAN = target.name.name
''               If Not Part_A_General.CheckRepPAN(UCase(Trim(Range(TargetRepPAN).value))) Then
''                   fmsgboxStatus " Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule Part A General" ', vbOKCancel
''                   Sheet1.Range(TargetRepPAN).Select
''                   Sheet1.Range(TargetRepPAN).value = ""
''
''               End If
''               If Range(TargetRepPAN).value = Sheet1.Range("Sheet1.pan").value Then
'                If UCase(Range(TargetRepPAN).value) = UCase(Sheet1.Range("Sheet1.pan").value) Then 'Added by Aavula Naresh for AY 2025-26
'                    fmsgboxStatus "PAN of the representative cannot be same as assessee PAN in Part-A General Information." ', vbOKCancel
'                    Sheet1.Range(TargetRepPAN).Select
'                    Sheet1.Range(TargetRepPAN).value = ""
'               End If
'
'               Sheet1.Range(TargetRepPAN).value = UCase(Range(TargetRepPAN).value)
'               Sheet52.Range("sheet9.PAN").value = UCase(Range(TargetRepPAN).value)
'           End If


           If Not Intersect(Range("sheet1.PinCode"), Range(Target.Address)) Is Nothing Then
               If Not (Part_A_General.ValidatePinCode()) Then
                   Sheet1.Range("sheet1.PinCode").value = ""
                   Sheet1.Range("sheet1.PinCode").Select
               Else
                   If (Target.value) = "999999" Then
                   If Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
                       If Mid(Sheet1.Range("sheet1.StateCode1"), 1, 2) <> "99" Then
                           Sheet1.Range("sheet1.PinCode").value = ""
                           fmsgboxStatus " INVALID PINCODE"
                           Sheet1.Range("sheet1.PinCode").Select
                       End If
                   End If
                   End If
               End If
           End If
        If Not Intersect(Range("sheet1.capacityassessee"), Range(Target.Address)) Is Nothing Then
        
               If Trim(Range("sheet1.capacityassessee").value) <> "00/00/0000" Then
                   If Not CheckDate10IFBefore(Trim(Range("sheet1.capacityassessee").value)) Then

                       Sheet1.Range("sheet1.capacityassessee").Select
                       Sheet1.Range("sheet1.capacityassessee").value = ""
                   End If
               End If
           End If
           If Not Intersect(Range("sheet1.DOB"), Range(Target.Address)) Is Nothing Then
'               If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
               If Trim(Range("sheet1.DOB").value) <> "" Then 'Added by Aavula Naresh
                   If Not Part_A_General.CheckDOB(Trim(Range("sheet1.DOB").value)) Then
                       Sheet1.Range("sheet1.DOB").Select
                       Sheet1.Range("sheet1.DOB").value = ""
                   Else
                     Sheet1.Range("sheet1.DOB").value = AssesseeDob
                   End If
               End If
           End If

            If Not Intersect(Range("sheet1.DOC"), Range(Target.Address)) Is Nothing Then
'               If Trim(Range("sheet1.DOC").value) <> "00/00/0000" Then
               If Trim(Range("sheet1.DOC").value) <> "" Then
                   If Not Part_A_General.CheckDOC(Trim(Range("sheet1.DOC").value)) Then
                       Sheet1.Range("sheet1.DOC").Select
                       Sheet1.Range("sheet1.DOC").value = ""
                   End If
                   If Not ChkMinInclusiveDate(Dformat(Range("sheet1.DOC").value, ""), Dformat(Range("sheet1.DOB").value, "")) Then
                         MsgBox ("Date of formation of business field in Sheet : PART A-GENERAL  should be on or before Date of Commencement")
                         Range("sheet1.DOC").Select
                       Range("sheet1.DOC").value = ""
                       Range("sheet1.DOB").value = ""
                   Else
                   End If
               End If
           End If
          If Not Intersect(Range("sheet1.DOB"), Range(Target.Address)) Is Nothing Then
'               If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
               If Trim(Range("sheet1.DOB").value) <> "" Then
                    If Not ChkMinInclusiveDate(Dformat(Range("sheet1.DOC").value, ""), Dformat(Range("sheet1.DOB").value, "")) Then
                         MsgBox ("Date of formation of business field in Sheet : PART A-GENERAL  should be on or before Date of Commencement")
                         Range("sheet1.DOC").Select
                       Range("sheet1.DOC").value = ""
                       Range("sheet1.DOB").value = ""
                     Else
                    End If
               End If
           End If
           If Not Intersect(Range("sheet1.EmailAddress"), Range(Target.Address)) Is Nothing Then
           
               If Not Part_A_General.CheckEmailAddress(Trim(Range("sheet1.EmailAddress").value), validateType) Then

                   If validateType = "1" Then fmsgboxStatus "*Email ID in Sheet PART A-GENERAL , contains special characters"
                   If validateType = "2" Then fmsgboxStatus "*EmailAddress-1 in Sheet : PART A-GENERAL  should not exceed 125 characters " ', vbOKOnly
                   If validateType = "3" Then fmsgboxStatus "*EmailAddress-1 in Sheet : PART A-GENERAL  is invalid" ', vbOKOnly

                   Range("sheet1.EmailAddress").Select
                   Range("sheet1.EmailAddress").value = ""
               End If

           End If
           
           If Not Intersect(Range("sheet1.EmailAddressSec"), Range(Target.Address)) Is Nothing Then
               If Not Part_A_General.CheckEmailAddress(Trim(Range("sheet1.EmailAddressSec").value), validateType) Then

                   If validateType = "1" Then fmsgboxStatus "*Email ID in Sheet PART A-GENERAL , contains special characters" ', vbOKOnly
                   If validateType = "2" Then fmsgboxStatus "*EmailAddress-2 in Sheet : PART A-GENERAL  should not exceed 125 characters " ', vbOKOnly
                   If validateType = "3" Then fmsgboxStatus "*EmailAddress-2 in Sheet : PART A-GENERAL  is invalid" ', vbOKOnly

                   Range("sheet1.EmailAddressSec").Select
                   Range("sheet1.EmailAddressSec").value = ""
               End If

           End If
              If Not Intersect(Range("sheet1.STDcode"), Range(Target.Address)) Is Nothing Then
                If Not (Part_A_General.ValidateSTDcode()) Then
                    Sheet1.Range("sheet1.STDcode").value = ""
                    Sheet1.Range("sheet1.STDcode").Select
                    GoTo endline
                End If
                If Sheet1.Range("sheet1.PhoneNo").value <> "" Then ValidatePhoneNo
                STDcode = Sheet1.Range("sheet1.STDcode").value
              End If

            If Not Intersect(Range("sheet1.PhoneNo"), Range(Target.Address)) Is Nothing Then
                If Part_A_General.ValidateSTDcode Then
                   If Not (Part_A_General.ValidatePhoneNo()) Then
                       GoTo endline
                   End If
                Else
                   Sheet1.Range("sheet1.PhoneNo").value = ""
                   Sheet1.Range("sheet1.STDcode").Select
                   GoTo endline
                End If
                PhoneNo = Sheet1.Range("sheet1.PhoneNo").value
            End If

            If Not Intersect(Range("Countrycode1"), Range(Target.Address)) Is Nothing Then
'            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" Then
            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-Foreign" Then
            If Len(Sheet1.Range("Countrycode1").MergeArea.value) > 5 Then
            fmsgboxStatus ("Country Code for Mobile No 1 in Sheet : Part A-General cannot exceed 5 characters ")
            Sheet1.Range("Countrycode1").MergeArea.value = ""
            End If
            End If

            If Sheet1.Range("Countrycode1").MergeArea.value = "0" Or Sheet1.Range("Countrycode1").MergeArea.value = "00" Or Sheet1.Range("Countrycode1").MergeArea.value = "000" Or Sheet1.Range("Countrycode1").MergeArea.value = "0000" Or Sheet1.Range("Countrycode1").MergeArea.value = "00000" Then
            fmsgboxStatus ("Please enter valid Country Code for Mobile No 1 in Sheet : Part A-General ")
            Sheet1.Range("Countrycode1").MergeArea.value = ""
            End If

            End If

            If Not Intersect(Range("Countrycode2"), Range(Target.Address)) Is Nothing Then
'            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" Then
            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-Foreign" Then
            If Len(Sheet1.Range("Countrycode2").MergeArea.value) > 5 Then
            fmsgboxStatus ("Country Code for Mobile No 2 in Sheet : Part A-General cannot exceed 5 characters ")
            Sheet1.Range("Countrycode2").MergeArea.value = ""
            End If
            End If

            If Sheet1.Range("Countrycode2").MergeArea.value = "0" Or Sheet1.Range("Countrycode2").MergeArea.value = "00" Or Sheet1.Range("Countrycode2").MergeArea.value = "000" Or Sheet1.Range("Countrycode2").MergeArea.value = "0000" Or Sheet1.Range("Countrycode2").MergeArea.value = "00000" Then
            fmsgboxStatus ("Please enter valid Country Code for Mobile No 2 in Sheet : Part A-General ")
            Sheet1.Range("Countrycode2").MergeArea.value = ""
            End If

            End If
            
            
            
            
            'Added by Aavula Naresh for AY 2026-27
            If Not Intersect(Range("Sheet1.RepEmail"), Range(Target.Address)) Is Nothing Then
           
               If Not Part_A_General.CheckEmailAddress(Trim(Range("Sheet1.RepEmail").value), validateType) Then

                   If validateType = "1" Then fmsgboxStatus "*Email-ID of the representative assessee in Sheet PART A-GENERAL , contains special characters"
                   If validateType = "2" Then fmsgboxStatus "*""Email-Id of the representative assessee cannot be more than 125 characters in schedule Part A General""" ', vbOKOnly
                   If validateType = "3" Then fmsgboxStatus "*""Please enter correct email-id of the representative assessee in schedule Part A General""" ', vbOKOnly

                   Range("Sheet1.RepEmail").Select
                   Range("Sheet1.RepEmail").value = ""
               End If

           End If
           
           
           
           
           
           
'           If Not Intersect(Range("Countrycode3"), Range(target.Address)) Is Nothing Then
'            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
'            If Len(Sheet1.Range("Countrycode3").MergeArea.value) > 5 Then
'            fmsgboxStatus ("""Please enter a valid Contact Number of representative assessee in schedule Part A General""")
'            Sheet1.Range("Countrycode3").MergeArea.value = ""
'            End If
'            End If
'
'''            If Sheet1.Range("Countrycode3").MergeArea.value = "0" Or Sheet1.Range("Countrycode3").MergeArea.value = "00" Or Sheet1.Range("Countrycode3").MergeArea.value = "000" Or Sheet1.Range("Countrycode3").MergeArea.value = "0000" Or Sheet1.Range("Countrycode3").MergeArea.value = "00000" Then
''            fmsgboxStatus ("""Please enter a valid 10 digit Contact Number of  representative assessee in schedule Part A General""")
''            Sheet1.Range("Countrycode3").MergeArea.value = ""
''            End If
'            End If
            
            
'            If Not Intersect(Sheet1.Range("BAC115.DateOfFiling"), Sheet1.Range(target.Address)) Is Nothing Then
'                If Not target.value = "" Then
'                    If Not CheckDateFilinginSheet_DOF1B(target.value) Then
'                    Sheet1.Range("BAC115.DateOfFiling").value = ""
'                    End If
'                End If
'            End If

'            If Not Intersect(Sheet1.Range("BAC115.DateOfFiling_New_Yes"), Sheet1.Range(target.Address)) Is Nothing Then
'                If Not target.value = "" Then
'                    If Not CheckDateFilinginSheet_DOF1B(target.value) Then
'                    Sheet1.Range("BAC115.DateOfFiling_New_Yes").value = ""
'                    End If
'                End If
'            End If
'
'            If Not Intersect(Sheet1.Range("BAC115.DateOfFiling_New_No"), Sheet1.Range(target.Address)) Is Nothing Then
'                If Not target.value = "" Then
'                    If Not CheckDateFilinginSheet_DOF1B(target.value) Then
'                    Sheet1.Range("BAC115.DateOfFiling_New_No").value = ""
'                    End If
'                End If
'            End If
'
'            If Not Intersect(Sheet1.Range("BAC115.DateOfFiling_New_NA"), Sheet1.Range(target.Address)) Is Nothing Then
'                If Not target.value = "" Then
'                    If Not CheckDateFilinginSheet_DOF1B(target.value) Then
'                    Sheet1.Range("BAC115.DateOfFiling_New_NA").value = ""
'                    End If
'                End If
'            End If
endline:
End Sub

Sub PartAGen_sub3(ByVal Target As Range)

''added by Chetan C M on 27/01/2026 for AY 2026-27
'--start--
LockUnlockPLSec44

'LockUnlockEIincome
'--end

Dim validateType As String

 If Not Intersect(Range("sheet1.MobileNo"), Range(Target.Address)) Is Nothing Then
             If Not (Part_A_General.ChkMobileNo()) Then
                   Sheet1.Range("sheet1.MobileNo").value = ""
                   Sheet1.Range("sheet1.MobileNo").Select

            End If
        End If
        If Not Intersect(Range("sheet1.MobileNoSec"), Range(Target.Address)) Is Nothing Then
             If Not (Part_A_General.ValidateMobileNoSec()) Then
                   Sheet1.Range("sheet1.MobileNoSec").value = ""
                   Sheet1.Range("sheet1.MobileNoSec").Select

            End If
        End If

If Not Intersect(Range("sheet1.SeventhProvisoFlag"), Range(Target.Address)) Is Nothing Then

   Sheet1.Unprotect Password:=getmsgstate
    

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

Sheet1.Unprotect Password:=getmsgstate
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag").MergeArea.value = "(Select)"

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.value = "(Select)"

        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If
WClockunlockProviso
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.DepositAmountFlag"), Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "N") Then
        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "Y") Then
        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.AggrigateAmountFlag"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.AggrigateAmountFlag1"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

'If Not Intersect(Range("sheet1.LiableSec44AEflg"), Range(target.Address)) Is Nothing Then
'
'   Sheet1.Unprotect Password:=getmsgstate
'If (Mid(Range("sheet1.LiableSec44AEflg").value, 1, 1) = "Y") Then
'        Range("sheet1.LiableSec5cflg").MergeArea.ClearContents
'        Range("sheet1.LiableSec5cflg").MergeArea.Locked = True
'        Range("sheet1.LiableSec5cflg").MergeArea.Interior.Color = (&HD8D8D8)
'Else
'        Range("sheet1.LiableSec5cflg").MergeArea.Locked = False
'        Range("sheet1.LiableSec5cflg").MergeArea.Interior.Color = (&HCCFFCC)
'End If
'Sheet1.Protect Password:=getmsgstate
'End If

If Not Intersect(Range("sheet1.LiableSec5cflg"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
'If (Mid(Range("sheet1.LiableSec5cflg").value, 1, 1) = "Y") Then

        If Sheet1.Range("sheet1.LiableSec5cflg").value = "More than Rs. 1 crore and up to Rs. 10 crores" Then 'Added by Aavula Naresh on 11/03/2026
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Locked = False
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Interior.Color = (&HCCFFCC)
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Locked = False
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Interior.Color = (&HCCFFCC)
Else
        Range("sheet1.LiableSecAggrigateflg").MergeArea.ClearContents
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Locked = True
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.ClearContents
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Locked = True
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Interior.Color = (&HD8D8D8)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.DepositAmount"), Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.DepositAmount").value < 10000000 Or Sheet1.Range("Sheet1.DepositAmount").value = 10000000) Then
            fmsgbox "*Amount entered cannot be less than or equal to Rs 1 Crore" & Chr(10) & Chr(13)
            Sheet1.Range("Sheet1.DepositAmount").value = " "

        End If
End If

If Not Intersect(Range("sheet1.AggrigateAmount"), Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.AggrigateAmount").value < 200000 Or Sheet1.Range("Sheet1.AggrigateAmount").value = 200000) Then
            fmsgbox "*Amount entered cannot be less than or equal to Rs 2 Lakhs" & Chr(10) & Chr(13)
            Sheet1.Range("Sheet1.AggrigateAmount").value = " "

        End If
End If

If Not Intersect(Range("sheet1.AggrigateAmount1"), Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.AggrigateAmount1").value < 100000 Or Sheet1.Range("Sheet1.AggrigateAmount1").value = 100000) Then
            fmsgbox "*Amount entered cannot be less than or equal to Rs 1 Lakhs" & Chr(10) & Chr(13)
            Sheet1.Range("Sheet1.AggrigateAmount1").value = " "
        End If
End If

If Not Intersect(Range("sheet1.ReturnFileSec"), Range(Target.Address)) Is Nothing Then
        sReceiptNo = ""
        sNoticeDate = ""
        sNoticeNo = ""
        Sheet1.Unprotect Password:=getmsgstate
        LockUnlockFilingStatusDtls
        LockUnlockFilingStatusDtls1
        Sheet52.Unprotect Password:=getmsgstate
            If Mid(Trim(Range("sheet1.ReturnFileSec1")), 1, 2) = "19" Then
                 Sheet52.Range("Veri.92CD").EntireRow.Hidden = False
             Else
                 Sheet52.Range("Veri.92CD").EntireRow.Hidden = True
            End If
        If Mid(Trim(Range("sheet1.ReturnFileSec1")), 1, 2) = "20" Then
'            fmsgboxStatus "* 119(2)(b)-after condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)." ', vbInformation, "ITR3"
          'Changed by sai on 20/09/2024
'            fmsgboxStatus "* 139(9A) - After condonation of delay u/s 119(2)(b) should be filed after expiry of due date mentioned in 139(1)/139(4)." ', vbInformation, "ITR3"
            
            'Added by Aavula on 22/10/2024
            
'           fmsgboxStatus " Section 139(9A)- After condonation of delay  u/s 119(2)(b) should be filed after expiry of due date mentioned in 139(1)/139(4) in Part A - General" ', vbInformation, "ITR3"
           fmsgboxStatus " ""Section 139(9A) should be filed after expiry of due date mentioned in 139(1)/139(4) in Part A - General""" ', vbInformation, "ITR3"
            
            Range(Target.Address).Select
            End If
       If UCase(Mid(Trim(Sheet1.Range("sheet1.ReturnFileSec").value), 1, 6)) = "139(4)" Then
            LockUnlock80
       Else
            LockUnlock80
       End If
        Sheet1.Protect Password:=getmsgstate
        Sheet52.Protect Password:=getmsgstate

End If

If Not Intersect(Range("sheet1.OrigRetFiledDate"), Range(Target.Address)) Is Nothing Then
'               If Trim(Range("sheet1.OrigRetFiledDate").value) <> "00/00/0000" Then
               If Trim(Range("sheet1.OrigRetFiledDate").value) <> "" Then
                   If Not Part_A_General.CheckDateBefore(Trim(Range("sheet1.OrigRetFiledDate").value)) Then
                       Range("sheet1.OrigRetFiledDate").Select
                       Range("sheet1.OrigRetFiledDate").value = ""
'                   Else
'                   Range("sheet1.OrigRetFiledDate").value = EfilingDate
                   End If
'Change-20.12.2022.102.018A
'                   If Dformat(Target.value, "") > Dformat(Sheet51.Range("sheet9.Date").value, "") Then
                   If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
'end---
'                    MsgBox ("Date cannot be greater than today's date")
                    MsgBox ("Date of filing of original return shall not be beyond system date.")

                    Target.value = ""
                  End If
               End If

           End If
          If Not Intersect(Range("sheet1.NoticeDate"), Range(Target.Address)) Is Nothing Then
'               If Trim(Range("sheet1.NoticeDate").value) <> "00/00/0000" Then
               If Trim(Range("sheet1.NoticeDate").value) <> "" Then

                   If Not Part_A_General.CheckNoticeDateBefore(Trim(Range("sheet1.NoticeDate").value)) Then
                       Range("sheet1.NoticeDate").Select
                       Range("sheet1.NoticeDate").value = ""
                   Else
                        'Change-20.12.2022.102.019B
                        If Dformat(Target.value, "") > Dformat(Sheet52.Range("sheet9.Date").value, "") Then
                            MsgBox ("Date cannot be greater than today's date")
                            Target.value = ""
                        Else
                            Range("sheet1.NoticeDate").value = noticeDate
                        End If
                        'end--
                   End If
               End If

           End If

       If Not Intersect(Range("sheet1.PortugeseCC5A"), Range(Target.Address)) Is Nothing Then
               Application.ScreenUpdating = False
           If UCase(Mid(Target.value, 1, 1)) = "N" Then
                   LOCK5A
           ElseIf UCase(Mid(Target.value, 1, 1)) = "Y" Then
                   fmsgboxStatus "*You have selected to be governed by Sec 5A. Please Fill Schedule 5A."
                   fmsgboxStatus "*Please enter your share of Receipts in all income  heads other than ""Income from Salary"" .The share of Receipts of the spouse should be filled in Schedule 5A"
                   UNLOCK5A
            Else
                UNLOCK5A
           End If
           Application.ScreenUpdating = True
        End If
        
''by Sai on 12/04/2025*****************************************************************************************
' If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(target.Address)) Is Nothing Then
'        If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
'            Sheet13.Unprotect Password:=sPassword
'            Application.EnableEvents = False
'
'            Sheet13.Range("STCG.A8_CertiStatus").Locked = True
'            Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HD8D8D8)
'            Sheet13.Range("STCG.A8_CertiStatus").ClearContents
'
'
'            Sheet13.Range("LTCG.B10_CertiStatus").Locked = True
'            Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HD8D8D8)
'            Sheet13.Range("LTCG.B10_CertiStatus").ClearContents
'
'            Sheet13.Protect Password:=sPassword
'            Application.EnableEvents = True
'        Else
'            Sheet13.Unprotect Password:=sPassword
'            Application.EnableEvents = False
'
'            Sheet13.Range("STCG.A8_CertiStatus").Locked = False
'            Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HCCFFCC)
'            Sheet13.Range("STCG.A8_CertiStatus").value = "(Select)"
'
'
'            Sheet13.Range("LTCG.B10_CertiStatus").Locked = False
'            Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HCCFFCC)
'            Sheet13.Range("LTCG.B10_CertiStatus").value = "(Select)"
'
'            Sheet13.Protect Password:=sPassword
'            Application.EnableEvents = True
'        End If
'End If
'
''***************************************************************************************************************
'
'
'    If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(target.Address)) Is Nothing Then
'      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
'         Sheet14.Unprotect Password:=sPassword
'         Sheet13.Unprotect Password:=sPassword
'         Sheet19.Unprotect Password:=sPassword
'
'         Sheet14.Range("os.TRC").Locked = False
'         Sheet14.Range("os.TRC").Interior.Color = (&HCCFFCC)
'         Sheet14.Range("os.TRC").value = "(Select)"
'
''          UNLOCKRANGEOS_DTAA 'For AY 2025-26 by Sai on 28/02/2025
'
'         '---'Uncommenting by Bindu on 11th feb 26 as per V0.01
'         Sheet13.Range("STCG.A8_CertiStatus").Locked = False
'         Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HCCFFCC)
'         Sheet13.Range("STCG.A8_CertiStatus").value = "(Select)"
'
'
''          UNLOCKRANGEcg_8tabA1toA7  'For AY 2025-26 by Sai on 28/02/2025
'
'         Sheet13.Range("LTCG.B10_CertiStatus").Locked = False
'         Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HCCFFCC)
'         Sheet13.Range("LTCG.B10_CertiStatus").value = "(Select)"
'
''           UNLOCKRANGEcg_12tbB1toB11 'For AY 2025-26 by Sai on 28/02/2025
'           '---'End - By Bindu on 11th Feb 26
'
'       Sheet19.Unprotect Password:=sPassword ' Added by Malli on 04/07/2025
'
'         Sheet19.Range("scvia.Section80QQB").Locked = True
'         Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HD8D8D8)
'         Sheet19.Range("scvia.Section80QQB").ClearContents
'
'         Sheet19.Unprotect Password:=sPassword
'         Sheet19.Range("scvia.Section80RRB").Locked = True
'         Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HD8D8D8)
'         Sheet19.Range("scvia.Section80RRB").ClearContents
'
'         Sheet19.Unprotect Password:=sPassword
'         Sheet14.Range("OSIncreliefus89A").Locked = True
'         Sheet14.Range("OSIncreliefus89A").Interior.Color = (&HD8D8D8)
'         Sheet14.Range("OSIncreliefus89A").ClearContents
'
'         Sheet19.Protect Password:=sPassword
'         Sheet14.Protect Password:=sPassword
'         Sheet13.Protect Password:=sPassword
'       Else
'        Sheet14.Unprotect Password:=sPassword
'         Sheet13.Unprotect Password:=sPassword
'         Sheet19.Unprotect Password:=sPassword
'
'          Sheet14.Range("os.TRC").Locked = True
'          Sheet14.Range("os.TRC").Interior.Color = (&HD8D8D8)
'          Sheet14.Range("os.TRC").ClearContents
'
'
''           LOCKRANGEOS_DTAA 'For AY 2025-26 by Sai on 28/02/2025
'
'           '--'Uncommenting by Bindu on 11th feb 26 as per V0.01
'           Sheet13.Range("STCG.A8_CertiStatus").Locked = True
'          Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HD8D8D8)
'          Sheet13.Range("STCG.A8_CertiStatus").ClearContents
'
''           LOCKRANGEcg_8tabA1toA7  'For AY 2025-26 by Sai on 28/02/2025
'
'          Sheet13.Range("LTCG.B10_CertiStatus").Locked = True
'          Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HD8D8D8)
'          Sheet13.Range("LTCG.B10_CertiStatus").ClearContents
'
''           LOCKRANGEcg_12tbB1toB11 'For AY 2025-26 by Sai on 28/02/2025



'by Sai on 12/04/2025*****************************************************************************************
 If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then
        If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
            Sheet13.Unprotect Password:=sPassword
            Application.EnableEvents = False
            
            Sheet13.Range("STCG.A8_CertiStatus").Locked = True
            Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HD8D8D8)
            Sheet13.Range("STCG.A8_CertiStatus").ClearContents
            
            
            Sheet13.Range("LTCG.B10_CertiStatus").Locked = True
            Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HD8D8D8)
            Sheet13.Range("LTCG.B10_CertiStatus").ClearContents
    
            Sheet13.Protect Password:=sPassword
            Application.EnableEvents = True
        Else
            Sheet13.Unprotect Password:=sPassword
            Application.EnableEvents = False
            
            Sheet13.Range("STCG.A8_CertiStatus").Locked = False
            Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HCCFFCC)
            Sheet13.Range("STCG.A8_CertiStatus").value = "(Select)"
            
            
            Sheet13.Range("LTCG.B10_CertiStatus").Locked = False
            Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HCCFFCC)
            Sheet13.Range("LTCG.B10_CertiStatus").value = "(Select)"
    
            Sheet13.Protect Password:=sPassword
            Application.EnableEvents = True
        End If
End If
        
'***************************************************************************************************************

        
    If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
         Sheet14.Unprotect Password:=sPassword
         Sheet13.Unprotect Password:=sPassword
         Sheet19.Unprotect Password:=sPassword

         Sheet14.Range("os.TRC").Locked = False
         Sheet14.Range("os.TRC").Interior.Color = (&HCCFFCC)
         Sheet14.Range("os.TRC").value = "(Select)"
      
'          UNLOCKRANGEOS_DTAA 'For AY 2025-26 by Sai on 28/02/2025

'         Sheet13.Range("STCG.A8_CertiStatus").Locked = False
'         Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HCCFFCC)
'         Sheet13.Range("STCG.A8_CertiStatus").value = "(Select)"
        
'          UNLOCKRANGEcg_8tabA1toA7  'For AY 2025-26 by Sai on 28/02/2025
          
'         Sheet13.Range("LTCG.B10_CertiStatus").Locked = False
'         Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HCCFFCC)
'         Sheet13.Range("LTCG.B10_CertiStatus").value = "(Select)"
          
'           UNLOCKRANGEcg_12tbB1toB11 'For AY 2025-26 by Sai on 28/02/2025

       Sheet19.Unprotect Password:=sPassword ' Added by Malli on 04/07/2025
       
         Sheet19.Range("scvia.Section80QQB").Locked = True
         Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HD8D8D8)
         Sheet19.Range("scvia.Section80QQB").ClearContents
         
         Sheet19.Unprotect Password:=sPassword
         Sheet19.Range("scvia.Section80RRB").Locked = True
         Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HD8D8D8)
         Sheet19.Range("scvia.Section80RRB").ClearContents
         
         Sheet19.Unprotect Password:=sPassword
         Sheet14.Range("OSIncreliefus89A").Locked = True
         Sheet14.Range("OSIncreliefus89A").Interior.Color = (&HD8D8D8)
         Sheet14.Range("OSIncreliefus89A").ClearContents

         Sheet19.Protect Password:=sPassword
         Sheet14.Protect Password:=sPassword
         Sheet13.Protect Password:=sPassword
       Else
        Sheet14.Unprotect Password:=sPassword
         Sheet13.Unprotect Password:=sPassword
         Sheet19.Unprotect Password:=sPassword
         
          Sheet14.Range("os.TRC").Locked = True
          Sheet14.Range("os.TRC").Interior.Color = (&HD8D8D8)
          Sheet14.Range("os.TRC").ClearContents
           
'           LOCKRANGEOS_DTAA 'For AY 2025-26 by Sai on 28/02/2025
           
'           Sheet13.Range("STCG.A8_CertiStatus").Locked = True
'          Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HD8D8D8)
'          Sheet13.Range("STCG.A8_CertiStatus").ClearContents

'           LOCKRANGEcg_8tabA1toA7  'For AY 2025-26 by Sai on 28/02/2025

'          Sheet13.Range("LTCG.B10_CertiStatus").Locked = True
'          Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HD8D8D8)
'          Sheet13.Range("LTCG.B10_CertiStatus").ClearContents

'           LOCKRANGEcg_12tbB1toB11 'For AY 2025-26 by Sai on 28/02/2025

'--Ends by Bindu on 11th feb 26
           
         Sheet19.Unprotect Password:=sPassword
         Sheet19.Range("scvia.Section80QQB").Locked = False
         Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HCCFFCC)
         Sheet19.Range("scvia.Section80QQB").value = 0
         
         Sheet19.Unprotect Password:=sPassword
         Sheet19.Range("scvia.Section80RRB").Locked = False
         Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HCCFFCC)
         Sheet19.Range("scvia.Section80RRB").value = 0
         
         Sheet14.Range("OSIncreliefus89A").Locked = False
         Sheet14.Range("OSIncreliefus89A").Interior.Color = (&HCCFFCC)

         Sheet19.Protect Password:=sPassword
         Sheet14.Protect Password:=sPassword
         Sheet13.Protect Password:=sPassword
        End If
            lockunlock_OS_Stch_Ltcg
            LockUnlockIBANSELECT
            LockUnlockClaim115H
            LockUnlockQVIA
            NewResidentialwork
            LockUnlockTR
            LockUnlockFSI
            
           Sheet1.Range("sheet1.ResidentialStatus1_Condition").value = "(Select)"
           If Trim(Mid(Target.value, 1, 3)) <> "NRI" Then
           sPassword = EfilingCommon.getmsgstate
           Sheet1.Unprotect Password:=sPassword
           Sheet1.Range("sheet1.NRI_PE").MergeArea.Locked = True
           Sheet1.Range("sheet1.NRI_PE").Interior.Color = (&HD8D8D8)
            Sheet1.Range("sheet1.NRI_PE").value = ""
           Sheet1.Protect Password:=sPassword
               LOCKCG
               UnLockBp
           Else
           sPassword = EfilingCommon.getmsgstate
           Sheet1.Unprotect Password:=sPassword
           Sheet1.Range("sheet1.NRI_PE").MergeArea.Locked = False
           Sheet1.Range("sheet1.NRI_PE").Interior.Color = (&HCCFFCC)
           Sheet1.Range("sheet1.NRI_PE").value = "(Select)"
               UNLOCKcg
               LockBp
           Sheet1.Protect Password:=sPassword
            End If
            If Not isdropdownblank(Target.value) Then
             ResetOSDropDown (Mid(Target.value, 1, 2))
            End If
       End If
       
'           If Not Intersect(Range("sheet1.AsseseeRepFlg"), Range(target.Address)) Is Nothing Then
'           'Commented by Aavula For AY 2026-27
''                   RepInfo
'           End If
           
If Not Intersect(Range("sheet1.LiableSec44ABflg"), Range(Target.Address)) Is Nothing Then

   If Mid(Range("sheet1.LiableSec44ABflg").value, 1, 1) = "Y" Then
       fmsgboxStatus "*Part A - OI should be filled mandatorily " ', vbOKOnly
   Else:
       fmsgboxStatus "*If not liable under Section 44AB then Audit details need not be fillled" ', vbOKOnly
   End If
End If
If Not Intersect(Range("sheet1.LiableSec92Eflg"), Range(Target.Address)) Is Nothing Then
    LockUnlock92Edia
End If

If Not Intersect(Range("sheet1.LiableSec92Eflgdia"), Range(Target.Address)) Is Nothing Then
    LockUnlockDOA92E
    LockUnlockDOA92E_AckNum 'Newly added by Sai For AY 2024-25
End If

If Not Intersect(Range("sheet1.AsseseeRepFlg"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword
  
    If UCase(Sheet1.Range("sheet1.AsseseeRepFlg").value) = UCase("Y - Yes") Then

      Sheet1.Unprotect Password:=getmsgstate
      Application.EnableEvents = False
      
       Range("sheet1.RepName").MergeArea.Locked = False
       Range("sheet1.RepName").MergeArea.Interior.Color = (&HCCFFCC)
       
       Range("Sheet1.RepEmail").MergeArea.Locked = False
       Range("Sheet1.RepEmail").MergeArea.Interior.Color = (&HCCFFCC)
       
       Range("sheet1.Repcountrycode").MergeArea.Locked = False
       Range("sheet1.Repcountrycode").MergeArea.Interior.Color = (&HCCFFCC)
       
       Range("sheet1.RepMobileNo").MergeArea.Locked = False
       Range("Sheet1.RepMobileNo").MergeArea.Interior.Color = (&HCCFFCC)
        
    ElseIf UCase(Sheet1.Range("Sheet1.AsseseeRepFlg").value) = UCase("N -No") Then

       Range("sheet1.RepName").MergeArea.Locked = True
       Range("sheet1.RepName").MergeArea.Interior.Color = (&HD8D8D8)
       Range("sheet1.RepName").MergeArea.ClearContents
       
       Range("Sheet1.RepEmail").MergeArea.Locked = True
       Range("Sheet1.RepEmail").MergeArea.Interior.Color = (&HD8D8D8)
       Range("Sheet1.RepEmail").MergeArea.ClearContents

       
       Range("sheet1.Repcountrycode").MergeArea.Locked = True
       Range("sheet1.Repcountrycode").MergeArea.Interior.Color = (&HD8D8D8)
       Range("sheet1.Repcountrycode").MergeArea.ClearContents
              
       Range("Sheet1.RepMobileNo").MergeArea.Locked = True
       Range("Sheet1.RepMobileNo").MergeArea.Interior.Color = (&HD8D8D8)
       Range("Sheet1.RepMobileNo").MergeArea.ClearContents
      
      ElseIf UCase(Sheet1.Range("Sheet1.AsseseeRepFlg").value) = UCase("(Select)") Then

       Range("sheet1.RepName").MergeArea.Locked = True
       Range("sheet1.RepName").MergeArea.Interior.Color = (&HD8D8D8)
       Range("sheet1.RepName").MergeArea.ClearContents
       
       Range("Sheet1.RepEmail").MergeArea.Locked = True
       Range("Sheet1.RepEmail").MergeArea.Interior.Color = (&HD8D8D8)
       Range("Sheet1.RepEmail").MergeArea.ClearContents

       
       Range("sheet1.Repcountrycode").MergeArea.Locked = True
       Range("sheet1.Repcountrycode").MergeArea.Interior.Color = (&HD8D8D8)
       Range("sheet1.Repcountrycode").MergeArea.ClearContents
              
       Range("Sheet1.RepMobileNo").MergeArea.Locked = True
       Range("Sheet1.RepMobileNo").MergeArea.Interior.Color = (&HD8D8D8)
       Range("Sheet1.RepMobileNo").MergeArea.ClearContents
        
        End If

   End If
 Sheet1.Protect Password:=sPassword
'End If


If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then

  sPassword = EfilingCommon.getmsgstate
  
  Sheet4.Unprotect Password:=sPassword
  Sheet7.Unprotect Password:=sPassword
  


If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR" Then

   If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
   
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("sheet11.Section44AD").formula = "=0"
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("sheet11.Section44ADA").formula = "=0"

        LOCKPL2

   Else
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("sheet11.Section44AD").formula = "=PL_61ii"
        Sheet9.Unprotect Password:=sPassword
        Sheet9.Range("sheet11.Section44ADA").formula = "=PL_62ii"
   UNLOCKPL2
   End If
   Sheet7.Range("AOIFlag_1").Locked = True
   Sheet7.Range("AOIFlag_1").Interior.Color = (&HD8D8D8)
   Sheet7.Range("AOIFlag_1").ClearContents

Else
   Sheet9.Unprotect Password:=sPassword
   Sheet9.Range("sheet11.Section44AD").formula = "=PL_61ii"
   Sheet9.Unprotect Password:=sPassword
   Sheet9.Range("sheet11.Section44ADA").formula = "=PL_62ii"
   
    UNLOCKPL2
   Sheet7.Range("AOIFlag_1").Locked = False
   Sheet7.Range("AOIFlag_1").Interior.Color = (&HCCFFCC)
   Sheet7.Range("AOIFlag_1").ClearContents
End If


'If Range("sheet1.ResidentialStatus1").value = "RES - Resident" Then
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then
Sheet1.Unprotect Password:=getmsgstate

        
        Range("sheet1.NRI_SEP").MergeArea.Locked = True
        Range("sheet1.NRI_SEP").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.NRI_SEP").value = "NA"
        Sheet1.Unprotect Password:=getmsgstate
        Range("sheet1.NRI_SEP_AggrPayment").MergeArea.Locked = True
        Range("sheet1.NRI_SEP_AggrPayment").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.NRI_SEP_AggrPayment").value = ""
        
        Sheet1.Unprotect Password:=getmsgstate
        Range("sheet1.NRI_SEP_Usr").MergeArea.Locked = True
        Range("sheet1.NRI_SEP_Usr").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.NRI_SEP_Usr").value = ""

Else
Sheet1.Unprotect Password:=sPassword
        Range("sheet1.NRI_SEP").MergeArea.Locked = False
        Range("sheet1.NRI_SEP").MergeArea.Interior.Color = (&HCCFFCC)
        Range("sheet1.NRI_SEP").value = "(Select)"

Sheet1.Unprotect Password:=sPassword
        Range("sheet1.NRI_SEP_AggrPayment").MergeArea.Locked = False
        Range("sheet1.NRI_SEP_AggrPayment").MergeArea.Interior.Color = (&HCCFFCC)
Sheet1.Unprotect Password:=sPassword
        Range("sheet1.NRI_SEP_Usr").MergeArea.Locked = False
        Range("sheet1.NRI_SEP_Usr").MergeArea.Interior.Color = (&HCCFFCC)
End If
 Sheet1.Protect Password:=sPassword
 Sheet7.Protect Password:=sPassword
 Sheet9.Protect Password:=sPassword
 Sheet4.Protect Password:=sPassword

End If


End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

'sudha


Application.EnableEvents = False
    Dim PinCode As Range
        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.PinCode").value <> "" Then
                'if state and pincode doesn't match, empty the pincode.

                    'For Each PinCode In Sheet33.Range("Duplicate_Pincode_list")
                    'If Sheet1.Range("sheet1.PinCode").value <> PinCode Then


               Dim state1 As Variant

               state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, SheetALL.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Sheet1.Range("sheet1.StateCode1").value, 4)) Or StateMatchesPin(Target.value)) Then
                        Sheet1.Range("sheet1.PinCode").value = ""

                End If

            End If
            'Next PinCode
        End If


    
        
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then
         Dim DupFlag As Boolean
         DupFlag = False
                 'If the pincode is not in the duplicate list, populate, else don't populate---
                 For Each PinCode In SheetALL.Range("Duplicate_Pincode_list")
                    If Sheet1.Range("sheet1.PinCode").value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode
                 
                
                If DupFlag = False Then
                    Dim StateName
                    Dim StateCode
                    
                    StateName = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, SheetALL.Range("All_PinCode_V"), 2, False)
                    If IsError(StateName) Then

                        Sheet1.Range("sheet1.PinCode").value = ""

                        
                    Else
                        StateCode = Application.VLookup(StateName, SheetALL.Range("ES:ET"), 2, False)
                        Sheet1.Range("sheet1.StateCode1").value = StateCode
                        Sheet1.Range("sheet1.CountryCode").value = "91-INDIA"
                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
                    End If
                 End If
                 
                 If DupFlag = True Then
                                 Application.EnableEvents = False

                    If Sheet1.Range("sheet1.StateCode1").value <> "" Then
                    If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode1").value) Then
                       
                    Sheet1.Range("sheet1.StateCode1").value = "(Select)"
                    Else
                    
                    Sheet1.Range("sheet1.PinCode").value = PinCode
                     

                    Application.EnableEvents = True
                            
                        End If
                    End If
                 End If

         End If
Application.EnableEvents = True
 sPassword = EfilingCommon.getmsgstate
 
 
 
 Application.EnableEvents = False
        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.PinCode2").value <> "" Then
                'if state and pincode doesn't match, empty the pincode.

               Dim State2 As Variant

               state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode2").value, SheetALL.Range("All_PinCode_V"), 2, False)
                If Not (State2 = UCase(Mid(Sheet1.Range("sheet1.StateCode2").value, 4)) Or StateMatchesPin(Target.value)) Then
                        Sheet1.Range("sheet1.PinCode2").value = ""

                End If

            End If
        End If
   
        
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
         DupFlag = False
                 'If the pincode is not in the duplicate list, populate, else don't populate---
                 For Each PinCode In SheetALL.Range("Duplicate_Pincode_list")
                    If Sheet1.Range("sheet1.PinCode2").value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode
                 
                
                If DupFlag = False Then
                    
                    StateName = Application.VLookup(Sheet1.Range("sheet1.PinCode2").value, SheetALL.Range("All_PinCode_V"), 2, False)
                    If IsError(StateName) Then

                        Sheet1.Range("sheet1.PinCode2").value = ""

                        
                    Else
                        StateCode = Application.VLookup(StateName, SheetALL.Range("ES:ET"), 2, False)
                        Sheet1.Range("sheet1.StateCode2").value = StateCode
                        Sheet1.Range("sheet1.CountryCode2").value = "91-INDIA"
                        
                    End If
                 End If
                 
                 If DupFlag = True Then
                                 Application.EnableEvents = False

                    If Sheet1.Range("sheet1.StateCode2").value <> "" Then
                    If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode2").value) Then
                       
                    Sheet1.Range("sheet1.StateCode2").value = "(Select)"
                    Else
                    
                    Sheet1.Range("sheet1.PinCode2").value = PinCode
                     

                    Application.EnableEvents = True
                            
                        End If
                    End If
                 End If

         End If
Application.EnableEvents = True
 sPassword = EfilingCommon.getmsgstate
 
 
 
 
'sudha

'**************************************************************
'Added by sai on 12/03/2025 for AY 2025-26
If Not Application.Intersect(Sheet1.Range("sheet1.ResidentialStatus1"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("Sheet1.115H"), Sheet1.Range(Target.Address)) Is Nothing Then
    
    'If UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) = "RES" And UCase(Mid(Sheet1.Range("Sheet1.115H").value, 1, 3)) = "YES" Then
    
    'SIT-118438 by sadineni on 28/04/2026
    'If UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) <> "NRI" And UCase(Mid(Sheet1.Range("Sheet1.115H").value, 1, 3)) = "YES" Then '
'Changed by sadineni on 04/06/2026 as per BA-Manjima
    If UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) <> "NRI" Or UCase(Mid(Sheet1.Range("Sheet1.115H").value, 1, 3)) = "YES" Then
       UNLOCKRANGECG_115ACA
    Else
      LOCKRANGECG_115ACA
    End If
If UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) = "RES" Then
   
   Sheet7.Unprotect Password:=sPassword
   Sheet7.Range("AOIFlag_1").Locked = False
   Sheet7.Range("AOIFlag_1").Interior.Color = (&HCCFFCC)
   Sheet7.Range("AOIFlag_1").ClearContents
   
      Else
      
   Sheet7.Unprotect Password:=sPassword
   Sheet7.Range("AOIFlag_1").Locked = True
   Sheet7.Range("AOIFlag_1").Interior.Color = (&HD8D8D8)
   Sheet7.Range("AOIFlag_1").ClearContents
   

   
Sheet7.Protect Password:=sPassword
End If

End If
'****************************************************************

CalculateTaxCheckFlag = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Dim validateType As String
Dim i
validateType = ""
Sheet1.Unprotect Password:=getmsgstate



If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword
  If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
         If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") Or (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I") Then
            Lock80U
            Lock80DD
            End If
  Else
            UnLock80U
            UnLock80DD
            
  End If
End If




'----------------------
'''''''''''''''''''''''''
'Added by Aavula Naresh for AY 2025-26
'
'If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(target.Address)) Is Nothing Then
'  sPassword = EfilingCommon.getmsgstate
'  Sheet1.Unprotect Password:=sPassword
'    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
'              LockEI4
'    Else
'              UnLockEI4
'    End If
'End If

'''''''''''''''''''''''''



'Change-16.05.2023.101 -> Compile error
PartAGen_sub1 Target
PartAGen_sub2 Target
PartAGen_sub3 Target


If Not Intersect(Range("sheet1.LiableSec44ABflg"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword

  If Mid(Range("sheet1.LiableSec44ABflg").value, 1, 1) <> "Y" Then
      
        
  'Newly added by Sai=======================================================
        Range("Sheet1.Condition_44AB").MergeArea.Locked = True
        Range("Sheet1.Condition_44AB").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.Condition_44AB").MergeArea.ClearContents
  '=========================================================================
  
  
  
       Range("sheet1.AuditAccountantFlg").value = "(Select)"
        Range("sheet1.AuditAccountantFlg").MergeArea.Locked = True
        Range("sheet1.AuditAccountantFlg").MergeArea.ClearContents
        Range("sheet1.AuditAccountantFlg").MergeArea.Interior.Color = (&HD8D8D8)


'Commented by Aavula For AY 2026-27

        Range("sheet1.AuditReportFurnishDate").MergeArea.Locked = True
        Range("sheet1.AuditReportFurnishDate").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditReportFurnishDate").MergeArea.ClearContents

'        Range("sheet1.AuditorName").MergeArea.Locked = True
'        Range("sheet1.AuditorName").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AuditorName").MergeArea.ClearContents

'        Range("sheet1.AuditorMemNo").MergeArea.Locked = True
'        Range("sheet1.AuditorMemNo").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AuditorMemNo").MergeArea.ClearContents

        Range("sheet1.AudFrmName").MergeArea.Locked = True
        Range("sheet1.AudFrmName").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmName").MergeArea.ClearContents
'
'        Range("sheet1.AudFrmRegNo").MergeArea.Locked = True
'        Range("sheet1.AudFrmRegNo").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AudFrmRegNo").MergeArea.ClearContents
'
        Range("sheet1.AudFrmPAN").MergeArea.Locked = True
        Range("sheet1.AudFrmPAN").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmPAN").MergeArea.ClearContents

 
        Range("sheet1.AudFrmAadhaar").MergeArea.Locked = True
        Range("sheet1.AudFrmAadhaar").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmAadhaar").MergeArea.ClearContents
'''''''''''

'        Range("sheet1.AuditDate").MergeArea.Locked = True
'        Range("sheet1.AuditDate").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AuditDate").MergeArea.ClearContents
        
        'Newly added by Sai For AY 2024-25----------------------------------------------
        Range("Sheet1.AckNumAuditReport").MergeArea.Locked = True
        Range("Sheet1.AckNumAuditReport").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.AckNumAuditReport").MergeArea.ClearContents
        
        'Commented by Aavula For AY 2026-27
'        Range("Sheet1.UDINAuditNum").MergeArea.Locked = True
'        Range("Sheet1.UDINAuditNum").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("Sheet1.UDINAuditNum").MergeArea.ClearContents
    'End Change by Sai--------------------------------------------------------------
    Else:
        
        Range("Sheet1.Condition_44AB").MergeArea.Locked = False
        Range("Sheet1.Condition_44AB").MergeArea.Interior.Color = (&HCCFFCC)
        
    
        Range("sheet1.AuditAccountantFlg").MergeArea.Locked = False
        Range("sheet1.AuditAccountantFlg").MergeArea.Interior.Color = (&HCCFFCC)

   End If
 Sheet1.Protect Password:=sPassword
End If
If Not Intersect(Range("sheet1.AuditAccountantFlg"), Range(Target.Address)) Is Nothing Then

  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword
  If Mid(Range("sheet1.AuditAccountantFlg").value, 1, 1) <> "Y" Then
  
   'Commented by Aavula For AY 2026-27
       Range("sheet1.AuditReportFurnishDate").MergeArea.Locked = True
        Range("sheet1.AuditReportFurnishDate").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditReportFurnishDate").MergeArea.ClearContents

'        Range("sheet1.AuditorName").MergeArea.Locked = True
'        Range("sheet1.AuditorName").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AuditorName").MergeArea.ClearContents

'        Range("sheet1.AuditorMemNo").MergeArea.Locked = True
'        Range("sheet1.AuditorMemNo").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AuditorMemNo").MergeArea.ClearContents
'
        Range("sheet1.AudFrmName").MergeArea.Locked = True
        Range("sheet1.AudFrmName").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmName").MergeArea.ClearContents
'
'        Range("sheet1.AudFrmRegNo").MergeArea.Locked = True
'        Range("sheet1.AudFrmRegNo").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AudFrmRegNo").MergeArea.ClearContents
'
        Range("sheet1.AudFrmPAN").MergeArea.Locked = True
        Range("sheet1.AudFrmPAN").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmPAN").MergeArea.ClearContents
        

        Range("sheet1.AudFrmAadhaar").MergeArea.Locked = True
        Range("sheet1.AudFrmAadhaar").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmAadhaar").MergeArea.ClearContents

'        Range("sheet1.AuditDate").MergeArea.Locked = True
'        Range("sheet1.AuditDate").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("sheet1.AuditDate").MergeArea.ClearContents
        
    'Newly added by Sai For AY 2024-25 Change------------------------------------------
        Range("Sheet1.AckNumAuditReport").MergeArea.Locked = True
        Range("Sheet1.AckNumAuditReport").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.AckNumAuditReport").MergeArea.ClearContents
        
        'Commented by Aavula For AY 2026-27

'        Range("Sheet1.UDINAuditNum").MergeArea.Locked = True
'        Range("Sheet1.UDINAuditNum").MergeArea.Interior.Color = (&HD8D8D8)
'        Range("Sheet1.UDINAuditNum").MergeArea.ClearContents
      '================================================================================
        
   Else:
       Range("sheet1.AuditReportFurnishDate").MergeArea.Locked = False
       Range("sheet1.AuditReportFurnishDate").MergeArea.Interior.Color = (&HCCFFCC)

'       Range("sheet1.AuditorName").MergeArea.Locked = False
'       Range("sheet1.AuditorName").MergeArea.Interior.Color = (&HCCFFCC)

'       Range("sheet1.AuditorMemNo").MergeArea.Locked = False
'       Range("sheet1.AuditorMemNo").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AudFrmName").MergeArea.Locked = False
       Range("sheet1.AudFrmName").MergeArea.Interior.Color = (&HCCFFCC)
'
'       Range("sheet1.AudFrmRegNo").MergeArea.Locked = False
'       Range("sheet1.AudFrmRegNo").MergeArea.Interior.Color = (&HCCFFCC)
'
       Range("sheet1.AudFrmPAN").MergeArea.Locked = False
       Range("sheet1.AudFrmPAN").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.AudFrmAadhaar").MergeArea.Locked = False
        Range("sheet1.AudFrmAadhaar").MergeArea.Interior.Color = (&HCCFFCC)

'       Range("sheet1.AuditDate").MergeArea.Locked = False
'       Range("sheet1.AuditDate").MergeArea.Interior.Color = (&HCCFFCC)
       
       'Newly added by Sai-------------------------------------------------------------
     
       Range("Sheet1.AckNumAuditReport").MergeArea.Locked = False
       Range("Sheet1.AckNumAuditReport").MergeArea.Interior.Color = (&HCCFFCC)
       
       'Commented by Aavula For AY 2026-27

'       Range("Sheet1.UDINAuditNum").MergeArea.Locked = False
'       Range("Sheet1.UDINAuditNum").MergeArea.Interior.Color = (&HCCFFCC)
     '=================================================================================
       

   End If
 Sheet1.Protect Password:=sPassword
End If

'Newly addded by Sai For 44AB Condition Dropdowns==========================================
        If Not Application.Intersect(Sheet1.Range("Sheet1.Condition_44AB"), Sheet1.Range(Target.Address)) Is Nothing Then
'
              sPassword = EfilingCommon.getmsgstate
               Sheet1.Unprotect Password:=sPassword
            If Sheet1.Range("Sheet1.Condition_44AB").value = "bii-Assessee falling u/s 44AD/44ADA/44AE/44BB but not opting for offering income on presumptive basis" Then

'               Range("Sheet1.bii").ClearContents
'               Range("Sheet1.bii").Locked = False
'               Range("Sheet1.bii").Interior.Color = (&HCCFFCC)
               
               
               Range("Sheet1.bii").MergeArea.Locked = False
               Range("Sheet1.bii").MergeArea.Interior.Color = (&HCCFFCC)
               Range("Sheet1.bii").MergeArea.ClearContents
               
               
               Range("Sheet1.bii_1").MergeArea.Locked = False
               Range("Sheet1.bii_1").MergeArea.Interior.Color = (&HCCFFCC)
               Range("Sheet1.bii_1").MergeArea.ClearContents
               
               
               Range("Sheet1.bii_2").MergeArea.Locked = False
               Range("Sheet1.bii_2").MergeArea.Interior.Color = (&HCCFFCC)
               Range("Sheet1.bii_2").MergeArea.ClearContents
               
               
               Range("Sheet1.bii_3").MergeArea.Locked = False
               Range("Sheet1.bii_3").MergeArea.Interior.Color = (&HCCFFCC)
               Range("Sheet1.bii_3").MergeArea.ClearContents


            Else

                 
                 Range("Sheet1.bii").MergeArea.Locked = True
                 Range("Sheet1.bii").MergeArea.Interior.Color = (&HD8D8D8)


                 Range("Sheet1.bii_1").MergeArea.Locked = True
                 Range("Sheet1.bii_1").MergeArea.Interior.Color = (&HD8D8D8)


                 Range("Sheet1.bii_2").MergeArea.Locked = True
                 Range("Sheet1.bii_2").MergeArea.Interior.Color = (&HD8D8D8)


                 Range("Sheet1.bii_3").MergeArea.Locked = True
                 Range("Sheet1.bii_3").MergeArea.Interior.Color = (&HD8D8D8)

           
          End If
           Sheet1.Protect Password:=sPassword
    End If
'End By Sai=========================================================================




'SIT-73616 Updated by sai on 25/04/2024


 If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("sheet1.LiableSec92Eflg")) Is Nothing Then
 
 
    'If Sheet1.Range("Sheet1.DueDate").value <> "30/11/2024 or extended" And Sheet1.Range("sheet1.LiableSec92Eflg").value = "Y - Yes" Then
    'sadineni 27/10/2025
     If Mid(Sheet1.Range("Sheet1.DueDate").value, 1, 10) <> "30/11/2025" And Sheet1.Range("sheet1.LiableSec92Eflg").value = "Y - Yes" Then
    
     If Trim(Range("sheet1.LiableSec92Eflg").value) <> "" Then
     
'          If Dformat(Trim(Range("sheet1.AuditDateSec92E").value), "yyyy-mm-dd") <> Dformat(Trim("30/11/2024"), "yyyy-mm-dd") Then
 
        fmsgboxStatus "*""Please select correct due date for filing return of income as you have selected Yes for " & " ""Are you liable for Audit u/s 92E """""
 
'            Target.MergeArea.ClearContents
'            Target.MergeArea.Select
'          End If
        End If
    End If
End If





 'For SIT-71390 added by sai on 14/03/2024
    If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("sheet1.LiableSec44ABflg")) Is Nothing Then

'        If Sheet1.Range("Sheet1.DueDate").value = "31/07/2024" And Sheet1.Range("sheet1.LiableSec44ABflg").value = "Y - Yes" Then
        If Sheet1.Range("Sheet1.DueDate").value = "31/08/2026" And Sheet1.Range("sheet1.LiableSec44ABflg").value = "Y - Yes" Then
                fmsgboxStatus "*""Please select correct due date for filing return of income as you have selected Yes for ""Are you liable for audit under section 44AB?"""
         End If
         
    End If
 
    



'newly added by sai For 10IE Date format check===========================================================================

    If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
    If Trim(Range("Sheet1.Form10IEADate").value) <> "00/00/0000" And Range("Sheet1.Form10IEADate").value <> "" Then
            If Not FormatNCheckDate(Trim(Range("Sheet1.Form10IEADate").value)) Then
                fmsgboxStatus """Please enter date in dd/mm/yyyy format""."        ', vbOKOnly, mdCommon.vbMessageCaption
                Sheet1.Range("Sheet1.Form10IEADate").Select
                Sheet1.Range("Sheet1.Form10IEADate").value = ""
                
            End If
    End If
    End If



               'newly added by sai For 10IE Date format prior to 01/04/2024
    If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
         If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
        If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2024"), "yyyy-mm-dd") Then

               fmsgboxStatus "*""Date of filing of  Form 10IEA cannot be prior to 01/04/2024""."
            Sheet1.Range("Sheet1.Form10IEADate").Select
            Sheet1.Range("Sheet1.Form10IEADate").value = ""

        End If
    End If
    End If
    
    
    

    
    
    
    
    
    
    
      'Commented by Aavula For AY 2026-27

''Added by Aavula Naresh
'If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("BAC115.DateOfFiling")) Is Nothing Then
'If Dformat(Trim(Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") > Dformat(Trim("15/12/2024"), "yyyy-mm-dd") And Sheet1.Range("Sheet1.115BAC").value = "Yes" Then
'Sheet1.Unprotect Password:=sPassword
'               Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
'               Range("BAC115.AcknowledgementNo").MergeArea.Interior.Color = (&HD8D8D8)
'               Sheet1.Protect Password:=sPassword
'            Else
'
'            Sheet1.Unprotect Password:=sPassword
'               Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
'               Range("BAC115.AcknowledgementNo").MergeArea.Interior.Color = (&HCCFFCC)
'               Range("BAC115.AcknowledgementNo").MergeArea.ClearContents
'            Sheet1.Protect Password:=sPassword
'End If
'End If
'''''''''


'======================Commented by Aavula Naresh for AY 2025-26======================


''newly added by sai For 10IE Date Logic with respect to Due date, Vice Versa====================================
'
'  If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'     If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
'      If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") <= Dformat(Trim("31/07/2024"), "yyyy-mm-dd") Then
'        If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date" Then
''            MsgBox ("If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly")
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly"
'
'            Target.MergeArea.ClearContents
'            Target.MergeArea.Select
'        End If
'     End If
'  End If
' End If
'
'==============='Commented by Aavula Naresh for AY 2025-26=====================

' If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'       If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
'         If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Trim("31/07/2024"), "yyyy-mm-dd") Then
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
''            MsgBox ("If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly")
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, within due date" & Chr(34) & ", then date of filing of Form should be within due date u/s 139(1). Please enter date of filing correctly"
'
'            Target.MergeArea.ClearContents
'             Target.MergeArea.Select
'          End If
'       End If
'     End If
'End If

'===================Commented by Aavula Naresh for AY 2025-26==============================

'For not greater than system date condition added by sai 06/12/2024
'=======================================================
' If Not Application.Intersect(Sheet1.Range("Sheet1.Form10IEADate"), Sheet1.Range(target.Address)) Is Nothing Then
'        If Sheet1.Range("Sheet1.Form10IEADate").value <> "" Then
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
'                  fmsgboxStatus """Date of filing of  Form 10IEA cannot be greater than system date."
'                  Sheet1.Range("Sheet1.Form10IEADate").value = ""
'           End If
'        End If
' End If

'===========================Commented by Aavula Naresh for AY 2025-26======================================

''For "Yes-within due date" Logic (IPIP-55865)Starts

'For Due date 31/07/2024
' If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'    If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
'       If Sheet1.Range("Sheet1.DueDate").value = "31/07/2024" Then
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Trim("31/07/2024"), "yyyy-mm-dd") Then
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, within due date" & Chr(34) & ", then date of filing of Form should be within due date u/s 139(1). Please enter date of filing correctly"
'
'            target.MergeArea.ClearContents
''             Target.MergeArea.Select
'           End If
'        End If
'      End If
'   End If
'End If

'For Due date 31/10/2024
'===========================Commented by Aavula Naresh for AY 2025-26=========================

' If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'    If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
''       If Sheet1.Range("Sheet1.DueDate").value = "31/10/2024" Then
'    'By sai on 16/10/2024
'       If Sheet1.Range("Sheet1.DueDate").value = "31/10/2024 or extended" Then
'
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
''           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Trim("31/10/2024"), "yyyy-mm-dd") Then
'         'Changed by sai on 17/10/2024
'
'         'Changed by Bindu on 27th - As Due Date extended to 15th Nov
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Trim("15/11/2024"), "yyyy-mm-dd") Then
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, within due date" & Chr(34) & ", then date of filing of Form should be within due date u/s 139(1). Please enter date of filing correctly"
'
'            target.MergeArea.ClearContents
'
'           End If
'        End If
'      End If
'   End If
'End If
'=============================Commented by Aavula Naresh for AY 2025-26================

'For Due date 30/11/2024
' If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'    If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
'
''       If Sheet1.Range("Sheet1.DueDate").value = "30/11/2024" Then
'
'    'Changed by Sai on 06/12/2024 - As Due Date extended to 15th Dec
'If Sheet1.Range("Sheet1.DueDate").value = "30/11/2024 or extended" Then
'
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
''           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Trim("30/11/2024"), "yyyy-mm-dd") Then
'
'        'Changed by Sai on 06/12/2024 - As Due Date extended to 15th Dec
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") > Dformat(Trim("15/12/2024"), "yyyy-mm-dd") Then
'
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, within due date" & Chr(34) & ", then date of filing of Form should be within due date u/s 139(1). Please enter date of filing correctly"
'
'            target.MergeArea.ClearContents
'
'           End If
'        End If
'      End If
'   End If
'End If

'For Yes-but beyond due date Logic

'For Due date 31/07/2024
' If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'    If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
'       If Sheet1.Range("Sheet1.DueDate").value = "31/07/2024" Then
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date" Then
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") <= Dformat(Trim("31/07/2024"), "yyyy-mm-dd") Then
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly"
'
'             target.MergeArea.ClearContents
'
'           End If
'        End If
'      End If
'   End If
'End If

'For Due date 30/11/2024

' If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'    If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
''       If Sheet1.Range("Sheet1.DueDate").value = "31/10/2024" Then
'
'   'By sai on 16/10/2024
'
'       If Sheet1.Range("Sheet1.DueDate").value = "31/10/2024 or extended" Then
'
'
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date" Then
'         'Change by Bindu on 27/11/24 as Due Date extended to  15th Nov
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") <= Dformat(Trim("15/11/2024"), "yyyy-mm-dd") Then
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly"
'
'            target.MergeArea.ClearContents
'
'           End If
'        End If
'      End If
'   End If
'End If

'For Due date 30/11/2024
' If Not Application.Intersect(Sheet1.Range(target.Address), Sheet1.Range("Sheet1.Form10IEADate")) Is Nothing Then
'    If Trim(Range("Sheet1.Form10IEADate").value) <> "" Then
''       If Sheet1.Range("Sheet1.DueDate").value = "30/11/2024" Then
'
'       'Changed by Sai on 06/12/24 as Due Date extended to 15th Dec
'       If Sheet1.Range("Sheet1.DueDate").value = "30/11/2024 or extended" Then
'
'
'         If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date" Then
''           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") <= Dformat(Trim("30/11/2024"), "yyyy-mm-dd") Then
'
'        'Changed by Sai on 06/12/24 as Due Date extended to 15th Dec
'           If Dformat(Trim(Range("Sheet1.Form10IEADate").value), "yyyy-mm-dd") <= Dformat(Trim("15/12/2024"), "yyyy-mm-dd") Then
'
'            fmsgboxStatus "*If option is selected as " & Chr(34) & "Yes, but beyond due date" & Chr(34) & ", then date of filing of Form should be after due date u/s 139(1). Please enter date of filing correctly"
'
'            target.MergeArea.ClearContents
'
'           End If
'        End If
'      End If
'   End If
'End If


'=======================(IPIP-55865) End===========================================================


       'Commented by Aavula For AY 2026-27
       
'If Not Intersect(Range("sheet1.AuditorMemNo"), Range(target.Address)) Is Nothing Then
''   If Sheet1.Range("sheet1.AuditorMemNo").value = 0 Then
''       fmsgboxStatus "*Membership Number should not contain only Zeroes"
''       Sheet1.Range("sheet1.AuditorMemNo").value = ""
''   GoTo endl
''   End If
'   Dim lenofMemNo, zerostoadd As Long
'      lenofMemNo = Len(Sheet1.Range("sheet1.AuditorMemNo").value)
''      If lenofMemNo <> 6 Then
''           zerostoadd = 6 - lenofMemNo
''               If zerostoadd = 1 Then
''                   Sheet1.Range("sheet1.AuditorMemNo").value = "0" + Sheet1.Range("sheet1.AuditorMemNo").value
''               ElseIf zerostoadd = 2 Then
''                   Sheet1.Range("sheet1.AuditorMemNo").value = "00" + Sheet1.Range("sheet1.AuditorMemNo").value
''               ElseIf zerostoadd = 3 Then
''                   Sheet1.Range("sheet1.AuditorMemNo").value = "000" + Sheet1.Range("sheet1.AuditorMemNo").value
''               ElseIf zerostoadd = 4 Then
''                   Sheet1.Range("sheet1.AuditorMemNo").value = "0000" + Sheet1.Range("sheet1.AuditorMemNo").value
''               ElseIf zerostoadd = 5 Then
''                   Sheet1.Range("sheet1.AuditorMemNo").value = "00000" + Sheet1.Range("sheet1.AuditorMemNo").value
''               End If
''               fmsgboxStatus "*Membership Number should not be less than 6 digits, hence it will be prefixed with zero"
''       End If
'
''       If Not chkNumeric(Sheet1.Range("sheet1.AuditorMemNo").value) Then
''       fmsgboxStatus "*Membership Number should not contain special charachters"
''       Sheet1.Range("sheet1.AuditorMemNo").Select
''       Sheet1.Range("sheet1.AuditorMemNo").value = ""
''       End If
'endl:
'End If
''''''''''''''''''''''



'If Not Intersect(Range("sheet1.AudFrmRegNo"), Range(target.Address)) Is Nothing Then
'   If Sheet1.Range("sheet1.AudFrmRegNo").value = 0 Then
'       fmsgboxStatus "*Membership Number should not contain only Zeroes"
'       Sheet1.Range("sheet1.AudFrmRegNo").value = ""
'    GoTo endl1
'   End If
'      lenofMemNo = Len(Sheet1.Range("sheet1.AudFrmRegNo").value)
'      If lenofMemNo <> 8 Then
'           zerostoadd = 8 - lenofMemNo
'               If zerostoadd = 1 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "0" + Sheet1.Range("sheet1.AudFrmRegNo").value
'               ElseIf zerostoadd = 2 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "00" + Sheet1.Range("sheet1.AudFrmRegNo").value
'               ElseIf zerostoadd = 3 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "000" + Sheet1.Range("sheet1.AudFrmRegNo").value
'               ElseIf zerostoadd = 4 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "0000" + Sheet1.Range("sheet1.AudFrmRegNo").value
'               ElseIf zerostoadd = 5 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "00000" + Sheet1.Range("sheet1.AudFrmRegNo").value
'                ElseIf zerostoadd = 6 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "000000" + Sheet1.Range("sheet1.AudFrmRegNo").value
'                ElseIf zerostoadd = 7 Then
'                   Sheet1.Range("sheet1.AudFrmRegNo").value = "0000000" + Sheet1.Range("sheet1.AudFrmRegNo").value
'               End If
'               fmsgboxStatus "*Proprietorship/firm registration number should not be less than 8 digits, hence it will be prefixed with zero"
'               GoTo endl1
'       End If
'
'       If Not CheckSpecialCharacter(Sheet1.Range("sheet1.AudFrmRegNo").value) Then
'       fmsgboxStatus "*Proprietorship/firm registration number should not contain special charachters"
'       Sheet1.Range("sheet1.AudFrmRegNo").Select
'       Sheet1.Range("sheet1.AudFrmRegNo").value = ""
'       GoTo endl1
'       End If
'endl1:
'End If

    If Not Intersect(Range("sheet1.AudFrmPAN"), Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
        If Not Part_A_General.CheckAudFrmPAN(UCase(Trim(Range("sheet1.AudFrmPAN").value))) Then
               fmsgboxStatus "* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule Part A General"""
               Range("sheet1.AudFrmPAN").value = ""
        End If
    End If
    

    If Not Intersect(Range("sheet1.AudFrmAadhaar"), Range(Target.Address)) Is Nothing Then
       Target.value = (Target.value)
        If Not ValidateAadhaar(UCase(Trim(Range("sheet1.AudFrmAadhaar").value))) Then
               fmsgboxStatus "*""Please enter the valid Aadhaar number in schedule Part A General.""" & Chr(13)
               Range("sheet1.AudFrmAadhaar").value = ""
        End If
    End If
    
    If Not Intersect(Range("sheet1.CountryCode"), Range(Target.Address)) Is Nothing Then
      If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"
            'Sheet1.Range("sheet1.PinCode") = ""
'            If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-STATE OUTSIDE INDIA" Then

            If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-Foreign" Then
                fmsgboxStatus "*Country Should not be India for Selected State" & Chr(13)
                Sheet1.Range("sheet1.CountryCode").value = "(Select)"
                Sheet1.Range("sheet1.CountryCode").Select
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = True
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf Sheet1.Range("sheet1.CountryCode").value = "(Select)" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"

            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = True
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
'            If UCase(Sheet1.Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" And Sheet1.Range("sheet1.StateCode1").value <> "(Select)" Then


'            If UCase(Sheet1.Range("sheet1.StateCode1").value) <> "99-Foreign" And Sheet1.Range("sheet1.StateCode1").value <> "(Select)" Then
            
'updated by sadineni on 08/06/2026
            If UCase(Sheet1.Range("sheet1.StateCode1").value) <> UCase("99-Foreign") And Sheet1.Range("sheet1.StateCode1").value <> "(Select)" Then
            
                fmsgboxStatus "*Country Should be India for Selected State" & Chr(13)
                Sheet1.Range("sheet1.CountryCode").value = ""
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = False
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
      End If
    End If
    If Not Intersect(Range("HASZIP"), Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)

        If Mid(Trim(Sheet1.Range("HASZIP").value), 1, 1) = "Y" Then
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.ZipCode").MergeArea.value = "XXXXXX"
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
        End If
    End If
    
    
    
 If Not Intersect(Range("sheet1.CountryCode2"), Range(Target.Address)) Is Nothing Then
      If UCase(Sheet1.Range("sheet1.CountryCode2")) = "91-INDIA" Then
            Sheet1.Range("sheet1.StateCode2") = "(Select)"
            'Sheet1.Range("sheet1.PinCode2") = ""
'            If UCase(Sheet1.Range("sheet1.StateCode2")) = "99-STATE OUTSIDE INDIA" Then

            If UCase(Sheet1.Range("sheet1.StateCode2")) = UCase("99-Foreign") Then
'            If UCase(Sheet1.Range("sheet1.StateCode2").value) <> UCase("99-Foreign") Then
                fmsgboxStatus "*Country Should not be India for Selected State" & Chr(13)
                Sheet1.Range("sheet1.CountryCode2").value = "(Select)"
                Sheet1.Range("sheet1.CountryCode2").Select
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP2").MergeArea.value = ""
               Sheet1.Range("HASZIP2").MergeArea.Locked = True
               Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf Sheet1.Range("sheet1.CountryCode2").value = "(Select)" Then
            Sheet1.Range("sheet1.StateCode2") = "(Select)"

            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP2").MergeArea.value = ""
               Sheet1.Range("HASZIP2").MergeArea.Locked = True
               Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf UCase(Sheet1.Range("sheet1.CountryCode2").value) <> "91-INDIA" Then
'            If UCase(Sheet1.Range("sheet1.StateCode2").value) <> "99-STATE OUTSIDE INDIA" And Sheet1.Range("sheet1.StateCode2").value <> "(Select)" Then

'Updated by sadineni on 08/06/2026
            If UCase(Sheet1.Range("sheet1.StateCode2").value) <> UCase("99-Foreign") And Sheet1.Range("sheet1.StateCode2").value <> "(Select)" Then
            
            'If UCase(Sheet1.Range("sheet1.StateCode2").value) <> "99-Foreign" And Sheet1.Range("sheet1.StateCode2").value <> "(Select)" Then
                fmsgboxStatus "*Country Should be India for Selected State" & Chr(13)
                Sheet1.Range("sheet1.CountryCode2").value = ""
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
               Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("HASZIP2").MergeArea.value = ""
               Sheet1.Range("HASZIP2").MergeArea.Locked = False
               Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
      End If
    End If
    If Not Intersect(Range("HASZIP2"), Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)

        If Mid(Trim(Sheet1.Range("HASZIP2").value), 1, 1) = "Y" Then
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.ZipCode2").MergeArea.value = "XXXXXX"
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
        End If
    End If
    
    
    
'Change-20.01.2023.101.25B
If Not Application.Intersect(Sheet1.Range("Sheet1.SEBI_Regn_No"), Sheet1.Range(Target.Address)) Is Nothing Then
    ValidateSEBI Target
End If

'Change-20.01.2023.101.24B
If (Not Intersect(Sheet1.Range("Sheet1.115H"), Sheet1.Range(Target.Address)) Is Nothing) Or _
    (Not Intersect(Sheet1.Range("Foreign_institutinal_investment"), Sheet1.Range(Target.Address)) Is Nothing) Then
    LockUnlock115AD_FII_115H Target
End If


'Change-16.02.2023.103.25B
If Not Intersect(Sheet1.Range("Foreign_institutinal_investment"), Sheet1.Range(Target.Address)) Is Nothing Then
    LockUnLockSEBI Target
End If

'Change-20.01.2023.101.24C
If Not Intersect(Sheet1.Range("Sheet1.115H"), Sheet1.Range(Target.Address)) Is Nothing Then
    LockUnlock115AD_FII_115H2 Target
End If
    
'    If Not Intersect(Range("Sheet1.115H"), Range(Target.Address)) Is Nothing Then
'        Application.EnableEvents = False
'       Target.value = UCase(Target.value)
'        LockUnlockPartBTTI
'        If Mid(Trim(Sheet1.Range("Sheet1.115H").value), 1, 1) = "Y" Then
''               UnLock115AD
'               UNLOCKcg
'               ResetOSDropDown (Mid(Target.value, 1, 2))
'        Else
'            If Trim(Mid(Sheet1.Range("sheet1.ResidentialStatus1"), 1, 2)) <> "NR" Then
'                LOCKCG
'                ResetOSDropDown (Mid(Target.value, 1, 2))
'            Else
'                UNLOCKcg
'                ResetOSDropDown (Mid(Target.value, 1, 2))
'            End If
'        End If
'
'        If Mid(Trim(Sheet1.Range("Sheet1.115H").value), 1, 1) = "N" Then
'        Lock115AD
'        End If
'    End If
    Sheet19.Unprotect Password:=getmsgstate
    Sheet1.Unprotect Password:=getmsgstate

    If (Mid(Range("sheet1.Status").value, 1, 1) = "" Or Mid(Range("sheet1.Status").value, 1, 1) = "(" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 1) = "" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 1) = "(" Or Range("sheet1.DOB").value = "") Then
              'Commented by sai for 2023-24 old
'            If Sheet1.Range("OptingNewTaxRegime").value = "Not opting" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then

    'Newly modified by Sai For AY 2024-25
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
    
    'Modified by Aavula Naresh For AY 2025-26

'    If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Or Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "Yes" Then
     If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTax").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "Yes" Or Sheet1.Range("Sheet1.Form10IEAB").value = "Yes" Then
                 
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = False
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                 'Chandru
'                  Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = False
'                Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
            End If

    End If

    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then
                 Sheet19.Range("scvia.Section80TTA").value = ""
                 Sheet19.Unprotect Password:=getmsgstate
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = True
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)
        Else
        'Comment by Sai For AY 2023-24
'            If Sheet1.Range("OptingNewTaxRegime").value = "Not opting" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then

              'For AY 2024-25 Modified by Sai
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
    
    'Modified by Aavula Naresh For AY 2025-26
Sheet19.Unprotect Password:=getmsgstate
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Or Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "Yes" Then
If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTax").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "Yes" Or Sheet1.Range("Sheet1.Form10IEAB").value = "Yes" Then
            
            
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = False
                Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
            End If
                 Sheet19.Range("scvia.Section80TTB").value = ""
                 Sheet19.Unprotect Password:=getmsgstate
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
    End If
                Sheet19.Protect Password:=getmsgstate

    Sheet19.Unprotect Password:=getmsgstate
    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then
'                Sheet19.Range("scvia.Section80TTB").value = ""

        'Commented by Sai For AY 2023-24
'               If Sheet1.Range("OptingNewTaxRegime").value = "Not opting" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then
        
        'Commented by Aavula Naresh For AY 2024-25
        'For AY 2024-25 Modified by Sai=====================================================================
'               If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within due date" Then
    
        'Modified by Aavula Naresh For AY 2025-26
               
'                If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "Yes" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "Yes" Or Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "Yes" Then
           If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTax").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "Yes" Or Sheet1.Range("Sheet1.Form10IEAB").value = "Yes" Then
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = False
                    Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                End If
    End If
    If (Mid(Range("sheet1.Status").value, 1, 1) = "I" And Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI") Or (Mid(Range("sheet1.Status").value, 1, 1) = "H") Then
                Sheet19.Range("scvia.Section80TTB").value = ""
                Sheet19.Unprotect Password:=getmsgstate
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
    End If
                Sheet19.Protect Password:=getmsgstate

'sudha commented

'        Dim PinCode As Range
''Change-ANK
'        Application.EnableEvents = False
'
'                 For Each PinCode In SheetALL.Range("PincodeList")
'                      Dim code As String
'                    code = Left(CStr(Sheet1.Range("sheet1.PinCode").value), 3)
'                    If CStr(PinCode.value) = CStr(code) Then
'                        Sheet1.Range("sheet1.StateCode1").value = SheetALL.Range("CX" & PinCode.row).value
'                        Exit For
'                    End If
'                Next PinCode
'        Application.EnableEvents = False
        
        
'sudha commented

'---
'     If Not IsEmpty(Trim(target.value)) Then
'           If (Target.name.name) = "sheet1.StateCode1" Then
           If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
               If Not Part_A_General.CheckState(UCase(Trim(Range("sheet1.StateCode1").value)), validateType) Then
                   If validateType = "1" Then fmsgboxStatus "*State in Sheet : PART A-GENERAL  is Mandatory" & Chr(13) ', vbOKOnly
               End If
           End If
           
           
           If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
               If Not Part_A_General.CheckState2(UCase(Trim(Range("sheet1.StateCode2").value)), validateType) Then
                   If validateType = "1" Then fmsgboxStatus "*State in Sheet : PART A-GENERAL  is Mandatory" & Chr(13) ', vbOKOnly
               End If
           End If
           
'    End If
    If Not Application.Intersect(Sheet1.Range("sheet1.MiddleName"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("sheet1.FirstName"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("sheet1.SurNameOrOrgName"), Sheet1.Range(Target.Address)) Is Nothing Then
'    If (Target.name.name) = "sheet1.MiddleName" Or (Target.name.name) = "sheet1.FirstName" Or (Target.name.name) = "sheet1.SurNameOrOrgName" Then
       Sheet52.Range("sheet9.AssesseeVerName").value = UCase(Sheet31.Range("FullName").value)
    End If

'   Change-22.12.2022.102.029PAG
   'Commented by Aavula For AY 2026-27
If Not Application.Intersect(Sheet1.Range("sheet1.AuditReportFurnishDate"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Not Target.value = "" Then
             If Not CheckFurnishAuditDate_Before(Target.value, "Date of furnishing of audit report") Then
                    Target.value = ""
                    Target.Select
             End If
        End If
        
               
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
        MsgBox "Date of furnishing of audit report  in sheet Audit information should not be greater than system date."
        Target.value = ""
        Target.Select
        End If

' Added by sudha as per SIT-91484
 
    If Target.value <> "" Then
    If Mid(Target.value, 3, 1) <> "/" Or Mid(Target.value, 6, 1) <> "/" Then
       MsgBox "Please enter valid date in dd/mm/yyyy format in schedule Part A General", , "Format ( / )"
       Target.value = ""
       End If
       End If

 '-----------------------------------------------


End If
'end Change=======

'Commented by Naresh - 12/03/2026
'Change-23.01.2023.102.033PAG
'If Not Application.Intersect(Sheet1.Range("sheet1.AuditDate"), Sheet1.Range(target.Address)) Is Nothing Then
'        If Not target.value = "" Then
'             If Not CheckAuditDate_Before(target.value, "Date of audit reporting") Then
'                    target.value = ""
'                    target.Select
'             End If
'        End If

'        If Dformat(target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
'        MsgBox "Invalid date: date cannot be greater than date of filing of the return"
'        target.value = ""
'        target.Select
'        End If
'End If
'end Change=======
'''''''''''''''''''''''''''''''''''''''


'Change-23.01.2023.102.036PAG
If Not Application.Intersect(Sheet1.Range("sheet1.AuditDateSec92E"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Not Target.value = "" Then
             If Not CheckAuditDate_Before(Target.value, "Date of audit report") Then
                    Target.value = ""
                    Target.Select
             End If
        End If
        
               
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
        MsgBox "Invalid date: date cannot be greater than date of filing of the return"
        Target.value = ""
        Target.Select
        End If
End If
'end Change=======
'Change-23.01.2023.102.039PAG
If Not Application.Intersect(Sheet1.Range("sheet1.DateOfAudit"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Not Target.value = "" Then
             If Not CheckAuditDate_Before_diii(Target.value, "Date of audit report") Then
                    Target.value = ""
                    Target.Select
             End If
        End If
        
               
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
        MsgBox "Invalid date: date cannot be greater than date of filing of the return"
        Target.value = ""
        Target.Select
        End If
End If
'end Change=======
'Change-23.01.2023.102.039PAG
If Not Application.Intersect(Sheet1.Range("Audit.Date"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Not Target.value = "" Then
             If Not CheckAuditDateLastTable_Before(Target.value, "Date of audit report") Then
                    Target.value = ""
                    Target.Select
             End If
        End If
'        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
'        MsgBox "Invalid date: date cannot be greater than date of filing of the return"
'        Target.value = ""
'        Target.Select
'        End If
               

End If


'======================


'sudha_2025-26

If Not Application.Intersect(Sheet1.Range("sheet1.EmailAddress"), Sheet1.Range(Target.Address)) Is Nothing Then

If Not ValidateEmailNoConsecutiveSpecialChars(Sheet1.Range("sheet1.EmailAddress")) Then
MsgBox "Consecutive 2 special characters is not allowed in Email Address-1."
Target.value = ""
End If
End If

If Not Application.Intersect(Sheet1.Range("sheet1.EmailAddressSec"), Sheet1.Range(Target.Address)) Is Nothing Then
If Not ValidateEmailNoConsecutiveSpecialChars(Sheet1.Range("sheet1.EmailAddressSec")) Then
MsgBox "Consecutive 2 special characters is not allowed in Email Address-2."
Target.value = ""


End If
End If


'end Change=======

'Added by Aavula Naresh for AY 2026-27

If Not Application.Intersect(Sheet1.Range("Sheet1.RepEmail"), Sheet1.Range(Target.Address)) Is Nothing Then

If Not ValidateEmailNoConsecutiveSpecialChars(Sheet1.Range("Sheet1.RepEmail")) Then
MsgBox "Consecutive 2 special characters is not allowed."
Target.value = ""
End If
End If

'Added by Aavula Naresh for AY 2026-27



'If Not Application.Intersect(Sheet1.Range("Alternate_Address"), Sheet1.Range(target.Address)) Is Nothing Then
'    LockUnlock_Alternate_Address
'End If

If Not Application.Intersect(Sheet1.Range("Secondary_Address"), Sheet1.Range(Target.Address)) Is Nothing Then
    AutoPopulateSecondaryAddress
End If


If Not Application.Intersect(Sheet1.Range("Sheet1.BusiProfAssyr"), Sheet1.Range(Target.Address)) Is Nothing Then
    BusiProfAssyr_26_27
End If

If Not Intersect(Range("Sheet1.Form10IEANewTax"), Range(Target.Address)) Is Nothing Then
            Application.ScreenUpdating = False
            
        If UCase(Mid(Target.value, 1, 1)) = "Y" Then
                MsgBox """Furnish ITR in New tax regime."""
            ElseIf UCase(Mid(Target.value, 1, 1)) = "N" Then
                MsgBox """Furnish ITR in Old tax regime."""
        End If
End If

If Not Intersect(Range("Sheet1.Form10IEANewTaxIB"), Range(Target.Address)) Is Nothing Then
            Application.ScreenUpdating = False
        If UCase(Mid(Target.value, 1, 1)) = "Y" Then
                 MsgBox """Furnish ITR in Old tax regime."""
            ElseIf UCase(Mid(Target.value, 1, 1)) = "N" Then
                 MsgBox """Furnish ITR in New tax regime."""
        End If
End If

'''''''''''''''''''''''''''''''''''''''''

    If Not Intersect(Range("sheet1.RepMobileNo"), Range(Target.Address)) Is Nothing Then
             If Not (Part_A_General.RepChkMobileNo()) Then
                   Sheet1.Range("sheet1.RepMobileNo").value = ""
'                   Sheet1.Range("sheet1.RepMobileNo").Select

            End If
        End If

     If Not Intersect(Range("sheet1.Repcountrycode"), Range(Target.Address)) Is Nothing Then
'            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" Then
            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-Foreign" Then
               If Len(Sheet1.Range("sheet1.Repcountrycode").value) > 5 Then
                  fmsgboxStatus ("countrycode for Mobile No 1 in Sheet : Part A-General cannot exceed 5 characters ")
                  Sheet1.Range("sheet1.Repcountrycode").MergeArea.value = ""
                End If
     End If

      End If
      
      'Added by Aavula Naresh for AY 2026-27 as per 0.2V
      If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("sheet1.LiableSecAggrigateflg")) Is Nothing Then
        If Sheet1.Range("sheet1.LiableSec44AEflg").value = "No" And Sheet1.Range("sheet1.LiableSec5cflg").value = "More than Rs. 1 crore and up to Rs. 10 crores" And Sheet1.Range("sheet1.LiableSecAggrigateflg").value = "More than 5%" Then

                fmsgboxStatus "*""Please check your liability for audit as per the inputs provided at sl. no. a2i, a2ii or a2iii in Schedule Part A General Information"""
           'Malli_SIT-117030  09/06/2026
           'Sheet1.Range("sheet1.LiableSecAggrigateflg").MergeArea.value = "(Select)"
           '----------------------------
        End If
    End If
     
       If Not Application.Intersect(Sheet1.Range(Target.Address), Sheet1.Range("sheet1.LiableSecAggrigate1flg")) Is Nothing Then
        If Sheet1.Range("sheet1.LiableSec44AEflg").value = "No" And Sheet1.Range("sheet1.LiableSec5cflg").value = "More than Rs. 1 crore and up to Rs. 10 crores" And Sheet1.Range("sheet1.LiableSecAggrigate1flg").value = "More than 5%" Then

                fmsgboxStatus "*""Please check your liability for audit as per the inputs provided at sl. no. a2i, a2ii or a2iii in Schedule Part A General Information"""
           'Malli_SIT-117030 09/06/2026
           'Sheet1.Range("sheet1.LiableSecAggrigate1flg").MergeArea.value = "(Select)"
           '---------------------------
        End If
    End If
    
ID_Nature_Employment
 
        
endline:
Sheet1.Activate
Application.EnableEvents = True
Application.ScreenUpdating = True
Sheet1.Protect Password:=getmsgstate
ThisWorkbook.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Changedpk(ByVal Target As Range)

LockUnlockPLSec44   'added by Chetan C M on 27/01/2026 for AY 2026-27

'On Error GoTo endline
CalculateTaxCheckFlag = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Dim validateType As String
validateType = ""
Sheet1.Unprotect Password:=getmsgstate
If Not Application.Intersect(Sheet1.Range("NRI_Juridiction"), Sheet1.Range(Target.Address)) Is Nothing Then
If 1 < Application.CountIf(Range("NRI_Juridiction"), Target.value) Then
             fmsgboxStatus "* Same dropdown cannot be selected twice "
             Range(Target.Address).value = "(Select)"
        End If
End If

If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Target.value = "H-HUF" Then
        Sheet4.Unprotect Password:=getmsgstate
        Sheet4.Range("PL_62i").value = ""
        Sheet4.Range("PL_62i").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("PL_62i").MergeArea.Locked = True
        Sheet4.Range("PL_62ii").value = ""
        Sheet4.Range("PL_62ii").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("PL_62ii").MergeArea.Locked = True
        Sheet4.Range("Sec44AD.NOBName1").value = ""
        Sheet4.Range("Sec44AD.NOBName1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("Sec44AD.NOBName1").Locked = True
        Sheet4.Range("Sec44AD.NOBCode1").value = ""
        Sheet4.Range("Sec44AD.NOBCode1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("Sec44AD.NOBCode1").Locked = True
        Sheet4.Range("Sec44AD.NOBDescription1").value = ""
        Sheet4.Range("Sec44AD.NOBDescription1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet4.Range("Sec44AD.NOBDescription1").Locked = True
        Sheet4.Protect Password:=getmsgstate
        Sheet9.Unprotect Password:=getmsgstate
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").value = ""
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HD8D8D8) ', "AU"
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Locked = True
        Sheet9.Protect Password:=getmsgstate
        Sheet4.Protect Password:=getmsgstate
        
        
        
'Commented by sudha_2025-26

'        Sheet19.Unprotect Password:=getmsgstate
'        Sheet19.Range("scvia.Section80E").value = ""
'        Sheet19.Range("scvia.Section80E").Interior.Color = (&HD8D8D8) ', "AU"
'        Sheet19.Range("scvia.Section80E").Locked = True
'        Sheet19.Range("scvia.Section80EE").value = ""
'        Sheet19.Range("scvia.Section80EE").Interior.Color = (&HD8D8D8) ', "AU"
'        Sheet19.Range("scvia.Section80EE").Locked = True
'        Sheet19.Protect Password:=getmsgstate
        
        LockUnlockESOP
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUS").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUS").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUS").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUK").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUK").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountUK").MergeArea.Locked = True
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountCan").MergeArea.value = ""
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountCan").MergeArea.Interior.Color = (&HD8D8D8)
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncomeNotified89A_AmountCan").MergeArea.Locked = True
'        Sheet14.Unprotect Password:=getmsgstate
'        Sheet14.Range("OSIncomeNotifiedOther89A").MergeArea.value = ""
'        Sheet14.Unprotect Password:=getmsgstate
'        Sheet14.Range("OSIncomeNotifiedOther89A").MergeArea.Interior.Color = (&HD8D8D8)
'        Sheet14.Unprotect Password:=getmsgstate
'        Sheet14.Range("OSIncomeNotifiedOther89A").MergeArea.Locked = True
        Sheet14.Protect Password:=getmsgstate
        
    Else

        Sheet4.Unprotect Password:=getmsgstate
        Sheet4.Range("PL_62i").value = ""
        Sheet4.Range("PL_62i").Interior.Color = (&HCCFFCC)  ', "AU"
        Sheet4.Range("PL_62i").MergeArea.Locked = False
        Sheet4.Range("PL_62ii").value = ""
        Sheet4.Range("PL_62ii").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("PL_62ii").MergeArea.Locked = False
        Sheet4.Range("Sec44AD.NOBName1").value = ""
        Sheet4.Range("Sec44AD.NOBName1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("Sec44AD.NOBName1").Locked = False
         Sheet4.Range("Sec44AD.NOBCode1").value = ""
        Sheet4.Range("Sec44AD.NOBCode1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("Sec44AD.NOBCode1").Locked = False
         Sheet4.Range("Sec44AD.NOBDescription1").value = ""
        Sheet4.Range("Sec44AD.NOBDescription1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet4.Range("Sec44AD.NOBDescription1").Locked = False
        Sheet4.Protect Password:=getmsgstate
        Sheet9.Unprotect Password:=getmsgstate
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").value = ""
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet9.Range("sheet11.ExpDebToPLOthHeads1").Locked = False
        Sheet9.Protect Password:=getmsgstate
       
'Commented by sudha_2025-26

'        Sheet19.Unprotect Password:=getmsgstate
'        Sheet19.Range("scvia.Section80E").value = ""
'        Sheet19.Range("scvia.Section80E").Interior.Color = (&HCCFFCC) ', "AU"
'        Sheet19.Range("scvia.Section80E").Locked = False
'        Sheet19.Range("scvia.Section80EE").value = ""
'        Sheet19.Range("scvia.Section80EE").Interior.Color = (&HCCFFCC) ', "AU"
'        Sheet19.Range("scvia.Section80EE").Locked = False
'        Sheet19.Protect Password:=getmsgstate
        
        
        LockUnlockESOP
        Sheet14.Unprotect Password:=getmsgstate
        Sheet14.Range("OSIncreliefus89A").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncreliefus89A").Locked = False
        Sheet14.Range("OSIncomeNotified89A_AmountUS").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeNotified89A_AmountUS").Locked = False
        Sheet14.Range("OSIncomeNotified89A_AmountUK").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeNotified89A_AmountUK").Locked = False
        Sheet14.Range("OSIncomeNotified89A_AmountCan").Interior.Color = (&HCCFFCC) ', "AU"
        Sheet14.Range("OSIncomeNotified89A_AmountCan").Locked = False
'        Sheet14.Range("OSIncomeNotifiedOther89A").Interior.Color = (&HCCFFCC) ', "AU"
'        Sheet14.Range("OSIncomeNotifiedOther89A").Locked = False
        Sheet14.Protect Password:=getmsgstate
    End If
    
End If
If Not Application.Intersect(Sheet1.Range("sheet1.ResidentialStatus1"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("sheet1.ResidentialStatus1").value, 1, 1) = "R") Then
        Range("sheet1.NRI_SEP").value = "NA"
        Range("sheet1.NRI_SEP").MergeArea.Locked = True
        Range("sheet1.NRI_SEP").MergeArea.Interior.Color = (&HD8D8D8)
Else
        Range("sheet1.NRI_SEP").MergeArea.Locked = False
        Range("sheet1.NRI_SEP").MergeArea.Interior.Color = (&HCCFFCC)
        Range("sheet1.NRI_SEP").value = "(Select)"
End If
Sheet1.Protect Password:=getmsgstate
End If
Sheet1.Unprotect Password:=getmsgstate
If Not Application.Intersect(Sheet1.Range("Sheet1.Foreign_exchange"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet22.Unprotect Password:=getmsgstate
If Range("Sheet1.Foreign_exchange").value = "No" Then
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").value = 0
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").Interior.Color = (&HD8D8D8)
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").Locked = True
Else
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").MergeArea.Locked = False
        Sheet22.Range("AMT.AdjustedUnderSec115JC3a").MergeArea.Interior.Color = (&HCCFFCC)
        
End If
Sheet1.Unprotect Password:=getmsgstate
Sheet22.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg").value, 1, 1) = "N") Or Range("clauseiv7provisio139iFlg").value = "(Select)" Then

        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)

ElseIf (Mid(Range("clauseiv7provisio139iFlg").value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_1").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_2").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_3").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_4").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)

End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_1"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_1").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_2"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_2").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_3"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_3").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_4"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 1) = "Y") Or (Mid(Range("clauseiv7provisio139iFlg_4").value, 1, 2) = "(S") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_1"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_1").value < 6000000 Or Sheet1.Range("clauseiv7provisio139iAmount_1").value = 6000000 Or Sheet1.Range("clauseiv7provisio139iAmount_1").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if ""total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_1").value = ""
        End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_2"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_2").value < 1000000 Or Sheet1.Range("clauseiv7provisio139iAmount_2").value = 1000000 Or Sheet1.Range("clauseiv7provisio139iAmount_2").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if the total gross receipts of the person in profession is less than or equal to ten lakh rupees during the previous year" & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_2").value = ""
        End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_3"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_3").value < 25000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_3").value = ""
        End If
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_4"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_4").value < 5000000 Or Sheet1.Range("clauseiv7provisio139iAmount_4").value = "") Then
            fmsgbox "* Please select drop down as ""No"" if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_4").value = ""
        End If
End If

If Not Application.Intersect(Sheet1.Range("OptingNewTaxRegime"), Sheet1.Range(Target.Address)) Is Nothing Then
     If (Target.value = "Continue to opt") Or (Target.value = "Opting in now") Then
      NewBacYesValueChange
     Else
      NewBacNoValueChange
    End If
  End If
 
If Not Application.Intersect(Sheet1.Range("sheet1.ReturnFileSec"), Sheet1.Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
Sheet52.Unprotect Password:=getmsgstate
    If Sheet1.Range(Target.Address).value = "139(4)- After due date" Then
        Dim vdate As Variant
        vdate = Sheet52.Range("sheet9.Date").value
    If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2024-08-01") Then 'Chandru
        fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2022 or extended due date"
        Range("sheet1.ReturnFileSec").value = "(Select)"
    End If
    End If
    Sheet1.Protect Password:=getmsgstate
    Sheet52.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("Audit.Act"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Sheet1.Range(Target.Address).value = "Others" Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(Target.Address, "F", "O")).value = ""
            Sheet1.Range(Replace(Target.Address, "F", "O")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(Target.Address, "F", "O")).MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(Target.Address, "F", "O")).value = ""
            Sheet1.Range(Replace(Target.Address, "F", "O")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(Target.Address, "F", "O")).MergeArea.Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If
If Not Intersect(Range("NRI_Juridiction"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Range(Target.Address).value) = "9998-Not Applicable (Not Resident in any Country)" Then
          fmsgboxStatus "*In field Tax Identification Number,""Please enter Passport No. along with the name of Country in which Passport is issued""" & Chr(13)
End If
End If
        If Not Intersect(Range("Sheet1.CompPAN"), Range(Target.Address)) Is Nothing Then
                 Target.value = UCase(Target.value)
        End If
         If Not Intersect(Range("Sheet1.FirmPAN"), Range(Target.Address)) Is Nothing Then
                 Target.value = UCase(Target.value)
        End If
If Not Intersect(Range("sheet1.Status"), Range(Target.Address)) Is Nothing Then
            Dim i As Long
            For i = 1 To Sheet24.Range("EI.NatureOfIncome").count
                If Mid(Sheet1.Range(Target.name.name).value, 1, 1) = "H" Or Mid(Sheet1.Range(Target.name.name).value, 1, 1) = "(" Then
                If Mid(Sheet24.Range("EI.NatureOfIncome").Cells.item(i).value, 1, 11) = "Sec 10(12A)" Or Mid(Sheet24.Range("EI.NatureOfIncome").Cells.item(i).value, 1, 11) = "Sec 10(12B)" Then
                Sheet24.Range("EI.NatureOfIncome").Cells.item(i).value = "(Select)"
                End If
                End If
            Next
        Application.EnableEvents = False
                        StatusBasedDefaultValue
                        LockUnlock80D
                        LockUnlockQVIA
            LockUnlockAadhaar
            LockUnlockClaim115H
            HideUnhideCheck
            LockUnlockSalary
            LockUnlockTDS
              NewResidentialwork
            If Mid(Sheet1.Range(Target.name.name).value, 1, 1) = "H" Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet9.Unprotect Password:=getmsgstate
                Sheet1.Range("sheet1.PortugeseCC5A").value = ""
                Sheet1.Range("sheet1.PortugeseCC5A").Interior.Color = (&HD8D8D8)
                Sheet1.Range("sheet1.PortugeseCC5A").MergeArea.Locked = True
                LOCK5A
                LOCK80EE_A_B
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").value = ""
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HD8D8D8)
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").MergeArea.Locked = True

                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").value = ""
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").Interior.Color = (&HD8D8D8)
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").MergeArea.Locked = True
            Sheet1.Protect Password:=getmsgstate
            Sheet9.Protect Password:=getmsgstate
            Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet9.Unprotect Password:=getmsgstate
                Sheet1.Range("sheet1.PortugeseCC5A").value = "No"
                Sheet1.Range("sheet1.PortugeseCC5A").Interior.Color = (&HCCFFCC)
                Sheet1.Range("sheet1.PortugeseCC5A").MergeArea.Locked = False
                LOCK5A
                UNLOCK80EE_A_B

                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").value = ""
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").Interior.Color = (&HCCFFCC)
                Sheet9.Range("Sheet11.ExpDebToPLOthHeads1").MergeArea.Locked = False

                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").value = ""
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").Interior.Color = (&HCCFFCC)
                Sheet9.Range("sheet10.IncRecCredPLOthHeads1").MergeArea.Locked = False

            Sheet1.Protect Password:=getmsgstate
            Sheet9.Protect Password:=getmsgstate
            End If
            If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) <> "I" Then
                Sheet52.Range("sheet9.PAN").value = ""
            Else
                Sheet52.Range("sheet9.PAN").value = Sheet1.Range("sheet1.PAN").value
            End If
             Application.EnableEvents = True
        End If
If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then
            LockUnlock115AD
End If

If Not Intersect(Range("IsDirectorCompany"), Range(Target.Address)) Is Nothing Then
            LockUnlockCheck
End If
If Not Intersect(Range("IsPartnerFirm"), Range(Target.Address)) Is Nothing Then
            LockUnlockFirm
End If
If Not Intersect(Range("IsShareOwner"), Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Trim(Sheet1.Range("IsShareOwner").value), 1, 1)) = "Y" Then
        UnLockShares
    Else
        LockShares
    End If
End If

       If Not Intersect(Range("sheet1.FirstName"), Range(Target.Address)) Is Nothing Then
           Sheet1.Range("sheet1.FirstName").value = UCase(Sheet1.Range("sheet1.FirstName").value)
       End If
       If Not Intersect(Range("sheet1.MiddleName"), Range(Target.Address)) Is Nothing Then
           Sheet1.Range("sheet1.MiddleName").value = UCase(Sheet1.Range("sheet1.MiddleName").value)
       End If
       If Not Intersect(Range("sheet1.SurNameOrOrgName"), Range(Target.Address)) Is Nothing Then
           Sheet1.Range("sheet1.SurNameOrOrgName").value = UCase(Sheet1.Range("sheet1.SurNameOrOrgName").value)
       End If
       'dpkpd
        If Not Intersect(Range("sheet1.PAN"), Range(Target.Address)) Is Nothing Then
               Dim TargetPAN As String
               TargetPAN = ""
               TargetPAN = Target.name.name
                Sheet1.Range(TargetPAN).value = UCase(Range(TargetPAN).value)
                   If Not Part_A_General.CheckPAN(UCase(Trim(Range(TargetPAN).value))) Then
                        fmsgboxStatus " Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."
                        Sheet1.Range(TargetPAN).Select
                        Sheet1.Range(TargetPAN).value = ""
                    Else
                       If Mid(Range(TargetPAN).value, 4, 1) = "P" Then
                            Sheet52.Range("sheet9.PAN").value = UCase(Range(TargetPAN).value)
                       End If

                   End If
           End If
           If Not Intersect(Range("sheet1.RepPAN"), Range(Target.Address)) Is Nothing Then
                   Dim TargetRepPAN As String
                   TargetRepPAN = ""
                   TargetRepPAN = Target.name.name
               If Not Part_A_General.CheckRepPAN(UCase(Trim(Range(TargetRepPAN).value))) Then
                   fmsgboxStatus " Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule Part A General" ', vbOKCancel
                   Sheet1.Range(TargetRepPAN).Select
                   Sheet1.Range(TargetRepPAN).value = ""

               End If
               If Range(TargetRepPAN).value = Sheet1.Range("Sheet1.pan").value Then
                    fmsgboxStatus "PAN of the representative cannot be same as assessee PAN in Part-A General Information." ', vbOKCancel
                    Sheet1.Range(TargetRepPAN).Select
                    Sheet1.Range(TargetRepPAN).value = ""
               End If
               Sheet1.Range(TargetRepPAN).value = UCase(Range(TargetRepPAN).value)
               Sheet52.Range("sheet9.PAN").value = UCase(Range(TargetRepPAN).value)
           End If
           If Not Intersect(Range("sheet1.PinCode"), Range(Target.Address)) Is Nothing Then
               If Not (Part_A_General.ValidatePinCode()) Then
                   Sheet1.Range("sheet1.PinCode").value = ""
                   Sheet1.Range("sheet1.PinCode").Select
               Else
                   If (Target.value) = "999999" Then
                   If Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" Then
                       If Mid(Sheet1.Range("sheet1.StateCode1"), 1, 2) <> "99" Then
                           Sheet1.Range("sheet1.PinCode").value = ""
                           fmsgboxStatus " INVALID PINCODE"
                           Sheet1.Range("sheet1.PinCode").Select
                       End If
                   End If
                   End If
               End If
           End If
        If Not Intersect(Range("sheet1.capacityassessee"), Range(Target.Address)) Is Nothing Then
               If Trim(Range("sheet1.capacityassessee").value) <> "00/00/0000" Then
                   If Not CheckDate10IFBefore(Trim(Range("sheet1.capacityassessee").value)) Then
                   
                       Sheet1.Range("sheet1.capacityassessee").Select
                       Sheet1.Range("sheet1.capacityassessee").value = ""
                   End If
               End If
           End If
           If Not Intersect(Range("sheet1.DOB"), Range(Target.Address)) Is Nothing Then
               If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
                   If Not Part_A_General.CheckDOB(Trim(Range("sheet1.DOB").value)) Then
                       Sheet1.Range("sheet1.DOB").Select
                       Sheet1.Range("sheet1.DOB").value = ""
                   Else
                     Sheet1.Range("sheet1.DOB").value = AssesseeDob
                   End If
               End If
           End If
           
            If Not Intersect(Range("sheet1.DOC"), Range(Target.Address)) Is Nothing Then
               If Trim(Range("sheet1.DOC").value) <> "00/00/0000" Then
                   If Not Part_A_General.CheckDOC(Trim(Range("sheet1.DOC").value)) Then
                       Sheet1.Range("sheet1.DOC").Select
                       Sheet1.Range("sheet1.DOC").value = ""
                   End If
                   If Not ChkMinInclusiveDate(Dformat(Range("sheet1.DOC").value, ""), Dformat(Range("sheet1.DOB").value, "")) Then
                         MsgBox ("Date of formation of business field in Sheet : PART A-GENERAL  should be on or before Date of Commencement")
                         Range("sheet1.DOC").Select
                       Range("sheet1.DOC").value = ""
                       Range("sheet1.DOB").value = ""
                   Else
                   End If
               End If
           End If
          If Not Intersect(Range("sheet1.DOB"), Range(Target.Address)) Is Nothing Then
               If Trim(Range("sheet1.DOB").value) <> "00/00/0000" Then
                    If Not ChkMinInclusiveDate(Dformat(Range("sheet1.DOC").value, ""), Dformat(Range("sheet1.DOB").value, "")) Then
                         MsgBox ("Date of formation of business field in Sheet : PART A-GENERAL  should be on or before Date of Commencement")
                         Range("sheet1.DOC").Select
                       Range("sheet1.DOC").value = ""
                       Range("sheet1.DOB").value = ""
                     Else
                    End If
               End If
           End If
           If Not Intersect(Range("sheet1.EmailAddress"), Range(Target.Address)) Is Nothing Then
               If Not Part_A_General.CheckEmailAddress(Trim(Range("sheet1.EmailAddress").value), validateType) Then

                   If validateType = "1" Then fmsgboxStatus "*Email ID in Sheet PART A-GENERAL , contains special characters"
                   If validateType = "2" Then fmsgboxStatus "*EmailAddress-1 in Sheet : PART A-GENERAL  should not exceed 125 characters " ', vbOKOnly
                   If validateType = "3" Then fmsgboxStatus "*EmailAddress-1 in Sheet : PART A-GENERAL  is invalid" ', vbOKOnly

                   Range("sheet1.EmailAddress").Select
                   Range("sheet1.EmailAddress").value = ""
               End If

           End If
           If Not Intersect(Range("sheet1.EmailAddressSec"), Range(Target.Address)) Is Nothing Then
               If Not Part_A_General.CheckEmailAddress(Trim(Range("sheet1.EmailAddressSec").value), validateType) Then

                   If validateType = "1" Then fmsgboxStatus "*Email ID in Sheet PART A-GENERAL , contains special characters" ', vbOKOnly
                   If validateType = "2" Then fmsgboxStatus "*EmailAddress-2 in Sheet : PART A-GENERAL  should not exceed 125 characters " ', vbOKOnly
                   If validateType = "3" Then fmsgboxStatus "*EmailAddress-2 in Sheet : PART A-GENERAL  is invalid" ', vbOKOnly

                   Range("sheet1.EmailAddressSec").Select
                   Range("sheet1.EmailAddressSec").value = ""
               End If

           End If
              If Not Intersect(Range("sheet1.STDcode"), Range(Target.Address)) Is Nothing Then
                If Not (Part_A_General.ValidateSTDcode()) Then
                    Sheet1.Range("sheet1.STDcode").value = ""
                    Sheet1.Range("sheet1.STDcode").Select
                    GoTo endline
                End If
                If Sheet1.Range("sheet1.PhoneNo").value <> "" Then ValidatePhoneNo
                STDcode = Sheet1.Range("sheet1.STDcode").value
              End If

            If Not Intersect(Range("sheet1.PhoneNo"), Range(Target.Address)) Is Nothing Then
                If Part_A_General.ValidateSTDcode Then
                   If Not (Part_A_General.ValidatePhoneNo()) Then
                       GoTo endline
                   End If
                Else
                   Sheet1.Range("sheet1.PhoneNo").value = ""
                   Sheet1.Range("sheet1.STDcode").Select
                   GoTo endline
                End If
                PhoneNo = Sheet1.Range("sheet1.PhoneNo").value
            End If
            
            If Not Intersect(Range("Countrycode1"), Range(Target.Address)) Is Nothing Then
            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" Then
            If Len(Sheet1.Range("Countrycode1").MergeArea.value) > 5 Then
            fmsgboxStatus ("Country Code for Mobile No 1 in Sheet : Part A-General cannot exceed 5 characters ")
            Sheet1.Range("Countrycode1").MergeArea.value = ""
            End If
            End If

            If Sheet1.Range("Countrycode1").MergeArea.value = "0" Or Sheet1.Range("Countrycode1").MergeArea.value = "00" Or Sheet1.Range("Countrycode1").MergeArea.value = "000" Or Sheet1.Range("Countrycode1").MergeArea.value = "0000" Or Sheet1.Range("Countrycode1").MergeArea.value = "00000" Then
            fmsgboxStatus ("Please enter valid Country Code for Mobile No 1 in Sheet : Part A-General ")
            Sheet1.Range("Countrycode1").MergeArea.value = ""
            End If

            End If
            
            If Not Intersect(Range("Countrycode2"), Range(Target.Address)) Is Nothing Then
            If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Or UCase(Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" Then
            If Len(Sheet1.Range("Countrycode2").MergeArea.value) > 5 Then
            fmsgboxStatus ("Country Code for Mobile No 2 in Sheet : Part A-General cannot exceed 5 characters ")
            Sheet1.Range("Countrycode2").MergeArea.value = ""
            End If
            End If

            If Sheet1.Range("Countrycode2").MergeArea.value = "0" Or Sheet1.Range("Countrycode2").MergeArea.value = "00" Or Sheet1.Range("Countrycode2").MergeArea.value = "000" Or Sheet1.Range("Countrycode2").MergeArea.value = "0000" Or Sheet1.Range("Countrycode2").MergeArea.value = "00000" Then
            fmsgboxStatus ("Please enter valid Country Code for Mobile No 2 in Sheet : Part A-General ")
            Sheet1.Range("Countrycode2").MergeArea.value = ""
            End If

            End If

        If Not Intersect(Range("sheet1.MobileNo"), Range(Target.Address)) Is Nothing Then
             If Not (Part_A_General.ChkMobileNo()) Then
                   Sheet1.Range("sheet1.MobileNo").value = ""
                   Sheet1.Range("sheet1.MobileNo").Select

            End If
        End If
        If Not Intersect(Range("sheet1.MobileNoSec"), Range(Target.Address)) Is Nothing Then
             If Not (Part_A_General.ValidateMobileNoSec()) Then
                   Sheet1.Range("sheet1.MobileNoSec").value = ""
                   Sheet1.Range("sheet1.MobileNoSec").Select

            End If
        End If

If Not Intersect(Range("sheet1.SeventhProvisoFlag"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)

ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = False
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.DepositAmountFlag").MergeArea.value = "(Select)"

        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
        
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg").MergeArea.value = "(Select)"
        
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_1").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_2").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_3").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_4").MergeArea.value = "(Select)"

        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)

        
End If

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag").MergeArea.value = "(Select)"

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If

If (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").value = "(Select)" Then

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)


ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.value = "(Select)"

        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.DepositAmountFlag"), Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "N") Then
        Range("Sheet1.DepositAmount").MergeArea.ClearContents
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.DepositAmountFlag").value, 1, 1) = "Y") Then
        Range("Sheet1.DepositAmount").MergeArea.Locked = False
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.AggrigateAmountFlag"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag").value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.AggrigateAmountFlag1"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag1").value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.LiableSec44AEflg"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("sheet1.LiableSec44AEflg").value, 1, 1) = "Y") Then
        Range("sheet1.LiableSec5cflg").MergeArea.ClearContents
        Range("sheet1.LiableSec5cflg").MergeArea.Locked = True
        Range("sheet1.LiableSec5cflg").MergeArea.Interior.Color = (&HD8D8D8)
Else
        Range("sheet1.LiableSec5cflg").MergeArea.Locked = False
        Range("sheet1.LiableSec5cflg").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.LiableSec5cflg"), Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("sheet1.LiableSec5cflg").value, 1, 1) = "Y") Then
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Locked = False
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Interior.Color = (&HCCFFCC)
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Locked = False
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Interior.Color = (&HCCFFCC)
Else
        Range("sheet1.LiableSecAggrigateflg").MergeArea.ClearContents
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Locked = True
        Range("sheet1.LiableSecAggrigateflg").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.ClearContents
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Locked = True
        Range("sheet1.LiableSecAggrigate1flg").MergeArea.Interior.Color = (&HD8D8D8)
End If
Sheet1.Protect Password:=getmsgstate
End If

If Not Intersect(Range("sheet1.DepositAmount"), Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.DepositAmount").value < 10000000 Or Sheet1.Range("Sheet1.DepositAmount").value = 10000000) Then
            fmsgbox "*Amount entered cannot be less than or equal to Rs 1 Crore" & Chr(10) & Chr(13)
            Sheet1.Range("Sheet1.DepositAmount").value = " "

        End If
End If

If Not Intersect(Range("sheet1.AggrigateAmount"), Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.AggrigateAmount").value < 200000 Or Sheet1.Range("Sheet1.AggrigateAmount").value = 200000) Then
            fmsgbox "*Amount entered cannot be less than or equal to Rs 2 Lakhs" & Chr(10) & Chr(13)
            Sheet1.Range("Sheet1.AggrigateAmount").value = " "

        End If
End If

If Not Intersect(Range("sheet1.AggrigateAmount1"), Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.AggrigateAmount1").value < 100000 Or Sheet1.Range("Sheet1.AggrigateAmount1").value = 100000) Then
            fmsgbox "*Amount entered cannot be less than or equal to Rs 1 Lakhs" & Chr(10) & Chr(13)
            Sheet1.Range("Sheet1.AggrigateAmount1").value = " "
        End If
End If

If Not Intersect(Range("sheet1.ReturnFileSec"), Range(Target.Address)) Is Nothing Then
        sReceiptNo = ""
        sNoticeDate = ""
        sNoticeNo = ""
        Sheet1.Unprotect Password:=getmsgstate
        LockUnlockFilingStatusDtls
        LockUnlockFilingStatusDtls1
        Sheet52.Unprotect Password:=getmsgstate
            If Mid(Trim(Range("sheet1.ReturnFileSec1")), 1, 2) = "19" Then
                 Sheet52.Range("Veri.92CD").EntireRow.Hidden = False
             Else
                 Sheet52.Range("Veri.92CD").EntireRow.Hidden = True
            End If
        If Mid(Trim(Range("sheet1.ReturnFileSec1")), 1, 2) = "20" Then
            fmsgboxStatus "* 119(2)(b)-after condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)." ', vbInformation, "ITR3"
            Range(Target.Address).Select
            End If
       If UCase(Mid(Trim(Sheet1.Range("sheet1.ReturnFileSec").value), 1, 6)) = "139(4)" Then
            LockUnlock80
       Else
            LockUnlock80
       End If
        Sheet1.Protect Password:=getmsgstate
        Sheet52.Protect Password:=getmsgstate
        
End If
          If Not Intersect(Range("sheet1.OrigRetFiledDate"), Range(Target.Address)) Is Nothing Then
               If Trim(Range("sheet1.OrigRetFiledDate").value) <> "00/00/0000" Then
                   If Not Part_A_General.CheckDateBefore(Trim(Range("sheet1.OrigRetFiledDate").value)) Then
                       Range("sheet1.OrigRetFiledDate").Select
                       Range("sheet1.OrigRetFiledDate").value = ""
                       
                   Else
                   Range("sheet1.OrigRetFiledDate").value = EfilingDate
                   End If
               End If

           End If
          If Not Intersect(Range("sheet1.NoticeDate"), Range(Target.Address)) Is Nothing Then
               If Trim(Range("sheet1.NoticeDate").value) <> "00/00/0000" Then
                   If Not Part_A_General.CheckNoticeDateBefore(Trim(Range("sheet1.NoticeDate").value)) Then
                       Range("sheet1.NoticeDate").Select
                       Range("sheet1.NoticeDate").value = ""
                   Else
                   Range("sheet1.NoticeDate").value = noticeDate
                   End If
               End If

           End If

       If Not Intersect(Range("sheet1.PortugeseCC5A"), Range(Target.Address)) Is Nothing Then
               Application.ScreenUpdating = False
           If UCase(Mid(Target.value, 1, 1)) = "N" Then
                   LOCK5A
           ElseIf UCase(Mid(Target.value, 1, 1)) = "Y" Then
                   fmsgboxStatus "*You have selected to be governed by Sec 5A. Please Fill Schedule 5A."
                   fmsgboxStatus "*Please enter your share of Receipts in all income  heads other than ""Income from Salary"" .The share of Receipts of the spouse should be filled in Schedule 5A"
                   UNLOCK5A
            Else
                UNLOCK5A
           End If
           Application.ScreenUpdating = True
        End If
    If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then
      If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
         Sheet14.Unprotect Password:=sPassword
         Sheet13.Unprotect Password:=sPassword
         Sheet19.Unprotect Password:=sPassword

         Sheet14.Range("os.TRC").Locked = False
         Sheet14.Range("os.TRC").Interior.Color = (&HCCFFCC)
         Sheet14.Range("os.TRC").value = "(Select)"

         Sheet13.Range("STCG.A8_CertiStatus").Locked = False
         Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HCCFFCC)
         Sheet13.Range("STCG.A8_CertiStatus").value = "(Select)"

         Sheet13.Range("LTCG.B10_CertiStatus").Locked = False
         Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HCCFFCC)
         Sheet13.Range("LTCG.B10_CertiStatus").value = "(Select)"
         
         Sheet19.Range("scvia.Section80QQB").Locked = True
         Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HD8D8D8)
         Sheet19.Range("scvia.Section80QQB").ClearContents
         
         Sheet19.Range("scvia.Section80RRB").Locked = True
         Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HD8D8D8)
         Sheet19.Range("scvia.Section80RRB").ClearContents
         
         Sheet19.Protect Password:=sPassword
         Sheet14.Protect Password:=sPassword
         Sheet13.Protect Password:=sPassword
       Else
        Sheet14.Unprotect Password:=sPassword
         Sheet13.Unprotect Password:=sPassword
         Sheet19.Unprotect Password:=sPassword
          Sheet14.Range("os.TRC").Locked = True
          Sheet14.Range("os.TRC").Interior.Color = (&HD8D8D8)
          Sheet14.Range("os.TRC").ClearContents

           Sheet13.Range("STCG.A8_CertiStatus").Locked = True
          Sheet13.Range("STCG.A8_CertiStatus").Interior.Color = (&HD8D8D8)
          Sheet13.Range("STCG.A8_CertiStatus").ClearContents

          Sheet13.Range("LTCG.B10_CertiStatus").Locked = True
          Sheet13.Range("LTCG.B10_CertiStatus").Interior.Color = (&HD8D8D8)
          Sheet13.Range("LTCG.B10_CertiStatus").ClearContents
          
         Sheet19.Range("scvia.Section80QQB").Locked = False
         Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HCCFFCC)
         Sheet19.Range("scvia.Section80QQB").value = 0
         
         Sheet19.Range("scvia.Section80RRB").Locked = False
         Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HCCFFCC)
         Sheet19.Range("scvia.Section80RRB").value = 0
         
         Sheet19.Protect Password:=sPassword
         Sheet14.Protect Password:=sPassword
         Sheet13.Protect Password:=sPassword
        End If
            lockunlock_OS_Stch_Ltcg
            LockUnlockIBANSELECT
            LockUnlockClaim115H
            LockUnlockQVIA
            NewResidentialwork
            LockUnlockTR
            LockUnlockFSI
           Sheet1.Range("sheet1.ResidentialStatus1_Condition").value = "(Select)"
           If Trim(Mid(Target.value, 1, 3)) <> "NRI" Then
           sPassword = EfilingCommon.getmsgstate
           Sheet1.Unprotect Password:=sPassword
           Sheet1.Range("sheet1.NRI_PE").MergeArea.Locked = True
           Sheet1.Range("sheet1.NRI_PE").Interior.Color = (&HD8D8D8)
            Sheet1.Range("sheet1.NRI_PE").value = ""
           Sheet1.Protect Password:=sPassword
               LOCKCG
               UnLockBp
           Else
           sPassword = EfilingCommon.getmsgstate
           Sheet1.Unprotect Password:=sPassword
           Sheet1.Range("sheet1.NRI_PE").MergeArea.Locked = False
           Sheet1.Range("sheet1.NRI_PE").Interior.Color = (&HCCFFCC)
           Sheet1.Range("sheet1.NRI_PE").value = "(Select)"
               UNLOCKcg
               LockBp
           Sheet1.Protect Password:=sPassword
            End If
            If Not isdropdownblank(Target.value) Then
             ResetOSDropDown (Mid(Target.value, 1, 2))
            End If
       End If
           If Not Intersect(Range("sheet1.AsseseeRepFlg"), Range(Target.Address)) Is Nothing Then
                   RepInfo
           End If
If Not Intersect(Range("sheet1.LiableSec44ABflg"), Range(Target.Address)) Is Nothing Then

   If Mid(Range("sheet1.LiableSec44ABflg").value, 1, 1) = "Y" Then
       fmsgboxStatus "*Part A - OI should be filled mandatorily " ', vbOKOnly
   Else:
       fmsgboxStatus "*If not liable under Section 44AB then Audit details need not be fillled" ', vbOKOnly
   End If
End If
If Not Intersect(Range("sheet1.LiableSec92Eflg"), Range(Target.Address)) Is Nothing Then
    LockUnlock92Edia
End If
If Not Intersect(Range("sheet1.LiableSec92Eflgdia"), Range(Target.Address)) Is Nothing Then
    LockUnlockDOA92E
End If
If Not Intersect(Range("sheet1.AsseseeRepFlg"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword
    If Mid(Range("sheet1.AsseseeRepFlg").value, 1, 1) = "N" Then
       Range("sheet1.RepName").MergeArea.Locked = True
       Range("sheet1.RepName").MergeArea.Interior.Color = (&HD8D8D8)
       Range("sheet1.RepName").MergeArea.ClearContents
       
       
'       Range("sheet1.RepCapacity").MergeArea.Locked = True
'       Range("sheet1.RepCapacity").MergeArea.Interior.Color = (&HD8D8D8)
'       Range("sheet1.RepCapacity").MergeArea.ClearContents
'
'
'       Range("sheet1.RepAddress").MergeArea.Locked = True
'       Range("sheet1.RepAddress").MergeArea.Interior.Color = (&HD8D8D8)
'       Range("sheet1.RepAddress").MergeArea.ClearContents
'
'       Range("sheet1.RepPAN").MergeArea.Locked = True
'       Range("sheet1.RepPAN").MergeArea.Interior.Color = (&HD8D8D8)
'       Range("sheet1.RepPAN").MergeArea.ClearContents
'
'       Range("sheet1.RepAadhar").MergeArea.Locked = True
'       Range("sheet1.RepAadhar").MergeArea.Interior.Color = (&HD8D8D8)
'       Range("sheet1.RepAadhar").MergeArea.ClearContents
   Else:
   
       Range("sheet1.RepName").MergeArea.Locked = False
       Range("sheet1.RepName").MergeArea.Interior.Color = (&HCCFFCC)
       
'       Range("sheet1.RepCapacity").value = "(Select)"
'       Range("sheet1.RepCapacity").MergeArea.Locked = False
'       Range("sheet1.RepCapacity").MergeArea.Interior.Color = (&HCCFFCC)
'       Range("sheet1.RepAddress").MergeArea.Locked = False
'       Range("sheet1.RepAddress").MergeArea.Interior.Color = (&HCCFFCC)
'       Range("sheet1.RepPAN").MergeArea.Locked = False
'       Range("sheet1.RepPAN").MergeArea.Interior.Color = (&HCCFFCC)
'
'       Range("sheet1.RepAadhar").MergeArea.Locked = False
'       Range("sheet1.RepAadhar").MergeArea.Interior.Color = (&HCCFFCC)


   End If
 Sheet1.Protect Password:=sPassword
End If
If Not Intersect(Range("sheet1.ResidentialStatus1"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet9.Unprotect Password:=sPassword
  Sheet4.Unprotect Password:=sPassword
  Sheet7.Unprotect Password:=sPassword
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR" Then
   If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
        Sheet9.Range("sheet11.Section44AD").formula = "=0"
        Sheet9.Range("sheet11.Section44ADA").formula = "=0"
        LOCKPL2
   Else
        Sheet9.Range("sheet11.Section44AD").formula = "=PL_61ii"
        Sheet9.Range("sheet11.Section44ADA").formula = "=PL_62ii"
   UNLOCKPL2
   End If
   Sheet7.Range("AOIFlag_1").Locked = True
   Sheet7.Range("AOIFlag_1").Interior.Color = (&HD8D8D8)
   Sheet7.Range("AOIFlag_1").ClearContents

Else
   Sheet9.Range("sheet11.Section44AD").formula = "=PL_61ii"
   Sheet9.Range("sheet11.Section44ADA").formula = "=PL_62ii"
    UNLOCKPL2
   Sheet7.Range("AOIFlag_1").Locked = False
   Sheet7.Range("AOIFlag_1").Interior.Color = (&HCCFFCC)
   Sheet7.Range("AOIFlag_1").ClearContents
End If
 Sheet7.Protect Password:=sPassword
 Sheet9.Protect Password:=sPassword
 Sheet4.Protect Password:=sPassword

End If
If Not Intersect(Range("sheet1.LiableSec44ABflg"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword

  If Mid(Range("sheet1.LiableSec44ABflg").value, 1, 1) <> "Y" Then
       Range("sheet1.AuditAccountantFlg").value = "(Select)"
        Range("sheet1.AuditAccountantFlg").MergeArea.Locked = True
        Range("sheet1.AuditAccountantFlg").MergeArea.ClearContents
        Range("sheet1.AuditAccountantFlg").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.AuditReportFurnishDate").MergeArea.Locked = True
        Range("sheet1.AuditReportFurnishDate").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditReportFurnishDate").MergeArea.ClearContents

        Range("sheet1.AuditorName").MergeArea.Locked = True
        Range("sheet1.AuditorName").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditorName").MergeArea.ClearContents

        Range("sheet1.AuditorMemNo").MergeArea.Locked = True
        Range("sheet1.AuditorMemNo").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditorMemNo").MergeArea.ClearContents

        Range("sheet1.AudFrmName").MergeArea.Locked = True
        Range("sheet1.AudFrmName").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmName").MergeArea.ClearContents

        Range("sheet1.AudFrmRegNo").MergeArea.Locked = True
        Range("sheet1.AudFrmRegNo").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmRegNo").MergeArea.ClearContents

        Range("sheet1.AudFrmPAN").MergeArea.Locked = True
        Range("sheet1.AudFrmPAN").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmPAN").MergeArea.ClearContents

        Range("sheet1.AudFrmAadhaar").MergeArea.Locked = True
        Range("sheet1.AudFrmAadhaar").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmAadhaar").MergeArea.ClearContents

        Range("sheet1.AuditDate").MergeArea.Locked = True
        Range("sheet1.AuditDate").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditDate").MergeArea.ClearContents
    Else:
        Range("sheet1.AuditAccountantFlg").MergeArea.Locked = False
        Range("sheet1.AuditAccountantFlg").MergeArea.Interior.Color = (&HCCFFCC)

   End If
 Sheet1.Protect Password:=sPassword
End If
If Not Intersect(Range("sheet1.AuditAccountantFlg"), Range(Target.Address)) Is Nothing Then
  sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword
  If Mid(Range("sheet1.AuditAccountantFlg").value, 1, 1) <> "Y" Then
       Range("sheet1.AuditReportFurnishDate").MergeArea.Locked = True
        Range("sheet1.AuditReportFurnishDate").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditReportFurnishDate").MergeArea.ClearContents

        Range("sheet1.AuditorName").MergeArea.Locked = True
        Range("sheet1.AuditorName").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditorName").MergeArea.ClearContents

        Range("sheet1.AuditorMemNo").MergeArea.Locked = True
        Range("sheet1.AuditorMemNo").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditorMemNo").MergeArea.ClearContents

        Range("sheet1.AudFrmName").MergeArea.Locked = True
        Range("sheet1.AudFrmName").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmName").MergeArea.ClearContents

        Range("sheet1.AudFrmRegNo").MergeArea.Locked = True
        Range("sheet1.AudFrmRegNo").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmRegNo").MergeArea.ClearContents

        Range("sheet1.AudFrmPAN").MergeArea.Locked = True
        Range("sheet1.AudFrmPAN").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmPAN").MergeArea.ClearContents

        Range("sheet1.AudFrmAadhaar").MergeArea.Locked = True
        Range("sheet1.AudFrmAadhaar").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AudFrmAadhaar").MergeArea.ClearContents

        Range("sheet1.AuditDate").MergeArea.Locked = True
        Range("sheet1.AuditDate").MergeArea.Interior.Color = (&HD8D8D8)
        Range("sheet1.AuditDate").MergeArea.ClearContents
   Else:
       Range("sheet1.AuditReportFurnishDate").MergeArea.Locked = False
      Range("sheet1.AuditReportFurnishDate").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AuditorName").MergeArea.Locked = False
       Range("sheet1.AuditorName").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AuditorMemNo").MergeArea.Locked = False
       Range("sheet1.AuditorMemNo").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AudFrmName").MergeArea.Locked = False
       Range("sheet1.AudFrmName").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AudFrmRegNo").MergeArea.Locked = False
       Range("sheet1.AudFrmRegNo").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AudFrmPAN").MergeArea.Locked = False
       Range("sheet1.AudFrmPAN").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.AudFrmAadhaar").MergeArea.Locked = False
        Range("sheet1.AudFrmAadhaar").MergeArea.Interior.Color = (&HCCFFCC)

       Range("sheet1.AuditDate").MergeArea.Locked = False
       Range("sheet1.AuditDate").MergeArea.Interior.Color = (&HCCFFCC)

   End If
 Sheet1.Protect Password:=sPassword
End If


If Not Intersect(Range("sheet1.AuditorMemNo"), Range(Target.Address)) Is Nothing Then
   If Sheet1.Range("sheet1.AuditorMemNo").value = 0 Then
       fmsgboxStatus "*Membership Number should not contain only Zeroes"
       Sheet1.Range("sheet1.AuditorMemNo").value = ""
   GoTo endl
   End If
   Dim lenofMemNo, zerostoadd As Long
      lenofMemNo = Len(Sheet1.Range("sheet1.AuditorMemNo").value)
      If lenofMemNo <> 6 Then
           zerostoadd = 6 - lenofMemNo
               If zerostoadd = 1 Then
                   Sheet1.Range("sheet1.AuditorMemNo").value = "0" + Sheet1.Range("sheet1.AuditorMemNo").value
               ElseIf zerostoadd = 2 Then
                   Sheet1.Range("sheet1.AuditorMemNo").value = "00" + Sheet1.Range("sheet1.AuditorMemNo").value
               ElseIf zerostoadd = 3 Then
                   Sheet1.Range("sheet1.AuditorMemNo").value = "000" + Sheet1.Range("sheet1.AuditorMemNo").value
               ElseIf zerostoadd = 4 Then
                   Sheet1.Range("sheet1.AuditorMemNo").value = "0000" + Sheet1.Range("sheet1.AuditorMemNo").value
               ElseIf zerostoadd = 5 Then
                   Sheet1.Range("sheet1.AuditorMemNo").value = "00000" + Sheet1.Range("sheet1.AuditorMemNo").value
               End If
               fmsgboxStatus "*Membership Number should not be less than 6 digits, hence it will be prefixed with zero"
       End If

       If Not chkNumeric(Sheet1.Range("sheet1.AuditorMemNo").value) Then
       fmsgboxStatus "*Membership Number should not contain special charachters"
       Sheet1.Range("sheet1.AuditorMemNo").Select
       Sheet1.Range("sheet1.AuditorMemNo").value = ""
       End If
endl:
End If

If Not Intersect(Range("sheet1.AudFrmRegNo"), Range(Target.Address)) Is Nothing Then
   If Sheet1.Range("sheet1.AudFrmRegNo").value = 0 Then
       fmsgboxStatus "*Membership Number should not contain only Zeroes"
       Sheet1.Range("sheet1.AudFrmRegNo").value = ""
    GoTo endl1
   End If
      lenofMemNo = Len(Sheet1.Range("sheet1.AudFrmRegNo").value)
      If lenofMemNo <> 8 Then
           zerostoadd = 8 - lenofMemNo
               If zerostoadd = 1 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "0" + Sheet1.Range("sheet1.AudFrmRegNo").value
               ElseIf zerostoadd = 2 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "00" + Sheet1.Range("sheet1.AudFrmRegNo").value
               ElseIf zerostoadd = 3 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "000" + Sheet1.Range("sheet1.AudFrmRegNo").value
               ElseIf zerostoadd = 4 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "0000" + Sheet1.Range("sheet1.AudFrmRegNo").value
               ElseIf zerostoadd = 5 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "00000" + Sheet1.Range("sheet1.AudFrmRegNo").value
                ElseIf zerostoadd = 6 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "000000" + Sheet1.Range("sheet1.AudFrmRegNo").value
                ElseIf zerostoadd = 7 Then
                   Sheet1.Range("sheet1.AudFrmRegNo").value = "0000000" + Sheet1.Range("sheet1.AudFrmRegNo").value
               End If
               fmsgboxStatus "*Proprietorship/firm registration number should not be less than 8 digits, hence it will be prefixed with zero"
               GoTo endl1
       End If

       If Not CheckSpecialCharacter(Sheet1.Range("sheet1.AudFrmRegNo").value) Then
       fmsgboxStatus "*Proprietorship/firm registration number should not contain special charachters"
       Sheet1.Range("sheet1.AudFrmRegNo").Select
       Sheet1.Range("sheet1.AudFrmRegNo").value = ""
       GoTo endl1
       End If
endl1:
End If
    If Not Intersect(Range("sheet1.AudFrmPAN"), Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
        If Not Part_A_General.CheckAudFrmPAN(UCase(Trim(Range("sheet1.AudFrmPAN").value))) Then
               fmsgboxStatus "*Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."
               Range("sheet1.AudFrmPAN").value = ""
        End If
    End If
    If Not Intersect(Range("sheet1.AudFrmAadhaar"), Range(Target.Address)) Is Nothing Then
       Target.value = (Target.value)
        If Not ValidateAadhaar(UCase(Trim(Range("sheet1.AudFrmAadhaar").value))) Then
               fmsgboxStatus "*Invalid Aadhaar." & Chr(10) & Chr(13) & "Please enter the valid Aadhaar number of the proprietorship in schedule Part A General."
               Range("sheet1.AudFrmAadhaar").value = ""
        End If
    End If
    If Not Intersect(Range("sheet1.CountryCode"), Range(Target.Address)) Is Nothing Then
      If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"
            'Sheet1.Range("sheet1.PinCode") = ""
            If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-STATE OUTSIDE INDIA" Then
                fmsgboxStatus "*Country Should not be India for Selected State" & Chr(13)
                Sheet1.Range("sheet1.CountryCode").value = "(Select)"
                Sheet1.Range("sheet1.CountryCode").Select
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = True
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf Sheet1.Range("sheet1.CountryCode").value = "(Select)" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"

            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               'Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = True
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
            If UCase(Sheet1.Range("sheet1.StateCode1").value) <> "99-STATE OUTSIDE INDIA" And Sheet1.Range("sheet1.StateCode1").value <> "(Select)" Then
                fmsgboxStatus "*Country Should be India for Selected State" & Chr(13)
                Sheet1.Range("sheet1.CountryCode").value = ""
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = False
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
      End If
    End If
    If Not Intersect(Range("HASZIP"), Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)

        If Mid(Trim(Sheet1.Range("HASZIP").value), 1, 1) = "Y" Then
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.ZipCode").MergeArea.value = "XXXXXX"
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
        End If
    End If
    If Not Intersect(Range("Sheet1.115H"), Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
        LockUnlockPartBTTI
        If Mid(Trim(Sheet1.Range("Sheet1.115H").value), 1, 1) = "Y" Then
               UnLock115AD
               UNLOCKcg
               ResetOSDropDown (Mid(Target.value, 1, 2))
        Else
            If Trim(Mid(Sheet1.Range("sheet1.ResidentialStatus1"), 1, 2)) <> "NR" Then
                LOCKCG
                ResetOSDropDown (Mid(Target.value, 1, 2))
            Else
                UNLOCKcg
                ResetOSDropDown (Mid(Target.value, 1, 2))
            End If
        End If

        If Mid(Trim(Sheet1.Range("Sheet1.115H").value), 1, 1) = "N" Then
        Lock115AD
        End If
    End If
    Sheet19.Unprotect Password:=getmsgstate
    Sheet1.Unprotect Password:=getmsgstate
    If (Mid(Range("sheet1.Status").value, 1, 1) = "" Or Mid(Range("sheet1.Status").value, 1, 1) = "(" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 1) = "" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 1) = "(" Or Range("sheet1.DOB").value = "") Then
            If Sheet1.Range("OptingNewTaxRegime").value = "Not opting" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = False
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
            End If
                Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = False
                Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
    End If

    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then
                 Sheet19.Range("scvia.Section80TTA").value = ""
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = True
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)
        Else
            If Sheet1.Range("OptingNewTaxRegime").value = "Not opting" Or Sheet1.Range("OptingNewTaxRegime").value = "Opt out" Then
                 Sheet19.Range("scvia.Section80TTA").MergeArea.Locked = False
                Sheet19.Range("scvia.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
            End If
                 Sheet19.Range("scvia.Section80TTB").value = ""
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
    End If
                Sheet19.Protect Password:=getmsgstate

    Sheet19.Unprotect Password:=getmsgstate
    If calculateAge(Trim(Range("sheet1.DOB").value)) >= 60 And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And Mid(Range("sheet1.Status").value, 1, 1) = "I" Then
'                Sheet19.Range("scvia.Section80TTB").value = ""
                Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = False
                Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)

    End If
    If (Mid(Range("sheet1.Status").value, 1, 1) = "I" And Mid(Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI") Or (Mid(Range("sheet1.Status").value, 1, 1) = "H") Then
                Sheet19.Range("scvia.Section80TTB").value = ""
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Locked = True
                 Sheet19.Range("scvia.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
    End If
                Sheet19.Protect Password:=getmsgstate
    
        Dim PinCode As Range
                 For Each PinCode In SheetALL.Range("All_Pincode_List")
                      Dim code As String
                    code = Left(CStr(Sheet1.Range("sheet1.PinCode").value), 6)
                    If CStr(PinCode.value) = CStr(code) Then
                        Sheet1.Range("sheet1.StateCode1").value = SheetALL.Range("EM" & PinCode.row).value
                        Exit For
                    End If
                Next PinCode
     
     If Not IsEmpty(Trim(Target.value)) Then
           If (Target.name.name) = "sheet1.StateCode1" Then
               If Not Part_A_General.CheckState(UCase(Trim(Range("sheet1.StateCode1").value)), validateType) Then
                   If validateType = "1" Then fmsgboxStatus "*State in Sheet : PART A-GENERAL  is Mandatory" & Chr(13) ', vbOKOnly
               End If
           End If
    End If
    If (Target.name.name) = "sheet1.MiddleName" Or (Target.name.name) = "sheet1.FirstName" Or (Target.name.name) = "sheet1.SurNameOrOrgName" Then
       Sheet52.Range("sheet9.AssesseeVerName").value = UCase(Sheet31.Range("FullName").value)
    End If
endline:
Application.EnableEvents = True
Application.ScreenUpdating = True
Sheet1.Protect Password:=getmsgstate
End Sub


Function CheckDateBeforeinSheet1(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2017, for A.Y. 2017-18

   CheckDateBeforeinSheet1 = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckDateBeforeinSheet1 = False
          fmsgboxStatus (errmsg & " Part-A-General  must be a valid dd/mm/yyyy format")
           Exit Function
       End If

       If Not ChkMinDOBDate1(dob, "31/03/2022") Then
       
           fmsgboxStatus (errmsg & " Part-A-General should be on or after 01/04/2022")
           CheckDateBeforeinSheet1 = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function


Function CheckDateBeforeinSheet2(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2017, for A.Y. 2017-18

   CheckDateBeforeinSheet2 = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckDateBeforeinSheet2 = False
          fmsgboxStatus (errmsg & " Part-A-General  must be a valid dd/mm/yyyy format")
           Exit Function
       End If

       If Not ChkMinDOBDate1(dob, "31/03/2022") Then
           fmsgboxStatus (errmsg & " in Part-A-General should not be less than 01/04/2022.")
           CheckDateBeforeinSheet2 = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function

Private Sub Worksheet_Deactivate()
Application.EnableEvents = False


Dim cnt00001
Dim cnttemp, cell
cnttemp = 0
For Each cell In Sheet2.Range("NOB.Code").Cells
    If Not isdropdownblank(cell.value) Then
        cnttemp = cnttemp + 1
    End If
Next cell
 cnt00001 = Application.CountIf(Sheet2.Range("NOB.Code"), "00001-Share of Income from firm only")
    If cnt00001 > 0 And cnt00001 = cnttemp And Sheet9.Range("sheet11.DebPLUs35ExcessAmt") = 0 And Sheet9.Range("sheet10.ProfBfrTaxPL") = 0 And SheetALL.Range("bacValue") = 2 Then
        ThisWorkbook.Unprotect Password:=getmsgstate
        Sheet56.Unprotect Password:=getmsgstate
        Sheet56.Visible = xlSheetVisible
        Sheet56.Protect Password:=getmsgstate
        ThisWorkbook.Protect Password:=getmsgstate

    Else
        ThisWorkbook.Unprotect Password:=getmsgstate
        Sheet56.Unprotect Password:=getmsgstate
        Sheet56.Visible = xlSheetHidden
        Sheet56.Protect Password:=getmsgstate
        ThisWorkbook.Protect Password:=getmsgstate

    End If

getExemption
Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Dim DateRangecells, dateCtr As Variant
Dim x As Variant

Application.EnableEvents = False



If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet1.Range(targetaddress).Value = ""
ElseIf ActiveCell.Validation.Type = 1 Then
 Application.OnKey "%{DOWN}", ""
ElseIf ActiveCell.Validation.Type = 3 Then

    Application.OnKey "%{DOWN}"
End If
Else
Application.OnKey "%{DOWN}", ""
End If



'Set DateRangecells = Sheet1.Range("sheet1.DateOfAudit").Cells
'dateCtr = Sheet1.Range("sheet1.DateOfAudit").count
'
'For x = 1 To dateCtr
'   If Not DateRangecells.item(x).value = "" Then
'         If Not CheckDateBeforeinSheet2(DateRangecells.item(x).value, "Date Of Audit") Then
'         DateRangecells.item(x).value = ""
'         DateRangecells.item(x).Select
'      Else
'      DateRangecells.item(x).value = CreditDate
'      End If
'   End If
'Next
'
'
'Set DateRangecells = Sheet1.Range("Audit.Date").Cells
'dateCtr = Sheet1.Range("Audit.Date").count
'
'For x = 1 To dateCtr
'   If Not DateRangecells.item(x).value = "" Then
'         If Not CheckDateBeforeinSheet2(DateRangecells.item(x).value, "Date Of Audit") Then
'         DateRangecells.item(x).value = ""
'         DateRangecells.item(x).Select
'      Else
'      DateRangecells.item(x).value = CreditDate
'      End If
'   End If
'Next

'
'If (Target.name.name) = "sheet1.AuditReportFurnishDate" Then
'   If Not CheckDateBeforeinSheet2(Sheet1.Range("sheet1.AuditReportFurnishDate").value, "Date of furnishing Audit Report") Then
'       Sheet1.Range("sheet1.AuditReportFurnishDate").value = ""
'       Sheet1.Range("sheet1.AuditReportFurnishDate").Select
'   Else
'       Sheet1.Range("sheet1.AuditReportFurnishDate").value = CreditDate
'   End If
'
'End If
'
'If (Target.name.name) = "sheet1.AuditDate" Then
'   If Not CheckDateBeforeinSheet2(Sheet1.Range("sheet1.AuditDate").value, "Date Of Audit") Then
'       Sheet1.Range("sheet1.AuditDate").value = ""
'       Sheet1.Range("sheet1.AuditDate").Select
'   End If
'End If
'
'If (Target.name.name) = "sheet1.AuditDateSec92E" Then
'   If Not CheckDateBeforeinSheet2(Sheet1.Range("sheet1.AuditDateSec92E").value, "Date Of Audit") Then
'       Sheet1.Range("sheet1.AuditDateSec92E").value = ""
'       Sheet1.Range("sheet1.AuditDateSec92E").Select
'   End If
'End If

Application.EnableEvents = True
End Sub

Function getExemption()
Dim age As Double
Dim resStatus As String
Dim Status As String
Dim exemption As Double



age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").value)
resStatus = Sheet1.Range("sheet1.ResidentialStatus1").value               'RES , NRI
Status = Sheet1.Range("sheet1.Status").value
       If (Mid(Status, 1, 1) = "I" And age > 59 And age < 80 And (Mid(resStatus, 1, 3) = "RES" Or Mid(resStatus, 1, 3) = "NOR")) Then
           exemption = 300000
       ElseIf (Mid(Status, 1, 1) = "I" And age >= 80 And (Mid(resStatus, 1, 3) = "RES" Or Mid(resStatus, 1, 3) = "NOR")) Then
           exemption = 500000
       ElseIf ((Mid(Status, 1, 1) = "I" Or Mid(Status, 1, 1) = "H") And (Mid(resStatus, 1, 3) = "RES" Or Mid(resStatus, 1, 3) = "NOR")) Then
           exemption = 250000
       Else
           exemption = 0
       End If
 '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet31.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

   Sheet31.Range("exemptionAmt").value = exemption
   'Sheet31.Range("Age").value = age
'----------------Lock Password-------------------START---
 Sheet31.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Function



Function Lock_UnLockTDS2_Col9()
On Error GoTo error1

 '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet28.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
   If UCase(Sheet1.Range("sheet1.PortugeseCC5A").value) = "YES" Then

       Sheet28.Range("TDS2.AmtClaimedBySpouse").Locked = False
       Sheet28.Range("TDS2.AmtClaimedBySpouse").Interior.Color = (&HCCFFCC)
       Sheet28.Range("TDS3.AmtClaimedBySpouse").Locked = False
       Sheet28.Range("TDS3.AmtClaimedBySpouse").Interior.Color = (&HCCFFCC)





   Else:
       Sheet28.Range("TDS2.AmtClaimedBySpouse").Locked = True
       Sheet28.Range("TDS2.AmtClaimedBySpouse").Interior.Color = (&HD8D8D8)
       Sheet28.Range("TDS2.AmtClaimedBySpouse").ClearContents
       Sheet28.Range("TDS3.AmtClaimedBySpouse").Locked = True
       Sheet28.Range("TDS3.AmtClaimedBySpouse").Interior.Color = (&HD8D8D8)
       Sheet28.Range("TDS3.AmtClaimedBySpouse").ClearContents
   End If

error1:
'----------------Lock Password-------------------START---
 Sheet28.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Function



Function LockUnlockVIA()
On Error GoTo error1
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword

   If Mid(UCase(Sheet1.Range("sheet1.LiableSec44ABflg").value), 1, 1) = "Y" Then

       Sheet19.Range("scvia.Section80JJAA").Locked = False
       Sheet19.Range("scvia.Section80JJAA").Interior.Color = (&HCCFFCC)

   Else
       Sheet19.Range("scvia.Section80JJAA").Locked = True
       Sheet19.Range("scvia.Section80JJAA").Interior.Color = (&HD8D8D8)
       Sheet19.Range("scvia.Section80JJAA").ClearContents
   End If

error1:
 Sheet19.Protect Password:=sPassword
End Function



Function LockUnlockDOA92E()
On Error GoTo error1
   sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword

   If Mid(UCase(Sheet1.Range("sheet1.LiableSec92Eflgdia").value), 1, 1) = "Y" Then

       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Locked = False
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Interior.Color = (&HCCFFCC)
 
   Else
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Locked = True
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.ClearContents

   End If

error1:
 Sheet1.Protect Password:=sPassword
End Function

'newly added by Sai For AY 2024-25
Function LockUnlockDOA92E_AckNum()
On Error GoTo error1
   sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword

   If Mid(UCase(Sheet1.Range("sheet1.LiableSec92Eflgdia").value), 1, 1) = "Y" Then

       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Locked = False
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Interior.Color = (&HCCFFCC)
       
   Else
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Locked = True
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.ClearContents
       
   End If

error1:
 Sheet1.Protect Password:=sPassword
End Function
Function LockUnlock92Edia()
On Error GoTo error1
   sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword

   If Mid(UCase(Sheet1.Range("sheet1.LiableSec92Eflg").value), 1, 1) = "Y" Then

       Sheet1.Range("sheet1.LiableSec92Eflgdia").MergeArea.Locked = False
       Sheet1.Range("sheet1.LiableSec92Eflgdia").MergeArea.Interior.Color = (&HCCFFCC)
     
     If Mid(UCase(Sheet1.Range("sheet1.LiableSec92Eflgdia").value), 1, 1) = "Y" Then
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Locked = False
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Interior.Color = (&HCCFFCC)
       
       
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Locked = False
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Interior.Color = (&HCCFFCC)
       
       
       
     End If
     
   Else
       Sheet1.Range("sheet1.LiableSec92Eflgdia").MergeArea.Locked = True
       Sheet1.Range("sheet1.LiableSec92Eflgdia").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.LiableSec92Eflgdia").MergeArea.value = "(Select)"
       
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Locked = True
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.AuditDateSec92E").MergeArea.value = ""
       
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Locked = True
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("Sheet1.AckNumAuditReport92E").MergeArea.value = ""
       
       
       
   End If

error1:
 Sheet1.Protect Password:=sPassword
End Function

Function ChkMinDOBDate1(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinDOBDate1 = True
    If Len(dob) > 0 Then
       If val(Mid(dob, 7, 4)) < 2022 Then
           ChkMinDOBDate1 = False
           Exit Function
       ElseIf (val(Mid(dob, 7, 4)) = 2022 And val(Mid(dob, 4, 2)) <= 3) Then
           ChkMinDOBDate1 = False
           Exit Function
       End If
   End If
End Function

Sub LockUnlockFilingStatusDtls()
On Error Resume Next

    Dim ReturnSec As String
    ReturnSec = Sheet1.Range("sheet1.ReturnFileSec1").value

    Sheet1.Unprotect Password:=getmsgstate
    Sheet17.Unprotect Password:=getmsgstate
    Application.EnableEvents = False

'    If Mid(ReturnSec, 1, 2) = "12" Or Mid(ReturnSec, 1, 2) = "13" Or Mid(ReturnSec, 1, 2) = "20" Or Mid(ReturnSec, 1, 2) = "21" Then
'IPIP-101888 sadinnei on 10/07/2026
If Mid(ReturnSec, 1, 2) = "12" Or Mid(ReturnSec, 1, 2) = "13" Or Mid(ReturnSec, 1, 2) = "14" Or Mid(ReturnSec, 1, 2) = "16" Or Mid(ReturnSec, 1, 2) = "17" Or Mid(ReturnSec, 1, 2) = "18" Or Mid(ReturnSec, 1, 2) = "19" Or Mid(ReturnSec, 1, 2) = "20" Or Mid(ReturnSec, 1, 2) = "21" Then
    Sheet17.Range("yr2007.HPLossCF10").Locked = False
    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").Locked = False
    Sheet17.Range("yr2007.LossFrmSpecBusCF10").Locked = False
    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").Locked = False
    Sheet17.Range("yr2007.STCGLossCF10").Locked = False
    Sheet17.Range("yr2007.LTCGLossCF10").Locked = False
    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").Locked = False

    Sheet17.Range("yr2007.HPLossCF10").formula = "=sheet16.BalHPlossCurYrAftSetoff"
    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").formula = "=sheet16.BalBusLossAftSetoff"
    Sheet17.Range("yr2007.LossFrmSpecBusCF10").formula = "=ABS(MIN(0,sheet12.AdjustedPLFrmSpecuBus))"
    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").formula = "=ABS(MIN(0,sheet12.AdjustedPLFrmSpecifiedBus))"
    'Sheet17.Range("yr2007.STCGLossCF10").formula = "=IHLA.Eviii2_StclSetoff15Per+IHLA.Eviii3_StclSetoff30Per+IHLA.Eviii4_StclSetoffAppRate+IHLA.Eviii5_StclSetoffDTAARate"
  'For AY 2025-26 by sai on 16/04/2025
     Sheet17.Range("yr2007.STCGLossCF10").formula = "=IHLA.Eviii2_StclSetoff15Per+IHLA.Exii2_StclSetoff20Per+IHLA.Eviii3_StclSetoff30Per+IHLA.Eviii4_StclSetoffAppRate+IHLA.Eviii5_StclSetoffDTAARate"
    
    'Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"
  'For AY 2025-26 by sai on 16/04/2025
     Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Exii8_LtclSetOff12.5Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"

    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").formula = "=ABS(MIN(0,os.BalanceOwnRaceHorse))"

    Sheet17.Range("yr2007.HPLossCF10").Interior.Color = RGB(255, 255, 204)
    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").Interior.Color = RGB(255, 255, 204)
    Sheet17.Range("yr2007.LossFrmSpecBusCF10").Interior.Color = RGB(255, 255, 204)
    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").Interior.Color = RGB(255, 255, 204)
    Sheet17.Range("yr2007.STCGLossCF10").Interior.Color = RGB(255, 255, 204)
    Sheet17.Range("yr2007.LTCGLossCF10").Interior.Color = RGB(255, 255, 204)
    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").Interior.Color = RGB(255, 255, 204)

    Else

    Sheet17.Range("yr2007.HPLossCF10").Locked = True
    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").Locked = True
    Sheet17.Range("yr2007.LossFrmSpecBusCF10").Locked = True
    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").Locked = True
    Sheet17.Range("yr2007.STCGLossCF10").Locked = True
    Sheet17.Range("yr2007.LTCGLossCF10").Locked = True
    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").Locked = True

    Sheet17.Range("yr2007.HPLossCF10").formula = "=sheet16.BalHPlossCurYrAftSetoff"
    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").formula = "=sheet16.BalBusLossAftSetoff"
    Sheet17.Range("yr2007.LossFrmSpecBusCF10").formula = "=ABS(MIN(0,sheet12.AdjustedPLFrmSpecuBus))"
    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").formula = "=ABS(MIN(0,sheet12.AdjustedPLFrmSpecifiedBus))"
    
    'Sheet17.Range("yr2007.STCGLossCF10").formula = "=IHLA.Eviii2_StclSetoff15Per+IHLA.Eviii3_StclSetoff30Per+IHLA.Eviii4_StclSetoffAppRate+IHLA.Eviii5_StclSetoffDTAARate"
    'Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"
    
'For AY 2025-26 by sai on 16/04/2025
    Sheet17.Range("yr2007.STCGLossCF10").formula = "=IHLA.Eviii2_StclSetoff15Per+IHLA.Exii2_StclSetoff20Per+IHLA.Eviii3_StclSetoff30Per+IHLA.Eviii4_StclSetoffAppRate+IHLA.Eviii5_StclSetoffDTAARate"
    Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Exii8_LtclSetOff12.5Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"
    
    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").formula = "=ABS(MIN(0,os.BalanceOwnRaceHorse))"



    Sheet17.Range("yr2007.HPLossCF10").Interior.ColorIndex = 2
    Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").Interior.ColorIndex = 2
    Sheet17.Range("yr2007.LossFrmSpecBusCF10").Interior.ColorIndex = 2
    Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").Interior.ColorIndex = 2
    Sheet17.Range("yr2007.STCGLossCF10").Interior.ColorIndex = 2
    Sheet17.Range("yr2007.LTCGLossCF10").Interior.ColorIndex = 2
    Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").Interior.ColorIndex = 2

    End If
    Application.EnableEvents = True
    

    Sheet17.Protect Password:=getmsgstate
     Sheet1.Protect Password:=getmsgstate
    End Sub

Sub LockUnlockFilingStatusDtls1()
On Error Resume Next

               Dim ReturnSec As String
                ReturnSec = Sheet1.Range("sheet1.ReturnFileSec1").value

             Sheet1.Unprotect Password:=getmsgstate
             
        Application.EnableEvents = False

        If ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 13) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 14) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 15) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 16)) Then

                Sheet1.Unprotect Password:=getmsgstate
                Range("sheet1.ReceiptNo").MergeArea.Locked = True
                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.ReceiptNo").value = ""
                Sheet1.Unprotect Password:=getmsgstate
                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.OrigRetFiledDate").value = ""
                
                Range("sheet1.NoticeDate").MergeArea.Locked = False
                Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)

                Range("Sheet1.UniqueNo").MergeArea.Locked = False
                Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HCCFFCC)


            ElseIf (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18) Then
                Range("sheet1.ReceiptNo").MergeArea.Locked = False
                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)

                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

                Range("sheet1.NoticeDate").MergeArea.Locked = False
                Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)

                Range("Sheet1.UniqueNo").MergeArea.Locked = False
                Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HCCFFCC)


             ElseIf (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 20) Then
                Sheet1.Unprotect Password:=getmsgstate
                
                Range("sheet1.ReceiptNo").MergeArea.Locked = True
                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.ReceiptNo").value = ""
                
                Sheet1.Unprotect Password:=getmsgstate
                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.OrigRetFiledDate").value = ""
                
                
                Range("sheet1.NoticeDate").MergeArea.Locked = False
                Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)
                Range("sheet1.NoticeDate").value = ""
                
                Range("Sheet1.UniqueNo").MergeArea.Locked = False
                Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HCCFFCC)


            ElseIf ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12)) Then
            
                Sheet1.Unprotect Password:=getmsgstate
                Range("sheet1.ReceiptNo").MergeArea.Locked = True
                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.ReceiptNo").value = ""
           
                Sheet1.Unprotect Password:=getmsgstate
                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
               
                Sheet1.Unprotect Password:=getmsgstate
                Range("sheet1.NoticeDate").MergeArea.Locked = True
                Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)
                Range("sheet1.NoticeDate").value = ""
                
                Sheet1.Unprotect Password:=getmsgstate
                Range("Sheet1.UniqueNo").MergeArea.Locked = True
                Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HD8D8D8)
                Range("Sheet1.UniqueNo").value = ""

ElseIf ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17)) Or ((Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 21)) Then

                Range("sheet1.ReceiptNo").MergeArea.Locked = False
                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)

                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

                Range("sheet1.NoticeDate").MergeArea.ClearContents
                Range("sheet1.NoticeDate").MergeArea.Locked = True
                Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)

                Range("Sheet1.UniqueNo").MergeArea.ClearContents
                Range("Sheet1.UniqueNo").MergeArea.Locked = True
                Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HD8D8D8)

        ElseIf Mid(ReturnSec, 1, 2) = "19" Then

         Sheet1.Range("sheet1.ReceiptNo").MergeArea.Locked = False
         Sheet1.Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)

         Sheet1.Range("sheet1.NoticeDate").MergeArea.Locked = False
         Sheet1.Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)

         'Range("Sheet1.UniqueNo").MergeArea.ClearContents
        
        Range("Sheet1.UniqueNo").MergeArea.Locked = True
        Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.UniqueNo").MergeArea.ClearContents
        
         Sheet1.Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
         Sheet1.Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)
    Else
        Sheet1.Unprotect Password:=getmsgstate
         Sheet1.Range("sheet1.ReceiptNo").MergeArea.Locked = True
         Sheet1.Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)
         Sheet1.Range("sheet1.ReceiptNo").MergeArea.ClearContents
         Sheet1.Unprotect Password:=getmsgstate
         Sheet1.Unprotect Password:=getmsgstate
         Sheet1.Range("sheet1.NoticeDate").MergeArea.Locked = True
         Sheet1.Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)
         Sheet1.Range("sheet1.NoticeDate").MergeArea.ClearContents
         
         Sheet1.Unprotect Password:=getmsgstate
         Sheet1.Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
         Sheet1.Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)
         Sheet1.Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
         
         Sheet1.Unprotect Password:=getmsgstate
         Sheet1.Range("Sheet1.UniqueNo").MergeArea.Locked = True
         Sheet1.Range("Sheet1.UniqueNo").MergeArea.Interior.Color = (&HD8D8D8)
         Sheet1.Range("sheet1.UniqueNo").MergeArea.ClearContents

    End If
 Sheet19.Unprotect Password:=getmsgstate
   If Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12 Then
   
         Sheet19.Range("scvia.Section80IAB").ClearContents
         Sheet19.Range("scvia.Section80IAB").Locked = True
         Sheet19.Range("scvia.Section80IAB").Interior.Color = (&HD8D8D8)

         Sheet19.Range("scvia.Section80IBA").ClearContents
         Sheet19.Range("scvia.Section80IBA").Locked = True
         Sheet19.Range("scvia.Section80IBA").Interior.Color = (&HD8D8D8)

         Sheet19.Range("scvia.Section80JJA").ClearContents
         Sheet19.Range("scvia.Section80JJA").Locked = True
         Sheet19.Range("scvia.Section80JJA").Interior.Color = (&HD8D8D8)

         'Sheet19.Range("scvia.Section80JJAA").ClearContents
         
'SIT-91088 sadineni 15_09_2025
'         Sheet19.Range("scvia.Section80JJAA").Locked = True
'         Sheet19.Range("scvia.Section80JJAA").Interior.Color = (&HD8D8D8)
'--
         Sheet19.Range("scvia.Section80QQB").ClearContents
         Sheet19.Range("scvia.Section80QQB").Locked = True
         Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HD8D8D8)

          Sheet19.Range("scvia.Section80RRB").ClearContents
         Sheet19.Range("scvia.Section80RRB").Locked = True
         Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HD8D8D8)
   Else
        'sheet19.Range("scvia.Section80IAB").ClearContents
         Sheet19.Range("scvia.Section80IAB").Locked = False
         Sheet19.Range("scvia.Section80IAB").Interior.Color = (&HCCFFCC)

         'sheet19.Range("scvia.Section80IBA").ClearContents
         Sheet19.Range("scvia.Section80IBA").Locked = False
         Sheet19.Range("scvia.Section80IBA").Interior.Color = (&HCCFFCC)

         'sheet19.Range("scvia.Section80JJA").ClearContents
         Sheet19.Range("scvia.Section80JJA").Locked = False
         Sheet19.Range("scvia.Section80JJA").Interior.Color = (&HCCFFCC)
         
'SIT-91088 sadineni 15_09_2025

'         Sheet19.Range("scvia.Section80JJAA").ClearContents
'         Sheet19.Range("scvia.Section80JJAA").Locked = False
'         Sheet19.Range("scvia.Section80JJAA").Interior.Color = (&HCCFFCC)
'--
        ' sheet19.Range("scvia.Section80QQB").ClearContents
          Sheet19.Range("scvia.Section80QQB").Locked = False
          Sheet19.Range("scvia.Section80QQB").Interior.Color = (&HCCFFCC)

         'sheet19.Range("scvia.Section80RRB").ClearContents
          Sheet19.Range("scvia.Section80RRB").Locked = False
          Sheet19.Range("scvia.Section80RRB").Interior.Color = (&HCCFFCC)
   End If
   Application.EnableEvents = True
   
 Sheet19.Protect Password:=getmsgstate
    Sheet1.Protect Password:=getmsgstate
End Sub
Sub NewResidentialwork()
Application.EnableEvents = False
sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword
'            NRI_Hidden
'            Individual_Hidden
            If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
            Sheet1.Range("Individual_Hidden").EntireRow.Hidden = False

            Else
            'Sheet1.Range("Individual_Hidden").value = "(Select)"
            Sheet1.Range("Individual_Hidden").EntireRow.Hidden = True
            End If


            If (Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" And Mid(Sheet1.Range("Sheet1.ResidentialStatus1").value, 1, 3) = "NRI") Then
            Sheet1.Range("NRI_Hidden").EntireRow.Hidden = False
            Dim chCells As Variant
            For Each chCells In Range("NRI_Juridiction")
            If chCells.MergeCells Then
            chCells.MergeArea.value = "(Select)"
            Else
            chCells.value = "(Select)"
            End If
            Next
            'Sheet1.Range("NRI_Juridiction").ClearContents
            Sheet1.Range("NRI_Taxpayer").ClearContents
            Sheet1.Shapes("Button 403").Visible = msoTrue
            Sheet1.Shapes("Button 403").Left = Range("Resident_Button").Left + 3
            Sheet1.Shapes("Button 403").Top = Range("Resident_Button").Top + 3
            
            Else
            
            Sheet1.Range("NRI_Hidden").EntireRow.Hidden = True
            Sheet1.Shapes("Button 403").Visible = msoFalse
            Sheet1.Shapes("Button 403").Visible = msoFalse
            End If
Application.EnableEvents = True
          Sheet1.Protect Password:=sPassword
End Sub
Sub HideUnhideCheck()
On Error Resume Next
Sheet1.Unprotect Password:=getmsgstate
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" Then
        'Sheet1.Range("IsDirectorCompany").EntireRow.Hidden = False
        Sheet1.Range("IsDirectorCompany").MergeArea.Locked = False
        Sheet1.Range("IsDirectorCompany").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IsDirectorCompany").value = "(Select)"
        LockUnlockCheck
    Else

        Sheet1.Range("IsDirectorCompany").MergeArea.Locked = True
        Sheet1.Range("IsDirectorCompany").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IsDirectorCompany").value = ""
        LockUnlockCheck
    End If
Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockUnlockCheck()
On Error Resume Next
    If UCase(Mid(Trim(Sheet1.Range("IsDirectorCompany").value), 1, 1)) = "Y" Then

        UnLockDirector

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Protect Password:=getmsgstate
    Else
        LockDirector

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Protect Password:=getmsgstate
    End If
End Sub

Sub LockDirector()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet1.Unprotect Password:=sPassword

        rngname = "Sheet1.CompanyName||Sheet1.CompanyType||Sheet1.CompPAN||Sheet1.Shares||Sheet1.DIN"
        LOCKRANGEPARTA (rngname)

    Sheet1.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnLockDirector()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant
    Dim chCells As Variant
    sPassword = getmsgstate
    Sheet1.Unprotect Password:=sPassword

        rngname = "Sheet1.CompanyName||Sheet1.CompanyType||Sheet1.CompPAN||Sheet1.Shares||Sheet1.DIN"
        UNLOCKRANGEPARTA (rngname)
        For Each chCells In Range("Sheet1.CompanyType")
            If chCells.MergeCells Then
            chCells.MergeArea.value = "(Select)"
            Else
            chCells.value = "(Select)"
            End If
            Next


    Sheet1.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub LockShares()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet1.Unprotect Password:=sPassword

        rngname = "Share_NameofCompany||Share_TypeofCompany||Share_PAN||Share_OP_Count||Share_OP_Cost||Share_Acquired_Count||Share_Acquired_Date||Share_Acquired_FaceValue||Share_Acquired_Issue||Share_Acquired_Cost||Share_Transferred_Count||Share_Transferred_Sale||Share_CB_Count||Share_CB_Cost"
        LOCKRANGEPARTA (rngname)

    Sheet1.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnLockShares()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet1.Unprotect Password:=sPassword

        rngname = "Share_NameofCompany||Share_TypeofCompany||Share_PAN||Share_OP_Count||Share_OP_Cost||Share_Acquired_Count||Share_Acquired_Date||Share_Acquired_FaceValue||Share_Acquired_Issue||Share_Acquired_Cost||Share_Transferred_Count||Share_Transferred_Sale||Share_CB_Count||Share_CB_Cost"
        UNLOCKRANGEPARTA (rngname)
        Dim chCells As Variant
            For Each chCells In Range("Share_TypeofCompany")
            If chCells.MergeCells Then
            chCells.MergeArea.value = "(Select)"
            Else
            chCells.value = "(Select)"
            End If
            Next

    Sheet1.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub LOCKRANGEPARTA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet1.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub

Sub UNLOCKRANGEPARTA(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet1.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub
Sub LockUnlockFirm()
On Error Resume Next
    If UCase(Mid(Trim(Sheet1.Range("IsPartnerFirm").value), 1, 1)) = "Y" Then
        UnLockFirm

    Else
        LockFirm

    End If
End Sub

Sub LockFirm()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet1.Unprotect Password:=sPassword

        rngname = "Sheet1.NameOfFirm||Sheet1.FirmPAN"
        LOCKRANGEPARTA (rngname)

    Sheet1.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnLockFirm()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet1.Unprotect Password:=sPassword

        rngname = "Sheet1.NameOfFirm||Sheet1.FirmPAN"
        UNLOCKRANGEPARTA (rngname)

    Sheet1.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKPL1()
Application.EnableEvents = False
'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_PL1 As Variant


rngname_PL1 = "PL_61a||PL_61b||PL_61iia||PL_61iib||PL_62i||PL_62ii"
UNLOCKRANGEPL1 (rngname_PL1)

'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub


Sub LOCKPL1()
Application.EnableEvents = False

'----------------Unlock Password-------------------START---
  sPassword = EfilingCommon.getmsgstate
  Sheet4.Unprotect Password:=sPassword
'   Sheet24.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_PL1 As Variant
rngname_PL1 = "PL_61a||PL_61b||PL_61iia||PL_61iib||PL_62i||PL_62ii"
LOCKRANGEPL1 (rngname_PL1)




'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'   Sheet24.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
End Sub

Sub LOCKRANGEPL1(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))
           If chcell.MergeCells Then
           chcell.MergeArea.value = ""
           chcell.MergeArea.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           Else
           chcell.value = ""
           chcell.Locked = True
           chcell.Interior.Color = (&HD8D8D8)
           End If
           Next
       Next
End Sub

Sub UNLOCKRANGEPL1(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
   rangenamestring = Split(rangenamestring, "||")
       For i = 0 To UBound(rangenamestring)
           For Each chcell In Sheet4.Range(rangenamestring(i))
           If chcell.MergeCells Then
           chcell.MergeArea.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           Else
           chcell.Locked = False
           chcell.Interior.Color = (&HCCFFCC)
           End If

           Next
       Next
End Sub

'Sub LOCKRANGEPL2(rangenamestring As Variant)
'Dim i As Long
'Dim chcell As Range
'   rangenamestring = Split(rangenamestring, "||")
'       For i = 0 To UBound(rangenamestring)
'           For Each chcell In Sheet4.Range(rangenamestring(i))
'           If chcell.MergeCells Then
'           chcell.MergeArea.value = ""
'           chcell.MergeArea.Locked = True
'           chcell.Interior.Color = (&HD8D8D8)
'           Else
'           chcell.value = ""
'           chcell.Locked = True
'           chcell.Interior.Color = (&HD8D8D8)
'           End If
'           Next
'       Next
'End Sub
'
'Sub UNLOCKRANGEPL2(rangenamestring As Variant)
'Dim i As Long
'Dim chcell As Range
'   rangenamestring = Split(rangenamestring, "||")
'       For i = 0 To UBound(rangenamestring)
'           For Each chcell In Sheet4.Range(rangenamestring(i))
'
'           If chcell.MergeCells Then
'           chcell.MergeArea.Locked = False
'           chcell.Interior.Color = (&HCCFFCC)
'           Else
'           chcell.Locked = False
'           chcell.Interior.Color = (&HCCFFCC)
'           End If
'
'           Next
'       Next
'End Sub

'Sub UNLOCKPL2()
'Application.EnableEvents = False
''----------------Unlock Password-------------------START---
'  sPassword = EfilingCommon.getmsgstate
'  Sheet4.Unprotect Password:=sPassword
''   Sheet24.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'
'Dim rngname_PL2 As Variant
'
'
'rngname_PL2 = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription||PL_61a||PL_61b||PL_61iia||PL_61iib||Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1||PL_62i||PL_62ii"
'UNLOCKRANGEPL2 (rngname_PL2)
'
''----------------Lock Password-------------------START---
'  Sheet4.Protect Password:=sPassword
''   Sheet24.Protect Password:=sPassword
''----------------Lock Password-------------------END-----
'
'Application.EnableEvents = True
'End Sub
'
'
'Sub LOCKPL2()
'Application.EnableEvents = False
'
''----------------Unlock Password-------------------START---
'  sPassword = EfilingCommon.getmsgstate
'  Sheet4.Unprotect Password:=sPassword
''   Sheet24.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'Dim rngname_PL2 As Variant
'rngname_PL2 = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription||PL_61a||PL_61b||PL_61iia||PL_61iib||Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1||PL_62i||PL_62ii"
'LOCKRANGEPL2 (rngname_PL2)
'
'
'
'
''----------------Lock Password-------------------START---
'  Sheet4.Protect Password:=sPassword
''   Sheet24.Protect Password:=sPassword
''----------------Lock Password-------------------END-----
'
'Application.EnableEvents = True
'End Sub
Sub LockUnlock115AD()
On Error Resume Next

    If UCase(Mid(Trim(Sheet1.Range("sheet1.ResidentialStatus1").value), 1, 3)) = "NRI" Or (Sheet1.Range("sheet1.ResidentialStatus1").value) = "(Select)" Then
        UnLock115AD

    Else
        Lock115AD

    End If
End Sub

Sub Lock115AD()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet54.Unprotect Password:=sPassword

    'changed by Chetan C M on 30/01/2026 for AY 2026-27
    'start--
        rngname = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
        'rngname = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
    '--end
    
        LOCKRANGE115AD (rngname)

    Sheet54.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub LOCKRANGE115AD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet54.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
LockFlag115A = True
End Sub
Sub UNLOCKRANGE115AD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet54.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    If rangenamestring(i) = "ShareAcq_115AD_1" Then
                        chcell.value = "(Select)"
                    End If
                    
                    'commented by Chetan C M on 30/01/2026 for AY 2026-27
                    'start--
'                    'Added by Aavula Naresh
'                    If rangenamestring(i) = "ShareTrans_115AD_1" Then
'                        chcell.value = "(Select)"
'                    End If
'                    '''''
                    '--end
                    
                End If
            Next
        Next
LockFlag115A = False
End Sub

Sub UnLock115AD()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet54.Unprotect Password:=sPassword

    'changed by Chetan C M on 30/01/2026 for AY 2026-27
    'start--
        rngname = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
        'rngname = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||ExpenditureWholly_115AD"
    '--end
        
        UNLOCKRANGE115AD (rngname)

    Sheet54.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub LockUnlockTDS()
On Error Resume Next

    If UCase(Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1)) = "H" Then
    'Or (Sheet1.Range("sheet1.ResidentialStatus1").value) = "(Select)" Then
        LockTDS
    Else
        UnLockTDS

    End If
End Sub

Sub LockTDS()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet28.Unprotect Password:=sPassword

        rngname = "TDS1.TAN||TDS1.EmployerOrDeductorOrCollecterName||TDS1.IncChrgSal||TDS1.TotalTDSSal"
        LOCKRANGETDS (rngname)

    Sheet28.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnLockTDS()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet28.Unprotect Password:=sPassword

        rngname = "TDS1.TAN||TDS1.EmployerOrDeductorOrCollecterName||TDS1.IncChrgSal||TDS1.TotalTDSSal"
        UNLOCKRANGETDS (rngname)

    Sheet28.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub LOCKRANGETDS(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet28.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub
Sub UNLOCKRANGETDS(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet28.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub
Sub LockUnlockESOP()
On Error Resume Next

    If Sheet1.Range("sheet1.Status").value = "H-HUF" Then
      Sheet46.Unprotect Password:=getmsgstate
                
        Sheet46.Range("ESOP.TaxPrevAY").ClearContents
        Sheet46.Range("ESOP.TaxPrevAY").Locked = True
        Sheet46.Range("ESOP.TaxPrevAY").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.SecType").ClearContents
        Sheet46.Range("ESOP.SecType").Locked = True
        Sheet46.Range("ESOP.SecType").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.CeasedEmployee").ClearContents
        Sheet46.Range("ESOP.CeasedEmployee").Locked = True
        Sheet46.Range("ESOP.CeasedEmployee").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.DateOfCeasing").ClearContents
        Sheet46.Range("ESOP.DateOfCeasing").Locked = True
        Sheet46.Range("ESOP.DateOfCeasing").Interior.Color = "&HD8D8D8"
        
'Change-13.05.2023.101 -> SIT-51187
        Sheet46.Range("ESOP.TaxPrevAY2").ClearContents
        Sheet46.Range("ESOP.TaxPrevAY2").Locked = True
        Sheet46.Range("ESOP.TaxPrevAY2").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.SecType2").ClearContents
        Sheet46.Range("ESOP.SecType2").Locked = True
        Sheet46.Range("ESOP.SecType2").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.CeasedEmployee2").ClearContents
        Sheet46.Range("ESOP.CeasedEmployee2").Locked = True
        Sheet46.Range("ESOP.CeasedEmployee2").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.DateOfCeasing2").ClearContents
        Sheet46.Range("ESOP.DateOfCeasing2").Locked = True
        Sheet46.Range("ESOP.DateOfCeasing2").Interior.Color = "&HD8D8D8"
'----

        Sheet46.Range("ESOP.TaxPrevAY3").ClearContents
        Sheet46.Range("ESOP.TaxPrevAY3").Locked = True
        Sheet46.Range("ESOP.TaxPrevAY3").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.SecType3").ClearContents
        Sheet46.Range("ESOP.SecType3").Locked = True
        Sheet46.Range("ESOP.SecType3").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.CeasedEmployee3").ClearContents
        Sheet46.Range("ESOP.CeasedEmployee3").Locked = True
        Sheet46.Range("ESOP.CeasedEmployee3").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.DateOfCeasing3").ClearContents
        Sheet46.Range("ESOP.DateOfCeasing3").Locked = True
        Sheet46.Range("ESOP.DateOfCeasing3").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.TaxPrevAY4").ClearContents
        Sheet46.Range("ESOP.TaxPrevAY4").Locked = True
        Sheet46.Range("ESOP.TaxPrevAY4").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.SecType4").ClearContents
        Sheet46.Range("ESOP.SecType4").Locked = True
        Sheet46.Range("ESOP.SecType4").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.CeasedEmployee4").ClearContents
        Sheet46.Range("ESOP.CeasedEmployee4").Locked = True
        Sheet46.Range("ESOP.CeasedEmployee4").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.DateOfCeasing4").ClearContents
        Sheet46.Range("ESOP.DateOfCeasing4").Locked = True
        Sheet46.Range("ESOP.DateOfCeasing4").Interior.Color = "&HD8D8D8"
        
                Sheet46.Range("ESOP.TaxPrevAY5").ClearContents
        Sheet46.Range("ESOP.TaxPrevAY5").Locked = True
        Sheet46.Range("ESOP.TaxPrevAY5").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.SecType5").ClearContents
        Sheet46.Range("ESOP.SecType5").Locked = True
        Sheet46.Range("ESOP.SecType5").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.CeasedEmployee5").ClearContents
        Sheet46.Range("ESOP.CeasedEmployee5").Locked = True
        Sheet46.Range("ESOP.CeasedEmployee5").Interior.Color = "&HD8D8D8"
        
        Sheet46.Range("ESOP.DateOfCeasing5").ClearContents
        Sheet46.Range("ESOP.DateOfCeasing5").Locked = True
        Sheet46.Range("ESOP.DateOfCeasing5").Interior.Color = "&HD8D8D8"
        
        LockESOP
        Sheet46.Protect Password:=getmsgstate
    Else
        Sheet46.Unprotect Password:=getmsgstate
                       
        Sheet46.Range("ESOP.TaxPrevAY").Locked = False
        Sheet46.Range("ESOP.TaxPrevAY").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.SecType").Locked = False
        Sheet46.Range("ESOP.SecType").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.CeasedEmployee").Locked = False
        Sheet46.Range("ESOP.CeasedEmployee").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.DateOfCeasing").Locked = False
        Sheet46.Range("ESOP.DateOfCeasing").Interior.Color = "&HCCFFCC"
        
'Change-13.05.2023.101 -> SIT-51187
        Sheet46.Range("ESOP.TaxPrevAY2").Locked = False
        Sheet46.Range("ESOP.TaxPrevAY2").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.SecType2").Locked = False
        Sheet46.Range("ESOP.SecType2").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.CeasedEmployee2").Locked = False
        Sheet46.Range("ESOP.CeasedEmployee2").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.DateOfCeasing2").Locked = False
        Sheet46.Range("ESOP.DateOfCeasing2").Interior.Color = "&HCCFFCC"
        
        
        Sheet46.Range("ESOP.TaxPrevAY3").Locked = False
        Sheet46.Range("ESOP.TaxPrevAY3").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.SecType3").Locked = False
        Sheet46.Range("ESOP.SecType3").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.CeasedEmployee3").Locked = False
        Sheet46.Range("ESOP.CeasedEmployee3").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.DateOfCeasing3").Locked = False
        Sheet46.Range("ESOP.DateOfCeasing3").Interior.Color = "&HCCFFCC"
        
                Sheet46.Range("ESOP.TaxPrevAY4").Locked = False
        Sheet46.Range("ESOP.TaxPrevAY4").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.SecType4").Locked = False
        Sheet46.Range("ESOP.SecType4").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.CeasedEmployee4").Locked = False
        Sheet46.Range("ESOP.CeasedEmployee4").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.DateOfCeasing4").Locked = False
        Sheet46.Range("ESOP.DateOfCeasing4").Interior.Color = "&HCCFFCC"
        
                Sheet46.Range("ESOP.TaxPrevAY5").Locked = False
        Sheet46.Range("ESOP.TaxPrevAY5").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.SecType5").Locked = False
        Sheet46.Range("ESOP.SecType5").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.CeasedEmployee5").Locked = False
        Sheet46.Range("ESOP.CeasedEmployee5").Interior.Color = "&HCCFFCC"
                
        Sheet46.Range("ESOP.DateOfCeasing5").Locked = False
        Sheet46.Range("ESOP.DateOfCeasing5").Interior.Color = "&HCCFFCC"
'----
        UnLockESOP
        Sheet46.Protect Password:=getmsgstate
    End If
End Sub

Sub LockESOP()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet46.Unprotect Password:=sPassword

        rngname = "ESOPCurrAY||ESOP.Date||ESOP.AttributedTax"
        LOCKRANGEESOP (rngname)

    Sheet46.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnLockESOP()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet46.Unprotect Password:=sPassword

         rngname = "ESOPCurrAY||ESOP.Date||ESOP.AttributedTax"
        UNLOCKRANGEESOP (rngname)

    Sheet46.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub LOCKRANGEESOP(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet46.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub

Sub UNLOCKRANGEESOP(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet46.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub



Sub LockUnlock80()
On Error Resume Next

    If UCase(Mid(Trim(Sheet1.Range("sheet1.ReturnFileSec").value), 1, 6)) = "139(4)" Then
    
    'Or (Sheet1.Range("sheet1.ResidentialStatus1").value) = "(Select)" Then
        Lock80
    Else
        UnLock80

    End If
End Sub
'Change.23.01.2023.102.Sch80F====
Sub Lock80()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet37.Unprotect Password:=sPassword

'        rngname = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura||IB80.DeductFruitVegUs80_IB_11A||IC80.DeductInSikkim||IC80.DeductInHimachalP||IC80.DeductInUttaranchal"
        
        'Added by Aavula Naresh for AY 2025-26
'        rngname = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura||IB80.DeductFruitVegUs80_IB_11A"
        
        rngname = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura||IC80.Sikkim||IB80.DeductFruitVegUs80_IB_11A"
        LOCKRANGE80 (rngname)

    Sheet37.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnLock80()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet37.Unprotect Password:=sPassword

'         rngname = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura||IB80.DeductFruitVegUs80_IB_11A||IC80.DeductInSikkim||IC80.DeductInHimachalP||IC80.DeductInUttaranchal"
        
        'Added by Aavula Naresh for A.Y 2025-26
        
'         rngname = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura||IB80.DeductFruitVegUs80_IB_11A"
         rngname = "IA80.DeductProfUs80_IA_4_i||IA80.DeductProfUs80_IA_4_ii||IA80.DeductProfUs80_IA_4_iii||IA80.DeductProfUs80_IA_4_iv||IB80.DeductMultiplexUs80_IB_7A||IB80.DeductConvCentUs80_IB_7B||IB80.DeductMinOilUs80_IB_9||IB80.DeductHousUs80_IB_10||IB80.DeductFoodGrainUs80_IB_11A||IB80.DeductRurHospUs80_IB_11B||IB80.DeductHospAnyAreaUs80IB_11C||IC80.Assam||IC80.ArunachalPradesh||IC80.Manipur||IC80.Mizoram||IC80.Meghalaya||IC80.Nagaland||IC80.Tripura||IC80.Sikkim||IB80.DeductFruitVegUs80_IB_11A"
       
        '''''''
        UNLOCKRANGE80 (rngname)

    Sheet37.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
'Change End=====
Sub LOCKRANGE80(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet37.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub
Sub UNLOCKRANGE80(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet37.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub
Sub LockUnlockSalary89A()
sPassword = getmsgstate
        Sheet39.Unprotect Password:=sPassword
If UCase(Mid(Trim(Sheet1.Range("sheet1.ResidentialStatus1").value), 1, 3)) = "NRI" Then
        Sheet39.Range("Increliefus89A").Locked = True
        Sheet39.Range("Increliefus89A").Interior.Color = "&HD8D8D8"
        Sheet39.Range("Increliefus89A").value = ""
        Sheet39.Protect Password:=sPassword
    Else
'Change-26.05.2023.101 -> Discussed with Aishwarya on 24.05.2023
        If (Sheet1.Range("sheet1.Status").value <> "H-HUF") Then
         Sheet39.Unprotect Password:=sPassword
            Sheet39.Range("Increliefus89A").Locked = False
            Sheet39.Range("Increliefus89A").Interior.Color = "&HCCFFCC"
            Sheet39.Protect Password:=sPassword
'---
        End If
    End If
End Sub
Sub LockUnlockSalary()
If UCase(Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1)) = "H" Then
    'Or (Sheet1.Range("sheet1.ResidentialStatus1").value) = "(Select)" Then
        Lock_Salary

        sPassword = getmsgstate
        Sheet39.Unprotect Password:=sPassword

        Sheet39.Range("Salary.Nature1").Locked = True
        Sheet39.Range("Salary.Nature1").Interior.Color = "&HD8D8D8"
        Sheet39.Range("Salary.Nature1").value = ""

        Sheet39.Range("Salary.Description1").Locked = True
        Sheet39.Range("Salary.Description1").Interior.Color = "&HD8D8D8"
        Sheet39.Range("Salary.Description1").value = ""

        Sheet39.Range("Salary.Amount1").Locked = True
        Sheet39.Range("Salary.Amount1").Interior.Color = "&HD8D8D8"
        Sheet39.Range("Salary.Amount1").value = ""

        Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
        Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
        Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
            
        Sheet39.Range("Increliefus89A").Locked = True
        Sheet39.Range("Increliefus89A").Interior.Color = "&HD8D8D8"
        Sheet39.Range("Increliefus89A").value = ""
        
        Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
        Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
        Sheet39.Range("SAL.DeductionUnderSection1").value = ""
        AddTotalSalary1
        Sheet39.AddTotalSalary

        Sheet39.Protect Password:=sPassword
    Else
        UnLock_Salary

        sPassword = getmsgstate
        Sheet39.Unprotect Password:=sPassword

        Sheet39.Range("Salary.Nature1").Locked = False
        Sheet39.Range("Salary.Nature1").Interior.Color = "&HCCFFCC"
'        Sheet39.Range("Salary.Nature1").value = "(Select)"

        Sheet39.Range("Salary.Amount1").Locked = False
        Sheet39.Range("Salary.Amount1").Interior.Color = "&HCCFFCC"

        Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
        Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
        
        Sheet39.Range("Increliefus89A").Locked = False
        Sheet39.Range("Increliefus89A").Interior.Color = "&HCCFFCC"
        
        Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
        Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"

        Sheet39.Protect Password:=sPassword
    End If
    
     'Malli_03/10/2025
    SalaryExemptAllowance1013A
    '-------------------
    
End Sub
Sub Lock_Salary()
Dim i As Long
Dim scount, rangenamestring As Variant
Dim chcell  As Range
Application.EnableEvents = False
'Change-23.05.2023.101 -> Issue number 2 Production build 1
'   rangenamestring = "SAL.NameOfEmployer||sheet.EmployerCategory||SAL.PANofEmployer||SAL.AddrDetail||SAL.StateCode||SAL.CityOrTownOrDistrict||SAL.PinCode||Salary.Section||Salary.salaryDescription||Salary.SalaryAmount||Salary.ValueSection||Salary.ValueDescription||Salary.ValueAmount||Salary.ProfitSection||Salary.ProfitDescription||Salary.ProfitAmount||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A"
    rangenamestring = "SAL.NameOfEmployer||sheet.EmployerCategory||SAL.PANofEmployer||SAL.AddrDetail||SAL.StateCode||SAL.CityOrTownOrDistrict||SAL.PinCode||Salary.Section||Salary.salaryDescription||Salary.SalaryAmount||Salary.ValueSection||Salary.ValueDescription||Salary.ValueAmount||Salary.ProfitSection||Salary.ProfitDescription||Salary.ProfitAmount||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||IncomeTaxablePrevYear89A"
    rangenamestring = Split(rangenamestring, "||")
  sPassword = getmsgstate
   Sheet39.Unprotect Password:=sPassword
    For scount = 1 To Sheet39.Range("SalarySectionCount").value
        For i = 0 To UBound(rangenamestring)
             For Each chcell In Sheet39.Range(rangenamestring(i) & scount)
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
    Next
    'Application.EnableEvents = True
    Sheet39.Protect Password:=sPassword
End Sub

Sub UnLock_Salary()
Dim i As Long
Dim chcell As Range
Dim scount, rangenamestring As Variant
 Application.EnableEvents = False
 'Change-23.05.2023.101 -> Issue number 2 Production build 1
' rangenamestring = "SAL.NameOfEmployer||sheet.EmployerCategory||SAL.PANofEmployer||SAL.AddrDetail||SAL.StateCode||SAL.CityOrTownOrDistrict||SAL.PinCode||Salary.Section||Salary.SalaryAmount||Salary.ValueSection||Salary.ValueAmount||Salary.ProfitSection||Salary.ProfitAmount||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A"
 rangenamestring = "SAL.NameOfEmployer||sheet.EmployerCategory||SAL.PANofEmployer||SAL.AddrDetail||SAL.StateCode||SAL.CityOrTownOrDistrict||SAL.PinCode||Salary.Section||Salary.SalaryAmount||Salary.ValueSection||Salary.ValueAmount||Salary.ProfitSection||Salary.ProfitAmount||IncomeNotified89A_AmountUS||IncomeNotified89A_AmountUK||IncomeNotified89A_AmountCan||IncomeNotifiedOther89A||IncomeTaxablePrevYear89A"
    rangenamestring = Split(rangenamestring, "||")
  sPassword = getmsgstate
   Sheet39.Unprotect Password:=sPassword
    For scount = 1 To Sheet39.Range("SalarySectionCount").value
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet39.Range(rangenamestring(i) & scount)
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
    Next
     Application.EnableEvents = True
       Sheet39.Protect Password:=sPassword
End Sub
Sub LockUnlockFSI()
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 1) = "(" Then
            Lock_FSI
Else
            UnLock_FSI
End If
End Sub

Sub Lock_FSI()
Dim i As Long
Dim Fsicount, rangenamestring As Variant
Dim chcell  As Range
Application.EnableEvents = False
   rangenamestring = "FSI_CountryCode||FSI_TaxIdentificationNo||FSI_IncFromSal||FSI_TXNOSAL||FSI_TXNISAL||FSI_DTAASAL||FSI_IncFromHP||FSI_TXNOHP||FSI_TXNIHP||FSI_DTAAHP||FSI_IncFromBP||FSI_TXNOBP||FSI_TXNIBP||FSI_DTAABP||FSI_IncCapGain||FSI_TXNOCapGain||FSI_TXnICapGain||FSI_DTAACapGain||FSI_IncOthSrc||FSI_TXNOOthSrc||FSI_TXNIOthSrc||FSI_DTAAOthSrc"
    rangenamestring = Split(rangenamestring, "||")
  sPassword = getmsgstate
   Sheet40.Unprotect Password:=sPassword
    For Fsicount = 1 To Sheet40.Range("FSI_Count").value
        For i = 0 To UBound(rangenamestring)
             For Each chcell In Sheet40.Range(rangenamestring(i) & Fsicount)
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
    Next
'Konda-------SIT-95967---28/08/2025----
  If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
  Sheet27.Unprotect Password:=sPassword
    Sheet27.Range("TR_Country").ClearContents
    Sheet27.Unprotect Password:=sPassword
    Sheet27.Range("TR_TIN").ClearContents
    Sheet27.Unprotect Password:=sPassword
    Sheet27.Range("TR_TaxPaidOutsideIndia").value = 0
    Sheet27.Unprotect Password:=sPassword
    Sheet27.Range("TR_TaxReliefOutsideIndia").value = 0
    Sheet27.Unprotect Password:=sPassword
    Sheet27.Range("TR_ReliefClaimedUsSection").value = "(Select)"
End If
    '---------------SIT-95967--------------
    
    'Application.EnableEvents = True
    Sheet40.Protect Password:=sPassword
End Sub

Sub UnLock_FSI()
Dim i As Long
Dim chcell As Range
Dim Fsicount, rangenamestring As Variant
 Application.EnableEvents = False
 rangenamestring = "FSI_CountryCode||FSI_TaxIdentificationNo||FSI_IncFromSal||FSI_TXNOSAL||FSI_TXNISAL||FSI_DTAASAL||FSI_IncFromHP||FSI_TXNOHP||FSI_TXNIHP||FSI_DTAAHP||FSI_IncFromBP||FSI_TXNOBP||FSI_TXNIBP||FSI_DTAABP||FSI_IncCapGain||FSI_TXNOCapGain||FSI_TXnICapGain||FSI_DTAACapGain||FSI_IncOthSrc||FSI_TXNOOthSrc||FSI_TXNIOthSrc||FSI_DTAAOthSrc"
    rangenamestring = Split(rangenamestring, "||")
sPassword = getmsgstate
   Sheet40.Unprotect Password:=sPassword
    For Fsicount = 1 To Sheet40.Range("FSI_Count").value
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet40.Range(rangenamestring(i) & Fsicount)
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
    Next
     Application.EnableEvents = True
       Sheet40.Protect Password:=sPassword
End Sub

Sub LockUnlockTR()
On Error Resume Next
Application.ScreenUpdating = False
 Application.EnableEvents = False
'----------------Unlock Password-------------------START---
    sPassword = getmsgstate
   Sheet27.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 1) = "(" Then
Sheet27.Range("TR_TaxPaidOutsideIndFlg").value = ""
Sheet27.Range("TR_TaxPaidOutsideIndFlg").Locked = True
Sheet27.Range("TR_TaxPaidOutsideIndFlg").Interior.Color = (&HD8D8D8)


Sheet27.Range("TR_AmtTaxRefunded").value = ""
Sheet27.Range("TR_AmtTaxRefunded").Locked = True
Sheet27.Range("TR_AmtTaxRefunded").Interior.Color = (&HD8D8D8)

Sheet27.Range("TR_AssmtYrTaxRelief").value = ""
Sheet27.Range("TR_AssmtYrTaxRelief").Locked = True
Sheet27.Range("TR_AssmtYrTaxRelief").Interior.Color = (&HD8D8D8)

Else
Sheet27.Range("TR_TaxPaidOutsideIndFlg").value = "(Select)"
Sheet27.Range("TR_TaxPaidOutsideIndFlg").Locked = False
Sheet27.Range("TR_TaxPaidOutsideIndFlg").Interior.Color = (&HCCFFCC)


Sheet27.Range("TR_AmtTaxRefunded").value = ""
Sheet27.Range("TR_AmtTaxRefunded").Locked = False
Sheet27.Range("TR_AmtTaxRefunded").Interior.Color = (&HCCFFCC)

Sheet27.Range("TR_AssmtYrTaxRelief").value = ""
Sheet27.Range("TR_AssmtYrTaxRelief").Locked = False
Sheet27.Range("TR_AssmtYrTaxRelief").Interior.Color = (&HCCFFCC)


End If
'----------------Lock Password-------------------START---
   Sheet27.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

 'Application.EnableEvents = True
 Application.ScreenUpdating = True
End Sub

'Sub lockunlock_OS_Stch_Ltcg()
'Application.EnableEvents = False
'    Sheet13.Unprotect Password:=sPassword
'    Sheet14.Unprotect Password:=sPassword
'    Dim temp As Long
'    Dim temp1 As Long
'    Dim Temp2 As Long
'     For temp = 1 To Sheet14.Range("os.TRC").count
'      If Sheet14.Range("os.TRC").Cells.item(temp).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'        Sheet14.Range("os.ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet14.Range("os.Rate_DTAA").Cells.item(temp).value, Sheet14.Range("os.Rate_ITAct").Cells.item(temp).value)
'         Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
'        Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.ColorIndex = 2
'
'
'        'added by Chetan C M on for AY 2026-27
'        '===start
''*******************************************************************************
''SIT -116737 commented by sadineni on 03/04/2026
'        '---Table 2d---
''        Sheet14.Range("os.SourceDescription").ClearContents
''        Sheet14.Range("os.SourceDescription").Locked = True
''        Sheet14.Range("os.SourceDescription").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.SourceAmount").ClearContents
''        Sheet14.Range("os.SourceAmount").Locked = True
''        Sheet14.Range("os.SourceAmount").Interior.Color = "&HD8D8D8"
''        '----2d
''
''        '---Table 2e---
''        Sheet14.Range("os.PTIDescription").ClearContents
''        Sheet14.Range("os.PTIDescription").Locked = True
''        Sheet14.Range("os.PTIDescription").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.PTIAmount").ClearContents
''        Sheet14.Range("os.PTIAmount").Locked = True
''        Sheet14.Range("os.PTIAmount").Interior.Color = "&HD8D8D8"
'        '---2e
''*******************************************************************************
'        '---Table 2f---
''        Sheet14.Range("os.Amount_income").ClearContents
''        Sheet14.Range("os.Amount_income").Locked = True
''        Sheet14.Range("os.Amount_income").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.NatureOfIncome").ClearContents
''        Sheet14.Range("os.NatureOfIncome").Locked = True
''        Sheet14.Range("os.NatureOfIncome").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.Country").ClearContents
''        Sheet14.Range("os.Country").Locked = True
''        Sheet14.Range("os.Country").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.Article_DTAA").ClearContents
''        Sheet14.Range("os.Article_DTAA").Locked = True
''        Sheet14.Range("os.Article_DTAA").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.Rate_DTAA").ClearContents
''        Sheet14.Range("os.Rate_DTAA").Locked = True
''        Sheet14.Range("os.Rate_DTAA").Interior.Color = "&HD8D8D8"
''
''        Sheet14.Range("os.SectionAct").ClearContents
''        Sheet14.Range("os.SectionAct").Locked = True
''        Sheet14.Range("os.SectionAct").Interior.Color = "&HD8D8D8"
'
''        Sheet14.Range("os.Rate_ITAct").ClearContents
''        Sheet14.Range("os.Rate_ITAct").Locked = True
''        Sheet14.Range("os.Rate_ITAct").Interior.Color = "&HD8D8D8"
'        '---2f
'        '===end
'
'     Else
'       Sheet14.Range("os.ApplRate").Cells.item(temp).ClearContents
'       Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
'       Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
'
'       'added by Chetan C M on for AY 2026-27
'    '===start
'
'
''************************************************************************************
'        '---Table 2d---
''SIT -116737 commented by sadineni on 03/04/2026
''        Sheet14.Range("os.SourceDescription").ClearContents
''        Sheet14.Range("os.SourceDescription").Locked = False
''        Sheet14.Range("os.SourceDescription").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.SourceAmount").ClearContents
''        Sheet14.Range("os.SourceAmount").Locked = False
''        Sheet14.Range("os.SourceAmount").Interior.Color = "&HCCFFCC"
''        '----2d
''
''        '---Table 2e---
''        Sheet14.Range("os.PTIDescription").ClearContents
''        Sheet14.Range("os.PTIDescription").Locked = False
''        Sheet14.Range("os.PTIDescription").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.PTIAmount").ClearContents
''        Sheet14.Range("os.PTIAmount").Locked = False
''        Sheet14.Range("os.PTIAmount").Interior.Color = "&HCCFFCC"
'        '---2e
''***********************************************************************************
'        '---Table 2f---
''        Sheet14.Range("os.Amount_income").ClearContents
''        Sheet14.Range("os.Amount_income").Locked = False
''        Sheet14.Range("os.Amount_income").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.NatureOfIncome").ClearContents
''        Sheet14.Range("os.NatureOfIncome").Locked = False
''        Sheet14.Range("os.NatureOfIncome").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.Country").ClearContents
''        Sheet14.Range("os.Country").Locked = False
''        Sheet14.Range("os.Country").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.Article_DTAA").ClearContents
''        Sheet14.Range("os.Article_DTAA").Locked = False
''        Sheet14.Range("os.Article_DTAA").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.Rate_DTAA").ClearContents
''        Sheet14.Range("os.Rate_DTAA").Locked = False
''        Sheet14.Range("os.Rate_DTAA").Interior.Color = "&HCCFFCC"
''
''        Sheet14.Range("os.SectionAct").ClearContents
''        Sheet14.Range("os.SectionAct").Locked = False
''        Sheet14.Range("os.SectionAct").Interior.Color = "&HCCFFCC"
'
''        Sheet14.Range("os.Rate_ITAct").ClearContents
''        Sheet14.Range("os.Rate_ITAct").Locked = False
''        Sheet14.Range("os.Rate_ITAct").Interior.Color = "&HCCFFCC"
'        '---2f
'
'    '===end
'
'    End If
'   Next
'
'
'
'       For temp1 = 1 To Sheet13.Range("STCG.A8_CertiStatus").count
'      If Sheet13.Range("STCG.A8_CertiStatus").Cells.item(temp1).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).value = WorksheetFunction.Min(Sheet13.Range("STCG.A8_RatePTreaty").Cells.item(temp1).value, Sheet13.Range("STCG.A8_RateIT").Cells.item(temp1).value)
'         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Locked = True
'        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Interior.ColorIndex = 2
'     Else
'
'        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Locked = True
'       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Interior.Color = "&HD8D8D8"
'       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).ClearContents
'
'    End If
'   Next
'
'       For Temp2 = 1 To Sheet13.Range("LTCG.B10_CertiStatus").count
'      If Sheet13.Range("LTCG.B10_CertiStatus").Cells.item(Temp2).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
'         Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).value = WorksheetFunction.Min(Sheet13.Range("LTCG.B10_RatepTreaty").Cells.item(Temp2).value, Sheet13.Range("LTCG.B10_RatePITAct").Cells.item(Temp2).value)
'         Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Locked = True
'        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Interior.ColorIndex = 2
'     Else
'
'       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).ClearContents
'       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Locked = True
'       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Interior.Color = "&HD8D8D8"
'    End If
'   Next
'
'   Sheet13.Protect Password:=sPassword
'    Sheet14.Protect Password:=sPassword
'
'    Application.EnableEvents = True
'End Sub


Sub lockunlock_OS_Stch_Ltcg()
Application.EnableEvents = False
    Sheet13.Unprotect Password:=sPassword
    Sheet14.Unprotect Password:=sPassword
    Dim temp As Long
    Dim temp1 As Long
    Dim Temp2 As Long
     For temp = 1 To Sheet14.Range("os.TRC").count
      If Sheet14.Range("os.TRC").Cells.item(temp).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
        Sheet14.Range("os.ApplRate").Cells.item(temp).value = WorksheetFunction.Min(Sheet14.Range("os.Rate_DTAA").Cells.item(temp).value, Sheet14.Range("os.Rate_ITAct").Cells.item(temp).value)
         Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
        Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.ColorIndex = 2
     Else
       Sheet14.Range("os.ApplRate").Cells.item(temp).ClearContents
       Sheet14.Range("os.ApplRate").Cells.item(temp).Locked = True
       Sheet14.Range("os.ApplRate").Cells.item(temp).Interior.Color = "&HD8D8D8"
    End If
   Next



       For temp1 = 1 To Sheet13.Range("STCG.A8_CertiStatus").count
      If Sheet13.Range("STCG.A8_CertiStatus").Cells.item(temp1).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).value = WorksheetFunction.Min(Sheet13.Range("STCG.A8_RatePTreaty").Cells.item(temp1).value, Sheet13.Range("STCG.A8_RateIT").Cells.item(temp1).value)
         Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Locked = True
        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Interior.ColorIndex = 2
     Else

        Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Locked = True
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).Interior.Color = "&HD8D8D8"
       Sheet13.Range("STCG.A8_ApplRate").Cells.item(temp1).ClearContents

    End If
   Next

       For Temp2 = 1 To Sheet13.Range("LTCG.B10_CertiStatus").count
      If Sheet13.Range("LTCG.B10_CertiStatus").Cells.item(Temp2).value = "Yes" Or Mid(Sheet1.Range("Sheet1.ResidentialStatus1"), 1, 3) <> "NRI" Then
         Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).value = WorksheetFunction.Min(Sheet13.Range("LTCG.B10_RatepTreaty").Cells.item(Temp2).value, Sheet13.Range("LTCG.B10_RatePITAct").Cells.item(Temp2).value)
         Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Locked = True
        Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Interior.ColorIndex = 2
     Else

       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).ClearContents
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Locked = True
       Sheet13.Range("LTCG.B10_ApplRate").Cells.item(Temp2).Interior.Color = "&HD8D8D8"
    End If
   Next

   Sheet13.Protect Password:=sPassword
    Sheet14.Protect Password:=sPassword

    Application.EnableEvents = True
End Sub










Sub LOCK80EE_A_B()
On Error Resume Next
Application.ScreenUpdating = False
 'Application.EnableEvents = False
'----------------Unlock Password-------------------START---
'   sPassword = Common.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----


'Commented by sudha_2025-26

'Sheet19.Range("scvia.Section80EEA").value = ""
'Sheet19.Range("scvia.Section80EEA").Locked = True
'Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HD8D8D8)
'
'
'Sheet19.Range("scvia.Section80EEB").value = ""
'Sheet19.Range("scvia.Section80EEB").Locked = True
'Sheet19.Range("scvia.Section80EEB").Interior.Color = (&HD8D8D8)

'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

 'Application.EnableEvents = True
 Application.ScreenUpdating = True
End Sub

Sub UNLOCK80EE_A_B()
On Error Resume Next
Application.ScreenUpdating = False
 'Application.EnableEvents = False
'----------------Unlock Password-------------------START---
'   sPassword = Common.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----


'Commented by sudha_2025-26

'Sheet19.Range("scvia.Section80EEA").value = ""
'Sheet19.Range("scvia.Section80EEA").Locked = False
'Sheet19.Range("scvia.Section80EEA").Interior.Color = (&HCCFFCC)
'
'
'Sheet19.Range("scvia.Section80EEB").value = ""
'Sheet19.Range("scvia.Section80EEB").Locked = False
'Sheet19.Range("scvia.Section80EEB").Interior.Color = (&HCCFFCC)

'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

 'Application.EnableEvents = True
 Application.ScreenUpdating = True
End Sub


Sub LockUnlock80D()
Application.EnableEvents = False
Sheet55.Unprotect Password:=getmsgstate
If Mid(Trim(Sheet1.Range("sheet1.Status")), 1, 1) = "H" Then

        Sheet55.Range("DropDown_ValueOf_SC_80D").ClearContents
        Sheet55.Range("DropDown_ValueOf_SC_80D").Locked = True
        Sheet55.Range("DropDown_ValueOf_SC_80D").Interior.Color = (&HD8D8D8)

        Sheet55.Range("Parents_80D").ClearContents
        Sheet55.Range("Parents_80D").Locked = True
        Sheet55.Range("Parents_80D").Interior.Color = (&HD8D8D8)

'        Sheet55.Range("Health_Insurance2_80D").ClearContents
'        Sheet55.Range("Health_Insurance2_80D").Locked = True
'        Sheet55.Range("Health_Insurance2_80D").Interior.Color = (&HD8D8D8)

        Sheet55.Range("Preventive_Health2_80D").ClearContents
        Sheet55.Range("Preventive_Health2_80D").Locked = True
        Sheet55.Range("Preventive_Health2_80D").Interior.Color = (&HD8D8D8)

         Sheet55.Range("Parents_SC_80D").ClearContents
        Sheet55.Range("Parents_SC_80D").Locked = True
        Sheet55.Range("Parents_SC_80D").Interior.Color = (&HD8D8D8)

'        Sheet55.Range("Health_Insurance3_80D").ClearContents
'        Sheet55.Range("Health_Insurance3_80D").Locked = True
'        Sheet55.Range("Health_Insurance3_80D").Interior.Color = (&HD8D8D8)

        Sheet55.Range("Preventive_Health3_80D").ClearContents
        Sheet55.Range("Preventive_Health3_80D").Locked = True
        Sheet55.Range("Preventive_Health3_80D").Interior.Color = (&HD8D8D8)

        Sheet55.Range("Medical_Expenditure2_80D").ClearContents
        Sheet55.Range("Medical_Expenditure2_80D").Locked = True
        Sheet55.Range("Medical_Expenditure2_80D").Interior.Color = (&HD8D8D8)
        
        
        
        
        'Malli_05/07/2025
        
        Sheet55.Range("NameInsurerA2.80D").ClearContents
        Sheet55.Range("NameInsurerA2.80D").Locked = True
        Sheet55.Range("NameInsurerA2.80D").Interior.Color = (&HD8D8D8)
        
        Sheet55.Range("PolicyNumA2.80D").ClearContents
        Sheet55.Range("PolicyNumA2.80D").Locked = True
        Sheet55.Range("PolicyNumA2.80D").Interior.Color = (&HD8D8D8)
        
        Sheet55.Range("AmtA2.80D").ClearContents
        Sheet55.Range("AmtA2.80D").Locked = True
        Sheet55.Range("AmtA2.80D").Interior.Color = (&HD8D8D8)
        
        
        Sheet55.Range("NameInsurerB2.80D").ClearContents
        Sheet55.Range("NameInsurerB2.80D").Locked = True
        Sheet55.Range("NameInsurerB2.80D").Interior.Color = (&HD8D8D8)
        
        Sheet55.Range("PolicyNumB2.80D").ClearContents
        Sheet55.Range("PolicyNumB2.80D").Locked = True
        Sheet55.Range("PolicyNumB2.80D").Interior.Color = (&HD8D8D8)
        
        Sheet55.Range("AmtB2.80D").ClearContents
        Sheet55.Range("AmtB2.80D").Locked = True
        Sheet55.Range("AmtB2.80D").Interior.Color = (&HD8D8D8)
        
        '----------------------------

         Else

        Sheet55.Range("DropDown_ValueOf_SC_80D").value = "(Select)"
        Sheet55.Range("DropDown_ValueOf_SC_80D").Locked = False
        Sheet55.Range("DropDown_ValueOf_SC_80D").Interior.Color = (&HCCFFCC)


        Sheet55.Range("Parents_80D").Locked = True
        Sheet55.Range("Parents_80D").Interior.Color = 13434879
'        Sheet55.range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Sheet55.Range("Parents_80D").formula = "=MIN(25000,SUM(Health_Insurance2_80D,Preventive_Health2_80D))"

'        Sheet55.Range("Health_Insurance2_80D").Locked = False
'        Sheet55.Range("Health_Insurance2_80D").Interior.Color = (&HCCFFCC)

        Sheet55.Range("Preventive_Health2_80D").Locked = False
        Sheet55.Range("Preventive_Health2_80D").Interior.Color = (&HCCFFCC)


        Sheet55.Range("Parents_SC_80D").Locked = True
        Sheet55.Range("Parents_SC_80D").Interior.Color = 13434879
'        Sheet55.range("Parents_80D").Font.Color = RGB(0, 0, 0)
        Sheet55.Range("Parents_SC_80D").formula = "=MIN(50000,SUM(Health_Insurance3_80D,Preventive_Health3_80D,Medical_Expenditure2_80D))"

'        Sheet55.Range("Health_Insurance3_80D").Locked = False
'        Sheet55.Range("Health_Insurance3_80D").Interior.Color = (&HCCFFCC)

        Sheet55.Range("Preventive_Health3_80D").Locked = False
        Sheet55.Range("Preventive_Health3_80D").Interior.Color = (&HCCFFCC)

        Sheet55.Range("Medical_Expenditure2_80D").Locked = False
        Sheet55.Range("Medical_Expenditure2_80D").Interior.Color = (&HCCFFCC)
        
        'Malli_07/05/2025
         
        Sheet55.Range("NameInsurerA2.80D").Locked = False
        Sheet55.Range("NameInsurerA2.80D").Interior.Color = (&HCCFFCC)
        
        Sheet55.Range("PolicyNumA2.80D").Locked = False
        Sheet55.Range("PolicyNumA2.80D").Interior.Color = (&HCCFFCC)
        
        Sheet55.Range("AmtA2.80D").Locked = False
        Sheet55.Range("AmtA2.80D").Interior.Color = (&HCCFFCC)
        
        Sheet55.Range("NameInsurerB2.80D").Locked = False
        Sheet55.Range("NameInsurerB2.80D").Interior.Color = (&HCCFFCC)
        
        Sheet55.Range("PolicyNumB2.80D").Locked = False
        Sheet55.Range("PolicyNumB2.80D").Interior.Color = (&HCCFFCC)
        
        Sheet55.Range("AmtB2.80D").Locked = False
        Sheet55.Range("AmtB2.80D").Interior.Color = (&HCCFFCC)
         '-----------------

  End If
  'Application.EnableEvents = True
            Sheet55.Protect Password:=getmsgstate
End Sub



Function getRangeName(ByVal Target As Range) As String
    Dim start As Long
    start = Len(name)

    If InStr(1, Target.name.name, "'") > 0 Then
        start = start + 3 + 1
    Else
        start = start + 1 + 1
    End If

    If InStr(1, Target.name.name, ActiveSheet.name) = 0 Then
        start = 1
    End If
    getRangeName = Mid(Target.name.name, start)
End Function


Sub Lock80U()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant



    sPassword = getmsgstate
    Sheet59.Unprotect Password:=sPassword

        rngname = "Naturedisability_80U||Disability_80U||DatefilingFm10IA_80U||AckNoFm10IAfiled_80U||FormAckNum11A_80U||UDIDNum_80U"
        LOCKRANGE80U (rngname)

    Sheet59.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub LOCKRANGE80U(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet59.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub

Sub UnLock80U()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant



    sPassword = getmsgstate
    Sheet59.Unprotect Password:=sPassword

        rngname = "Naturedisability_80U||Disability_80U||DatefilingFm10IA_80U||AckNoFm10IAfiled_80U||FormAckNum11A_80U||UDIDNum_80U"
        UNLOCKRANGE80U (rngname)

    Sheet59.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKRANGE80U(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet59.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub
Sub Lock80DD()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet59.Unprotect Password:=sPassword

        rngname = "Naturedisability_80DD||Disability_80DD||Typedependent_80DD||PANdependent_80DD||Aadhaardependent_80DD||DatefilingFm10IA_80DD||AckNoFm10IAfiled_80DD||FormAckNum11A_80DD||UDIDNum_80DD"
        
LOCKRANGE80DD (rngname)

    Sheet59.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub LOCKRANGE80DD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet59.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
End Sub

Sub UnLock80DD()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim rngname As Variant

    sPassword = getmsgstate
    Sheet59.Unprotect Password:=sPassword

        rngname = "Naturedisability_80DD||Disability_80DD||Typedependent_80DD||PANdependent_80DD||Aadhaardependent_80DD||DatefilingFm10IA_80DD||AckNoFm10IAfiled_80DD||FormAckNum11A_80DD||UDIDNum_80DD"
        UNLOCKRANGE80DD (rngname)

    Sheet59.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKRANGE80DD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet59.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                End If
            Next
        Next
End Sub

Function StateMatchesPin(StateCode As Variant) As Boolean
    StateMatchesPin = False
    Dim state1, State2 As Variant
    
    
    state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, SheetALL.Range("Duplicate_V"), 2, False)
    State2 = Application.VLookup(Sheet1.Range("sheet1.PinCode").value, SheetALL.Range("Duplicate_V"), 3, False)

    If IsError(state1) Or IsError(State2) Then
    StateMatchesPin = False
    Exit Function
    End If
    
    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = State2 Then
    StateMatchesPin = True
    End If
    
End Function
Public Function ChkMaxDate_filing_2425(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

    ChkMaxDate_filing_2425 = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        If Year > 2024 Then
            ChkMaxDate_filing_2425 = False
            Exit Function
        Else
            If Year = 2024 Then
                If month > 12 Then
                    ChkMaxDate_filing_2425 = False
                    Exit Function
                Else
'                    If month = 8 Then
                    If month = 12 Then

                        If dat > 15 Then
                           ChkMaxDate_filing_2425 = False
                            Exit Function
                        Else
'                            If dat = 1 Then
                            If dat = 15 Then
                               ChkMaxDate_filing_2425 = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function
'Public Function ChkMaxDate_filing_2425(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
'On Error Resume Next
'Dim Year, month, dat As Variant
'
'    ChkMaxDate_filing_2425 = True
'     If Len(dateEntered) > 0 Then
'
'     Year = val(Mid(dateEntered, 7, 4))
'     month = val(Mid(dateEntered, 4, 2))
'     dat = val(Mid(dateEntered, 1, 2))
'
''        If Year > 2023 Then
'        'PAG_C1 AY 2024-25 Change
'        If Year > 2024 Then
'            ChkMaxDate_filing_2425 = False
'            Exit Function
'        Else
'            If Year = 2024 Then
'                If month > 8 Then
'                    ChkMaxDate_filing_2425 = False
'                    Exit Function
'                Else
'                    If month = 8 Then
'                        If dat > 1 Then
'                           ChkMaxDate_filing_2425 = False
'                            Exit Function
'                        Else
'                            If dat = 1 Then
'                               ChkMaxDate_filing_2425 = False
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'     End If
'End Function

''''''''''''''''''
'Added by Aavula Naresh for AY 2025-26

'
'Sub LockEI4()
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Dim rngname As Variant
'
'
'
'    sPassword = getmsgstate
'    Sheet24.Unprotect Password:=sPassword
'
'        rngname = "SCHEI_DTAA_AmountOfIncome||SCHEI_DTAA_NatureOfIncome||SCHEI_DTAA_CountryCode||SCHEI_DTAA_Article||SCHEI_DTAA_HeadOfIncome||SCHEI_TRCObtained"
'        LOCKRANGEEI4 (rngname)
'
'    Sheet24.Protect Password:=sPassword
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'End Sub
'
'Sub LOCKRANGEEI4(rangenamestring As Variant)
'Dim i As Long
'Dim chcell As Range
'    rangenamestring = Split(rangenamestring, "||")
'        For i = 0 To UBound(rangenamestring)
'            For Each chcell In Sheet24.Range(rangenamestring(i))
'                If chcell.MergeCells Then
'                    chcell.MergeArea.Locked = True
'                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
'                    chcell.value = ""
'                Else
'                    chcell.Locked = True
'                    chcell.Interior.Color = "&HD8D8D8"
'                    chcell.value = ""
'                End If
'            Next
'        Next
'End Sub
'
'
'Sub UnLockEI4()
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Dim rngname As Variant
'
'
'
'    sPassword = getmsgstate
'    Sheet24.Unprotect Password:=sPassword
'
'
'        rngname = "SCHEI_DTAA_AmountOfIncome||SCHEI_DTAA_NatureOfIncome||SCHEI_DTAA_CountryCode||SCHEI_DTAA_Article||SCHEI_DTAA_HeadOfIncome||SCHEI_TRCObtained"
'
'        UNLOCKRANGEEI4 (rngname)
'
'    Sheet24.Protect Password:=sPassword
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'End Sub
'
'
'
'Sub UNLOCKRANGEEI4(rangenamestring As Variant)
'Dim i As Long
'Dim chcell As Range
'    rangenamestring = Split(rangenamestring, "||")
'        For i = 0 To UBound(rangenamestring)
'            For Each chcell In Sheet24.Range(rangenamestring(i))
'                If chcell.MergeCells Then
'                    chcell.MergeArea.Locked = False
'                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
'                Else
'                    chcell.Locked = False
'                    chcell.Interior.Color = "&HCCFFCC"
'                 If rangenamestring(i) = "SCHEI_DTAA_CountryCode" Then 'Added by Aavula 2025-26
'                 chcell.value = "(Select)"
'                 End If
'
'                End If
'            Next
'        Next
'End Sub
'Newly added by sai on 25/02/2025 for AY 2025-26
Sub LOCKRANGEOS_DTAA()
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant
rangenamestring = "os.Amount_income||os.NatureOfIncome||os.Country||os.Article_DTAA||os.Rate_DTAA||os.TRC||os.SectionAct||os.Rate_ITAct"
    rangenamestring = Split(rangenamestring, "||")
    Application.EnableEvents = False
    Sheet14.Unprotect Password:=getmsgstate
    
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet14.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
 Application.EnableEvents = True
'Sheet14.Protect Password:=getmsgstate
End Sub

Sub UNLOCKRANGEOS_DTAA()
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant

rangenamestring = "os.Amount_income||os.NatureOfIncome||os.Country||os.Article_DTAA||os.Rate_DTAA||os.TRC||os.SectionAct||os.Rate_ITAct"
    rangenamestring = Split(rangenamestring, "||")
    
    Application.EnableEvents = False
    Sheet14.Unprotect Password:=getmsgstate
    
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet14.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    If rangenamestring(i) = "os.NatureOfIncome" Or rangenamestring(i) = "os.Country" Or rangenamestring(i) = "os.TRC" Or rangenamestring(i) = "os.SectionAct" Then
                        chcell.value = "(Select)"
                    End If
                End If
            Next
        Next
Application.EnableEvents = True
'Sheet14.Protect Password:=getmsgstate
End Sub
'Added by sai on 12/03/2025 for AY 2025-26

Sub UNLOCKRANGECG_115ACA()
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant
'rangenamestring = "LTCG.Bii3a_FullConsideration||LTCG.Bii3a_FullConsideration_ii||LTCG.Bii3bi_AquisitCost||LTCG.Bii3bi_AquisitCost_ii||LTCG.Bii3bii_ImproveCost||LTCG.Bii3bii_ImproveCost_ii||LTCG.Bii3biii_ExpOnTrans||LTCG.Bii3biii_ExpOnTrans_ii||LTCG.B04di_ExemptionGrandTotal||LTCG.B04di_ExemptionGrandTotal_ii"
    'Newly updated by Bindu as per DE Sheet on 18th Feb 26
 rangenamestring = "LTCG.Bii3a_FullConsideration_ii||LTCG.Bii3bi_AquisitCost_ii||LTCG.Bii3bii_ImproveCost_ii||LTCG.Bii3biii_ExpOnTrans_ii||LTCG.B04di_ExemptionGrandTotal_ii"
   
    rangenamestring = Split(rangenamestring, "||")
    Application.EnableEvents = False
    Sheet13.Unprotect Password:=getmsgstate
    
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
                
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    chcell.value = ""
                
            Next
        Next
 Application.EnableEvents = True
Sheet13.Protect Password:=getmsgstate
End Sub
'Added by sai on 12/03/2025 for AY 2025-26

Sub LOCKRANGECG_115ACA()
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant
'rangenamestring = "LTCG.Bii3a_FullConsideration||LTCG.Bii3a_FullConsideration_ii||LTCG.Bii3bi_AquisitCost||LTCG.Bii3bi_AquisitCost_ii||LTCG.Bii3bii_ImproveCost||LTCG.Bii3bii_ImproveCost_ii||LTCG.Bii3biii_ExpOnTrans||LTCG.Bii3biii_ExpOnTrans_ii||LTCG.B04di_ExemptionGrandTotal||LTCG.B04di_ExemptionGrandTotal_ii"
'Updated by Bindu on 17th Feb 26 as per DE Sheet
  rangenamestring = "LTCG.Bii3a_FullConsideration_ii||LTCG.Bii3bi_AquisitCost_ii||LTCG.Bii3bii_ImproveCost_ii||LTCG.Bii3biii_ExpOnTrans_ii||LTCG.B04di_ExemptionGrandTotal_ii"
   
    rangenamestring = Split(rangenamestring, "||")
    Application.EnableEvents = False
    Sheet13.Unprotect Password:=getmsgstate
    
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
                
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                
            Next
        Next
 Application.EnableEvents = True
Sheet13.Protect Password:=getmsgstate
End Sub

Function CheckDateFilinginSheet_DOF1B(dob As Variant) As Boolean
On Error Resume Next
CheckDateFilinginSheet_DOF1B = True

    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDateDOF1B(dob) Then
            CheckDateFilinginSheet_DOF1B = False
            MsgBox ("""Please enter date in dd/mm/yyyy format"".")
            Exit Function
        End If
    End If
 End Function
 
 Sub AutoPopulateSecondaryAddress()
    Application.EnableEvents = False
 
    If Sheet1.Range("Secondary_Address").value = "Yes" Then
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceNo2").value = Sheet1.Range("sheet1.ResidenceNo").value
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.LocalityOrArea2").value = Sheet1.Range("sheet1.LocalityOrArea").value
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceName2").value = Sheet1.Range("sheet1.ResidenceName").value
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.RoadOrStreet2").value = Sheet1.Range("sheet1.RoadOrStreet").value
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").value = Sheet1.Range("sheet1.CityOrTownOrDistrict").value
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.PinCode").Locked = True Then
            Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Else
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.PinCode2").value = Sheet1.Range("sheet1.PinCode").value
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        End If
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.StateCode2").value = Sheet1.Range("sheet1.StateCode1").value
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CountryCode2").value = Sheet1.Range("sheet1.CountryCode").value
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("HASZIP").Locked = True Then
            Sheet1.Range("HASZIP2").MergeArea.ClearContents
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"
        Else
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("HASZIP2").value = Sheet1.Range("HASZIP").value
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
        End If
'         If Sheet1.Range("HASZIP").value = "Yes" Then
'            Sheet1.Range("sheet1.ZipCode2").value = "XXXXXX"
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
'        End If
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.ZipCode").Locked = True Then
            If Sheet1.Range("HASZIP").value = "Yes" Then
                Sheet1.Range("sheet1.ZipCode2").value = "XXXXXX"
                Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
                Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
            Else
            Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
            End If
        Else
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.ZipCode2").value = Sheet1.Range("sheet1.ZipCode").value
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
        End If
Else

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = False

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.ClearContents

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = False

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = False

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.ClearContents

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = False

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.ClearContents

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = False

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents

        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.StateCode2").value = "(Select)"

        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
 
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.CountryCode2").value = "(Select)"

        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("HASZIP2").MergeArea.ClearContents

        Sheet1.Range("HASZIP2").MergeArea.Locked = False

        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"

    End If
 
End Sub
 
 Sub LockUnlock_Alternate_Address()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
     
     
     If Sheet1.Range("Alternate_Address").value = "Yes" Then
     
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = False
            Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HCCFFCC"
            
             
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = False
            Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = False
            Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = False
            Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = False
            Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HCCFFCC"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"
            Sheet1.Range("sheet1.StateCode2").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HCCFFCC"
            Sheet1.Range("sheet1.CountryCode2").MergeArea.value = "(Select)"
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HCCFFCC"
            
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("HASZIP2").MergeArea.Locked = False
'            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HCCFFCC"
        
    Else
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.ClearContents
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = True
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.ResidenceNo2").value = ""
        
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.ClearContents
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = True
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.ResidenceName2").value = ""
        
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.ClearContents
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = True
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.RoadOrStreet2").value = ""
        
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.ClearContents
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = True
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.LocalityOrArea2").value = ""
        
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.ClearContents
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = True
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").value = ""
        
        Sheet1.Range("sheet1.StateCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.StateCode2").value = ""
        
        Sheet1.Range("sheet1.CountryCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.CountryCode2").value = ""
        
        Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents
        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.PinCode2").value = ""
        
        
'        Sheet1.Range("HASZIP2").MergeArea.ClearContents
'        Sheet1.Range("HASZIP2").MergeArea.Locked = True
'        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("HASZIP2").value = ""
'
'        Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents
'        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
'        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("sheet1.ZipCode2").value = ""
        
        
     End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockUnlock80E_80EE()
On Error Resume Next

    If UCase(Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1)) = "H" Then
    
        sheetVisibleForYes_80E
    Else
        sheetVisibleForNo_80E

    End If
End Sub
Sub sheetVisibleForYes_80E()

ThisWorkbook.Unprotect Password:=getmsgstate
      
      Sheet60.Unprotect Password:=getmsgstate
      Sheet60.Visible = xlSheetHidden
      Sheet60.Protect Password:=getmsgstate
      
ThisWorkbook.Protect Password:=getmsgstate
End Sub

Sub sheetVisibleForNo_80E()

ThisWorkbook.Unprotect Password:=getmsgstate
    
    Sheet60.Unprotect Password:=getmsgstate
    Sheet60.Visible = xlSheetVisible
    Sheet60.Protect Password:=getmsgstate
    
    ThisWorkbook.Protect Password:=getmsgstate
    
    End Sub
Sub ID_Nature_Employment()

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    'Unprotect only required sheet
    Sheet24.Unprotect Password:=getmsgstate

    With Sheet24.Range("Others.NOI") ' ? ensure this named range exists
        .Validation.Delete

        If Sheet1.Range("sheet1.ResidentialStatus1").value = "NRI - Non Resident" Then
            .Validation.add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="=Nature_TP"
        Else
            .Validation.add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="=Nature_TP1"
        End If
    End With

ExitHandler:
    Sheet24.Protect Password:=getmsgstate
    Application.EnableEvents = True

End Sub
''''''''''''''''''
Sub test()
Application.EnableEvents = True
End Sub


