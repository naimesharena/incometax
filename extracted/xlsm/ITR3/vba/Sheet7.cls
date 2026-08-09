Attribute VB_Name = "Sheet7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim CreditDate As String

Dim sPassword As String

Private Sub Worksheet_Activate()
On Error Resume Next

'For Rebate 87A by sai on 06/01/2025===================================================
If CalculateTaxCheckFlag = False Then

  Application.EnableEvents = False
            Sheet7.Unprotect Password:=getmsgstate
            Sheet7.Range("Sheet9.RebateUs88E").Locked = True
            
    'IPIP-71564 uncommented by sai on 19/07/2025 for rebate issue
            Sheet7.Range("Sheet9.RebateUs88E").formula = "=Rebate87Aformula_new"
            
    'commented by sai on 19/07/2025 for rebate issue
'           Sheet7.Range("Sheet9.RebateUs88E").value = 0
            Sheet7.Range("Sheet9.RebateUs88E").Interior.ColorIndex = 2
            Sheet7.Protect Password:=getmsgstate
End If

'sadineni_10/07/2025
'**************************************************************************************
Sheet7.Unprotect Password:=getmsgstate
Sheet7.Range("Sheet9.Surcharge_ii").value = Sheet33.Range("Surcharge_ii").value
'**************************************************************************************
  Application.EnableEvents = True
  
'======================================================================================

 sPassword = getmsgstate
    Sheet7.Unprotect Password:=sPassword
    
    ActiveSheet.Shapes("Button 175").Left = Range("PartB.Button1").Left + 2
    ActiveSheet.Shapes("Button 175").Top = Range("PartB.Button1").Top + 2


    'If Range("Sheet9.BalTaxPayable1").value > 0 Then
    '    ActiveSheet.Shapes("Button 201").Visible = True
    'Else
        ActiveSheet.Shapes("Button 201").Visible = False
    'End If

    If Sheet7.Range("Sheet9.BalTaxPayable1").value > 0 Then
         Sheet7.Shapes("Button 1930").Visible = True
       Else
           Sheet7.Shapes("Button 1930").Visible = False
    End If

    'ActiveSheet.Shapes("Button 196").Left = Range("PartB.Button2").Left + 2
    'ActiveSheet.Shapes("Button 196").Top = Range("PartB.Button2").Top - 1
    
    
     sPassword = getmsgstate
    Sheet7.Protect Password:=sPassword

End Sub


Sub LockUnlockBANKDTL()
On Error Resume Next
    If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then
        UnlockBANK

    ElseIf Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 1) = "N" And Mid(Sheet7.Range("Sheet9.IBANSELECT").value, 1, 1) = "N" Then

     LockBANK
     fmsgboxStatus "Currently refund will be credited to a bank account located in India and not to a bank account located in a foreign country"

     ElseIf Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 1) = "N" And Mid(Sheet7.Range("Sheet9.IBANSELECT").value, 1, 1) = "Y" Then

     UnlockBANK
    Else
     UnlockBANK
    End If
End Sub

Sub LockBANK()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    sPassword = getmsgstate
    Sheet7.Unprotect Password:=sPassword

    Dim rngname_BANK As Variant

   ' rngname_BANK = "Sheet9.MICRCode||Sheet9.BankName||Sheet9.BankAccountNumber" '||Sheet9.CashDeposited"
   ' LOCKRANGEBANK (rngname_BANK)
     
     'Malli
    'rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||tempxml" '||Sheet9.OthCashDeposited"
    rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount" '||Sheet9.OthCashDeposited"
    LOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "Sheet9.IBAN||Sheet9.IBANBANK||Sheet9.IBANCOUNTRY||Sheet9.IBANACCNO"
    UNLOCKRANGEBANK (rngname_BANK)

       
    
'Newly added by sai on 13/09/2024 =======================================================
Sheet7.Unprotect Password:=sPassword

Sheet7.Activate

Sheet7.Buttons("Button 1939").Visible = False

Sheet7.Range("k116").Font.ThemeColor = xlThemeColorDark1 'And Sheet7.Range("k116").Font.TintAndShade = -0.249977111117893

Sheet7.Range("k116").Font.TintAndShade = -0.249977111117893

'=========================================================================================


    Sheet7.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub LockBANK1()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    sPassword = getmsgstate
    Sheet7.Unprotect Password:=sPassword

    Dim rngname_BANK As Variant

   ' rngname_BANK = "Sheet9.MICRCode||Sheet9.BankName||Sheet9.BankAccountNumber" '||Sheet9.CashDeposited"
   ' LOCKRANGEBANK (rngname_BANK)

    'Malli
    'rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||tempxml" '||Sheet9.OthCashDeposited"
    rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount" '||Sheet9.OthCashDeposited"
    UNLOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "Sheet9.IBAN||Sheet9.IBANBANK||Sheet9.IBANCOUNTRY||Sheet9.IBANACCNO"
    UNLOCKRANGEBANK (rngname_BANK)

    Sheet7.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnlockBANK()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    sPassword = getmsgstate
    Sheet7.Unprotect Password:=sPassword

    Dim rngname_BANK As Variant

    'rngname_BANK = "Sheet9.MICRCode||Sheet9.BankName||Sheet9.BankAccountNumber" '||Sheet9.CashDeposited"
    'UNLOCKRANGEBANK (rngname_BANK)
     
     'Malli
    'rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||tempxml" '||Sheet9.OthCashDeposited"
    rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount" '||Sheet9.OthCashDeposited"
    UNLOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "Sheet9.IBAN||Sheet9.IBANBANK||Sheet9.IBANCOUNTRY||Sheet9.IBANACCNO" '
    LOCKRANGEBANK (rngname_BANK)
'=================================================
'Newly added by sai on 13/09/2024

Sheet7.Unprotect Password:=sPassword

Sheet7.Activate

Sheet7.Buttons("Button 1939").Visible = True

Sheet7.Range("k116").Font.Color = -16776961
'
''==================================================

  

    Sheet7.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKRANGEBANK(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet7.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                    'chCell.value = ""
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    'chCell.value = ""

                End If
            Next
        Next
End Sub

Sub LOCKRANGEBANK(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet7.Range(rangenamestring(i))
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

Private Sub Worksheet_Calculate()

End Sub

'Worksheet Change

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False
'Sheet7.Unprotect Password:=sPassword
Dim count, i As Variant
Dim allowifscnnn, tempvar As Variant


'For Rebate 87A Changes by sai on 06/01/2025
'========================================================================
If Not Application.Intersect(Sheet7.Range("Sheet9.RebateUs88E"), Sheet7.Range(Target.Address)) Is Nothing Then
  
    CalculateTaxCheckFlag = True
    Call calculateTax(False)
    CalculateTaxCheckFlag = True
  
End If

'======================================================================

'-------------Bank Code------------------------------------
Dim BankNameFromIFSC As Variant
Dim Address As Variant
Dim rangecell As Range
Dim sourceAddress As String

'Newly added by Bindu

If Not Application.Intersect(Range("Sheet9.OthBankName"), Range(Target.Address)) Is Nothing Then
    
    Target.value = UCase(Target.value)

End If

If Not Application.Intersect(Range("Sheet9.OthMICRCode"), Range(Target.Address)) Is Nothing Then
    For Each rangecell In Sheet7.Range("Sheet9.OthMICRCode")
        rangecell.value = UCase(rangecell.value)
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(rangecell.value)) Then

                allowifscnnn = False
                tempvar = Sheet1.Range("sheet1.ResidentialStatus1").value
                If Mid(tempvar, 1, 3) = "NRI" Then
                   allowifscnnn = True
                End If

                    fmsgboxStatus "Please enter valid IFS Code."
                    Sheet7.Range(Target.Address).value = ""
                    Sheet7.Range(Target.Address).Select
                    sourceAddress = rangecell.Address
                    Address = Replace(sourceAddress, "G", "H")
                    Sheet7.Range(Address).value = ""
                'End If
                
             ElseIf Not ValidateIFSCCode(Target.value) Then
                'fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                fmsgboxStatus "IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                Range(Target.Address) = ""
                Range(Target.Address).Select

            Else
                sourceAddress = rangecell.Address
                Address = Replace(sourceAddress, "G", "H")
                   'NEWLY ADDED BY BINDU ON 27/11/24 AS PER REQUEST FROM BA TEAM
                
                If Mid(rangecell, 1, 4) = "GIFT" Then
                    BankNameFromIFSC = GetBankName(rangecell)
                Else
                    BankNameFromIFSC = GetBankName(Mid(rangecell, 1, 4))
                End If
                
                If Mid(rangecell, 1, 4) = "GIFT" And BankNameFromIFSC <> "" Then
                   Sheet7.Range(Address).value = BankNameFromIFSC
                ElseIf Mid(rangecell, 1, 4) <> "GIFT" Then
                   Sheet7.Range(Address).value = BankNameFromIFSC
                End If
              '--ends
            
        'SIT-117310 by sadineni on 11/04/2026
'****************************************************************************************
                If Mid(rangecell, 1, 4) <> "GIFT" And rangecell <> "" Then
                    If (Sheet7.Range(Address).value = "") Then
                        fmsgboxStatus """Please ensure the IFSC entered in Schedule Part B TTI in point 13 is valid"""
                        Address = Replace(Address, "H", "G")
                        Sheet7.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
'*******************************************************************************************
            
            End If
        End If
    Next
    GoTo endline
End If


If Target.name.name = "Sheet9.IBANSELECT" Then
    LockUnlockBANKDTL
    GoTo endline
End If

If Target.name.name = "AOIFlag_1" Then

CalculateTaxCheckFlag = False

If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NRI" And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) <> "NOR" Then
    If Mid(Sheet7.Range("AOIFlag_1").value, 1, 1) <> "N" Then
        UNLOCKFA
    Else
        LOCKFA
    End If
Else
    LOCKFA
End If
End If

If (Target.name.name) = "Sheet9.Section89" Then
        CalculateTaxCheckFlag = False

            If Range("Sheet9.Section89").value = "" Then
                Range("Sheet9.Section89").value = 0
           End If
'            If Range("Sheet9.Section89").value > Range("Sheet9.Section89New").value Then
'                Range("Sheet9.Section89").value = 0
'                MsgBox "Section 89 cannot exceed Rs." & Range("Sheet9.Section89New").value
'            End If

    End If


If Target.name.name = "Sheet9.Section89" Then
 CalculateTaxCheckFlag = False
 End If


endline:
Application.EnableEvents = True
End Sub




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
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

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

'Sub PPPPP()
'Application.EnableEvents = True
'End Sub

