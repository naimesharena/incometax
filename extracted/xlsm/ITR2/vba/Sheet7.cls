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
Dim ColCount4_BA As Variant

Private Sub Worksheet_Activate()
On Error Resume Next

'=======For Rebate 87A by sai on 14/01/2025
If CalculateTaxCheckFlag = False Then

Application.EnableEvents = False
            Sheet7.Unprotect Password:=getmsgstate
            Sheet7.Range("Sheet9.RebateUs88E").Locked = True
            Sheet7.Range("Sheet9.RebateUs88E").formula = "=Rebate87Aformula_new" 'uncommented by Sadineni_22_07_2025
            'Sheet7.Range("Sheet9.RebateUs88E").value = 0
            Sheet7.Range("Sheet9.RebateUs88E").Interior.ColorIndex = 2
            Sheet7.Protect Password:=getmsgstate
End If

Application.EnableEvents = True

'==================================================



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
    
    

'Newly added by Bindu
Dim count54Acc, count54IFSC
Dim countIFSC_PartB, TotalDiffRow
Dim rowcount, rowcount54F
Dim Acc54FColumn, IFSC54FColumn
Dim cChell As Variant
Dim ccount
Dim IFSCColumn, AccColumn, TypeColumn

Dim mIntCells
Dim mIntCtr
Dim rangecells As Range

Dim sPassword

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

    'rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||tempxml" '||Sheet9.OthCashDeposited"
    rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount"

    
    LOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "Sheet9.IBAN||Sheet9.IBANBANK||Sheet9.IBANCOUNTRY||Sheet9.IBANACCNO"
    UNLOCKRANGEBANK (rngname_BANK)
    
    'sudha
    Sheet7.Range("Sheet9.IBANCOUNTRY").value = "(Select)"
    'sudha
    
     'Malli----
    Sheet7.Unprotect Password:=sPassword

Sheet7.Activate
Sheet7.Buttons("Button 1939").Visible = False
Sheet7.Range("k116").Font.ThemeColor = xlThemeColorDark1 'And Sheet7.Range("k116").Font.TintAndShade = -0.249977111117893
Sheet7.Range("k116").Font.TintAndShade = -0.249977111117893
 
    '---------

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

   ' rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||tempxml" '||Sheet9.OthCashDeposited"
    rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount"

    
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

    rngname_BANK = "Sheet9.OthMICRCode||Sheet9.OthBankName||Sheet9.OthBankAccountNumber||Sheet9.TypeOfAccount" '||Sheet9.OthCashDeposited"
   
    UNLOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "Sheet9.IBAN||Sheet9.IBANBANK||Sheet9.IBANCOUNTRY||Sheet9.IBANACCNO" '
    LOCKRANGEBANK (rngname_BANK)

'Malli----------
Sheet7.Unprotect Password:=sPassword

Sheet7.Activate
Sheet7.Buttons("Button 1939").Visible = True
Sheet7.Range("k116").Font.Color = -16776961
'-----------------

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
                    'chcell.value = "(Select)"
                    'chCell.value = ""
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    'chcell.value = "(Select)"
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
Sheet7.Unprotect Password:=getmsgstate

Dim count, i As Variant
Dim allowifscnnn, tempvar As Variant

'Dim graosstag As Boolean
'Newly added by Bindu

'added by sai for Rebate 87A enhancement on 014/01/2025
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
                fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                Range(Target.Address) = ""
                Range(Target.Address).Select

            Else
                sourceAddress = rangecell.Address
                Address = Replace(sourceAddress, "G", "H")
                'Malli--------------AY_2025_26
                'AY_2024_25 OLD
'                BankNameFromIFSC = GetBankName(Mid(rangecell, 1, 4))
'                Sheet7.Range(Address).value = BankNameFromIFSC
                
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
                
                If Mid(rangecell, 1, 4) <> "GIFT" And rangecell <> "" Then
                    If (Sheet7.Range(Address).value = "") Then
                     
                     fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)" & Chr(13)
                     Address = Replace(Address, "H", "G")
                     Sheet7.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                '--------------------------
                '-----------------------------------
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
           
                   'Ankita_08/09/2025_SIT-89279
            If Range("Sheet9.Section89").value > Range("Sheet9.Section89New").value Then
                Range("Sheet9.Section89").value = 0
                MsgBox "Section 89 cannot exceed Rs." & Range("Sheet9.Section89New").value
            End If

    End If


If Target.name.name = "Sheet9.Section89" Then
 CalculateTaxCheckFlag = False
 End If
'sudha



'If Not Application.Intersect(Sheet7.Range("Sheet9.BalTaxPayable"), Sheet7.Range(Target.Address)) Is Nothing Then


'Ankita_08/09/2025_SIT-89279 =========================================
'If Application.Intersect(Sheet7.Range("Sheet9.BalTaxPayable"), Sheet7.Range(Target.name)) Is Nothing Then
'If Sheet7.Range("Sheet9.BalTaxPayable1").value > 0 Then
'MsgBox "Please ensure that the taxes are paid before the submission of the return, else return shall be treated as defective.", vbInformation, "Info"
'End If
'End If


'sudha

endline:
Application.EnableEvents = True
Sheet7.Protect Password:=getmsgstate
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

Sub TESST()
Application.EnableEvents = True
End Sub
Sub setTableInfo_AccountNumber()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet9.OthBankAccountNumber").count
    Set rangecells = Range("Sheet9.OthBankAccountNumber").Cells
    For mIntCtr = 1 To mIntCells
            If Not Trim(rangecells.item(mIntCtr).value) = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount4_BA = ccount
End Sub
