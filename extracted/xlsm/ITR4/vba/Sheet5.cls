Attribute VB_Name = "Sheet5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim sPassword As Variant


Public IBANSELECT As Variant
Public IBAN As Variant
Public IBANNameOfBank As Variant
Public IBANCOUNTRY As Variant
Public IBANACCNO As Variant
Public IDcurrCnt1, IDcurrCnt2, IDcurrCnt3, IDcurrCnt4, IDcurrCnt5 As Variant
Public currCnt1, currCnt2, currCnt3, currCnt4, currCnt5, currCnt6, currCnt7, currCnt8, currCnt9, currCnt10, currCnt11, currCnt12, currCnt13, currCnt14, currCnt15 As Variant
Sub a()
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False

 Sheet5.Unprotect Password:=getmsgstate

    ActiveSheet.Shapes("Button 99").Left = Range("TPV.Button").Left + 5
    ActiveSheet.Shapes("Button 99").Top = Range("TPV.Button").Top + 5

    ActiveSheet.Shapes("Button 159").Left = Range("others_button").Left + 5
    ActiveSheet.Shapes("Button 159").Top = Range("others_button").Top + 5

    'If Range("IncD.BalTaxPayable").value > 0 Then
    '    ActiveSheet.Shapes("Button 111").Visible = True
    'Else
        ActiveSheet.Shapes("Button 111").Visible = False
    'End If


    ActiveSheet.Shapes("EPAYButton").Left = Range("Button_epaytax").Left + 1
    ActiveSheet.Shapes("EPAYButton").Top = Range("Button_epaytax").Top + 1


    If Sheet5.Range("IncD.BalTaxPayable").value > 0 Then
        Sheet5.Shapes("EPAYButton").Visible = True
          Else
         Sheet5.Shapes("EPAYButton").Visible = False
    End If
    

'Ayush_Legacy6
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>29/05/2025

If Sheet8.Range("bacValue").value = 1 Then

Dim Sheet5_Nature_cellscnt, i As Variant

Sheet5_Nature_cellscnt = Sheet5.Range("Sheet5.Nature").Rows.count

        For i = 1 To Sheet5_Nature_cellscnt

                If Sheet5.Range("Sheet5.Nature").item(i, 1).value = "Sec 10(17)-Allowance MP/MLA/MLC" Then

                     Sheet5.Range("Sheet5.Nature").item(i, 1).value = "(Select)"

                     Sheet5.Range("Sheet5.Nature").item(i, 1).Offset(0, 1).value = ""

                     Sheet5.Range("Sheet5.Nature").item(i, 1).Offset(0, 2).value = ""

                 End If

        Next

End If

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 





endline:
Application.EnableEvents = True
'Sheet5.Protect Password:=sPassword
 Sheet5.Protect Password:=getmsgstate

End Sub


'Sub LockUnlockBANKDTL()
'On Error Resume Next
'    If Mid(Sheet5.Range("Sheet9.IBANSELECT").value, 1, 1) <> "N" Then
'        UnlockBANK
'    ElseIf Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" And Mid(Sheet5.Range("Sheet9.IBANSELECT").value, 1, 1) = "N" Then
'
'     LockBANK
'     MsgBox "Currently Refund will be credited to a bank account located in India and not to a bank account located in a foreign country", vbExclamation
'
'    Else
'     LockBANK
'    End If
'End Sub

Sub LockBANK()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    sPassword = getmsgstate
    Sheet5.Unprotect Password:=sPassword

    Dim rngname_BANK As Variant

    rngname_BANK = "Refund.IFSC||Refund.BankName||Refund.AcntNo" '||Refund.CashDeposited"
    LOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo" '||SchBA.DepositedCash"
    LOCKRANGEBANK (rngname_BANK)


    Sheet5.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UnlockBANK()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    sPassword = getmsgstate
    Sheet5.Unprotect Password:=sPassword

    Dim rngname_BANK As Variant

    rngname_BANK = "Refund.IFSC||Refund.BankName||Refund.AcntNo" '||Refund.CashDeposited"
    UNLOCKRANGEBANK (rngname_BANK)

    rngname_BANK = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo" '||SchBA.DepositedCash"
    UNLOCKRANGEBANK (rngname_BANK)

    Sheet5.Protect Password:=sPassword

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub UNLOCKRANGEBANK(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet5.Range(rangenamestring(i))
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
            For Each chcell In Sheet5.Range(rangenamestring(i))
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


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim TargtVal, targetadd As Variant
Dim IffCode As Range
Dim IfsCode As Range


Application.EnableEvents = False
Application.ScreenUpdating = False

'CalculateTaxCheckFlag = False


'Added by Riyaz on 05/03/2026 for AY 2026-27
If Not Application.Intersect(Sheet5.Range("Sheet5.Nature"), Sheet5.Range(Target.Address)) Is Nothing Then
    
    If Sheet8.Range("bacValue").value = 1 Then
                Sheet5.Unprotect Password:=getmsgstate
                Application.EnableEvents = False
                
                Dim Sub_Cat As Variant
                Sub_Cat = Target.Address
                Sub_Cat = Replace(Sub_Cat, "$", "")
                Sub_Cat = Replace(Sub_Cat, "G", "H")
                
            If Target.value = "Agricultural  & related incomes" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Agri_dropdown"
            ElseIf Target.value = "Compensation/other sums received by government or other approved entities" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Comp_dropdown"
            ElseIf Target.value = "Income from specified Investments" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Inco_dropdown"
            ElseIf Target.value = "Specified sums received by armed forces personnel" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Spec__dropdown"
            ElseIf Target.value = "Sums received by Senior Citizens/Minors " Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Seminew_dropdown"
            ElseIf Target.value = "Sums received by specified Category of Taxpayers" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Taxp_dropdown"
            ElseIf Target.value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Yojana_dropdown"
            ElseIf Target.value = "Other Incomes" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
                Sheet5.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
            ElseIf Target.value = "(Select)" Or Target.value = "" Then
                Sheet5.Range(Sub_Cat).Validation.Delete
                Sheet5.Range(Sub_Cat).value = ""
            End If
        Else
            Sheet5.Unprotect Password:=getmsgstate
            Application.EnableEvents = False
            
            Dim Sub_Cat1 As Variant
            Sub_Cat1 = Target.Address
            Sub_Cat1 = Replace(Sub_Cat1, "$", "")
            Sub_Cat1 = Replace(Sub_Cat1, "G", "H")
            
        If Target.value = "Agricultural  & related incomes" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Agri_dropdown"
        ElseIf Target.value = "Compensation/other sums received by government or other approved entities" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Comp_dropdown"
        ElseIf Target.value = "Income from specified Investments" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Inco_dropdown"
        ElseIf Target.value = "Specified sums received by armed forces personnel" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Spec__dropdown"
        ElseIf Target.value = "Sums received by Senior Citizens/Minors " Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Semiold_dropdown"
        ElseIf Target.value = "Sums received by specified Category of Taxpayers" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Taxp_dropdown"
        ElseIf Target.value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Yojana_dropdown"
        ElseIf Target.value = "Other Incomes" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
            Sheet5.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
        ElseIf Target.value = "(Select)" Or Target.value = "" Then
            Sheet5.Range(Sub_Cat1).Validation.Delete
            Sheet5.Range(Sub_Cat1).value = ""
        End If
           
            Sheet5.Protect Password:=getmsgstate
            Application.EnableEvents = True
    End If
End If

'Ankita_15/06/2026=======
If Not Application.Intersect(Sheet5.Range("Sheet5.Description"), Sheet5.Range(Target.Address)) Is Nothing Then
    Dim subcategory As String
            
   ' If Target.value = "Income exempt as per CBDT Circular/Notification" Or Target.value = "Receipts not in the nature of income" Then
    If Target.value = "Income exempt as per CBDT Circular" Or Target.value = "Income exempt as per CBDT Notification" Or Target.value = "Receipts not in the nature of income" Then
        Sheet5.Unprotect Password:=getmsgstate
        Sheet5.Range(Replace(Target.Address, "H", "I")).ClearContents
        Sheet5.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HCCFFCC"
        Sheet5.Range(Replace(Target.Address, "H", "I")).Locked = False
        Sheet5.Range(Replace(Target.Address, "H", "I")).value = ""
        Sheet5.Protect Password:=getmsgstate
    Else
        Sheet5.Unprotect Password:=getmsgstate
        Sheet5.Range(Replace(Target.Address, "H", "I")).ClearContents
        Sheet5.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HD8D8D8"
        Sheet5.Range(Replace(Target.Address, "H", "I")).Locked = True
        Sheet5.Range(Replace(Target.Address, "H", "I")).value = ""
        Sheet5.Protect Password:=getmsgstate
    End If
    
        Sheet5.Unprotect Password:=getmsgstate
'       Sheet5.Range(Replace(Target.Address, "H", "G")).value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.VLookup(Target.value, Sheet31.Range("Sub_Category_lookup"), 2, 0), "")
        Sheet5.Protect Password:=getmsgstate
End If

'Ankita_15/06/2026===

If Not Application.Intersect(Sheet5.Range("Sheet5.Description"), Sheet5.Range(Target.Address)) Is Nothing Then
    If Not Target.value = "(Select)" Or Target.value = "" Then
    If Target.value = "Income exempt as per CBDT Circular" Or Target.value = "Income exempt as per CBDT Notification" Or Target.value = "Receipts not in the nature of income" Then
         Else
        If WorksheetFunction.CountIf(Sheet5.Range("Sheet5.Description"), Target.value) > 1 Then
        Application.EnableEvents = False
            MsgBox "You cannot select twice"
            Target.value = ""
        Application.EnableEvents = True
        End If
      End If
    End If
End If

'Newly added by Bindu on 31st Jan 2025 as per V3 DE sheet
If Not Application.Intersect(Sheet5.Range("IncD.Cost_LTCG"), Sheet5.Range(Target.Address)) Is Nothing Then
  If Sheet5.Range("IncD.Sale_LTCG") > 0 Then
            If Sheet5.Range("IncD.Cost_LTCG") > 0 And (Sheet5.Range("IncD.Sale_LTCG") - Sheet5.Range("IncD.Cost_LTCG")) > 125000 Then
                fmsgbox """In ITR 4, the maximum gains as per Section 112A can be INR 1,25,000/-. Please file ITR 3/5 if gains u/s 112A is more than INR 1,25,000/-"""
                Sheet5.Range("IncD.Cost_LTCG").ClearContents
            End If
    End If
End If

If Not Application.Intersect(Sheet5.Range("IncD.Sale_LTCG"), Sheet5.Range(Target.Address)) Is Nothing Then
   If Sheet5.Range("IncD.Cost_LTCG") > 0 Then
        If Sheet5.Range("IncD.Sale_LTCG") > 0 And (Sheet5.Range("IncD.Sale_LTCG") - Sheet5.Range("IncD.Cost_LTCG")) > 125000 Then
            fmsgbox """In ITR 4, the maximum gains as per Section 112A can be INR 1,25,000/-. Please file ITR 3/5 if gains u/s 112A is more than INR 1,25,000/-"""
            Sheet5.Range("IncD.Sale_LTCG").ClearContents
        End If
    End If
End If


    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")


''TAX DETAILS-C7 AY_2024_25 Added 'malli
If Target.Column = 9 Then
        If Not Target.value = "(Select)" Then
        'If 1 < Application.CountIf(Range("SchBA.Type_of_Account"), Target.value) Then
             'MsgBox "* B22_Section under which TYPE of account in Same drop downs cannot be selected more than once in  Taxes paid and verification"
         '    fmsgbox "* TYPE of account in Same drop downs cannot be selected more than once in  Taxes paid and verification"
          '   Range(Target.Address).value = "(Select)"
        'End If
        End If
End If


'If Not checkfieldSuperSpecialcharacter(Target.value) Then
'         MsgBox ("Please do not enter Special characters < , >" & ",&" & Chr(13))
'        Target.value = ""
'End If

'Chanage.02.03.2023.102.IDS.98/99

'If Not Application.Intersect(Sheet5.Range("Sheet5.Nature"), Sheet5.Range(Target.Address)) Is Nothing Then
'    If Sheet8.Range("bacValue").value = 1 Then
'    If Sheet5.Range(Target.Address).value = "Sec 10(17)-Allowance MP/MLA/MLC" Then
'    Sheet5.Range(Target.Address).value = ""
'            Sheet5.Range(Target.Address).Select
'
'      Sheet5.Range(Replace(Target.Address)).value = "(Select)"
'
'    End If
'    End If
'
'End If

'Added by Riyaz on 05/03/2026
'If Not Application.Intersect(Sheet5.Range("Sheet5.Description"), Sheet5.Range(Target.Address)) Is Nothing Then
'        If 1 < Application.CountIf(Sheet5.Range("Sheet5.Description"), Target.value) Then
'             fmsgbox "* Same drop downs cannot be selected twice"
'             Range(Target.Address).value = "(Select)"
'        End If
'End If

'Changed by Riyaz on 05/04/2026
If Not Application.Intersect(Sheet5.Range("Sheet5.Description"), Sheet5.Range(Target.Address)) Is Nothing Then
    If Target.value = "10(32)-Minor child’s income—small exemption" Then
        DPMAddComment45 Target
    ElseIf Target.value = "10(26)-Any income as referred to in section 10(26)" Then
        DPMAddComment46 Target
    ElseIf Target.value = "10(26AAA)-Any income as referred to in section 10(26AAA)" Then
        DPMAddComment47 Target
    ElseIf Target.value = "10(12A)-Any payment from the National Pension System Trust to an assessee" Then
        DPMAddComment48 Target
    ElseIf Target.value = "10(12AA)-any payment from the National Pension System Trust " Then
        DPMAddComment49 Target
    ElseIf Target.value = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee" Then
        DPMAddComment50 Target
    ElseIf Target.value = "10(12BA)-partial withdrawal made from the National Pension System" Then
        DPMAddComment51 Target
    Else
        DPMADeleteComment45 Target
    End If
End If
'End Change IDS98/99


If Not Application.Intersect(Sheet5.Range("Sheet5.Amount"), Sheet5.Range(Target.Address)) Is Nothing Then

   ' If Mid(Sheet5.Range(Replace(Target.Address, "I", "G")).value, 1, 4) = "Agri" Then
   If Mid(Sheet5.Range(Replace(Target.Address, "I", "H")).value, 1, 5) = "10(1)" Then
        If Sheet5.Range(Target.Address).value > 5000 Then
            fmsgbox "* If sec 10(1) agricultural income is more than Rs 5000/- then use ITR 3 or 5"
            Sheet5.Range(Target.Address).value = ""
            Sheet5.Range(Target.Address).Select

            'Sheet5.Range(Replace(Target.Address, "I", "G")).value = "(Select)"
        End If
    End If

    If Mid(Sheet5.Range(Replace(Target.Address, "I", "G")).value, 1, 10) = "Sec 10(34)" Then
        If Sheet5.Range(Target.Address).value > 1000000 Then
            fmsgbox "*Sec 10(34) can not be more than Rs. 10,00,000/- in exempt income . Please fill ITR-3/ITR-5 as applicable"
            Sheet5.Range(Target.Address).value = ""
            Sheet5.Range(Target.Address).Select
           ' Sheet5.Range(Replace(Target.Address, "I", "G")).value = "(Select)"
        End If
    End If
'Change-28.11.2022.103.06
    TAX_LTCG_AMOUNT Target

End If

                Dim nOfCount  As Long
'Change-28.11.2022.103.05A
'               nOfCount = CountDropdown(Trim(Target.value), Sheet5.Range("Sheet5.Nature").Cells, Sheet5.Range("Sheet5.Nature").count)
'                nOfCount = CountDropdown23_24(Trim(Target.value), Sheet5.Range("Sheet5.Description").Cells, Sheet5.Range("Sheet5.Description").count, Target)
                'If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Then
'                If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Then
'                    fmsgbox "* Do not enter same Nature of income twice" ',' vbOKOnly, "Error(s)!"
'                    Target.value = "(Select)"
'                End If

'---End Change




If Target.Validation.Type = 3 Then


    If Not Application.Intersect(Sheet5.Range("Sheet5.Nature"), Sheet5.Range(Target.Address)) Is Nothing Then


'        If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
'                MsgBox ("Nature of income should not contain Special characters < , >" & ",&" & Chr(13))
'               Range(Target.Address).value = ""
'        End If


'Commented by Riyaz on 05/03/2026
'        If Sheet5.Range(Target.Address).value = "Any Other" Then
'            sPassword = getmsgstate
'            Sheet5.Unprotect Password:=getmsgstate
'            Sheet5.Range(Replace(Target.Address, "G", "H")).value = ""
'            Sheet5.Range(Replace(Target.Address, "G", "H")).Interior.Color = (&HCCFFCC)  ', "AU"
'            Sheet5.Range(Replace(Target.Address, "G", "H")).Locked = False
'
'            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
'
'            Sheet5.Protect Password:=sPassword
'        Else
'            sPassword = getmsgstate
'            Sheet5.Unprotect Password:=sPassword
'            Sheet5.Range(Replace(Target.Address, "G", "H")).value = "Not Applicable"
'            Sheet5.Range(Replace(Target.Address, "G", "H")).Interior.Color = (&HD8D8D8)
'            Sheet5.Range(Replace(Target.Address, "G", "H")).Locked = True
'            Sheet5.Protect Password:=sPassword
'       End If




    End If


'If Target.name.name = "Ver.Capacity" Then
'
'
'    End If




'
'    If Target.name.name = "Sheet9.IBANSELECT" Then
'        LockUnlockBANKDTL
'        GoTo endline
'    End If

End If

'    If Target.name.name = "AIR.TaxExmpIntInc" Then
'        If Sheet5.Range("AIR.TaxExmpIntInc").value > 5000 Then
'                MsgBox "If Agricultural income is greater than 5000  use ITR-3/5"
'                Sheet5.Range("AIR.TaxExmpIntInc").value = ""
'        End If
'    End If




If Target.Validation.Type = 6 Then
    Target.value = UCase(Target.value)

        Set IffCode = Range("Refund.IFSC")
        If Not Application.Intersect(IffCode, Range(Target.Address)) Is Nothing Then
                Call CheckIFSC1(TargtVal, targetadd)
               GoTo endline
        End If

       



'--------17/12/2024----
'17/12/2024
 Set IfsCode = Range("SchBA.IFSC")
'        If Not Application.Intersect(IfsCode, Range(target.Address)) Is Nothing Then
'
'
'            Call CheckIFSC1(TargtVal, targetadd)
'            GoTo endline
'        End If

    Dim sourceAddress, Address, Address1 As Variant
    Dim rangecell As Range
    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet5.Range("SchBA.IFSC")
            rangecell.value = UCase(rangecell.value)
     
     If rangecell.value <> "" Then
           
           If Not CheckIFSC(UCase(TargtVal)) Then
                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet5.Range(targetadd) = ""
                    Sheet5.Range(targetadd).Offset(0, 1) = ""
                    GoTo endline
           
            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet5.Range(targetadd) = ""
                    Sheet5.Range(targetadd).Offset(0, 1) = ""
                    GoTo endline
     
                  'Call CheckIFSC1(TargtVal, targetadd)
                  'GoTo endline
            Else:
                       sourceAddress = rangecell.Address
                       Address = Replace(sourceAddress, "F", "G")
                
                     If Mid(rangecell, 1, 4) = "GIFT" Then
                             BankNameFromIFSC = GetBankName(rangecell)
                     Else
                             BankNameFromIFSC = GetBankName(Mid(rangecell, 1, 4))
                     End If
                
                
                If Mid(rangecell, 1, 4) = "GIFT" And BankNameFromIFSC <> "" Then
                   Sheet5.Range(Address).value = BankNameFromIFSC
                ElseIf Mid(rangecell, 1, 4) <> "GIFT" Then
                   Sheet5.Range(Address).value = BankNameFromIFSC
                End If
                
                If Mid(rangecell, 1, 4) <> "GIFT" Then
                    If (Sheet5.Range(Address).value = "") Then
                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet5.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
         End If
    Next
    GoTo endline
End If


'-----------------
'Commented by Riyaz on 05/03/2026
'        If Not Intersect(Range("Sheet5.Description"), Range(Target.Address)) Is Nothing Then
'            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).value) Then
'                fmsgbox ("*Description should not contain Special characters < , > " & ",& " & "in in Sheet Taxes paid and verification" & Chr(13))
'               Range(Target.Address).value = ""
'               Range(Target.Address).Select
'            End If
'            End If


        If Target.name.name = "Ver.PAN" Then
                Target.value = UCase(Target.value)
            If Range("Ver.PAN").value <> "" Then
                If Not CheckPAN(UCase(Range("Ver.PAN").value)) Then
                    fmsgbox "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
                    Range("Ver.PAN").value = ""
                    Range("Ver.PAN").Select
                    GoTo endline
                End If
            Else
                Range("Ver.PAN").value = UCase(Range("Ver.PAN").value)
            End If

'            If calculateAge(Trim(Range("sheet1.DOB").value)) < 18 And Mid(Range("sheet1.status").value, 1, 1) = "I" Then
'                If Sheet5.Range("Ver.PAN").value = Sheet1.Range("sheet1.PAN").value Then
'                    fmsgbox "*If status is individual & date of birth is on or after 2 April 2002 then in verification PAN cannot be same as mentioned in general information."
'                     Sheet5.Range("Ver.PAN").value = ""
'                End If
'            End If

        End If





'    If Target.name.name = "Ver.Date" Then
'        If Range("Ver.Date").value <> "" Then
'           If Not FormatNCheckDate(Range("Ver.Date").value) Then
'                If Not ChkMinInclusiveDate(Dformat(Range("Ver.Date").value, "yyyy-mm-dd"), "2020-04-01") Then
'                    fmsgbox "* Verification Date should be on or after 01/04/2020"
'                    Range("Ver.Date").value = ""
'                    Range("Ver.Date").Select
'                    GoTo endline
'                End If
'            End If
'        End If
'    End If


  If Target.name.name = "Ver.IdentificationNoOfTRP" Then
        If Target.value <> "" Then
            If IsNumeric(Sheet5.Range("Ver.IdentificationNoOfTRP").value) Then
                If Len(Sheet5.Range("Ver.IdentificationNoOfTRP").value) = 6 And Len(Sheet5.Range("Ver.IdentificationNoOfTRP").value) > 0 Then
                    Sheet5.Range("Ver.IdentificationNoOfTRP").value = UCase(Sheet5.Range("Ver.IdentificationNoOfTRP").value)
                Else
                    fmsgbox "* Please enter valid TRP Identification number"
                    Sheet5.Range("Ver.IdentificationNoOfTRP").value = ""
                    Sheet5.Range("Ver.IdentificationNoOfTRP").Select
                End If
            Else
                If Mid(Sheet5.Range("Ver.IdentificationNoOfTRP").value, 1, 1) = "T" Or Mid(Sheet5.Range("Ver.IdentificationNoOfTRP").value, 1, 1) = "t" Then
                    If IsNumeric(Mid(Sheet5.Range("Ver.IdentificationNoOfTRP").value, 2, 9)) And Len(Mid(Sheet5.Range("Ver.IdentificationNoOfTRP").value, 2, 9)) = 9 Then
                        Sheet5.Range("Ver.IdentificationNoOfTRP").value = UCase(Sheet5.Range("Ver.IdentificationNoOfTRP").value)
                    Else
                        fmsgbox "* Please enter valid TRP Identification number"
                        Sheet5.Range("Ver.IdentificationNoOfTRP").value = ""
                        Sheet5.Range("Ver.IdentificationNoOfTRP").Select
                    End If
                Else
                    fmsgbox "* Please enter valid TRP Identification number"
                    Sheet5.Range("Ver.IdentificationNoOfTRP").value = ""
                    Sheet5.Range("Ver.IdentificationNoOfTRP").Select
                End If
            End If
        End If
    End If


    If Target.name.name = "Refund.AcntNo" Then
        If Not ValidateBankAccountNumber(Sheet5.Range("Refund.AcntNo").value) Then
                fmsgbox errmsgVerification ', vbOKOnly, vbMessageCaption
                Sheet5.Range("Refund.AcntNo").value = ""
                Sheet5.Range("Refund.AcntNo").Select
        End If
            Range("Refund.AcntNo").value = UCase(Range("Refund.AcntNo").value)
            GoTo endline
    End If

End If

If (Target.name.name) = "Ver.Capacity" Then
        If ((Mid(Sheet1.Range("sheet1.status").value, 1, 1) = "I") Or (Mid(Sheet1.Range("sheet1.status").value, 1, 1) = "(")) And Mid(Range("Ver.Capacity").value, 1, 1) = "S" Then
            sPassword = mdCommon.getmsgstate
            Sheet5.Unprotect Password:=sPassword
            Sheet5.Range("Ver.AssesseeVerName").value = UCase(Sheet1.Range("sheet1.FirstName").value + " " + Sheet1.Range("sheet1.MiddleName").value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").value)
            Sheet5.Protect Password:=sPassword
            'GoTo endline
        Else
            Sheet5.Range("Ver.AssesseeVerName").value = ""
        End If

'        If calculateAge(Trim(Range("sheet1.DOB").value)) < 18 And Mid(Range("sheet1.status").value, 1, 1) = "I" Then
'          If Sheet5.Range("Ver.Capacity").value <> "Representative" Then
'            fmsgbox "*If status is individual & date of birth is on or after 2 April 2002 then in verification drop down cannot be selected as self/Karta/ Partner"
'            Sheet5.Range("Ver.Capacity").value = ""
'          End If
'        End If


    End If

'  If Not Application.Intersect(Sheet5.Range("Sheet5.Amount"), Sheet5.Range(Target.Address)) Is Nothing Then
'       If Sheet5.Range(Target.Address).value > 5000 Then
'           If Mid(Sheet5.Range(Replace(Target.Address, "I", "G")).value, 1, 1) = "1" Then
'                MsgBox "If agricultural income is more than Rs 5000/- then use ITR 2 or 3"
'                Sheet5.Range(Target.Address).value = ""
'           End If
'       End If
'
'       If Sheet5.Range(Target.Address).value > 1000000 Then
'           If Mid(Sheet5.Range(Replace(Target.Address, "I", "G")).value, 1, 1) = "2" Then
'                MsgBox "Sec 10(34) can not be more than Rs. 10,00,000/-"
'                Sheet5.Range(Target.Address).value = ""
'           End If
'       End If
'  End If


endline:
''Change.02.03.2023.102.IDS98/99 A
'If Not Application.Intersect(Sheet5.Range("Sheet5.Nature"), Sheet5.Range(Target.Address)) Is Nothing Then
'DPMADeleteComment45 Target
'End If
''End Change IDS98/99
Application.EnableEvents = True
Sheet5.Protect Password:=sPassword
Application.ScreenUpdating = True
End Sub
'Change-28.11.2022.103.05B
'Function CountDropdown(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
'    Dim currCnt  As Long
'    Dim rangeValue As Variant
'    Dim i As Long
'
'    currCnt1 = 0
'     currCnt2 = 0
'      currCnt3 = 0
'      currCnt4 = 0
'      currCnt5 = 0
'      currCnt6 = 0
'      currCnt7 = 0
'      currCnt8 = 0
'      currCnt9 = 0
'     currCnt10 = 0
'      currCnt11 = 0
'      currCnt12 = 0
'      currCnt13 = 0
'      currCnt14 = 0
'      currCnt15 = 0
'
'    ReDim rangeValue(rangeCount)
'      For i = 1 To rangeCount
'        If Not rangecells.item(i).value = "(Select)" Then
'            rangeValue(i) = rangecells.item(i).value
'        End If
'      Next
'
'    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
'      For i = 1 To UBound(rangeValue)
'       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then
'
'       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
'       If Mid(Trim(sDropdownValue), 1, 4) = "Agri" Then
'       currCnt1 = currCnt1 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(34)" Then
'       currCnt2 = currCnt2 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(19)" Then
'       currCnt3 = currCnt3 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 12) = "Sec 10(10BC)" Then
'       currCnt4 = currCnt4 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(10D)" Then
'       currCnt5 = currCnt5 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(11)" Then
'       currCnt6 = currCnt6 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(26)" Then
'       currCnt7 = currCnt7 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 13) = "Sec 10(26AAA)" Then
'       currCnt8 = currCnt8 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(12)" Then
'       currCnt9 = currCnt9 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(13)" Then
'       currCnt10 = currCnt10 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(16)" Then
'       currCnt11 = currCnt11 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(17)" Then
'       currCnt12 = currCnt12 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(17A)" Then
'       currCnt13 = currCnt13 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(18)" Then
'       currCnt14 = currCnt14 + 1
'       End If
'
'       If Mid(Trim(sDropdownValue), 1, 4) = "Defe" Then
'       currCnt15 = currCnt15 + 1
'       End If
'
'
'       End If
'      Next
'    End If
'    CountDropdown = currCnt1
'End Function


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

'Malli--------17/12/2024----
Function GetBankName(IFSC As Variant) As String
On Error Resume Next
GetBankName = ""
Dim rangecells As Range
Dim Address As Variant
For Each rangecells In Sheet8.Range("IFSC_StartLetter")

If UCase(rangecells.value) = UCase(IFSC) Then
Address = rangecells.Address
Address = Replace(Address, "E", "F")
GetBankName = Sheet8.Range(Address).value
Exit Function
End If

Next

'If GetBankName = "" Then
'GetBankName = False
'End If

End Function
