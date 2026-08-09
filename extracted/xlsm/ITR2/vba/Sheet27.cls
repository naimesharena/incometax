Attribute VB_Name = "Sheet27"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String


Private Sub Worksheet_Activate()
 On Error Resume Next
 Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   sourceSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----


sourceSheet.Shapes("Button 79").Left = Range("Button_FAA1").Left
sourceSheet.Shapes("Button 79").Top = Range("Button_FAA1").Top + 3

sourceSheet.Shapes("Button 80").Left = Range("Button_FAA2").Left
sourceSheet.Shapes("Button 80").Top = Range("Button_FAA2").Top + 3

sourceSheet.Shapes("Button 82").Left = Range("Button_FAA3").Left
sourceSheet.Shapes("Button 82").Top = Range("Button_FAA3").Top + 3

sourceSheet.Shapes("Button 84").Left = Range("Button_FAA4").Left
sourceSheet.Shapes("Button 84").Top = Range("Button_FAA4").Top + 3

sourceSheet.Shapes("Button 18").Left = Range("Button_FA2").Left
sourceSheet.Shapes("Button 18").Top = Range("Button_FA2").Top + 3

sourceSheet.Shapes("Button 19").Left = Range("Button_FA3").Left
sourceSheet.Shapes("Button 19").Top = Range("Button_FA3").Top + 3

sourceSheet.Shapes("Button 20").Left = Range("Button_FA4").Left
sourceSheet.Shapes("Button 20").Top = Range("Button_FA4").Top + 3

sourceSheet.Shapes("Button 21").Left = Range("Button_FA5").Left
sourceSheet.Shapes("Button 21").Top = Range("Button_FA5").Top + 3

sourceSheet.Shapes("Button 22").Left = Range("Button_FA6").Left
sourceSheet.Shapes("Button 22").Top = Range("Button_FA6").Top + 3

'Lock - Unlock FA---------------------------------

'-------------------------------------------------

sourceSheet.Protect Password:=sPassword
Sheet27.Shapes("Button TR").Visible = msoFalse
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Application.EnableEvents = False
 Dim sourceSheet As Worksheet
 Set sourceSheet = ThisWorkbook.Sheets("TR_FA")


If Not Intersect(Range("FA_A2_Grossinterest_Nature"), Range(Target.Address)) Is Nothing Then
    If Range(Target.Address).value = "" Or Range(Target.Address).value = "(Select)" Then
        '------------ Sheet unprotect-----------
        sourceSheet.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------


         Range(Replace(Target.Address, "N", "O")).Interior.Color = (&HD8D8D8)
            Range(Replace(Target.Address, "N", "O")).value = ""
            Range(Replace(Target.Address, "N", "O")).Locked = True


       '-------------Sheet Protect --------------
           sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    Else
    '------------ Sheet unprotect-----------
        Sheet27.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------

          Range(Replace(Target.Address, "N", "O")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "N", "O")).value = ""
        Range(Replace(Target.Address, "N", "O")).Locked = False



        '-------------Sheet Protect --------------
        sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    End If
    GoTo endline
End If



If Not Intersect(Range("FA_E_IncomeAccured"), Range(Target.Address)) Is Nothing Then
    If Range(Target.Address).value <> "No" Then
        '------------ Sheet unprotect-----------
        sourceSheet.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------
        Range(Replace(Target.Address, "L", "M")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "L", "M")).value = ""
        Range(Replace(Target.Address, "L", "M")).Locked = False

        Range(Replace(Target.Address, "L", "N")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "L", "N")).value = ""
        Range(Replace(Target.Address, "L", "N")).Locked = False

        Range(Replace(Target.Address, "L", "O")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "L", "O")).value = ""
        Range(Replace(Target.Address, "L", "O")).Locked = False

        Range(Replace(Target.Address, "L", "P")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "L", "P")).value = ""
        Range(Replace(Target.Address, "L", "P")).Locked = False


       '-------------Sheet Protect --------------
           sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    Else
    '------------ Sheet unprotect-----------
        Sheet27.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------

            Range(Replace(Target.Address, "L", "M")).Interior.Color = (&HD8D8D8)
            Range(Replace(Target.Address, "L", "M")).value = ""
            Range(Replace(Target.Address, "L", "M")).Locked = True

            Range(Replace(Target.Address, "L", "N")).Interior.Color = (&HD8D8D8)
            Range(Replace(Target.Address, "L", "N")).value = ""
            Range(Replace(Target.Address, "L", "N")).Locked = True

            Range(Replace(Target.Address, "L", "O")).Interior.Color = (&HD8D8D8)
            Range(Replace(Target.Address, "L", "O")).value = ""
            Range(Replace(Target.Address, "L", "O")).Locked = True

            Range(Replace(Target.Address, "L", "P")).Interior.Color = (&HD8D8D8)
            Range(Replace(Target.Address, "L", "P")).value = ""
            Range(Replace(Target.Address, "L", "P")).Locked = True

        '-------------Sheet Protect --------------
        sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    End If
    GoTo endline
End If


If Not Intersect(Range("FA_F_IncomeDerivedIsTaxable"), Range(Target.Address)) Is Nothing Then
    If Range(Target.Address).value <> "No" Then
        '------------ Sheet unprotect-----------
        sourceSheet.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------
        Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "P", "Q")).value = ""
        Range(Replace(Target.Address, "P", "Q")).Locked = False

        Range(Replace(Target.Address, "P", "R")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "P", "R")).value = ""
        Range(Replace(Target.Address, "P", "R")).Locked = False

        Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "P", "S")).value = ""
        Range(Replace(Target.Address, "P", "S")).Locked = False

        Range(Replace(Target.Address, "P", "T")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "P", "T")).value = ""
        Range(Replace(Target.Address, "P", "T")).Locked = False

       '-------------Sheet Protect --------------
           sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    Else
    '------------ Sheet unprotect-----------
        Sheet27.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------

        Range(Replace(Target.Address, "P", "Q")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "P", "Q")).value = ""
        Range(Replace(Target.Address, "P", "Q")).Locked = True

        Range(Replace(Target.Address, "P", "R")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "P", "R")).value = ""
        Range(Replace(Target.Address, "P", "R")).Locked = True

        Range(Replace(Target.Address, "P", "S")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "P", "S")).value = ""
        Range(Replace(Target.Address, "P", "S")).Locked = True

        Range(Replace(Target.Address, "P", "T")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "P", "T")).value = ""
        Range(Replace(Target.Address, "P", "T")).Locked = True

        '-------------Sheet Protect --------------
        sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    End If
    GoTo endline
End If


If Not Intersect(Range("FA_G_WhetherTaxable"), Range(Target.Address)) Is Nothing Then
    If Range(Target.Address).value <> "No" Then
        '------------ Sheet unprotect-----------
        sourceSheet.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------

        Range(Replace(Target.Address, "K", "L")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "K", "L")).value = ""
        Range(Replace(Target.Address, "K", "L")).Locked = False

        Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "K", "M")).value = ""
        Range(Replace(Target.Address, "K", "M")).Locked = False

        Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HCCFFCC)
        Range(Replace(Target.Address, "K", "N")).value = ""
        Range(Replace(Target.Address, "K", "N")).Locked = False


       '-------------Sheet Protect --------------
           sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    Else
    '------------ Sheet unprotect-----------
        Sheet27.Unprotect Password:=getmsgstate
        'Application.EnableEvents = False
        '---------------------------------------

        Range(Replace(Target.Address, "K", "L")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "K", "L")).value = ""
        Range(Replace(Target.Address, "K", "L")).Locked = True

        Range(Replace(Target.Address, "K", "M")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "K", "M")).value = ""
        Range(Replace(Target.Address, "K", "M")).Locked = True

        Range(Replace(Target.Address, "K", "N")).Interior.Color = (&HD8D8D8)
        Range(Replace(Target.Address, "K", "N")).value = ""
        Range(Replace(Target.Address, "K", "N")).Locked = True


       '-------------Sheet Protect --------------
        sourceSheet.Protect Password:=getmsgstate
        'Application.EnableEvents = True
        '----------------------------------------
    End If
    GoTo endline
End If


 If (Target.name.name) = "TR_TaxPaidOutsideIndFlg" Then
            LockUnlockTR_ReliefRefund
 End If
endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error GoTo endline
On Error Resume Next
     Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Application.EnableEvents = False

   Dim halfval, SecHval, firstval, secval As String
    firstval = ""
    secval = ""

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

    If Target.Validation.Type = 6 Then
        If Target.name.name = "TR_AssmtYrTaxRelief" Then
            halfval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 1, 4)
            SecHval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 5, 1)
            firstval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 3, 2)
            secval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 6, 2)
          If sourceSheet.Range("TR_AssmtYrTaxRelief").value <> "" Then
            If SecHval <> "-" Then
                fmsgboxStatus "Enter in format YYYY-YY, eg 2024-25" ', vbOKOnly, "Assessment Year" 'dpk1601
                sourceSheet.Range("TR_AssmtYrTaxRelief").value = ""
                sourceSheet.Range("TR_AssmtYrTaxRelief").Select
                GoTo endline
            End If

            If Not IsNumeric(halfval) Then
                fmsgboxStatus "Enter in format YYYY-YY, eg 2024-25" ', vbOKOnly, "Assessment Year" 'dpk1601
                sourceSheet.Range("TR_AssmtYrTaxRelief").value = ""
                sourceSheet.Range("TR_AssmtYrTaxRelief").Select
                GoTo endline
            End If

            If Not IsNumeric(secval) Then
                fmsgboxStatus "Enter in format YYYY-YY, eg 2024-25" ', vbOKOnly, "Assessment Year" 'dpk1601
                sourceSheet.Range("TR_AssmtYrTaxRelief").value = ""
                sourceSheet.Range("TR_AssmtYrTaxRelief").Select
                GoTo endline
            End If

            Dim firstval_1 As String
             firstval_1 = Mid(firstval + 1, 1, 2)
             If (Len(firstval_1) = 1) Then
             firstval_1 = "0" & firstval_1
             End If


            If secval <> IIf(Len(firstval + 1) > 2, Mid(firstval + 1, 2, 2), firstval_1) Then
                fmsgboxStatus "Year should be consecutive" ', vbOKOnly, "Assessment Year"
                sourceSheet.Range("TR_AssmtYrTaxRelief").value = ""
                sourceSheet.Range("TR_AssmtYrTaxRelief").Select
                GoTo endline
            End If

        End If
    End If
    End If




endline:

Application.EnableEvents = True
End Sub



Sub LockUnlockTR_ReliefRefund()

Application.EnableEvents = False
 Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   sourceSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
If sourceSheet.Range("TR_TaxPaidOutsideIndFlg").value <> "YES" Then

     sourceSheet.Range("TR_AmtTaxRefunded").Locked = True
     sourceSheet.Range("TR_AmtTaxRefunded").ClearContents
     sourceSheet.Range("TR_AmtTaxRefunded").Interior.Color = (&HD8D8D8)

     sourceSheet.Range("TR_AssmtYrTaxRelief").Locked = True
     sourceSheet.Range("TR_AssmtYrTaxRelief").Interior.Color = (&HD8D8D8)
    sourceSheet.Range("TR_AssmtYrTaxRelief").ClearContents

Else
      sourceSheet.Range("TR_AmtTaxRefunded").Locked = False
      sourceSheet.Range("TR_AmtTaxRefunded").Interior.Color = (&HCCFFCC)

      sourceSheet.Range("TR_AssmtYrTaxRelief").Locked = False
      sourceSheet.Range("TR_AssmtYrTaxRelief").Interior.Color = (&HCCFFCC)

End If


'----------------Lock Password-------------------START---
   sourceSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

