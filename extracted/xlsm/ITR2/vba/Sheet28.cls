Attribute VB_Name = "Sheet28"
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
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet28.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet28.Shapes("Button 38").Left = Range("Button_TCS").Left
Sheet28.Shapes("Button 38").Top = Range("Button_TCS").Top + 3

Sheet28.Shapes("Button 36").Left = Range("Button_TDS1").Left
Sheet28.Shapes("Button 36").Top = Range("Button_TDS1").Top + 3

Sheet28.Shapes("Button 37").Left = Range("Button_TDS2").Left
Sheet28.Shapes("Button 37").Top = Range("Button_TDS2").Top + 3

Sheet28.Shapes("Button 39").Left = Range("Button_TDS3").Left
Sheet28.Shapes("Button 39").Top = Range("Button_TDS3").Top + 3

Sheet28.Protect Password:=sPassword

fmsgboxStatus "Please verify the TAN details and the Amount of credit available in your Form 16 and 26AS" & Chr(10) & Chr(10) ', vbOKOnly, "Alert"


End Sub
Private Sub Worksheet_Change(ByVal Target As Range)

CalculateTaxCheckFlag = False

On Error GoTo endline

ChangeTCSTable Target

Dim i As Variant
Dim rangecell As Range
Application.EnableEvents = False

If Not IsEmpty(Trim(Target.value)) Then
If Target.Validation.Type = 6 Then
    Target.value = UCase(Target.value)
    If Not Application.Intersect(Sheet28.Range("TDS3.PAN"), Sheet28.Range(Target.Address)) Is Nothing Then
    If Not CheckPAN(Sheet28.Range(Target.Address).value) Then
        fmsgboxStatus "Please enter a valid PAN in TDS" ', vbOKOnly, "ITR-3: 2022-23"
        Sheet28.Range(Target.Address).value = ""
        Sheet28.Range(Target.Address).Select
    End If
End If
     GoTo endline
End If
End If

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet28.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

'AY_2025-26 'Shrutika
'//start-------------------------
'TCS POPUP
If Not Intersect(Sheet28.Range("TCS1.Collected37i"), Sheet28.Range(Target.Address)) Is Nothing Or _
   Not Intersect(Sheet28.Range("TCS1.ClaimedOwnHands"), Sheet28.Range(Target.Address)) Is Nothing Then
 
   If Target.value > 0 Then
        If (Not Intersect(Sheet28.Range("TCS1.Collected37i"), Sheet28.Range(Target.Address)) Is Nothing And Sheet28.Range(Replace(Target.Address, "L", "F")).value <> "Self") Or _
           (Not Intersect(Sheet28.Range("TCS1.ClaimedOwnHands"), Sheet28.Range(Target.Address)) Is Nothing And Sheet28.Range(Replace(Target.Address, "M", "F")).value <> "Self") Then
                 MsgBox """Before claiming credit for TCS from the hands of ""Other person"", please ensure that such other person also files the return of income, disclosing the TCS collected in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TCS of the said return of Income." & vbNewLine & "However, if the TCS claimed is out of the minor's payments, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent by disclosing the PAN of Minor), after ensuring that the sources for minor's payments are relevant in the entries made/ available in the return of income.""", , "Alert"
        End If
   End If
 
End If
 
 
'TDS2/TDS3 POPUP
  
If Not Intersect(Sheet28.Range("TDS2.DeductedTDS5A"), Sheet28.Range(Target.Address)) Is Nothing Or _
   Not Intersect(Sheet28.Range("TDS2.ClaimedInOwnHands"), Sheet28.Range(Target.Address)) Is Nothing Or _
   Not Intersect(Sheet28.Range("TDS3.DeductedTDS5A"), Sheet28.Range(Target.Address)) Is Nothing Or _
   Not Intersect(Sheet28.Range("TDS3.ClaimedInOwnHands"), Sheet28.Range(Target.Address)) Is Nothing Then
   If Target.value > 0 Then
        If (Not Intersect(Sheet28.Range("TDS2.ClaimedInOwnHands"), Sheet28.Range(Target.Address)) Is Nothing And Sheet28.Range(Replace(Target.Address, "N", "E")).value <> "Self") Or _
           (Not Intersect(Sheet28.Range("TDS2.DeductedTDS5A"), Sheet28.Range(Target.Address)) Is Nothing And Sheet28.Range(Replace(Target.Address, "M", "E")).value <> "Self") Or _
           (Not Intersect(Sheet28.Range("TDS3.ClaimedInOwnHands"), Sheet28.Range(Target.Address)) Is Nothing And Sheet28.Range(Replace(Target.Address, "O", "E")).value <> "Self") Or _
           (Not Intersect(Sheet28.Range("TDS3.DeductedTDS5A"), Sheet28.Range(Target.Address)) Is Nothing And Sheet28.Range(Replace(Target.Address, "N", "E")).value <> "Self") Then
              MsgBox """Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return." & vbNewLine & "However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income.""", , "Alert"
        End If
   End If
End If
'---------------------------End\\
'--------------------


If Not Intersect(Sheet28.Range("TDS2.TDSCreditName"), Sheet28.Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HD8D8D8"

    Else
        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"

    End If
End If
'
'If Not Intersect(Range("TDS2.HeadOfIncome"), Range(Target.Address)) Is Nothing Then
'    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "N" Then
'        Sheet28.Range(Replace(Target.Address, "T", "S")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "T", "S")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "T", "S")).Interior.Color = "&HD8D8D8"
'     Else
'        Sheet28.Range(Replace(Target.Address, "T", "S")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "T", "S")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "T", "S")).Interior.Color = "&HCCFFCC"
'    End If
'End If

If Not Intersect(Range("TDS2.DeductedInOwnHands"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
      If Sheet28.Range(Replace(Target.Address, "L", "K")).value > 0 Or _
         Sheet28.Range(Replace(Target.Address, "L", "J")).value <> "(Select)" And Sheet28.Range(Replace(Target.Address, "L", "J")).value <> "" Then

            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 2"
      End If
    End If
End If

If Not Intersect(Range("TDS2.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "K", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "K", "L")).value > 0 Then
          
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 2"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS2.DeductedYear"), Range(Target.Address)) Is Nothing Then
    If Not isdropdownblank(Sheet28.Range(Target.Address).value) Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "J", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "J", "L")).value > 0 Then
         
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 2"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS3.TDSCreditName"), Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "O")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "T")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "T")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "T")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "O")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "T")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "T")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "T")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HD8D8D8"

    Else
        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "O")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "T")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "T")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "T")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

    End If
End If

If Not Intersect(Range("TDS3.DeductedInOwnHands"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
        If Sheet28.Range(Replace(Target.Address, "M", "L")).value > 0 Or _
           Sheet28.Range(Replace(Target.Address, "M", "K")).value <> "(Select)" And Sheet28.Range(Replace(Target.Address, "M", "K")).value <> "" Then
           
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 3"
        End If
    End If
End If

If Not Intersect(Range("TDS3.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "L", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "L", "M")).value > 0 Then
          
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 3"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS3.DeductedYear"), Range(Target.Address)) Is Nothing Then
    If Not isdropdownblank(Sheet28.Range(Target.Address).value) Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "K", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "L", "N")).value > 0 Then
          
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 3"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS2.ClaimedInOwnHands"), Range(Target.Address)) Is Nothing Then

        If Sheet28.Range(Target.Address).value > 0 Then
            Sheet28.Range(Replace(Target.Address, "O", "T")).Locked = False
            Sheet28.Range(Replace(Target.Address, "O", "T")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(Target.Address, "O", "U")).Locked = False
            Sheet28.Range(Replace(Target.Address, "O", "U")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(Target.Address, "O", "U")).value = "(Select)"
        Else
            Sheet28.Range(Replace(Target.Address, "O", "T")).ClearContents
            Sheet28.Range(Replace(Target.Address, "O", "T")).Locked = True
            Sheet28.Range(Replace(Target.Address, "O", "T")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(Target.Address, "O", "U")).ClearContents
            Sheet28.Range(Replace(Target.Address, "O", "U")).Locked = True
            Sheet28.Range(Replace(Target.Address, "O", "U")).Interior.Color = "&HD8D8D8"


    End If
End If


If Not Intersect(Range("TDS3.ClaimedInOwnHands"), Range(Target.Address)) Is Nothing Then

        If Sheet28.Range(Target.Address).value > 0 Then
            Sheet28.Range(Replace(Target.Address, "P", "U")).Locked = False
            Sheet28.Range(Replace(Target.Address, "P", "U")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(Target.Address, "P", "V")).Locked = False
            Sheet28.Range(Replace(Target.Address, "P", "V")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(Target.Address, "P", "V")).value = "(Select)"
        Else
            Sheet28.Range(Replace(Target.Address, "P", "U")).ClearContents
            Sheet28.Range(Replace(Target.Address, "P", "U")).Locked = True
            Sheet28.Range(Replace(Target.Address, "P", "U")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(Target.Address, "P", "V")).ClearContents
            Sheet28.Range(Replace(Target.Address, "P", "V")).Locked = True
            Sheet28.Range(Replace(Target.Address, "P", "V")).Interior.Color = "&HD8D8D8"


    End If
End If

If Not Intersect(Range("TDS2.DeductedInOwnHands"), Range(Target.Address)) Is Nothing Then
Sheet28.Unprotect Password:=getmsgstate
        If Sheet28.Range(Target.Address).value = 0 Then

            Sheet28.Range(Replace(Target.Address, "L", "J")).Locked = False
            Sheet28.Range(Replace(Target.Address, "L", "J")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(Target.Address, "L", "J")).value = "(Select)"
            
            Sheet28.Range(Replace(Target.Address, "L", "K")).Locked = False
            Sheet28.Range(Replace(Target.Address, "L", "K")).Interior.Color = "&HCCFFCC"
            Sheet28.Protect Password:=getmsgstate
        Else
        Sheet28.Unprotect Password:=getmsgstate
            Sheet28.Range(Replace(Target.Address, "L", "J")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "J")).Locked = True
            Sheet28.Range(Replace(Target.Address, "L", "J")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(Target.Address, "L", "K")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "K")).Locked = True
            Sheet28.Range(Replace(Target.Address, "L", "K")).Interior.Color = "&HD8D8D8"
            Sheet28.Protect Password:=getmsgstate

    End If
End If


If Not Intersect(Range("TDS2.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
Sheet28.Unprotect Password:=getmsgstate
        If Sheet28.Range(Target.Address).value = 0 And Sheet28.Range(Replace(Target.Address, "K", "E")).value <> "Other Person" Then
                    
            Sheet28.Range(Replace(Target.Address, "K", "L")).Locked = False
            Sheet28.Range(Replace(Target.Address, "K", "L")).Interior.Color = "&HCCFFCC"
            Sheet28.Protect Password:=getmsgstate
        Else
Sheet28.Unprotect Password:=getmsgstate
            Sheet28.Range(Replace(Target.Address, "K", "L")).ClearContents
            Sheet28.Range(Replace(Target.Address, "K", "L")).Locked = True
            Sheet28.Range(Replace(Target.Address, "K", "L")).Interior.Color = "&HD8D8D8"
            Sheet28.Protect Password:=getmsgstate

    End If
End If

If Not Intersect(Range("TDS3.DeductedInOwnHands"), Range(Target.Address)) Is Nothing Then
Sheet28.Unprotect Password:=getmsgstate
        If Sheet28.Range(Target.Address).value = 0 Then
        
            Sheet28.Range(Replace(Target.Address, "M", "K")).Locked = False
            Sheet28.Range(Replace(Target.Address, "M", "K")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(Target.Address, "M", "K")).value = "(Select)"
            
            Sheet28.Range(Replace(Target.Address, "M", "L")).Locked = False
            Sheet28.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
            Sheet28.Protect Password:=getmsgstate
        Else
        Sheet28.Unprotect Password:=getmsgstate
        
            Sheet28.Range(Replace(Target.Address, "M", "K")).ClearContents
            Sheet28.Range(Replace(Target.Address, "M", "K")).Locked = True
            Sheet28.Range(Replace(Target.Address, "M", "K")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(Target.Address, "M", "L")).ClearContents
            Sheet28.Range(Replace(Target.Address, "M", "L")).Locked = True
            Sheet28.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
Sheet28.Protect Password:=getmsgstate
    End If
End If

If Not Intersect(Range("TDS3.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
Sheet28.Unprotect Password:=getmsgstate
        If Sheet28.Range(Target.Address).value = 0 Then
                    
            Sheet28.Range(Replace(Target.Address, "L", "M")).Locked = False
            Sheet28.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
            Sheet28.Protect Password:=getmsgstate
        Else
Sheet28.Unprotect Password:=getmsgstate
            Sheet28.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet28.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
Sheet28.Protect Password:=getmsgstate
    End If
End If

'If Not Intersect(Range("TCS.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
'Sheet28.Unprotect Password:=getmsgstate
'        If Sheet28.Range(Target.Address).value = 0 Then
'
'            Sheet28.Range(Replace(Target.Address, "H", "I")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HCCFFCC"
'            Sheet28.Protect Password:=getmsgstate
'        Else
'Sheet28.Unprotect Password:=getmsgstate
'            Sheet28.Range(Replace(Target.Address, "H", "I")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "H", "I")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "H", "I")).Interior.Color = "&HD8D8D8"
'Sheet28.Protect Password:=getmsgstate
'    End If
'End If

'If Not Intersect(Range("TCS.TotalTCS"), Range(Target.Address)) Is Nothing Then
'Sheet28.Unprotect Password:=getmsgstate
'        If Sheet28.Range(Target.Address).value = 0 Then
'            Sheet28.Range(Replace(Target.Address, "I", "G")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "I", "G")).Interior.Color = "&HCCFFCC"
'            Sheet28.Range(Replace(Target.Address, "I", "G")).value = "(Select)"
'
'            Sheet28.Range(Replace(Target.Address, "I", "H")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "I", "H")).Interior.Color = "&HCCFFCC"
'            Sheet28.Protect Password:=getmsgstate
'        Else
'        Sheet28.Unprotect Password:=getmsgstate
'            Sheet28.Range(Replace(Target.Address, "I", "G")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "I", "G")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "I", "G")).Interior.Color = "&HD8D8D8"
'
'            Sheet28.Range(Replace(Target.Address, "I", "H")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "I", "H")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "I", "H")).Interior.Color = "&HD8D8D8"
'Sheet28.Protect Password:=getmsgstate
'    End If
'End If



'
'If Not Intersect(Range("TCS1.CreditName"), Range(Target.Address)) Is Nothing Then
'    TCS_Change Target
'End If
'
'If Not Intersect(Range("TCS1.UnclaimedBFAmount"), Range(Target.Address)) Is Nothing Then
'    TCS_UnclaimedBFAmount Target
'End If
'
'If Not Intersect(Range("TCS1.CollectedOwnHands"), Range(Target.Address)) Is Nothing Then
'    TCS_CollectedOwnHands Target
'End If
'
'If Not Intersect(Range("TCS1.ClaimedOwnHands"), Range(Target.Address)) Is Nothing Then
'    TCS_ClaimedOwnHands Target
'End If


'================================================================================================================

Dim TDS_Section_warning1, TDS_Section_warning2, TDS_section_error As Range
Dim TDS_Section_warning_count1, TDS_Section_warning_count2, TDS_section_error_count, iCount_warning1, iCount_warning2, iCount_error As Long

'Newly added by sai on 24th April As per DE V6
If Not Intersect(Range("TDS2.TdsDeducted"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("TDS3.TdsIsDeducted"), Range(Target.Address)) Is Nothing Then

     If Target.value <> "" Or UCase(Target.value) <> UCase("(Select)") Then
       Set TDS_Section_warning1 = SheetALL.Range("TDS_Section_Warning1").Cells
       TDS_Section_warning_count1 = SheetALL.Range("TDS_Section_warning1").count
       
       For iCount_warning1 = 1 To TDS_Section_warning_count1
         
         If UCase(Trim(Target.value)) = UCase(Trim(TDS_Section_warning1.item(iCount_warning1).value)) Then
            'MsgBox "*Incomes chargeable to tax at special rates are not expected to be reported in ITR  4. Accordingly, TDS relating to the special rates income namely 194B,194BB,194BA,194IA,194IC,194LA,194S are also not expected to be claimed. Thus, please ensure that the return being filed in ITR 4 is not defective.", vbInformation
          'MsgBox "*The Income on which the TDS under this section has been deducted is payable to a non-resident.Returns for income of non-residents are required to be filed in ITR 3/5. Thus, kindly ensure that the return being filed in ITR 4 is not defective.", vbInformation
           MsgBox "*The income on which TDS under this section has been deducted is not expected to be reported in ITR 2. Kindly ensure that the return being filed in ITR 2 is not defective.", vbInformation
          
          End If
       Next
       
'       Set TDS_Section_warning2 = SheetALL.Range("TDS_Section_warning2").Cells
'       TDS_Section_warning_count2 = SheetALL.Range("TDS_Section_warning2").count
'
'       For iCount_warning2 = 1 To TDS_Section_warning_count2
'
'         If UCase(Trim(Target.value)) = UCase(Trim(TDS_Section_warning2.item(iCount_warning2).value)) Then
'            MsgBox "*The Income on which the TDS under this section has been deducted is payable to a non-resident.Returns for income of non-residents are required to be filed in ITR 3/5. Thus, kindly ensure that the return being filed in ITR 4 is not defective.", vbInformation
'          End If
'       Next
       
    
       Set TDS_section_error = SheetALL.Range("TDS_Section_Error").Cells
       TDS_section_error_count = SheetALL.Range("TDS_Section_Error").count
       
        For iCount_error = 1 To TDS_section_error_count
       
          If UCase(Trim(Target.value)) = UCase(Trim(TDS_section_error.item(iCount_error).value)) Then
            MsgBox "*TDS relating to salaries is required to be claimed in Schedule TDS 1."
            Target.value = "(Select)"
          End If
       
        Next
     
     
     End If

End If

'TDS-2

'TDS Credit
If Not Intersect(Range("TDS2.TDSCreditName"), Range(Target.Address)) Is Nothing Then

If Target.value = "Other Person" Then
MsgBox "*Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return. However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent (by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income.", vbInformation

End If
End If

'Deduction in the hands of spouce

If Not Intersect(Range("TCS.ClaimedOwnHands"), Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then
MsgBox "*Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return. However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent (by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income", vbInformation

End If
End If


'Claimed in own hands

If Not Intersect(Range("TDS2.ClaimedInOwnHands"), Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then
MsgBox "*Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return. However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent (by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income.", vbInformation

End If
End If

'=====================================

'TDS-3

'TDS Credit
If Not Intersect(Range("TDS3.TDSCreditName"), Range(Target.Address)) Is Nothing Then

If Target.value = "Other Person" Then
MsgBox "*Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return. However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent (by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income.", vbInformation

End If
End If

'Deduction in the hands of spouce

If Not Intersect(Range("TDS3.DeductedIncome5A"), Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then
MsgBox "*Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return. However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent (by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income.", vbInformation

End If
End If


'Claimed in own hands

'If Not Intersect(Range("TDS3.ClaimedInOwnHands"), Range(Target.Address)) Is Nothing Then
'
'If Target.value > 0 Then
'MsgBox "*Before claiming credit for TDS from hands of ""Other person"", please ascertain that such other person also files the return of income, disclosing the TDS deducted in his own hands and the extent of credit that can be claimed by you by providing your PAN in schedule TDS of the said return of Income. Also ensure that the corresponding income is offered to tax in the return. However, if the TDS claimed is out of the minor's income, and the minor's income is clubbed in the hands of the parent, then schedule SPI needs to be filled by Parent (by disclosing the PAN of Minor), after ensuring that the minor's income is clubbed and offered to tax in the return of income.", vbInformation
'
'End If
'End If

'=====================================

endline:

Sheet28.Protect Password:=sPassword
Application.EnableEvents = True

End Sub

Private Sub Worksheet_Changeold1(ByVal Target As Range)

CalculateTaxCheckFlag = False

On Error GoTo endline
Dim i As Variant
Dim rangecell As Range
Application.EnableEvents = False

If Not IsEmpty(Trim(Target.value)) Then
If Target.Validation.Type = 6 Then
    Target.value = UCase(Target.value)
    If Not Application.Intersect(Sheet28.Range("TDS3.PAN"), Sheet28.Range(Target.Address)) Is Nothing Then
    If Not CheckPAN(Sheet28.Range(Target.Address).value) Then
        fmsgboxStatus "Please enter a valid PAN in TDS" ', vbOKOnly, "ITR-3: 2022-23"
        Sheet28.Range(Target.Address).value = ""
        Sheet28.Range(Target.Address).Select
    End If
End If
     GoTo endline
End If
End If

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet28.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

If Not Intersect(Range("TDS2.TDSCreditName"), Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "O")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "K")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "K")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "O")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "K")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "K")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HD8D8D8"

    Else
        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "O")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "O")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "K")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "K")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"

    End If
End If

If Not Intersect(Range("TDS2.HeadOfIncome"), Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "N" Then
        Sheet28.Range(Replace(Target.Address, "T", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "T", "S")).Locked = True
        Sheet28.Range(Replace(Target.Address, "T", "S")).Interior.Color = "&HD8D8D8"
     Else
        Sheet28.Range(Replace(Target.Address, "T", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "T", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "T", "S")).Interior.Color = "&HCCFFCC"
    End If
End If

If Not Intersect(Range("TDS2.DeductedInOwnHands"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
      If Sheet28.Range(Replace(Target.Address, "K", "J")).value > 0 Or _
         Sheet28.Range(Replace(Target.Address, "K", "I")).value <> "(Select)" And Sheet28.Range(Replace(Target.Address, "K", "I")).value <> "" Then

            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 2"
      End If
    End If
End If

If Not Intersect(Range("TDS2.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "J", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "J", "K")).value > 0 Then
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 2"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS2.DeductedYear"), Range(Target.Address)) Is Nothing Then
    If Not isdropdownblank(Sheet28.Range(Target.Address).value) Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "I", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "I", "K")).value > 0 Then
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 2"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS3.TDSCreditName"), Range(Target.Address)) Is Nothing Then
    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HD8D8D8"

    Else
        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "F")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "F")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "F")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"

        Sheet28.Range(Replace(Target.Address, "E", "P")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "P")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "P")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "Q")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "Q")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "R")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "R")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "R")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "E", "S")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "S")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "S")).Interior.Color = "&HCCFFCC"

        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"

    End If
End If

If Not Intersect(Range("TDS3.DeductedInOwnHands"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
        If Sheet28.Range(Replace(Target.Address, "L", "K")).value > 0 Or _
           Sheet28.Range(Replace(Target.Address, "L", "J")).value <> "(Select)" And Sheet28.Range(Replace(Target.Address, "L", "J")).value <> "" Then
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 3"
        End If
    End If
End If

If Not Intersect(Range("TDS3.BroughtFwdTDSAmt"), Range(Target.Address)) Is Nothing Then
    If (Sheet28.Range(Target.Address).value > 0 And Sheet28.Range(Target.Address).value <> "") Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "K", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "K", "L")).value > 0 Then
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 3"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS3.DeductedYear"), Range(Target.Address)) Is Nothing Then
    If Not isdropdownblank(Sheet28.Range(Target.Address).value) Then
        If UCase(Mid(Sheet28.Range(Replace(Target.Address, "J", "E")).value, 1, 1)) = "S" Then
          If Sheet28.Range(Replace(Target.Address, "J", "L")).value > 0 Then
            fmsgboxStatus """Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Sch TDS 3"
          End If
        End If
    End If
End If

If Not Intersect(Range("TDS2.ClaimedInOwnHands"), Range(Target.Address)) Is Nothing Then

        If Sheet28.Range(Target.Address).value > 0 Then
            Sheet28.Range(Replace(Target.Address, "N", "S")).Locked = False
            Sheet28.Range(Replace(Target.Address, "N", "S")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(Target.Address, "N", "T")).Locked = False
            Sheet28.Range(Replace(Target.Address, "N", "T")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(Target.Address, "N", "T")).value = "(Select)"
        Else
            Sheet28.Range(Replace(Target.Address, "N", "S")).ClearContents
            Sheet28.Range(Replace(Target.Address, "N", "S")).Locked = True
            Sheet28.Range(Replace(Target.Address, "N", "S")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(Target.Address, "N", "T")).ClearContents
            Sheet28.Range(Replace(Target.Address, "N", "T")).Locked = True
            Sheet28.Range(Replace(Target.Address, "N", "T")).Interior.Color = "&HD8D8D8"


    End If
End If


If Not Intersect(Range("TDS3.ClaimedInOwnHands"), Range(Target.Address)) Is Nothing Then

        If Sheet28.Range(Target.Address).value > 0 Then
            Sheet28.Range(Replace(Target.Address, "O", "T")).Locked = False
            Sheet28.Range(Replace(Target.Address, "O", "T")).Interior.Color = "&HCCFFCC"

            Sheet28.Range(Replace(Target.Address, "O", "U")).Locked = False
            Sheet28.Range(Replace(Target.Address, "O", "U")).Interior.Color = "&HCCFFCC"
            Sheet28.Range(Replace(Target.Address, "O", "U")).value = "(Select)"
        Else
            Sheet28.Range(Replace(Target.Address, "O", "T")).ClearContents
            Sheet28.Range(Replace(Target.Address, "O", "T")).Locked = True
            Sheet28.Range(Replace(Target.Address, "O", "T")).Interior.Color = "&HD8D8D8"

            Sheet28.Range(Replace(Target.Address, "O", "U")).ClearContents
            Sheet28.Range(Replace(Target.Address, "O", "U")).Locked = True
            Sheet28.Range(Replace(Target.Address, "O", "U")).Interior.Color = "&HD8D8D8"


    End If
End If


'TCS NEW






endline:

Sheet28.Protect Password:=sPassword
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









