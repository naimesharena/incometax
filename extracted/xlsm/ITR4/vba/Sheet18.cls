Attribute VB_Name = "Sheet18"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False

fmsgboxStatus " Please verify the Tax payment details (BSR code, Challan details and date of payment and amount - not to be rounded off ) as available in your form 26 AS" ', vbOKOnly, "Schedule IT "

If Mid(Sheet1.Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
'Ayush_UR
MsgBox "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax , credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also in Schedule IT-2.Please be noted that credit will be allowed only once." & vbCrLf & "Also, ""REGULAR ASSESSMENT TAX"" if any paid needs to be disclosed in sl.no.8 of Part B_ATI and schedule IT-2 at sl.no.14B."

MsgBox "If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments in Schedule IT-1."
'MsgBox "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax, credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also at Sl. No. 14B of Part B-ATI. Please be noted that credit will be allowed only once. Also, REGULAR ASSESSMENT TAX if any paid needs to be disclosed at Sl. No. 8 & Sl. No. 14B of Part B-ATI. If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments at Sl. No. 14A Part B-ATI "
End If

    ActiveSheet.Shapes("Button 10").Left = Range("SchTDS.Button4").Left
    ActiveSheet.Shapes("Button 10").Top = Range("SchTDS.Button4").Top

'    ActiveSheet.Shapes("Button 30").Left = Range("SchTDS3.Button2").Left
'    ActiveSheet.Shapes("Button 30").Top = Range("SchTDS3.Button2").Top
'
'    ActiveSheet.Shapes("Button 26").Left = Range("SchTDS3.Button3").Left
'    ActiveSheet.Shapes("Button 26").Top = Range("SchTDS3.Button3").Top



endline:
Application.EnableEvents = True
Sheet18.Protect Password:=getmsgstate
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False

CalculateTaxCheckFlag = False

Sheet18.Unprotect Password:=getmsgstate

'Target.value = UCase(Target.value)

If Not checkfieldSuperSpecialcharacter(Target.value) Then
         fmsgbox ("* Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
End If
If Not Application.Intersect(Sheet18.Range("TaxP.DateDep"), Sheet18.Range(Target.Address)) Is Nothing Then
        'Change-13.12.2022.103.02
        'If Not (ValidateDateDep_IT()) Then
        
        'IT-E2 AY_2023_2024 Old
        'If Not (ValidateDateDep_ITAY23_24(Target)) Then''''malli
        
        ''IT-C2 AY_2024_2025 change Malli
        If Not (ValidateDateDep_ITAY24_25(Target)) Then
        'malli
'            MsgBox "INVALID Date of Deposit"
         End If
        
         

        GoTo endline

    End If
    Sheet18.Unprotect Password:=getmsgstate

    
'Sudha_2025-26
 If Not Application.Intersect(Sheet18.Range("TaxP.BSRCode"), Sheet18.Range(Target.Address)) Is Nothing Then

         If Target.value = "0000000" Then
         fmsgbox ("* Please enter a valid 7 characters BSR Code" & Chr(13))
         Target.value = ""
     End If

    End If

'=====================================================================
    Sheet18.Protect Password:=getmsgstate


endline:
Application.EnableEvents = True

End Sub

Function ValidateDateDep_IT() As Boolean

    Dim rangecells As Range

    Dim i As Long

    Dim DateDep_IT

Dim mIntCells As Long

 Dim mIntCtr As Long

 mIntCells = Sheet18.Range("TaxP.DateDep").count

    

    ValidateDateDep_IT = True

    Dim end_IT

    end_IT = 0

    Set rangecells = Sheet18.Range("TaxP.DateDep").Cells

 For mIntCtr = 1 To mIntCells

     If Not rangecells.item(mIntCtr).value = "" Then

         end_IT = end_IT + 1

     End If

 Next

    ReDim DateDep_IT(end_IT)

    

    For i = 1 To end_IT

        DateDep_IT(i) = rangecells.item(i).value

        If Not chkCompulsory(DateDep_IT(i)) Then

            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory")

            ValidateDateDep_IT = False

            Exit Function

        End If

        

        If Not CheckDateddmmyyyy(DateDep_IT(i)) Then

            ValidateDateDep_IT = False

'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")

             fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")


            Exit Function

        Else

            DateDep_IT(i) = Dformat(DateDep_IT(i), "")

            'If Not ChkMinInclusiveDate(DateDep_IT(i), "2021-04-01") Then dpk1201
'            If Not ChkMinInclusiveDate(DateDep_IT(i), "2022-04-01") Then
            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then


                'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  should not be less than 01/04/2021") dpk1201
'                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  should not be less than 01/04/2022")
                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  should not be less than 01/04/2023")


                ValidateDateDep_IT = False

                Exit Function

            End If
            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
            
''                If ChkMinInclusiveDate(DateDep_IT(i), "2024-01-01") Then
''
''                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
'Ayush_UR
'Ayush_25-26
                If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
    
                    fmsgboxsmall ("Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.")
    
                End If
            End If



'            If DateDep_IT(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgboxsmall "* Date of Deposit at Sr. No  " & i & " in Sheet IT can not be more than system date in Income Details"
'
'                            rangecells.item(i).value = ""
'
'                            ValidateDateDep_IT = False
'
'                Exit Function
'
'            End If
'
        End If

    Next

End Function

 
Sub ttt()
Application.EnableEvents = True

End Sub
