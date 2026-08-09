Attribute VB_Name = "Sheet46"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_Deactivate()
sPassword = getmsgstate
Sheet46.Protect Password:=sPassword
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False

If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If

If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing"), Sheet46.Range(Target.Address)) Is Nothing Then
Sheet46.Unprotect Password:=getmsgstate
'AY_2023-24 OLD  'Malli
'        If Target.value <> "" And ChkMaxDOBDateCease(Sheet46.Range(Target.Address).value, "31/03/2022") Then
'        fmsgbox ("The date should not be before 01/04/2022 for A.Y. 2023-24 In schedule ESOP")
'AY_2024-25 NEW  'Malli
  'If Target.value <> "" And ChkMaxDOBDateCease(Sheet46.Range(Target.Address).value, "31/03/2022") Then

' If target.value <> "" And Dformat(Sheet46.Range(target.Address).value, "") <= Dformat("31/03/2020", "") Then

'Added by Riyaz on 16/10/2025 for SIT-102023
        If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not CheckDate2425_ESOP(Trim(Target.value)) Then
'            MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
            Target.ClearContents
        End If
        End If
'End

 If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2020", "") Then

       'fmsgbox ("""The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP""")
        fmsgbox ("""The date should not be before 01/04/2020""")

        Sheet46.Range(Target.Address).value = ""
        GoTo endline
        End If
Sheet46.Protect Password:=getmsgstate

Sheet46.Unprotect Password:=getmsgstate
        Dim DateOfCeasing As Variant
'        VerificationDate = Sheet52.Range("sheet9.Date").value
'        DateOfCeasing = Sheet46.Range("ESOP.DateOfCeasing").value
'        If Dformat(DateOfCeasing, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
'            fmsgbox ("""Date of Ceasing cannot be after System Date in Schedule ESOP""")
'            Sheet46.Range(Target.Address).value = ""
'            GoTo endline
'        End If
        
'        If Not ChkMaxDOBDate(Sheet46.Range(Target.Address).value, "31/03/2023") Then
'Malli
'        If Dformat(Sheet46.Range(Target.Address).value, "") > Dformat("31/03/2023", "") Then
'                fmsgbox ("Date of Ceasing should not be after 31/03/2023 for AY 2023-24 In schedule ESOP ")
'If Dformat(Sheet46.Range(Target.Address).value, "") > Dformat("31/03/2024", "") Then
'                fmsgbox ("""Date of Ceasing should not be after 31/03/2024 for AY 2024-25 In schedule ESOP""")
'                Target.value = ""
'                GoTo endline
'        End If
        
Sheet46.Protect Password:=getmsgstate

End If


'SIT-73599 Newly added by sai on 29/04/2024
'===================================================================
        If Not Intersect(Range("ESOP.DateOfCeasing"), Range(Target.Address)) Is Nothing Then
                If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
                    fmsgbox (" Date cannot exceed User system date")
                     Target.value = ""
                End If
        End If
'====================================================================

'Change Satyaprakash ======
If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee2"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If

'2
If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing2"), Sheet46.Range(Target.Address)) Is Nothing Then
Sheet46.Unprotect Password:=getmsgstate
'Change-28.01.2023.101.ESOP.7B
'        If ChkMaxDOBDateDef(Sheet46.Range(Target.Address).value, "31/03/2022") Then
'        fmsgbox ("The date should not be before 01/04/2022 for A.Y. 2023-24 In schedule ESOP")
'        Sheet46.Range(Target.Address).value = ""
'        GoTo endline
'        End If

    'AY_2024_25 'Malli
'    If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2022", "") Then
'        MsgBox ("The date should not be before 01/04/2022 for A.Y. 2023-24 In schedule ESOP")
' If target.value <> "" And Dformat(Sheet46.Range(target.Address).value, "") <= Dformat("31/03/2021", "") Then
 
'Added by Riyaz on 16/10/2025 for SIT-102023
        If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not CheckDate2425_ESOP(Trim(Target.value)) Then
'            MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
            Target.ClearContents
        End If
        End If
'End

  If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2021", "") Then
        MsgBox ("""The date should not be before 01/04/2021""")
        Sheet46.Range(Target.Address).value = ""
        GoTo endline
  End If
  
'---
Sheet46.Protect Password:=getmsgstate

Sheet46.Unprotect Password:=getmsgstate
        Dim DateOfCeasing1 As Variant
'        VerificationDate = Sheet52.Range("sheet9.Date").value
'        DateOfCeasing1 = Sheet46.Range("ESOP.DateOfCeasing2").value
'        If Dformat(DateOfCeasing1, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
'            fmsgbox ("""Date of Ceasing cannot be after System Date in Schedule ESOP""")
'            Sheet46.Range(Target.Address).value = ""
'            GoTo endline
'        End If
'
'Change-28.01.2023.101.ESOP.7A
'        If Not ChkMaxDOBDate(Sheet46.Range(Target.Address).value, "31/03/2023") Then
'                fmsgbox ("Date of Ceasing should not be after 31/03/2023 for AY 2023-24 In schedule ESOP ")
'                Target.value = ""
'                GoTo endline
'        End If

        'AY_2024_25 'Malli
'        If Dformat(Sheet46.Range(Target.Address).value, "") > Dformat("31/03/2024", "") Then
'            MsgBox ("""Date of Ceasing should not be after 31/03/2024 for AY 2024-25 In schedule ESOP""")
'            Target.value = ""
'            GoTo endline
'        End If
''---
'2023-24 - should not be before 01/04/2023
'2022-23 - should not be before 01/04/2022
'2021-22 - should not be before 01/04/2021

'2021-22> Error Msg >AY 2022-23
'
'2022-23>Error Msg > AY 2023-24
'
'2023-24>Error Msg > AY 2024-25

Sheet46.Protect Password:=getmsgstate

End If


'SIT-73599 Newly added by sai on 29/04/2024
'=========================================================
        If Not Intersect(Range("ESOP.DateOfCeasing2"), Range(Target.Address)) Is Nothing Then
          
                If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
                    fmsgbox (" ""Date cannot exceed User system date""")
                     Target.value = ""
        
                End If
                
        End If
'==========================================================


'Malli=====================================================
If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee3"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If


'If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing3"), Sheet46.Range(Target.Address)) Is Nothing Then
'Sheet46.Unprotect Password:=getmsgstate
'
' If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2023", "") Then
'        MsgBox ("""The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP""")
'        Sheet46.Range(Target.Address).value = ""
'        GoTo endline
'End If
'---

'3

If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing3"), Sheet46.Range(Target.Address)) Is Nothing Then
Sheet46.Unprotect Password:=getmsgstate
 
 
'Newly modified by Sai  for SIT-69593
' If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2023", "") Then
'If target.value <> "" And Dformat(Sheet46.Range(target.Address).value, "") <= Dformat("31/03/2022", "") Then

'Added by Riyaz on 16/10/2025 for SIT-102023
        If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not CheckDate2425_ESOP(Trim(Target.value)) Then
'            MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
            Target.ClearContents
        End If
        End If
'End

If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2022", "") Then

'        MsgBox ("""The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP""")
         MsgBox ("""The date should not be before 01/04/2022""")
        Sheet46.Range(Target.Address).value = ""
        GoTo endline
End If
Sheet46.Protect Password:=getmsgstate

Sheet46.Unprotect Password:=getmsgstate
        Dim DateOfCeasing3 As Variant
'        VerificationDate = Sheet52.Range("sheet9.Date").value
'        DateOfCeasing3 = Sheet46.Range("ESOP.DateOfCeasing3").value
'        If Dformat(DateOfCeasing3, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
'            fmsgbox ("""Date of Ceasing cannot be after System Date in Schedule ESOP""")
'            Sheet46.Range(Target.Address).value = ""
'            GoTo endline
'        End If
        
 
        'AY_2024_25 'Malli
'        If Dformat(Sheet46.Range(Target.Address).value, "") > Dformat("31/03/2024", "") Then
'            MsgBox ("""Date of Ceasing should not be after 31/03/2024 for AY 2024-25 In schedule ESOP""")
'            Target.value = ""
'            GoTo endline
'        End If
 
Sheet46.Protect Password:=getmsgstate

End If

     'SIT-73599 Newly added by sai on 29/04/2024
'=========================================================
        If Not Intersect(Range("ESOP.DateOfCeasing3"), Range(Target.Address)) Is Nothing Then
          
                If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
                    fmsgbox (" ""Date cannot exceed User system date""")
                     Target.value = ""
        
                End If
                
        End If
'====================================================================================================

'Newly added by sai on 21/02/2025 for AY 2025-26

  If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee4"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate

End If

'4

If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing4"), Sheet46.Range(Target.Address)) Is Nothing Then
    Sheet46.Unprotect Password:=getmsgstate


'Added by Riyaz on 16/10/2025 for SIT-102023
        If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not CheckDate2425_ESOP(Trim(Target.value)) Then
'            MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
            Target.ClearContents
        End If
        End If
'End

      If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2023", "") Then
             MsgBox ("""The date should not be before 01/04/2023""")
             Sheet46.Range(Target.Address).value = ""
             GoTo endline
      End If
     Sheet46.Protect Password:=getmsgstate

End If


    If Not Intersect(Range("ESOP.DateOfCeasing4"), Range(Target.Address)) Is Nothing Then
            If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
                fmsgbox (" ""Date cannot exceed User system date""")
                 Target.value = ""
            End If
    End If

'======================================================================================================
'End Change Satyaprakash===
'Change-ANK


''''''''''''''''''''''''''Added by Aavula Naresh for 2026-27'''''''

If Not Application.Intersect(Sheet46.Range("ESOP.CeasedEmployee5"), Sheet46.Range(Target.Address)) Is Nothing Then
   Sheet46.Unprotect Password:=getmsgstate
   If Sheet46.Range(Target.Address).value = "No" Then
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = True
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
   Else
        Sheet46.Range(Replace(Target.Address, "I", "J")).ClearContents
        Sheet46.Range(Replace(Target.Address, "I", "J")).Locked = False
        Sheet46.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
        
   End If
   Sheet46.Protect Password:=getmsgstate
End If


If Not Application.Intersect(Sheet46.Range("ESOP.DateOfCeasing5"), Sheet46.Range(Target.Address)) Is Nothing Then
    Sheet46.Unprotect Password:=getmsgstate


        If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not CheckDate2425_ESOP(Trim(Target.value)) Then
'            MsgBox """ Date of Ceasing is invalid in schedule ESOP. Please enter a valid Date in dd/mm/yyyy format.""" & Chr(13), , "Error(s)"
            Target.ClearContents
        End If
        End If

      If Target.value <> "" And Dformat(Sheet46.Range(Target.Address).value, "") <= Dformat("31/03/2024", "") Then
             MsgBox ("""The date should not be before 01/04/2024""")
             Sheet46.Range(Target.Address).value = ""
             GoTo endline
      End If
     Sheet46.Protect Password:=getmsgstate

End If

    If Not Intersect(Range("ESOP.DateOfCeasing5"), Range(Target.Address)) Is Nothing Then
            If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
                fmsgbox (" ""Date cannot exceed User system date""")
                 Target.value = ""
            End If
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Added by Aavula
Dim i As Long
If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then
    i = Target.row
    
      If Not Target.value = "" Then
           If Not CheckDate2425_ESOP(Target.value) Then
              Target.value = ""
              Target.Select
           End If

        If Sheet46.Cells(i, Sheet46.Range("ESOPCurrAY").Column).value = "2021-22" Then
            If Dformat(Target.value, "") < "2020-04-01" Then
                fmsgbox ("The date should not be before 01/04/2020")
                Target.value = ""
                GoTo endline
                End If
        End If

        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
            fmsgbox (" Date cannot be after System Date in Schedule ESOP")
             Target.value = ""
            GoTo endline
        End If
End If
End If

'Added by Aavula
If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then
    i = Target.row
    
      If Not Target.value = "" Then
           If Not CheckDate2425_ESOP(Target.value) Then
              Target.value = ""
              Target.Select
           End If

        If Sheet46.Cells(i, Sheet46.Range("ESOPCurrAY").Column).value = "2022-23" Then
            If Dformat(Target.value, "") < "2021-04-01" Then
                fmsgbox ("The date should not be before 01/04/2021")
                Target.value = ""
                GoTo endline
                End If
        End If

        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
            fmsgbox (" Date cannot be after System Date in Schedule ESOP")
             Target.value = ""
            GoTo endline
        End If
End If

'        If Sheet46.Cells(i, Sheet46.Range("ESOPCurrAY").Column).value = "2021-22" Then
'            If Dformat(Target.value, "") > "2021-03-31" Then
'                fmsgbox ("Date should not be after 31/03/2021 for AY 2021-22 In schedule ESOP ")
'                Target.value = ""
'                GoTo endline
'                End If
'        End If

End If

''Added by Aavula

If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then
    i = Target.row

  If Not Target.value = "" Then
           If Not CheckDate2425_ESOP(Target.value) Then
              Target.value = ""
              Target.Select
           End If
           
           
        If Sheet46.Cells(i, Sheet46.Range("ESOPCurrAY").Column).value = "2023-24" Then
            If Dformat(Target.value, "") < "2022-04-01" Then
                fmsgbox ("The date should not be before 01/04/2022")
                Target.value = ""
                GoTo endline
                End If
        End If
'
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
            fmsgbox (" Date cannot be after System Date in Schedule ESOP")
             Target.value = ""
            GoTo endline
        End If
End If
End If

'''''''''''''''''''''''''

'added by sai on 21/02/2025 for AY 2025-26===========================================================================================

  If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then
    i = Target.row
    
           If Not Target.value = "" Then
           If Not CheckDate2425_ESOP(Target.value) Then
              Target.value = ""
              Target.Select
           End If

        If Sheet46.Cells(i, Sheet46.Range("ESOPCurrAY").Column).value = "2024-25" Then
            If Dformat(Target.value, "") < "2023-04-01" Then
                fmsgbox ("The date should not be before 01/04/2023")
                Target.value = ""
                GoTo endline
                End If
        End If
        End If
'
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
            fmsgbox (" Date cannot be after System Date in Schedule ESOP")
             Target.value = ""
            GoTo endline
        End If
End If

'====================================================================================================
'Added by Aavula Naresh for AY 2026-27

If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then
    i = Target.row
    
        If Not Target.value = "" Then
           If Not CheckDate2425_ESOP(Target.value) Then
              Target.value = ""
              Target.Select
           End If

           If Sheet46.Cells(i, Sheet46.Range("ESOPCurrAY").Column).value = "2025-26" Then
               If Dformat(Target.value, "") < "2024-04-01" Then
                    fmsgbox ("The date should not be before 01/04/2024")
                    Target.value = ""
                    GoTo endline
               End If
            End If
        End If
        
        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(Sheet52.Range("sheet9.Date").value, "yyyy-mm-dd") Then
            fmsgbox (" Date cannot be after System Date in Schedule ESOP")
             Target.value = ""
            GoTo endline
        End If
End If



'    For i = 1 To Sheet46.Range("ESOP.Date").count
'        If Not Intersect(Range("ESOP.Date"), Range(Target.Address)) Is Nothing Then

            'Change-24.02.2023.103.ESOP
            'If ChkMaxDOBDateDef(Target.value, "31/03/2021") Then
            'Change.14.03.2023.102.IDS.10
            
            'AY_2024_25 Change 'Malli
'            If Target.value <> "" And Dformat(Target.value, "") < Dformat("31/03/2022", "") Then
'                'fmsgbox ("The date should not be before 01/04/2021 for A.Y. 2023-24 In schedule ESOP")
'                fmsgbox ("The date should not be before 01/04/2022 for A.Y. 2023-24 In schedule ESOP")

'Comment by aavula
'                If Target.value <> "" And Dformat(Target.value, "") < Dformat("31/03/2023", "") Then
'                fmsgbox ("""The date should not be before 01/04/2023 for A.Y. 2024-25 In schedule ESOP""")
'                Target.value = ""
'                GoTo endline
'            End If
            'End Change.102.IDS.10

'Comment by Aavula
'        VerificationDate = Sheet52.Range("sheet9.Date").value
'
'        If Dformat(Target.value, "yyyy-mm-dd") > Dformat(VerificationDate, "yyyy-mm-dd") Then
'        fmsgbox ("""Date cannot be after System Date in Schedule ESOP""")
'         Target.value = ""
'        GoTo endline
'        End If
'==================
        'Change-24.02.2023.103.ESOP
        'If (Not ChkMaxDOBDate(Target.value, "31/03/2023")) Then
        
        'AY_2024_25 Change 'Malli
'        If Dformat(Target.value, "") > Dformat("31/03/2023", "") Then
'                fmsgbox ("Date should not be after 31/03/2023 for AY 2023-24 In schedule ESOP ")

'Comment by Aavula
'            If Dformat(Target.value, "") > Dformat("31/03/2024", "") Then
'                fmsgbox ("""Date should not be after 31/03/2024 for AY 2024-25 In schedule ESOP""")
'                Target.value = ""
'                GoTo endline
'        End If

        'End If
'
'Next

'{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{commented by sai on 21/02/2025 but again need to check
'If (Not Intersect(Sheet46.Range("ESOP.SecType"), Range(target.Address)) Is Nothing) Or _
'    (Not Intersect(Sheet46.Range("ESOP.SecType2"), Range(target.Address)) Is Nothing) Then
'Sheet46.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'Dim listString As String
'    listString = "(Select),"
'If Sheet46.Range("ESOP.SecType").value <> "Not sold" Then
'    listString = listString + "2021-22,"
'End If
'
'If Sheet46.Range("ESOP.SecType2").value <> "Not sold" Then
'    listString = listString + "2022-23,"
'End If
'
''added by sai on 21/02/2025========================================
'If Sheet46.Range("ESOP.SecType").value <> "Not sold" Then
'    listString = listString + "2023-24,"
'End If
'
'If Sheet46.Range("ESOP.SecType2").value <> "Not sold" Then
'    listString = listString + "2024-25,"
'End If
''=================================================================
'
'Sheet46.Range("ESOPCurrAY").ClearContents
'Sheet46.Range("ESOP.Date").ClearContents
'Sheet46.Range("ESOP.AttributedTax").ClearContents
'
'Sheet46.Range("ESOPCurrAY").Validation.Delete
'Sheet46.Range("ESOPCurrAY").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listString
'
'Sheet46.Protect Password:=getmsgstate
'Application.EnableEvents = True
'
'
'End If
'                                    }}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}



'================================================
'Newly added by Malli
If Not Intersect(Sheet46.Range("ESOP.PAN"), Range(Target.Address)) Is Nothing Then

Dim TargetPAN As String
               TargetPAN = ""
               TargetPAN = Target.name.name
                Sheet46.Range("ESOP.PAN").value = UCase(Range(TargetPAN).value)
                If Not Part_A_General.CheckPAN(UCase(Trim(Range(TargetPAN).value))) Then
                     fmsgboxStatus """Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule ESOP"""
                     Sheet46.Range(TargetPAN).Select
                     Sheet46.Range(TargetPAN).value = ""
                 
                End If
End If





'Updated by Jyothi malli


If Not Intersect(Sheet46.Range("ESOP.DPIIT"), Range(Target.Address)) Is Nothing Then
   If Trim(Range("ESOP.DPIIT").value) <> "" Then
   Dim TargetDPIIT
        TargetDPIIT = ""
        TargetDPIIT = Target.name.name
        Sheet46.Range("ESOP.DPIIT").value = UCase(Range(TargetDPIIT).value)
        If Not Part_A_General.ChkDPIIT(UCase(Trim(Range(TargetDPIIT).value))) Then
             fmsgboxStatus """Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"""
             Sheet46.Range(TargetDPIIT).Select
             Sheet46.Range(TargetDPIIT).value = ""
        End If
        If Not checkallfieldspecialcharacter_80U_80DD(UCase(Trim(Range(TargetDPIIT).value))) Then
             fmsgboxStatus """Please provide valid DPIIT Registeration number of the employer which shall begin with DIPP followed by 3 to 5 numeric characters"""
             Sheet46.Range(TargetDPIIT).Select
             Sheet46.Range(TargetDPIIT).value = ""

        End If
   End If
End If
'----------------------------------


endline:
Application.EnableEvents = True
End Sub





Function CheckDate2425_ESOP(dob As Variant) As Boolean
On Error Resume Next
CheckDate2425_ESOP = True

    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDateVDA(dob) Then
            CheckDate2425_ESOP = False
            '01/10/2025
            'MsgBox ("Date of must be a valid dd/mm/yyyy format")
            MsgBox ("""Date must be a valid dd/mm/yyyy format"""), , "Error(s):"
            '-------------------------
            Exit Function
        End If
    End If
 End Function
