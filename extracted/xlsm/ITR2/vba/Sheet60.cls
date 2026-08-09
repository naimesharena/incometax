Attribute VB_Name = "Sheet60"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Public lockcount_80EEA, lockcount_80EE As Variant
Public lock_80EEA_flag, lock_80EE_flag As Boolean
Public lock_80EEA_flag_val, lock_80EE_flag_val As Variant


 
Private Sub Worksheet_Activate()
Sheet60.Activate
Sheet60.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Calculate()

End Sub


'Shrutika 14/05/2025_80E

Function ValidateDate_80E(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80E
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet60.Range("loanDate.80E").count
    ValidateDate_80E = True

    Dim end_80E
    end_80E = 0
    Set rangecells = Sheet60.Range("loanDate.80E").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80E = end_80E + 1
        End If
    Next

    ReDim Date_80E(end_80E)

    For i = 1 To end_80E
        Date_80E(i) = rangecells.item(i).value
        If Not chkCompulsory(Date_80E(i)) Then
            fmsgboxsmall ("* ""Please provide Date of sanction of loan in schedule 80E"" at Sr. No  " & i & "")
            ValidateDate_80E = False
            Exit Function
        End If

        If Not CheckDateddmmyyyy(Date_80E(i)) Then
            ValidateDate_80E = False
              fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")

            Target.value = ""
            Exit Function
        Else
            Date_80E(i) = Dformat(Date_80E(i), "")
            
            If Not ChkMaxDate_24b(Trim(Date_80E(i)), "31-03-2025") Then
                fmsgboxsmall (" Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80E = False
                 

          Exit Function
        End If
End If
Next
End Function
'---------------------------------

'Shrutika 19/05/2025_80EE


Function ValidateDate_80EE(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80EE
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet60.Range("loanDate.80EE").count
    ValidateDate_80EE = True

    Dim end_80EE
    end_80EE = 0
    Set rangecells = Sheet60.Range("loanDate.80EE").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80EE = end_80EE + 1
        End If
    Next

    ReDim Date_80EE(end_80EE)

    For i = 1 To end_80EE
        Date_80EE(i) = rangecells.item(i).value
        If Not chkCompulsory(Date_80EE(i)) Then
            fmsgboxsmall ("Please provide ""Date of sanction of loan"" in schedule 80EE at Sr. No  " & i & "")
            ValidateDate_80EE = False
            Exit Function
        End If

        If Not CheckDateddmmyyyy(Date_80EE(i)) Then
            ValidateDate_80EE = False
               fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_80EE(i) = Dformat(Date_80EE(i), "")
            
            If Not ChkMaxDate_24b(Trim(Date_80EE(i)), "31-03-2025") Then
                fmsgboxsmall ("* Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80EE = False
                 

          Exit Function
        End If
        
        End If
    Next

End Function
        
        
'-------------------------------

'Shrutika_19/05/2025_80EEA

Function ValidateDate_80EEA(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80EEA
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet60.Range("loanDate.80EEA").count
    ValidateDate_80EEA = True

    Dim end_80EEA
    end_80EEA = 0
    Set rangecells = Sheet60.Range("loanDate.80EEA").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80EEA = end_80EEA + 1
        End If
    Next

    ReDim Date_80EEA(end_80EEA)

    For i = 1 To end_80EEA
        Date_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(Date_80EEA(i)) Then
            fmsgboxsmall ("* Please provide ""Date of sanction of loan"" in schedule 80EEA at Sr. No  " & i & "")
            ValidateDate_80EEA = False
            Exit Function
        End If

        If Not CheckDateddmmyyyy(Date_80EEA(i)) Then
            ValidateDate_80EEA = False
                fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
  
            Target.value = ""
            Exit Function
        Else
            Date_80EEA(i) = Dformat(Date_80EEA(i), "")
            
            If Not ChkMaxDate_24b(Trim(Date_80EEA(i)), "31-03-2025") Then
                fmsgboxsmall (" Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80EEA = False
          Exit Function
        End If
            
        End If
    Next

End Function
'---------------------------------------



'Shrutik_19/05/2025_80EEB

Function ValidateDate_80EEB(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80EEB
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet60.Range("loanDate.80EEB").count
    ValidateDate_80EEB = True

    Dim end_80EEB
    end_80EEB = 0
    Set rangecells = Sheet60.Range("loanDate.80EEB").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80EEB = end_80EEB + 1
        End If
    Next

    ReDim Date_80EEB(end_80EEB)

    For i = 1 To end_80EEB
        Date_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(Date_80EEB(i)) Then
            fmsgboxsmall ("* ""Please provide Date of sanction of loan in schedule 80EEB"" at Sr. No  " & i & "")
            ValidateDate_80EEB = False
            Exit Function
        End If

        If Not CheckDateddmmyyyy(Date_80EEB(i)) Then
            ValidateDate_80EEB = False
                fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")

            Target.value = ""
            Exit Function
        Else
            Date_80EEB(i) = Dformat(Date_80EEB(i), "")
            
            If Not ChkMaxDate_24b(Trim(Date_80EEB(i)), "31-03-2025") Then
                fmsgboxsmall ("* Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80EEB = False
                 

          Exit Function
        End If
            
        End If
    Next

End Function


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim TargtVal, targetadd As Variant
Dim IffCode As Range
Dim IfsCode As Range
Dim IfsCode2 As Range

Application.EnableEvents = False
Application.ScreenUpdating = False

'--------------------------------------
'Sheet60.Unprotect Password:=sPassword
'Target.Columns.AutoFit
'--------------------------------------

CalculateTaxCheckFlag = False
Sheet60.Unprotect Password:=getmsgstate



Dim lock80EE_chk, lock80EEA_chk As Boolean


If Not Application.Intersect(Sheet60.Range("Stampduty.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("LoanfrmBankOrInstitute.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("bankName.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanAccNum.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanDate.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanAmt.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanOutstanding.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("Intrst.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Then


     lock80EE_80EEAcount


If lockcount_80EEA > 0 Then
  lockcount_80EE = 0
  lock80EE_80EEA_chk
Else
  lock80EE_80EEA_chk
End If
 

End If


If Not Application.Intersect(Sheet60.Range("LoanfrmBankOrInstitute.80EE"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("bankName.80EE"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanAccNum.80EE"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanDate.80EE"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanAmt.80EE"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("loanOutstanding.80EE"), Sheet60.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet60.Range("Intrst.80EE"), Sheet60.Range(Target.Address)) Is Nothing Then

 

     lock80EE_80EEAcount

'80EE
If lockcount_80EE > 0 Then
  lockcount_80EEA = 0
  lock80EE_80EEA_chk
  
  Else
  lock80EE_80EEA_chk
End If
 

End If

'----------------------------

If Not Application.Intersect(Sheet60.Range("loanDate.80E"), Sheet60.Range(Target.Address)) Is Nothing Then
        If Not (ValidateDate_80E(Target)) Then
         End If
           GoTo endline
End If

If Not Application.Intersect(Sheet60.Range("loanDate.80EE"), Sheet60.Range(Target.Address)) Is Nothing Then
        If Not (ValidateDate_80EE(Target)) Then
         End If
           GoTo endline
End If

If Not Application.Intersect(Sheet60.Range("loanDate.80EEA"), Sheet60.Range(Target.Address)) Is Nothing Then
        If Not (ValidateDate_80EEA(Target)) Then
         End If
           GoTo endline
End If


If Not Application.Intersect(Sheet60.Range("loanDate.80EEB"), Sheet60.Range(Target.Address)) Is Nothing Then
        If Not (ValidateDate_80EEB(Target)) Then
         End If
           GoTo endline
End If
'--------------------------------------

endline:

Application.EnableEvents = True
Sheet60.Protect Password:=sPassword
Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Function lock80EE_80EEAcount()

Dim ResidentialHP_80E, Stampduty_80EEA As Variant




lockcount_80EE = 0
lockcount_80EEA = 0


'------------------------------------------------------------------------
'80EE

    Dim mIntCells As Long
    Dim rangecells, rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8, rangecells9 As Range
     Dim rangecells11, rangecells22, rangecells33, rangecells44, rangecells55, rangecells66, rangecells77, rangecells88, rangecells99 As Range
    
    Dim mIntCtr As Long
    Dim ccount, ccount1, ccount2, ccount3, ccount4, ccount5, ccount6, ccount7, ccount8 As Long
     Dim ccount9, ccount33, ccount44, ccount55, ccount66, ccount77, ccount88 As Long
    
    mIntCells = Range("LoanfrmBankOrInstitute.80EE").count
    Set rangecells9 = Range("LoanfrmBankOrInstitute.80EE").Cells
    Set rangecells33 = Range("bankName.80EE").Cells
    Set rangecells44 = Range("loanAccNum.80EE").Cells
    Set rangecells55 = Range("loanDate.80EE").Cells
    Set rangecells66 = Range("loanAmt.80EE").Cells
    Set rangecells77 = Range("loanOutstanding.80EE").Cells
    Set rangecells88 = Range("Intrst.80EE").Cells
    
    For mIntCtr = 1 To mIntCells
            If Not (rangecells9.item(mIntCtr).value = "" Or rangecells9.item(mIntCtr).value = "(Select)") Then
                ccount9 = ccount9 + 1
            End If
 
            If Not (rangecells33.item(mIntCtr).value = "") Then
                ccount33 = ccount33 + 1
            End If
            If Not (rangecells44.item(mIntCtr).value = "") Then
                ccount44 = ccount44 + 1
            End If
            If Not (rangecells55.item(mIntCtr).value = "") Then
                ccount55 = ccount55 + 1
            End If
            If Not (rangecells66.item(mIntCtr).value = "") Then
                ccount66 = ccount66 + 1
            End If
            If Not (rangecells77.item(mIntCtr).value = "") Then
                ccount77 = ccount77 + 1
            End If
            If Not (rangecells88.item(mIntCtr).value = "") Then
                ccount88 = ccount88 + 1
            End If
            
            
            
    Next
    

    lockcount_80EE = Application.WorksheetFunction.Max(0, ccount9, ccount33, ccount44, ccount55, ccount66, ccount77, ccount88)
mIntCells = Range("LoanfrmBankOrInstitute.80EEA").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
    Set rangecells3 = Range("bankName.80EEA").Cells
    Set rangecells4 = Range("loanAccNum.80EEA").Cells
    Set rangecells5 = Range("loanDate.80EEA").Cells
    Set rangecells6 = Range("loanAmt.80EEA").Cells
    Set rangecells7 = Range("loanOutstanding.80EEA").Cells
    Set rangecells8 = Range("Intrst.80EEA").Cells
    
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
            If Not (rangecells3.item(mIntCtr).value = "") Then
                ccount3 = ccount3 + 1
            End If
            If Not (rangecells4.item(mIntCtr).value = "") Then
                ccount4 = ccount4 + 1
            End If
            If Not (rangecells5.item(mIntCtr).value = "") Then
                ccount5 = ccount5 + 1
            End If
            If Not (rangecells6.item(mIntCtr).value = "") Then
                ccount6 = ccount6 + 1
            End If
            If Not (rangecells7.item(mIntCtr).value = "") Then
                ccount7 = ccount7 + 1
            End If
            If Not (rangecells8.item(mIntCtr).value = "") Then
                ccount8 = ccount8 + 1
            End If
            
            
            
    Next
    
If Sheet60.Range("Stampduty.80EEA").value > 0 Then
Stampduty_80EEA = 1
Else
Stampduty_80EEA = 0
End If

    lockcount_80EEA = WorksheetFunction.Max(0, Stampduty_80EEA, ccount, ccount3, ccount4, ccount5, ccount6, ccount7, ccount8)
'-------------------------------------------------------------------------
End Function



Function lock80EE_80EEA_chk()

 lock_80EEA_flag_val = 0
 lock_80EE_flag_val = 0
lock_80EEA_flag = True
lock_80EE_flag = True
            
        If lockcount_80EEA > 0 Then

            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").ClearContents
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").Locked = True
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("bankName.80EE").ClearContents
            Sheet60.Range("bankName.80EE").Locked = True
            Sheet60.Range("bankName.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAccNum.80EE").ClearContents
            Sheet60.Range("loanAccNum.80EE").Locked = True
            Sheet60.Range("loanAccNum.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanDate.80EE").ClearContents
            Sheet60.Range("loanDate.80EE").Locked = True
            Sheet60.Range("loanDate.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAmt.80EE").ClearContents
            Sheet60.Range("loanAmt.80EE").Locked = True
            Sheet60.Range("loanAmt.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanOutstanding.80EE").ClearContents
            Sheet60.Range("loanOutstanding.80EE").Locked = True
            Sheet60.Range("loanOutstanding.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("Intrst.80EE").ClearContents
            Sheet60.Range("Intrst.80EE").Locked = True
            Sheet60.Range("Intrst.80EE").Interior.Color = "&HD8D8D8"
            
            lock_80EE_flag = False
            lock_80EE_flag_val = 1

          ElseIf lockcount_80EE > 0 Then
            '80EEA
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("Stampduty.80EEA").ClearContents
            Sheet60.Range("Stampduty.80EEA").Locked = True
            Sheet60.Range("Stampduty.80EEA").Interior.Color = "&HD8D8D8"
            
            
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Locked = True
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Interior.Color = "&HD8D8D8"
            
            
           Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("bankName.80EEA").ClearContents
            Sheet60.Range("bankName.80EEA").Locked = True
            Sheet60.Range("bankName.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAccNum.80EEA").ClearContents
            Sheet60.Range("loanAccNum.80EEA").Locked = True
            Sheet60.Range("loanAccNum.80EEA").Interior.Color = "&HD8D8D8"
            
           Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanDate.80EEA").ClearContents
            Sheet60.Range("loanDate.80EEA").Locked = True
            Sheet60.Range("loanDate.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAmt.80EEA").ClearContents
            Sheet60.Range("loanAmt.80EEA").Locked = True
            Sheet60.Range("loanAmt.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanOutstanding.80EEA").ClearContents
            Sheet60.Range("loanOutstanding.80EEA").Locked = True
            Sheet60.Range("loanOutstanding.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("Intrst.80EEA").ClearContents
            Sheet60.Range("Intrst.80EEA").Locked = True
            Sheet60.Range("Intrst.80EEA").Interior.Color = "&HD8D8D8"
            
             lock_80EEA_flag = False
             lock_80EEA_flag_val = 1
            
            
       Else
    

            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").ClearContents
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").Locked = False
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").value = "(Select)"   'Ankita_30/04/2025
            Sheet60.Range("LoanfrmBankOrInstitute.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("bankName.80EE").ClearContents
            Sheet60.Range("bankName.80EE").Locked = False
            Sheet60.Range("bankName.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAccNum.80EE").ClearContents
            Sheet60.Range("loanAccNum.80EE").Locked = False
            Sheet60.Range("loanAccNum.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanDate.80EE").ClearContents
            Sheet60.Range("loanDate.80EE").Locked = False
            Sheet60.Range("loanDate.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAmt.80EE").ClearContents
            Sheet60.Range("loanAmt.80EE").Locked = False
            Sheet60.Range("loanAmt.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanOutstanding.80EE").ClearContents
            Sheet60.Range("loanOutstanding.80EE").Locked = False
            Sheet60.Range("loanOutstanding.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("Intrst.80EE").ClearContents
            Sheet60.Range("Intrst.80EE").Locked = False
            Sheet60.Range("Intrst.80EE").Interior.Color = "&HCCFFCC"

          
           Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("Stampduty.80EEA").ClearContents
            Sheet60.Range("Stampduty.80EEA").Locked = False
            Sheet60.Range("Stampduty.80EEA").Interior.Color = "&HCCFFCC"
            
            
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Locked = False
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").value = "(Select)"   'Ankita_30/04/2025
            Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Interior.Color = "&HCCFFCC"
        
            
            
           Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("bankName.80EEA").ClearContents
            Sheet60.Range("bankName.80EEA").Locked = False
            Sheet60.Range("bankName.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAccNum.80EEA").ClearContents
            Sheet60.Range("loanAccNum.80EEA").Locked = False
            Sheet60.Range("loanAccNum.80EEA").Interior.Color = "&HCCFFCC"
            
           Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanDate.80EEA").ClearContents
            Sheet60.Range("loanDate.80EEA").Locked = False
            Sheet60.Range("loanDate.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanAmt.80EEA").ClearContents
            Sheet60.Range("loanAmt.80EEA").Locked = False
            Sheet60.Range("loanAmt.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("loanOutstanding.80EEA").ClearContents
            Sheet60.Range("loanOutstanding.80EEA").Locked = False
            Sheet60.Range("loanOutstanding.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet60.Unprotect Password:=getmsgstate
            Sheet60.Range("Intrst.80EEA").ClearContents
            Sheet60.Range("Intrst.80EEA").Locked = False
            Sheet60.Range("Intrst.80EEA").Interior.Color = "&HCCFFCC"
            
            lock_80EEA_flag = True
            lock_80EE_flag = True
            lock_80EEA_flag_val = 0
            lock_80EE_flag_val = 0
           End If

End Function
Public Function ChkMaxDate_24b(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_24b = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

        If Year > 2025 Then
            ChkMaxDate_24b = False
            Exit Function
        Else

                If Year = 2025 Then
                If month > 4 Then
                    ChkMaxDate_24b = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_24b = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_24b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function






