Attribute VB_Name = "Sheet22"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim validateType As String
Dim cell As Variant
'Malli----------------
            Dim PinCode As Range
            Dim DupFlag As Boolean
            Dim StateName
            Dim StateCode
            Dim state1 As Variant
'-----------------------------
   validateType = ""
   Application.EnableEvents = False

'----------------Unlock Password-------------------START---
'   sPassword = mdCommon.getmsgstate
'   ActiveSheet.Unprotect Password:=sPassword
    Sheet22.Unprotect Password:=getmsgstate
'----------------Unlock Password-------------------END-----

  'To add Income to House Property
 ' AddTotIncomeUnderHouseProperty

  'Set cell = Sheet22.Range("HP.Co.Pan").Cells


    Dim HpCount As Long

       For HpCount = 1 To Sheet22.Range("PropertySectionCount").value

       For Each cell In Sheet22.Range("HP.Co.Pan" & HpCount).Cells
        Sheet22.Range("" & cell.Address).value = UCase(Sheet22.Range("" & cell.Address).value)
       Next

       For Each cell In Sheet22.Range("HP.PANofTenant" & HpCount).Cells
        Sheet22.Range("" & cell.Address).value = UCase(Sheet22.Range("" & cell.Address).value)
       Next
       
       For Each cell In Sheet22.Range("HP.AadharofTenant" & HpCount).Cells
        Sheet22.Range("" & cell.Address).value = UCase(Sheet22.Range("" & cell.Address).value)
       Next


        For Each cell In Sheet22.Range("HP.TANofTenant" & HpCount).Cells
        Sheet22.Range("" & cell.Address).value = UCase(Sheet22.Range("" & cell.Address).value)
       Next


'
    If Not Application.Intersect(Range("HP.OwnerProperty" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
       '   If Target.name.name = ("HP.OwnerProperty" & HpCount) Then
           If Mid(Sheet22.Range("HP.OwnerProperty" & HpCount).value, 1, 1) = "O" Then
                Sheet22.Range("HP.OwnerPropertyDescription" & HpCount).Locked = False
                'Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).ClearContents
                Sheet22.Range("HP.OwnerPropertyDescription" & HpCount).Interior.Color = (&HCCFFCC)

           Else
                Sheet22.Range("HP.OwnerPropertyDescription" & HpCount).Locked = True
                Sheet22.Range("HP.OwnerPropertyDescription" & HpCount).ClearContents
                Sheet22.Range("HP.OwnerPropertyDescription" & HpCount).Interior.Color = (&HD8D8D8)


         End If
        End If
        
        
        If Not Intersect(Range("HP.ifLetOut" & HpCount), Range(Target.Address)) Is Nothing Then
           If Mid(Sheet22.Range("HP.ifLetOut" & HpCount).value, 1, 1) = "S" Or Mid(Sheet22.Range("HP.ifLetOut" & HpCount).value, 1, 1) = "D" Then
                
                Sheet22.Range("HP.NameofTenant" & HpCount).Locked = True
                Sheet22.Range("HP.NameofTenant" & HpCount).ClearContents
                Sheet22.Range("HP.NameofTenant" & HpCount).Interior.Color = (&HD8D8D8)

                Sheet22.Range("HP.PANofTenant" & HpCount).Locked = True
                Sheet22.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet22.Range("HP.PANofTenant" & HpCount).ClearContents
                
                Sheet22.Range("HP.AadharofTenant" & HpCount).Locked = True
                Sheet22.Range("HP.AadharofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet22.Range("HP.AadharofTenant" & HpCount).ClearContents
                
                Sheet22.Range("HP.TANofTenant" & HpCount).Locked = True
                Sheet22.Range("HP.TANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet22.Range("HP.TANofTenant" & HpCount).ClearContents
                If Mid(Sheet22.Range("HP.ifLetOut" & HpCount).value, 1, 1) = "S" Then
                    
                    Sheet22.Range("HP.AnnualLetableValue" & HpCount).Locked = True
                    Sheet22.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HD8D8D8)
                    Sheet22.Range("HP.AnnualLetableValue" & HpCount).ClearContents
    
                    Sheet22.Range("HP.RentNotRealized" & HpCount).Locked = True
                    Sheet22.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HD8D8D8)
                    Sheet22.Range("HP.RentNotRealized" & HpCount).ClearContents
    
                    Sheet22.Range("HP.LocalTaxes" & HpCount).Locked = True
                    Sheet22.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HD8D8D8)
                    Sheet22.Range("HP.LocalTaxes" & HpCount).ClearContents
                    
                    Else
                    
                    Sheet22.Range("HP.AnnualLetableValue" & HpCount).Locked = False
                    Sheet22.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HCCFFCC)
                    Sheet22.Range("HP.AnnualLetableValue" & HpCount).ClearContents
    
                    Sheet22.Range("HP.RentNotRealized" & HpCount).Locked = False
                    Sheet22.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HCCFFCC)
                    Sheet22.Range("HP.RentNotRealized" & HpCount).ClearContents
    
                    Sheet22.Range("HP.LocalTaxes" & HpCount).Locked = False
                    Sheet22.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HCCFFCC)
                    Sheet22.Range("HP.LocalTaxes" & HpCount).ClearContents
                    
                End If
                '---------------------------------------------------------------------------
                Sheet22.Range("HP.PANofTenant" & HpCount).Locked = True
                Sheet22.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet22.Range("HP.PANofTenant" & HpCount).ClearContents
                
            Else
                Sheet22.Range("HP.NameofTenant" & HpCount).Locked = False
                Sheet22.Range("HP.NameofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.NameofTenant" & HpCount).ClearContents

                Sheet22.Range("HP.PANofTenant" & HpCount).Locked = False
                Sheet22.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.PANofTenant" & HpCount).ClearContents
                
                Sheet22.Range("HP.AadharofTenant" & HpCount).Locked = False
                Sheet22.Range("HP.AadharofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.AadharofTenant" & HpCount).ClearContents
                
                Sheet22.Range("HP.TANofTenant" & HpCount).Locked = False
                Sheet22.Range("HP.TANofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.TANofTenant" & HpCount).ClearContents
                
                Sheet22.Range("HP.AnnualLetableValue" & HpCount).Locked = False
                Sheet22.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.AnnualLetableValue" & HpCount).ClearContents

                Sheet22.Range("HP.RentNotRealized" & HpCount).Locked = False
                Sheet22.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.RentNotRealized" & HpCount).ClearContents

                Sheet22.Range("HP.LocalTaxes" & HpCount).Locked = False
                Sheet22.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet22.Range("HP.LocalTaxes" & HpCount).ClearContents
            End If
        End If
        
        

            
    If Not Application.Intersect(Range("HP.Co.Pan" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
        Range(Target.Address).value = UCase(Range(Target.Address).value)
        
            If Not ValidateTenantPan1(Range(Target.Address).value) Then
                fmsgboxStatus "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP"
                Range(Target.Address).value = ""
                Range(Target.Address).Select
            End If
        End If
    '--end

       If Not Application.Intersect(Range("HP.PANofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
        Range(Target.Address).value = UCase(Range(Target.Address).value)

        'If Target.name.name("HP.PANofTenant" & HpCount) Then
        '    Sheet22.Range("HP.PANofTenant" & HpCount).value = UCase(Sheet22.Range("HP.PANofTenant" & HpCount).value)
        '    MsgBx_TenantPan = ""
        '    Sheet22.Range("HP.PANofTenant" & HpCount).value = UCase(Sheet22.Range("HP.PANofTenant" & HpCount).value)
            If Not ValidateTenantPan1(Range(Target.Address).value) Then
                'fmsgboxStatus "Enter a Valid PAN fomat :: 5 Aplphabets,4 digits follwed by 1 character"
                fmsgboxStatus """Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP"""
                Range(Target.Address).value = ""
                Range(Target.Address).Select
            End If
        'End If
        End If

        If Not Application.Intersect(Range("HP.TANofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
         If Not (CheckTAN1(Range(Target.Address).value) Or ValidateTenantPan1(Range(Target.Address).value)) Then
                'fmsgboxStatus "Enter a Valid PAN/TAN fomat "
                'Malli
                'fmsgboxStatus """Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet or First 4 Alphabets, next 5 digits, then 1 Alphabet in schedule HP"""
                 'By Bindu on 24th Feb 26
                 fmsgboxStatus """Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet in case of TAN or First 4 Alphabets, next 5 digits, then 1 Alphabet in case of PAN in schedule HP"""
                Range(Target.Address).value = ""
                Range(Target.Address).Select
            End If
        'End If
        End If
    
    
    
    
                   
       If Not Application.Intersect(Range("HP.AadharofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
         If Trim(Range(Target.Address).value) <> "" Then
    
    
               If Not (IsNumeric(Range(Target.Address).value)) Then
    
                   fmsgboxStatus """Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                   Range(Target.Address).value = ""
                   Range(Target.Address).Select
               End If
           End If
       End If

 
'   ' If Target.Validation.Type = 2 Then
      If Not Intersect(Sheet22.Range("HP.CoOwnedYN" & HpCount), Range(Target.Address)) Is Nothing Or _
         Not Intersect(Sheet22.Range("HP.Co.Share" & HpCount), Range(Target.Address)) Is Nothing Then
       LockUnlockCoOwnersDetails (HpCount)
      End If
 '   End If
'
  

'
      '  If (Target.name.name) = "HP.ifLetOut" & HpCount Or (Target.name.name) = "HP.IntOnBorwCap" & HpCount Then
      If Not Application.Intersect(Sheet22.Range("HP.ifLetOut" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet22.Range("HP.IntOnBorwCap" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Then
         '------------
                 mdHouseProperty.CheckHousePropertyIncome_New (HpCount)
        End If
'
     '    If (Target.name.name) = "HP.IntOnBorwCap" & HpCount Then
     If Not Application.Intersect(Sheet22.Range("HP.IntOnBorwCap" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Then
     

            If Range("HP.ifLetOut" & HpCount).value = "Let Out" Then
                If Range("HP.TotalIncomeChargeableUnHP").value < -200000 Then
                    MsgBox "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in schedule HP" ', vbOKOnly, mdCommon.vbMessageCaption

                End If
            End If
        End If
'
'
        If Mid(Range("HP.CoOwnedYN" & HpCount).value, 1, 1) = "Y" Then
           Dim temp As String
           Dim rangecells As Range
           Dim mIntCtr As Long, i As Long
           Dim summedvalue As Double
           Set rangecells = Range("HP.Co.Share" & HpCount).Cells
          mIntCtr = Range("HP.Co.Share" & HpCount).count
          summedvalue = 0
          For i = 1 To mIntCtr
           If Not rangecells.item(i) = "" Then
                summedvalue = summedvalue + rangecells.item(i)
            End If
           Next


        End If
'
'
'
'
'
'
'

    If Not Application.Intersect(Sheet22.Range("HP.PinCode" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Then

    Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
      'pinrng_chk = Target.address
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "J", "H")
      cntryrng_chk = Replace(pinrng_chk, "J", "I")
      ZIPrng_chk = Replace(pinrng_chk, "J", "K")

      Set ws = Worksheets("HP")

      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)




    End If

'   '--------------------------------------------------------------------------
   

           
           If Not Application.Intersect(Sheet22.Range("HP.StateCode" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Or _
              Not Application.Intersect(Sheet22.Range("HP.CountryCode" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Or _
              Not Application.Intersect(Sheet22.Range("HP.PinCode" & HpCount), Sheet22.Range(Target.Address)) Is Nothing Then
         '------------
         Sheet22.Unprotect Password:=getmsgstate
           ' Sheet22.Unprotect getmsgstate
            If isdropdownblank(Sheet22.Range("HP.StateCode" & HpCount).value) Then
                Sheet22.Range("HP.CountryCode" & HpCount).value = "91-INDIA"

                Sheet22.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet22.Range("HP.ZipCode" & HpCount).Locked = False
                Sheet22.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"

                Sheet22.Range("HP.PinCode" & HpCount).ClearContents
                Sheet22.Range("HP.PinCode" & HpCount).Locked = False
                Sheet22.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
            
            ElseIf Mid(Sheet22.Range("HP.StateCode" & HpCount).value, 1, 2) = "99" Then
             
                        If Mid(Sheet22.Range("HP.CountryCode" & HpCount).value, 1, 2) = "91" Then
                        Sheet22.Range("HP.CountryCode" & HpCount).value = "(Select)"
                        End If
                        
                        Sheet22.Range("HP.ZipCode" & HpCount).ClearContents
                        Sheet22.Range("HP.ZipCode" & HpCount).Locked = False
                        Sheet22.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"
        
                        Sheet22.Range("HP.PinCode" & HpCount).ClearContents
                        Sheet22.Range("HP.PinCode" & HpCount).Locked = True
                        Sheet22.Range("HP.PinCode" & HpCount).Interior.Color = "&HD8D8D8"
            ElseIf Mid(Sheet22.Range("HP.StateCode" & HpCount).value, 1, 2) <> "99" Then
                     
                        If (Target.name.name) = "HP.CountryCode" & HpCount And Mid(Sheet22.Range("HP.CountryCode" & HpCount).value, 1, 2) <> "91" Then  'Malli 17/10/2024
                         fmsgboxStatus """Country cannot be other than India as you have selected an Indian state in schedule HP."""
                        End If
                     
                        Sheet22.Range("HP.CountryCode" & HpCount).value = "91-INDIA"
                        Sheet22.Range("HP.ZipCode" & HpCount).ClearContents
                        Sheet22.Range("HP.ZipCode" & HpCount).Locked = True
                        Sheet22.Range("HP.ZipCode" & HpCount).Interior.Color = "&HD8D8D8"
        
                        If Sheet22.Range("HP.PinCode" & HpCount).value = "" Then
                        Sheet22.Range("HP.PinCode" & HpCount).ClearContents
                        Sheet22.Range("HP.PinCode" & HpCount).Locked = False
                        Sheet22.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
                        End If
                        
                        Sheet22.Range("HP.ZipCode" & HpCount).ClearContents
                        Sheet22.Range("HP.ZipCode" & HpCount).Locked = True
                        Sheet22.Range("HP.ZipCode" & HpCount).Interior.Color = "&HD8D8D8"
                        
                    End If
            End If
             
      Next
    
'Added by Riyaz
Dim R As Long
For R = 1 To Sheet22.Range("PropertySectionCount").value
            If Range("HP.ifLetOut" & R).value = "Let Out" Then
                If Range("HP.TotalIncomeChargeableUnHP").value < -200000 Then
                    fmsgboxsmall "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in schedule HP" ', vbOKOnly, mdCommon.vbMessageCaption
                    
                End If
            End If
 
   If Mid(Range("HP.ifLetOut" & R).value, 1, 1) = "S" Then
        Dim hpInterest1, salIncome1 As Variant
           hpInterest1 = Range("HP.IntOnBorwCap" & R).value
            If Range("HP.ifLetOut" & R).value = "Self Occupied" Then
             
                     If Not mdHouseProperty.CheckHousePropertyIncome_New(hpInterest1) Then
                     hpInterest1 = Range("HP.IntOnBorwCap" & R).value
                     End If
                     
                     Dim CheckSelfOccupied_count As Variant
                     Dim CheckSelfOccupied_IntOnBorwCap_count As Variant
                     CheckSelfOccupied_count = CheckSelfOccupied_count + 1
                     CheckSelfOccupied_IntOnBorwCap_count = CheckSelfOccupied_IntOnBorwCap_count + hpInterest1
                     
                     'Malli-------------------
    
                     If CheckSelfOccupied_count > 2 Then
                        fmsgboxsmall "* ""Self occupied property cannot be selected more than twice""" & Chr(13)
                        Range("HP.ifLetOut" & R).value = "(Select)"
                        CheckSelfOccupied_count = 0
                     End If
    
                     If CheckSelfOccupied_IntOnBorwCap_count > 200000 Then
              fmsgboxsmall "*""Amount of deduction for interest u/s 24(b) in case of Self Occupied House Property or aggregate of such properties  cannot exceed 2 lakh. Therefore, interest is being restrcited to 2 lakh""" & Chr(13)

                     Range("HP.IntOnBorwCap" & R).value = ""
                        CheckSelfOccupied_IntOnBorwCap_count = 0
                     End If
    '------------------------
                End If
    End If
    

        If Sheet22.Range("HP.AnnualLetableValue" & R).value = "" Or Sheet22.Range("HP.AnnualLetableValue" & R).value = 0 Then
        If Sheet22.Range("HP.LocalTaxes" & R).value > 0 Then
        fmsgboxsmall "Tax paid to local authorities can be claimed only if income from house property is declared in schedule HP" & Chr(13)
        Target.value = ""
        End If
        End If
Next
'----End----

  Dim c As Long
    For c = 1 To Sheet22.Range("PropertySectionCOunt").value
    If Not Application.Intersect(Range("HP.ifLetOut" & c).Cells, Range(Target.Address)) Is Nothing Then
            If Sheet8.Range("bacValue").value = 1 And Mid(Sheet22.Range("HP.ifLetOut" & c).value, 1, 1) = "S" Then
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & c).Locked = True
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet22.Range("bankName.24b" & c).Locked = True
                Sheet22.Range("bankName.24b" & c).ClearContents
                Sheet22.Range("bankName.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet22.Range("loanAccNum.24b" & c).Locked = True
                Sheet22.Range("loanAccNum.24b" & c).ClearContents
                Sheet22.Range("loanAccNum.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet22.Range("loanDate.24b" & c).Locked = True
                Sheet22.Range("loanDate.24b" & c).ClearContents
                Sheet22.Range("loanDate.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet22.Range("loanAmt.24b" & c).Locked = True
                Sheet22.Range("loanAmt.24b" & c).ClearContents
                Sheet22.Range("loanAmt.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet22.Range("loanOutstanding.24b" & c).Locked = True
                Sheet22.Range("loanOutstanding.24b" & c).ClearContents
                Sheet22.Range("loanOutstanding.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet22.Range("Intrst.24b" & c).Locked = True
                Sheet22.Range("Intrst.24b" & c).ClearContents
                Sheet22.Range("Intrst.24b" & c).Interior.Color = (&HD8D8D8)
                
            Else
                
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & c).Locked = False
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet22.Range("bankName.24b" & c).Locked = False
                Sheet22.Range("bankName.24b" & c).ClearContents
                Sheet22.Range("bankName.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet22.Range("loanAccNum.24b" & c).Locked = False
                Sheet22.Range("loanAccNum.24b" & c).ClearContents
                Sheet22.Range("loanAccNum.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet22.Range("loanDate.24b" & c).Locked = False
                Sheet22.Range("loanDate.24b" & c).ClearContents
                Sheet22.Range("loanDate.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet22.Range("loanAmt.24b" & c).Locked = False
                Sheet22.Range("loanAmt.24b" & c).ClearContents
                Sheet22.Range("loanAmt.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet22.Range("loanOutstanding.24b" & c).Locked = False
                Sheet22.Range("loanOutstanding.24b" & c).ClearContents
                Sheet22.Range("loanOutstanding.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet22.Range("Intrst.24b" & c).Locked = False
                Sheet22.Range("Intrst.24b" & c).ClearContents
                Sheet22.Range("Intrst.24b" & c).Interior.Color = (&HCCFFCC)
                
            End If
        End If
    Next
 



Dim s As Long
Dim SR As Long
SR = 0
Sheet22.Unprotect Password:=getmsgstate
For s = 1 To Sheet22.Range("PropertySectionCount").value

    If Sheet8.Range("bacValue").value = 1 And Mid(Sheet22.Range("HP.ifLetOut" & s).value, 1, 1) = "S" Then
        Sheet22.Range("HP.IntOnBorwCap" & s).value = 0
    ElseIf Sheet8.Range("bacValue").value = 2 And Mid(Sheet22.Range("HP.ifLetOut" & s).value, 1, 1) = "S" Then
            Sheet22.Range("HP.IntOnBorwCap" & s).value = Application.WorksheetFunction.Min(200000 - SR, Sheet22.Range("TotAmt.24b" & s).value)
            SR = SR + Sheet22.Range("HP.IntOnBorwCap" & s).value
    ElseIf (Sheet22.Range("HP.AnnualLetableValue" & s).value = 0 Or Sheet22.Range("HP.AnnualLetableValue" & s).value = "") And Mid(Sheet22.Range("HP.ifLetOut" & s).value, 1, 1) <> "S" Then
            Sheet22.Range("HP.IntOnBorwCap" & s).value = 0
    Else
        Sheet22.Range("HP.IntOnBorwCap" & s).value = Sheet22.Range("TotAmt.24b" & s).value
    End If

Next
Sheet22.Protect Password:=getmsgstate
    
endline:
'----------------Lock Password-------------------START---
   Sheet22.Protect Password:=getmsgstate
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Sub


Sub LockUnlock_24bHP()
On Error Resume Next
'Application.EnableEvents = False
    Sheet22.Unprotect Password:=getmsgstate
    Dim HpCount1 As Long
 
       For HpCount1 = 1 To Sheet22.Range("PropertySectionCount").value
    'If Not Application.Intersect(Range("HP.ifLetOut" & HpCount1).Cells, Range(Target.Address)) Is Nothing Then
            If Sheet8.Range("bacValue").value = 1 And Mid(Sheet22.Range("HP.ifLetOut" & HpCount1).value, 1, 1) = "S" Then
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Locked = True
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Interior.Color = (&HD8D8D8)

                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("bankName.24b" & HpCount1).Locked = True
                Sheet22.Range("bankName.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("bankName.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAccNum.24b" & HpCount1).Locked = True
                Sheet22.Range("loanAccNum.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAccNum.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanDate.24b" & HpCount1).Locked = True
                Sheet22.Range("loanDate.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanDate.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAmt.24b" & HpCount1).Locked = True
                Sheet22.Range("loanAmt.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAmt.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanOutstanding.24b" & HpCount1).Locked = True
                Sheet22.Range("loanOutstanding.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanOutstanding.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("Intrst.24b" & HpCount1).Locked = True
                Sheet22.Range("Intrst.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("Intrst.24b" & HpCount1).Interior.Color = (&HD8D8D8)
            Else
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Locked = False
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Interior.Color = (&HCCFFCC)

                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("bankName.24b" & HpCount1).Locked = False
                Sheet22.Range("bankName.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("bankName.24b" & HpCount1).Interior.Color = (&HCCFFCC)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAccNum.24b" & HpCount1).Locked = False
                Sheet22.Range("loanAccNum.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAccNum.24b" & HpCount1).Interior.Color = (&HCCFFCC)

                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanDate.24b" & HpCount1).Locked = False
                Sheet22.Range("loanDate.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanDate.24b" & HpCount1).Interior.Color = (&HCCFFCC)

                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAmt.24b" & HpCount1).Locked = False
                Sheet22.Range("loanAmt.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanAmt.24b" & HpCount1).Interior.Color = (&HCCFFCC)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanOutstanding.24b" & HpCount1).Locked = False
                Sheet22.Range("loanOutstanding.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("loanOutstanding.24b" & HpCount1).Interior.Color = (&HCCFFCC)
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("Intrst.24b" & HpCount1).Locked = False
                Sheet22.Range("Intrst.24b" & HpCount1).ClearContents
                Sheet22.Unprotect Password:=getmsgstate
                Sheet22.Range("Intrst.24b" & HpCount1).Interior.Color = (&HCCFFCC)
            End If
        'End If
    Next
 
 
    Sheet22.Protect Password:=getmsgstate
End Sub
