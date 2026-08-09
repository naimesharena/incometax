Attribute VB_Name = "Sheet8"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim MsgBx_TenantPan As String

Dim ActiveCellRange As String



Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Sheet8.Shapes("Button 14").Left = Range("Button_HP").Left
Sheet8.Shapes("Button 14").Top = Range("Button_HP").Top

Sheet8.Shapes("Button 17").Left = Range("Button_HP2").Left
Sheet8.Shapes("Button 17").Top = Range("Button_HP2").Top

Sheet8.Shapes("Button 75").Left = Range("Button_HP2").Left + 180
Sheet8.Shapes("Button 75").Top = Range("Button_HP2").Top


Sheet8.Protect Password:=sPassword

Sheet8.Activate
'    Dim HpCount As Long
'
'    For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value
'        Application.EnableEvents = False
'        sPassword = EfilingCommon.getmsgstate
'        ActiveSheet.Unprotect Password:=sPassword
'        If SheetALL.Range("bacValue").value = 1 Then
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
'        Else
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
'        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
'        End If
'        Application.EnableEvents = True
'        ActiveSheet.Protect Password:=sPassword
'    Next


End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim validateType As String
Dim cell As Variant

'By Sai 12/09/2024=========================
            Dim PinCode As Range
            Dim DupFlag As Boolean
            Dim StateName
            Dim StateCode
            Dim state1 As Variant
'==========================================

   validateType = ""
   Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

  'To add Income to House Property
'  AddTotIncomeUnderHouseProperty

  'Set cell = Sheet8.Range("HP.Co.Pan").Cells


    Dim HpCount As Long

       For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value

       For Each cell In Sheet8.Range("HP.Co.Pan" & HpCount).Cells
        Sheet8.Range("" & cell.Address).value = UCase(Sheet8.Range("" & cell.Address).value)
       Next

       For Each cell In Sheet8.Range("HP.PANofTenant" & HpCount).Cells
        Sheet8.Range("" & cell.Address).value = UCase(Sheet8.Range("" & cell.Address).value)
       Next
       
       For Each cell In Sheet8.Range("HP.AadharofTenant" & HpCount).Cells
        Sheet8.Range("" & cell.Address).value = UCase(Sheet8.Range("" & cell.Address).value)
       Next


        For Each cell In Sheet8.Range("HP.TANofTenant" & HpCount).Cells
        Sheet8.Range("" & cell.Address).value = UCase(Sheet8.Range("" & cell.Address).value)
       Next


       If Not Application.Intersect(Range("HP.PANofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
        Range(Target.Address).value = UCase(Range(Target.Address).value)

        'If Target.name.name("HP.PANofTenant" & HpCount) Then
        '    Sheet8.Range("HP.PANofTenant" & HpCount).value = UCase(Sheet8.Range("HP.PANofTenant" & HpCount).value)
        '    MsgBx_TenantPan = ""
        '    Sheet8.Range("HP.PANofTenant" & HpCount).value = UCase(Sheet8.Range("HP.PANofTenant" & HpCount).value)
            If Not ValidateTenantPan(Range(Target.Address).value) Then
                fmsgboxStatus "Enter a Valid PAN fomat :: 5 Aplphabets,4 digits follwed by 1 character"
                Range(Target.Address).value = ""
                Range(Target.Address).Select
            End If
        'End If
        End If

        If Not Application.Intersect(Range("HP.TANofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
        Range(Target.Address).value = UCase(Range(Target.Address).value)

        'If Target.name.name("HP.PANofTenant" & HpCount) Then
        '    Sheet8.Range("HP.PANofTenant" & HpCount).value = UCase(Sheet8.Range("HP.PANofTenant" & HpCount).value)
        '    MsgBx_TenantPan = ""
        '    Sheet8.Range("HP.PANofTenant" & HpCount).value = UCase(Sheet8.Range("HP.PANofTenant" & HpCount).value)
            If Not (CheckTAN(Range(Target.Address).value) Or ValidateTenantPan(Range(Target.Address).value)) Then
                fmsgboxStatus "Enter a Valid PAN/TAN fomat "
                Range(Target.Address).value = ""
                Range(Target.Address).Select
            End If
        'End If
        End If
        
'Commented by Jyoti2025 as it was auto populated field and changed the logic as per DE V0.5 and V0.6
'*************

' 'ramya11=====================================================================================================
'        If SheetALL.Range("bacValue").value = 1 Then
'        'If Range("HP.ifLetOut1").value = "Let Out" Or Range("HP.ifLetOut1").value = "Deemed Let Out" Then
'        If Range("HP.ifLetOut" & HpCount).value = "Let Out" Or Range("HP.ifLetOut" & HpCount).value = "Deemed Let Out" Or Range("HP.ifLetOut" & HpCount).value = "Self Occupied" Then
''ramya11=========================================================================================================
'                    Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
'        Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
'        Else
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
'            Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
'        End If
'        End If
'************
 
 
  '===================================Need to Check 27/02/2024
  
'        If Range("HP.ifLetOut" & HpCount).value = "Let Out" Or Range("HP.ifLetOut" & HpCount).value = "Deemed Let Out" And Range("HP.AnnualLetableValue" & HpCount).value = "" Then
'                    Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = True
'                    Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HD8D8D8)
'
'          ElseIf Range("HP.ifLetOut" & HpCount).value = "Let Out" Or Range("HP.ifLetOut" & HpCount).value = "Deemed Let Out" And Range("HP.AnnualLetableValue" & HpCount).value <> "" Or Range("HP.AnnualLetableValue" & HpCount).value > 0 Then
'            Else
'                    Sheet8.Range("HP.IntOnBorwCap" & HpCount).Locked = False
'                    Sheet8.Range("HP.IntOnBorwCap" & HpCount).ClearContents
'                    Sheet8.Range("HP.IntOnBorwCap" & HpCount).Interior.Color = (&HCCFFCC)
'        End If

 '==============================================================

    If Target.Validation.Type = 2 Then
        If Not Intersect(Sheet8.Range("HP.Co.Share" & HpCount), Range(Target.Address)) Is Nothing Then
         LockUnlockCoOwnersDetails (HpCount)
        End If
    End If


     If Target.Validation.Type = 3 Then

        If (Target.name.name) = "HP.StateCode" & HpCount Then
            Sheet8.Unprotect getmsgstate
            If isdropdownblank(Sheet8.Range("HP.StateCode" & HpCount).value) Then
                Sheet8.Range("HP.CountryCode" & HpCount).value = "(Select)"

                Sheet8.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet8.Range("HP.ZipCode" & HpCount).Locked = False
                Sheet8.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"

                Sheet8.Range("HP.PinCode" & HpCount).ClearContents
                Sheet8.Range("HP.PinCode" & HpCount).Locked = False
                Sheet8.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
            ElseIf Mid(Sheet8.Range("HP.StateCode" & HpCount).value, 1, 2) = "99" Then
                Sheet8.Range("HP.CountryCode" & HpCount).value = "(Select)"
                Sheet8.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet8.Range("HP.ZipCode" & HpCount).Locked = False
                Sheet8.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"

                Sheet8.Range("HP.PinCode" & HpCount).ClearContents
                Sheet8.Range("HP.PinCode" & HpCount).Locked = True
                Sheet8.Range("HP.PinCode" & HpCount).Interior.Color = "&HD8D8D8"
             ElseIf Mid(Sheet8.Range("HP.StateCode" & HpCount).value, 1, 2) <> "99" Then
                Sheet8.Range("HP.CountryCode" & HpCount).value = "91-INDIA"
                Sheet8.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet8.Range("HP.ZipCode" & HpCount).Locked = True
                Sheet8.Range("HP.ZipCode" & HpCount).Interior.Color = "&HD8D8D8"

                If Sheet8.Range("HP.PinCode" & HpCount).value = "" Then  'by sai
                Sheet8.Range("HP.PinCode" & HpCount).ClearContents
                Sheet8.Range("HP.PinCode" & HpCount).Locked = False
                Sheet8.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
                End If
                
            End If
            
            
'Newly added by sai=================================================================
     If Not Intersect(Range("HP.StateCode" & HpCount), Range(Target.Address)) Is Nothing Then
         If Mid(Sheet8.Range("HP.StateCode" & HpCount).value, 1, 2) = "99" Then
             Sheet8.Range("HP.ZipCode" & HpCount).Locked = False
             Sheet8.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"
             Sheet8.Range("HP.ZipCode" & HpCount).ClearContents
             Sheet8.Range("HP.PinCode" & HpCount).Locked = True
             Sheet8.Range("HP.PinCode" & HpCount).Interior.Color = "&HD8D8D8"
             Sheet8.Range("HP.PinCode" & HpCount).ClearContents
             
       ElseIf Mid(Sheet8.Range("HP.StateCode" & HpCount).value, 1, 2) <> "99" Then
            If Sheet8.Range("HP.StateCode" & HpCount).value <> "(Select)" Then
             Sheet8.Range("HP.PinCode" & HpCount).Locked = False
             Sheet8.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
             Sheet8.Range("HP.PinCode" & HpCount).ClearContents
             Sheet8.Range("HP.ZipCode" & HpCount).Locked = True
             Sheet8.Range("HP.ZipCode" & HpCount).Interior.Color = "&HD8D8D8"
             Sheet8.Range("HP.ZipCode" & HpCount).ClearContents
            End If
          End If
            
       End If
  '==============================================================================================
            
'By Sai on 12/09/2024
'=========================================================
            If Not Application.Intersect(Sheet8.Range("HP.StateCode" & HpCount), Sheet8.Range(Target.Address)) Is Nothing Then
                     If Sheet8.Range(Replace(Target.Address, "H", "J")).value <> "" Then
                         state1 = Application.VLookup(Sheet8.Range(Replace(Target.Address, "H", "J")).value, SheetALL.Range("All_PinCode_V"), 2, False)
                         If Not (state1 = UCase(Mid(Target.value, 4)) Or EfilingCommon.StateMatchesPin(Target.value, Sheet8.Range(Replace(Target.Address, "H", "J")).value)) Then
                                 'Target.value = ""
                                 Sheet8.Range(Replace(Target.Address, "H", "J")).value = ""
                         End If
                    End If
            End If
'==========================================================


            Sheet8.Protect getmsgstate
            GoTo endline
          End If

    
        If (Target.name.name) = "HP.CoOwnedYN" & HpCount Then
           LockUnlockCoOwnersDetails (HpCount)
        End If

        If (Target.name.name) = "HP.ifLetOut" & HpCount Or (Target.name.name) = "HP.IntOnBorwCap" & HpCount Then
                 SchHP.CheckHousePropertyIncome (HpCount)
        End If
        
         If (Target.name.name) = "HP.IntOnBorwCap" & HpCount Then

            If Range("HP.ifLetOut" & HpCount).value = "Let Out" Then
                If Range("HP.TotalIncomeChargeableUnHP").value < -200000 Then
                    MsgBox "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
                    'Range("HP.IntOnBorwCap" & HpCount).value = ""
                End If
            End If
        End If


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

'            If Range("HP.SharePercent" & HpCount).value + summedvalue > 100 Then
'            fmsgboxStatus "Sum of Assessee percentage and Co-Owner(s) Percentage must be be equal to 100 %"
'            Range("HP.SharePercent" & HpCount).ClearContents
'            Range("HP.Co.Share" & HpCount).ClearContents
'        End If
        End If



        If Target.name.name = ("HP.OwnerProperty" & HpCount) Then
           If Mid(Sheet8.Range("HP.OwnerProperty" & HpCount).value, 1, 1) = "O" Then
                Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).Locked = False
                'Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).ClearContents
                Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).Interior.Color = (&HCCFFCC)

           Else
                Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).Locked = True
                Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).ClearContents
                Sheet8.Range("HP.OwnerPropertyDescription" & HpCount).Interior.Color = (&HD8D8D8)


         End If
        End If

        If Target.name.name = ("HP.ifLetOut" & HpCount) Then
           If Mid(Sheet8.Range("HP.ifLetOut" & HpCount).value, 1, 1) = "S" Then
                'Sheet8.Range("HP.NameofTenant" & HpCount).Locked = True
                Sheet8.Range("HP.NameofTenant" & HpCount).ClearContents
                'Sheet8.Range("HP.NameofTenant" & HpCount).Interior.Color = (&HD8D8D8)

                'Sheet8.Range("HP.PANofTenant" & HpCount).Locked = True
                'Sheet8.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet8.Range("HP.PANofTenant" & HpCount).ClearContents
                Sheet8.Range("HP.TANofTenant" & HpCount).ClearContents
                Sheet8.Range("HP.AnnualLetableValue" & HpCount).Locked = True
                Sheet8.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet8.Range("HP.AnnualLetableValue" & HpCount).ClearContents

'Change-23.05.2023.101 -> Issue number 3 in production build 1
                Sheet8.Range("HP.RentNotRealized" & HpCount).Locked = True
                Sheet8.Range("HP.RentNotRealized" & HpCount).ClearContents
                Sheet8.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HD8D8D8)
                
                Sheet8.Range("HP.LocalTaxes" & HpCount).Locked = True
                Sheet8.Range("HP.LocalTaxes" & HpCount).ClearContents
                Sheet8.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HD8D8D8)

'----
                'Sheet8.Range("HP.RentNotRealized" & HpCount).Locked = True
                'Sheet8.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet8.Range("HP.RentNotRealized" & HpCount).ClearContents

                'Sheet8.Range("HP.LocalTaxes" & HpCount).Locked = True
                'Sheet8.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet8.Range("HP.LocalTaxes" & HpCount).ClearContents
            Else:
                'Sheet8.Range("HP.NameofTenant" & HpCount).Locked = False
                'Sheet8.Range("HP.NameofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet8.Range("HP.NameofTenant" & HpCount).ClearContents

                'Sheet8.Range("HP.PANofTenant" & HpCount).Locked = False
                'Sheet8.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet8.Range("HP.PANofTenant" & HpCount).ClearContents
                Sheet8.Range("HP.TANofTenant" & HpCount).ClearContents
                Sheet8.Range("HP.AnnualLetableValue" & HpCount).Locked = False
                Sheet8.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet8.Range("HP.AnnualLetableValue" & HpCount).ClearContents

'Change-23.05.2023.101 -> Issue number 3 in production build 1
                Sheet8.Range("HP.RentNotRealized" & HpCount).Locked = False
                Sheet8.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HCCFFCC)

                Sheet8.Range("HP.LocalTaxes" & HpCount).Locked = False
                Sheet8.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HCCFFCC)
'-----

                'Sheet8.Range("HP.RentNotRealized" & HpCount).Locked = False
                'Sheet8.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet8.Range("HP.RentNotRealized" & HpCount).ClearContents

                'Sheet8.Range("HP.LocalTaxes" & HpCount).Locked = False
                'Sheet8.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet8.Range("HP.LocalTaxes" & HpCount).ClearContents
            End If
        End If
    End If
    
      'Change 01, CDB - , 30.06.2022
      
      'Commented by sai 12/09/2024====================================
      
      
'        If Not Intersect(Range("HP.PinCode" & HpCount), Range(Target.Address)) Is Nothing Then
'
'           Dim PinCode As Range
'           Dim found As Variant
'            If (Target.value <> "") Then
'           found = False
'
'           For Each PinCode In SheetALL.Range("PincodeList")
'               Dim code As String
'               code = Left(CStr(Target.value), 3)
'               If CStr(PinCode.value) = CStr(code) Then
'                   Range("HP.StateCode" & HpCount).value = SheetALL.Range("CX" & PinCode.row).value
'                   found = True
'                   Exit For
'               End If
'            Next PinCode
'            If found = False Then
'            MsgBox "Enter correct Pincode"
'            Target.value = ""
'            End If
'            End If
'            End If
'
'---End Change
    'End by sai 12/09/2024====================================
    
 '   Next
        
  'by sai 12/09/2024====================================================================================================

If Not Application.Intersect(Sheet8.Range("HP.PinCode" & HpCount), Sheet8.Range(Target.Address)) Is Nothing Then
  
    Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
    Dim ws As Worksheet
      'pinrng_chk = Target.address
        pinrng_chk = Replace(Target.Address, "$", "")
        staterng_chk = Replace(pinrng_chk, "J", "H")
        cntryrng_chk = Replace(pinrng_chk, "J", "I")
        ZIPrng_chk = Replace(pinrng_chk, "J", "K")
      
      Set ws = Worksheets("House Property")
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
End If
    Next
  
    
'HP income Validation on Data Entry commented by sai on 27/05/2025
'***************************************************************************************************************
'        If (target.name.name) = "HP.IntOnBorwCap1" Then
'
'            If Range("HP.ifLetOut1").value = "Let Out" Then
'                If Range("HP.TotalIncomeChargeableUnHP").value < -200000 Then
'                    fmsgboxsmall "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in Sheet Income Details" ', vbOKOnly, mdCommon.vbMessageCaption
'                End If
'            End If
'        End If
'
'         If (target.name.name) = "HP.IntOnBorwCap1" Then
'            Dim hpInterest, salIncome As Variant
'            hpInterest = Range("HP.IntOnBorwCap1").value
'                    If Range("HP.ifLetOut1").value = "Self Occupied" Then
'                    SchHP.CheckHousePropertyIncome (hpInterest)
'                    End If
'        End If
'
'   If Mid(Range("HP.ifLetOut1").value, 1, 1) = "S" Then
'
'
'         Dim hpInterest1, salIncome1 As Variant
'           hpInterest1 = Range("HP.IntOnBorwCap1").value
'            If Range("HP.ifLetOut1").value = "Self Occupied" Then
'                SchHP.CheckHousePropertyIncome (hpInterest1)
'            End If
'    End If
'***************************************************************************************************************



'Added by sai on 27/05/2025
'*******************************************************************************************************************



Dim r As Long
For r = 1 To Sheet8.Range("PropertySectionCOunt").value
'Ayush_13/06
    If Not Application.Intersect(Sheet8.Range("Intrst.24b" & r), Sheet8.Range(Target.Address)) Is Nothing Then
            If Range("HP.ifLetOut" & r).value = "Let Out" Then
                If Range("HP.TotalIncomeChargeableUnHP").value < -200000 Then
                    fmsgboxsmall "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in schedule HP" ', vbOKOnly, mdCommon.vbMessageCaption
                    'Range("HP.IntOnBorwCap" & R).value = ""
                End If
            End If
        End If
        
 
   If Mid(Range("HP.ifLetOut" & r).value, 1, 1) = "S" Then
        Dim hpInterest1, salIncome1 As Variant
        
            Sheet8.Range("HP.IntOnBorwCap" & r).value = Sheet8.Range("TotAmt.24b" & r).value
            
           hpInterest1 = Range("HP.IntOnBorwCap" & r).value
           
            If Range("HP.ifLetOut" & r).value = "Self Occupied" Then

                ' SchHP.CheckHousePropertyIncome (hpInterest1)
                     If Not SchHP.CheckHousePropertyIncome(hpInterest1) Then
            'by sai on 26/05/2025
'                           Range("HP.IntOnBorwCap" & R).value = ""
                            hpInterest1 = Range("HP.IntOnBorwCap" & r).value
                     End If
                     'HP.IntOnBorwCap1
                     Dim CheckSelfOccupied_count As Variant
                     Dim CheckSelfOccupied_IntOnBorwCap_count As Variant
                     CheckSelfOccupied_count = CheckSelfOccupied_count + 1
                     CheckSelfOccupied_IntOnBorwCap_count = CheckSelfOccupied_IntOnBorwCap_count + hpInterest1

    
                     If CheckSelfOccupied_count > 2 Then
                        fmsgboxsmall "* ""Self occupied property cannot be selected more than twice""" & Chr(13)
                        Range("HP.ifLetOut" & r).value = "(Select)"
                        CheckSelfOccupied_count = 0
                     End If
    
'                     If CheckSelfOccupied_IntOnBorwCap_count > 200000 Then
'                     'Jyoti 26_06_2025
'                        'fmsgboxsmall "* ""Interest cannot be more than Rs. 2,00,000 for category two self-occupied in schedule HP.""" & Chr(13)
'                        fmsgboxsmall "* "" Amount of deduction for interest u/s 24(b) in case of Self Occupied House Property or aggregate of such properties  cannot exceed 2 lakh. Therefore, interest is being restrcited to 2 lakh""" & Chr(13)
'
'
'
''                       Range("HP.IntOnBorwCap" & R).value = ""
'                        CheckSelfOccupied_IntOnBorwCap_count = 0
'                     End If

                End If
    End If
    

        If Sheet8.Range("HP.AnnualLetableValue" & r).value = "" Or Sheet8.Range("HP.AnnualLetableValue" & r).value = 0 Then
            If Sheet8.Range("HP.LocalTaxes" & r).value > 0 Then
                fmsgboxsmall "Tax paid to local authorities can be claimed only if income from house property is declared in schedule HP" & Chr(13)
                Target.value = ""
            End If
        End If
Next
'***************************************************************************************************************************************
    

    '--------
    'Jyoti 19/05/2025
    
Dim TargtVal, targetadd As Variant

Application.EnableEvents = False
Application.ScreenUpdating = False

CalculateTaxCheckFlag = False
Sheet8.Unprotect Password:=getmsgstate

 'Jyoti 26_06_2025
    'start--
    Dim c As Long
    For c = 1 To Sheet8.Range("PropertySectionCOunt").value
    If Not Application.Intersect(Range("HP.ifLetOut" & c).Cells, Range(Target.Address)) Is Nothing Then
'        If Target.name.name = ("HP.ifLetOut" & c) Then
            If SheetALL.Range("bacValue").value = 1 And Mid(Sheet8.Range("HP.ifLetOut" & c).value, 1, 1) = "S" Then
                Sheet8.Range("LoanfrmBankOrInstitute.24b" & c).Locked = True
                Sheet8.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet8.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HD8D8D8)
                              
                
                Sheet8.Range("bankName.24b" & c).Locked = True
                Sheet8.Range("bankName.24b" & c).ClearContents
                Sheet8.Range("bankName.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet8.Range("loanAccNum.24b" & c).Locked = True
                Sheet8.Range("loanAccNum.24b" & c).ClearContents
                Sheet8.Range("loanAccNum.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet8.Range("loanDate.24b" & c).Locked = True
                Sheet8.Range("loanDate.24b" & c).ClearContents
                Sheet8.Range("loanDate.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet8.Range("loanAmt.24b" & c).Locked = True
                Sheet8.Range("loanAmt.24b" & c).ClearContents
                Sheet8.Range("loanAmt.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet8.Range("loanOutstanding.24b" & c).Locked = True
                Sheet8.Range("loanOutstanding.24b" & c).ClearContents
                Sheet8.Range("loanOutstanding.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet8.Range("Intrst.24b" & c).Locked = True
                Sheet8.Range("Intrst.24b" & c).ClearContents
                Sheet8.Range("Intrst.24b" & c).Interior.Color = (&HD8D8D8)
                
            Else
                
                Sheet8.Range("LoanfrmBankOrInstitute.24b" & c).Locked = False
                Sheet8.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet8.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet8.Range("bankName.24b" & c).Locked = False
                Sheet8.Range("bankName.24b" & c).ClearContents
                Sheet8.Range("bankName.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet8.Range("loanAccNum.24b" & c).Locked = False
                Sheet8.Range("loanAccNum.24b" & c).ClearContents
                Sheet8.Range("loanAccNum.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet8.Range("loanDate.24b" & c).Locked = False
                Sheet8.Range("loanDate.24b" & c).ClearContents
                Sheet8.Range("loanDate.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet8.Range("loanAmt.24b" & c).Locked = False
                Sheet8.Range("loanAmt.24b" & c).ClearContents
                Sheet8.Range("loanAmt.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet8.Range("loanOutstanding.24b" & c).Locked = False
                Sheet8.Range("loanOutstanding.24b" & c).ClearContents
                Sheet8.Range("loanOutstanding.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet8.Range("Intrst.24b" & c).Locked = False
                Sheet8.Range("Intrst.24b" & c).ClearContents
                Sheet8.Range("Intrst.24b" & c).Interior.Color = (&HCCFFCC)
                
            End If
        End If
    Next
'--end

'Jyoti 26_06_2025
Dim s As Long
Dim SR As Long
SR = 0
Sheet8.Unprotect Password:=sPassword
For s = 1 To Sheet8.Range("PropertySectionCOunt").value

    If SheetALL.Range("bacValue").value = 1 And Mid(Sheet8.Range("HP.ifLetOut" & s).value, 1, 1) = "S" Then
        Sheet8.Range("HP.IntOnBorwCap" & s).value = 0
    ElseIf SheetALL.Range("bacValue").value = 2 And Mid(Sheet8.Range("HP.ifLetOut" & s).value, 1, 1) = "S" Then
            Sheet8.Range("HP.IntOnBorwCap" & s).value = Application.WorksheetFunction.Min(200000 - SR, Sheet8.Range("TotAmt.24b" & s).value)
            SR = SR + Sheet8.Range("HP.IntOnBorwCap" & s).value
    ElseIf (Sheet8.Range("HP.AnnualLetableValue" & s).value = 0 Or Sheet8.Range("HP.AnnualLetableValue" & s).value = "") And Mid(Sheet8.Range("HP.ifLetOut" & s).value, 1, 1) <> "S" Then
            Sheet8.Range("HP.IntOnBorwCap" & s).value = 0
    Else
        Sheet8.Range("HP.IntOnBorwCap" & s).value = Sheet8.Range("TotAmt.24b" & s).value
    End If

Next

AddTotIncomeUnderHouseProperty

Sheet8.Protect Password:=sPassword


'---
endline:
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Sub

Sub AddTotIncomeUnderHouseProperty()
    Dim TotIncomeRangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, Totsections As Long

    tot = 0
    TotIncomeRangeName = "HP.IncomeOfHP"
    Totsections = Sheet8.Range("PropertySectionCOunt").value
        For i = 1 To Totsections
        newDefinedName = TotIncomeRangeName & i
        tot = tot + Sheet8.Range(newDefinedName)
        Next
    Sheet8.Range("HP.TotalIncomeChargeableUnHP").value = tot + Sheet8.Range("HP.PassTroughIncome").value '+ Sheet8.Range("HP.RentArearsSec25BAfter30pcDeduct").value
End Sub

Private Sub Worksheet_Deactivate()

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
'Jyoti 19/05/2025
Function ValidateDate_24b(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim DateDep_IT
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet8.Range("loanDate.24b").count
    ValidateDate_24b = True

    Dim end_24
    end_24 = 0
    Set rangecells = Sheet8.Range("loanDate.24b").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_24 = end_24 + 1
        End If
    Next

    ReDim Date_24b(end_24)

    For i = 1 To end_24
        Date_24b(i) = rangecells.item(i).value
        If Not chkCompulsory(Date_24b(i)) Then
          
            fmsgboxsmall ("Please provide Date of taking loan in schedule 24(b) at Sr. No  " & i & "")
            ValidateDate_24b = False
            Exit Function
        End If

        If Not CheckDateddmmyyyy(Date_24b(i)) Then
            ValidateDate_24b = False
                fmsgboxsmall ("""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_24b(i) = Dformat(Date_24b(i), "")
            
'            If Not ChkMaxDate_24b(Trim(Date_24b(i)), "31-03-2025") Then
'                fmsgboxsmall (" Date can not be after 31/03/2025 at Sr. No " & i & ".")
                
                            If Not ChkMaxDate_24b(Trim(Date_24b(i)), "31-03-2026") Then
                fmsgboxsmall (" Date can not be after 31/03/2026 at Sr. No " & i & ".")
                
                Target.value = ""
                ValidateDate_24b = False
                 

          Exit Function
        End If
            

        End If
    Next

End Function
'--




Sub test()
Application.EnableEvents = True
End Sub
