Attribute VB_Name = "Calculations"

Option Explicit
Public Temp_allowance As Double
Public Temp_CG As Double
Public Temp_SG As Double
Public Temp_PSU As Double
Public Temp_Pens As Double
Public Temp_OTH As Double
Public Temp_Select_Null As Double


'Central Government
'State Government
'Pu
'Pe
'O
'for Schedule S(Salary)

Function calSchS()

On Error Resume Next

Dim GrossSalary As Double
Dim GrossSalaryTotal As Double
Dim salaryCount  As Long
Dim Deduction80CCD2Total As Double
Dim Temp_allowance, Temp_DA_BS, Temp_HRA, Temp_HRA_Sec13A As Long
Dim Temp_Salary_171, Temp_Inc89A, Temp_Inc89AOth As Variant
Dim i, j As Long
Dim contribution_flag As Boolean

'----------------Unlock Password-------------------START---
'    sPassword = Common.getmsgstate
    Sheet7.Unprotect Password:=sPassword 'partbtti
    Sheet39.Unprotect Password:=sPassword 'salary
    Sheet19.Unprotect Password:=sPassword  'via
'----------------Unlock Password-------------------END-----

salaryCount = Sheet39.Range("SalarySectionCount").value
GrossSalary = 0
GrossSalaryTotal = 0
Deduction80CCD2Total = 0
Temp_allowance = 0
Temp_CG = 0
 Temp_SG = 0
 Temp_PSU = 0
 Temp_Pens = 0
 Temp_OTH = 0
 Temp_Select_Null = 0
 Temp_DA_BS = 0
 Temp_HRA = 0
 Temp_HRA_Sec13A = 0
 Temp_Salary_171 = 0
 Temp_Inc89A = 0
 Temp_Inc89AOth = 0
ReDim exemptUs10Total(salaryCount) As Double

For i = 1 To salaryCount
    GrossSalary = Sheet39.Range("SAL.Salary" & i).value
    GrossSalaryTotal = GrossSalaryTotal + GrossSalary
    
    'Sheet39.Range("SAL.TotalGrossSalary").value = GrossSalaryTotal
    If Not (Sheet39.Range("sheet.EmployerCategory" & i).value = "CG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "SG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "PSU-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "Others-Pensioners") Then
    Temp_Salary_171 = Temp_Salary_171 + Sheet39.Range("SAL.AllowancesNotExempt" & i).value
    End If
    
'    Temp_Salary_171 = Temp_Salary_171 + Sheet39.Range("SAL.AllowancesNotExempt" & i).value
    Temp_Inc89A = Temp_Inc89A + Sheet39.Range("IncomeNotified89A" & i).value
    Temp_Inc89AOth = Temp_Inc89AOth + Sheet39.Range("IncomeNotifiedOther89A" & i).value
    Deduction80CCD2Total = Deduction80CCD2Total + Sheet39.Range("SAL.AllowancesNotExempt" & i).value + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value
'Temp_allowance = Temp_allowance2 + Sheet31.Range("SAL.Salary" & i).value

For j = 1 To Range("Salary.Section" & i).count
If Not (Sheet39.Range("sheet.EmployerCategory" & i).value = "CG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "SG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "PSU-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "Others-Pensioners") Then

    If Range("Salary.Section" & i).Cells.item(j).value = "Basic Salary" Or Range("Salary.Section" & i).Cells.item(j).value = "Dearness Allowance(DA)" Then
      Temp_DA_BS = Temp_DA_BS + Range("Salary.SalaryAmount" & i).Cells.item(j).value
    End If
End If


If Range("Salary.Section" & i).Cells.item(j).value = "Basic Salary" Then
 Temp_allowance = Temp_allowance + Range("Salary.SalaryAmount" & i).Cells.item(j).value
End If

If Range("Salary.Section" & i).Cells.item(j).value = "House Rent Allowance(HRA)" Then
 Temp_HRA = Temp_HRA + Range("Salary.SalaryAmount" & i).Cells.item(j).value
End If
Next

If Sheet39.Range("sheet.EmployerCategory" & i).value = "Central Government" Then
 Temp_CG = Temp_CG + 1
End If

If Sheet39.Range("sheet.EmployerCategory" & i).value = "State Government" Then
Temp_SG = Temp_SG + 1
End If

If Sheet39.Range("sheet.EmployerCategory" & i).value = "Public Sector Undertaking" Then
Temp_PSU = Temp_PSU + 1
End If

'If Mid(Sheet39.Range("sheet.EmployerCategory" & i).value, 1, 2) = "Pe" Then
If (Sheet39.Range("sheet.EmployerCategory" & i).value = "CG-Pensioners") Or (Sheet39.Range("sheet.EmployerCategory" & i).value = "SG-Pensioners") Or (Sheet39.Range("sheet.EmployerCategory" & i).value = "PSU-Pensioners") Or (Sheet39.Range("sheet.EmployerCategory" & i).value = "Others-Pensioners") Then
Temp_Pens = Temp_Pens + 1
End If

If Sheet39.Range("sheet.EmployerCategory" & i).value = "OTHERS" Then
'SIT-118500 by sadineni on 06/05/2026
' If Sheet39.Range("sheet.EmployerCategory" & i).value = "Others" Then
   Temp_OTH = Temp_OTH + 1
 End If

If isdropdownblank(Sheet39.Range("sheet.EmployerCategory" & i).value) Then
 Temp_Select_Null = Temp_Select_Null + 1
End If


Next
Dim k As Long
For k = 1 To Range("Salary.Nature1").count

'If Mid(Range("Salary.Nature1").Cells.item(k).value, 1, 11) = "Sec 10(13A)" Then
'Temp_HRA_Sec13A = Temp_HRA_Sec13A + Range("Salary.Amount1").Cells.item(k).value
'End If

'For AY 2025-26 Newly added by sai on 25/06/2025 as it was changed to static field
'***********************************************************************************
If Sheet39.Range("Sheet39.HRA").value > 0 Then
   Temp_HRA_Sec13A = Sheet39.Range("Sheet39.HRA").value
End If
'************************************************************************************

Next

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_HRA_Sec13A").value = Temp_HRA_Sec13A

Sheet19.Unprotect Password:=sPassword  'via
Sheet39.Range("Temp_Allowance").value = Temp_allowance

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_Salary_171").value = Temp_Salary_171

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_Inc89A").value = Temp_Inc89A

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_Inc89AOth").value = Temp_Inc89AOth

    
Sheet19.Unprotect Password:=sPassword  'via
'"" or (S
Sheet19.Range("Sheet11.Salaries80CCD2").value = Deduction80CCD2Total - Sheet39.Range("SAL.ExemptUSectionOth1").value


Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_CG").value = Temp_CG

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_SG").value = Temp_SG

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_PSU").value = Temp_PSU

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_Pensioner").value = Temp_Pens

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_OTH").value = Temp_OTH

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_Null").value = Temp_Select_Null

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_DA_BS").value = Temp_DA_BS

Sheet19.Unprotect Password:=sPassword  'via
Sheet19.Range("Temp_HRA").value = Temp_HRA
'----------------Lock Password-------------------START---
    Sheet7.Protect Password:=sPassword
    Sheet39.Protect Password:=sPassword
    Sheet19.Protect Password:=sPassword  'via

    
 'added by Malli
 
 
For i = 1 To salaryCount

For j = 1 To Range("Salary.Section" & i).count
    Sheet19.Unprotect Password:=sPassword
    If Sheet39.Range("Salary.Section" & i).Cells.item(j).value = "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH" Then

       contribution_flag = True
       GoTo contrubution
    Else
       contribution_flag = False
       
    End If

Next

Next
    
contrubution:
'If contribution_flag = True And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "2001-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2007-04-01") Then
'If contribution_flag = True And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "2001-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2008-04-01") Then
    
    'changed by Chetan C M for AY 2026-27
    'start--
'    If contribution_flag = True And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "1998-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2008-04-01") Then
    If contribution_flag = True And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "1999-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2009-04-01") Then
    '-end
    
    Sheet19.Unprotect Password:=sPassword
    If Sheet19.Range("Anyother80CCH").value <= 0 Or Sheet19.Range("Anyother80CCH").value = "" Then
        Sheet19.Range("Anyother80CCH").Locked = False
        Sheet19.Range("Anyother80CCH").Interior.Color = "&HCCFFCC"
        Exit Function
    End If
       
Else
       Sheet19.Unprotect Password:=sPassword  'via
       Sheet19.Range("Anyother80CCH").ClearContents
       Sheet19.Unprotect Password:=sPassword
       Sheet19.Range("Anyother80CCH").Locked = True
       Sheet19.Range("Anyother80CCH").Interior.Color = "&HD8D8D8"

End If
    
    
    Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Function

