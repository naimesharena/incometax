Attribute VB_Name = "Sheet15"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
'//V0.7_AY_2025_26  Malli_05/05/2025 start development//

Application.EnableEvents = False

Dim IfsCode As Range
Dim sPassword

Dim IffCode As Range
Dim TargtVal, targetadd As Variant

TargtVal = Target.value
targetadd = Target.Address
targetadd = Replace(targetadd, "$", "")

Sheet15.Unprotect Password:=getmsgstate


    If Not Application.Intersect(Sheet15.Range("Sch10of13A_PlaceofWrk"), Sheet15.Range(Target.Address)) Is Nothing Then


     If Target.value <> "" And UCase(Target.value) <> UCase("(Select)") Then
            Sheet15.Unprotect Password:=sPassword

            Sheet15.Range("Sch10of13A_BasicSalary").Locked = False
            Sheet15.Range("Sch10of13A_BasicSalary").Interior.Color = "&HCCFFCC"

            Sheet15.Range("Sch10of13A_DearAllowance").Locked = False
            Sheet15.Range("Sch10of13A_DearAllowance").Interior.Color = "&HCCFFCC"

        Else
            Sheet15.Unprotect Password:=sPassword

            Sheet15.Range("Sch10of13A_BasicSalary").ClearContents
            Sheet15.Range("Sch10of13A_BasicSalary").Locked = True
            Sheet15.Range("Sch10of13A_BasicSalary").Interior.Color = "&HD8D8D8"

            Sheet15.Range("Sch10of13A_DearAllowance").ClearContents
            Sheet15.Range("Sch10of13A_DearAllowance").Locked = True
            Sheet15.Range("Sch10of13A_DearAllowance").Interior.Color = "&HD8D8D8"
        End If

    End If
    
    
If Not Application.Intersect(Sheet15.Range("Sch10of13A_BasicSalary"), Sheet15.Range(Target.Address)) Is Nothing Or _
   Not Application.Intersect(Sheet15.Range("Sch10of13A_DearAllowance"), Sheet15.Range(Target.Address)) Is Nothing Or _
   Not Application.Intersect(Sheet15.Range("Sch10of13A_50Por40Pofsalary"), Sheet15.Range(Target.Address)) Is Nothing Then
  
              Dim Per50Salaryaspersection17_i As Variant
              Per50Salaryaspersection17_i = Sheet15.Range("Sch10of13A_DetlsofSalpersec17of1").value * 0.5
                If Sheet15.Range("Sch10of13A_50Por40Pofsalary").value > Per50Salaryaspersection17_i Then
                'fmsgbox "* ""The value in this field shall not be more than 50% of field 4 ""Details of Salary as per Section 17(1)"""""
                fmsgbox "* ""The value in this Field 6 ""50% /40% of salary (C)"" shall not be more than 50% of Field 4 ""Details of Salary as per Section 17(1)"""""
                Sheet15.Range("Sch10of13A_50Por40Pofsalary").value = ""
                End If
  
End If

endline:

Application.EnableEvents = True
'sPassword = EfilingCommon.getmsgstate
'Sheet15.Protect Password:=sPassword
Sheet15.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Sub nnnnnn()
Application.EnableEvents = True
End Sub


