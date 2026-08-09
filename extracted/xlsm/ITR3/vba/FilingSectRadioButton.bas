Attribute VB_Name = "FilingSectRadioButton"
'Sub RadioButton1_Click()
'
'Dim op As OptionButton
'Set op = Sheet1.Shapes("Option Button 452").OLEFormat.Object
'op.value = False
'
'Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
'
'    Dim formula As Variant
'    Dim cellrange As Range
'    Set cellrange = Sheet1.Range("sheet1.ReturnFileSec")
''Added by Sai on 05/11/2024
'    If Dformat(Sheet52.Range("sheet9.Date").value, "DD/MM/YYYY") > Dformat("31/12/2024", "DD/MM/YYYY") Then
''        formula = "=Returnfiledundersection_new"
'        formula = "=ReturnFileUnderSection1"
'
'      Else
''        formula = "=ReturnFileUnderSection1"
'        formula = "=Returnfiledundersection_new"
'    End If
'
'
'
'
'    sPassword = EfilingCommon.getmsgstate
'    Sheet1.Unprotect Password:=sPassword
'
'    With cellrange.Validation
'        .Delete
'        .add Type:=xlValidateList, _
'         AlertStyle:=xlValidAlertStop, _
'         Operator:=xlBetween, _
'         Formula1:=formula
'    End With
'
'    Sheet1.Protect Password:=sPassword
'    Sheet1.Activate
'End Sub

Sub RadioButton1_Click()

Dim op As OptionButton
Set op = Sheet1.Shapes("Option Button 452").OLEFormat.Object
op.value = False

Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"

    Dim cellrange As Range

    Set cellrange = Sheet1.Range("sheet1.ReturnFileSec")
    formula = "=ReturnFileUnderSection1"

    sPassword = EfilingCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword

    With cellrange.Validation
        .Delete
        .add Type:=xlValidateList, _
         AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, _
         Formula1:=formula
    End With

    Sheet1.Protect Password:=sPassword
    Sheet1.Activate
End Sub
Sub RadioButton2_Click()

Dim op As OptionButton
Set op = Sheet1.Shapes("Option Button 451").OLEFormat.Object
op.value = False
Dim op2 As OptionButton
Set op2 = Sheet1.Shapes("Option Button 452").OLEFormat.Object
op2.value = True

Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"

    Dim cellrange As Range

    Set cellrange = Sheet1.Range("sheet1.ReturnFileSec")
    formula = "=ReturnFileUnderSection2with153C"

    sPassword = EfilingCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword

    With cellrange.Validation
        .Delete
        .add Type:=xlValidateList, _
         AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, _
         Formula1:=formula
    End With

    Sheet1.Protect Password:=sPassword
    Sheet1.Activate
End Sub


