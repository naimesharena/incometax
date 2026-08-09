Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
ThisWorkbook.Unprotect Password:=getmsgstate
End Sub

Private Sub Workbook_Open()
mdInitializer.Intialize
ThisWorkbook.Protect Password:=getmsgstate

CalculateTaxCheckFlag = False
Sheet1.Shapes("Button 240").Visible = True
Sheet1.Shapes("Button 242").Visible = True
Sheet1.Shapes("Button 243").Visible = True



''Ayush_29/07
'Dim due1, due2 As String
'
'due1 = "2025-09-15"
'
'due2 = "2025-09-16"
'
'Dim sheet1_ReturnFileSec_duplicate As Variant
'
'sheet1_ReturnFileSec_duplicate = Sheet1.Range("sheet1.ReturnFileSec").value
'
''Updated by Sadineni_29/07/2025
'
'If Dformat(Sheet5.Range("Ver.Date"), "") <= due1 Then
'
'    'Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'
'    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(4)-Belated" Or Sheet1.Range("sheet1.ReturnFileSec") = "(Select)" Then
'
'           Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'
''    Else
'
''        Sheet1.Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
'
'    End If
'
'     If Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
'
'       If Range("sheet1.ReceiptNo").MergeArea.Locked = False Or Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False Then
'
'        Sheet1.Unprotect Password:=getmsgstate
'
'        Application.EnableEvents = False
'
'                Range("sheet1.ReceiptNo").value = ""
'
'                Range("sheet1.ReceiptNo").MergeArea.Locked = True
'
'                Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)
'
'                Range("sheet1.OrigRetFiledDate").value = ""
'
'                Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
'
'                Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)
'
'        Application.EnableEvents = True
'
'        End If
'
'     End If
'
'End If
'
'
'If Dformat(Sheet5.Range("Ver.Date"), "") >= due2 And Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date" Then
'
'    Sheet1.Range("sheet1.ReturnFileSec").value = "139(4)-Belated"
'
'        ElseIf Dformat(Sheet5.Range("Ver.Date"), "") >= due2 Then
'
'        Sheet1.Range("sheet1.ReturnFileSec").value = sheet1_ReturnFileSec_duplicate
'
'End If
'
'
''*****************************************************

'Ayush_22/07

'Dim due1 As String
'
'
'due1 = "2025-09-15"
'
'
'
'If Dformat(Sheet5.Range("Ver.Date"), "") > due1 Then
'
'Sheet1.Range("sheet1.ReturnFileSec").value = "139(4)-After due date"
'
'Else
'
'Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'
'End If
'
End Sub
 
