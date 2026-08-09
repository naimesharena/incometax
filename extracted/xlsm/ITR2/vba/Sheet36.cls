Attribute VB_Name = "Sheet36"
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

'setTblinfo_QDTradingConcern
'setTblinfo_QDRawMaterial
'setTblinfo_QDFinishrByProd
'Sch80G.setTblinfo_Per10080G
'Sch80G.setTblinfo_PerYES10080G
'Sch80G.setTblinfo_PerNO5080G
'Sch80G.setTblinfo_Per5080G
'setTblinfo_hprptfrm
'
'Sheet36.Range("PreXML_QD_A").value = end_QDTradingConcern
'Sheet36.Range("PreXML_QD_B").value = end_QDRawMaterial
'S'heet36.Range("PreXML_QD_C").value = end_QDFinishrByProd
'Sheet36.Range("PreXML_SCH80GA").value = end_Per10080G
'Sheet36.Range("PreXML_SCH80GB").value = end_PerYES10080G
'Sheet36.Range("PreXML_SCH80GC").value = end_PerNO5080G
'Sheet36.Range("PreXML_SCH80GD").value = end_Per5080G
'Sheet36.Range("PreXML_SCHS").value = noOfSalFilled
'Sheet36.Range("PreXML_SCHP").value = end_hprptfrm

endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Deactivate()
'ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
ThisWorkbook.Unprotect Password:=getmsgstate
Sheet36.Visible = xlSheetHidden
ThisWorkbook.Protect Password:=getmsgstate
'ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate
CalculateTaxCheckFlag = True
End Sub
