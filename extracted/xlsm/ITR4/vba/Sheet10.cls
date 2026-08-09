Attribute VB_Name = "Sheet10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Dim sPassword As Variant

Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False
fmsgbox "* Please verify that the TDS statement is submitted for complete financial year"
fmsgbox "* To Compute Tax and Interest using this utility,you must click on 'Calculate Tax' button and verify the figures before saving the XML,if you  have not done so,please do it and then Generate XML" ', vbInformation, vbMessageCaption
fmsgbox "* To Save XML Click on the 'Save XML' Button on this Sheet" ', vbInformation, vbMessageCaption
setTblinfo_TDS1
setTblinfo_TDS2
setTblinfo_TDS2ii
setTblinfo_IT
setTblinfo_TCS

sPassword = mdCommon.getmsgstate
Sheet10.Unprotect Password:=sPassword

Sheet10.Range("C12").value = mdTDS.end_TDS1
Sheet10.Range("C13").value = mdTDS.end_TDS2
Sheet10.Range("C14").value = mdTDS.end_TDS2ii
Sheet10.Range("C15").value = mdTDS.end_IT
Sheet10.Range("C16").value = mdTDS.end_TCS

sPassword = mdCommon.getmsgstate
Sheet10.Protect Password:=sPassword

endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Deactivate()
ThisWorkbook.Unprotect Password:=getmsgstate
Sheet10.Visible = xlSheetHidden
ThisWorkbook.Protect Password:=getmsgstate
End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub
