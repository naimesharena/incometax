Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
ThisWorkbook.Close
End Sub
Private Sub Workbook_Open()
    SheetsNotRequiredforITR2
    CalculateTaxCheckFlag = False
    Initializer.Intialize
    Application.DisplayFormulaBar = False
    ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate
End Sub

Sub SheetsNotRequiredforITR2()
ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
Sheet2.Visible = xlSheetHidden
Sheet3.Visible = xlSheetHidden
Sheet4.Visible = xlSheetHidden
Sheet5.Visible = xlSheetHidden
Sheet6.Visible = xlSheetHidden
Sheet9.Visible = xlSheetHidden
Sheet10.Visible = xlSheetHidden
Sheet11.Visible = xlSheetHidden
Sheet12.Visible = xlSheetHidden
Sheet15.Visible = xlSheetHidden
Sheet18.Visible = xlSheetHidden
Sheet37.Visible = xlSheetHidden
Sheet42.Visible = xlSheetHidden
Sheet48.Visible = xlSheetHidden
Sheet49.Visible = xlSheetHidden
Sheet50.Visible = xlSheetHidden
Sheet51.Visible = xlSheetHidden
Sheet57.Visible = xlSheetHidden
ThisWorkbook.Protect Password:=getmsgstate

End Sub
