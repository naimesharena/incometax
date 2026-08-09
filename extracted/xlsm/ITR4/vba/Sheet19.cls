Attribute VB_Name = "Sheet19"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
 Application.ScreenUpdating = False

 Application.EnableEvents = False

 

Sheet19.Unprotect Password:=getmsgstate
 
    Target.EntireColumn.AutoFit

Sheet19.Protect Password:=mdCommon.getmsgstate

endline:

Application.EnableEvents = True

Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
